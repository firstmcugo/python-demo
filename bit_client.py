# -*- coding: utf-8 -*-
"""
比特浏览器分组 / 窗口列表查询

按需求：
    1. POST /group/list  {page, pageSize, all:true}  -> 拿到所有分组的 id + 名称
    2. 用配置里的分组名称匹配出 groupId
    3. POST /browser/list {page, pageSize, groupId}  -> 分页拉该分组下的窗口
       注意：page 从 0 开始；pageSize 最大 100，传更大的值也只返回 100 条

字段注意：/group/list 返回的分组名字段是 **groupName**（不是 name）。
"""
from __future__ import annotations

import logging

import requests

import config as C

log = logging.getLogger("bit")

MAX_PAGE_SIZE = 100          # 接口硬上限

# 新窗口统一用这个平台（批量建窗 / 封禁重建 都走它，别各写一份）
DEFAULT_PLATFORM = "https://www.rednote.com/"


def _post(path: str, payload: dict, timeout: int = 60) -> dict:
    try:
        r = requests.post(f"{C.BIT_API}{path}", json=payload, timeout=timeout)
        j = r.json()
    except Exception as e:
        raise RuntimeError(f"请求 {path} 失败: {e}")
    if not j.get("success"):
        raise RuntimeError(f"{path} 返回失败: {j.get('msg')}")
    return j.get("data") or {}


# ------------------------------------------------------------------ 分组
def bit_groups(page_size: int = 100) -> list[dict]:
    """拉取所有分组（自动翻页）"""
    out: list[dict] = []
    page = 0
    while True:
        data = _post("/group/list", {"page": page, "pageSize": min(page_size, MAX_PAGE_SIZE),
                                     "all": True})
        batch = data.get("list") or []
        out.extend(batch)
        total = int(data.get("totalNum") or 0)
        if len(batch) < page_size or len(out) >= total:
            break
        page += 1
    return out


def group_name_of(g: dict) -> str:
    """兼容 groupName / name 两种字段"""
    return (g.get("groupName") or g.get("name") or "").strip()


def resolve_group_id(name: str | None, groups: list[dict] | None = None) -> tuple[str | None, str]:
    """
    按分组名称找 groupId。返回 (id, 说明)。

    匹配顺序：完全相等 -> 名称包含（唯一命中才用）-> 失败
    name 为空时返回 (None, "未配置分组")，调用方按全量窗口处理。
    """
    want = (name or "").strip()
    if not want:
        return None, "未配置分组（按全量窗口处理）"

    groups = groups if groups is not None else bit_groups()
    names = [group_name_of(g) for g in groups]

    for g in groups:
        if group_name_of(g) == want:
            return g["id"], f"分组精确匹配: {want}"

    hits = [g for g in groups if want in group_name_of(g)]
    if len(hits) == 1:
        return hits[0]["id"], f"分组模糊匹配: {want} -> {group_name_of(hits[0])}"
    if len(hits) > 1:
        raise RuntimeError(f"分组名 {want!r} 匹配到多个: {[group_name_of(g) for g in hits]}，请写全名")

    raise RuntimeError(f"找不到分组 {want!r}；现有分组: {names}")


def describe_groups(groups: list[dict] | None = None) -> str:
    groups = groups if groups is not None else bit_groups()
    return "\n".join(f"  {group_name_of(g)}  (id={g['id']}, 窗口数={g.get('browserCount')})"
                     for g in groups) or "  （无分组）"


# ------------------------------------------------------------------ 窗口列表
def list_browsers(group_name: str | None = None, page_size: int = MAX_PAGE_SIZE,
                  group_id: str | None = None, **filters) -> list[dict]:
    """
    分页拉取窗口列表。

    group_name: 分组名称（None/空 -> 全量）；也可直接传 group_id
    filters:    name / remark / seq / minSeq / maxSeq / sort / ownedByMe / opened
    """
    gid = group_id
    if gid is None and group_name:
        gid, how = resolve_group_id(group_name)
        log.info("%s (groupId=%s)", how, gid)

    size = max(1, min(page_size, MAX_PAGE_SIZE))
    out: list[dict] = []
    page = 0
    while True:
        payload = {"page": page, "pageSize": size}
        if gid:
            payload["groupId"] = gid
        payload.update({k: v for k, v in filters.items() if v is not None})

        data = _post("/browser/list", payload)
        batch = data.get("list") or []
        out.extend(batch)
        total = int(data.get("totalNum") or 0)
        if len(batch) < size or len(out) >= total:
            break
        page += 1
        if page > 1000:            # 兜底，避免异常数据导致死循环
            log.warning("窗口列表翻页超过 1000 页，提前结束")
            break

    log.info("窗口列表: 分组=%s 共 %d 个（翻页 %d 次）", gid or "全量", len(out), page + 1)
    return out


def list_browser_ids(group_name: str | None = None, **filters) -> list[str]:
    return [b["id"] for b in list_browsers(group_name=group_name, **filters)]


# ------------------------------------------------------------------ 窗口生命周期
def close_browser(browser_id: str, timeout: int = 60) -> tuple[bool, str]:
    """关闭窗口（可再从比特里打开）"""
    try:
        r = requests.post(f"{C.BIT_API}/browser/close", json={"id": browser_id}, timeout=timeout)
        j = r.json()
        if j.get("success"):
            return True, "已关闭"
        return False, str(j.get("msg") or "关闭失败")
    except Exception as e:
        return False, f"关闭异常: {str(e)[:100]}"


def delete_browser(browser_id: str, timeout: int = 60) -> tuple[bool, str]:
    """
    **彻底删除**浏览器窗口：POST /browser/delete {"id": "..."}

    ⚠ 这是彻底删除接口，删除后无法从回收站找回。
    用于「账号被封禁/禁言」的窗口：这类窗口的 IP/指纹已被平台标记，
    继续复用只会拖累后续账号，必须连窗口一起清掉。

    返回 (是否成功, 平台消息)。
    """
    try:
        r = requests.post(f"{C.BIT_API}/browser/delete", json={"id": browser_id},
                          timeout=timeout)
        try:
            j = r.json()
        except Exception:
            return False, f"返回非 JSON: {(r.text or '')[:120]}"
        if j.get("success"):
            return True, str(j.get("msg") or "已彻底删除")
        return False, str(j.get("msg") or f"删除失败(http {r.status_code})")
    except Exception as e:
        return False, f"删除异常: {str(e)[:100]}"


def create_browser(name: str, group_id: str | None = None, platform: str | None = None,
                   proxy_type: str = "noproxy", proxy_info: str = "",
                   fingerprint: dict | None = None, remark: str | None = None,
                   timeout: int = 60) -> tuple[str | None, str]:
    """
    创建浏览器窗口：POST /browser/add

    fingerprint: 指纹对象；
                - None 或 {} → 交给比特随机生成（官方：传 {} 即可，指纹值留空会随机）
                - dict  → 用传入的指纹（如需固定某些字段）

    返回 (新窗口 id, 说明)；失败时 id 为 None。
    """
    gid = group_id
    if not gid:
        try:
            gid, how = resolve_group_id(C.BROWSER_GROUP or None)
            if not gid:
                # 没配置分组：让比特自己建 API 分组
                gid = None
        except Exception as e:
            log.warning("解析分组失败，交给比特默认分组: %s", str(e)[:80])
            gid = None

    clean_name = (name or "").strip()
    if not clean_name:
        log.warning("create_browser 名称为空，比特会自动命名（建议调用方传入原名）")

    payload: dict = {
        "name": clean_name,
        "proxyMethod": 2,
        "proxyType": proxy_type or "noproxy",
        "platform": platform or DEFAULT_PLATFORM,
        # 必传字段；空对象 = 随机指纹
        "browserFingerPrint": fingerprint if isinstance(fingerprint, dict) else {},
    }
    if gid:
        payload["groupId"] = gid
    if remark:
        payload["remark"] = remark

    # 代理字段
    if (proxy_type or "").lower() not in ("", "noproxy"):
        parsed = {}
        try:
            from proxy_book import split_proxy
            parsed = split_proxy(proxy_info)
        except Exception:
            parsed = {}
        if parsed:
            payload.update({"host": parsed["host"], "port": parsed["port"]})
            if parsed.get("user"):
                payload["proxyUserName"] = parsed["user"]
            if parsed.get("password"):
                payload["proxyPassword"] = parsed["password"]
            if parsed.get("is_ipv6"):
                payload["isIpv6"] = "true"
        else:
            log.warning("窗口 %s 代理信息无法解析(%r)，按直连创建", name, proxy_info)
            payload["proxyType"] = "noproxy"

    try:
        r = requests.post(f"{C.BIT_API}/browser/add", json=payload, timeout=timeout)
        try:
            j = r.json()
        except Exception:
            return None, f"返回非 JSON: {(r.text or '')[:120]}"
        if j.get("success"):
            data = j.get("data") or {}
            bid = data.get("id") or (data.get("ids") or [None])[0]
            return bid, str(j.get("msg") or "已创建")
        return None, str(j.get("msg") or f"创建失败(http {r.status_code})")
    except Exception as e:
        return None, f"创建异常: {str(e)[:100]}"


def replace_browser(old_id: str, name: str, platform: str | None = None,
                    proxy_type: str = "noproxy", proxy_info: str = "",
                    group_id: str | None = None, tag: str | None = None,
                    remark: str | None = None, fingerprint: dict | None = None,
                    log_=log) -> tuple[str | None, str, bool]:
    """
    删旧建新：先**关闭**再彻底删除 old_id，然后用同样的名称/平台/代理创建一个新窗口（随机指纹）。

    和 xhs_create.py 的批量建窗走同一个 create_browser 入口，参数语义完全一致：
        group_id 不传 -> 按 C.BROWSER_GROUP 解析（与批量建窗默认相同）
        platform 不传 -> DEFAULT_PLATFORM
        fingerprint 不传 -> {} 交给比特随机生成

    tag: 建完打的标签（新窗口必然未登录，传「未登陆」最合适；登录成功会自动清）

    返回 (新窗口 id, 说明, 旧窗口是否删除成功)。
    """
    # 关键：一定要先 /browser/close 再 /browser/delete。
    # 窗口进程还活着时 delete 会失败，失败后仍建同名新窗口 -> 出现两个同名窗口。
    ok_del, msg_del = kill_and_delete(old_id, log_=log_)
    if not ok_del:
        log_.warning("旧窗口删除失败（可能仍在运行），继续创建新窗口；请检查是否有同名窗口")

    # 分组解析和 xhs_create.py 走同一条路：不传就按 C.BROWSER_GROUP 解析
    gid = group_id
    if gid is None:
        try:
            gid, how = resolve_group_id(C.BROWSER_GROUP or None)
            log_.info("%s (groupId=%s)", how, gid)
        except Exception as e:
            log_.warning("分组解析失败，交给比特默认分组: %s", str(e)[:80])
            gid = None

    new_id, msg_new = create_browser(name=name, group_id=gid, platform=platform,
                                     proxy_type=proxy_type, proxy_info=proxy_info,
                                     fingerprint=fingerprint, remark=remark)
    log_.info("创建新窗口 %s: %s (%s)", name, "成功" if new_id else "失败", msg_new)

    if new_id and tag:
        try:
            ok_t, msg_t = tag_browser(new_id, tag_name=tag)
            log_.info("新窗口 %s 打标签「%s」: %s", new_id[:8], tag,
                      "成功" if ok_t else f"失败({msg_t})")
        except Exception as e:
            log_.warning("新窗口打标签异常: %s", str(e)[:80])

    return new_id, msg_new, ok_del


# ------------------------------------------------------------------ 标签
DEFAULT_TAG_COLOR = "#FFF18C"


def list_tags(timeout: int = 30) -> list[dict]:
    """标签列表 /browserTag/list"""
    try:
        r = requests.post(f"{C.BIT_API}/browserTag/list",
                          json={"page": 0, "pageSize": 200}, timeout=timeout)
        j = r.json()
        return j.get("data") or []
    except Exception as e:
        log.warning("取标签列表失败: %s", str(e)[:80])
        return []


def ensure_tag(tag_name: str, tag_color: str = DEFAULT_TAG_COLOR,
               timeout: int = 30) -> str | None:
    """
    确保标签存在：先查列表，没有就 POST /browserTag/create 建一个。
    返回标签 id；失败返回 None。
    """
    name = (tag_name or "").strip()
    if not name:
        return None
    for t in list_tags(timeout):
        if (t.get("tagName") or "").strip() == name:
            return t.get("id")
    try:
        r = requests.post(f"{C.BIT_API}/browserTag/create",
                          json={"tagName": name, "tagColor": tag_color}, timeout=timeout)
        j = r.json()
        if j.get("success"):
            log.info("已创建标签「%s」", name)
        else:
            log.warning("创建标签「%s」失败: %s", name, j.get("msg"))
    except Exception as e:
        log.warning("创建标签「%s」异常: %s", name, str(e)[:80])
    # create 不返回 id，重新查一次
    for t in list_tags(timeout):
        if (t.get("tagName") or "").strip() == name:
            return t.get("id")
    return None


def update_tag_relation(browser_id: str, add_tag_ids: list[str] | None = None,
                        remove_tag_ids: list[str] | None = None,
                        timeout: int = 30) -> tuple[bool, str]:
    """
    更新窗口与标签的绑定关系：POST /browserTag/updateRelation

    body: {"browserId": "...", "addTagIds": [...], "removeTagIds": [...]}
    返回 (是否成功, 平台消息)。
    """
    payload = {
        "browserId": browser_id,
        "addTagIds": add_tag_ids or [],
        "removeTagIds": remove_tag_ids or [],
    }
    try:
        r = requests.post(f"{C.BIT_API}/browserTag/updateRelation", json=payload, timeout=timeout)
        try:
            j = r.json()
        except Exception:
            return False, f"返回非 JSON: {(r.text or '')[:120]}"
        if j.get("success"):
            return True, str(j.get("msg") or "已更新标签关系")
        return False, str(j.get("msg") or f"失败(http {r.status_code})")
    except Exception as e:
        return False, f"异常: {str(e)[:100]}"


def browser_tags(browser_id: str, timeout: int = 30) -> list[dict]:
    """读窗口当前绑定的标签"""
    try:
        r = requests.post(f"{C.BIT_API}/browser/detail", json={"id": browser_id}, timeout=timeout)
        return (r.json().get("data") or {}).get("browserTags") or []
    except Exception:
        return []


def untag_browser(browser_id: str, tag_name: str = "未登陆",
                  timeout: int = 30) -> tuple[bool, str]:
    """
    解除窗口上的某个标签（登录成功后清掉「未登陆」用）。
    窗口本来就没这个标签时直接返回成功。
    """
    current = browser_tags(browser_id, timeout)
    targets = [t.get("id") for t in current
               if (t.get("tagName") or "").strip() == tag_name and t.get("id")]
    if not targets:
        return True, f"窗口没有标签「{tag_name}」"

    ok, msg = update_tag_relation(browser_id, remove_tag_ids=targets, timeout=timeout)
    if ok:
        left = browser_tags(browser_id, timeout)
        names = [t.get("tagName") for t in left]
        log.info("窗口 %s 已移除标签「%s」（剩余: %s）", browser_id[:8], tag_name, names or "无")
        return True, f"已移除标签「{tag_name}」"
    log.warning("移除标签「%s」失败: %s", tag_name, msg)
    return False, f"移除失败: {msg}"


def tag_browser(browser_id: str, tag_name: str = "未登陆",
                tag_color: str = DEFAULT_TAG_COLOR,
                remark: str | None = None, timeout: int = 30) -> tuple[bool, str]:
    """
    给窗口打标签：
      1) ensure_tag：标签不存在就创建（POST /browserTag/create）
      2) update_tag_relation：把标签绑到窗口（POST /browserTag/updateRelation）
      3) 校验浏览器详情里的 browserTags
      4) 附带把 remark 写上，方便人工识别

    返回 (标签是否真的绑上, 说明)。
    """
    tid = ensure_tag(tag_name, tag_color, timeout)
    if not tid:
        return False, f"标签「{tag_name}」创建/获取失败"

    # 已经带这个标签就不用重复绑
    current = browser_tags(browser_id, timeout)
    if any((t.get("id") == tid) for t in current):
        log.info("窗口 %s 已有标签「%s」", browser_id[:8], tag_name)
        return True, f"已带标签「{tag_name}」"

    ok, msg = update_tag_relation(browser_id, add_tag_ids=[tid], timeout=timeout)
    if not ok:
        log.warning("绑定标签「%s」失败: %s", tag_name, msg)
        return False, f"绑定失败: {msg}"

    # 校验
    now = browser_tags(browser_id, timeout)
    bound = any((t.get("id") == tid) for t in now)
    if bound:
        names = ", ".join((t.get("tagName") or "") for t in now)
        log.info("窗口 %s 已打标签「%s」（当前标签: %s）", browser_id[:8], tag_name, names)

    # 附带备注，方便界面上直接看出原因
    if remark:
        try:
            det = {}
            try:
                r = requests.post(f"{C.BIT_API}/browser/detail",
                                  json={"id": browser_id}, timeout=timeout)
                det = r.json().get("data") or {}
            except Exception:
                pass
            requests.post(f"{C.BIT_API}/browser/update", json={
                "id": browser_id,
                "proxyMethod": det.get("proxyMethod", 2),
                "proxyType": det.get("proxyType", "noproxy"),
                "browserFingerPrint": {},
                "remark": remark[:100],
            }, timeout=timeout)
        except Exception as e:
            log.debug("写 remark 失败: %s", str(e)[:60])

    return bound, (f"已打标签「{tag_name}」" if bound else "标签已绑定但校验未读到")


def kill_and_delete(browser_id: str, log_=log) -> tuple[bool, str]:
    """先关窗口再删除（窗口没关时常删不掉），返回删除结果"""
    ok_c, msg_c = close_browser(browser_id)
    log_.info("关闭窗口 %s: %s", browser_id[:8], msg_c)
    import time as _t
    _t.sleep(1.5)                       # 给比特一点时间释放进程
    ok_d, msg_d = delete_browser(browser_id)
    log_.info("删除窗口 %s: %s (%s)", browser_id[:8], "成功" if ok_d else "失败", msg_d)
    return ok_d, msg_d


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")
    gs = bit_groups()
    print("分组列表:")
    print(describe_groups(gs))
    print(f"\n配置的分组名: {C.BROWSER_GROUP!r}")
    if C.BROWSER_GROUP:
        gid, how = resolve_group_id(C.BROWSER_GROUP, gs)
        print(f"解析结果: {how} -> {gid}")
        bs = list_browsers(group_id=gid)
        print(f"该分组窗口数: {len(bs)}")
