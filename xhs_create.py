# -*- coding: utf-8 -*-
"""
批量创建比特浏览器窗口

用法：
    # 建 小红书测试101 ~ 小红书测试110（分组取 .env 的 BIT_BROWSER_GROUP）
    python xhs_create.py --start 101 --end 110

    # 从现有最大序号往后建 5 个
    python xhs_create.py --count 5

    # 代理按「代理信息.xlsx」里同名行取（没写就直连）
    python xhs_create.py --start 101 --end 110 --proxy-from-excel

    # 全部统一走同一个代理
    python xhs_create.py --start 101 --end 105 --proxy-type socks5 \
        --proxy-info 1.2.3.4:1080:user:pass

    # 干跑：只打印将要建什么，不真的建
    python xhs_create.py --start 101 --end 110 --dry-run

    # 建完打「未登陆」标签，方便后面扫描
    python xhs_create.py --start 101 --end 110 --tag 未登陆

    # 批量删（同名窗口）
    python xhs_create.py --start 101 --end 110 --delete

重名策略（--on-exist）：
    skip     默认，已存在就跳过
    error    遇到重名直接停
    replace  删掉旧窗口再建（会丢登录态，慎用）
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import sys
import time
from pathlib import Path

import config as C
import bit_client as B

log = logging.getLogger("create")

DEFAULT_PREFIX = "小红书测试"
DEFAULT_PLATFORM = B.DEFAULT_PLATFORM      # 与封禁重建共用同一个常量
RESULT_FILE = Path(__file__).with_name("create_result.json")


# ------------------------------------------------------------------ 序号
def parse_range(spec: str) -> list[int]:
    """'1-5,8,10-12' -> [1,2,3,4,5,8,10,11,12]"""
    out: list[int] = []
    for part in (spec or "").replace("，", ",").split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, _, b = part.partition("-")
            try:
                lo, hi = int(a), int(b)
            except ValueError:
                raise ValueError(f"区间写法错误: {part!r}")
            if lo > hi:
                lo, hi = hi, lo
            out.extend(range(lo, hi + 1))
        else:
            out.append(int(part))
    # 去重保序
    seen, uniq = set(), []
    for n in out:
        if n not in seen:
            seen.add(n)
            uniq.append(n)
    return uniq


def seq_of(b: dict) -> int:
    """比特内部的 seq 字段（全局自增，不代表窗口名里的数字）"""
    try:
        return int(b.get("seq") or 0)
    except (TypeError, ValueError):
        return 0


def name_number(name: str, prefix: str) -> int | None:
    """从窗口名里取序号：'小红书测试101' -> 101"""
    name = (name or "").strip()
    if prefix and not name.startswith(prefix):
        return None
    tail = name[len(prefix):].strip() if prefix else name
    return int(tail) if tail.isdigit() else None


# ------------------------------------------------------------------ 代理
def resolve_proxy(name: str, args, book: dict) -> tuple[str, str, str]:
    """
    返回 (proxy_type, proxy_info, 说明)

    优先级：命令行 --proxy-type/--proxy-info  >  代理信息.xlsx 同名行  >  直连
    """
    if args.proxy_type:
        return args.proxy_type, args.proxy_info or "", "命令行统一指定"

    if args.no_proxy_excel:
        return "noproxy", "", "指定 --no-proxy-excel（直连）"

    info = book.get(name)
    if not info:
        return "noproxy", "", "代理表里没有该窗口 → 直连"
    from proxy_book import normalize_proxy_type, describe
    ptype = normalize_proxy_type(info.get("proxy_type", ""))
    if ptype == "noproxy":
        return "noproxy", "", "代理表标了直连"
    if not info.get("proxy_info"):
        return "noproxy", "", "代理表该行没写代理信息 → 直连"
    return ptype, info.get("proxy_info", ""), describe(info)


# ------------------------------------------------------------------ 主流程
def build_plan(args, existing: dict[str, dict]) -> list[dict]:
    """算出要建哪些窗口"""
    if args.names:
        names = [n.strip() for n in args.names.replace("，", ",").split(",") if n.strip()]
        return [{"name": n, "seq": None} for n in names]

    prefix = args.prefix or DEFAULT_PREFIX

    if args.count:
        prefix = args.prefix or DEFAULT_PREFIX
        nums = [n for n in (name_number(x, prefix) for x in existing) if n is not None]
        start = (max(nums) if nums else 0) + 1
        log.info("--count: 现有「%s」最大序号 %s，从 %d 开始建", prefix,
                 max(nums) if nums else None, start)
        seqs = list(range(start, start + args.count))
    else:
        if args.start is None:
            raise SystemExit("要么给 --count，要么给 --start/--end（或 --start + 数量）")
        if args.end is None:
            end = args.start + (args.num or 1) - 1
        else:
            end = args.end
        if args.start > end:
            raise SystemExit(f"--start {args.start} 大于 --end {end}")
        seqs = list(range(args.start, end + 1))

    return [{"name": f"{prefix}{n}", "seq": n} for n in seqs]


def do_delete(targets: list[dict], existing: dict[str, dict], dry: bool) -> list[dict]:
    results = []
    for t in targets:
        name = t["name"]
        hit = existing.get(name)
        if not hit:
            log.info("[跳过] %s 不存在", name)
            results.append({"name": name, "ok": False, "reason": "不存在"})
            continue
        bid = hit["id"]
        if dry:
            log.info("[干跑] 将删除 %s (%s)", name, bid[:8])
            results.append({"name": name, "ok": True, "dry": True, "id": bid})
            continue
        ok, msg = B.kill_and_delete(bid)
        log.info("[%s] 删除 %s (%s): %s", "成功" if ok else "失败", name, bid[:8], msg)
        results.append({"name": name, "ok": ok, "id": bid, "msg": msg})
        time.sleep(0.3)
    return results


def main(args) -> int:
    try:
        browsers = B.list_browsers(group_name=args.group or None)
    except Exception as e:
        log.error("拉窗口列表失败（比特浏览器客户端没开？）: %s", e)
        return 2

    existing = {(b.get("name") or "").strip(): b for b in browsers if b.get("name")}
    log.info("当前窗口 %d 个（分组=%s）", len(existing), args.group or "全量")

    targets = build_plan(args, existing)

    # 代理表默认读取：代理来源就是 代理信息.xlsx（按窗口名同名行匹配）
    book: dict = {}
    if args.no_proxy_excel:
        if args.proxy_type:
            log.info("代理来源：命令行统一指定 %s", args.proxy_type)
        else:
            log.info("代理来源：--no-proxy-excel，全部直连")
    else:
        from proxy_book import load_proxy_book
        book = load_proxy_book()
        if not book:
            log.warning("代理信息表为空或读不到，全部按直连处理")
        else:
            log.info("代理来源：代理信息.xlsx（%d 条），按窗口名同名行匹配", len(book))
            missing = [t["name"] for t in targets if t["name"] not in book]
            if missing:
                log.warning("代理表里没有这些窗口，将按直连建：%s", ", ".join(missing[:10])
                            + (" ..." if len(missing) > 10 else ""))

    if args.delete:
        results = do_delete(targets, existing, args.dry_run)
        out = RESULT_FILE.with_name("delete_result.json")
        out.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")
        ok = sum(1 for r in results if r["ok"])
        print(f"\n删除完成: 成功 {ok} / {len(results)}")
        print(f"结果: {out}")
        return 0 if ok == len(results) else 1

    # 解析分组
    gid = None
    if args.group:
        try:
            gid, how = B.resolve_group_id(args.group)
            log.info("%s (groupId=%s)", how, gid)
        except Exception as e:
            log.error("分组解析失败: %s", e)
            return 2

    results: list[dict] = []
    created = skipped = failed = 0

    for i, t in enumerate(targets, 1):
        name = t["name"]
        tag = f"[{i}/{len(targets)}]"

        if name in existing:
            if args.on_exist == "skip":
                log.info("%s [跳过] %s 已存在 (%s)", tag, name, existing[name]["id"][:8])
                results.append({"name": name, "ok": True, "skipped": True, "id": existing[name]["id"]})
                skipped += 1
                continue
            if args.on_exist == "error":
                log.error("%s %s 已存在，按 --on-exist error 停止", tag, name)
                break
            if args.on_exist == "replace":
                ptype, pinfo, pdesc = resolve_proxy(name, args, book)
                if args.dry_run:
                    log.info("%s [干跑] 将删除旧窗口并重建 %s | 代理=%s | 指纹=比特随机生成",
                             tag, name, pdesc)
                    results.append({"name": name, "ok": True, "dry": True, "replace": True})
                    continue
                # 走和封禁重建完全相同的入口：先关→再删→再建（比特随机指纹、同分组）
                new_id, msg, ok_del = B.replace_browser(
                    existing[name]["id"], name, platform=args.platform,
                    proxy_type=ptype, proxy_info=pinfo,
                    group_id=gid, tag=args.tag, fingerprint={},
                )
                if new_id:
                    log.info("%s 已重建 %s (%s)，旧窗口删除=%s | 代理=%s",
                             tag, name, new_id[:8], "成功" if ok_del else "失败", pdesc)
                    results.append({"name": name, "ok": True, "id": new_id,
                                    "replaced": True, "old_deleted": ok_del,
                                    "proxy": pdesc, "fingerprint": "bit-random"})
                    created += 1
                else:
                    log.error("%s 重建 %s 失败: %s", tag, name, msg)
                    results.append({"name": name, "ok": False, "reason": msg})
                    failed += 1
                continue

        ptype, pinfo, pdesc = resolve_proxy(name, args, book)

        if args.dry_run:
            log.info("%s [干跑] 将创建 %s | 平台=%s | 代理=%s | 指纹=比特随机生成",
                     tag, name, args.platform, pdesc)
            results.append({"name": name, "ok": True, "dry": True,
                            "proxy_type": ptype, "proxy": pdesc, "fingerprint": "bit-random"})
            continue

        bid, msg = B.create_browser(
            name=name, group_id=gid, platform=args.platform,
            proxy_type=ptype, proxy_info=pinfo, fingerprint={},
        )
        if not bid:
            log.error("%s 创建 %s 失败: %s", tag, name, msg)
            results.append({"name": name, "ok": False, "reason": msg})
            failed += 1
            continue

        log.info("%s 已创建 %s (%s) | 代理=%s | 指纹=%s", tag, name, bid[:8], pdesc, fpdesc)
        results.append({"name": name, "ok": True, "id": bid,
                        "proxy_type": ptype, "proxy": pdesc})

        if args.tag:
            try:
                ok, tmsg = B.tag_browser(bid, tag_name=args.tag)
                if not ok:
                    log.warning("    打标签失败: %s", tmsg)
            except Exception as e:
                log.warning("    打标签异常: %s", str(e)[:80])

        created += 1
        if args.sleep:
            time.sleep(args.sleep)

    RESULT_FILE.write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n完成: 新建 {created} / 跳过 {skipped} / 失败 {failed} / 计划 {len(targets)}")
    print(f"结果: {RESULT_FILE}")
    return 0 if failed == 0 else 1


def build_parser():
    p = argparse.ArgumentParser(description="批量创建比特浏览器窗口")
    p.add_argument("--prefix", default=DEFAULT_PREFIX, help=f"窗口名前缀，默认 {DEFAULT_PREFIX}")
    p.add_argument("--start", type=int, help="起始序号")
    p.add_argument("--end", type=int, help="结束序号（含）")
    p.add_argument("--num", type=int, help="不给 --end 时，从 --start 起建几个")
    p.add_argument("--count", type=int, help="从现有最大序号往后建几个")
    p.add_argument("--names", help="直接指定窗口名，逗号分隔（优先于序号）")
    p.add_argument("--group", default=C.BROWSER_GROUP or None, help="分组名，默认取 .env")
    p.add_argument("--platform", default=DEFAULT_PLATFORM, help=f"平台 URL，默认 {DEFAULT_PLATFORM}")
    p.add_argument("--proxy-type", choices=["noproxy", "socks5", "http", "https"], help="统一代理类型（优先于代理表）")
    p.add_argument("--proxy-info", default="", help="统一代理信息 host:port:user:pass")
    p.add_argument("--no-proxy-excel", action="store_true",
                   help="不读代理信息表（默认会按窗口名自动读表取代理）")
    p.add_argument("--proxy-from-excel", action="store_true",
                   help="(兼容保留) 默认已经读表，这个开关不再需要")
    p.add_argument("--tag", help="创建后打的标签名（如 未登陆）")
    p.add_argument("--on-exist", choices=["skip", "error", "replace"], default="skip",
                   help="重名处理，默认 skip")
    p.add_argument("--delete", action="store_true", help="反向操作：批量删除同名窗口")
    p.add_argument("--sleep", type=float, default=0.3, help="每个窗口之间的间隔秒")
    p.add_argument("--dry-run", action="store_true", help="只打印不执行")
    p.add_argument("-v", "--verbose", action="store_true")
    return p


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    a = build_parser().parse_args()
    logging.basicConfig(level=logging.DEBUG if a.verbose else logging.INFO,
                        format="%(levelname)-7s %(message)s")
    sys.exit(main(a))
