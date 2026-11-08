# -*- coding: utf-8 -*-
"""
代理信息表读取：窗口名称 -> 账号平台 / 代理类型 / 代理信息

路径由 .env 的 PROXY_EXCEL 指定（默认 桌面/代理信息.xlsx），工作表「导入窗口模版」：
    第1-3行是说明，数据从第4行开始
    A列 窗口名称      如 小红书测试1
    C列 账号平台      https://www.rednote.com/
    I列 代理类型      socks5 / http / https / 直连模式
    J列 代理信息      host:port:user:pass   （IPv6 前缀 ipv6:）

用法：
    book = load_proxy_book()                 # {name: {...}}
    info = book.get("小红书测试1")
    payload = build_proxy_payload(info)      # 直接喂给 /browser/add
"""
from __future__ import annotations

import logging
from pathlib import Path

import config as C

log = logging.getLogger("proxy")

# 直连/空的写法归一
DIRECT_ALIASES = {"", "直连", "直连模式", "noproxy", "none", "null", "-"}


def _cell(v) -> str:
    return str(v).strip() if v is not None else ""


def load_proxy_book(path: str | Path | None = None) -> dict[str, dict]:
    """
    读代理信息表，返回 {窗口名称: {platform, proxy_type, proxy_info, row}}。
    文件不存在或读取失败返回空 dict（调用方按"无代理信息"处理）。
    """
    p = Path(path or C.PROXY_EXCEL)
    if not p.exists():
        log.warning("代理信息表不存在: %s", p)
        return {}

    try:
        from openpyxl import load_workbook
        wb = load_workbook(p, data_only=True)
    except Exception as e:
        log.error("读取代理信息表失败: %s", e)
        return {}

    # 优先用「导入窗口模版」表，找不到用第一张
    ws = wb["导入窗口模版"] if "导入窗口模版" in wb.sheetnames else wb.worksheets[0]

    book: dict[str, dict] = {}
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i <= 3:                       # 前 3 行是表头/说明
            continue
        name = _cell(row[0]) if len(row) > 0 else ""
        if not name:
            continue
        book[name] = {
            "row": i,
            "name": name,
            "platform": _cell(row[2]) if len(row) > 2 else "",
            "proxy_type": _cell(row[8]) if len(row) > 8 else "",
            "proxy_info": _cell(row[9]) if len(row) > 9 else "",
            "remark": _cell(row[15]) if len(row) > 15 else "",
        }
    log.info("代理信息表载入 %d 个窗口（%s）", len(book), p)
    return book


def split_proxy(proxy_info: str) -> dict:
    """
    解析 "host:port:user:pass"，支持 "ipv6:[addr]:port:user:pass"。

    返回 {host, port, user, password, is_ipv6}；解析不出 port 时返回 {}。
    """
    s = (proxy_info or "").strip()
    if not s:
        return {}

    is_ipv6 = False
    if s.lower().startswith("ipv6:"):
        is_ipv6 = True
        s = s[5:].strip()

    # IPv6: [addr]:port:user:pass
    if s.startswith("["):
        end = s.find("]")
        if end <= 0:
            return {}
        host = s[1:end]
        rest = s[end + 1:].lstrip(":")
        parts = rest.split(":")
        if not parts or not parts[0].isdigit():
            return {}
        out = {"host": host, "port": int(parts[0]), "is_ipv6": True}
        if len(parts) >= 2:
            out["user"] = parts[1]
        if len(parts) >= 3:
            out["password"] = ":".join(parts[2:])
        return out

    parts = s.split(":")
    if len(parts) < 2 or not parts[1].isdigit():
        return {}
    out = {"host": parts[0], "port": int(parts[1]), "is_ipv6": is_ipv6}
    if len(parts) >= 3:
        out["user"] = parts[2]
    if len(parts) >= 4:
        out["password"] = ":".join(parts[3:])
    return out


def normalize_proxy_type(raw: str) -> str:
    v = (raw or "").strip().lower()
    if v in DIRECT_ALIASES:
        return "noproxy"
    if v in ("http", "https", "socks5", "ssh"):
        return v
    # 未识别的写法按 socks5 处理（表里绝大多数是 socks5）
    log.warning("未知代理类型 %r，按 socks5 处理", raw)
    return "socks5"


def build_proxy_payload(info: dict | None) -> dict:
    """
    把表里一行转成 /browser/add 需要的代理字段。
    直连或无信息时返回 {"proxyMethod": 2, "proxyType": "noproxy"}。
    """
    if not info:
        return {"proxyMethod": 2, "proxyType": "noproxy"}

    ptype = normalize_proxy_type(info.get("proxy_type", ""))
    if ptype == "noproxy":
        return {"proxyMethod": 2, "proxyType": "noproxy"}

    parsed = split_proxy(info.get("proxy_info", ""))
    if not parsed:
        log.warning("窗口 %s 的代理信息无法解析: %r，按直连处理",
                    info.get("name"), info.get("proxy_info"))
        return {"proxyMethod": 2, "proxyType": "noproxy"}

    payload = {
        "proxyMethod": 2,
        "proxyType": ptype,
        "host": parsed["host"],
        "port": parsed["port"],
    }
    if parsed.get("user"):
        payload["proxyUserName"] = parsed["user"]
    if parsed.get("password"):
        payload["proxyPassword"] = parsed["password"]
    if parsed.get("is_ipv6"):
        payload["isIpv6"] = "true"
    return payload


def describe(info: dict | None) -> str:
    if not info:
        return "无代理信息"
    ptype = normalize_proxy_type(info.get("proxy_type", ""))
    parsed = split_proxy(info.get("proxy_info", ""))
    if ptype == "noproxy":
        return f"{info.get('name')} 直连"
    if parsed:
        return (f"{info.get('name')} {ptype} {parsed['host']}:{parsed['port']}"
                + ("(有账号)" if parsed.get("user") else ""))
    return f"{info.get('name')} {ptype} (代理信息解析失败)"


if __name__ == "__main__":
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    logging.basicConfig(level=logging.INFO, format="%(levelname)-7s %(message)s")

    book = load_proxy_book()
    print(f"\n共 {len(book)} 个窗口：")
    for i, (name, info) in enumerate(list(book.items())[:10], 1):
        print(f"  {i:>2}. {describe(info)}")
        print(f"      platform={info['platform']!r}")
        print(f"      payload ={build_proxy_payload(info)}")