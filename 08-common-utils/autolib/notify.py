# -*- coding: utf-8 -*-
"""通知推送：企业微信群机器人 + 金山文档运行记录。

设计原则：通知是旁支，失败只返回消息、绝不抛异常影响主流程。
只用标准库 http.client，零第三方依赖。
"""

from __future__ import annotations

import http.client
import json
import time
from datetime import datetime

RETRY_TIMES = 3
RETRY_DELAY = 2
TIMEOUT = 10


def post_https_json(host, path, payload, headers=None, timeout=TIMEOUT):
    """标准库 HTTPS POST，返回 (状态码, 响应文本)。"""
    conn = http.client.HTTPSConnection(host, timeout=timeout)
    try:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        conn.request(
            "POST", path, body=body, headers=headers or {"Content-Type": "application/json"}
        )
        resp = conn.getresponse()
        return resp.status, resp.read().decode("utf-8", "replace")
    finally:
        conn.close()


def send_wecom(key_or_url: str, content: str, retries=RETRY_TIMES):
    """企业微信群机器人 text 消息。key_or_url 可以是 key，也可以是完整 webhook URL。"""
    if not key_or_url:
        return False, "[跳过] 未配置企业微信机器人"
    if key_or_url.startswith("http"):
        url = key_or_url
    else:
        url = "https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=" + key_or_url
    host = "qyapi.weixin.qq.com"
    path = "/" + url.split("qyapi.weixin.qq.com/", 1)[-1]
    payload = {"msgtype": "text", "text": {"content": content}}
    last = ""
    for attempt in range(1, retries + 1):
        try:
            code, text = post_https_json(host, path, payload)
            if code == 200:
                try:
                    if json.loads(text).get("errcode") in (0, None):
                        return True, "[OK] 企业微信推送成功"
                except Exception:  # noqa: BLE001
                    return True, "[OK] 企业微信推送成功"
                last = f"企业微信返回: {text[:200]}"
            else:
                last = f"HTTP {code}: {text[:200]}"
        except Exception as exc:  # noqa: BLE001
            last = str(exc)
        if attempt < retries:
            time.sleep(RETRY_DELAY)
    return False, f"[失败] 企业微信推送失败: {last}"


def write_kdocs_record(file_id, token, script_id, sheet_name, fields: dict,
                       retries=RETRY_TIMES):
    """按金山文档 AirScript 接口写入一行运行记录。"""
    if not (file_id and token and script_id):
        return False, "[跳过] 未配置金山文档信息"
    payload = {
        "Context": {
            "argv": {"sheet_name": sheet_name, "sheet_data": [{"fields": fields}]}
        }
    }
    headers = {"Content-Type": "application/json", "AirScript-Token": token}
    path = f"/api/v3/ide/file/{file_id}/script/{script_id}/sync_task"
    last = ""
    for attempt in range(1, retries + 1):
        try:
            code, text = post_https_json("www.kdocs.cn", path, payload, headers)
            if code == 200:
                return True, "[OK] 金山文档写入成功"
            if 400 <= code < 500:
                return False, f"[失败] 客户端错误 ({code}): {text[:200]}"
            last = f"服务端错误 HTTP {code}: {text[:200]}"
        except Exception as exc:  # noqa: BLE001
            last = str(exc)
        if attempt < retries:
            time.sleep(RETRY_DELAY)
    return False, f"[失败] 金山文档写入失败: {last}"


def build_summary(app_name, status, start_ts, stages=None, note="",
                  extra_lines=None) -> str:
    """拼一份适合推送的纯文本摘要。"""
    lines = [
        f"应用：{app_name}",
        f"状态：{status}",
        f"运行时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
        f"运行时长：{round(time.time() - start_ts, 1)} 秒",
    ]
    if stages:
        parts = [f"{s['name']} {s['secs']}s({'完成' if s['ok'] else '失败'})" for s in stages]
        lines.append("阶段：" + " | ".join(parts))
    for line in extra_lines or []:
        lines.append(line)
    if note:
        lines.append(f"备注：{note}")
    return "\n".join(lines)


def notify_all(config: dict, content: str, status="成功", start_ts=None,
               fields: dict | None = None, stages=None, note="", verbose=True):
    """按 config.notify 一次做完：金山文档记录 + 企业微信推送。

    config 形如：
      {"notify": {"wecom_key": "...", "kdocs": {"file_id":..., "token":...,
                  "script_id":..., "sheet_name":"数据表",
                  "field_mapping": {"app_name":"应用名称", ...}},
                  "app_name": "我的脚本"}}
    返回 [(ok, msg), ...]
    """
    nt = (config or {}).get("notify", {}) or {}
    if not nt:
        return [(False, "[跳过] 未配置 notify")]

    results = []
    kd = nt.get("kdocs") or {}
    if kd.get("file_id"):
        fm = kd.get("field_mapping") or {
            "app_name": "应用名称",
            "duration": "运行时长",
            "run_time": "运行时间",
            "account": "账号标识",
            "status": "运行状态",
        }
        payload_fields = dict(fields or {})
        payload_fields.setdefault(fm.get("app_name", "应用名称"), nt.get("app_name", "脚本"))
        payload_fields.setdefault(fm.get("status", "运行状态"), status)
        payload_fields.setdefault(
            fm.get("account", "账号标识"), nt.get("account_id", "脚本")
        )
        payload_fields.setdefault(
            fm.get("duration", "运行时长"),
            round(time.time() - (start_ts or time.time()), 2),
        )
        payload_fields.setdefault(
            fm.get("run_time", "运行时间"), datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        )
        results.append(
            ("金山文档", *write_kdocs_record(
                kd.get("file_id"), kd.get("token"), kd.get("script_id"),
                kd.get("sheet_name", "数据表"), payload_fields,
            ))
        )

    if nt.get("wecom_key"):
        results.append(("企业微信", *send_wecom(nt["wecom_key"], content)))

    if verbose:
        for name, ok, msg in results:
            print(f"[通知] {name}: {msg}")
    return results
