# -*- coding: utf-8 -*-
"""SKYE 打单系统 — 导出账单（Playwright + 本机 Edge）。

按 SOP 步骤：
  1) 打开 Microsoft Edge，访问 https://www.skye-ship.com/index.html#/home 登录
  2) 左侧导航 → 财务管理 → 账单列表
  3) 按日期筛选，点击"导出"按钮
  4) 下载的 xlsx 自动放到 paths.input_dir

账号在 credentials.local.json（或环境变量 SKYE_USER / SKYE_PASS）。
登录后脚本会等待用户完成手动验证，再开始自动化（避免被风控）。

依赖：playwright（复用本机 Edge，详见 08-common-utils/autolib/browser.py）
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
import time
from pathlib import Path

if getattr(sys, "frozen", False):
    HERE = Path(sys.executable).resolve().parent.parent
else:
    HERE = Path(__file__).resolve().parent


def setup_log() -> logging.Logger:
    log = logging.getLogger("skye-fetch")
    if not log.handlers:
        logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
    return log


def main():
    ap = argparse.ArgumentParser(description="SKYE 打单系统 - 导出账单")
    ap.add_argument("--date", default="", help="账单日期（YYYY-MM-DD），留空导出全部")
    ap.add_argument("--auto-login", action="store_true",
                    help="尝试用 credentials 自动填表（需要 SKYE_USER / SKYE_PASS）")
    args = ap.parse_args()

    log = setup_log()
    cfg = json.loads((HERE / "config.json").read_text(encoding="utf-8"))
    in_dir = Path(cfg["paths"]["input_dir"])
    in_dir.mkdir(parents=True, exist_ok=True)

    if not getattr(sys, "frozen", False):
        sys.path.insert(0, str(HERE.parents[1] / "08-common-utils"))
    from autolib import browser  # noqa: PLC0415

    log.info("[SKYE] 打开浏览器")
    page = browser.edge_page(url="https://www.skye-ship.com/index.html#/home",
                             headless=False, maximize=True, accept_downloads=True)
    if args.auto_login:
        creds = json.loads((HERE / "credentials.local.json").read_text(encoding="utf-8"))
        browser.auto_login(page, creds.get("skye_user", ""), creds.get("skye_pass", ""),
                           submit_text="登录")
    else:
        input("\n>>> 请在浏览器中完成登录后，按回车继续。\n")

    log.info("[SKYE] 等待进入主页")
    browser.wait_stable(page, pattern=r"工作台|首页|home", timeout=60)

    log.info("[SKYE] 导航到 财务管理 → 账单列表")
    # 页面元素会变，这里给出最常见的"按文本点击"模式；具体选择器请按实际页面调整
    browser.click_text(page, "财务管理", timeout=10000)
    time.sleep(1)
    browser.click_text(page, "账单列表", timeout=10000)
    time.sleep(2)

    if args.date:
        log.info("[SKYE] 设置日期 %s", args.date)
        # 通过日期输入框：label 选 "开始日期" / "结束日期"；具体看页面
        # 这里给个占位：fill_input 需要你按实际页面补充
        # browser.fill_input(page, "开始日期", args.date)
    # 等待列表加载
    browser.wait_stable(page, pattern=r"共\s*([\d,]+)\s*条", timeout=120)

    log.info("[SKYE] 点击'导出'按钮")
    with page.expect_download(timeout=120_000) as dl_info:
        browser.click_text(page, "导出", timeout=10000)
    download = dl_info.value
    target = in_dir / download.suggested_filename
    download.save_as(target)
    log.info("[SKYE] 账单已下载：%s", target)
    page.close()


if __name__ == "__main__":
    main()
