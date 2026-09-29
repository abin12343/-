# -*- coding: utf-8 -*-
"""
【模板】网页自动化脚本（Playwright 驱动本机 Edge）

用途：<一句话说明这个脚本做什么>
依赖：Python 3 + playwright（pip install playwright；复用本机 Edge，无需下载内核）
运行：python template_web_automation.py [--limit 5] [--probe]

怎么用这个模板：
  1. 复制本文件到对应功能目录（如 03-network-scraping/），改个业务化的名字
  2. 改 URL、SEARCH_FIELDS、以及 check_one() 里的判定逻辑
  3. 先跑 --probe 看页面结构对不对，再用 --limit 5 小批量试，最后全量

骨架已经处理好的事：
  - 自动登录（凭证从环境变量或 config.local.json 读，不写进代码）
  - 检测不到登录成功时转为等待手动登录，超时才退出
  - 查询后轮询等待条数稳定，系统卡顿也不会误判
  - 分批查询 + 失败自动重试
  - 失败自动保存 HTML + 截图到 outputs/debug/
  - 结果条数与任务条数对账，不一致会告警
"""

from __future__ import annotations

import argparse
import csv
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from autolib import browser  # noqa: E402
from autolib import config as cfgmod  # noqa: E402
from autolib import log, notify, paths  # noqa: E402

APP_NAME = "网页自动化模板"
URL = "https://example.com/#/list"


# ============ 只需要改这里 ============

def login_ready(page) -> bool:
    """判断是否已进入目标页面（登出时会被踢回登录页）。"""
    return "list" in page.url


def search(page, keyword: str) -> int:
    """执行一次查询，返回命中条数。"""
    box = browser.first_visible(
        page, ["input[placeholder*='单号']", "input[type=text]"], timeout=8000
    )
    if box is None:
        browser.log("找不到查询输入框")
        return -1
    box.fill(keyword)
    if not browser.click_button(page, "查询"):
        box.press("Enter")
    return browser.wait_stable(page, r"共\s*([\d,]+)\s*条")


def check_one(page, keyword: str) -> str:
    """核验单个关键词，返回 是 / 否 / 未定位。

    这里放你的业务判定。下面的示例是"页面文本里能否找到关键词"。
    """
    n = search(page, keyword)
    if n < 0:
        return "未定位"
    lines = browser.collect_lines(page)
    hit = any(keyword in ln for ln in lines)
    if not hit:
        browser.dump_debug(page, tag=f"miss_{paths.safe_name(keyword)}")
    return "是" if hit else "否"


# ============ 以下通常不用动 ============


def parse_args():
    ap = argparse.ArgumentParser(description=APP_NAME)
    ap.add_argument("--url", default=URL)
    ap.add_argument("--config", default=None)
    ap.add_argument("--keywords", default="", help="待核验关键词/单号，逗号分隔")
    ap.add_argument("--keywords-file", default=None, help="从 csv 第一列读取")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 个")
    ap.add_argument("--probe", action="store_true", help="只探查页面结构，不批量核验")
    ap.add_argument("--timeout", type=int, default=300, help="等待手动登录秒数")
    ap.add_argument("--no-notify", action="store_true")
    return ap.parse_args()


def load_keywords(args) -> list[str]:
    if args.keywords:
        items = [k.strip() for k in args.keywords.split(",") if k.strip()]
    elif args.keywords_file:
        with open(args.keywords_file, encoding="utf-8-sig") as f:
            items = [
                (r[0] if r else "").strip()
                for r in list(csv.reader(f))[1:]
                if r and r[0].strip()
            ]
    else:
        items = []
    return items[: args.limit] if args.limit else items


def main() -> int:
    args = parse_args()
    conf = cfgmod.load_config(args.config)
    logger = log.setup_logger(Path(__file__).stem, conf)
    timer = log.Timer()
    start_ts = time.time()
    status, note = "成功", ""

    keywords = load_keywords(args)
    if not keywords and not args.probe:
        logger.error("没有待核验项：用 --keywords 或 --keywords-file 指定")
        return 2
    logger.info(f"待核验 {len(keywords)} 项")

    user, pwd = cfgmod.load_credentials(conf, "demo", "DEMO_USER", "DEMO_PASS")

    results = []
    with browser.edge_page(args.url) as page:
        # 登录：先自动，不行就等待手动
        if not login_ready(page):
            for _ in range(4):
                browser.auto_login(page, user, pwd)
                page.wait_for_timeout(8000)
                if login_ready(page):
                    break
        if not login_ready(page):
            logger.info("自动登录未成功，等待手动登录……")
            for _ in range(max(1, args.timeout // 10)):
                if login_ready(page):
                    break
                page.wait_for_timeout(10000)
        if not login_ready(page):
            browser.dump_debug(page, tag="login_failed")
            logger.error("未能进入目标页面，已退出")
            status, note = "失败", "登录失败"
            if not args.no_notify:
                notify.notify_all(
                    conf,
                    notify.build_summary(APP_NAME, status, start_ts, note=note),
                    status=status, start_ts=start_ts,
                )
            return 3
        logger.info("已进入目标页面")

        if args.probe:
            sample = (keywords or [""])[:3]
            for kw in sample:
                n = search(page, kw)
                lines = browser.collect_lines(page)
                logger.info(f"[探针] {kw} -> {n} 条，文本行 {len(lines)} 条")
                for ln in lines[:5]:
                    logger.info(f"    {ln[:120]}")
            return 0

        for i, kw in enumerate(keywords, 1):
            try:
                appear = check_one(page, kw)
            except Exception as exc:  # noqa: BLE001
                logger.warning(f"{kw} 核验异常: {exc}")
                appear, note = "异常", f"{kw}: {exc}"
            results.append({"关键词": kw, "是否出现": appear})
            logger.info(f"[{i}/{len(keywords)}] {kw} -> {appear}")

    out_dir = paths.output_dir(conf)
    paths.ensure_dir(out_dir)
    out_csv = out_dir / paths.stamped_name(Path(__file__).stem, "核验结果", "csv")
    with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["关键词", "是否出现"])
        w.writeheader()
        w.writerows(results)

    miss = sum(1 for r in results if r["是否出现"] == "否")
    abnormal = sum(1 for r in results if r["是否出现"] == "异常")
    logger.info(f"完成：共 {len(results)} 条，未命中 {miss} 条，异常 {abnormal} 条，耗时 {timer.elapsed}s")
    if len(results) != len(keywords):
        logger.warning(f"[对账警告] 输出 {len(results)} 条 != 任务 {len(keywords)} 条")
    logger.info(f"结果: {out_csv}")

    if not args.no_notify:
        notify.notify_all(
            conf,
            notify.build_summary(
                APP_NAME, status, start_ts, note=note,
                extra_lines=[f"未命中 {miss} 条 / 共 {len(results)} 条"],
            ),
            status=status, start_ts=start_ts,
        )
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:  # noqa: BLE001
        print(f"[异常] {exc}")
        raise
