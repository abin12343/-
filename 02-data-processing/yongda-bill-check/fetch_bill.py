# -*- coding: utf-8 -*-
"""
用途：打开 Microsoft Edge 访问打单系统(永达)，导出客户账单，取第一份 xlsx 作为原始数据。
流程对应《永达.docx》：
  1) 打开打单系统-永达（或网址 https://label.wingtatusa.com/#/login）
  2) 登录（账号 TTTX；可手动用 Edge 已保存账号，脚本会等待你登录完成）
  3) 财务管理 -> 客户账单 -> 全部 -> 勾选 -> 导出账单 -> 得到压缩包/账单
  4) 解压后多份 xlsx 只取第一份，源文件不修改
依赖：python3 + playwright（pip install playwright；使用本机 Edge，无需下载浏览器内核）
运行：python fetch_bill.py [--url ...] [--out 存放目录]
     首次登录建议手动在打开的 Edge 页完成，脚本检测到“财务管理”后自动继续。
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import os
import re
import sys
import time
import zipfile
from pathlib import Path


LOGIN_URL = "https://label.wingtatusa.com/#/login"


def log(msg):
    print(f"[{_dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def click_text(page, text, exact=False, timeout=20000):
    loc = page.get_by_text(text, exact=exact)
    loc.first.wait_for(state="visible", timeout=timeout)
    loc.first.click()
    log(f"点击: {text}")


def click_tab(page, label, timeout=20000):
    """点击 Element UI 页签（如 全部/已核销）。"""
    loc = page.locator(".el-tabs__item", has_text=label).first
    loc.wait_for(state="visible", timeout=timeout)
    loc.click()
    log(f"点击页签: {label}")


def maximize_window(page):
    """把浏览器窗口铺满屏幕，避免右下列表列不加载。"""
    try:
        page.evaluate("window.moveTo(0,0); window.resizeTo(screen.availWidth, screen.availHeight);")
        page.wait_for_timeout(800)
    except Exception:  # noqa: BLE001
        pass


def find_download_dir(out_dir):
    d = Path(out_dir)
    d.mkdir(parents=True, exist_ok=True)
    return d


def pick_first_xlsx(folder):
    cands = sorted(
        [f for f in Path(folder).rglob("TTTX_UPS*.xlsx") if not f.name.startswith("~$")],
        key=lambda f: f.name,
    )
    return cands[0] if cands else None


def first_visible(page, selectors, timeout=3000):
    for sel in selectors:
        loc = page.locator(sel).first
        if not loc.count():
            continue
        try:
            loc.wait_for(state="visible", timeout=timeout)
            return loc
        except Exception:  # noqa: BLE001
            continue
    return None


def wait_first_visible(page, selectors, timeout=10):
    """轮询等待任一选择器可见（页面加载慢时用于登录框）。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        loc = first_visible(page, selectors, timeout=800)
        if loc is not None:
            return loc
        time.sleep(0.5)
    return None


def try_auto_login(page, username, password):
    """自动填账号密码并登录：先等待输入框出现（最多约12秒）。"""
    if not username or not password:
        return False
    try:
        ubox = wait_first_visible(page, [
            "input[placeholder*='用户名']", "input[placeholder*=账号 i]",
            "input[placeholder*=用户 i]", "input[placeholder*=账 i]",
            "input[type=text]", "input[type=tel]",
        ], timeout=10)
        pbox = wait_first_visible(page, [
            "input[placeholder*='密码']", "input[placeholder*=密 i]",
            "input[type=password]",
        ], timeout=10)
        if ubox is None or pbox is None:
            log("自动登录未开始：页面上找不到可见的账号/密码输入框（登录框未加载、页面结构不同或在 iframe 内）")
            return False
        ubox.fill(username)
        pbox.fill(password)
        login_btn = first_visible(page, [
            "button:has-text('登')", "button:has-text('登 录')",
            "button:has-text('登 陆')", "input[type=button][value*=登]",
            "button:has-text('Sign')", "button:has-text('Login')",
        ])
        if login_btn is not None:
            login_btn.click()
            log("已点击登录按钮")
        else:
            pbox.press("Enter")
            log("已回车提交登录")
        return True
    except Exception as exc:  # noqa: BLE001
        log(f"自动登录尝试异常: {exc}")
        return False


def dump_debug(page, folder: Path, tag: str):
    folder.mkdir(parents=True, exist_ok=True)
    html = folder / f"{tag}.html"
    png = folder / f"{tag}.png"
    try:
        html.write_text(page.content(), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    try:
        page.screenshot(path=str(png), full_page=True)
    except Exception:  # noqa: BLE001
        pass
    return html, png


def summarize_dom(page):
    """输出页面上表格/复选框的简况，便于定位勾选控件。"""
    try:
        rows = page.locator("table tr").count()
        print(f"[结构] table行数: {rows}")
        if rows:
            for i in range(min(rows, 6)):
                t = page.locator("table tr").nth(i).inner_text(timeout=2000).replace("\n", " | ")[:160]
                print(f"[结构] 第{i}行: {t}")
    except Exception:  # noqa: BLE001
        print("[结构] 未找到 table")
    try:
        vrows = page.locator(".vxe-body--row")
        print(f"[结构] vxe数据行数: {vrows.count()}")
        if vrows.count():
            print("[结构] 首行文本: " + vrows.first.inner_text(timeout=3000).replace("\n", " | ")[:200])
            cb = vrows.first.locator(".vxe-cell--checkbox").count()
            print(f"[结构] 首行复选框: {cb}")
    except Exception:  # noqa: BLE001
        print("[结构] 未找到 vxe 数据行")
    try:
        boxes = page.locator("input[type=checkbox]")
        print(f"[结构] checkbox数量: {boxes.count()}")
        for i in range(min(boxes.count(), 8)):
            b = boxes.nth(i)
            try:
                cls = b.get_attribute("class") or ""
                chk = b.is_checked()
                en = b.is_enabled()
                print(f"[结构] checkbox#{i} checked={chk} enabled={en} class={cls[:60]}")
            except Exception:  # noqa: BLE001
                pass
    except Exception:  # noqa: BLE001
        print("[结构] 未找到 checkbox")


def main():
    ap = argparse.ArgumentParser(description="永达打单系统账单获取")
    ap.add_argument("--url", default=LOGIN_URL)
    ap.add_argument("--out", default=None, help="导出文件存放目录（默认输入目录下自动新建）")
    ap.add_argument("--timeout", type=int, default=300, help="等待手动登录秒数")
    ap.add_argument("--probe", action="store_true", help="只抓页面结构，不导出")
    ap.add_argument("--probe-out", default=None, help="探针页面快照目录")
    ap.add_argument("--login-test", action="store_true", help="只测自动登录，不导出")
    args = ap.parse_args()

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log("缺少 playwright：请先执行  pip install playwright  后重试")
        sys.exit(2)

    base = Path(args.out) if args.out else None
    if base is None:
        here = Path(__file__).resolve().parent
        cfgp = here / "config.json"
        if cfgp.exists():
            cfg = json.loads(cfgp.read_text(encoding="utf-8"))
            base = Path(cfg["paths"]["input_dir"])
    if base is None:
        base = Path.home() / "Downloads"
    stamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = base / f"客户账单导出_auto_{stamp}"
    out_dir.mkdir(parents=True, exist_ok=True)
    log(f"导出目录: {out_dir}")

    from run_yongda_check import load_credentials
    username, password = load_credentials("yongda", "YONGDA_USER", "YONGDA_PASS")
    if username:
        log("已读取本地凭证，将自动登录")

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=False, args=["--start-maximized"])
        ctx = browser.new_context(accept_downloads=True, no_viewport=True)
        page = ctx.new_page()
        page.goto(args.url, wait_until="domcontentloaded", timeout=60000)
        maximize_window(page)
        log("浏览器已打开；如需登录请在页面完成（脚本检测到“财务管理”后自动继续）")

        logged_in = False
        if username and password:
            for _ in range(6):
                try_auto_login(page, username, password)
                try:
                    page.get_by_text("财务管理", exact=True).first.wait_for(
                        state="visible", timeout=12000
                    )
                    logged_in = True
                    break
                except Exception:  # noqa: BLE001
                    time.sleep(2)
        if username and password and not logged_in:
            log("已尝试自动填表 4 次仍未登录成功，接下来等待手动登录")
        if not logged_in:
            try:
                page.get_by_text("财务管理", exact=True).first.wait_for(
                    state="visible", timeout=args.timeout * 1000
                )
                logged_in = True
            except Exception:  # noqa: BLE001
                log("自动登录未成功且等待超时，未检测到“财务管理”，已退出")
                dbg = Path(os.environ.get("TEMP", ".")) / "yongda_probe"
                dump_debug(page, dbg, "login_failed")
                ctx.close()
                browser.close()
                sys.exit(3)
        log("登录成功")
        if args.login_test:
            log("自动登录自测通过：已进入系统")
            ctx.close()
            browser.close()
            return

        click_text(page, "财务管理")
        click_text(page, "客户账单")
        if args.probe:
            time.sleep(8)
            probe_dir = Path(args.probe_out) if args.probe_out else Path(os.environ.get("TEMP", ".")) / "yongda_probe"
            dump_debug(page, probe_dir, "1_客户账单")
            try:
                click_tab(page, "全部")
                log("已点击“全部”")
            except Exception:  # noqa: BLE001
                log("未找到“全部”，继续（可能默认已全部）")
            time.sleep(12)
            dump_debug(page, probe_dir, "2_全部")
            summarize_dom(page)
            try:
                rows = page.locator(".vxe-body--row")
                if rows.count():
                    rows.first.locator(".vxe-cell--checkbox").first.wait_for(
                        state="visible", timeout=10000
                    )
                    rows.first.locator(".vxe-cell--checkbox").first.click()
                    time.sleep(2)
                    export_btn = page.locator("button:has-text('导出账单')").first
                    disabled = export_btn.get_attribute("disabled")
                    print(f"[探针] 勾选首行后导出按钮 disabled={disabled}")
                    dump_debug(page, probe_dir, "3_已勾选首行")
                else:
                    print("[探针] 无 vxe 数据行")
            except Exception as exc:  # noqa: BLE001
                print(f"[探针] 勾选首行失败: {exc}")
            print(f"[探针] 页面快照目录: {probe_dir}")
            ctx.close()
            browser.close()
            sys.exit(0)
        # 系统慢：文档要求增加等待时间；先切到“全部”页签让数据行出现
        time.sleep(6)
        click_tab(page, "全部")
        rows = page.locator(".vxe-body--row")
        rows.first.wait_for(state="visible", timeout=60000)
        log(f"账单列表已加载，共 {rows.count()} 行")

        # 只勾选第一行（vxe 表格自绘复选框；不点表头“全选”）
        first_row = rows.first
        first_row.locator(".vxe-cell--checkbox").first.wait_for(
            state="visible", timeout=15000
        )
        first_row.locator(".vxe-cell--checkbox").first.click()
        log("已勾选第一行账单")

        export_btn = page.locator("button:has-text('导出账单')").first
        enabled = False
        for _ in range(30):
            if export_btn.get_attribute("disabled") is None:
                enabled = True
                break
            time.sleep(1)
        if not enabled:
            log("勾选后“导出账单”仍未启用，尝试直接点击第一行后再等一次")
            first_row.click()
            time.sleep(2)
            for _ in range(20):
                if export_btn.get_attribute("disabled") is None:
                    enabled = True
                    break
                time.sleep(1)
        if not enabled:
            dump_debug(page, Path(out_dir), "select_failed")
            log("仍无法启用“导出账单”，页面快照已存，请人工处理")
            ctx.close()
            browser.close()
            sys.exit(5)

        log("等待下载……（系统慢，最多等 5 分钟）")
        dl_dir = Path(out_dir)
        downloaded = []
        try:
            with page.expect_download(timeout=300000) as dl_info:
                click_text(page, "导出账单")
                # SOP：导出后点击“客户账单带出压缩包”；若按钮不存在（直接下载）则忽略
                try:
                    page.get_by_text("客户账单带出压缩包", exact=False).first.click(timeout=5000)
                except Exception:  # noqa: BLE001
                    pass
            dl = dl_info.value
            target = dl_dir / (dl.suggested_filename or "账单下载.zip")
            dl.save_as(str(target))
            downloaded.append(target)
            log(f"下载完成: {target}")
        except Exception:  # noqa: BLE001
            log("未等到自动下载；请在弹出的页面手动下载到上述导出目录，然后按回车")
            input(">>> 下载完成后按回车继续……")

        # 处理压缩包：解压后取第一份 xlsx
        for f in downloaded:
            if f.suffix.lower() == ".zip":
                with zipfile.ZipFile(f) as zf:
                    zf.extractall(out_dir)
                log(f"已解压 {f.name}")
        first = pick_first_xlsx(out_dir)
        if first is None and downloaded:
            first = downloaded[0] if downloaded[0].suffix.lower() == ".xlsx" else None
        if first is None:
            log("没有找到 TTTX_UPS*.xlsx；请检查下载目录")
            ctx.close()
            browser.close()
            sys.exit(4)
        log(f"采用第一份原始账单: {first}")
        log(f"下一步运行: python run_yongda_check.py --input {first}")
        ctx.close()
        browser.close()


if __name__ == "__main__":
    main()
