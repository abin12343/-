# -*- coding: utf-8 -*-
"""Playwright 网页自动化辅助（复用本机 Edge，无需下载浏览器内核）。

只在真正使用时才 import playwright，没装会给出明确提示而不是莫名崩溃。

典型用法：

    from autolib import browser

    with browser.edge_page("https://example.com") as page:
        browser.auto_login(page, user, pwd)
        browser.wait_stable(page, r"共\\s*([\\d,]+)\\s*条")
"""

from __future__ import annotations

import re
import time
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

from . import paths


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def _require_playwright():
    try:
        from playwright.sync_api import sync_playwright  # noqa: PLC0415
    except ImportError as exc:
        raise SystemExit(
            "缺少 playwright：请先执行  pip install playwright\n"
            "（本库复用本机 Edge，无需执行 playwright install）"
        ) from exc
    return sync_playwright


@contextmanager
def edge_page(url=None, headless=False, maximize=True, accept_downloads=True,
              timeout=60000):
    """打开本机 Edge 并返回 page；退出时自动关闭浏览器。"""
    sync_playwright = _require_playwright()
    with sync_playwright() as p:
        browser = p.chromium.launch(
            channel="msedge", headless=headless,
            args=["--start-maximized"] if maximize else None,
        )
        ctx = browser.new_context(accept_downloads=accept_downloads, no_viewport=True)
        page = ctx.new_page()
        if url:
            page.goto(url, wait_until="domcontentloaded", timeout=timeout)
            if maximize:
                maximize_window(page)
        try:
            yield page
        finally:
            ctx.close()
            browser.close()


def maximize_window(page) -> None:
    """把窗口铺满屏幕。很多后台系统靠右的列（如"应收"）不铺满就不渲染。"""
    try:
        page.evaluate(
            "window.moveTo(0,0); window.resizeTo(screen.availWidth, screen.availHeight);"
        )
        page.wait_for_timeout(800)
    except Exception:  # noqa: BLE001
        pass


def first_visible(page, selectors, timeout=3000):
    """在一组选择器里找第一个可见元素；都找不到返回 None。"""
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


def click_text(page, text, exact=False, timeout=20000) -> bool:
    try:
        loc = page.get_by_text(text, exact=exact).first
        loc.wait_for(state="visible", timeout=timeout)
        loc.click()
        log(f"点击: {text}")
        return True
    except Exception as exc:  # noqa: BLE001
        log(f"点击失败 {text}: {exc}")
        return False


def click_button(page, text, timeout=5000) -> bool:
    """点击可见按钮（按文字精确/包含匹配）。"""
    try:
        btns = page.locator(f'button:has-text("{text}"):visible')
        for i in range(btns.count()):
            t = btns.nth(i).inner_text().strip().replace(" ", "")
            if t == text.replace(" ", "") or text in t:
                btns.nth(i).click(timeout=timeout)
                return True
    except Exception:  # noqa: BLE001
        pass
    return False


def auto_login(page, username, password,
               user_selectors=None, pass_selectors=None,
               btn_selectors=None) -> bool:
    """自动填账号密码并提交。找不到输入框时返回 False，交给调用方等待手动登录。"""
    if not (username and password):
        return False
    u_sel = user_selectors or [
        "input[placeholder*='用户名']", "input[placeholder*=账号 i]",
        "input[placeholder*=用户 i]", "input[type=text]", "input[type=tel]",
    ]
    p_sel = pass_selectors or ["input[placeholder*='密码']", "input[type=password]"]
    b_sel = btn_selectors or [
        "button:has-text('登')", "button:has-text('Sign')", "button:has-text('Login')",
    ]
    try:
        ubox = first_visible(page, u_sel, timeout=8000)
        pbox = first_visible(page, p_sel, timeout=8000)
        if ubox is None or pbox is None:
            log("自动登录未开始：找不到可见的账号/密码输入框（可能未加载、结构不同或在 iframe 内）")
            return False
        ubox.fill(username)
        pbox.fill(password)
        btn = first_visible(page, b_sel, timeout=3000)
        if btn is not None:
            btn.click()
            log("已点击登录按钮")
        else:
            pbox.press("Enter")
            log("已回车提交登录")
        return True
    except Exception as exc:  # noqa: BLE001
        log(f"自动登录异常: {exc}")
        return False


def wait_stable(page, pattern=r"共\s*([\d,]+)\s*条", timeout=90, stable_rounds=2,
                loading_words=("加载中", "查询中")):
    """轮询等待页面条数稳定，返回条数；超时返回最近一次的值（-1 表示从未读到）。

    后台系统常常"先显示旧数据再刷新"，直接 sleep 会误判，所以要等稳定。
    """
    rx = re.compile(pattern)
    deadline = time.time() + max(timeout, 10)
    last, stable = -1, 0
    while time.time() < deadline:
        try:
            body = page.locator("body").inner_text(timeout=8000)
        except Exception:  # noqa: BLE001
            time.sleep(2)
            continue
        if any(w in body for w in loading_words):
            stable = 0
        else:
            m = rx.search(body)
            cnt = int(m.group(1).replace(",", "")) if m else -1
            if cnt == last:
                stable += 1
                if cnt >= 0 and stable >= stable_rounds:
                    return cnt
            else:
                last, stable = cnt, 0
        time.sleep(1)
    log(f"等待结果超时，返回最近条数 {last}")
    return last


def scroll_until(page, keyword, max_rounds=40, step=600, horizontal=True) -> bool:
    """滚动（横向或纵向）直到页面出现关键词，如把"应收"列滚进视野。"""
    sels = (
        ".vxe-table--body-wrapper", ".vxe-table--render-wrapper",
        ".el-table__body-wrapper", ".vxe-table--main-wrapper", ".el-scrollbar__wrap",
    )
    prop = "scrollLeft" if horizontal else "scrollTop"
    for _ in range(max_rounds):
        try:
            if keyword in page.locator("body").inner_text(timeout=8000):
                return True
        except Exception:  # noqa: BLE001
            pass
        moved = False
        for sel in sels:
            loc = page.locator(sel)
            for i in range(loc.count()):
                try:
                    el = loc.nth(i)
                    if el.count() and el.is_visible():
                        if el.evaluate(
                            f"(e)=>{{const b=e.{prop}; e.{prop}+={step}; return e.{prop}!==b;}}"
                        ):
                            moved = True
                            break
                except Exception:  # noqa: BLE001
                    continue
            if moved:
                break
        if not moved:
            break
        page.wait_for_timeout(500)
    try:
        return keyword in page.locator("body").inner_text(timeout=8000)
    except Exception:  # noqa: BLE001
        return False


def collect_lines(page, selectors=("div.line-break", "tr"), max_rounds=20,
                  scroll_step=900):
    """边滚动边收集文本行（虚拟表格按滚动分批渲染，不滚会漏）。"""
    seen, lines = set(), []
    for _ in range(max_rounds):
        added = 0
        for sel in selectors:
            try:
                for t in page.locator(sel).all_inner_texts():
                    key = re.sub(r"\s+", " ", t).strip()
                    if key and key not in seen:
                        seen.add(key)
                        lines.append(key)
                        added += 1
            except Exception:  # noqa: BLE001
                continue
        moved = False
        for sel in (".vxe-table--body-wrapper", ".el-table__body-wrapper",
                    ".el-scrollbar__wrap"):
            loc = page.locator(sel)
            for i in range(loc.count()):
                try:
                    el = loc.nth(i)
                    if el.count() and el.is_visible():
                        if el.evaluate(
                            f"(e)=>{{const b=e.scrollTop; e.scrollBy(0,{scroll_step});"
                            " return e.scrollTop!==b;}"
                        ):
                            moved = True
                            break
                except Exception:  # noqa: BLE001
                    continue
            if moved:
                break
        if not moved:
            try:
                page.mouse.wheel(0, scroll_step)
            except Exception:  # noqa: BLE001
                pass
        page.wait_for_timeout(500)
        if not added and not moved:
            break
    return lines


def dump_debug(page, folder=None, tag="debug"):
    """失败时快照：存 HTML + 截图，便于事后排查。返回 (html路径, png路径)。"""
    d = Path(folder) if folder else paths.output_dir() / "debug"
    d.mkdir(parents=True, exist_ok=True)
    html, png = d / f"{tag}.html", d / f"{tag}.png"
    try:
        html.write_text(page.content(), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    try:
        page.screenshot(path=str(png), full_page=True)
    except Exception:  # noqa: BLE001
        pass
    log(f"已保存调试快照: {d / tag}.*")
    return html, png
