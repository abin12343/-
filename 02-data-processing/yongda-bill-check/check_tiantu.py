# -*- coding: utf-8 -*-
"""
用途：天图系统逐类核验（对账流程第二阶段，分批高效版）。
输入：run_yongda_check.py 生成的《*_天图核验清单_*.csv》。
流程：
  1) 打开天图运单页并登录（TIANTU_USER/TIANTU_PASS 或手动）；
  2) 点击“展开”，清空“创建时间”；
  3) 在“关键字”输入框把【同一种期望标记】的单号用逗号一起输入查询
     （如 偏远 62 单分成几批，每批最多 --batch 个）；
  4) 自动滚动结果区收集每个单号明细行（红色标记），批内缺失的单号自动单查兜底；
  5) 输出《天图核验结果_*.csv》：是 / 否 / 未定位 / 异常(条数)。
运行：python check_tiantu.py [--batch 20] [--limit 30] [--nos ...]
依赖：python3 + playwright（复用本机 Edge）
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import json
import os
import re
import sys
import time
from pathlib import Path


TIANTU_URL = "https://fba.tttxex.com/#/document/waybillCollection/waybill/list"

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent.parent
else:
    BASE_DIR = Path(__file__).resolve().parent


def log(msg):
    print(f"[{_dt.datetime.now():%H:%M:%S}] {msg}", flush=True)


def wait_visible(loc, timeout=1500):
    """等待元素出现且可见；兼容新版 Playwright（is_visible 不再接受 timeout 参数）。"""
    if not loc.count():
        return False
    try:
        loc.wait_for(state="visible", timeout=timeout)
        return True
    except Exception:  # noqa: BLE001
        return False


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


def try_auto_login(page, username, password):
    if not username or not password:
        return False
    try:
        ubox = first_visible(page, [
            "input[type=text]", "input[name*=user i]", "input[name*=account i]",
            "input[placeholder*=账 i]", "input[placeholder*=用户 i]", "input[placeholder*=手机 i]",
        ])
        pbox = first_visible(page, [
            "input[type=password]", "input[name*=pass i]", "input[placeholder*=密 i]",
        ])
        if ubox is None or pbox is None:
            log("自动登录未开始：页面上找不到可见的账号/密码输入框（登录框未加载、页面结构不同或在 iframe 内）")
            return False
        ubox.fill(username)
        pbox.fill(password)
        btn = first_visible(page, [
            "button:has-text('登')", "button:has-text('Sign')", "button:has-text('Login')",
        ])
        if btn is not None:
            btn.click()
            log("已点击登录按钮")
        else:
            pbox.press("Enter")
        return True
    except Exception as exc:  # noqa: BLE001
        log(f"自动登录异常: {exc}")
        return False


def dump_debug(page, folder: Path, tag: str):
    folder.mkdir(parents=True, exist_ok=True)
    try:
        (folder / f"{tag}.html").write_text(page.content(), encoding="utf-8")
    except Exception:  # noqa: BLE001
        pass
    try:
        page.screenshot(path=str(folder / f"{tag}.png"), full_page=True)
    except Exception:  # noqa: BLE001
        pass


def maximize_window(page):
    """把浏览器窗口铺满屏幕，确保右列（应收）与全部分区数据加载。"""
    try:
        page.evaluate("window.moveTo(0,0); window.resizeTo(screen.availWidth, screen.availHeight);")
        page.wait_for_timeout(800)
    except Exception:  # noqa: BLE001
        pass


def logged_in(page):
    if "waybill/list" in page.url or "waybillCollection" in page.url:
        return True
    try:
        if wait_visible(page.get_by_text("运单号", exact=False).first, timeout=1500):
            return True
    except Exception:  # noqa: BLE001
        pass
    return False


def click_visible_button(page, text):
    btns = page.locator(f'button:has-text("{text}"):visible')
    for i in range(btns.count()):
        t = btns.nth(i).inner_text().strip().replace(" ", "")
        if t == text.replace(" ", "") or text in t:
            btns.nth(i).click()
            return True
    return False


def ensure_expanded_search(page):
    """确保高级搜索区展开（能看到 创建时间 字段）。"""
    try:
        editor = page.locator(".search-item-page", has_text="创建时间").first
        for _ in range(2):
            if wait_visible(editor, timeout=1500):
                return
            if click_visible_button(page, "展开"):
                log("已点击“展开”")
                page.wait_for_timeout(1500)
            else:
                return
    except Exception:  # noqa: BLE001
        pass


def clear_create_time(page):
    """真正清空“创建时间”（悬停点击清除图标；无图标则键盘删除）。"""
    try:
        editor = page.locator(".search-item-page", has_text="创建时间").first
        if not wait_visible(editor, timeout=3000):
            return
        editor.hover()
        page.wait_for_timeout(500)
        for sel in (".el-icon-circle-close", ".el-range__close-icon"):
            ic = editor.locator(sel).first
            if wait_visible(ic, timeout=1500):
                ic.click()
                log("已清空“创建时间”")
                page.wait_for_timeout(500)
                return
        for ph in ("开始日期", "结束日期"):
            inp = editor.locator(f"input[placeholder='{ph}']").first
            if wait_visible(inp, timeout=1500):
                inp.click()
                inp.press("Control+A")
                inp.press("Delete")
        page.keyboard.press("Escape")
        log("已通过键盘清空“创建时间”")
        page.wait_for_timeout(500)
    except Exception as exc:  # noqa: BLE001
        log(f"清空“创建时间”失败(继续): {exc}")


def ensure_keyword_mode(page):
    try:
        sel = page.locator(".el-select").first
        cur = (sel.inner_text(timeout=2000) or "").strip()
        if "关键字" in cur:
            return
        sel.click()
        page.locator(".el-select-dropdown__item", has_text="关键字").last.click()
        page.wait_for_timeout(1000)
        log("已切换查询类型为 关键字")
    except Exception:  # noqa: BLE001
        pass


def query_input(page):
    return first_visible(page, [
        "input[placeholder*='查询单号']",
        "input[placeholder*=单号 i]",
        "input[type=text]",
    ], timeout=8000)


def reset_search(page):
    try:
        click_visible_button(page, "重置")
        page.wait_for_timeout(2000)
    except Exception:  # noqa: BLE001
        pass


def ensure_all_tab(page):
    """按用户要求：天图核验需在“全部”状态下查看。"""
    try:
        loc = page.get_by_text(re.compile(r"^全部\s*\(\d+\)\s*$")).first
        if loc.count() and loc.is_visible(timeout=3000):
            loc.click()
            page.wait_for_timeout(1500)
            log("已切换到 全部")
    except Exception:  # noqa: BLE001
        pass


def run_query(page, query_text, wait_ms=15000):
    """输入关键字并查询，轮询等待结果条数稳定后返回（系统卡顿时不超时误判）。"""
    ensure_expanded_search(page)
    clear_create_time(page)
    ensure_keyword_mode(page)
    inp = query_input(page)
    inp.wait_for(state="visible", timeout=8000)
    inp.fill(query_text)
    if not click_visible_button(page, "查询"):
        inp.press("Enter")
    deadline = time.time() + max(wait_ms, 90)
    last = -1
    stable = 0
    while time.time() < deadline:
        try:
            body = page.locator("body").inner_text(timeout=8000)
        except Exception:  # noqa: BLE001
            time.sleep(2)
            continue
        loading = ("加载中" in body) or ("查询中" in body)
        m = re.search(r"共\s*([\d,]+)\s*条", body)
        cnt = int(m.group(1).replace(",", "")) if m else -1
        if loading:
            stable = 0
        else:
            if cnt == last:
                stable += 1
            else:
                last = cnt
                stable = 0
            if cnt >= 0 and stable >= 2:
                # 用户口径：先等查询数据加载，再切到“全部”查看
                ensure_all_tab(page)
                page.wait_for_timeout(2000)
                return cnt
        time.sleep(1)
    log(f"等待结果超时，返回最近条数 {last}")
    return last


def run_query_retry(page, query_text, wait_ms=15000, tries=3):
    """查询失败/超时自动重试，最多 tries 次。"""
    for attempt in range(1, tries + 1):
        if attempt > 1:
            reset_search(page)
            log(f"重试第 {attempt} 次：{query_text[:80]}")
        try:
            n = run_query(page, query_text, wait_ms=wait_ms)
            if n >= 0:
                return n
        except Exception as exc:  # noqa: BLE001
            log(f"查询异常（第 {attempt} 次）: {exc}")
        time.sleep(2)
    return -1


def _collect_once(page):
    seen_lines = set()
    seen_marks = set()
    lines = []
    colors = []
    for _ in range(20):
        added = 0
        sources = page.locator("div.line-break").all_inner_texts()
        sources += page.locator("tr").all_inner_texts()
        for t in sources:
            key = re.sub(r"\s+", " ", t).strip()
            if key and key not in seen_lines:
                seen_lines.add(key)
                lines.append(key)
                added += 1
        # 每一轮滚动后都采集一次“标记”元素的颜色（虚拟表格会按滚动分批渲染）
        for e in collect_mark_colors(page):
            ckey = (e["text"], e["color"], str(e["row"])[:80])
            if ckey not in seen_marks:
                seen_marks.add(ckey)
                colors.append(e)
        scrolled = False
        for sel in (".vxe-table--body-wrapper", ".el-table__body-wrapper", ".el-scrollbar__wrap"):
            loc = page.locator(sel)
            for i in range(loc.count()):
                el = loc.nth(i)
                try:
                    if el.count() and el.is_visible():
                        moved = el.evaluate(
                            "(e)=>{const b=e.scrollTop; e.scrollBy(0, 900); return e.scrollTop!==b;}"
                        )
                        if moved:
                            scrolled = True
                            break
                except Exception:  # noqa: BLE001
                    continue
            if scrolled:
                break
        if not scrolled:
            try:
                page.mouse.wheel(0, 900)
            except Exception:  # noqa: BLE001
                pass
        page.wait_for_timeout(500)
        if not added and not scrolled:
            break
    return lines, colors


def collect_detail_lines(page):
    """滚动收集；第一轮没读到就等 3 秒再补一轮（系统卡顿时结果行渲染慢）。"""
    lines, colors = _collect_once(page)
    if not lines:
        page.wait_for_timeout(3000)
        lines2, colors2 = _collect_once(page)
        lines = lines + lines2
        colors = colors + colors2
    return lines, colors


def query_batch(page, nos, wait_ms=18000):
    n = run_query_retry(page, ",".join(nos), wait_ms=wait_ms)
    lines, colors = collect_detail_lines(page)
    return n, lines, colors


def query_single(page, no):
    n = run_query_retry(page, no, wait_ms=10000)
    lines, colors = collect_detail_lines(page)
    return n, lines, colors


def scroll_to_receivable(page):
    """横向滚动主表直到出现 应收 列。"""
    for _ in range(40):
        if "应收" in page.locator("body").inner_text(timeout=8000):
            return True
        moved = False
        for sel in (".vxe-table--body-wrapper", ".vxe-table--render-wrapper",
                    ".vxe-table--layout-wrapper", ".vxe-table--main-wrapper",
                    ".el-table__body-wrapper", ".vxe-table--header-wrapper"):
            loc = page.locator(sel)
            for i in range(loc.count()):
                el = loc.nth(i)
                try:
                    if wait_visible(el, timeout=600):
                        mv = el.evaluate(
                            "(e)=>{const b=e.scrollLeft; e.scrollLeft+=600; return e.scrollLeft!==b;}"
                        )
                        if mv:
                            moved = True
                            break
                except Exception:  # noqa: BLE001
                    continue
            if moved:
                break
        if not moved:
            break
        page.wait_for_timeout(700)
    return "应收" in page.locator("body").inner_text(timeout=8000)


def scroll_right_max(page, steps=60):
    """一直横向滚动到最右侧，确保运踪/地址等靠右内容被虚拟表格加载。"""
    for _ in range(steps):
        moved = False
        for sel in (".vxe-table--body-wrapper", ".vxe-table--render-wrapper",
                    ".vxe-table--layout-wrapper", ".vxe-table--main-wrapper",
                    ".el-table__body-wrapper", ".vxe-table--header-wrapper",
                    ".el-scrollbar__wrap"):
            loc = page.locator(sel)
            for i in range(loc.count()):
                el = loc.nth(i)
                try:
                    if el.count() and el.is_visible():
                        mv = el.evaluate(
                            "(e)=>{const b=e.scrollLeft; e.scrollLeft+=900; return e.scrollLeft!==b;}"
                        )
                        if mv:
                            moved = True
                            break
                except Exception:  # noqa: BLE001
                    continue
            if moved:
                break
        if not moved:
            break
        page.wait_for_timeout(400)


def open_waybill_detail(page, no):
    """双击包含该单号的运单行，进入“运单信息明细”。"""
    candidates = [
        page.locator("tr").filter(has_text=no),
        page.locator(".vxe-body--row").filter(has_text=no),
        page.locator(".el-table__row").filter(has_text=no),
    ]
    for rows in candidates:
        for i in range(min(rows.count(), 10)):
            try:
                row = rows.nth(i)
                if row.is_visible():
                    row.dblclick(timeout=3000)
                    page.wait_for_timeout(2500)
                    return True
            except Exception:  # noqa: BLE001
                continue
    try:
        clicked = page.evaluate(
            """(no) => {
                const els = document.querySelectorAll('tr, [class*="row"], [class*="item"]');
                for (const el of els) {
                    const t = (el.innerText || '').replace(/\\s+/g, ' ');
                    if (t && t.includes(no) && t.length > 20) {
                        el.dispatchEvent(new MouseEvent('dblclick', {bubbles: true, cancelable: true, view: window}));
                        return true;
                    }
                }
                return false;
            }""",
            no,
        )
        if clicked:
            page.wait_for_timeout(2500)
            return True
    except Exception:  # noqa: BLE001
        pass
    return False


def click_detail_tab(page, tab_name):
    """在运单信息明细中点击指定页签。"""
    for _ in range(5):
        try:
            loc = page.get_by_text(tab_name, exact=False).first
            if loc.count() and loc.is_visible(timeout=1500):
                loc.click(timeout=3000)
                page.wait_for_timeout(1800)
                return True
        except Exception:  # noqa: BLE001
            pass
        page.wait_for_timeout(1200)
    return False


def click_trace_tab(page):
    """在运单信息明细中点击“运踪信息”。"""
    return click_detail_tab(page, "运踪信息")


def collect_detail_texts(page):
    """收集明细弹层/抽屉/正文文本，address、corrected 可能出现在这些区域。"""
    texts = []
    for sel in (".el-dialog", ".el-drawer", ".vxe-modal--wrapper",
                "[class*='detail']", "[class*='trace']", "[class*='logistics']"):
        try:
            loc = page.locator(sel)
            for i in range(min(loc.count(), 10)):
                txt = (loc.nth(i).inner_text(timeout=1500) or "").strip()
                if txt:
                    texts.append(txt)
        except Exception:  # noqa: BLE001
            continue
    try:
        body = page.locator("body").inner_text(timeout=8000)
        if body.strip():
            texts.append(body)
    except Exception:  # noqa: BLE001
        pass
    return texts


def close_waybill_detail(page):
    """关闭运单信息明细，避免影响下一个单号。"""
    for sel in (".el-dialog__headerbtn", ".el-drawer__close-btn",
                ".el-drawer__header button", ".el-icon-close", "button[aria-label='Close']"):
        try:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible(timeout=1000):
                loc.click(timeout=2000)
                page.wait_for_timeout(1200)
                return True
        except Exception:  # noqa: BLE001
            continue
    try:
        page.keyboard.press("Escape")
        page.wait_for_timeout(1200)
    except Exception:  # noqa: BLE001
        pass
    return False


def query_address_single(page, no):
    """地址更正核验：按 SOP 单查 -> 双击进运单明细 -> 运踪信息 -> 识别 address/corrected。"""
    n = run_query_retry(page, no, wait_ms=10000)
    page.wait_for_timeout(1500)
    opened = open_waybill_detail(page, no)
    if opened:
        click_trace_tab(page)
        # 向下滚动运踪区域，确保长节点列表都被渲染
        for _ in range(8):
            try:
                page.evaluate(
                    """() => {
                        const els = document.querySelectorAll('.el-scrollbar__wrap, [class*="scroll"], [class*="timeline"], .el-dialog, .el-drawer');
                        for (const el of els) {
                            if (el.scrollHeight > el.clientHeight) el.scrollTop += 1200;
                        }
                        window.scrollBy(0, 900);
                    }"""
                )
                page.wait_for_timeout(500)
            except Exception:  # noqa: BLE001
                pass
    texts = collect_detail_texts(page)
    seg = "\n".join(texts)
    if missing_text_keywords([seg], ADDRESS_KEYWORDS):
        try:
            dump_debug(page, default_output_dir() / "tiantu_debug", f"addr_no_{no}")
        except Exception:  # noqa: BLE001
            pass
    close_waybill_detail(page)
    if not seg:
        # 兜底：没有弹层时直接抓列表行文本
        lines, _colors = collect_detail_lines(page)
        seg = "\n".join(ln for ln in lines if no in ln)
    return n, [seg] if seg else [], opened


def query_receivable_single(page, no):
    """住宅私人核验：单查后横向滚动到 应收，再用多种方式提取该单号应收相关文本。
    只要应收文字里包含 住宅私人/私人住宅/住宅地址费/私人地址 等字样即算“是”。"""
    n = run_query_retry(page, no, wait_ms=10000)
    page.wait_for_timeout(1500)
    visible = scroll_to_receivable(page)
    page.wait_for_timeout(1000)
    lines, _colors = collect_detail_lines(page)
    # 策略1：优先用包含该单号的数据行文本
    own = [ln for ln in lines if no in ln]
    seg = "\n".join(own) if own else ""
    # 策略1b：JS 直接提取包含该单号的行/抽屉完整文本（兼容非 table 结构的明细区）
    try:
        raw_rows = page.evaluate(
            """(no) => {
                const seen = new Set();
                const out = [];
                const sels = [
                    'tbody tr', '.vxe-body--row', '.el-table__row', '.line-break',
                    '.el-drawer', '.el-dialog', '[class*="detail"]'
                ];
                for (const sel of sels) {
                    for (const el of document.querySelectorAll(sel)) {
                        const t = (el.innerText || '').replace(/\\s+/g, ' ').trim();
                        if (t && t.includes(no) && t.length <= 60000 && !seen.has(t)) {
                            seen.add(t);
                            out.push(t);
                        }
                    }
                }
                return out.sort((a, b) => b.length - a.length).slice(0, 10);
            }""",
            no,
        )
        if raw_rows:
            seg = seg + "\n" + "\n".join(raw_rows)
    except Exception:  # noqa: BLE001
        pass
    # 策略2：直接抓取页面上所有含住宅关键词的可见元素文本（单查结果页仅显示该单号，命中即有效）
    kw_hits = []
    for kw in ("住宅私人地址费", "私人住宅地址费", "住宅私人", "私人住宅", "住宅地址费"):
        try:
            loc = page.get_by_text(kw, exact=False)
            for i in range(min(loc.count(), 20)):
                txt = loc.nth(i).inner_text(timeout=2000) or ""
                if txt.strip():
                    kw_hits.append(txt.strip())
        except Exception:  # noqa: BLE001
            pass
    if kw_hits:
        seg = seg + "\n" + "\n".join(kw_hits)
    # 策略3：用 JS 提取整个 table 或明细面板的文本，作为兜底
    body = page.locator("body").inner_text(timeout=8000)
    # 找到最后一个“应收”表头，取其后较完整内容；同时保留整个 body 做兜底
    idxs = [m.start() for m in re.finditer("应收", body)]
    if idxs:
        seg = seg + "\n" + body[idxs[-1]:]
    else:
        seg = seg + "\n" + body
    # 兜底：应收/列表未命中时，进入运单明细读“费用信息”
    if not receivable_keyword_text(seg):
        try:
            opened = open_waybill_detail(page, no)
            if opened:
                click_detail_tab(page, "费用信息")
                detail_texts = collect_detail_texts(page)
                seg = seg + "\n" + "\n".join(detail_texts)
                close_waybill_detail(page)
        except Exception:  # noqa: BLE001
            pass
    return n, seg, visible


def receivable_keyword_text(seg):
    """应收文字命中口径：包含关键词即可（先去空白兼容夹杂空格/换行），无需独立显示。"""
    compact = re.sub(r"\s+", "", seg or "")
    for kw in ("住宅私人地址费", "私人住宅地址费", "住宅私人", "私人住宅", "住宅地址费"):
        if kw in compact:
            return f"应收含 {kw}"
    return ""


def lines_for_no(lines, no):
    return [ln for ln in lines if no in ln]


def mark_in_line(line, mark):
    """按用户口径判定标记：偏远 不算 超偏远/极度偏远。"""
    text = line
    if mark == "偏远":
        return re.search(r"(?<!超)(?<!极)偏远", text) is not None
    if mark == "超偏远":
        return "超偏远" in text
    if mark in ("超长", "超重", "住宅私人", "address"):
        return mark in text
    return mark in text


RED_COLOR_MARKS = ("偏远", "超偏远", "超长", "超重")
ADDRESS_KEYWORDS = ("address", "corrected")


def missing_text_keywords(lines, keywords):
    """返回 lines 文本里缺失的关键词列表（忽略大小写）。"""
    text = " ".join(lines).lower()
    return [kw for kw in keywords if kw not in text]

_MARK_COLOR_JS = r"""
() => {
  const marks = ['偏远', '超偏远', '超长', '超重'];
  const out = [];
  const seen = new Set();
  const els = document.querySelectorAll('div,span,td,font,b,em,strong,a');
  for (const el of els) {
    const t = (el.innerText || '').trim();
    if (!t || t.length > 6 || !marks.includes(t)) continue;
    const cs = getComputedStyle(el);
    const anc = el.closest('div.line-break,tr,.el-table__row,.vxe-body--row');
    const row = anc ? (anc.innerText || '') : '';
    const key = t + '|' + cs.color + '|' + row.slice(0, 60);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({
      text: t,
      color: cs.color || '',
      row: row.replace(/\s+/g, ' ').trim().slice(0, 260)
    });
  }
  return out;
}
"""


def _is_red_color(color):
    """判断 CSS 颜色是否为红色系（宽容处理 #ff0000 / rgb(255,0,0) / 橙红等）。"""
    s = str(color or "").strip().lower()
    m = re.search(r"rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)", s)
    if m:
        r, g, b = (float(x) for x in m.groups())
    else:
        m = re.fullmatch(r"#([0-9a-f]{6})", s)
        if m:
            h = m.group(1)
            r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
        elif s == "red":
            r, g, b = 255, 0, 0
        else:
            return False
    return r >= 100 and (r - g) >= 40 and (r - b) >= 40


def collect_mark_colors(page):
    """采集页面上精确等于 偏远/超偏远/超长/超重 的元素及其颜色。"""
    try:
        raw = page.evaluate(_MARK_COLOR_JS)
    except Exception:  # noqa: BLE001
        return []
    out = []
    for e in raw or []:
        if not isinstance(e, dict):
            continue
        out.append({
            "text": str(e.get("text") or "").strip(),
            "color": str(e.get("color") or ""),
            "row": str(e.get("row") or ""),
        })
    return out


def resolve_appearance(no, mark, n, lines, colors, kw=None):
    """返回 是/否/未定位/异常；红色类标记按“运单号行内有红色标记”判定。"""
    if kw:
        return "是"
    if not lines:
        return "未定位" if n == 1 else f"异常({n}条)"
    if mark in RED_COLOR_MARKS:
        entries = [e for e in colors if str(e.get("text") or "").strip() == mark]
        if entries:
            # 页面上能取到该标记的独立元素：必须出现红色才算“是”
            red_hit = any(
                _is_red_color(e.get("color")) and no in str(e.get("row") or "")
                for e in entries
            )
            return "是" if red_hit else "否"
        # 结构上取不到独立标记元素（标记混在整段文本里），退回纯文字判断
    text_ok = mark_in_line(" ".join(lines), mark)
    return "是" if text_ok else "否"


def latest_task_csv(out_dir: Path):
    cands = sorted(
        out_dir.glob("*_天图核验清单_*.csv"),
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    return cands[0] if cands else None


def default_output_dir():
    """统一读取 config.json 的输出目录，保证与核价引擎同一处。"""
    try:
        cfg_file = BASE_DIR / "config.json"
        cfg = json.loads(
            cfg_file.read_text(encoding="utf-8")
        )
        p = Path(cfg["paths"]["output_dir"])
        return p if p.is_absolute() else cfg_file.resolve().parent / p
    except Exception:  # noqa: BLE001
        return BASE_DIR / "outputs"


def main():
    ap = argparse.ArgumentParser(description="天图系统核验（分批）")
    ap.add_argument("--csv", default=None)
    ap.add_argument("--url", default=TIANTU_URL)
    ap.add_argument("--batch", type=int, default=13, help="每批同标记单号数量（实测13个能一次返回）")
    ap.add_argument("--limit", type=int, default=0, help="只处理前 N 个任务")
    ap.add_argument("--nos", default="", help="只测指定单号（逗号分隔）")
    ap.add_argument("--timeout", type=int, default=300)
    ap.add_argument("--probe", action="store_true")
    ap.add_argument("--probe-out", default=None)
    args = ap.parse_args()

    out_dir = default_output_dir()
    csv_path = Path(args.csv) if args.csv else latest_task_csv(out_dir)
    if csv_path is None or not csv_path.exists():
        log(f"找不到核验清单 csv: {csv_path}")
        sys.exit(2)
    with open(csv_path, encoding="utf-8-sig") as f:
        tasks = list(csv.DictReader(f))
    nos_filter = [x.strip() for x in args.nos.split(",") if x.strip()]
    if nos_filter:
        tasks = [t for t in tasks if t.get("客户单号(去后缀)", "").strip() in nos_filter]
    if args.limit:
        tasks = tasks[: args.limit]
    log(f"核验清单: {csv_path}，共 {len(tasks)} 条")

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        log("缺少 playwright：pip install playwright")
        sys.exit(2)

    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    from run_yongda_check import load_credentials
    user, password = load_credentials("tiantu", "TIANTU_USER", "TIANTU_PASS")
    if user:
        log("已读取本地凭证，将自动登录")

    # 按期望标记分组，同一组内按出现顺序去重
    groups = {}
    order = []
    for t in tasks:
        mark = t.get("期望标记", "").strip() or "?"
        no = t.get("客户单号(去后缀)", "").strip()
        if not no:
            continue
        if mark not in groups:
            groups[mark] = []
            order.append(mark)
        if no not in groups[mark]:
            groups[mark].append(no)

    with sync_playwright() as p:
        browser = p.chromium.launch(channel="msedge", headless=False, args=["--start-maximized"])
        ctx = browser.new_context(accept_downloads=True, no_viewport=True)
        page = ctx.new_page()
        page.goto(args.url, wait_until="domcontentloaded", timeout=60000)
        maximize_window(page)
        for _ in range(4):
            if logged_in(page):
                break
            try_auto_login(page, user, password)
            page.wait_for_timeout(8000)
        if not logged_in(page):
            log("等待手动登录……")
            for _ in range(max(1, args.timeout // 10)):
                if logged_in(page):
                    break
                page.wait_for_timeout(10000)
        if not logged_in(page):
            log("未能确认已进入运单列表页")
            dump_debug(page, Path(os.environ.get("TEMP", ".")) / "tiantu_probe", "login_failed")
            ctx.close()
            browser.close()
            sys.exit(3)
        log("已进入天图运单页")
        page.wait_for_timeout(3000)

        if args.probe:
            log("探针模式：抽查页面“标记”元素的真实颜色，不做完整核验……")
            for pmark in order:
                if pmark not in RED_COLOR_MARKS:
                    continue
                nos = groups[pmark][: max(1, min(3, args.limit or 3))]
                if not nos:
                    continue
                try:
                    reset_search(page)
                    n, lines, colors = query_batch(page, nos)
                    log(
                        f"[探针] 期望[{pmark}] {','.join(nos)} -> {n} 条, "
                        f"文本行 {len(lines)} 条, 标记元素 {len(colors)} 个"
                    )
                    for e in colors[:80]:
                        hit = any(x in e["row"] for x in nos)
                        log(
                            f"  标记={e['text']} color={e['color']} "
                            f"red={_is_red_color(e['color'])} 行含查询单号={hit} "
                            f"row={e['row'][:100]}"
                        )
                    if args.probe_out:
                        dump_debug(page, Path(args.probe_out), f"probe_{pmark}")
                except Exception as exc:  # noqa: BLE001
                    log(f"[探针] {pmark} 查询异常: {exc}")
            ctx.close()
            browser.close()
            log("探针结束")
            return

        no_info = {}
        processed = 0
        total_batches = sum(max(1, (len(v) + args.batch - 1) // args.batch) for v in groups.values())

        def verify_chunk(chunk, depth=0):
            """批查询；没读全就拆半重试，最小到单查，保证不漏。"""
            nonlocal processed
            if len(chunk) <= 1:
                no = chunk[0]
                try:
                    reset_search(page)
                    n1, lines1, colors1 = query_single(page, no)
                    no_info[(no, mark)] = {
                        "lines": lines_for_no(lines1, no), "n": n1, "colors": colors1,
                    }
                    log(f"  单查 {no} -> {n1} 条")
                except Exception as exc:  # noqa: BLE001
                    no_info[(no, mark)] = {"lines": [], "n": -1, "error": str(exc)[:150]}
                return
            try:
                reset_search(page)
                n, lines, colors = query_batch(page, chunk)
                processed += 1
                log(f"[批] {mark} x{len(chunk)} -> {n} 条")
            except Exception as exc:  # noqa: BLE001
                log(f"[批] 查询失败，拆半重试: {exc}")
                verify_chunk(chunk[: len(chunk) // 2], depth + 1)
                verify_chunk(chunk[len(chunk) // 2:], depth + 1)
                return
            found_map = {no: lines_for_no(lines, no) for no in chunk}
            for no, ls in found_map.items():
                if ls:
                    no_info[(no, mark)] = {"lines": ls, "n": n, "colors": colors}
            missing = [no for no in chunk if not found_map[no]]
            if missing:
                if len(missing) == len(chunk):
                    mid = len(chunk) // 2
                    log(f"[批] {len(chunk)} 个单号只返回 {n} 条，拆半重试")
                    verify_chunk(chunk[:mid], depth + 1)
                    verify_chunk(chunk[mid:], depth + 1)
                else:
                    log(f"[批] {len(missing)}/{len(chunk)} 缺失，拆半补查")
                    verify_chunk(missing, depth + 1)

        for mark in order:
            nos = groups[mark]
            if mark == "住宅私人":
                # 应收列在右侧，横向滚动加载，需要逐单核对
                for no in nos:
                    try:
                        reset_search(page)
                        n, seg, visible = query_receivable_single(page, no)
                        found_text = receivable_keyword_text(seg)
                        if not found_text:
                            dump_debug(page, out_dir / "tiantu_debug", f"resi_no_{no}")
                            log(f"  应收未命中，已保存调试截图/页面: outputs/tiantu_debug/resi_no_{no}")
                        no_info[(no, mark)] = {
                            "lines": [seg.strip()] if seg.strip() else [],
                            "n": n,
                            "kw": found_text,
                        }
                        log(f"  应收核验 {no} -> {n} 条 | {found_text or '未找到住宅/私人字样'}")
                    except Exception as exc:  # noqa: BLE001
                        no_info[(no, mark)] = {"lines": [], "n": -1, "error": str(exc)[:150]}
            elif mark == "address":
                # 地址更正需看靠右的运踪/地址内容，逐单核验避免漏掉 corrected
                for no in nos:
                    try:
                        reset_search(page)
                        n, seg_lines, _visible = query_address_single(page, no)
                        no_info[(no, mark)] = {"lines": seg_lines, "n": n}
                        log(f"  地址更正核验 {no} -> {n} 条")
                    except Exception as exc:  # noqa: BLE001
                        no_info[(no, mark)] = {"lines": [], "n": -1, "error": str(exc)[:150]}
            else:
                for start in range(0, len(nos), args.batch):
                    chunk = nos[start:start + args.batch]
                    verify_chunk(chunk)

        # 对账：输入任务数 == 输出结果数；失败/缺失的任务补查一次
        for t in tasks:
            no = t.get("客户单号(去后缀)", "").strip()
            mark = t.get("期望标记", "").strip()
            key = (no, mark)
            if key in no_info and no_info[key].get("n", -1) >= 0:
                continue
            try:
                reset_search(page)
                if mark == "住宅私人":
                    n, seg, vis = query_receivable_single(page, no)
                    found_text = receivable_keyword_text(seg)
                    if not found_text:
                        dump_debug(page, out_dir / "tiantu_debug", f"resi_no_{no}")
                    no_info[key] = {"lines": [seg.strip()] if seg.strip() else [], "n": n, "kw": found_text}
                elif mark == "address":
                    n, seg_lines, _vis = query_address_single(page, no)
                    no_info[key] = {"lines": seg_lines, "n": n}
                else:
                    n, lines1, colors1 = query_single(page, no)
                    no_info[key] = {
                        "lines": lines_for_no(lines1, no), "n": n, "colors": colors1,
                    }
                log(f"[对账补查] {no} {mark} -> {n} 条")
            except Exception as exc:  # noqa: BLE001
                no_info[key] = {"lines": [], "n": -1, "error": str(exc)[:150]}
        got = sum(1 for t in tasks
                  if (t.get("客户单号(去后缀)", "").strip(), t.get("期望标记", "").strip()) in no_info)
        log(f"[对账] 任务 {len(tasks)} 条，已记录结果 {got} 条")

        results = []
        for t in tasks:
            no = t.get("客户单号(去后缀)", "").strip()
            mark = t.get("期望标记", "").strip()
            r = no_info.get((no, mark), {"lines": [], "n": -1})
            lines = r.get("lines", [])
            n = r.get("n", -1)
            missing_kws = []
            if mark == "address":
                # 地址更正需同时确认 address 与 corrected 两个关键词
                if not lines:
                    appear = "未定位" if n == 1 else f"异常({n}条)"
                else:
                    missing_kws = missing_text_keywords(lines, ADDRESS_KEYWORDS)
                    appear = "是" if not missing_kws else "否"
            else:
                appear = resolve_appearance(
                    no, mark, n, lines, r.get("colors", []), kw=r.get("kw")
                )
            results.append({
                "单号": no,
                "费用名称": t.get("费用名称", ""),
                "期望标记": mark,
                "是否出现": appear,
                "命中条数": n,
                "明细": " | ".join(lines[:2])[:400],
                "缺失关键词": "、".join(missing_kws),
            })

        out_csv = out_dir / f"天图核验结果_{ts}.csv"
        with open(out_csv, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.DictWriter(
                f,
                fieldnames=["单号", "费用名称", "期望标记", "是否出现", "命中条数", "明细", "缺失关键词"],
            )
            w.writeheader()
            w.writerows(results)
        miss = sum(1 for r in results if r["是否出现"] == "否")
        abnormal = sum(1 for r in results if str(r["是否出现"]).startswith(("异常", "未定位")))
        log(f"核验完成：共 {len(results)} 条，明确未出现标记 {miss} 条，需人工复核 {abnormal} 条")
        if len(results) != len(tasks):
            log(f"[对账警告] 输出 {len(results)} 条 != 任务 {len(tasks)} 条，请检查！")
        else:
            log(f"[对账] 输出 {len(results)}/{len(tasks)} 条，条数一致")
        log(f"结果: {out_csv}")
        ctx.close()
        browser.close()


if __name__ == "__main__":
    main()
