# -*- coding: utf-8 -*-
"""
用途：把《天图核验结果_*_final.csv》中“是否出现=否”的问题单，
      以及可选的其他问题清单，按 运单号/费用名称/金额USD/备注 列追加到用户选择的 xlsx
      （备注列填写核价差异或天图缺标记原因）。
流程：每次运行都会弹出文件选择窗口 -> 自动定位含“永达”的 sheet/表头 -> 写入首个空白行，
      同一 (运单号, 费用名称) 默认跳过重复。
运行：python write_problems_to_stats.py [--tiantu-result <csv>]
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent.parent
else:
    BASE_DIR = Path(__file__).resolve().parent


def latest_file(out_dir: Path, pattern: str):
    cands = sorted(
        out_dir.glob(pattern),
        key=lambda f: f.stat().st_mtime,
        reverse=True,
    )
    return cands[0] if cands else None


def default_output_dir():
    """统一读取 config.json 的输出目录，保证与核价/天图同一处。"""
    try:
        cfg_file = BASE_DIR / "config.json"
        cfg = json.loads(
            cfg_file.read_text(encoding="utf-8")
        )
        p = Path(cfg["paths"]["output_dir"])
        return p if p.is_absolute() else cfg_file.resolve().parent / p
    except Exception:  # noqa: BLE001
        return BASE_DIR / "outputs"


def _fmt_money(v):
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return f"{float(v):.2f}"
    return "" if v is None else str(v)


def _norm_h(v):
    return re.sub(r"\s+", "", str(v or ""))


def diff_remark(amount, quote, diff):
    """核价差异行备注：R-P = 报价-账单。"""
    return (
        f"核价差异 R-P={_fmt_money(diff)}"
        f"（账单P={_fmt_money(amount)}，SOP报价R={_fmt_money(quote)}）"
    )


def _cell(row, idx):
    if idx is None or idx - 1 >= len(row):
        return None
    return row[idx - 1]


def _find_header_row(ws, required=("客户单号",), extra=()):
    """在表头行里定位需要的列；required 全部命中才返回。"""
    for ridx, row in enumerate(ws.iter_rows(min_row=1, max_row=5, values_only=True), start=1):
        cols = {}
        for cidx, c in enumerate(row, start=1):
            t = _norm_h(c)
            if not t:
                continue
            if t == "客户单号":
                cols.setdefault("no", cidx)
            elif t in ("翻译", "费用名称"):
                cols.setdefault("fee", cidx)
            elif t == "费用金额P":
                cols["amount"] = cidx
            elif t == "报价R":
                cols["quote"] = cidx
            elif t == "差异R-P":
                cols["diff"] = cidx
            elif t in extra:
                cols.setdefault(t, cidx)
            elif "金额" in t and "amount" not in cols:
                cols["amount"] = cidx
            elif "备注" in t:
                cols.setdefault("remark", cidx)
        if all(x in cols for x in required):
            return ridx, cols
    return None, None


def load_pricing_items(out_dir: Path):
    """从最新《核价结果_*.xlsx》的问题清单 + 百磅核对表读取核价问题（含备注文案）。"""
    import openpyxl

    xlsx = latest_file(out_dir, "*_核价结果_*.xlsx")
    if xlsx is None:
        return []
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    items = []

    if "问题清单" in wb.sheetnames:
        ws = wb["问题清单"]
        hrow, cols = _find_header_row(ws, required=("no", "fee"))
        if hrow:
            for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
                if not row or not any(v not in (None, "") for v in row):
                    continue
                no = _cell(row, cols.get("no"))
                fee = _cell(row, cols.get("fee"))
                amount = _cell(row, cols.get("amount"))
                quote = _cell(row, cols.get("quote"))
                diff = _cell(row, cols.get("diff"))
                if no not in (None, "") and fee not in (None, ""):
                    items.append({
                        "no": str(no).strip(),
                        "fee": str(fee).strip(),
                        "amount": amount,
                        "remark": diff_remark(amount, quote, diff),
                    })
    if "百磅核对(按票)" in wb.sheetnames:
        ws = wb["百磅核对(按票)"]
        hrow, ci = _find_header_row(
            ws,
            required=("no",),
            extra=("件数", "计费重量合计lb", "费用金额P合计", "报价R合计", "差异R-P", "档位", "状态"),
        )
        if hrow is None:
            return items
        for row in ws.iter_rows(min_row=hrow + 1, values_only=True):
            if not row or not any(v not in (None, "") for v in row):
                continue
            status = _cell(row, ci.get("状态"))
            if str(status or "") != "diff":
                continue
            no = _cell(row, ci.get("no"))
            n = _cell(row, ci.get("件数"))
            lb = _cell(row, ci.get("计费重量合计lb"))
            band = _cell(row, ci.get("档位"))
            amount = _cell(row, ci.get("费用金额P合计"))
            quote = _cell(row, ci.get("报价R合计"))
            diff = _cell(row, ci.get("diff"))
            head = f"{n}件 总重{lb}lb {band}档"
            items.append({
                "no": str(no or "").strip(),
                "fee": "百磅计费运费(按票)",
                "amount": amount,
                "remark": head + " " + diff_remark(amount, quote, diff),
            })
    return items


def main():
    out_dir = default_output_dir()
    ap = argparse.ArgumentParser(description="问题写入统计表（弹窗选文件）")
    ap.add_argument("--tiantu-result", default=None, help="天图核验结果 csv（缺省取最新 final）")
    ap.add_argument("--checklist", default=None, help="天图核验清单 csv（用于取金额）")
    ap.add_argument("--no-dedupe", action="store_true", help="允许重复写入")
    ap.add_argument("--write-target", default=None, help="直接写入指定 xlsx（跳过弹窗）")
    ap.add_argument("--dry-run", action="store_true", help="只打印待写入内容，不弹窗")
    args = ap.parse_args()

    import run_yongda_check as engine

    result_path = Path(args.tiantu_result) if args.tiantu_result else latest_file(
        out_dir, "天图核验结果_*.csv"
    )
    if result_path is None or not result_path.exists():
        print(f"找不到天图核验结果: {result_path}")
        sys.exit(2)
    checklist_path = Path(args.checklist) if args.checklist else latest_file(
        out_dir, "*_天图核验清单_*.csv"
    )

    with open(result_path, encoding="utf-8-sig") as f:
        results = list(csv.DictReader(f))
    amounts = {}
    fees = {}
    if checklist_path and checklist_path.exists():
        with open(checklist_path, encoding="utf-8-sig") as f:
            for t in csv.DictReader(f):
                key = (t.get("客户单号(去后缀)", "").strip(), t.get("期望标记", "").strip())
                if key not in amounts and t.get("金额USD"):
                    try:
                        amounts[key] = float(t["金额USD"])
                    except ValueError:
                        pass
                fees.setdefault(key, t.get("费用名称", "").strip())

    mark_remarks = {
        "偏远": "天图核验：运单号列未见红色“偏远”标记",
        "超偏远": "天图核验：运单号列未见红色“超偏远”标记",
        "住宅私人": "天图核验：应收列未见“住宅私人/私人住宅”等字样",
        "超长": "天图核验：运单号列未见红色“超长”标记",
        "超重": "天图核验：运单号列未见红色“超重”标记",
        "address": "天图核验：运踪信息未见 address 记录",
    }
    tiantu_items = []
    for r in results:
        if r.get("是否出现", "") != "否":
            continue
        no = r.get("单号", "").strip()
        mark = r.get("期望标记", "").strip()
        key = (no, mark)
        fee = r.get("费用名称", "").strip() or fees.get(key, "")
        amount = amounts.get(key)
        if amount is None:
            # 结果文件没有金额时尝试解析
            amount = r.get("金额USD")
        missing_field = str(r.get("缺失关键词", "") or "").strip()
        if missing_field:
            parts = [
                f"运踪信息未见 {kw}"
                for kw in missing_field.split("、")
                if kw.strip()
            ]
            remark = "天图核验：" + "；".join(parts)
        else:
            remark = mark_remarks.get(
                mark, f"天图核验：未找到期望标记“{mark}”"
            )
        tiantu_items.append({"no": no, "fee": fee or mark, "amount": amount, "remark": remark})

    pricing_items = load_pricing_items(out_dir)
    items = list(pricing_items) + list(tiantu_items)
    # 去重：同一 (单号, 费用名称) 只保留一条；重复时合并两条备注
    index = {}
    dedup = []
    for it in items:
        key = (str(it.get("no") or "").strip(), str(it.get("fee") or "").strip())
        if not key[0] or not key[1]:
            continue
        if key in index and not args.no_dedupe:
            old = dedup[index[key]]
            new_rm = str(it.get("remark") or "").strip()
            old_rm = str(old.get("remark") or "").strip()
            if new_rm and new_rm not in old_rm:
                old["remark"] = (old_rm + "；" + new_rm) if old_rm else new_rm
            continue
        index[key] = len(dedup)
        dedup.append(it)
    items = dedup

    if not items:
        print("没有需要写入的问题，流程结束。")
        return
    print(f"待写入 {len(items)} 条（核价 {len(pricing_items)} + 天图缺标记 {len(tiantu_items)}，去重后）：")
    for it in items:
        print(f"  {it['no']} | {it['fee']} | {it['amount']} | {it.get('remark') or ''}")
    if args.dry_run:
        print("[dry-run] 未弹出窗口、未写入。")
        return

    if args.write_target:
        target = Path(args.write_target)
        if not target.exists():
            print(f"找不到 --write-target 文件: {target}")
            return
    else:
        print("请在弹出窗口中选择要写入的统计表 xlsx……")
        target = engine.choose_target_via_dialog()
    if not target:
        print("未选择文件，取消写入。")
        return
    try:
        added, skipped, filled, sheet = engine.append_to_workbook(
            target, items, dedupe=not args.no_dedupe
        )
        print(
            f"已写入 {target} [工作表:{sheet}]：新增 {added} 行，"
            f"跳过重复 {skipped} 行，补写备注 {filled} 行"
        )
    except PermissionError:
        print("写入失败：文件可能正被 Excel 打开。请先关闭该文件后重新运行本步骤。")
    except Exception as exc:  # noqa: BLE001
        print(f"写入失败: {exc}")
        print("请确认选择的是 .xlsx 且未被其他程序占用。")


if __name__ == "__main__":
    main()
