# -*- coding: utf-8 -*-
"""
用途：永达 UPS 账单核价检查（第一阶段）
      - 定位最新导出的 TTTX_UPS*.xlsx
      - 按翻译表填充费用中文名
      - 按报价表/固定价/分区阶梯/百磅规则计算报价 R
      - 计算差异 S = R - P，输出问题清单与汇总报告
依赖：Python3 + openpyxl
运行：python run_yongda_check.py [--input 账单或目录] [--tolerance 0.01] [--config config.json]
只读输入文件，输出写入 config.paths.output_dir，日志写入 log_dir。
"""

from __future__ import annotations

import argparse
import bisect
import csv
import datetime as _dt
import json
import logging
import math
import os
import re
import shutil
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter


LOG = logging.getLogger("yongda_check")

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent.parent
else:
    BASE_DIR = Path(__file__).resolve().parent


# ---------- 通用工具 ----------

def norm(s):
    """去掉表头里的空白，便于匹配。"""
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def to_float(v):
    """宽松转金额/重量；支持文本里的逗号、负号与首尾空格。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("￥", "").replace("$", "")
    if s in ("", "-", "--", "None"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def find_row_by_prefix(rows, prefix, col=0):
    for i, row in enumerate(rows):
        if row and norm(row[col]).startswith(prefix):
            return i
    return None


# ---------- 1. 配置与输入定位 ----------

def load_config(path):
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_credentials(prefix, env_user, env_pass):
    """凭证来源优先级：环境变量 > 本地 credentials.local.json（不入库）。"""
    user = os.environ.get(env_user, "")
    password = os.environ.get(env_pass, "")
    fp = BASE_DIR / "credentials.local.json"
    if fp.exists() and (not user or not password):
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
            user = user or data.get(f"{prefix}_user", "")
            password = password or data.get(f"{prefix}_pass", "")
        except Exception:  # noqa: BLE001
            pass
    return user, password


def pick_bill_file(cfg, cli_input):
    """定位最新一次导出的账单；同一批多份 xlsx 只取第一份（按文件名排序）。"""
    if cli_input:
        p = Path(cli_input)
        if p.is_dir():
            cands = sorted(
                [f for f in p.glob("TTTX_UPS*.xlsx")
                 if not f.name.startswith("~$") and "原始备份" not in f.name],
                key=lambda f: f.name,
            )
            if cands:
                return cands[0]
            cands = sorted(
                [f for f in p.rglob("TTTX_UPS*.xlsx")
                 if not f.name.startswith("~$") and "原始备份" not in f.name],
                key=lambda f: f.name,
            )
            if cands:
                return cands[0]
        elif p.is_file():
            return p
        else:
            raise FileNotFoundError(cli_input)
    else:
        root = Path(cfg["paths"]["input_dir"])
        export_dirs = sorted(
            [d for d in root.iterdir() if d.is_dir() and "客户账单导出" in d.name],
            key=lambda d: d.name, reverse=True,
        )
        if export_dirs:
            for d in export_dirs:
                cands = sorted(
                    [f for f in d.glob("TTTX_UPS*.xlsx")
                     if not f.name.startswith("~$") and "原始备份" not in f.name],
                    key=lambda f: f.name,
                )
                if cands:
                    return cands[0]  # 多个数据只取第一个
        cands = sorted(
            [f for f in root.rglob("TTTX_UPS*.xlsx")
             if not f.name.startswith("~$") and "原始备份" not in f.name],
            key=lambda f: f.stat().st_mtime, reverse=True,
        )
        if not cands:
            raise FileNotFoundError(f"在 {cfg['paths']['input_dir']} 下没有找到 TTX_UPS*.xlsx")
        return cands[0]


# ---------- 2. 翻译表 ----------

def load_translation(path):
    """Sheet1 优先，Sheet2 兜底；返回 (dict en->zh, 冲突清单)。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    mapping = {}
    conflicts = set()
    prefer = ["Sheet1", "Sheet2"]
    sheets = [s for s in prefer if s in wb.sheetnames]
    sheets += [s for s in wb.sheetnames if s not in prefer]
    for sn in sheets:
        ws = wb[sn]
        for row in ws.iter_rows(values_only=True):
            if not row or len(row) < 2:
                continue
            en, zh = row[0], row[1]
            if en is None or zh is None:
                continue
            en = str(en).strip()
            zh = str(zh).strip()
            if not en or not zh or en == "费用名称":
                continue
            if en in mapping:
                if mapping[en] != zh:
                    conflicts.add((en, mapping[en], zh, sn))
            else:
                mapping[en] = zh
    return mapping, sorted(conflicts)


# ---------- 3. 报价表 ----------

def _parse_zone_int(cell):
    m = re.fullmatch(r"\s*Zone\s*(\d+)\s*", str(cell or ""), re.I)
    return int(m.group(1)) if m else None


def load_ground_quote(path):
    """主 sheet：识别两段 Zone 表头（商业/住宅），返回 {block:{zone:[(lb,rate)]}}。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb["sheet"]
    rows = list(ws.iter_rows(values_only=True))
    # 找表头行：至少两个单元格含 'Zone'
    header_i = None
    for i, row in enumerate(rows[:10]):
        zc = [j for j, v in enumerate(row) if _parse_zone_int(v) is not None]
        if len(zc) >= 7:
            header_i = i
            break
    if header_i is None:
        raise ValueError("报价表 sheet 找不到 Zone 表头行")
    header = rows[header_i]
    zone_cols = [(j, _parse_zone_int(header[j])) for j in range(len(header))]
    zone_cols = [(j, z) for j, z in zone_cols if z is not None]
    # 按列位置聚成两段（商业靠左、住宅靠右）
    clusters = []
    for j, z in zone_cols:
        if clusters and j - clusters[-1][-1][0] <= 1:
            clusters[-1].append((j, z))
        else:
            clusters.append([(j, z)])
    if len(clusters) < 2:
        raise ValueError("报价表 Zone 表头未识别出商业/住宅两段")
    blocks = {"commercial": clusters[0], "residential": clusters[1]}
    quote = {}
    for bname, cluster in blocks.items():
        first_zone_col = cluster[0][0]
        weight_col = first_zone_col - 1
        wcell = norm(header[weight_col]) if weight_col < len(header) else ""
        if "LB" not in wcell and "重量" not in wcell:
            raise ValueError(f"{bname} 段重量列识别失败: {wcell!r}")
        table = {}
        for j, z in cluster:
            table[z] = []
        for row in rows[header_i + 1 :]:
            lb = to_float(row[weight_col]) if weight_col < len(row) else None
            if lb is None:
                continue
            filled = False
            for j, z in cluster:
                rv = to_float(row[j]) if j < len(row) else None
                if rv is not None:
                    table[z].append((lb, rv))
                    filled = True
            if not filled:
                break
        for z in table:
            table[z].sort(key=lambda x: x[0])
        quote[bname] = table
    return quote


def lookup_zone_rate(table, zone, weight, round_up=True):
    """返回 (rate, matched_lb) 或 (None, None)。按 UPS 规则向上取整到下一个整磅档。"""
    if zone not in table:
        return None, None
    lbs = [x[0] for x in table[zone]]
    if not lbs or weight is None or weight <= 0:
        return None, None
    target = math.ceil(weight - 1e-9) if round_up else weight
    i = bisect.bisect_left(lbs, target)
    if i >= len(lbs):
        return None, None
    return table[zone][i][1], lbs[i]


def load_hundredweight_quote(path, block_choice=1):
    """Sheet1：解析两套百磅费率表，返回 {block: {zone: {band: rate}}}。"""
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    target = None
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows(values_only=True):
            if any(_parse_zone_int(c) is not None for c in row) and norm(row[0]).startswith("Weight"):
                target = sn
                break
        if target:
            break
    if target is None:
        raise ValueError("百磅报价表找不到（需含 Weight + Zone 044 的表）")
    rows = list(wb[target].iter_rows(values_only=True))
    blocks = {}
    # 块内：先找表头（含 Zone），再读 '200 to 499'/'200 - 499' 等行，遇下一个 Zone 表头结束
    i = 0
    while i < len(rows):
        row = rows[i]
        zone_cols = [(j, _parse_zone_int(c)) for j, c in enumerate(row)]
        zone_cols = [(j, z) for j, z in zone_cols if z is not None]
        if not zone_cols or not (norm(row[0]).startswith("Weight") or any(z >= 44 for _, z in zone_cols)):
            i += 1
            continue
        table = {z: {} for _, z in zone_cols}
        i += 1
        while i < len(rows):
            r2 = rows[i]
            if r2 and any(_parse_zone_int(c) is not None for c in r2):
                break  # 下一块表头
            label = str(r2[0] or "").strip() if r2 else ""
            band = None
            if "200 to 499" in label or "200 - 499" in label or "200-499" in label:
                band = "200-499"
            elif "500 to 999" in label or "500 - 999" in label or "500-999" in label:
                band = "500-999"
            elif label.startswith("1000"):
                band = "1000+"
            if band:
                for j, z in zone_cols:
                    rv = to_float(r2[j]) if j < len(r2) else None
                    if rv is not None:
                        table[z][band] = rv
            i += 1
        blocks[len(blocks) + 1] = table
    if not blocks:
        raise ValueError("百磅报价表解析不到费率块")
    return blocks.get(block_choice, blocks[1]), blocks


# ---------- 4. 账单解析 ----------

EXPECTED_HEADERS = {
    "发票日期": "invoice_date", "交易日期": "trade_date", "系统单号": "sys_no",
    "客户单号": "customer_no", "客户编码": "customer_code", "主单号": "main_tracking",
    "子单号": "sub_tracking", "分区": "zone", "重量lbs": "weight_lb", "尺寸inch": "dims",
    "计费重量lbs": "billed_lb", "计费尺寸inch": "billed_dims", "费用类型": "fee_type",
    "费用说明": "fee_name", "翻译": "trans_zh", "费用金额": "amount",
     "第一次收取费用": "first_charged", "报价表": "quote_r", "差异": "diff_s",
 }

# 系统原始导出必含的基础列（不含 SOP 中人工新增的 翻译/第一次收取费用/报价表/差异 四列）
REQUIRED_HEADERS = [
    "系统单号", "客户单号", "客户编码", "主单号", "子单号", "分区",
    "计费重量lbs", "费用类型", "费用说明", "费用金额",
]


def read_bill(path):
    wb = openpyxl.load_workbook(path, data_only=False, read_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    header = [norm(c) for c in rows[0]]
    col = {}
    missing = []
    for zh in REQUIRED_HEADERS:
        key = EXPECTED_HEADERS[zh]
        z = norm(zh)
        idx = None
        for i, h in enumerate(header):
            if h == z:
                idx = i
                break
        if idx is None:
            for i, h in enumerate(header):
                if z[:4] and h.startswith(z[:4]):
                    idx = i
                    break
        if idx is None:
            missing.append(zh)
        else:
            col[key] = idx
    if missing:
        raise ValueError(f"账单缺少列: {missing}; 实际表头: {header}")
    optional_cols = {}
    for zh, key in EXPECTED_HEADERS.items():
        if zh in REQUIRED_HEADERS:
            continue
        z = norm(zh)
        idx = None
        for i, h in enumerate(header):
            if h == z:
                idx = i
                break
        if idx is None:
            for i, h in enumerate(header):
                if z[:4] and h.startswith(z[:4]):
                    idx = i
                    break
        if idx is not None:
            optional_cols[key] = idx
    out = []
    for rno, r in enumerate(rows[1:], start=2):
        def get(k):
            if k in col and col[k] < len(r):
                return r[col[k]]
            if k in optional_cols and optional_cols[k] < len(r):
                return r[optional_cols[k]]
            return None
        fee_name = str(get("fee_name") or "").strip()
        amount = to_float(get("amount"))
        if not fee_name and amount is None and not str(get("customer_no") or "").strip():
            continue
        rec = {
            "row": rno,
            "invoice_date": str(get("invoice_date") or ""),
            "trade_date": str(get("trade_date") or ""),
            "sys_no": str(get("sys_no") or ""),
            "customer_no": str(get("customer_no") or ""),
            "customer_code": str(get("customer_code") or ""),
            "main_tracking": str(get("main_tracking") or ""),
            "sub_tracking": str(get("sub_tracking") or ""),
            "zone": str(get("zone") or "").strip(),
            "weight_lb": to_float(get("weight_lb")),
            "dims": str(get("dims") or ""),
            "billed_lb": to_float(get("billed_lb")),
            "billed_dims": str(get("billed_dims") or ""),
            "fee_type": str(get("fee_type") or ""),
            "fee_name": fee_name,
            "amount": amount,
            "trans_zh": str(get("trans_zh") or ""),
        }
        rec["customer_key"] = rec["customer_no"].split("_")[0]
        out.append(rec)
    return out


# ---------- 5. 核价 ----------

def price_record(rec, cfg, quote, cwt):
    """返回 (R, 状态, 备注)。"""
    fee = rec["fee_name"]
    category = cfg["fee_map"].get(fee)
    if category is None:
        return None, "rule-unknown", f"无规则: {fee}"
    meta = cfg["categories"][category]
    kind = meta["kind"]
    zone = rec["zone"]
    lb = rec["billed_lb"]

    if kind == "info":
        return None, "info", meta.get("note", "")
    try:
        zi = int(zone)
    except (TypeError, ValueError):
        return None, "zone-out-of-scope", f"分区无效: {zone!r}"
    if zi not in (2, 3, 4, 5, 6, 7, 8):
        return None, "zone-out-of-scope", f"分区{zone}不在2-8范围，跳过"
    if kind == "fixed":
        return float(cfg["prices"]["fixed_rates"][category]), "priced", meta.get("note", "")
    if kind == "correction-copy":
        return None, "correction-copy", meta.get("note", "")
    if kind == "zone_price":
        table = cfg["prices"]["zone_tiers"][meta["table_key"]]
        if str(zi) in table:
            return float(table[str(zi)]), "priced", meta.get("note", "")
        try:
            if zi >= 7 and "7plus" in table:
                return float(table["7plus"]), "priced", meta.get("note", "")
        except ValueError:
            pass
        return None, "lookup-failed", f"分区不在阶梯表: {zone}"
    if kind == "lookup":
        block = quote.get(meta["block"], {})
        rate, matched_lb = lookup_zone_rate(block, zi, lb, cfg["run"].get("round_up_lookup", True))
        if rate is None:
            valid = sorted(block.keys())
            return None, "lookup-failed", f"报价表无匹配(分区{zone}/重量{lb!r}); 可用分区: {valid}"
        return round(rate, 2), "priced", f"命中 {matched_lb}lb 档"
    if kind == "lookup_diff":
        # 住宅调整 = 同一分区/重量的住宅报价 - 商业报价（真实账单验证：21.67-16.70=4.97）
        res_rate, res_lb = lookup_zone_rate(
            quote.get("residential", {}), zi, lb, cfg["run"].get("round_up_lookup", True)
        )
        comm_rate, comm_lb = lookup_zone_rate(
            quote.get("commercial", {}), zi, lb, cfg["run"].get("round_up_lookup", True)
        )
        if res_rate is None or comm_rate is None:
            return None, "lookup-failed", f"住宅/商业报价无匹配(分区{zone}/重量{lb!r})"
        return round(res_rate - comm_rate, 2), "priced", (
            f"住宅{res_lb}lb {res_rate} - 商业{comm_lb}lb {comm_rate}"
        )
    if kind == "hundredweight":
        return None, "grouped-cwt", "百磅按票透视汇总核对"
    return None, "rule-unknown", f"类别类型未实现: {kind}"


def build_r_formulas(records, cfg, quote=None):
    """按 SOP：
    - 运费更正/运费修正：通过 G 列子单号找运费行，把运费行 R 公式复制到修正行 R，
      只把公式里 K（重量）的行号改成修正行当前行，其它引用保持运费行不变；
    - 商业/住宅运费 R 列写 VLOOKUP 公式；住宅调整写住宅-商业差公式；
    - 差异 S 列写 =R-P 公式。
    """
    quote_name = Path(cfg["paths"]["quote_file"]).name
    freight_by_sub = {}
    for r in records:
        cat = cfg["fee_map"].get(r["fee_name"])
        if cat in ("ground_commercial", "ground_residential") and r.get("quote") is not None:
            freight_by_sub.setdefault(r.get("sub_tracking", ""), r)
    # 第一遍：修正行先关联运费行（只处理运费更正/修正，其它关联行不做）
    for r in records:
        if r["status"] != "correction-copy":
            continue
        fr = freight_by_sub.get(r.get("sub_tracking", ""))
        if fr is None:
            r["status"] = "correction-no-freight"
            r["note"] = "未找到同 G 列子单号的运费行，未填 R"
            continue
        r["quote"] = fr["quote"]
        r["status"] = "priced"
        r["_freight_row"] = fr
        note = f"R 公式取自运费行(Excel行{fr['row']}, 子单号 {r.get('sub_tracking', '')})"
        # 公式复制到修正行后只有 K 取当前行，因此有报价表时用“当前行重量 + 运费行分区”重算 R，
        # 与 Excel 公式计算结果保持一致
        fr_cat = cfg["fee_map"].get(fr.get("fee_name", ""))
        meta = cfg["categories"].get(fr_cat, {})
        block = meta.get("block") if meta.get("kind") == "lookup" else None
        if block and quote is not None:
            try:
                zi = int(fr["zone"])
            except (TypeError, ValueError):
                zi = None
            if zi is not None and r.get("billed_lb") is not None:
                rate, _matched = lookup_zone_rate(
                    quote.get(block, {}), zi, r["billed_lb"],
                    cfg["run"].get("round_up_lookup", True),
                )
                if rate is not None:
                    r["quote"] = round(rate, 2)
                    note += f"；按当前行 K{r['row']} 与运费行 H{fr['row']} 重算"
        r["note"] = note
    # 第二遍：商业/住宅运费、住宅调整写公式
    for r in records:
        cat = cfg["fee_map"].get(r["fee_name"])
        if r.get("quote") is None:
            continue
        row = r["row"]
        if cat == "ground_commercial":
            r["r_formula"] = (
                f"=VLOOKUP(VALUE(K{row}),'[{quote_name}]sheet'!$B:$I,VALUE(H{row}),0)"
            )
        elif cat == "ground_residential":
            r["r_formula"] = (
                f"=VLOOKUP(VALUE(K{row}),'[{quote_name}]sheet'!$L:$S,VALUE(H{row}),0)"
            )
        elif cat == "residential_adjustment":
            r["r_formula"] = (
                f"=VLOOKUP(VALUE(K{row}),'[{quote_name}]sheet'!$L:$S,VALUE(H{row}),0)"
                f"-VLOOKUP(VALUE(K{row}),'[{quote_name}]sheet'!$B:$I,VALUE(H{row}),0)"
            )
    # 第三遍：修正行复制运费行公式
    for r in records:
        cat = cfg["fee_map"].get(r["fee_name"])
        if cat == "shipping_correction" and r.get("quote") is not None:
            fr = r.get("_freight_row")
            if fr is not None and fr.get("r_formula"):
                row = r["row"]
                # 只把 K（重量）的行号改成修正行当前行，H（分区）等其它引用保持原样
                r["r_formula"] = re.sub(r"\bK\d+", f"K{row}", fr["r_formula"])
    # S 列公式
    for r in records:
        if r.get("status") in ("priced", "ok", "diff") and r.get("amount") is not None:
            r["s_formula"] = f"=R{r['row']}-P{r['row']}"
    return records


def summarize_hundredweight(records, cfg, cwt):
    """按票(客户单号)透视百磅行：总重选档，报价=总重*费率/100。"""
    groups = {}
    order = []
    for r in records:
        cat = cfg["fee_map"].get(r["fee_name"])
        if cat != "hundredweight":
            continue
        key = r["customer_no"] or r["sys_no"] or r["main_tracking"]
        if key not in groups:
            groups[key] = {
                "key": key, "row": r["row"], "main_tracking": r["main_tracking"],
                "customer_no": r["customer_no"], "customer_key": r["customer_key"],
                "sys_no": r["sys_no"],
                "zones": [], "lb": 0.0, "amount": 0.0, "n": 0,
            }
            order.append(key)
        g = groups[key]
        g["n"] += 1
        g["zones"].append(r["zone"])
        if r["billed_lb"]:
            g["lb"] += r["billed_lb"]
        if r["amount"]:
            g["amount"] += r["amount"]
    # 百磅允许单独容差：导出表把每行重量取整，总重*费率/100 与账单分项合计
    # 通常只差几分钱，默认 0.10 内视为正常舍入
    base_tol = cfg["run"].get("amount_tolerance", 0.01)
    tol = cfg["run"].get("hundredweight_tolerance", base_tol)
    out = []
    for key in order:
        g = groups[key]
        g["lb"] = round(g["lb"], 2)
        g["amount"] = round(g["amount"], 2)
        from collections import Counter
        zone = Counter([z for z in g["zones"] if z]).most_common(1)
        zone = zone[0][0] if zone else ""
        zi = int(zone) if str(zone).isdigit() else None
        lb = g["lb"]
        band = "200-499" if lb < 500 else ("500-999" if lb < 1000 else "1000+")
        rate = None
        note = f"{g['n']}件 总重{lb:g}lb {band}档"
        if zi is not None:
            rate = cwt.get(zi, {}).get(band)
        if rate is None:
            g.update(zone=zone, status="lookup-failed", band=band, rate=None, quote=None,
                     diff=None, note=note + f" 无费率(分区{zone or '?'})")
            out.append(g)
            continue
        quote = round(lb * rate / 100.0, 2)
        diff = round(quote - g["amount"], 2)
        status = "diff" if abs(diff) > tol else "ok"
        g.update(zone=zone, band=band, rate=rate, quote=quote, diff=diff, status=status,
                 note=note + f" 费率{rate}/100lb")
        out.append(g)
    return out


# ---------- 5.5 加工副本（增加四列）与天图核验任务 ----------

def build_verify_tasks(records, cfg):
    """按费用类别生成去重后的天图核验任务（单号按 SOP 去掉 _ 后缀）。"""
    checks = cfg.get("tiantu_checks", {})
    tasks = []
    seen = set()
    for r in records:
        cat = cfg["fee_map"].get(r["fee_name"])
        if not cat or cat not in checks:
            continue
        key = (r["customer_no"] or r["sys_no"] or r["main_tracking"]).split("_")[0]
        if (cat, key) in seen:
            continue
        seen.add((cat, key))
        tasks.append({
            "row": r["row"],
            "no": key,
            "customer_no": r["customer_no"],
            "cat": cat,
            "fee": r["trans_zh"] or r["fee_name"],
            "mark": checks[cat]["mark"],
            "note": checks[cat].get("note", ""),
            "amount": r["amount"],
        })
    return tasks


def write_enriched_copy(src_path, records, out_path):
    """把原始账单复制为加工副本：原数据逐行保留，仅插入并填充四列。"""
    src = openpyxl.load_workbook(src_path, read_only=True, data_only=False)
    ws = src[src.sheetnames[0]]
    rows = ws.iter_rows(values_only=True)
    header = [c for c in next(rows)]
    amount_idx = None
    trans_idx = None
    for i, h in enumerate(header):
        t = norm(h)
        if t == "费用金额":
            amount_idx = i
        if t == "翻译":
            trans_idx = i
    rec_by_row = {r["row"]: r for r in records}
    out = openpyxl.Workbook()
    o = out.active
    o.title = ws.title or "Sheet0"
    if amount_idx is None:
        raise ValueError("原始账单找不到“费用金额”列，无法插入四列")
    if trans_idx is not None:
        # 输入已含四列（例如加工过的旧文件）：原样复制并只更新数值
        o.append(list(header))
        idxmap = {}
        for i, h in enumerate(header):
            t = norm(h)
            if t == "翻译":
                idxmap["trans"] = i
            elif t == "报价表":
                idxmap["quote"] = i
            elif t == "差异":
                idxmap["diff"] = i
        for rno, row in enumerate(rows, start=2):
            vals = list(row)
            rec = rec_by_row.get(rno)
            if rec is not None:
                if "trans" in idxmap:
                    vals[idxmap["trans"]] = rec["trans_zh"]
                if "quote" in idxmap:
                    vals[idxmap["quote"]] = rec.get("r_formula") or rec["quote"]
                if "diff" in idxmap:
                    vals[idxmap["diff"]] = rec.get("s_formula") or rec["diff"]
            o.append(vals)
    else:
        o.append(_insert_helper_cols(header, amount_idx, None))
        for rno, row in enumerate(rows, start=2):
            rec = rec_by_row.get(rno)
            o.append(_insert_helper_cols(list(row), amount_idx, rec))
    src.close()
    out.save(out_path)
    return out_path


def backup_original(path):
    bak = path.with_name(f"{path.stem}_原始备份.xlsx")
    if not bak.exists():
        shutil.copy2(path, bak)
        print(f"已生成原始备份: {bak}")
    return bak


def add_hundredweight_sheets(wb_path, records, cwt_groups):
    """在原表/加工副本新增 sheet：百磅原始拷贝 + 透视核对（重复运行先清旧表再重建）。"""
    wb = openpyxl.load_workbook(wb_path)
    for sn in ("百磅原始数据", "百磅透视核对"):
        if sn in wb.sheetnames:
            del wb[sn]
    raw = [r for r in records if r["fee_name"] == "Ground Hundredweight"]
    ws1 = wb.create_sheet("百磅原始数据")
    ws1.append(["客户单号", "分区H", "计费重量K", "费用金额P", "原Excel行"])
    for r in raw:
        ws1.append([r["customer_no"], r["zone"], r["billed_lb"], r["amount"], r["row"]])
    ws2 = wb.create_sheet("百磅透视核对")
    ws2.append(["客户单号", "分区", "件数", "总重量lb", "费用P合计", "档位",
                "费率/100lb", "报价R合计", "差异R-P", "状态"])
    for g in cwt_groups:
        ws2.append([g["customer_no"], g.get("zone", ""), g["n"], round(g["lb"], 2),
                    round(g["amount"], 2), g.get("band", ""), g.get("rate"),
                    g.get("quote"), g.get("diff"), g["status"]])
    wb.save(wb_path)
    return wb_path


def enrich_source_in_place(bill_path, records, cwt_groups, quote_path):
    """把四列 + 百磅 sheets 写回原表（先备份，重复运行不重复加列）；报价表复制到同目录供公式引用。"""
    backup_original(bill_path)
    qdst = bill_path.parent / Path(quote_path).name
    if not qdst.exists():
        shutil.copy2(quote_path, qdst)
    tmp = bill_path.with_name(f"{bill_path.stem}.tmp_enrich.xlsx")
    try:
        write_enriched_copy(bill_path, records, tmp)
        try:
            os.replace(tmp, bill_path)
            final = bill_path
        except OSError:
            # 原表可能正被 Excel 打开：退而生成同目录 _已核价 副本
            final = bill_path.with_name(f"{bill_path.stem}_已核价.xlsx")
            os.replace(tmp, final)
            print(f"原表 {bill_path.name} 可能正被占用，已生成: {final.name}（关闭原表后重跑可写回原表）")
        add_hundredweight_sheets(final, records, cwt_groups)
    finally:
        if tmp.exists():
            tmp.unlink()
    return final


def _insert_helper_cols(row, amount_idx, rec):
    """在 费用金额 前插 翻译、后插 第一次收取费用/报价表/差异。"""
    vals = []
    for i, v in enumerate(row):
        if i == amount_idx:
            vals.append("翻译" if rec is None else (rec["trans_zh"] or None))
        vals.append(v)
        if i == amount_idx:
            if rec is None:
                vals += ["第一次收取费用", "报价表", "差异"]
            else:
                vals += [None,
                         rec.get("r_formula") if rec.get("r_formula") else rec["quote"],
                         rec.get("s_formula") if rec.get("s_formula") else rec["diff"]]
    return vals


def write_verify_csv(out_path, tasks):
    with open(out_path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Excel行号", "客户单号(去后缀)", "费用类别", "费用名称", "期望标记", "金额USD", "说明"])
        for t in tasks:
            w.writerow([t["row"], t["no"], t["cat"], t["fee"], t["mark"], t["amount"], t["note"]])


# ---------- 6. 输出 ----------

def write_xlsx(out_path, records, summary_rows, problem_rows, pending_rows, cwt_rows, verify_rows=None):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "核价明细"
    cols = [
        "Excel行号", "发票日期", "交易日期", "系统单号", "客户单号", "客户单号(去后缀)",
        "主单号", "子单号", "分区", "计费重量lbs", "费用类型", "费用名称(EN)", "翻译",
        "费用金额P", "报价R", "差异R-P", "状态", "备注",
    ]
    ws.append(cols)
    for r in records:
        ws.append([
            r["row"], r["invoice_date"], r["trade_date"], r["sys_no"], r["customer_no"],
            r["customer_key"], r["main_tracking"], r["sub_tracking"], r["zone"],
            r["billed_lb"], r["fee_type"], r["fee_name"], r["trans_zh"], r["amount"],
            r["quote"], r["diff"], r["status"], r["note"],
        ])
    if cwt_rows:
        ws_c = wb.create_sheet("百磅核对(按票)")
        ws_c.append(["首行", "客户单号", "客户单号(去后缀)", "系统单号", "主单号", "分区",
                     "件数", "计费重量合计lb", "费用金额P合计", "档位", "费率/100lb",
                     "报价R合计", "差异R-P", "状态", "备注"])
        for g in cwt_rows:
            ws_c.append([
                g["row"], g["customer_no"], g["customer_key"], g["sys_no"],
                g["main_tracking"], g.get("zone", ""), g["n"],
                round(g["lb"], 2), round(g["amount"], 2), g.get("band", ""),
                g.get("rate"), g.get("quote"), g.get("diff"), g["status"], g.get("note", ""),
            ])
    for sheet_name, title, data in [
        ("问题清单", "问题行(|差异|>容差)", problem_rows),
        ("待处理清单", "未核价/待补规则行", pending_rows),
    ]:
        if not data:
            continue
        ws2 = wb.create_sheet(sheet_name)
        ws2.append(title and [title])
        ws2.append(cols)
        for r in data:
            ws2.append([
                r["row"], r["invoice_date"], r["trade_date"], r["sys_no"], r["customer_no"],
                r["customer_key"], r["main_tracking"], r["sub_tracking"], r["zone"],
                r["billed_lb"], r["fee_type"], r["fee_name"], r["trans_zh"], r["amount"],
                r["quote"], r["diff"], r["status"], r["note"],
            ])
    ws3 = wb.create_sheet("费用汇总")
    ws3.append(["费用名称(EN)", "翻译", "类别", "行数", "金额P合计", "报价R合计", "差异合计", "有差异行数", "说明"])
    for s in summary_rows:
        ws3.append(s)
    if verify_rows:
        ws4 = wb.create_sheet("待天图核验")
        ws4.append(["Excel行号", "客户单号(去后缀)", "费用类别", "费用名称", "期望标记", "金额USD", "说明"])
        for t in verify_rows:
            ws4.append([t["row"], t["no"], t["cat"], t["fee"], t["mark"], t["amount"], t["note"]])
    # 简单列宽
    widths = {1: 10, 2: 12, 3: 12, 4: 18, 5: 22, 6: 18, 7: 22, 8: 22, 9: 6,
              10: 12, 11: 8, 12: 34, 13: 18, 14: 12, 15: 12, 16: 12, 17: 14, 18: 40}
    for sheet in wb.worksheets:
        for idx, w in widths.items():
            if idx <= sheet.max_column:
                sheet.column_dimensions[get_column_letter(idx)].width = w
        for cell in sheet[1]:
            cell.font = Font(bold=True)
        sheet.freeze_panes = "A2"
    wb.save(out_path)


def write_report(out_path, lines):
    with open(out_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines) + "\n")


# ---------- 6.5 问题汇总写入用户选择的 xlsx ----------

def _fmt_money(v):
    """金额格式化：数值保留两位小数，空值返回空串。"""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, (int, float)):
        return f"{float(v):.2f}"
    return "" if v is None else str(v)


def diff_remark(amount, quote, diff):
    """核价差异行写入统计表备注列的文案。"""
    return (
        f"核价差异 R-P={_fmt_money(diff)}"
        f"（账单P={_fmt_money(amount)}，SOP报价R={_fmt_money(quote)}）"
    )


def cwt_diff_remark(g):
    """百磅票级差异写入统计表备注列的文案。"""
    head = f"{g.get('n', '')}件 总重{g.get('lb', '')}lb {g.get('band', '')}档"
    return head + " " + diff_remark(g.get("amount"), g.get("quote"), g.get("diff"))


def build_problem_items(problem_rows, cwt_groups):
    """把行级问题与百磅票级问题整理成可写入汇总表的记录（含备注列文案）。"""
    items = []
    for r in problem_rows:
        no = r["customer_no"] or r["sys_no"] or r["main_tracking"]
        items.append({
            "no": no,
            "fee": r["trans_zh"] or r["fee_name"],
            "amount": r["amount"],
            "remark": diff_remark(r["amount"], r["quote"], r["diff"]),
        })
    for g in cwt_groups:
        if g["status"] != "diff":
            continue
        no = g["customer_no"] or g["sys_no"] or g["main_tracking"]
        items.append({
            "no": no,
            "fee": "百磅计费运费(按票)",
            "amount": g["amount"],
            "remark": cwt_diff_remark(g),
        })
    return items


def _norm_h(v):
    return norm(v)


def locate_target_sheet(wb):
    """优先 名称含'永达' 的 sheet；否则找带 运单号/费用名称 表头的 sheet；再退回活动 sheet。"""
    for sn in wb.sheetnames:
        if "永达" in sn:
            return wb[sn]
    best, best_score = None, 0
    for sn in wb.sheetnames:
        ws = wb[sn]
        for row in ws.iter_rows(min_row=1, max_row=6, values_only=True):
            score = 0
            for c in row:
                t = _norm_h(c)
                if "运单号" in t or "客户单号" in t:
                    score += 2
                if "费用名称" in t:
                    score += 2
                if "金额" in t or "备注" in t:
                    score += 1
            if score > best_score:
                best, best_score = ws, score
    return best or wb.active


def map_target_headers(ws, max_scan=8):
    for ridx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1):
        colmap = {}
        for cidx, c in enumerate(row, start=1):
            t = _norm_h(c)
            if not t:
                continue
            if "运单号" in t:
                colmap.setdefault("no", cidx)
            elif "客户单号" in t:
                colmap.setdefault("no", cidx)
            elif "费用名称" in t:
                colmap.setdefault("fee", cidx)
            elif ("金额" in t or ("费用" in t and "名称" not in t)) and "amount" not in colmap:
                colmap["amount"] = cidx
            elif "备注" in t or "原因" in t or "说明" in t:
                colmap.setdefault("remark", cidx)
        if "no" in colmap and "fee" in colmap:
            return ridx, colmap
    return None, None


def append_to_workbook(path, items, dedupe=True):
    """把 items 追加到目标文件首个空白行；返回 (新增行数, 跳过重复数, 补备注行数, sheet名)。"""
    wb = openpyxl.load_workbook(path)
    ws = locate_target_sheet(wb)
    header_row, colmap = map_target_headers(ws)
    if header_row is None:
        raise ValueError(f"在 {ws.title} 找不到含 运单号/费用名称 的表头")
    existing = {}
    for ridx, row in enumerate(ws.iter_rows(min_row=header_row + 1, values_only=True),
                               start=header_row + 1):
        no = row[colmap["no"] - 1] if colmap["no"] - 1 < len(row) else None
        fee = row[colmap["fee"] - 1] if colmap["fee"] - 1 < len(row) else None
        if no not in (None, "") or fee not in (None, ""):
            key = (str(no or "").strip(), str(fee or "").strip())
            existing.setdefault(key, ridx)
    # 找表头下第一个真正空白行（避免 max_row 里存在格式残留空行）
    start_row = header_row + 1
    while start_row < header_row + 5000:
        vals = [ws.cell(row=start_row, column=c).value for c in colmap.values()]
        if all(v in (None, "") for v in vals):
            break
        start_row += 1
    added = skipped = filled = 0
    for it in items:
        key = (str(it["no"] or "").strip(), str(it["fee"] or "").strip())
        if dedupe and key in existing:
            # 原行备注为空时补写备注，避免重复问题永远缺原因
            if "remark" in colmap:
                old_row = existing[key]
                cur = ws.cell(row=old_row, column=colmap["remark"]).value
                if cur in (None, "") and it.get("remark") not in (None, ""):
                    ws.cell(row=old_row, column=colmap["remark"], value=str(it["remark"]))
                    filled += 1
            skipped += 1
            continue
        r = start_row + added
        ws.cell(row=r, column=colmap["no"], value=it["no"])
        ws.cell(row=r, column=colmap["fee"], value=it["fee"])
        if "amount" in colmap:
            amt = it.get("amount")
            if isinstance(amt, (int, float)):
                amt = round(float(amt), 2)
            ws.cell(row=r, column=colmap["amount"], value=amt)
        if "remark" in colmap:
            rm = it.get("remark")
            if rm not in (None, ""):
                ws.cell(row=r, column=colmap["remark"], value=str(rm))
        added += 1
        existing.setdefault(key, r)
    wb.save(path)
    return added, skipped, filled, ws.title


def choose_target_via_dialog():
    """弹窗选择要写入问题汇总的 xlsx；校验不通过会提示并重新弹出，取消返回 None。"""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    initial = ""
    try:
        cfg = json.loads((BASE_DIR / "config.json").read_text(encoding="utf-8"))
        p = Path(cfg["paths"]["input_dir"])
        initial = str(p) if p.exists() else ""
    except Exception:  # noqa: BLE001
        pass
    for _ in range(10):
        path = _ask_open_file(
            title="选择要写入问题汇总的 xlsx（如 打单问题统计表.xlsx）",
            filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
            initialdir=initial or "",
        )
        if not path:
            return None
        err = validate_stats_target(path)
        if err is None:
            return path
        messagebox.showerror("文件不符合要求", err)
    print("已连续多次选择不符合要求的文件，取消写入。")
    return None


def _ask_open_file(title, filetypes, initialdir=None):
    """独立临时 Tk 窗口弹出文件选择；关闭后窗口即销毁，避免隐藏窗口导致后续弹窗卡住。"""
    import tkinter as tk
    from tkinter import filedialog

    root = tk.Tk()
    root.withdraw()
    try:
        return filedialog.askopenfilename(
            title=title,
            filetypes=filetypes,
            initialdir=initialdir or "",
        )
    finally:
        root.destroy()


def validate_stats_target(path):
    """校验统计表文件；通过返回 None，否则返回可展示的错误说明。"""
    p = Path(path or "")
    if not p.exists():
        return "文件不存在，请重新选择。"
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        return f"不是 Excel 工作簿：{p.name}，请选择 .xlsx 文件。"
    try:
        with open(p, "r+b"):
            pass
    except PermissionError:
        return f"文件正被占用（可能已在 Excel 中打开），请先关闭再重新选择：{p.name}"
    except OSError as exc:  # noqa: BLE001
        return f"无法写入该文件，请检查权限后重新选择：{exc}"
    try:
        wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001
        return f"无法打开该 Excel 文件：{exc}"
    try:
        ws = locate_target_sheet(wb)
        header_row, colmap = map_target_headers(ws)
        if header_row is None:
            return "该文件没有包含“运单号/费用名称”表头的工作表，请选择《打单问题统计表》。"
    finally:
        wb.close()
    return None


# ---------- 7. 主流程 ----------

def main():
    ap = argparse.ArgumentParser(description="永达 UPS 账单核价检查")
    ap.add_argument("--config", default=str(BASE_DIR / "config.json"))
    ap.add_argument("--input", default=None, help="账单文件或目录，缺省自动选最新")
    ap.add_argument("--tolerance", type=float, default=None)
    ap.add_argument("--hundredweight-block", type=int, default=None, help="百磅费率第1/2套")
    ap.add_argument("--no-write-dialog", action="store_true", help="跑完不弹窗写入汇总表")
    ap.add_argument("--write-target", default=None, help="指定目标 xlsx 直接写入（跳过弹窗）")
    ap.add_argument("--no-dedupe", action="store_true", help="允许重复写入汇总表")
    args = ap.parse_args()
    cfg = load_config(args.config)
    cfg_file = Path(args.config)
    if not cfg_file.is_absolute():
        cfg_file = (Path.cwd() / cfg_file).resolve()
    cfg_dir = cfg_file.resolve().parent
    for key in ("input_dir", "output_dir", "log_dir", "translation_file", "quote_file"):
        raw = cfg["paths"].get(key)
        if raw:
            p = Path(raw)
            if not p.is_absolute():
                cfg["paths"][key] = str((cfg_dir / p).resolve())
    if args.tolerance is not None:
        cfg["run"]["amount_tolerance"] = args.tolerance
    if args.hundredweight_block is not None:
        cfg["run"]["hundredweight_block"] = args.hundredweight_block

    log_dir = Path(cfg["paths"]["log_dir"])
    log_dir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=log_dir / "yongda-bill-check.log",
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        encoding="utf-8",
    )

    bill_path = pick_bill_file(cfg, args.input)
    LOG.info("开始处理账单: %s", bill_path)
    print(f"账单文件: {bill_path}")

    records = read_bill(bill_path)
    if not records:
        raise SystemExit("账单没有数据行")
    translation, trans_conflicts = load_translation(cfg["paths"]["translation_file"])
    quote = load_ground_quote(cfg["paths"]["quote_file"])
    block_choice = cfg["run"].get("hundredweight_block", 1)
    cwt, cwt_blocks = load_hundredweight_quote(cfg["paths"]["quote_file"], block_choice)
    print(f"账单行数: {len(records)}; 报价表商业/住宅分区已加载; 百磅取第{block_choice}套费率")

    tol = cfg["run"].get("amount_tolerance", 0.01)
    missing_trans = sorted({r["fee_name"] for r in records if r["fee_name"] not in translation})
    for r in records:
        r["trans_zh"] = translation.get(r["fee_name"], "")
        r["quote"], r["status"], r["note"] = price_record(r, cfg, quote, cwt)
        cat = cfg["fee_map"].get(r["fee_name"])
        if r["status"] == "priced" and r["amount"] is not None:
            eff_tol = cfg["categories"].get(cat, {}).get("tolerance", tol)
            r["diff"] = round(r["quote"] - r["amount"], 2)
            r["status"] = "diff" if abs(r["diff"]) > eff_tol else "ok"
        else:
            r["diff"] = None
            if r["status"] == "priced" and r["amount"] is None:
                r["status"] = "amount-missing"
    # 运费修正规则 + R/S 公式（按 SOP）
    build_r_formulas(records, cfg, quote)
    for r in records:
        if r["status"] == "priced" and r["amount"] is not None and r.get("diff") is None:
            cat = cfg["fee_map"].get(r["fee_name"])
            eff_tol = cfg["categories"].get(cat, {}).get("tolerance", tol)
            r["diff"] = round(r["quote"] - r["amount"], 2)
            r["status"] = "diff" if abs(r["diff"]) > eff_tol else "ok"
    cwt_groups = summarize_hundredweight(records, cfg, cwt)
    verify_tasks = build_verify_tasks(records, cfg)

    # 汇总
    by_en = {}
    for r in records:
        cat = cfg["fee_map"].get(r["fee_name"])
        if cat == "hundredweight":
            continue  # 百磅在下方按票单独汇总
        key = (r["fee_name"], r["trans_zh"], cfg["fee_map"].get(r["fee_name"]))
        s = by_en.setdefault(key, {
            "n": 0, "sp": 0.0, "sr": 0.0, "sd": 0.0, "ndiff": 0, "note": ""
        })
        s["n"] += 1
        if r["amount"] is not None:
            s["sp"] += r["amount"]
        if r["quote"] is not None:
            s["sr"] += r["quote"]
        if r["diff"] is not None:
            s["sd"] += r["diff"]
            if r["status"] == "diff":
                s["ndiff"] += 1
        if cat:
            s["note"] = cfg["categories"][cat].get("note", "")
    summary_rows = []
    for (en, zh, cat), s in sorted(by_en.items()):
        summary_rows.append([en, zh, cat, s["n"], round(s["sp"], 2),
                             round(s["sr"], 2), round(s["sd"], 2), s["ndiff"], s["note"]])
    cwt_raw_n = sum(g["n"] for g in cwt_groups)
    cwt_sum_p = sum(g["amount"] for g in cwt_groups)
    cwt_sum_r = sum(g["quote"] for g in cwt_groups if g["quote"] is not None)
    cwt_sum_d = sum(g["diff"] for g in cwt_groups if g["diff"] is not None)
    cwt_n_diff = sum(1 for g in cwt_groups if g["status"] == "diff")
    summary_rows.append([
        "Ground Hundredweight", "百磅计费运费", "hundredweight", cwt_raw_n,
        round(cwt_sum_p, 2), round(cwt_sum_r, 2), round(cwt_sum_d, 2), cwt_n_diff,
        "按票透视汇总（差异按票计）",
    ])

    problem_rows = [r for r in records if r["status"] == "diff"]
    pending_rows = [r for r in records if r["status"] not in ("ok", "diff", "info", "grouped-cwt")]
    info_rows = [r for r in records if r["status"] == "info"]

    ts = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    stem = bill_path.stem
    out_dir = Path(cfg["paths"]["output_dir"])
    out_dir.mkdir(parents=True, exist_ok=True)
    xlsx_path = out_dir / f"{stem}_核价结果_{ts}.xlsx"
    txt_path = out_dir / f"{stem}_报告_{ts}.txt"
    write_xlsx(xlsx_path, records, summary_rows, problem_rows, pending_rows, cwt_groups, verify_tasks)
    enriched_path = out_dir / f"{stem}_加工副本_{ts}.xlsx"
    verify_csv_path = out_dir / f"{stem}_天图核验清单_{ts}.csv"
    try:
        write_enriched_copy(bill_path, records, enriched_path)
        print(f"加工副本(含四列): {enriched_path}")
        add_hundredweight_sheets(enriched_path, records, cwt_groups)
        qdst_out = out_dir / Path(cfg["paths"]["quote_file"]).name
        if not qdst_out.exists():
            shutil.copy2(cfg["paths"]["quote_file"], qdst_out)
    except Exception as exc:  # noqa: BLE001
        print(f"加工副本生成失败: {exc}")
    write_verify_csv(verify_csv_path, verify_tasks)
    print(f"待天图核验清单: {verify_csv_path}")
    if cfg["run"].get("enrich_source_in_place"):
        try:
            final = enrich_source_in_place(bill_path, records, cwt_groups, cfg["paths"]["quote_file"])
            print(f"四列及百磅sheet已写入: {final}")
        except Exception as exc:  # noqa: BLE001
            print(f"写回原表失败: {exc}")

    total_p = sum(r["amount"] for r in records if r["amount"] is not None)
    total_r = sum(r["quote"] for r in records if r["quote"] is not None)
    cust_all = {r["customer_no"] for r in records if r["customer_no"]}
    cust_key = {r["customer_key"] for r in records if r["customer_key"]}
    lines = [
        f"永达账单核价报告",
        f"生成时间: {_dt.datetime.now():%Y-%m-%d %H:%M:%S}",
        f"账单文件: {bill_path}",
        f"账单行数: {len(records)} | 客户单号数: {len(cust_all)} | 去后缀唯一数: {len(cust_key)}",
        f"费用金额P合计: {total_p:.2f} | 报价R合计(仅已核价): {total_r:.2f} | 差异容差: {tol}",
        "",
        "== 费用汇总 ==",
    ]
    lines += [
        " | ".join(f"{c}" for c in ["费用名称", "翻译", "类别", "行数", "P合计", "R合计", "差异合计", "差异行数"])
    ]
    for row in summary_rows:
        lines.append(" | ".join(str(c) for c in row[:8]))
    lines += ["", f"== 行级问题(差异>容差): {len(problem_rows)} =="]
    for r in problem_rows[:80]:
        lines.append(
            f"行{r['row']} {r['customer_no']} 分区{r['zone']} K={r['billed_lb']} "
            f"{r['fee_name']} P={r['amount']} R={r['quote']} 差异={r['diff']} {r['note']}"
        )
    if len(problem_rows) > 80:
        lines.append(f"... 其余 {len(problem_rows) - 80} 行见 Excel")
    lines += ["", f"== 百磅核对(按票透视): {len(cwt_groups)} 票（差异 {cwt_n_diff} 票）=="]
    for g in cwt_groups:
        lines.append(
            f"行{g['row']} {g['customer_no']} 分区{g.get('zone','')} {g['n']}件 "
            f"总重{g['lb']:g}lb P={g['amount']:.2f} {g.get('band','')} "
            f"R={g.get('quote')} 差异={g.get('diff')} {g['status']} {g.get('note','')}"
        )
    lines += ["", f"== 待天图核验(去重按票): {len(verify_tasks)} 条 =="]
    mark_stat = {}
    for t in verify_tasks:
        mark_stat.setdefault(t["mark"], 0)
        mark_stat[t["mark"]] += 1
    lines.append("按期望标记: " + ", ".join(f"{k}={v}" for k, v in sorted(mark_stat.items())))
    for t in verify_tasks[:60]:
        lines.append(f"行{t['row']} {t['no']} {t['fee']} -> 期望[{t['mark']}]")
    if len(verify_tasks) > 60:
        lines.append(f"... 其余 {len(verify_tasks) - 60} 条见 Excel/CSV")
    lines += ["", f"== 待处理/未核价: {len(pending_rows)} =="]
    pend_stat = {}
    for r in pending_rows:
        pend_stat.setdefault(r["status"], 0)
        pend_stat[r["status"]] += 1
    lines.append("按状态: " + ", ".join(f"{k}={v}" for k, v in sorted(pend_stat.items())))
    for r in pending_rows[:40]:
        lines.append(f"行{r['row']} {r['fee_name']} P={r['amount']} 状态={r['status']} {r['note']}")
    if len(pending_rows) > 40:
        lines.append(f"... 其余 {len(pending_rows) - 40} 行见 Excel")
    lines += ["", f"== 信息类(不核价): {len(info_rows)} =="]
    inf_stat = {}
    for r in info_rows:
        inf_stat.setdefault(r["fee_name"], 0)
        inf_stat[r["fee_name"]] += 1
    lines += [f"{k}: {v} 行" for k, v in sorted(inf_stat.items())]
    lines += ["", "== 翻译覆盖 =="]
    lines.append(f"账单费用名称未翻译: {missing_trans if missing_trans else '无'}")
    lines.append(f"翻译表内部冲突 {len(trans_conflicts)} 处（保留首次出现值）:")
    for c in trans_conflicts[:10]:
        lines.append(f"  {c[0]}: {c[1]} vs {c[2]} ({c[3]})")
    lines += ["", "== 校准提醒 =="]
    lines += [
        "1. 加工副本已插入 翻译/第一次收取费用/报价表/差异 四列：翻译、报价、差异已填，",
        "   第一次收取费用 留空，等待天图核验后回填。",
        "2. 天图核验按费用类别生成去重清单（偏远/超偏远/住宅私人/超长/超重/address），",
        "   由 check_tiantu.py 逐类核验后回写。",
        "3. 超偏远(DAS-Extended)按固定价2.85核算，2.82/2.87 视为容差(0.05)内一致。",
        "4. 百磅按票透视：总重<500用200-499档，<1000用500-999档，>=1000用1000+档，",
        "   报价=总重*费率/100，仅取2-8分区费率；费率取第2套(Rate Per Hundredweight Unit)。",
        "   导出表把每行重量取整，|差异|<=0.10 视为正常舍入（config: hundredweight_tolerance）。",
        "5. 燃油附加费、UPS折扣及运费更正关联行不核算（信息类不计入问题）。",
        f"输出: {xlsx_path}",
        f"报告: {txt_path}",
        f"加工副本: {enriched_path}",
        f"天图核验清单: {verify_csv_path}",
    ]
    report_text = "\n".join(lines)
    print(report_text)
    write_report(txt_path, lines)
    LOG.info("完成: %s (问题行 %s, 待处理 %s)", xlsx_path, len(problem_rows), len(pending_rows))

    if not args.no_write_dialog:
        items = build_problem_items(problem_rows, cwt_groups)
        target = args.write_target
        if not target:
            print("请在弹出窗口中选择要写入问题汇总的 xlsx（建议《打单问题统计表》）……")
            target = choose_target_via_dialog()
        if target:
            try:
                added, skipped, filled, sheet = append_to_workbook(
                    target, items, dedupe=not args.no_dedupe
                )
                msg = (
                    f"已写入 {target} [工作表:{sheet}]：新增 {added} 行，"
                    f"跳过重复 {skipped} 行，补写备注 {filled} 行"
                )
                print(msg)
                LOG.info("汇总写入 %s (%s)", msg, target)
            except Exception as exc:  # noqa: BLE001
                print(f"写入 {target} 失败: {exc}")
                LOG.exception("汇总写入失败: %s", target)
        else:
            print("未选择目标文件，跳过汇总写入；问题清单见报告与核价结果。")


if __name__ == "__main__":
    main()
