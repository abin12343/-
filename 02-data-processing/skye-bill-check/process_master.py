# -*- coding: utf-8 -*-
"""加工 Master sheet：按 SOP 规则执行核价，直接修改原表。

SOP 步骤映射：
  1) 在 Master 表头下方新增“翻译行”：填写 Service Code 的中文匹配行
  2) 在 J 列(Transportation Charge)右侧插入 2 列：K=报价表, L=差异
  3) 运费核对：筛选 Number of Pieces==1 的行，按 (Service Code, Rated Weight, Zone Code)
     查报价表，K 列填值，L 列填 Excel 公式 =K{r}-J{r}
  4) 附加费核对：按英文表头名定位费用列（SOP 中文名列于 config.master.fee_checks），
     按 (费用项, Zone, 商业/住宅) 查报价表，差异 > 容差时写入"报价差异"sheet
  5) 天图核验清单：**所有被收费的**附加费行（不只差异行）都要去天图验标记，
     由 write_tiantu_checklist 按费用名归类后生成

注意：SOP 文档里的列字母（P/Q/R/T/V/X/Y/AA/AB/AF）是人工插入 K/L 两列**之后**看到的列标，
比账单原始列整整 +2。代码若拿它直接索引原表会整体错位，故一律改用英文表头名定位：
  N=Additional Handling  O=AHS - Weight  P=AHS - Dimensions  R=Oversize Charge
  T=Address Correction   V=DAS Comm      W=DAS Resi
  Y=DAS Extended Comm    Z=DAS Extended Resi   AD=Residential
"""

from __future__ import annotations

import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import openpyxl

from quote_engine import QuoteBook, _to_float, _to_int, _col_letter_to_index, _index_to_col_letter


# ============== 数据结构 ==============

@dataclass
class PriceRow:
    row: int
    sheet: str
    reference_no: str
    service_code: str
    fee_name: str
    amount: float | None
    quote: float | None
    diff: float | None
    note: str = ""
    zone: int | str | None = None  # SOP 035：报价差异表要带分区


# ============== 翻译表 ==============

def load_translation(path) -> dict[str, str]:
    if not path or not Path(path).is_file():
        return {}
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    result: dict[str, str] = {}
    for name in wb.sheetnames[:2]:
        ws = wb[name]
        for r in ws.iter_rows(min_row=2, values_only=True):
            if not r or r[0] is None:
                continue
            k = str(r[0]).strip()
            v = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
            if k and v and k not in result:
                result[k] = v
    wb.close()
    return result


# ============== 辅助 ==============

def _col_letter_by_header(ws, header_row: int, name: str) -> int | None:
    for cell in ws[header_row]:
        if cell.value is not None and str(cell.value).strip() == name:
            return cell.column
    return None


def _ensure_cols(ws, header_row: int, cfg: dict) -> dict[str, str]:
    """幂等插入辅助列。返回 {逻辑名: 列字母}，其中 amount 是金额列（供差异公式引用）。

    关键：金额列必须**按表头名**定位。
    """
    master = cfg["master"]
    cols: dict[str, str] = {}

    # 报价表/差异：紧跟 Transportation Charge 右侧
    amt_idx = _col_letter_by_header(ws, header_row, "Transportation Charge")
    if amt_idx is None:
        amt_idx = _col_letter_to_index(master["columns"]["amount"])

    q_idx = _col_letter_by_header(ws, header_row, "报价表")
    d_idx = _col_letter_by_header(ws, header_row, "差异")
    if q_idx is None:
        ws.insert_cols(amt_idx + 1, amount=2)
        ws.cell(row=header_row, column=amt_idx + 1).value = "报价表"
        ws.cell(row=header_row, column=amt_idx + 2).value = "差异"
        q_idx, d_idx = amt_idx + 1, amt_idx + 2

    cols["quote"] = _index_to_col_letter(q_idx)
    cols["diff"] = _index_to_col_letter(d_idx)
    cols["amount"] = _index_to_col_letter(amt_idx)
    return cols


def _header_translations(
    ws, header_row: int, translation: dict, extra: dict | None = None,
) -> dict[int, str]:
    """第 1 行英文表头 -> 中文。

    优先用 config 内置的 extra（翻译表 Sheet1 里没有的 6 个基础表头），
    其余查翻译表；查不到的不写。
    """
    out: dict[int, str] = {}
    for cell in ws[header_row]:
        if cell.value is None:
            continue
        h = str(cell.value).strip()
        zh = (extra or {}).get(h) or translation.get(h)
        if zh is None and h in ("报价表", "差异"):
            zh = h          # 辅助列本身已是中文，人工模板里也写在第 2 行
        if zh:
            out[cell.column] = zh
    return out


def _ensure_translation_row(
    ws,
    header_row: int,
    translation: dict,
    max_row: int,
    extra: dict | None = None,
) -> int:
    """确保表头下方有一行中文翻译行（同人工模板的第 2 行）。

    幂等判据：该行已与翻译表对上的单元格 ≥ 2 个，就认为翻译行已存在。
    """
    trans_row = header_row + 1
    pairs = _header_translations(ws, header_row, translation, extra)
    if not pairs:
        return trans_row
    if trans_row <= max_row:
        hits = sum(
            1 for col, zh in pairs.items()
            if str(ws.cell(row=trans_row, column=col).value or "").strip() == zh
        )
        if hits < 2:
            ws.insert_rows(trans_row, amount=1)
    for col, zh in pairs.items():
        ws.cell(row=trans_row, column=col).value = zh
    return trans_row


def _hundredweight_freight(quote: QuoteBook, match, weight, zone, qcfg: dict):
    """百磅服务的运费报价，返回 (报价, 报价取自哪个 sheet, 哪个阶梯 block)。

    百磅档（200lb+/500lb+）只在重量 >= 200lb 时适用；低于门槛的票不能拿百磅单价去乘，
    否则会被 `Min 81` 保底价兜成 81（用户 2026-09-14 报的 5 行就是这样：
    26lb 的票报 81）。这类票改查**对应渠道的基础运费阶梯**——人工模板里
    31lb/Zone3 = 8.83 正是 FedEx Multiweigh 表内 FedEx Ground 阶梯的价格。

    只有配置了 `below_hwt="base"` 的服务才这么兜底（FedEx Multiweigh 表里正好并排
    放了一段 FedEx Ground 阶梯 = block 1）；`UPS GROUND HWT`、`FedEx-MWT-01-*` 没配
    → 低于门槛就不报价，返回 (None, None, None)。**注意**：`FedEx-MWT-01-*` 表里其实
    也有同样的 Ground 阶梯（block 1，与账单个别行能对上），但 <200lb 的 MWT 行都是
    合并计费的多件单行，逐行比 Ground 价会全是假差异，故不配。多件单本来也会被
    "件数=1 才核运费"这条挡掉。

    百磅档本身是"档位单价 × 总重 × 折扣"，不是阶梯查表，写不出 VLOOKUP，
    故第 2/3 项返回 None。
    """
    qcfg = qcfg or {}
    w = _to_float(weight) or 0.0
    if w >= qcfg.get("hundredweight_min_weight", 200):
        info = quote.hwt_price(match.sheet, match.block, weight, zone)
        return (info["final"] if info else None), None, None
    if match.below_hwt == "base":
        return (quote.base_price(match.sheet, match.base_block, weight, zone,
                                 round_up=qcfg.get("round_up_lookup", True)),
                match.sheet, match.base_block)
    return None, None, None


def _price_master_freight(
    ws, row: int, service_code: str, weight, zone, amount,
    quote: QuoteBook, cols: dict, cfg_master: dict, qcfg: dict,
    weight_letter: str | None = None,
) -> tuple[float | None, str]:
    match = quote.service_resolver(service_code)
    if not match:
        return None, ""
    price_sheet, price_block = None, None
    if match.kind == "weight_zone":
        price_sheet, price_block = match.sheet, match.block
        qp = quote.base_price(price_sheet, price_block, weight, zone,
                              round_up=qcfg.get("round_up_lookup", True))
    elif match.kind == "hundredweight":
        qp, price_sheet, price_block = _hundredweight_freight(quote, match, weight, zone, qcfg)
    else:
        return None, ""
    if qp is None:
        return None, ""
    q_letter = cols.get("quote")
    d_letter = cols.get("diff")
    # 金额列取 _ensure_cols 定位到的实际字母，不能用 config 里写死的 "J"
    j = cols.get("amount") or cfg_master["columns"]["amount"]
    if q_letter:
        w_letter = weight_letter or cfg_master["columns"]["weight"]
        # 报价表列保留公式（同人工模板）：跨工作簿 VLOOKUP 打到报价表的阶梯上；
        # 表结构不合模板形状时（如百磅档）退回数值，保证这列不会是空的。
        formula = (quote.vlookup_formula(price_sheet, price_block, row, w_letter, zone,
                                         weight=weight,
                                         round_up=qcfg.get("round_up_lookup", True))
                   if price_sheet is not None else None)
        ws[f"{q_letter}{row}"] = formula if formula else round(qp, 2)
    if d_letter:
        ws[f"{d_letter}{row}"] = f"={q_letter}{row}-{j}{row}"
    return round(qp, 2), f"={q_letter}{row}-{j}{row}"


def _price_master_surcharge(
    ws, row: int, fee_check: dict, zone, kind: str | None,
    quote: QuoteBook, qcfg: dict, amount, columns: dict, sheet_for_kind: str,
    tolerance: float, refno: str = "", sc: str = "",
    fallback_sheets: list[str] | None = None, multiplier: int = 1,
) -> PriceRow | None:
    item = fee_check.get("quote_item")
    if not item and fee_check.get("surcharge_item"):
        idx = 0 if kind == "商业" else 1
        items = fee_check["surcharge_item"]
        if idx < len(items):
            item = items[idx]
    if not item:
        return None
    unit = quote.surcharge_price(sheet_for_kind, 0, item, zone=zone, kind=kind,
                                 fallback_sheets=fallback_sheets)
    if unit is None:
        return None
    cap = quote.surcharge_cap(sheet_for_kind, 0, item, zone=zone, kind=kind,
                              fallback_sheets=fallback_sheets)
    # multiplier 由调用方决定：按箱计费的附加费传本票件数（Master 是整票汇总），
    # 按票计费的传 1；报价表的 `封顶MWT`/`封顶HWT` 列是**整票上限**，超了就按封顶收
    qp = round(unit * multiplier, 2)
    capped = cap is not None and qp > cap
    if capped:
        qp = round(cap, 2)
    # 统一差异口径：报价 - 账单金额。
    diff = round(qp - (amount or 0), 2)
    note = f"{fee_check['zh']} R={qp} 账单={amount} 差异={diff:+.2f}"
    if multiplier != 1:
        note += f"（单价{unit}×{multiplier}件）"
    if capped:
        note += f"（封顶{cap}）"
    return PriceRow(
        row=row, sheet="Master", reference_no=refno, service_code=sc,
        fee_name=fee_check["zh"], amount=amount, quote=qp, diff=diff, note=note,
        zone=zone,   # SOP 035：报价差异表要带分区（附加费也是按分区定价的）
    )


# ============== 主流程 ==============

def process_master(
    in_path, cfg: dict, quote: QuoteBook, translation: dict,
    log: logging.Logger,
) -> tuple[list[PriceRow], list[PriceRow]]:
    in_path = Path(in_path)
    wb = openpyxl.load_workbook(in_path)
    sheet_name = cfg["master"]["sheet"]
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise KeyError(f"账单缺少 {sheet_name} sheet")


    ws = wb[sheet_name]
    header_row = cfg["master"].get("header_row", 1)

    if _col_letter_by_header(ws, header_row, "Service Code") is None:
        raise ValueError("Master 表头缺少 'Service Code'")

    # 插入辅助列
    cols = _ensure_cols(ws, header_row, cfg)
    log.info("[Master] 辅助列：%s", cols)

    # 列字母（注意：insert_cols 后 openpyxl 的列对象已经右移，但 cell.coordinate 的列字母会更新）
    c = cfg["master"]["columns"]
    # 用表头名称直接定位（比字母更可靠，不受 insert 影响）
    def _col_by_name(name: str) -> int:
        return _col_letter_by_header(ws, header_row, name) or 1

    a_service = _col_by_name("Service Code")
    a_order   = _col_by_name("Order Code")
    a_refno   = _col_by_name("Reference No")
    a_pieces  = _col_by_name("Number of Pieces")
    a_weight  = _col_by_name("Rated Weight")
    a_zone    = _col_by_name("Zone Code")
    a_amount  = _col_by_name("Transportation Charge")
    # 在首行下维护一行中文翻译行（同人工模板第 2 行）
    trans_row = _ensure_translation_row(
        ws, header_row, translation, ws.max_row,
        cfg["master"].get("translation", {}).get("extra_headers"),
    )
    # SOP：超大尺寸费用要先看该单有没有收"私人住宅附加费(Residential)"来判断走商业还是住宅报价
    a_resi    = _col_letter_by_header(ws, header_row, "Residential")

    # 若插入了翻译行，数据从其下一行开始；历史文件兼容仍从首数据行开始
    data_start = header_row + 1
    if trans_row == header_row + 1:
        data_start = trans_row + 1


    freight_problems: list[PriceRow] = []
    surcharge_problems: list[PriceRow] = []
    # SOP：**所有被收费的**附加费行都要去天图核验标记（不只核价有差异的那些），
    # 所以这里单独收集一份，供 write_tiantu_checklist 生成核验清单。
    tiantu_rows: list[PriceRow] = []
    qcfg = cfg.get("quote", {})
    tolerance = cfg["run"]["amount_tolerance"]

    for r in range(data_start, ws.max_row + 1):
        if all(ws.cell(row=r, column=k).value in (None, "", "-") for k in range(1, ws.max_column + 1)):
            continue
        sc = ws.cell(row=r, column=a_service).value
        if not sc:
            continue
        sc = str(sc).strip()

        pieces = _to_int(ws.cell(row=r, column=a_pieces).value)
        weight = _to_float(ws.cell(row=r, column=a_weight).value)
        zone   = _to_int(ws.cell(row=r, column=a_zone).value)
        amount = _to_float(ws.cell(row=r, column=a_amount).value)
        refno  = ws.cell(row=r, column=a_refno).value or ""
        refno  = str(refno).split("-")[0] if refno else ""

        if pieces == 1 and weight and zone and amount is not None:
            match = quote.service_resolver(sc)
            if match:
                qp, _ = _price_master_freight(
                    ws, r, sc, weight, zone, amount, quote, cols, cfg["master"], qcfg,
                    weight_letter=_index_to_col_letter(a_weight),
                )
                if qp is None:
                    # 该报价的行这次不给价（如低于百磅门槛且无基础阶梯），
                    # 必须把上一次运行留下的值擦掉，否则重复处理同一张账单结果会不一样
                    for k in ("quote", "diff"):
                        if cols.get(k):
                            ws[f"{cols[k]}{r}"] = None
                if qp is not None and abs(amount - qp) > tolerance:
                    freight_problems.append(PriceRow(
                        row=r, sheet=sheet_name, reference_no=refno, service_code=sc,
                        fee_name="运费",
                        amount=amount, quote=qp, diff=round(qp - amount, 2),
                        note=f"运费 R={qp} 账单={amount} 差异={round(qp-amount,2):+.2f}",
                        zone=zone,
                    ))

        for fee in cfg["master"]["fee_checks"]:
            # 优先按真实英文表头名定位（不受 insert_cols 与列插入顺序影响）
            hdr_name = fee.get("header")
            col_idx = _col_letter_by_header(ws, header_row, hdr_name) if hdr_name else None
            if col_idx is None:
                col_idx = _col_letter_to_index(fee["col"])
            fee_amt = _to_float(ws.cell(row=r, column=col_idx).value)
            if not fee_amt or abs(fee_amt) < 0.01:
                continue
            # 收了这个费就得去天图验标记。收集放在"有无分区 / 服务能否匹配"的判断**之前**，
            # 否则这两类行会被天图核验漏掉。
            rec = PriceRow(
                row=r, sheet=sheet_name, reference_no=refno, service_code=sc,
                fee_name=fee["zh"], amount=fee_amt, quote=None, diff=None,
                note=f"{fee['zh']} 账单={fee_amt}", zone=zone,
            )
            tiantu_rows.append(rec)
            if not zone:
                continue
            match = quote.service_resolver(sc)
            if not match:
                continue
            sheet_for_kind = match.sheet
            # DAS Comm/DAS Resi 这类费用列本身就写明商业还是住宅，优先按费用列来
            # （服务只能给个默认值：同一张 MWT 单子也照样会收"住宅偏远附加费"）
            kind = fee.get("kind") or match.surcharge_kind
            # SOP 066-068：超大尺寸费用走商业还是住宅，看本单有无收 Residential(私人住宅附加费)
            if fee.get("infer_by_residential") and a_resi:
                resi_amt = _to_float(ws.cell(row=r, column=a_resi).value)
                kind = "住宅" if (resi_amt and abs(resi_amt) >= 0.01) else "商业"
            entry = next((e for e in cfg["quote"]["service_map"] if e["match"] in sc), None)
            fb = (entry or {}).get("fallback_sheets") or []
            # 报价表里绝大多数附加费按箱计费（Master 是整票汇总，故乘本票件数），
            # 但超重/超尺寸这类是**按票**收的：`per_piece: false` 时直接用单价（用户 2026-09-17 指出）
            prob = _price_master_surcharge(
                ws, r, fee, zone, kind, quote, qcfg, fee_amt,
                c, sheet_for_kind, tolerance, refno, sc,
                fallback_sheets=fb,
                multiplier=(pieces or 1) if fee.get("per_piece", True) else 1,
            )
            if prob is None:
                surcharge_problems.append(PriceRow(
                    row=r, sheet=sheet_name, reference_no=refno, service_code=sc,
                    fee_name=fee["zh"],
                    amount=fee_amt, quote=None, diff=None,
                    note=f"{fee['zh']} 账单={fee_amt} (报价表无对应项)",
                    zone=zone,
                ))
                rec.note = f"{fee['zh']} 账单={fee_amt} (报价表无对应项)"
                continue
            if abs((prob.diff or 0)) > tolerance:
                surcharge_problems.append(prob)
            # 有报价的行，把报价/差异补进清单行的说明里（人工对照天图时要用）
            rec.quote, rec.diff, rec.note = prob.quote, prob.diff, prob.note

    # 报价表列里是跨工作簿 VLOOKUP，本机没有 Excel 没法预算缓存值；
    # 打开时强制整本重算，Excel 才会去读报价表把数字算出来（否则单元格看起来是空的）
    wb.calculation.fullCalcOnLoad = True
    wb.save(in_path)
    wb.close()
    log.info("[Master] 已直接写入原表：%s", in_path)
    log.info("[Master] 运费差异=%d 行，附加费差异/无报价=%d 行，待天图核验的收费行=%d 行",
             len(freight_problems), len(surcharge_problems), len(tiantu_rows))
    return freight_problems, surcharge_problems, tiantu_rows
