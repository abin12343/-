# -*- coding: utf-8 -*-
"""openpyxl 辅助。

覆盖仓库里反复出现的几件事：
  - 表头不在第一行 / 有合并行时定位表头
  - 按中文别名模糊映射列（写死列号最容易在表格加列后错位）
  - 追加行并去重
  - 保存时文件被 Excel 占用要能降级，不能直接崩

只在调用时 import openpyxl，未安装给出明确提示。
"""

from __future__ import annotations

import re
from pathlib import Path

from .paths import backup_path


def _require_openpyxl():
    try:
        import openpyxl  # noqa: PLC0415
    except ImportError as exc:
        raise SystemExit("缺少 openpyxl：请先执行  pip install openpyxl") from exc
    return openpyxl


def norm(value) -> str:
    """规范化表头文本：去空白，便于匹配（"费 用 金额" == "费用金额"）。"""
    return re.sub(r"\s+", "", str(value)) if value is not None else ""


def find_header_row(worksheet, required=(), extra=(), max_scan=10):
    """在表格前 max_scan 行里定位表头行。

    :param required: 必须命中的规范化表头名列表，如 ("费用金额",)
    :param extra: 额外要提取的列，命中即记录
    :return: (行号, {逻辑名: 列号})；找不到返回 (None, None)
    """
    for ridx, row in enumerate(
        worksheet.iter_rows(min_row=1, max_row=max_scan, values_only=True), start=1
    ):
        cols = {}
        for cidx, cell in enumerate(row, start=1):
            t = norm(cell)
            if not t:
                continue
            for name in required:
                if t == name:
                    cols.setdefault(name, cidx)
            for name in extra:
                if t == name:
                    cols.setdefault(name, cidx)
        if all(name in cols for name in required):
            return ridx, cols
    return None, None


def map_columns(header_values, aliases: dict):
    """按别名映射列。

    aliases 形如 {"no": ["客户单号", "系统单号"], "fee": ["费用名称", "翻译"]}
    返回 {"no": 列号, ...}；未命中的键不出现在结果里。
    """
    cols = {}
    for cidx, cell in enumerate(header_values, start=1):
        t = norm(cell)
        if not t:
            continue
        for key, names in aliases.items():
            if key in cols:
                continue
            for n in names:
                if t == n or (len(n) > 2 and n in t):
                    cols[key] = cidx
                    break
    return cols


def read_rows(worksheet, header_row=1, min_row=None, skip_empty=True,
              aliases: dict | None = None):
    """把工作表读成 dict 列表。

    :param aliases: 传了就把命中的列**同时**映射成逻辑名，
                    返回的 dict 里既有逻辑名也有原始表头名，取用哪个都行。
                    这样表格改列名也不会让业务代码崩。
    """
    rows = list(worksheet.iter_rows(min_row=header_row, values_only=True))
    if not rows:
        return []
    headers = [norm(h) for h in rows[0]]
    idx_to_key = {}
    if aliases:
        for key, col in map_columns(headers, aliases).items():
            idx_to_key[col - 1] = key

    out = []
    for r in rows[(min_row or 1):]:
        if skip_empty and not any(v not in (None, "") for v in r):
            continue
        item = {}
        for i, value in enumerate(r):
            if i in idx_to_key:
                item[idx_to_key[i]] = value
            if i < len(headers) and headers[i]:
                item.setdefault(headers[i], value)
        out.append(item)
    return out


def first_blank_row(worksheet, start_row=1) -> int:
    """返回首个真正空白的行号（跳过只有空字符串的行）。"""
    ridx = start_row
    for row in worksheet.iter_rows(min_row=start_row, values_only=True):
        if any(v not in (None, "") for v in row):
            ridx += 1
            continue
        return ridx
    return ridx


def append_rows(worksheet, items, colmap, dedupe_keys=("no", "fee"),
                remark_key="remark", fill_blank_remark=True):
    """把 items 追加到工作表首个空白行。

    :param items: [{"no":..., "fee":..., "amount":..., "remark":...}, ...]
    :param colmap: {"no": 列号, "fee": 列号, "amount": 列号, "remark": 列号}
    :return: (新增行数, 跳过行数, 补写备注行数)
    """
    def key_of(values):
        key = tuple(
            str(values[colmap[k] - 1]).strip()
            for k in dedupe_keys
            if k in colmap and colmap[k] - 1 < len(values)
        )
        return key if key and all(key) else None

    # 先扫出已有行的 (单号, 名称) -> 行号，用于去重与补写备注
    index = {}
    for ridx, row in enumerate(worksheet.iter_rows(values_only=True), start=1):
        key = key_of(row)
        if key:
            index.setdefault(key, ridx)

    added = skipped = filled = 0
    row_ptr = first_blank_row(worksheet)
    for it in items:
        key = tuple(str(it.get(k, "") or "").strip() for k in dedupe_keys)
        complete = len(key) == len(dedupe_keys) and all(key)
        if complete and key in index:
            if fill_blank_remark and "remark" in colmap and it.get("remark"):
                cell = worksheet.cell(row=index[key], column=colmap["remark"])
                if cell.value in (None, ""):
                    cell.value = it["remark"]
                    filled += 1
            skipped += 1
            continue
        for key_name, col in colmap.items():
            value = it.get(key_name)
            if value is not None:
                worksheet.cell(row=row_ptr, column=col).value = value
        if complete:
            index[key] = row_ptr
        row_ptr += 1
        added += 1
    return added, skipped, filled


def safe_save(workbook, path, suffix="_已生成"):
    """保存工作簿；文件被 Excel 占用时自动写到 xx{suffix}.xlsx，避免整个流程崩掉。

    返回 (实际保存路径, 是否降级)
    """
    p = Path(path)
    try:
        workbook.save(p)
        return p, False
    except PermissionError:
        alt = p.with_name(f"{p.stem}{suffix}{p.suffix}")
        workbook.save(alt)
        return alt, True


def copy_with_backup(src, backup=True):
    """复制一份源文件（默认先备份原文件），返回 (工作簿, 备份路径或 None)。"""
    openpyxl = _require_openpyxl()
    src = Path(src)
    bak = None
    if backup:
        bak = backup_path(src)
        import shutil

        shutil.copy2(src, bak)
    return openpyxl.load_workbook(src), bak


def validate_bill_like(path, required=("费用金额",), any_of=("系统单号", "客户单号")):
    """校验 xlsx 是否像目标表格；通过返回 None，否则返回错误说明。

    用于弹窗选文件后的即时校验，避免跑到一半才发现选错。
    """
    p = Path(path or "")
    if not p.exists():
        return "文件不存在，请重新选择。"
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        return f"不是 Excel 工作簿：{p.name}"
    openpyxl = _require_openpyxl()
    wb = None
    try:
        wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), ())
        header = "".join(norm(c) for c in row)
    except Exception as exc:  # noqa: BLE001
        return f"无法打开该 Excel 文件：{exc}"
    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:  # noqa: BLE001
                pass
    miss = [k for k in required if k not in header]
    if miss:
        return f"缺少必需表头 {'/'.join(miss)}，请选择正确的文件。"
    if any_of and not any(k in header for k in any_of):
        return f"表头中未找到 {'/'.join(any_of)} 任一列，请选择正确的文件。"
    return None
