# -*- coding: utf-8 -*-
"""环洋账单加工与核价（账单2《成本明细》）。

严格照 SOP 走，加工程序与人工模板逐列对齐：

  1. 账单1（成本汇总）按《主单号》匹配出《客户单号》，在账单2 的 A 列右侧
     插入一列「运单号」，**粘贴为值**（SOP 原话）
     —— 用值不用 VLOOKUP：报价表/账单路径每次都重选，跨工作簿引用一断就是满屏 #REF!
  2. 在「费用名称」右侧插入列：翻译 / 费用明细 / 报价表 / 差异；计费重紧随其后
     差异 = 报价表 - 费用
  3. 「翻译」引用《打单费用名称中英文翻译》表，把英文费用名译成中文
  4. 按「翻译」分派报价规则，查报价表填「报价表」
  5. 基础运费按 计费重 × 分区 查阶梯；附加费按 产品名称 定位 sheet 再找费用项

插列一律**按表头名**定位、按列号降序插入，所以重复跑不会插歪、也不会插两次。
"""

from __future__ import annotations

from pathlib import Path
from copy import copy

import formulas
from quote_engine import QuoteBook, col_letter, strip_waybill_suffix, to_float, to_int

# 账单2 的列头（人工模板的最终列序，用于自检）
EXPECTED_HEADERS = [
    "账单日期", "运单号", "主单号", "子单号", "产品名称", "费用",
    "费用名称", "翻译", "费用明细", "报价表", "差异", "计费重",
    "账单实重", "邮编", "分区",
]


def _norm(value) -> str:
    import re

    return re.sub(r"\s+", "", str(value)) if value is not None else ""


def header_map(ws, header_row=1) -> dict:
    """表头文本 → 列号。"""
    return {
        _norm(c.value): c.column
        for c in ws[header_row]
        if c.value not in (None, "")
    }


# ── 翻译表 ───────────────────────────────────────────────────────────────


def build_translation_map(path, sheet="Sheet1", overrides=None, log=None) -> dict:
    """读《打单费用名称中英文翻译》，返回 英文费用名 → 中文。

    这张表**有重复键**：同一个英文名出现两次会给出不同中文，靠第 3 列 status
    区分——带「更新」的是后来修订的口径，优先取它。

    :param overrides: 手动定死若干个英文名（用于冲突键。例如 'Address Correction'
        在表里既有「地址更正」又有修订版的「拦截改派」，而 SOP 按「地址更正」分类，
        就必须在这里定死，否则 SOP 那条筛选永远匹配不上）
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        sheet = wb.sheetnames[0]
    ws = wb[sheet]
    out, updated = {}, set()
    for r in range(2, ws.max_row + 1):
        en = ws.cell(row=r, column=1).value
        zh = ws.cell(row=r, column=2).value
        st = _norm(ws.cell(row=r, column=3).value)
        if en in (None, "") or zh in (None, ""):
            continue
        key = str(en).strip()
        if key in updated:
            continue
        if key not in out or st == "更新":
            out[key] = str(zh).strip()
            if st == "更新":
                updated.add(key)
    wb.close()

    for en, zh in (overrides or {}).items():
        if not en or str(en).startswith("_"):
            continue
        out[str(en).strip()] = str(zh).strip()
        if log:
            log(f"[翻译] 手动指定 {en!r} → {zh!r}")
    if log:
        log(f"[翻译] 载入 {len(out)} 条")
    return out


def load_translation_table(path, sheet="Sheet1"):
    """翻译表的**原始两列** + 实际生效的 sheet 名。返回 (sheet名, [(英文, 中文), ...])。

    `build_translation_map()` 把数据归一过（strip、去重、更新行优先），那份结果给
    Python 用；公式走的是 **Excel 的匹配规则**，验算也得按原表的样子走一遍，
    所以这里再来读一次原始值。空行也留着——VLOOKUP 的首匹配会被它改变落点。
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    used = sheet if sheet in wb.sheetnames else wb.sheetnames[0]
    ws = wb[used]
    rows = [(ws.cell(row=r, column=1).value, ws.cell(row=r, column=2).value)
            for r in range(2, ws.max_row + 1)]
    wb.close()
    return used, rows


# ── 列处理 ───────────────────────────────────────────────────────────────


def _fill_column(ws, col, values, first_row=2):
    for i, v in enumerate(values):
        if v is not None:
            ws.cell(row=first_row + i, column=col).value = v


def _copy_column(ws, src_col: int, dst_col: int) -> None:
    """复制整列内容与样式。"""
    for r in range(1, ws.max_row + 1):
        src = ws.cell(row=r, column=src_col)
        dst = ws.cell(row=r, column=dst_col)
        if src.has_style:
            dst._style = copy(src._style)
        dst.number_format = src.number_format
        dst.alignment = copy(src.alignment)
        dst.protection = copy(src.protection)
        dst._hyperlink = copy(src.hyperlink) if src.hyperlink else None
        dst.comment = copy(src.comment) if src.comment else None
        dst.value = src.value


def _move_column(ws, src_col: int, dst_col: int) -> None:
    """将一列移动到目标位置，保留其余列顺序。"""
    if src_col == dst_col:
        return
    if src_col > dst_col:
        ws.insert_cols(dst_col, 1)
        _copy_column(ws, src_col + 1, dst_col)
        ws.delete_cols(src_col + 1, 1)
    else:
        ws.insert_cols(dst_col + 1, 1)
        _copy_column(ws, src_col, dst_col + 1)
        ws.delete_cols(src_col, 1)


def ensure_columns(ws, log=None) -> dict:
    """把 运单号 / 翻译 / 费用明细 / 报价表 / 差异 列补齐（缺哪补哪，已存在就不动）。

    按列号**降序**插入：先插靠右的列，再插靠左的「运单号」，
    这样前面插列不会把后面算好的位置顶偏。
    """
    hmap = header_map(ws)

    if "费用名称" not in hmap:
        raise ValueError("账单2 里找不到「费用名称」列，请确认选对文件")

    # 按目标顺序检查：翻译 / 费用明细 / 报价表 / 差异，缺哪补哪
    needed = []
    at = hmap["费用名称"] + 1
    if "翻译" not in hmap:
        needed.append(("翻译", at))
        at += 1
    else:
        at = hmap["翻译"] + 1

    if "费用明细" not in hmap:
        needed.append(("费用明细", at))
        at += 1
    else:
        at = hmap["费用明细"] + 1

    if "报价表" not in hmap:
        needed.append(("报价表", at))
        at += 1
    else:
        at = hmap["报价表"] + 1

    if "差异" not in hmap:
        needed.append(("差异", at))

    # 降序插入，避免列号错位
    for name, pos in reversed(needed):
        ws.insert_cols(pos, 1)
        ws.cell(row=1, column=pos).value = name
        if log:
            log(f"[列] 在第 {pos} 列插入：{name}")

    hmap = header_map(ws)
    if "运单号" not in hmap:
        ws.insert_cols(2, 1)
        ws.cell(row=1, column=2).value = "运单号"
        if log:
            log("[列] 在「账单日期」右侧插入 1 列：运单号")

    cols = header_map(ws)
    # 固定成品顺序：报价表 → 差异 → 计费重。
    # 兼容旧版已加工账单（报价表 → 计费重 → 差异），实际移动整列而非只改表头。
    if all(k in cols for k in ("报价表", "差异", "计费重")):
        if cols["报价表"] < cols["计费重"] < cols["差异"]:
            _move_column(ws, cols["差异"], cols["计费重"])
            cols = header_map(ws)
            if log:
                log(
                    f"[列] 已调整顺序：报价表={cols['报价表']}、"
                    f"差异={cols['差异']}、计费重={cols['计费重']}"
                )
    missing = [h for h in EXPECTED_HEADERS if h not in cols]
    if log:
        log("[列] 最终列序：" + " | ".join(
            f"{c}={ws.cell(row=1, column=cols[c]).value}" for c in sorted(cols, key=cols.get)
        ))
        if missing:
            log(f"[警告] 账单2 缺少这些列：{missing}")
    return cols


# ── 报价规则 ─────────────────────────────────────────────────────────────


def load_rules(cfg) -> dict:
    """config.fee_rules → {中文费用名: 规则}。"""
    rules = {}
    for rule in cfg.get("fee_rules") or []:
        for zh in rule.get("zh") or []:
            rules[_norm(zh)] = rule
    return rules


def price_row(rule, quote, product, weight, zone, round_up=True):
    """按规则查一条报价，返回 (价格, 说明, 出处)。

    出处（`QuoteRef`）是「这个价落在报价表哪一格」，给写公式用；查不到就是 None。
    """
    if rule.get("kind") == "base":
        lad = "residential" if rule.get("ladder") == "residential" else "commercial"
        ref = quote.base_ref(product, weight, zone, lad, round_up=round_up)
        if ref is None:
            return None, f"报价表无 {product} 的{lad}阶梯价（重={weight} 区={zone}）", None
        return ref.price, f"{lad}阶梯 重{weight} 区{zone}", ref
    ref, why = quote.surcharge_ref(
        product,
        rule.get("items") or [],
        zone=zone,
        weight=weight,
        section_prefer=rule.get("section_prefer"),
    )
    return (ref.price if ref else None), why, ref


# ── 写公式还是写值 ───────────────────────────────────────────────────────


class FormulaGate:
    """逐行、逐列决定「这一格写公式还是写值」。

    跨工作簿引用一断就是满屏 `#REF!`，所以每条公式在写下去之前都先按 **Excel 的规矩**
    验算一遍：算出来的数与引擎的数**相等**才写公式，不等就写值，并把原因记进报告。
    宁可少写一条公式，也不能写出一个静默的错价。

    代价是每条公式都要多算一遍——只算这一次，不影响天图那段。
    """

    def __init__(self, quote, cfg, ref_dir=None, translation_file=None, log=print):
        run_cfg = cfg.get("run") or {}
        self.on = bool(run_cfg.get("write_formulas", False))
        self.allow_abs = bool(run_cfg.get("write_formulas_allow_absolute", False))
        self.quote = quote
        self.ref_dir = ref_dir
        self.log = log
        self.stat = {"公式_翻译": 0, "公式_报价": 0, "公式_降级": 0}
        self.reasons = {}                 # 降级原因 -> 条数
        self._samples = {}                # 每种原因只打几行日志，别刷屏
        self._tokens = {}                 # sheet -> token 或 None（别反复 resolve）
        self.tr_token, self.tr_a, self.tr_b = None, [], []

        if not self.on:
            return
        if translation_file:
            self.tr_token, self.tr_a, self.tr_b = self._open_translation(
                translation_file, (cfg.get("translation") or {}).get("sheet", "Sheet1")
            )
        if self.tr_token is None:
            # 只说一声，不记进「降级」——降级要一格一格地数，才能和「已翻译」对上账
            log("[公式] 翻译表不在账单同目录：Excel 的外部引用只能写裸文件名或绝对路径，"
                "这一列会写值")

    def _open_translation(self, translation_file, sheet):
        """打开翻译表，返回 (token, 英文列, 中文列)。

        公式里的裸文件名是给 **Excel** 看的：它会按**账单所在目录**去找同名文件。
        所以验算也必须拿那一份——发布包自带一份翻译表，而账单文件夹里往往也有
        一份，两份不是同一个文件时验出来的结论就不作数。优先账单旁边那份。
        """
        use, name = Path(translation_file), Path(translation_file).name
        if self.ref_dir is not None and name:
            beside = Path(self.ref_dir) / name
            if beside.is_file() and beside.resolve() != use.resolve():
                use = beside
                self.log(f"[公式] 翻译表优先用账单旁边那份：{beside}")
        try:
            used_sheet, rows = load_translation_table(use, sheet)
        except Exception as exc:  # noqa: BLE001
            self._note(f"翻译表读不出来（{exc}）")
            return None, [], []
        token = formulas.ext_ref(use, self.ref_dir, used_sheet,
                                 allow_absolute=self.allow_abs)
        return token, [a for a, _ in rows], [b for _, b in rows]

    def _token(self, sheet):
        if sheet not in self._tokens:
            self._tokens[sheet] = formulas.ext_ref(
                self.quote.path, self.ref_dir, sheet, allow_absolute=self.allow_abs
            )
        return self._tokens[sheet]

    def _note(self, reason, sample=None):
        self.reasons[reason] = self.reasons.get(reason, 0) + 1
        self.stat["公式_降级"] += 1
        seen = self._samples.get(reason, 0)
        if seen < 3 and self.log:
            extra = f"（如 {sample}）" if sample else ""
            self.log(f"[公式] 退回写值：{reason}{extra}")
        self._samples[reason] = seen + 1

    def _read(self, sheet):
        return lambda r, c: self.quote.raw_cell(sheet, r, c)

    # -- 三列各自一道闸 ---------------------------------------------------

    def translation(self, row, fee_col, name, engine_zh):
        """翻译列：写 VLOOKUP，还是写引擎译好的中文。"""
        if not engine_zh:
            # 引擎没译出来：留空，连公式也不写——写出来 Excel 会显示 #N/A
            return None
        if not self.on or not name:
            return engine_zh
        if not self.tr_token:
            self._note("翻译表不在账单同目录，翻译列写值", name)
            return engine_zh
        got, sure = formulas.vlookup_first(self.tr_a, self.tr_b, name)
        if not sure or str(got if got is not None else "").strip() != engine_zh:
            # 首匹配跟引擎结论不一样：重名键、表里有首尾空格、或大小写差异。
            # 引擎的规则（含 config 里的 overrides）比 VLOOKUP 复杂，写出来就对不上。
            self._note("翻译表首匹配与引擎结论不一致，翻译列写值", name)
            return engine_zh
        self.stat["公式_翻译"] += 1
        return formulas.translation_formula(row, fee_col, self.tr_token)

    def quote_price(self, row, ref, weight_col, round_up, engine_price,
                    weight_raw, weight_engine):
        """报价列：写公式，还是写引擎算好的价。

        `weight_raw` 是计费重单元格里的**原文**，`weight_engine` 是引擎从它抠出来的数。
        公式将来读的是单元格，所以必须确认 Excel 从那一格读出来的数与引擎一致——
        `'26.000'` 这种数字文本两边都是 26，而带千分位的 `'1,234.5'` 引擎抠得出、
        Excel 只会给 #VALUE!，那就不能写公式。
        """
        want = round(engine_price, 2)
        if not self.on or ref is None:
            return want
        weight = formulas.xl_number(weight_raw)
        if weight is None or weight != weight_engine:
            self._note("计费重那一格 Excel 读不出同样的数（文本/公式/带千分位）",
                       f"第{row}行 {weight_raw!r}")
            return want
        token = self._token(ref.sheet)
        if not token:
            self._note("报价表不在账单同目录，报价列写值", f"第{row}行 {ref.sheet}")
            return want
        read = self._read(ref.sheet)
        if ref.ladder is not None:
            f, val = formulas.base_formula(
                token, ref.ladder, ref.zone, weight, read,
                round_up=round_up,
                weight_cell=f"${col_letter(weight_col)}{row}" if weight_col else None,
            )
        else:
            f, val = formulas.surcharge_formula(token, ref, read)
        if f is None:
            self._note("报价表这一档写不成公式（原表没升序/区间有空档/命中的是空格子）",
                       f"第{row}行 {ref.why}")
            return want
        if val != want:
            self._note("公式算出来的价与引擎对不上", f"第{row}行 {ref.why} 公式={val} 引擎={want}")
            return want
        self.stat["公式_报价"] += 1
        return f

    def diff(self, row, price_col, fee_col, price_2dp, fee, engine_diff):
        """差异列：写 `=ROUND(报价-费用,2)`，还是写引擎算好的差。

        报价单元格里已经是 2 位小数了，而引擎是拿**未舍入**的报价减费用——
        先舍再加偶尔会差一分，所以这里也必须验。
        """
        want = round(engine_diff, 2)
        if not self.on:
            return want
        if formulas.diff_value(price_2dp, fee) != want:
            self._note("差异写公式会与引擎差一分（报价先舍后减）", f"第{row}行")
            return want
        self.stat["公式_差异"] = self.stat.get("公式_差异", 0) + 1
        return formulas.diff_formula(row, price_col, fee_col)


# ── 主流程 ───────────────────────────────────────────────────────────────


def _backup_path(path: Path) -> Path:
    """备份路径：<名>_原始备份.xlsx，已存在就往后编号。"""
    cand = path.with_name(f"{path.stem}_原始备份{path.suffix}")
    i = 1
    while cand.exists():
        cand = path.with_name(f"{path.stem}_原始备份{i}{path.suffix}")
        i += 1
    return cand


def _safe_save(wb, path: Path, suffix="_已生成"):
    """保存；文件被 Excel 占用时另存，别让整个流程崩掉。返回 (实际路径, 是否降级)。"""
    try:
        wb.save(path)
        return path, False
    except PermissionError:
        alt = path.with_name(f"{path.stem}{suffix}{path.suffix}")
        wb.save(alt)
        return alt, True


def process(bill2_path, summary_index, quote, translation, cfg, log=print,
            backup=True, write_formulas=None, ref_dir=None, translation_file=None):
    """加工账单2 并回填报价。

    :param summary_index: load_summary_index() 的结果（主单号 → 客户单号…）
    :param quote: QuoteBook
    :param translation: {英文费用名: 中文}
    :param write_formulas: 覆盖 cfg.run.write_formulas（None = 用配置里的值）
    :param ref_dir: 写公式时「账单所在的目录」——Excel 按它解析裸文件名。
        `_dry_run` 会把账单复制到临时目录，那时**必须**传原始账单目录，否则基准全错。
    :param translation_file: 翻译表的路径（要写公式才需要）
    :return: (保存路径, 统计 dict, 行列表)
    """
    import shutil

    import openpyxl

    src = Path(bill2_path)
    bak = None
    if backup:
        # 已存在就不覆盖：否则反复跑会把"已加工"的表当成原始表存下来，备份就没意义了
        bak = _backup_path(src)
        shutil.copy2(src, bak)
        log(f"[备份] 原账单已备份到 {bak.name}")

    if write_formulas is not None:
        cfg = dict(cfg)
        cfg["run"] = dict(cfg.get("run") or {}, write_formulas=bool(write_formulas))
    gate = FormulaGate(quote, cfg, ref_dir=ref_dir if ref_dir is not None else src.parent,
                       translation_file=translation_file, log=log)

    wb = openpyxl.load_workbook(src)
    ws = wb[wb.sheetnames[0]]
    cols = ensure_columns(ws, log=log)
    max_row = ws.max_row

    c_主单 = cols.get("主单号")
    c_运单 = cols["运单号"]
    c_费用 = cols.get("费用")
    c_费用名 = cols["费用名称"]
    c_翻译 = cols["翻译"]
    c_报价 = cols["报价表"]
    c_差异 = cols["差异"]
    c_产品 = cols.get("产品名称")
    c_计费重 = cols.get("计费重")
    c_分区 = cols.get("分区")

    run_cfg = cfg.get("run") or {}
    round_up = bool(run_cfg.get("round_up_lookup", True))
    tol = float(run_cfg.get("amount_tolerance", 0.05))
    skip_zero = bool(run_cfg.get("skip_zero_amount", True))
    rules = load_rules(cfg)

    stat = {
        "总行数": max_row - 1,
        "已填运单号": 0, "运单号缺失": 0,
        "已翻译": 0, "无翻译": 0,
        "已报价": 0, "未匹配报价": 0, "未匹配明细": {},
        "零额跳过": 0, "差异超容差": 0, "差异样例": [],
        "公式_翻译": 0, "公式_报价": 0, "公式_差异": 0, "公式_降级": 0,
    }
    unknown_fee = {}

    # 先整列读出，避免边读边写在插入列后错位
    rows = []
    for r in range(2, max_row + 1):
        rows.append({
            "row": r,
            "主单号": ws.cell(row=r, column=c_主单).value if c_主单 else None,
            "费用": to_float(ws.cell(row=r, column=c_费用).value) if c_费用 else None,
            "费用名称": ws.cell(row=r, column=c_费用名).value,
            "产品名称": ws.cell(row=r, column=c_产品).value if c_产品 else None,
            "计费重": to_float(ws.cell(row=r, column=c_计费重).value) if c_计费重 else None,
            # 计费重的**原文**：写公式时公式读的是单元格，单元格是文本的话
            # Excel 的 CEILING 给 #VALUE!，而 to_float 还抠得出数——必须分开留
            "计费重原文": ws.cell(row=r, column=c_计费重).value if c_计费重 else None,
            "分区": to_int(ws.cell(row=r, column=c_分区).value) if c_分区 else None,
        })

    # 1) 运单号（值，不是公式）
    for it in rows:
        mno = _norm(it["主单号"])
        info = summary_index.get(mno) if mno else None
        cust = info.get("客户单号") if info else None
        it["运单号"] = strip_waybill_suffix(cust) if cust else ""
        if it["运单号"]:
            stat["已填运单号"] += 1
        elif mno:
            stat["运单号缺失"] += 1
        ws.cell(row=it["row"], column=c_运单).value = it["运单号"] or None

    # 2) 翻译 / 3) 报价表 / 4) 差异
    for it in rows:
        name = str(it["费用名称"] or "").strip()
        zh = translation.get(name, "") if name else ""
        it["翻译"] = zh
        if zh:
            stat["已翻译"] += 1
        elif name:
            stat["无翻译"] += 1
            unknown_fee[name] = unknown_fee.get(name, 0) + 1
        ws.cell(row=it["row"], column=c_翻译).value = gate.translation(
            it["row"], c_费用名, name, zh
        )

        fee = it["费用"]
        quote_price = None
        ref = None
        rule = rules.get(_norm(zh)) if zh else None
        if rule is None:
            if zh and _norm(zh) not in rules:
                pass  # 不是 SOP 要核的费用类别，留空（如燃油、折扣）
        elif fee in (None, 0) and skip_zero:
            stat["零额跳过"] += 1
        else:
            quote_price, why, ref = price_row(
                rule, quote, it["产品名称"], it["计费重"], it["分区"], round_up
            )
            if quote_price is None:
                stat["未匹配报价"] += 1
                key = f"{zh} | {why}"
                stat["未匹配明细"][key] = stat["未匹配明细"].get(key, 0) + 1
            else:
                stat["已报价"] += 1

        if quote_price is not None:
            ws.cell(row=it["row"], column=c_报价).value = gate.quote_price(
                it["row"], ref, c_计费重, round_up, quote_price,
                it["计费重原文"], it["计费重"],
            )
            if fee is not None:
                diff = round(quote_price - fee, 2)
                ws.cell(row=it["row"], column=c_差异).value = gate.diff(
                    it["row"], c_报价, c_费用, round(quote_price, 2), fee, quote_price - fee
                )
                if abs(diff) > tol:
                    stat["差异超容差"] += 1
                    if len(stat["差异样例"]) < 12:
                        stat["差异样例"].append(
                            (it["运单号"], zh, it["产品名称"], it["计费重"],
                             it["分区"], fee, round(quote_price, 2), diff)
                        )
        else:
            ws.cell(row=it["row"], column=c_报价).value = None
            ws.cell(row=it["row"], column=c_差异).value = None

    saved, degraded = _safe_save(wb, src)
    wb.close()
    if degraded:
        log(f"[提示] 原文件被占用，已另存为 {saved.name}")
    stat["未翻译费用名"] = sorted(unknown_fee.items(), key=lambda kv: -kv[1])[:20]
    stat.update(gate.stat)                       # 公式条的计数由闸门自己记
    stat["公式开关"] = gate.on
    stat["降级明细"] = sorted(gate.reasons.items(), key=lambda kv: -kv[1])
    stat["备份"] = str(bak) if bak else ""
    stat["保存"] = str(saved)
    return saved, stat, rows


def build_checklist(rows, cfg, translation, rules, log=print):
    """按 SOP 生成《天图核验清单》。

    「二者缺一不可」的类别（住宅偏远/超偏远）在 config 里配了多个标记，
    这里就拆成多条任务——引擎按标记分组批量查，同单号不同标记互不影响。
    返回 (任务列表, 单号→账单信息)。
    """
    tasks, seen = [], set()
    for it in rows:
        rule = rules.get(_norm(it.get("翻译"))) if it.get("翻译") else None
        if not rule:
            continue
        marks = rule.get("tiantu") or []
        no = strip_waybill_suffix(it.get("运单号"))
        if not no or not marks:
            continue
        for mark in marks:
            key = (no, mark)
            if key in seen:
                continue
            seen.add(key)
            tasks.append({
                "row": it.get("row", ""),
                "no": no,
                "cat": it.get("翻译") or "",
                "fee": rule.get("zh", [""])[0] if rule.get("zh") else "",
                "mark": mark,
                "amount": it.get("费用"),
                "note": rule.get("note", ""),
            })
    log(f"[清单] 天图核验任务 {len(tasks)} 条（去重后），涉及单号 {len({t['no'] for t in tasks})} 个")
    return tasks
