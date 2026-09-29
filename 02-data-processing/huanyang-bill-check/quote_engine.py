# -*- coding: utf-8 -*-
"""环洋《HYE美国海外事业部大货价格表》解析引擎。

环洋报价表每张渠道 sheet 的布局都不一样（21 张表，起码三种排版），
所以一律靠"文字线索"定位，绝不写死列号——列一挪也不会算错价。

两条线索：

  1) 阶梯表：某行的单元格文本是 lbs./Lbs./LB（重量表头），
     它右边紧接着一串 Zone 2 / Zone 3 / ... / Zone 8 / Zone 44 / Zone 45 / Zone 46。
     这一行就是表头行，下面是「重量 × 分区 → 单价」阶梯。
     块标题（UPS Ground Commercial / Home Delivery / Ground MWT …）写在表头行上方。
     同一张表可以有好几段阶梯：商业、住宅、百磅、MWT，各自独立。

  2) 附加费表：某单元格 X 的**右邻居**文本含「收费标准」，
     那么 X 就是区块标题（费用名称 / GROUND费用名称 / HOME DELIVERY费用名称 / 通用附加费），
     右邻居列就是价格列。区块内逐行读「费用项名 → 价」。
     若价格格里写的是分区（2区 / 3-4区 / Zone 7+），真价在再右一列，
     而且费用项名会留空、沿用上一行（合并单元格的常见写法）。

对外只暴露三个查询（其余是内部实现）：
    sheet_for_product(product)                                 产品名称 → sheet
    base_price(product, weight, zone)                          基础运费
    surcharge_price(product, items, zone, weight, section_prefer)  附加费
"""

from __future__ import annotations

import math
import re
from pathlib import Path

# ── 文本归一化 ───────────────────────────────────────────────────────────

_WS = re.compile(r"\s+")
# 产品名称/表名里的地理前缀，匹配报价表时一律忽略（表名写的是"美西 美中 美东UPS GROUND-BD"，
# 账单里写的是"美中UPS GROUND-BD"，去掉前缀才能对上）
_REGION_PREFIX = ("美西", "美中", "美东")
# 报价表里不是渠道的表，归一化后为空或没意义，直接排除
_META_SHEETS = ("产品介绍", "产品特点", "目录", "说明")

_ZONE_HEADER = re.compile(r"^\s*(?:zone|zones|分区)\s*([0-9]+)\s*(?:-\s*([0-9]+))?\s*(\+|plus)?\s*$", re.I)
_ZONE_KEY = re.compile(r"([0-9]+)\s*(?:-\s*([0-9]+))?\s*(?:\+|plus)?", re.I)
_WEIGHT_HEADER = re.compile(r"^\s*lbs?\.?\s*$", re.I)
_NUM = re.compile(r"-?\d+(?:\.\d+)?")


def _norm(value) -> str:
    """去空白、转大写，用于表头/名称比较。"""
    return _WS.sub("", str(value)) if value is not None else ""


def _product_key(name: str) -> str:
    """产品名称/表名归一化：去地理前缀、去非字母数字、大写。"""
    s = str(name or "").upper()
    for pre in _REGION_PREFIX:
        s = s.replace(pre, "")
    # "Fedex Fedex Ground-F" 这种把厂商名写重了的，压成一次
    s = re.sub(r"\b(FEDEX|UPS)\s+(?=\1\b)", "", s)
    return re.sub(r"[^0-9A-Z]", "", s)


def to_float(value):
    """尽量转 float；失败返回 None。'6.26' / 6.26 / ' 1,234.5 ' 都能吃。"""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    m = _NUM.search(str(value).replace(",", ""))
    return float(m.group()) if m else None


def to_int(value):
    f = to_float(value)
    return None if f is None else int(f)


def parse_zones(text):
    """把分区写法解析成 int 集合。

    覆盖：'Zone 2' / '2区' / 'Zone 3-4' / '3-4区' / 'Zone 7+' / '7+区' / 'Zones 44-46'。
    '+' 表示"该号段及以上"，同时带上 44/45/46（阿拉斯加/夏威夷/波多黎各）。
    """
    s = str(text or "")
    m = _ZONE_KEY.search(s)
    if not m:
        return set()
    lo = int(m.group(1))
    hi = int(m.group(2)) if m.group(2) else None
    if hi is not None:
        return set(range(lo, hi + 1))
    if "+" in s or "＋" in s or re.search(r"plus", s, re.I):
        return set(list(range(lo, 9)) + [44, 45, 46])
    return {lo}


def _looks_like_zone_label(value) -> bool:
    """价格格里写的是不是分区。

    必须排除纯数字字符串——'2区' 能被 to_float 抠出 2.0，
    只靠"能不能转成数字"判断会把分区号当成价格读（这里踩过）。
    """
    if not isinstance(value, str):
        return False
    s = value.strip()
    if not s or re.fullmatch(r"[-+]?[\d.,\s]+", s):
        return False
    return bool(_ZONE_KEY.search(s))


def _section_matches(title, pref) -> bool:
    """区块名匹配（宽松）。

    报价表是人工维护的，区块标题拼写/截断不齐：FedEx Ground-C 写 'HOME DELIVERY',
    Ground-Q 写 'HOME DELIVER'（少个 Y）。所以去空格后双向包含 + 前 6 个字符前缀
    都算命中——要求精确相等的话住宅类的价会全部落到商业区块。
    """
    t = _WS.sub("", str(title or "")).upper()
    p = _WS.sub("", str(pref or "")).upper()
    if not p:
        return True
    if not t:
        # 标题为空（'费用名称' 那种通用区块）不算命中：否则 '' in 'GROUND' 恒真，
        # 住宅的价会被通用区块先截胡
        return False
    if p in t or t in p:
        return True
    n = min(len(p), len(t), 6)
    return n >= 4 and t[:n] == p[:n]


def col_letter(idx: int) -> str:
    """1 → A。仅用于日志/报告可读性。"""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# ── 数据结构 ─────────────────────────────────────────────────────────────


class Ladder:
    """一段「重量 × 分区 → 单价」阶梯。"""

    def __init__(self, title, header_row, weight_col, zone_cols, rows, min_by_zone=None,
                 rung_rows=None, match_safe=False, min_cells=None):
        self.title = title
        self.header_row = header_row
        self.weight_col = weight_col
        self.zone_cols = zone_cols          # {zone:int -> 列号}
        self.rows = rows                    # 升序 [(重量, {zone: 单价})]
        self.min_by_zone = min_by_zone or {}  # 表里 'Min' 行的分区保底价
        self.is_hwt = False                 # 百磅/MWT 段（>=200lb 才用）
        # ── 以下只为「把这次查价写成 Excel 公式」提供出处（formulas.py 用），
        #    解析与查价的行为一个字不变 ──
        self.rung_rows = rung_rows or []    # 源表顺序 [(重量, 行号)]，取自 rows.sort() **之前**
        self.match_safe = match_safe        # 见 _parse_ladder：原表升序+区间连续+纯数字
        self.min_cells = min_cells or {}    # {zone: (行, 列)} 'Min' 保底价所在格

    @property
    def data_row_range(self):
        """阶梯数据行区间 (首行, 末行)，给 MATCH 的区间用。"""
        if not self.rung_rows:
            return (None, None)
        return (self.rung_rows[0][1], self.rung_rows[-1][1])

    @property
    def is_residential(self) -> bool:
        """只看标题里的英文标识。

        不能拿中文"住宅"当判据：FedEx Ground-C 的阶梯标题是
        「Fedex Ground Commercial（住宅地址需要额外增加住宅附加费）」，
        那是备注不是属性，按中文判会把这唯一的通用阶梯误标成住宅梯。
        """
        t = self.title.upper()
        return "RESIDENTIAL" in t or "HOME DELIVERY" in t

    def weight_range(self):
        if not self.rows:
            return (None, None)
        return (self.rows[0][0], self.rows[-1][0])

    def _pick_row(self, w):
        """命中档位在 rows（已升序）里的下标；轻于阶梯起点时是 0。"""
        idx = None
        for i, (row_w, _) in enumerate(self.rows):
            if row_w <= w:
                idx = i
            else:
                break
        return 0 if idx is None else idx

    def lookup(self, weight, zone, round_up=True):
        """查价。重量按 ceil 取档（行业口径），分区必须精确命中该段阶梯。"""
        col = self.zone_cols.get(int(zone)) if zone is not None else None
        if col is None or not self.rows:
            return None
        w = float(weight)
        w = math.ceil(w) if round_up else int(w)
        price = self.rows[self._pick_row(w)][1].get(int(zone))
        if price is None:
            return None
        floor = self.min_by_zone.get(int(zone))
        if floor is not None:
            price = max(price, floor)
        return round(price, 4)

    def cell_for_weight(self, weight, zone, round_up=True):
        """这次查价命中的**价单元格** (行, 列)；给不出就 None。

        只有 match_safe 的阶梯给得出：rung_rows 是**源表顺序**的，排序之后还能
        按下标对得上，靠的就是"原表本来就是升序、压根没被排动过"。
        """
        col = self.zone_cols.get(int(zone)) if zone is not None else None
        if col is None or not self.rows or not self.match_safe:
            return None
        w = float(weight)
        w = math.ceil(w) if round_up else int(w)
        idx = self._pick_row(w)
        if idx >= len(self.rung_rows):
            return None
        return (self.rung_rows[idx][1], col)

    def __repr__(self):
        lo, hi = self.weight_range()
        return f"<Ladder {self.title!r} {lo}-{hi} zones={sorted(self.zone_cols)}>"


class SurchargeEntry:
    """一个附加费项：要么定额，要么按分区定价。"""

    def __init__(self, key, flat=None, zones=None, flat_cell=None, zone_cells=None):
        self.key = key
        self.flat = flat                    # float 或 None
        self.zones = zones or {}            # {zone:int -> 价} 或 {}
        # 出处的单元格坐标，给「写成公式」用（见 formulas.py）
        self.flat_cell = flat_cell          # (行, 列)
        self.zone_cells = zone_cells or {}  # {zone:int -> (行, 列)}

    def cell_for(self, zone=None):
        """这条报价落在哪个单元格。定额项跟分区无关。"""
        if self.flat is not None:
            return self.flat_cell
        return self.zone_cells.get(int(zone)) if zone is not None else None

    def price(self, zone=None):
        if self.flat is not None:
            return round(self.flat, 4)
        if zone is not None:
            p = self.zones.get(int(zone))
            return round(p, 4) if p is not None else None
        return None

    def __repr__(self):
        return f"<Entry {self.key!r} flat={self.flat} zones={sorted(self.zones)}>"


class SurchargeSection:
    def __init__(self, title, entries, price_col=None, key_col=None):
        self.title = title
        self.entries = entries              # [SurchargeEntry]
        self.price_col = price_col          # 价格列（出处用）
        self.key_col = key_col              # 费用项名列（出处用）

    def find(self, keys):
        """按候选名找项：先精确匹配，再子串匹配（表里写法常有出入）。"""
        norm = [_norm(k).upper() for k in keys]
        for e in self.entries:
            if _norm(e.key).upper() in norm:
                return e
        for e in self.entries:
            eu = _norm(e.key).upper()
            for n in norm:
                if n and (n in eu or eu in n):
                    return e
        return None

    def __repr__(self):
        return f"<Section {self.title!r} n={len(self.entries)}>"


class QuoteRef:
    """一次查价的**出处**：价落在报价表的哪一格。

    `price` 跟引擎平时返回的是**同一个数**——公式只是这个价的另一种写法，
    不是另算一遍。`sheet` / `ladder` / `entry` 记下"凭什么算出这个数"，
    公式生成器拿它去拼 INDEX/MATCH 或直接引用单元格。
    """

    def __init__(self, sheet, price, why="", cell=None, ladder=None, entry=None, zone=None):
        self.sheet = sheet
        self.price = price
        self.why = why                  # 与旧 surcharge_price 的说明串一致
        self.cell = cell                # 基础运费命中的价单元格 (行, 列)
        self.ladder = ladder            # 基础运费才有
        self.entry = entry              # 附加费才有
        self.zone = zone

    def target_cell(self):
        """价单元格。基础运费是命中的那一档，附加费是这项自己的格子。"""
        if self.ladder is not None:
            return self.cell
        if self.entry is not None:
            return self.entry.cell_for(self.zone)
        return None

    def __repr__(self):
        return f"<QuoteRef {self.sheet} {self.price} @{self.target_cell()}>"


# ── 解析 ─────────────────────────────────────────────────────────────────


def _cell(ws, r, c):
    return ws.cell(row=r, column=c).value


def _parse_ladder(ws, header_row, weight_col, max_col, max_row):
    """从表头行往下读阶梯数据，直到重量列不是数字。"""
    zone_cols = {}
    for c in range(weight_col + 1, max_col + 1):
        m = _ZONE_HEADER.match(str(_cell(ws, header_row, c) or ""))
        if m:
            lo = int(m.group(1))
            hi = int(m.group(2)) if m.group(2) else lo
            for z in range(lo, hi + 1):
                zone_cols[z] = c
        elif zone_cols:
            break                        # 分区表头连续；断了就结束
    if not zone_cols:
        return None

    rows, min_by_zone, min_cells, rung_rows = [], {}, {}, []
    all_numeric = True                   # 源表重量列是不是**真数字**（不是 '200lb-500lb' 这种文本）
    blank_streak = 0
    for r in range(header_row + 1, max_row + 1):
        raw = _cell(ws, r, weight_col)
        w = to_float(raw)
        src_num = isinstance(raw, (int, float)) and not isinstance(raw, bool)
        if w is None:
            txt = str(raw or "")
            # 百磅表写的是 '200lb-500lb' / '500+ lbs' / 'Min'，按文本分别处理
            if re.search(r"\bmin\b", txt, re.I):
                for z, c in zone_cols.items():
                    v = to_float(_cell(ws, r, c))
                    if v is not None:
                        min_by_zone[z] = v
                        min_cells[z] = (r, c)
                continue
            m = _NUM.search(txt)
            if m and zone_cols:
                w = float(m.group())
            else:
                blank_streak += 1
                if blank_streak >= 3:
                    break
                continue
        blank_streak = 0
        prices = {}
        for z, c in zone_cols.items():
            v = to_float(_cell(ws, r, c))
            if v is not None:
                prices[z] = v
        if prices:
            rows.append((w, prices))
            rung_rows.append((w, r))
            all_numeric = all_numeric and src_num
        elif rows:
            break                          # 阶梯结束（后面是附加费区）
    if not rows:
        return None
    # MATCH(...,1) 是拿**原表顺序**做二分查找的：升序与否必须取 sort 之前的源顺序，
    # 下面这句 rows.sort() 会把降序表也"修"成升序，把问题掩盖掉。
    src_w = [w for w, _ in rung_rows]
    src_r = [r for _, r in rung_rows]
    contiguous = src_r == list(range(src_r[0], src_r[-1] + 1))
    rows.sort(key=lambda x: x[0])
    return Ladder("", header_row, weight_col, zone_cols, rows, min_by_zone,
                  rung_rows=rung_rows,
                  match_safe=(all_numeric and contiguous and src_w == sorted(src_w)),
                  min_cells=min_cells)


def _ladder_title(ws, header_row, weight_col):
    """块标题：表头行上方最近的非空单元格（同列起往左找）。"""
    for r in range(header_row - 1, max(header_row - 4, 0), -1):
        v = _cell(ws, r, weight_col)
        if v not in (None, ""):
            return str(v).strip()
    return ""


def _parse_surcharge_sections(ws, max_row, max_col):
    """找所有「X 的右邻居含 收费标准」的单元格，作为区块起点。"""
    anchors = []
    for r in range(1, max_row + 1):
        for c in range(1, max_col):
            nxt = str(_cell(ws, r, c + 1) or "")
            if "收费标准" in nxt:
                title = str(_cell(ws, r, c) or "").strip()
                # 去掉 '费用名称' 这种通用后缀，留 'GROUND' / 'HOME DELIVERY' / 'HWT'
                clean = re.sub(r"费用名称|收费标准|[:：]", "", title).strip()
                anchors.append((r, c + 1, clean))
    anchors.sort()

    sections = []
    for i, (start_row, price_col, title) in enumerate(anchors):
        end_row = anchors[i + 1][0] - 1 if i + 1 < len(anchors) else max_row
        entries = _read_section(ws, start_row + 1, end_row, price_col)
        sections.append(SurchargeSection(title, entries,
                                         price_col=price_col, key_col=price_col - 1))
    return sections


def _read_section(ws, r0, r1, price_col):
    """读一个附加费区块：'费用项名 → 价'，分区定价时价在再右一列。"""
    entries, cur_key = [], ""
    blank_streak = 0
    for r in range(r0, r1 + 1):
        key_raw = _cell(ws, r, price_col - 1)
        key = str(key_raw).strip() if key_raw not in (None, "") else ""
        if key:
            cur_key = key
        val = _cell(ws, r, price_col)
        if val in (None, ""):
            blank_streak += 1
            if blank_streak >= 4:
                break
            continue
        blank_streak = 0
        if not cur_key:
            continue

        if _looks_like_zone_label(val):
            # 价格格里写的是分区 → 真价在右一列；同一项的后续分区行会留空费用项名
            zones = parse_zones(val)
            cell = (r, price_col + 1)
            price = to_float(_cell(ws, r, price_col + 1))
            if price is None or not zones:
                continue
            if entries and entries[-1].key == cur_key:
                # 合并单元格续行：同一个费用项横跨好几行分区，价仍然一格一格地记
                entries[-1].zones.update({z: price for z in zones})
                entries[-1].zone_cells.update({z: cell for z in zones})
            else:
                entries.append(SurchargeEntry(cur_key, zones={z: price for z in zones},
                                              zone_cells={z: cell for z in zones}))
        else:
            price = to_float(val)
            if price is None:
                continue
            entries.append(SurchargeEntry(cur_key, flat=price, flat_cell=(r, price_col)))
    return entries


# ── 报价簿 ───────────────────────────────────────────────────────────────


class QuoteBook:
    """一份报价表工作簿。解析结果按 sheet 缓存，重复查询不重复解析。"""

    def __init__(self, path, product_aliases=None, metadata_sheets=None,
                 hwt_min_weight=200):
        self.path = Path(path)
        # 空值 / 以 '_' 开头的键（_comment 之类）都当没配，否则会把产品映射到空 sheet 名
        self.product_aliases = {
            _product_key(k): str(v)
            for k, v in (product_aliases or {}).items()
            if k and not str(k).startswith("_") and v
        }
        self.hwt_min_weight = hwt_min_weight
        self._wb = None
        self._cache = {}                    # sheet -> (ladders, sections)
        self._index = None
        self.missing_products = set()       # 本次运行没配上 sheet 的产品名

    # -- 打开 -------------------------------------------------------------

    def _load(self):
        if self._wb is not None:
            return
        import openpyxl

        if not self.path.is_file():
            raise FileNotFoundError(f"报价表不存在：{self.path}")
        # data_only=True 读缓存值——报价表里分区价常是公式，只有值才有用
        self._wb = openpyxl.load_workbook(self.path, data_only=True)

    def _sheet_index(self):
        if self._index is None:
            self._load()
            self._index = {}
            for name in self._wb.sheetnames:
                if name in _META_SHEETS:
                    continue
                self._index[name] = _product_key(name)
        return self._index

    # -- 产品 → sheet -----------------------------------------------------

    def sheet_for_product(self, product):
        """产品名称 → 报价表 sheet 名；配不上返回 None 并记进 missing_products。"""
        key = _product_key(product)
        if not key:
            return None
        if key in self.product_aliases:
            name = self.product_aliases[key]
            if name in self._sheet_index():
                return name
            self.missing_products.add(f"{product}（别名指向的 sheet「{name}」不存在）")
            return None

        index = self._sheet_index()
        for name, k in index.items():
            if k and k == key:
                return name
        # 退一步：互为子串，且必须唯一命中，否则宁可判"配不上"也不乱配
        hits = [name for name, k in index.items() if k and (k in key or key in k)]
        if len(hits) == 1:
            return hits[0]
        self.missing_products.add(str(product))
        return None

    # -- 解析一张 sheet ---------------------------------------------------

    def _parsed(self, sheet):
        if sheet in self._cache:
            return self._cache[sheet]
        self._load()
        ws = self._wb[sheet]
        max_row, max_col = ws.max_row, ws.max_column

        ladders = []
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                v = _cell(ws, r, c)
                if isinstance(v, str) and _WEIGHT_HEADER.match(v):
                    lad = _parse_ladder(ws, r, c, max_col, max_row)
                    if lad:
                        lad.title = _ladder_title(ws, r, c)
                        lad.is_hwt = bool(
                            re.search(r"hundredweight|百磅|\bmwt\b", lad.title, re.I)
                        )
                        ladders.append(lad)
        sections = _parse_surcharge_sections(ws, max_row, max_col)
        self._cache[sheet] = (ladders, sections)
        return self._cache[sheet]

    def ladders(self, sheet):
        return self._parsed(sheet)[0]

    def sections(self, sheet):
        return self._parsed(sheet)[1]

    def raw_cell(self, sheet, row, col):
        """原表某一格的**原始值**（data_only：公式取缓存值）。

        只给「写公式」的闸门验算用：公式将来在 Excel 里读的就是这一格，
        验算必须读同一个地方，否则验的是另一份数据。
        """
        self._load()
        if sheet not in self._wb.sheetnames:
            return None
        return self._wb[sheet].cell(row=row, column=col).value

    def describe(self, sheet):
        """给人看的解析摘要，用于跑完后的报告。"""
        ladders, sections = self._parsed(sheet)
        lines = [f"[{sheet}]"]
        for lad in ladders:
            lo, hi = lad.weight_range()
            kind = "住宅" if lad.is_residential else "商业"
            if lad.is_hwt:
                kind += "·百磅"
            lines.append(
                f"  阶梯·{kind} {lad.title[:38]!r} 重量 {lo:g}-{hi:g} "
                f"分区 {sorted(lad.zone_cols)} 保底 {sorted(lad.min_by_zone.items()) or None}"
            )
        for sec in sections:
            lines.append(f"  附加费区块 {sec.title or '(通用)'!r} {len(sec.entries)} 项")
        return "\n".join(lines)

    # -- 基础运费 ---------------------------------------------------------

    def _pick_ladder(self, sheet, kind, weight=None):
        """商业/住宅各挑一段阶梯。

        住宅 = 标题带 Residential / Home Delivery；
        商业 = 其余里取"档位最多"的那段——百磅/MWT 块只有三四档，
               会被自动跳过，选中通用的 Ground 阶梯。
        但计费重 >= hwt_min_weight 时要反过来优先百磅块：UPS Ground-BD 那张表里
        百磅价和地面价差一个量级（商业偏远 22.5 vs 2.25），用错整列都错。
        """
        ladders = self.ladders(sheet)
        if not ladders:
            return None
        resi = [l for l in ladders if l.is_residential]
        comm = [l for l in ladders if not l.is_residential]
        pool = resi or comm if kind == "residential" else (comm or resi)
        if not pool:
            return None
        w = to_float(weight)
        if w is not None and w >= self.hwt_min_weight:
            hwt = [l for l in pool if l.is_hwt]
            if hwt:
                return max(hwt, key=lambda l: len(l.rows))
        return max(pool, key=lambda l: len(l.rows))

    def base_ref(self, product, weight, zone, kind="commercial", round_up=True):
        """基础运费：返回 QuoteRef（含出处），查不到返回 None。"""
        sheet = self.sheet_for_product(product)
        if sheet is None:
            return None
        lad = self._pick_ladder(sheet, kind, weight)
        if lad is None:
            return None
        price = lad.lookup(weight, zone, round_up=round_up)
        if price is None:
            return None
        return QuoteRef(sheet, price,
                        why=f"{sheet}/阶梯·{'住宅' if lad.is_residential else '商业'}",
                        cell=lad.cell_for_weight(weight, zone, round_up=round_up),
                        ladder=lad, zone=to_int(zone))

    def base_price(self, product, weight, zone, kind="commercial", round_up=True):
        """基础运费。kind: commercial / residential。查不到返回 None。"""
        ref = self.base_ref(product, weight, zone, kind=kind, round_up=round_up)
        return ref.price if ref else None

    # -- 附加费 -----------------------------------------------------------

    def surcharge_ref(self, product, items, zone=None, weight=None,
                      section_prefer=None, hwt_keywords=("HWT", "MWT", "百磅")):
        """按附加费项名查价：返回 (QuoteRef | None, 说明)。

        :param items: 候选费用项名（表里写法不统一，多给几个）
        :param section_prefer: 区块名关键词优先级，如 ["HOME DELIVERY", "住宅"]
        :param weight: 计费重；>=hwt_min_weight 时优先查 HWT/MWT 区块（那套价通常高一个量级）
        """
        sheet = self.sheet_for_product(product)
        if sheet is None:
            return None, f"产品「{product}」没配上报价表 sheet"

        sections = self.sections(sheet)
        if not sections:
            return None, f"报价表 {sheet} 里没解析到附加费区块"

        order = []
        if weight is not None and to_float(weight) is not None and float(weight) >= self.hwt_min_weight:
            for sec in sections:
                if any(_section_matches(sec.title, k) for k in hwt_keywords):
                    order.append(sec)
        for pref in (section_prefer or []):
            for sec in sections:
                if _section_matches(sec.title, pref) and sec not in order:
                    order.append(sec)
        for sec in sections:
            if sec not in order:
                order.append(sec)               # 兜底：按文档顺序都试一遍

        for sec in order:
            entry = sec.find(items)
            if entry is None:
                continue
            price = entry.price(zone)
            if price is None:
                continue
            where = sec.title or "通用"
            tag = entry.key if entry.flat is not None else f"{entry.key}·{zone}区"
            why = f"{sheet}/{where}/{tag}"
            return QuoteRef(sheet, price, why=why, entry=entry,
                            zone=to_int(zone)), why
        return None, f"{sheet} 里没找到 {'/'.join(items)} 的报价"

    def surcharge_price(self, product, items, zone=None, weight=None,
                        section_prefer=None, hwt_keywords=("HWT", "MWT", "百磅")):
        """:return: (价格, 说明) —— 查不到是 (None, 原因)"""
        ref, why = self.surcharge_ref(product, items, zone=zone, weight=weight,
                                      section_prefer=section_prefer,
                                      hwt_keywords=hwt_keywords)
        return (ref.price if ref else None), why


# ── 汇总运费（账单1 成本汇总，用于自检基础运费） ─────────────────────────


def load_summary_index(summary_path, summary_sheet="成本汇总",
                       log=None) -> dict:
    """读账单1 的《成本汇总》，返回 主单号 → {客户单号, 产品名称, 分区, 预报重, 派送费}。

    SOP 里这句 VLOOKUP 本身是自相矛盾的（要求按主单号匹配却给了向右的列号），
    这里改成先建索引再回填值——跨工作簿引用一旦报价表换路径就会全变 #REF!。
    """
    import openpyxl

    # 不能用 read_only=True：环洋这两份账单的表维度信息是坏的，
    # 只读模式会把它当成 A1:A1，读出来只有一个单元格。
    wb = openpyxl.load_workbook(summary_path, data_only=True)
    if summary_sheet not in wb.sheetnames:
        if log:
            log(f"[警告] {summary_path} 里没有 {summary_sheet}，改用第一个 sheet")
        summary_sheet = wb.sheetnames[0]
    ws = wb[summary_sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return {}
    header = [_norm(h) for h in rows[0]]

    def col(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None

    c_no = col("主单号")
    c_cust = col("客户单号")
    c_prod = col("产品名称")
    c_zone = col("分区")
    c_wt = col("预报重")
    c_fee = col("派送费")
    if c_no is None or c_cust is None:
        raise ValueError("《成本汇总》缺 主单号 或 客户单号 列，请确认选对了账单1")

    index = {}
    for row in rows[1:]:
        if c_no >= len(row) or row[c_no] in (None, ""):
            continue
        key = _norm(row[c_no])
        item = {
            "客户单号": row[c_cust] if c_cust is not None else None,
            "产品名称": row[c_prod] if c_prod is not None else None,
            "分区": to_int(row[c_zone]) if c_zone is not None else None,
            "预报重": to_float(row[c_wt]) if c_wt is not None else None,
            "派送费": to_float(row[c_fee]) if c_fee is not None else None,
        }
        # 同一个主单号可能在总表里出现多次（多件），保留第一条用于回填
        index.setdefault(key, item)
    return index


def strip_waybill_suffix(no):
    """运单号去后缀：只取 '_' 前面的部分（SOP：分列后取前半段）。"""
    return str(no or "").split("_", 1)[0].strip()
