# -*- coding: utf-8 -*-
"""SKYE 报价表解析与查价。

报价表是渠道方提供的 Excel，每个 sheet 通常包含两块：
  1) 基础运费表：weight × Zone 2..8，行 = 重量（LB），列 = 分区
  2) 附加费表：费用项 × 分区/类型，价格列对应一个数字

HWT/MWT 百磅表是另一种结构：行是分档（200LB+ / 500LB+ / Min），列是分区。

本模块只关心查价，列位置、合并、附注、缩进都自动忽略。
所有解析都做缓存（同一进程多次查价只解析一次）。
"""

from __future__ import annotations

import math
import os
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


# ============== 工具函数 ==============

def _norm(s) -> str:
    """表头文本归一：去空白与小尾巴。"""
    if s is None:
        return ""
    s = str(s).replace("\n", "").replace("\r", "").strip()
    return re.sub(r"\s+", "", s)


def _to_float(v) -> float | None:
    if v is None or v == "" or v == "-":
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _to_int(v) -> int | None:
    f = _to_float(v)
    if f is None:
        return None
    return int(f)


def _col_letter_to_index(letter: str) -> int:
    """A -> 1, Z -> 26, AA -> 27, BK -> 63 ..."""
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def _index_to_col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


# ============== 基础运费表（weight × zone） ==============

_ZONE_HEADER_RE = re.compile(r"^zone\s*-?\s*(\d+)$", re.I)
_ZONE_RANGE_RE = re.compile(r"^zones?\s*(\d+)\s*[-–~]\s*(\d+)$", re.I)
_ZONE_PLUS_RE  = re.compile(r"^zones?\s*(\d+)\s*\+$", re.I)
# FedEx 的分档行写作 "200LB+/lb"，比 UPS 的 "200lb+" 多一个 /lb 尾巴，故后缀可选
_HWT_TIER_RE   = re.compile(r"^(\d+)\s*l?b\s*\+(?:\s*/\s*l?b)?\s*$", re.I)
_HWT_MIN_RE    = re.compile(r"^min(?:imum)?$", re.I)


def _extract_num(v) -> float | None:
    """从 '$81/每票'、'81元'、'USD 81' 之类文本里抠出第一个数字。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"-?\d+(?:\.\d+)?", str(v).replace(",", ""))
    return float(m.group(0)) if m else None


def _norm_item(s) -> str:
    """费用项名的比较形式：去掉所有空白与连字符。

    报价表里的名字常带英文后缀或换行（`偏远费- Remote`、`额外处理费\\n超重费`、
    `住宅派送费Residential Surcharge`），config 里写的是干净的中文名。
    """
    return re.sub(r"[\s\-]+", "", str(s or ""))


def _parse_zone_header(s) -> tuple[str, int | tuple[int, int]]:
    """解析一个表头 cell：'Zone 2' / 'Zone-2' / 'Zones 3-4' / 'Zones 7+' / 'zones3-4'。

    返回 (canonical, key)，canonical 用于显示、key 用于查表。
    key 是 int 或 (int,int) 或 '7+'(str 哨兵)。
    """
    n = _norm(s).lower()
    m = _ZONE_HEADER_RE.match(n)
    if m:
        return n, int(m.group(1))
    m = _ZONE_RANGE_RE.match(n)
    if m:
        return n, (int(m.group(1)), int(m.group(2)))
    m = _ZONE_PLUS_RE.match(n)
    if m:
        return n, f"{m.group(1)}+"
    return n, None


def _zone_match(zone_key, target_zone: int) -> bool:
    """target_zone 命中这个 zone 列定义吗？"""
    if isinstance(zone_key, int):
        return zone_key == target_zone
    if isinstance(zone_key, tuple):
        return zone_key[0] <= target_zone <= zone_key[1]
    if isinstance(zone_key, str) and zone_key.endswith("+"):
        try:
            return target_zone >= int(zone_key[:-1])
        except ValueError:
            return False
    return False


# ============== 附加费表 ==============

_VALID_KINDS = {
    "商业", "住宅",
    "Remote", "Alaska", "Hawaii",  # 偏远/超偏远的特殊区段
    "普通签名", "成人签名",         # 签名费分项
    "拦截附加费", "退回运费", "改派成功与否以服务商实际情况为准", "改派运费",
    "/",                            # 占位行（“Q:/” 这种）
}


def _parse_kind(s) -> str | None:
    """识别表头的"类型"：商业/住宅/Remote/Alaska/Hawaii/普通签名 等。

    对数字 (2.8/25.25 等) 返回 None——避免把价格误判成 kind。
    """
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return None
    n = _norm(s)
    if n in _VALID_KINDS:
        return n
    return None


# ============== 数据类 ==============

@dataclass
class BaseTable:
    """一张基础运费表：weight(int) -> {zone_key -> price}。

    lookup(weight, zone, round_up=True) -> price or None
    """
    weights: list[int] = field(default_factory=list)
    data: dict[int, dict[Any, float]] = field(default_factory=dict)
    weight_col: int = -1
    zone_cols: list[tuple[int, Any]] = field(default_factory=list)  # (col_index, zone_key)
    block_index: int = 0

    def lookup(self, weight: float, zone: int, round_up: bool = True) -> float | None:
        if not self.weights:
            return None
        # 找到 weight 行的索引
        w_int = int(math.ceil(weight)) if round_up else int(weight)
        idx = -1
        for i, w in enumerate(self.weights):
            if w >= w_int:
                idx = i
                break
        if idx == -1:
            idx = len(self.weights) - 1  # 超出最大行用最后一档
        chosen_weight = self.weights[idx]
        row = self.data.get(chosen_weight, {})
        for col, zkey in self.zone_cols:
            if _zone_match(zkey, zone):
                return row.get(col)
        return None


@dataclass
class HwtTable:
    """百磅表：分档 -> {zone_key -> rate}；额外有 Min 价。

    compute(total_weight, zone) -> {rate_total, min_charge, final, picked_tier}
    """
    tiers: list[tuple[str, dict[Any, float]]] = field(default_factory=list)  # [(tier_label, {col: rate})]
    min_prices: dict[Any, float] = field(default_factory=dict)  # {col: min_price}
    zone_cols: list[tuple[int, Any]] = field(default_factory=list)
    block_index: int = 0

    def compute(self, total_weight: float, zone: int) -> dict | None:
        if not self.tiers:
            return None
        # 选档：>=500 用 500LB+，否则 200LB+
        if total_weight >= 500:
            label, row = self.tiers[-1]
        else:
            label, row = self.tiers[0]
        for col, zkey in self.zone_cols:
            if _zone_match(zkey, zone):
                rate = row.get(col)
                if rate is None:
                    return None
                total = round(total_weight * rate, 2)
                minp = self.min_prices.get(col)
                final = max(total, minp) if minp is not None else total
                return {
                    "tier": label,
                    "rate": rate,
                    "weight": total_weight,
                    "zone": zone,
                    "rate_total": total,
                    "min_price": minp,
                    "final": final,
                }
        return None


@dataclass
class SurchargeTable:
    """附加费表：按费用项聚合。

    surcharge[费用项] = {
        "by_zone":   {zone_key -> price},        # 按分区取价
        "by_kind":   {kind -> price},            # 按商业/住宅取价
        "flat":      float,                      # 唯一固定价
        "ranges":    [(zone_key_set, price)],    # 复杂多档（备用）
        "caps_*":    同结构，存"封顶价"（整票上限），取不到就是 None
    }

    封顶价来自报价表的 `封顶MWT` / `封顶HWT` 列：写数字的才有封顶，
    `无` / `不封顶` / `是` / 空 都当没有封顶（`是` 那几行报价表没给数值）。
    """
    items: dict[str, dict] = field(default_factory=dict)
    block_index: int = 0

    def lookup(self, item: str, zone: int | None = None, kind: str | None = None) -> float | None:
        info = self.items.get(item)
        if not info:
            return None
        if kind and kind in info["by_kind"]:
            return info["by_kind"][kind]
        if zone is not None:
            for zkey, price in info["by_zone"].items():
                if _zone_match(zkey, zone):
                    return price
        return info["flat"]


@dataclass
class ServiceMatch:
    kind: str          # "weight_zone" | "hundredweight"
    sheet: str
    block: int
    surcharge_kind: str | None  # "商业" / "住宅" / None
    base_block: int = 0         # 同一张表里"基础运费阶梯"是第几个 block
                                # （百磅表与基础阶梯可能并存，如 FedEx Multiweigh）
    below_hwt: str | None = None  # 百磅服务但重量 <200lb 时怎么办：
                                  # "base"=改查本表的基础运费阶梯；None=不报价（留空）


# ============== 核心：QuoteBook ==============

class QuoteBook:
    """懒加载的报价表查询器。"""

    def __init__(self, path, service_map: list[dict],
                 hwt_discount: float = 1.0, discount_min_price: bool = False,
                 book_dir=None):
        self.path = Path(path)
        # 账单所在目录。生成跨工作簿公式时用它把报价表路径转成**相对路径**：
        # 写死 `'C:\Users\TT1\Desktop\…\[报价表.xlsx]Sheet'!` 换台电脑就找不到文件，
        # 公式全变 #REF!/#N/A。相对路径只要报价表与账单保持同样的相对位置就能用。
        self._book_dir = Path(book_dir) if book_dir else None
        self._service_map = list(service_map or [])
        # 百磅实际结算价 = 档位价 × hwt_discount（账单普遍比报价低约 3.75%）。
        # discount_min_price=True 时保底价一并打折，否则保底价原价兜底。
        self._hwt_discount = hwt_discount
        self._discount_min_price = discount_min_price
        self._base_cache: dict[tuple[str, int], BaseTable] = {}
        self._hwt_cache:  dict[tuple[str, int], HwtTable]  = {}
        self._sur_cache:  dict[tuple[str, int], SurchargeTable] = {}
        self._wb = None  # 延迟 openpyxl 导入
        self._openpyxl = None

    # ---------- 入口 ----------

    def _load(self):
        if self._wb is not None:
            return
        try:
            import openpyxl  # noqa: PLC0415
        except ImportError as exc:  # noqa: BLE001
            raise SystemExit("缺少 openpyxl：请先执行 pip install openpyxl") from exc
        self._openpyxl = openpyxl
        if not self.path.is_file():
            raise FileNotFoundError(f"报价表文件不存在：{self.path}")
        self._wb = openpyxl.load_workbook(self.path, data_only=True, read_only=True)

    def _sheet(self, name: str):
        self._load()
        if name not in self._wb.sheetnames:
            raise KeyError(f"报价表 sheet 不存在：{name}（实际：{self._wb.sheetnames}）")
        return self._wb[name]

    # ---------- Service → sheet 解析 ----------

    def service_resolver(self, service_code: str) -> ServiceMatch | None:
        """匹配 Service Code —— 取**最长**的那条规则。

        必须最长优先：'Fedex-MWT-01-ORD' 同时命中 'FedEx-MWT'(LAX) 和 'MWT-01-ORD'，
        若按配置顺序取第一个就会把 ORD 的运单算到 LAX 报价表上。
        """
        if not service_code:
            return None
        sc = str(service_code).strip().lower()
        best: dict | None = None
        best_len = -1
        for rule in self._service_map:
            m = str(rule.get("match", "")).strip().lower()
            if m and m in sc and len(m) > best_len:
                best, best_len = rule, len(m)
        if best is None:
            return None
        return ServiceMatch(
            kind=best["kind"],
            sheet=best["sheet"],
            block=int(best.get("block", 0)),
            surcharge_kind=best.get("surcharge_kind"),
            base_block=int(best.get("base_block", 0)),
            below_hwt=best.get("below_hwt"),
        )

    # ---------- 基础运费 ----------

    def base(self, sheet: str, block: int = 0) -> BaseTable:
        self._load()
        key = (sheet, block)
        if key not in self._base_cache:
            self._base_cache[key] = self._parse_base_table(self._sheet(sheet), block)
        return self._base_cache[key]

    def base_price(self, sheet: str, block: int, weight, zone, round_up: bool = True) -> float | None:
        bt = self.base(sheet, block)
        return bt.lookup(_to_float(weight) or 0.0, _to_int(zone) or 0, round_up=round_up)

    # ---------- 生成查价公式（跨工作簿 VLOOKUP，同人工模板） ----------

    def external_ref(self, sheet: str) -> str:
        """external reference 前缀：`'[<相对目录>\\][<报价表文件名>]<sheet>'!`

        目录**相对账单所在目录**写（Excel 原生写法）：
          * 与账单同目录 → `'[报价表.xlsx]UPS-Ground商业'!`
          * 在账单上一级   → `'..\\[报价表.xlsx]UPS-Ground商业'!`
        不同盘符没法相对，只能退回绝对路径。给不出相对路径时（没传 book_dir）
        保持旧的绝对路径写法。
        """
        rel = self._rel_dir()
        return f"'{rel}[{self.path.name}]{sheet}'!"

    def _rel_dir(self) -> str:
        """报价表目录相对账单目录的写法，带尾部分隔符；同目录返回空串。"""
        if self._book_dir is None:
            parent = str(self.path.parent)
            return "" if parent in (".", "") else f"{parent}\\"
        try:
            rel = os.path.relpath(self.path.parent, self._book_dir)
        except ValueError:          # 不同盘符（Windows 才抛）
            return f"{self.path.parent}\\"
        return "" if rel == "." else f"{rel}\\"

    def vlookup_formula(self, sheet: str, block: int, row: int, weight_letter: str,
                        zone, weight=None, round_up: bool = True) -> str | None:
        """查基础运费阶梯的 VLOOKUP 文本；表结构不符合模板形状时返回 None。

        形状：=VLOOKUP(ROUNDUP(H{row},0),'<相对目录>\\[<报价表>]<sheet>'!$B:$I,{zone},0)
          * 区间首列 = 重量列（阶梯表里重量列正好在 Zone 2 左边一列）
          * 区间末列 = 最后一个分区列
          * 第 3 参数 = zone：区间内第 zone 列正好就是 Zone {zone}（Zone 2 落在第 2 列）
        所以先校验 "列序 - 重量列序 + 1 == 分区号"，不成立就不用公式（宁可退化成数值）。
        取整必须跟 `base_price(round_up=…)` 一致：阶梯只按整数磅排，小数重量精确匹配会 #N/A，
        故 round_up=True 时包 ROUNDUP（与 `round_up_lookup: true` 对应）。给了 weight 时再校验
        这一票确实落在阶梯行上——超过阶梯上限的票 `base_price` 会夹到最后一档，那种情况公式
        给不出值，同样退回写数值。
        """
        bt = self.base(sheet, block)
        zc = bt.zone_cols or []
        zi = _to_int(zone)
        if bt.weight_col < 0 or len(zc) < 2 or zi is None:
            return None
        if not all(col - bt.weight_col + 1 == _to_int(z) for col, z in zc):
            return None
        target = [(col, z) for col, z in zc if _zone_match(z, zi)]
        if not target:
            return None
        wf = _to_float(weight) if weight is not None else None
        if wf is not None:
            w_int = int(math.ceil(wf)) if round_up else int(wf)
            row_vals = bt.data.get(w_int)
            if not row_vals or not any(row_vals.get(col) for col, _ in target):
                return None
        first = _index_to_col_letter(bt.weight_col + 1)
        last = _index_to_col_letter(max(col for col, _ in zc) + 1)
        key = f"ROUNDUP({weight_letter}{row},0)" if round_up else f"INT({weight_letter}{row})"
        return (f"=VLOOKUP({key},{self.external_ref(sheet)}"
                f"${first}:${last},{zi},0)")

    # ---------- 百磅 ----------

    def hwt(self, sheet: str, block: int = 0) -> HwtTable:
        self._load()
        key = (sheet, block)
        if key not in self._hwt_cache:
            self._hwt_cache[key] = self._parse_hwt_table(self._sheet(sheet), block)
        return self._hwt_cache[key]

    def hwt_price(self, sheet: str, block: int, total_weight, zone) -> dict | None:
        ht = self.hwt(sheet, block)
        info = ht.compute(_to_float(total_weight) or 0.0, _to_int(zone) or 0)
        if not info:
            return None
        f = self._hwt_discount
        if f != 1.0:
            info["list_rate_total"] = info["rate_total"]
            info["list_final"] = info["final"]
            info["rate_total"] = round(info["rate_total"] * f, 2)
            floor = info["min_price"]
            if floor is not None and self._discount_min_price:
                floor = round(floor * f, 2)
            info["final"] = max(info["rate_total"], floor) if floor is not None else info["rate_total"]
            info["discount"] = f
        return info

    # ---------- 附加费 ----------

    def surcharge(self, sheet: str, block: int = 0) -> SurchargeTable:
        self._load()
        key = (sheet, block)
        if key not in self._sur_cache:
            self._sur_cache[key] = self._parse_surcharge_table(self._sheet(sheet), block)
        return self._sur_cache[key]

    def _surcharge_lookup(
        self, sheet: str, block: int, item: str, zone: int | None = None,
        kind: str | None = None, fallback_sheets: list[str] | None = None,
    ) -> tuple[float | None, float | None]:
        """按 sheet -> fallback_sheets 的顺序查一个费用项，返回 (单价, 封顶价)。

        单价和封顶必须取自同一行，所以两个值一起返回。
        先精确匹配费用项名，再按去空格/连字符后的前缀匹配——**先在本 sheet 里找**，
        找不到才去 fallback：HWT/MWT 表里才是带封顶的那一份，落到普通 Ground 表就丢封顶。
        """
        for sh in [sheet] + list(fallback_sheets or []):
            st = self.surcharge(sh, block if sh == sheet else 0)
            info = st.items.get(item)
            if info is None:
                key = _norm_item(item)
                for k, v in st.items.items():
                    if key and _norm_item(k).startswith(key):
                        info = v
                        break
            if not info:
                continue
            if kind and kind in info["by_kind"]:
                return info["by_kind"][kind], info["caps_by_kind"].get(kind)
            if zone is not None:
                for zkey, price in info["by_zone"].items():
                    if _zone_match(zkey, zone):
                        cap = info["caps_by_zone"].get(zkey)
                        if cap is None and info["caps"]:
                            cap = min(info["caps"])
                        return price, cap
            flats = info.get("flats") or ([info["flat"]] if info["flat"] is not None else [])
            if flats:
                return min(flats), (min(info["caps"]) if info["caps"] else None)
        return None, None

    def surcharge_price(
        self, sheet: str, block: int, item: str, zone: int | None = None, kind: str | None = None,
        fallback_sheets: list[str] | None = None,
    ) -> float | None:
        """在主 sheet 查不到费用项时，按 fallback_sheets 顺序兜底。"""
        return self._surcharge_lookup(sheet, block, item, zone, kind, fallback_sheets)[0]

    def surcharge_cap(
        self, sheet: str, block: int, item: str, zone: int | None = None, kind: str | None = None,
        fallback_sheets: list[str] | None = None,
    ) -> float | None:
        """该费用项整票的封顶价（报价表 `封顶MWT`/`封顶HWT` 列），没有封顶返回 None。"""
        return self._surcharge_lookup(sheet, block, item, zone, kind, fallback_sheets)[1]

    # ===========================================================
    # 解析实现
    # ===========================================================

    def _parse_base_table(self, ws, block_index: int) -> BaseTable:
        """从 sheet 中找 N 个 base block；按 block_index 选第几个。

        一个 block = 一段连续的'weight × Zone 2..8'表。识别规则：
          - 行内至少 3 个 cell 的 normalize 后能命中 zone 头 (Zone N / Zones X-Y / Zones N+)
          - 同一行内，weight 列位于首个 zone 列左侧
        """
        rows = list(ws.iter_rows(values_only=True))
        blocks: list[BaseTable] = []
        cur: BaseTable | None = None
        for r in rows:
            # 找这一行能命中的 zone 列
            zone_cols: list[tuple[int, Any]] = []
            for i, c in enumerate(r):
                canon, zkey = _parse_zone_header(c)
                if zkey is not None:
                    zone_cols.append((i, zkey))
            if len(zone_cols) >= 3:
                # 只保留从首个 zone 列起的**连续**区段。
                # UPS-Ground商业 这类表在同一行里横向排了多组 Zone 表
                # （基础运费 / 另一区域 / 附加费），全收会把不同表的价格混进同一行。
                zc = sorted(zone_cols)
                contig = [zc[0]]
                for col, zk in zc[1:]:
                    if col == contig[-1][0] + 1:
                        contig.append((col, zk))
                    else:
                        break
                zone_cols = contig

                # 找 weight 列：第一个 zone 列左侧最近的"看起来像 weight"的列
                first_zone_col = zone_cols[0][0]
                wcol = -1
                for j in range(first_zone_col - 1, -1, -1):
                    cell = r[j]
                    if cell is None or (isinstance(cell, str) and not cell.strip()):
                        continue
                    wcol = j
                    break
                # 只有"真正的表头行"（weight 列是文字，如 '重量（LB）'）才开新 block。
                # 否则只是在同一张表里又遇到一行 Zone 头（数据行里夹着的附加费区表头），
                # 保持当前 block 继续收数据 —— 否则基础运费表会被截断在表头出现的那一行。
                is_header_row = wcol >= 0 and isinstance(r[wcol], str)
                if cur is None:
                    cur = BaseTable(weight_col=wcol, zone_cols=zone_cols, block_index=len(blocks))
                    blocks.append(cur)
                elif cur.weight_col == wcol:
                    cur.zone_cols = zone_cols
                elif is_header_row:
                    cur = BaseTable(weight_col=wcol, zone_cols=zone_cols, block_index=len(blocks))
                    blocks.append(cur)
                # 其余情况：忽略该 Zone 头行，继续往当前 block 收数据
            # 解析 weight 行：一行可能同时属于多个 block。
            # 例如 UPS-Ground商业 在列 32-37 横向嵌了一张附加费表，它的表头落在第 32 行，
            # 而这一行同时也是基础运费表的数据行（LB=28）。只写 cur 会把基础运费截断在 27 档。
            for bt in blocks:
                wc = bt.weight_col
                if wc < 0 or wc >= len(r):
                    continue
                wnum = _to_int(r[wc])
                if wnum is None or not (1 <= wnum <= 2000):
                    continue
                if wnum not in bt.data:
                    bt.weights.append(wnum)
                bt.data[wnum] = {col: _to_float(r[col]) for col, _ in bt.zone_cols}
        if not blocks:
            return BaseTable(block_index=0)
        return blocks[min(block_index, len(blocks) - 1)]

    def _parse_hwt_table(self, ws, block_index: int) -> HwtTable:
        """解析 HWT / MWT 表：行是分档（200LB+, 500LB+ / Min）。"""
        rows = list(ws.iter_rows(values_only=True))
        ht = HwtTable(block_index=block_index)
        # 找 header 行：含 "Zone 2..8"
        header_row = None
        for r in rows:
            zcols = [(i, _parse_zone_header(c)[1]) for i, c in enumerate(r) if _parse_zone_header(c)[1] is not None]
            if len(zcols) >= 3:
                header_row = r
                ht.zone_cols = zcols
                break
        if header_row is None:
            return ht
        # 找 weight 列：第一个 zone 列左侧第一个非空 cell
        first_zone_col = ht.zone_cols[0][0]
        wcol = -1
        for j in range(first_zone_col - 1, -1, -1):
            cell = header_row[j]
            if cell is None or (isinstance(cell, str) and not cell.strip()):
                continue
            wcol = j
            break
        if wcol < 0:
            return ht

        # FedEx 风格：保底价不在 Min 行，而在表头的 'Minimum Charge' 列（值形如 '$81/每票'）
        min_col = None
        for j, c in enumerate(header_row):
            cn = _norm(c).lower()
            if cn.startswith("minimum") or cn == "min":
                min_col = j
                break

        # 解析数据行
        for r in rows:
            label = r[wcol] if wcol < len(r) else None
            if label is None:
                continue
            label_n = _norm(label).lower()
            m = _HWT_TIER_RE.match(label_n)
            if m:
                tier_label = m.group(0)
                row = {col: _to_float(r[col]) for col, _ in ht.zone_cols}
                ht.tiers.append((tier_label, row))
                # 分档行同一行的 Minimum Charge 列即为保底价。
                # 只认第一个档位行：FedEx 第二个档位行该列写的是"单箱最低重25lb"（是重量不是钱）
                if min_col is not None and min_col < len(r) and not ht.min_prices:
                    mv = _extract_num(r[min_col])
                    if mv is not None:
                        ht.min_prices = {col: mv for col, _ in ht.zone_cols}
            elif _HWT_MIN_RE.match(label_n):
                ht.min_prices = {col: _to_float(r[col]) for col, _ in ht.zone_cols}
        return ht

    def _parse_surcharge_table(self, ws, block_index: int) -> SurchargeTable:
        """解析附加费表：识别 序号/费用项/分区/价格 表头，向下扫描。"""
        rows = list(ws.iter_rows(values_only=True))
        st = SurchargeTable(block_index=block_index)
        # 找表头行
        header_idx = None
        header_cols: dict[str, int] = {}
        for ridx, r in enumerate(rows):
            for i, c in enumerate(r):
                cn = _norm(c)
                if cn == "序号" and "序号" not in header_cols:
                    header_cols["序号"] = i
                elif cn == "费用项" and "费用项" not in header_cols:
                    header_cols["费用项"] = i
                elif cn == "分区" and "分区" not in header_cols:
                    header_cols["分区"] = i
                elif cn == "价格" and "价格" not in header_cols:
                    header_cols["价格"] = i
                elif cn.startswith("封顶") and "封顶" not in header_cols:
                    # 列名是 `封顶MWT`（FedEx MWT 表）/ `封顶HWT`（UPS HWT 表）
                    header_cols["封顶"] = i
            if "序号" in header_cols and "费用项" in header_cols:
                header_idx = ridx
                break
        if header_idx is None:
            return st
        # 扫描
        cur_item: str | None = None
        for r in rows[header_idx + 1:]:
            if "费用项" in header_cols and header_cols["费用项"] < len(r):
                v = r[header_cols["费用项"]]
                if v is not None and str(v).strip():
                    cur_item = str(v).replace("\n", " ").strip()
            if not cur_item:
                continue
            # 封顶价：只有数字才算；`无` / `不封顶` / `是` / 空 都当没封顶
            # （用严格的 _to_float，别用能从文本里抠数字的 _extract_num）
            cap = None
            if "封顶" in header_cols and header_cols["封顶"] < len(r):
                cap = _to_float(r[header_cols["封顶"]])
            zone = None
            if "分区" in header_cols and header_cols["分区"] < len(r):
                cell = r[header_cols["分区"]]
                if cell is not None and str(cell).strip():
                    z = _parse_zone_header(cell)
                    if z[1] is not None:
                        zone = z[1]
                    else:
                        # 可能是 "商业" / "住宅"
                        kind = _parse_kind(cell)
                        if kind:
                            price = _to_float(r[header_cols.get("价格", -1)]) if "价格" in header_cols else None
                            if price is not None:
                                self._add_surcharge(st, cur_item, kind=kind, price=price, cap=cap)
                            continue
            if "价格" in header_cols and header_cols["价格"] < len(r):
                price = _to_float(r[header_cols["价格"]])
                if price is None and "分区" in header_cols and header_cols["分区"] < len(r):
                    # 部分费用（地址修正/超不可发）把价格直接写在“分区”格里；
                    # 当价格列为空时，把分区的数字视为固定价。
                    price = _to_float(r[header_cols["分区"]])
                if price is None:
                    continue
                if zone is not None:
                    self._add_surcharge(st, cur_item, zone=zone, price=price, cap=cap)
                else:
                    self._add_surcharge(st, cur_item, flat=price, cap=cap)
        return st

    @staticmethod
    def _add_surcharge(st: SurchargeTable, item: str, zone=None, kind: str | None = None,
                       flat: float | None = None, price: float | None = None,
                       cap: float | None = None):
        info = st.items.setdefault(item, {"by_zone": {}, "by_kind": {}, "flats": [],
                                         "caps_by_zone": {}, "caps_by_kind": {}, "caps": []})
        if zone is not None:
            info["by_zone"][zone] = price
            if cap is not None:
                info["caps_by_zone"][zone] = cap
        if kind is not None:
            info["by_kind"][kind] = price
            if cap is not None:
                info["caps_by_kind"][kind] = cap
        if flat is not None:
            # 同一费用项在表中可能分多行（例如 UPS 主动 / 收件人主动），全部保留，
            # 查价时取最小。避免后一行覆盖前一行。
            if flat not in info["flats"]:
                info["flats"].append(flat)
            if info.get("flat") is None:
                info["flat"] = flat
            if cap is not None:
                info["caps"].append(cap)


# ============== 自测（直接运行本文件时跑一遍） ==============

def _selftest():
    if len(sys.argv) < 2:
        print("usage: python quote_engine.py <报价表.xlsx>")
        return
    import json  # noqa: PLC0415
    qb = QuoteBook(sys.argv[1], [
        {"match": "UPS-Ground-C",   "kind": "weight_zone", "sheet": "UPS-Ground住宅", "block": 0, "surcharge_kind": "住宅"},
        {"match": "UPS-Ground-T01", "kind": "weight_zone", "sheet": "UPS-Ground商业", "block": 0, "surcharge_kind": "商业"},
        {"match": "UPS GROUND HWT", "kind": "hundredweight", "sheet": "UPS Ground HWT", "block": 0, "surcharge_kind": "商业"},
        {"match": "FedEx Home Delivery", "kind": "weight_zone", "sheet": "FedEx Home Delivery  ", "block": 0, "surcharge_kind": "住宅"},
    ])
    print("T01 26LB zone 8:", qb.base_price("UPS-Ground商业", 0, 26, 8))   # 期望 14.79
    print("C  23LB zone 5:", qb.base_price("UPS-Ground住宅", 0, 23, 5))   # 期望 8.53
    print("HWT 245LB zone 7:", qb.hwt_price("UPS Ground HWT", 0, 245, 7))
    print("HWT 342LB zone 8:", qb.hwt_price("UPS Ground HWT", 0, 342, 8))
    print("UPS-Ground住宅 异形 Zone 6:", qb.surcharge_price("UPS-Ground住宅", 0, "额外处理费-异形费", zone=6, kind="住宅"))
    print("UPS-Ground住宅 偏远 商业:",   qb.surcharge_price("UPS-Ground住宅", 0, "偏远费", kind="商业"))
    print("UPS-Ground住宅 偏远 住宅:",   qb.surcharge_price("UPS-Ground住宅", 0, "偏远费", kind="住宅"))
    print("UPS-Ground住宅 超偏远 商业:", qb.surcharge_price("UPS-Ground住宅", 0, "超偏远费", kind="商业"))
    print("UPS-Ground住宅 地址修正:",    qb.surcharge_price("UPS-Ground住宅", 0, "地址修正"))


if __name__ == "__main__":
    _selftest()
