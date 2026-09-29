# -*- coding: utf-8 -*-
"""把「这次匹配」写成 Excel 公式。

SOP 的原话本来就是公式：「在H列**引用**"打单费用名称中英文翻译"表，匹配出中文」、
「差异=报价表-费用」（「粘贴为值」只针对运单号那一步）。永达那边也是这么落地的。

改回公式的代价是跨工作簿引用一断就是满屏 `#REF!`，所以这里只**生成字符串**，
写不写由 process_details.py 的闸门逐行决定：**先按 Excel 的规矩验算，与引擎的价
相等才写**。为此本模块里「拼公式」和「验算」是同一个函数——两边用的是同一批事实，
想漂也漂不开。

纯函数：不读文件、不碰 openpyxl。原表数据由调用方通过 `read_cell` 递进来。
"""
from __future__ import annotations

import math
import re
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

from quote_engine import col_letter, to_float

# 「看着像数字的文本」：Excel 会把它当数字用，`VALUE()` 也认这一种
_NUMERIC_TEXT = re.compile(r"^-?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$")


def xl_number(raw):
    """这一格在 Excel 的算术里会当几用；当不了数字就是 None。

    导出的账单**整列计费重都是文本**（`'26.000'`），而 Excel 对看着像数字的文本会
    隐式转成数字——`CEILING("26.000",1)` 在 Excel 里照样是 27。所以「不是 float 就
    不写公式」这把尺子太粗，会把所有行都毙掉。

    这里按 Excel 的转换规则判：纯数字文本（可带正负号、小数点、指数）算数，
    别的一概不算。千分位逗号、`inf`/`nan`、下划线分隔这些 Python 认、Excel 不认的，
    正则本身就挡住了。
    """
    if raw is None or isinstance(raw, bool):
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip()
    if not _NUMERIC_TEXT.match(s):
        return None
    return float(s)


def excel_round(value, digits=2):
    """Excel 的 `ROUND`：**四舍五入（逢半远离零）**。

    跟 Python 的 `round()` 不是一回事——`round` 是逢半取偶（banker's）：
    `round(17.625, 2)` 给 17.62 而 Excel 给 17.63。闸门要拿 Excel 的语义去验，
    就必须用它自己的规则，否则「验过」是假的。
    """
    if value is None:
        return None
    q = Decimal(1).scaleb(-digits)
    return float(Decimal(repr(float(value))).quantize(q, rounding=ROUND_HALF_UP))


# ── 外部引用 ─────────────────────────────────────────────────────────────


def ext_ref(path, ref_dir, sheet, allow_absolute=False):
    """外部引用的 `'[文件.xlsx]sheet'` 段；写不出可解析的形式就 None。

    Excel 只有两种形式：

      * **裸文件名**：`'[报价表.xlsx]sheet'!$B:$I` —— 按**引用方工作簿自己所在的
        目录**解析。所以账单和报价表必须放在同一个文件夹里。
      * **完整绝对路径**：目录在中括号**外面** —— `'C:\\dir\\[报价表.xlsx]sheet'!$B:$I`。

    `..\\` 那种「父目录相对路径」Excel 不认（永达的 SOP 只写裸文件名、其脚本宁可用
    「把报价表复制到账单旁边」来兜，就是因为相对路径走不通），所以这里不生成。
    默认 `allow_absolute=False`：报价表不在账单旁边就返回 None，由调用方退回写值。

    :param sheet: 目标 sheet 名
    :param ref_dir: 引用方（账单）所在目录
    """
    p = Path(str(path))
    name = p.name
    if not name or any(ch in name for ch in "[]"):
        return None                 # 文件名里带中括号，Excel 语法上没法表达
    if not sheet or any(ch in str(sheet) for ch in "[]"):
        return None

    same_dir = False
    if ref_dir is not None:
        try:
            # p.resolve() 按**当前工作目录**解析，跟程序当初打开它是同一个基准
            same_dir = p.resolve().parent == Path(ref_dir).resolve()
        except OSError:
            same_dir = False

    if same_dir:
        head = ""
    elif allow_absolute and p.is_absolute():
        head = str(p.parent) + "\\"
    else:
        return None

    token = f"{head}[{name}]{sheet}"
    # 整个 `路径[文件]sheet` 段用单引号包住（不必要也合法），名里的单引号双写
    return "'" + token.replace("'", "''") + "'"


def cell_ref(token, row, col):
    """`'[报价表.xlsx]sheet'!$C$5` —— 一律行列都锁。

    阶梯区间、保底价、附加费价格子都不该随复制填充漂移。
    """
    return f"{token}!${col_letter(col)}${row}"


# ── 翻译列 ───────────────────────────────────────────────────────────────


def translation_formula(row, fee_col, token, key_col="A", val_col="B"):
    """`=VLOOKUP(TRIM(<费用名称>),'[翻译表.xlsx]Sheet1'!$A:$B,2,0)`

    用 `$A:$B` 而不是写死结束行：翻译表往后加行，公式不用改。
    `TRIM` 是因为账单那边的费用名常带首尾空格——引擎的取数键也是 strip 过的。
    """
    if not token:
        return None
    return (f"=VLOOKUP(TRIM({col_letter(fee_col)}{row}),"
            f"{token}!${key_col}:${val_col},2,0)")


def vlookup_first(a_col, b_col, want):
    """复刻 `VLOOKUP(...,0)` 的首匹配，返回 (命中值, 是否可信)。

    Excel 的 VLOOKUP **大小写不敏感**、文本与数字还会互相靠拢（`="123"=123` 为真）。
    这些宽松规则我复刻不全，所以：**只有「宽松首匹配」和「严格首匹配」落在同一行、
    且该行是完全相等**时才判可信，否则退回写值。

    宁可多退回几个值（少写一条公式），也不能写出一个跟引擎结论对不上的公式——
    后者会静默地错。
    """
    w = str(want if want is not None else "").strip()
    if not w:
        return None, False
    loose = [str(x if x is not None else "").strip().casefold() for x in a_col]
    w_loose = w.casefold()
    i_loose = next((i for i, v in enumerate(loose) if v == w_loose), None)
    i_exact = next((i for i, x in enumerate(a_col)
                    if isinstance(x, str) and x.strip() == w), None)
    if i_loose is None or i_loose != i_exact:
        return None, False
    return b_col[i_exact], True


# ── 报价列 ───────────────────────────────────────────────────────────────


def base_formula(token, ladder, zone, weight, read_cell, round_up=True,
                 weight_cell=None):
    """基础运费（计费重 × 分区 查阶梯）。

    公式形状：

        档位命中   =ROUND(INDEX(<分区列>,MATCH(CEILING(<计费重>,1),<重量列>,1)),2)
        另有保底   ... MAX(INDEX(...), <Min 单元格>) ...
        轻于第一档 IFERROR(INDEX(...), <第一档单元格>)

    几个刻意的选择：

    * `MATCH(...,1)` 要求重量列**原表升序**。降序的表引擎会在解析时 `sort()` 掉，
      公式却不会——所以调用方必须先确认 `ladder.match_safe`（原表升序、区间连续、
      档位是纯数字），否则不写公式。
    * `IFERROR` **只在该行重量确实低于第一档时**才裹。否则它会把断链的 `#REF!`
      也一并吞成第一档的价，那是个静默的错价，比报错更糟。
    * 保底写**单元格**而不是字面量，否则改了报价表的保底价公式不会跟。
    * 取整用 `TRUNC` 不是 `INT`：Python 的 `int()` 向零截断，Excel 的 `INT()` 是
      向下取整，负数上会差一。

    :param read_cell: `(行, 列) -> 原表值`，验算与拼公式用的是同一批数
    :param weight: 账单行的计费重（引擎读到的那个数）
    :return: (公式, Excel 语义下的值)；值为 None 表示这条不能写成公式
    """
    z = int(zone)
    col = ladder.zone_cols.get(z)
    r0, r1 = ladder.data_row_range
    weight = xl_number(weight)
    if col is None or r0 is None or weight is None or not ladder.match_safe:
        return None, None

    key = math.ceil(weight) if round_up else int(weight)

    # 按 Excel 的 MATCH(...,1) 找落点：区间里最后一个 <= 查找值的行
    hit_r, prev = None, None
    for r in range(r0, r1 + 1):
        wv = read_cell(r, ladder.weight_col)
        if not isinstance(wv, (int, float)) or isinstance(wv, bool):
            return None, None           # 区间里混了非数字 → 二分查找的结果没法照搬
        if prev is not None and wv < prev:
            return None, None           # 没升序就谈不上 MATCH(...,1)
        prev = wv
        if wv <= key:
            hit_r = r
    below_first = hit_r is None

    src_r = r0 if below_first else hit_r
    raw = read_cell(src_r, col)
    if raw is None or isinstance(raw, str):
        # 空格子在 Excel 里 INDEX 出来是 **0**，而引擎那边是"没这个价"——差得太远
        return None, None
    value = float(raw)

    floor_cell = ladder.min_cells.get(z)
    floor = None
    if floor_cell:
        fraw = read_cell(floor_cell[0], floor_cell[1])
        if isinstance(fraw, (int, float)) and not isinstance(fraw, bool):
            floor = float(fraw)
    if floor is not None:
        value = max(value, floor)

    wcol, zcol = col_letter(ladder.weight_col), col_letter(col)
    wrange = f"{token}!${wcol}${r0}:${wcol}${r1}"
    zrange = f"{token}!${zcol}${r0}:${zcol}${r1}"
    if weight_cell is None:
        return None, None
    key_expr = (f"CEILING({weight_cell},1)" if round_up
                else f"TRUNC({weight_cell})")
    inner = f"INDEX({zrange},MATCH({key_expr},{wrange},1))"
    if below_first:
        inner = f"IFERROR({inner},{cell_ref(token, r0, col)})"
    if floor_cell:
        inner = f"MAX({inner},{cell_ref(token, floor_cell[0], floor_cell[1])})"
    return f"=ROUND({inner},2)", excel_round(value, 2)


def surcharge_formula(token, ref, read_cell):
    """附加费：**直接指向**报价表里那个价单元格（`='[报价表.xlsx]sheet'!$G$34`）。

    不做 `MATCH(...,0)`：`SurchargeSection.find()` 是去空白、忽略大小写、还带**子串
    兜底**的三段式匹配，`MATCH` 复刻不了——凡是靠子串兜底命中的项，公式一律 `#N/A`。
    引擎已经把确切出处解析出来了，直接指过去既忠实又省事。

    :return: (公式, Excel 语义下的值)；值为 None 表示这条不能写成公式
    """
    cell = ref.target_cell()
    if cell is None or not token:
        return None, None
    raw = read_cell(cell[0], cell[1])
    if raw is None or isinstance(raw, str):
        return None, None
    return f"={cell_ref(token, cell[0], cell[1])}", excel_round(float(raw), 2)


def diff_formula(row, price_col, fee_col):
    """`差异 = 报价表 - 费用`。

    `ROUND(...,2)` 是为了消掉浮点尾巴（Excel 里 `I-G` 常算出 `0.30000000000000004`）。
    注意这跟引擎的 `round(报价原价 - 费用, 2)` **不总是相等**——报价单元格里已经是
    2 位小数了，先舍再加会差一分。所以差异列也要过闸门：对不上就写值。
    """
    return f"=ROUND({col_letter(price_col)}{row}-{col_letter(fee_col)}{row},2)"


def diff_value(price_2dp, fee):
    """差异公式的 Excel 语义值：拿**已舍到 2 位的报价**减费用，再舍一次。"""
    if price_2dp is None or fee is None:
        return None
    return excel_round(float(price_2dp) - float(fee), 2)
