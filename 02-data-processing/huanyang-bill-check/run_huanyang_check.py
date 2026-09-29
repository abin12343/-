# -*- coding: utf-8 -*-
"""环洋核价：命令行入口（可不开弹窗，用于自测/计划任务）。

    python run_huanyang_check.py --bill2 <成本明细.xlsx> --bill1 <成本汇总.xlsx> \
        --quote <报价表.xlsx> [--translation <翻译表.xlsx>] [--out <目录>]

跑完产出：
  - 加工后的账单2（原文件就地回填，原文件先备份成 <名>_原始备份.xlsx）
  - 《环洋核价报告_<时间戳>.txt》
  - 《环洋天图核验清单_<时间戳>.csv》—— 交给 永达 check_tiantu.py 消费

天图核验是独立一步（要开浏览器登录），由 main_flow.py / selftest_tiantu.py 驱动。
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from datetime import datetime
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from process_details import (  # noqa: E402
    build_checklist,
    build_translation_map,
    load_rules,
    process,
)
from quote_engine import QuoteBook, load_summary_index  # noqa: E402

HERE = Path(__file__).resolve().parent

if getattr(sys, "frozen", False):
    HERE = Path(sys.executable).resolve().parent.parent
    # 冻结后 PYTHONUTF8 不生效，管道默认 cp936；调用方按 utf-8 解码我们的输出
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def load_config(path=None) -> dict:
    cfg_path = Path(path) if path else HERE / "config.json"
    cfg = json.loads(cfg_path.read_text(encoding="utf-8")) if cfg_path.exists() else {}
    local = cfg_path.with_name(cfg_path.stem + ".local" + cfg_path.suffix)
    if local.exists():
        cfg = _deep_merge(cfg, json.loads(local.read_text(encoding="utf-8")))
    return cfg


def _deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = _deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def _log_factory(quiet=False):
    def log(msg):
        if not quiet:
            print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)

    return log


def pick_files(cfg, args):
    """定位 账单1/账单2/报价表。命令行优先，其次 config.paths。"""
    paths = cfg.get("paths") or {}
    bills = cfg.get("bills") or {}
    out = {}

    if args.bill2:
        out["bill2"] = Path(args.bill2)
    else:
        out["bill2"] = None

    if args.bill1:
        out["bill1"] = Path(args.bill1)
    else:
        out["bill1"] = None

    # 没给账单时，从 input_dir 里按文件名猜：成本明细=账单2，其余 xlsx=账单1
    if not out["bill2"] or not out["bill1"]:
        d = Path(paths.get("input_dir") or ".")
        detail = bills.get("detail_match", "成本明细")
        summary = bills.get("summary_match", "成本汇总")
        cands = [p for p in d.glob("*.xlsx") if not p.name.startswith("~$")
                 and "原始备份" not in p.name and "已生成" not in p.name]
        for p in cands:
            if not out["bill2"] and detail in p.name:
                out["bill2"] = p
            elif not out["bill1"] and summary not in p.name and "价格表" not in p.name \
                    and "翻译" not in p.name and "统计表" not in p.name:
                out["bill1"] = p

    out["quote"] = Path(args.quote) if args.quote else (
        Path(paths["quote_file"]) if paths.get("quote_file") else None
    )
    # 发布包在另一台电脑时 config 里的开发机绝对路径会失效；优先找账单同目录报价表。
    if out.get("quote") and not out["quote"].is_file() and out.get("bill2"):
        qdir = Path(out["bill2"]).parent
        candidates = [p for p in qdir.glob("*.xlsx")
                      if "价格" in p.name or "报价" in p.name]
        if candidates:
            out["quote"] = sorted(candidates)[0]
    tr = Path(args.translation) if args.translation else None
    if tr is None and out.get("bill2"):
        sibling = Path(out["bill2"]).parent / "打单费用名称中英文翻译.xlsx"
        if sibling.is_file():
            tr = sibling
    if tr is None and paths.get("translation_file"):
        tr = Path(paths["translation_file"])
    out["translation"] = tr
    out["stats"] = Path(args.stats) if args.stats else (
        Path(paths["stats_file"]) if paths.get("stats_file") else None
    )
    if out.get("stats") and not out["stats"].is_file() and out.get("bill2"):
        s = Path(out["bill2"]).parent / "打单问题统计表（模板）.xlsx"
        if s.is_file():
            out["stats"] = s
    return out


def write_checklist_csv(tasks, out_dir: Path) -> Path:
    """写《环洋天图核验清单》。字段与 永达 check_tiantu.py 的契约保持一致。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"环洋_天图核验清单_{ts}.csv"
    fields = ["Excel行号", "客户单号(去后缀)", "费用类别", "费用名称",
              "期望标记", "金额USD", "说明"]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for t in tasks:
            w.writerow({
                "Excel行号": t.get("row", ""),
                "客户单号(去后缀)": t["no"],
                "费用类别": t.get("cat", ""),
                "费用名称": t.get("fee", ""),
                "期望标记": t["mark"],
                "金额USD": t.get("amount", ""),
                "说明": t.get("note", ""),
            })
    return path


def write_report(stat, cfg, out_dir: Path, checklist: Path | None,
                 quote=None, log=print) -> Path:
    """写人看的核价报告。"""
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = out_dir / f"环洋核价报告_{ts}.txt"
    L = []
    add = L.append

    add("=" * 78)
    add(f"环洋核价报告  {datetime.now():%Y-%m-%d %H:%M:%S}")
    add("=" * 78)
    add(f"账单2（成本明细）：{stat.get('保存', '')}")
    add(f"原始备份          ：{stat.get('备份', '') or '（未备份）'}")
    if checklist:
        add(f"天图核验清单      ：{checklist}")
    add("")
    add("-" * 78)
    add("一、加工统计")
    add("-" * 78)
    add(f"  总行数          {stat['总行数']}")
    add(f"  填出运单号      {stat['已填运单号']}（缺失 {stat['运单号缺失']}）")
    add(f"  已翻译          {stat['已翻译']}（无翻译 {stat['无翻译']}）")
    add(f"  已填报价        {stat['已报价']}")
    add(f"  未匹配到报价    {stat['未匹配报价']}")
    add(f"  零额行跳过      {stat['零额跳过']}")
    add(f"  差异超容差      {stat['差异超容差']}（容差 {cfg.get('run', {}).get('amount_tolerance')}）")
    if stat.get("公式开关"):
        add(f"  公式列         翻译 {stat.get('公式_翻译', 0)} / 报价 {stat.get('公式_报价', 0)}"
            f" / 差异 {stat.get('公式_差异', 0)}"
            f"（退回写值 {stat.get('公式_降级', 0)}）")
    else:
        add("  公式列         （关：run.write_formulas=false，三列都写值）")

    if stat["差异样例"]:
        add("")
        add("-" * 78)
        add("二、差异明细（|报价表-费用| 超容差，按出现顺序取前 12 条）")
        add("-" * 78)
        add(f"  {'运单号':<20}{'费用类别':<20}{'产品名称':<26}{'重':>7}{'区':>4}{'账单':>9}{'报价':>9}{'差异':>9}")
        for no, zh, prod, wt, z, fee, qp, diff in stat["差异样例"]:
            add(f"  {str(no):<20}{str(zh):<20}{str(prod)[:24]:<26}"
                f"{wt if wt is not None else '':>7}{z if z is not None else '':>4}"
                f"{fee:>9.2f}{qp:>9.2f}{diff:>9.2f}")

    if stat["未匹配明细"]:
        add("")
        add("-" * 78)
        add("三、未匹配到报价的原因（Top 20）")
        add("-" * 78)
        for k, n in sorted(stat["未匹配明细"].items(), key=lambda kv: -kv[1])[:20]:
            add(f"  {n:>5} 行  {k}")

    if stat["未翻译费用名"]:
        add("")
        add("-" * 78)
        add("四、翻译表里没有的费用名（这些行不会填报价表）")
        add("-" * 78)
        for k, n in stat["未翻译费用名"]:
            add(f"  {n:>5} 行  {k}")

    if stat.get("降级明细"):
        add("")
        add("-" * 78)
        add("七、本来要写公式、验算没通过而退回写值的原因")
        add("-" * 78)
        add("  （公式写下去之前都按 Excel 的规矩验算过：算出来的数与引擎的数相等才写。")
        add("    这里列出的行写的是值，不影响报价结论，只是以后改计费重不会自动跟着变。）")
        for k, n in stat["降级明细"]:
            add(f"  {n:>5} 行  {k}")

    if quote is not None and quote.missing_products:
        add("")
        add("-" * 78)
        add("五、报价表里找不到对应 sheet 的产品（需人工在 config.quote.product_aliases 指定）")
        add("-" * 78)
        for p in sorted(quote.missing_products):
            add(f"  {p}")

    if quote is not None:
        add("")
        add("-" * 78)
        add("六、报价表解析结果（自检用：阶梯档位/区块识别是否正常）")
        add("-" * 78)
        for sh in quote._sheet_index():
            add(quote.describe(sh))

    text = "\n".join(L) + "\n"
    path.write_text(text, encoding="utf-8")
    return path


def main(argv=None):
    ap = argparse.ArgumentParser(description="环洋账单加工与核价")
    ap.add_argument("--bill2", help="账单2《成本明细》xlsx（加工对象）")
    ap.add_argument("--bill1", help="账单1《成本汇总》xlsx（提供 主单号→客户单号）")
    ap.add_argument("--quote", help="报价表 xlsx")
    ap.add_argument("--translation", help="《打单费用名称中英文翻译》xlsx")
    ap.add_argument("--stats", help="《打单问题统计表》xlsx（本脚本只记路径，不写入）")
    ap.add_argument("--out", help="输出目录（报告与清单）")
    ap.add_argument("--config", help="配置文件路径")
    ap.add_argument("--no-backup", action="store_true", help="不备份原账单")
    ap.add_argument("--dry-run", action="store_true", help="只生成清单与报告，不回填账单")
    ap.add_argument("--quiet", action="store_true")
    args = ap.parse_args(argv)

    log = _log_factory(args.quiet)
    cfg = load_config(args.config)
    f = pick_files(cfg, args)
    out_dir = Path(args.out) if args.out else Path(
        (cfg.get("paths") or {}).get("output_dir") or (HERE / "outputs")
    )

    missing = [k for k in ("bill2", "bill1", "quote") if not f.get(k)]
    if missing:
        log(f"缺少必需输入：{missing}。用 --bill2/--bill1/--quote 指定，或在 config.paths 里配好。")
        return 2
    for k in ("bill2", "bill1", "quote"):
        if not f[k].is_file():
            log(f"{k} 不存在：{f[k]}")
            return 2

    log(f"账单2（成本明细）：{f['bill2'].name}")
    log(f"账单1（成本汇总）：{f['bill1'].name}")
    log(f"报价表            ：{f['quote'].name}")

    # 1) 账单1 → 主单号→客户单号
    summary_sheet = (cfg.get("bills") or {}).get("summary_sheet", "成本汇总")
    summary = load_summary_index(f["bill1"], summary_sheet, log=log)
    log(f"[账单1] 读到 {len(summary)} 个主单号")

    # 2) 翻译表
    tr_cfg = cfg.get("translation") or {}
    if f.get("translation") and f["translation"].is_file():
        translation = build_translation_map(
            f["translation"], tr_cfg.get("sheet", "Sheet1"),
            overrides=tr_cfg.get("overrides"), log=log,
        )
    else:
        translation = {}
        log("[警告] 没配翻译表，翻译列会全空")

    # 3) 报价表
    quote = QuoteBook(
        f["quote"],
        product_aliases=(cfg.get("quote") or {}).get("product_aliases"),
        hwt_min_weight=(cfg.get("quote") or {}).get("hwt_min_weight", 200),
    )

    # 4) 加工
    if args.dry_run:
        log("[dry-run] 跳过回填，仅统计")
        saved, stat, rows = _dry_run(f["bill2"], summary, quote, translation, cfg, log,
                                     translation_file=f.get("translation"))
    else:
        saved, stat, rows = process(
            f["bill2"], summary, quote, translation, cfg, log=log,
            backup=not args.no_backup,
            translation_file=f.get("translation"),
        )

    # 5) 天图核验清单
    tasks = build_checklist(rows, cfg, translation, load_rules(cfg), log=log)
    checklist = write_checklist_csv(tasks, out_dir) if tasks else None
    if checklist:
        log(f"[清单] {checklist}")

    report = write_report(stat, cfg, out_dir, checklist, quote=quote, log=log)
    log(f"[报告] {report}")
    log(f"完成：回填报价 {stat['已报价']} 行，未匹配 {stat['未匹配报价']} 行，"
        f"差异超容差 {stat['差异超容差']} 行")
    return 0


def _dry_run(bill2_path, summary, quote, translation, cfg, log, translation_file=None):
    """不写盘的试跑：在临时副本上把加工逻辑跑一遍，只为拿统计和清单。"""
    import shutil
    import tempfile

    tmp = Path(tempfile.mkdtemp(prefix="hy_dry_"))
    copy = tmp / Path(bill2_path).name
    shutil.copy2(bill2_path, copy)
    # ref_dir 必须给**原账单所在目录**：副本躺在临时目录里，报价表/翻译表都不在它旁边，
    # 拿副本目录当基准的话每条公式都会「不在同目录」而退回写值，试跑就白跑了
    saved, stat, rows = process(copy, summary, quote, translation, cfg, log=log,
                                backup=False, ref_dir=Path(bill2_path).parent,
                                translation_file=translation_file)
    stat["保存"] = f"{bill2_path}（dry-run，未改动原文件）"
    stat["备份"] = ""
    return saved, stat, rows


if __name__ == "__main__":
    sys.exit(main())
