# -*- coding: utf-8 -*-
"""天图核验自测：确认「天图检测」这一段是可信的，再看它有多慢。

分两层，第一层不联网、秒出，第二层才开浏览器：

  A. 契约自检（离线）
     - 引擎文件在不在、清单字段全不全；
     - 环洋用到的每个【期望标记】引擎是否认识。引擎只对
       `住宅私人`（逐单查应收）、`address`（逐单查运踪）、
       `偏远/超偏远/超长/超重`（红色标记）有专门分支，写错一个字就会
       被当成"整段文本里找这个词"，恒判「否」——这种错从报告上看不出来，
       只会表现为"这单天图确实没有"，所以必须在跑之前挡掉。
     - 引擎根目录下有没有 credentials.local.json（引擎只认那一份）。

  B. 双跑一致性（联网）
     抽 limit 个单号，用同一份清单**连跑两遍**，按 (单号, 期望标记) 逐条比对。
     天图的应收格子是异步填的，历史上出现过「读一次读空就当否」导致
     同一批单号两次结果不一致（有 是→否 翻转）。翻转数就是这段逻辑的
     可信度指标：翻转 0 条才说明判定稳定，而不是运气好。

跑法：
    python selftest_tiantu.py --checklist outputs/环洋_天图核验清单_xxx.csv
    python selftest_tiantu.py            # 用 outputs 里最新的那份清单
结果写成《环洋_天图自测报告_<时间戳>.txt》。
"""

from __future__ import annotations

import csv
import json
import logging
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

# 引擎对这些标记有专门分支，其余标记会退化成"整段文本里找这个词"
_SPECIAL_MARKS = ("住宅私人", "address")
_RED_MARKS = ("偏远", "超偏远", "超长", "超重")
KNOWN_MARKS = _SPECIAL_MARKS + _RED_MARKS

_CHECKLIST_FIELDS = ["Excel行号", "客户单号(去后缀)", "费用类别", "费用名称",
                     "期望标记", "金额USD", "说明"]


def _pad(text, width: int) -> str:
    """按**显示宽度**补齐，不是按字符数。

    报告是给人看的：「偏远」是 2 个字符但占 4 格，用 `f"{m:<10}"` 补出来的列
    在等宽字体里是歪的（中文标记全挤在一起）。东亚宽/全角算 2 格，其余算 1 格。
    """
    s = str(text or "")
    used = sum(2 if unicodedata.east_asian_width(ch) in ("W", "F") else 1 for ch in s)
    return s + " " * max(width - used, 0)


class _Log:
    """包一层日志。

    调用方传进来的可能是 logging.Logger（main_flow 那条路），也可能是 print 式的
    可调用对象（单跑本模块调试时最省事）——两种都得能吃。
    """

    def __init__(self, log=None):
        self._logger = log if isinstance(log, logging.Logger) else None
        self._fn = None if self._logger else log

    def _emit(self, level, msg, args):
        if self._logger is not None:
            getattr(self._logger, level)(msg, *args)
        elif self._fn is not None:
            self._fn((msg % args) if args else msg)

    def info(self, msg, *args): self._emit("info", msg, args)
    def warning(self, msg, *args): self._emit("warning", msg, args)
    def error(self, msg, *args): self._emit("error", msg, args)


def _fields(path: Path) -> list[str]:
    with path.open(encoding="utf-8-sig", newline="") as f:
        return next(csv.reader(f), [])


def _read_tasks(path: Path) -> list[dict]:
    with path.open(encoding="utf-8-sig", newline="") as f:
        return [r for r in csv.DictReader(f) if (r.get("客户单号(去后缀)") or "").strip()]


def check_contract(cfg, engine: Path, base: Path, checklist: Path, log) -> list[str]:
    """离线自检，返回问题列表（空 = 通过）。

    :param base: 引擎认的根目录（见 main_flow._engine_base）——凭证就放在那里。
    """
    problems = []

    if not engine.is_file():
        problems.append(f"引擎文件不存在：{engine}（config.tiantu.engine）")
    else:
        log.info(f"[自检] 引擎：{engine}")
        cred = Path(base) / "credentials.local.json"
        if cred.is_file():
            try:
                d = json.loads(cred.read_text(encoding="utf-8"))
                user = d.get("tiantu_user") or d.get("user") or ""
                log.info(f"[自检] 凭证：{cred}（账号 {user or '未填'}）")
                if not user:
                    problems.append(f"{cred} 里没有 tiantu_user")
                if not (d.get("tiantu_pass") or d.get("pass")):
                    problems.append(f"{cred} 里没有 tiantu_pass")
            except Exception as exc:  # noqa: BLE001
                problems.append(f"{cred} 读不了：{exc}")
        else:
            problems.append(
                f"引擎根目录下没有 credentials.local.json：{cred}。"
                "引擎只认这一份（发布包放在发布根，源码态在引擎脚本旁），"
                "config 里的账号是打单系统的、登不上天图。"
            )

    if not checklist.is_file():
        problems.append(f"核验清单不存在：{checklist}")
        return problems
    fields = _fields(checklist)
    missing = [h for h in ("客户单号(去后缀)", "费用名称", "期望标记") if h not in fields]
    if missing:
        problems.append(f"清单缺字段 {missing}（现有：{fields}），引擎会读不到任务")

    tasks = _read_tasks(checklist)
    log.info(f"[自检] 清单：{checklist.name}，任务 {len(tasks)} 条，"
        f"涉及单号 {len({t['客户单号(去后缀)'].strip() for t in tasks})} 个")
    if not tasks:
        problems.append("清单里一条任务都没有")

    marks = sorted({(t.get("期望标记") or "").strip() for t in tasks} - {""})
    bad = [m for m in marks if m not in KNOWN_MARKS]
    log.info(f"[自检] 期望标记：{'、'.join(marks) or '（无）'}")
    if bad:
        problems.append(
            f"清单里有引擎不认识的标记 {bad}。引擎只对 {list(KNOWN_MARKS)} 有专门分支，"
            "其余会退化成纯文本查找、恒判「否」。"
        )
    n_slow = sum(1 for t in tasks
                 if (t.get("期望标记") or "").strip() in _SPECIAL_MARKS)
    if n_slow:
        log.info(f"[自检] 其中 {n_slow} 条走逐单核验（每单 20-30 秒），是整段最慢的部分")
    return problems


def pick_sample(tasks: list[dict], limit: int, prefer_marks=()) -> list[dict]:
    """抽 limit 条做样本：按标记分组后轮流取，保证各类标记都被覆盖到。

    只从一类标记里取的话，自测通过也说明不了别的标记是稳的。
    """
    groups: dict[str, list[dict]] = {}
    seen = set()
    for t in tasks:
        no = (t.get("客户单号(去后缀)") or "").strip()
        mark = (t.get("期望标记") or "").strip()
        if not no or not mark or (no, mark) in seen:
            continue
        seen.add((no, mark))
        groups.setdefault(mark, []).append(t)

    order = [m for m in prefer_marks if m in groups]
    order += [m for m in sorted(groups, key=lambda k: -len(groups[k])) if m not in order]

    out, ptr = [], 0
    while len(out) < limit and any(ptr < len(groups[m]) for m in order):
        for m in order:
            if len(out) >= limit:
                break
            if ptr < len(groups[m]):
                out.append(groups[m][ptr])
        ptr += 1
    return out


def write_sample_csv(sample: list[dict], path: Path) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=_CHECKLIST_FIELDS, extrasaction="ignore")
        w.writeheader()
        w.writerows(sample)
    return path


def compare(rows_a: list[dict], rows_b: list[dict]) -> dict:
    """按 (单号, 期望标记) 比对两遍结果。"""
    def idx(rows):
        out = {}
        for r in rows:
            key = ((r.get("单号") or "").strip(), (r.get("期望标记") or "").strip())
            if all(key):
                out[key] = r
        return out

    a, b = idx(rows_a), idx(rows_b)
    same, flips, only_a, only_b = [], [], [], []
    detail = []
    for key in sorted(set(a) | set(b)):
        ra, rb = a.get(key), b.get(key)
        if ra is None:
            only_b.append(key)
        elif rb is None:
            only_a.append(key)
        else:
            va = (ra.get("是否出现") or "").strip()
            vb = (rb.get("是否出现") or "").strip()
            rec = {
                "no": key[0], "mark": key[1], "a": va, "b": vb,
                "na": ra.get("命中条数", ""), "nb": rb.get("命中条数", ""),
                "fee": (ra.get("费用名称") or "").strip(),
            }
            if va == vb:
                same.append(rec)
            else:
                flips.append(rec)
            detail.append(rec)
    return {"same": same, "flips": flips, "only_a": only_a, "only_b": only_b,
            "detail": detail, "total": len(set(a) | set(b))}


def run_selftest(cfg, checklist: Path, out_dir: Path, log, limit=None) -> Path:
    """执行自测并写报告。返回报告路径。"""
    import main_flow as MF  # noqa: PLC0415  延迟导入：main_flow 也 import 本模块

    log = _Log(log)
    st = cfg.get("selftest") or {}
    limit = int(limit or st.get("limit") or 6)
    engine, _head = MF._engine_cmd(cfg)
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    L: list[str] = []
    add = L.append

    add("=" * 78)
    add(f"环洋天图核验自测报告  {datetime.now():%Y-%m-%d %H:%M:%S}")
    add("=" * 78)
    add("")

    # ---- A. 契约自检 ----
    add("-" * 78)
    add("A. 契约自检（离线）")
    add("-" * 78)
    problems = check_contract(cfg, engine, MF._engine_base(engine), Path(checklist), log)
    if problems:
        for p in problems:
            add(f"  [不通过] {p}")
        add("")
        add("契约自检未通过，后续双跑没有意义，已停止。")
        path = out_dir / f"环洋_天图自测报告_{ts}.txt"
        path.write_text("\n".join(L) + "\n", encoding="utf-8")
        for p in problems:
            log.error("[自检] 不通过：%s", p)
        log.error("[自检] 报告：%s", path)
        return path
    add("  [通过] 引擎、凭证、清单字段、期望标记均正常")
    add("")

    # ---- 抽样 ----
    tasks = _read_tasks(Path(checklist))
    sample = pick_sample(tasks, limit, st.get("marks") or [])
    if not sample:
        add("清单里没有可用样本（单号/标记为空），自测结束。")
        path = out_dir / f"环洋_天图自测报告_{ts}.txt"
        path.write_text("\n".join(L) + "\n", encoding="utf-8")
        log.warning("[自检] 没有可用样本")
        return path

    sample_csv = write_sample_csv(sample, out_dir / f"环洋_天图自测样本_{ts}.csv")
    add("-" * 78)
    add(f"B. 双跑一致性（样本 {len(sample)} 条，取自 {Path(checklist).name}）")
    add("-" * 78)
    add(f"  样本清单：{sample_csv.name}")
    from collections import Counter

    add("  样本构成：" + "、".join(
        f"{m} {n}" for m, n in Counter(
            (t.get("期望标记") or "").strip() for t in sample
        ).items()
    ))
    for t in sample:
        add("    " + _pad(t.get("客户单号(去后缀)", ""), 24)
            + _pad(t.get("期望标记") or "", 12) + (t.get("费用类别") or ""))
    add("")
    log.info("[自检] 样本 %d 条：%s", len(sample),
             "、".join(f"{t['客户单号(去后缀)']}/{t['期望标记']}" for t in sample))

    # ---- 连跑两遍 ----
    results: list[tuple[Path, str]] = []
    timing: list[str] = []
    for i in (1, 2):
        log.info("[自检] 第 %d 遍开始（天图要开浏览器登录，请勿打扰）", i)
        t0 = time.time()
        res = MF.run_tiantu_check(cfg, sample_csv, out_dir, 0, log)
        secs = round(time.time() - t0, 2)
        if not res:
            add(f"  第 {i} 遍没有拿到结果，自测中止（看上面的 [天图] 日志）。")
            path = out_dir / f"环洋_天图自测报告_{ts}.txt"
            path.write_text("\n".join(L) + "\n", encoding="utf-8")
            log.error("[自检] 第 %d 遍失败，报告：%s", i, path)
            return path
        rows = list(csv.DictReader(res.open(encoding="utf-8-sig")))
        results.append((res, rows))
        # 引擎把运行日志写在自己的输出目录里（跟结果 CSV 同处），这里回捞一份，
        # 自测报告才能顺带回答「这么点单要跑多久、慢在哪一类查询」。
        run_log = _latest_run_log(out_dir, started_after=t0 - 1)
        timing.append(MF.summarize_tiantu_log(run_log, secs))
        log.info("[自检] 第 %d 遍完成：%s（%d 行，%ss）", i, res.name, len(rows), secs)

    (csv_a, rows_a), (csv_b, rows_b) = results
    cmp = compare(rows_a, rows_b)
    add(f"  第 1 遍：{csv_a.name}（{len(rows_a)} 行）")
    add(f"  第 2 遍：{csv_b.name}（{len(rows_b)} 行）")
    add("")
    add("  两遍耗时与查询次数（天图慢就慢在查询次数上）")
    for i, s in enumerate(timing, 1):
        add(f"    第 {i} 遍：{s}")
    add("")
    add(f"  逐条比对（按 单号+期望标记，共 {cmp['total']} 条）")
    add(f"    一致          {len(cmp['same'])}")
    add(f"    翻转          {len(cmp['flips'])}")
    add(f"    只第一遍有    {len(cmp['only_a'])}")
    add(f"    只第二遍有    {len(cmp['only_b'])}")

    by_mark = {}
    for r in cmp["detail"]:
        d = by_mark.setdefault(r["mark"], [0, 0])
        d[0] += 1
        if r["a"] != r["b"]:
            d[1] += 1
    add("")
    add("  按标记拆分（条数 / 翻转）")
    for m in sorted(by_mark):
        n, f = by_mark[m]
        add("    " + _pad(m, 14) + f"{n:>5} / {f}")

    if cmp["flips"]:
        add("")
        add("  ▸ 翻转明细（同一单号同一标记，两遍结论不同）")
        add("    " + _pad("单号", 24) + _pad("标记", 12)
            + _pad("第一遍", 14) + _pad("第二遍", 14) + "命中(1/2)")
        for r in cmp["flips"]:
            add("    " + _pad(r["no"], 24) + _pad(r["mark"], 12)
                + _pad(r["a"], 14) + _pad(r["b"], 14)
                + f"{r['na']}/{r['nb']}")
        add("")
        add("  翻转不为 0 说明判定不稳：应收/运踪格子是异步填的，别把「否」直接当结论，")
        add("  这一批要么重跑要么人工看。")
    else:
        add("")
        add("  ▸ 翻转 0 条：两次结论完全一致，判定稳定。")

    add("")
    add("  备注：本自测只证明「同一份清单两次跑结论一致」，不代表结论本身正确。")
    add("        真实有/无仍需人工抽验（历史上翻转 0 的批次里也可能存在系统性误判）。")

    path = out_dir / f"环洋_天图自测报告_{ts}.txt"
    path.write_text("\n".join(L) + "\n", encoding="utf-8")
    log.info("[自检] 一致 %d / 翻转 %d（样本 %d 条），报告：%s",
             len(cmp["same"]), len(cmp["flips"]), cmp["total"], path)
    if cmp["flips"]:
        log.warning("[自检] 有翻转，天图判定不稳定，请人工复核上面列出的单号")
    return path


def _latest_run_log(out_dir: Path, started_after: float) -> Path | None:
    """取本轮引擎运行日志（`环洋_天图运行_*.log`）。

    必须卡时间：同一目录里还躺着上一次跑的单遍日志和上一轮自测的日志，
    挑最新且**晚于 i 遍开始时刻**的那一份，才不会把上一遍的查询次数算到这一遍头上。
    """
    if not out_dir.is_dir():
        return None
    cands = [f for f in out_dir.glob("*_天图运行_*.log")
             if f.stat().st_mtime >= started_after]
    return max(cands, key=lambda f: f.stat().st_mtime) if cands else None


def _latest_checklist(out_dir: Path) -> Path | None:
    if not out_dir.is_dir():
        return None
    cands = sorted(out_dir.glob("*_天图核验清单_*.csv"),
                   key=lambda f: f.stat().st_mtime, reverse=True)
    return cands[0] if cands else None


def main(argv=None) -> int:
    import argparse  # noqa: PLC0415

    import main_flow as MF  # noqa: PLC0415

    ap = argparse.ArgumentParser(description="环洋天图核验自测")
    ap.add_argument("--checklist", help="核验清单 csv（默认取 outputs 里最新的）")
    ap.add_argument("--limit", type=int, default=0, help="样本条数（默认 config.selftest.limit）")
    args = ap.parse_args(argv)

    cfg = MF._load_cfg()
    log = MF._setup_logger(cfg)
    out_dir = MF._out_dir(cfg)
    checklist = Path(args.checklist) if args.checklist else _latest_checklist(out_dir)
    if not checklist:
        log.error("找不到核验清单，先用 run_huanyang_check.py 或 main_flow.py 跑一次核价。")
        return 2
    run_selftest(cfg, checklist, out_dir, log, limit=args.limit or None)
    return 0


if __name__ == "__main__":
    sys.exit(main())
