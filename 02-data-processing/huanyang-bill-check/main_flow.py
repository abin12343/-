# -*- coding: utf-8 -*-
"""环洋打单数据整理 — 总流程。

  1) 弹窗选**账单1（成本汇总）** —— 提供 主单号 → 客户单号
  2) 弹窗选**账单2（成本明细）** —— 加工对象（运单号/翻译/报价表/差异 都写在这张）
     （SOP 要求两个账单各弹一次：账单1 是"总表"，账单2 才是要加工的那张）
  3) 弹窗选**报价表 xlsx** —— 报价表时常更换，每次都要选
  4) 加工 + 核价（本进程内直接调 process_details，不另起 exe）
  5) 天图核验（复用 永达 check_tiantu.py，跑完统计查询次数与耗时）
  6) 把「天图未查到标记」登记进《打单问题统计表》的「环洋」sheet
  7) 通知（失败不抛）

运行：python main_flow.py [--bill1 x.xlsx] [--bill2 y.xlsx] [--quote z.xlsx]
                        [--stats s.xlsx] [--no-dialog] [--dry-run]
                        [--tiantu-limit N] [--selftest]
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

if getattr(sys, "frozen", False):
    HERE = Path(sys.executable).resolve().parent.parent
    # 冻结后 PYTHONUTF8 不生效，管道默认 cp936 写中文；父流程按 utf-8 解码并靠中文
    # 措辞数查询次数，不强制 utf-8 会整段乱码、次数数成 0。
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
else:
    HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

PY = sys.executable
_EXCEL_TYPES = [("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")]


# ============== 基础工具 ==============

def _safe_print(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        sys.stdout.buffer.write((str(msg) + "\n").encode(enc, "replace"))
        sys.stdout.flush()


def _load_cfg() -> dict:
    p = HERE / "config.json"
    cfg = json.loads(p.read_text(encoding="utf-8")) if p.is_file() else {}
    local = HERE / "config.local.json"
    if local.is_file():
        cfg = _deep_merge(cfg, json.loads(local.read_text(encoding="utf-8")))
    return cfg


def _deep_merge(base: dict, over: dict) -> dict:
    out = dict(base)
    for k, v in (over or {}).items():
        out[k] = _deep_merge(out[k], v) if isinstance(v, dict) and isinstance(out.get(k), dict) else v
    return out


def _load_notify_local(cfg: dict) -> dict:
    """本机 notify.local.json 覆盖通知字段（不入库）。"""
    p = HERE / "notify.local.json"
    if not p.is_file():
        return cfg
    try:
        local = json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return cfg
    nt = cfg.setdefault("notify", {})
    for k in ("wecom_key", "app_name", "account_id"):
        if local.get(k):
            nt[k] = local[k]
    if local.get("kdocs"):
        nt["kdocs"] = {**nt.get("kdocs", {}), **local["kdocs"]}
    return cfg


def _resolve_path(raw, fallback: Path | None = None) -> Path | None:
    """解析配置路径；发布包内相对路径优先，相邻文件作为跨电脑兜底。"""
    if not raw:
        return fallback
    p = Path(str(raw)).expanduser()
    if p.is_file() or p.is_dir():
        return p
    if not p.is_absolute():
        q = HERE / p
        if q.exists():
            return q
    return fallback


def _resolve_translation(cfg: dict, bill2: str, explicit: str = "") -> str | None:
    """翻译表优先级：命令行 → 账单同目录 → 发布根 → config 路径。"""
    candidates = []
    if explicit:
        candidates.append(Path(explicit))
    if bill2:
        d = Path(bill2).resolve().parent
        candidates.extend([d / "打单费用名称中英文翻译.xlsx", d / "翻译表.xlsx"])
    candidates.extend([HERE / "打单费用名称中英文翻译.xlsx", HERE / "翻译表.xlsx"])
    raw = (cfg.get("paths") or {}).get("translation_file")
    if raw:
        candidates.append(Path(raw))
    for p in candidates:
        if p.is_file():
            return str(p)
    return None


def _out_dir(cfg: dict) -> Path:
    try:
        p = Path(cfg["paths"]["output_dir"])
        return p if p.is_absolute() else HERE / p
    except Exception:  # noqa: BLE001
        return HERE / "outputs"


def _log_dir(cfg: dict) -> Path:
    try:
        p = Path(cfg["paths"]["log_dir"])
        return p if p.is_absolute() else HERE / p
    except Exception:  # noqa: BLE001
        return HERE / "logs"


def _setup_logger(cfg: dict) -> logging.Logger:
    log = logging.getLogger("huanyang")
    log.setLevel(logging.INFO)
    if log.handlers:
        return log
    d = _log_dir(cfg)
    try:
        d.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(d / "huanyang.log", encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        log.addHandler(fh)
    except Exception as exc:  # noqa: BLE001
        _safe_print(f"[警告] 无法写日志文件（{exc}）")
    ch = logging.StreamHandler(sys.stdout)
    ch.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%H:%M:%S"))
    log.addHandler(ch)
    return log


def _autolib_path() -> None:
    """开发态把 08-common-utils 挂到 sys.path；打包后 autolib 已进 exe，不用管。"""
    if getattr(sys, "frozen", False):
        return
    sys.path.insert(0, str(HERE.parents[1] / "08-common-utils"))


def _autolib():
    _autolib_path()
    import autolib.dialog as dialog  # noqa: PLC0415
    import autolib.excel as excel  # noqa: PLC0415
    return dialog, excel


# ============== 弹窗 ==============

def _pick(title: str, initialdir: str, validator=None, max_tries=10) -> str | None:
    """弹窗选一个 xlsx，选错重来。无 GUI 环境返回 None。"""
    dialog, _ = _autolib()
    return dialog.pick_until_valid(title, _EXCEL_TYPES, initialdir, validator, max_tries)


def _squash(value) -> str:
    """去掉所有空白再比对：表头里「Zone 2」可能写成「Zone  2」/换行，直接比会漏。"""
    return re.sub(r"\s+", "", str(value)) if value is not None else ""


def _file_err(path: str, need_all=(), need_any=(), scan_sheets=25, scan_rows=6) -> str | None:
    """选错文件时给一句人看得懂的理由，返回 None 表示放行。

    need_all 允许命中 **sheet 名**：账单1 整本只有一个「成本汇总」页，页名本身就是凭据。
    need_any 先看第一个 sheet 的表头，不中再翻前 scan_sheets 页的前 scan_rows 行 ——
    报价表第一页是「产品介绍」说明页，价目表在后面二十来页里，只看第一页会把
    **正确的报价表拦下来**，用户就一直卡在重选（这里踩过）。
    """
    p = Path(path or "")
    if not p.is_file():
        return "文件不存在，请重新选择。"
    import openpyxl

    try:
        wb = openpyxl.load_workbook(p, data_only=True)
    except Exception as exc:  # noqa: BLE001
        return f"打不开该 Excel：{exc}"

    try:
        names = wb.sheetnames
        header = "".join(_squash(c.value)
                         for c in next(wb[names[0]].iter_rows(min_row=1, max_row=1)))
        name_blob = "".join(names)

        miss = [h for h in need_all if _squash(h) not in header and _squash(h) not in name_blob]
        if miss:
            return f"看起来不是目标文件：找不到 {('/'.join(miss))}。"

        if need_any:
            want = [_squash(h) for h in need_any]
            hit = any(w in header for w in want) or any(w in name_blob for w in want)
            if not hit:
                for sn in names[:scan_sheets]:
                    ws = wb[sn]
                    blob = "".join(_squash(v)
                                   for row in ws.iter_rows(min_row=1, max_row=scan_rows,
                                                           values_only=True)
                                   for v in row)
                    if any(w in blob for w in want):
                        hit = True
                        break
            if not hit:
                return f"看起来不是目标文件：里面没有 {('/'.join(need_any))} 任一内容。"
    finally:
        wb.close()
    return None


def choose_summary_via_dialog(cfg) -> str | None:
    """账单1 —— 《成本汇总》（提供 主单号→客户单号）。"""
    b = cfg.get("bills") or {}
    initial = ""
    try:
        d = Path(cfg["paths"].get("input_dir") or "")
        initial = str(d) if d.is_dir() else ""
    except Exception:  # noqa: BLE001
        pass
    sheet = b.get("summary_sheet", "成本汇总")
    return _pick(
        f"① 请选择【账单1 = {sheet}】xlsx（提供 主单号→客户单号）",
        initial,
        lambda p: _file_err(p, need_all=("主单号",), need_any=("客户单号",)),
    )


def choose_detail_via_dialog(cfg) -> str | None:
    """账单2 —— 《成本明细》（加工对象）。"""
    b = cfg.get("bills") or {}
    initial = ""
    try:
        d = Path(cfg["paths"].get("input_dir") or "")
        initial = str(d) if d.is_dir() else ""
    except Exception:  # noqa: BLE001
        pass
    need = tuple(b.get("detail_headers") or ["费用名称", "费用"])
    return _pick(
        f"② 请选择【账单2 = {b.get('detail_match', '成本明细')}】xlsx（加工对象）",
        initial,
        lambda p: _file_err(p, need_all=need),
    )


def choose_quote_via_dialog(cfg) -> str | None:
    """报价表 —— 时常更换，每次都要选。"""
    initial = ""
    q = (cfg.get("paths") or {}).get("quote_file")
    if q and Path(q).parent.is_dir():
        initial = str(Path(q).parent)
    return _pick(
        "③ 请选择【环洋报价表】xlsx（HYE美国海外事业部大货价格表*）",
        initial,
        lambda p: _file_err(p, need_any=("Zone 2",)),
    )


def choose_stats_via_dialog(cfg) -> str | None:
    """《打单问题统计表》（天图未查到标记写这里）。"""
    initial = ""
    s = (cfg.get("paths") or {}).get("stats_file")
    if s and Path(s).parent.is_dir():
        initial = str(Path(s).parent)
    sheet = (cfg.get("paths") or {}).get("stats_sheet", "环洋")
    return _pick(
        f"④ 请选择【打单问题统计表】xlsx（天图未查到标记写入「{sheet}」页）",
        initial,
        lambda p: _file_err(p, need_any=("运单号",)),
    )


# ============== 天图引擎调用 ==============

def _engine_cmd(cfg: dict) -> tuple[Path, list]:
    """天图引擎路径与启动命令前缀。

    冻结后不能再用 sys.executable 当解释器（那会变成"用主流程去跑引擎脚本"），
    所以优先用发布包自带的 <HERE>/check_tiantu/check_tiantu.exe。
    """
    if getattr(sys, "frozen", False):
        exe = HERE / "check_tiantu" / "check_tiantu.exe"
        if exe.is_file():
            return exe, [str(exe)]
    raw = str((cfg.get("tiantu") or {}).get("engine") or "")
    engine = Path(raw).expanduser()
    if raw and not engine.is_absolute():
        engine = HERE / engine
    if engine.suffix.lower() == ".exe":
        return engine, [str(engine)]
    return engine, [PY, str(engine)]


def _engine_base(engine: Path) -> Path:
    """引擎认的「根目录」——credentials.local.json 和 config.json 都在这里。

    引擎（run_yongda_check.load_credentials / check_tiantu.default_output_dir）冻结态
    取 sys.executable 的父目录的**父目录**，也就是发布根；源码态取脚本所在目录。
    不按这个规则找的话，发布包里会跑去 check_tiantu/ 里找一个不存在的凭证文件。
    """
    return engine.parent.parent if engine.suffix.lower() == ".exe" else engine.parent


def _tiantu_dirs(engine: Path, our_out: Path) -> list[Path]:
    """结果 CSV 的候选目录。

    永达引擎把结果写到**它自己** config.json 的 output_dir（yongda-bill-check/outputs，
    发布包里则是发布根/outputs），不是我们的 outputs —— 两个都要找。
    """
    base = _engine_base(engine)
    dirs = [our_out, base / "outputs"]
    try:
        p = Path(json.loads((base / "config.json").read_text(encoding="utf-8"))
                 ["paths"]["output_dir"])
        dirs.append(p if p.is_absolute() else base / p)
    except Exception:  # noqa: BLE001
        pass
    out, seen = [], set()
    for d in dirs:
        if not d.is_dir():
            continue
        k = str(d.resolve()).lower()
        if k not in seen:
            seen.add(k)
            out.append(d)
    return out


def find_tiantu_result(engine: Path, our_out: Path, newer_than: float = 0.0) -> Path | None:
    """找最新的《天图核验结果_*.csv》。newer_than 用来排除上一次运行的旧结果。"""
    best, best_m = None, 0.0
    for d in _tiantu_dirs(engine, our_out):
        try:
            for f in d.glob("天图核验结果_*.csv"):
                m = f.stat().st_mtime
                if m >= newer_than and m > best_m:
                    best, best_m = f, m
        except OSError:
            continue
    return best


def _child_env() -> dict:
    """子进程一律按 UTF-8 说话（父进程可能跑在 cp936 控制台）。"""
    import os

    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def _csv_rows(path: Path) -> int:
    try:
        with path.open(encoding="utf-8-sig") as f:
            return max(0, sum(1 for line in f if line.strip()) - 1)
    except OSError:
        return 0


def _count_per_waybill(path: Path) -> int:
    """数引擎要**逐单**核验的行数。

    引擎里 mark 为「住宅私人」和 address 的分支不走批量查询，而是一单一单读
    （每个单号一次查询，实测 20-30 秒），是整段天图最慢的部分。
    """
    try:
        with path.open(encoding="utf-8-sig", newline="") as f:
            return sum(
                1 for r in csv.DictReader(f)
                if (r.get("期望标记") or "").strip() in ("住宅私人", "address")
            )
    except OSError:
        return 0


def _run_teeing(cmd: list, log_path: Path, timeout: int, idle_timeout=300) -> tuple[int, str]:
    """跑子进程，输出同时进控制台和日志（天图要跑十几分钟，现场输出不能藏）。

    两个中止条件：总时长超 timeout，或连续 idle_timeout 秒没有任何输出。
    后者才是真正判"卡死"的依据——总时长只能按清单长度估，估小了会把还在正常
    干活的引擎杀掉，整轮白跑。
    """
    log_path.parent.mkdir(parents=True, exist_ok=True)
    state = {"reason": "", "last": time.time()}
    with log_path.open("w", encoding="utf-8", errors="replace") as f:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, encoding="utf-8", errors="replace", bufsize=1,
            env=_child_env(),
        )
        stop = threading.Event()

        def _kill(reason: str):
            state["reason"] = reason
            try:
                proc.kill()
            except Exception:  # noqa: BLE001
                pass

        def _watch():
            deadline = time.time() + max(1, int(timeout))
            while not stop.wait(2):
                if time.time() - state["last"] > idle_timeout:
                    _kill(f"卡死：{idle_timeout} 秒没有任何输出")
                    return
                if time.time() >= deadline:
                    _kill(f"总时长超过上限 {timeout} 秒")
                    return

        threading.Thread(target=_watch, daemon=True).start()
        try:
            if proc.stdout is not None:
                for line in proc.stdout:
                    state["last"] = time.time()
                    _safe_print(line.rstrip("\n"))
                    f.write(line)
                    f.flush()
        finally:
            stop.set()
            try:
                proc.wait(timeout=60)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()
        return proc.returncode or 0, state["reason"]


def summarize_tiantu_log(log_path: Path, secs: float) -> str:
    """从引擎日志里数查询次数——天图慢就慢在查询次数上。

    引擎的每类查询都有固定措辞：
      `[批] 偏远 x13 -> 5 条`    一次批量查询
      `拆半重试/拆半补查`         批量没读全 → 拆半再查（可能退化成逐单）
      `单查 XXX -> N 条` / `应收核验` / `地址更正核验` / `[对账补查]`
    """
    if not log_path or not log_path.is_file():
        return f"总用时 {secs}s（没有日志，无法拆分查询次数）"
    pats = {
        "批量查询": re.compile(r"\[批\] .* x\d+ -> "),
        "拆半补查": re.compile(r"拆半(重试|补查)"),
        "单查": re.compile(r"^\S*\s*单查 "),
        "应收核验": re.compile(r"应收核验 "),
        "地址核验": re.compile(r"地址更正核验 "),
        "对账补查": re.compile(r"\[对账补查\]"),
    }
    counts = dict.fromkeys(pats, 0)
    for line in log_path.read_text(encoding="utf-8", errors="replace").splitlines():
        for k, pat in pats.items():
            if pat.search(line):
                counts[k] += 1
    total = sum(counts.values())
    per = (secs / total) if total else 0
    detail = "、".join(f"{k} {v}" for k, v in counts.items() if v)
    return (f"共 {total} 次查询（{detail or '无'}），用时 {secs}s，"
            f"平均 {per:.1f}s/次；日志 {log_path.name}")


def run_tiantu_check(cfg, checklist: Path, out_dir: Path, limit: int,
                     log: logging.Logger) -> Path | None:
    """复用 永达 check_tiantu.py 跑核验。

    引擎**没有** --result 参数（输出目录它自己决定），所以只能传它认识的
    --csv/--limit/--batch/--url，跑完再去候选目录里捞结果。
    """
    engine, head = _engine_cmd(cfg)
    if not engine.is_file():
        log.error("[天图] 找不到核验引擎：%s（检查 config.tiantu.engine）", engine)
        return None
    if not checklist or not checklist.is_file():
        log.error("[天图] 没有核验清单，跳过。")
        return None

    cmd = [*head, "--csv", str(checklist)]
    if limit > 0:
        cmd += ["--limit", str(int(limit))]
    t = cfg.get("tiantu") or {}
    batch = int(t.get("batch") or 13)
    cmd += ["--batch", str(batch)]
    if t.get("url"):
        cmd += ["--url", str(t["url"])]

    # 耗时随**批次数**增长（实测 45-60 秒/批）。但只按批次数估会严重低估：
    # 「住宅私人」和 address 是一单一单查的，每单 20-30 秒。宁松勿紧——
    # 进程跑完自己就退了，不会真等到上限。
    n = min(_csv_rows(checklist), limit) if limit > 0 else _csv_rows(checklist)
    batches = max(1, -(-n // batch))
    one_by_one = _count_per_waybill(checklist)
    if limit > 0:
        one_by_one = min(one_by_one, limit)
    timeout = max(
        int(t.get("timeout") or 300),
        batches * int(t.get("per_batch_seconds") or 60)
        + one_by_one * int(t.get("per_waybill_seconds") or 30) + 180,
    )
    idle = int(t.get("idle_seconds") or 300)
    log.info("[天图] 清单 %d 条 → 约 %d 批 + %d 条逐单核验；总时长上限 %d 秒、静默上限 %d 秒",
             n, batches, one_by_one, timeout, idle)
    log.info("[天图] 启动：%s", " ".join(cmd))

    started = time.time()
    run_log = out_dir / f"环洋_天图运行_{datetime.now():%Y%m%d_%H%M%S}.log"
    code, killed = _run_teeing(cmd, run_log, timeout, idle)
    secs = round(time.time() - started, 2)
    if killed:
        log.error("[天图] 引擎被中止（%s），本轮结果不可用。日志：%s", killed, run_log)
        return None
    if code != 0:
        log.error("[天图] 引擎退出码 %s（多为网络/登录失败），本轮结果不可用。日志：%s",
                  code, run_log)
        return None

    log.info("[天图] %s", summarize_tiantu_log(run_log, secs))
    res = find_tiantu_result(engine, out_dir, newer_than=started - 1)
    if not res:
        log.warning("[天图] 未生成结果 CSV（找过：%s）",
                    [str(d) for d in _tiantu_dirs(engine, out_dir)])
        return None
    log.info("[天图] 结果：%s", res)
    return res


# ============== 统计表 ==============

# 《打单问题统计表》各客户 sheet 的列序/列数都不一样
# （环洋 运单号|金额USD|费用名称，永达 多一个备注），一律按表头名定位，不写死列号。
STATS_ALIASES = {
    "no":     ["运单号", "客户单号", "系统单号"],
    "amount": ["金额USD", "费用金额", "金额"],
    "fee":    ["费用名称", "费用类别"],
    "remark": ["备注", "说明"],
}


def _num(value):
    """CSV 读出来一律是字符串，直接写进 Excel 会变成**文本**（那一列求和会算成 0）。

    转不成数字就原样返回，宁可难看也别把内容丢了。
    """
    s = str(value if value is not None else "").strip()
    if not s:
        return ""
    try:
        return float(s)
    except ValueError:
        return s


def collect_tiantu_missing(checklist_csv: Path | None, tiantu_csv: Path | None,
                           log: logging.Logger) -> list[dict]:
    """挑出「天图查不到标记」的问题，供写入《打单问题统计表》。

    引擎结果 CSV 的列是 `单号/费用名称/期望标记/是否出现/命中条数/明细/缺失关键词`
    （**不是**清单那套 `客户单号(去后缀)/费用类别/依据`）。金额按 (单号, 费用名称)
    去本次清单里补，同时用清单把上一次运行的旧结果剔掉。

    结论分两档写到备注里，因为处理方式不一样：
      `否`            → "天图未查到标记：X"  —— 确认没有，按 SOP 登记
      `未定位/异常`   → "天图未能定位该单（需人工复核）：X" —— 是**没读到那行**，
                        不代表天图没有。照实登记但写清楚，免得人工当成"确认没有"处理掉。
    """
    if not tiantu_csv or not tiantu_csv.is_file():
        log.warning("[统计表] 没有天图结果 CSV，本次不产出天图问题。")
        return []

    known: dict[tuple[str, str], dict] = {}
    if checklist_csv and checklist_csv.is_file():
        with checklist_csv.open(encoding="utf-8-sig") as f:
            for r in csv.DictReader(f):
                no = (r.get("客户单号(去后缀)") or "").strip()
                fee = (r.get("费用名称") or "").strip()
                if no and fee:
                    known[(no, fee)] = r
    if not known:
        log.warning("[统计表] 本次核验清单为空或缺失，无法确认天图结果属于本张账单，跳过。")
        return []

    items, seen, unmatched = [], set(), []
    total = 0
    by_verdict: dict[str, int] = {}
    with tiantu_csv.open(encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            total += 1
            verdict = (r.get("是否出现") or "").strip()
            by_verdict[verdict] = by_verdict.get(verdict, 0) + 1
            if verdict == "是":
                continue
            no = (r.get("单号") or "").strip()
            fee = (r.get("费用名称") or "").strip()
            if not no or not fee:
                continue
            cl = known.get((no, fee))
            if cl is None:
                unmatched.append(f"{no}/{fee}")
                continue
            if (no, fee) in seen:
                continue
            seen.add((no, fee))
            mark = (r.get("期望标记") or cl.get("期望标记") or "").strip()
            why = (r.get("缺失关键词") or "").strip()
            if verdict == "否":
                remark = "天图未查到标记：" + mark + (f"（缺 {why}）" if why else "")
            else:
                # 「未定位 / 异常(N条)」是**没读到那一行**，不等于天图没有这个标记
                # （引擎自己也是把它们算作"需人工复核"而不是"明确未出现"）。
                # 照实登记、但在备注里写清楚，免得人工当成"确认没有"处理掉。
                remark = f"天图未能定位该单（{verdict}，需人工复核）：{mark}"
            items.append({
                "no": no,
                "fee": fee,
                "amount": _num(cl.get("金额USD")),
                "remark": remark,
            })
    if unmatched:
        log.warning("[统计表] 天图结果里 %d 条不在本次核验清单内，已忽略：%s",
                    len(unmatched), "、".join(unmatched[:8]))
    log.info("[统计表] 天图结果 %d 行（%s），登记 %d 条",
             total, "、".join(f"{k or '空'} {v}" for k, v in sorted(by_verdict.items())),
             len(items))
    return items


def write_to_stats(stats_file, sheet_name: str, items: list[dict], dry_run: bool,
                   log: logging.Logger) -> bool:
    """把「天图未查到标记」追加到《打单问题统计表》的目标 sheet。

    列序按表头名适配（见 STATS_ALIASES），按 (运单号, 费用名称) 去重，重复行补空备注。
    表里没有「备注」列时自动在表头行末尾补一个，否则「天图未查到 偏远」这个关键
    信息会随备注一起丢掉，人工看不出为什么登记这行。
    """
    if not items:
        log.info("[统计表] 没有「天图未查到标记」的问题，不写。")
        return True
    if dry_run:
        log.info("[dry-run] 拟写入《打单问题统计表》%d 条：", len(items))
        for it in items[:10]:
            log.info("  %s", it)
        return True

    p = Path(stats_file or "")
    if not p.is_file():
        log.error("[统计表] 文件不存在：%s", p)
        return False
    _, excel = _autolib()
    import openpyxl  # noqa: PLC0415

    try:
        wb = openpyxl.load_workbook(p)
    except Exception as exc:  # noqa: BLE001
        log.error("[统计表] 打不开（是不是在 Excel 里开着？）：%s -> %s", p, exc)
        return False
    try:
        if sheet_name not in wb.sheetnames:
            log.error("[统计表] 找不到 sheet「%s」，现有：%s", sheet_name, wb.sheetnames)
            return False
        ws = wb[sheet_name]
        hrow, _ = excel.find_header_row(ws, required=("运单号",))
        if hrow is None:
            log.error("[统计表] sheet「%s」找不到含「运单号」的表头行。", sheet_name)
            return False
        colmap = excel.map_columns([c.value for c in ws[hrow]], STATS_ALIASES)
        if "no" not in colmap or "fee" not in colmap:
            log.error("[统计表] sheet「%s」表头缺列（至少要有 运单号 与 费用名称）：%s",
                      sheet_name, [c.value for c in ws[hrow]])
            return False
        if "remark" not in colmap and any(it.get("remark") for it in items):
            last = max((c.column for c in ws[hrow] if c.value not in (None, "")), default=0)
            ws.cell(row=hrow, column=last + 1).value = "备注"
            colmap["remark"] = last + 1
            log.info("[统计表] 「%s」原本没有「备注」列，已在表头行第 %d 列补上",
                     sheet_name, last + 1)
        added, skipped, filled = excel.append_rows(ws, items, colmap)
        saved, degraded = excel.safe_save(wb, p)
    finally:
        wb.close()
    log.info("[统计表] %s「%s」新增 %d 行 / 重复跳过 %d 行 / 补备注 %d 行%s",
             Path(saved).name, sheet_name, added, skipped, filled,
             "（原文件被占用，已写副本）" if degraded else "")
    return True


# ============== 通知 ==============

def notify_all(cfg: dict, content: str, status="成功", start_ts=None,
               stages=None, note="") -> None:
    try:
        _autolib_path()
        from autolib.notify import notify_all as _n  # noqa: PLC0415

        _n(cfg, content, status=status, start_ts=start_ts, stages=stages, note=note)
    except Exception as exc:  # noqa: BLE001
        _safe_print(f"[通知] 跳过（{exc}）")


# ============== 主流程 ==============

def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="环洋打单数据整理 — 总流程")
    ap.add_argument("--bill1", help="账单1《成本汇总》xlsx")
    ap.add_argument("--bill2", help="账单2《成本明细》xlsx")
    ap.add_argument("--quote", help="报价表 xlsx")
    ap.add_argument("--stats", help="打单问题统计表 xlsx")
    ap.add_argument("--translation", help="《打单费用名称中英文翻译》xlsx（默认取 config）")
    ap.add_argument("--no-dialog", action="store_true", help="不弹窗，全部走命令行/config")
    ap.add_argument("--dry-run", action="store_true", help="不回填账单、不写统计表")
    ap.add_argument("--tiantu-limit", type=int, default=0, help="天图只核验前 N 条")
    ap.add_argument("--skip-tiantu", action="store_true")
    ap.add_argument("--selftest", action="store_true",
                    help="天图自测：抽少量单号跑两遍比对，然后退出")
    ap.add_argument("--selftest-limit", type=int, default=0,
                    help="天图自测抽几个单号（默认 config.selftest.limit）")
    args = ap.parse_args(argv)

    import run_huanyang_check as R  # noqa: PLC0415
    from process_details import (  # noqa: PLC0415
        build_checklist, build_translation_map, load_rules, process,
    )
    from quote_engine import QuoteBook, load_summary_index  # noqa: PLC0415

    cfg = _load_notify_local(_load_cfg())
    log = _setup_logger(cfg)
    out_dir = _out_dir(cfg)
    out_dir.mkdir(parents=True, exist_ok=True)
    start_ts = time.time()
    stages = []

    def _lap(name, ok, t0):
        secs = round(time.time() - t0, 2)
        stages.append({"name": name, "secs": secs, "ok": ok})
        log.info("[阶段] %s 用时 %ss（%s）", name, secs, "完成" if ok else "失败")
        return secs

    # ---- 选文件（SOP 要求两个账单各弹一次）----
    if args.no_dialog:
        f = R.pick_files(cfg, args)
        bill1 = str(f.get("bill1") or "")
        bill2 = str(f.get("bill2") or "")
        quote = str(f.get("quote") or "")
        stats = str(f.get("stats") or "")
    else:
        bill1 = args.bill1 or choose_summary_via_dialog(cfg) or ""
        if not bill1:
            log.error("取消选择账单1，流程结束。")
            return 1
        bill2 = args.bill2 or choose_detail_via_dialog(cfg) or ""
        if not bill2:
            log.error("取消选择账单2，流程结束。")
            return 1
        quote = args.quote or choose_quote_via_dialog(cfg) or ""
        if not quote:
            log.error("取消选择报价表，流程结束。")
            return 1
        stats = args.stats or choose_stats_via_dialog(cfg) or ""

    log.info("账单1（成本汇总）：%s", bill1)
    log.info("账单2（成本明细）：%s", bill2)
    log.info("报价表            ：%s", quote)
    log.info("统计表            ：%s", stats or "（未选，跳过登记）")
    if not Path(bill1).is_file() or not Path(bill2).is_file() or not Path(quote).is_file():
        log.error("输入文件不存在，流程结束。")
        return 1

    # ---- 1) 加工 + 核价 ----
    t0 = time.time()
    ok = False
    try:
        summary = load_summary_index(bill1, (cfg.get("bills") or {}).get("summary_sheet", "成本汇总"),
                                     log=lambda m: log.info("%s", m))
        tr_cfg = cfg.get("translation") or {}
        tr_file = _resolve_translation(cfg, bill2, args.translation)
        if tr_file and Path(tr_file).is_file():
            translation = build_translation_map(
                tr_file, tr_cfg.get("sheet", "Sheet1"),
                overrides=tr_cfg.get("overrides"), log=lambda m: log.info("%s", m),
            )
        else:
            translation = {}
            raise FileNotFoundError(
                "没找到翻译表：请把打单费用名称中英文翻译.xlsx 放在账单旁边或发布包根目录"
            )
        quote_book = QuoteBook(
            quote,
            product_aliases=(cfg.get("quote") or {}).get("product_aliases"),
            hwt_min_weight=(cfg.get("quote") or {}).get("hwt_min_weight", 200),
        )
        if args.dry_run:
            saved, stat, rows = R._dry_run(bill2, summary, quote_book, translation, cfg,
                                           lambda m: log.info("%s", m),
                                           translation_file=tr_file)
        else:
            saved, stat, rows = process(
                bill2, summary, quote_book, translation, cfg,
                log=lambda m: log.info("%s", m),
                backup=bool((cfg.get("run") or {}).get("backup_original", True)),
                translation_file=tr_file,
            )
        ok = True
    except Exception as exc:  # noqa: BLE001
        log.exception("[核价] 失败：%s", exc)
        _lap("加工核价", False, t0)
        notify_all(cfg, f"环洋打单数据整理\n状态：失败\n原因：{exc}", "失败", start_ts, stages)
        return 1
    _lap("加工核价", ok, t0)

    if stat.get("总行数", 0) and stat.get("已翻译", 0) == 0:
        log.error("翻译结果为 0 行，已中止，避免生成空报价结果。")
        return 1

    # 记录核价统计信息
    price_summary = f"已报价 {stat['已报价']} 行，未匹配 {stat['未匹配报价']} 行，差异超容差 {stat['差异超容差']} 行"

    # ---- 2) 天图核验清单 ----
    tasks = build_checklist(rows, cfg, translation, load_rules(cfg),
                            log=lambda m: log.info("%s", m))
    checklist = R.write_checklist_csv(tasks, out_dir) if tasks else None

    # ---- 3) 天图核验 ----
    tiantu_csv = None
    if checklist and not args.skip_tiantu:
        t0 = time.time()
        if args.selftest:
            import selftest_tiantu  # noqa: PLC0415

            selftest_tiantu.run_selftest(cfg, checklist, out_dir, log,
                                         limit=args.selftest_limit or None)
            return 0
        tiantu_csv = run_tiantu_check(cfg, checklist, out_dir, args.tiantu_limit, log)
        _lap("天图核验", bool(tiantu_csv), t0)

    # ---- 4) 统计表 ----
    items = collect_tiantu_missing(checklist, tiantu_csv, log)
    stats_ok = True
    if stats:
        stats_ok = write_to_stats(
            stats, (cfg.get("paths") or {}).get("stats_sheet", "环洋"),
            items, args.dry_run, log,
        )
    else:
        log.info("[统计表] 未选择统计表，跳过登记。")

    # ---- 5) 报告 + 通知 ----
    t0 = time.time()
    report = R.write_report(stat, cfg, out_dir, checklist, quote=quote_book, log=log)
    _lap("报告", True, t0)
    log.info("[报告] %s", report)

    bad = stat["差异超容差"]
    summary_msg = [
        "环洋打单数据整理",
        f"状态：{'成功' if stats_ok else '部分成功（统计表未写入）'}",
        f"账单2：{Path(bill2).name}",
        f"填出运单号 {stat['已填运单号']} / 总行 {stat['总行数']}",
        f"已填报价 {stat['已报价']} 行，未匹配 {stat['未匹配报价']} 行",
        f"差异超容差 {bad} 行",
        f"天图清单 {len(tasks)} 条；天图问题 {len(items)} 条已登记统计表",
        f"报告：{report.name}",
    ]
    for line in summary_msg:
        log.info("[结果] %s", line)
    notify_all(cfg, "\n".join(summary_msg), "成功" if stats_ok else "部分成功",
               start_ts, stages)
    return 0


if __name__ == "__main__":
    sys.exit(main())
