# -*- coding: utf-8 -*-
"""SKYE 打单数据整理 — 总流程。

  1) 弹窗选账单 xlsx；再弹窗选**尾程渠道报价表** xlsx（报价表时常更换，故每次都要选）
  2) run_skye_check.py    核价 + 直接修改原表 + 报价差异 + 天图核验清单
  3) check_tiantu.py (复用 永达)   逐单去天图系统查标记
  4) 把「天图未查到标记」的问题写进独立《打单问题统计表》（写前弹窗确认目标文件）
  5) 通知（autolib.notify.notify_all，失败不抛）

运行：python main_flow.py [--input <账单.xlsx>] [--quote <报价表.xlsx>]
                        [--stats <统计表.xlsx>] [--no-dialog]
                        [--dry-run] [--tiantu-limit N]
"""

from __future__ import annotations

import argparse
import csv
import json
import logging
import os
import re
import subprocess
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

if getattr(sys, "frozen", False):
    HERE = Path(sys.executable).resolve().parent.parent
    # 冻结后 PYTHONUTF8 不生效，管道默认按 cp936 写中文；父流程按 utf-8 解码并靠
    # 中文措辞统计查询次数，不强制 utf-8 会整段乱码、次数数成 0。
    for _s in (sys.stdout, sys.stderr):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass
else:
    HERE = Path(__file__).resolve().parent
PY = sys.executable

_EXCEL_TYPES = [("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")]


# ============== 工具 ==============

def _safe_print(msg: str) -> None:
    try:
        print(msg)
    except UnicodeEncodeError:
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        sys.stdout.buffer.write((msg + "\n").encode(enc, "replace"))
        sys.stdout.flush()


def _load_cfg() -> dict:
    return json.loads((HERE / "config.json").read_text(encoding="utf-8"))


def _load_notify_local(cfg: dict) -> dict:
    """从本机 notify.local.json 覆盖通知相关字段（不入库）。"""
    p = HERE / "notify.local.json"
    if not p.is_file():
        return cfg
    try:
        local = json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return cfg
    nt = cfg.setdefault("notify", {})
    if local.get("wecom_key"):
        nt["wecom_key"] = local["wecom_key"]
    if local.get("kdocs"):
        nt["kdocs"] = {**nt.get("kdocs", {}), **local["kdocs"]}
    if local.get("app_name"):
        nt["app_name"] = local["app_name"]
    if local.get("account_id"):
        nt["account_id"] = local["account_id"]
    return cfg


def _default_output_dir() -> Path:
    try:
        cfg = _load_cfg()
        p = Path(cfg["paths"]["output_dir"])
        return p if p.is_absolute() else HERE / p
    except Exception:  # noqa: BLE001
        return HERE / "outputs"


def _latest(out_dir: Path, pattern: str) -> Path | None:
    cands = sorted(out_dir.glob(pattern), key=lambda f: f.stat().st_mtime, reverse=True)
    return cands[0] if cands else None


def _latest_since(out_dir: Path, pattern: str, since: float) -> Path | None:
    """只认本次运行之后生成的产物。

    否则引擎跑失败时，_latest() 会把**上一次**的天图核验清单塞给天图阶段，
    结果是拿旧清单去查今天的天图 → 统计表里出现早已处理完的单号。
    """
    try:
        cands = [f for f in out_dir.glob(pattern) if f.stat().st_mtime >= since - 1]
    except OSError:
        return None
    return max(cands, key=lambda f: f.stat().st_mtime) if cands else None


def _add_autolib_path() -> None:
    """开发态把 08-common-utils 挂到 sys.path。

    打包后 `HERE.parents[1]` 指向发布目录的祖父目录（不存在），但 autolib 已经被
    PyInstaller 打进 exe，直接 import 就能找到 —— 所以冻结时什么都不做。
    """
    if getattr(sys, "frozen", False):
        return
    sys.path.insert(0, str(HERE.parents[1] / "08-common-utils"))


def _autolib():
    """懒加载 08-common-utils/autolib（弹窗/excel 辅助）。"""
    _add_autolib_path()
    import autolib.dialog as dialog  # noqa: PLC0415
    import autolib.excel as excel  # noqa: PLC0415
    return dialog, excel


def _openpyxl_read_sheets(path: Path) -> list[str]:
    import openpyxl  # noqa: PLC0415
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        return list(wb.sheetnames)
    finally:
        wb.close()


# ============== 弹窗 ==============

def choose_bill_via_dialog() -> str | None:
    import tkinter as tk
    from tkinter import filedialog
    initial = ""
    try:
        cfg = _load_cfg()
        p = Path(cfg["paths"]["input_dir"])
        initial = str(p) if p.exists() else ""
    except Exception:  # noqa: BLE001
        pass
    for _ in range(10):
        root = tk.Tk()
        root.withdraw()
        try:
            path = filedialog.askopenfilename(
                title="请选择 SKYE 账单 xlsx（如 M8123*.xlsx）",
                filetypes=[("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")],
                initialdir=initial or str(Path.home()),
            )
        finally:
            root.destroy()
        if not path:
            return None
        # 简单校验
        from tkinter import messagebox
        err = _validate_bill_file(path)
        if err is None:
            return path
        messagebox.showerror("文件不符合要求", err)
    _safe_print("连续多次选错文件，流程取消。")
    return None


def _validate_bill_file(path: str) -> str | None:
    p = Path(path or "")
    if not p.exists():
        return "文件不存在，请重新选择。"
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        return f"不是 Excel 工作簿：{p.name}"
    try:
        import openpyxl  # noqa: PLC0415
        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        try:
            sheets = wb.sheetnames
        finally:
            wb.close()
    except Exception as exc:  # noqa: BLE001
        return f"无法打开文件：{exc}"
    if "Master" not in sheets:
        return "选择的文件不是 SKYE 账单（缺少 Master sheet），请重新选择。"
    return None


def _validate_quote_file(path: str) -> str | None:
    """报价表校验：必须是含渠道 sheet 的 xlsx（不校验具体渠道名，报价表会换版）。"""
    p = Path(path or "")
    if not p.exists():
        return "文件不存在，请重新选择。"
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        return f"不是 Excel 工作簿：{p.name}"
    try:
        sheets = _openpyxl_read_sheets(p)
    except Exception as exc:  # noqa: BLE001
        return f"无法打开文件：{exc}"
    if not sheets:
        return "工作簿里没有 sheet。"
    if not any(("MWT" in s) or ("Ground" in s) for s in sheets):
        return (f"这不像尾程渠道报价表（没找到 MWT/Ground 渠道 sheet）。\n"
                f"文件：{p.name}\n现有 sheet：{'、'.join(sheets[:12])}")
    return None


def choose_quote_via_dialog(default_quote: str | None = None) -> str | None:
    """弹窗选**尾程渠道报价表** xlsx。

    报价表时常更换，所以每次都问一次。标题特意写明「尾程渠道报价表」——
    它和选账单的弹窗长得几乎一样，标题不区分的话极易选错文件。
    """
    dialog, _ = _autolib()
    initial = ""
    for cand in (default_quote, ""):
        if cand:
            d = Path(cand)
            if d.parent.is_dir():
                initial = str(d.parent)
                break
    if not initial:
        try:
            p = Path(_load_cfg()["paths"]["quote_file"])
            if p.parent.is_dir():
                initial = str(p.parent)
        except Exception:  # noqa: BLE001
            pass
    if not initial:
        initial = str(Path.home())
    return dialog.pick_until_valid(
        "请选择【尾程渠道报价表】xlsx（如 2026xxxx-M8123-SKYE尾程渠道报价.xlsx）",
        _EXCEL_TYPES, initial, validator=_validate_quote_file,
    )


def _validate_stats_file(path: str, sheet_name: str | None = None) -> str | None:
    p = Path(path or "")
    if not p.exists():
        return "文件不存在，请重新选择。"
    if p.suffix.lower() not in (".xlsx", ".xlsm"):
        return f"不是 Excel 工作簿：{p.name}"
    try:
        sheets = _openpyxl_read_sheets(p)
    except Exception as exc:  # noqa: BLE001
        return f"无法打开文件：{exc}"
    if not sheets:
        return "工作簿里没有 sheet。"
    # 选错文件（比如把账单选进来）在这里就拦住，别等写的时候才发现没这张表
    if sheet_name and sheet_name not in sheets:
        return f"《{p.name}》里没有 sheet「{sheet_name}」，请选择打单问题统计表。"
    return None


def choose_stats_via_dialog(cfg: dict) -> str | None:
    """弹窗确认《打单问题统计表》的目标文件。

    标题写明「打单问题统计表（天图未查到标记写这里）」，和账单/报价表两个弹窗区分开。
    """
    dialog, _ = _autolib()
    default = str(cfg.get("paths", {}).get("stats_file") or "")
    sheet_name = cfg.get("paths", {}).get("stats_sheet", "skye")
    initial = ""
    if default:
        d = Path(default)
        if d.parent.is_dir():
            initial = str(d.parent)
    return dialog.pick_until_valid(
        f"请选择【打单问题统计表】xlsx（含 sheet「{sheet_name}」，天图未查到标记的问题写到这里）",
        _EXCEL_TYPES, initial or str(Path.home()),
        validator=lambda p: _validate_stats_file(p, sheet_name),
    )


# ============== 阶段 ==============

def run_stage(name: str, script: Path, extra: list | None = None):
    """拉起一个子脚本。子进程缺失/崩溃时返回非零码而不是让主流程炸掉。"""
    _safe_print(f"\n{'='*60}\n[主流程] {name}: {script.name}\n{'='*60}")
    t0 = time.time()
    if getattr(sys, "frozen", False):
        exe = HERE / script.stem / f"{script.stem}.exe"
        cmd = [str(exe)] + (extra or [])
        if not exe.is_file():
            _safe_print(f"[主流程] 找不到可执行文件：{exe}")
            return 127, round(time.time() - t0, 2)
    else:
        cmd = [PY, str(script)] + (extra or [])
        if not script.is_file():
            _safe_print(f"[主流程] 找不到脚本：{script}")
            return 127, round(time.time() - t0, 2)
    try:
        code = subprocess.run(cmd).returncode
    except Exception as exc:  # noqa: BLE001
        _safe_print(f"[主流程] {name} 启动失败：{exc}")
        return 1, round(time.time() - t0, 2)
    secs = round(time.time() - t0, 2)
    _safe_print(f"[主流程] {name} 用时 {secs}s（退出码 {code}）")
    return code, secs


def _engine_cmd(cfg: dict) -> tuple[Path, list]:
    """天图引擎的路径与启动命令前缀。

    打包后发布根下是 `<HERE>/check_tiantu/check_tiantu.exe`，而 config 里写的是开发机的
    绝对路径 —— 所以冻结时优先用发布包自带的引擎 exe；并且**不能**再用 `sys.executable`
    当解释器（冻结后那就是 main_flow.exe 自己，会变成"用主流程去跑引擎脚本"）。
    """
    if getattr(sys, "frozen", False):
        exe = HERE / "check_tiantu" / "check_tiantu.exe"
        if exe.is_file():
            return exe, [str(exe)]
    raw = str(cfg.get("tiantu", {}).get("engine") or "")
    engine = Path(raw).expanduser()
    if raw and not engine.is_absolute():
        engine = HERE / engine
    if engine.suffix.lower() == ".exe":
        return engine, [str(engine)]
    return engine, [PY, str(engine)]


def _tiantu_search_dirs(engine: Path, our_out: Path) -> list[Path]:
    """天图结果 CSV 的候选目录。

    永达的 check_tiantu.py 把结果写到**它自己** config.json 的 paths.output_dir
    （默认 yongda-bill-check/outputs），不是我们的 outputs。两个目录都要找。
    """
    dirs = [our_out]
    # 引擎所在目录的 outputs/
    cand = engine.parent / "outputs"
    if cand.is_dir():
        dirs.append(cand)
    # 引擎自己 config.json 里配的 output_dir
    try:
        cfg_txt = (engine.parent / "config.json").read_text(encoding="utf-8")
        p = Path(json.loads(cfg_txt)["paths"]["output_dir"])
        if not p.is_absolute():
            p = engine.parent / p
        if p.is_dir():
            dirs.append(p)
    except Exception:  # noqa: BLE001
        pass
    # 去重保序
    seen: set[str] = set()
    out: list[Path] = []
    for d in dirs:
        key = str(d.resolve()).lower()
        if key in seen:
            continue
        seen.add(key)
        out.append(d)
    return out


def find_tiantu_result(engine: Path, our_out: Path, newer_than: float = 0.0) -> Path | None:
    """在候选目录里找最新的《天图核验结果_*.csv》。

    :param newer_than: 只认修改时间晚于该时间戳的文件，避免拿到上一次运行的旧结果
    """
    best: Path | None = None
    best_m = 0.0
    for d in _tiantu_search_dirs(engine, our_out):
        try:
            for f in d.glob("天图核验结果_*.csv"):
                m = f.stat().st_mtime
                if m < newer_than or m <= best_m:
                    continue
                best, best_m = f, m
        except OSError:
            continue
    return best


def preflight_tiantu(url: str, log: logging.Logger) -> bool:
    """起浏览器之前先探一下站点可达性，避免白白等一条 Playwright 堆栈。

    注意：这里用 urllib（它认 http_proxy/https_proxy），而 Chromium 走系统代理设置，
    两者结论可能不一致。探测失败只告警、不阻断——真正的判定交给引擎自己。
    """
    import os
    import urllib.request
    import ssl

    host = url.split("//", 1)[-1].split("/")[0].split("#")[0]
    if not host:
        return True
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE
    try:
        with urllib.request.urlopen(f"https://{host}/", timeout=10, context=ctx) as r:
            ok = 200 <= getattr(r, "status", 200) < 500
    except Exception as exc:  # noqa: BLE001
        log.warning("[天图] 预检：直连 %s 失败（%s）", host, type(exc).__name__)
        if os.environ.get("https_proxy") or os.environ.get("HTTPS_PROXY"):
            log.warning("[天图] 预检：检测到 https_proxy=%s。"
                        "Chromium 默认不读该环境变量而走系统代理设置，"
                        "两者不通会导致 Playwright 报 ERR_CONNECTION_CLOSED。",
                        os.environ.get("https_proxy") or os.environ.get("HTTPS_PROXY"))
        return False
    if not ok:
        log.warning("[天图] 预检：%s 返回异常状态", host)
    return ok


def _csv_data_rows(path: Path) -> int:
    """数清单里有几行任务（不含表头）。只用来估超时，不用精确。"""
    try:
        with path.open(encoding="utf-8-sig") as f:
            return max(0, sum(1 for line in f if line.strip()) - 1)
    except OSError:
        return 0


def _count_per_waybill_rows(path: Path) -> int:
    """数出引擎要**逐单**核验的行数。

    `check_tiantu.py` 里 `mark == "住宅私人"` 的分支不走批量查询，而是一单一单去读应收栏
    （每个单号一次查询，实测 20-30 秒），是整段天图最慢的部分。
    """
    try:
        with path.open(encoding="utf-8-sig", newline="") as f:
            return sum(1 for r in csv.DictReader(f)
                       if (r.get("期望标记") or "").strip() == "住宅私人")
    except OSError:
        return 0


def _child_env() -> dict:
    """子进程一律按 UTF-8 说话。

    打包后父进程可能跑在 cp936 控制台，子引擎（另一个 exe）就会按 cp936 往管道里写中文，
    而我们按 utf-8 解码、且 summarize_tiantu_log 要靠中文措辞数查询次数 → 会数成 0 次。
    """
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def _run_teeing(cmd: list, log_path: Path, timeout: int,
                idle_timeout: int = 300) -> tuple[int, str]:
    """跑子进程，输出同时进控制台和日志文件（天图要跑十几分钟，现场输出不能藏起来）。

    引擎自己会逐批打印进度，这些行既是给人看的，也是自测的原始数据：
    退出后由 summarize_tiantu_log() 数出"查了多少次"。

    两个中止条件：**总时长**超过 timeout，或者**连续 idle_timeout 秒一行输出都没有**。
    后者才是真正判"卡死"的依据 —— 天图整段要跑几十分钟，总时长只能按清单长度估，
    估小了就会把还在正常干活的引擎杀掉，整轮白跑。

    :return: (退出码, 中止原因)；正常跑完中止原因为空串
    """
    log_path.parent.mkdir(parents=True, exist_ok=True)
    state = {"reason": "", "last": time.time()}
    with log_path.open("w", encoding="utf-8", errors="replace") as f:
        proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                text=True, encoding="utf-8", errors="replace", bufsize=1,
                                env=_child_env())
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
                # 引擎卡住时可能长时间不输出，所以不能只靠"读到的行"判活
                if time.time() - state["last"] > idle_timeout:
                    _kill(f"卡死：{idle_timeout} 秒没有任何输出")
                    return
                if time.time() >= deadline:
                    _kill(f"总时长超过上限 {timeout} 秒")
                    return

        watchdog = threading.Thread(target=_watch, daemon=True)
        watchdog.start()
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
    """从引擎日志里数出查询次数——天图慢就慢在查询次数上。

    永达引擎的每类查询都有固定措辞：
      `[批] 偏远 x13 -> 5 条`        一次批量查询
      `[批] … 拆半重试/拆半补查`      批量没读全 → 拆半再查（最费时：可能退化成逐单）
      `单查 XXX -> N 条` / `应收核验 …` / `地址更正核验 …` / `[对账补查] …`
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
    counts = {k: 0 for k in pats}
    for line in log_path.read_text(encoding="utf-8", errors="replace").splitlines():
        if pats["批量查询"].search(line):
            counts["批量查询"] += 1
        if pats["拆半补查"].search(line):
            counts["拆半补查"] += 1
        if pats["单查"].search(line):
            counts["单查"] += 1
        if pats["应收核验"].search(line):
            counts["应收核验"] += 1
        if pats["地址核验"].search(line):
            counts["地址核验"] += 1
        if pats["对账补查"].search(line):
            counts["对账补查"] += 1
    total = sum(counts.values())
    per = (secs / total) if total else 0
    detail = "、".join(f"{k} {v}" for k, v in counts.items() if v)
    return (f"自测：共 {total} 次查询（{detail or '无'}），"
            f"用时 {secs}s，平均 {per:.1f}s/次；日志 {log_path.name}")


def run_tiantu_check(cfg: dict, checklist: Path, out_dir: Path, limit: int, log: logging.Logger) -> Path | None:
    """复用 永达 check_tiantu.py。

    注意：永达的脚本**不支持** --result 参数（它自己决定输出目录），
    所以这里只能传它认识的 --csv/--limit/--batch/--url，跑完再去它的输出目录捞结果。
    """
    engine, cmd_head = _engine_cmd(cfg)
    if not engine.is_file():
        log.warning("[天图] 未配置 tiantu.engine（或路径不存在），跳过天图核验。")
        return None
    if not checklist or not checklist.is_file():
        log.warning("[天图] 没有核验清单，跳过。")
        return None

    cmd = [*cmd_head, "--csv", str(checklist)]
    if limit > 0:
        cmd += ["--limit", str(int(limit))]
    batch = int(cfg.get("tiantu", {}).get("batch") or 13)
    if batch:
        cmd += ["--batch", str(batch)]
    url = cfg.get("tiantu", {}).get("url")
    if url:
        cmd += ["--url", str(url)]

    # 引擎按"标记"分组、每批最多 batch 个单号查一次，耗时随**批次数**增长（实测约 45-60 秒一批）：
    # 清单是"所有收费行"后会有上百条，写死 300 秒必然超时 → 按清单长度估一个上限（宁松勿紧，
    # 进程跑完就退出，不会真等到上限）。
    # 但**只按批次数估会严重低估**：标记为「住宅私人」的单号引擎是一单一单查应收栏的，
    # 每单 20-30 秒，上百条单号时这一项比所有批量查询加起来还慢（2026-09-17 现场：
    # 清单 234 条估出 1260 秒，跑到一半被自己的超时掐掉，用户反馈「天图核验步骤失败」）。
    n_rows = _csv_data_rows(checklist)
    if limit > 0:
        n_rows = min(n_rows, int(limit))
    batches = max(1, -(-n_rows // batch))
    per_batch = int(cfg.get("tiantu", {}).get("per_batch_seconds") or 60)
    n_one_by_one = _count_per_waybill_rows(checklist)
    if limit > 0:
        n_one_by_one = min(n_one_by_one, int(limit))
    per_one = int(cfg.get("tiantu", {}).get("per_waybill_seconds") or 30)
    timeout = max(int(cfg.get("tiantu", {}).get("timeout") or 300),
                  batches * per_batch + n_one_by_one * per_one + 180)
    idle_timeout = int(cfg.get("tiantu", {}).get("idle_seconds") or 300)
    log.info("[天图] 清单 %d 条 → 约 %d 批 + %d 条逐单核验，总时长上限 %d 秒、静默上限 %d 秒",
             n_rows, batches, n_one_by_one, timeout, idle_timeout)

    if url and not preflight_tiantu(str(url), log):
        log.warning("[天图] 站点探测不通，仍会尝试启动浏览器（本机直连通常是通的）。")
    log.info("[天图] 启动：%s", " ".join(cmd))

    started = time.time()
    run_log = out_dir / f"{checklist.stem.split('_')[0]}_天图运行_{time.strftime('%Y%m%d_%H%M%S')}.log"
    code, killed = _run_teeing(cmd, run_log, timeout, idle_timeout)
    secs = round(time.time() - started, 2)
    if killed:
        log.error("[天图] 引擎被中止（%s），本轮天图核验结果不可用。日志：%s", killed, run_log)
        return None
    if code != 0:
        log.error("[天图] 引擎退出码 %s（多为网络/登录失败），本轮天图核验结果不可用。日志：%s",
                  code, run_log)
        return None

    # 自测：天图的耗时几乎只跟"查了多少次"成正比（每次查询 5-20 秒），
    # 把这些数字打出来，下次要缩短时间就知道该动哪里。
    log.info("[天图] %s", summarize_tiantu_log(run_log, secs))

    res = find_tiantu_result(engine, out_dir, newer_than=started - 1)
    if not res:
        log.warning("[天图] 未生成结果 CSV（搜索目录：%s）",
                    [str(d) for d in _tiantu_search_dirs(engine, out_dir)])
        return None
    log.info("[天图] 结果：%s", res)
    return res


def collect_tiantu_missing(checklist_csv: Path | None, tiantu_csv: Path | None,
                           log: logging.Logger) -> list[dict]:
    """挑出「天图系统里查不到标记」的问题，供写入《打单问题统计表》。

    SOP：核价差异走「报价差异」sheet，**不进**统计表；统计表只登记天图未查到的。

    永达结果 CSV 的列是 `单号/费用名称/期望标记/是否出现/命中条数/明细/缺失关键词`
    （**不是**清单那套 `客户单号(去后缀)/费用类别/依据`——早先按清单列名读，
    结果 8 条"否"全部读成空单号并塌成 1 行）。金额与 Excel 行号按 (单号, 费用名称)
    去本次核验清单里补。
    """
    if not tiantu_csv or not tiantu_csv.is_file():
        log.warning("[统计表] 没有天图结果 CSV，本次不产出天图问题。")
        return []

    # 本次账单的核验清单：既提供金额/说明，也用来剔掉上一次运行的旧结果
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

    items: list[dict] = []
    seen: set[tuple[str, str]] = set()
    unmatched: list[str] = []
    total = 0
    ok_cnt = 0
    with tiantu_csv.open(encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            total += 1
            if (r.get("是否出现") or "").strip() == "是":
                ok_cnt += 1
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
            items.append({
                "no": no,
                "fee": fee,
                "amount": cl.get("金额USD", ""),
                "remark": "天图未查到标记：" + mark + (f"（缺 {why}）" if why else ""),
            })
    if unmatched:
        log.warning("[统计表] 天图结果里 %d 条不在本次核验清单内，已忽略：%s",
                    len(unmatched), "、".join(unmatched[:8]))
    log.info("[统计表] 天图结果 %d 行（查到 %d / 未查到 %d），登记未查到 %d 条",
             total, ok_cnt, total - ok_cnt, len(items))
    return items


# 独立《打单问题统计表》里各客户 sheet 的列序/列数都不一样
# （环洋 运单号|金额USD|费用名称、中盟 费用名称|金额USD|运单号、永达 多一个备注），
# 所以一律按表头名定位，绝不写死列号。
STATS_ALIASES = {
    "no":     ["运单号", "客户单号", "系统单号"],
    "amount": ["金额USD", "费用金额", "金额"],
    "fee":    ["费用名称", "费用类别"],
    "remark": ["备注", "说明"],
}


def write_to_stats(stats_file, sheet_name: str, items: list[dict], dry_run: bool,
                   log: logging.Logger) -> bool:
    """把「天图未查到标记」的问题追加到独立《打单问题统计表》的目标 sheet。

    目标 sheet 的列序/列数按表头名适配（见 STATS_ALIASES），按 (运单号, 费用名称) 去重；
    重复行若原备注为空则补写备注。**表里没有「备注」列时自动补一个**（加在表头行末尾），
    否则「天图未查到 偏远」这种关键信息会随备注一起被丢掉，人工看不出为什么登记这行。
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
            log.error("[统计表] sheet「%s」表头缺列（至少要 运单号 与 费用名称）：%s",
                      sheet_name, [c.value for c in ws[hrow]])
            return False
        if "remark" not in colmap and any(it.get("remark") for it in items):
            last = max((c.column for c in ws[hrow] if c.value not in (None, "")),
                       default=0)
            ws.cell(row=hrow, column=last + 1).value = "备注"
            colmap["remark"] = last + 1
            log.info("[统计表] sheet「%s」原本没有「备注」列，已在表头行第 %d 列补上",
                     sheet_name, last + 1)
        added, skipped, filled = excel.append_rows(ws, items, colmap)
        saved, degraded = excel.safe_save(wb, p)
    finally:
        wb.close()
    log.info("[统计表] %s「%s」新增 %d 行 / 重复跳过 %d 行 / 补备注 %d 行%s",
             Path(saved).name, sheet_name, added, skipped, filled,
             "（原文件被占用，已写副本）" if degraded else "")
    return True


def _drop_legacy_stats_sheet(bill_path: Path, log: logging.Logger) -> int:
    """删掉账单里由旧流程写入的「打单问题统计」sheet。

    旧流程把统计表直接建在账单里，而且读错了天图结果的列名（塌成 1 行空单号），
    又在封顶修复之前跑过，留下了一批"其实没有差异"的行。统计表现在写独立文件，
    账单里不该再有这张表。账单在 Excel 里开着时删不掉，只告警不阻断。
    """
    if not bill_path or not bill_path.is_file():
        return 0
    import openpyxl  # noqa: PLC0415
    try:
        wb = openpyxl.load_workbook(bill_path)
    except Exception as exc:  # noqa: BLE001
        log.warning("[清理] 打不开账单，跳过旧统计 sheet 清理：%s", exc)
        return 0
    try:
        targets = [n for n in wb.sheetnames
                   if n == "打单问题统计" or n.startswith("打单问题统计-")]
        if not targets:
            return 0
        for n in targets:
            del wb[n]
        wb.save(bill_path)
    except PermissionError:
        log.warning("[清理] 账单正被 Excel 占用，删不掉旧 sheet「打单问题统计」，请关闭 Excel 后重跑。")
        return 0
    except Exception as exc:  # noqa: BLE001
        log.warning("[清理] 删除旧统计 sheet 失败：%s", exc)
        return 0
    finally:
        wb.close()
    log.info("[清理] 已从账单删除旧 sheet：%s", "、".join(targets))
    return len(targets)


# ============== 通知 ==============

def notify_all(cfg: dict, status: str, start_ts: float, stage_reports: list[dict], note: str = ""):
    _add_autolib_path()
    try:
        from autolib.notify import notify_all as _notify_all, build_summary  # noqa: PLC0415
    except Exception:  # noqa: BLE001
        return
    nt = cfg.get("notify", {}) or {}
    app_name = nt.get("app_name", "SKYE打单脚本")
    content = build_summary(app_name, status, start_ts, stages=stage_reports, note=note)
    _notify_all(cfg, content, status=status, start_ts=start_ts, verbose=True)


# ============== main ==============

def main():
    ap = argparse.ArgumentParser(description="SKYE 打单数据整理全流程")
    ap.add_argument("--input", default=None, help="账单 xlsx（不传则弹窗选）")
    ap.add_argument("--quote", default=None,
                    help="尾程渠道报价表 xlsx（不传则弹窗选；--no-dialog 时用配置值）")
    ap.add_argument("--stats", default=None,
                    help="《打单问题统计表》xlsx（不传则弹窗确认；--no-dialog 时用配置值）")
    ap.add_argument("--no-dialog", action="store_true", help="全程不弹窗，一律用参数/配置里的路径")
    ap.add_argument("--skip-engine", action="store_true")
    ap.add_argument("--skip-tiantu", action="store_true")
    ap.add_argument("--skip-write", action="store_true")
    ap.add_argument("--tiantu-limit", type=int, default=0)
    ap.add_argument("--dry-run", action="store_true")
    args = ap.parse_args()

    cfg = _load_notify_local(_load_cfg())
    log = logging.getLogger("skye-main")
    if not log.handlers:
        logging.basicConfig(level=logging.INFO, format="[%(asctime)s] %(message)s", datefmt="%H:%M:%S")
    out_dir = _default_output_dir()

    bill_arg = args.input
    if not bill_arg and not args.no_dialog:
        _safe_print("\n" + "="*60 + "\n[主流程] 弹出窗口选账单\n" + "="*60)
        bill_arg = choose_bill_via_dialog()
    if not bill_arg:
        _safe_print("未选择账单，流程取消。")
        return

    # 报价表时常更换 → 每次跑都让用户选一次。标题写明「尾程渠道报价表」，
    # 与选账单的弹窗区分开（两个弹窗长得像，标题不区分极易选错）。
    quote_arg = args.quote
    if not quote_arg and not args.no_dialog and not args.skip_engine:
        _safe_print("\n" + "="*60 + "\n[主流程] 弹出窗口选报价表\n" + "="*60)
        quote_arg = choose_quote_via_dialog()
        if not quote_arg:
            _safe_print("未选择报价表，流程取消。")
            return
    if quote_arg:
        _safe_print(f"[主流程] 报价表：{quote_arg}")

    start_ts = time.time()
    stages: list[dict] = []
    status = "成功"
    note = ""
    stats_used: str | None = None

    try:
        if not args.skip_engine:
            eng_args = ["--input", bill_arg]
            if quote_arg:
                eng_args += ["--quote", quote_arg]
            code, secs = run_stage("核价引擎", HERE / "run_skye_check.py", eng_args)
            stages.append({"name": "核价引擎", "secs": secs, "ok": code == 0})
            if code != 0:
                status = "失败"
                note = f"核价引擎退出码 {code}"
                raise SystemExit(code)
            if not args.dry_run:
                # 旧流程往账单里塞的那张「打单问题统计」表要清掉（统计表现在写独立文件）
                _drop_legacy_stats_sheet(Path(bill_arg), log)

        # 只认本次引擎生成的清单：否则引擎没跑成时会拿上一次的清单去查天图
        checklist = (_latest(out_dir, "*_天图核验清单_*.csv") if args.skip_engine
                     else _latest_since(out_dir, "*_天图核验清单_*.csv", start_ts))
        if checklist is None:
            log.warning("[天图] 本次没有生成核验清单，天图核验与统计表登记都跳过。")

        tiantu_result: Path | None = None
        if not args.skip_tiantu:
            if checklist:
                _safe_print(f"\n{'='*60}\n[主流程] 天图核验: {_engine_cmd(cfg)[0].name}"
                            f"（清单 {checklist.name}）\n{'='*60}")
                t0 = time.time()
                tiantu_result = run_tiantu_check(cfg, checklist, out_dir, args.tiantu_limit, log)
                secs = round(time.time() - t0, 2)
                _safe_print(f"[主流程] 天图核验 用时 {secs}s（{'完成' if tiantu_result else '无结果'}）")
                stages.append({"name": "天图核验", "secs": secs, "ok": tiantu_result is not None})
                if tiantu_result is None:
                    # 天图失败不阻断：但**不要**退回旧结果，否则统计表会记上早已处理完的单号
                    if status == "成功":
                        status = "部分成功"
                    note = (note + "；" if note else "") + "天图核验未出结果（多为网络/登录失败）"
            else:
                log.warning("[天图] 未找到核验清单，跳过。")
                stages.append({"name": "天图核验", "secs": 0, "ok": False})

        if not args.skip_write:
            if tiantu_result is None and args.skip_tiantu:
                # --skip-tiantu 是明确要求复用上一次的结果；结果仍按本次清单过滤
                eng, _ = _engine_cmd(cfg)
                tiantu_result = find_tiantu_result(eng, out_dir) if eng.is_file() else None
                if tiantu_result:
                    log.warning("[统计表] --skip-tiantu：复用上一次的天图结果 %s（按本次核验清单过滤）",
                                tiantu_result.name)
            missing = collect_tiantu_missing(checklist, tiantu_result, log)

            if args.dry_run:
                _safe_print(f"[dry-run] 拟写入《打单问题统计表》{len(missing)} 条：")
                for it in missing[:10]:
                    _safe_print(f"  {it}")
                stages.append({"name": "写入统计表", "secs": 0, "ok": True})
            else:
                stats_sheet = cfg.get("paths", {}).get("stats_sheet", "skye")
                if not missing:
                    # 没有要登记的就别弹窗打扰用户
                    log.info("[统计表] 本次没有「天图未查到标记」的问题，不弹窗、不写入。")
                    stages.append({"name": "写入统计表", "secs": 0, "ok": True})
                else:
                    # 统计表和报价表一样经常换（不同月份/不同客户各一份），所以每次都要
                    # 让用户选一次；--stats 或 --no-dialog 才走配置里的默认文件。
                    stats_used = args.stats or None
                    if not stats_used and not args.no_dialog:
                        _safe_print("\n" + "="*60 + "\n[主流程] 弹出窗口选统计表\n" + "="*60)
                        stats_used = choose_stats_via_dialog(cfg)
                    if not stats_used:
                        stats_used = cfg.get("paths", {}).get("stats_file")
                        if stats_used:
                            reason = "未选统计表" if not args.no_dialog else "--no-dialog"
                            _safe_print(f"[主流程] {reason}，按配置写：{stats_used}")
                    ok = write_to_stats(stats_used, stats_sheet, missing, False, log)
                    stages.append({"name": "写入统计表", "secs": 0, "ok": ok})
                    if not ok and status == "成功":
                        status = "部分成功"
                        note = (note + "；" if note else "") + "统计表写入失败"

        _safe_print("\n" + "="*60 + "\n[主流程] 全部完成\n" + "="*60)
        _safe_print(f"  原表: {bill_arg}（已加工，含 报价差异 sheet）")
        for label, pat in [
            ("天图核验清单",     "*_天图核验清单_*.csv"),
            ("报告",             "*_报告_*.txt"),
        ]:
            f = _latest(out_dir, pat)
            _safe_print(f"  {label}: {f if f else '未生成'}")
        # 结果 CSV 由天图引擎写到它自己的输出目录，两个目录都找。
        # 只认**本次清单之后**产出的结果：发布包/上一轮的 outputs/ 里可能还躺着旧的
        # 天图核验结果_*.csv，不带 newer_than 就会把旧文件当成本次结果打出来（2026-09-17 现场）
        eng, _ = _engine_cmd(cfg)
        since = 0.0
        if checklist and Path(checklist).is_file():
            since = Path(checklist).stat().st_mtime - 1
        tr = find_tiantu_result(eng, out_dir, newer_than=since) if eng.is_file() else None
        _safe_print(f"  天图核验结果: {tr if tr else '未生成'}")
        _safe_print(f"  打单问题统计表: {stats_used or '（本次未写入）'}"
                    f"（sheet {cfg.get('paths', {}).get('stats_sheet', 'skye')}）")
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001
        status = "失败"
        note = f"主流程异常: {exc}"
        raise
    finally:
        notify_all(cfg, status, start_ts, stages, note)


if __name__ == "__main__":
    main()
