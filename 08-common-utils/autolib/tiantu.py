# -*- coding: utf-8 -*-
"""天图系统核验适配器。

仓库现状：
  - `02-data-processing/yongda-bill-check/check_tiantu.py`（1027 行）是稳定可用的
    Playwright 引擎，能直接消费我们生成的《天图核验清单》CSV（字段
    `客户单号(去后缀)` + `期望标记`），输出《天图核验结果_*.csv》。
  - 直接复制 1000+ 行是浪费，把它做成一个薄适配器由新项目调用即可。

调用契约（保持稳定）：
  - `run(checklist_csv, out_csv=None, limit=0, **extra)` —— 生成结果 CSV
    （是/否/未定位/异常），并返回 `Path`（结果文件路径）。
  - 通过环境变量 `TIANTU_USER/TIANTU_PASS` 或配置传入凭据。

环境变量优先于 `extra["env"]` 显式传入的环境。
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path


# 任务清单字段（与 check_tiantu.py 兼容）。保持稳定。
CSV_FIELDS = [
    "Excel行号", "客户单号(去后缀)", "费用类别",
    "费用名称", "期望标记", "金额USD", "说明",
]

# 结果 CSV 字段（与 check_tiantu.py 兼容）
RESULT_FIELDS = [
    "客户单号(去后缀)", "费用类别", "费用名称", "金额USD",
    "期望标记", "是否出现", "依据",
]


def write_checklist(out_path, tasks):
    """把去重后的核验任务写为 check_tiantu.py 能识别的 CSV。

    :param tasks: list[dict]，必须含 no / cat / fee / mark / amount（其余可选）
    """
    import csv  # noqa: PLC0415
    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    with out.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS)
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
    return out


def find_engine(config: dict | None) -> Path:
    """从配置里找天图核验引擎脚本；找不到给明确报错。"""
    if not config:
        raise FileNotFoundError("未配置 tiantu.engine，无法调用天图核验。")
    p = Path(config.get("engine") or "").expanduser()
    if not p.is_file():
        raise FileNotFoundError(
            f"天图核验引擎不存在：{p}。"
            f"可在 config.tiantu.engine 改写为本地路径，"
            f"或运行永达对账模块下的 check_tiantu.py。"
        )
    return p


# 引擎只认这些参数，多传一个就会 argparse 报错退出码 2
_ENGINE_ARGS = ("--csv", "--limit", "--batch", "--url", "--timeout")


def run(checklist_csv, out_csv=None, limit: int = 0,
        config: dict | None = None, env: dict | None = None,
        timeout: int | None = None) -> Path:
    """驱动天图核验引擎子进程；返回结果 CSV 路径。

    注意：引擎**没有** --result 参数，输出目录由它自己的 config.json 决定
    （永达项目是 yongda-bill-check/outputs）。所以 out_csv 只是"期望路径"，
    真正的结果要靠 find_result() 去引擎的输出目录捞。
    """
    engine = find_engine(config)
    out_csv = Path(out_csv) if out_csv else None
    # 开发环境运行 .py 需要当前 Python；发布包中的天图引擎是独立 .exe，
    # 不能再把主流程 ZhongmengTool.exe 当作 Python 解释器调用。
    cmd = ([str(engine)] if engine.suffix.lower() == ".exe"
           else [sys.executable, str(engine)]) + ["--csv", str(checklist_csv)]
    if limit and limit > 0:
        cmd += ["--limit", str(int(limit))]
    if config and config.get("batch"):
        cmd += ["--batch", str(int(config["batch"]))]
    if config and config.get("url"):
        cmd += ["--url", config["url"]]
    if config and config.get("timeout"):
        cmd += ["--timeout", str(int(config["timeout"]))]

    full_env = os.environ.copy()
    if env:
        full_env.update({k: str(v) for k, v in env.items()})
    started = time.time()
    subprocess.run(cmd, check=True, env=full_env, timeout=timeout)

    res = find_result(engine, newer_than=started - 1)
    if res is None:
        raise FileNotFoundError(
            f"天图核验引擎未生成结果 CSV（已搜索 {[str(d) for d in _result_dirs(engine)]}）"
        )
    return res


def _result_dirs(engine: Path) -> list[Path]:
    """结果 CSV 的候选目录：引擎同级 outputs/ + 引擎自己 config.json 的 output_dir。"""
    dirs: list[Path] = []
    cand = engine.parent / "outputs"
    if cand.is_dir():
        dirs.append(cand)
    try:
        cfg = json.loads((engine.parent / "config.json").read_text(encoding="utf-8"))
        p = Path(cfg["paths"]["output_dir"])
        if not p.is_absolute():
            p = engine.parent / p
        if p.is_dir():
            dirs.append(p)
    except Exception:  # noqa: BLE001
        pass
    if not dirs:
        dirs.append(engine.parent)
    seen: set[str] = set()
    out: list[Path] = []
    for d in dirs:
        k = str(d.resolve()).lower()
        if k not in seen:
            seen.add(k)
            out.append(d)
    return out


def find_result(engine, newer_than: float = 0.0) -> Path | None:
    """在引擎的输出目录里找最新的《天图核验结果_*.csv》。"""
    engine = Path(engine)
    best: Path | None = None
    best_m = 0.0
    for d in _result_dirs(engine):
        try:
            for f in d.glob("天图核验结果_*.csv"):
                m = f.stat().st_mtime
                if m < newer_than or m <= best_m:
                    continue
                best, best_m = f, m
        except OSError:
            continue
    return best
