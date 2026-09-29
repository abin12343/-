# -*- coding: utf-8 -*-
"""路径与文件命名工具。

约定：
  - 输出统一写到 config.paths.output_dir，缺省用脚本同级的 outputs/
  - 产物文件名带时间戳，避免相互覆盖
  - 需要串联多个脚本时，用 latest() 按 mtime 取最新产物
"""

from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path

STAMP_FMT = "%Y%m%d_%H%M%S"


def base_dir(depth=0):
    """当前脚本/EXE 所在目录；depth 用于向上回退层级（模板在子目录时用）。

    PyInstaller 打包后 sys.executable 位于 <dist>/<name>/<name>.exe，
    因此回退一级作为基准目录。
    """
    if getattr(sys, "frozen", False):
        p = Path(sys.executable).resolve().parent
        return p.parent if depth == 0 else p.parents[depth]
    here = Path(__file__).resolve()
    return here.parents[depth + 1]


def repo_root():
    """仓库根目录（08-common-utils 的上一级）。找不到时返回 base_dir()。"""
    p = base_dir()
    for parent in [p, *p.parents]:
        if (parent / "README.md").exists() and (parent / "02-data-processing").exists():
            return parent
    return p


def ensure_dir(path) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


def output_dir(config: dict | None = None, fallback=None) -> Path:
    """取输出目录：优先 config['paths']['output_dir']，否则 outputs/。"""
    try:
        raw = (config or {})["paths"]["output_dir"]
        p = Path(raw)
        return p if p.is_absolute() else (base_dir() / p)
    except Exception:  # noqa: BLE001
        return Path(fallback) if fallback else base_dir() / "outputs"


def log_dir(config: dict | None = None, fallback=None) -> Path:
    """取日志目录：优先 config['paths']['log_dir']，否则仓库根 logs/。"""
    try:
        raw = (config or {})["paths"]["log_dir"]
        p = Path(raw)
        return p if p.is_absolute() else (base_dir() / p)
    except Exception:  # noqa: BLE001
        return Path(fallback) if fallback else repo_root() / "logs"


def stamp() -> str:
    """当前时间戳字符串，用于文件名。"""
    return datetime.now().strftime(STAMP_FMT)


def stamped_name(prefix: str, suffix: str, ext: str, ts: str | None = None) -> str:
    """拼出带时间戳的文件名，如 TTTX_UPS_核价结果_20260909_184000.xlsx。"""
    t = ts or stamp()
    ext = ext.lstrip(".")
    parts = [p for p in (prefix, suffix, t) if p]
    return "_".join(parts) + (f".{ext}" if ext else "")


def latest(directory, pattern: str) -> Path | None:
    """按修改时间取目录里最新的匹配文件；没有返回 None。"""
    d = Path(directory)
    if not d.exists():
        return None
    cands = sorted(d.glob(pattern), key=lambda f: f.stat().st_mtime, reverse=True)
    return cands[0] if cands else None


def safe_name(name: str) -> str:
    """把字符串清理成可用作文件名的形式（去掉 Windows 非法字符）。"""
    return re.sub(r'[\\/:*?"<>|\r\n\t]', "_", str(name)).strip().strip(".")


def backup_path(path) -> Path:
    """生成不冲突的备份路径：<原名>_备份.xlsx / _备份1.xlsx ..."""
    p = Path(path)
    candidate = p.with_name(f"{p.stem}_备份{p.suffix}")
    i = 1
    while candidate.exists():
        candidate = p.with_name(f"{p.stem}_备份{i}{p.suffix}")
        i += 1
    return candidate
