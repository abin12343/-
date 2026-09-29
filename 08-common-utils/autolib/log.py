# -*- coding: utf-8 -*-
"""日志工具。

统一写到 logs/<name>.log（UTF-8），同时输出到控制台。
控制台编码不支持某些字符（如 emoji）时自动降级，不让打印中断流程。
"""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

from . import paths

_FMT = "%(asctime)s %(levelname)s %(message)s"
_DATEFMT = "%Y-%m-%d %H:%M:%S"
_setup_done: set = set()


def setup_logger(
    name: str,
    config: dict | None = None,
    level: int = logging.INFO,
    filename: str | None = None,
    console: bool = True,
) -> logging.Logger:
    """创建（或复用）logger。重复调用同一 name 不会重复挂 handler。"""
    log = logging.getLogger(name)
    log.setLevel(level)
    if name in _setup_done:
        return log

    directory = paths.log_dir(config)
    try:
        directory.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(
            directory / (filename or f"{name}.log"), encoding="utf-8"
        )
        fh.setFormatter(logging.Formatter(_FMT, datefmt=_DATEFMT))
        log.addHandler(fh)
    except Exception as exc:  # noqa: BLE001
        print(f"[警告] 无法写日志文件（{exc}），仅输出到控制台", file=sys.stderr)

    if console:
        ch = logging.StreamHandler(sys.stdout)
        ch.setFormatter(logging.Formatter(_FMT, datefmt=_DATEFMT))
        log.addHandler(ch)

    _setup_done.add(name)
    return log


def safe_print(msg, log: logging.Logger | None = None) -> None:
    """兼容 Windows 控制台编码：打不出就降级，不让流程中断。"""
    text = str(msg)
    try:
        print(text)
    except UnicodeEncodeError:
        enc = getattr(sys.stdout, "encoding", None) or "utf-8"
        try:
            sys.stdout.buffer.write((text + "\n").encode(enc, "replace"))
            sys.stdout.flush()
        except Exception:  # noqa: BLE001
            pass
    if log is not None:
        log.info(text)


class Timer:
    """计时器：用于统计阶段耗时，写进运行报告/通知。"""

    def __init__(self):
        self.start = datetime.now()
        self._last = self.start

    def lap(self) -> float:
        """返回距上次 lap/开始的秒数，并重置计时点。"""
        now = datetime.now()
        secs = (now - self._last).total_seconds()
        self._last = now
        return round(secs, 2)

    @property
    def elapsed(self) -> float:
        return round((datetime.now() - self.start).total_seconds(), 2)

    def __str__(self) -> str:
        return f"{self.elapsed}s"
