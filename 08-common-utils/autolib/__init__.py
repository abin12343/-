# -*- coding: utf-8 -*-
"""
autolib —— MyScripts 仓库通用自动化工具库

把仓库里反复出现的样板代码收拢成一处，新脚本直接 import 即可：
  config   配置加载（config.json + config.local.json 覆盖）
  log      日志（统一写入 logs/，控制台同步输出）
  paths    路径工具（输出目录、带时间戳文件名、取最新文件）
  dialog   弹窗选文件（独立 Tk 窗口，关闭即销毁）
  notify   通知（企业微信机器人 + 金山文档运行记录）
  browser  Playwright 辅助（复用本机 Edge，延迟导入）
  excel    openpyxl 辅助（表头定位、列映射、安全保存、去重追加）
  tiantu   天图系统核验适配器（薄封装，驱动 02-data-processing/yongda-bill-check/check_tiantu.py）

使用方式（脚本在任意目录时）：

    import sys
    from pathlib import Path
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))   # 指向 08-common-utils
    from autolib import config, log, paths, dialog, notify

依赖按需引入：browser 需要 playwright，excel 需要 openpyxl，未安装时 import 对应
模块会给出明确提示，而不是在运行时才炸。
"""

from . import paths as _paths

__all__ = [
    "base_dir",
    "config",
    "log",
    "paths",
    "dialog",
    "notify",
    "browser",
    "excel",
    "tiantu",
]

__version__ = "0.1.0"


def base_dir():
    """当前脚本/EXE 所在目录。

    源码运行 -> 该脚本文件所在目录；PyInstaller 打包后 -> dist 根目录。
    """
    return _paths.base_dir()
