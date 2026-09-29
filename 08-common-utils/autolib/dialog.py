# -*- coding: utf-8 -*-
"""弹窗选文件。

坑点（仓库里踩过）：Tk 根窗口如果复用或者不销毁，第二次弹窗会卡住不显示。
这里每个弹窗都创建独立的 Tk 根窗口，withdraw 后立刻 destroy，用完即弃。
无图形界面环境（如计划任务后台运行）会返回 None，调用方应能优雅降级。
"""

from __future__ import annotations

import sys
from pathlib import Path

_EXCEL = [("Excel 工作簿", "*.xlsx"), ("所有文件", "*.*")]
_CSV = [("CSV 文件", "*.csv"), ("所有文件", "*.*")]


def _root():
    import tkinter as tk

    root = tk.Tk()
    root.withdraw()
    # 根窗口**一直**保持置顶：filedialog 是根窗口的 owned window，owner 置顶弹窗才会
    # 跟着置顶。子进程（main_flow.exe）不是前台进程，Windows 的前台锁定会把新弹出的
    # 选择框压在用户正在看的 Edge/Excel 底下 —— 用户反馈「弹窗没有弹出」就是这个
    # （2026-09-17 实测：把别的窗口切到前台后弹窗确实生成了，但不在最上层）。
    # 置顶只影响可见性，不抢焦点，用户点一下即可。
    try:
        root.attributes("-topmost", True)
        root.update()
    except Exception:  # noqa: BLE001
        pass
    return root


def _run(fn):
    try:
        root = _root()
    except Exception as exc:  # noqa: BLE001
        print(f"[提示] 无法弹出选择窗口（{exc}）；后台运行时请用命令行参数指定文件")
        return None
    try:
        from tkinter import filedialog

        return fn(filedialog)
    finally:
        try:
            root.destroy()
        except Exception:  # noqa: BLE001
            pass


def ask_open_file(title="请选择文件", filetypes=None, initialdir=None):
    """弹窗选一个已存在的文件；取消返回 None。"""
    return _run(
        lambda fd: fd.askopenfilename(
            title=title,
            filetypes=filetypes or [("所有文件", "*.*")],
            initialdir=str(initialdir) if initialdir else "",
        )
        or None
    )


def ask_open_excel(title="请选择 Excel 文件", initialdir=None):
    return ask_open_file(title, _EXCEL, initialdir)


def ask_open_csv(title="请选择 CSV 文件", initialdir=None):
    return ask_open_file(title, _CSV, initialdir)


def ask_save_as(title="保存为", default_name="output.xlsx", filetypes=None,
                initialdir=None):
    """弹窗选择保存位置；取消返回 None。"""
    return _run(
        lambda fd: fd.asksaveasfilename(
            title=title,
            initialfile=default_name,
            filetypes=filetypes or _EXCEL,
            initialdir=str(initialdir) if initialdir else "",
            defaultextension=Path(default_name).suffix,
        )
        or None
    )


def ask_directory(title="请选择目录", initialdir=None):
    """弹窗选目录；取消返回 None。"""
    return _run(
        lambda fd: fd.askdirectory(
            title=title, initialdir=str(initialdir) if initialdir else ""
        )
        or None
    )


def _error_box(message: str) -> None:
    """「文件不符合要求」提示框。自带根窗口并显式置顶，用完即弃。"""
    try:
        root = _root()
    except Exception:  # noqa: BLE001
        print(f"[提示] {message}")
        return
    try:
        from tkinter import messagebox

        # 必须显式传 parent：上一个根窗口已销毁，_default_root 指向的是那个死对象
        messagebox.showerror("文件不符合要求", message, parent=root)
    except Exception:  # noqa: BLE001
        print(f"[提示] {message}")
    finally:
        try:
            root.destroy()
        except Exception:  # noqa: BLE001
            pass


def pick_until_valid(title, filetypes=None, initialdir=None, validator=None,
                     max_tries=10, on_error=None):
    """反复弹窗直到选到合法文件。

    :param validator: 接收路径，合法返回 None，不合法返回可展示的错误说明
    :param on_error: 错误提示回调 (message)；缺省用 messagebox.showerror
    :return: 合法路径，或用户取消/超过次数后 None
    """
    for _ in range(max_tries):
        path = ask_open_file(title, filetypes, initialdir)
        if not path:
            return None
        if validator is None:
            return path
        err = validator(path)
        if err is None:
            return path
        if on_error:
            on_error(err)
        else:
            _error_box(err)
    print("已连续多次选择不符合要求的文件，取消操作。")
    return None


def confirm(message: str, title="请确认") -> bool:
    """弹窗确认；无 GUI 环境返回 False（宁可不执行危险操作）。"""
    try:
        from tkinter import messagebox

        root = _root()
    except Exception:  # noqa: BLE001
        print(f"[提示] 无法弹出确认框：{message}")
        return False
    try:
        return bool(messagebox.askyesno(title, message, parent=root))
    except Exception:  # noqa: BLE001
        print(f"[提示] 无法弹出确认框：{message}")
        return False
    finally:
        try:
            root.destroy()
        except Exception:  # noqa: BLE001
            pass


def alert(message: str, title="提示") -> None:
    try:
        from tkinter import messagebox

        root = _root()
    except Exception:  # noqa: BLE001
        print(f"[提示] {message}")
        return
    try:
        messagebox.showinfo(title, message, parent=root)
    except Exception:  # noqa: BLE001
        print(f"[提示] {message}")
    finally:
        try:
            root.destroy()
        except Exception:  # noqa: BLE001
            pass


def has_gui() -> bool:
    """当前环境能否弹窗（计划任务/远程会话下可能为 False）。"""
    if sys.platform != "win32":
        import os

        return bool(os.environ.get("DISPLAY"))
    return True
