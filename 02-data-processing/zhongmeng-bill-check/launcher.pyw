# -*- coding: utf-8 -*-
"""中盟账单整理图形启动器。

采用环洋/永达同样的发布结构：启动器只负责显示日志并拉起发布根目录
下的 main_flow/main_flow.exe，所有文件选择和业务处理仍由主流程完成。
"""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import threading
import time
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext

if getattr(sys, "frozen", False):
    BASE = Path(sys.executable).resolve().parent.parent
else:
    BASE = Path(__file__).resolve().parent

APP_TITLE = "中盟账单数据整理"
LOG_FILE = BASE / "运行日志.txt"
OUT_DIR = BASE / "outputs"


class _Null:
    def write(self, *_args):
        pass
    def flush(self):
        pass


if sys.stdout is None:
    sys.stdout = sys.stderr = _Null()


def main_flow_cmd() -> list[str]:
    if getattr(sys, "frozen", False):
        return [str(BASE / "main_flow" / "main_flow.exe")]
    return [sys.executable, str(BASE / "main_flow.py")]


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.proc = None
        self.q: queue.Queue[str] = queue.Queue()
        self.log_lines: list[str] = []
        root.title(APP_TITLE)
        root.geometry("980x620")
        root.minsize(720, 460)
        bar = tk.Frame(root); bar.pack(fill="x", padx=10, pady=(10, 0))
        self.btn_run = tk.Button(bar, text="开始处理账单", width=16, height=2,
                                 command=self.start, font=("Microsoft YaHei", 11, "bold"))
        self.btn_run.pack(side="left")
        tk.Button(bar, text="打开输出文件夹", width=14, height=2,
                  command=self.open_outputs).pack(side="left", padx=8)
        tk.Button(bar, text="复制日志", width=10, height=2,
                  command=self.copy_log).pack(side="left")
        self.status = tk.Label(bar, text="就绪", anchor="e", font=("Microsoft YaHei", 10))
        self.status.pack(side="right")
        tk.Label(root, text=f"工作目录：{BASE}", anchor="w", fg="#555").pack(fill="x", padx=12, pady=(6, 0))
        self.text = scrolledtext.ScrolledText(root, wrap="word", font=("Consolas", 10), state="disabled")
        self.text.pack(fill="both", expand=True, padx=10, pady=8)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.poll()

    def write(self, line: str):
        self.log_lines.append(line)
        self.text.configure(state="normal"); self.text.insert("end", line); self.text.see("end"); self.text.configure(state="disabled")

    def poll(self):
        try:
            while True: self.write(self.q.get_nowait())
        except queue.Empty:
            pass
        if self.proc is not None and self.proc.poll() is not None:
            code = self.proc.returncode; self.proc = None
            self.btn_run.configure(state="normal"); self.status.configure(text=f"结束（退出码 {code}）")
            self.flush_log()
            if code == 0:
                if messagebox.askyesno(APP_TITLE, "处理结束，是否打开输出文件夹？"): self.open_outputs()
            else:
                messagebox.showwarning(APP_TITLE, f"处理结束但退出码为 {code}。\n请查看窗口日志或点击【复制日志】。")
        self.root.after(200, self.poll)

    def flush_log(self):
        try: LOG_FILE.write_text("".join(self.log_lines), encoding="utf-8", errors="replace")
        except OSError: pass

    def start(self):
        if self.proc is not None: return
        cmd = main_flow_cmd()
        if not Path(cmd[0]).is_file():
            messagebox.showerror(APP_TITLE, f"找不到主程序：{cmd[0]}"); return
        self.log_lines.clear(); self.text.configure(state="normal"); self.text.delete("1.0", "end"); self.text.configure(state="disabled")
        self.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] 启动主流程\n")
        self.write(f"[启动器] 工作目录：{BASE}\n\n")
        self.btn_run.configure(state="disabled"); self.status.configure(text="运行中…")
        env = os.environ.copy(); env["PYTHONUTF8"] = "1"; env["PYTHONIOENCODING"] = "utf-8"
        try:
            self.proc = subprocess.Popen(cmd, cwd=str(BASE), stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                                         text=True, encoding="utf-8", errors="replace", bufsize=1, env=env)
        except Exception as exc:
            self.proc = None; self.btn_run.configure(state="normal"); self.status.configure(text="启动失败")
            messagebox.showerror(APP_TITLE, f"启动失败：{exc}"); return
        threading.Thread(target=self._pump, daemon=True).start()

    def _pump(self):
        proc = self.proc
        if proc is None or proc.stdout is None: return
        for line in proc.stdout: self.q.put(line)
        proc.wait()

    def open_outputs(self):
        try: OUT_DIR.mkdir(parents=True, exist_ok=True); os.startfile(str(OUT_DIR))
        except Exception as exc: messagebox.showerror(APP_TITLE, f"打不开输出文件夹：{exc}")

    def copy_log(self):
        self.root.clipboard_clear(); self.root.clipboard_append("".join(self.log_lines)); self.status.configure(text="日志已复制")

    def on_close(self):
        if self.proc is not None and self.proc.poll() is None:
            if not messagebox.askyesno(APP_TITLE, "任务还在运行，确定要退出吗？"): return
            try: self.proc.kill()
            except OSError: pass
        self.flush_log(); self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk(); App(root); root.mainloop()
