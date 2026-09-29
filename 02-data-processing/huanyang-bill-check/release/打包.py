# -*- coding: utf-8 -*-
"""一键打包环洋工具（EXE 版）。

    python 打包.py            # 打包 + 组装发布包
    python 打包.py --zip      # 顺便压成 zip

做的事：
  1) 用三个 spec 各打一份 onedir（HuanyangTool 启动器 / main_flow 主流程 / check_tiantu 天图引擎）；
  2) 按 SKYE 那套布局组装发布包：
        环洋账单整理工具_EXE版/
          HuanyangTool/          ← 图形启动器
          main_flow/             ← 主流程
          check_tiantu/          ← 天图引擎（永达那份）
          config.json            已按发布根相对化
          credentials.local.json  从永达那份里**只取 tiantu_***（见下）
          notify.local.json
          打单费用名称中英文翻译.xlsx
          outputs/ logs/  启动环洋工具.bat  使用说明.md

两个容易踩的点：
  - **credentials.local.json 必须放在发布根**，不是 check_tiantu/ 里。
    引擎（check_tiantu → run_yongda_check.load_credentials）冻结态取
    `sys.executable` 的父目录的父目录当根，也就是发布根。放错地方的表现是
    「点了开始、浏览器没起来、日志说没凭证」。
  - 发布根的 config.json 同时被主流程和引擎读（引擎只取 paths.output_dir），
    所以 output_dir 必须写成相对发布根（"outputs"），写死开发机绝对路径的话
    换台电脑结果 CSV 会落到找不到的地方。
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent              # .../huanyang-bill-check/release
PROJ = HERE.parent                                  # .../huanyang-bill-check
BUILD = HERE / "build"
DIST = HERE / "dist"
PKG = HERE / "环洋账单整理工具_EXE版"

YONGDA = PROJ.parent / "yongda-bill-check"
SPECS = ["HuanyangTool", "main_flow", "check_tiantu"]


def _log(msg):
    print(f"[打包] {msg}", flush=True)


def pyinstaller() -> list:
    """找 PyInstaller。本机 venv 里没装，装在 Anaconda 那份里。"""
    exe = Path(r"D:/Anaconda/Scripts/pyinstaller.exe")
    if exe.is_file():
        return [str(exe)]
    return [sys.executable, "-m", "PyInstaller"]


def build_one(name: str) -> Path:
    BUILD.mkdir(parents=True, exist_ok=True)
    DIST.mkdir(parents=True, exist_ok=True)
    cmd = [*pyinstaller(), "--noconfirm", "--clean",
           "--distpath", str(DIST), "--workpath", str(BUILD),
           str(BUILD / f"{name}.spec")]
    _log(f"打包 {name} …")
    r = subprocess.run(cmd, cwd=str(BUILD), text=True,
                       encoding="utf-8", errors="replace")
    if r.returncode != 0:
        raise SystemExit(f"打包 {name} 失败（退出码 {r.returncode}）")
    out = DIST / name
    if not out.is_dir():
        raise SystemExit(f"打包 {name} 没产出目录：{out}")
    return out


def write_credentials(dst: Path) -> str:
    """发布包只需要天图账号。

    直接从永达那份里取 tiantu_user/tiantu_pass，而不是整份拷过来 ——
    永达那份还带着打单系统账号，环洋流程根本不用，没必要跟着一起发出去。
    """
    src = YONGDA / "credentials.local.json"
    if not src.is_file():
        _log(f"没找到 {src}，留一份模板，请手工填写后再发")
        shutil.copy2(PROJ / "credentials.example.json", dst)
        return "模板"
    d = json.loads(src.read_text(encoding="utf-8"))
    sub = {k: d[k] for k in ("tiantu_user", "tiantu_pass") if d.get(k)}
    if len(sub) < 2:
        _log(f"{src} 里缺 tiantu_user/tiantu_pass，留模板")
        shutil.copy2(PROJ / "credentials.example.json", dst)
        return "模板（不完整）"
    dst.write_text(json.dumps(sub, ensure_ascii=False, indent=2) + "\n",
                   encoding="utf-8")
    return f"已内置（账号 {sub['tiantu_user']}）"


def assemble(dist: dict[str, Path]) -> None:
    PKG.mkdir(parents=True, exist_ok=True)
    for name, src in dist.items():
        dst = PKG / name
        if dst.exists():
            shutil.rmtree(dst)
        shutil.copytree(src, dst)
        _log(f"{name}/ → 发布包")

    # 同名文件：config.json 已按发布根写好，别用开发目录那份覆盖
    for f in ("notify.example.json", "credentials.example.json"):
        shutil.copy2(PROJ / f, PKG / f)

    # 发布包不得携带开发机绝对路径。主流程和天图引擎都从发布根读取这份配置。
    source_cfg = PROJ / "config.json"
    cfg = json.loads(source_cfg.read_text(encoding="utf-8"))
    paths = cfg.setdefault("paths", {})
    paths.update({
        "input_dir": ".",
        "quote_file": "",
        "translation_file": "打单费用名称中英文翻译.xlsx",
        "stats_file": "打单问题统计表（模板）.xlsx",
        "output_dir": "outputs",
        "log_dir": "logs",
    })
    (PKG / "config.json").write_text(
        json.dumps(cfg, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    _log("config.json：已改为发布包相对路径（不再引用开发机路径）")

    tr = PROJ / "打单费用名称中英文翻译.xlsx"
    if not tr.is_file():
        tr = Path(r"C:/Users/TT1/Desktop/环洋/打单费用名称中英文翻译.xlsx")
    if tr.is_file():
        shutil.copy2(tr, PKG / tr.name)
        _log(f"{tr.name} → 发布包")
    else:
        _log("⚠ 没找到翻译表，请手工放进发布包（否则翻译列会全空）")

    state = write_credentials(PKG / "credentials.local.json")
    _log(f"credentials.local.json：{state}")

    nf = PROJ / "notify.local.json"
    if nf.is_file():
        shutil.copy2(nf, PKG / "notify.local.json")
        _log("notify.local.json：已内置")
    else:
        _blank_placeholders(PROJ / "notify.example.json", PKG / "notify.local.json")
        _log("notify.local.json：留模板（通知会静默跳过，不影响主流程）")

    for d in ("outputs", "logs"):
        (PKG / d).mkdir(exist_ok=True)


def _blank_placeholders(src: Path, dst: Path) -> None:
    """模板 → notify.local.json 时把「请填写…」清成空串。

    留着那串占位文字的话，autolib 会认为"已配置"（非空就发），于是每次运行都去
    POST 一个带空格的假 URL，日志里冒出一句让人摸不着头脑的
    `URL can't contain control characters`。清空后走的是 `[跳过] 未配置` 那条路，
    用户填上真 key 才发。只动发布包，不改项目里的 example，也不动永达共用的 autolib。
    """
    data = json.loads(src.read_text(encoding="utf-8"))

    def blank(v):
        if isinstance(v, dict):
            return {k: blank(x) for k, x in v.items()}
        if isinstance(v, str) and v.strip().startswith("请填写"):
            return ""
        return v

    dst.write_text(json.dumps(blank(data), ensure_ascii=False, indent=2) + "\n",
                   encoding="utf-8")


def main() -> int:
    ap = argparse.ArgumentParser(description="打包环洋账单整理工具（EXE 版）")
    ap.add_argument("--zip", action="store_true", help="打包完压成 zip")
    ap.add_argument("--only", help="只打某一个（HuanyangTool/main_flow/check_tiantu）")
    ap.add_argument("--no-build", action="store_true", help="跳过 PyInstaller，只重组发布包")
    args = ap.parse_args()

    names = [args.only] if args.only else SPECS
    dist = {}
    if args.no_build:
        for n in names:
            d = DIST / n
            if not d.is_dir():
                raise SystemExit(f"没有现成的 {d}，去掉 --no-build 重打一次")
            dist[n] = d
    else:
        for n in names:
            dist[n] = build_one(n)

    if not args.only:
        assemble(dist)

    _log(f"完成：{PKG}")
    if args.zip:
        z = shutil.make_archive(str(HERE / f"{PKG.name}"), "zip", root_dir=HERE,
                                base_dir=PKG.name)
        _log(f"压缩包：{z}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
