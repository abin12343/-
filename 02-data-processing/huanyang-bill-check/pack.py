#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
环洋账单核价工具打包脚本

自动执行：
1. 运行 PyInstaller
2. 复制必需文件到发布目录
3. 整理目录结构
4. 生成压缩包

用法：
    python pack.py
"""

import shutil
import subprocess
import sys
import zipfile
from datetime import datetime
from pathlib import Path

HERE = Path(__file__).resolve().parent
PARENT = HERE.parent.parent
RELEASE_DIR = HERE / "release"
SPEC_FILE = HERE / "build.spec"


def log(msg):
    print(f"[{datetime.now():%H:%M:%S}] {msg}", flush=True)


def run_pyinstaller():
    """运行 PyInstaller 打包"""
    log("开始打包...")
    if not SPEC_FILE.exists():
        log(f"错误：找不到 {SPEC_FILE}")
        return False

    cmd = [sys.executable, "-m", "PyInstaller", str(SPEC_FILE), "--clean", "--noconfirm"]
    log(f"执行：{' '.join(cmd)}")

    result = subprocess.run(cmd, cwd=HERE)
    if result.returncode != 0:
        log("打包失败")
        return False

    log("打包完成")
    return True


def copy_resources(dist_dir):
    """复制必需的资源文件到发布目录"""
    log("复制资源文件...")

    # 创建输出和日志目录
    (dist_dir / "outputs").mkdir(exist_ok=True)
    (dist_dir / "logs").mkdir(exist_ok=True)

    # 复制翻译表（如果存在）
    translation_files = [
        HERE / "打单费用名称中英文翻译.xlsx",
        PARENT / "打单费用名称中英文翻译.xlsx",
    ]
    for src in translation_files:
        if src.exists():
            shutil.copy2(src, dist_dir / "打单费用名称中英文翻译.xlsx")
            log(f"  复制：{src.name}")
            break

    # 复制天图引擎目录
    tiantu_src = HERE / "dist" / "check_tiantu"
    tiantu_dst = dist_dir / "check_tiantu"
    if tiantu_src.exists():
        if tiantu_dst.exists():
            shutil.rmtree(tiantu_dst)
        shutil.copytree(tiantu_src, tiantu_dst)
        log(f"  复制：check_tiantu/")

    # 创建 README.txt
    readme = dist_dir / "使用说明.txt"
    readme.write_text("""环洋账单核价工具使用说明
======================

1. 运行方式
   - 双击"环洋对账工具.exe"启动程序
   - 按提示依次选择：账单1（成本汇总）、账单2（成本明细）、报价表

2. 必需文件
   - 打单费用名称中英文翻译.xlsx（已包含）
   - credentials.local.json（首次使用需配置天图登录凭证）

3. 输出文件
   - outputs/ 目录：核价报告、天图核验清单、天图核验结果
   - 原账单文件会被回填报价列（原文件自动备份）

4. 配置文件
   - config.json：报价规则、费用类别配置
   - credentials.local.json：天图登录凭证（参考 credentials.example.json）
   - notify.local.json：通知配置（可选，参考 notify.example.json）

5. 命令行用法
   环洋对账工具.exe --help

更多信息请参考项目 README.md
""", encoding="utf-8")
    log("  创建：使用说明.txt")


def create_archive(dist_dir):
    """创建 ZIP 压缩包"""
    log("创建压缩包...")

    timestamp = datetime.now().strftime("%Y%m%d")
    zip_name = f"环洋对账工具_EXE版_{timestamp}.zip"
    zip_path = RELEASE_DIR / zip_name

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in dist_dir.rglob("*"):
            if file.is_file():
                arcname = file.relative_to(dist_dir.parent)
                zf.write(file, arcname)

    size_mb = zip_path.stat().st_size / 1024 / 1024
    log(f"压缩包：{zip_name}（{size_mb:.1f} MB）")
    return zip_path


def main():
    log("=" * 60)
    log("环洋账单核价工具打包")
    log("=" * 60)

    # 1. 运行 PyInstaller
    if not run_pyinstaller():
        return 1

    # 2. 定位发布目录
    dist_dir = HERE / "dist" / "环洋对账工具_EXE版"
    if not dist_dir.exists():
        log(f"错误：找不到发布目录 {dist_dir}")
        return 1

    # 3. 复制资源文件
    copy_resources(dist_dir)

    # 4. 移动到 release 目录
    RELEASE_DIR.mkdir(exist_ok=True)
    release_target = RELEASE_DIR / "环洋对账工具_EXE版"
    if release_target.exists():
        shutil.rmtree(release_target)
    shutil.move(str(dist_dir), str(release_target))
    log(f"发布目录：{release_target}")

    # 5. 创建压缩包
    zip_path = create_archive(release_target)

    log("=" * 60)
    log("打包完成！")
    log(f"  发布目录：{release_target}")
    log(f"  压缩包：{zip_path}")
    log("=" * 60)

    return 0


if __name__ == "__main__":
    sys.exit(main())
