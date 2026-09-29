# -*- mode: python ; coding: utf-8 -*-
"""
环洋账单核价工具 PyInstaller 打包配置

用法：
    pyinstaller build.spec

生成目录结构：
    release/环洋对账工具_EXE版/
    ├── 环洋对账工具.exe          # 主程序（带图形界面）
    ├── check_tiantu/
    │   └── check_tiantu.exe     # 天图核验引擎（复用永达）
    ├── config.json
    ├── credentials.example.json
    ├── notify.example.json
    ├── 打单费用名称中英文翻译.xlsx
    ├── outputs/                 # 输出目录
    └── logs/                    # 日志目录
"""

import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files

block_cipher = None
HERE = Path('.').resolve()
PARENT = HERE.parent.parent

# ============ 主程序：环洋对账工具.exe ============

a_main = Analysis(
    ['main_flow.py'],
    pathex=[str(HERE), str(PARENT / '08-common-utils')],
    binaries=[],
    datas=[
        ('config.json', '.'),
        ('credentials.example.json', '.'),
        ('notify.example.json', '.'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell._writer',
        'autolib',
        'autolib.dialog',
        'autolib.excel',
        'autolib.notify',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'pytest',
        'IPython',
        'jupyter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz_main = PYZ(a_main.pure, a_main.zipped_data, cipher=block_cipher)

exe_main = EXE(
    pyz_main,
    a_main.scripts,
    [],
    exclude_binaries=True,
    name='环洋对账工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll_main = COLLECT(
    exe_main,
    a_main.binaries,
    a_main.zipfiles,
    a_main.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='环洋对账工具_EXE版',
)

# ============ 天图引擎：check_tiantu.exe（复用永达）============

yongda_check = PARENT / 'yongda-bill-check' / 'check_tiantu.py'
if yongda_check.exists():
    a_tiantu = Analysis(
        [str(yongda_check)],
        pathex=[str(PARENT / 'yongda-bill-check'), str(PARENT / '08-common-utils')],
        binaries=[],
        datas=[],
        hiddenimports=[
            'openpyxl',
            'playwright',
            'autolib',
        ],
        hookspath=[],
        hooksconfig={},
        runtime_hooks=[],
        excludes=[
            'matplotlib',
            'numpy',
            'pandas',
            'pytest',
            'IPython',
            'jupyter',
        ],
        win_no_prefer_redirects=False,
        win_private_assemblies=False,
        cipher=block_cipher,
        noarchive=False,
    )

    pyz_tiantu = PYZ(a_tiantu.pure, a_tiantu.zipped_data, cipher=block_cipher)

    exe_tiantu = EXE(
        pyz_tiantu,
        a_tiantu.scripts,
        [],
        exclude_binaries=True,
        name='check_tiantu',
        debug=False,
        bootloader_ignore_signals=False,
        strip=False,
        upx=True,
        console=True,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        icon=None,
    )

    coll_tiantu = COLLECT(
        exe_tiantu,
        a_tiantu.binaries,
        a_tiantu.zipfiles,
        a_tiantu.datas,
        strip=False,
        upx=True,
        upx_exclude=[],
        name='check_tiantu',
    )
