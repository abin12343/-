@echo off
chcp 65001 >nul
echo ========================================
echo 环洋账单核价工具 - 打包发布
echo ========================================
echo.

REM 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo 错误：未找到 Python，请先安装 Python 3.8+
    pause
    exit /b 1
)

REM 检查 PyInstaller
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo 安装 PyInstaller...
    pip install pyinstaller
)

REM 执行打包
python pack.py
if errorlevel 1 (
    echo.
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo 打包成功！
echo 发布文件位于：release\环洋对账工具_EXE版\
echo.
pause
