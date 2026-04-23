@echo off
chcp 65001 >nul
echo ====================================
echo   QMT自动登录工具 - 一键打包
echo ====================================
echo.

:: 检查Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ 未检测到 Python，请先安装 Python 3.8+
    echo    下载地址: https://www.python.org/downloads/
    pause
    exit
)

echo ✅ 检测到 Python 环境
echo.

:: 安装依赖
echo 📦 正在安装依赖...
pip install -r requirements.txt -q
pip install pyinstaller -q

echo.

:: 打包
echo 🔨 正在打包为 exe...
pyinstaller --onefile --noconsole --name "QMT自动登录工具" main.py

if exist "dist\QMT自动登录工具.exe" (
    echo.
    echo ====================================
    echo   ✅ 打包成功！
    echo   📍 exe 位置: dist\QMT自动登录工具.exe
    echo ====================================
    echo.
    echo 按任意键打开所在目录...
    explorer dist
) else (
    echo.
    echo ❌ 打包失败，请检查错误信息
)

pause
