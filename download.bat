@echo off
chcp 65001 >nul
echo ====================================
echo   QMT自动登录工具 - 下载器
echo ====================================
echo.
echo 正在下载文件...

:: 下载 base64 文件（将下面的内容粘贴到 base64.txt 同目录下）
powershell -Command "Invoke-WebRequest -Uri 'http://10.0.23.160:9001/qmt-autorun.tar.gz' -OutFile 'qmt-autorun.tar.gz'"

if exist "qmt-autorun.tar.gz" (
    echo.
    echo ✅ 下载成功！正在解压...
    tar -xzvf qmt-autorun.tar.gz
    echo.
    echo ====================================
    echo   ✅ 完成！
    echo   📍 文件已解压到当前目录
    echo   📍 运行 一键打包.bat 开始打包
    echo ====================================
) else (
    echo.
    echo ❌ 下载失败，请手动下载文件后放入当前目录
    echo 手动解码命令: tar -xzvf qmt-autorun.tar.gz
)

pause
