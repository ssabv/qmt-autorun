"""
打包 QMT 自动登录工具 为 exe
"""
import os
import sys
import subprocess

def build_exe():
    print("正在安装依赖...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-r", "requirements.txt", "-q"])
    
    print("正在安装 PyInstaller...")
    subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller", "-q"])
    
    print("正在打包为 exe...")
    # 使用 --noconsole 隐藏控制台窗口（GUI程序不需要）
    # 使用 --onefile 打包成单个文件
    cmd = [
        "pyinstaller",
        "--onefile",
        "--noconsole",
        "--name", "QMT自动登录工具",
        "--icon", "icon.ico",  # 可选：添加图标
        "main.py"
    ]
    
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        print("\n✅ 打包成功！")
        print("exe 文件位置: dist/QMT自动登录工具.exe")
    else:
        print("\n❌ 打包失败，请检查错误信息")


if __name__ == "__main__":
    build_exe()
