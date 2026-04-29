import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import schedule
import time
import threading
import subprocess
import pyautogui
import datetime

# 尝试导入 pywin32，如果未安装则跳过窗口激活（仅警告）
try:
    import win32gui
    import win32con
    import win32api
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# 确保路径存在
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
else:
    APP_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(APP_DIR, "config.json")

# 全局变量
root = None
status_label = None
current_time_value = None
remaining_value = None
running = False
schedule_thread = None


def load_config():
    """加载配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        "account": "",
        "password": "",
        "program_path": "",
        "delay_seconds": "3"
    }


def save_config(config):
    """保存配置"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def find_qmt_window():
    """遍历所有窗口，找到 QMT 相关窗口"""
    result = {'hwnd': None, 'title': ''}

    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            # QMT/XTP/迅投 是常见关键词，根据实际调整
            keywords = ["QMT", "XTP", "迅投", "极速策略", "QMT交易", "xtp"]
            if any(kw.lower() in title.lower() for kw in keywords):
                result['hwnd'] = hwnd
                result['title'] = title
                return False  # 找到就停止枚举
    try:
        win32gui.EnumWindows(enum_handler, None)
    except Exception:
        pass
    return result['hwnd'], result['title']


def bring_window_to_front(hwnd):
    """强制将窗口置顶并激活"""
    try:
        # 如果窗口最小化了，恢复
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        # 把窗口提到最前
        win32gui.SetForegroundWindow(hwnd)
        # 多次激活确保生效
        win32api.SendMessage(hwnd, win32con.WM_ACTIVATE, 0, 0)
        win32api.SendMessage(hwnd, win32con.WM_ACTIVATESELF, 0, 0)
        time.sleep(2)
        return True
    except Exception as e:
        print(f"窗口激活失败: {e}")
        return False


def do_login():
    """执行登录操作"""
    config = load_config()

    program_path = config["program_path"]
    if not program_path or not os.path.exists(program_path):
        messagebox.showerror("错误", "程序路径无效！")
        return False

    program_name = os.path.basename(program_path)

    # 先杀掉已存在的 QMT 进程
    print(f"检查并关闭已运行的 QMT...")
    os.system(f'taskkill /f /im "{program_name}" 2>nul')
    time.sleep(1)

    # 启动程序
    print(f"启动 {program_path}...")
    subprocess.Popen(f'"{program_path}"', shell=True)

    # 等待程序启动
    delay = int(config.get("delay_seconds", "3"))
    print(f"等待 {delay} 秒...")
    time.sleep(delay)

    # ★★★ 关键修复：用 pywin32 强制把 QMT 窗口置顶 ★★★
    window_found = False
    if HAS_WIN32:
        print("正在查找 QMT 窗口...")
        max_wait = 20
        hwnd, title = None, ""
        for i in range(max_wait):
            hwnd, title = find_qmt_window()
            if hwnd:
                break
            time.sleep(1)
            print(f"  等待窗口出现... ({i+1}/{max_wait})")

        if hwnd:
            print(f"找到窗口: {title}")
            if bring_window_to_front(hwnd):
                print("窗口已激活")
                window_found = True
            else:
                print("窗口激活失败，尝试直接截图调试...")
                try:
                    pyautogui.screenshot(os.path.join(APP_DIR, 'debug_fail.png'))
                    print("调试截图已保存: debug_fail.png")
                except:
                    pass
        else:
            print("未找到 QMT 窗口，保存调试截图...")
            try:
                pyautogui.screenshot(os.path.join(APP_DIR, 'debug_no_window.png'))
                print("调试截图已保存: debug_no_window.png")
            except:
                pass
    else:
        print("警告: 未安装 pywin32，无法自动激活窗口")

    # 延时确保稳定
    time.sleep(1)

    # 输入账号
    print("输入账号...")
    pyautogui.write(config["account"])
    time.sleep(0.5)

    # 按 Tab 跳到密码框
    print("按 Tab...")
    pyautogui.press('tab')
    time.sleep(0.5)

    # 输入密码
    print("输入密码...")
    pyautogui.write(config["password"])
    time.sleep(0.5)

    # 按回车登录
    print("按回车...")
    pyautogui.press('enter')

    print("登录操作完成")
    return True


def run_schedule():
    """运行定时任务"""
    global running
    while running:
        schedule.run_pending()
        time.sleep(1)


def start_scheduler():
    """启动调度器"""
    global running, schedule_thread

    config = load_config()
    if not config["account"] or not config["password"]:
        messagebox.showwarning("警告", "请先填写账号和密码！")
        return

    running = True

    # 清除旧任务
    schedule.clear()

    # 设置周一到周五 8:00 的任务
    schedule.every().monday.at("08:00").do(do_login)
    schedule.every().tuesday.at("08:00").do(do_login)
    schedule.every().wednesday.at("08:00").do(do_login)
    schedule.every().thursday.at("08:00").do(do_login)
    schedule.every().friday.at("08:00").do(do_login)

    # 启动调度线程
    schedule_thread = threading.Thread(target=run_schedule, daemon=True)
    schedule_thread.start()

    update_status()
    messagebox.showinfo("成功", "定时任务已启动！\n周一至周五 8:00 将自动登录。")


def stop_scheduler():
    """停止调度器"""
    global running
    running = False
    schedule.clear()
    update_status()
    messagebox.showinfo("提示", "定时任务已停止。")


def get_next_run_time():
    """计算下次执行时间"""
    now = datetime.datetime.now()
    weekday = now.weekday()

    if weekday >= 5:
        days_until = 7 - weekday
        next_time = now + datetime.timedelta(days=days_until)
        return next_time.replace(hour=8, minute=0, second=0, microsecond=0)

    current_time = now.time()
    target_time = datetime.time(8, 0, 0)

    if current_time < target_time:
        return now.replace(hour=8, minute=0, second=0, microsecond=0)
    else:
        next_time = now + datetime.timedelta(days=1)
        while next_time.weekday() >= 5:
            next_time += datetime.timedelta(days=1)
        return next_time.replace(hour=8, minute=0, second=0, microsecond=0)


def countdown_to_next_run():
    """计算距离下次执行的倒计时"""
    next_run = get_next_run_time()
    now = datetime.datetime.now()
    delta = next_run - now

    if delta.total_seconds() < 0:
        return "已错过，等待下次..."

    hours, remainder = divmod(int(delta.total_seconds()), 3600)
    minutes, seconds = divmod(remainder, 60)
    return f"{hours}小时 {minutes}分 {seconds}秒"


def update_status():
    """更新状态显示"""
    now = datetime.datetime.now()

    current_time_value.config(text=now.strftime('%H:%M:%S'))

    if running:
        countdown = countdown_to_next_run()
        remaining_value.config(text=countdown)
        status_label.config(text="✅ 运行中", foreground="#2e7d32")
    else:
        remaining_value.config(text="未启动")
        status_label.config(text="⏸️ 已停止", foreground="#666666")


def manual_run():
    """手动执行"""
    if messagebox.askyes("确认", "立即执行登录操作？"):
        threading.Thread(target=do_login, daemon=True).start()
        messagebox.showinfo("提示", "登录操作已执行！")


def create_ui():
    """创建界面"""
    global root, status_label, current_time_value, remaining_value

    root = tk.Tk()
    root.title("QMT 自动登录工具 (增强版)")
    root.geometry("460x420")
    root.minsize(460, 420)

    main_frame = ttk.Frame(root, padding="15")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 标题
    title_label = ttk.Label(main_frame, text="QMT 自动登录工具", font=("微软雅黑", 14, "bold"))
    title_label.pack(pady=(0, 5))

    # 副标题
    sub_label = ttk.Label(main_frame, text="(窗口自动聚焦增强版)",
                          font=("微软雅黑", 9), foreground="#888888")
    sub_label.pack(pady=(0, 8))

    # 时间显示
    time_frame = ttk.Frame(main_frame)
    time_frame.pack(fill=tk.X, pady=(0, 10))

    ttk.Label(time_frame, text="当前:",
              font=("微软雅黑", 11)).pack(side=tk.LEFT)
    current_time_value = ttk.Label(time_frame, text="--:--:--",
                                    font=("微软雅黑", 11, "bold"), foreground="#1976d2")
    current_time_value.pack(side=tk.LEFT, padx=(3, 12))

    ttk.Label(time_frame, text="剩余:",
              font=("微软雅黑", 11)).pack(side=tk.LEFT)
    remaining_value = ttk.Label(time_frame, text="--小时--分--秒",
                                font=("微软雅黑", 11, "bold"), foreground="#d32f2f")
    remaining_value.pack(side=tk.LEFT, padx=(3, 0))

    # 配置区
    config_frame = ttk.LabelFrame(main_frame, text="配置", padding="6")
    config_frame.pack(fill=tk.X, pady=(0, 6))

    ttk.Label(config_frame, text="账号:").grid(row=0, column=0, sticky=tk.W, pady=2)
    account_var = tk.StringVar(value=load_config().get("account", ""))
    ttk.Entry(config_frame, textvariable=account_var, width=32).grid(
        row=0, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    ttk.Label(config_frame, text="密码:").grid(row=1, column=0, sticky=tk.W, pady=2)
    password_var = tk.StringVar(value=load_config().get("password", ""))
    ttk.Entry(config_frame, textvariable=password_var, show="*", width=32).grid(
        row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    ttk.Label(config_frame, text="QMT路径:").grid(row=2, column=0, sticky=tk.W, pady=2)
    path_var = tk.StringVar(value=load_config().get("program_path", ""))
    ttk.Entry(config_frame, textvariable=path_var, width=32).grid(
        row=2, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    ttk.Label(config_frame, text="延迟(秒):").grid(row=3, column=0, sticky=tk.W, pady=2)
    delay_var = tk.StringVar(value=load_config().get("delay_seconds", "3"))
    ttk.Spinbox(config_frame, from_=1, to=30, textvariable=delay_var, width=8).grid(
        row=3, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    def save_settings():
        config = {
            "account": account_var.get(),
            "password": password_var.get(),
            "program_path": path_var.get(),
            "delay_seconds": delay_var.get()
        }
        save_config(config)
        messagebox.showinfo("成功", "配置已保存！")

    # 按钮区
    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill=tk.X, pady=(0, 6))

    ttk.Button(btn_frame, text="保存配置", command=save_settings).pack(
        side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_frame, text="立即测试", command=manual_run).pack(side=tk.LEFT)

    # 定时区
    schedule_frame = ttk.LabelFrame(
        main_frame, text="定时任务 (周一至周五 08:00)", padding="6")
    schedule_frame.pack(fill=tk.X, pady=(0, 6))

    btn_scheduler_frame = ttk.Frame(schedule_frame)
    btn_scheduler_frame.pack()

    ttk.Button(btn_scheduler_frame, text="启动定时",
               command=start_scheduler).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_scheduler_frame, text="停止定时",
               command=stop_scheduler).pack(side=tk.LEFT)

    status_label = ttk.Label(schedule_frame, text="⏸️ 已停止",
                             font=("微软雅黑", 12, "bold"), foreground="#666666")
    status_label.pack(pady=(6, 0))

    # 提示
    if not HAS_WIN32:
        warn_label = ttk.Label(main_frame,
                               text="⚠️ 建议运行: pip install pywin32 以启用窗口自动聚焦",
                               font=("微软雅黑", 8), foreground="#e65100")
        warn_label.pack(side=tk.BOTTOM, pady=(4, 0))

    tip_label = ttk.Label(main_frame,
                          text="提示: 调试截图会保存在程序同目录 (debug_*.png)",
                          font=("微软雅黑", 8), foreground="#888888")
    tip_label.pack(side=tk.BOTTOM, pady=(4, 0))

    def update_loop():
        update_status()
        root.after(1000, update_loop)

    update_loop()

    root.mainloop()


if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    create_ui()
