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
    
    # 直接输入账号（不聚焦窗口）
    print("输入账号...")
    pyautogui.write(config["account"])
    time.sleep(3)
    print("按 Tab...")
    pyautogui.press('tab')
    time.sleep(3)
    print("输入密码...")
    pyautogui.write(config["password"])
    time.sleep(3)
    print("按回车...")
    pyautogui.press('enter')
    
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
    
    # 如果是周末，下个执行日是周一
    if weekday >= 5:
        days_until = 7 - weekday
        next_time = now + datetime.timedelta(days=days_until)
        return next_time.replace(hour=8, minute=0, second=0, microsecond=0)
    
    # 工作日
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
    
    # 更新顶部时间显示
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
    if messagebox.askyesno("确认", "立即执行登录操作？"):
        threading.Thread(target=do_login, daemon=True).start()
        messagebox.showinfo("提示", "登录操作已执行！")


def create_ui():
    """创建界面"""
    global root, status_label, current_time_value, remaining_value
    
    root = tk.Tk()
    root.title("QMT 自动登录工具")
    root.geometry("420x400")
    root.minsize(420, 400)
    
    # 主容器
    main_frame = ttk.Frame(root, padding="15")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 标题
    title_label = ttk.Label(main_frame, text="QMT 自动登录工具", font=("微软雅黑", 14, "bold"))
    title_label.pack(pady=(0, 8))
    
    # 时间显示区域
    time_frame = ttk.Frame(main_frame)
    time_frame.pack(fill=tk.X, pady=(0, 10))
    
    ttk.Label(time_frame, text="当前:", font=("微软雅黑", 11)).pack(side=tk.LEFT)
    current_time_value = ttk.Label(time_frame, text="--:--:--", font=("微软雅黑", 11, "bold"), foreground="#1976d2")
    current_time_value.pack(side=tk.LEFT, padx=(3, 12))
    
    ttk.Label(time_frame, text="剩余:", font=("微软雅黑", 11)).pack(side=tk.LEFT)
    remaining_value = ttk.Label(time_frame, text="--小时--分--秒", font=("微软雅黑", 11, "bold"), foreground="#d32f2f")
    remaining_value.pack(side=tk.LEFT, padx=(3, 0))
    
    # 配置区域
    config_frame = ttk.LabelFrame(main_frame, text="配置", padding="6")
    config_frame.pack(fill=tk.X, pady=(0, 6))
    
    # 账号
    ttk.Label(config_frame, text="账号:").grid(row=0, column=0, sticky=tk.W, pady=2)
    account_var = tk.StringVar(value=load_config().get("account", ""))
    ttk.Entry(config_frame, textvariable=account_var, width=30).grid(row=0, column=1, sticky=tk.W, padx=(5, 0), pady=2)
    
    # 密码
    ttk.Label(config_frame, text="密码:").grid(row=1, column=0, sticky=tk.W, pady=2)
    password_var = tk.StringVar(value=load_config().get("password", ""))
    ttk.Entry(config_frame, textvariable=password_var, show="*", width=30).grid(row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)
    
    # 程序路径
    ttk.Label(config_frame, text="QMT路径:").grid(row=2, column=0, sticky=tk.W, pady=2)
    path_var = tk.StringVar(value=load_config().get("program_path", ""))
    ttk.Entry(config_frame, textvariable=path_var, width=30).grid(row=2, column=1, sticky=tk.W, padx=(5, 0), pady=2)
    
    # 启动延迟
    ttk.Label(config_frame, text="延迟(秒):").grid(row=3, column=0, sticky=tk.W, pady=2)
    delay_var = tk.StringVar(value=load_config().get("delay_seconds", "3"))
    ttk.Spinbox(config_frame, from_=1, to=30, textvariable=delay_var, width=8).grid(row=3, column=1, sticky=tk.W, padx=(5, 0), pady=2)
    
    def save_settings():
        config = {
            "account": account_var.get(),
            "password": password_var.get(),
            "program_path": path_var.get(),
            "delay_seconds": delay_var.get()
        }
        save_config(config)
        messagebox.showinfo("成功", "配置已保存！")
    
    # 按钮区域
    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill=tk.X, pady=(0, 6))
    
    ttk.Button(btn_frame, text="保存配置", command=save_settings).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_frame, text="立即测试", command=manual_run).pack(side=tk.LEFT)
    
    # 定时控制区域
    schedule_frame = ttk.LabelFrame(main_frame, text="定时任务 (周一至周五 08:00)", padding="6")
    schedule_frame.pack(fill=tk.X, pady=(0, 6))
    
    btn_scheduler_frame = ttk.Frame(schedule_frame)
    btn_scheduler_frame.pack()
    
    ttk.Button(btn_scheduler_frame, text="启动定时", command=start_scheduler).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_scheduler_frame, text="停止定时", command=stop_scheduler).pack(side=tk.LEFT)
    
    # 状态显示
    status_label = ttk.Label(schedule_frame, text="⏸️ 已停止", font=("微软雅黑", 12, "bold"), foreground="#666666")
    status_label.pack(pady=(6, 0))
    
    # 底部提示
    tip_label = ttk.Label(main_frame, text="提示: 确保 QMT 程序路径正确，账号密码已保存", 
                          font=("微软雅黑", 8), foreground="#888888")
    tip_label.pack(side=tk.BOTTOM, pady=(4, 0))
    
    # 启动时间更新循环
    def update_loop():
        update_status()
        root.after(1000, update_loop)
    
    update_loop()
    
    root.mainloop()


if __name__ == "__main__":
    pyautogui.FAILSAFE = True
    create_ui()
