import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import sys
import schedule
import time
import threading
import datetime

# 核心改进：使用 pywinauto 而不是纯 pyautogui
try:
    from pywinauto.application import Application, ProcessNotFoundError
    HAS_PYWINAUTO = True
except ImportError:
    HAS_PYWINAUTO = False

# pyautogui 用于键盘模拟（窗口激活后用它最可靠）
try:
    import pyautogui
    HAS_PYAUTOGUI = True
    pyautogui.FAILSAFE = True
except ImportError:
    HAS_PYAUTOGUI = False

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
log_text = None  # 日志文本框


def log(msg):
    """写日志到界面"""
    print(msg)
    if log_text:
        log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%H:%M:%S')} {msg}\n")
        log_text.see(tk.END)


def load_config():
    """加载配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            pass
    return {
        "exe_path": "",
        "password": "",
        "delay_seconds": "5",
        "user_id": ""
    }


def save_config(cfg):
    """保存配置"""
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def is_qmt_running(exe_path):
    """检查QMT是否在运行"""
    if not HAS_PYWINAUTO:
        return None
    try:
        app = Application(backend="uia").connect(path=exe_path, timeout=2)
        return app
    except Exception:
        return None


def check_logged_in(app):
    """检查QMT是否已登录"""
    try:
        win = app.top_window()
        title = win.window_text()
        log(f"当前窗口: {title}")

        # 已登录窗口通常有这些关键词
        logged_in_kw = ["委托", "持仓", "交易", "行情", "资产", "策略", "今日"]
        if any(kw in title for kw in logged_in_kw):
            return True

        # 登录窗口通常有这些关键词
        login_kw = ["登录", "密码", "账号", "验证码", "Login", "Password"]
        if any(kw in title for kw in login_kw):
            return False

        # 兜底：如果有密码输入框则是登录界面
        try:
            for child in win.children():
                if 'Edit' in str(child.class_name()) and child.is_visible():
                    text = child.window_text()
                    if '密码' in text or 'password' in text.lower():
                        return False
        except:
            pass

        return True  # 默认认为已登录
    except Exception as e:
        log(f"检查登录状态异常: {e}")
        return False


def bring_window_to_front(app):
    """用 pywinauto 强制把 QMT 窗口置顶"""
    try:
        win = app.top_window()
        log(f"找到窗口: {win.window_text()}")
        # 恢复窗口（如果最小化了）
        if win.is_minimized():
            win.restore()
        # 强制激活
        win.set_focus()
        win.topup()
        time.sleep(1.5)
        log("窗口已激活")
        return True
    except Exception as e:
        log(f"窗口激活失败: {e}")
        return False


def do_login():
    """执行登录操作"""
    config = load_config()
    exe_path = config.get("exe_path", "").strip()
    password = config.get("password", "").strip()

    if not exe_path or not os.path.exists(exe_path):
        log("错误: QMT程序路径无效！")
        return False

    if not password:
        log("错误: 密码不能为空！")
        return False

    log("=" * 40)
    log("开始登录流程...")
    log(f"程序路径: {exe_path}")

    # ===== 步骤1：检查QMT是否已在运行 =====
    app = is_qmt_running(exe_path)
    if app:
        log("检测到QMT已在运行")
        if check_logged_in(app):
            log("QMT已登录，无需重复登录")
            return True
        log("QMT未登录，准备关闭重启...")
        try:
            app.kill()
            log("已关闭旧进程")
            time.sleep(2)
        except Exception as e:
            log(f"关闭旧进程失败: {e}")

    # ===== 步骤2：启动QMT（用 pywinauto） =====
    log("正在启动QMT...")
    try:
        if not HAS_PYWINAUTO:
            log("错误: 未安装 pywinauto，无法启动程序")
            return False
        app = Application(backend="uia").start(exe_path, timeout=15)
        log("QMT进程已启动")
    except Exception as e:
        log(f"启动QMT失败: {e}")
        return False

    # ===== 步骤3：等待登录窗口出现 =====
    delay = int(config.get("delay_seconds", "5"))
    log(f"等待 {delay} 秒让程序加载...")
    time.sleep(delay)

    # ===== 步骤4：检查是否已登录（启动时可能直接进了） =====
    try:
        app2 = Application(backend="uia").connect(path=exe_path, timeout=5)
        if check_logged_in(app2):
            log("启动后检测已登录，登录成功！")
            return True
    except Exception:
        pass

    # ===== 步骤5：窗口聚焦 =====
    log("正在进行窗口聚焦...")
    try:
        app3 = Application(backend="uia").connect(path=exe_path, timeout=5)
        if not bring_window_to_front(app3):
            log("窗口激活失败，尝试备用方案...")
            if HAS_PYAUTOGUI:
                pyautogui.press('alt')
                time.sleep(0.5)
                pyautogui.press('tab')
                time.sleep(1)
    except Exception as e:
        log(f"窗口连接失败: {e}")
        if HAS_PYAUTOGUI:
            log("尝试 alt-tab 切换窗口...")
            pyautogui.press('alt')
            time.sleep(0.5)
            pyautogui.press('tab')
            time.sleep(1)

    # ===== 步骤6：输入密码 =====
    if HAS_PYAUTOGUI:
        try:
            # 焦点切到密码框（先按Tab让焦点移动）
            log("切换到密码框...")
            pyautogui.press('tab')
            time.sleep(0.5)
            pyautogui.press('tab')  # 再按一次确保在密码框
            time.sleep(0.5)

            # 全选清空
            pyautogui.hotkey('ctrl', 'a')
            time.sleep(0.3)

            # 输入密码
            log("输入密码...")
            pyautogui.typewrite(password, interval=0.05)
            time.sleep(0.5)

            log("按回车提交登录...")
            pyautogui.press('enter')
        except Exception as e:
            log(f"键盘模拟失败: {e}")
            log("请手动输入密码并登录...")
            time.sleep(10)
    else:
        log("警告: 未安装 pyautogui，请手动登录")
        time.sleep(15)

    # ===== 步骤7：等待登录完成 =====
    log("等待登录完成...")
    timeout = 60
    for i in range(timeout // 2):
        time.sleep(2)
        try:
            app4 = Application(backend="uia").connect(path=exe_path, timeout=3)
            if check_logged_in(app4):
                log("✓ 登录成功！")
                return True
        except Exception:
            pass
        if i % 5 == 0:
            log(f"  等待中... ({i*2}/{timeout}秒)")

    log("登录超时，请检查QMT窗口状态")
    return False


def run_schedule():
    """运行定时任务"""
    global running
    while running:
        schedule.run_pending()
        time.sleep(1)


def start_scheduler():
    """启动调度器"""
    global running, schedule_thread

    cfg = load_config()
    if not cfg.get("exe_path") or not cfg.get("password"):
        messagebox.showwarning("警告", "请先填写程序路径和密码！")
        return

    running = True
    schedule.clear()

    # 周一到周五 8:00 执行
    for day in ["monday", "tuesday", "wednesday", "thursday", "friday"]:
        getattr(schedule.every(), day).at("08:00").do(do_login)

    schedule_thread = threading.Thread(target=run_schedule, daemon=True)
    schedule_thread.start()

    update_status()
    messagebox.showinfo("成功", "定时任务已启动！\n周一至周五 08:00 将自动登录 QMT")


def stop_scheduler():
    """停止调度器"""
    global running
    running = False
    schedule.clear()
    update_status()
    messagebox.showinfo("提示", "定时任务已停止")


def get_next_run_time():
    """计算下次执行时间"""
    now = datetime.datetime.now()
    weekday = now.weekday()

    if weekday >= 5:
        days_until = 7 - weekday
        next_time = now + datetime.timedelta(days=days_until)
        return next_time.replace(hour=8, minute=0, second=0, microsecond=0)

    target_time = datetime.time(8, 0, 0)
    if now.time() < target_time:
        return now.replace(hour=8, minute=0, second=0, microsecond=0)
    else:
        next_time = now + datetime.timedelta(days=1)
        while next_time.weekday() >= 5:
            next_time += datetime.timedelta(days=1)
        return next_time.replace(hour=8, minute=0, second=0, microsecond=0)


def countdown_to_next_run():
    """倒计时"""
    next_run = get_next_run_time()
    delta = next_run - datetime.datetime.now()
    if delta.total_seconds() < 0:
        return "已错过，等待下次..."
    h, r = divmod(int(delta.total_seconds()), 3600)
    m, s = divmod(r, 60)
    return f"{h}小时 {m}分 {s}秒"


def update_status():
    """更新状态"""
    now = datetime.datetime.now()
    current_time_value.config(text=now.strftime('%H:%M:%S'))
    if running:
        remaining_value.config(text=countdown_to_next_run())
        status_label.config(text="✅ 运行中", foreground="#2e7d32")
    else:
        remaining_value.config(text="未启动")
        status_label.config(text="⏸️ 已停止", foreground="#666666")


def manual_run():
    """手动执行"""
    if messagebox.askyesno("确认", "立即执行登录？"):
        threading.Thread(target=do_login, daemon=True).start()


def create_ui():
    """创建界面"""
    global root, status_label, current_time_value, remaining_value, log_text

    root = tk.Tk()
    root.title("QMT 自动登录工具 (pywinauto版)")
    root.geometry("520x520")
    root.minsize(520, 520)

    main_frame = ttk.Frame(root, padding="12")
    main_frame.pack(fill=tk.BOTH, expand=True)

    # 标题
    title = ttk.Label(main_frame, text="QMT 自动登录工具", font=("微软雅黑", 14, "bold"))
    title.pack(pady=(0, 2))
    sub = ttk.Label(main_frame, text="基于 pywinauto + pyautogui（窗口自动聚焦）",
                    font=("微软雅黑", 8), foreground="#888")
    sub.pack(pady=(0, 8))

    # 时间显示
    time_frame = ttk.Frame(main_frame)
    time_frame.pack(fill=tk.X, pady=(0, 8))
    ttk.Label(time_frame, text="当前:").pack(side=tk.LEFT)
    current_time_value = ttk.Label(time_frame, text="--:--:--",
                                   font=("微软雅黑", 11, "bold"), foreground="#1976d2")
    current_time_value.pack(side=tk.LEFT, padx=(3, 12))
    ttk.Label(time_frame, text="剩余:").pack(side=tk.LEFT)
    remaining_value = ttk.Label(time_frame, text="--小时--分--秒",
                                font=("微软雅黑", 11, "bold"), foreground="#d32f2f")
    remaining_value.pack(side=tk.LEFT, padx=(3, 0))

    # 配置区
    cfg_frame = ttk.LabelFrame(main_frame, text="配置", padding="8")
    cfg_frame.pack(fill=tk.X, pady=(0, 6))

    ttk.Label(cfg_frame, text="QMT程序路径:").grid(row=0, column=0, sticky=tk.W, pady=2)
    exe_var = tk.StringVar(value=load_config().get("exe_path", ""))
    ttk.Entry(cfg_frame, textvariable=exe_var, width=38).grid(
        row=0, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    ttk.Label(cfg_frame, text="资金账号:").grid(row=1, column=0, sticky=tk.W, pady=2)
    uid_var = tk.StringVar(value=load_config().get("user_id", ""))
    ttk.Entry(cfg_frame, textvariable=uid_var, width=38).grid(
        row=1, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    ttk.Label(cfg_frame, text="登录密码:").grid(row=2, column=0, sticky=tk.W, pady=2)
    pwd_var = tk.StringVar(value=load_config().get("password", ""))
    ttk.Entry(cfg_frame, textvariable=pwd_var, show="*", width=38).grid(
        row=2, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    ttk.Label(cfg_frame, text="等待加载(秒):").grid(row=3, column=0, sticky=tk.W, pady=2)
    delay_var = tk.StringVar(value=load_config().get("delay_seconds", "5"))
    ttk.Spinbox(cfg_frame, from_=3, to=30, textvariable=delay_var, width=8).grid(
        row=3, column=1, sticky=tk.W, padx=(5, 0), pady=2)

    def save():
        cfg = {
            "exe_path": exe_var.get().strip(),
            "user_id": uid_var.get().strip(),
            "password": pwd_var.get().strip(),
            "delay_seconds": delay_var.get().strip()
        }
        save_config(cfg)
        messagebox.showinfo("成功", "配置已保存！")

    # 按钮区
    btn_frame = ttk.Frame(main_frame)
    btn_frame.pack(fill=tk.X, pady=(0, 6))
    ttk.Button(btn_frame, text="保存配置", command=save).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(btn_frame, text="立即测试", command=manual_run).pack(side=tk.LEFT)

    # 定时区
    sched_frame = ttk.LabelFrame(main_frame, text="定时任务 (周一至周五 08:00)", padding="6")
    sched_frame.pack(fill=tk.X, pady=(0, 6))
    bf = ttk.Frame(sched_frame)
    bf.pack()
    ttk.Button(bf, text="启动定时", command=start_scheduler).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(bf, text="停止定时", command=stop_scheduler).pack(side=tk.LEFT)
    status_label = ttk.Label(sched_frame, text="⏸️ 已停止",
                             font=("微软雅黑", 12, "bold"), foreground="#666")
    status_label.pack(pady=(6, 0))

    # 日志区
    log_frame = ttk.LabelFrame(main_frame, text="运行日志", padding="4")
    log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 4))
    log_scroll = ttk.Scrollbar(log_frame)
    log_scroll.pack(side=tk.RIGHT, fill=tk.Y)
    log_text = tk.Text(log_frame, height=6, font=("Consolas", 8),
                        yscrollcommand=log_scroll.set, state=tk.DISABLED)
    log_text.pack(fill=tk.BOTH, expand=True)
    log_scroll.config(command=log_text.yview)

    def write_log(msg):
        log_text.config(state=tk.NORMAL)
        log_text.insert(tk.END, f"{datetime.datetime.now().strftime('%H:%M:%S')} {msg}\n")
        log_text.see(tk.END)
        log_text.config(state=tk.DISABLED)

    # 注入全局 log 函数
    globals()['log_text'] = log_text

    # 环境检测
    env_warn = []
    if not HAS_PYWINAUTO:
        env_warn.append("⚠️ 缺少 pywinauto")
    if not HAS_PYAUTOGUI:
        env_warn.append("⚠️ 缺少 pyautogui")
    if env_warn:
        warn = ttk.Label(main_frame, text=" | ".join(env_warn) +
                         "\n安装命令: pip install pywinauto pyautogui",
                         foreground="#e65100", font=("微软雅黑", 8))
        warn.pack(side=tk.BOTTOM, pady=(2, 0))
        write_log("⚠️ 依赖未安装完整，部分功能可能异常")
        write_log("安装: pip install pywinauto pyautogui pytesseract opencv-python")
    else:
        write_log("✓ 环境检测通过")

    write_log(f"配置路径: {CONFIG_FILE}")

    # 每秒更新时间
    def loop():
        update_status()
        root.after(1000, loop)
    loop()

    root.mainloop()


if __name__ == "__main__":
    create_ui()
