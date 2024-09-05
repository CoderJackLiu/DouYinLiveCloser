import tkinter as tk
from tkinter import messagebox, ttk
import ctypes
import sys
import os
from datetime import datetime, timedelta
import time
from pynput.keyboard import Controller, Key
import pygetwindow as gw
from pywinauto.application import Application

# 创建键盘控制器实例
keyboard = Controller()

# 全局变量
target_time = None
# active_windows = ["直播伴侣", "超级鼠标连点器"]
active_windows = ["直播伴侣"]


# 检查管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


# 提示用户是否重新启动为管理员权限
def request_admin_permission():
    if not is_admin():
        # 弹出确认对话框
        result = messagebox.askyesno("需要管理员权限", "此操作需要管理员权限，是否重新启动为管理员？")
        if result:
            # 如果用户选择了 "是"，重新启动并请求管理员权限
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()


# 激活指定窗口
def activate_windows():
    for window_name in active_windows:
        windows = gw.getWindowsWithTitle(window_name)
        if windows:
            try:
                window = windows[0]
                app = Application(backend="win32").connect(handle=window._hWnd)
                win = app.window(handle=window._hWnd)
                win.set_focus()
                time.sleep(0.2)  # 短暂停顿
            except Exception as e:
                print(f"无法激活窗口 {window_name}: {e}")


# 执行 F8 快捷键
def execute_f8():
    activate_windows()
    keyboard.press(Key.f8)
    keyboard.release(Key.f8)


# 倒计时逻辑
def start_countdown():
    global target_time
    now = datetime.now()
    remaining_time = target_time - now
    countdown_button.config(state=tk.DISABLED)

    if remaining_time.total_seconds() <= 0:
        execute_f8()
    else:
        for i in range(int(remaining_time.total_seconds()), 0, -1):
            countdown_label.config(text=f"倒计时: {i}s")
            root.update()
            time.sleep(1)

        countdown_label.config(text="任务执行完毕")
        execute_f8()

    countdown_button.config(state=tk.NORMAL)


# 设置倒计时1秒
def set_one_second_countdown():
    global target_time
    target_time = datetime.now() + timedelta(seconds=1)
    start_countdown()


# 验证时间并启动倒计时
def validate_and_start_countdown():
    year = int(year_var.get())
    month = int(month_var.get())
    day = int(day_var.get())
    hour = int(hour_var.get())
    minute = int(minute_var.get())

    selected_time = datetime(year, month, day, hour, minute)
    now = datetime.now()

    if selected_time <= now:
        countdown_label.config(text="无效的时间，请选择晚于当前的时间")
    else:
        global target_time
        target_time = selected_time
        start_countdown()


# 创建主窗口
root = tk.Tk()
root.title("倒计时执行任务")
root.geometry("400x400")

# 检查并请求管理员权限
request_admin_permission()

# 日期选择部分
date_frame = tk.Frame(root)
date_frame.pack(pady=10)

# 年选择框
year_var = tk.StringVar()
year_label = tk.Label(date_frame, text="年:")
year_label.pack(side='left', padx=5)
year_combobox = ttk.Combobox(date_frame, textvariable=year_var, values=[str(i) for i in range(datetime.now().year, datetime.now().year + 10)], state='readonly', width=5)
year_combobox.pack(side='left')
year_combobox.set(datetime.now().strftime("%Y"))

# 月选择框
month_var = tk.StringVar()
month_label = tk.Label(date_frame, text="月:")
month_label.pack(side='left', padx=5)
month_combobox = ttk.Combobox(date_frame, textvariable=month_var, values=[f"{i:02d}" for i in range(1, 13)], state='readonly', width=3)
month_combobox.pack(side='left')
month_combobox.set(datetime.now().strftime("%m"))

# 日选择框
day_var = tk.StringVar()
day_label = tk.Label(date_frame, text="日:")
day_label.pack(side='left', padx=5)
day_combobox = ttk.Combobox(date_frame, textvariable=day_var, values=[f"{i:02d}" for i in range(1, 32)], state='readonly', width=3)
day_combobox.pack(side='left')
day_combobox.set(datetime.now().strftime("%d"))

# 时间选择部分
time_frame = tk.Frame(root)
time_frame.pack(pady=10)

# 小时选择框
hour_var = tk.StringVar()
hour_label = tk.Label(time_frame, text="小时:")
hour_label.pack(side='left', padx=5)
hour_combobox = ttk.Combobox(time_frame, textvariable=hour_var, values=[f"{i:02d}" for i in range(24)], state='readonly', width=3)
hour_combobox.pack(side='left')
hour_combobox.set(datetime.now().strftime("%H"))

# 分钟选择框
minute_var = tk.StringVar()
minute_label = tk.Label(time_frame, text="分钟:")
minute_label.pack(side='left', padx=5)
minute_combobox = ttk.Combobox(time_frame, textvariable=minute_var, values=[f"{i:02d}" for i in range(60)], state='readonly', width=3)
minute_combobox.pack(side='left')
minute_combobox.set(datetime.now().strftime("%M"))

# 添加倒计时1秒按钮
one_second_button = tk.Button(root, text="倒计时1秒", command=set_one_second_countdown, font=("Arial", 12))
one_second_button.pack(pady=10)

# 添加验证并启动倒计时按钮
countdown_button = tk.Button(root, text="启动倒计时", command=validate_and_start_countdown, font=("Arial", 12))
countdown_button.pack(pady=10)

# 显示倒计时标签
countdown_label = tk.Label(root, text="请选择时间或点击倒计时1秒", font=("Arial", 16))
countdown_label.pack(pady=20)

# 启动主循环
root.mainloop()
