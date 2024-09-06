import tkinter as tk
from tkinter import messagebox, ttk
import ctypes
import sys
import os
import json
from datetime import datetime, timedelta
import time
from pynput.mouse import Button, Controller
import pygetwindow as gw
from pywinauto.application import Application

# 定义版本号和窗口名称为变量
version = "1.0"
default_window_name = "直播伴侣"

# 创建鼠标控制器实例
mouse = Controller()

# 文件路径来保存偏移值
config_file = "offset_config.json"

# 读取配置文件
def load_config():
    if os.path.exists(config_file):
        with open(config_file, "r") as file:
            return json.load(file)
    return {
        "horizontal_offset_close": 0,
        "vertical_offset_close": 0,
        "horizontal_offset_confirm": 0,
        "vertical_offset_confirm": 0
    }

# 保存配置文件
def save_config(config):
    with open(config_file, "w") as file:
        json.dump(config, file)

# 载入配置
config = load_config()

# 全局变量
target_time = None
# 窗口名称，初始固定为 "直播伴侣"
window_options = [default_window_name]

# 偏移值（从配置文件中读取）
horizontal_offset_close = config.get("horizontal_offset_close", 0)
vertical_offset_close = config.get("vertical_offset_close", 0)
horizontal_offset_confirm = config.get("horizontal_offset_confirm", 0)
vertical_offset_confirm = config.get("vertical_offset_confirm", 0)

# 检查管理员权限
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# 提示用户是否重新启动为管理员权限
def request_admin_permission():
    if not is_admin():
        result = messagebox.askyesno("需要管理员权限", "此操作需要管理员权限，是否重新启动为管理员？")
        if result:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
            sys.exit()

# 模拟鼠标左键按下和释放
def simulate_left_click():
    mouse.press(Button.left)
    time.sleep(0.5)  # 模拟按下持续时间
    mouse.release(Button.left)

# 激活指定窗口并移动鼠标到窗口底部中间位置，带偏移
def activate_windows_and_move_mouse():
    selected_window = window_combobox.get()

    windows = gw.getWindowsWithTitle(selected_window)
    if windows:
        try:
            window = windows[0]
            app = Application(backend="win32").connect(handle=window._hWnd)
            win = app.window(handle=window._hWnd)
            win.set_focus()
            time.sleep(0.2)  # 短暂停顿

            # 获取窗口的位置矩形
            rect = win.rectangle()

            # 计算窗口底部中间位置并应用偏移
            center_x = rect.left + (rect.right - rect.left) // 2
            center_y = rect.bottom

            # 打印vertical_offset horizontal_offset
            print(f"关播按钮偏移:水平偏移={horizontal_offset_close}, 垂直偏移={vertical_offset_close}")

            center_x += horizontal_offset_close
            center_y += vertical_offset_close

            # 将鼠标移动到窗口底部中间位置
            print(f"鼠标移动到窗口底部中间位置: X={center_x}, Y={center_y}")
            mouse.position = (center_x, center_y)
            time.sleep(0.5)  # 停顿一小段时间

            # 模拟鼠标左键点击
            simulate_left_click()

            time.sleep(2)  # 停顿一小段时间

            # 模拟确认窗口
            simulate_confirmation(win)

        except Exception as e:
            print(f"无法激活窗口 {selected_window}: {e}")
            messagebox.showerror("激活失败", f"无法激活窗口 {selected_window}: {e}")
    else:
        messagebox.showerror("激活失败", f"未找到窗口: {selected_window}")

# 模拟确认窗口，移动到中间位置并带偏移
def simulate_confirmation(window):
    # 获取窗口的位置矩形
    rect = window.rectangle()

    # 计算窗口中间位置并应用偏移
    center_x = rect.left + (rect.right - rect.left) // 2
    center_y = rect.top + (rect.bottom - rect.top) // 2

    # 打印确认按钮的偏移
    print(f"确认按钮偏移:水平偏移={horizontal_offset_confirm}, 垂直偏移={vertical_offset_confirm}")

    center_x += horizontal_offset_confirm
    center_y += vertical_offset_confirm

    # 将鼠标移动到确认窗口中间位置
    print(f"鼠标移动到确认按钮位置: X={center_x}, Y={center_y}")
    mouse.position = (center_x, center_y)
    time.sleep(0.5)

    # 模拟鼠标左键点击
    simulate_left_click()

# 倒计时逻辑
def start_countdown():
    global target_time
    now = datetime.now()
    remaining_time = target_time - now

    # 禁用倒计时按钮、时间选择框和偏移控件
    disable_ui()

    if remaining_time.total_seconds() <= 0:
        activate_windows_and_move_mouse()
    else:
        for i in range(int(remaining_time.total_seconds()), 0, -1):
            remaining = target_time - datetime.now()
            formatted_remaining_time = format_remaining_time(remaining)
            formatted_target_time = target_time.strftime("%Y-%m-%d %H:%M:%S")
            countdown_label.config(text=f"剩余时间: {formatted_remaining_time}\n目标时间: {formatted_target_time}", font=("Arial", 20))
            root.update()
            time.sleep(1)

        countdown_label.config(text="任务执行完毕", font=("Arial", 20))
        activate_windows_and_move_mouse()

    # 重新启用倒计时按钮、时间选择框和偏移控件
    enable_ui()

# 设置倒计时1秒
def set_one_second_countdown():
    global target_time
    target_time = datetime.now() + timedelta(seconds=1)

    # 禁用倒计时按钮、时间选择框和偏移控件
    disable_ui()

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

        # 禁用倒计时按钮、时间选择框和偏移控件
        disable_ui()

        start_countdown()

# 禁用所有倒计时按钮、时间选择框和偏移控件
def disable_ui():
    countdown_button.config(state=tk.DISABLED)
    one_second_button.config(state=tk.DISABLED)
    year_combobox.config(state=tk.DISABLED)
    month_combobox.config(state=tk.DISABLED)
    day_combobox.config(state=tk.DISABLED)
    hour_combobox.config(state=tk.DISABLED)
    minute_combobox.config(state=tk.DISABLED)
    window_combobox.config(state=tk.DISABLED)
    horizontal_offset_close_spinbox.config(state=tk.DISABLED)
    vertical_offset_close_spinbox.config(state=tk.DISABLED)
    horizontal_offset_confirm_spinbox.config(state=tk.DISABLED)
    vertical_offset_confirm_spinbox.config(state=tk.DISABLED)

# 启用所有倒计时按钮、时间选择框和偏移控件
def enable_ui():
    countdown_button.config(state=tk.NORMAL)
    one_second_button.config(state=tk.NORMAL)
    year_combobox.config(state="readonly")
    month_combobox.config(state="readonly")
    day_combobox.config(state="readonly")
    hour_combobox.config(state="readonly")
    minute_combobox.config(state="readonly")
    window_combobox.config(state="readonly")
    horizontal_offset_close_spinbox.config(state=tk.NORMAL)
    vertical_offset_close_spinbox.config(state=tk.NORMAL)
    horizontal_offset_confirm_spinbox.config(state=tk.NORMAL)
    vertical_offset_confirm_spinbox.config(state=tk.NORMAL)

# 格式化剩余时间为 年-月-日 时:分:秒，省略为 0 的单位
def format_remaining_time(td):
    days = td.days
    seconds = td.seconds
    hours, seconds = divmod(seconds, 3600)
    minutes, seconds = divmod(seconds, 60)

    years = days // 365
    months = (days % 365) // 30
    remaining_days = (days % 365) % 30

    parts = []
    if years > 0:
        parts.append(f"{years}年")
    if months > 0:
        parts.append(f"{months}月")
    if remaining_days > 0:
        parts.append(f"{remaining_days}日")
    if hours > 0 or len(parts) > 0:
        parts.append(f"{hours}时")
    if minutes > 0 or len(parts) > 0:
        parts.append(f"{minutes}分")
    parts.append(f"{seconds}秒")

    return ''.join(parts)

# 更新水平偏移值（关播按钮）
def update_horizontal_offset_close(*args):
    global horizontal_offset_close
    horizontal_offset_close = int(horizontal_offset_close_var.get())
    config["horizontal_offset_close"] = horizontal_offset_close
    save_config(config)

# 更新垂直偏移值（关播按钮）
def update_vertical_offset_close(*args):
    global vertical_offset_close
    vertical_offset_close = int(vertical_offset_close_var.get())
    config["vertical_offset_close"] = vertical_offset_close
    save_config(config)

# 更新水平偏移值（确认按钮）
def update_horizontal_offset_confirm(*args):
    global horizontal_offset_confirm
    horizontal_offset_confirm = int(horizontal_offset_confirm_var.get())
    config["horizontal_offset_confirm"] = horizontal_offset_confirm
    save_config(config)

# 更新垂直偏移值（确认按钮）
def update_vertical_offset_confirm(*args):
    global vertical_offset_confirm
    vertical_offset_confirm = int(vertical_offset_confirm_var.get())
    config["vertical_offset_confirm"] = vertical_offset_confirm
    save_config(config)

# 创建主窗口
root = tk.Tk()
root.title(f"深锶-直播伴侣自动关播工具V{version}")
root.geometry("450x600")
root.minsize(width=450, height=600)

# 检查并请求管理员权限
request_admin_permission()

# 窗口选择部分
window_frame = tk.Frame(root)
window_frame.pack(pady=10)

# 窗口选择下拉框
window_label = tk.Label(window_frame, text="窗口选择:")
window_label.pack(side="left", padx=5)

# 直播伴侣为固定选项
window_var = tk.StringVar(value=default_window_name)
window_combobox = ttk.Combobox(window_frame, textvariable=window_var, values=window_options, state='readonly', width=20)
window_combobox.pack(side="left")

# 日期选择部分
date_frame = tk.Frame(root)
date_frame.pack(pady=10)

# 年选择框
year_var = tk.StringVar()
year_combobox = ttk.Combobox(date_frame, textvariable=year_var, values=[str(i) for i in range(datetime.now().year, datetime.now().year + 10)], state='readonly', width=5)
year_combobox.pack(side='left')
year_combobox.set(datetime.now().strftime("%Y"))
year_label = tk.Label(date_frame, text="年")
year_label.pack(side='left', padx=5)

# 月选择框
month_var = tk.StringVar()
month_combobox = ttk.Combobox(date_frame, textvariable=month_var, values=[f"{i:02d}" for i in range(1, 13)], state='readonly', width=3)
month_combobox.pack(side='left')
month_combobox.set(datetime.now().strftime("%m"))
month_label = tk.Label(date_frame, text="月")
month_label.pack(side='left', padx=5)

# 日选择框
day_var = tk.StringVar()
day_combobox = ttk.Combobox(date_frame, textvariable=day_var, values=[f"{i:02d}" for i in range(1, 32)], state='readonly', width=3)
day_combobox.pack(side='left')
day_combobox.set(datetime.now().strftime("%d"))
day_label = tk.Label(date_frame, text="日")
day_label.pack(side='left', padx=5)

# 时间选择部分
time_frame = tk.Frame(root)
time_frame.pack(pady=10)

# 小时选择框
hour_var = tk.StringVar()
hour_combobox = ttk.Combobox(time_frame, textvariable=hour_var, values=[f"{i:02d}" for i in range(24)], state='readonly', width=3)
hour_combobox.pack(side='left')
hour_combobox.set(datetime.now().strftime("%H"))
hour_label = tk.Label(time_frame, text="时")
hour_label.pack(side='left', padx=5)

# 分钟选择框
minute_var = tk.StringVar()
minute_combobox = ttk.Combobox(time_frame, textvariable=minute_var, values=[f"{i:02d}" for i in range(60)], state='readonly', width=3)
minute_combobox.pack(side='left')
minute_combobox.set(datetime.now().strftime("%M"))
minute_label = tk.Label(time_frame, text="分")
minute_label.pack(side='left', padx=5)

# 偏移设置部分
offset_frame = tk.Frame(root)
offset_frame.pack(pady=10)

# 创建偏移量的变量
horizontal_offset_close_var = tk.StringVar(value=str(horizontal_offset_close))
vertical_offset_close_var = tk.StringVar(value=str(vertical_offset_close))
horizontal_offset_confirm_var = tk.StringVar(value=str(horizontal_offset_confirm))
vertical_offset_confirm_var = tk.StringVar(value=str(vertical_offset_confirm))

# 追踪偏移量的变化
horizontal_offset_close_var.trace_add("write", update_horizontal_offset_close)
vertical_offset_close_var.trace_add("write", update_vertical_offset_close)
horizontal_offset_confirm_var.trace_add("write", update_horizontal_offset_confirm)
vertical_offset_confirm_var.trace_add("write", update_vertical_offset_confirm)

# 关播按钮水平偏移设置
horizontal_offset_close_label = tk.Label(offset_frame, text="关播按钮水平偏移:")
horizontal_offset_close_label.grid(row=0, column=0, padx=5, pady=5)
horizontal_offset_close_spinbox = tk.Spinbox(offset_frame, from_=-500, to=500, textvariable=horizontal_offset_close_var, width=5)
horizontal_offset_close_spinbox.grid(row=0, column=1, padx=5, pady=5)

# 关播按钮垂直偏移设置
vertical_offset_close_label = tk.Label(offset_frame, text="关播按钮垂直偏移:")
vertical_offset_close_label.grid(row=0, column=2, padx=5, pady=5)
vertical_offset_close_spinbox = tk.Spinbox(offset_frame, from_=-500, to=500, textvariable=vertical_offset_close_var, width=5)
vertical_offset_close_spinbox.grid(row=0, column=3, padx=5, pady=5)

# 新行：确认按钮水平偏移设置
horizontal_offset_confirm_label = tk.Label(offset_frame, text="确认按钮水平偏移:")
horizontal_offset_confirm_label.grid(row=1, column=0, padx=5, pady=5)
horizontal_offset_confirm_spinbox = tk.Spinbox(offset_frame, from_=-500, to=500, textvariable=horizontal_offset_confirm_var, width=5)
horizontal_offset_confirm_spinbox.grid(row=1, column=1, padx=5, pady=5)

# 新行：确认按钮垂直偏移设置
vertical_offset_confirm_label = tk.Label(offset_frame, text="确认按钮垂直偏移:")
vertical_offset_confirm_label.grid(row=1, column=2, padx=5, pady=5)
vertical_offset_confirm_spinbox = tk.Spinbox(offset_frame, from_=-500, to=500, textvariable=vertical_offset_confirm_var, width=5)
vertical_offset_confirm_spinbox.grid(row=1, column=3, padx=5, pady=5)

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
