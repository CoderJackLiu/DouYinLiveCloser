import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import pygetwindow as gw
from pywinauto.application import Application
import pyautogui
import warnings
from datetime import datetime, timedelta
import time

# 忽略 32 位和 64 位不匹配的警告
warnings.filterwarnings("ignore", category=UserWarning)

# 创建全局变量用于倒计时
target_time = None


def move_mouse_with_retry(target_x, target_y, retries):
    """
    尝试将鼠标移动到目标位置。如果移动失败则每隔0.5秒重试，直到达到最大重试次数。
    :param target_x: 鼠标目标X坐标
    :param target_y: 鼠标目标Y坐标
    :param retries: 重试次数
    :return: None
    """
    for attempt in range(retries):
        # 移动鼠标到目标位置
        pyautogui.moveTo(target_x, target_y)

        # 获取当前鼠标的位置
        current_x, current_y = pyautogui.position()

        # 打印当前鼠标位置和目标位置
        print(f"重试 {attempt + 1}: 当前鼠标位置: X={current_x}, Y={current_y}，目标: X={target_x}, Y={target_y}")

        # 检查是否已经到达目标位置
        if current_x == target_x and current_y == target_y:
            print("鼠标已到达目标位置")
            return

        # 等待0.5秒后再次尝试
        time.sleep(0.5)

    # 如果重试次数用尽后还未成功，则打印失败信息
    print("鼠标未能到达目标位置，已超出重试次数")


def list_window_titles():
    # 获取所有窗口
    all_windows = gw.getWindowsWithTitle('')

    # 清空下拉列表中的旧数据
    combobox['values'] = []

    # 过滤空白标题
    valid_titles = [window.title for window in all_windows if window.title.strip()]

    # 将获取的窗口标题填充到下拉选择框中
    if valid_titles:
        combobox['values'] = valid_titles
        combobox.current(0)  # 默认选中第一个窗口
    else:
        combobox['values'] = ["没有找到任何窗口"]
        combobox.current(0)
        messagebox.showwarning("警告", "没有找到任何窗口。")


def SetMousePosition():
    pyautogui.moveTo(3192, 640)


def bring_window_to_front():
    selected_title = window_title_var.get()
    if not selected_title or selected_title == "没有找到任何窗口":
        return

    windows = gw.getWindowsWithTitle(selected_title)

    if windows:
        try:
            # 选择第一个匹配的窗口，使用其句柄进行连接
            window = windows[0]
            handle = window._hWnd

            # 使用pywinauto通过句柄连接窗口
            app = Application(backend="win32").connect(handle=handle)
            win = app.window(handle=handle)
            win.set_focus()  # 将窗口置于前台并聚焦

            # 获取窗口位置并将鼠标移动到窗口的右上角，向左和向下偏移 20 像素
            rect = win.rectangle()
            # 停顿一秒
            time.sleep(1)

            # 打印鼠标需要移动到的目标位置
            print(f"鼠标目标位置: X={rect.right - 20}, Y={rect.top + 20}")
            pyautogui.moveTo(rect.right - 20, rect.top + 20)
            time.sleep(1)

            # 调用封装的移动函数，重试10次
            move_mouse_with_retry(rect.right - 20, rect.top + 20, retries=10)

            # pyautogui.moveTo(rect.right - 20, rect.top + 20)

            # 停顿一秒
            time.sleep(1)

            # 模拟鼠标点击
            pyautogui.click()

        except Exception as e:
            # 只有在发生错误时弹出错误信息
            messagebox.showerror("错误", f"无法激活窗口: {e}")


def validate_time():
    year = int(year_var.get())
    month = int(month_var.get())
    day = int(day_var.get())
    hour = int(hour_var.get())
    minute = int(minute_var.get())

    selected_datetime = datetime(year, month, day, hour, minute)
    now = datetime.now()

    if selected_datetime <= now:
        messagebox.showwarning("无效时间", "请选择晚于当前时间的日期和时间。")
    else:
        global target_time
        target_time = selected_datetime
        countdown_label.config(text=f"目标时间: {target_time.strftime('%Y-%m-%d %H:%M:%S')}")
        update_countdown()  # 启动倒计时


# 添加倒计时5秒的逻辑
def set_five_seconds_countdown():
    global target_time
    now = datetime.now()
    target_time = now + timedelta(seconds=2)  # 设置目标时间为当前时间加5秒
    countdown_label.config(text=f"目标时间: {target_time.strftime('%Y-%m-%d %H:%M:%S')}")
    update_countdown()  # 启动倒计时


# 更新倒计时
def update_countdown():
    if target_time:
        now = datetime.now()
        remaining_time = target_time - now

        if remaining_time.total_seconds() > 0:
            countdown_label.config(text=f"倒计时: {str(remaining_time).split('.')[0]}")
            root.after(1000, update_countdown)
        else:
            countdown_label.config(text="倒计时结束！")
            bring_window_to_front()  # 倒计时结束时，打开窗口并移动鼠标


# 添加打印鼠标位置的函数
def print_mouse_position():
    # 获取鼠标当前位置
    mouse_x, mouse_y = pyautogui.position()
    print(f"当前鼠标位置: X={mouse_x}, Y={mouse_y}")

    # 每隔500毫秒调用一次自己
    root.after(500, print_mouse_position)


# 获取当前日期和时间
now = datetime.now()

# 创建主窗口
root = tk.Tk()
root.title("窗口管理工具")
root.geometry("500x500")

# 添加下拉选择框用于选择窗口标题
window_title_var = tk.StringVar()
combobox = ttk.Combobox(root, textvariable=window_title_var, state='readonly')
combobox.pack(pady=10, fill='x')

# 添加一个刷新按钮用于获取最新的窗口列表
refresh_button = tk.Button(root, text="刷新窗口列表", command=list_window_titles)
refresh_button.pack(pady=10)

# 添加一个按钮来将选中的窗口置顶
bring_to_front_button = tk.Button(root, text="置顶选中窗口", command=bring_window_to_front)
bring_to_front_button.pack(pady=10)

# 添加一个按钮来将选中的窗口置顶
bring_to_front_button = tk.Button(root, text="设置鼠标位置", command=SetMousePosition)
bring_to_front_button.pack(pady=10)

# 添加日期选择器部分
date_frame = tk.Frame(root)
date_frame.pack(pady=10)

# 年选择框
year_var = tk.StringVar()
year_label = tk.Label(date_frame, text="年:")
year_label.pack(side='left', padx=5)
year_combobox = ttk.Combobox(date_frame, textvariable=year_var, values=[str(i) for i in range(now.year, now.year + 10)],
                             state='readonly', width=5)
year_combobox.pack(side='left')
year_combobox.set(now.strftime("%Y"))

# 月选择框
month_var = tk.StringVar()
month_label = tk.Label(date_frame, text="月:")
month_label.pack(side='left', padx=5)
month_combobox = ttk.Combobox(date_frame, textvariable=month_var, values=[f"{i:02d}" for i in range(1, 13)],
                              state='readonly', width=3)
month_combobox.pack(side='left')
month_combobox.set(now.strftime("%m"))

# 日选择框
day_var = tk.StringVar()
day_label = tk.Label(date_frame, text="日:")
day_label.pack(side='left', padx=5)
day_combobox = ttk.Combobox(date_frame, textvariable=day_var, values=[f"{i:02d}" for i in range(1, 32)],
                            state='readonly', width=3)
day_combobox.pack(side='left')
day_combobox.set(now.strftime("%d"))

# 添加时间选择器部分
time_frame = tk.Frame(root)
time_frame.pack(pady=10)

# 小时选择框
hour_var = tk.StringVar()
hour_label = tk.Label(time_frame, text="小时:")
hour_label.pack(side='left', padx=5)
hour_combobox = ttk.Combobox(time_frame, textvariable=hour_var, values=[f"{i:02d}" for i in range(24)],
                             state='readonly', width=3)
hour_combobox.pack(side='left')
hour_combobox.set(now.strftime("%H"))

# 分钟选择框
minute_var = tk.StringVar()
minute_label = tk.Label(time_frame, text="分钟:")
minute_label.pack(side='left', padx=5)
minute_combobox = ttk.Combobox(time_frame, textvariable=minute_var, values=[f"{i:02d}" for i in range(60)],
                               state='readonly', width=3)
minute_combobox.pack(side='left')
minute_combobox.set(now.strftime("%M"))

# 添加一个按钮用于验证选中的日期和时间
validate_time_button = tk.Button(root, text="设置时间", command=validate_time)
validate_time_button.pack(pady=10)

# 添加倒计时5秒按钮
set_five_seconds_button = tk.Button(root, text="倒计时5秒", command=set_five_seconds_countdown)
set_five_seconds_button.pack(pady=10)

# 添加倒计时显示标签
countdown_label = tk.Label(root, text="未设置倒计时", font=("Arial", 24))
countdown_label.pack(pady=20)

pyautogui.FAILSAFE = False
pyautogui.FAILSAFE_POINTS = [(1, 1)]

# 初始化时，刷新一次窗口列表
list_window_titles()

# 在主循环启动之前，调用这个函数来开始定时打印
print_mouse_position()

# 启动主循环
root.mainloop()
