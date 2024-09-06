import tkinter as tk
from pynput.mouse import Button, Controller, Listener
import time

# 创建鼠标控制器实例
mouse = Controller()

# 创建 Tkinter 窗口
root = tk.Tk()
root.attributes('-topmost', True)
root.attributes('-fullscreen', True)  # 全屏透明窗口
root.config(bg='black', alpha=0.01)   # 设置背景透明

# 创建画布，绘制圆圈
canvas = tk.Canvas(root, bg="black", highlightthickness=0)
canvas.pack(fill=tk.BOTH, expand=True)

# 设置圆圈半径
circle_radius = 30

def draw_circle(x, y):
    """在指定位置绘制一个圆圈"""
    circle = canvas.create_oval(x - circle_radius, y - circle_radius, x + circle_radius, y + circle_radius, outline='red', width=2)
    root.update()
    time.sleep(0.5)  # 圆圈显示的时间
    canvas.delete(circle)

def on_click(x, y, button, pressed):
    """鼠标按下时显示圆圈"""
    if button == Button.left and pressed:
        draw_circle(x, y)

# 监听鼠标事件
listener = Listener(on_click=on_click)
listener.start()

# 启动 Tkinter 主循环
root.mainloop()
