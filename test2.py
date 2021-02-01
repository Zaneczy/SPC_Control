import sys
import tkinter as Tk
import matplotlib
from numpy import arange, sin, pi
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
from matplotlib.backend_bases import key_press_handler
from matplotlib.figure import Figure

matplotlib.use('TkAgg')
root = Tk.Tk()
root.title("matplotlib in TK")
# 设置图形尺寸与质量
f = Figure(figsize=(5, 4), dpi=100)
# a = f.add_subplot(111)
# t = arange(0.0, 3, 0.01)
# s = sin(2 * pi * t)

average_data=['100.0112', '99.9930', '99.9746', '99.9732', '100.0196', '100.0102', '100.0010', '99.9974', '100.0054', '99.9712', '99.9968', '99.9644', '100.0094', '99.9964', '99.9748', '99.9896', '100.0370', '100.0038', '99.9874', '99.9602', '99.9992', '100.0254', '100.0530', '99.9910', '100.0272']
# 绘图
plt = f.add_subplot(111)
x_data = [i + 1 for i in range(len(average_data))]
y_data = list(map(float, average_data))
plt.plot(x_data, y_data, 'bo-', linewidth=1, label='average')
plt.plot(x_data, [100 for i in x_data], 'g', linewidth=1)
plt.plot(x_data, [99.96 for i in x_data], 'y--', linewidth=1)
plt.plot(x_data, [99.98 for i in x_data], 'y--', linewidth=1)
plt.plot(x_data, [100.02 for i in x_data], 'y--', linewidth=1)
plt.plot(x_data, [100.04 for i in x_data], 'y--', linewidth=1)
plt.plot(x_data, [99.94 for i in x_data], 'r', linewidth=1)
plt.plot(x_data, [100.06 for i in x_data], 'r', linewidth=1)
# plt.title('SPC')
# plt.legend()
# plt.xlabel('x')
# plt.ylabel('y')
# my_x_ticks = np.arange(0, 26, 1)
# my_y_ticks = np.arange(99.9, 100.1, 0.02)
# plt.xticks(my_x_ticks)
# plt.yticks(my_y_ticks)


# 绘制图形
# a.plot(t, s)
# 把绘制的图形显示到tkinter窗口上
canvas = FigureCanvasTkAgg(f, master=root)
canvas.draw()
canvas.get_tk_widget().pack(side=Tk.TOP, fill=Tk.BOTH, expand=1)
# 把matplotlib绘制图形的导航工具栏显示到tkinter窗口上
# toolbar = NavigationToolbar2Tk(canvas, root)
# toolbar.update()
# canvas._tkcanvas.pack(side=Tk.TOP, fill=Tk.BOTH, expand=1)


# 定义并绑定键盘事件处理函数
def on_key_event(event):
    print('you pressed %s' % event.key)
    key_press_handler(event, canvas, toolbar)
    canvas.mpl_connect('key_press_event', on_key_event)


# 按钮单击事件处理函数
def _quit():
    # 结束事件主循环，并销毁应用程序窗口
    root.quit()
    root.destroy()
    button = Tk.Button(master=root, text='Quit', command=_quit)
    button.pack(side=Tk.BOTTOM)


Tk.mainloop()