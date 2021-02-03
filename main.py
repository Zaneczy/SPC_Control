#!/usr/bin/python3
# coding=utf-8
import tkinter
from tkinter import messagebox
import xlrd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg,NavigationToolbar2Tk
from matplotlib.figure import Figure
import configparser as cp


class NewConfigParser(cp.ConfigParser):
    def optionxform(self, optionstr):
        return optionstr


class BaseWindow:

    def __init__(self, root):
        self.root = root
        self.root.config()
        self.root.geometry("1100x575+450+200")
        BaseFrame(self.root)


class BaseFrame:

    def __init__(self, root, page=1):
        self.root = root
        self.root.config(bg='#CCFFFF')
        self.root.title('SPC Control')
        self.base_frame = tkinter.Frame(self.root)
        self.base_frame.config(bg='#CCFFFF')
        self.base_frame.pack()
        self.cf = Config()
        self.page = page
        self._menu()
        self._button()
        self._page()

    def _button(self):
        tables = self.cf.get_all()
        table_show = tables[0+((self.page-1) * 15):15+((self.page-1) * 15)]
        button_frame1 = tkinter.Frame(self.base_frame)
        button_frame2 = tkinter.Frame(self.base_frame)
        button_frame3 = tkinter.Frame(self.base_frame)
        button_frame1.pack()
        button_frame2.pack()
        button_frame3.pack()
        for i in range(len(table_show)):
            if i < 5:
                self.graph_button(button_frame1, table_show[i])
            elif i < 10:
                self.graph_button(button_frame2, table_show[i])
            else:
                self.graph_button(button_frame3, table_show[i])

    def _menu(self):
        menu = tkinter.Menu(self.root)
        menu2 = tkinter.Menu(menu, tearoff=0)
        menu.add_cascade(label='Set', menu=menu2)
        menu.add_command(label='Refresh', command=self.refresh)
        menu2.add_command(label='New', command=self.add_spc)
        menu3 = tkinter.Menu(menu2, tearoff=0)
        menu2.add_cascade(label='Delete', menu=menu3)
        tables = self.cf.get_all()
        for i in tables:
            self.menu3_command(menu3, i)
        self.root.config(menu=menu)

    def _page(self):
        page_frame = tkinter.Frame(self.base_frame)
        page_frame.pack(side=tkinter.BOTTOM)
        for i in range(5):
            if i+1 == self.page:
                self.page_button_click(page_frame, i+1)
            else:
                self.page_button(page_frame, i+1)

    def page_button(self, page_frame, num):
        page = tkinter.Button(page_frame, text=num, command=lambda:self.change_page(num), width=1, height=1)
        page.pack(side=tkinter.LEFT)

    def page_button_click(self, page_frame, num):
        page = tkinter.Button(page_frame, text=num, width=1, height=1, relief='sunken')
        page.pack(side=tkinter.LEFT)

    def change_page(self, num):
        self.base_frame.destroy()
        BaseFrame(self.root, num)

    def add_spc(self):
        Dialog()

    def menu3_command(self, menu3, table):
        menu3.add_command(label=f'{table[1]}({table[0]})', command=lambda: self.del_spc(table))

    def del_spc(self, table):
        answer = messagebox.askyesno('确认提示', f'确定删除{table[1]}({table[0]})？')
        if answer:
            self.cf.remove(table[0])
            self.refresh()

    def graph_button(self, frame, table):
        btn = tkinter.Button(frame, text=f'{table[1]} ({table[0]})', command=lambda:self.graphs(table), height=10,
                             width=30, activebackground='#c8c8c8')
        btn.pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=1)

    def graphs(self, table):
        # self.base_frame.destroy()
        # LineGraphs(self.root, table)
        try:
            self.base_frame.destroy()
            LineGraphs(self.root, table)
        except xlrd.biffh.XLRDError:
            messagebox.showerror(title='错误', message=f'Excel文件中并没有表"{table[0]}"')
            self.refresh()
        except:
            messagebox.showinfo(title='数据出错', message=f'表{table[0]}中的数据不符合格式，请检查')
            self.refresh()

    def refresh(self):
        self.base_frame.destroy()
        BaseFrame(self.root)


class Dialog:
    def __init__(self):
        self.dialog = tkinter.Toplevel()
        self.dialog.title('新建')
        self.dialog.geometry("220x80+550+300")
        self.setup_UI()

    def setup_UI(self):
        row1 = tkinter.Frame(self.dialog)
        row1.pack(fill="x")
        tkinter.Label(row1, text='Sheet：', width=8).pack(side=tkinter.LEFT)
        self.sheet = tkinter.StringVar()
        # self.sheet.set('Sheet')
        tkinter.Entry(row1, textvariable=self.sheet, width=20).pack(side=tkinter.LEFT)

        row2 = tkinter.Frame(self.dialog)
        row2.pack(fill="x", ipadx=1, ipady=1)
        tkinter.Label(row2, text='Name：', width=8).pack(side=tkinter.LEFT)
        self.name = tkinter.StringVar()
        # self.name.set('Name')
        tkinter.Entry(row2, textvariable=self.name, width=20).pack(side=tkinter.LEFT)

        row3 = tkinter.Frame(self.dialog)
        row3.pack()
        tkinter.Button(row3, text="确定", command=self.ok).pack(side=tkinter.LEFT)
        tkinter.Button(row3, text="取消", command=self.cancel).pack(side=tkinter.LEFT)

    def ok(self):
        sheet = self.sheet.get()
        name = self.name.get()
        cf = Config()
        sheets = cf.get_all_key()
        if sheet == '' or name == '':
            messagebox.showinfo(title='提示', message='不能为空')
        elif sheet in sheets:
            name = cf.get_value(sheet)
            messagebox.showinfo(title='提示', message=f'{sheet}已存在于{name}')
        else:
            try:
                cf = Config()
                cf.update_value(sheet, name)
                messagebox.showinfo(title='提示', message='创建成功，请刷新')
            except:
                messagebox.showwarning(title='提示', message='创建失败')
        self.dialog.destroy()

    def cancel(self):
        self.dialog.destroy()


class LineGraphs:
    def __init__(self, root, table):
        self.root = root
        self.root.config(menu='')
        self.cf = Config()
        self.table_key = table[0]
        self.table_name = table[1]
        self._get_data()
        self.frame()
        self._draw()
        self.label()
        self.checkout_button()
        self.button()
        self.entry_button()

    def _get_data(self):
        rexl = Rexcel(self.table_key)
        self.all_data, self.average_data, self.differential_data = rexl.read_excel()
        self.value = rexl.read_value()

    def _draw(self):
        self.draw_plt()
        self.draw_line()
        self.draw_canvas()

    def back(self):
        self.frame1.destroy()
        BaseFrame(self.root)

    def get_table_name(self):
        return self.cf.get_value(self.table_key)

    def entry_button(self):
        self.entry = tkinter.Entry(self.frame_bottom, show='')
        self.entry.pack()
        btn = tkinter.Button(self.frame_bottom, text='修改线性图名称', command=self.change_name)
        btn.pack()
        # self.cf.update_value(self.table_id, new_name)

    def change_name(self):
        name = self.entry.get()
        if name == '':
            return
        answer = messagebox.askyesno(title='修改名称', message=f'确定修改线性图的名称为"{name}"？')
        if answer:
            self.cf.update_value(self.table_key, name)
            self.frame1.destroy()
            LineGraphs(self.root, [self.table_key, name])

    def frame(self):
        self.root.title(f'SPC Control - {self.table_name}')
        self.frame1 = tkinter.Frame(self.root)
        self.frame1.pack(side=tkinter.TOP, fill=tkinter.BOTH, expand=1)
        self.frame_left = tkinter.Frame(self.frame1)
        self.frame_left.pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=1)
        self.frame_right = tkinter.Frame(self.frame1)
        self.frame_right.pack(side=tkinter.RIGHT, fill=tkinter.BOTH, expand=1)

        self.frame_bottom = tkinter.Frame(self.frame_right)
        self.frame_bottom.pack(side=tkinter.BOTTOM)

    def draw_plt(self):
        self.f = Figure(figsize=(6, 4), dpi=120)
        self.plt = self.f.add_subplot(111)

    def draw_line(self, abnormal=[]):
        x_data = [i + 1 for i in range(len(self.average_data))]
        # print('x_data', x_data)
        y_data = self.average_data
        # print('y_data', y_data)
        cl = self.value[0]
        ucl = self.value[3]
        lcl = self.value[8]

        self.plt.plot(x_data, [cl for i in x_data], 'g', linewidth=1)
        self.plt.plot(x_data, [ucl for i in x_data], 'r', linewidth=1)
        self.plt.plot(x_data, [lcl for i in x_data], 'r', linewidth=1)
        self.plt.plot(x_data, [self.value[4] for i in x_data], 'y--', linewidth=1)
        self.plt.plot(x_data, [self.value[5] for i in x_data], 'y--', linewidth=1)
        self.plt.plot(x_data, [self.value[6] for i in x_data], 'y--', linewidth=1)
        self.plt.plot(x_data, [self.value[7] for i in x_data], 'y--', linewidth=1)

        self.plt.plot(x_data, y_data, 'bo-', linewidth=1, label='average')
        # print('abnormal', abnormal)
        x_data_a = []
        for i in set(abnormal):
            x_data_a.append(i + 1)
        y_data_a = []
        for i in x_data_a:
            y_data_a.append(y_data[i - 1])
        # print('x_data_a', x_data_a)
        # print('y_data_a', y_data_a)
        self.plt.plot(x_data_a, y_data_a, 'ro-', linewidth=1)

    def draw_canvas(self):
        self.canvas = FigureCanvasTkAgg(self.f, master=self.frame_left)
        self.canvas.draw()
        self.canvas.get_tk_widget().pack(side=tkinter.LEFT, fill=tkinter.BOTH, expand=1)

    def show_info(self, string):
        if string == '数据异常：\n':
            messagebox.showinfo(title='result', message='测试通过，数据正常')
        else:
            messagebox.showwarning(title='result', message=string)

    def button(self):
        btn_check = tkinter.Button(self.frame_right, text="检验数据", command=self.checkout, activeforeground='green')
        btn_check.pack()
        # btn_refresh = tkinter.Button(self.root, text="刷新", command=self.refresh, activeforeground='green')
        # btn_refresh.pack()
        btn_back = tkinter.Button(self.frame_right, text='返回主窗口', command=self.back)
        btn_back.pack()

    def refresh(self):
        self.frame1.update()
        print('update')

    def label(self):
        lab = tkinter.Label(self.frame_right, text='检验规则', height=2, font='Helvetica -18 bold')
        lab.pack(side=tkinter.TOP)

    def checkout_button(self):
        rule_list = ['1.任意1点超出3σ以外',
        '2.连续9点落在中心线同一侧',
        '3.连续6点递增或递减',
        '4.连续14点中相邻点交替上下',
        '5.连续3点中有2点落在中心线同侧2σ以外',
        '6.连续5点中有4点落在中心线同侧1σ以外',
        '7.连续15点落在中心线±1σ以内',
        '8.连续8点落在中心线两侧且无一落在±1σ以内']
        self.var_list = []
        for i in range(8):
            self.var_list.append(tkinter.IntVar())
        for i in range(len(self.var_list)):
            frame_x = tkinter.Frame(self.frame_right)
            frame_x.pack(side=tkinter.TOP, fill=tkinter.BOTH)
            cb = tkinter.Checkbutton(frame_x, text=rule_list[i], variable=self.var_list[i])
            cb.pack(side=tkinter.LEFT, expand=0)

    def checkout(self):
        '''1.任意1点超出3σ以外
        2.连续9点落在中心线同一侧
        3.连续6点递增或递减
        4.连续14点中相邻点交替上下
        5.连续3点中有2点落在中心线同侧2σ以外
        6.连续5点中有4点落在中心线同侧1σ以外
        7.连续15点落在中心线±1σ以内
        8.连续8点落在中心线两侧且无一落在±1σ以内'''
        cl = self.value[0]
        ucl_1q = self.value[5]
        ucl_2q = self.value[4]
        ucl = self.value[3]
        lcl_1q = self.value[6]
        lcl_2q = self.value[7]
        lcl = self.value[8]
        data = self.average_data
        string = '数据异常：\n'

        abnormal_list = []
        rule_list = []
        for i in self.var_list:
            rule_list.append(i.get())
        # print(rule_list)

        # 1.任意1点超出3σ以外
        if rule_list[0]:
            for i in range(len(data)):
                if data[i] > ucl or data[i] < lcl:
                    string += f'第{i + 1}个数据的值{data[i]}超出了中间线3σ；\n'
                    abnormal_list.append(i)
                    # print(abnormal_list)

        # 2.连续9点落在中心线同一侧
        if rule_list[1]:
            list_2up = []
            list_2down = []
            for i in range(len(data)):
                if data[i] > cl:
                    list_2up.append(data[i])
                    list_2down.clear()
                    if len(list_2up) >= 9:
                        string += f'数据组{list_2up}超过9个数据连续落在中心线上方；\n'
                        for j in range(9):
                            abnormal_list.append(i - j)
                elif data[i] < cl:
                    list_2up.clear()
                    list_2down.append(data[i])
                    if len(list_2down) >= 9:
                        string += f'数据组{list_2down}超过9个数据连续落在中心线下方；\n'
                        for j in range(9):
                            abnormal_list.append(i - j)
                else:
                    list_2up.append(data[i])
                    list_2down.append(data[i])
                    if len(list_2up) >= 9:
                        string += f'数据组{list_2up}超过9个数据连续落在中心线上方；\n'
                        for j in range(9):
                            abnormal_list.append(i - j)
                    elif len(list_2down) >= 9:
                        string += f'数据组{list_2down}超过9个数据连续落在中心线下方；\n'
                        for j in range(9):
                            abnormal_list.append(i - j)

        # 3.连续6点递增或递减
        if rule_list[2]:
            list_3up = []
            list_3down = []
            for i in range(len(data)):
                if i == 0:
                    pass
                else:
                    if data[i] > data[i - 1]:
                        list_3up.append(data[i])
                        list_3down.clear()
                        if len(list_3up) >= 6:
                            string += f'数据组{list_3up}超过6个数连续递增；\n'
                            for j in range(6):
                                abnormal_list.append(i - j)
                    elif data[i] < data[i - 1]:
                        list_3up.clear()
                        list_3down.append(data[i])
                        if len(list_3down) >= 6:
                            string += f'数据组{list_3down}超过6个数连续递减；\n'
                            for j in range(6):
                                abnormal_list.append(i - j)
                    else:
                        list_3up.append(data[i])
                        list_3down.append(data[i])
                        if len(list_3up) >= 6:
                            string += f'数据组{list_3up}超过6个数连续递增(包括相等)；\n'
                            for j in range(6):
                                abnormal_list.append(i - j)
                        elif len(list_3down) >= 6:
                            string += f'数据组{list_3down}超过6个数连续递减(包括相等)；\n'
                            for j in range(6):
                                abnormal_list.append(i - j)

        # 4.连续14点中相邻点交替上下
        if rule_list[3]:
            list_4 = []
            for i in range(len(data)):
                if i == 0:
                    list_4.append(data[i])
                elif i == len(data)-1:
                    pass
                else:
                    if data[i] > data[i-1] and data[i] > data[i+1]:
                        list_4.append(data[i])
                        if len(list_4) >= 13:
                            list_4.append(data[i+1])
                            string += f'数据组{list_4}超过14个连续数交替上下浮动；\n'
                            for j in range(14):
                                abnormal_list.append(i - j)
                    elif data[i] < data[i-1] and data[i] < data[i+1]:
                        list_4.append(data[i])
                        if len(list_4) >= 13:
                            list_4.append(data[i+1])
                            string += f'数据组{list_4}超过14个连续数交替上下浮动；\n'
                            for j in range(14):
                                abnormal_list.append(i - j)
                    else:
                        list_4.clear()

        # 5.连续3点中有2点落在中心线同侧2σ以外
        if rule_list[4]:
            for i in range(len(data)-2):
                if data[i] > ucl_2q:
                    if data[i+1] > ucl_2q or data[i+2] > ucl_2q:
                        string += f'连续3点数据[{data[i]}, {data[i+1]}, {data[i+2]}],其中超过两点高于中心线2σ以上；\n'
                        for j in range(3):
                            abnormal_list.append(i + j)
                if data[i] < lcl_2q:
                    if data[i+1] < lcl_2q or data[i+2] < lcl_2q:
                        string += f'连续3点数据[{data[i]}, {data[i+1]}, {data[i+2]}],其中超过两点低于中心线2σ以上；\n'
                        for j in range(3):
                            abnormal_list.append(i + j)

        # 6.连续5点中有4点落在中心线同侧1σ以外
        if rule_list[5]:
            for i in range(len(data)-4):
                list_6 = []
                list_6up = []
                list_6down = []
                for j in range(5):
                    list_6.append(data[i+j])
                    if data[i+j] >= ucl_1q:
                        list_6up.append(data[i+j])
                    elif data[i+j] <= lcl_1q:
                        list_6down.append(data[i+j])
                if len(list_6up) >= 4:
                    string += f'连续5点数据{list_6},其中超过4点高于中心线1σ以上；\n'
                    for j in range(5):
                        abnormal_list.append(i + j)
                elif len(list_6down) >= 4:
                    string += f'连续5点数据{list_6},其中超过4点低于中心线1σ以上；\n'
                    for j in range(5):
                        abnormal_list.append(i + j)

        # 7.连续15点落在中心线±1σ以内
        if rule_list[6]:
            for i in range(len(data)-14):
                list_7 = []
                for n in range(15):
                    list_7.append(data[i+n])
                    if data[i+n] > ucl_1q or data[i+n] < lcl_1q:
                        break
                    elif len(list_7) == 15:
                        string += f'连续15点数据{list_7}落在中心线±1σ以内；\n'
                        for j in range(15):
                            abnormal_list.append(i + j)

        # 8.连续8点落在中心线两侧且无一落在±1σ以内
        if rule_list[7]:
            for i in range(len(data)-7):
                list_8 = []
                for n in range(8):
                    list_8.append(data[i+n])
                    if lcl_1q < data[i+n] < ucl_1q:
                        break
                    elif len(list_8) == 8:
                        string += f'连续8点数据{list_8}无一落在中心线±1σ以内；\n'
                        for j in range(8):
                            abnormal_list.append(i + j)

        self.show_info(string)
        self.draw_line(abnormal_list)
        self.canvas.draw()


class Rexcel:

    def __init__(self, sheet_name):
        self.cf = Config()
        self.data = xlrd.open_workbook(self.cf.get_excel())
        # self.table = self.data.sheets()[table]
        self.table = self.data.sheet_by_name(sheet_name)
        self.ncols = self.table.ncols

    def read_excel(self):
        all_data = []
        average_data = []
        differential_data = []
        for i in range(self.ncols):
            if i == 0:
                pass
            else:
                data = self.table.col_values(colx=i, start_rowx=2, end_rowx=7)
                all_data.append(data)
                average_data.append(self.get_average(data))
                differential_data.append(self.get_differential(data))
        # print(all_data)
        # print(average_data)
        # print(differential_data)
        return all_data, average_data, differential_data

    def read_value(self):
        value_a = self.table.row_values(rowx=self.table.nrows - 1, end_colx=9)
        value = [float('{:.4f}'.format(i)) for i in value_a]
        # print('value', value)
        return value

    def read_sheets(self):
        a = self.data.get_sheets()
        print(a)

    def get_average(self, list):
        data_sum = 0
        for i in list:
            data_sum += i
        return float(format(data_sum/5, '.4f'))

    def get_differential(self, list):
        differential = max(list) - min(list)
        return float(format(differential, '.3f'))


class Config:

    def __init__(self):
        self.cf = NewConfigParser(allow_no_value=True)
        self.cf.read('config.ini')

    def get_excel(self):
        return self.cf.get('File', 'Excel')

    def get_value(self, key):
        return self.cf.get('Button', key)

    def update_value(self, key, value):
        self.cf.set('Button', key, value)
        self.cf.write(open('config.ini', 'w'))

    def remove(self, key):
        self.cf.remove_option('Button', key)
        self.cf.write(open('config.ini', 'w'))

    def get_all(self):
        return self.cf.items('Button')

    def get_all_key(self):
        return self.cf.options('Button')


if __name__ == '__main__':
    root = tkinter.Tk()
    BaseWindow(root)
    root.mainloop()
    # x = Rexcel('Sheet1')
    # print(x.read_value())
    # c = Config()
    # a = c.get_all()
    # print(len(a))
    # cf = Config()
    # cf.update_value(5, 'eee')
    # a = cf.get_all_key()
    # print(a)
    # cf.remove(5)
    # x = Rexcel('Sheet1')
    # print(x.read_value())
