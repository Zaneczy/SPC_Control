#!/usr/bin/python3
import tkinter
from tkinter import *
import matplotlib.pyplot as plt
import xlrd
import numpy as np

def getdata():
    data = xlrd.open_workbook(r'data.xlsx')
    # 获取表数据
    table = data.sheets()[0]
    print(table)
    # 获取表名
    sheel_name = data.sheet_names()
    print(sheel_name)
    # 获取表中的有效行数
    nrows = table.nrows
    print(nrows)
    # 获取表中第3行的数据
    value_a = table.row_values(rowx=2)
    print(value_a)
    value_b = table.row_values(rowx=2, start_colx=1)
    print(value_b)
    # 获取表中的有效列数
    ncols = table.ncols
    print(ncols)
    # 获取表中第2列的数据
    col_values_a = table.col_values(colx=1)
    print(col_values_a)
    col_values_b = table.col_values(colx=1, start_rowx=2, end_rowx=7)
    print(col_values_b)
    all_data = []
    average_data = []
    differential_data = []
    for i in range(ncols):
        if i == 0:
            pass
        else:
            data = table.col_values(colx=i, start_rowx=2, end_rowx=7)
            all_data.append(data)
            average_data.append(get_average(data))
            differential_data.append(get_differential(data))
    print('all_data', all_data)
    print('average_data', average_data)
    print('differential_data', differential_data)
    return all_data, average_data, differential_data


def draw(average_list):
    x_data = [i + 1 for i in range(len(average_list))]
    print('x_data', x_data)
    y_data = list(map(float, average_list))
    print(y_data)
    plt.plot(x_data, y_data, 'bo-', linewidth=1, label='average')
    plt.plot(x_data, [100 for i in x_data], 'g', linewidth=1)
    plt.plot(x_data, [99.96 for i in x_data], 'y--', linewidth=1)
    plt.plot(x_data, [99.98 for i in x_data], 'y--', linewidth=1)
    plt.plot(x_data, [100.02 for i in x_data], 'y--', linewidth=1)
    plt.plot(x_data, [100.04 for i in x_data], 'y--', linewidth=1)
    plt.plot(x_data, [99.94 for i in x_data], 'r', linewidth=1)
    plt.plot(x_data, [100.06 for i in x_data], 'r', linewidth=1)
    plt.title('SPC')
    plt.legend()
    plt.xlabel('x')
    plt.ylabel('y')
    my_x_ticks = np.arange(0, 26, 1)
    my_y_ticks = np.arange(99.9, 100.1, 0.02)
    print(my_x_ticks)
    print(my_y_ticks)
    plt.xticks(my_x_ticks)
    plt.yticks(my_y_ticks)
    # plt.savefig('spc.png')
    plt.show()


def get_average(list):
    data_sum = 0
    for i in list:
        data_sum += i
    return float(format(data_sum/5, '.4f'))

def get_differential(list):
    differential = max(list) - min(list)
    return float('{:.3f}'.format(differential))


def tkt():
    window = Tk()
    window.title("First Window")
    window.geometry("600x400")
    lbl = Label(window, text="Hello")
    lbl.grid(column=0, row=0)

    def clicked():
        lbl.configure(text="Button was clicked!")

    btn = Button(window, text="Click Me", command=clicked)
    btn.grid(column=1, row=0)
    window.mainloop()


def read_excel():
    data = xlrd.open_workbook(r'test_data.xlsx')
    table = data.sheets()[0]
    print(table.nrows)
    print(table.ncols)

    value_a = table.row_values(rowx=table.nrows-1, end_colx=9)
    key_data = [float('{:.4f}'.format(i)) for i in value_a]
    print('key_data', key_data)

    all_data = []
    average_data = []
    differential_data = []
    for i in range(table.ncols):
        if i == 0:
            pass
        else:
            data = table.col_values(colx=i, start_rowx=2, end_rowx=7)
            all_data.append(data)
            average_data.append(get_average(data))
            differential_data.append(get_differential(data))
    print('all_data', all_data)
    print('average_data', average_data)
    print('differential_data', differential_data)
    return key_data, all_data, average_data, differential_data

def draw_l(average_data, key_data):
    x_data = [i + 1 for i in range(len(average_data))]
    print('x_data', x_data)
    y_data = average_data
    print('y_data', y_data)
    cl = key_data[0]
    ucl = key_data[3]
    lcl = key_data[8]
    plt.plot(x_data, y_data, 'bo-', linewidth=1, label='average')
    plt.plot(x_data, [cl for i in x_data], 'g', linewidth=1)
    plt.plot(x_data, [ucl for i in x_data], 'r', linewidth=1)
    plt.plot(x_data, [lcl for i in x_data], 'r', linewidth=1)
    plt.plot(x_data, [key_data[4] for i in x_data], 'y--', linewidth=1)
    plt.plot(x_data, [key_data[5] for i in x_data], 'y--', linewidth=1)
    plt.plot(x_data, [key_data[6] for i in x_data], 'y--', linewidth=1)
    plt.plot(x_data, [key_data[7] for i in x_data], 'y--', linewidth=1)
    plt.title('SPC')
    plt.legend()
    plt.xlabel('x')
    plt.ylabel('y')
    # my_x_ticks = np.arange(0, 26, 1)
    # my_y_ticks = np.arange(99.9, 100.1, 0.02)
    # print(my_x_ticks)
    # print(my_y_ticks)
    # plt.xticks(my_x_ticks)
    # plt.yticks(my_y_ticks)
    # plt.savefig('spc.png')
    plt.show()

def checkout(average_data, value):
    '''1.任意1点超出3σ以外
    2.连续9点落在中心线同一侧
    3.连续6点递增或递减
    4.连续14点中相邻点交替上下
    5.连续3点中有2点落在中心线同侧2σ以外
    6.连续5点中有4点落在中心线同侧1σ以外
    7.连续15点落在中心线±1σ以内
    8.连续8点落在中心线两侧且无一落在±1σ以内'''
    cl = value[0]
    ucl_1q = value[5]
    ucl_2q = value[4]
    ucl = value[3]
    lcl_1q = value[6]
    lcl_2q = value[7]
    lcl = value[8]
    data = average_data
    string = '数据异常：\n'


    # 1.任意1点超出3σ以外
    for i in range(len(data)):
        if data[i] > ucl or data[i] < lcl:
            string += f'第{i+1}组数据的值{data[i]}超出了中间线3σ\n'

    # 2.连续9点落在中心线同一侧
    list_2up = []
    list_2down = []
    for i in range(len(data)):
        if data[i] > cl:
            list_2up.append(data[i])
            list_2down.clear()
            if len(list_2up) >= 9:
                string += f'数据组{list_2up}超过9个数据连续落在中心线上方\n'
        elif data[i] < cl:
            list_2up.clear()
            list_2down.append(data[i])
            if len(list_2down) >= 9:
                string += f'数据组{list_2down}超过9个数据连续落在中心线下方\n'
        else:
            list_2up.append(data[i])
            list_2down.append(data[i])
            if len(list_2up) >= 9:
                string += f'数据组{list_2up}超过9个数据连续落在中心线上方\n'
            elif len(list_2down) >= 9:
                string += f'数据组{list_2down}超过9个数据连续落在中心线下方\n'

    # 3.连续6点递增或递减
    list_3up = []
    list_3down = []
    for i in range(len(data)):
        if i == 0:
            pass
        else:
            if data[i] > data[i-1]:
                list_3up.append(data[i])
                list_3down.clear()
                if len(list_3up) >= 6:
                    string += f'数据组{list_3up}超过6个数连续递增\n'
            elif data[i] < data[i-1]:
                list_3up.clear()
                list_3down.append(data[i])
                if len(list_3down) >= 6:
                    string += f'数据组{list_3down}超过6个数连续递减\n'
            else:
                list_3up.append(data[i])
                list_3down.append(data[i])
                if len(list_3up) >= 6:
                    string += f'数据组{list_3up}超过6个数连续递增\n'
                elif len(list_3down) >= 6:
                    string += f'数据组{list_3down}超过6个数连续递减\n'

    # 4.连续14点中相邻点交替上下
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
                    string += f'数据组{list_4}超过14个连续数交替上下浮动\n'
            elif data[i] < data[i-1] and data[i] < data[i+1]:
                list_4.append(data[i])
                if len(list_4) >= 13:
                    list_4.append(data[i+1])
                    string += f'数据组{list_4}超过14个连续数交替上下浮动\n'
            else:
                list_4.clear()

    # 5.连续3点中有2点落在中心线同侧2σ以外
    for i in range(len(data)-2):
        if data[i] > ucl_2q:
            if data[i+1] > ucl_2q or data[i+2] > ucl_2q:
                string += f'连续3点数据[{data[i]}, {data[i+1]}, {data[i+2]}],其中超过两点高于中心线2σ以上\n'
        if data[i] < lcl_2q:
            if data[i+1] < lcl_2q or data[i+2] < lcl_2q:
                string += f'连续3点数据[{data[i]}, {data[i+1]}, {data[i+2]}],其中超过两点低于中心线2σ以上\n'

    # 6.连续5点中有4点落在中心线同侧1σ以外
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
            string += f'连续5点数据{list_6},其中超过4点高于中心线1σ以上\n'
        elif len(list_6down) >= 4:
            string += f'连续5点数据{list_6},其中超过4点低于中心线1σ以上\n'

    # 7.连续15点落在中心线±1σ以内
    for i in range(len(data)-14):
        list_7 = []
        for j in range(15):
            list_7.append(data[i+j])
            if data[i+j] > ucl_1q or data[i+j] < lcl_1q:
                break
            elif len(list_7) == 15:
                string += f'连续15点数据{list_7}落在中心线±1σ以内\n'

    # 8.连续8点落在中心线两侧且无一落在±1σ以内
    for i in range(len(data)-7):
        list_8 = []
        for j in range(8):
            list_8.append(data[i+j])
            if lcl_1q < data[i+j] < ucl_1q:
                break
            elif len(list_8) == 8:
                string += f'连续8点数据{list_8}无一落在中心线±1σ以内\n'
    print(string)


def test():
    # data = [100.0112, 99.913, 99.9746, 99.8932, 100.0196, 99.9302, 100.001, 99.9974, 100.0054, 99.9712, 99.9968, 99.9644, 100.0094, 99.9964, 99.9748, 99.9896, 100.037, 100.0038, 99.9874, 99.9602, 99.9992, 100.0254, 100.053, 99.991, 100.0272]
    # for i in range(len(data)-4):
    #     list_6a = []
    #     for j in range(5):
    #         list_6a.append(data[i+j])
    dict_1 = {}
    dict_1.update({5:88})
    dict_1.update({6:54})
    print(list(dict_1.keys()))
    print(dict_1.values())
    test1(dict_1)
    test1()

def test1(a={}):
    print('kwargs', a)


def windows1():
    master = tkinter.Tk()
    master.title('窗口1')
    master.geometry('500x300+500+300')
    but = tkinter.Button(text=string,command=windows2)
    but.pack()
    master.mainloop()

string='窗口2字符'

def windows2():
    master = tkinter.Tk()
    master.title(string)
    master.geometry('500x300')
    master.mainloop()

if __name__ == '__main__':
    # all_list, average_list, differential_list = getdata()
    # draw(average_list)
    # value, all_data, average_data, differential_data = read_excel()
    # draw_l(average_data, value)
    # checkout(average_data, value)
    # test()
    windows1()
