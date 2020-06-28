# coding:utf-8
"""
BSD 2-Clause License

Copyright (c) 2020, Sea Zhou (WeChat: uefi64)
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this
   list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.
"""

"""
Name        : LcfcExcelFilter.py
Usage       : LcfcExcelFilter.py
EXE Package : pyinstaller -F -w -i LCFC.ico LcfcExcelFilter.py
Description : Filter the data file by script file, and output the final result to result file
Author      : dahai.zhou@lcfuturecenter.com
Change      :
Data           Author        Version    Description
2020/06/10     dahai.zhou    1.00       Initial release
(please search and modify tool_version and tool_release_date if update version)
"""

import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
import pandas as pd
import xlwings
import pythoncom
import openpyxl  # pandas may need it
import os
import threading
from LcfcGetRawFileLib import lcfc_get_raw_file

tool_version = 'V1.00'
tool_release_date = '2020/06/10'


def process_special_function(df_data, df_script, result_file, wb, ws, app):
    """
    Return:
    0  = Hit and Success
    1  = Hit but error
    2  = Hit and user cancel
    50 = No hit
    """
    if len(df_script) != 1:
        return 50
    df1 = df_script.dropna(axis=1)
    if len(df1.columns) != 1:
        return 50
    if df1.iloc[0, 0] != '{DivideByMe}':
        return 50

    # Here process the {DivideByMe} command
    divide_column = df1.columns[0]
    divide_group = df_data[divide_column].value_counts()
    tmp = len(divide_group) + len(df_data[df_data[divide_column].isnull()])
    if tmp >= 10:
        if not tk.messagebox.askokcancel('',
                                         '该操作将会生成%d个文件, 确认吗?' % tmp,
                                         default=tk.messagebox.CANCEL):
            l_4.config(text='用户取消', bg='red', fg='white', width=10, font=('Arial', 16))
            wb.close()
            app.quit()
            return 2

    for i in divide_group.keys():
        rng = ws.range('A2').expand('table')
        rng.api.EntireRow.Delete()
        data_to_write = df_data[df_data[divide_column] == i].values.tolist()
        ws.range('A2').expand('table').value = data_to_write
        try:
            wb.save(result_file[:-5] + '_' + i + result_file[-5:])
        except:
            error_handling(wb, app, '错误: 保存文件失败, 可能是路径不存在或者文件被使用中')
            return 1

    # Process None data in the divide_column
    if len(df_data[df_data[divide_column].isnull()]) != 0:
        ws.range('A2').expand('table').value = df_data[df_data[divide_column].isnull()].values.tolist()
        wb.save(result_file[:-5] + '_' + result_file[-5:])

    # Hit and success here
    success_handling(wb, app)
    return 0


def error_handling(wb, app, error_str):
    l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
    tk.messagebox.showerror('', error_str)
    wb.close()
    app.quit()


def success_handling(wb, app):
    l_4.config(text='处理成功', bg='green', fg='white', width=10, font=('Arial', 16))
    tk.messagebox.showinfo('', '处理完成!')
    wb.close()
    app.quit()


def b_1_process():
    f = tk.filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')])
    if f != '':
        entry_1_text.set(f)


def b_2_process():
    f = tk.filedialog.askopenfilename(filetypes=[('Excel', '*.xlsx')])
    if f != '':
        entry_2_text.set(f)


def b_3_process():
    f = tk.filedialog.asksaveasfilename(defaultextension='*.xlsx', initialfile='Result.xlsx',
                                        filetypes=[('Excel', '*.xlsx')])
    if f != '':
        entry_3_text.set(f)


def b_4_1_process():
    help_str = \
        '''本工具会去将数据文件的内容按照脚本文件来进行过滤, 最终生成一个过滤之后的结果文件. \
使用场景是有些Excel文件太大, 直接打开会卡顿, 此时可以利用该工具.
我们将sample脚本文件生成在该处, 请查阅使用说明: %s\\Script.xlsx\n
%s    %s    by RD\\dahai.zhou''' % (os.getcwd(), tool_version, tool_release_date)
    tk.messagebox.showinfo('帮助', help_str)
    tmp = open("Script.xlsx", "wb+")
    tmp.write(lcfc_get_raw_file(3))
    tmp.close()


def b_4_2_process():
    t = threading.Thread(target=b_4_2_process_thread)
    t.start()


def b_4_2_process_thread():
    l_4.config(text='处理中...', bg='yellow', fg='white', width=10, font=('Arial', 16))
    data_file = entry_1_text.get()
    script_file = entry_2_text.get()
    result_file = entry_3_text.get()
    if data_file == '' or script_file == '' or result_file == '':
        l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
        tk.messagebox.showerror('', '错误: 请输入完整的信息')
        return

    # Read data to DF
    pythoncom.CoInitialize()  # This is a must for multi-thread
    app = xlwings.App(visible=False, add_book=False)
    wb = app.books.open(data_file)
    ws = wb.sheets[0]
    ws_data = ws.range('A2').expand('table').value
    ws_header = ws.range('A1').expand('right').value
    if len(ws_data) != ws.used_range.shape[0] - 1:
        error_handling(wb, app, '错误: 数据文件第一列内包含空值, 无法处理, 请修正后再继续')
        return
    if len(ws_data[0]) != ws.used_range.shape[1] or \
            len(ws_data[0]) != len(ws_header):
        error_handling(wb, app, '错误: 数据文件第一行或第二行内包含空值, 无法处理, 请修正后再继续')
        return
    df_data = pd.DataFrame(ws_data, columns=ws.range('A1').expand('right').value)
    df_script = pd.DataFrame(pd.read_excel(script_file))

    try:
        for i in df_script.columns:
            temp = df_data[i]
    except KeyError:
        error_handling(wb, app, '错误: 数据文件和脚本文件的第一行不匹配')
        return

    # Process hidden function, if hit, then skip the rest code, error handling is in the routine
    result = process_special_function(df_data, df_script, result_file, wb, ws, app)
    if result == 0 or result == 1 or result == 2:
        return
    elif result == 50:  # No hit, continue
        pass

    # Process data
    df_result = df_data
    for i in df_script.columns:
        if not df_script[i].dropna().empty:
            condition_series = (df_result[i] != df_result[i])  # 若有任何条件, 则默认全False
        else:
            condition_series = (df_result[i] == df_result[i])  # 若无任何条件, 则默认全True
        for j in df_script[i]:
            if pd.isna(j):
                continue
            condition_series = condition_series | (df_result[i] == j)
            # print (j, type(j), df_result['ENDTIME'][0], type(df_result['ENDTIME'][0]), df_result['ENDTIME'][0] == j)
        df_result = df_result[condition_series]

    # Save data to new file
    #df_result.to_excel(result_file)
    rng = ws.range('A2').expand('table')
    rng.api.EntireRow.Delete()
    ws.range('A2').expand('table').value = df_result.values.tolist()
    try:
        wb.save(result_file)
    except:
        error_handling(wb, app, '错误: 保存文件失败, 可能是路径不存在或者文件被使用中')
        return

    # Success here
    success_handling(wb, app)
    return


# main
window = tk.Tk()
window.title('LCFC Excel Filter %s' % tool_version)
window.geometry('500x300')

with open('tmp.png', 'wb') as tmp:
    tmp.write(lcfc_get_raw_file(1))
window.iconphoto(True, tk.PhotoImage(file='tmp.png'))
os.remove('tmp.png')

window.resizable(0, 0)

l_1 = tk.Label(window, text='请选择原始数据文件:')
l_1.place(x=20, y=20)
entry_1_text = tk.StringVar()
#entry_1_text.set(r'C:\Users\Sea\Desktop\HR\LcfcExcelFilter\DL班次_2017&08_1.xlsx')
entry_1 = tk.Entry(window, width=55, textvariable=entry_1_text)
entry_1['state'] = 'readonly'
entry_1.place(x=20, y=40)
b_1 = tk.Button(window, width=6, text="浏览", command=b_1_process)
b_1.place(x=420, y=35)

l_2 = tk.Label(window, text='请选择脚本文件:')
l_2.place(x=20, y=80)
entry_2_text = tk.StringVar()
#entry_2_text.set(r'C:\Users\Sea\Desktop\HR\LcfcExcelFilter\Script.xlsx')
entry_2 = tk.Entry(window, width=55, textvariable=entry_2_text)
entry_2['state'] = 'readonly'
entry_2.place(x=20, y=100)
b_2 = tk.Button(window, width=6, text="浏览", command=b_2_process)
b_2.place(x=420, y=95)

l_3 = tk.Label(window, text='生成文件名(软件同目录下):')
l_3.place(x=20, y=140)
entry_3_text = tk.StringVar()
entry_3_text.set(r'Result.xlsx')
entry_3 = tk.Entry(window, width=55, textvariable=entry_3_text)
# entry_3['state'] = 'readonly'
entry_3.place(x=20, y=160)
# b_3 = tk.Button(window, width=6, text="浏览", command=b_3_process)
# b_3.place(x=420, y=155)

l_4 = tk.Label(window, text='', width=10, font=('Arial', 16))
l_4.place(x=20, y=220)
b_4_1 = tk.Button(window, width=6, bd=3, text="帮助", command=b_4_1_process)
b_4_1.place(x=320, y=215)
b_4_2 = tk.Button(window, width=6, bd=3, text="开始处理", command=b_4_2_process)
b_4_2.place(x=420, y=215)

window.withdraw()
window.deiconify()
window.mainloop()
