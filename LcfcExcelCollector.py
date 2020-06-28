# coding:utf-8
"""
BSD 2-Clause License

Copyright (c) 2020, Sea Zhou, WeChat: uefi64
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
Name        : LcfcExcelCollector.py
Usage       : LcfcExcelCollector.py
EXE Package : pyinstaller -F -w -i LCFC.ico LcfcExcelCollector.py
Description : Collect info from all Excel files in the target folder, and merge into a new one
              according the rule described in the script file
Author      : dahai.zhou@lcfuturecenter.com
Change      :
Data           Author        Version    Description
2020/05/25     dahai.zhou    1.00       Initial release
2020/05/26     dahai.zhou    1.01       Switch the position of select script and select folder
2020/06/01     dahai.zhou    1.02       Remove the "生成文件" button
(please search and modify tool_version and tool_release_date if update version)
"""

import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog
import pandas as pd
import win32com.client as win32
import openpyxl  # pandas may need it
import os
import threading
from LcfcGetRawFileLib import lcfc_get_raw_file

tool_version = 'V1.02'
tool_release_date = '2020/06/01'


def b_1_process():
    f = tk.filedialog.askdirectory()
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
        '''本工具会去数据文件夹里面搜寻所有的Excel文件, 并根据脚本文件\
的内容来处理所有的数据, 最终生成一个结果文件.
我们将sample脚本文件生成在该处, 请查阅使用说明: %s\\Script.xlsx\n
%s    %s    by RD\\dahai.zhou''' % (os.getcwd(), tool_version, tool_release_date)
    tk.messagebox.showinfo('帮助', help_str)
    tmp = open("Script.xlsx", "wb+")
    tmp.write(lcfc_get_raw_file(2))
    tmp.close()


def b_4_2_process():
    t = threading.Thread(target=b_4_2_process_thread)
    t.start()


def b_4_2_process_thread():
    l_4.config(text='处理中...', bg='yellow', fg='white', width=10, font=('Arial', 16))
    data_folder = entry_1_text.get()
    script_file = entry_2_text.get()
    result_file = entry_3_text.get()
    if data_folder == '' or script_file == '' or result_file == '':
        l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
        tk.messagebox.showerror('', '错误: 请输入完整的信息')
        return

    try:
        df = pd.DataFrame(pd.read_excel(script_file))
    except:
        l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
        tk.messagebox.showerror('', '错误: 脚本文件读取出错')
        return

    try:
        for i in df.columns:
            sheet_location = df.loc[0, i]
            sheet = sheet_location.split('|')[0]
            lcoation = sheet_location.split('|')[1]
    except:
        l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
        tk.messagebox.showerror('', '错误: 脚本文件解析出错')
        return

    df_result = df.copy()
    excel = win32.DispatchEx('Excel.Application')

    file_list = os.listdir(data_folder)
    for i in file_list:
        try:
            wb = excel.Workbooks.Open(data_folder + '\\' + i)
        except:
            excel.Application.Quit()
            l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
            tk.messagebox.showerror('', '解析出错: %s不是文件' % i)
            return
        for j in df.columns:
            sheet_location = df.loc[0, j]
            sheet = sheet_location.split('|')[0]
            lcoation = sheet_location.split('|')[1]
            try:
                ws = wb.Worksheets(sheet)
            except:
                wb.Close()
                excel.Application.Quit()
                l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
                tk.messagebox.showerror('', '解析出错: 文件%s找不到%s工作簿' % (i, sheet))
                return
            ws.Activate()
            try:
                df_result.loc[file_list.index(i), j] = ws.Range(lcoation).Value
            except:
                wb.Close()
                excel.Application.Quit()
                l_4.config(text='处理失败', bg='red', fg='white', width=10, font=('Arial', 16))
                tk.messagebox.showerror('', '解析出错: 文件%s找不到%s单元格' % (i, lcoation))
                return
        wb.Close()

    df_result.to_excel(result_file)
    excel.Application.Quit()
    l_4.config(text='处理成功', bg='green', fg='white', width=10, font=('Arial', 16))
    tk.messagebox.showinfo('', '处理完成!')
    return


# main
window = tk.Tk()
window.title('LCFC Excel Collector %s' % tool_version)
window.geometry('500x300')

with open('tmp.png', 'wb') as tmp:
    tmp.write(lcfc_get_raw_file(1))
window.iconphoto(True, tk.PhotoImage(file='tmp.png'))
os.remove('tmp.png')

window.resizable(0, 0)

l_1 = tk.Label(window, text='请选择数据所在文件夹:')
l_1.place(x=20, y=20)
entry_1_text = tk.StringVar()
entry_1 = tk.Entry(window, width=55, textvariable=entry_1_text)
entry_1['state'] = 'readonly'
entry_1.place(x=20, y=40)
b_1 = tk.Button(window, width=6, text="浏览", command=b_1_process)
b_1.place(x=420, y=35)

l_2 = tk.Label(window, text='请选择脚本文件:')
l_2.place(x=20, y=80)
entry_2_text = tk.StringVar()
entry_2 = tk.Entry(window, width=55, textvariable=entry_2_text)
entry_2['state'] = 'readonly'
entry_2.place(x=20, y=100)
b_2 = tk.Button(window, width=6, text="浏览", command=b_2_process)
b_2.place(x=420, y=95)

l_3 = tk.Label(window, text='生成文件名:')
l_3.place(x=20, y=140)
entry_3_text = tk.StringVar()
entry_3_text.set('Result.xlsx')
entry_3 = tk.Entry(window, width=55, textvariable=entry_3_text)
#entry_3['state'] = 'readonly'
entry_3.place(x=20, y=160)
#b_3 = tk.Button(window, width=6, text="浏览", command=b_3_process)
#b_3.place(x=420, y=155)

l_4 = tk.Label(window, text='', width=10, font=('Arial', 16))
l_4.place(x=20, y=220)
b_4_1 = tk.Button(window, width=6, bd=3, text="帮助", command=b_4_1_process)
b_4_1.place(x=320, y=215)
b_4_2 = tk.Button(window, width=6, bd=3, text="开始处理", command=b_4_2_process)
b_4_2.place(x=420, y=215)

window.withdraw()
window.deiconify()
window.mainloop()
