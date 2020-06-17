# coding:utf-8
"""
*****************************************************************************
*  Copyright (c) 2012 - 2020, Hefei LCFC Information Technology Co.Ltd.
*  And/or its affiliates. All rights reserved.
*  Hefei LCFC Information Technology Co.Ltd. PROPRIETARY/CONFIDENTIAL.
*  Use is subject to license terms.
*****************************************************************************
Name        : LcfcMailer.py
Usage       : LcfcMailer.py (It will fetch MailList.xlsx in the same folder)
EXE Package : pyinstaller -F -i LCFC.ico LcfcMailer.py
Description : Get info from script file MailList.xlsx, and send emails automaticlly via Outlook
Author      : dahai.zhou@lcfuturecenter.com
Change      :
Data           Author        Version    Description
2020/05/18     dahai.zhou    1.00       Initial release
2020/05/20     dahai.zhou    1.01       Support multi-attachment
                                        Add author name
2020/06/15     dahai.zhou    1.02       Generate template script if not exist
2020/06/17     dahai.zhou    1.03       Optimize
(please search and modify tool_version and tool_release_date if update version)
"""

import pandas as pd
import win32com.client as win32
import xlrd  # pandas may need it
import time
import sys
import os
from LcfcGetRawFileLib import lcfc_get_raw_file

tool_version = 'V1.03'
tool_release_date = '2020/06/17'

# 全局变量区, 不要关心他们的值, 后面会被覆盖掉
mail_addressee = 'dahai.zhou;'  # 收件人邮箱列表
mail_cc = 'dahai.zhou;'  # 抄送邮箱列表
mail_bcc = 'dahai.zhou;'  # 密件抄送邮箱列表
mail_subject = '[Python培训]HR+SCM实践期任务'  # 主题
mail_content = 'html format content'  # 正文
mail_attachment_path = [r"D:\Outlook.txt"]  # 附件列表


def send_email(o):
    mail = o.CreateItem(0)
    mail.To = '' if pd.isna(mail_addressee) else mail_addressee
    mail.CC = '' if pd.isna(mail_cc) else mail_cc
    mail.BCC = '' if pd.isna(mail_bcc) else mail_bcc
    mail.Subject = mail_subject
    for i in mail_attachment_path:
        mail.Attachments.Add(i)
    mail.HTMLBody = mail_content
    mail.Send()


# main
print('''
 =====================================================
|                 LCFC Mailer %s                   |
|                                       %s    |
|                                       RD/dahai.zhou |
 =====================================================
''' % (tool_version, tool_release_date))
try:
    df = pd.DataFrame(pd.read_excel(r'MailList.xlsx'))
except:
    print("没有找到脚本文件MailList.xlsx")
    with open('MailList.xlsx', 'wb') as tmp:
        tmp.write(lcfc_get_raw_file(4))
        print("生成脚本文件在该处, 请编辑后使用: %s\\MailList.xlsx" % (os.getcwd()))
    os.system("pause")
    sys.exit(1)

first_in = True
for i in df.index:
    mail_addressee = df.loc[i, '收件人']
    mail_cc = df.loc[i, '抄送']
    mail_bcc = df.loc[i, '密件抄送']

    mail_subject = df.loc[i, '主题']
    mail_subject_var = df.loc[i, '主题变量']
    if not pd.isna(mail_subject_var):
        mail_subject_var_list = mail_subject_var.split('|')
        mail_subject = mail_subject.format(*mail_subject_var_list)

    mail_content_path = df.loc[i, '正文']
    read = open(mail_content_path, encoding='gb2312')  # 打开需要发送的文件
    mail_content = read.read()  # 读取html文件中的内容
    read.close()
    mail_content_var = df.loc[i, '正文变量']
    if not pd.isna(mail_content_var):
        mail_subject_var_list = mail_content_var.split('|')
        for j in mail_subject_var_list:
            org_str = '{' + str(mail_subject_var_list.index(j)) + '}'
            mail_content = mail_content.replace(org_str, j)

    mail_attachment_var = df.loc[i, '附件']
    if not pd.isna(mail_attachment_var):
        mail_attachment_path = mail_attachment_var.split('|')
    else:
        mail_attachment_path = []

    if first_in:
        first_in = False
        print("找到脚本并解析成功, 即将开始发送邮件, 请确保Outlook处于打开状态\n按任意键开始发送, 关闭程序可以取消发送")
        os.system("pause")
    outlook = win32.Dispatch("outlook.Application")
    send_email(outlook)
    print("发送邮件%d成功!" % i)
    # We will not exit Outlook because it should be always opened for the most
    time.sleep(1)

print('总共发送了%d封邮件' % len(df))
os.system("pause")
