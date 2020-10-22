# coding:utf-8

import os
import docx
import pandas as pd
import win32com.client as win32

df_name = pd.DataFrame(pd.read_excel('NameList.xlsx'))
docx_count = int(len(df_name) / 3) + 1

# 合并
word = win32.gencache.EnsureDispatch("Word.Application")
word.Visible = True
new_document = word.Documents.Add()
for i in range(docx_count):
    new_document.Application.Selection.Range.InsertFile(os.path.join(os.getcwd(), 'Temp.docx'))
new_document.SaveAs(os.path.join(os.getcwd(), 'Final.docx'))
new_document.Close()
word.Quit()

# 覆盖值
doc = docx.Document(r'Final.docx')
index = 0
for child in doc.element.body.iter():
    if child.tag.endswith('AlternateContent'):
        for c in child.iter():
            if c.tag.endswith('main}r'):
                if int(index / 4) < len(df_name):
                    _tmp = df_name.iat[int(index / 4), index % 2]
                    c.text = _tmp if len(_tmp) != 2 else _tmp[0] + ' ' + _tmp[1]
                    index += 1
doc.save(r'Final.docx')
