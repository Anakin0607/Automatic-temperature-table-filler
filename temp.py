# -*- coding: utf-8 -*-
import docx
import random
path='temp.docx'
doc=docx.Document(path)
tbs=doc.tables
tb=tbs[0]
def fill(a,Text):
    row=tb.rows[a]
    for cell in row.cells[1:]:
        cell.text=Text
    return

def Rand(a):
    row=tb.rows[a]
    for cell in row.cells[1:]:
        tmp=random.uniform(36.2,36.8)
        cell.text=str(round(tmp,1))
    return
Name=u'Anakin0607' #把这里修改成测量人
Place=u'家' #把这里修改成活动地点，修改这两处时只修改引号内的内容即可，请勿将u一并修改
Rand(3)
Rand(5)
fill(4,Name)
fill(6,Name)
fill(7,Place)
doc.save(path)


