# !/usr/python
# -*- coding:utf-8 -*-

from xlc import excel
from xlc import xmlExcel

e = xmlExcel.xmlExcel.parseXML("../test.xml")
ef = e.createExcelTp()
for sheet in ef:
    print(sheet[0])
    print(sheet[2])
print('>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>')
file = excel.excel("./test.xlsx", "../test.xml")
test_data = []
for i in range(11):
    test_data.append(i)
for i in range(11):
    file.addRow('挂单记录','挂单表',test_data)