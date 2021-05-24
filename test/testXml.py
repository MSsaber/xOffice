# !/usr/python
# -*- coding:utf-8 -*-

import uuid
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
test_uuid = ''
for i in range(11):
    test_data.append(i)
for i in range(11):
    test_data[2] = str(uuid.uuid4())
    if i == 4:
        test_uuid = test_data[2]
    file.addRow('挂单记录','挂单表',test_data)
print("resRow(" + test_uuid +") : " + str(file.findRow('挂单记录','挂单表','单号',test_uuid)))
print("resCol(" + test_uuid +") : " + str(file.findCol('挂单记录','挂单表','单号')))
file.setCell('挂单记录','挂单表','挂单数量', 4, 100)
test_data.clear()
for i in range(5):
    test_data.append(i*20)
print(file.setDatas('挂单记录','挂单表',6, 4, test_data))