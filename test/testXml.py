# !/usr/python
# -*- coding:utf-8 -*-

from xlc import xmlExcel

e = xmlExcel.xmlExcel.parseXML("../test.xml")
excel = e.createExcelTp()
for sheet in excel:
    print(sheet[0])
    print(sheet[2])