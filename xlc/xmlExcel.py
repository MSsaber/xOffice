# !/usr/python
# -*- coding:utf-8 -*-

import xml,sys
from xlc import xmlSheet
from xml.etree.ElementTree import parse

def testNode(node):
    print(node.tag)
    print(node.attrib)
    if sys.version_info.minor < 9:
        print(node.getchildren())
    else:
        print(list(node))
    if len(node.text): print(node.text)

class xmlExcel:
    def __init__(self, **argparam):
        self.sheets = []
        if 'name' in argparam.keys():
            self.name = argparam['name']
        if 'attrib' in argparam.keys():
            self.attrib = argparam['attrib']
        if 'nodes' in argparam.keys():
            self.nodes = argparam['nodes']
            self.__genSheet(self.nodes)

    def __genSheet(self, nodes):
        if nodes == None: return
        for node in nodes:
            ns = None
            if sys.version_info.minor < 9:
                ns = node.getchildren()
            else:
                ns = list(node)
            if node.tag == 'sheet' and 'title' in node.attrib.keys():
                self.sheets.append(xmlSheet.xmlSheet(name=node.attrib['title'], attrib=node.attrib,
                                                    nodes=ns))

    def createExcelTp(self):
        excel = []
        for s in self.sheets:
            excel.append((s.name, s.attrib, s.createSheet()))
        return excel

    def filename(self):
        return self.name

    def changeTableName(self, sheetname, tableold, tablenew):
        for s in self.sheets:
            if sheetname == s.name:
                pass

    def parseXML(filename):
        dom = parse(filename)
        root = dom.getroot()
        if root.tag != 'excel': raise Exception("Invalid excel format")
        #xmlExcel.traverseNode(root, testNode)
        nodes = None
        if sys.version_info.minor < 9:
            nodes = root.getchildren()
        else:
            nodes = list(root)
        return xmlExcel(name=filename, attrib=root.attrib, nodes=nodes)

    def traverseNode(node, func):
        for childnode in node:
            if func != None: func(childnode)
            xmlExcel.traverseNode(childnode, func)
