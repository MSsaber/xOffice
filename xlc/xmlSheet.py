# !/usr/python
# -*- coding:utf-8 -*-

import xml

class xmlSheet:
    def __init__(self, **argparam):
        self.tables = []
        self.tag = 'sheet'
        if 'name' in argparam.keys():
            self.name = argparam['name']
        if 'attrib' in argparam.keys():
            self.attrib = argparam['attrib']
        if 'nodes' in argparam.keys():
            self.nodes = argparam['nodes']
            self.__genTable(self.nodes)

    def __genTable(self, nodes):
        if nodes == None: return
        for node in nodes:
            if node.tag == 'table' and 'name' in node.attrib.keys():
                self.tables.append(xmlTable(name=node.attrib['name'], attrib=node.attrib,
                                            nodes=node.getchildren()))

    def createSheet(self):
        sheet = []
        for t in self.tables:
            sheet.append((t.name, t.attrib, t.createTable()))
        return sheet

class xmlTable:
    def __init__(self, **argparam):
        self.header = []
        self.tag = 'table'
        if 'name' in argparam.keys():
            self.name = argparam['name']
        if 'attrib' in argparam.keys():
            self.attrib = argparam['attrib']
        if 'nodes' in argparam.keys():
            self.nodes = argparam['nodes']
            self.__genHeader(self.nodes)

    def __genHeader(self, nodes):
        if nodes == None: return
        for node in nodes:
            if node.tag == "header" and "name" in node.attrib.keys():
                self.header.append(xmlHeader(name=node.attrib['name'], attrib=node.attrib,
                                            text=node.text))

    def createTable(self):
        table = []
        for h in self.header:
            table.append((h.name, h.attrib, h.eles))
        return table

class xmlHeader:
    def __init__(self, **argparam):
        self.tag = "header"
        self.eles = []
        if 'name' in argparam.keys():
            self.name = argparam['name']
        if 'attrib' in argparam.keys():
            self.attrib = argparam['attrib']
        if 'text' in argparam.keys():
            self.text = argparam['text']
            self.__parseText(self.text)

    def __parseText(self, text):
        text=text.replace('\n','')
        text=text.replace(' ','')
        self.eles = text.split(',')