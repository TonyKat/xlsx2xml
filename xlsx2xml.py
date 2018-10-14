#!/usr/bin/python
# -*- coding: utf-8 -*-
import datetime
import time
import xlrd
import xml.dom.minidom
import os



def translate_excel_to_xml(xlsx_path, name):
    data = xlrd.open_workbook(xlsx_path)
    print("translate " + xlsx_path + " ... ")
    table = data.sheets()[0]

    nrows = table.nrows
    ncols = table.ncols

    # create .xml
    doc = xml.dom.minidom.Document()
    root = doc.createElement('root')
    doc.appendChild(root)

    for nrow in range(0, nrows):
        if nrow == 0:
            continue
        item = doc.createElement('item')
        for ncol in range(0, ncols):
            key = "%s" % table.cell(nrow, ncol).value
            value = table.cell(nrow, ncol).value

            t = table.cell(nrow, ncol).ctype

            if t == 0:
                key = ""
                value = ""
            if t == 2:
                value = str(int(value))
            if t == 3:
                key = datetime.datetime(*xlrd.xldate_as_tuple(float(key), data.datemode))
                value = key.strftime('%d.%m.%Y')
                key = value

            if ncol == 0:
                k = doc.createElement(str(table.cell(0, 0).value))
                v = doc.createTextNode(value)
                k.appendChild(v)
                item.appendChild(k)
            elif ncol == 1:
                k = doc.createElement(str(table.cell(0, 1).value))
                v = doc.createTextNode(value)
                k.appendChild(v)
                item.appendChild(k)
            elif ncol == 2:
                k = doc.createElement(str(table.cell(0, 2).value))
                v = doc.createTextNode(value)
                k.appendChild(v)
                item.appendChild(k)
            elif ncol == 3:
                k = doc.createElement(str(table.cell(0, 3).value))
                v = doc.createTextNode(value)
                k.appendChild(v)
                item.appendChild(k)
            else:
                k = doc.createElement(str(table.cell(0, ncol).value))
                v = doc.createTextNode(value)
                k.appendChild(v)
                item.appendChild(k)
        root.appendChild(item)
    xml_name = name.strip().split('.')[0] + '.xml'
    xml_path = os.path.join(xml_dir, xml_name)

    f = open(xml_path, 'w')
    f.write(doc.toprettyxml())
    f.close()


if __name__ == "__main__":
    time_begin = time.time()
    xlsx_dir = '\\xlsx_files\\'
    xml_dir = '\\xlsx_to_xml\\'
    for name in os.listdir(xlsx_dir):
        if name.endswith('.xlsx'):
            xlsx_path = os.path.join(xlsx_dir, name)
            translate_excel_to_xml(xlsx_path, name)

    print('\nВремя исполнения программы: ', time.time() - time_begin)
