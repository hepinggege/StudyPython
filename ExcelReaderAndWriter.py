# -*- coding:utf-8 -*-
import xlrd
import xlwt
import sys
import os

dirname = "C:\\Users\\PC\Desktop\\result_excel"
datavalue = []


def open_xls(file):
    fh = xlrd.open_workbook(file)
    return fh


def getsheet(fh):
    return fh.sheets()


def getnrows(fh, sheet):
    table = fh.sheets()[sheet]
    return table.nrows


def getFilect(file, shnum):
    fh = open_xls(file)
    table = fh.sheets()[shnum]
    num = table.nrows
    for row in range(num):
        rdata = table.row_values(row)
        datavalue.append(rdata)
    return datavalue


if __name__ == '__main__':
    if len(sys.argv) >= 2:
        dirname = sys.argv[1]

    dir_list = os.listdir(dirname)
    print(os.path.split(dirname)[1])
    dstfile = dirname + ".xls"
    workbook = xlwt.Workbook(encoding='utf-8')
    lastcasename = 0
    isFirstExcel = 0
    beginRow = 0
    for root, dirs, files in os.walk(dirname):
        for file in files:
            file_extension = os.path.splitext(file)[1]
            if file_extension == '.xls':
                workfile = open_xls(dstfile)
                file = root + "/" + file
                # print(file)
                caseworkfile = open_xls(file)
                sheets = getsheet(caseworkfile)
                for i in range(len(sheets)):
                    if isFirstExcel == 0:
                        worksheet = workbook.add_sheet(str(sheets[i].name))
                    else:
                        worksheet = workbook.get_sheet(str(sheets[i].name))
                    workbook.save(dstfile)
                    workfile = open_xls(dstfile)
                    readvalue = getFilect(file, i)
                    alength = len(readvalue)
                    currownum = getnrows(workfile, i)
                    if currownum == 0:
                        beginRow = 0
                    else:
                        beginRow = 2
                    for a in range(beginRow, len(readvalue)):
                        for b in range(len(readvalue[a])):
                            c = readvalue[a][b]
                            if beginRow == 0:
                                worksheet.write(a + currownum, b, c)
                            else:
                                worksheet.write(a + currownum - 2, b, c)
                    datavalue = []
                isFirstExcel = 1
                workbook.save(dstfile)

    workbook.save(dstfile)
