# -*- coding:utf-8 -*-

import sys
import xlwt
import xlrd
import openpyxl
from xlutils.copy import copy

myStyle = xlwt.easyxf('font: name Times New Roman, color-index red, bold on', num_format_str='#,##0.00')


def is_num(s):
    try:
        num = float(s)
        return True, num
    except ValueError:
        return False, 0


base_xls = sys.argv[1]
sec_xls = sys.argv[2]

wb = xlrd.open_workbook(base_xls)
vb = xlrd.open_workbook(sec_xls)
vb_new = copy(vb)
# 获取workbook中所有的表格
sheet1 = wb.sheets()
sheet2 = vb.sheets()

for i in range(len(sheet1)):
    sheet_w = sheet1[i]
    sheet_v = sheet2[i]
    sheet_new = vb_new.get_sheet(i)
    for r in range(2, sheet_w.nrows):
        for c in range(1, sheet_w.ncols):
            value1 = sheet_w.cell(r, c).value
            value2 = sheet_v.cell(r, c).value
            print(value1)
            print(value2)
            isTrue1, num1 = is_num(value1)
            isTrue2, num2 = is_num(value2)
            if isTrue1 and isTrue2:
                a = num2 - num1
                if num1 != 0 and a / num1 > 0:
                    sheet_new.write(r, c, value2, myStyle)

vb_new.save(sec_xls)







