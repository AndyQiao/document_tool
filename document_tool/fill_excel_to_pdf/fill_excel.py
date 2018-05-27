#! /usr/bin/env python   
# -*- coding: utf-8 -*-

import codecs
import os
import xlrd
import xlwt


#读取excel列
def ReadExcelCol(file_name, col, row_begin):
    workbook = xlrd.open_workbook(file_name)
    sheet = workbook.sheet_by_index(0)

    nrows = sheet.nrows
    #ncols = sheet1.ncols

    col_data = sheet.col_values(col)
    return col_data[row_begin:]





#写入
def Write2Excel(names, levels):
    workbook = xlwt.Workbook(encoding = 'utf-8')
    worksheet = workbook.add_sheet('My Worksheet')
    for row in range(0, len(levels)):
        worksheet.write(row+1, 0, names[row])
        for col in range(0,len(levels[row])):
            worksheet.write(row+1, col+1, levels[row][col])
    workbook.save('out.xls')


# 读&写
def ReadAndWrite(names, from_file, to_file, row, col):
    workbook = xlrd.open_workbook(from_file)
    sheet = workbook.sheet_by_index(0)

    for name in names:
        sheet1 = workbook.add_sheet('1')
        sheet1 = sheet
        sheet1.write(row, col, names[row])
        workbook.save(name+to_file)




read_file_path = 'shop_president.xlsx'
to_file_name1 = '-超市店长胜任力测评个人反馈报告.xlsx'
to_file_name2 = '-超市店长胜任力测评综合评估报告.xls'
to_file_name3 = '-采购总监资深采购胜任力测评个人反馈报告.xls'
to_file_name4 = '-采购总监资深采购胜任力测评综合评估报告.xls'

from_file_name1 = 'shop_president_source1.xlsx'
from_file_name2 = 'shop_president_source2.xls'
from_file_name3 = 'shop_buyyer_source3.xls'
from_file_name4 = 'shop_buyyer_source4.xls'

names = ReadExcelCol(read_file_path, 0, 1)

ReadAndWrite(names, from_file_name1, to_file_name1, 41, 17)

