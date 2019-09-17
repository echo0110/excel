

# -*- coding: utf-8 -*-
import xlrd
import xlwt
from datetime import date, datetime

# from openpyxl import load_workbook

from openpyxl import Workbook

from openpyxl import *

from openpyxl import load_workbook
from openpyxl.styles import colors, Font, Fill, NamedStyle
from openpyxl.styles import PatternFill, Border, Side, Alignment







def read_excel():

    print("Hello World")
    # 打开文件
    # workbook = xlrd.open_workbook(r'D:\test3.xls')
    wb = load_workbook(r'D:\test3.xlsx')
    # 读取sheetname
    print('输出文件所有工作表名：\n', wb.sheetnames)
    # # 获取所有sheet
    # print
    # workbook.sheet_names()  # [u'sheet1', u'sheet2']
    # sheet1_name = workbook.sheet_names()[0]
    # sheet2_name = workbook.sheet_names()[1]

    # #
    # 根据sheet索引或者名称获取sheet内容
    # sheet2 = workbook.sheet_by_index(0)  # sheet索引从0开始
    # sheet2 = workbook.sheet_by_name('sheet2')

    # table = wb.sheets()[0]  # 通过索引顺序获取
    # names = workbook.sheet_names()  # 返回book中所有工作表的名字
    # filename = r'D:\test.xls'
    # wb = load_workbook(filename)
    # sheet1_name = workbook.sheet_names()[0]
    # ws = workbook.active
    # workbook.delete_cols(3)  # 删除第 3 列数据
    # print("delete  cols")
    # workbook.save(r'D:\test.xls')
    # ws = wb['test']
    # ws2 = wb[sheet_names[0]]
    # ws.delete_cols(1, 4)  # 从第m列开始，删除n列
    # ws = wb.get_sheet_by_name(sheetNames[0])
    # sheet = wb.get_sheet_by_name("第一行")
    sheet_names = wb.sheetnames  # 返回一个列表
    ws2 = wb[sheet_names[0]]  # index为0为第一张表
    ws2.delete_rows(1)  # 删除第n行
    ws2.delete_rows(2)   # 删除第n行
    ws2.delete_rows(3)  # 删除第n行
    ws2.delete_cols(1, 3)  # 从第m列开始，删除n列
    wb.save(r'D:\test3.xlsx')
    print("delete  cols3")


if __name__ == '__main__':
    read_excel()
