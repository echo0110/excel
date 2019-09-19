

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
import pandas as pd
import numpy as np






# def read_excel():
#
#     print("Hello World")
#     # 打开文件
#     # workbook = xlrd.open_workbook(r'D:\test5.xls')
#     wb = load_workbook(r'D:\test5.xlsx')
#     # # 读取sheetname
#     # print('输出文件所有工作表名：\n', wb.sheetnames)
#     # # 获取所有sheet
#     # print
#     # workbook.sheet_names()  # [u'sheet1', u'sheet2']
#     # sheet1_name = workbook.sheet_names()[0]
#     # sheet2_name = workbook.sheet_names()[1]
#
#     # #
#     # 根据sheet索引或者名称获取sheet内容
#     # sheet2 = workbook.sheet_by_index(0)  # sheet索引从0开始
#     # sheet2 = workbook.sheet_by_name('sheet2')
#
#     # table = wb.sheets()[0]  # 通过索引顺序获取
#     # names = workbook.sheet_names()  # 返回book中所有工作表的名字
#     # filename = r'D:\test.xls'
#     # wb = load_workbook(filename)
#     # sheet1_name = workbook.sheet_names()[0]
#     # ws = workbook.active
#     # workbook.delete_cols(3)  # 删除第 3 列数据
#     # print("delete  cols")
#     # workbook.save(r'D:\test.xls')
#     # ws = wb['test']
#     # ws2 = wb[sheet_names[0]]
#     # ws.delete_cols(1, 4)  # 从第m列开始，删除n列
#     # ws = wb.get_sheet_by_name(sheetNames[0])
#     # sheet = wb.get_sheet_by_name("第一行")
#     # sheet_names = wb.sheetnames  # 返回一个列表
#     # ws2 = wb[sheet_names[0]]  # index为0为第一张表
#     # ws2.delete_rows(1)  # 删除第n行
#     # ws2.delete_rows(2)   # 删除第n行
#     # ws2.delete_rows(3)  # 删除第n行
#     # ws2.delete_cols(1, 3)  # 从第m列开始，删除n列
#     # wb.save(r'D:\test3.xlsx')
#     # print("delete  success\n")
#
#     # data = pd.read_excel(r'D:\test3.xlsx')
#     # df = pd.DataFrame(pd.read_excel(r'D:\test5.xlsx'))
#     # print (df)
#     # sheet1 = wb.sheet_by_index(0)  # sheet索引从0开始
#     sheet_names = wb.sheetnames
#     sheet1 = wb[sheet_names[0]]  # sheet索引从0开始
#     print(sheet1.col_values(9)[5])   # 获取第四行内容,值
#     # print(sheet1.row_values(3).replace("1", '10'))
#     # print (sheet.col_values(9)[5].replace("1", '10'))
#     # wb5 = load_workbook(r'D:\test5.xlsx')
#     # print (df.sheet_names()[0])
#     # df.replace('1', 10)
#     # df.replace('regex='CSC-J' , value='test'')
#     # data[data.姓名 == '张三'].语文
#     # wb5.save(r'D:\test5.xlsx')
#     print("replace  success\n")


# if __name__ == '__main__':
#     read_excel()


# -*- coding: utf-8 -*-
import xlrd
import xlwt
from datetime import date, datetime
from xlutils.copy import copy


def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook(r'D:\test5.xlsx')
    # new_workbook = copy(workbook)  # 将xlrd对象拷贝转化为xlwt对象
    # new_worksheet = new_workbook.get_sheet(0)  # 获取转化后工作簿中的第一个表格

    # wbook = copy(workbook)  # wbook即为xlwt.WorkBook对象
    # wsheet = wbook.get_sheet(0)  # 通过get_sheet()获取的sheet有write()方法
    # wr = xlwt.open_workbook(r'D:\test5.xlsx')
    # workbook = load_workbook(r'D:\test5.xlsx')

    # 根据sheet索引或者名称获取sheet内容
    sheet2 = workbook.sheet_by_index(0)  # sheet索引从0开始
    # sheet2_wr = wr.sheet_by_index(0)  # sheet索引从0开始
    # sheet2 = workbook.sheet_by_name('sheet2')

    # sheet的名称，行数，列数
    print(sheet2.name, sheet2.nrows, sheet2.ncols)

    # 获取整行和整列的值（数组）
    rows = sheet2.row_values(3)  # 获取第四行内容
    cols = sheet2.col_values(2)  # 获取第三列内容
    cols = sheet2.col_values(8)  # 获取第8列内容
    print(rows)
    print(cols)


    # 获取单元格内容
    cols = sheet2.col_values(8)  # 获取第8列内容

    print(sheet2.cell(0, 0))
    print(sheet2.col_values(7)[4])
    print(sheet2.col_values(7)[4].replace('CSC-J', ''))
    # write = sheet2.col_values(7)[4].replace('CSC-J', 'niu')
    # new_worksheet.write(7, 4, write)
    # new_workbook.save(r'D:\test5.xlsx')  # 保存工作簿
    print(sheet2.col_values(7)[4])
    print("xls格式表格【追加】写入数据成功！")

    # workbook.save(r'D:\test5.xlsx')

    # workbook.save(r'D:\test5.xlsx')
    # workbook.save(r'D:\test5.xlsx')
    # print(sheet2.cell(i, j))
    # 用for循环的方式连续向单元格中写入内容
    # for i in range(4, 20):
    #     for j in range(7, 9):
            # print(sheet2.cell(i, j))   # replace('CSC-J', 'niu')
            # sheet2.cell(i, j)
            # sheet2.write(i, j, 'niu')
            # print(sheet2.cell(i, j))
     # 保存excel文件，写入完成后，一定要保存


    # print(sheet2.cell(1, 0).value.encode('utf-8'))
    # print(sheet2.cell_value(1, 0)# workbook.save(r'D:\test5.xlsx') .encode('utf-8'))
    # print(sheet2.row(1)[0].value.encode('utf-8'))

    # 获取单元格内容的数据类型
    # print(sheet2.cell(1, 0).ctype)



if __name__ == '__main__':
    read_excel()





# -*- coding:utf-8 -*-
# from xlrd import open_workbook
# from xlutils.copy import copy
# import re
#
#
#
# def getrule(rfile='D:/test1.txt'):
#     try:
#         rdict = {}
#         with open(rfile, 'r') as f:
#             for line in f:
#                 rline = line.split('->')
#                 rdict[rline[0].strip()] = rline[1].strip()
#                 print rdict[rline[0].strip()]
#         return rdict
#     except Exception, e:
#         print e
#
#
# if __name__ == '__main__':
#     excelfile = 'D:/test1.xls'
#     rdict = getrule()
#     print rdict
#     rb = open_workbook(excelfile)
#     rs = rb.sheet_by_index(0)
#     wb = copy(rb)
#     ws = wb.get_sheet(0)
#     nrows = rs.nrows
#     ncols = rs.ncols
#     table = rb.sheets()[0]
#     prices = table.row_values(0)[0]
#     print prices
#     c = prices.replace('1', '88')
#     print c
#
#     strinfo = re.compile('1')
#     b = strinfo.sub('python', prices)
#     print b
#     for i in range(0,nrows):
#         for j in range(0,ncols):
#             cvalue = rs.cell(i, j).value
#             # if type(cvalue).__name__ == 'float':
#             #     cvalue = str(int(cvalue))
#             # if rdict.has_key(cvalue):
#             #     print '%s is replaced by %s' % (cvalue, rdict[cvalue])
#             #     ws.write(i, j, rdict[cvalue])
#             for repStr in rdict.keys():
#                 cvalue = cvalue.replace(repStr, rdict[repStr])
#                 ws.write(i, j, cvalue)
#     wb.save(excelfile)
#     print 'OK!'



