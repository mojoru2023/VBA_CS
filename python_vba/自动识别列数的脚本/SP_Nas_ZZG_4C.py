#
#
#
# # IW ---SN
# # 3----254
# import string
# import time
# # -*- coding: utf-8 -*-
#
# # 读取页面文本
# # 按照标题，保存整个文本
#
#
#
#
#
# import csv
# import datetime
#
# import os
# import re
# import time
# import sys
# type = sys.getfilesystemencoding()
# import pymysql
# import xlrd
# import requests
# from requests.exceptions import RequestException
# from lxml import etree
#
# # 1. 先读取excel文件中列数，然后+1，列数*2-1
# # 2. 相对为位置
# # 3.然后直接使用
#
#
# def removeDot(item):
#     f_l = []
#     for it in item:
#         f_str = "".join(it.split(","))
#         f_l.append(f_str)
#
#     return f_l
#
#
# def remove_block(items):
#     new_items = []
#     for it in items:
#         f = "".join(it.split())
#         new_items.append(f)
#     return new_items
#
#
#
# # print(sh.nrows)#有效数据行数
# # print(sh.ncols)#有效数据列数
# # print(sh.cell(0,0).value)#输出第一行第一列的值
# # print(sh.row_values(0))#输出第一行的所有值
# # #将数据和标题组合成字典
# # print(dict(zip(sh.row_values(0),sh.row_values(1))))
# # #遍历excel，打印所有数据
# # for i in range(sh.nrows):
# #     print(sh.row_values(i))
#
#
#
#
#
# ff_List=[]
# alphabet =string.ascii_uppercase
# head_str='Public num As Integer\nSub s()\nDim i As Integer\nFor i = 2 To 9999\nIf Range("C" & i) = "" Then\nnum = i\nExit For\nEnd If'
# tail_str='Next\nRange("{0}1:{1}" & num).Select\nActiveSheet.Shapes.AddChart2(227, xlLine).Select\nActiveChart.SetSourceData Source:=Range("Sheet1!${0}$1:${1}$" & num)\nEnd Sub'
# str_1 = ' Range("{0}1").Select\nActiveCell.FormulaR1C1 = "=RC[-{2}]"\nRange("{0}" & i).Select \n ActiveCell.FormulaR1C1 = "=RC[-{2}]/R2C{1}-1"'
#
#
#
#
# for i1 in alphabet:
#     ff_List.append(i1)
#     for i2 in alphabet:
#         f_al =i1+i2
#         ff_List.append(f_al)
#         for i3 in alphabet:
#             f_al1= i1+i2+i3
#             ff_List.append(f_al1)
#
#
#
#
# def into_file(f_name,item_name):
#     try:
#         with open('{0}.bas'.format(f_name), 'a') as file_handle:
#
#             file_handle.write(item_name + "")  # 写入
#             file_handle.write('\n')  # 有时放在循环里面需要自动转行，不然会覆盖上一条数据
#             print("{0} 整理完毕".format(f_name))
#     except:
#         pass
#
#
# if __name__=="__main__":
#
#     # 先按照字符串长度排序(小->大)，再按照字母排序
#     f_alist=sorted(ff_List, key=lambda i: len(i), reverse=False)
#     # 1.读取excel文件
#
#     # 打开excel
#     wb = xlrd.open_workbook('SP_Nas_ZZG.xlsx')
#     # 按工作簿定位工作表
#     sh = wb.sheet_by_name('Sheet1')
#     f_ncols = sh.ncols
#     # start_index1 = f_alist.index("W")  # 修改处1
#     # end_index1 = f_alist.index("AN")  # 修改处2
#     start_index = f_ncols + 1
#     end_index = f_ncols * 2-3
#     len_str = end_index - start_index + 3
#
#
#
#     into_file("SP_Nas_ZZG",head_str)
#     for a_i,num in zip(f_alist[start_index:end_index+1],range(3,end_index-start_index+3+1)):
#
#         f_code=str_1.format(a_i,num,len_str)
#         into_file("SP_Nas_ZZG",f_code)
#
#     into_file("SP_Nas_ZZG", tail_str.format(f_alist[start_index],f_alist[end_index]))
#
#
#




# 放弃vba,直接上python

# -*- coding: utf-8 -*-

# 读取页面文本
# 按照标题，保存整个文本


import csv
import datetime
import numpy as np


import os
import re
import time
import sys

type = sys.getfilesystemencoding()
import pymysql
import xlrd






def writerDt_csv(headers, rowsdata):
    # rowsdata列表中的数据元组,也可以是字典数据
    with open('SP_Nas_ZZG.csv', 'w', newline='') as f:
        f_csv = csv.writer(f)
        f_csv.writerow(headers)
        f_csv.writerows(rowsdata)


def read_xlrd(excelFile):
    data = xlrd.open_workbook(excelFile)
    table = data.sheet_by_index(0)
    dataFile = []
    for rowNum in range(table.nrows):
        dataFile.append(table.row_values(rowNum))

    # # if 去掉表头
    # if rowNum > 0:

    return dataFile


# xlsx---list_url----单页url
def get_allURL():

    lpath = os.getcwd()
    for item in ["SP_Nas_ZZG.xlsx"]:

        excelFile = '{0}\\{1}'.format(lpath,item)
        full_items = read_xlrd(excelFile=excelFile)


        #2---362
        for num in range(2,len(full_items[0])-1):
            _list= []

            for item in full_items:

                _list.append((item[num]))
            operate_list.append(_list)







if __name__ == '__main__':
    operate_list = []
    f_lst = []
    last_max_min_ = []
    headers=["js_industryDT","last_","max_","_min","_stdev"]

    get_allURL()
    for a in operate_list:

        f_1_0 = a[1:][0]

        for i in a[1:]:

            if f_1_0 !=0:
                f_r = round((i - f_1_0) / f_1_0, 4)
                last_max_min_.append(f_r)
            else:
                last_max_min_.append(0)
        stdev = last_max_min_[-100:]
        arr_std = np.std(stdev, ddof=1)

        _last = last_max_min_[-1]
        _max = max(last_max_min_)
        _min = min(last_max_min_)
        f_lst.append((a[0], _last, _max, _min, round(arr_std, 6)))
    writerDt_csv(headers,f_lst)
    print(f_lst)



