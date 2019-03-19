# -*- coding: utf-8 -*-
"""
@author zhuyan
"""
import xlrd
import xlwt
import os
import csv


def file_name(file_dir):
    file_list = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if (os.path.splitext(file)[1] == '.xlsx') | (os.path.splitext(file)[1] == '.xls'):
                file_list.append(os.path.join(root, file))
    return file_list


if __name__ == '__main__':
    path = "C:\\Users\9\Desktop\户籍人口"  # 原Excel所在文件夹
    k = file_name(path)  # 获取原文件夹下所有的xls文件，并将其全路径保存在一个数组中。
    print(k)
    count = 0
    huji = open('Sheet.csv', 'a', encoding='utf-8', newline='')
    csv_write = csv.writer(huji, dialect='excel')
    for current_file_name in k:
        count = count + 1
        print("当前文件:" + str(count) + "/" + str(k.__len__()))  # 输出当前文件数/总文件数
        print(current_file_name)
        data = xlrd.open_workbook(current_file_name)
        table = data.sheet_names()  # 打开原文件中的Sheet页
        if table.__contains__('Sheet'):
            table.remove('Sheet')
        if table.__contains__('Sheet1'):
            table.remove('Sheet1')
        if table.__contains__('Sheet4'):
            table.remove('Sheet4')
        if table.__contains__('bean'):
            table.remove('bean')
        if table.__contains__('省市'):
            table.remove('省市')
        if table.__contains__('市县'):
            table.remove('市县')
        if table.__contains__('国籍选项字段'):
            table.remove('国籍选项字段')
        row = []
        row.append(table)
        row.append(current_file_name.replace("C:\\Users\9\Desktop\\",''))
        if row[0].__len__() != 0:
            print("存在问题：")
            print(row)
            csv_write.writerow(row)
