# -*- coding: utf-8 -*-
"""
@author zhuyan
"""
import xlrd
import xlwt
import os


def file_name(file_dir):
    file_list = []
    for root, dirs, files in os.walk(file_dir):
        for file in files:
            if (os.path.splitext(file)[1] == '.xlsx') | (os.path.splitext(file)[1] == '.xls'):
                file_list.append(os.path.join(root, file))
    return file_list


if __name__ == '__main__':
    path = "C:\\Users\9\Desktop\街道信息"  # 原Excel所在文件夹
    new_path = "C:\\Users\9\Desktop\已完成"  # 新Excel保存路径，如原文件夹下有子文件夹，需新建对应同名子文件夹
    k = file_name(path)  # 获取原文件夹下所有的xls文件，并将其全路径保存在一个数组中。
    print(k)
    count = 0
    for current_file_name in k:
        count = count + 1
        write_book = xlwt.Workbook()  # 打开一个新Excel
        print("当前文件:" + str(count) + "/" + str(k.__len__()))  # 输出当前文件数/总文件数
        print(current_file_name + " 正在转存。。。")
        sheet = write_book.add_sheet('Sheet')  # 新文件中添加一个Sheet页，以及名称
        data = xlrd.open_workbook(current_file_name)
        new_name = current_file_name.replace(path, new_path)
        table = [u"Sheet"]  # 打开原文件中的Sheet页
        for x in table:
            if current_file_name.__contains__("志愿者"):
                x = "Sheet1"
            if data.sheet_names().__contains__(x):
                table = data.sheet_by_name(x)
                name = -1
                meaning = -1
                factor = -1
                default_value = -1
                factors = -1
                # 最大行
                rows = table.nrows
                # 最大列
                cols = table.ncols
                for i in range(0, rows):
                    u = 0
                    current_row = "当前行：" + str(i + 1) + "/" + str(rows)
                    if ((((i + 1) / rows) * 100) % 10 == 0) | (i + 1 == rows):
                        print(current_row)
                    for j in range(0, cols):
                        filled = str(table.row_values(i)[j])
                        if i <= 2 & j == 0:
                            sheet.write(i, j, "1")  # 前三行第一列写入“1”
                        else:
                            sheet.write(i, j, filled.strip())
                # write_book.save(new_name)
                print(new_name+" 已完成")
                print(" ")
        # os.remove(current_file_name)  # 将完成的文件删除
