#!/usr/bin/env python
# -*- coding: UTF-8 -*-
"""
*************************************************
@Project -> File   ：dropout-master -> excel
@IDE    ：PyCharm
@Author ：Zhuge hammer
@Date   ：2021/12/10 15:59
@Desc   ：
*************************************************
"""
import xlrd
import numpy as np
import pandas as pd


# 提取互评表某一列
def extract(path, col_num):
    data = xlrd.open_workbook(path, encoding_override='utf-8')
    table = data.sheets()[0]  # 选定表
    nrows = table.nrows  # 获取行号

    # 总成绩列表
    sum_list = []
    for i in range(3, nrows):  # 第0行为表头
        alldata = table.row_values(i)  # 循环输出excel表中每一行，即所有数据
        result = alldata[col_num]  # 取出表中第x列数据
        sum_list.append(result)
    return sum_list


def handle_list(excel_num, col_num):
    # 互评分数列表
    hlist = []
    # 自评分数列表
    zlist = []
    for i in range(1, excel_num+1):
        excel_path = 'dataset\\' + str(i) + '.xls'  # excel文件所在路径
        col_list = extract(excel_path, col_num)
        # 自评分数取50%
        zlist.append(col_list[i-1]/2)
        # 将每个人的自评分数置0
        col_list[i-1] = 0
        # 写入互评分数列表
        hlist.append(col_list)

    # 求和除以互评表数量
    result = lambda lst: [sum(i) / (2*(excel_num-1))for i in (list(zip(*lst)))]
    # 最后的互评均分
    input_rlist = result(hlist)
    # 最后的个人总分
    last_list = list(map(lambda x: x[0] + x[1], zip(input_rlist, zlist)))
    return last_list


def get_input_list(excel_nums):
    datas = []
    # 取学生姓名
    std_name = extract('dataset\\1.xls', 1)
    names = std_name[0:excel_nums]
    # 循环从第2列取到第8列
    for i in range(2, 8):
        datas.append(handle_list(excel_nums, i))
    datas.append(names)
    return datas


if __name__ == "__main__":
    # 只需修改表格数量，这里测试取3
    input_data = get_input_list(39)
    column_list = ['姓名', '有信仰', '讲政治', '重品行', '争先锋', '守纪律', '总分']
    df_data = {'有信仰': input_data[0], '讲政治': input_data[1], '重品行': input_data[2], '争先锋': input_data[3], '守纪律': input_data[4], '总分': input_data[5], '姓名': input_data[6]}
    df = pd.DataFrame(df_data)
    df.to_excel("评议表总分统计.xlsx", index=False, columns=column_list)

