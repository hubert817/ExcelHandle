#!/usr/bin/python
# -*- coding: UTF-8 -*-

import sys
sys.path.append('D:\study\Python\shell\parse')
from WriteExcel import Write
from ReadExcel import Read
from YouDao import Translation

"""
Write 引用xlsxwriter库,写入Excel表里数据
Read 引用pandas库,读取Excel表里数据
Translation 引用requests库,调用有道翻译接口, 翻译专业名称
"""


def Classification():
    f = open('course.txt', encoding='utf-8')
    line = f.readline()
    while line:
        file_data = Read('1.xlsx', line.strip())
        file_name = Translation(line.strip())
        if len(file_data) and file_name is not None:
            Write(file_data, file_name)
        line = f.readline()
    f.close()


if __name__ == '__main__':
    Classification()