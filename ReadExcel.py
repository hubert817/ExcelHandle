#!/usr/bin/python
# -*- coding: UTF-8 -*-

import pandas as pd


def Read(read_file, read_name):
    df = pd.read_excel(read_file)
    excel_data = []
    #获取excel表里数据,dict类型输出
    for i in df.index.values:
        row_data = df.ix[i,['论文情况','平台','批次','招生顾问','来源','学生姓名','性别','民族','生日','身份证号','报名编号','学号','奥鹏卡号','身份证地址','学校','层次','专业','平台手机','学位','学位班','用户名','密码','毕业情况','第一次交费','第一次学费收款人','第二次交费','第二次学费收款人','统考费','统考费收款人','学位班学费','学位收款人','学校报名费','学校报名费交款人','学校学费','第一次缴费','第一次缴费人','第二次缴费','第二次缴费人','新华社信息采集费（元）','备注']].to_dict()
        excel_data.append(row_data)
    filter_data = []
    #过滤专业人数,待写入新的excel的dict数据
    for j in excel_data:
        if j.get('专业', '') == read_name:
            filter_data.append(j)

    return filter_data

