#!/usr/bin/python
# -*- coding: UTF-8 -*-


import xlsxwriter


def Write (write_data, write_file):
    workbook = xlsxwriter.Workbook('D:/study/Python/shell/parse/excel/'+write_file+'.xlsx')
    worksheet = workbook.add_worksheet()

    # 设定格式，等号左边格式名称自定义，字典中格式为指定选项
    # bold：加粗，num_format:数字格式
    bold_format = workbook.add_format({'bold': True})
    # money_format = workbook.add_format({'num_format': '$#,##0'})
    # date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})

    # 将二行二列设置宽度为15(从0开始)
    worksheet.set_column(1, 1, 15)

    # 用符号标记位置，例如：A列1行
    worksheet.write('A1', '论文情况', bold_format)
    worksheet.write('B1', '平台', bold_format)
    worksheet.write('C1', '批次', bold_format)
    worksheet.write('D1', '招生顾问', bold_format)
    worksheet.write('E1', '来源', bold_format)
    worksheet.write('F1', '学生姓名', bold_format)
    worksheet.write('G1', '性别', bold_format)
    worksheet.write('H1', '民族', bold_format)
    worksheet.write('I1', '生日', bold_format)
    worksheet.write('J1', '身份证号', bold_format)
    worksheet.write('K1', '报名编号', bold_format)
    worksheet.write('L1', '学号', bold_format)
    worksheet.write('M1', '奥鹏卡号', bold_format)
    worksheet.write('N1', '身份证地址', bold_format)
    worksheet.write('O1', '学校', bold_format)
    worksheet.write('P1', '层次', bold_format)
    worksheet.write('Q1', '专业', bold_format)
    worksheet.write('R1', '平台手机', bold_format)
    worksheet.write('S1', '学位', bold_format)
    worksheet.write('T1', '学位班', bold_format)
    worksheet.write('U1', '用户名', bold_format)
    worksheet.write('V1', '密码', bold_format)
    worksheet.write('W1', '毕业情况', bold_format)
    worksheet.write('X1', '第一次交费', bold_format)
    worksheet.write('Y1', '第一次学费收款人', bold_format)
    worksheet.write('Z1', '第二次交费', bold_format)
    worksheet.write('AA1', '第二次学费收款人', bold_format)
    worksheet.write('AB1', '统考费', bold_format)
    worksheet.write('AC1', '统考费收款人', bold_format)
    worksheet.write('AD1', '学位班学费', bold_format)
    worksheet.write('AE1', '学位收款人', bold_format)
    worksheet.write('AF1', '学校报名费', bold_format)
    worksheet.write('AG1', '学校报名费交款人', bold_format)
    worksheet.write('AH1', '学校学费', bold_format)
    worksheet.write('AI1', '第一次缴费', bold_format)
    worksheet.write('AJ1', '第一次缴费人', bold_format)
    worksheet.write('AK1', '第二次缴费', bold_format)
    worksheet.write('AL1', '第二次缴费人', bold_format)
    worksheet.write('AM1', '新华社信息采集费（元）', bold_format)
    worksheet.write('AN1', '备注', bold_format)

    row = 1
    col = 0
    for item in (write_data):
        # 使用write_string方法，指定数据格式写入数据
        worksheet.write_string(row, col, str(item['论文情况']))
        worksheet.write_string(row, col + 1, str(item['平台']))
        worksheet.write_string(row, col + 2, str(item['批次']))
        worksheet.write_string(row, col + 3, str(item['招生顾问']))
        worksheet.write_string(row, col + 4, str(item['来源']))
        worksheet.write_string(row, col + 5, str(item['学生姓名']))
        worksheet.write_string(row, col + 6, str(item['性别']))
        worksheet.write_string(row, col + 7, str(item['民族']))
        worksheet.write_string(row, col + 8, str(item['生日']))
        worksheet.write_string(row, col + 9, str(item['身份证号']))
        worksheet.write_string(row, col + 10, str(item['报名编号']))
        worksheet.write_string(row, col + 11, str(item['学号']))
        worksheet.write_string(row, col + 12, str(item['奥鹏卡号']))
        worksheet.write_string(row, col + 13, str(item['身份证地址']))
        worksheet.write_string(row, col + 14, str(item['学校']))
        worksheet.write_string(row, col + 15, str(item['层次']))
        worksheet.write_string(row, col + 16, str(item['专业']))
        worksheet.write_string(row, col + 17, str(item['平台手机']))
        worksheet.write_string(row, col + 18, str(item['学位']))
        worksheet.write_string(row, col + 19, str(item['学位班']))
        worksheet.write_string(row, col + 20, str(item['用户名']))
        worksheet.write_string(row, col + 21, str(item['密码']))
        worksheet.write_string(row, col + 22, str(item['毕业情况']))
        worksheet.write_string(row, col + 23, str(item['第一次交费']))
        worksheet.write_string(row, col + 24, str(item['第一次学费收款人']))
        worksheet.write_string(row, col + 25, str(item['第二次交费']))
        worksheet.write_string(row, col + 26, str(item['第二次学费收款人']))
        worksheet.write_string(row, col + 27, str(item['统考费']))
        worksheet.write_string(row, col + 28, str(item['统考费收款人']))
        worksheet.write_string(row, col + 29, str(item['学位班学费']))
        worksheet.write_string(row, col + 30, str(item['学位收款人']))
        worksheet.write_string(row, col + 31, str(item['学校报名费']))
        worksheet.write_string(row, col + 32, str(item['学校报名费交款人']))
        worksheet.write_string(row, col + 33, str(item['学校学费']))
        worksheet.write_string(row, col + 34, str(item['第一次缴费']))
        worksheet.write_string(row, col + 35, str(item['第一次缴费人']))
        worksheet.write_string(row, col + 36, str(item['第二次缴费']))
        worksheet.write_string(row, col + 37, str(item['第二次缴费人']))
        worksheet.write_string(row, col + 38, str(item['新华社信息采集费（元）']))
        worksheet.write_string(row, col + 39, str(item['备注']))

        row += 1
    workbook.close()