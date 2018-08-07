# -*- coding: utf-8 -*-
#python
#用于实现将excel文件转为markdown文件

import codecs
import xlrd



def excelToMd(origin_url,save_url,save_as):
    # 打开文件
    workbook = xlrd.open_workbook(origin_url)
    # 获取所有sheet
    # print(workbook.sheet_names())  # [u'sheet1', u'sheet2']
    # sheet2_name = workbook.sheet_names()

    # 根据sheet索引或者名称获取sheet内容
    sheet2 = workbook.sheet_by_index(0)  # sheet索引从0开始
    # sheet2 = workbook.sheet_by_name('sheet2')
    md_url_file = '%s%s.md' % (save_url, save_as)  # 文件名连接第一列和第二列
    file = codecs.open(md_url_file, 'w', "utf-8")  # 写入文件名
    # sheet的名称，行数，列数
    # print(sheet2.name, sheet2.nrows, sheet2.ncols)
    rowfilter = []
    str = ''
    for i in range(0,sheet2.nrows):
        if i > 0:
            row = list(map(lambda x : '^' if x == '' else x,sheet2.row_values(i)))
            res = list(map(lambda x: x.replace('\n', '<br/>'), row))
            rowfilter.append('|'.join(res))
        else:
            row = sheet2.row_values(i)
            res = list(map(lambda x: x.replace('\n', '<br/>'), row))
            rowfilter.append('|'.join(res))
            split_str = '|---' * sheet2.ncols
            rowfilter.append(split_str)
    for i in rowfilter:
        if rowfilter.index(i) == 1:
            str = i + '|\n'
        else:
            str = '|' + i + '|\n'
        # print(str)
        file.write(str)
    file.close()



if __name__ == '__main__':
    # 指定excel的地址
    excel_url = r'F:\\markdown\\control_plan.xlsx'

    # 指定mark文档输入地址

    md_url = 'F:\\markdown\\'
    save_as = 'control_plan'
    excelToMd(excel_url,md_url,save_as)