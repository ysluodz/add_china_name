# -*- coding:utf-8 -*-
# __author__ = 'Administrator'
import os
import re

import openpyxl
import pandas as pd
import xlrd as xlrd
import xlwt
from openpyxl import load_workbook
from pandas import DataFrame
from xlutils.copy import copy
import sys

reload(sys)
sys.setdefaultencoding('utf8')

def add_china_name(result):
    workbook = xlrd.open_workbook('pearn_crr_data.xlsx')
    booksheet = workbook.sheet_by_index(0)
    wbk = xlwt.Workbook('pearn_crr_data_1.xlsx')
    #wb = copy(workbook)

    cell_11 = booksheet.cell_value(1, 1)
    cell_12 = booksheet.cell_value(1, 2)
    name1 = result[cell_11]
    name2 = result[cell_12]

    wbk.get_sheet(0).write(1, 1, cell_11 + '('+name1 + ')')
    wbk.get_sheet(0).write(1, 2, cell_12 + '('+name2 + ')')
    wbk.save('pearn_crr_data1.xlsx')

    print cell_11, cell_12, name1.encode('utf-8'), name2.encode('utf-8')


def add_china_name1(result):
    workbook = xlrd.open_workbook('pearn_crr_data.xlsx')
    booksheet = workbook.sheet_by_index(0)

    for i in xrange(1,381093):
        cell_11 = booksheet.cell_value(1, 1)
        cell_12 = booksheet.cell_value(1, 2)

        name1 = result[cell_11]
        name2 = result[cell_12]

        booksheet.write(1, 1, cell_11 + '('+name1 + ')')
        booksheet.write(1, 2, cell_12 + '('+name2 + ')')

        workbook.save('pearn_crr_data.xlsx')
        print cell_11, cell_12, name1.encode('utf-8'), name2.encode('utf-8')


def id_name():
    file_path = "id_name.csv"
    workbook = xlrd.open_workbook('namecsv/id_name.xls')
    booksheet = workbook.sheet_by_index(0)
    result = {}
    for i in xrange(1, 1242):
        id = booksheet.cell_value(i, 3)
        name = booksheet.cell_value(i, 5)
        result[id] = name

    # df = DataFrame(result)
    # df.to_csv('namecsv/id_name.csv',index=False,encoding="utf-8")
    return result

def get_target_files(root_dir, data_file, pname):
    """
    给定目录，模式符合pname的文件
    :param root_dir:
    :param data_file:
    :param pname:
    :return:
    """
    dir_file_list = os.listdir(root_dir)
    for item in dir_file_list:
        path = os.path.join(root_dir, item)
        if os.path.isfile(path):
            if re.search(pname, item):
                data_file.append(path)

        elif os.path.isdir(path):
            get_target_files(path, data_file, pname)


def pandasReadxls(dict_id_name, file_path):
    print file_path
    file_name = file_path.split("/")[-1]
    df = pd.read_excel(file_path, None)
    sheets = df.keys()

    writer = pd.ExcelWriter("ok/" + file_name)

    for sheet in sheets:
        df = pd.read_excel(file_path, sheet_name=sheet)

        name_point1_list = df['point_id'].tolist()
        print len(name_point1_list)
        print sheet
        name1_list = []

        for i in xrange(len(name_point1_list)):

            name1 = dict_id_name.get(name_point1_list[i])
            if name1 != None and name1 != '':
                new_name1 = name_point1_list[i]+'('+name1+')'
                name1_list.append(new_name1)
                df.iat[i, 0] = new_name1

        df.to_excel(writer, sheet, index=False)

    writer.save()


def test():
     df = pd.read_excel('test.xlsx', sheet_name='Sheet3')
     print df
     df.iat[0, 0] = '123'
     df.to_excel('test_save.xlsx',sheet_name='sheet1')



def to_csv():
    """
    转化为csv文件
    """
    sheets = [u'皮尔逊相关系数(5min10min)', u'皮尔逊相关系数-排除无法计算的点(5min10min)',u'正相关top10',
              u'负相关top10', u'正相关top20', u'负相关top20']
    for sheet in sheets:
        print sheet
        df = pd.read_excel('pearn_crr_data.xlsx', sheet_name=sheet)

        name_point1_list = df['name_point1'].tolist()
        name_point2_list = df['name_point2'].tolist()

        data = DataFrame()
        data['name_point1'] = name_point1_list
        data['name_point2'] = name_point2_list
        data.to_csv(sheet+'.csv')

def add_name(result):
    path_dir = 'E:\\shihua\\coding\\new_work\\namecsv\\'

    files = ['5min10min','5min10min_no','top10fu','top10zheng','top20fu','top20zheng']

    for file in files:
        df = pd.read_csv(path_dir+file+'.csv')

        name_point1_list = df['name_point1'].tolist()
        name_point2_list = df['name_point2'].tolist()
        # print name_point1_list

        name1_list = []
        name2_list = []
        for i in xrange(len(name_point1_list)):
            name1 = result.get(name_point1_list[i])
            if name1 != None and name1 != '':
                new_name1 = name_point1_list[i]+'('+name1+')'
                name1_list.append(new_name1)
                df.iat[i, 1] = new_name1

            name2 = result.get(name_point2_list[i])
            if name2 != None and name1 != '':
                new_name2 = name_point2_list[i]+'('+name2+')'
                name2_list.append(new_name2)
                df.iat[i, 2] = new_name2

        # df['name_point1'] = name1_list
        # df['name_point2'] = name2_list

        df.to_csv("E:\shihua\coding\\new_work\\namecsv\\"+file + '_new.csv')


def csv2excel():

    df = pd.read_excel('pearn_crr_data.xlsx', sheet_name=u'皮尔逊相关系数(5min10min)')
    data = pd.read_csv('E:\shihua\coding\\new_work\\namecsv\\5min10min_new.csv')

    name_point1_list = data['name_point1']
    name_point2_list = data['name_point2']
    df['name_point1'] = name_point1_list
    df['name_point2'] = name_point2_list

    df.to_excel('kk.xlsx', u'皮尔逊相关系数(5min10min)')


if __name__ == "__main__":

    dict_id_name = id_name()
    root_dir = "ori_data/"
    save_path = "ok/"
    data_file = []
    get_target_files(root_dir, data_file, ".xls")

    for path in data_file:

        pandasReadxls(dict_id_name,path)


    # test()
    # to_csv()
    # add_name(result)
    # csv2excel()