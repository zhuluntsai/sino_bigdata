#!/usr/bin/env python
# coding: utf-8

import docx
import xlsxwriter
import xml.etree.ElementTree as ET
import os
import shutil
import re
import xlrd
import csv
from copy import deepcopy
import ast
import numpy as np
import math
import itertools

def extract_type(string):
    pattern = r'TYPE\s[IVX]+'
    result = re.findall(pattern, string)
    return result

def countsheet(excel, keyword):
    count = 0 #keyword數量
    count_blank = 0 #空打數量
    sheet_list = []
    
    for sheetName in excel.sheet_names():
        if sheetName.find(keyword)!=-1 and sheetName.find("空打")==-1:
            sheet_list.append(sheetName)
            #print(sheetName)
            count += 1
        elif sheetName.find(keyword)!=-1 and sheetName.find("空打")!=-1:
            count_blank += 1
    return count, count_blank, sheet_list
        
def find_value(value, front, back):
    return value.split(front)[-1].split(back)[0].strip()

def read_excel(self, excelName, quantityFile, regulationFile, num_workItemType_design):
    excel = xlrd.open_workbook(excelName)
    num_workItemType_quantity = 0
    list_workItemType_quantity = []
    list_workItemType_quantity_nested = []

    concrete_list = []
    concrete_type_list = []

    def counttype(sheetName, keyword):
        count = 0 #keyword數量
        type_list = []
        
        sheet = excel.sheet_by_name(sheetName)
        for row in range(8,sheet.nrows):
            if str(sheet.cell_value(row, 1)).find(keyword)!=-1:
                type_list.append(sheet.cell_value(row, 1))
                count += 1 
        return count, type_list
    
    count, count_blank, sheet_list = countsheet(excel, "連續壁")
    self.amount_type_list = [ find_value(s, '(', ')').replace('系列', '').replace('TYPE', 'Type') for s in sheet_list]

    type_list_total = []
    for sheetName in sheet_list:
        count, type_list = counttype(sheetName, "Type")
        num_workItemType_quantity += count
        type_list_total.append(type_list)

    list_workItemType_quantity_nested = type_list_total
    # 使用 chain.from_iterable() 方法打平嵌套列表
    list_workItemType_quantity = list(itertools.chain.from_iterable(type_list_total))
    # print(num_workItemType_quantity)
    # print(list_workItemType_quantity_nested)
    # print(list_workItemType_quantity)

    for i in range(num_workItemType_quantity - 1):
        quantityFile.append( deepcopy(quantityFile[0]) )

    for i in range(num_workItemType_quantity):
        element = quantityFile[i]
        element.set('TYPE', list_workItemType_quantity[i])

    #混凝土強度
    for i in range(len(list_workItemType_quantity_nested)):
        start_row = 8
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            tar_col = 0
            row_first = True
            type_value = list_workItemType_quantity_nested[i][j]
    #         value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Concrete/Total/Value')
            for row in range(start_row,sheet.nrows):
                if str(sheet.cell_value(row, 1)).find("水中混凝土")!=-1 and row_first:
                    #print(sheet.cell_value(row, col))
                    value_Str = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Concrete/Strength/Value')
                    #print(str(sheet.cell_value(row, col)).split(value.get('unit'))[0])
                    value_Str.text = str(sheet.cell_value(row, 1)).split(value_Str.get('unit'))[0]
                    start_row = row+1
                    row_first = False

    #混凝土設計深度
    for i in range(len(list_workItemType_quantity_nested)):
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            tar_col = 0
            row_first = True
            col_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Concrete/Depth/Value')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find(list_workItemType_quantity_nested[i][0])!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("設計深度\n(m)")!=-1 and col_first:
                        #print(sheet.cell_value(row, col))
                        tar_col = col
                        col_first = False
            #print(str(sheet.cell_value(tar_row, tar_col)))
            value.text = str(sheet.cell_value(tar_row, tar_col))

    #混凝土設計厚度
    for i in range(len(list_workItemType_quantity_nested)):
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            tar_col = 0
            row_first = True
            col_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Concrete/Thickness/Value')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("水中混凝土")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("厚度\n(m)")!=-1 and col_first:
                        #print(sheet.cell_value(row, col))
                        tar_col = col
                        col_first = False
            #print(str(sheet.cell_value(tar_row, tar_col)))
            value.text = str(sheet.cell_value(tar_row, tar_col))

    for i in range(len(list_workItemType_quantity_nested)):
        start_row = 8
        tar_col = 0
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            row_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Concrete/Total/Value')
            for row in range(start_row,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("Type")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("數量")!=-1:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
            value.text = str(sheet.cell_value(tar_row, tar_col))
            start_row = tar_row+1

    for i in range(len(list_workItemType_quantity_nested)):
        start_row = 8
        tar_col = 0
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            row_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Concrete/Length/Value')
            for row in range(start_row,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("水中混凝土")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("設計長度\n(m)")!=-1:
                        #print(sheet.cell_value(row, col))
                        tar_col = col
            #print(str(sheet.cell_value(tar_row, tar_col)))
            value.text = str(sheet.cell_value(tar_row, tar_col))
            start_row = tar_row+1

    for i in range(len(list_workItemType_quantity_nested)):
        start_row = 8
        tar_col = 0
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            row_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('U_Board/Total/Value')
            for row in range(start_row,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("末端板")!=-1 or str(sheet.cell_value(row, 1)).find("U型端板")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(8, col)).find("數量")!=-1:
                        #print(sheet.cell_value(row, col))
                        tar_col = col
            # print(str(sheet.cell_value(tar_row, tar_col)))
            value.text = str(sheet.cell_value(tar_row, tar_col))
            start_row = tar_row+1

    #鋼筋籠
    for i in range(len(list_workItemType_quantity_nested)):
        start_row = 8
        tar_col = 0
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            row_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('RebarCageGroup').find("./*[@Description='SD420W']")
            for row in range(start_row,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("SD420W")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(8, col)).find("數量")!=-1:
                        #print(sheet.cell_value(row, col))
                        tar_col = col
            # print(str(sheet.cell_value(tar_row, tar_col)))
            value.text = str(sheet.cell_value(tar_row, tar_col))
            start_row = tar_row+1

    #鋼筋籠
    for i in range(len(list_workItemType_quantity_nested)):
        start_row = 8
        tar_col = 0
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            row_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('RebarCageGroup').find("./*[@Description='SD280W']")
            for row in range(start_row,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("SD280W")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(8, col)).find("數量")!=-1:
                        #print(sheet.cell_value(row, col))
                        tar_col = col
            # print(str(sheet.cell_value(tar_row, tar_col)))
            value.text = str(sheet.cell_value(tar_row, tar_col))
            start_row = tar_row+1


    def findSheet(excel, target):
        targetsheet = []
        count = 0
        for sheet in excel.sheets():
            if sheet.name.find(target)!=-1 and sheet.name.find("混凝土版")==-1:
                targetsheet.append(sheet.name)
                # print(sheet.name, count)
            count += 1
        return targetsheet

    targetsheets = findSheet(excel, "中間柱")
    self.middle_type_list = targetsheets

    group_list_middlecolumn = np.reshape(np.arange(0,len(targetsheets)), (-1,1)).tolist()
    if len(targetsheets) != num_workItemType_design and self.is_pass == -1:
        self.is_pass = False
        return concrete_list, concrete_type_list, list_workItemType_quantity

    group_list_middlecolumn = self.group_array

    #施作深度
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            # print('type_value', type_value)
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/Length')
            # print('node', node)
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
            # print('施作深度', str(sheet.cell_value(tar_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
        except:
            pass        

    #開挖深度
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/Depth')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1 and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(sheet.cell_value(tar_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
        except:
            pass

    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        # type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/Steel/Type')
        #     value = quantityFile[i].find('MiddleColumn/Steel/Type/Value')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
            # print(str(sheet.cell_value(tar_row, tar_col)))
        #     print(node.get('Row'), node.get('Col'))
            node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
        except:
            pass

    #型鋼長度
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/Steel/Length')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1 and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(sheet.cell_value(tar_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
        except:
            pass

    #開挖面以上重
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            # print(type_value)
        #         print(quantityFile.find(f"./*[@TYPE='{type_value}']"))
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/Steel/TotalUpper')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1 and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(sheet.cell_value(tar_row, tar_col)))
            if node.find('Value').text == None:
                node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
            else:
                node.find('Value').text = str(float(node.find('Value').text) + sheet.cell_value(tar_row, tar_col))
        except:
            pass

    #埋入鑽掘樁重
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/Steel/TotalLower')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1 and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(sheet.cell_value(tar_row, tar_col)))
            if node.find('Value').text == None:
                node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
            else:
                node.find('Value').text = str(float(node.find('Value').text) + sheet.cell_value(tar_row, tar_col))
        except:
            pass

    #鑽掘樁直徑
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/DrilledPile/Diameter')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1 and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(int(re.sub('[^0-9.]', '', sheet.cell_value(tar_row, tar_col)))/10))
            node.find('Value').text = str(int(re.sub('[^0-9.]', '', sheet.cell_value(tar_row, tar_col)))/10)
        except:
            pass

    #樁身埋入深度
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/DrilledPile/Length')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find(node.get('Col'))!=-1 and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(sheet.cell_value(tar_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
        except:
            pass

    #中間柱混凝土強度
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        row_first = True
        type_value = group_list_middlecolumn[i][0]
        node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/DrilledPile/Concrete/Strength')
        for row in range(8,sheet.nrows):
            if str(sheet.cell_value(row, 1)).find(node.get('Row'))!=-1 and row_first:
                #print(sheet.cell_value(row, col))
                tar_row = row
                row_first = False
        # print(str(sheet.cell_value(tar_row, 1)).split(node.find('Value').get('unit'))[0])
        node.find('Value').text = str(sheet.cell_value(tar_row, 1)).split(node.find('Value').get('unit'))[0]

    #支數
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/DrilledPile/Count')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)) == node.get('Col') and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(sheet.cell_value(tar_row, tar_col)))
            if node.find('Value').text == None:
                node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
            else:
                node.find('Value').text = str(float(node.find('Value').text) + sheet.cell_value(tar_row, tar_col))
        except:
            pass

    #鋼筋籠總重
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        tar_col = 0
        row_first = True
        col_first = True
    #     type_value = group_list_middlecolumn[i][0]
        try:
            type_value = group_list_middlecolumn[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/RebarCage/Total')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, col)).find(node.get('Row'))!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)) == node.get('Col') and col_first:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
                        col_first = False
            # print(str(sheet.cell_value(tar_row, tar_col)))
            if node.find('Value').text == None:
                node.find('Value').text = str(sheet.cell_value(tar_row, tar_col))
            else:
                node.find('Value').text = str(float(node.find('Value').text) + sheet.cell_value(tar_row, tar_col))
        except:
            pass

    targetsheets = findSheet(excel, "支撐")

    def writeSupFenQuantity(support,total_arr,no_arr,num_arr,type_arr,unitWeight_arr,length_arr):
        for i in range(len(no_arr)-1):
            # print(len(no_arr))
    #         support.append( deepcopy(support[0]) )
            support.insert(0, deepcopy(support[0]))
            i+=1
        for i in range(len(no_arr)):
    #         support[i].find('Total/Value').text = str(format(total_arr[i], '.1f'))
            support[i].find('Layer/Value').text = str(no_arr[i])
            support[i].find('Type/Value').text = str(type_arr[i])
            support[i].find('Count/Value').text = str(num_arr[i])
    #         support[i].find('UnitWeight/Value').text = str(format(unitWeight_arr[i], '.3f'))
    #         support[i].find('Length/Value').text = str(format(length_arr[i], '.1f'))
            i+=1

    sheet = excel.sheet_by_name(targetsheets[0])
    tar_row = []
    type_last = ''
    group_arr = []

    #node = quantityFile[i].find('SupportGroup/Support/Total')
    for row in range(8,sheet.nrows):
        if str(sheet.cell_value(row, 1)).find('支撐')!=-1:
        #print(sheet.cell_value(row, col))
            tar_row.append(row)
    if len(tar_row) != num_workItemType_quantity and self.is_pass == -1:
        self.is_pass = False
        return concrete_list, concrete_type_list

    group_arr = self.group_array
    # print(group_arr)
    # print(tar_row)
    count = 0
    for row in tar_row:
    #     print(sheet.cell_value(row+1, col))
        total = []
        no = []
        num = []
        Type = []
        unitWeight = []
        length = []
        try:
            type_value = group_arr[count][0]
            WorkItemType = group_arr[count][0]
            # print('WorkItemType', WorkItemType)
            for i in range(row+1,sheet.nrows):
                if sheet.cell_value(i,4) != '':
    #                 total.append(sheet.cell_value(i,3))
                    no.append(sheet.cell_value(i,4))
                    num.append(sheet.cell_value(i,5)[0])
                    Type.append(sheet.cell_value(i,5)[1:])
    #                 unitWeight.append(sheet.cell_value(i,6))
    #                 length.append(sheet.cell_value(i,7))
                else:
                    break;
            if type_value != type_last:
                writeSupFenQuantity(quantityFile.find(f"./*[@TYPE='{type_value}']").find('SupportGroup'),total,no,num,Type,unitWeight,length)
                type_last = type_value
            count += 1
        #     writeSupFenQuantity(designFile[count][3],no,num_2,type_2)
    #         print(total)
            # print(no)
            # print(num)
            # print(Type)
    #         print(unitWeight)
    #         print(length)
        except:
            pass

    tar_row = []
    type_last = ''

    for row in range(8,sheet.nrows):
        if str(sheet.cell_value(row, 1)).find('橫擋')!=-1:
            tar_row.append(row)

    count = 0
    for row in tar_row:
    #     print(sheet.cell_value(row+1, col))
        total = []
        no = []
        num = []
        Type = []
        unitWeight = []
        length = []
        try:
            type_value = group_arr[count][0]
            WorkItemType = group_arr[count][0]
            # print('WorkItemType', WorkItemType)
            for i in range(row+1,sheet.nrows):
                if sheet.cell_value(i,4) != '':
    #                 total.append(sheet.cell_value(i,3))
                    no.append(sheet.cell_value(i,4))
                    num.append(sheet.cell_value(i,5)[0])
                    Type.append(sheet.cell_value(i,5)[1:])
    #                 unitWeight.append(sheet.cell_value(i,6))
    #                 length.append(sheet.cell_value(i,7))
                else:
                    break;
            if type_value != type_last:
                writeSupFenQuantity(quantityFile.find(f"./*[@TYPE='{type_value}']").find('FenceGroup'),total,no,num,Type,unitWeight,length)
                type_last = type_value
            count += 1
        #     writeSupFenQuantity(designFile[count][3],no,num_2,type_2)
    #         print(total)
            # print(no)
            # print(num)
            # print(Type)
    #         print(unitWeight)
    #         print(length)
        except:
            pass

    #總計
    tar_row = 0
    row_first = True
    value = quantityFile[0].find('Total_SupFen/Value')
    for row in range(8,sheet.nrows):
        if str(sheet.cell_value(row, 1)).find("總 計")!=-1 and row_first:
            #print(sheet.cell_value(row, col))
            tar_row = row
            row_first = False
    # print(str(sheet.cell_value(tar_row, 3)))
    value.text = str(sheet.cell_value(tar_row, 3))

    
    for i in range(num_workItemType_quantity - 1):
        regulationFile.append( deepcopy(regulationFile[0]) )

    for i in range(num_workItemType_quantity):
        element = regulationFile[i]
        element.set('TYPE', list_workItemType_quantity[i])

    
    

    #TITLE
    for i in range(len(list_workItemType_quantity_nested)):
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            type_value = list_workItemType_quantity_nested[i][j]
            element = regulationFile.find(f"./*[@TYPE='{type_value}']")
            # print(str(sheet.cell_value(5, 0)))
    #         print(re.sub('[^0-9]', '', sheet.cell_value(tar_row, tar_col))
            element.set('Title', str(sheet.cell_value(5, 0)))
            concrete_list.append(str(sheet.cell_value(5, 0)).replace('\u3000', '').replace('\n', ''))

    #混凝土TYPE
    for i in range(len(list_workItemType_quantity_nested)):
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            tar_col = 0
            row_first = True
            col_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            element = regulationFile.find(f"./*[@TYPE='{type_value}']").find('Concrete')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("水中混凝土")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("備  註")!=-1 and col_first:
                        #print(sheet.cell_value(row, col))
                        tar_col = col
                        col_first = False
            # print(extract_type(str(sheet.cell_value(tar_row, tar_col)))[0])
            concrete_type_list.append(extract_type(str(sheet.cell_value(tar_row, tar_col)))[0])

    #         print(re.sub('[^0-9]', '', sheet.cell_value(tar_row, tar_col))
            element.set('Type', extract_type(str(sheet.cell_value(tar_row, tar_col)))[0])

    targetsheets = findSheet(excel, "中間柱")
    #中間柱混凝土TYPE
    for i in range(len(group_list_middlecolumn)):
        sheet = excel.sheet_by_name(targetsheets[i])
        tar_row = 0
        row_first = True
        try:
            type_value = group_list_middlecolumn[i][0]
            element = regulationFile.find(f"./*[@TYPE='{type_value}']").find('MiddleColumn/DrilledPile/Concrete')
            for row in range(8,sheet.nrows):
                if str(sheet.cell_value(row, 1)).find("回填材料")!=-1 and row_first:
                    #print(sheet.cell_value(row, col))
                    tar_row = row
                    row_first = False
        #     print(str(sheet.cell_value(tar_row, 1)))
            # print(extract_type(str(sheet.cell_value(tar_row, 1)))[0])
        #         print(re.sub('[^0-9]', '', sheet.cell_value(tar_row, tar_col))
            element.set('Type', extract_type(str(sheet.cell_value(tar_row, 1)))[0])
        except:
            pass

    return concrete_list, concrete_type_list, list_workItemType_quantity


if __name__ == '__main__':
    read_excel()