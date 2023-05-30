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
from budget_check import read_budget
from excel_check import read_excel
from word_check import read_word
from regulation_check import check_regulation

prefix = '{http://pcstd.pcc.gov.tw/2003/eTender}'
class Word2Xml():
    def __init__(self):
        self.amount_type_list = []
        self.middle_type_list = []
        self.is_pass = -1
        self.group_array = ''
        self.type_list = []
        self.thickness_list = []
        self.compare_dict = {}

        self.wordName = ''
        self.excelName = ''
        self.schemaName = ''
        self.drawing_schema = ''
        self.treeName = ''
        self.budget_path = ''
        self.output_path = ''

        self.key_dict = {
            '混凝土強度': 'Concrete/Strength',
            '混凝土深度': 'Concrete/Depth',
            '混凝土厚度': 'Concrete/Thickness',
            '混凝土面積': 'Concrete/Total',
            '混凝土長度': 'Concrete/Length',
            
            '型鋼尺寸': 'MiddleColumn/Steel/Type',
            '型鋼長度': 'MiddleColumn/Steel/Length',
            '中間柱支數': 'MiddleColumn/DrilledPile/Count',
            '開挖面以上': 'MiddleColumn/Steel/TotalUpper',
            '埋入鑽掘樁': 'MiddleColumn/Steel/TotalLower',

            '鑽掘樁直徑': 'MiddleColumn/DrilledPile/Diameter',
            '施作深度': 'MiddleColumn/Length',
            '開挖深度': 'MiddleColumn/Depth',
            '實作深度（樁身埋入深度）': 'MiddleColumn/DrilledPile/Length',

            '鋼筋強度': 'RebarCageGroup/RebarCage/Strength',
            # '主筋': '',
            # '水平筋': '',
            # '剪力筋': '',

            # '鋼筋籠鋼筋重': 'RebarCage/Rebar/Total',
            '支撐/圍囹鋼材噸數': 'Total_SupFen',
            # '末端版重': 'EndPanel/Total',

            '鑽掘樁混凝土強度': 'MiddleColumn/DrilledPile/Concrete/Strength',
            '鑽掘樁鋼筋籠鋼筋重': 'MiddleColumn/RebarCage/Total',
            
            # '導溝行徑米': 'GuideWall/Total',
        }

    def get_concrete_strength(self, drawing_schema, drawing_shema_path):
        drawing_schema_root = ET.parse(drawing_schema).getroot()
        concrete_strength = []
        for i in range(1, 5, 1):
            path = drawing_shema_path.replace('Strength', f'Strength{i}')
            try:
                value = drawing_schema_root.find(f"File/[@Description='設計圖說']/Drawing[@Description='結構一般説明']/{path}/Value").text
                concrete_strength.append(value)
            except:
                pass

        return concrete_strength

    def value_compare(self, key, value):
        # flatten
        new_value = []
        for v in value:
            if isinstance(v, list):
                for vv in v:
                    new_value.append(vv)
            else:
                new_value.append(v)

        value = new_value
        if key == 'Concrete/Strength' and value[0] != '' and value[0] != None:
            value[0] = str(int(value[0]) + 35)

        # float
        try:
            # if less than 2 value, don't comparing
            count = [1 for v in value if v != '' and v != None]
            if sum(count) == 1:
                return ''
            elif sum(count) == 0:
                return 'NA'
            
            value = [float(v) for v in value if v != '' and v != None]
            delta = np.diff(np.sort(value))

            threshold_value = [v * self.threshold for v in value]
            # print(value, threshold_value, delta)
            
            return all([ d <= threshold_value[i] for i, d in enumerate(delta) ])

        # string
        except:
            value = [v.casefold() for v in value if v != '' and v != None]
            return len(set(value)) == 1

    def get_drawing(self, drawing_shema_path, t):
            def find_drawing(root, drawing, t, drawing_path):
                try:
                    return root.find(f"File/[@Description='設計圖說']/Drawing[@Description='{drawing}']/WorkItemType[@Description='{t}']/{drawing_path}/Value").text
                except:
                    return ''
            
            drawing_schema_root = ET.parse(self.drawing_schema).getroot()
            drawing_value = []

            for drawing in ['立面圖', '配筋圖', '平面圖', '結構一般説明']:
                value = find_drawing(drawing_schema_root, drawing, t, drawing_shema_path)
                drawing_value.append(value)

            return drawing_value

    def find_type_drawing(self, drawing_schema):
        drawing_schema_root = ET.parse(drawing_schema).getroot()
        type_list = []
        count_blank = 0

        drawing = '配筋圖'
        for t in drawing_schema_root.find(f"File/[@Description='設計圖說']/Drawing[@Description='{drawing}']"):
            type_list.append(t.attrib['Description'])
            if "空打" in t.attrib['Description']:
                count_blank += 1

        return len(type_list), count_blank, type_list

    def get_value(self, key_dict, key, t):
        try:
            schema_path = key_dict[key]
        except:
            pass

        try:
            design = self.designFile.find(f"./*[@TYPE='{t}']").find(schema_path + '/Value').text
        except:
            design = ''
        
        try:
            quantity = self.quantityFile.find(f"./*[@TYPE='{t}']").find(schema_path + '/Value').text 

            if schema_path == 'MiddleColumn/Steel/Length':
                quantity = float(quantity) + float(self.quantityFile.find(f"./*[@TYPE='{t}']").find('MiddleColumn/Depth/Value').text)
        except:
            quantity = ''

        try:
            drawing = self.get_drawing(schema_path, t.replace('Type', 'TYPE'))
        except:
            drawing = ''
        
        try:
            t = t.replace('Type', 'TYPE')
            budget = self.budgetFile.find(f"./*[@TYPE='{t}']").find(schema_path + '/Value').text 
        except:
            budget = ''
        
        return design, quantity, drawing, budget

    def export_report(self, wordName, excelName, schemaName, drawing_schema, treeName, budget_path, output_path, threshold, station_code):
        self.wordName = wordName
        self.excelName = excelName
        self.schemaName = schemaName
        self.drawing_schema = drawing_schema
        self.treeName = treeName
        self.budget_path = budget_path
        self.output_path = output_path
        self.threshold = threshold

        excel_type_list = []
        drawing_type_list = []
        budget_type_list = []
        
        #複製 schema 為 tree
        shutil.copy(schemaName, treeName)
        tree = ET.parse(treeName)
        root = tree.getroot()
        self.designFile = root[0]
        num_workItemType_design = 0
        self.quantityFile = root[2]
        self.regulationFile = root[3]
        root.append(deepcopy(root[2]))
        root[4].set('Description', '預算書')
        self.budgetFile = root[4]

        if '請選擇' not in wordName:
            num_workItemType_design = read_word(wordName, drawing_schema, self.designFile)
            tree.write(treeName)

        if '請選擇' not in excelName:
            self.concrete_list, self.concrete_type_list, excel_type_list = read_excel(self, excelName, self.quantityFile, self.regulationFile, num_workItemType_design)
            tree.write(treeName)

        if '請選擇' not in budget_path:
            budget_type_list, self.thickness_list, self.compare_dict = read_budget(self.budgetFile, budget_path, station_code)
            tree.write(treeName)

        if '請選擇' not in drawing_schema:
            _, _, drawing_type_list = self.find_type_drawing(drawing_schema)

        length = 0
        all_type_list = [excel_type_list, drawing_type_list, budget_type_list]
        for t in all_type_list:
            if len(t) > length:
                length = len(t)
                self.type_list = t
            
        self.type_list = [t.replace('TYPE', 'Type') for t in self.type_list]
        with open(output_path, 'w', encoding='BIG5', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["項目", 'TYPE', "設計計算書數值", "數量計算書數值", "立面圖", "配筋圖", "平面圖", "結構一般說明", "預算書數值", "是否一致", "備註"])
            for key in self.key_dict:

                for t in self.type_list:
                    design, quantity, drawing, budget = self.get_value(self.key_dict, key, t)
                    compare_result = self.value_compare(self.key_dict[key], [design, quantity, drawing, budget])

                    row = [key, f'{t}', design, quantity, drawing, budget, compare_result]

                    # 圖說結構一般說明
                    if self.key_dict[key] == 'Concrete/Strength':
                        concrete_strength = self.get_concrete_strength(drawing_schema, self.key_dict[key])
                        row.append(f'降階, 結構一般說明: {concrete_strength}')

                    new_row = []
                    for v in row:
                        if isinstance(v, list):
                            for vv in v:
                                new_row.append(vv)
                        else:
                            new_row.append(v)
                    row = new_row

                    if compare_result != 'NA':
                        writer.writerow(row)

        
        
        if '請選擇' not in excelName:
            check_regulation(self)
