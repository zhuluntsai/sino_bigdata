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
        self.group_array = []
        self.type_list = []
        self.thickness_list = []
        self.compare_dict = {}
        self.concrete_strength = []
        self.rebar_strength = []

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
            
            '中間柱型鋼尺寸': 'MiddleColumn/Steel/Type',
            '中間柱型鋼長度': 'MiddleColumn/Steel/Length',
            '中間柱支數': 'MiddleColumn/DrilledPile/Count',
            '開挖面以上': 'MiddleColumn/Steel/TotalUpper',
            '埋入鑽掘樁': 'MiddleColumn/Steel/TotalLower',

            '鑽掘樁直徑': 'MiddleColumn/DrilledPile/Diameter',
            '鑽掘樁施作深度': 'MiddleColumn/Length',
            '開挖深度': 'MiddleColumn/Depth',
            '鑽掘樁實作深度（樁身埋入深度）': 'MiddleColumn/DrilledPile/Length',

            '鑽掘樁混凝土強度': 'MiddleColumn/DrilledPile/Concrete/Strength',
            '鑽掘樁鋼筋籠鋼筋重': 'MiddleColumn/RebarCage/Total',
            '支撐/圍囹鋼材噸數': 'Total_SupFen',
            
            '鋼筋強度': 'RebarCageGroup/RebarCage/Strength',

            # '鋼筋籠鋼筋重': 'RebarCage/Rebar/Total',
            # '末端版重': 'EndPanel/Total',
            # '導溝行徑米': 'GuideWall/Total',
        }
        self.concrete_type_dict = {
            'TYPE I': ['雙牆', '中間樁(柱)', '第1型水泥'],
            'TYPE II': ['單牆', '複合牆', '永久性', '抗浮樁', '第2型水泥']}
        self.rebar_dict = {
            '垂直筋（擋土）': 'RebarGroup/VertRebar/Retaining',
            '垂直筋（開挖）': 'RebarGroup/VertRebar/Excavation',
            '水平筋': 'RebarGroup/HorznRebar',
            '剪力筋': 'RebarGroup/ShearRebar',
        }
        self.rebar_item_dict = {
            '開始深度': 'DepthStart', 
            '結束深度': 'DepthEnd', 
            '設計': 'Type',
        }

    def get_strength(self, drawing_schema):
        drawing_schema_root = ET.parse(drawing_schema).getroot()
        strength_dict = {'Concrete': [],
                         'Rebar': [],
                        }
        for key in strength_dict.keys():
            for i in range(1, 5, 1):
                path = key + f'/Strength{i}'

                # 結構一般說明
                try:
                    value = drawing_schema_root.find(f"File/[@Description='設計圖說']/Drawing[@Description='結構一般說明']/{path}/Value").text
                    strength_dict[key].append(value)
                except Exception as e:
                    pass
                # 結構一般説明
                try:
                    value = drawing_schema_root.find(f"File/[@Description='設計圖說']/Drawing[@Description='結構一般説明']/{path}/Value").text
                    strength_dict[key].append(value)
                except Exception as e:
                    pass

        return strength_dict

    def get_drawing(self, drawing_shema_path, t):
        def find_drawing(root, drawing, t, drawing_path):
            try:
                return root.find(f"File/[@Description='設計圖說']/Drawing[@Description='{drawing}']/WorkItemType[@Description='{t}']/{drawing_path}/Value").text
            except:
                return ''
        
        drawing_schema_root = ET.parse(self.drawing_schema).getroot()
        drawing_value = []

        for drawing in ['立面圖', '配筋圖', '平面圖']:
            value = find_drawing(drawing_schema_root, drawing, t, drawing_shema_path)
            drawing_value.append(value)

        return drawing_value

    def get_type_drawing(self, drawing_schema):
        print('抓取設計圖說')
        drawing_schema_root = ET.parse(drawing_schema).getroot()
        type_list = []
        count_blank = 0

        drawing = '配筋圖'
        for t in drawing_schema_root.find(f"File/[@Description='設計圖說']/Drawing[@Description='{drawing}']"):
            type_list.append(t.attrib['Description'])
            if "空打" in t.attrib['Description']:
                count_blank += 1

        print('設計圖說抓取完成')
        return len(type_list), count_blank, type_list

    def value_compare(self, key, value):
        # flatten
        new_value = []
        for v in value:
            if isinstance(v, list):
                for vv in v:
                    new_value.append(vv)
            else:
                new_value.append(v)

        # if less than 2 value, don't comparing
        value = [ v for v in new_value if v != '' and v != None]
        if len(value) <= 1:
            return ''

        if key == 'Concrete/Strength' and value[0] != '' and value[0] != None:
            value[0] = str(int(value[0]) + 35)

        # float
        try:            
            benchmark = float(value[0])
            threshold_value = benchmark * self.threshold
            return all([ abs(float(v) - benchmark) <= threshold_value for v in value ])

        # string
        except:
            value = [ v.casefold() for v in value ]
            return len(set(value)) == 1

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
            drawing = ['', '', '']
        
        try:
            budget = self.budgetFile.find(f"./*[@TYPE='{t.replace('Type', 'TYPE')}']").find(schema_path + '/Value').text 
        except:
            budget = ''
        
        return design, quantity, drawing, budget

    def export_report(self, wordName='', excelName='', schemaName='', drawing_schema='', treeName='tree.xml', budget_path='', output_path='', threshold=0.05, station_code=''):
        self.wordName = wordName
        self.excelName = excelName
        self.schemaName = schemaName
        self.drawing_schema = drawing_schema
        self.treeName = treeName
        self.budget_path = budget_path
        self.output_path = output_path
        self.threshold = threshold

        word_type_list = []
        excel_type_list = []
        drawing_type_list = []
        budget_type_list = []
        num_workItemType_design = 0
        
        # copy schema as tree
        shutil.copy(schemaName, treeName)
        tree = ET.parse(treeName)
        root = tree.getroot()
        self.designFile = root[0]
        self.quantityFile = root[2]
        self.regulationFile = root[3]
        root.append(deepcopy(root[2]))
        root[4].set('Description', '預算書')
        self.budgetFile = root[4]

        # read excel
        if '請選擇' not in excelName and excelName != '':
            self.concrete_list, self.concrete_type_list, excel_type_list = read_excel(self, excelName, self.quantityFile, self.regulationFile, num_workItemType_design)
            tree.write(treeName)

        # read drawing
        if '請選擇' not in drawing_schema:
            strength_dict = self.get_strength(drawing_schema)
            self.concrete_strength, self.rebar_strength = strength_dict['Concrete'], strength_dict['Rebar']
            try:
                _, _, drawing_type_list = self.get_type_drawing(drawing_schema)
            except:
                print('無法從設計圖說schema抓取type資訊')
                return

        # read budget
        if '請選擇' not in budget_path and excelName != '':
            budget_type_list, self.thickness_list, self.compare_dict = read_budget(self.budgetFile, budget_path, station_code)
            tree.write(treeName)

        length = 0
        all_type_dict = {'數量計算書': excel_type_list,
                         '設計圖說': drawing_type_list,
                         '預算書': budget_type_list,
                         }
        
        # find longest type list as the main type list
        for t in all_type_dict:
            print(t, all_type_dict[t])
            if len(all_type_dict[t]) > length:
                length = len(all_type_dict[t])
                self.type_list = all_type_dict[t]
        
        # read word
        if '請選擇' not in wordName:
            word_type_list, num_workItemType_design = read_word(wordName, self.designFile, self.type_list)
            tree.write(treeName)
    
        print('設計計算書', word_type_list)
        self.type_list = [t.replace('TYPE', 'Type') for t in self.type_list]

        # output csv
        with open(output_path, 'w', encoding='BIG5', newline='') as f:
            writer = csv.writer(f)
            writer.writerow(["項目", 'TYPE', "設計計算書數值", "數量計算書數值", "立面圖", "配筋圖", "平面圖", "預算書數值", "是否一致", "備註"])
            
            # main key dict
            for key in self.key_dict:

                for t in self.type_list:
                    design, quantity, drawing, budget = self.get_value(self.key_dict, key, t)
                    compare_result = self.value_compare(self.key_dict[key], [design, drawing, quantity, budget])

                    row = [key, f'{t}', design, quantity, drawing, budget, compare_result]

                    # 圖說結構一般說明
                    if self.key_dict[key] == 'Concrete/Strength':
                        row.append(f'降階, 結構一般說明: {self.concrete_strength}')
                    elif self.key_dict[key] == 'RebarCageGroup/RebarCage/Strength':
                        row.append(f'結構一般說明: {self.rebar_strength}')

                    new_row = []
                    for v in row:
                        if isinstance(v, list):
                            for vv in v:
                                new_row.append(vv)
                        else:
                            new_row.append(v)
                    row = new_row

                    # if compare_result != '':
                    writer.writerow(row)   

            # support, fence   
            if '請選擇' not in wordName and '請選擇' not in excelName:
                # print(self.quantityFile[1].find('SupportGroup')[0].tag)
                # print(self.quantityFile[1].find('SupportGroup')[0].find('Type/Value').text)
                # print(self.quantityFile[1].find('SupportGroup')[0].get('TYPE'))

                for i in range(len(self.group_array)):
                    try:
                        type_value = self.group_array[i][0]
                        # print("len(self.group_array) = " + str(len(self.group_array)))
                        # print(type_value)
                        for j in range(len(self.designFile.find(f"./*[@TYPE='{type_value}']").find('SupportGroup').findall('Support'))):
                            # print("num_sup = " + str(len(self.designFile.find(f"./*[@TYPE='{type_value}']").find('SupportGroup').findall('Support'))))
                            try:
                                designSup = self.designFile[i].find('SupportGroup')[j]
                                # quantitySup = self.quantityFile[i].find('SupportGroup')[j]
                                quantitySup = self.quantityFile.find(f"./*[@TYPE='{type_value}']").find('SupportGroup')[j]
                                # print('i', i, 'j', j, 'designSup: ', designSup, 'quantitySup: ', quantitySup)
                                # print('designSup: ', designSup.find('Count/Value').text, 'quantitySup: ', quantitySup.find('Count/Value').text)
                                writer.writerow(['支撐層次',type_value,designSup.find('Layer/Value').text,quantitySup.find('Layer/Value').text,'','','','',float(designSup.find('Layer/Value').text)==float(quantitySup.find('Layer/Value').text)])
                                writer.writerow(['支撐桿件',type_value,designSup.find('Type/Value').text,quantitySup.find('Type/Value').text,'','','','',designSup.find('Type/Value').text==quantitySup.find('Type/Value').text.strip()])
                                writer.writerow(['支撐支數',type_value,designSup.find('Count/Value').text,quantitySup.find('Count/Value').text,'','','','',float(designSup.find('Count/Value').text)==float(quantitySup.find('Count/Value').text)])
                            except:
                                # print(i, j)
                                pass
                    except:
                        pass
                
                for i in range(len(self.group_array)):
                    type_value = self.group_array[i][0]
                    # print("len(self.group_array) = " + str(len(self.group_array)))
                    # print(type_value)
                    for j in range(len(self.designFile.find(f"./*[@TYPE='{type_value}']").find('FenceGroup').findall('Fence'))):
                        # print("num_sup = " + str(len(self.designFile.find(f"./*[@TYPE='{type_value}']").find('FenceGroup').findall('Fence'))))
                        try:
                            designSup = self.designFile[i].find('FenceGroup')[j]
                            # quantitySup = self.quantityFile[i].find('SupportGroup')[j]
                            quantitySup = self.quantityFile.find(f"./*[@TYPE='{type_value}']").find('FenceGroup')[j]
                            # print('i', i, 'j', j, 'designSup: ', designSup, 'quantitySup: ', quantitySup)
                            # print('designSup: ', designSup.find('Count/Value').text, 'quantitySup: ', quantitySup.find('Count/Value').text)
                            writer.writerow(['圍囹層次',type_value,designSup.find('Layer/Value').text,quantitySup.find('Layer/Value').text,'','','','',float(designSup.find('Layer/Value').text)==float(quantitySup.find('Layer/Value').text)])
                            writer.writerow(['圍囹桿件',type_value,designSup.find('Type/Value').text,quantitySup.find('Type/Value').text,'','','','',designSup.find('Type/Value').text==quantitySup.find('Type/Value').text.strip()])
                            writer.writerow(['圍囹支數',type_value,designSup.find('Count/Value').text,quantitySup.find('Count/Value').text,'','','','',float(designSup.find('Count/Value').text)==float(quantitySup.find('Count/Value').text)])
                        except:
                            # print(i, j)
                            pass
                
            # rebar
            if '請選擇' not in wordName and '請選擇' not in drawing_schema:
                for key in self.rebar_dict:
                    for t in word_type_list:
                        try:
                            design_rebar = self.designFile.find(f"./*[@TYPE='{t}']").find(self.rebar_dict[key])
                            design_rebar_length = len(design_rebar)
                            drawing_rebar = ET.parse(self.drawing_schema).getroot().find(f"File/[@Description='設計圖說']/Drawing[@Description='配筋圖']/WorkItemType[@Description='{t.replace('Type', 'TYPE')}']/{self.rebar_dict[key].replace('RebarGroup/', '')}")
                            drawing_rebar_length = len(drawing_rebar)
                            for i in range(max(design_rebar_length, drawing_rebar_length)):
                                for item_key in self.rebar_item_dict:
                                    item = self.rebar_item_dict[item_key]
                                    schema_path = f'{self.rebar_dict[key]}/Rebar/{item}'
                                    
                                    try:
                                        design = design_rebar[i].find(f'{item}/Value').text
                                        if item_key == '設計':
                                            design = design.replace(' ', '') + '0'
                                    except:
                                        design = ''

                                    try:
                                        if 'DepthStart' in item:
                                            item = 'StartDepth'
                                        elif 'DepthEnd' in item:
                                            item = 'EndDepth'
                                        
                                        drawing = drawing_rebar[i].find(f'{item}/Value').text
                                    except:
                                        drawing = ''

                                    compare_result = self.value_compare(self.rebar_dict[key], [design, drawing]) 
                                    
                                    row = [key + item_key, f'{t}', design, '', '', drawing, '', '', compare_result]     
                                    new_row = []
                                    for v in row:
                                        if isinstance(v, list):
                                            for vv in v:
                                                new_row.append(vv)
                                        else:
                                            new_row.append(v)
                                    row = new_row

                                    # if compare_result != '':
                                    writer.writerow(row)
                        except Exception as e:
                            # print(e)
                            pass
        
        # regulation check
        if '請選擇' not in excelName and '請選擇' not in drawing_schema and '請選擇' not in budget_path:
            try:
                check_regulation(self)
            except Exception as e:
                pass
