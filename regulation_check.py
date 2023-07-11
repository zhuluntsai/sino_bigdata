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
from budget_check import find_budget

prefix = '{http://pcstd.pcc.gov.tw/2003/eTender}'
concrete_type_dict = {
    'TYPE I': ['雙牆', '中間樁(柱)', '第1型水泥'],
    'TYPE II': ['單牆', '複合牆', '永久性', '抗浮樁', '第2型水泥']}

def get_unit(self, quantityFile, drawing_schema, budget_path, type_list):
    path = 'Concrete/Total'
    t = type_list[0]

    drawing_schema_root = ET.parse(drawing_schema).getroot()
    drawing = drawing_schema_root.find(f"File/[@Description='設計圖說']/Drawing[@Description='立面圖']/WorkItemType[@Description='{t.replace('Type', 'TYPE')}']/{path}/Value").attrib['unit']

    quantity = quantityFile.find(f"./*[@TYPE='{t}']").find(path + '/Value').attrib['unit']

    budget_root = ET.parse(budget_path).getroot()
    value = [v.replace('TYPE S0', t).replace('Type', 'TYPE').replace('000cm', f'{self.thickness_list[0]}cm') if 'TYPE' in v else v for v in self.compare_dict[path] ]
    xpath = f"{prefix+value[0]}/{prefix}PayItem/{prefix}PayItem/[{prefix}Description='{value[1]}']/{prefix}Unit"
    budget = budget_root.find(xpath).text

    return drawing, quantity, budget
    
def concrete_type_classification(concrete):
    for key in concrete_type_dict:
        for keyword in concrete_type_dict[key]:
            if keyword in concrete:
                return key

def get_budget_list(compare_dict, budget_path, type_list, thickness_list, path):
    budget_root = ET.parse(budget_path).getroot()
    budget_type_list = []

    for t, thickness in zip(type_list, thickness_list):
        value = [v.replace('TYPE S0', t).replace('Type', 'TYPE').replace('000cm', f'{thickness}cm') if 'TYPE' in v else v for v in compare_dict[path] ]
        value[0] = '**'
        budget_value = find_budget(value, budget_root)
        budget_type_list.append(budget_value)
    
    return budget_type_list

def check_regulation(self):    
    output_path = self.output_path.replace('.csv', '_regulation.csv')
    with open(output_path, 'w', encoding='BIG5', newline='') as f:
        writer = csv.writer(f)

        writer.writerow(["規範校合"])
        writer.writerow(['項目', 'TYPE', '規範', "設計計算書數值", "數量計算書數值", "設計圖說", "預算書數值", "是否一致", "備註"])

        key = '計量方式連續壁'
        description = '第02266章\n連續壁\n頁碼：02266-16'
        drawing, quantity, budget = get_unit(self, self.quantityFile, self.drawing_schema, self.budget_path, self.type_list)
        unit_compare_result = self.value_compare('', [drawing, quantity, budget])
        row = [key, '', description, '', quantity, drawing, budget, unit_compare_result]
        writer.writerow(row)

        key = '混凝土強度'
        description = '第03010章\n水中混凝土：245kgf/cm2等級混凝土'
        _, quantity, _, budget = self.get_value(self.key_dict, key, self.type_list[0])
        drawing = self.concrete_strength[0]
        strength_compare_result = self.value_compare('', [drawing, quantity, budget])
        row = [key, '', description, '', quantity, drawing, budget, strength_compare_result]
        writer.writerow(row)

        key = '鋼筋強度'
        path = 'RebarCageGroup/RebarCage/Strength'
        description = '第3210章\n鋼筋\n頁碼：03210-4'
        drawing = self.get_concrete_strength(self.drawing_schema, 'Rebar/Strength')[0]
        budget_list = get_budget_list(self.compare_dict, self.budget_path, self.type_list, self.thickness_list, path)
        for t, budget in zip(self.type_list, budget_list):
            _, quantity, _, _ = self.get_value(self.key_dict, key, t)
            budget = f"{budget.split('產品，鋼筋，SD')[-1].split('W')[0].strip()}0"
            strength_compare_result = self.value_compare('', [drawing, quantity, budget])
            row = [key, f'TYPE {t}', description, '', quantity, drawing, budget, strength_compare_result]
            writer.writerow(row)
        
        writer.writerow(['項目', 'TYPE', '規範', "數量計算書混凝土", "數量計算書混凝土型式", "數量計算書混凝土TYPE", '預算書混凝土TYPE', "是否一致", "備註"])
        key = '混凝土型式'
        path = 'Concrete/Strength'
        description = '第03010章\n卜特蘭水泥混凝土\n頁碼：03010-9'
        budget_type_list = get_budget_list(self.compare_dict, self.budget_path, self.type_list, self.thickness_list, path)
        for concrete, quantity_type, budget_type, t in zip(self.concrete_list, self.concrete_type_list, budget_type_list, self.type_list):
            concrete_type = concrete_type_classification(concrete)
            budget_type = concrete_type_classification(budget_type)
            concrete_type_compare_result = self.value_compare('', [concrete_type, quantity_type, budget_type])
            row = [key, f'TYPE {t}', description, concrete, concrete_type, quantity_type, budget_type, concrete_type_compare_result]
            writer.writerow(row)


if __name__ == '__main__':
    check_regulation()