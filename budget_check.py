# -*- coding: utf-8 -*-
import xml.etree.cElementTree as ET
import argparse, csv
import numpy as np
from copy import deepcopy

prefix = '{http://pcstd.pcc.gov.tw/2003/eTender}'

def find_budget(target, root):
    # exclude special case
    if target[0] in ['', '*', '**', '***', '#']:
        return find_budget_in(target, root)

    # common case
    if target[0] == 'DetailList':
        xpath = f"{prefix+target[0]}/{prefix}PayItem/{prefix}PayItem/[{prefix}Description='{target[1]}']/{prefix}Quantity"
        
    elif target[0] == 'CostBreakdownList':
        xpath = f"{prefix+target[0]}/"
        for i in range(len(target) - 1):
            xpath += f"{prefix}WorkItem/[{prefix}Description='{target[i + 1]}']/"
        xpath += f"{prefix}Quantity"

    return root.find(xpath).text

# target value is in item
def find_budget_in(target, root):
    tag = target.pop(0)
    keyword = target.pop(-3)
    front = target.pop(-2) # keyword2
    back = target.pop(-1)
    is_pass = 1
    
    # if same item name
    if tag == '*' or tag == '#' or tag == '**':
        is_pass = 0

    if target[0] == 'DetailList' and tag != '#':
        xpath = f"{prefix+target[0]}/{prefix}PayItem/{prefix}PayItem"
    elif target[0] == 'DetailList' and tag == '#':
        xpath = f"{prefix+target[0]}/{prefix}PayItem"
    elif target[0] == 'CostBreakdownList':
        xpath = f"{prefix+target[0]}/"
        for i in range(len(target) - 1):
            xpath += f"{prefix}WorkItem/[{prefix}Description='{target[i + 1]}']/"
    
    find = root.findall(xpath)
    count = 0
    count_list = []
    value_list = []
    
    for f in find:
        for ff in f:
            if all(k in ff.text for k in keyword.split(',')):
                if tag != '*' and tag != '#' and tag != '**' and tag != '***':
                    return find_value(ff.text, front, back)
                is_pass = 1

                # collect value with keyword
                if tag == '***':
                    value_list.append(find_value(ff.text, front, back))
            
            if is_pass and tag == '#':
                for fff in ff:
                    count += 1
                    if front in fff.text and keyword in f[0].text:
                        count_list.append(count)
                        if len(set(np.diff(count_list))) <= 1:
                            # value_list.append(fff.text.split('，')[-1])
                            value_list.append(fff.text)
            
            # if found keyword in description, return the next item with keyword2(front)
            if is_pass and tag == '*' and front in ff.text:
                try:
                    return find_value(ff.text, front, back)
                except:
                    return f.find(f"{prefix}Quantity").text
            
            if is_pass and tag == '**' and front in ff.text:
                return f.find(f"{prefix}Quantity").text
            #     return ff.text

    if tag == '#' or tag == '***':      
        return value_list

def find_value(value, front, back):
    return value.split(front)[-1].split(back)[0].strip()

def compare_budget(key, value, budget_root, keyword_list, t=''):    
    value = [v.replace(v, t) if any(k in v for k in keyword_list) else v for v in value ] 
    budget_value = find_budget(value, budget_root)

    if key == 'Concrete/Thickness':
        budget_value = float(budget_value) / 100
    elif key == 'MiddleColumn/DrilledPile/Diameter':
        budget_value = float(budget_value) / 10
    elif key == 'RebarCageGroup/RebarCage/Strength':
        budget_value = find_value(budget_value, 'SD', 'W') + '0'

    return budget_value

def read_budget(budgetFile, budget_path, station_code):
    print('抓取預算書')
    budget_root = ET.parse(budget_path).getroot()

    # station_code = '明挖覆蓋隧道'
    # station_code = 'LG10站'
    station = f'開挖支撐及保護，{station_code}站'

    keyword_dict = {
        '連續壁': 'DetailList',
        '排樁': 'DetailList',
        '鋼板樁': 'CostBreakdownList',
        '基樁': 'DetailList',
        }

    # prepare schema
    type_list =[ t for k, l in keyword_dict.items() for t in find_budget(['#', l, f'{station_code}', k, ''], budget_root) ]
    type_list = [ t for t in type_list if '壁體敲除' not in t]
    # delete duplicate
    type_list = list(dict.fromkeys(type_list))

    for i in range(len(type_list)-1):
        budgetFile.append(deepcopy(budgetFile[0]))
        budgetFile[i].set('TYPE', type_list[i])
    budgetFile[-1].set('TYPE', type_list[-1])

    # prepare middle coloumn schema
    middle_column_list = find_budget(['#', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', ''], budget_root)
    middleColumnGroup = budgetFile.find(f"./*[@TYPE='{type_list[0]}']").find('MiddleColumnGroup')
    for i in range(len(middle_column_list)-1):
        middleColumnGroup.insert(0, deepcopy(middleColumnGroup[0]))

    compare_dict = {
        'DiaphragmWall/Thickness': ['', 'DetailList', '連續壁，(含導溝,TYPE S0', '厚', 'cm'],
        'DiaphragmWall/Concrete/Total': ['DetailList', '連續壁，(含導溝，厚000cm)，TYPE S0'],
        'DiaphragmWall/Concrete/Strength': ['*', 'CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '混凝土澆置', '材料費，', 'kgf/cm2'],
        'DiaphragmWall/GuideWall/Total': ['CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '產品，預拌混凝土材料費，210kgf/cm2，第1型水泥'],
        'DiaphragmWall/RebarWeight': ['CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '產品，鋼筋，SD420W'],
        'DiaphragmWall/U_Board/Total': ['CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '產品，金屬材料，鋼料，末端板，分隔板'],
        
        'Rowpile/Height': ['', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '全套管式鑽掘混凝土基樁', '施作深度', '公尺'],
        'Rowpile/Diameter': ['', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '全套管式鑽掘混凝土基樁', 'D=', 'mm'],
        'Rowpile/Concrete/Strength': ['', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '場鑄', '土', 'kgf'],
        'Rowpile/Count': ['CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '全套管式鑽掘混凝土基樁，D=800mm，施作深度23公尺'],
        'Rowpile/RebarWeight': ['**', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '全套管式鑽掘混凝土基樁，D=800mm，施作深度23公尺', '', '產品，鋼筋', ''],
        'Rowpile/Concrete/Total': ['**', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '全套管式鑽掘混凝土基樁，D=800mm，施作深度23公尺', '', '預拌混凝土', ''],
        
        'Beam/Concrete/Total': ['**', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '', '場鑄', ''],
        'Beam/formwork': ['**', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '', '模板', ''],
        'Beam/RebarWeight': ['**', 'CostBreakdownList', '全套管式鑽掘混凝土基樁，排樁', '', '鋼筋', ''],

        'Sheetpile/Total': ['CostBreakdownList', station, '臨時擋土樁設施，鋼板樁，L=0000m'],
        'Sheetpile/Height': ['', 'CostBreakdownList', station, '臨時擋土樁設施，鋼板樁', 'L=', 'm'],
        
        'Total_SupFen': ['CostBreakdownList', station, '臨時擋土支撐工法，支撐系統之型鋼組立'],
        'TotalAmount': ['CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S0'],
        'RebarCageGroup/RebarCage/Strength': ['**', 'CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '鋼筋籠組立及吊裝', '產品，鋼筋', ''],
    }   
    
    middle_column_compare_dict = {
        'Steel/TotalUpper': ['*', 'CostBreakdownList', station, '中間樁(柱)', '臨時擋土支撐工法，支撐系統之型鋼拆除', ''],
        'Steel/TotalLower': ['CostBreakdownList', station, '產品，結構用鋼材，H型鋼'],
        
        'DrilledPile/Diameter': ['', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', 'D=', 'mm'],
        'DrilledPile/Length': ['', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '實作深度', '公尺'],
        'DrilledPile/Concrete/Total': ['**', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '', '預拌混凝土材料費', ''],
        'DrilledPile/Concrete/Strength': ['', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '預拌混凝土材料費', '材料費，', 'kgf/cm2'],
        'DrilledPile/Backfill/Total': ['**', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '', '構造物回填', ''],
        'DrilledPile/Count': ['CostBreakdownList', station, '全套管式鑽掘混凝土基樁'],
        
        'Rebar/Total': ['CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '產品，鋼筋，SD420W'],
        'TotalLength': ['CostBreakdownList', station, '鑽掘樁，中間樁吊裝'],
        'Length': ['', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '施作深度', '公尺'],     
    }
    
    keyword_list = ['連續壁', '排樁', '鋼板樁']
    for t in type_list:
        for key in list(compare_dict.keys()):
            try:
                budget_value = compare_budget(key, compare_dict[key], budget_root, keyword_list, t)
                budgetFile.find(f"./*[@TYPE='{t}']").find(key + '/Value').text = str(budget_value)
            except:
                pass
                # print('error', t, key)

    keyword_list = ['連續壁', '排樁', '鋼板樁', '基樁']
    for m, t in zip(middleColumnGroup, middle_column_list):
        for key in list(middle_column_compare_dict.keys()):
            try:
                budget_value = compare_budget(key, middle_column_compare_dict[key], budget_root, keyword_list, t)
                m.find(key + '/Value').text = str(budget_value)
            except:
                pass
                # print(t, key)
        
    print('預算書抓取完成')
    return type_list, middle_column_list