# -*- coding: utf-8 -*-
import xml.etree.cElementTree as ET
import argparse, csv
import numpy as np

prefix = '{http://pcstd.pcc.gov.tw/2003/eTender}'

def find_budget(target, root):
    # exclude special case
    if target[0] == '' or target[0] == '*' or target[0] == '#':
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
    if tag == '*' or tag == '#':
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
    type_list = []
    
    for f in find:
        for ff in f:
            if all(k in ff.text for k in keyword.split(',')):
                if tag != '*' and tag != '#':
                    return find_value(ff.text, front, back)
                is_pass = 1
            
            if is_pass and tag == '#':
                
                for fff in ff:
                    count += 1
                    if front in fff.text:
                        count_list.append(count)
                        if len(set(np.diff(count_list))) <= 1:
                            type_list.append(find_value(fff.text, 'TYPE', '連續壁'))

            # if found keyword in description, return the next item with keyword2(front)
            if is_pass and tag == '*' and front in ff.text :
                try:
                    return find_value(ff.text, front, back)
                except:
                    return f.find(f"{prefix}Quantity").text

    if tag == '#':            
        return type_list

def find_value(value, front, back):
    return value.split(front)[-1].split(back)[0].strip()

def compare_budget(key, value, budget_root, t='', thickness=''):    
    value = [v.replace('S0', t).replace('000cm', f'{thickness}cm') if 'TYPE' in v else v for v in value ]
    budget_value = find_budget(value, budget_root)

    if key == 'Concrete/Thickness':
        budget_value = float(budget_value) / 100
    elif key == 'MiddleColumn/DrilledPile/Diameter':
        budget_value = float(budget_value) / 10

    return budget_value

def find_type(budget_root, station):
    station_code = find_value(station, '開挖支撐及保護，', '站')
    value = ['#', 'DetailList', f'{station_code}站結構工程', '連續壁，(含導溝，厚', '']

    return find_budget(value, budget_root)

def find_thickness(budget_root, type_list):
    value = ['', 'DetailList', '連續壁，(含導溝,TYPE S0', '厚', 'cm']
    return [ find_budget([v.replace('S0', t) if 'TYPE' in v else v for v in value], budget_root) for t in type_list ]

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--amount_path', default='compare/tree.xml')
    parser.add_argument('--budget_path', default='CQ881標土建工程CQ881-11-04_bp_rbid.xml')
    parser.add_argument('--output_path', default='output.csv')
    args = parser.parse_args()

    f = open(args.output_path, 'w')
    writer = csv.writer(f)

    budget_root = ET.parse(args.budget_path).getroot()

    station = '開挖支撐及保護，LG09站'
    type_list = find_type(budget_root, station)
    thickness_list = find_thickness(budget_root, type_list)

    compare_dict = {
        'Concrete/Thickness': ['', 'DetailList', '連續壁，(含導溝,TYPE S0', '厚', 'cm'],
        'Concrete/Total': ['DetailList', '連續壁，(含導溝，厚000cm)，TYPE S0'],
        'Concrete/Strength': ['*', 'CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '混凝土澆置', '材料費，', 'kgf/cm2'],
        
        'GuideWall/Total': ['CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '產品，預拌混凝土材料費，210kgf/cm2，第1型水泥'],
        'RebarCage/Rebar/Total': ['CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '產品，鋼筋，SD420W'],
        'EndPanel/Total': ['CostBreakdownList', '連續壁，(含導溝，厚000cm)，TYPE S0', '產品，金屬材料，鋼料，末端板，分隔板'],
        
        'Total_SupFen': ['CostBreakdownList', station, '臨時擋土支撐工法，支撐系統之型鋼組立'],
        'Total_SupFen2': ['CostBreakdownList', station, '臨時擋土支撐工法，支撐系統之型鋼拆除'],
        
        'MiddleColumn/Steel/TotalUpper': ['*', 'CostBreakdownList', station, '中間樁(柱)', '臨時擋土支撐工法，支撐系統之型鋼拆除', ''],
        'MiddleColumn/Steel/TotalLower': ['CostBreakdownList', station, '產品，結構用鋼材，H型鋼'],
        'MiddleColumn/DrilledPile/Diameter': ['', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', 'D=', 'mm'],
        'MiddleColumn/Length': ['', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '施作深度', '公尺'],
        'MiddleColumn/DrilledPile/Length': ['', 'CostBreakdownList', station, '全套管式鑽掘混凝土基樁', '實作深度', '公尺'],
    }    

    compare_result = {}
    for t, thickness in zip(type_list, thickness_list):
        compare_result[f'TYPE {t}'] = {}

        for key in list(compare_dict.keys()):
            if not any('TYPE' in v for v in compare_dict[key]):
                t = type_list[0]
            budget_value = compare_budget(key, compare_dict[key], budget_root, t, thickness)
            compare_result[f'TYPE {t}'][key] = budget_value
            row = key, f'TYPE {t}', budget_value
            writer.writerow(row)

    first_type = compare_result[f'TYPE {type_list[0]}']
    diameter = int(first_type['MiddleColumn/DrilledPile/Diameter']*10)
    depth = first_type['MiddleColumn/Length']
    real_depth = first_type['MiddleColumn/DrilledPile/Length']
    pile_path = f'全套管式鑽掘混凝土基樁，D={diameter}mm，施作深度{depth}公尺，實作深度{real_depth}公尺'

    compare_dict2 = {
        'MiddleColumn/DrilledPile/Count': ['CostBreakdownList', station, pile_path],
        'MiddleColumn/DrilledPile/Concrete/Strength': ['', 'CostBreakdownList', station, pile_path, '產品，預拌混凝土材料費', '材料費，', 'kgf/cm2'],
        'MiddleColumn/RebarCage/Total': ['CostBreakdownList', station, pile_path, '產品，鋼筋，SD420W'],
    }
    for t, thickness in zip(type_list, thickness_list):
        t = type_list[0]
        for key in list(compare_dict2.keys()):
            budget_value = compare_budget(key, compare_dict2[key], budget_root, t, thickness)
            if key == 'MiddleColumn/RebarCage/Total':
                budget_value = float(compare_result[f'TYPE {t}']['MiddleColumn/DrilledPile/Count']) * float(budget_value)
            compare_result[f'TYPE {t}'][key] = budget_value
            row = key, f'TYPE {t}', budget_value
            writer.writerow(row)

if __name__ == '__main__':
    main()