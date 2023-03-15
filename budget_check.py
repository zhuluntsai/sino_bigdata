# -*- coding: utf-8 -*-
import xml.etree.cElementTree as ET
import argparse, csv
import numpy as np

prefix = '{http://pcstd.pcc.gov.tw/2003/eTender}'

def find_amount(key, root, i):
    try:
        return root.find(f"File/[@Description='數量計算書']/WorkItemType[@Description='TYPE S{i}']/{key}/Value").text
    except:
        # print(f"File/[@Description='數量計算書']/WorkItemType[@Description='TYPE S1']/{key}/Value")
        return 0

def find_budget(target, root):
    # exclude special case
    if target[0] == '' or target[0] == '*':
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
    if tag == '*':
        is_pass = 0

    if target[0] == 'DetailList':
        xpath = f"{prefix+target[0]}/{prefix}PayItem/{prefix}PayItem"
    
    elif target[0] == 'CostBreakdownList':
        xpath = f"{prefix+target[0]}/"
        for i in range(len(target) - 1):
            xpath += f"{prefix}WorkItem/[{prefix}Description='{target[i + 1]}']/"
    
    find = root.findall(xpath)
    for f in find:
        for ff in f:
            if all(k in ff.text for k in keyword.split(',')):
                if tag != '*':
                    return find_number(ff.text, front, back)
                is_pass = 1

            # if found keyword in description, return the next item with keyword2(front)
            if front in ff.text and is_pass and tag == '*':
                try:
                    return find_number(ff.text, front, back)
                except:
                    return f.find(f"{prefix}Quantity").text

def find_number(value, front, back):
    return value.split(front)[-1].split(back)[0]

def compare_budget(key, value, amount_root, budget_root, i = 1):
    target_value = find_amount(key, amount_root, i)
    budget_value = find_budget(value, budget_root)

    if key == 'Concrete/Thickness':
        budget_value = float(budget_value) / 100
    elif key == 'MiddleColumn/DrilledPile/Diameter':
        budget_value = float(budget_value) / 10

    delta = abs(round(float(target_value)) - round(float(budget_value))) < 0.1
    return key, target_value, budget_value, delta

def main(num_workItemType):
    parser = argparse.ArgumentParser()
    parser.add_argument('--amount_path', default='compare/tree.xml')
    parser.add_argument('--budget_path', default='CQ881標土建工程CQ881-11-04_bp_rbid.xml')
    parser.add_argument('--output_path', default='output.csv')
    args = parser.parse_args()

    f = open(args.output_path, 'w')
    writer = csv.writer(f)

    amount_root = ET.parse(args.amount_path).getroot()
    budget_root = ET.parse(args.budget_path).getroot()

    compare_dict = {
        'Concrete/Thickness': ['', 'DetailList', '連續壁，(含導溝,TYPE S0', '厚', 'cm'],
        'Concrete/Total': ['DetailList', '連續壁，(含導溝，厚100cm)，TYPE S0'],
        'Concrete/Strength': ['*', 'CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S0', '混凝土澆置', '材料費，', 'kgf/cm2'],
        
        'GuideWall/Total': ['CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S0', '產品，預拌混凝土材料費，210kgf/cm2，第1型水泥'],
        'RebarCage/Rebar/Total': ['CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S0', '產品，鋼筋，SD420W'],
        'EndPanel/Total': ['CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S0', '產品，金屬材料，鋼料，末端板，分隔板'],
        
        'SupportGroup/SteelWeight': ['CostBreakdownList', '開挖支撐及保護，LG09站', '臨時擋土支撐工法，支撐系統之型鋼組立'],
        'SupportGroup/SteelWeight2': ['CostBreakdownList', '開挖支撐及保護，LG09站', '臨時擋土支撐工法，支撐系統之型鋼拆除'],
        'SupportGroup/Support/Count': ['CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁，D=1000mm，施作深度27公尺，實作深度5公尺'],
        
        'MiddleColumn/Depth': ['*', 'CostBreakdownList', '開挖支撐及保護，LG09站', '中間樁(柱)', '臨時擋土支撐工法，支撐系統之型鋼拆除', ''],
        'MiddleColumn/Steel/Length': ['CostBreakdownList', '開挖支撐及保護，LG09站', '產品，結構用鋼材，H型鋼'],
        'MiddleColumn/DrilledPile/Diameter': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁', 'D=', 'mm'],
        'MiddleColumn/Length': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁', '施作深度', '公尺'],
        'MiddleColumn/DrilledPile/Length': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁', '實作深度', '公尺'],
        'MiddleColumn/DrilledPile/Strength': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁，D=1000mm，施作深度27公尺，實作深度5公尺', '產品，預拌混凝土材料費', '材料費，', 'kgf/cm2'],
        'MiddleColumn/DrilledPile/SteelCageWeight': ['CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁，D=1000mm，施作深度27公尺，實作深度5公尺', '產品，鋼筋，SD420W'],
    }

    thickness = [100] * (num_workItemType + 1)
    for key in list(compare_dict.keys()):
        
        # if there is any key has 'TYPE'
        if any('TYPE' in v for v in compare_dict[key]):
            # get the index of the item with TYPE
            type_index = [i for i, v in enumerate(compare_dict[key]) if 'TYPE' in v][0]

            for i in range(1, num_workItemType + 1):
                compare_dict[key][type_index] = compare_dict[key][type_index].replace(f'TYPE S{i - 1}', f'TYPE S{i}').replace(f'{thickness[i - 1]}cm', f'{thickness[i]}cm')
                value = compare_dict[key].copy()
                try:
                    key, target_value, budget_value, delta = compare_budget(key, value, amount_root, budget_root, i)
                    row = key, f'TYPE S{i}', target_value, budget_value, delta
                    writer.writerow(row)

                    if key == 'Concrete/Thickness':
                        thickness[i] = int(budget_value * 100)

                except:
                    pass
        else:
            value = compare_dict[key].copy()
            key, target_value, budget_value, delta = compare_budget(key, value, amount_root, budget_root)
            row = key, target_value, budget_value, delta
            writer.writerow(row)
        

if __name__ == '__main__':
    num_workItemType = 3
    main(num_workItemType)