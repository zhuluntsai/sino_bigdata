import xml.etree.cElementTree as ET

prefix = '{http://pcstd.pcc.gov.tw/2003/eTender}'

def find_budget(target):
    xml_path = '(空白標單)臺北都會區大眾捷運系統萬大線(第二期工程)-CQ881標土建工程CQ881-11-04_bp_rbid.xml'
    tree = ET.parse(xml_path)
    root = tree.getroot()

    if target[0] == '' or target[0] == '*':
        return find_budget_in(target, root)

    if target[0] == 'DetailList':
        xpath = f"{prefix+target[0]}/{prefix}PayItem/{prefix}PayItem/[{prefix}Description='{target[1]}']/{prefix}Quantity"
        
    elif target[0] == 'CostBreakdownList':
        xpath = f"{prefix+target[0]}/"
        for i in range(len(target) - 1):
            xpath += f"{prefix}WorkItem/[{prefix}Description='{target[i + 1]}']/"
        xpath += f"{prefix}Quantity"

    return root.find(xpath).text

def find_budget_in(target, root):
    tag = target.pop(0)
    keyword = target.pop(-3)
    front = target.pop(-2) # keyword2
    back = target.pop(-1)
    is_pass = 1
    
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

            if front in ff.text and is_pass and tag == '*':
                return f.find(f"{prefix}Quantity").text


def find_number(value, front, back):
    return value.split(front)[-1].split(back)[0]

def find_amount(key):
    xml_path = 'schema.xml'

    tree = ET.parse(xml_path)
    root = tree.getroot()
    
    value = root.find(f"File/[@Description='數量計算書']/WorkItemType[@Description='TYPE S1']/{key}/Value")
    return value.text

def compare(key, compare_dict):
    target_value = find_amount(key)
    budget_value = find_budget(compare_dict[key])
    delta = abs(round(float(target_value)) - round(float(budget_value))) < 0.1
    print(key)
    print(target_value, '\t', budget_value, '\t', delta)


def main():
    compare_dict = {
        'Concrete/Total': ['DetailList', '連續壁，(含導溝，厚100cm)，TYPE S1'],
        'Concrete/Thickness': ['', 'DetailList', '連續壁，(含導溝,TYPE S1', '厚', 'cm'],
        'Concrete/Strength': ['', 'CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S1', '產品，預拌混凝土材料費', '材料費，', 'kgf/cm2'],
        
        'GuideWall/Total': ['CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S1', '產品，預拌混凝土材料費，210kgf/cm2，第1型水泥'],
        'RebarCage/Rebar/Total': ['CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S1', '產品，鋼筋，SD420W'],
        'EndPanel/Total': ['CostBreakdownList', '連續壁，(含導溝，厚100cm)，TYPE S1', '產品，金屬材料，鋼料，末端板，分隔板'],
        'SupportGroup/SteelWeight': ['CostBreakdownList', '開挖支撐及保護，LG09站', '臨時擋土支撐工法，支撐系統之型鋼組立'],
        'SupportGroup/SteelWeight2': ['CostBreakdownList', '開挖支撐及保護，LG09站', '臨時擋土支撐工法，支撐系統之型鋼拆除'],
        
        'MiddleColumn/Steel/Count': ['CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁，D=1000mm，施作深度27公尺，實作深度5公尺'],
        'MiddleColumn/Steel/Above': ['*', 'CostBreakdownList', '開挖支撐及保護，LG09站', '中間樁(柱)', '臨時擋土支撐工法，支撐系統之型鋼拆除', ''],
        'MiddleColumn/Steel/Under': ['CostBreakdownList', '開挖支撐及保護，LG09站', '產品，結構用鋼材，H型鋼'],
        
        'MiddleColumn/DrilledPile/Diameter': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁', 'D=', 'mm'],
        'MiddleColumn/DrilledPile/Depth': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁', '施作深度', '公尺'],
        'MiddleColumn/DrilledPile/RealDepth': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁', '實作深度', '公尺'],
        'MiddleColumn/DrilledPile/Strength': ['', 'CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁，D=1000mm，施作深度27公尺，實作深度5公尺', '產品，預拌混凝土材料費', '材料費，', 'kgf/cm2'],
        'MiddleColumn/DrilledPile/Backfill': ['CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁，D=1000mm，施作深度27公尺，實作深度5公尺', '構造物回填，借土，第Ⅰ類材料'],
        'MiddleColumn/DrilledPile/SteelCageWeight': ['CostBreakdownList', '開挖支撐及保護，LG09站', '全套管式鑽掘混凝土基樁，D=1000mm，施作深度27公尺，實作深度5公尺', '產品，鋼筋，SD420W'],
    }
    # key = list(compare_dict.keys())[3]
    # compare(key, compare_dict)

    # key = list(compare_dict.keys())[2]
    # compare(key, compare_dict)

    for key in list(compare_dict.keys()):
        compare(key, compare_dict)
        

if __name__ == '__main__':
    main()