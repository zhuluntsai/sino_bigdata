import os
from word2xml import Word2Xml

def compare(w2x):
    global type_multiple_select

    prefix = '/home/user/Documents/weilun/sinotech/'
    wordName = 'word-preprocess/data/LG09站地下擋土壁及支撐系統20221212圍囹正確版_修改換行符.docx'
    excelName = 'word-preprocess/data/CQ881標LG09站地工數量-1111230更新.xls'
    drawing_schema = 'word-preprocess/data/drawing_schema.xml'
    schemaName = 'word-preprocess/data/schema.xml'
    budget_path = 'word-preprocess/data/CQ881標土建工程CQ881-11-04_bp_rbid.xml'
    output_path = 'report3.csv'
    treeName = 'tree.xml'
    threshold = 0.05
    station_code = 'LG09'

    if w2x.is_pass != -1:
        group_array = [[] for _ in range(len(w2x.middle_type_list))]
        input_list = [(0, ), (0, ), (1, )]
        # input_list = [(0, ), (0, ), (0, ), (2, )]
        for i, select in enumerate(input_list):
            for s in select:
                group_array[s].append(w2x.amount_type_list[i])
            
        w2x.group_array = group_array

    # os.system(f'python word2xml.py --word_path {wordName} --excel_path {excelName} --schema_path {schemaName} --budget_path {budget_path} --output_path {output_path} --tree_path {treeName}')
    w2x.export_report(
        wordName=wordName, 
        excelName=excelName,
        schemaName=schemaName,
        drawing_schema=drawing_schema,
        budget_path=budget_path,
        output_path=output_path,
        treeName=treeName,
        threshold=threshold,
        station_code=station_code,)

    print(w2x.amount_type_list)
    print(w2x.middle_type_list)
    print(f'比對報告已儲存在 {output_path}') 

    # if amount of word and excel doesn't match, add compare button
    if not w2x.is_pass:
        # type_multiple_select = typeMultipleSelect(root, amount_type_list=word2Xml.amount_type_list, middle_type_list=word2Xml.middle_type_list)

        w2x.is_pass = True 
        return w2x       



box_list = []
w2x = Word2Xml()

w2x = compare(w2x)
w2x = compare(w2x)

# amount_type_list = ['TYPE S1','TYPE S2','TYPE S3']
# middle_type_list = ['TYPE S1','TYPE S3']
# [[0, 1], [2]]

# amount_type_list = ['TYPE T1','TYPE T1A','TYPE T2']
# middle_type_list = ['中間柱1左','中間柱1中','中間柱1右','中間柱2', '中間柱3']
# [[0], [0], [0], [2]]