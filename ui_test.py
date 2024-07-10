import os
from word2xml import Word2Xml

def compare(w2x):
    global type_multiple_select

    wordName = 'word-preprocess/data/LG09站地下擋土壁及支撐系統20221212圍囹正確版_修改換行符.docx'
    # wordName = '請選擇'
    excelName = 'word-preprocess/data/CQ881標LG09站地工數量-1111230更新.xls'
    # excelName = '請選擇'
    # drawing_schema = 'word-preprocess/data/drawing_schema.xml'
    drawing_schema = '請選擇'
    budget_path = 'word-preprocess/data/CQ881標土建工程CQ881-11-04_bp_rbid.xml'
    # budget_path = '請選擇'
    schemaName = 'word-preprocess/data/schema.xml'
    output_path = 'report3.csv'
    treeName = 'tree.xml'
    threshold = 0.05
    station_code = 'LG09'
    input_list = [(0, ), (0, ), (1, )]

    wordName = 'LG10/(已取代)03-2.LG10站地下擋土壁及支撐系統.docx'
    # wordName = '請選擇'
    excelName = 'LG10/(已修改支撐TYPE、調整中間樁TYPE)0080B-CQ881標LG10站地工數量-11004062.xls'
    # excelName = '請選擇'
    # drawing_schema = 'word-preprocess/data/LG10_drawing_schema.xml'
    drawing_schema = '請選擇'
    budget_path = 'word-preprocess/data/CQ881標土建工程CQ881-11-04_bp_rbid.xml'
    # budget_path = '請選擇'
    station_code = 'LG10'
    input_list = [(0, 3, 4, ), (1, )]

    # wordName = 'unearthed_section/明挖覆蓋隧道及出土段地下擋土壁及支撐系統.docx'
    # excelName = 'unearthed_section/0080B_CQ881明挖覆蓋隧道及出土段地工數量-1090427.xls'
    # drawing_schema = '請選擇'
    # budget_path = 'word-preprocess/data/CQ881標土建工程CQ881-11-04_bp_rbid.xml'
    schemaName = 'schema_rowpile.xml'

    # wordName = 'Y37/02-2.計算書0606B_Y37_已取代空白.docx'
    # excelName = 'Y37/CF752-Y37站地工數量-1130606_擋土開挖、建物保護_拆分TYPE工作表.xls'
    # drawing_schema = '請選擇'
    # budget_path = '請選擇'
    # # budget_path = 'Y37/(預算書)臺北都會區大眾捷運系統環狀線東環段CF120區段標-CF752標土建工程CF752-01-01_ap_bdgt.xml'
    # station_code = 'Y37'
    # input_list = [(0,), (1,)]

    wordName = 'Y38/02-2.計算書0606B_Y38+中央避車線_已取代空白-1130708.docx'
    excelName = 'Y38/CF761-Y38站地工數量-1130408_擋土開挖、建物保護_拆分TYPE工作表.xlsx'
    drawing_schema = '請選擇'
    budget_path = '請選擇'
    budget_path = 'Y38/(預算書)臺北都會區大眾捷運系統環狀線東環段CF761標土建工程CF761-C-05_ap_bdgt.xml'
    station_code = 'Y38'
    input_list = [(1,), (), (0,), (), 
                  (), (), (2,), (), 
                  (3,), (), (4,), (), 
                  (), (), (), ()]

    # wordName = 'Y39/計算書0606B_Y39合併.docx'
    # excelName = 'Y39/Y39-Volume0514(君毅)_擋土開挖、建物保護、推管工作井、監測儀器.xls'
    # drawing_schema = '請選擇'
    # budget_path = '請選擇'
    # budget_path = 'Y39/(預算書)臺北都會區大眾捷運系統環狀線東環段CF762標土建工程CF762-C-05_ap_bdgt.xml'
    # station_code = 'Y39'
    # input_list = [(0,), (1,), (1,), (2,), (2,),
    #               (3,), (3,), (3,), (3,), (3,),
    #               (3,)]

    if w2x.is_pass != -1:
        group_array = [[] for _ in range(len(w2x.middle_type_list))]
        for i, select in enumerate(input_list):
            for s in select:
                group_array[s].append(w2x.amount_type_list[i])
            
        group_array = [ g for g in group_array if len(g) != 0 ]
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

    if w2x.is_pass:
        print(f'比對報告已儲存在 {output_path}') 

    # if amount of word and excel doesn't match, add compare button
    if not w2x.is_pass:
        # type_multiple_select = typeMultipleSelect(root, amount_type_list=word2Xml.amount_type_list, middle_type_list=word2Xml.middle_type_list)

        print(w2x.middle_type_list)
        print(w2x.amount_type_list)
        w2x.is_pass = True 
        return w2x       



box_list = []
w2x = Word2Xml()

w2x = compare(w2x)
# w2x = compare(w2x)

# amount_type_list = ['TYPE S1','TYPE S2','TYPE S3']
# middle_type_list = ['TYPE S1','TYPE S3']
# [[0, 1], [2]]

# amount_type_list = ['TYPE T1','TYPE T1A','TYPE T2']
# middle_type_list = ['中間柱1左','中間柱1中','中間柱1右','中間柱2', '中間柱3']
# [[0], [0], [0], [2]]