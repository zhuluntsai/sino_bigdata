# from gensim.models import word2vec
from gensim.models import Word2Vec
import spacy, jieba
import numpy as np
import pandas as pd
import xml.etree.cElementTree as ET
from ckiptagger import data_utils, construct_dictionary, WS, POS, NER
from typing import List

def read_xml():
    xml_path = '(空白標單)臺北都會區大眾捷運系統萬大線(第二期工程)-CQ881標土建工程CQ881-11-04_bp_rbid.xml'

    ws = WS("data")

    tree = ET.parse(xml_path)
    root = tree.getroot()

    # print('payitem: ', root[0]) 
    # print('payitem: [0]: ', root[1][0])

    text_list = []

    for r in root[1][7]:
        if 'PayItem' in r.tag:
            for d in r :
                if 'Description' in d.tag and '連續壁' in d.text and 'C' in d.text and d.text != '                                                                                                                                                                                                        ':
                    # print(d.text)
                    text_list.append(d.text)

    print(text_list)
    print(ws(text_list))

def get_text(d):
    if 'unit' in d.attrib:
        return d.attrib['unit']
    elif 'Value' in d.tag and 'frame' in d.attrib:
        return d.attrib['frame'] + ' ' + d.text
    else:
        return d.text 
    
def get_value(d):
    if 'unit' in d.attrib:
        return float(d.text)

def read_knowledge():
    knowledge_path = '樹狀圖.xml'

    tree = ET.parse(knowledge_path)
    root = tree.getroot()

    text_dict = {}

    for r in root[3]:

        if 'WorkItem' in r.tag:            
            workItem = ''
            for d in r:
                workItem += get_text(d) + ' '

                # for dd in d.iter():
                #     print(dd.text.replace('\n', ''))
            text_dict[workItem.strip()] = get_value(d)

    return text_dict

def read_excel():
    excel_path = 'CQ881標LG09站地工數量-1100330更新.xls'
    
    excel = pd.read_excel(excel_path, sheet_name=None)
    target = 'Type S1'

    for k in list(excel.keys()):
        df = excel[k]

        for index, row in df.iterrows():
            # print(index, row.str())
            if target in row.values:
                print(k, row.values)

def get_budget(d):
    if 'language' in d.attrib and d.attrib == 'zh-TW':
        return d.text
    elif 'Value' in d.tag and 'frame' in d.attrib:
        return d.attrib['frame'] + ' ' + d.text
    else:
        return d.text 

def read_budget():
    xml_path = '(空白標單)臺北都會區大眾捷運系統萬大線(第二期工程)-CQ881標土建工程CQ881-11-04_bp_rbid.xml'

    tree = ET.parse(xml_path)
    root = tree.getroot()

    text_dict = {}
    description = ''
    quantity = ''
    unit = ''

    for r in root[1]:
        for d in r:


            if 'PayItem' in d.tag:
                description = ''
                quantity = ''
                unit = ''

                for dd in d:
                    if 'Description' in dd.tag and '連續壁' in dd.text and 'TYPE' in dd.text:
                        description = dd.text
                    elif 'Quantity' in dd.tag:
                        quantity = dd.text
                    elif 'Unit' in dd.tag:
                        unit = dd.text
                    else:
                        pass

                if description != '' and quantity != '' and unit != '':
                    print(f'{description}, {quantity}, {unit}')

            # print(d.text)

            # if '明挖覆蓋隧道工程' in d.text:
            #     print(d)
            #     for dd in r:
            #         print(dd.text)


            # print(workItem)

            # for dd in d.iter():
            #     print(dd.text.replace('\n', ''))
        # text_dict[workItem.strip()] = get_value(d)

    
    return text_dict

def read_amount_xml():
    xml_path = '樹狀圖_new.xml'

    tree = ET.parse(xml_path)
    root = tree.getroot()

    text_dict = {}
    description = ''
    quantity = ''
    unit = ''
    for r in root:
        if r.findtext('Description') == '數量計算書':
            print(r.tag)
            for d in r:
                print(d.tag)
                if d.findtext('Description') == '數量':
                    print(d.tag)


            # if 'PayItem' in d.tag:
            #     description = ''
            #     quantity = ''
            #     unit = ''


def get_word_embedding(text: str, model):
    cut = jieba.lcut(text)
    cut_item = []
    word_embeddings = []

    for c in cut:
        try: 
            word_embeddings.append(model.wv[c])
            cut_item.append(c)
        except:
            pass
        
    # print('original text: ', text_list[i])
    # print('word segmentation: ', cut)
    # print('ws with word embeddings: ', cut_item)
    
    return np.average(np.array(word_embeddings), axis=0)

def cosine_similarity(item0, item1):
    return np.dot(item0, item1) / (np.linalg.norm(item0) * np.linalg.norm(item1))

def calculate_word_emebedding(target: str, text_list: List[str]):
    model_path = 'word2vec.zh.300.model'
    model = Word2Vec.load(model_path, mmap='r')

    target_embed = get_word_embedding(target, model)

    cosine_similarity_list = []
    for i in range(len(text_list)):
        text = text_list[i]
        source_embed = get_word_embedding(text, model)
        cosine_similarity_list.append(cosine_similarity(target_embed, source_embed))

    return text_list[np.argmax(cosine_similarity_list)]

def main():
    compare_dict = {
        'Total': ['連續壁型式(編號) TYPE S1', '連續壁，(含導溝，厚100cm)，TYPE S1']
    }
    # read_xml()
    # read_excel()

    # read_budget()
    read_amount_xml()
    exit()

    # text_dict = read_knowledge()
    text_dict = read_budget()
    target_item = 'LG09車站雙牆系統連續壁，（含導溝，厚100cm），產品，預拌混凝土材料費，245kgf/cm2，第一型'
    target_item = '連續壁，(含導溝，厚100cm)，TYPE S1'
    target_item = '連續壁型式(編號) TYPE S1'
    target_value = 6713.6

    print(text_dict)

    text_list = list(text_dict.keys())
    source_item = calculate_word_emebedding(target_item, text_list)
    source_value = text_dict[source_item]

    print('target: ', target_item)
    print('target value: ', target_value)
    print('source: ', source_item)
    print('source value: ', source_value)
    print('same: ', target_value == source_value)


if __name__ == '__main__':
    main()