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

def extract_type(string):
    pattern = r'TYPE\s[IVX]+'
    result = re.findall(pattern, string)
    return result

#找工作表
def countsheet(excel, keywords): #keywords= ["連續壁", "排樁"]
    count = 0 #keyword數量
    count_blank = 0 #空打數量
    sheet_list = []
    
    for sheetName in excel.sheet_names():
        for keyword in keywords:
            if keyword in sheetName:
                sheet_list.append(sheetName)
                #print('sheetName為:', sheetName)
                count += 1
                break # 如果找到一個匹配的關鍵字，就跳過剩下的關鍵字檢查
        #elif sheetName.find(keyword)!=-1 and sheetName.find("空打")!=-1:
            #count_blank += 1
    #print('countsheet函式返回三個值:',count, count_blank, sheet_list) 
    print("sheet_list為",sheet_list)
    #4 0 ['LG10連續壁100cm(TYPE T1)', 'LG10連續壁100cm (空打)(TYPE T1A)', 'LG10連續壁80cm(TYPE T2)', 'LG10排樁80cm(TYPE T3)']
    return count, count_blank, sheet_list
    

        
def find_value(value, front, back):
    #print('find_value函式:',value.split(front)[-1].split(back)[0].strip())
    return value.split(front)[-1].split(back)[0].strip()

def read_excel(self, excelName, quantityFile, regulationFile, num_workItemType_design):
    print('抓取數量計算書')
    excel = xlrd.open_workbook(excelName)
    num_workItemType_quantity = 0
    list_workItemType_quantity = []
    list_workItemType_quantity_nested = []

    concrete_list = []
    concrete_type_list = []

    def counttype(sheetName, keywords): #keywords = ["Type", "80cmφ擋土排樁"]
        count = 0 #keyword數量
        type_list = []
        
        sheet = excel.sheet_by_name(sheetName)
        for row in range(8, sheet.nrows):
            cell_value = str(sheet.cell_value(row, 1))
            for keyword in keywords:
                if keyword in cell_value:
                    #t = cell_value.replace(' (空打)', '')
                    t = cell_value
                    # add space between type and type name
                    #if 'e ' not in t and 'Type' in t:
                    #    t = t.replace('e', 'e ')
                    #if 'Type' in t and 'Type ' not in t:
                    #    t = t.replace('Type', 'Type ')
                    t = t.split('-')[0]
                    type_list.append(t)
                    count += 1
                    break # 找到一個匹配的關鍵字後跳過其餘關鍵字
        return count, type_list
    #找工作表
    keywords= ["連續壁", "排樁"]
    count, count_blank, sheet_list = countsheet(excel, keywords)
    self.amount_type_list = [ find_value(s, '(', ')').replace('系列', '').replace('TYPE', 'Type') for s in sheet_list]
    
    #從工作表中找類型
    type_list_total = []
    keywords = ["Type","80cmφ擋土排樁"]
    for sheetName in sheet_list:
        count, type_list = counttype(sheetName, keywords)
        num_workItemType_quantity += count
        type_list_total.append(type_list)   
    
    print(type_list_total)

    list_workItemType_quantity_nested = type_list_total
    #使用 chain.from_iterable() 方法打平嵌套列表
    list_workItemType_quantity = list(itertools.chain.from_iterable(type_list_total))
    #list_workItemType_quantity = sheet_list
    print("num_workItemType_quantity為",num_workItemType_quantity) #4
    print("list_workItemType_quantity_nested為",list_workItemType_quantity_nested)
    #[['Type T1'], ['Type T1A'], ['Type T2'], ['80cmφ擋土排樁']]
    print("list_workItemType_quantity為",list_workItemType_quantity)
    # ['Type T1', 'TypeT1A (空打)', 'Type T2', '80cmφ擋土排樁']
    #num_workItemType_quantity = len(list_workItemType_quantity)

    for i in range(num_workItemType_quantity - 1):
        quantityFile.append( deepcopy(quantityFile[0]) )
    print("類型copy",i) #2
    
    #設定xml的TYPE名稱 
    list_workItemType_quantity = sheet_list
    print(sheet_list)
    print(num_workItemType_quantity)
    print(list_workItemType_quantity)

    for i in range(num_workItemType_quantity):
        element = quantityFile[i]
        try:
            element.set('TYPE', list_workItemType_quantity[i])
        except:
            pass
        #print(ET.tostring(element, encoding='unicode'))  # 打印設置 TYPE 屬性後的 XML 結構
    print(i)  # 顯示 XML 結構設置的類型

    #混凝土強度
    #找到所有需要寫入的節點
    strength_nodes = quantityFile.findall(".//DiaphragmWall/Concrete/Strength/Value")
    #print(f"Found {len(strength_nodes)} strength nodes")
    if len(strength_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(strength_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
        # 混凝土強度
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                tar_col = 0
                row_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

                for row in range(start_row, sheet.nrows):
                    if str(sheet.cell_value(row, 1)).find("水中混凝土") != -1 and row_first:
                        try:
                            if node_index >= len(strength_nodes):
                                #print("Warning: More values found than nodes available.")
                                break

                            value_Str = strength_nodes[node_index]
                            value_Str.text = str(sheet.cell_value(row, 1)).split(value_Str.get('unit'))[0]
                            start_row = row + 1
                            row_first = False
                            #print(f"Set value for node {node_index}: {value_Str.text}")
                            node_index += 1
                            break  
                        except Exception as e:
                            print(f"Error processing sheet {sheet_list[i]}: {e}")
                            continue

    #混凝土設計深度
    depth_nodes = quantityFile.findall(".//DiaphragmWall/Depth/Value")
    #print(f"Found {len(depth_nodes)} design_depth_nodes")
    if len(depth_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(depth_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
        # 設計深度
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                tar_col = 0
                row_first = True
                col_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

                for row in range(start_row, sheet.nrows):
                    for col in range(sheet.ncols):
                        if str(sheet.cell_value(row, 1)).find(list_workItemType_quantity_nested[i][0]) != -1 and row_first:
                            tar_row = row
                            row_first = False
                        if str(sheet.cell_value(row, col)).find("設計深度\n(m)") != -1 and col_first:
                            tar_col = col
                            col_first = False

                    if not row_first and not col_first:
                        break  

                if tar_row and tar_col:
                    try:
                        if node_index >= len(depth_nodes):
                            #print("Warning: More values found than nodes available.")
                            break

                        value_Depth = depth_nodes[node_index]
                        value_Depth.text = str(sheet.cell_value(tar_row, tar_col))
                        #print(f"Set depth value for node {node_index}: {value_Depth.text}")
                        node_index += 1
                    except Exception as e:
                        print(f"Error processing sheet {sheet_list[i]}: {e}")
                        continue

    #混凝土設計厚度
    thickness_nodes = quantityFile.findall(".//DiaphragmWall/Thickness/Value")
    if len(thickness_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(thickness_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
    #混凝土設計厚度
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                tar_col = 0
                row_first = True
                col_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

                for row in range(start_row, sheet.nrows):
                    for col in range(sheet.ncols):
                        if str(sheet.cell_value(row, 1)).find("水中混凝土") != -1 and row_first:
                            tar_row = row
                            row_first = False
                        if str(sheet.cell_value(row, col)).find("厚度\n(m)") != -1 and col_first:
                            tar_col = col
                            col_first = False
                    if not row_first and not col_first:
                        break

                try:
                    if node_index >= len(thickness_nodes):
                        #print("Warning: More values found than nodes available.")
                        break

                    value_Thk = thickness_nodes[node_index]
                    value_Thk.text = str(sheet.cell_value(tar_row, tar_col)).split(value_Thk.get('unit'))[0]
                    #print(f"Set value for node {node_index}: {value_Thk.text}")
                    node_index += 1
                except Exception as e:
                    print(f"Error processing sheet {sheet_list[i]}: {e}")
                    continue               
    # 混凝土總量
    # 找到所有需要寫入的節點
    total_nodes = quantityFile.findall(".//DiaphragmWall/Concrete/Total/Value")
    if len(total_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(total_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
    # 混凝土總量
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            tar_col = 0
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                row_first = True
                col_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

            for row in range(start_row, sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("Type") != -1 and row_first:
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("數量") != -1 and str(sheet.cell_value(row, col)).find("計價數量") == -1 and col_first:
                        tar_col = col
                        col_first = False
                if not row_first and not col_first:
                    break

            try:
                if node_index >= len(total_nodes):
                    #print("Warning: More values found than nodes available.")
                    break

                value_Total = total_nodes[node_index]
                value_Total.text = str(sheet.cell_value(tar_row, tar_col)).split(value_Total.get('unit'))[0]
                #print(f"Set value for node {node_index}: {value_Total.text}")
                node_index += 1
            except Exception as e:
                print(f"Error processing sheet {sheet_list[i]}: {e}")
                continue

    # 混凝土設計長度
    # 找到所有需要寫入的節點
    length_nodes = quantityFile.findall(".//DiaphragmWall/Length/Value")
    if len(length_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(length_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
    #混凝土設計長度
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            tar_col = 0
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                row_first = True
                col_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

            for row in range(start_row, sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("水中混凝土") != -1 and row_first:
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("設計長度\n(m)") != -1 and col_first:
                        tar_col = col
                        col_first = False
                if not row_first and not col_first:
                    break

            try:
                if node_index >= len(length_nodes):
                    #print("Warning: More values found than nodes available.")
                    break

                value_Length = length_nodes[node_index]
                value_Length.text = str(sheet.cell_value(tar_row, tar_col)).split(value_Length.get('unit'))[0]
                #print(f"Set value for node {node_index}: {value_Length.text}")
                node_index += 1
            except Exception as e:
                print(f"Error processing sheet {sheet_list[i]}: {e}")
                continue

    # U 板數量
    # 找到所有需要寫入的節點
    uboard_nodes = quantityFile.findall(".//DiaphragmWall/U_Board/Total/Value")
    if len(uboard_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(uboard_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
    # U 板數量
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            tar_col = 0
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                row_first = True
                col_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

            for row in range(start_row, sheet.nrows):
                for col in range(sheet.ncols):
                    if (str(sheet.cell_value(row, 1)).find("末端板") != -1 or str(sheet.cell_value(row, 1)).find("U型端板") != -1) and row_first:
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(8, col)).find("數量") != -1 and col_first:
                        tar_col = col
                        col_first = False
                if not row_first and not col_first:
                    break

            try:
                if node_index >= len(uboard_nodes):
                    #print("Warning: More values found than nodes available.")
                    break

                value_UBoard = uboard_nodes[node_index]
                value_UBoard.text = str(sheet.cell_value(tar_row, tar_col)).split(value_UBoard.get('unit'))[0]
                #print(f"Set value for node {node_index}: {value_UBoard.text}")
                node_index += 1
            except Exception as e:
                print(f"Error processing sheet {sheet_list[i]}: {e}")
                continue
    
    # 鋼筋籠數值 SD420W
    # 找到所有需要寫入的節點
    rebarcage_nodes = quantityFile.findall('.//DiaphragmWall/RebarWeight[@Description="SD420W鋼筋噸數"]/Value')
    if len(rebarcage_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(rebarcage_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
    # 鋼筋籠數值
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            tar_col = 0
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                row_first = True
                col_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

            for row in range(start_row, sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("SD420W") != -1 and row_first:
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(8, col)).find("數量") != -1 and col_first:
                        tar_col = col
                        col_first = False
                if not row_first and not col_first:
                    break

            try:
                if node_index >= len(rebarcage_nodes):
                    #print("Warning: More values found than nodes available.")
                    break

                value_RebarCage = rebarcage_nodes[node_index]
                value_RebarCage.text = str(sheet.cell_value(tar_row, tar_col)).split(value_RebarCage.get('unit'))[0]
                #print(f"Set value for node {node_index}: {value_RebarCage.text}")
                node_index += 1
            except Exception as e:
                print(f"Error processing sheet {sheet_list[i]}: {e}")
                continue


    rebarcage_nodes = quantityFile.findall('.//DiaphragmWall/RebarWeight[@Description="SD280W鋼筋噸數"]/Value')
    if len(rebarcage_nodes) < len(list_workItemType_quantity):
        print(f"Error: Not enough nodes to assign values. Found {len(rebarcage_nodes)} nodes, but need {len(list_workItemType_quantity)}.")
    else:
    # 鋼筋籠數值
        node_index = 0
        for i in range(len(list_workItemType_quantity_nested)):
            start_row = 8
            tar_col = 0
            for j in range(len(list_workItemType_quantity_nested[i])):
                sheet = excel.sheet_by_name(sheet_list[i])
                tar_row = 0
                row_first = True
                col_first = True
                #print(f"Processing sheet: {sheet_list[i]}")

            for row in range(start_row, sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("SD280W") != -1 and row_first:
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(8, col)).find("數量") != -1 and col_first:
                        tar_col = col
                        col_first = False
                if not row_first and not col_first:
                    break

            try:
                if node_index >= len(rebarcage_nodes):
                    #print("Warning: More values found than nodes available.")
                    break

                value_RebarCage = rebarcage_nodes[node_index]
                value_RebarCage.text = str(sheet.cell_value(tar_row, tar_col)).split(value_RebarCage.get('unit'))[0]
                #print(f"Set value for node {node_index}: {value_RebarCage.text}")
                node_index += 1
            except Exception as e:
                print(f"Error processing sheet {sheet_list[i]}: {e}")
                continue


    #排樁
    def Rowpilecounttype(sheetName, keyword):  #keyword = 80cmφ擋土排樁
        count = 0 #keyword數量
        type_list = []
        
        sheet = excel.sheet_by_name(sheetName)
        for row in range(8, sheet.nrows):
            if str(sheet.cell_value(row, 1)).find(keyword)!=-1:
                t = sheet.cell_value(row, 1)
                #print(t)
                # add spcae between type and type name
                #if 'e ' not in t:
                #    t = t.replace('e', 'e ')
                type_list.append(t)
                count += 1 
        
        #print('counttype函式:',count, 'type_list為:',type_list) #1  #['80cmφ擋土排樁']
        return count, type_list
    
    #選定目標sheet與設定XML(排樁)
    try:
        keywords= ["排樁"]
        count, count_blank, sheet_list = countsheet(excel, keywords)
        sheet = excel.sheet_by_name(sheet_list[0])
        print("sheet為：",sheet)
        depth_list = []
        for row in range(8, sheet.nrows):
            row_value = str(sheet.cell_value(row, 1))
            if row_value.find("排樁") != -1 and '80cm' in row_value:
                depth = row_value
                print(row_value)
                depth_list.append(depth)
                print('row_value為:',row_value)
                print('depth_list為:',depth_list)
    
        #設定XML TYPE (排樁)
        #for i in range(len(depth_list)):
        #    quantityFile.append( deepcopy(quantityFile[0]) )

        #for i in range(len(depth_list)):
        #    element = quantityFile[i]
        #    element.set('TYPE', depth_list[i])
        
        #排樁
        type_list_total = []
        #keywords = ["80cmφ擋土排樁"]
        for sheetName in sheet_list:
            count, type_list = Rowpilecounttype(sheetName,"80cmφ擋土排樁")
            num_workItemType_quantity += count
            type_list_total.append(type_list)

        type_list_total = [[i] for i in type_list_total[0]] 
    except:
        pass

    #print('type_list_total為:',type_list_total) #[['80cmφ擋土排樁']]

    list_workItemType_quantity_nested = type_list_total
    list_workItemType_quantity = list(itertools.chain.from_iterable(type_list_total))
    print('num_workItemType_quantity為:',num_workItemType_quantity) #1
    print('list_workItemType_quantity_nested為:',list_workItemType_quantity_nested) #[['80cmφ擋土排樁']]
    print('list_workItemType_quantity為:',list_workItemType_quantity) #['80cmφ擋土排樁']
    single_list = sheet_list
    nested_list = [[item] for item in single_list]
    print(nested_list)
    type_list_total = nested_list 
    Rowpile_list = type_list_total
    
    #排樁型式
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Rowpile/Type')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('排樁型式', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
        
        except Exception as error:
            print("排樁形式error",error)
    
    #支數
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Rowpile/Count')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('支數', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
        
        except Exception as error:
            print("支數error",error)
    
    #樁長(高度)
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Rowpile/Height')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('樁長', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
        
        except Exception as error:
            print("樁長error",error)

    #樁徑
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Rowpile/Diameter')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('樁徑', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
        
        except Exception as error:
            print("樁徑error",error)

    #排樁混凝土強度
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        found = False 
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Rowpile/Concrete/Strength')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                  if str(sheet.cell_value(row, 1)).find(node.get('Row')) != -1:
                      get_row = row
                      #print(f"Row found at: {get_row}")
                  if get_row :  
                      found = True
                      break   
                if found:
                   break  
            node.find('Value').text = str(sheet.cell_value(get_row, 1)).split(node.find('Value').get('unit'))[0]
        
        except Exception as error:
            print("排樁混凝土強度error",error)
    
    #繫梁混凝土強度
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        found = False 
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Beam/Concrete/Strength')
            for row in range(15,sheet.nrows):    #指定15列開始
                for col in range(sheet.ncols):
                  if str(sheet.cell_value(row, 1)).find(node.get('Row')) != -1:
                      get_row = row                   
                      #print(f"Row found at: {get_row}")
                  if get_row :  
                      found = True
                      break   
                if found:
                   break  
            node.find('Value').text = str(sheet.cell_value(get_row, 1)).split(node.find('Value').get('unit'))[0]

        except Exception as error:
            print("繫梁混凝土強度error",error)

    #排樁(鋼筋頓數)
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Rowpile/RebarWeight')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('鋼筋頓數', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))

        except Exception as error:
            print("鋼筋頓數error",error)
    
    #繫梁(鋼筋頓數）
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Beam/RebarWeight')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('鋼筋頓數', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
       
        except Exception as error:
            print("鋼筋頓數error",error)
    
    #模板面積
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Beam/formwork')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('模板面積', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
        
        except Exception as error:
            print("模板面積error",error)

    #排樁混凝土體積
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Rowpile/Concrete/Total')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('排樁混凝土體積', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
       
        except Exception as error:
            print("排樁混凝土體積error",error)

    #繫梁混凝土體積
    for i in range(len(Rowpile_list)):
        #sheet = excel.sheet_by_name(targetsheets[i])
        get_row = 0
        tar_col = 0
        found = False
        #type_value = group_list_middlecolumn[i][0]
        try:
            type_value = Rowpile_list[i][0]
            node = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Beam/Concrete/Total')
            for row in range(8,sheet.nrows):
                for col in range(sheet.ncols):
                 if str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                    get_row = row
                    #print(f"Row found at: {get_row}")
                 if str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                    tar_col = col
                    #print(f"Col found at: {tar_col}")
                 if get_row and tar_col:  
                    found = True
                    #print(f"Values will be set from Row: {get_row}, Col: {tar_col}")
                    break  
                if found:
                 break  
            #print(get_row, tar_col)
            #print('繫梁混凝土體積', str(sheet.cell_value(get_row, tar_col)))
            node.find('Value').text = str(sheet.cell_value(get_row, tar_col))
        
        except Exception as error:
            print("繫梁混凝土體積error",error)
    
    #支撐
    def findSheet(excel, target):
        targetsheet = []
        count = 0
        for sheet in excel.sheets():
            if sheet.name.find(target)!=-1 and sheet.name.find("混凝土版")==-1:
                targetsheet.append(sheet.name)
                # print(sheet.name, count)
            count += 1
        return targetsheet
    targetsheets = findSheet(excel, "支撐")
    #print('支撐目標工作表:',targetsheets) #['LG10支撐']
    
    def writeSupFenQuantity(support,total_arr,no_arr,num_arr,type_arr,unitWeight_arr,length_arr):
        for i in range(len(no_arr)-1):
            #print('len(no_arr)為:',len(no_arr))
           #support.append( deepcopy(support[0]) )
            support.insert(0, deepcopy(support[0]))
            i+=1
        for i in range(len(no_arr)):
            #support[i].find('Total/Value').text = str(format(total_arr[i], '.1f'))
            support[i].find('Layer/Value').text = str(no_arr[i])
            support[i].find('Type/Value').text = str(type_arr[i])
            support[i].find('Count/Value').text = str(num_arr[i])
            support[i].find('UnitWeight/Value').text = str(format(unitWeight_arr[i], '.3f'))
            support[i].find('Length/Value').text = str(format(length_arr[i], '.1f'))
            i+=1
    
    sheet = excel.sheet_by_name(targetsheets[0])
    #print(sheet) #Sheet  6:<G10支撐> 
    tar_row = []
    type_last = ''
    group_arr = []
    #node = quantityFile[i].find('SupportGroup/Support/Total')
    
    #支撐
    for row in range(8,sheet.nrows):
        if str(sheet.cell_value(row, 1)).find('支撐')!=-1:
        #print(sheet.cell_value(row, col))
            tar_row.append(row)
    #print(tar_row)
    tar_row = tar_row[:4]
    #print(tar_row)

    #if len(tar_row) != num_workItemType_quantity and self.is_pass == -1:
    #    self.is_pass = False
    #    return concrete_list, concrete_type_list
    
    #定義關鍵字列表
    keywords= ["連續壁", "排樁"]
    count, count_blank, sheet_list = countsheet(excel, keywords)
    self.amount_type_list = [ find_value(s, '(', ')').replace('系列', '').replace('TYPE', 'Type') for s in sheet_list]

    type_list_total = []
    num_middle_column_list = []
    keywords = ["Type","80cmφ擋土排樁"]
    for sheetName in sheet_list:
        #print(sheet_list)
        count, type_list = counttype(sheetName, keywords)
        #num_workItemType_quantity += count
        type_list_total.append(type_list) 
        num_middle_column_list.append(type_list)

    single_list = sheet_list
    nested_list = [[item] for item in single_list]
    #print(nested_list)
    type_list_total = nested_list
    group_arr = type_list_total #[['Type T1'], ['TypeT1A (空打)'], ['Type T2'], ['80cmφ擋土排樁']]
    count = 0
    
    for row in tar_row:
        #print(sheet.cell_value(row+1, col))
        total = []
        no = []
        num = []
        Type = []
        unitWeight = []
        length = []
        try:
            type_value = group_arr[count][0]
            #print(count, type_value)
            WorkItemType = group_arr[count][0]
            #print('WorkItemType', WorkItemType)
            for i in range(row+1,sheet.nrows):
                #print(i, sheet.cell_value(i,4), sheet.cell_value(i,5) != '')
                if sheet.cell_value(i,5) != '':
    #                 total.append(sheet.cell_value(i,3))
                    no.append(sheet.cell_value(i,4))
                    num.append(sheet.cell_value(i,5)[0])
                    Type.append(sheet.cell_value(i,5)[1:])
                    unitWeight.append(sheet.cell_value(i,6))
                    length.append(sheet.cell_value(i,7))
                else:
                    break
                    
            #print('no: ', no, num, Type,unitWeight, length)
            writeSupFenQuantity(quantityFile.find(f"./*[@TYPE='{type_value}']").find('SupportGroup'),total,no,num,Type,unitWeight,length)
            # if type_value != type_last:
            # writeSupFenQuantity(quantityFile.find(f"./*[@DEPTH='{type_value}']").find('SupportGroup'),total,no,num,Type,unitWeight,length)
            # print('type_value: ', type_value)
            # type_last = type_value
            count += 1
            # writeSupFenQuantity(designFile[count][3],no,num_2,type_2)
            # print(total)
            # print(no)
            # print(num)
            # print(Type)
            # print(unitWeight)
            # print(length)

        except Exception as error:
            print(error)
    
    #橫擋
    tar_row = []
    type_last = ''

    for row in range(8,sheet.nrows):
        if str(sheet.cell_value(row, 1)).find('橫擋')!=-1:
            tar_row.append(row)
    #print(tar_row)
    tar_row = tar_row[:4]
    #print(tar_row)
    count = 0
    for row in tar_row:
    #     print(sheet.cell_value(row+1, col))
        total = []
        no = []
        num = []
        Type = []
        unitWeight = []
        length = []
        try:
            type_value = group_arr[count][0]
            #print(count, type_value)
            WorkItemType = group_arr[count][0]
            # print('WorkItemType', WorkItemType)
            for i in range(row+1,sheet.nrows):
                #print(i, sheet.cell_value(i,4), sheet.cell_value(i,5) != '')
                if sheet.cell_value(i,5) != '':
    #                 total.append(sheet.cell_value(i,3))
                    no.append(sheet.cell_value(i,4))
                    num.append(sheet.cell_value(i,5)[0])
                    Type.append(sheet.cell_value(i,5)[1:])
                    unitWeight.append(sheet.cell_value(i,6))
                    length.append(sheet.cell_value(i,7))
                else:
                    break
                    
            #print('no: ', no, num, Type)
            writeSupFenQuantity(quantityFile.find(f"./*[@TYPE='{type_value}']").find('FenceGroup'),total,no,num,Type,unitWeight,length)
            # if type_value != type_last:
            # writeSupFenQuantity(quantityFile.find(f"./*[@DEPTH='{type_value}']").find('FenceGroup'),total,no,num,Type,unitWeight,length)
            # type_last = type_value
            count += 1
            # writeSupFenQuantity(designFile[count][3],no,num_2,type_2)
            # print(total)
            # print(no)
            # print(num)
            # print(Type)
            # print(unitWeight)
            # print(length)
        except Exception as error:
            print(error)
    
    #總計(總計包含TYEP T1~T3)寫在第一個workitemtype
    tar_row = 0
    row_first = True
    value = quantityFile[0].find('Total_SupFen/Value')
    for row in range(8,sheet.nrows):
        if str(sheet.cell_value(row, 1)).find("總 計")!=-1 and row_first:
            #print(sheet.cell_value(row, col))
            tar_row = row
            row_first = False
    #print(str(sheet.cell_value(tar_row, 3)))
    value.text = str(sheet.cell_value(tar_row, 3))
    
    #中間柱工作表
    def findallSheet(excel, target1, target2):
        middle_column_list = []
        count = 0
        for sheet in excel.sheets():
            if target1 in sheet.name and target2 in sheet.name:
                middle_column_list.append(sheet.name)
                #print("有中間柱和TYPE字樣的工作表", sheet.name, count)
            count += 1
        return middle_column_list

    middle_column_list = findallSheet(excel, "中間柱", "TYPE")
    #print('全部中間柱和TYPE目標工作表:', middle_column_list)

    #middleColumnGroup = quantityFile.find(f"./*[@TYPE='{type_list[0]}']").find('MiddleColumnGroup')
    #print(type_list)
    #for i in range(len(middle_column_list)-1):
    #    middleColumnGroup.insert(0, deepcopy(middleColumnGroup[0]))
    #    print("中間柱copy的i",i)
    
    #group_list_middlecolumn = np.reshape(np.arange(0,len(targetsheets)), (-1,1)).tolist() 
    #print('group_list_middlecolumn為:',group_list_middlecolumn) 
    #if len(targetsheets) == num_workItemType_design and self.is_pass == -1:  
    #    self.is_pass = False
    #    return concrete_list, concrete_type_list, list_workItemType_quantity

    group_list_middlecolumn = type_list_total 
    #sheet = excel.sheet_by_name(middle_column_list[i]) 
    #print(middle_column_list)
    
    # 中間柱長度
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/Length")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/Length")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #開挖深度
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/Depth")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/Depth")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")

    
    #型鋼尺寸
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/Steel/Type")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/Steel/Type")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    
    #型鋼長度
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/Steel/Length")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/Steel/Length")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    row_found = False
                    col_found = False
                    for col in range(sheet.ncols):
                        cell_content = str(sheet.cell_value(row, col))
                        if cell_content.find(node.get('Row')) != -1:
                            get_row = row
                            row_found = True
                        if cell_content.startswith("長度") and "(m)" in cell_content:
                            tar_col = col
                            col_found = True

                        if row_found and col_found:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #開挖面以上重    
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/Steel/TotalUpper")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/Steel/TotalUpper")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #埋入鑽掘樁重
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/Steel/TotalLower")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/Steel/TotalLower")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #鑽掘樁直徑
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Diameter")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Diameter")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(int(re.sub('[^0-9.]', '', sheet.cell_value(get_row, tar_col))) / 10)
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #樁身埋入深度
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Length")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Length")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #中間柱混凝土強度
    middle_column_nodes = []
    for type_value in sheet_list :
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Concrete/Strength")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Concrete/Strength")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if get_row:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    node.find('Value').text = str(sheet.cell_value(get_row, 1)).split(node.find('Value').get('unit'))[0]
                    #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {node.find('Value').text}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #支數
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Count")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Count")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #鋼筋籠總重
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/RebarCage/Total")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/RebarCage/Total")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    
    #回填材料
    middle_column_nodes = []
    for type_value in sheet_list:
        nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{type_value}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Backfill/Total")
        middle_column_nodes.extend(nodes)

    # 確認節點數量與需要寫入的數值數量是否匹配
    if len(middle_column_nodes) < len(middle_column_list):
        print(f"Error: Not enough nodes to assign values. Found {len(middle_column_nodes)} nodes, but need {len(middle_column_list)}.")
    else:
        # 依序將數值寫入對應的節點
        used_types = set()
        for sheet_name in middle_column_list:
            # 抓取對應的 TYPE
            matched_type = None
            for type_value in sheet_list:
                # 比對 TYPE 和 TYPE 後面的字母及數字，去除空格後進行比對
                try:
                    type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                try:
                    sheet_suffix = sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "")
                except IndexError:
                    continue

                if type_suffix == sheet_suffix:
                    matched_type = type_value
                    break
            if matched_type is None:
                print(f"No matching TYPE found for sheet {sheet_name}")
                continue

            # 確認 TYPE 尚未被使用
            if matched_type in used_types:
                print(f"Type '{matched_type}' already used, skipping.")
                continue
            used_types.add(matched_type)

            # 抓取對應的節點
            nodes = quantityFile.findall(f".//WorkItemType[@TYPE='{matched_type}']/MiddleColumnGroup/MiddleColumn/DrilledPile/Backfill/Total")
            if not nodes:
                print(f"Error: No nodes found for TYPE '{matched_type}'")
                continue

            sheet = excel.sheet_by_name(sheet_name)
            get_row = 0
            tar_col = 0
            found = False
            try:
                node = nodes[0]
                for row in range(8, sheet.nrows):
                    for col in range(sheet.ncols):
                        if node.get('Row') and str(sheet.cell_value(row, col)).find(node.get('Row')) != -1:
                            get_row = row
                        if node.get('Col') and str(sheet.cell_value(row, col)).find(node.get('Col')) != -1:
                            tar_col = col
                        if get_row and tar_col:
                            found = True
                            break
                    if found:
                        break
                if node is not None:
                    value_node = node.find('Value')
                    if value_node is not None:
                        value_node.text = str(sheet.cell_value(get_row, tar_col))
                        #print(f"Set value for sheet {sheet_name} at node for TYPE '{matched_type}': {value_node.text}")
                    else:
                        print(f"No 'Value' node found for node: {node}")
                else:
                    print(f"Node for sheet {sheet_name} is None.")
            except Exception as error:
                print(f"Error processing sheet {sheet_name}: {error}")

    # 檢查 sheet_list 中沒有對應的 TYPE
    for type_value in sheet_list:
        try:
            type_suffix = type_value.split('(TYPE ')[1].split(')')[0].replace(" ", "")
        except IndexError:
            continue

        if all(type_suffix not in sheet_name.split('(TYPE ')[1].split(')')[0].replace(" ", "") for sheet_name in middle_column_list):
            print(f"Skipping TYPE '{type_value}' as there is no corresponding middle column sheet.")
    

    """
    #regulation
    for i in range(len(depth_list) - 1):
        regulationFile.append( deepcopy(regulationFile[0]) )

    for i in range(len(depth_list)):
        element = regulationFile[i]
        element.set('DEPTH', depth_list[i])
    
    
    
    for i in range(len(list_workItemType_quantity_nested)):
        start_row = 8
        tar_col = 0
        for j in range(len(list_workItemType_quantity_nested[i])):
            sheet = excel.sheet_by_name(sheet_list[i])
            tar_row = 0
            row_first = True
            type_value = list_workItemType_quantity_nested[i][j]
            value = quantityFile.find(f"./*[@TYPE='{type_value}']").find('Concrete/Total/Value')
            for row in range(start_row,sheet.nrows):
                for col in range(sheet.ncols):
                    if str(sheet.cell_value(row, 1)).find("Type")!=-1 and row_first:
                        #print(sheet.cell_value(row, col))
                        tar_row = row
                        row_first = False
                    if str(sheet.cell_value(row, col)).find("數量")!=-1 and str(sheet.cell_value(row, col)).find("計價數量")!=0:
                        #print(sheet.cell_value(i, j))
                        tar_col = col
            value.text = str(sheet.cell_value(tar_row, tar_col))
            start_row = tar_row+1

    
    """
    print('數量計算書抓取完成')
    return concrete_list, concrete_type_list, sheet_list