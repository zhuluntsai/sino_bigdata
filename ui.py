import tkinter as tk
import tkinter.ttk as ttk

from tkinter import filedialog
import os

root = tk.Tk()
root.title('文件比對程式')
root.geometry('+300+300')
root.grid_columnconfigure(0, weight=1)

# https://stackoverflow.com/questions/73658643/bind-button-open-file-dialog-to-a-text-field-in-tkinter
class fileSelect(tk.Frame):
    def __init__(self, parent, fileDescription='', ext=''):
        tk.Frame.__init__(self, parent)
    
        self.filePath = tk.StringVar()
        self.ext = ext
        self.label = tk.Label(self, text=fileDescription, width=10)
        self.label.grid(row=0, column=0)

        self.textbox = tk.Entry(self, textvariable=self.filePath, width=25, foreground='grey')
        self.textbox.insert(0, f'請選擇 .{self.ext} 檔')
        self.textbox.grid(row=0, column=1)

        self.button = tk.Button(self, text="瀏覽", command=self.setFilePath, width=6)
        self.button.grid(row=0, column=2, padx=10)

    def setFilePath(self):
        file_selected = filedialog.askopenfilename(parent=root, filetypes=(('', f'*.{self.ext}'), ))
        self.filePath.set(file_selected)
        self.textbox.configure(foreground='black')

    @property
    def file_path(self):
        return self.filePath.get()
    
class fileSaveAs(fileSelect):
    def __init__(self, parent, folderDescription, ext):
        fileSelect.__init__(self, parent, folderDescription, ext)

        self.textbox.delete(0, 'end')
        self.textbox.insert(0, f'請輸入比對報告檔名')
    
    def setFilePath(self):
        file_selected = filedialog.asksaveasfilename(filetypes=(('', f'*.{self.ext}'), ))
        self.filePath.set(file_selected)
        self.textbox.configure(foreground='black')
    
class typeSelect(tk.Frame):
    def __init__(self, parent, label, type_list):
        tk.Frame.__init__(self, parent)
    
        self.label = tk.Label(self, text=label, width=10)
        self.label.grid(row=0, column=0)

        self.combobox = ttk.Combobox(self, values=type_list)
        self.combobox.grid(row=0, column=1)

    @property
    def combo_box(self):
        return self.combobox.get()
    
class typeMultipleSelect(tk.Frame):
    def __init__(self, parent, amount_type_list, middle_type_list):
        tk.Frame.__init__(self, parent)
    
        self.label = []
        self.listbox = []

        self.label.append(tk.Label(self, text='數量計算書'))
        self.label[0].grid(row=0, column=0)

        self.listbox.append(tk.Label(self, text='中間柱、支撐'))
        self.listbox[0].grid(row=1, column=0)

        for i, amount in enumerate(amount_type_list, 1):
            self.label.append(tk.Label(self, text=amount))
            self.label[i].grid(row=0, column=i)

            self.listbox.append(tk.Listbox(self, height=8, width=10, selectmode=tk.MULTIPLE, exportselection=False))
            self.listbox[i].grid(row=1, column=i)
            for middle in middle_type_list:
                self.listbox[i].insert(tk.END, middle)

    @property
    def combo_box(self):
        return self.combobox.get()

pass_list = []
def compare():
    wordName = wordName_select.file_path
    excelName = excelName_select.file_path
    schemaName = schemaName_select.file_path
    budget_path = budget_path_select.file_path
    output_path = output_path_select.file_path
    treeName = 'tree.xml'

    wordName = 'word-preprocess/data/LG09站地下擋土壁及支撐系統20221212圍囹正確版_修改換行符.docx'
    excelName = 'word-preprocess/data/CQ881標LG09站地工數量-1111129開發用.xls'
    schemaName = 'word-preprocess/data/schema.xml'
    budget_path = 'word-preprocess/data/CQ881標土建工程CQ881-11-04_bp_rbid.xml'
    output_path = 'report.csv'
    treeName = 'tree.xml'
    
    # os.system(f'python word2xml.py --word_path {wordName} --excel_path {excelName} --schema_path {schemaName} --budget_path {budget_path} --output_path {output_path} --tree_path {treeName}')

    amount_type_list = ['TYPE S1','TYPE S2','TYPE S3']
    middle_type_list = ['TYPE S1','TYPE S3']
    # [[0, 1], [2]]

    # amount_type_list = ['TYPE T1','TYPE T1A','TYPE T2']
    # middle_type_list = ['中間柱1左','中間柱1中','中間柱1右','中間柱2', '中間柱3']
    # [[0], [0], [0], [2]]

    print(f'比對報告已儲存在 {output_path}') 
    print(len(pass_list))

    # if amount of word and excel doesn't match, add compare button
    if len(pass_list) == 0:
        global type_multiple_select
        type_multiple_select = typeMultipleSelect(root, amount_type_list=amount_type_list, middle_type_list=middle_type_list)
        type_multiple_select.grid(row=6, pady=5)
        compare_button.grid(row=7, pady=5, ipadx=50)
        root.update()

        pass_list.append(1)
    
    # compare with group array
    else:
        selection = []
        group_array = []
        for i, listbox in enumerate(type_multiple_select.listbox[1:]):
            select = listbox.curselection()
            print(select)
            for s in select:
                group_array.append(i)

        print(group_array)

wordName_select = fileSelect(root, '設計計算書', 'docx')
wordName_select.grid(row=0, pady=5)

excelName_select = fileSelect(root, '數量計算書', 'xls')
excelName_select.grid(row=1, pady=2)

schemaName_select = fileSelect(root, 'Schema', 'xml')
schemaName_select.grid(row=2, pady=2)

budget_path_select = fileSelect(root, '預算書', 'xml')
budget_path_select.grid(row=3, pady=2)

output_path_select = fileSaveAs(root, '輸出路徑', 'csv')
output_path_select.grid(row=4, pady=2)

compare_button = tk.Button(root, text="文件比對", command=compare)
compare_button.grid(row=5, pady=10, ipadx=50)

box_list = []

root.mainloop()