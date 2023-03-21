import tkinter as tk
from tkinter import filedialog
from wordpreprocess im

root = tk.Tk()
root.title('文件比對程式')
root.geometry('400x300+300+300')
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
    
def compare():
    wordName = wordName_select.file_path
    excelName = excelName_select.file_path
    schemaName = schemaName_select.file_path
    budget_path = budget_path_select.file_path
    output_path = output_path_select.file_path
    treeName = 'tree.xml'



    print('compare complete')

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
compare_button.grid(row=5,column=0)

root.mainloop()