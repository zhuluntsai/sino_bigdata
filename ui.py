import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.title('文件比對程式')
root.geometry('400x300+300+300')
root.grid_columnconfigure(0, weight=1)

# https://stackoverflow.com/questions/73658643/bind-button-open-file-dialog-to-a-text-field-in-tkinter
class file_select(tk.Frame):
    def __init__(self, parent, fileDescription='', ext=''):
        tk.Frame.__init__(self, parent)
    
        self.file_path = tk.StringVar()
        self.ext = ext
        self.label = tk.Label(self, text=fileDescription, width=10)
        self.label.grid(row=0, column=0)

        self.textbox = tk.Entry(self, textvariable=self.file_path, width=25, foreground='grey')
        self.textbox.insert(0, f'請選擇 .{self.ext} 檔')
        self.textbox.grid(row=0, column=1)

        self.button = tk.Button(self, text="瀏覽", command=self.setFilePath, width=6)
        self.button.grid(row=0, column=2, padx=10)

    def setFilePath(self):
        file_selected = filedialog.askopenfilename(parent=root, filetypes=(('', f'*.{self.ext}'), ))
        self.file_path.set(file_selected)
        self.textbox.configure(foreground='black')

    @property
    def folder_path(self):
        return self.file_path.get()
    
class folder_select(file_select):
    def __init__(self, parent, folderDescription):
        file_select.__init__(self, parent, folderDescription)

        self.textbox.delete(0, 'end')
        self.textbox.insert(0, f'請選擇資料夾')
    
    def setFolderPath(self):
        file_selected = filedialog.askdirectory()
        self.file_path.set(file_selected)
        self.textbox.configure(foreground='black')
    
def compare():
    wordName = wordName_select.folder_path
    print('compare')

wordName_select = file_select(root, '設計計算書', 'docx')
wordName_select.grid(row=0, pady=5)

excelName_select = file_select(root, '數量計算書', 'xls')
excelName_select.grid(row=1, pady=2)

schemaName_select = file_select(root, 'Schema', 'xml')
schemaName_select.grid(row=2, pady=2)

budget_path_select = file_select(root, '預算書', 'xml')
budget_path_select.grid(row=3, pady=2)

budget_path_select = folder_select(root, '輸出資料夾')
budget_path_select.grid(row=4, pady=2)



treeName = 'tree.xml'
output_path = 'report.csv'

root.mainloop()