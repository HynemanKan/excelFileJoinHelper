import os
import tkinter as tk
from file_selecter import FileSelecter
import tkinter.messagebox
import core
info = """注意：
  编辑所有需要对齐的Excel表格文件，使文档第一行为列名(不可有合并行/列)，将基准列放在第一列。
  基准列数据不应有重复，推荐使用学号。若有重复请按程序提示去重。
  表格请另存为最新的Excel2007格式（后缀名.xlsx）
  纯数字的文本内容必须转换为文本格式，否则生成文件或不可用
  如需对比其他文件请重新打开本软件
"""

class GUI:
    def __init__(self):
        self.top = tk.Tk()
        self.input_part = tk.Frame(self.top)
        self.bigFile_selecter = FileSelecter(self.input_part,entryWidth=65,frameLabel="主表格")
        self.bl1 = tk.Label(self.input_part, text=" ")
        self.bl1.pack()
        self.smallFile_selecter = FileSelecter(self.input_part, entryWidth=65, frameLabel="目标表格")
        self.input_part.pack()
        self.infoLabel = tk.Label(self.top,text=info,fg="red",relief=tk.GROOVE,justify=tk.LEFT)
        self.infoLabel.pack()
        self.btn_part = tk.Frame(self.top)
        self.btn_A2B = tk.Button(self.btn_part,text="目标表信息补全",command=self.btn_A2B_fun)
        self.btn_A2B.pack()
        self.btn_B2A = tk.Button(self.btn_part, text="主表格补全（缺失信息跳过）",command=self.btn_B2A_fun)
        self.btn_B2A.pack()
        self.btn_miss = tk.Button(self.btn_part, text="生成缺失名单",command=self.btn_miss_fun)
        self.btn_miss.pack()
        self.btn_part.pack()
        self.gen = core.ExcelJoin()
        self.fileNotLoad=True
        self.top.mainloop()

    def btn_A2B_fun(self):
        self.btn_A2B.setvar("state",tk.DISABLED)
        fileA = self.bigFile_selecter.getFilePath()
        fileB = self.smallFile_selecter.getFilePath()
        if os.path.exists(fileA) and os.path.exists(fileB):
            if self.fileNotLoad:
                self.gen.lodaXlsxFile(fileA,fileB)
                self.fileNotLoad=False
                self.bigFile_selecter.disableChange()
                self.smallFile_selecter.disableChange()
            self.gen.joinA2B()
            tkinter.messagebox.showinfo("success","已在当前文件将生成")
        else:
            tkinter.messagebox.showerror("error","请提供正确文件地址")
        self.btn_A2B.setvar("state", tk.NORMAL)


    def btn_B2A_fun(self):
        self.btn_B2A.setvar("state",tk.DISABLED)
        fileA = self.bigFile_selecter.getFilePath()
        fileB = self.smallFile_selecter.getFilePath()
        if os.path.exists(fileA) and os.path.exists(fileB):
            if self.fileNotLoad:
                self.gen.lodaXlsxFile(fileA,fileB)
                self.fileNotLoad=False
                self.bigFile_selecter.disableChange()
                self.smallFile_selecter.disableChange()
            self.gen.joinB2A()
            tkinter.messagebox.showinfo("success","已在当前文件将生成")
        else:
            tkinter.messagebox.showerror("error","请提供正确文件地址")
        self.btn_B2A.setvar("state", tk.NORMAL)


    def btn_miss_fun(self):
        self.btn_miss.setvar("state",tk.DISABLED)
        fileA = self.bigFile_selecter.getFilePath()
        fileB = self.smallFile_selecter.getFilePath()
        if os.path.exists(fileA) and os.path.exists(fileB):
            if self.fileNotLoad:
                self.gen.lodaXlsxFile(fileA,fileB)
                self.fileNotLoad=False
                self.bigFile_selecter.disableChange()
                self.smallFile_selecter.disableChange()
            self.gen.findANotInB()
            tkinter.messagebox.showinfo("success","已在当前文件将生成")
        else:
            tkinter.messagebox.showerror("error","请提供正确文件地址")
        self.btn_miss.setvar("state", tk.NORMAL)


    def mainloop(self):
        self.top.mainloop()

if __name__ == '__main__':
    gui = GUI()
    gui.mainloop()