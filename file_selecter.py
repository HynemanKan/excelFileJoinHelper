import tkinter as tk
import tkinter.filedialog as fd

class FileSelecter:
    def __init__(self,parent,frameLabel="",entryWidth=200):
        if frameLabel!="":
            self.frm = tk.LabelFrame(parent,text=frameLabel)
        else:
            self.frm = tk.Frame(parent)
        self.filePath = tk.StringVar()
        self.filePath.set("请选择文件")
        self.filePathEntry = tk.Entry(self.frm,textvariable=self.filePath,width=entryWidth)
        self.showDialogButton = tk.Button(self.frm,text="打开",command=self.showDialog)
        self.filePathEntry.pack(side=tk.LEFT)
        self.showDialogButton.pack(side=tk.LEFT)
        self.frm.pack()

    def showDialog(self):
        filename = fd.askopenfilename()
        if filename != '':
            self.filePath.set(filename)
        else:
            self.filePath.set("请选择文件")
    def getFilePath(self):
        return self.filePath.get()

    def disableChange(self):
        self.showDialogButton.setvar("state",tk.DISABLED)


if __name__ == '__main__':
    top = tk.Tk()
    test = FileSelecter(top,frameLabel="原始素材")
    top.mainloop()