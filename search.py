# 先写 API， 再写 GUI, 可以设置默认路径
import os
import xlrd
from glob import glob
from tkinter import *
from tkinter import messagebox, ttk
from tkinter.filedialog import askdirectory


def searchInExcel(path):
    file = xlrd.open_workbook(path)
    for sheet in file.sheets():
        for row_index, row in enumerate(sheet.get_rows()):
            # 此处不确定特殊格式是否会影响内容
            for col_index, cell in enumerate(row):
                value = cell.value
                if value and keyword.get() in str(value):
                    position = '{},{}'.format(row_index + 1, chr(col_index + 65))
                    result.append({
                        'file': path,
                        'value': value,
                        'sheet': sheet.name,
                        'cell': position,
                    })


def findInFolder(path):
    if not directory.get() or not keyword.get():
        messagebox.showinfo("小笨蛋", "路径和关键词写清了么？")
        return

    for row in tree.get_children():
        tree.delete(row)

    for item in glob(path + '/*'):
        if item.endswith('.xls') or item.endswith('.xlsx'):
            searchInExcel(item)
        elif os.path.isdir(item):
            findInFolder(item)


def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    size = '%dx%d+%d+%d' % (width, height, (screenwidth - width) / 2, (screenheight - height) / 2)
    root.geometry(size)


def selectPath():
    path_ = askdirectory()
    directory.set(path_)


def work():
    findInFolder(directory.get())
    if not result:
        messagebox.showinfo('结果', '没找到关键词')
    else:
        for index, item in enumerate(result):
            tree.insert('', index, values=(item['file'], item['value'], item['sheet'], item['cell']))

def openFile(event):
        item = tree.selection()[0]
        item_text = tree.item(item, "values")
        os.startfile(item_text[0])  # 输出所选行的第一列的值


root = Tk()
root.title('小黄专用多Excel查找,双击查找内容可打开文件')
directory = StringVar()
keyword = StringVar()
result = []

center_window(root, 600, 340)

fram = Frame(root)
fram.pack()

fram_1 = Frame(fram)
fram_1.pack(pady=3)
fram_2 = Frame(fram)
fram_2.pack(pady=3)

Label(fram_1, text="路径", width=5).pack(side=LEFT)
Entry(fram_1, textvariable=directory, state="readonly").pack(side=LEFT)
Button(fram_1, text="路径", command=selectPath).pack(side=LEFT, padx=10)

Label(fram_2, text="关键字", width=5).pack(side=LEFT)
Entry(fram_2, textvariable=keyword).pack(side=LEFT)
Button(fram_2, text="查找", command=work).pack(side=LEFT, padx=10)


tree = ttk.Treeview(root, columns=('file', 'value', 'sheet', 'cell'), show="headings")
tree.column('file')
tree.column('value', width=200)
tree.column('sheet', width=80)
tree.column('cell', width=80)

tree.heading('file', text='文件')
tree.heading('value', text='内容')
tree.heading('sheet',  text='sheet')
tree.heading('cell',  text='单元格')
tree.bind('<Double-1>', openFile)
tree.pack(fill='x')

root.mainloop()
