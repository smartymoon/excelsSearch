# 先写 API， 再写 GUI, 可以设置默认路径
import os
import xlrd
from glob import glob
from tkinter import *
from tkinter import messagebox, ttk
from tkinter.filedialog import askdirectory
import threading


def searchInExcel(path, word):
    print(path)
    try:
        file = xlrd.open_workbook(path)
        for sheet in file.sheets():
            for row_index, row in enumerate(sheet.get_rows()):
                # 此处不确定特殊格式是否会影响内容
                for col_index, cell in enumerate(row):
                    value = cell.value
                    if value and word in str(value):
                        result.append({
                            'file': path
                        })
                        file.release_resources()
                        return
    except Exception:
        # 记录不能打开的文件
        failResult.append(path)


def findInFolder(path):
    if not directory.get() or not keyword.get():
        messagebox.showinfo("小笨蛋", "路径和关键词写清了么？")
        return


    for item in glob(path + '/*'):
        if (item.endswith('.xls') or item.endswith('.xlsx')) and (not os.path.basename(item).startswith('~')):
            t = threading.Thread(target=searchInExcel, args=(item, keyword.get()))
            pools.append(t)
        elif os.path.isdir(item):
            findInFolder(item)

def selectPath():
    path_ = askdirectory()
    directory.set(path_)


def work():
    global pools, result, failResult
    pools = []
    result = []
    failResult = []
    findInFolder(directory.get())

    for row in tree.get_children():
        tree.delete(row)

    for row in failTree.get_children():
        failTree.delete(row)

    for t in pools:
        t.start()

    for t in pools:
        t.join()

    if not result:
        messagebox.showinfo('结果', '没找到关键词')
    else:
        for index, item in enumerate(result):
            print('aa', item['file'])
            tree.insert('', index, values=(item['file'],))

    for index, item in enumerate(failResult):
        failTree.insert('', index, values=(item,))


def openFile(event):
    item = tree.selection()[0]
    item_text = tree.item(item, "values")
    os.startfile(item_text[0])  # 输出所选行的第一列的值


root = Tk()
pools = []
root.title('多Excel查找,双击查找内容可打开文件')
directory = StringVar()
keyword = StringVar()
result = []
failIndex = 0
failResult = []

w, h = root.maxsize()
root.geometry("{}x{}".format(w, h))

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

tree = ttk.Treeview(root, columns=('file',), show="headings", height=15)
tree.column('file')
tree.heading('file', text='文件')
tree.bind('<Double-1>', openFile)
tree.pack(fill='x')

failTree = ttk.Treeview(root, columns=('file',), show="headings", height=15)
failTree.column('file')
failTree.heading('file', text='未能打开的文件')
failTree.pack(fill='x')

root.mainloop()
