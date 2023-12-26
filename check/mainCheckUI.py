import subprocess
import tkinter as tk
from tkinter import filedialog

import xlrd


from check.mainCheck import getResStr


def open_program_1():
    # 启动第二个程序
    subprocess.Popen(["python", "investCheckUI.py"])
    root.destroy()

# 定义一个Python函数，选择文件并将文件路径显示在小文本框中
def choose_file(entry_file_path):
    file_path = filedialog.askopenfilename()  # 弹出文件选择对话框，获取所选文件的路径
    if file_path:  # 如果用户选择了文件
        entry_file_path.delete(0, tk.END)  # 清空小文本框中的内容
        entry_file_path.insert(0, file_path)  # 在小文本框中显示所选文件的文件路径

# 定义一个Python函数，获取小文本框中的文件路径，并在大文本框中展示完整路径
def display_investCheck( text):
    jsb_path = entry_file_path_1.get()  # 获取小文本框中的文件路径
    design_path=entry_file_path_2.get()
    if jsb_path and design_path:  # 如果小文本框不为空
        res_Str=getResStr(jsb_path,design_path)
        text.insert(tk.END, f"{res_Str}\n")
    else:
        text.insert(tk.END, "未选择任何文件\n")

def display_itemCheck(entry_file_path, text):
    file_path = entry_file_path.get()  # 获取小文本框中的文件路径
    if file_path:  # 如果小文本框不为空
        wb = xlrd.open_workbook(file_path)
        sheet = wb.sheet_by_name("Sheet1")
        # 总投资
        investAll = float(sheet.row(0)[1].value)
        text.insert(tk.END, f"文件完整路径为：{investAll}\n")
    else:
        text.insert(tk.END, "未选择任何文件\n")

# 创建一个Tkinter应用程序窗口
root = tk.Tk()
root.title("审核")  # 设置窗口标题


# 添加第一行选择文件的控件
label_file_path_1 = tk.Label(root, text="文件1：")
label_file_path_1.grid(row=0, column=0)

entry_file_path_1 = tk.Entry(root)
entry_file_path_1.grid(row=0, column=1)

choose_file_btn_1 = tk.Button(root, text='选择文件1', command=lambda: choose_file(entry_file_path_1))
choose_file_btn_1.grid(row=0, column=2)

display_btn_1 = tk.Button(root, text='前往检查',command=open_program_1)
display_btn_1.grid(row=0, column=3)
#
# 添加第二行选择文件的控件
label_file_path_2 = tk.Label(root, text="文件2：")
label_file_path_2.grid(row=1, column=0)

entry_file_path_2 = tk.Entry(root)
entry_file_path_2.grid(row=1, column=1)

choose_file_btn_2 = tk.Button(root, text='选择文件2', command=lambda: choose_file(entry_file_path_2))
choose_file_btn_2.grid(row=1, column=2)

display_btn_2 = tk.Button(root, text='点击检查', command=lambda: display_investCheck(text))
display_btn_2.grid(row=1, column=3)

# 添加一个大文本框控件
text = tk.Text(root)
text.grid(row=2, columnspan=4)

# 运行窗口程序
# root.mainloop()
# 运行窗口程序
def start():
    root.mainloop()

