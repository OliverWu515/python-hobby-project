import os
import re
import tkinter
from tkinter import Tk, Label, Button, StringVar, Entry, Frame, scrolledtext, messagebox, filedialog
import xlrd

re_stuno = re.compile(r"[0-9A-Za-z]{9,12}")
name_inf = {}


def read_nameinf():
    clearall()
    global name_inf
    insert_message('表格要求：第一列是学号，第二列是姓名，学号姓名从第二行开始填写。\n')
    namelist_path = filedialog.askopenfilename(filetypes=[("xls表格", ".xls")], title='打开存有班级学号-姓名的xls表格')
    try:
        wb = xlrd.open_workbook(namelist_path)
        table = wb.sheet_by_index(0)
        nrows = table.nrows  # 获取表格行数
        for i in range(1, nrows):
            stuno = table.cell_value(i, 0)
            if isinstance(stuno, (float, int)):
                name_inf[str(int(stuno))] = str(table.cell_value(i, 1))
            elif isinstance(stuno, str):
                stuno_upper = stuno.upper()
                name_inf[stuno_upper] = str(table.cell_value(i, 1))
        insert_message('共找到' + str(len(name_inf)) + '个学号-姓名\n')
    except FileNotFoundError:
        messagebox.showerror(title='提示', message='未找到名单文件!')
        name_inf = None
    except Exception as e:
        messagebox.showerror(title='提示', message='错误信息为' + str(e))
        name_inf = None


# functions
def select_file():
    selected_directory_path = filedialog.askdirectory()
    select_path.set(selected_directory_path)


def process():
    clearall()
    inputdir = filepath_entry.get()
    if inputdir == '' or inputdir is None or inputdir.isspace():
        messagebox.showerror(title='提示', message='待转换文件夹不能为空！')
        return
    if name_inf == {} or name_inf is None:
        insert_message('还没有找到学号-姓名关系。可点击 选择学号-姓名文件 按钮来选取一个存有学号-姓名的xls表格。\n')
        insert_message('表格要求：第一列是学号，第二列是姓名，学号姓名从第二行开始填写。\n')
        return
    filetype = filetype_entry.get()
    if filetype == '' or filetype is None or filetype.isspace():
        messagebox.showerror(title='提示', message='文件类型不能为空！')
        return
    replaceby = replace_entry.get()
    if replaceby == '' or replaceby is None or replaceby.isspace():
        messagebox.showerror(title='提示', message='转换后文件名不能为空！')
        return
    file_list = []
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    for parent, dirnames, filenames in os.walk(inputdir):
        for fn in filenames:
            if os.path.splitext(fn)[1] == filetype:
                file_list.append(os.path.join(parent, fn))
    if len(file_list) == 0:
        messagebox.showerror(title='提示', message='未找到符合要求的文件！')
        return
    for filename in file_list:
        filename = filename.replace('/', '\\')
        if len(re.findall(re_stuno, filename)):
            stu_no = re.findall(re_stuno, filename)[0]
        else:
            continue
        try:
            rectified_name = replaceby.format(stu_no, name_inf[stu_no.upper()]) + filetype
        except KeyError:
            insert_message('未找到学号' + stu_no + '对应的姓名！\n')
            continue
        full_rectified_name = os.path.join(os.path.split(filename)[0], rectified_name)
        os.rename(filename, full_rectified_name)
        insert_message(os.path.split(filename)[1])
        insert_message('\n已更名为\n')
        insert_message(rectified_name)
        insert_message('\n')
    insert_message('\n全部完成')


def prompt():
    clearall()
    insert_message("提取学号以进行更名。\n先选取存有班级学号-姓名对应关系的xls表格。表格要求：第一列是学号，第二列是姓名，学号姓名从第二行开始填写。\n"
                   "“替换后文件名”中，{0}将会被替换成学号，{1}将会被替换成姓名。如：{0}-{1}-实验报告1 "
                   "将被替换为 学号-姓名-实验报告1。文件类型需要写完整扩展名，比如.pdf，而不是pdf。")


# 插入信息并保证插入后文本框仍不可编辑
def insert_message(message):
    scroll_inf.config(state="normal")
    scroll_inf.insert("end", message)
    scroll_inf.config(state="disabled")


def clearall():
    scroll_inf.config(state=tkinter.NORMAL)
    scroll_inf.delete("1.0", "end")
    scroll_inf.config(state=tkinter.DISABLED)


# 删除已有文件
def delete_file(dir):
    try:
        os.remove(dir)
        print("中间文件" + dir + "已被移除\n")
    except:
        messagebox.showerror(title='提示', message=dir + '无法删除，请手动删除！')


# GUI
master_window = Tk()
master_window.title('收作业文件更名小程序')
master_window.geometry('480x320')
master_window.resizable(False, False)

text_frame = Frame(master_window)
text_frame.pack(fill="x")
select_path = StringVar()
Label(text_frame, text='文件路径:').grid(row=0, column=0, padx=10)
filepath_entry = Entry(text_frame, textvariable=select_path, state="disabled")
filepath_entry.grid(row=0, column=1, padx=10)
filetype = StringVar()
Label(text_frame, text='文件类型:').grid(row=1, column=0, padx=10)
filetype_entry = Entry(text_frame, textvariable=filetype)
filetype_entry.grid(row=1, column=1, padx=10)
replaceby = StringVar(value="{0}-{1}")
Label(text_frame, text='替换后文件名格式:').grid(row=2, column=0, padx=10)
replace_entry = Entry(text_frame, textvariable=replaceby)
replace_entry.grid(row=2, column=1, padx=10)

button_frame = Frame(master_window)
button_frame.pack(fill="x")
Button(button_frame, text='选择文件夹', command=select_file, width=10).pack(side="left", padx=15)
Button(button_frame, text='选择学号-姓名文件', command=lambda: read_nameinf(), width=15).pack(side="left", padx=15)
Button(button_frame, text='使用说明', command=prompt, width=10).pack(side="left", padx=15)
Button(button_frame, text='执行转换', command=lambda: process(), width=10).pack(side="left", padx=15)
scroll_inf = scrolledtext.ScrolledText(master_window, font=('楷体', 14))
scroll_inf.pack(side="bottom", pady=10, fill="x")
scroll_inf.config(state="disabled")
master_window.mainloop()
