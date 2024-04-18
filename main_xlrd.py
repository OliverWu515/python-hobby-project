import os
import win32com.client as win32
import xlrd
import xlwt
import re
from tkinter import Tk, Label, Button, StringVar, Entry, Frame, scrolledtext, messagebox, filedialog

# regex patterns
re_bracket = re.compile(r'^\[(.+)]$')  # 检测中括号
re_number = re.compile(r'^\[(.+)]\[(.+)]$')  # 检测周数+地点(非实验课)/节数+周数(实验课)
re_lesson = re.compile(r'^第(\d+)-(\d+)节$')  # 检测节数
re_lesson_experiment = re.compile(r'^(\d+)-(\d+)节$')  # 检测节数(实验课)
re_name_experiment = re.compile(r'【实验】')  # 实验课程名称检测


# functions
def select_file():
    selected_file_path = filedialog.askopenfilename(filetypes=[("xlsx表格", ".xlsx")])
    select_path.set(selected_file_path)


def process():
    temp = filepath_entry.get()
    temp = temp.replace('/', '\\')
    inputdir = os.path.dirname(temp)
    # 三个参数：父目录；所有文件夹名（不含路径）；所有文件名
    # for parent, dirnames, filenames in os.walk(inputdir):
    #     for fn in filenames:
    #         if fn == "export.xls" or fn == "res.csv" or fn == "res.xls":
    #             delete_file(os.path.join(parent, fn))
    ret = transform(temp, 56)  # xls
    if ret is False:
        return
    xls_file_name = temp.replace(os.path.splitext(temp)[1], '.xls')
    combine = False
    show_info_as_teacher = False
    if messagebox.askyesno("请选择", "是否将同一门实验的不同时段课程合并为一个名称?"):
        combine = True
    if combine and messagebox.askyesno("请选择", "是否将同一门实验的不同时段课程"
                                              "（如实验一/二等）放在”老师“一栏以方便查阅?"):
        show_info_as_teacher = True
    data = xlrd.open_workbook(xls_file_name)
    table = data.sheet_by_index(0)
    # 获取表格行列数
    nrows = table.nrows
    ncols = table.ncols

    excel = xlwt.Workbook(encoding='utf-8', style_compression=0)  # 创建一个新的工作簿，以UTF-8编码
    ws_1 = excel.add_sheet('1', cell_overwrite_ok=True)  # 创建一个名为"1"的表单

    ind = 0
    for col_index in range(1, ncols):
        for row_index in range(3, nrows):
            if table.cell_value(row_index, col_index) != "":
                ar = table.cell_value(row_index, col_index)
                lar = ar.split('\n')
                ind = data_process(lar, ind, col_index - 1, ws_1, combine, show_info_as_teacher)
                if ind == -1:
                    return
    # 表头
    ws_1.write(0, 0, "课程名称")
    ws_1.write(0, 1, "星期")
    ws_1.write(0, 2, "开始节数")
    ws_1.write(0, 3, "结束节数")
    ws_1.write(0, 4, "老师")
    ws_1.write(0, 5, "地点")
    ws_1.write(0, 6, "周数")
    delete_file(xls_file_name)
    dir = os.path.join(inputdir, "res.xls")
    excel.save(dir)
    if transform(dir, 6):  # csv
        insert_message("写入csv完成\n")
    delete_file(dir)


def data_process(inf, index, week, ws_1, combine, show_info_as_teacher):
    for item in inf:
        try:
            if re.match(re_lesson, item):
                ws_1.write(index, 2, re.match(re_lesson, item).groups()[0])  # 开始节数
                ws_1.write(index, 3, re.match(re_lesson, item).groups()[1])  # 结束节数
            elif re.match(re_number, item):
                if re.match(re_number, item).groups()[0][-1] == '周':  # (非实验课)
                    ws_1.write(index, 6, re.match(re_number, item).groups()[0][:-1])  # 周数
                    ws_1.write(index, 5, re.match(re_number, item).groups()[1])  # 地点
                else:
                    ws_1.write(index, 6, re.match(re_number, item).groups()[1][:-1])  # 周数
                    st_end = re.match(re_number, item).groups()[0]
                    ws_1.write(index, 2, re.match(re_lesson_experiment, st_end).groups()[0])  # 开始节数
                    ws_1.write(index, 3, re.match(re_lesson_experiment, st_end).groups()[1])  # 结束节数
            elif re.match(re_bracket, item):  # 有中括号,教师名/实验室名
                if re.match(re_bracket, item).groups()[0][-1].isdigit():
                    ws_1.write(index, 5, re.match(re_bracket, item).groups()[0])  # 实验室名
                else:
                    ws_1.write(index, 4, re.match(re_bracket, item).groups()[0])  # 教师名
            else:  # 没有特殊格式,课程名
                index += 1
                item_modified = item
                if combine and re.match(re_name_experiment, item):  # 为实验课
                    item_modified = item.split(' ')[0]
                ws_1.write(index, 0, item_modified)  # 课程名
                ws_1.write(index, 1, week)  # 星期
                experiment_info = ""
                if show_info_as_teacher:
                    try:
                        experiment_info = item.split(' ')[1]
                    except IndexError:
                        experiment_info = ""
                    finally:
                        ws_1.write(index, 4, experiment_info)
        except Exception as e:
            insert_message(str(e) + '\n发生错误时处理的项：' + item)
            return -1
    return index


# 转格式函数
def transform(inputdir, type):
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = 0
        excel.DisplayAlerts = 0
        wb = excel.Workbooks.Open(inputdir)
        # xlsx: FileFormat=51
        # xls:  FileFormat=56
        # csv: FileFormat=6
        if type == 56:  # xls 方便利用xlrd读写
            wb.SaveAs(inputdir.replace(os.path.splitext(inputdir)[1], '.xls'), FileFormat=type)
        elif type == 6:  # csv
            wb.SaveAs(inputdir.replace(os.path.splitext(inputdir)[1], '.csv'), FileFormat=type)
        wb.Close()
        excel.Application.Quit()
    except Exception as e:
        insert_message("转换过程发生错误,错误信息为" + str(e) + '\n')
        return False
    return True


def prompt():
    insert_message('使用方法：（需安装有Excel软件！）\n'
                   '先从校网导出课表Excel文件，点击“选择文件”选择文件路径，'
                   '再点击“开始转换”转换后'
                   '会产生一个csv文件，将此csv文件导入至WakeUp课程表即可。\n\n'
                   "1.1版本更新说明：在原来的版本中同一门课程的实验由于课程名称中带有“实验一”等字样"
                   "而不会自动合并，导致“已添课程”中每节实验课都单独列出，不方便查看。现在每门课程的实验课在导出时会合并起来，方便查看。"
                   "同时通过将“实验一/二”等信息放在“老师”栏（实验课的“老师”栏往往空出），为不想要冗长的课程名称"
                   "又想要区分不同时段实验的同学提供了新选择。")


# 插入信息并保证插入后文本框仍不可编辑
def insert_message(message):
    scroll_inf.config(state="normal")
    scroll_inf.insert("end", message)
    scroll_inf.config(state="disabled")


# 删除已有文件
def delete_file(dir):
    try:
        os.remove(dir)
        insert_message("中间文件" + dir + "已被移除\n")
    except:
        messagebox.showerror(title='提示', message=dir + '无法删除，请手动删除！')


# GUI
master_window = Tk()
master_window.title('HITsz课表信息提取程序for WakeUp课程表 V1.1')
master_window.geometry('480x320')
master_window.resizable(False, False)
select_path = StringVar()
Label(master_window, text='文件路径:').pack(fill="x")
filepath_entry = Entry(master_window, textvariable=select_path, state="disabled")
filepath_entry.pack(fill="x")
button_frame = Frame(master_window)
button_frame.pack(fill="x")
Button(button_frame, text='选择文件', command=select_file, width=15).pack(side="left", padx=20)
Button(button_frame, text='使用说明', command=prompt, width=15).pack(side="left", padx=25)
Button(button_frame, text='执行转换', command=lambda: process(), width=15).pack(side="right", padx=25)
scroll_inf = scrolledtext.ScrolledText(master_window, font=('楷体', 14))
scroll_inf.pack(side="bottom", pady=10, fill="x")
scroll_inf.config(state="disabled")
master_window.mainloop()
