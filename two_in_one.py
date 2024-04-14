from PyPDF2 import PdfReader, PdfWriter, Transformation
import tkinter
from tkinter import scrolledtext, messagebox, filedialog, simpledialog
from pathvalidate import sanitize_filename


# 函数功能：返回某页pdf的尺寸
def check_size(page):
    if page.get('/Rotate', 0) in [90, 270]:
        # return page['/MediaBox'][2], page['/MediaBox'][3]
        return float(page['/MediaBox'][2]), float(page['/MediaBox'][3])
    else:
        # return page['/MediaBox'][3], page['/MediaBox'][2]
        return float(page['/MediaBox'][3]), float(page['/MediaBox'][2])
    # 原文链接：https://blog.csdn.net/u013546508/article/details/104674374/


# 函数功能：将两页pdf页面显示在一页纸上 (横竖均可，但横竖混杂文档无法做到)
def merge_two_in_one(inputdir, outputname):
    writer = PdfWriter()
    if inputdir == '':
        messagebox.showinfo(title='提示', message='先选择文件！')
        return
    if outputname == '' or outputname.isspace():
        messagebox.showinfo(title='提示', message='文件名不能为空！')
        return
    inputdir = inputdir.replace('/', '\\')
    split_name = inputdir.split('\\')
    outputname = sanitize_filename(outputname)
    split_name[-1] = outputname + '.pdf'
    outputdir = '\\'.join(split_name)
    reader = PdfReader(inputdir)
    total_page = len(reader.pages)  # 总页数
    for i in range(0, (total_page + 1) // 2):  # 默认向下取整
        # for i in range(0,1): # for test
        even, odd = 2 * i, 2 * i + 1  # even 奇数页偶数下标
        page_base = reader.pages[even]  # 以奇数页为基准（奇数页在左边或上面）
        height_base, width_base = check_size(page_base)
        if height_base < width_base:  # 横版
            height_base, width_base = width_base, height_base
            height_smaller = True
        else:  # 竖版
            height_smaller = False
        print(height_base, width_base)
        scaling_base_1 = height_base / (2 * width_base)
        scaling_base_2 = width_base / height_base
        scaling_base = min(scaling_base_1, scaling_base_2)
        print(scaling_base)
        if height_smaller:
            translation_base_1 = height_base / 2 + (height_base / 2 - width_base * scaling_base) / 2
            translation_base_2 = width_base - (width_base - height_base * scaling_base) / 2
        else:
            translation_base_1 = (width_base - height_base * scaling_base) / 2
            translation_base_2 = height_base - (height_base / 2 - width_base * scaling_base) / 2
        transformation_left = Transformation().scale(sx=scaling_base, sy=scaling_base). \
            rotate(-90).translate(tx=translation_base_1, ty=translation_base_2)
        print(translation_base_1, translation_base_2)
        page_base.add_transformation(transformation_left)
        # 尾页特判
        if i == (total_page + 1) // 2 - 1 and total_page % 2 == 1:
            page_base.rotate(-90)
            writer.add_page(page_base)
            break

        page_add = reader.pages[odd]
        height_add, width_add = check_size(page_add)
        if height_add < width_add:
            height_add, width_add = width_add, height_add
            height_smaller = True
        else:
            height_smaller = False
        scaling_add_1 = height_base / (2 * width_add)
        scaling_add_2 = width_base / height_add
        scaling_add = min(scaling_add_1, scaling_add_2)
        if height_smaller:
            translation_add_1 = (height_base / 2 - width_add * scaling_add) / 2
            translation_add_2 = translation_base_2
        else:
            translation_add_1 = translation_base_1
            translation_add_2 = height_base / 2 - (height_base / 2 - width_add * scaling_add) / 2
        transformation_right = Transformation().rotate(-90).scale(sx=scaling_add, sy=scaling_add). \
            translate(ty=translation_add_2, tx=translation_add_1)
        page_add.add_transformation(transformation_right)
        page_base.merge_page(page_add)
        page_base.rotate(-90)
        page_base.scale(sx=1.414, sy=1.414)  # A4放大为A3版式
        writer.add_page(page_base)
    with open(outputdir, "wb") as fp:
        writer.write(fp)
    messagebox.showinfo(message="转换成功！")


pdf_list = []


def clear_record():
    if not pdf_list:
        return
    res = messagebox.askyesno(title="提示",message="确定要清除所有记录吗？")
    if res:
        pdf_list.clear()
        output_record()


def clear_last_record():
    id = simpledialog.askinteger("输入序号","输入要删除的项的序号")
    if id is None or id<=0:
        messagebox.showerror(message="序号非法！")
        return
    try:
        del pdf_list[id-1]
        output_record()
    except IndexError:
        messagebox.showerror(message="序号非法！")


def output_record():
    clearall()
    if not pdf_list:
        insert_message("无记录")
        return
    insert_message("现有记录：\n")
    for ind in range(len(pdf_list)):
        message = str(ind+1)+ ". " +  str(pdf_list[ind]) + '\n'
        insert_message(message)


def merge_pdf_select_file():
    selected_file_path = tkinter.filedialog.askopenfilename(filetypes=[("PDF文件", ".pdf")])
    print(selected_file_path)
    if selected_file_path is None or selected_file_path == '':
        return
    start = simpledialog.askinteger("起始页码", "起始页码",initialvalue = 1)
    stop = simpledialog.askinteger("请输入终止页码", "默认为最后一页对应的页码",initialvalue=len(PdfReader(selected_file_path).pages))
    if start is None or stop is None:
        return
    if start <= 0 or stop > len(PdfReader(selected_file_path).pages):
        messagebox.showerror("错误提示", "超过有效页码范围，无效输入!")
        return
    pdf_list.append([selected_file_path, start, stop])
    output_record()


def merge_pdf(outputname):
    if outputname == '' or outputname.isspace():
        messagebox.showinfo(title='提示', message='文件名不能为空！')
        return
    selected_outputdirectory = tkinter.filedialog.askdirectory()
    if selected_outputdirectory == '' or selected_outputdirectory is None:
        messagebox.showinfo(title='提示', message='输出文件夹不能为空！')
        return
    selected_outputdirectory = selected_outputdirectory.replace('/', '\\')
    final_path = selected_outputdirectory + '\\' + sanitize_filename(outputname) +".pdf"
    if not pdf_list:
        return
    try:
        merger = PdfWriter()
        for item in pdf_list:
            source = open(item[0], 'rb')
            merger.append(fileobj=source, pages=(item[1] - 1, item[2]))
        output = open(final_path, "wb")
        merger.write(output)
        # Close File Descriptors
        merger.close()
        output.close()
        messagebox.showinfo(message="合并成功!")
    except Exception as e:
        messagebox.showerror(message='保存时发生错误，错误信息：\n' + str(e))


def scale_to_size():
    writer = PdfWriter()
    inputdir = tkinter.filedialog.askopenfilename()
    if inputdir == '' or inputdir is None:
        messagebox.showinfo(title='提示', message='输出文件夹不能为空！')
        return
    inputdir = inputdir.replace('/', '\\')
    outputdir =inputdir+'1.pdf'
    reader = PdfReader(inputdir)
    total_page = len(reader.pages)  # 总页数
    width = simpledialog.askfloat("宽度", "以cm为单位", initialvalue=21.0)
    height = simpledialog.askfloat("高度", "以cm为单位", initialvalue=29.7)
    for i in range(0, total_page):
        current_page = reader.pages[i]
        height_base, width_base = check_size(current_page)
        print(height_base, width_base)
        # scaling_base_2 = height_base / (height *28.35)
        # scaling_base_1 = width_base / (width *28.35)
        current_page.mediabox.lower_left=(0,0)
        current_page.mediabox.lower_right = (width*28.35, 0)
        current_page.mediabox.upper_left = (0, height*28.35)
        current_page.mediabox.upper_right = (width*28.35, height*28.35)
        writer.add_page(current_page)
    with open(outputdir, "wb") as fp:  # 原文件夹输出
        writer.write(fp)

# GUI
def select_file():
    selected_file_path = tkinter.filedialog.askopenfilename(filetypes=[("PDF文件", ".pdf")])
    select_path.set(selected_file_path)


def insert_message(message):
    scroll_inf.config(state=tkinter.NORMAL)
    scroll_inf.insert(tkinter.END, message)
    scroll_inf.config(state=tkinter.DISABLED)


def clearall():
    scroll_inf.config(state=tkinter.NORMAL)
    scroll_inf.delete("1.0", "end")
    scroll_inf.config(state=tkinter.DISABLED)

def prompt():
    clearall()
    with open("README.TXT",'r',encoding='utf-8') as f:
        content = f.read()
        insert_message(content)



# 窗口设置
master_window = tkinter.Tk()
master_window.title('PDF双页合一/合并程序')
master_window.geometry('480x320')
master_window.resizable(False, False)

select_path = tkinter.StringVar()
out_path = tkinter.StringVar()

filepath_frame = tkinter.Frame(master_window)
filepath_frame.pack(fill=tkinter.X, side=tkinter.TOP)
filepath_name = tkinter.Label(filepath_frame, text='文件路径：')
filepath_name.pack(side=tkinter.LEFT)
filepath_entry = tkinter.Entry(filepath_frame, textvariable=select_path, state=tkinter.DISABLED, width=200)
filepath_entry.pack(side=tkinter.RIGHT)

outputfilename_frame = tkinter.Frame(master_window)
outputfilename_frame.pack(fill=tkinter.X)
outputfilename_name = tkinter.Label(outputfilename_frame, text='输出文件名：')
outputfilename_name.pack(side=tkinter.LEFT)
outputfilename_entry = tkinter.Entry(outputfilename_frame, textvariable=out_path, width=200)
outputfilename_entry.pack(side=tkinter.RIGHT)

button_frame = tkinter.Frame(master_window)
button_frame.pack(fill=tkinter.X)
tkinter.Button(button_frame, text='(双页合一)选择文件', command=select_file, width=25).grid(row=0, column=0, padx=30)
tkinter.Button(button_frame, text='(双页合一)执行转换',
               command=lambda: merge_two_in_one(filepath_entry.get(), outputfilename_entry.get()), width=25). \
    grid(row=0, column=1, padx=30)
tkinter.Button(button_frame, text='(pdf合并)选择文件', command=merge_pdf_select_file, width=25).grid(row=1, column=0, padx=30)
tkinter.Button(button_frame, text='(pdf合并)清除指定记录', command=lambda: clear_last_record(), width=25).grid(row=1, column=1,
                                                                                                       padx=30)
tkinter.Button(button_frame, text='(pdf合并)清除所有记录', command=lambda: clear_record(), width=25).grid(row=2, column=0,
                                                                                                  padx=30)
tkinter.Button(button_frame, text='(pdf合并)输出记录', command=lambda: output_record(), width=25).grid(row=2, column=1,
                                                                                                 padx=30)
tkinter.Button(button_frame, text='(pdf合并)执行转换', command=lambda: merge_pdf(outputfilename_entry.get()), width=25). \
    grid(row=3, column=0, padx=30)
tkinter.Button(button_frame, text='使用方法', command=lambda: prompt(), width=25).grid(row=3, column=1, padx=30)
tkinter.Button(button_frame, text='PDF页面大小统一', command=lambda: scale_to_size(), width=25). \
    grid(row=4, column=0, padx=30)
scroll_inf = scrolledtext.ScrolledText(master_window, font=('楷体', 14))
scroll_inf.pack(side=tkinter.BOTTOM, pady=10, fill=tkinter.X)
scroll_inf.config(state=tkinter.DISABLED)
# insert_message('使用方法：选择PDF文件，再点击“执行转换”即可。默认源文件夹下输出。')
master_window.mainloop()
