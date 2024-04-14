import PyPDF2
from PyPDF2 import PdfReader, PdfWriter
import tkinter
from tkinter import scrolledtext, messagebox, filedialog, simpledialog
from pathvalidate import sanitize_filename

pdf_list = []


def clear_record():
    if not pdf_list:
        return
    res = messagebox.askyesno(title="提示", message="确定要清除所有记录吗？")
    if res:
        pdf_list.clear()
        output_record()


def clear_last_record():
    if len(pdf_list) == 0:
        return
    id = simpledialog.askinteger("输入序号", "输入要删除的项的序号")
    if id is None or id <= 0:
        messagebox.showerror(message="序号非法！")
        return
    try:
        del pdf_list[id - 1]
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
        message_index = "序号【" + str(ind + 1) + "】" + '\n'
        insert_message(message_index)
        message_filepath = "文件路径：" + str(pdf_list[ind][0]) + '\n'
        insert_message(message_filepath)
        message_startpage = "起始页码：" + str(pdf_list[ind][1]) + '\n'
        insert_message(message_startpage)
        message_endpage = "结束页码：" + str(pdf_list[ind][2]) + '\n'
        insert_message(message_endpage)
        insert_message("---------\n")


def get_item():
    selected_file_path = tkinter.filedialog.askopenfilename(filetypes=[("PDF文件", ".pdf")])
    print(selected_file_path)
    if selected_file_path is None or selected_file_path == '':
        return None
    reader = PdfReader(selected_file_path)
    pw = ""
    while reader.is_encrypted and reader.decrypt(pw) == PyPDF2.PasswordType.NOT_DECRYPTED:
        pw = simpledialog.askstring(title="", prompt="文件被加密，请输入密码")
    total_page_num = len(reader.pages)
    reader.stream.close()
    start = simpledialog.askinteger("起始页码", "起始页码", initialvalue=1)
    stop = simpledialog.askinteger("请输入终止页码", "默认为最后一页对应的页码", initialvalue=total_page_num)
    if start is None or stop is None:
        return None
    if start <= 0 or stop > total_page_num:
        messagebox.showerror("错误提示", "超过有效页码范围，无效输入!")
        return None
    return [selected_file_path, start, stop, pw]


def merge_pdf_select_file():
    temp = get_item()
    if temp is not None:
        pdf_list.append(temp)
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
    final_path = selected_outputdirectory + '\\' + sanitize_filename(outputname) + ".pdf"
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


def change_record():
    if len(pdf_list) == 0:
        return
    id = simpledialog.askinteger("输入序号", "输入要更改的项的序号")
    if id is None or id <= 0 or id > len(pdf_list):
        messagebox.showerror(message="序号非法！")
        return
    temp = get_item()
    if temp is not None:
        pdf_list[id - 1] = temp
        messagebox.showinfo(message="更改成功！")
    output_record()


def change_page():
    if len(pdf_list) == 0:
        return
    id = simpledialog.askinteger("输入序号", "输入要更改的项的序号")
    if id is None or id <= 0 or id > len(pdf_list):
        messagebox.showerror(message="序号非法！")
        return
    reader = PdfReader(pdf_list[id - 1][0])
    if reader.is_encrypted:
        reader.decrypt(pdf_list[id - 1][3])
    total_page_num = len(reader.pages)
    reader.stream.close()
    start = simpledialog.askinteger("起始页码", "起始页码", initialvalue=1)
    stop = simpledialog.askinteger("请输入终止页码", "默认为最后一页对应的页码", initialvalue=total_page_num)
    if start is None or stop is None:
        return None
    if start <= 0 or stop > total_page_num:
        messagebox.showerror("错误提示", "超过有效页码范围，无效输入!")
        return None
    pdf_list[id - 1][1] = start
    pdf_list[id - 1][2] = stop
    messagebox.showinfo(message="更改成功！")
    output_record()


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
    # try:
    #     f = open("README00.TXT",'r',encoding='utf-8')
    #     content = f.read()
    #     insert_message(content)
    # except:
    #     messagebox.showerror(message='未找到帮助文件！')
    insert_message("使用说明：\n点击“选择文件”来选定文件及要加入的页数，程序会显示已加入文件的文件路径及起止页码。"
                   "在文本框里输入 输出PDF的文件名（不需要带扩展名），"
                   "点击“执行转换”，会弹出提示框要求选择路径，选择路径并确定后会自动开始转换。"
                   "\n点击“清除指定记录”，输入想删除的记录序号，该记录即被删除。还可以更改某条记录的页码范围或更改整条记录。"
                   "\n【提示：该功能还可用于截取PDF的一部分或拆分PDF，只要选择单独的一个PDF文件，然后输入需要截取的那部分页面的页码即可。】")


# 窗口设置
master_window = tkinter.Tk()
master_window.title('PDF合并程序')
master_window.geometry('640x480')
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
tkinter.Button(button_frame, text='选择文件', command=merge_pdf_select_file, width=25).grid(row=0, column=0, padx=30)
tkinter.Button(button_frame, text='清除指定记录', command=lambda: clear_last_record(), width=25).grid(row=0, column=1,
                                                                                                padx=30)
tkinter.Button(button_frame, text='清除所有记录', command=lambda: clear_record(), width=25).grid(row=1, column=0,
                                                                                           padx=30)
tkinter.Button(button_frame, text='输出记录', command=lambda: output_record(), width=25).grid(row=1, column=1,
                                                                                          padx=30)
tkinter.Button(button_frame, text='执行转换', command=lambda: merge_pdf(outputfilename_entry.get()), width=25). \
    grid(row=3, column=1, padx=30)
tkinter.Button(button_frame, text='使用方法', command=lambda: prompt(), width=25).grid(row=3, column=0, padx=30)
tkinter.Button(button_frame, text='更改某条记录', command=lambda: change_record(), width=25). \
    grid(row=2, column=1, padx=30)
tkinter.Button(button_frame, text='更改某条记录的页码范围', command=lambda: change_page(), width=25). \
    grid(row=2, column=0, padx=30)
scroll_inf = scrolledtext.ScrolledText(master_window, font=('楷体', 14))
scroll_inf.pack(side=tkinter.BOTTOM, pady=10, fill=tkinter.X)
scroll_inf.config(state=tkinter.DISABLED)
output_record()
master_window.mainloop()
