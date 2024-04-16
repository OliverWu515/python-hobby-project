# python_hobby_project

### 介绍

用 python 写的两个小工具。软件包放在release 里。

- main_openpyxl.py（最新）、main_command_line.py、main_xlrd.py：

  - 动机：之前很喜欢一个名为 <i>WakeUp课程表 </i>的轻量级课表软件。这个软件本来可以自动抓取教务系统的课表（通过从html文件里获取信息），但是由于HITsz教务系统的特殊设计，其无法抓取成功。我又不想每学期或每次课表更新都手动输入一遍课表，于是就写了个小脚本，将从HITSZ教务系统官网下载得到的xlsx文件，转化为WakeUp课程表可识别的csv文件。

  - 本来是利用 xlrd 读写 Excel 文档的，结果发现这个库只能处理 xls 格式的表格，要先调用 Excel 将文档先转化成 xls 格式，有些卡顿，不太方便，而且 xlrd 已经停止维护了，遂改用 openpyxl 来读写。后来又搞了一个不带 tkinter 的命令行版本，更为简洁。

  - 使用说明硬编码在exe里了（笑），点击即可查看。输出的默认文件名是 res.csv。

  - 示例文件放在release中的 example.zip 压缩包中。

    | 待转换文件      | 对应的转换后示例文件 |
    | --------------- | -------------------- |
    | export0415.xlsx | res.csv              |
    | 21秋.xlsx       | res_21秋.csv         |

- merge_pdf_with_UI.py（最新）、two_in_one.py：

  - 动机：我曾得到一套题库，原来是一个很大的 PDF 文件，前半部分是各章习题，后半部分是各章答案。我想将其重新分章整理成 习题 + 答案的形式。于是就写了这个小脚本，利用PyPdf2库来辅助我。后来又用于将 iPad 上导出的 PDF 和电脑上的 PDF合并成同一份作业。
  - 使用说明：
    - 转换流程：点击“选择文件”来选定文件及要加入的页数，程序会显示已加入文件的文件路径及起止页码（一条这样的信息称为一条”记录“）。在文本框里输入 转换后文件的文件名（不需要带扩展名），点击“执行转换”，会弹出提示框要求选择路径，选择路径并确定后会自动开始转换。
    - 对于“记录”的操作：点击“清除指定记录”，输入想删除的记录序号，该记录即被删除；“清除所有记录”，字面意思。要单独更改某条记录的话，如果只是更改页码范围，就选择“更改某条记录的页码范围”，否则选取“更改某条记录”。
    - 特别地，该程序还可用于截取PDF的一部分或拆分PDF，只要选择单独的一个PDF文件，然后输入需要截取的那部分页面的页码即可。
  - 另一个 two_in_one.py 及其二进制程序包增加了一个双页合一功能。这个功能只是因为我打印电子书想省纸（笑）。双页合一往往是A4双页合并到A3，页面会扩大，所以又写了个很粗糙的页面大小规整功能。

- 程序注明new者用Nuitka打包，其余均用 pyinstaller 打包。

  - 用 pyinstaller 打包命令：pyinstaller -F –-noconsole xxx.py，其中-F是打包为单个文件，–-noconsole是程序运行时不出现控制台。如果电脑上有多个版本python解释器，最好是只有一个python环境下装有pyinstaller。
  - 用 Nuitka 打包命令是：python39 -m nuitka –-onefile --disable-console --show-memory --enable-plugins=tk-inter --output-dir=out --remove-output xxx.py
    - python39：我电脑上有多个版本python解释器，将python解释器程序改了名字，用python39来指明。若没有多个python解释器则直接写python即可。若电脑上装了多个python，最好的specify方法是：python解释器路径 + -m + 模块。但注意pyinstaller不是module，不能这样运行。
    - -m：即module
    - --onefile：打包成一个exe（或bin，其他操作系统）。如果是 --standalone，就是一个文件夹，其中包含软件本体和各种依赖库、待import的module等，像是绿色安装版。
    - --disable-console：程序运行时不出现控制台。
    - --show-memory：显示内存使用情况。
    - --enable-plugins=tk-inter：对于 tkinter、tensorflow等特殊的大包需要增加nuitka插件支持。列表可通过--plugin-list查看。
    - --output-dir=out：输出文件夹名为out
    - --remove-output：生成过程中在输出文件夹下生成 xxx.build 和 xxx.dist 文件夹（xxx为python文件名），分别存放生成过程的中间文件和 --standalone 形式的软件包。如果使用了 --standalone 命令，则 --remove-output 这条命令会在生成成功时清除 xxx.build 文件夹；如果使用了 --onefile 命令，则这条命令会把 xxx.dist 文件夹也清除掉。
    - xxx.py：最后跟上py源文件名。
    - 还有几条会用到的命令：
      - --nofollow-imports：所有在源文件里import的模块/包都不查找，软件包将不会包含与import的模块。对standalone模式不可用。--follow-imports与之相反。
      - --nofollow-import-to=xxx：不查找指定模块/包。--follow-import-to与之相反。
      - --include-package=xxx、--include-module=xxx：包含指定包和模块。
    - Details：[Nuitka User Manual — Nuitka the Python Compiler](https://nuitka.net/user-documentation/user-manual.html)
  - TODO：学习多文件打包方式。

### 备注

本人编程水平有限，所以

- 代码风格可能不够优雅；
- 程序结构可能比较混乱；
- 没有做跨平台支持；
- 还可能有一些小bug（）

敬请各位指正！！