from datetime import datetime
from pathlib import Path
from tkinter import *
from tkinter import filedialog

from docxtpl import DocxTemplate
import xlrd
import os
import configparser
import sys


def resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    print(os.path.join(base_path, relative_path))
    return os.path.join(base_path, relative_path)


def valid_count():
    config = configparser.ConfigParser()
    config.read(resource_path(os.path.join('res', 'conf.ini')), encoding="utf8")
    return config.getint("sys_config", "totalCount"), config.getint("sys_config", "usedCount")


def update_valid(count):
    config = configparser.ConfigParser()
    config.read(resource_path(os.path.join('res', 'conf.ini')), encoding="utf8")
    config.set("sys_config", "usedCount", repr(count))
    config.write(open(resource_path(os.path.join('res', 'conf.ini')), "w"))


class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master, bg='white')
        self.pack(expand=YES, fill=BOTH)
        self.window_init()
        self.createWidgets()

    def window_init(self):
        self.master.title('报告批处理系统')
        self.master.bg = 'white'
        width, height = self.master.maxsize()
        self.master.geometry("{}x{}".format(500, 500))

    def createWidgets(self):
        # # fm1
        self.fm1 = Frame(self, bg='white')
        self.openButton = Button(self.fm1, text='选择表格文件', bg='#e4e4e5', fg='black', font=('微软雅黑', 12),
                                 command=self.fileOpen)
        self.openButton.pack(expand=YES)
        self.fm1.pack(side=TOP, pady=10, expand=NO, fill='x')

        # fm2
        self.fm2 = Frame(self, bg='white')
        self.predictEntry = Text(self.fm2, font=('微软雅黑', 10), fg='#FF4081', state=DISABLED)
        self.predictEntry.pack(side=LEFT, fill='y', padx=20, expand=YES)
        self.fm2.pack(side=TOP, expand=YES, fill="y")

    def output_predict_sentence(self, r):
        # self.predictEntry.delete(0, END)
        self.predictEntry.config(state=NORMAL)
        self.predictEntry.insert(INSERT, r + "\n")
        self.predictEntry.config(state=DISABLED)

    def fileOpen(self):
        fileName = filedialog.askopenfilename(title='选择表格文件', filetypes=[('Excel', '*.xlsx')])
        self.read_excel(fileName)
        self.output_predict_sentence("结束")

    def read_excel(self, fileName):
        zpath = os.getcwd() + '/'
        try:
            self.output_predict_sentence("选择文件为：" + fileName)

            my_file = Path(fileName)
            if my_file.exists():
                pass
            else:
                self.output_predict_sentence("文件不存在,重新选择文件！")

            fileNames = fileName.split('/')
            my_dir_name = zpath + fileNames[len(fileNames) - 1].replace('.xlsx', '')
            my_dir = Path(my_dir_name)
            if my_dir.exists():
                pass
            else:
                os.makedirs(my_dir)
                self.output_predict_sentence("创建存储目录")
            # 打开excel
            x1 = xlrd.open_workbook(fileName)
            # 打开sheet1
            table = x1.sheet_by_index(0)
            nrows = table.nrows
            validCount = valid_count()

            if nrows - 2 + validCount[1] > validCount[0]:
                self.output_predict_sentence('数据异常，联系开发人员！')
                return

            self.output_predict_sentence('预计生成报告数:' + str(nrows - 2))
            self.output_predict_sentence("开始渲染报告！")

            for i in range(nrows - 2):
                reqTimeStr = str(table.cell_value(i + 2, 0)).strip()
                companyName = table.cell_value(i + 2, 1)
                if companyName is None:
                    break
                productNumber = str(table.cell_value(i + 2, 2)).strip()
                SCCJ = str(table.cell_value(i + 2, 3)).strip()
                productName = str(table.cell_value(i + 2, 4)).strip()
                productTime = table.cell_value(i + 2, 5)
                PH = table.cell_value(i + 2, 6)
                LC = str(table.cell_value(i + 2, 7)).strip()
                GCZCH = table.cell_value(i + 2, 8)
                YJZH = table.cell_value(i + 2, 9)
                CYWZ = str(table.cell_value(i + 2, 10)).strip()
                GH = str(table.cell_value(i + 2, 11)).strip()
                reportTime = str(table.cell_value(i + 2, 12)).strip()
                # 日期转换
                reqTime = datetime.strptime(reqTimeStr, '%Y.%m.%d')
                reportTime = datetime.strptime(reportTime, '%Y.%m.%d')

                tpl = DocxTemplate(resource_path(os.path.join('res', 'tempdoc.docx')))
                context = {
                    'companyName': companyName,
                    'productNumber': productNumber,
                    # 'SCCJ': SCCJ,
                    # 'productName': productName,
                    # 'productTime': productTime,
                    # 'PH': PH,
                    # 'LC': LC,
                    # 'GCZCH': GCZCH,
                    # 'YJZH': YJZH,
                    'CYWZ': CYWZ,
                    'GH': GH,
                    'reqTime': "{0:%Y}.{0:%m}.{0:%d}".format(reqTime),
                    'checkTime': "{0:%Y}.{0:%m}.{0:%d}".format(reqTime),
                    'reportTime': "{0:%Y}.{0:%m}.{0:%d}".format(reportTime),
                }

                if productName == 'None':
                    context['productName'] = ''
                else:
                    context['productName'] = productName

                if LC == 'None':
                    context['LC'] = ''
                else:
                    context['LC'] = LC

                if productTime is None:
                    context['productTime'] = ''
                else:
                    if isinstance(productTime, float):
                        context['productTime'] = int(float(productTime))
                    elif isinstance(productTime, int):
                        context['productTime'] = int(productTime)
                    else:
                        context['productTime'] = str(
                            productTime).replace('00:00:00+00:00', '')

                if PH is None:
                    context['PH'] = ''
                else:
                    if isinstance(PH, float):
                        context['PH'] = int(float(PH))
                    else:
                        context['PH'] = PH

                if SCCJ == 'None':
                    context['SCCJ'] = ''
                else:
                    context['SCCJ'] = SCCJ

                if YJZH is None:
                    context['YJZH'] = ''
                else:
                    if isinstance(YJZH, float):
                        context['YJZH'] = int(float(YJZH))
                    else:
                        context['YJZH'] = YJZH

                if GCZCH is None:
                    context['GCZCH'] = ''
                else:
                    if isinstance(GCZCH, float):
                        context['GCZCH'] = int(float(GCZCH))
                    else:
                        context['GCZCH'] = GCZCH

                temp = str(i + 1)
                saveFileName = my_dir_name + '/' + \
                               companyName.replace('有限公司', '').strip() + '_' + \
                               GH + "_" + temp + '.docx'
                self.output_predict_sentence("第" + temp + "文件：" + saveFileName)
                tpl.render(context)
                tpl.save(saveFileName)

            update_valid(nrows - 2 + validCount[1])

        except Exception as err:
            blogpath = resource_path(os.path.join('res', 'log_err.txt'))
            f = open(blogpath, 'w+')
            f.writelines(repr(err))
            f.close()
            self.output_predict_sentence("报告生成失败,原因:" + repr(err))


if __name__ == '__main__':
    app = Application()
    app.mainloop()
