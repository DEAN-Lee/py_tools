from win32com.client import Dispatch
from datetime import datetime, timedelta
from pathlib import Path
import os

print("运行开始！！")

filename = input('输入文件名： ')
print(filename)

zpath = os.getcwd() + '/'

my_file_name = zpath + filename

my_file = Path(my_file_name)
if my_file.exists():
    pass
else:
    print("文件不存在,重新输入文件名称！")
    os._exit()

my_dir_name = zpath + filename.replace('.xlsx', '')

my_dir = Path(my_dir_name)
if my_dir.exists():
    pass
else:
    os.makedirs(my_dir)
    print("创建文件存储目录")

app = Dispatch('Word.Application')
xlApp = Dispatch("Excel.Application")

xlBook = xlApp.Workbooks.Open(my_file_name)

table = xlBook.Worksheets('sheet1')
info = table.UsedRange
nrows = info.Rows.Count

print('生成报告数:' + str(nrows - 2))

for i in range(nrows):
    reqTimeStr = str(table.Cells(i + 3, 1).Value)
    companyName = table.Cells(i + 3, 2).Value
    if companyName is None:
        break
    productNumber = str(table.Cells(i + 3, 3).Value)
    SCCJ = str(table.Cells(i + 3, 4).Value)
    productName = str(table.Cells(i + 3, 5).Value)
    productTime = str(table.Cells(i + 3, 6).Value)
    PH = table.Cells(i + 3, 7).Value
    LC = str(table.Cells(i + 3, 8).Value)
    GCZCH = table.Cells(i + 3, 9).Value
    YJZH = str(table.Cells(i + 3, 10).Value)
    CYWZ = str(table.Cells(i + 3, 11).Value)
    GH = str(table.Cells(i + 3, 12).Value).strip()
    # 日期转换
    reqTime = datetime.strptime(reqTimeStr, '%Y.%m.%d')
    nextNow = reqTime + timedelta(days=1)

    doc = app.Documents.Add(zpath + 'templeteTrade.dotx')

    doc.Bookmarks("companyName").Range.Text = companyName
    doc.Bookmarks("productName").Range.Text = productName
    doc.Bookmarks("productNumber").Range.Text = productNumber

    if productTime == 'None':
        pass
    else:
        doc.Bookmarks("productTime").Range.Text = productTime.replace(
            '00:00:00+00:00', '')

    if PH is None:
        pass
    else:
        if isinstance(PH, float):
            doc.Bookmarks("PH").Range.Text = int(float(PH))
        else:
            doc.Bookmarks("PH").Range.Text = PH

    if SCCJ == 'None':
        pass
    else:
        doc.Bookmarks("SCCJ").Range.Text = SCCJ

    doc.Bookmarks("LC").Range.Text = LC

    if GCZCH is None:
        pass
    else:
        if isinstance(GCZCH, float):
            doc.Bookmarks("GCZCH").Range.Text = int(float(GCZCH))
        else:
            doc.Bookmarks("GCZCH").Range.Text = GCZCH

    doc.Bookmarks("CYWZ").Range.Text = CYWZ

    if YJZH == 'None':
        pass
    else:
        doc.Bookmarks("YJZH").Range.Text = YJZH

    doc.Bookmarks("GH").Range.Text = GH
    doc.Bookmarks(
        "reqTime").Range.Text = "{0:%Y}.{0:%m}.{0:%d}".format(reqTime)
    doc.Bookmarks(
        "checkTime").Range.Text = "{0:%Y}.{0:%m}.{0:%d}".format(reqTime)
    doc.Bookmarks(
        "reportTime").Range.Text = "{0:%Y}.{0:%m}.{0:%d}".format(nextNow)
    temp = str(i + 1)
    print(my_dir_name + '/' + companyName.replace('有限公司', '') +
          '_' + GH + "_" + temp + '.docx')
    doc.SaveAs(my_dir_name + '/' + companyName.replace('有限公司',
                                                       '') + '_' + GH + "_" + temp + '.docx')
    app.Documents.Close()

app.Quit()
xlBook.Close()
xlApp.Quit()
print("运行结束！！")
