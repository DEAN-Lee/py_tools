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
    print("文件不存在")
    os._exit()

my_dir_name = zpath + filename.replace('.xlsx', '')

my_dir = Path(my_dir_name)
if my_dir.exists():
    pass
else:
    os.makedirs(my_dir)
    print("创建目录")

app = Dispatch('Word.Application')
xlApp = Dispatch("Excel.Application")
xlBook = xlApp.Workbooks.Open(zpath + 'person.xlsx')

table = xlBook.Worksheets('sheet1')
info = table.UsedRange
nrows = info.Rows.Count

print('预计生成报告数:' + str(nrows - 2))

for i in range(nrows):
    number = table.Cells(i + 3, 1).Value
    if number is None:
        break
    name = str(table.Cells(i + 3, 2).Value)
    sex = str(table.Cells(i + 3, 3).Value)
    age = table.Cells(i + 3, 4).Value
    ID = str(table.Cells(i + 3, 5).Value)
    reqTimeStr = str(table.Cells(i + 3, 6).Value)
    # 日期转换
    reqTime = datetime.strptime(reqTimeStr, '%Y.%m.%d')
    nextNow = reqTime + timedelta(days=1)

    doc = app.Documents.Add(zpath + 'templete.dotx')
    doc.Bookmarks("name").Range.Text = name
    doc.Bookmarks("age").Range.Text = age
    doc.Bookmarks("sex").Range.Text = sex
    doc.Bookmarks("number").Range.Text = number
    doc.Bookmarks("ID").Range.Text = ID
    doc.Bookmarks(
        "checkTime").Range.Text = "{0:%Y}.{0:%m}.{0:%d}".format(reqTime)
    doc.Bookmarks("reqTime").Range.Text = "{0:%m}.{0:%d}".format(reqTime)
    doc.Bookmarks(
        "reportTime").Range.Text = "{0:%Y}.{0:%m}.{0:%d}".format(nextNow)
    print(my_dir_name + '/' + name + '_' + ID + '.docx')
    doc.SaveAs(my_dir_name + '/' + name + '_' + ID + '.docx')
    app.Documents.Close()

app.Quit()
xlBook.Close()
xlApp.Quit()
print("运行结束！！")
