from win32com.client import Dispatch
from datetime import datetime, timedelta
import os

print("运行开始！！")
zpath = os.getcwd() + '/'
zpath_new = os.getcwd() + '/data/'
app = Dispatch('Word.Application')
xlApp = Dispatch("Excel.Application")

xlBook = xlApp.Workbooks.Open(zpath + 'data.xlsx')

table = xlBook.Worksheets('sheet1')
info = table.UsedRange
nrows = info.Rows.Count

for i in range(nrows):
    companyName = table.Cells(i + 3, 1).Value
    if companyName is None:
        break
    productNumber = str(table.Cells(i + 3, 2).Value)
    SCCJ = str(table.Cells(i + 3, 3).Value)
    productName = str(table.Cells(i + 3, 4).Value)
    productTime = str(table.Cells(i + 3, 5).Value)
    PH = str(table.Cells(i + 3, 6).Value)
    LC = str(table.Cells(i + 3, 7).Value)
    GCZCH = table.Cells(i + 3, 8).Value
    YJZH = str(table.Cells(i + 3, 9).Value)
    CYWZ = str(table.Cells(i + 3, 10).Value)
    reqTimeStr = str(table.Cells(i + 3, 11).Value)
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

    if PH == 'None':
        pass
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
    print(zpath_new + companyName.replace('有限公司', '') +
          '_' + GH + "_" + temp + '.docx')
    doc.SaveAs(zpath_new + companyName.replace('有限公司', '') +
               '_' + GH + "_" + temp + '.docx')
    app.Documents.Close()

app.Quit()
xlBook.Close()
xlApp.Quit()
print("运行结束！！")
