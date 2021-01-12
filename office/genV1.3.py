from win32com.client import Dispatch
from datetime import datetime, timedelta
from pathlib import Path
from docxtpl import DocxTemplate
import win32timezone

import os

print("运行开始！！")
zpath = os.getcwd() + '/'
try:
    filename = input('输入文件名： ')
    print(filename)

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
        productNumber = str(table.Cells(i + 3, 3).Value).strip()
        SCCJ = str(table.Cells(i + 3, 4).Value).strip()
        productName = str(table.Cells(i + 3, 5).Value).strip()
        productTime = table.Cells(i + 3, 6).Value
        PH = table.Cells(i + 3, 7).Value
        LC = str(table.Cells(i + 3, 8).Value).strip()
        GCZCH = table.Cells(i + 3, 9).Value
        YJZH = table.Cells(i + 3, 10).Value
        CYWZ = str(table.Cells(i + 3, 11).Value).strip()
        GH = str(table.Cells(i + 3, 12).Value).strip()
        # 日期转换
        reqTime = datetime.strptime(reqTimeStr, '%Y.%m.%d')
        nextNow = reqTime + timedelta(days=1)

        tpl = DocxTemplate(zpath + 'tempdoc.docx')
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
            'reportTime': "{0:%Y}.{0:%m}.{0:%d}".format(nextNow),
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
                context['YJZH'] = GCZCH

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
        print(saveFileName)
        tpl.render(context)
        tpl.save(saveFileName)

    xlBook.Close()
    xlApp.Quit()
except Exception as err:
    print("err %s: " % err)
    blogpath = os.path.join(zpath, 'test1.txt')
    f = open(blogpath, 'w+')
    f.writelines(repr(err))
    f.close()

print("运行结束！！")
