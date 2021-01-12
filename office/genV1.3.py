from datetime import datetime, timedelta
from pathlib import Path
from docxtpl import DocxTemplate
import xlrd
import os


def read_excel():
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
        # 打开excel
        x1 = xlrd.open_workbook(my_file_name)
        # 打开sheet1
        table = x1.sheet_by_index(0)
        nrows = table.nrows

        print('生成报告数:' + str(nrows - 2))

        for i in range(nrows-2):
            reqTimeStr = str(table.cell_value(i + 2, 0))
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
            print(saveFileName)
            tpl.render(context)
            tpl.save(saveFileName)

    except Exception as err:
        print("err %s: " % err)
        blogpath = os.path.join(zpath, 'log_err.txt')
        f = open(blogpath, 'w+')
        f.writelines(repr(err))
        f.close()

if __name__ == '__main__':
    # 读取Excel
    print("运行开始！！")
    read_excel()
    print("运行结束！！")
