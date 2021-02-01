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

        my_dir_name = zpath + filename.replace('.xls', '')

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

        print('生成报告数:' + str(nrows - 1))

        for i in range(nrows - 1):
            number = table.cell_value(i + 1, 1)
            if number is None:
                break
            id = str(table.cell_value(i + 1, 2)).strip()
            name = str(table.cell_value(i + 1, 3)).strip()
            sex = str(table.cell_value(i + 1, 4)).strip()
            age = table.cell_value(i + 1, 5)


            tpl = DocxTemplate(zpath + 'person.docx')
            context = {
                'number': number,
                'id': id,
                'name': name,
                'sex': sex,
            }

            if age is None:
                context['age'] = ''
            else:
                if isinstance(age, float):
                    context['age'] = int(float(age))
                else:
                    context['age'] = age

            temp = str(i + 1)
            saveFileName = my_dir_name + '/' + \
                           name.strip() + '_' + \
                           "_" + temp + '.docx'
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
