# import qrcode
import xlrd
from MyQR import myqr
from PIL import Image
# import xlwings as xw
# 居中 插入
import os
import tkinter as tk

from openpyxl import load_workbook


def getInfo():
    try:
        data = xlrd.open_workbook(r"D:/111/aaa.xls")
        sheetname = "Sheet1"
        table = data.sheet_by_name(sheetname)
        data_cols = table.ncols
        for i in range(1):
            col_values = table.col_values(i)
            muti_code(data, i, col_values)
        return col_values
    except Exception as e:
        print(e)


def muti_code(data, i, col_values):
    s = 0

    wb = load_workbook('D:/111/aaa.xlsx')
    sht = wb["Sheet1"]
    for machCode in col_values:
        if machCode:
            filename = machCode
            # now_save = 'qr_png/jpg_code/{}.jpg'.format(machCode)
            makdirs(i)
            myqr.run(words=filename, version=4, save_name='./qr_png/jpg_code/{}/{}.jpg'.format('a',
                                                                                               machCode))
            # myqr.run(words=filename,version=4,save_name='qr_png/jpg_code/test/{}.jpg'.format(machCode))
            s += 1
            try:
                sht.add_image('./qr_png/jpg_code/{}/{}.jpg'.format('a', machCode), 'B' + s)
                wb.save()
            except Exception as e:
                print(e)
            # print("执行成功！第%s条！" % (str(s)))
        else:
            pass
    # print("执行成功！共%s条！" % (str(s)))


def makdirs(i):
    if not os.path.isdir('./qr_png/jpg_code/{}'.format('a')):  # 创建保存路径
        os.mkdir('./qr_png/jpg_code/{}'.format('a'))


getInfo()
