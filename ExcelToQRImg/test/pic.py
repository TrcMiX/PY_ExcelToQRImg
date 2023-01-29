# coding=UTF-8
import xlrd
import xlwt
import requests
import json
import xlsxwriter
# from xlutils.copy import copy
import glob
def  insert(workbook,sheet_name_xls,value_title,url,picturePath,):
    lengh=len(picturePath)
    sheet = workbook.add_worksheet(sheet_name_xls)  # 在工作簿中新建一个表格
    index = len(value_title)
    for i in range(0, index):
        for j in range(0, len(value_title[i])):
            sheet.write(i, j, value_title[i][j])  # 像表格中写入数据（对应的行和列）
            print("xls格式表格写入数据成功！")
    arrayList = []
    for t in range(0, lengh):  # 通过range()把行数生成一个可迭代对象
        files = {'photo': open(picturePath[t], 'rb')}
        # print(uploadpicture(url, files))
        list = []  ## 空列表
        list.append(url)  ## 使用 append() 添加元素
        print(picturePath[t])
        text = uploadpicture(url, files)
        list.append(text)
        list.append(picturePath[t])
        try:
            re = json.loads(text)
            if re['status']==200:
                list.append('成功')
            else:
                list.append('失败')
            data = re['data']
            list.append(str(data['score']))
            arrayList.append(list)
            print(list)
        except Exception as e:
            print(e)
            continue


    rows_old = 1  # 获取表格中已存在的数据的行数
    index1 = len(arrayList)  # 获取需要写入数据的行数
    for i in range(0, index1):
        for j in range(0, len(arrayList[i])):
            if j == 2:
                try:
                    sheet.insert_image(i + rows_old, j, arrayList[i][j], {'x_scale': 0.1, 'y_scale': 0.1})
                except IOError:
                    sheet.write(i + rows_old, j, arrayList[i][j])
            else:
                sheet.write(i + rows_old, j, arrayList[i][j])  # 追加写入数据，注意是从i+rows_old行开始写入

                # new_workbook.save(path)  # 保存工作簿
                # print("xls格式表格【追加】写入数据成功！")
def  uploadpicture(url,files ):
    r = requests.post(url, files=files)
    return r.text
workbook = xlsxwriter.Workbook("code2")  # 新建一个工作簿