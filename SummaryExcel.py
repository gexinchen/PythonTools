#汇总相同表头的Excel文件
#工具编写日期：2019.11.19
#作者：陈格欣
#第一版：初个版本，日后继续完善表头等内容判断

# coding:utf-8
import openpyxl
import xlrd
import os.path
import time

#开始时间
time_start = time.time()

#创建文件Summary_table.xlsx并保存
workbook = openpyxl.Workbook('Summary_table.xlsx')
workbook.save('Summary_table.xlsx')
#读取文件Summary_table.xlsx
workbook = openpyxl.load_workbook('Summary_table.xlsx')
#设置写入第0张sheet
worksheet = workbook.worksheets[0]

#Excel表路径
rootdir = "./ExcelFile"
files = os.listdir(rootdir)
#获取文件个数
num = len(files)

print ('读取数据中...请稍等...')
for file in files:
    #迭代文件夹，获取每个文件名称，获取相对路径
    path = './ExcelFile/' + file
    #通过相对路径读取表格
    sheets = xlrd.open_workbook(path)
    #读取指定Sheet名称
    #sheet = sheets.sheet_by_name('SheetName')
    #读取指定Sheet序号
    #两者二选一
    sheet = sheets.sheet_by_index(3)
    #获取最大行数
    rows = sheet.nrows
    #获取最大列数
    cols = sheet.ncols
    #除去表头，开始汇总的行数
    startrow = 4
    startrow -= 1
    for row in range(startrow,rows):
        list = sheet.row_values(row) + [file]
        worksheet.append(list)

print ('文件保存中...')
workbook.save('Summary_table.xlsx')
print ('文件保存完成...保存路径为该程序同目录下')
time_end = time.time()
print ('文件汇总完成...一共汇总了' + str(num) + '个文件' + '花了' + str(format(time_end-time_start,'0.2f')) + '秒')
