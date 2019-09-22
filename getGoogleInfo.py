#从谷歌地球kml文件中提取站点经纬度
#工具编写日期 2019.9.22
#作者：陈格欣
#第一版: 提取谷歌地球站点坐标经纬度
#第二版：更新提取规则

def getNum (kmlName):
    num = 0
    with open(kmlName,'r') as f:
        for i in f:
            if '<Placemark>' in i:
                return (num)
            else:
                num += 1

def getGoogleInfo (kmlName,tabName):
    import re
    import openpyxl
    import xlsxwriter
    workbook = xlsxwriter.Workbook(tabName)
    workbook.add_worksheet('sheet1')
    workbook.close()
    with open(kmlName,'r') as f:
        workbook = openpyxl.load_workbook(tabName)
        worksheet = workbook.worksheets[0]
        worksheet.append(['站名', '经度', '纬度'])
        list = []
        for i in f:
            for i in f.readlines()[getNum (kmlName):]:
                if '<name>' in i:
                    i = i[re.search('<name>', i).span()[1]:re.search('</name>', i).span()[0]]
                    list.append(i)
                if '<coordinates>' in i:
                    i = i[re.search('<coordinates>', i).span()[1]:re.search('</coordinates>', i).span()[0]]
                    longitude, latitude,other = i.split(',', 2)
                    list.append(float(longitude))
                    list.append(float(latitude))
                    worksheet.append(list)
                    list = []
        workbook.save(tabName)

#将本文件放置到需提取经纬度同文件路径
#将kml转化为txt
#参数1：txt文件名，也可以为路径
#参数2：保存到Excel文件名
getGoogleInfo ('新建文本文档.txt','3.xlsx')
