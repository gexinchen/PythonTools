#从谷歌地球kml文件中提取站点经纬度
#工具编写日期 2019.9.22
#作者：陈格欣

def getGoogleInfo (kmlName,tabName):
    import openpyxl
    import xlsxwriter
    workbook = xlsxwriter.Workbook(tabName)
    workbook.add_worksheet('sheet1')
    workbook.close()
    with open(kmlName,'r') as f:
        workbook = openpyxl.load_workbook(tabName)
        worksheet = workbook.worksheets[0]
        worksheet.append(['站名', '经度', '纬度'])
        workbook.save(tabName)
        list = []
        for i in f:
            if ('<name>' in i) and ('.kml' not in i):
                i = i[8:len(i)-8]
                list.append(i)
            if '<coordinates>' in i:
                i = i[16:len(i)-17]
                longitude,latitude = i.split(',', 1)
                longitude = float(longitude)
                latitude = float(latitude)
                list.append(longitude)
                list.append(latitude)
                workbook = openpyxl.load_workbook(tabName)
                worksheet = workbook.worksheets[0]
                worksheet.append(list)
                workbook.save(tabName)
                list = []

#将本文件放置到需提取经纬度同文件路径
#参数1：kml文件名，也可以为路径
#参数2：保存到Excel文件名
getGoogleInfo ('新建文本文档.kml','3.xlsx')
