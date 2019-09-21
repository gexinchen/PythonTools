def getFile(excelName,Router,Type,):
    import os
    import time
    import os.path
    import xlsxwriter
    time_start=time.time()
    print ('开始获取指定文件，路径，大小，创建时间')
    #创建.xlsx文件
    #workbook = xlsxwriter.Workbook('Star.xlsx')
    workbook = xlsxwriter.Workbook(excelName)
    worksheet = workbook.add_worksheet()
    title = [U'文件路径',U'文件大小',U'创建时间',U'CMD命令']
    worksheet.write_row('A1',title)
    num = 1
    Catalog = Router
    print ('请稍等')
    for i in os.walk(Catalog):
        fileName = i[2]
        print (fileName)
        for n in fileName:
            (file,fileType) = os.path.splitext(n)
            if fileType == type:
                #获取该类型文件完整路径
                fileRoute = i[0] + '\\' + file + '.pdf'
                #print (fileRoute)
                #获取该类型文件文件大小
                tempsize = os.path.getsize(fileRoute)
                fileSize = str(tempsize)
                #获取该类型文件的创建时间
                startinfo = os.stat(fileRoute)
                temptime = startinfo.st_mtime
                fileTime = str(temptime)
                cmd = 'copy "' + fileRoute + '" d:\图纸\\'
                #把获取到文件路径添加进EXCEL
                data = [fileRoute,fileSize,fileTime,cmd]
                q = q+1
                temp = str(q)
                row = 'A' + temp
                #在EXCEL里插入数据
                worksheet.write_row(row,data)
    workbook.close()
    time_end=time.time()
    print('已完成,获取.pdf总共使用了',(str(time_end-time_start))[0:4],'s')

if __name__=="__main__":
    getFile('Star','E://','.pdf')
