# -*- coding: utf-8 -*-
import xlrd
import xlwt
import os


def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception,e:
        print str(e)
        
def readByFileName(fileName,startRow,strSrc,strDst):
    AllData = [];
    try:
        strSerach = strDst.replace(u"；",";")
        searchArray =  []
        searchArray = strSerach.split(";")
##        print searchArray
        data = open_excel(fileName)
        for sheet in data.sheets():
            nrows = sheet.nrows #行数
            ncols = sheet.ncols #列数
    ##        print nrows,ncols
            listData = []
            try:
                for rownum in range(int(startRow),nrows):
                     #print rownum
                     row = sheet.row_values(rownum)
                     rowOut = []
                     rowOut.append(row[int(strSrc)])
                     for key in searchArray:
                         row[int(key)] = row[int(key)].replace(",","")
                         rowOut.append(row[int(key)])
##                     print rowOut
                     #print row
                     if rowOut:
                         listData.append(rowOut)
            except Exception,error:
                print str(error)
            AllData.append(listData)
    except Exception,error:
        print str(error)
    
    return AllData
def createFileIfNotExit(fileName,listValue):
    if(os.path.exists(fileName) == True):
        return
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1',cell_overwrite_ok=True)
    x = 0;
    y = 0;
    for value in listValue:
        for valueData in value:
##            print x,y
##            print valueData
            ws.write(x , y, valueData)
            y = y + 1
        x = x + 1
        y = 0
    wb.save(fileName)
    
##dataDict = readByFileName("test.xls")
##for name in dataDict:
##    print name
##    print dataDict[name]
