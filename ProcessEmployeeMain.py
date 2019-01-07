# -*- coding: utf-8 -*-
import os,sys
from PyQt4.QtGui import *  
from PyQt4.QtCore import *  

import ctypes
import time
import threading
import ProcessEmployee
import ProcessExcel
import xlrd,xlwt

sys.path.append(sys.path[0])
reload(sys)
sys.setdefaultencoding('utf-8')
from PyQt4 import QtCore, QtGui

g_file_name = "";
g_dataArray = []
g_tmpData = []
g_fileList = {}
items = []

def fTreeWidgetDoubleClick(item):
    name = str(item.text(0)).decode('utf8')
    print name
##    clearTableViewContent(g_main)
##    showTableViewContent(g_main,g_dataDict[name])
    
def fOpenFile():
    global g_tmpData
    global g_dataDict
    global g_fileList
    try:
        file_names = QtGui.QFileDialog.getOpenFileNames(None,"open file dialog","","Excel files(*.xls;|*.xlsx;)") 
        for file_name in file_names:
            file_name = str(file_name).decode("utf8")
            if(g_fileList.has_key(file_name)):
                continue
            showTreeViewContent(g_main,file_name)
            g_dataArray.append(readDataByFileName(file_name))
##            print g_dataArray
            showAllTableViewContent(g_main,g_dataArray)
            g_fileList[file_name] = "true"
    except Exception, error:
        QMessageBox.information(None,"Information",str(error))
    
    
def clearTreeViewContent():
    global g_tmpData
    global g_dataDict
    global g_fileList
    global g_dataArray
    global items
    g_main.treeWidget.setColumnCount(2);
    header = [u'文件列表',u' ']
    g_main.treeWidget.setHeaderLabels(header)
    g_main.treeWidget.clear()
    g_dataArray = []
    g_tmpData = []
    g_fileList = {}
    items = []
    
def showTreeViewContent(mainWindow,fileName):
    mainWindow.treeWidget.setColumnCount(2);
    header = [u'文件列表',u' ']
    mainWindow.treeWidget.setHeaderLabels(header)
    a = QtGui.QTreeWidgetItem();  
    a.setText(0, fileName);
    items.append(a)
    mainWindow.treeWidget.addTopLevelItems(items)

def clearTableViewContent():
    g_main.tableWidget.setRowCount(0);
    g_main.tableWidget.setColumnCount(0);
    g_main.tableWidgetResult.setRowCount(0);
    g_main.tableWidgetResult.setColumnCount(0);
    g_main.tableWidgetResultDst.setRowCount(0);
    g_main.tableWidgetResultDst.setColumnCount(0);
    
def showAllTableViewContent(mainWindow,dataDict):
    global g_tmpData
    allRows = 0;
    MaxCols = 0;
    mainWindow.tableWidget.clear()
    for sheetDatas in dataDict:
        for sheetData in sheetDatas:
            for data in sheetData:
                MaxCols = len(data)
                allRows = allRows + 1
    mainWindow.tableWidget.setRowCount(allRows)
    mainWindow.tableWidget.setColumnCount(MaxCols)
    i = 0
    dataDict.reverse()
    g_tmpData = []
    for sheetDatas in dataDict:
        for sheetData in sheetDatas:
            for data in sheetData:
                tmpArray = []
                j = 0
                for dataItem in data:
                    item = QtGui.QTableWidgetItem(dataItem)
                    mainWindow.tableWidget.setItem(i,j,item)
                    tmpArray.append(dataItem)
                    j = j + 1
                i = i + 1
                g_tmpData.append(tmpArray)

def readDataByFileName(name):
    try:
        if g_main.comboBoxType.currentIndex() == 0:
            g_main.spinBoxStart.setValue(5)
            g_main.spinBoxSrc.setValue(0)
            g_main.lineEditColumArray.setText("6")
            return ProcessExcel.readByFileName(name,"5","0",str(g_main.lineEditColumArray.text()).decode("utf8"))
        elif g_main.comboBoxType.currentIndex() == 1:
            g_main.spinBoxStart.setValue(6)
            g_main.spinBoxSrc.setValue(0)
            g_main.lineEditColumArray.setText("1;6;7")
            return ProcessExcel.readByFileName(name,"6","0",str(g_main.lineEditColumArray.text()).decode("utf8"))
        elif g_main.comboBoxType.currentIndex() == 2:
            return ProcessExcel.readByFileName(name,str(g_main.spinBoxStart.value()),
                                               str(g_main.spinBoxSrc.value()),
                                               str(g_main.lineEditColumArray.text()).decode("utf8"))
    except Exception,error:
        QMessageBox.information(None,"Information",str(error))

def fStartSearch():
    allRows = 0;
    MaxCols = 0;
    strSerach = str(g_main.lineEditKey.text()).decode("utf8")
    strSerach = strSerach.replace(u"；",";")
    searchArray =  []
    searchArray = strSerach.split(";")
    print searchArray
##    if g_main.comboBoxType.currentIndex() == 0:
    for keyName in searchArray:
        for tmpArray in g_tmpData:
            MaxCols = len(tmpArray)
            if tmpArray[0].find(keyName) >= 0:
                allRows = allRows + 1
    g_main.tableWidgetResult.clear()           
    g_main.tableWidgetResult.setRowCount(allRows)
    g_main.tableWidgetResult.setColumnCount(MaxCols)

    DstMaxCols = MaxCols;
    g_main.tableWidgetResultDst.clear()           
    g_main.tableWidgetResultDst.setRowCount(len(searchArray))
    g_main.tableWidgetResultDst.setColumnCount(DstMaxCols)
    
    dstInfo = [[]];
##    QMessageBox.information(None,"Information",str(g_main.lineEditKey.text()))
    
    i = 0
    dstInfo = [];
    
    
    for keyName in searchArray:
        tData = []
        count = [];
        for countNum in range (1,MaxCols):
            count.append(0)
        tData.append(keyName)
        for tmpArray in g_tmpData:
            if tmpArray[0].find(keyName) >= 0:
                j = 0;
                for tmpData in tmpArray:
                    item = QtGui.QTableWidgetItem(tmpData)
                    g_main.tableWidgetResult.setItem(i,j,item)
                    if j > 0:
                        try:
                            count[j - 1] = count[j - 1]+ float(tmpData)
                        except Exception,error:
                            print error
                            pass
                    j = j + 1
                i = i + 1
        for countData in count:
            tData.append(countData)
        dstInfo.append(tData)
    dstI = 0
##    print dstInfo
    for dstData in dstInfo:
        dstJ = 0
        for dstRowData in dstData:
            if type(dstRowData) == type(float(dstJ)):
                dstRowData = str(dstRowData)
            item = QtGui.QTableWidgetItem(dstRowData)
            g_main.tableWidgetResultDst.setItem(dstI,dstJ,item)
            dstJ = dstJ + 1
        dstI = dstI + 1

def fClearData():
    clearTableViewContent()
    clearTreeViewContent()


def fCheckOutResultDst():
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1',cell_overwrite_ok=True)
    filename = QtGui.QFileDialog.getSaveFileName(None,"choose file","","Excel files(*.xls;|*.xlsx;)");
    for i in range(0, g_main.tableWidgetResultDst.rowCount()):
        for j in range(0,g_main.tableWidgetResultDst.columnCount()):
            if(g_main.tableWidgetResultDst.item(i,j) != ""):
                info = str(g_main.tableWidgetResultDst.item(i,j).text()).decode("utf8")
                ws.write(i,j,info)
    wb.save(filename)

def fCheckOutResult():
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet('Sheet1',cell_overwrite_ok=True)
    filename = QtGui.QFileDialog.getSaveFileName(None,"choose file","","Excel files(*.xls;|*.xlsx;)");
    for i in range(0, g_main.tableWidgetResult.rowCount()):
        for j in range(0,g_main.tableWidgetResult.columnCount()):
            if(g_main.tableWidgetResult.item(i,j) != ""):
                info = str(g_main.tableWidgetResult.item(i,j).text()).decode("utf8")
                ws.write(i,j,info)
    wb.save(filename)
                
g_main = ProcessEmployee.Ui_Dialog()
if __name__ == "__main__":        
    app = QtGui.QApplication(sys.argv) 
    Form = QtGui.QMainWindow()
    g_main.setupUi(Form)
    g_main.treeWidget.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
    g_main.treeWidget.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarAsNeeded)
    g_main.treeWidget.itemDoubleClicked.connect(fTreeWidgetDoubleClick)
    
    g_main.pushButton.clicked.connect(fOpenFile)
    g_main.pushButtonSearch.clicked.connect(fStartSearch)
    g_main.pushButtonClear.clicked.connect(fClearData)
    g_main.pushButtonResultOut.clicked.connect(fCheckOutResult)
    g_main.pushButtonResultDstOut.clicked.connect(fCheckOutResultDst)
    Form.show()
    sys.exit(app.exec_())
