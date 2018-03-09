#-*-coding:utf-8-*-
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

import xlrd
from PyQt4 import QtGui, QtCore
from Ui_ExcelDiffer import Ui_ExcelDiffer

class Main(QtGui.QWidget, Ui_ExcelDiffer):
    def __init__(self, parent=None):
        super(Main, self).__init__(parent)
        self.setupUi(self)

        # 添加事件
        # 选择旧文件按钮点击事件
        self.originExcelChooseButton.clicked.connect(lambda :self.chooseExcelFile('originExcelChooseButton'))
        # 选择新文件按钮点击事件
        self.newExcelChooseButton.clicked.connect(lambda :self.chooseExcelFile('newExcelChooseButton'))
        # 差异对比按钮点击事件
        self.excelDiffButton.clicked.connect(lambda : self.excelDiffOperate(self.originExcelUrlLineEdit.text(), self.newExcelUrlLineEdit.text()))
        # 清空数据按钮点击事件
        self.refreshButton.clicked.connect(lambda : self.refreshWindow())
        # Sheet增删情况列表点击事件
        self.sheetListTableWidget.itemClicked.connect(self.sheetTableWidgetClicked)
        # 新增Sheet列表点击事件
        self.sheetAddTableWidget.itemClicked.connect(self.sheetTableWidgetClicked)
        # 删除Sheet列表点击事件
        self.sheetDelTableWidget.itemClicked.connect(self.sheetTableWidgetClicked)
        # 行增删列表点击事件
        self.rowAddDelTableWidget.itemClicked.connect(self.trackingRow)
        # 列增删列表点击事件
        self.columnAddDelTableWidget.itemClicked.connect(self.trackingCol)
        # 单元格改动列表点击事件
        self.cellModifyTableWidget.itemClicked.connect(self.trackingCell)

    # 选择文件并填充对应表格
    def chooseExcelFile(self, button):
        excelUrl = QtGui.QFileDialog.getOpenFileName(self,  "选取文件".decode('utf-8'),  ".","ExcelFile.xlsx(*.xlsx);;ExcelFile.xls(*.xls)")
        if len(excelUrl) == 0:
            return
        # 获取excelName
        urlStr = excelUrl.split('/')
        excelName = urlStr[len(urlStr)-1]
        # 读出数据并填充界面
        if button == 'originExcelChooseButton':
            self.originExcelUrlLineEdit.setText(excelUrl)
            # 读取excel数据
            originExcel = xlrd.open_workbook(str(excelUrl).decode('utf-8'))
            originExcelData, originExcelSheetSortList = self.readExcelData(originExcel)
            originExcelSheetLabelString = excelName + ':' + str(originExcelSheetSortList[0]).decode('utf-8')
            self.originExcelSheetLabel.setText(originExcelSheetLabelString)
            # 初始化旧Excel表格（默认为sheet1中的数据）
            rowCount = len(originExcelData[originExcelSheetSortList[0]])
            columnCount = len(originExcelData[originExcelSheetSortList[0]][0])
            self.originExcelViewTableWidget.setRowCount(rowCount)
            self.originExcelViewTableWidget.setColumnCount(columnCount)
            columnHeaderLabelList = self.createExcelColumnHeaderLabelList(columnCount)
            self.originExcelViewTableWidget.setHorizontalHeaderLabels(columnHeaderLabelList)
            # 将读出的Excel数据填入界面表单
            for row in range(0, rowCount):
                for col in range(0, columnCount):
                    # 数字处理，整形去掉小数点
                    cellValue = originExcelData[originExcelSheetSortList[0]][row][col]
                    if (type(cellValue) == float) and (cellValue % 1.0 == 0.0):
                        item = QtGui.QTableWidgetItem(str(int(cellValue)).decode('utf-8'))
                    else:
                        item = QtGui.QTableWidgetItem(str(cellValue).decode('utf-8'))
                    self.originExcelViewTableWidget.setItem(row, col, item)

        elif button == 'newExcelChooseButton':
            self.newExcelUrlLineEdit.setText(excelUrl)
            # 读取excel数据
            newExcel = xlrd.open_workbook(str(excelUrl).decode('utf-8'))
            newExcelData, newExcelSheetSortList = self.readExcelData(newExcel)
            newExcelSheetLabelString = excelName + ':' + str(newExcelSheetSortList[0]).decode('utf-8')
            self.newExcelSheetLabel.setText(newExcelSheetLabelString)
            # 初始化新Excel表格（默认为sheet1中的数据）
            rowCount = len(newExcelData[newExcelSheetSortList[0]])
            columnCount = len(newExcelData[newExcelSheetSortList[0]][0])
            self.newExcelViewTableWidget.setRowCount(rowCount)
            self.newExcelViewTableWidget.setColumnCount(columnCount)
            columnHeaderLabelList = self.createExcelColumnHeaderLabelList(columnCount)
            self.newExcelViewTableWidget.setHorizontalHeaderLabels(columnHeaderLabelList)
            # 将读出的Excel数据填入界面表单
            for row in range(0, rowCount):
                for col in range(0, columnCount):
                    # 数字处理，整形去掉小数点
                    cellValue = newExcelData[newExcelSheetSortList[0]][row][col]
                    if (type(cellValue) == float) and (cellValue % 1.0 == 0.0):
                        item = QtGui.QTableWidgetItem(str(int(cellValue)).decode('utf-8'))
                    else:
                        item = QtGui.QTableWidgetItem(str(cellValue).decode('utf-8'))
                    self.newExcelViewTableWidget.setItem(row, col, item)

    # 差异对比点击事件
    def excelDiffOperate(self, originExcelUrl, newExcelUrl):
        if originExcelUrl == '':
            print '旧Excel路径为空'.decode('utf-8')
            return
        if newExcelUrl == '':
            print '新Excel路径为空'.decode('utf-8')
            return

        self.refreshWindow()
        self.originExcelUrlLineEdit.setText(str(originExcelUrl).decode('utf-8'))
        self.newExcelUrlLineEdit.setText(str(newExcelUrl).decode('utf-8'))
        originExcel = xlrd.open_workbook(str(originExcelUrl).decode('utf-8'))
        newExcel = xlrd.open_workbook(str(newExcelUrl).decode('utf-8'))

        delSheetList, addSheetList, intersectionSheetList, excelDiffResult = self.findExcelDiff(originExcel,newExcel)
        # sheet增删结果显示
        self.sheetListLabel.setText(('共计新增 ' + str(len(addSheetList)) + ' Sheet' + ', 共计删除 '+ str(len(delSheetList)) + ' Sheet').decode('utf-8'))
        self.sheetAddLabel.setText(('共计新增 ' + str(len(addSheetList)) + ' Sheet').decode('utf-8'))
        self.sheetDelLabel.setText(('共计删除 '+ str(len(delSheetList)) + ' Sheet').decode('utf-8'))
        totalSheetList = intersectionSheetList + addSheetList + delSheetList
        self.sheetListTableWidget.setRowCount(len(totalSheetList))
        if len(addSheetList) == 0:
            self.sheetAddTableWidget.setRowCount(1)
        else:
            self.sheetAddTableWidget.setRowCount(len(addSheetList))
        if len(delSheetList) == 0:
            self.sheetDelTableWidget.setRowCount(1)
        else:
            self.sheetDelTableWidget.setRowCount(len(delSheetList))
        for row in range(0, len(totalSheetList)):
            sheetListItem = QtGui.QTableWidgetItem(str(totalSheetList[row]).decode('utf-8'))
            if sheetListItem.text() in addSheetList:
                sheetListItem.setForeground(QtGui.QBrush(QtGui.QColor(0,0,255)))
            if sheetListItem.text() in delSheetList:
                sheetListItem.setForeground(QtGui.QBrush(QtGui.QColor(255,0,0)))
            self.sheetListTableWidget.setItem(row, 0, sheetListItem)
        for row in range(0, len(addSheetList)):
            sheetAddItem = QtGui.QTableWidgetItem(str(addSheetList[row]).decode('utf-8'))
            sheetAddItem.setForeground(QtGui.QBrush(QtGui.QColor(0,0,255)))
            self.sheetAddTableWidget.setItem(row, 0, sheetAddItem)
        for row in range(0, len(delSheetList)):
            sheetDelItem = QtGui.QTableWidgetItem(str(delSheetList[row]).decode('utf-8'))
            sheetDelItem.setForeground(QtGui.QBrush(QtGui.QColor(255,0,0)))
            self.sheetDelTableWidget.setItem(row, 0, sheetDelItem)
        # 旧表，新表默认sheet差异结果显示在窗口中 （默认交集中第一个sheet，若无交集，则各自显示自己第一个sheet）
        if intersectionSheetList:
            self.showSheetDifferContent(intersectionSheetList[0], originExcelUrl, newExcelUrl)

    # 应用窗口复原（清空数据）
    def refreshWindow(self):
        self.originExcelUrlLineEdit.clear()
        self.newExcelUrlLineEdit.clear()
        self.originExcelViewTableWidget.clear()
        self.newExcelViewTableWidget.clear()
        self.sheetListTableWidget.clear()
        self.sheetAddTableWidget.clear()
        self.sheetDelTableWidget.clear()
        self.rowAddDelTableWidget.clear()
        self.columnAddDelTableWidget.clear()
        self.cellModifyTableWidget.clear()
        self.originExcelSheetLabel.setText(('旧Excel名:sheet名').decode('utf-8'))
        self.newExcelSheetLabel.setText(('新Excel名:sheet名').decode('utf-8'))
        self.sheetListLabel.setText(('Sheet增删情况').decode('utf-8'))
        self.sheetAddLabel.setText(('Sheet新增情况').decode('utf-8'))
        self.sheetDelLabel.setText(('Sheet删除情况').decode('utf-8'))
        self.rowAddDelLabel.setText(('行增删情况').decode('utf-8'))
        self.columnAddDelLabel.setText(('列增删情况').decode('utf-8'))
        self.cellModifyLabel.setText(('单元格改动情况').decode('utf-8'))
        self.originExcelViewTableWidget.setColumnCount(1)
        self.originExcelViewTableWidget.setHorizontalHeaderLabels(['A'])
        self.originExcelViewTableWidget.setRowCount(1)
        self.newExcelViewTableWidget.setColumnCount(1)
        self.newExcelViewTableWidget.setHorizontalHeaderLabels(['A'])
        self.newExcelViewTableWidget.setRowCount(1)
        self.sheetListTableWidget.setRowCount(1)
        self.sheetListTableWidget.setColumnCount(1)
        self.sheetListTableWidget.setHorizontalHeaderLabels(['Sheet名'.decode('utf-8')])
        self.sheetAddTableWidget.setRowCount(1)
        self.sheetAddTableWidget.setColumnCount(1)
        self.sheetAddTableWidget.setHorizontalHeaderLabels(['Sheet名'.decode('utf-8')])
        self.sheetDelTableWidget.setRowCount(1)
        self.sheetDelTableWidget.setColumnCount(1)
        self.sheetDelTableWidget.setHorizontalHeaderLabels(['Sheet名'.decode('utf-8')])
        self.rowAddDelTableWidget.setRowCount(2)
        self.rowAddDelTableWidget.setColumnCount(1)
        self.rowAddDelTableWidget.setVerticalHeaderLabels(['新增行'.decode('utf-8'),'删除行'.decode('utf-8')])
        self.columnAddDelTableWidget.setRowCount(2)
        self.columnAddDelTableWidget.setColumnCount(1)
        self.columnAddDelTableWidget.setVerticalHeaderLabels(['新增列'.decode('utf-8'),'删除列'.decode('utf-8')])
        self.cellModifyTableWidget.setRowCount(1)
        self.cellModifyTableWidget.setColumnCount(3)
        self.cellModifyTableWidget.setHorizontalHeaderLabels(['坐标'.decode('utf-8'),'旧值'.decode('utf-8'),'新值'.decode('utf-8')])

    # sheetTableWidget点击事件响应函数
    def sheetTableWidgetClicked(self, item):
        sheetName = item.text()
        originExcelUrl = self.originExcelUrlLineEdit.text()
        newExcelUrl = self.newExcelUrlLineEdit.text()
        self.showSheetDifferContent(sheetName, originExcelUrl, newExcelUrl)

    # Sheet内容及对比结果展示
    def showSheetDifferContent(self, sheetName, originExcelUrl, newExcelUrl):
        sheetName = str(sheetName).decode('utf-8')
        originExcel = xlrd.open_workbook(str(originExcelUrl).decode('utf-8'))
        newExcel = xlrd.open_workbook(str(newExcelUrl).decode('utf-8'))
        # 读取数据并获取两Excel差异
        originExcelData, originExcelSheetSortList = self.readExcelData(originExcel)
        newExcelData, newExcelSheetSortList = self.readExcelData(newExcel)
        delSheetList, addSheetList, intersectionSheetList, excelDiffResult = self.findExcelDiff(originExcel,newExcel)
        # 初始化界面
        self.originExcelSheetLabel.setText(('旧Excel名:sheet名').decode('utf-8'))
        self.newExcelSheetLabel.setText(('新Excel名:sheet名').decode('utf-8'))
        self.originExcelViewTableWidget.clear()
        self.newExcelViewTableWidget.clear()
        self.rowAddDelTableWidget.clear()
        self.columnAddDelTableWidget.clear()
        self.cellModifyTableWidget.clear()
        self.rowAddDelTableWidget.setRowCount(2)
        self.rowAddDelTableWidget.setColumnCount(1)
        self.rowAddDelTableWidget.setVerticalHeaderLabels(['新增行'.decode('utf-8'),'删除行'.decode('utf-8')])
        self.columnAddDelTableWidget.setRowCount(2)
        self.columnAddDelTableWidget.setColumnCount(1)
        self.columnAddDelTableWidget.setVerticalHeaderLabels(['新增列'.decode('utf-8'),'删除列'.decode('utf-8')])
        self.cellModifyTableWidget.setRowCount(1)
        self.cellModifyTableWidget.setColumnCount(3)
        self.cellModifyTableWidget.setHorizontalHeaderLabels(['坐标'.decode('utf-8'),'旧值'.decode('utf-8'),'新值'.decode('utf-8')])

        # 旧表，新表指定sheet差异结果显示在窗口中
        # 显示新旧表默认内容
        if sheetName in intersectionSheetList:
            # 初始化旧表Label
            originExcelUrlStr = originExcelUrl.split('/')
            originExcelName = originExcelUrlStr[len(originExcelUrlStr)-1]
            originExcelSheetLabelString = originExcelName + ':' + str(sheetName).decode('utf-8')
            self.originExcelSheetLabel.setText(originExcelSheetLabelString)
            # 初始化新表Label
            newExcelUrlStr = newExcelUrl.split('/')
            newExcelName = newExcelUrlStr[len(newExcelUrlStr)-1]
            newExcelSheetLabelString = newExcelName + ':' + str(sheetName).decode('utf-8')
            self.newExcelSheetLabel.setText(newExcelSheetLabelString)
            # 初始化旧Excel表格
            rowCount = len(originExcelData[sheetName])

            if originExcelData[sheetName]:
                columnCount = len(originExcelData[sheetName][0])
            else:
                columnCount = 0
            self.originExcelViewTableWidget.setRowCount(rowCount)
            self.originExcelViewTableWidget.setColumnCount(columnCount)
            columnHeaderLabelList = self.createExcelColumnHeaderLabelList(columnCount)
            self.originExcelViewTableWidget.setHorizontalHeaderLabels(columnHeaderLabelList)
            # 初始化新Excel表格（默认为sheet1中的数据）
            rowCount = len(newExcelData[sheetName])
            if newExcelData[sheetName]:
                columnCount = len(newExcelData[sheetName][0])
            else:
                columnCount = 0
            self.newExcelViewTableWidget.setRowCount(rowCount)
            self.newExcelViewTableWidget.setColumnCount(columnCount)
            columnHeaderLabelList = self.createExcelColumnHeaderLabelList(columnCount)
            self.newExcelViewTableWidget.setHorizontalHeaderLabels(columnHeaderLabelList)

            # 新旧表单元格交集显示（改动单元格标黄）
            originRowCount = len(originExcelData[sheetName])
            if originExcelData[sheetName]:
                originColCount = len(originExcelData[sheetName][0])
            else:
                originColCount = 0
            newRowCount = len(newExcelData[sheetName])
            if newExcelData[sheetName]:
                newColCount = len(newExcelData[sheetName][0])
            else:
                newColCount = 0

            modifyCellRowCount = len(excelDiffResult.get(sheetName).get('modifyCellMap'))
            if modifyCellRowCount == 0:
                modifyCellColCount = 0
            else:
                modifyCellColCount = len(excelDiffResult.get(sheetName).get('modifyCellMap')[0])

            cellModifyCount = 0
            for row in range(0, modifyCellRowCount):
                for col in range(0, modifyCellColCount):
                    originCellValue = originExcelData[sheetName][row][col]
                    newCellValue = newExcelData[sheetName][row][col]
                    # 数字处理，整形去掉小数点
                    originItem = self.digitItemHandel(originCellValue)
                    newItem = self.digitItemHandel(newCellValue)
                    # 单元格改动标黄
                    if excelDiffResult.get(sheetName).get('modifyCellMap')[row][col] == 1:
                        cellModifyCount += 1
                        originItem.setBackground(QtGui.QBrush(QtGui.QColor(255,215,0)))
                        newItem.setBackground(QtGui.QBrush(QtGui.QColor(255,215,0)))
                    self.originExcelViewTableWidget.setItem(row, col, originItem)
                    self.newExcelViewTableWidget.setItem(row, col, newItem)
            # 旧表删除行显示
            if originRowCount > modifyCellRowCount:
                for row in range(modifyCellRowCount, originRowCount):
                    for col in range(0, originColCount):
                        cellValue = originExcelData[sheetName][row][col]
                        # 数字处理，整形去掉小数点
                        item = self.digitItemHandel(cellValue)
                        item.setBackground(QtGui.QBrush(QtGui.QColor(255,182,193)))
                        self.originExcelViewTableWidget.setItem(row, col, item)
            # 旧表删除列显示
            if originColCount > modifyCellColCount:
                for row in range(0, originRowCount):
                    for col in range(modifyCellColCount, originColCount):
                        cellValue = originExcelData[sheetName][row][col]
                        # 数字处理，整形去掉小数点
                        item = self.digitItemHandel(cellValue)
                        item.setBackground(QtGui.QBrush(QtGui.QColor(255,182,193)))
                        self.originExcelViewTableWidget.setItem(row, col, item)
            # 新表增加行显示
            if newRowCount > modifyCellRowCount:
                for row in range(modifyCellRowCount, newRowCount):
                    for col in range(0, newColCount):
                        cellValue = newExcelData[sheetName][row][col]
                        # 数字处理，整形去掉小数点
                        item = self.digitItemHandel(cellValue)
                        item.setBackground(QtGui.QBrush(QtGui.QColor(135,206,250)))
                        self.newExcelViewTableWidget.setItem(row, col, item)
            # 新表增加列显示
            if newColCount > modifyCellColCount:
                for row in range(0, newRowCount):
                    for col in range(modifyCellColCount, newColCount):
                        cellValue = newExcelData[sheetName][row][col]
                        item = self.digitItemHandel(cellValue)
                        item.setBackground(QtGui.QBrush(QtGui.QColor(135,206,250)))
                        self.newExcelViewTableWidget.setItem(row, col, item)
            # 显示行列增删内容及单元格改动情况
            # 行增删情况
            self.rowAddDelLabel.setText(('共计新增 ' + str(len(excelDiffResult.get(sheetName).get('addRowList'))) + ' 行' + ', 共计删除 ' + str(len(excelDiffResult.get(sheetName).get('delRowList'))) + ' 行').decode('utf-8'))
            if excelDiffResult.get(sheetName).get('addRowList'):
                self.rowAddDelTableWidget.setColumnCount(len(excelDiffResult.get(sheetName).get('addRowList')))
                for col in range(0, len(excelDiffResult.get(sheetName).get('addRowList'))):
                    item = QtGui.QTableWidgetItem(str(excelDiffResult.get(sheetName).get('addRowList')[col]).decode('utf-8'))
                    item.setForeground(QtGui.QBrush(QtGui.QColor(0,0,255)))
                    self.rowAddDelTableWidget.setItem(0, col, item)
            elif excelDiffResult.get(sheetName).get('delRowList'):
                self.rowAddDelTableWidget.setColumnCount(len(excelDiffResult.get(sheetName).get('delRowList')))
                for col in range(0, len(excelDiffResult.get(sheetName).get('delRowList'))):
                    item = QtGui.QTableWidgetItem(str(excelDiffResult.get(sheetName).get('delRowList')[col]).decode('utf-8'))
                    item.setForeground(QtGui.QBrush(QtGui.QColor(255,0,0)))
                    self.rowAddDelTableWidget.setItem(1, col, item)
            # 列增删情况
            self.columnAddDelLabel.setText(('共计新增 ' + str(len(excelDiffResult.get(sheetName).get('addColList'))) + ' 列' + ', 共计删除 ' + str(len(excelDiffResult.get(sheetName).get('delColList'))) + ' 列').decode('utf-8'))
            if excelDiffResult.get(sheetName).get('addColList'):
                self.columnAddDelTableWidget.setColumnCount(len(excelDiffResult.get(sheetName).get('addColList')))
                for col in range(0, len(excelDiffResult.get(sheetName).get('addColList'))):
                    item = QtGui.QTableWidgetItem(str(excelDiffResult.get(sheetName).get('addColList')[col]).decode('utf-8'))
                    item.setForeground(QtGui.QBrush(QtGui.QColor(0,0,255)))
                    self.columnAddDelTableWidget.setItem(0, col, item)
            elif excelDiffResult.get(sheetName).get('delColList'):
                self.columnAddDelTableWidget.setColumnCount(len(excelDiffResult.get(sheetName).get('delColList')))
                for col in range(0, len(excelDiffResult.get(sheetName).get('delColList'))):
                    item = QtGui.QTableWidgetItem(str(excelDiffResult.get(sheetName).get('delColList')[col]).decode('utf-8'))
                    item.setForeground(QtGui.QBrush(QtGui.QColor(255,0,0)))
                    self.columnAddDelTableWidget.setItem(1, col, item)
            # 单元格改动
            self.cellModifyLabel.setText(('共计 ' + str(cellModifyCount) + ' 个单元格被改动').decode('utf-8'))
            if (cellModifyCount == 0):
                self.cellModifyTableWidget.setRowCount(1)
            else:
                self.cellModifyTableWidget.setRowCount(cellModifyCount)
            cellRowIndex = 0
            for row in range(0, len(excelDiffResult.get(sheetName).get('modifyCellMap'))):
                for col in range(0, len(excelDiffResult.get(sheetName).get('modifyCellMap')[row])):
                    if excelDiffResult.get(sheetName).get('modifyCellMap')[row][col] == 1:
                        rowIndex = row+1
                        colIndex = self.createExcelColumnHeaderLabel(col+1)
                        index = '(' + str(rowIndex) + ', ' + str(colIndex) + ')'
                        indexItem = QtGui.QTableWidgetItem(index)
                        indexItem.setForeground(QtGui.QBrush(QtGui.QColor(0,0,255)))

                        originCellValue = originExcelData[sheetName][row][col]
                        originCellValueItem = self.digitItemHandel(originCellValue)
                        newCellValue = newExcelData[sheetName][row][col]
                        newCellValueItem = self.digitItemHandel(newCellValue)

                        self.cellModifyTableWidget.setItem(cellRowIndex, 0, indexItem)
                        self.cellModifyTableWidget.setItem(cellRowIndex, 1, originCellValueItem)
                        self.cellModifyTableWidget.setItem(cellRowIndex, 2, newCellValueItem)
                        cellRowIndex += 1
        else:
            self.rowAddDelLabel.setText('行增删情况'.decode('utf-8'))
            self.columnAddDelLabel.setText('列增删情况'.decode('utf-8'))
            self.cellModifyLabel.setText('单元格改动情况'.decode('utf-8'))
            self.originExcelViewTableWidget.setColumnCount(1)
            self.originExcelViewTableWidget.setHorizontalHeaderLabels(['A'])
            self.originExcelViewTableWidget.setRowCount(1)
            self.newExcelViewTableWidget.setColumnCount(1)
            self.newExcelViewTableWidget.setHorizontalHeaderLabels(['A'])
            self.newExcelViewTableWidget.setRowCount(1)

        if sheetName in delSheetList:
            # 初始化旧表Label
            originExcelUrlStr = originExcelUrl.split('/')
            originExcelName = originExcelUrlStr[len(originExcelUrlStr)-1]
            originExcelSheetLabelString = originExcelName + ':' + str(sheetName).decode('utf-8')
            self.originExcelSheetLabel.setText(originExcelSheetLabelString)
            # 初始化新表Label
            newExcelUrlStr = newExcelUrl.split('/')
            newExcelName = newExcelUrlStr[len(newExcelUrlStr)-1]
            newExcelSheetLabelString = newExcelName + ':'
            self.newExcelSheetLabel.setText(str(newExcelSheetLabelString).decode('utf-8'))
            # 初始化旧Excel表格
            rowCount = len(originExcelData[sheetName])
            columnCount = len(originExcelData[sheetName][0])
            self.originExcelViewTableWidget.setRowCount(rowCount)
            self.originExcelViewTableWidget.setColumnCount(columnCount)
            columnHeaderLabelList = self.createExcelColumnHeaderLabelList(columnCount)
            self.originExcelViewTableWidget.setHorizontalHeaderLabels(columnHeaderLabelList)
            for row in range(0, rowCount):
                for col in range(0, columnCount):
                    # 数字处理，整形去掉小数点
                    cellValue = originExcelData[sheetName][row][col]
                    if (type(cellValue) == float) and (cellValue % 1.0 == 0.0):
                        item = QtGui.QTableWidgetItem(str(int(cellValue)).decode('utf-8'))
                    else:
                        item = QtGui.QTableWidgetItem(str(cellValue).decode('utf-8'))
                    self.originExcelViewTableWidget.setItem(row, col, item)

        if sheetName in addSheetList:
            # 初始化新表Label
            newExcelUrlStr = newExcelUrl.split('/')
            newExcelName = newExcelUrlStr[len(newExcelUrlStr)-1]
            newExcelSheetLabelString = newExcelName + ':' + str(sheetName).decode('utf-8')
            self.newExcelSheetLabel.setText(newExcelSheetLabelString)
            # 初始化旧表Label
            originExcelUrlStr = originExcelUrl.split('/')
            originExcelName = originExcelUrlStr[len(originExcelUrlStr)-1]
            originExcelSheetLabelString = originExcelName + ':'
            self.originExcelSheetLabel.setText(str(originExcelSheetLabelString).decode('utf-8'))
            # 初始化新Excel表格（默认为sheet1中的数据）
            rowCount = len(newExcelData[sheetName])
            columnCount = len(newExcelData[sheetName][0])
            self.newExcelViewTableWidget.setRowCount(rowCount)
            self.newExcelViewTableWidget.setColumnCount(columnCount)
            columnHeaderLabelList = self.createExcelColumnHeaderLabelList(columnCount)
            self.newExcelViewTableWidget.setHorizontalHeaderLabels(columnHeaderLabelList)
            for row in range(0, rowCount):
                for col in range(0, columnCount):
                    # 数字处理，整形去掉小数点
                    cellValue = newExcelData[sheetName][row][col]
                    if (type(cellValue) == float) and (cellValue % 1.0 == 0.0):
                        item = QtGui.QTableWidgetItem(str(int(cellValue)).decode('utf-8'))
                    else:
                        item = QtGui.QTableWidgetItem(str(cellValue).decode('utf-8'))
                    self.newExcelViewTableWidget.setItem(row, col, item)

    # 行定位函数
    def trackingRow(self, item):
        if not item.text():
            return
        row = int(item.text()) - 1
        self.originExcelViewTableWidget.verticalScrollBar().setSliderPosition(row)
        self.originExcelViewTableWidget.selectRow(row)
        self.newExcelViewTableWidget.verticalScrollBar().setSliderPosition(row)
        self.newExcelViewTableWidget.selectRow(row)

    # 列定位函数
    def trackingCol(self, item):
        if not item.text():
            return
        col = item.text()
        colNum = self.alphabetToNum(col)
        self.originExcelViewTableWidget.horizontalScrollBar().setSliderPosition(colNum)
        self.originExcelViewTableWidget.selectColumn(colNum)
        self.newExcelViewTableWidget.horizontalScrollBar().setSliderPosition(colNum)
        self.newExcelViewTableWidget.selectColumn(colNum)

    # 单元格改动定位函数
    def trackingCell(self, item):
        indexString =  str(self.cellModifyTableWidget.selectedItems()[0].text())
        indexList = indexString.split(',')
        row = int(indexList[0].strip().strip('(')) - 1
        col = self.alphabetToNum(indexList[1].strip().strip(')'))

        self.originExcelViewTableWidget.setCurrentCell(row,col)
        self.newExcelViewTableWidget.setCurrentCell(row,col)

    # 数字处理，整形去掉小数点
    def digitItemHandel(self, cellValue):
        if (type(cellValue) == float) and (cellValue % 1.0 == 0.0):
            item = QtGui.QTableWidgetItem(str(int(cellValue)).decode('utf-8'))
        else:
            item = QtGui.QTableWidgetItem(str(cellValue).decode('utf-8'))
        return item
    # 字母列标 转化为数字序号
    def alphabetToNum(self, alphabet):
        alphabet = str(alphabet)
        if len(alphabet) == 1:
            colNum = (ord(alphabet[0]) - 65) + 1
        elif len(alphabet) == 2:
            a = (ord(alphabet[0]) - 65) * 26
            b = (ord(alphabet[1]) - 65) + 1
            colNum = a + b + 26
        elif len(alphabet) == 3:
            a = (ord(alphabet[0]) - 65) * (26**2)
            b = (ord(alphabet[1]) - 65) * 26
            c = (ord(alphabet[2]) - 65) + 1
            colNum = a+b+c + 702
        return colNum - 1

    # 读取excel中的数据
    def readExcelData(self, excel):
        excelData = {}
        sheetSortList = []
        for sheet in excel.sheets():
            sheetSortList.append(sheet.name)
            table = excel.sheet_by_name(sheet.name)
            sheetData = [[None] * table.ncols for i in range(0, table.nrows)]
            for row in range(0, table.nrows):
                for column in range(0, table.ncols):
                    sheetData[row][column] = table.cell(row, column).value
            excelData[sheet.name] = sheetData
        return excelData, sheetSortList

    # 查找两Excel的差异
    def findExcelDiff(self, originExcel, targetExcel):
        originExcelData, originSheetSortList = self.readExcelData(originExcel)
        targetExceData, targetSheetSortList = self.readExcelData(targetExcel)
        # 两Excel的sheetName交集
        intersectionSheetList = [sheetName for sheetName in originSheetSortList if sheetName in targetSheetSortList]

        delSheetList = []  # 已删除sheet记录
        addSheetList = []  # 新增sheet记录
        for sheetName in originSheetSortList:
            if sheetName not in intersectionSheetList:
                delSheetList.append(sheetName)
        for sheetName in targetSheetSortList:
            if sheetName not in intersectionSheetList:
                addSheetList.append(sheetName)
        # printList(delSheetList)
        # printList(addSheetList)

        # 定义excel差异结果集
        excelDiffResult = {}
        #记录交集sheet中的差异
        for sheetName in intersectionSheetList:
            orginSheet = originExcel.sheet_by_name(sheetName)
            targetSheet = targetExcel.sheet_by_name(sheetName)

            # 确定原表与新表的边界
            rowBoundary = min(orginSheet.nrows, targetSheet.nrows)

            columnBoundary = min(orginSheet.ncols, targetSheet.ncols)

            # 定义sheet差异结果集
            sheetDiffResult = {}
            # 记录行列增删情况
            delRowList = []
            addRowList = []
            delColList = []
            addColList = []
            if orginSheet.nrows > rowBoundary:
                for rowIndex in range(rowBoundary+1, orginSheet.nrows+1):
                    delRowList.append(rowIndex)
            sheetDiffResult['delRowList'] = delRowList
            if targetSheet.nrows > rowBoundary:
                for rowIndex in range(rowBoundary+1, targetSheet.nrows+1):
                    addRowList.append(rowIndex)
            sheetDiffResult['addRowList'] = addRowList
            if orginSheet.ncols > columnBoundary:
                for colIndex in range(columnBoundary+1, orginSheet.ncols+1):
                    colLabel = self.createExcelColumnHeaderLabel(colIndex)
                    delColList.append(colLabel)
            sheetDiffResult['delColList'] = delColList
            if targetSheet.ncols > columnBoundary:
                for colIndex in range(columnBoundary+1, targetSheet.ncols+1):
                    colLabel = self.createExcelColumnHeaderLabel(colIndex)
                    addColList.append(colLabel)
            sheetDiffResult['addColList'] = addColList
            # 记录单元格修改情况
            modifyCellMap = [[0] * columnBoundary for i in range(0, rowBoundary)]
            for row in range(0, rowBoundary):
                for col in range(0, columnBoundary):
                    if orginSheet.cell(row, col).value != targetSheet.cell(row, col).value:
                        modifyCellMap[row][col] = 1

            # for row,col in zip(range(0, rowBoundary),range(0, columnBoundary)):
            #     if orginSheet.cell(row, col).value != targetSheet.cell(row, col).value:
            #         modifyCellMap[row][col] = 1
            sheetDiffResult['modifyCellMap'] = modifyCellMap
            # sheet差异结果集 以 sheetName为标记 存入excel差异结果集
            excelDiffResult[sheetName] = sheetDiffResult

        return delSheetList, addSheetList, intersectionSheetList, excelDiffResult

    # 生成Excel列标序列，Excel最大列数16384，最大列标XFD
    def createExcelColumnHeaderLabelList(self, columnNum):
        columnHeaderLabelList = []
        for a in range(1,columnNum+1):
            if a > 702:
                aa = a - 702
                first = chr(65 + (aa-1) / (26 ** 2))
                second = chr(65 + ((aa-1) % (26 ** 2))/ 26)
                third = chr(65 + ((aa-1) % 26))
                headerLabel = first + second + third
                columnHeaderLabelList.append(headerLabel)
            elif a > 26:
                aa = a - 26
                first = chr(65 + (aa-1) / 26)
                second = chr(65 + ((aa-1) % 26))
                headerLabel = first + second
                columnHeaderLabelList.append(headerLabel)
            else:
                first = chr(65 + a - 1)
                headerLabel = first
                columnHeaderLabelList.append(headerLabel)
        return columnHeaderLabelList

    # 生成Excel单个列标
    def createExcelColumnHeaderLabel(self, columnNum):
        a = columnNum
        headerLabel = ''
        if a > 702:
            aa = a - 702
            first = chr(65 + (aa-1) / (26 ** 2))
            second = chr(65 + ((aa-1) % (26 ** 2))/ 26)
            third = chr(65 + ((aa-1) % 26))
            headerLabel = first + second + third
        elif a > 26:
            aa = a - 26
            first = chr(65 + (aa-1) / 26)
            second = chr(65 + ((aa-1) % 26))
            headerLabel = first + second
        else:
            first = chr(65 + a - 1)
            headerLabel = first
        return headerLabel

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    main = Main();
    main.show()
    sys.exit(app.exec_())

