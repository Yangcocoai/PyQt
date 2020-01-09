
from PyQt5.Qt import *
import pandas as pd
import numpy as np
import xlrd

class Pushbotton(QPushButton):
    """
    重写 父类 点击 和双击 双操作
    """
    doubleClicked = pyqtSignal()
    clicked = pyqtSignal()

    def __init__(self, *args, **kwargs):
        QPushButton.__init__(self, *args, **kwargs)
        self.timer = QTimer()
        self.timer.setSingleShot(True)
        self.timer.timeout.connect(self.clicked.emit)
        super().clicked.connect(self.checkDoubleClick)

    @pyqtSlot()
    def checkDoubleClick(self):
        if self.timer.isActive():
            self.doubleClicked.emit()
            self.timer.stop()
        else:
            self.timer.start(250)




class excelreader():
    """
    xlrd: 对象 返回出入表格的表题、行列、数据
    """
    def __init__(self,file_name):

        self.xls = xlrd.open_workbook(file_name)

    def sheems_names(self):
        """
        返回表格的表题
        :return:
        """
        return self.xls.sheet_names()

    def sheet_size(self,sheet_name):
        """
        返回 表名对应的行列
        (行，列)
        """
        if not sheet_name in self.sheems_names():
            return 0,0
        sheet = self.xls.sheet_by_name(sheet_name)
        return sheet.nrows,sheet.ncols

    def sheet_content(self,sheet_name):
        """
        返回 表名对应的数据
        [[行],[行]]
        """
        if not sheet_name in self.sheems_names():
            return []
        sheet = self.xls.sheet_by_name(sheet_name)
        nr,nc = sheet.nrows,sheet.ncols
        return [[sheet.cell(r, c).value for c in range(nc)] for r in range(nr)]


class Widget(QMainWindow):

    def __init__(self,parent=None):
        super().__init__(parent)
        self.resize(950,850)
        self.setUI(self)
        self.Menu()

        self.puttonrow = 170
        # 添加按钮
        self.addtoput = QPushButton(self)
        self.addtoput.setText("+")
        self.addtoput.move(self.puttonrow, 50)
        self.addtoput.clicked.connect(self.Putton)

    def Menu(self):
        """菜单"""
        tb = self.addToolBar("打开")
        open = QAction("打开",self)
        open.triggered.connect(self.openfile)
        open.triggered.connect(self.creat_table_show)
        tb.addAction(open)
        save = QAction("保存",self)
        save.triggered.connect(self.savecao)
        tb.addAction(save)

    def setUI(self,Mainwindows):
        """多窗口"""
        self.tabdemo = QTabWidget(Mainwindows)
        self.tabdemo.resize(900,750)
        self.tabdemo.move(25,80)
        self.biaogeQTab("表格",10,10)

    def biaogeQTab(self,biao_name,row,column,data=None):
        """
        QTableWidget 表格 QGridLayout排版
        :param biao_name: 表题
        :param row: 行数
        :param column: 列数
        :param data: 打开表格文件传入的数据
        :return:
        """
        self.tad2 = QWidget(self.tabdemo)
        self.tabdemo.addTab(self.tad2,biao_name)
        self.tabdemo.setTabText(0,biao_name)
        grid = QGridLayout()
        self.tableWidget = QTableWidget(self.tad2)
        self.tableWidget.setColumnCount(column)
        self.tableWidget.setRowCount(row)
        if data:
            for i in range(len(data)):
                for a in range(len(data[i])):
                    self.tableWidget.setItem(i,a,QTableWidgetItem(data[i][a]))

        grid.addWidget(self.tableWidget)
        self.tad2.setLayout(grid)

    def openfile(self):

        ###获取路径===================================================================

        openfile_name = QFileDialog.getOpenFileName(self,'选择文件','','Excel files(*.xlsx , *.xls)')

        # print(openfile_name) # 获取用户选择的数据 openfile_name[0] = 文件路径
        global path_openfile_name

        ###获取路径====================================================================

        path_openfile_name = openfile_name[0]

    def creat_table_show(self):
        '''
        读取表格,显示在: tableWidget
        '''
        Xlrd = excelreader(path_openfile_name)
        sheems_names = Xlrd.sheems_names()
        for x in sheems_names:
            row,column = Xlrd.sheet_size(x)
            biaogedata = Xlrd.sheet_content(x)

            self.biaogeQTab(x,row,column,biaogedata)

    def putton_tab(self):
        """按钮单击 把按钮的文本添加到表格中"""
        tabrow = self.tableWidget.currentRow()  # 获取点击的行
        tabcolunm = self.tableWidget.currentColumn() # 获取点击的列
        print(tabrow,tabcolunm)
        puttext = self.sender()
        if tabrow == -1: # 没点击 值是-1
            tabrow = 0
            tabcolunm = 0
            tabdemo_text = QTableWidgetItem(puttext.text())
            self.tableWidget.setItem(tabrow,tabcolunm,QTableWidgetItem(tabdemo_text))
        else:
            tabdemo_text = QTableWidgetItem(puttext.text())
            self.tableWidget.setItem(tabrow,tabcolunm,QTableWidgetItem(tabdemo_text))

    def edit_putton(self):
        """按钮双击 编辑按钮的文本"""
        puttext = self.sender()
        edit_lable = QInputDialog()
        puttonname0, _ = edit_lable.getText(self,"编辑","编辑按钮的额文本")
        if puttonname0:
            puttext.setText(puttonname0)

    def Putton(self):
        """操作按钮，和表格交互的按钮"""
        self.puttonrow -= 120
        self.putton = Pushbotton(self)
        self.putton.setText("点击点击点击")
        self.putton.move(self.puttonrow,50)
        self.putton.setFixedSize(100,28)
        self.putton.clicked.connect(self.putton_tab)
        self.putton.doubleClicked.connect(self.edit_putton)
        self.puttonrow += 120
        self.addtoput.move(self.puttonrow, 50)
        self.puttonrow += 120
        self.putton.show()

    def savecao(self):
        print("保存")

        A = self.tableWidget.item(0, 0).text()  # 获取某行某列item中的x信息
        print(A)



        # file_path = QFileDialog.getSaveFileName(self, "保存文件", "./",
        #                                         "Excel files(*.xlsx , *.xls)")
        # print(file_path[0])




if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    dome = Widget()
    dome.show()
    sys.exit(app.exec_())

