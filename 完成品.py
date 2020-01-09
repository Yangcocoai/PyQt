
from PyQt5.Qt import *
import pandas as pd
import numpy as np


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

    def setUI(self,Mainwindows):
        """多窗口"""
        tabdemo = QTabWidget(Mainwindows)
        tabdemo.resize(900,750)
        tabdemo.move(25,80)

        self.tad2 = QWidget(tabdemo)
        tabdemo.addTab(self.tad2,"tab2")
        tabdemo.setTabText(0,"表格")
        grid = QGridLayout()
        self.tableWidget = QTableWidget(self.tad2)
        self.tableWidget.setColumnCount(10)
        self.tableWidget.setRowCount(10)
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
        ###===========读取表格，转换表格，===========================================
        if len(path_openfile_name) > 0:
            self.input_table = pd.read_excel(path_openfile_name)  # input_table 为数据

            input_table_rows = self.input_table.shape[0]  # input_table_rows = 行
            input_table_colunms = self.input_table.shape[1]  # input_table_colunms = 列

            # input_table_header = 第一行一般是标题头
            input_table_header = self.input_table.columns.values.tolist()

            ###===========读取表格，转换表格，============================================
            ###======================给tablewidget设置行列表头============================

            self.tableWidget.setColumnCount(input_table_colunms)
            self.tableWidget.setRowCount(input_table_rows)
            self.tableWidget.setHorizontalHeaderLabels(input_table_header)

            ###======================给tablewidget设置行列表头============================

            ###================遍历表格每个元素，同时添加到tablewidget中========================
            for i in range(input_table_rows):
                input_table_rows_values = self.input_table.iloc[[i]]
                # print(input_table_rows_values)
                input_table_rows_values_array = np.array(input_table_rows_values)
                input_table_rows_values_list = input_table_rows_values_array.tolist()[0]
                # print(input_table_rows_values_list)
                for j in range(input_table_colunms):
                    input_table_items_list = input_table_rows_values_list[j]
                    # print(input_table_items_list)
                    # print(type(input_table_items_list))

                    ###==============将遍历的元素添加到tablewidget中并显示=======================

                    input_table_items = str(input_table_items_list)
                    newItem = QTableWidgetItem(input_table_items)
                    newItem.setTextAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
                    self.tableWidget.setItem(i, j, newItem)

        ###================遍历表格每个元素，同时添加到tablewidget中========================
        else:
            self.centralWidget.show()

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

    def putton_tab(self):
        """按钮单击 把按钮的文本添加到表格中"""
        tabrow = self.tableWidget.currentRow()  # 获取点击的行
        tabcolunm = self.tableWidget.currentColumn() # 获取点击的列
        puttext = self.sender()
        if tabrow == -1: # 没点击 值是-1
            tabrow = 0
            tabcolunm = 0
            tabdemo_text = QTableWidgetItem(puttext.text())
            self.tableWidget.setItem(tabrow,tabcolunm,tabdemo_text)
        else:
            tabdemo_text = QTableWidgetItem(puttext.text())
            self.tableWidget.setItem(tabrow,tabcolunm,tabdemo_text)

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
        file_path = QFileDialog.getSaveFileName(self, "保存文件", "./",
                                                "Excel files(*.xlsx , *.xls)")
        print(file_path[0])




if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    dome = Widget()
    dome.show()
    sys.exit(app.exec_())

