# -*- coding: utf-8 -*-
import os

import xlrd
from PyQt5 import QtCore, QtWidgets
from scrapy.cmdline import execute

from ExcelCrawl.utils.common import get_desktop


class Ui_MainWindow(object):
    sheet_index = 0  # 第几个表
    filename = ""
    urlcol = 0
    savecol = -1
    hastitile = True

    mainWindow = None

    has_proxy = False  # 设置默认是否使用代理IP

    download_time = 1.5  # 设置默认加载时间(单位：秒)

    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(408, 433)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.horizontalLayout = QtWidgets.QHBoxLayout(self.centralwidget)
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.cb_sheet_name = QtWidgets.QComboBox(self.centralwidget)
        self.cb_sheet_name.setEnabled(False)
        self.cb_sheet_name.setObjectName("cb_sheet_name")
        self.gridLayout_2.addWidget(self.cb_sheet_name, 1, 2, 1, 2)
        self.ckb_hastitile = QtWidgets.QCheckBox(self.centralwidget)
        self.ckb_hastitile.setEnabled(False)
        self.ckb_hastitile.setText("")
        self.ckb_hastitile.setObjectName("ckb_hastitile")
        self.gridLayout_2.addWidget(self.ckb_hastitile, 4, 2, 1, 1)
        self.pb_start = QtWidgets.QPushButton(self.centralwidget)
        self.pb_start.setEnabled(False)
        self.pb_start.setAutoRepeatDelay(300)
        self.pb_start.setAutoRepeatInterval(100)
        self.pb_start.setAutoDefault(False)
        self.pb_start.setDefault(False)
        self.pb_start.setFlat(False)
        self.pb_start.setObjectName("pb_start")
        self.gridLayout_2.addWidget(self.pb_start, 7, 1, 1, 2)
        self.dsb_download_delay = QtWidgets.QDoubleSpinBox(self.centralwidget)
        self.dsb_download_delay.setEnabled(False)
        self.dsb_download_delay.setObjectName("dsb_download_delay")
        self.gridLayout_2.addWidget(self.dsb_download_delay, 5, 2, 1, 1)
        self.let_path_in = QtWidgets.QLineEdit(self.centralwidget)
        self.let_path_in.setEnabled(False)
        self.let_path_in.setObjectName("let_path_in")
        self.gridLayout_2.addWidget(self.let_path_in, 0, 2, 1, 1)
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setAlignment(QtCore.Qt.AlignCenter)
        self.label_7.setObjectName("label_7")
        self.gridLayout_2.addWidget(self.label_7, 3, 0, 1, 1)
        self.cb_out_col = QtWidgets.QComboBox(self.centralwidget)
        self.cb_out_col.setEnabled(False)
        self.cb_out_col.setObjectName("cb_out_col")
        self.gridLayout_2.addWidget(self.cb_out_col, 3, 2, 1, 2)
        self.pb_path_in = QtWidgets.QPushButton(self.centralwidget)
        self.pb_path_in.setObjectName("pb_path_in")
        self.gridLayout_2.addWidget(self.pb_path_in, 0, 3, 1, 1)
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setAlignment(QtCore.Qt.AlignCenter)
        self.label_8.setObjectName("label_8")
        self.gridLayout_2.addWidget(self.label_8, 4, 0, 1, 1)
        self.ckb_proxy_ip = QtWidgets.QCheckBox(self.centralwidget)
        self.ckb_proxy_ip.setEnabled(False)
        self.ckb_proxy_ip.setText("")
        self.ckb_proxy_ip.setObjectName("ckb_proxy_ip")
        self.gridLayout_2.addWidget(self.ckb_proxy_ip, 6, 2, 1, 1)
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setAlignment(QtCore.Qt.AlignCenter)
        self.label.setObjectName("label")
        self.gridLayout_2.addWidget(self.label, 0, 0, 1, 1)
        self.cb_url_col = QtWidgets.QComboBox(self.centralwidget)
        self.cb_url_col.setEnabled(False)
        self.cb_url_col.setObjectName("cb_url_col")
        self.gridLayout_2.addWidget(self.cb_url_col, 2, 2, 1, 2)
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setAlignment(QtCore.Qt.AlignCenter)
        self.label_6.setObjectName("label_6")
        self.gridLayout_2.addWidget(self.label_6, 2, 0, 1, 1)
        self.label_9 = QtWidgets.QLabel(self.centralwidget)
        self.label_9.setAlignment(QtCore.Qt.AlignCenter)
        self.label_9.setObjectName("label_9")
        self.gridLayout_2.addWidget(self.label_9, 5, 0, 1, 1)
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setAlignment(QtCore.Qt.AlignCenter)
        self.label_3.setObjectName("label_3")
        self.gridLayout_2.addWidget(self.label_3, 1, 0, 1, 1)
        self.label_10 = QtWidgets.QLabel(self.centralwidget)
        self.label_10.setAlignment(QtCore.Qt.AlignCenter)
        self.label_10.setObjectName("label_10")
        self.gridLayout_2.addWidget(self.label_10, 6, 0, 1, 1)
        self.horizontalLayout.addLayout(self.gridLayout_2)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 408, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.ckb_proxy_ip.setObjectName("ckb_proxy_ip")
        self.gridLayout_2.addWidget(self.ckb_proxy_ip, 6, 2, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 408, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.ckb_hastitile.setChecked(self.hastitile)
        self.ckb_proxy_ip.setChecked(self.has_proxy)
        self.dsb_download_delay.setProperty("value", self.download_time)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "新闻爬取"))
        self.pb_start.setText(_translate("MainWindow", "开始爬取"))
        self.let_path_in.setText(_translate("MainWindow", "无"))
        self.label_7.setText(_translate("MainWindow", "结果保存所在列:"))
        self.pb_path_in.setText(_translate("MainWindow", "浏览"))
        self.label_8.setText(_translate("MainWindow", "Excel首行是否标题:"))
        self.label.setText(_translate("MainWindow", "导入Excel文件路径:"))
        self.label_6.setText(_translate("MainWindow", "链接所在列:"))
        self.label_9.setText(_translate("MainWindow", "每个页面加载时间(秒):"))
        self.label_3.setText(_translate("MainWindow", "所在表名:"))
        self.label_10.setText(_translate("MainWindow", "是否使用代理IP:"))

        self._translate = _translate
        self.mainWindow = MainWindow
        self.pb_path_in.clicked.connect(lambda: self.wichBtn(self.pb_path_in))
        self.pb_start.clicked.connect(lambda: self.wichBtn(self.pb_start))
        self.ckb_hastitile.clicked.connect(lambda: self.wichBtn(self.ckb_hastitile))

        self.cb_url_col.currentIndexChanged[int].connect(self.sele_url_col)
        self.cb_out_col.currentIndexChanged[int].connect(self.sele_out_col)
        self.cb_sheet_name.currentIndexChanged[int].connect(self.sele_sheet_name)

    # 链接列下拉框触发
    def sele_url_col(self, index):
        self.urlcol = index
        # print("urlcol=" + str(self.urlcol))

    # 结果保存列下拉框触发
    def sele_out_col(self, index):
        self.savecol = index - 1
        # self.savecol=-1为添加到新的列
        # print("savecol=" + str(self.savecol))

    # 选择表名后触发
    def sele_sheet_name(self, index):
        self.sheet_index = index
        self.sheet = self.worksheet.sheet_by_name(self.sheet_names[index])
        self.cb_url_col.clear()
        self.cb_out_col.clear()

        if (self.sheet.nrows >= 1):
            self.cb_out_col.addItem("添加到新的列")
            if (self.hastitile):
                ncols_list = [str(i) for i in self.sheet.row_values(0)]
            else:
                ncols_list = ["第" + str(i + 1) + "列" for i in range(self.sheet.ncols)]
            self.cb_url_col.addItems(ncols_list)
            self.cb_out_col.addItems(ncols_list)

            self.cb_url_col.setEnabled(True)
            self.cb_out_col.setEnabled(True)
            self.pb_start.setEnabled(True)

            self.ckb_proxy_ip.setEnabled(True)
            self.dsb_download_delay.setEnabled(True)

        else:
            self.cb_url_col.setEnabled(False)
            self.cb_out_col.setEnabled(False)

            self.pb_start.setEnabled(False)

            self.ckb_proxy_ip.setEnabled(False)
            self.dsb_download_delay.setEnabled(False)

    def wichBtn(self, btn):
        if (btn is self.pb_path_in):
            # 默认可加：directory=r'D:\新建文件夹'
            filename = QtWidgets.QFileDialog.getOpenFileName(parent=None, caption="请选择Excel文件",
                                                             # directory=get_desktop(),
                                                             directory=os.getcwd(),
                                                             filter="Excel files(*.xlsx *.xls *.csv)")[0]
            if ("" != filename):
                self.filename = filename
                # 初始化
                self.sheet_index = 0  # 第几个表
                self.urlcol = 0
                self.savecol = -1
                self.hastitile = True
                self.ckb_hastitile.setChecked(self.hastitile)

                self.let_path_in.setText(self._translate("MainWindow", self.filename))
                self.worksheet = xlrd.open_workbook(self.filename)  # 打开Excel文件
                self.sheet_names = self.worksheet.sheet_names()  # 获取所有表名

                self.cb_sheet_name.clear()

                if (len(self.sheet_names) >= 1):
                    # 选择所在表（默认第一张表）
                    self.cb_sheet_name.addItems(self.sheet_names)
                    self.cb_sheet_name.setEnabled(True)
                    self.ckb_hastitile.setEnabled(True)
                    self.sele_sheet_name(self.sheet_index)
                else:
                    self.cb_url_col.clear()
                    self.cb_out_col.clear()
                    self.ckb_hastitile.setEnabled(False)
                    self.cb_sheet_name.setEnabled(False)
                    self.cb_url_col.setEnabled(False)
                    self.cb_out_col.setEnabled(False)

                    self.pb_start.setEnabled(False)
                    self.ckb_proxy_ip.setEnabled(False)
                    self.dsb_download_delay.setEnabled(False)

        # 开始检测点击
        elif (btn is self.pb_start):
            # self.urlcol=0
            # self.savecol=-1为添加到新的列

            if (-1 == self.savecol):
                self.savecol = self.sheet.ncols

            self.has_proxy = self.ckb_proxy_ip.isChecked()
            self.download_time = self.dsb_download_delay.value()

            self.mainWindow.close()

            execute(['scrapy', 'crawl', 'news_spider',
                     "-a", "filename={}".format(self.filename),
                     "-a", "sheet_index={}".format(self.sheet_index),
                     "-a", "urlcol={}".format(self.urlcol),
                     "-a", "savecol={}".format(self.savecol),
                     "-a", "hastitile={}".format(self.hastitile),
                     "-a", "has_proxy={}".format(self.has_proxy),
                     "-a", "download_time={}".format(self.download_time)])






        # 单击Excel首行是否标题
        elif (btn is self.ckb_hastitile):
            self.hastitile = bool(1 - self.hastitile)
            if (None is not self.sheet_names and len(self.sheet_names) >= 1):
                self.sele_sheet_name(self.sheet_index)

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     mainWindow = QMainWindow()
#     ui = Ui_MainWindow()
#     ui.setupUi(mainWindow)
#     mainWindow.show()
#     sys.exit(app.exec_())
