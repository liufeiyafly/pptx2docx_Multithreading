# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'pptx2docx4.16.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets


# 下面3行用来解决pyqt5生成exe没有图片
# from wx import img as wxi
# from zz import img as zzi
# import base64
#
# with open('wx.jpg', 'wb') as wx:
#     wx.write(base64.b64decode(wxi))
# with open('zz.jpg', 'wb') as zz:
#     zz.write(base64.b64decode(zzi))


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setFixedSize(520, 305)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setGeometry(QtCore.QRect(0, 0, 518, 285))
        self.tabWidget.setObjectName("tabWidget")
        self.tab_1 = QtWidgets.QWidget()
        self.tab_1.setObjectName("tab_1")
        self.radioButton_2 = QtWidgets.QRadioButton(self.tab_1)
        self.radioButton_2.setGeometry(QtCore.QRect(240, 206, 120, 16))
        self.radioButton_2.setObjectName("radioButton_2")
        self.checkBox = QtWidgets.QCheckBox(self.tab_1)
        self.checkBox.setGeometry(QtCore.QRect(365, 195, 120, 16))
        self.checkBox.setObjectName("checkBox")
        self.pushButton = QtWidgets.QPushButton(self.tab_1)
        self.pushButton.setGeometry(QtCore.QRect(0, 190, 75, 31))
        self.pushButton.setObjectName("pushButton")
        self.textEdit = QtWidgets.QTextEdit(self.tab_1)
        self.textEdit.setGeometry(QtCore.QRect(2, 2, 201, 183))
        self.textEdit.setObjectName("textEdit")
        self.pushButton_2 = QtWidgets.QPushButton(self.tab_1)
        self.pushButton_2.setGeometry(QtCore.QRect(260, 225, 201, 31))
        self.pushButton_2.setObjectName("pushButton_2")
        self.pushButton_3 = QtWidgets.QPushButton(self.tab_1)
        self.pushButton_3.setGeometry(QtCore.QRect(128, 190, 75, 31))
        self.pushButton_3.setObjectName("pushButton_3")
        self.progressBar = QtWidgets.QProgressBar(self.tab_1)
        self.progressBar.setGeometry(QtCore.QRect(1, 228, 239, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.radioButton = QtWidgets.QRadioButton(self.tab_1)
        self.radioButton.setGeometry(QtCore.QRect(240, 190, 120, 16))
        self.radioButton.setObjectName("radioButton")
        self.tableWidget = QtWidgets.QTableWidget(self.tab_1)
        self.tableWidget.setGeometry(QtCore.QRect(205, 2, 305, 60))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(2)
        self.tableWidget.setRowCount(1)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget.setHorizontalHeaderItem(1, item)

        self.tableWidget.verticalHeader().setVisible(False)

        self.tableWidget.setColumnWidth(0, 152)
        self.tableWidget.setColumnWidth(1, 151)

        self.tableWidget_2 = QtWidgets.QTableWidget(self.tab_1)
        self.tableWidget_2.setGeometry(QtCore.QRect(205, 64, 305, 121))
        self.tableWidget_2.setObjectName("tableWidget_2")
        self.tableWidget_2.setColumnCount(3)
        self.tableWidget_2.setRowCount(0)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.tableWidget_2.setHorizontalHeaderItem(2, item)
        self.tableWidget_2.horizontalHeader().setDefaultSectionSize(100)
        self.tableWidget_2.horizontalHeader().setMinimumSectionSize(25)
        self.tableWidget_2.verticalHeader().setDefaultSectionSize(30)
        self.tableWidget_2.verticalHeader().setMinimumSectionSize(25)

        self.tableWidget_2.verticalHeader().setVisible(False)
        # self.tableWidget_2.setRowHeight(0, 50)  # 设置第一行高度
        self.tableWidget_2.setColumnWidth(0, 101)
        self.tableWidget_2.setColumnWidth(1, 101)
        self.tableWidget_2.setColumnWidth(2, 101)

        self.tabWidget.addTab(self.tab_1, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.groupBox_2 = QtWidgets.QGroupBox(self.tab_2)
        self.groupBox_2.setGeometry(QtCore.QRect(0, 3, 511, 77))
        self.groupBox_2.setObjectName("groupBox_2")
        self.label = QtWidgets.QLabel(self.groupBox_2)
        self.label.setGeometry(QtCore.QRect(10, 5, 400, 81))
        self.label.setObjectName("label")
        self.groupBox_3 = QtWidgets.QGroupBox(self.tab_2)
        self.groupBox_3.setGeometry(QtCore.QRect(0, 200, 511, 51))
        self.groupBox_3.setObjectName("groupBox_3")
        self.lineEdit = QtWidgets.QLineEdit(self.groupBox_3)
        self.lineEdit.setGeometry(QtCore.QRect(10, 20, 491, 20))
        self.lineEdit.setObjectName("lineEdit")
        self.groupBox_4 = QtWidgets.QGroupBox(self.tab_2)
        self.groupBox_4.setGeometry(QtCore.QRect(0, 87, 511, 106))
        self.groupBox_4.setObjectName("groupBox_4")
        self.label_3 = QtWidgets.QLabel(self.groupBox_4)
        self.label_3.setGeometry(QtCore.QRect(10, 17, 200, 82))
        self.label_3.setObjectName("label_3")
        self.groupBox_4.raise_()
        self.groupBox_3.raise_()
        self.groupBox_2.raise_()
        self.tabWidget.addTab(self.tab_2, "")
        self.tab_3 = QtWidgets.QWidget()
        self.tab_3.setObjectName("tab_3")
        self.groupBox = QtWidgets.QGroupBox(self.tab_3)
        self.groupBox.setGeometry(QtCore.QRect(0, 8, 511, 245))
        self.groupBox.setTitle("")
        self.groupBox.setObjectName("groupBox")
        self.tabWidget.addTab(self.tab_3, "")
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 520, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        self.label_zhifubao = QtWidgets.QLabel(self.groupBox)
        self.label_zhifubao.setGeometry(QtCore.QRect(10, 2, 240, 240))
        self.label_zhifubao.setObjectName("label_zhifubao")
        self.label_zhifubao.setPixmap(QtGui.QPixmap('zz.jpg'))
        self.label_wx = QtWidgets.QLabel(self.groupBox)
        self.label_wx.setGeometry(QtCore.QRect(260, 2, 240, 240))
        self.label_wx.setObjectName("label_wx")
        self.label_wx.setPixmap(QtGui.QPixmap('wx.jpg'))

        self.retranslateUi(MainWindow)
        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "PPTX2DOCX"))
        self.radioButton_2.setText(_translate("MainWindow", "回车分隔段落"))
        self.checkBox.setText(_translate("MainWindow", "生成word文档"))
        self.pushButton.setText(_translate("MainWindow", "选择文件"))
        self.pushButton_2.setText(_translate("MainWindow", "开始"))
        self.pushButton_3.setText(_translate("MainWindow", "清空页面"))
        self.radioButton.setText(_translate("MainWindow", "空格分隔段落"))
        item = self.tableWidget.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "选择文件数"))
        item = self.tableWidget.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "完成文件数"))
        item = self.tableWidget_2.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "文件名"))
        item = self.tableWidget_2.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "中朝字数"))
        item = self.tableWidget_2.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "非中文单词"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_1), _translate("MainWindow", "提取"))
        self.groupBox_2.setTitle(_translate("MainWindow", "使用说明"))
        self.label.setText(_translate("MainWindow", "1.如果提取不了，可尝试将文件放在没有空格的文件夹下！\n"
                                                    "\n"
                                                    "注：生成Docx文件保存到源文件所在位置。"))
        self.groupBox_3.setTitle(_translate("MainWindow", "ppt批量转pptx懒人教程："))
        self.lineEdit.setText(_translate("MainWindow", "https://blog.csdn.net/wayne1000/article/details/104376239"))
        self.groupBox_4.setTitle(_translate("MainWindow", "关于"))
        self.label_3.setText(_translate("MainWindow", "当前版本：1.1.0\n"
                                                      "\n"
                                                      "更新时间：2020/04/16\n"
                                                      "\n"
                                                      "版权所有：萌萌的维尼"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "说明"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), _translate("MainWindow", "赞助"))
