# ！/usr//bin/env python3
# -*- coding:utf-8 -*-
'''
@Author: wayne
@Time: 2020/3/17 0017 下午 4:11
'''

# Form implementation generated from reading ui file 'wayne_new.ui'
#
# Created by: PyQt5 UI code generator 5.13.0
#
# WARNING! All changes made in this file will be lost!


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QCheckBox


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setEnabled(True)
        MainWindow.setFixedSize(500, 478)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        MainWindow.setSizePolicy(sizePolicy)
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(9)  # 编辑框textEdit里面的文字大小
        MainWindow.setFont(font)
        MainWindow.setMouseTracking(False)
        MainWindow.setAcceptDrops(False)
        MainWindow.setAutoFillBackground(True)
        MainWindow.setStyleSheet("")
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        # self.pushButton.setGeometry(QtCore.QRect(310, 107, 55, 22))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(9)
        self.pushButton.setFont(font)
        self.pushButton.setObjectName("pushButton")
        self.label = QtWidgets.QLabel(self.centralwidget)
        # self.label.setGeometry(QtCore.QRect(300, 140, 150, 21))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(9)
        self.label.setFont(font)
        self.label.setObjectName("label")

        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        # self.label_2.setGeometry(QtCore.QRect(10, 140, 250, 21))
        font = QtGui.QFont()
        font.setFamily("宋体")
        font.setPointSize(9)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.line = QtWidgets.QFrame(self.centralwidget)
        # self.line.setGeometry(QtCore.QRect(10, 130, 431, 16))
        self.line.setFrameShape(QtWidgets.QFrame.HLine)
        self.line.setFrameShadow(QtWidgets.QFrame.Sunken)
        self.line.setObjectName("line")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        # self.pushButton_2.setGeometry(QtCore.QRect(310, 78, 55, 23))
        font = QtGui.QFont()
        font.setFamily("Arial")
        self.pushButton_2.setFont(font)
        self.pushButton_2.setStyleSheet("")
        self.pushButton_2.setObjectName("pushButton_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        # self.label_3.setGeometry(QtCore.QRect(300, 10, 200, 70))
        font = QtGui.QFont()
        font.setFamily("宋体")
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        # self.textEdit.setGeometry(QtCore.QRect(10, 10, 281, 91))
        self.textEdit.setObjectName("textEdit")

        # 修改
        font = QtGui.QFont()
        font.setFamily("宋体")
        self.textBrowser = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser.setObjectName("textBrowser")

        font = QtGui.QFont()
        font.setFamily("宋体")
        self.textBrowser2 = QtWidgets.QTextBrowser(self.centralwidget)
        self.textBrowser2.setObjectName("textBrowser2")

        self.webtext = QtWidgets.QTextEdit(self.centralwidget)
        self.webtext.setText("https://blog.csdn.net/wayne1000/article/details/104376239")
        self.weblabel = QtWidgets.QLabel(self.centralwidget)
        self.weblabel.setText("ppt批量转pptx懒人教程：")
        # 大小
        self.textEdit.setGeometry(QtCore.QRect(10, 10, 300, 90))  # 文件信息
        self.line.setGeometry(QtCore.QRect(10, 340, 478, 16))  # 线
        self.textBrowser.setGeometry(QtCore.QRect(10, 110, 300, 38))  # 结果
        self.textBrowser2.setGeometry(QtCore.QRect(10, 160, 470, 100))  # 结果之字数统计

        self.label_2.setGeometry(QtCore.QRect(10, 350, 250, 21))  # version
        self.label.setGeometry(QtCore.QRect(350, 350, 150, 21))  # 作者

        self.label_3.setGeometry(QtCore.QRect(10, 270, 500, 70))  # 说明

        self.pushButton.setGeometry(QtCore.QRect(320, 125, 55, 22))  # 开始
        self.pushButton_2.setGeometry(QtCore.QRect(320, 78, 55, 23))  # ...

        self.webtext.setGeometry(QtCore.QRect(10, 421, 470, 30))
        self.weblabel.setGeometry(QtCore.QRect(10, 391, 500, 30))
        # 复选框
        self.checkbox = QCheckBox('生成.docx文件', self)
        self.checkbox.toggle()
        self.checkbox.setGeometry(QtCore.QRect(321, 100, 200, 23))
        # self.checkbox.move(321,100)

        # font = QtGui.QFont()
        # font.setFamily("黑体")
        # self.label

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 450, 23))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        # self.pushButton.clicked.connect(MainWindow.start)
        # self.pushButton_2.clicked.connect(MainWindow.choose_files)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "PPTX To DOCX"))
        self.pushButton.setText(_translate("MainWindow", "开始"))
        self.label.setText(_translate("MainWindow", "作者：萌萌的小维尼"))
        self.label_2.setText(_translate("MainWindow", "仅支持pptx格式 | version :1.0.1"))
        self.pushButton_2.setText(_translate("MainWindow", "..."))

        # 修改

        self.label_3.setText(_translate("MainWindow",
                                        "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0//EN\" \"http://www.w3.org/TR/REC-html40/strict.dtd\">\n"
                                        "<html><head><meta name=\"qrichtext\" content=\"1\" /><style type=\"text/css\">\n"
                                        "p, li { white-space: pre-wrap; }\n"
                                        "</style></head><body style=\" font-family:\'SimSun\'; font-size:9pt; font-weight:400; font-style:normal;\">\n"
                                        "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">1.打开文件(可多个)</p>\n"  #
                                        "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">2.支持拖拽单/多个文件到最上的控件内</p>\n"
                                        "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">文件夹名不要有空格~</p>\n"
                                        "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\"></p>\n"
                                        "<p style=\" margin-top:0px; margin-bottom:0px; margin-left:0px; margin-right:0px; -qt-block-indent:0; text-indent:0px;\">注：生成Docx文件保存到源文件所在位置。</p></body></html>"))
