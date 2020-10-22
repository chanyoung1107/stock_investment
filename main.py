# -*- coding: utf-8 -*-
import logging
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5 import QtWidgets
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import sys
from logging.handlers import TimedRotatingFileHandler
import time


getCodeListBtn = ''
codeNameList = []
perList = []



class StockStart(QWidget):
    def __init__(self):
        super().__init__()
        self.start()

    def start(self):
        global getCodeListBtn

        self.setWindowTitle('Kiwoom Stock Investment')
        self.setFixedSize(400, 250)
        self.setFocusPolicy(Qt.StrongFocus)

        self.kiwoom = QAxWidget("KHOPENAPI.KHOpenAPICtrl.1")
        self.kiwoom.dynamicCall("CommConnect()")

        self.kiwoom.OnReceiveTrData.connect(self.receive_trdata)

        layout = QtWidgets.QVBoxLayout()
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        getCodeListBtn = QtWidgets.QPushButton('Get Code List')
        layout.addWidget(getCodeListBtn)

        self.setLayout(layout)
        self.show()

        self.kiwoom.OnEventConnect.connect(self.checkStatus)


    def checkStatus(self, err_code):
        global getCodeListBtn

        if err_code == 0:
            getCodeListBtn.clicked.connect(self.getCodeList)


    def getCodeList(self):
        ret = self.kiwoom.dynamicCall("GetCodeListByMarket(QString)", ["0"])

        kospiNameList = []
        kospiCodeList = ret.split(';')

        # for i in kospiCodeList:
        #     name = self.kiwoom.dynamicCall("GetMasterCodeName(QString)", [i])
        #     kospiNameList.append(i + " : " + name)

        for i in range(kospiCodeList):
            self.kiwoom.dynamicCall("SetInputValue(QString, QString)", "종목코드", kospiCodeList[i])
            self.kiwoom.dynamicCall("CommRqData(QString, QString, int, QString)", "opt10001_req", "opt10001", 0, "0101")
            time.sleep(0.2)


    def receive_trdata(self, screen_no, rqname, trcode, recordname, prev_next, data_len, err_code, msg1, msg2):
        global codeNameList
        global perList

        if rqname == 'opt10001_req':
            name = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "종목명")
            per = self.kiwoom.dynamicCall("CommGetData(QString, QString, QString, int, QString)", trcode, "", rqname, 0, "PER")

            print(name.strip())
            # codeNameList.append(name.strip())
            # perList.append(per.strip())
            #
            #
            # print(perList)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWindow = StockStart()
    myWindow.show()
    app.exec_()