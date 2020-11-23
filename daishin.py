# -*- coding: utf-8 -*-
import logging
from PyQt5 import QtCore
from PyQt5 import QtGui
from PyQt5 import QtWidgets
from PyQt5.QtCore import *
import win32com.client
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QAxContainer import *
import sys
import time


getCodeListBtn = ''
codeNameList = []
perList = []



class StockStart(QWidget):

    def __init__(self):
        stokeType = []  # 업종 코드 리스트

        # CpCybos - CYBOS의 각종 상태를 확인할 수 있음. (모듈 위치: CpUtil.dll)
        cybos = win32com.client.Dispatch("CpUtil.CpCybos")
        # print(cybos.IsConnect)           # 연결상태 확인
        # 주식 종록에 대한 정보 확인
        cpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
        # print(cpStockCode.GetCount())        # 주식 상장(비상장 일부 포함) 갯수
        # print(cpStockCode.GetData(1, 1))     # 주식 종목(0: 종목코드, 1: 종목명, 2:둘다) / 인자값
        # 여러 종목의 필요 항목을 한번에 수신
        marketEye = win32com.client.Dispatch("CpSysDib.MarketEye")
        # 업종별 코드 리스트
        cpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")

        # typeList에 주식 종목코드 넣기
        for i in range(cpStockCode.GetCount()):
            stokeType.append(cpStockCode.GetData(0, i))

        fullData = []
        dataArr = []
        division = int(len(stokeType) / 60) + 1

        for i in range(len(stokeType)):
            dataArr.append(stokeType[i])
            if i % 60 == 0 and i != 0:
                fullData.append(dataArr)
                dataArr = []
            if i == len(stokeType) - 1:
                fullData.append(dataArr)


        # for i in range(len(fullData)):
        #     print(len(fullData[i]))

        # print(division)
        # print(len(fullData))
        # print(len(stokeType))
        # print(fullData)

        # for i in range(len(stokeType)):           # 너무 오래걸림
        #     marketEye.SetInputValue(0, 67)
        #     marketEye.SetInputValue(1, stokeType[i])
        #     marketEye.BlockRequest()
        #     print("PER " + str(i) + " : " + str(marketEye.getDataValue(0, 0)))  # 필드의 인자값, 종목의 인자값
        for i in range(len(fullData)):
            # print(fullData[i])
            # print(cpStockCode.GetData(1, 0))

            marketEye.SetInputValue(0, (20, 67, 75, 88, 89))
            marketEye.SetInputValue(1, fullData[i])
            marketEye.BlockRequest()

            # for j in range(len(fullData[i])):
                # print(cpStockCode.GetData(1, ))
                # print("총상장주식수 : " + str(marketEye.getDataValue(0, 0)))             # 4
                # print("PER : " + str(marketEye.getDataValue(1, 0)))                    # 20
                # print("부채비율 : " + str(marketEye.getDataValue(2, 0)))                 # 67
                # print("당기순이익 : " + str(marketEye.getDataValue(3, 0)))               # 70
                # print("BPS(주당순자산) : " + str(marketEye.getDataValue(4, 0)))          # 88
                # print(marketEye.getDataValue(0, 0) / marketEye.getDataValue(3, 0))     # correct!!
                # print(marketEye.getDataValue(0, 0) * marketEye.getDataValue(4, 0))

        # industryCodeList = cpCodeMgr.GetIndustryList()  # 업종별 리스트 호출
        # for industryCode in industryCodeList:
        #     print(industryCode, cpCodeMgr.GetIndustryName(industryCode))


if __name__ == '__main__':
    # app = QApplication(sys.argv)
    # myWindow = StockStart()
    # myWindow.show()
    # app.exec_()
    StockStart()