import sys
from PyQt5.QtWidgets import *
import win32com.client

class Cp7043:
    def __init__(self):
        # 통신 OBJECT 기본 세팅
        self.objRq = win32com.client.Dispatch("CpSysDib.CpSvrNew7043")
        self.objRq.SetInputValue(0, ord('0')) # 거래소 + 코스닥
        self.objRq.SetInputValue(1, ord('2'))  # 상승
        self.objRq.SetInputValue(2, ord('1'))  # 당일
        self.objRq.SetInputValue(3, 21)  # 전일 대비 상위 순
        self.objRq.SetInputValue(4, ord('1'))  # 관리 종목 제외
        self.objRq.SetInputValue(5, ord('0'))  # 거래량 전체
        self.objRq.SetInputValue(6, ord('0'))  # '표시 항목 선택 - '0': 시가대비
        self.objRq.SetInputValue(7, 0)  #  등락율 시작
        self.objRq.SetInputValue(8, 30)  # 등락율 끝
        self.ud_list = []
        self.name_list=[] 
    # 실제적인 7043 통신 처리
    def rq7043(self, retcode):
        self.objRq.BlockRequest()
        # 현재가 통신 및 통신 에러 처리
        rqStatus = self.objRq.GetDibStatus()
        rqRet = self.objRq.GetDibMsg1()
        print("통신상태", rqStatus, rqRet)
        if rqStatus != 0:
            return False

        cnt = self.objRq.GetHeaderValue(0)
        cntTotal  = self.objRq.GetHeaderValue(1)
        print(cnt, cntTotal)
 
        for i in range(cnt):
            code = self.objRq.GetDataValue(0, i)  # 코드
            
            if len(retcode) >=  30:       # 최대 30 종목만,
                break
            if self.objRq.GetDataValue(5, i)>=20:
                retcode.append(code)
                name = self.objRq.GetDataValue(1, i)  # 종목명
                self.name_list.append(name)
                diffflag = self.objRq.GetDataValue(3, i)
                diff = self.objRq.GetDataValue(4, i)
                ud = self.objRq.GetDataValue(5, i)  # 등락율
                self.ud_list.append(ud)
                vol = self.objRq.GetDataValue(6, i)  # 거래량
            
            # print(code, name, diffflag, diff, ud, vol)
    
    def Request(self, retCode):
        self.rq7043(retCode)
 
        # 연속 데이터 조회 - 200 개까지만.
        while self.objRq.Continue:
            self.rq7043(retCode)
            print(len(retCode))
            if self.ud_list[-1] >= 0:
                break
 
        # #7043 상승하락 서비스를 통해 받은 상승률 상위 200 종목
        size = len(retCode)
        for i in range(size):
            print(retCode[i], self.name_list[i], self.ud_list[i])
        return True

if __name__=="__main__":
    codes = []
    obj7043 = Cp7043()
    if obj7043.Request(codes) == False:
        exit()
    
    print("상승종목 개수:", len(codes))