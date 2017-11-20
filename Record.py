# -*- coding: utf-8 -*-
import datetime
import EmployeeCard
from CardIdLoader import CardIdLoader
#   年 月 日 時 分 秒       卡號 機號
# 2017 09 30 11 40 14 0000000014  01   01
# 2017103117040000000000160201
class PreloadCardId(object):
    def __init__(self):
        self.__CardIdLoader = CardIdLoader()

    def setRecord(self,StrRecord):
        self.Record = Record(StrRecord , self.__CardIdLoader)
        return self.Record

class Record(object):
    __EmployeeCard = EmployeeCard.EmployeeCard()
    __Record = ""
    Date = datetime.datetime.now() 
    Time = datetime.datetime.now() 
    CardId = ""
    MachineId = ""
    EmployeeId = ""
    EmployeeName = ""
    isManager = False
    EmployeeType = 0 # 0:一般員工 ; 1:管理職 ; 2:排班人員
    isHoliday = False
    HolidayStatus = 0
    IgnoreCode = 0
    IgnoreMsg = ""

    def __init__(self , Record ,CardIdLoader):
        self.Record = Record
        Record = self.Record
        self.Date = self.getDate(Record[0:4],Record[4:6],Record[6:8])
        self.Time = self.getTime(Record[8:10],Record[10:12])
        self.CardId = Record[14:24]
        self.MachineId = Record[24:16]
        Employee = self.searchEmployee(self.CardId)
        self.EmployeeId = Employee[0][0].encode("utf-8")
        self.EmployeeName = Employee[0][1].encode("utf-8")

        self.HolidayStatus = self.searchHoliday(self.CardId , self.Date)
        if(self.HolidayStatus == 0):
            self.isHoliday = True
        else:
            self.isHoliday = False
        if(CardIdLoader.isScheduling(self.CardId)):
            self.EmployeeType = 2
        else:
            if(CardIdLoader.isManager(self.CardId)):
                self.EmployeeType = 1
                self.isManager = True
            else:
                self.EmployeeType = 0

    def changeHour(self , Hour):
        self.Record = self.Record[0:8] + Hour + self.Record[10:14] + self.Record[14:]
        return self

    def setIgnoreCode(self , IgnoreCode , IgnoreMsg):
        self.IgnoreCode = IgnoreCode
        self.IgnoreMsg = IgnoreMsg

    def searchEmployee(self , CardId):
        try:
            Employee = self.__EmployeeCard.searchEmployee(CardNo = CardId)
            if(not Employee):
                return[[u'查無此人 '+self.CardId,u'']]
            return Employee
        except Exception as err:
            print(err)

    def searchHoliday(self , CardId , date):
        try:
            HolidayStatus = self.__EmployeeCard.searchEmployeeHoliday(CardNo = CardId , date = date)
            return HolidayStatus
        except Exception as err:
            print(err)

    def __str__(self):
        return self.Record

    def getDate(self , y , m , d ):
        return datetime.datetime.strptime( str(y)+'-'+str(m)+'-'+str(d) , "%Y-%m-%d")

    def getTime(self ,h ,M ):
        return datetime.datetime.strptime( str(h)+":"+str(M), "%H:%M") 
