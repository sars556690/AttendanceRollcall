# -*- coding: utf-8 -*-

import pyodbc
import datetime
class EmployeeCard:
    """docstring for EmployeeCard"""
    def __init__(self):
        self.cnxn = None
        try:
            self.cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.10.233;DATABASE=HRMDB;UID=odbc;PWD=batom5280')
            self.cursor = self.cnxn.cursor()
        except Exception as err:
            print(err)

    def searchEmployee(self , CardNo = ""):
        if(self.cnxn):
            EmployeeList = []
            sql = "select Employee.Code, Employee.CnName, Card.CardNo \
                    FROM dbo.Employee AS Employee \
                    LEFT OUTER JOIN dbo.Card AS Card \
                    ON Employee.EmployeeId = Card.EmployeeId \
                    where UseTypeId ='UseType_001' "
            if(CardNo!=""):
                sql += "and CardNo ='"+ str(CardNo) +"'"
            self.cursor.execute(sql)
            # row = self.cursor.fetchone()
            for i in self.cursor:    
                if i:
                    # 工號，姓名，卡號
                    EmployeeList.append([i[0] , i[1] , i[2]])
            return EmployeeList

    def searchEmployeeHoliday(self , CardNo=None , date=None ):
        if(self.cnxn != None):
            datetime.datetime
            sql = "SELECT AttendanceEmpRank.[Date]\
                    ,AttendanceEmpRank.[AttendanceHolidayTypeId]\
                    ,AttendanceHolidayType.Name\
                    FROM [HRMDB].[dbo].[AttendanceEmpRank] as AttendanceEmpRank\
                    LEFT OUTER JOIN [HRMDB].[dbo].[AttendanceHolidayType] as AttendanceHolidayType\
                    ON AttendanceEmpRank.AttendanceHolidayTypeId = AttendanceHolidayType.AttendanceHolidayTypeId\
                    LEFT OUTER JOIN [HRMDB].[dbo].Employee as Employee\
                    ON Employee.EmployeeId = AttendanceEmpRank.EmployeeId\
                    LEFT OUTER JOIN dbo.Card AS Card \
                    ON Employee.EmployeeId = Card.EmployeeId "
            if(CardNo != None or date != None):
                sql +=" where "
    
            if(CardNo != None):
                sql += "CardNo ='"+ str(CardNo) +"'"
    
            if(date !=None):
                if(CardNo != None):
                    sql +=" and "
                sql += "AttendanceEmpRank.Date = '"+ str(date) +"'"

            self.cursor.execute(sql)
            employee = self.cursor.fetchall()

            # 0 假日
            # 1 平日
            # 2 無資料
            if(employee):
                if(employee[0][1] != "DefaultHolidayType001"):
                    return 0
                else:
                    return 1
            else:
                return 2

# a=EmployeeCard()
# b = a.searchEmployeeHoliday( EmployeeCode = "00336" ,date = datetime.datetime(2017,10,10) )
# print(b)
# EmployeeList = a.searchEmployee(CardNo ="0000000336")
# print(EmployeeList[0][1].encode('utf-8'))