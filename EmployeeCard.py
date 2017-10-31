# -*- coding: utf-8 -*-

import pyodbc

class EmployeeCard:
    """docstring for EmployeeCard"""
    def __init__(self):
        cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=192.168.10.233;DATABASE=HRMDB;UID=odbc;PWD=batom5280')
        self.cursor = cnxn.cursor()
        print(cnxn)

    def searchEmployee(self , CardNo = ""):
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
                EmployeeList.append([i[0] , i[1] , i[2]])
        return EmployeeList
                
a=EmployeeCard()
EmployeeList = a.searchEmployee(CardNo ="0000000336")
print(EmployeeList[0][1].encode('utf-8'))