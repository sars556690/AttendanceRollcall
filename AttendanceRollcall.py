# -*- coding: utf-8 -*-
from Tkinter import *
from tkFileDialog import askopenfilename
import os
import ConfigParser
import codecs 
import datetime
import EmployeeCard
import xlsxwriter
import Record 
 # 2017093011401400000000140101

class Windows:

    def __init__(self, master):   
        self.time0700 = self.getTime(7,0)
        self.time1800 = self.getTime(18,00)
        self.time2155 = self.getTime(21,55)
        self.file_path=""                   # 瀏覽檔案路徑
        self.save_path=""                   # 儲存檔案路徑
        self.ManagerCards=[]                # 主管卡號
        self.SchedulingCards=[]             # 排班人員卡號
        self.Recoads = []                   # 已處理紀錄
        self.IgnoreRecords = []             # 已處理刪除紀錄
        self.IgnoreRecords2excels = []      # 已處理刪除紀錄(補刷卡)
        self.IgnoreRecords2excels_2 = []    # 已處理刪除紀錄(補刷卡)處理後
        self.PassCount = 0
        self.Label_Msg = ""
        self.IgnoreMsg = ["假日超時" ,"平日超時","主管假日超時","主管平日超時","查無資料(無排班)" ]

        # 搜尋員工卡號
        self.EmployeeCard = EmployeeCard.EmployeeCard()

        self.PreloadCardId = Record.PreloadCardId()
        # 設定載入
        config = ConfigParser.ConfigParser()
        config.readfp(codecs.open('config.ini', "r", "utf-8-sig"))
        self.save_path = config.get('system', 'path')

        # 設定UI
        master.title("移除打卡資料")
        master.resizable(0,0)

        # 設定UI 顯示訊息
        self.var = StringVar()
        msg = Label( master, textvariable=self.var).grid(row=4, column=0 , columnspan=2, sticky=W)

        # 設定UI 載入檔案欄
        Lable_file=Label(master, text="檔案").grid(row=1, column=0)
        self.entry_file=Entry(master)
        self.entry_file.grid(row=1, column=1)

        # 設定UI checkbox產生已移除紀錄檔案
        self.chkbtn_value_removed_file = IntVar(value=1)
        self.chkbtn_removed_file=Checkbutton(master , text="是否產生移除的紀錄檔案" , command = self.chkbtn_command_removed_file , variable=self.chkbtn_value_removed_file)
        self.chkbtn_removed_file.grid(row=2, column=0 ,columnspan=2, sticky=W)

        self.chkbtn_value_removed_excel = IntVar(value=1)
        self.chkbtn_removed_excel=Checkbutton(master , text="Excel" , variable=self.chkbtn_value_removed_excel)
        self.chkbtn_removed_excel.grid(row=3, column=0)

        self.chkbtn_value_removed_text = IntVar(value=0)
        self.chkbtn_removed_text=Checkbutton(master , text="Text" , variable=self.chkbtn_value_removed_text)
        self.chkbtn_removed_text.grid(row=3, column=1)

        # 設定UI 處理紀錄按鈕
        self.btn_Change= Button(master, text="完成" ,command=self.change_date)
        self.btn_Change.grid(row=2, column=3, sticky = W + E)

        # 設定UI 瀏覽檔案
        self.btn_Browse= Button(master, text="瀏覽", command=self.browse)
        self.btn_Browse.grid(row=1, column=3)

    def chkbtn_command_removed_file(self):
        if(self.chkbtn_value_removed_file.get()):
            self.chkbtn_removed_excel['state'] = 'normal'
            self.chkbtn_removed_text['state'] = 'normal'
        else:
            self.chkbtn_removed_excel['state'] = 'disabled'
            self.chkbtn_removed_text['state'] = 'disabled'
    def browse(self):
        # 瀏覽檔案
        opts = {}
        opts['filetypes'] = [('Text','.txt'),('all files','.*')]
        self.file_path = askopenfilename(**opts)
        self.entry_file.delete(0, END)
        self.entry_file.insert(0, self.file_path)

    def change_date(self):
        
        self.file_path = self.entry_file.get()
        if(self.file_path!="" and os.path.isfile(self.file_path)):
            self.Recoads = []
            self.IgnoreRecords =[]
            self.IgnoreRecords2excel = []
            self.IgnoreRecords2excel_2 = []
            self.Label_Msg = ''
            self.PassCount = 0
            file = open(self.file_path,'r+')
            
            fileRowCount = len(file.readlines())
            file.seek(0, 0)
            
            for i in range(0, fileRowCount):
                fileRow = file.readline()
                self.Rollcall(fileRow)

            # 被過濾資料比數
            hasRemoved = fileRowCount - self.PassCount

            # 儲存點名後的txt檔
            new_file_name = self.savePorcessedRecord()

            if(self.chkbtn_value_removed_file.get()):
                if(not os.path.isdir('Ignore/')):
                    os.makedirs('Ignore/')

                if(self.chkbtn_value_removed_excel.get()):
                    # 儲存過濾後excel檔
                    self.saveIgnoreRecord2xlsx(new_file_name)

                if(self.chkbtn_value_removed_text.get()):
                    # 儲存過濾後txt檔
                    self.saveIgnoreRecord(new_file_name)

            self.Label_Msg+="完成，已移去"+str(hasRemoved)+"項紀錄"
        else:
            self.Label_Msg+='檔案路徑有誤'

        # 顯示處理訊息
        self.var.set(self.Label_Msg)
        
    # 點名處理
    def Rollcall(self,fileRow):
        # 若為空白則不處理
        if(fileRow.lstrip()==""):
            return 

        record = self.PreloadCardId.setRecord(fileRow)
        # 若為排班人員則不處理
        if(record.EmployeeType == 2):
            self.Recoads.append(record)
            self.PassCount+=1 
        else:
            # 不為主管
            if(record.EmployeeType == 0):
                # 是否為假日
                if(record.isHoliday):
                    # 是否 18點 前下班
                    if(record.Time <= self.time1800):
                        self.Recoads.append(record)
                        self.PassCount+=1  
                    else:
                        record.setIgnoreCode(0, self.IgnoreMsg[0])
                        self.IgnoreRecords.append(record)
                        self.addintoIgnoreRecord(record)
                else:
                    if(record.HolidayStatus==2):
                        record.setIgnoreCode(4, self.IgnoreMsg[4])
                        self.IgnoreRecords.append(record)
                        self.addintoIgnoreRecord(record)
                    elif(record.HolidayStatus==1):
                        # 是否 21：55 前下班
                        if (record.Time <= self.time2155):
                            self.Recoads.append(record)
                            self.PassCount+=1 
                        else:
                            record.setIgnoreCode(1, self.IgnoreMsg[1])
                            self.IgnoreRecords.append(record)
                            self.addintoIgnoreRecord(record)
            # 主管
            elif(record.EmployeeType == 1):
                if(record.isHoliday):
                    # 是否在 7:30 及 18點 打卡
                    if(record.Time>=self.time0700 and record.Time <= self.time1800):
                        self.Recoads.append(record)
                        self.PassCount+=1 
                    else:
                        record.setIgnoreCode(2, self.IgnoreMsg[2])
                        self.IgnoreRecords.append(record)
                        self.addintoIgnoreRecord(record)
                else:
                    if(record.HolidayStatus==2):
                        record.setIgnoreCode(4, self.IgnoreMsg[4])
                        self.IgnoreRecords.append(record)
                        self.addintoIgnoreRecord(record)
                    elif(record.HolidayStatus==1):
                        # 是否在 7:30 及 18點 打卡
                        if(record.Time>=self.time0700 and record.Time <= self.time1800):
                            self.Recoads.append(record)
                            self.PassCount+=1 
                        else:
                            record.setIgnoreCode(3, self.IgnoreMsg[3])
                            self.IgnoreRecords.append(record)
                            self.addintoIgnoreRecord(record)
        return 

    # 剔除當天重複人員打卡資料
    # return Status 是否有重複
    def removeDuplicatedRecord(self , IgnoreRecord2excel , record):
        Status = False
        Two_Hours = 3600*2
        for d in IgnoreRecord2excel:
            if(d.CardId == record.CardId and d.Date == record.Date):
                if(abs((record.Date - d.Date).total_seconds()) < Two_Hours):
                    Status=True
        return Status

    # 加入IgnoreRecord 及 excel 的 List
    def addintoIgnoreRecord(self , record):
        isDuplicated = self.removeDuplicatedRecord(self.IgnoreRecords2excel , record)
        if(not isDuplicated):
            if(record.isManager):
                if(record.Time<self.time0700):
                    self.IgnoreRecords2excel.append(record)
                    new_record = record.changeHour("07")
                    self.IgnoreRecords2excel_2.append(new_record)
                elif(record.Time > self.time1800):
                    self.IgnoreRecords2excel.append(record)
                    new_record = record.changeHour("17")
                    self.IgnoreRecords2excel_2.append(new_record)
            else:
                self.IgnoreRecords2excel.append(record)
                self.IgnoreRecords2excel_2.append(record)

    # 儲存點名後的txt檔
    # return new_file_name 檔名
    def savePorcessedRecord(self):
        # 儲存檔案
        new_file_path=os.path.split(self.file_path)
        new_file_dir=''
        new_file_name= os.path.splitext(new_file_path[1])[0]
        if(self.save_path != ""):
            if(os.path.isdir(self.save_path)):
                self.Label_Msg = ''
                new_file_dir = self.save_path+'/'
            else:
                self.Label_Msg += '儲存路徑有誤\n'
    
        file = open(new_file_dir + self.getTodayString() + '.txt','w')
        for Recoad in self.Recoads:
            file.write(str(Recoad))
        return new_file_name

    # 儲存過濾後txt檔
    def saveIgnoreRecord(self , new_file_name):
        file = open('Ignore/' + new_file_name+ '_Ignore.txt','w')
        for IgnoreRecord in self.IgnoreRecords:
            file.write(self.processIgnore(IgnoreRecord) +"  "+IgnoreRecord.IgnoreMsg+"\n")

    # 儲存過濾後excel檔
    def saveIgnoreRecord2xlsx(self , new_file_name):
        workbook = xlsxwriter.Workbook('Ignore/'+ new_file_name + '_Ignore' + '.xlsx')
        worksheet = workbook.add_worksheet()
        worksheet.set_column(0,2, 20)
        format = workbook.add_format({'font_color': 'red',"align":"center" ,"num_format": "@"})
        Row = 0
        worksheet.write(Row, 0, u"工號", format)
        worksheet.write(Row, 1, u"刷卡日期", format)
        worksheet.write(Row, 2, u"刷卡時間", format)
        
        for record in self.IgnoreRecords2excel_2:
            IgnoreData = self.processIgnore2xlsx(record)
            Row += 1
            if(record.IgnoreCode < 2 or record.IgnoreCode==4):
                format = workbook.add_format({"align":"center" ,"num_format": "@",'bg_color': 'yellow'})
            else:
                format = workbook.add_format({"align":"center" ,"num_format": "@"})
            worksheet.write(Row, 0, IgnoreData[0], format)
            worksheet.write(Row, 1, IgnoreData[1], format)
            worksheet.write(Row, 2, IgnoreData[2], format)
            
        workbook.close() 
    
    def getDate(self , y , m , d ):
        return datetime.datetime.strptime( str(y)+'-'+str(m)+'-'+str(d) , "%Y-%m-%d")

    def getTime(self ,h ,M ):
        return datetime.datetime.strptime( str(h)+":"+str(M), "%H:%M")    

    def processIgnore(self ,recoad):
        Record = str(recoad)[0:4] + '-' + \
                 str(recoad)[4:6] + '-' + str(recoad)[6:8] + '  ' + \
                 str(recoad)[8:10] + ':' + str(recoad)[10:12] + '  ' + \
                 str(recoad)[14:24] + '  ' 

        Record += recoad.EmployeeId + "  " + recoad.EmployeeName
        return Record

    def processIgnore2xlsx(self ,recoad):

        EmployeeId = recoad.EmployeeId.decode("utf-8") 
        RecoadDate = str(recoad)[0:4] + '-' + str(recoad)[4:6] + '-' + str(recoad)[6:8]
        RecoadTime = str(recoad)[8:10] + ':' + str(recoad)[10:12]
        
        Record = [EmployeeId , RecoadDate ,  RecoadTime]

        return Record

    def getTodayString(self):
        today = datetime.datetime.now().timetuple()
        return str(today.tm_year) + str(today.tm_mon) + str(today.tm_mday)

root = Tk()
window=Windows(root)
root.mainloop()  
