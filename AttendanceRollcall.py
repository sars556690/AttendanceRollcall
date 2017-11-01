# -*- coding: utf-8 -*-
from Tkinter import *
from tkFileDialog import askopenfilename
import os
import ConfigParser
import codecs 
import datetime
import EmployeeCard
 # 2017093011401400000000140101

class Windows:
    def __init__(self, master):   
        self.file_path=""           # 瀏覽檔案路徑
        self.save_path=""           # 儲存檔案路徑
        self.IgnoreCard=[]          # 主管卡號
        self.SchedulingCard=[]      # 排班人員卡號
        self.Recoad = []            # 已處理紀錄
        self.IgnoreRecoad = []      # 已處理刪除紀錄

        # 搜尋員工卡號
        self.EmployeeCard = EmployeeCard.EmployeeCard()
        
        # 設定載入
        config = ConfigParser.ConfigParser()
        config.readfp(codecs.open('config.ini', "r", "utf-8-sig"))
        self.save_path = config.get('system', 'path')

        # 載入主管卡號
        IgnoreCardFile = open('IgnoreCard.txt','r')
        for IgnoreCard in IgnoreCardFile.readlines():
            self.IgnoreCard.append(IgnoreCard.rstrip('\n'))

        # 載入排班人員卡號
        SchedulingCardFile = open('SchedulingCard.txt','r')
        for SchedulingCard in SchedulingCardFile.readlines():
            self.SchedulingCard.append(SchedulingCard.rstrip('\n'))

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
        self.chkbtn_value_removed_file = IntVar()
        self.chkbtn_removed_file=Checkbutton(master , text="是否產生移除的紀錄檔案" , variable=self.chkbtn_value_removed_file)
        self.chkbtn_removed_file.grid(row=2, column=0 ,columnspan=2, sticky=W)

        # 設定UI 處理紀錄按鈕
        self.btn_Change= Button(master, text="完成" ,command=self.change_date)
        self.btn_Change.grid(row=2, column=3, sticky = W + E)

        # 設定UI 瀏覽檔案
        self.btn_Browse= Button(master, text="瀏覽", command=self.browse)
        self.btn_Browse.grid(row=1, column=3)

    def browse(self):
        # 瀏覽檔案
        opts = {}
        opts['filetypes'] = [('Text','.txt'),('all files','.*')]
        self.file_path = askopenfilename(**opts)
        self.entry_file.delete(0, END)
        self.entry_file.insert(0, self.file_path)

    def change_date(self):
        
        Label_Msg = ""
        self.file_path = self.entry_file.get()
        if(self.file_path!="" and os.path.isfile(self.file_path)):
            self.Recoad = []
            self.IgnoreRecoad =[]
            file = open(self.file_path,'r+')
            
            fileRowCount = len(file.readlines())
            file.seek(0, 0)
            PassCount = 0
            for i in range(0, fileRowCount):
                fileRow = file.readline()
                # 若為排班人員則不處理
                if(fileRow[14:24] in self.SchedulingCard):
                    self.Recoad.append(fileRow)
                    PassCount+=1 
                else:
                    time0700 = self.getTime(7,0)
                    time1800 = self.getTime(18,00)
                    time2155 = self.getTime(21,55)
                    recoadTime = self.getTime(fileRow[8:10],fileRow[10:12])
                    recoadDate = self.getDate(fileRow[0:4],fileRow[4:6],fileRow[6:8]).timetuple()
                    # 不為主管
                    if(fileRow[14:24] not in self.IgnoreCard):
                        # 是否為假日
                        if(recoadDate.tm_wday >= 5 ):
                            # 是否 18點 前下班
                            if(recoadTime < time1800):
                                self.Recoad.append(fileRow)
                                PassCount+=1 
                            else:
                                self.IgnoreRecoad.append(fileRow)
                         # 是否 21：55 前下班
                        elif (recoadTime < time2155):
                            self.Recoad.append(fileRow)
                            PassCount+=1 
                        else:
                            self.IgnoreRecoad.append(fileRow)
                    # 主管
                    elif(fileRow[14:24] in self.IgnoreCard):
                        # 是否在 7:30 及 18點 打卡
                        if(recoadTime>time0700 and recoadTime < time1800):
                            self.Recoad.append(fileRow)
                            PassCount+=1 
                        else:
                            self.IgnoreRecoad.append(fileRow)

            # 被過濾資料比數
            hasRemoved = fileRowCount - PassCount
            
            # 儲存檔案
            new_file_path=os.path.split(self.file_path)
            new_file_dir=''
            new_file_name= os.path.splitext(new_file_path[1])[0]
            if(self.save_path != ""):
                if(os.path.isdir(self.save_path)):
                    Label_Msg = ''
                    new_file_dir = self.save_path+'/'
                else:
                    Label_Msg += '儲存路徑有誤\n'
    
            file = open(new_file_dir + self.getTodayString() + '.txt','w')
            for Employee in self.Recoad:
                file.write(Employee)
    
            if(self.chkbtn_value_removed_file.get()):
                if(not os.path.isdir('Ignore/')):
                    os.makedirs('Ignore/')
                file = open('Ignore/' + new_file_name+ '_Ignore.txt','w')
                for IgnoreRecoad in self.IgnoreRecoad:
                    file.write(self.processIgnore(IgnoreRecoad))
            Label_Msg+="完成，已移去"+str(hasRemoved)+"項紀錄"
        else:
            Label_Msg+='檔案路徑有誤'

        # 顯示處理訊息
        self.var.set(Label_Msg)
    
    def getDate(self , y , m , d ):
        return datetime.datetime.strptime( str(y)+'-'+str(m)+'-'+str(d) , "%Y-%m-%d")

    def getTime(self ,h ,M ):
        return datetime.datetime.strptime( str(h)+":"+str(M), "%H:%M")    

    def processIgnore(self ,recoad):
        Employee = self.EmployeeCard.searchEmployee(CardNo = recoad[14:24])
        Record = recoad[0:4] + '-' + \
        		 recoad[4:6] + '-' + recoad[6:8] + '  ' + \
        		 recoad[8:10] + ':' + recoad[10:12] + '  ' + \
        		 recoad[14:24] + '  ' + Employee[0][0].encode("utf-8") + '  ' + \
        		 Employee[0][1].encode("utf-8") + '\n'
        return Record

    def getTodayString(self):
        today = datetime.datetime.now().timetuple()
        return str(today.tm_year) + str(today.tm_mon) + str(today.tm_mday)

root = Tk()
window=Windows(root)
root.mainloop()  
