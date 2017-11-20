# -*- coding: utf-8 -*-

class CardIdLoader(object):
    """docstring for DataLoader"""
    __ManagerCards=[]
    __SchedulingCards=[]
    def __init__(self):
        # 載入主管卡號
        self.ManagerCardFile = open('ManagerCard.txt','r')
        for ManagerCard in self.ManagerCardFile.readlines():
            self.__ManagerCards.append(ManagerCard.rstrip('\n'))

        # 載入排班人員卡號
        self.SchedulingCardFile = open('SchedulingCard.txt','r')
        for SchedulingCard in self.SchedulingCardFile.readlines():
            self.__SchedulingCards.append(SchedulingCard.rstrip('\n'))

    def isManager(self , CardId):
        return CardId in self.__ManagerCards

    def isScheduling(self , CardId):
        return CardId in self.__SchedulingCards


        