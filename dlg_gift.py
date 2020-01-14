# dlg_gift.py

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import PyQt5.QtWidgets as QtWidgets
from PyQt5 import uic
import sys

from model_price import gift1_a, gift1_b, gift1_c, gift1_d,\
    gift2_a, gift2_b, gift2_c, gift2_d,\
    gift3_a, gift3_b, gift3_c, gift3_d


class gift_append_dlg(QDialog):
    def __init__(self):

        QDialog.__init__(self)
        self.ui = uic.loadUi("setting/dialog_gift.ui", self)
        self.ui.show()
        """======================================== size setting========================================"""
        self.gift1.setColumnWidth(0, 100)
        self.gift1.setColumnWidth(1, 270)
        self.gift1.setColumnWidth(2, 150)
        self.gift2.setColumnWidth(0, 100)
        self.gift2.setColumnWidth(1, 270)
        self.gift2.setColumnWidth(2, 150)
        self.gift3.setColumnWidth(0, 100)
        self.gift3.setColumnWidth(1, 270)
        self.gift3.setColumnWidth(2, 150)

        """=========================================================================================="""
        self.win_gift_append.clicked.connect(self.gift_append_btn_click)
        self.win_find_btn.clicked.connect(self.find)
        self.win_init_btn.clicked.connect(self.gift_init)
        self.normal_shutdown_token = False

    def gift_init(self):


        self.gift1.clear(), self.gift2.clear(), self.gift3.clear()

        for x, y, z in zip(gift1_a, gift1_b, gift1_d):
            self.gift1.addTopLevelItem(QTreeWidgetItem([x, y, z]))

        for x, y, z in zip(gift2_a, gift2_b, gift2_d):
            self.gift2.addTopLevelItem(QTreeWidgetItem([x, y, z]))

        for x, y, z in zip(gift3_a, gift3_b, gift3_d):
            self.gift3.addTopLevelItem(QTreeWidgetItem([x, y, z]))

    def find(self):
        if self.win_find_text.text() != "":
            num = 0
            for i in range(self.gift1.topLevelItemCount()):
                if self.win_find_text.text() not in self.gift1.topLevelItem(num).text(1):
                    self.gift1.takeTopLevelItem(num)
                else:
                    num += 1
            num = 0
            for i in range(self.gift2.topLevelItemCount()):
                if self.win_find_text.text() not in self.gift2.topLevelItem(num).text(1):
                    self.gift2.takeTopLevelItem(num)
                else:
                    num += 1
            num = 0
            for i in range(self.gift3.topLevelItemCount()):
                if self.win_find_text.text() not in self.gift3.topLevelItem(num).text(1):
                    self.gift3.takeTopLevelItem(num)
                else:
                    num += 1

    def gift_append_btn_click(self):

        rank_list = []
        gift_list = []
        price_list = []

        for x in self.gift1.selectedItems():
            rank_list.append(x.text(0))
            gift_list.append(x.text(1))
            price_list.append(x.text(2))

        for x in self.gift2.selectedItems():
            rank_list.append(x.text(0))
            gift_list.append(x.text(1))
            price_list.append(x.text(2))

        for x in self.gift3.selectedItems():
            rank_list.append(x.text(0))
            gift_list.append(x.text(1))
            price_list.append(x.text(2))

        rank_list = [x for x in rank_list if x]
        gift_list = [x for x in gift_list if x]
        price_list = [x for x in price_list if x]

        gift_full_list = gift1_b + gift2_b + gift3_b
        company_full_list = gift1_c + gift2_c + gift3_c

        company_list = []
        for gift in gift_list:
            if not gift in gift_full_list : continue
            index = gift_full_list.index(gift)
            company_list.append(company_full_list[index])

        self.info = {"분류" : rank_list,
                    "사은품" : gift_list,
                     "업체명" : company_list,
                     "단가" : price_list}

        self.normal_shutdown_token = True
        self.close()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = gift_append_dlg()
    myWindow.show()
    app.exec_()