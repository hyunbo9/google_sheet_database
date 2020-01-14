# dlg_modify.py

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
import PyQt5.QtWidgets as QtWidgets
from PyQt5 import uic
import sys

from model_price import *


class model_modify_dlg(QDialog):

    def __init__(self):

        QDialog.__init__(self)
        self.ui = uic.loadUi("setting/dialog_modify.ui", self)

        self.treeWidget_2.setColumnWidth(0, 220)
        self.treeWidget_3.setColumnWidth(0,170)
        self.treeWidget_3.setColumnWidth(1,70)

        # button click connect
        self.init_btn.clicked.connect(self.model_init)
        self.model_search_btn.clicked.connect(self.find)
        self.model_modify.clicked.connect(self.model_modify_btn_click)

        self.normal_shutdown_token = False

    def model_init(self):
        self.treeWidget_1.clear()
        self.treeWidget_2.clear()
        self.treeWidget_3.clear()

        brand_name = self.QLabel_brand_label.text()

        water_purifier_model_var = brand_name + "_water_purifier_model"
        water_purifier_price_var = brand_name + "_water_purifier_price"
        air_cleaner_model_var = brand_name + "_air_cleaner_model"
        air_cleaner_price_var =brand_name + "_air_cleaner_price"
        etc_name_var = brand_name + "_etc_name"
        etc_model_var = brand_name + "_etc_model"
        etc_price_var = brand_name + "_etc_price"
        self.package_model_var = brand_name + "_package_model"
        self.package_price_var = brand_name + "_package_price"


        for i in range(len(globals()[water_purifier_model_var])):
            self.treeWidget_1.addTopLevelItem(QTreeWidgetItem(
                [str(globals()[water_purifier_model_var][i]), str(globals()[water_purifier_price_var][i])]))
        for i in range(len(globals()[air_cleaner_model_var])):
            self.treeWidget_2.addTopLevelItem(
                QTreeWidgetItem(
                    [str(globals()[air_cleaner_model_var][i]), str(globals()[air_cleaner_price_var][i])]))
        for i in range(len(globals()[etc_name_var])):
            self.treeWidget_3.addTopLevelItem(QTreeWidgetItem(
                [str(globals()[etc_name_var][i]), str(globals()[etc_model_var][i]),
                 str(globals()[etc_price_var][i])]))

    def remark__init(self):
        brand_name = self.QLabel_brand_label.text()
        remarks = brand_name + "_remarks"

        self.textEdit.setText(globals()[remarks])

    # lookup function
    def find(self):
        if self.win_model_search_text.text() != "":
            num = 0
            for i in range(self.treeWidget_1.topLevelItemCount()):
                if self.win_model_search_text.text() not in self.treeWidget_1.topLevelItem(num).text(0) and i != 0:
                    self.treeWidget_1.takeTopLevelItem(num)
                else:
                    num += 1
            num = 0
            for i in range(self.treeWidget_2.topLevelItemCount()):
                if self.win_model_search_text.text() not in self.treeWidget_2.topLevelItem(num).text(0) and i != 0:
                    self.treeWidget_2.takeTopLevelItem(num)
                else:
                    num +=1
            num = 0
            for i in range(self.treeWidget_3.topLevelItemCount()):
                if self.win_model_search_text.text() not in self.treeWidget_3.topLevelItem(num).text(0) and i != 0:
                    self.treeWidget_3.takeTopLevelItem(num)
                else:
                    num += 1

    def calculate_fee(self, model_list):

        brand_name = self.QLabel_brand_label.text()
        model_full_list = globals()[brand_name + "_water_purifier_model"]  + \
                    globals()[brand_name + "_air_cleaner_model"] + \
                    globals()[brand_name + "_etc_model"]

        fee_full_list = globals()[brand_name + "_water_purifier_fees"] + \
                        globals()[brand_name + "_air_cleaner_fees"] + \
                        globals()[brand_name + "_etc_fees"]



        # In case single
        if len(model_list) == 1:

            index = model_full_list.index(model_list[0])
            fee = fee_full_list[index]

            return int(fee)


        # In case package
        if len(model_list) != 1 :


            fee = 0
            for model in model_list:
                if not model in model_full_list : continue
                index = model_full_list.index(model)
                fee += int(fee_full_list[index].replace(",", ""))

            return int(fee)

    def calculate_price(self, model_list):
        brand_name = self.QLabel_brand_label.text()

        model_full_list = globals()[brand_name + "_water_purifier_model"] + \
                          globals()[brand_name + "_air_cleaner_model"] + \
                          globals()[brand_name + "_etc_model"]

        # In case single
        if len(model_list) == 1 and (not self.cross_package_token.checkState()):
            price_full_list = globals()[brand_name + "_water_purifier_price"] + \
                              globals()[brand_name + "_air_cleaner_price"] + \
                              globals()[brand_name + "_etc_price"]

            index = model_full_list.index(model_list[0])
            price = int(price_full_list[index].replace(",", ""))

            return price

        # In case package
        if len(model_list) != 1 or self.cross_package_token.checkState():
            package_price_full_list = globals()[brand_name + "_water_purifier_package_price"] + \
                                      globals()[brand_name + "_air_cleaner_package_price"] + \
                                      globals()[brand_name + "_etc_package_price"]

            price = 0
            for model in model_list:
                if model == "": continue
                if not model in model_full_list: continue
                index = model_full_list.index(model)
                price += int(package_price_full_list[index].replace(",", ""))

            return price

    def model_modify_btn_click(self):
        try:
            model_list = []

            for i in range(len(self.treeWidget_1.selectedItems())):
                water_purifier_model = self.treeWidget_1.selectedItems()[i].text(0)
                model_list.append(water_purifier_model)


            for i in range(len(self.treeWidget_2.selectedItems())):
                air_cleaner_model = self.treeWidget_2.selectedItems()[i].text(0)
                model_list.append(air_cleaner_model)


            for i in range(len(self.treeWidget_3.selectedItems())):
                etc_model = self.treeWidget_3.selectedItems()[i].text(1)
                model_list.append(etc_model)

            model_list = [x for x in model_list if not x == ""]

            price = self.calculate_price(model_list)

            # valve for SK, 교원, 쿠쿠
            if (self.QLabel_brand_label.text() == "SK" or \
                self.QLabel_brand_label.text() == "CUCKOO" or \
                self.QLabel_brand_label.text() == "KYOWON") and self.valve_token.checkState():
                price += 1000

            try:
                fees = self.calculate_fee(model_list)
            except Exception as ex:
                print("ex", ex)
                fees = "계산 불가"

            """=========================================== make dict =============================================="""

            model = (" + ").join(model_list)  # 정수기 + 공청기 + etc 문자열
            if not model:
                model = self.win_current_model.text()
                price = self.win_current_price.text()
                model_list = model.split(" + ")
                fees = self.calculate_fee(model_list)


            num_of_model = model.count(" + ")
            if num_of_model == 0:
                package = "단품"
            else:
                package = "패키지"

            if self.win_radio_card.isChecked():
                payment_method = "카드"
            elif self.win_radio_account.isChecked():
                payment_method = "자동 이체"

            if self.old_purifier_token.checkState():
                old_purifier_token = "yes"
            else:
                old_purifier_token = "no"

            if self.cross_package_token.checkState():
                cross_package_token = "yes"
            else:
                cross_package_token = "no"

            self.customer_info = {'모델명' : model, '렌탈료' : str(price), '색상' : self.win_color.text(),\
                                  'gift_address_token' : self.win_address_token.checkState(),\
                                  '패키지' : package, \
                                  '조리수 벨브' : self.valve_token.checkState(), \
                                  '계약자' : self.win_name.text(), '연락처' : self.win_phone_number.text(),\
                                  '설치주소' : self.win_address.text(), '생년월일' : self.win_birth_date.text(),\
                                  '설치 요청일' : self.win_request_date.selectedDate(), \
                                  '은행/카드사' : self.win_card_text.text(), \
                                  '지불 방법' : payment_method, \
                                  '계좌/카드번호' : self.win_card_account.text(), \
                                  '유효기간/예금주' : self.win_validity.text(), \
                                  'account_token': self.account_token.checkState(), \
                                  '주문 상태' : self.comboBox.currentText(),
                                  '폐 정수기 수거 요청' : old_purifier_token,
                                  '설치 완료일' : "",
                                  '수수료' : str(fees),
                                  "크로스 패키지": cross_package_token
                                  }

            self.normal_shutdown_token = True
            self.close()

        except Exception as ex:
            print("Error raise ", ex)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = model_modify_dlg()
    myWindow.show()
    app.exec_()