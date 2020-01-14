# main.py

import sys
from PyQt5 import uic, QtCore
import PyQt5.QtWidgets as QtWidgets
from PyQt5.QtWidgets import *
from functools import partial
import datetime

from credentials import google_sheet, main_sheet_first_row
from setting_from_excel import consultant_list, order_path_list
import dlg_model, dlg_gift, dlg_find, dlg_modify, dlg_enroll_complete, dlg_install_complete

#lib path append
libpaths = QtWidgets.QApplication.libraryPaths()
libpaths.append("setting/plugins")
QtWidgets.QApplication.setLibraryPaths(libpaths)

main_window = uic.loadUiType("setting/Main_window.ui")[0]

class Main_window(QMainWindow, main_window):
    def __init__(self):
        super().__init__()
        self.setupUi(self)
        """================================= column size setting == ===================================="""
        self.win_model_treeview.setColumnWidth(0, 55)       # 제품사
        self.win_model_treeview.setColumnWidth(1, 85)       # 주문 상태
        self.win_model_treeview.setColumnWidth(2, 95)       # 설치 요청일
        self.win_model_treeview.setColumnWidth(3, 85)       # 설치 완료일
        self.win_model_treeview.setColumnWidth(4, 200)       # 모델명
        self.win_model_treeview.setColumnWidth(5, 75)       # 렌탈료
        self.win_model_treeview.setColumnWidth(6, 75)       # 색상
        self.win_model_treeview.setColumnWidth(7, 45)       # 패키지 or 단품
        self.win_model_treeview.setColumnWidth(8, 70)       # 계약자
        self.win_model_treeview.setColumnWidth(9, 110)       # 연락처
        self.win_model_treeview.setColumnWidth(10, 70)       # 생년월일
        self.win_model_treeview.setColumnWidth(11, 400)       # 설치주소
        self.win_model_treeview.setColumnWidth(12, 45)       # 조리수 벨브
        self.win_model_treeview.setColumnWidth(13, 180)       # 카드사
        self.win_model_treeview.setColumnWidth(14, 300)       # 카드 번호
        """=====================================  GUI buttion ====================================="""
        self.win_append.clicked.connect(self.google_sheet_append_btn_click)                  # 신규등록 버튼
        self.win_modify.clicked.connect(self.motify_btn_click)                  # 수정 버튼
        self.win_find.clicked.connect(self.find_btn_click)                      # 조회 버튼
        self.win_int.clicked.connect(self.defalt_btn_click)                         # 초기화 버튼
        self.win_lg_enroll_btn.clicked.connect(self.enroll_complete)            # 등록 완료 버튼
        self.win_install_complete_btn.clicked.connect(self.install_complete)
        """=====================================  model ====================================="""
        self.win_cuckoo.clicked.connect(partial(self.model_append_btn_click, brand = "CUCKOO"))
        self.win_lg.clicked.connect(partial(self.model_append_btn_click, brand = "LG"))
        self.win_ruhens.clicked.connect(partial(self.model_append_btn_click, brand = "RUHENS"))
        self.win_sk.clicked.connect(partial(self.model_append_btn_click, brand = "SK"))
        self.win_hyundai.clicked.connect(partial(self.model_append_btn_click, brand = "HYUNDAI"))
        self.win_chungho.clicked.connect(partial(self.model_append_btn_click, brand = "CHUNGHO"))
        self.win_kyowon.clicked.connect(partial(self.model_append_btn_click, brand = "KYOWON"))
        self.win_welrix.clicked.connect(partial(self.model_append_btn_click, brand = "WELRIX"))
        self.win_novita.clicked.connect(partial(self.model_append_btn_click, brand = "NOVITA"))
        self.win_samsung.clicked.connect(partial(self.model_append_btn_click, brand = "SAMSUNG"))
        self.win_cherish.clicked.connect(partial(self.model_append_btn_click, brand = "CHERISH"))
        self.win_model_delete.clicked.connect(self.model_delete_btn_click)
        self.win_model_treeview.itemDoubleClicked.connect(self.model_modify)
        self.win_modify_order_state.clicked.connect(self.order_state_edit)
        """===================================== gift ====================================="""
        self.win_gift_append.clicked.connect(self.gift_append_btn_click)
        self.win_gift_delete.clicked.connect(self.gift_delete_btn_click)
        self.win_gift_force_enroll_btn.clicked.connect(self.gift_force_enroll)

        """===================================== consultant ====================================="""
        [self.win_order_path.addItem(x) for x in order_path_list]
        [self.comboBox.addItem(x) for x in consultant_list]

        """===================================== calculate fee ====================================="""
        self.win_payback_money.textEdited.connect(self.calculate_fee)

        self.customer_info = {}
        for cell in main_sheet_first_row:
            self.customer_info[cell] = ""

    """===================================== model ====================================="""

    def find_model_header(self):
        headerItem = self.win_model_treeview.headerItem()
        num_of_column = headerItem.columnCount()

        model_header = []
        for i in range(num_of_column):
            x = headerItem.text(i)
            model_header.append(x)

        return model_header

    def make_model_item_list_for_model_tree(self, dlg, model_header, brand):
        model_item_list = []
        for x in model_header:
            if x == "제품사":
                model_item_list.append(brand)
                continue
            elif x == "설치 요청일":
                model_item_list.append(dlg.customer_info['설치 요청일'].toString('yyyy.MM.dd (ddd)'))
                continue
            elif x == "렌탈료":
                model_item_list.append(str(dlg.customer_info['렌탈료']))
                continue
            elif x == "조리수 벨브":
                if dlg.customer_info['조리수 벨브']: valve = "yes"
                else: valve = "no"
                model_item_list.append(valve)
                continue

            model_item_list.append(dlg.customer_info[x])



        return model_item_list

    # create_model_tree
    def model_append_btn_click(self, brand):

        try:
            if self.win_time.text() != "" : return

            dlg = dlg_model.model_append_dlg()
            dlg.QLabel_brand_label.setText(brand)
            dlg.model_init()
            dlg.remark__init()

            # data restore
            dlg.win_name.setText(self.customer_info['계약자'])
            dlg.win_phone_number.setText(self.customer_info['연락처'])
            dlg.win_address.setText(self.customer_info['설치주소'])
            dlg.win_color.setText(self.customer_info['색상'])
            dlg.win_birth_date.setText(self.customer_info['생년월일'])
            if self.customer_info['설치 요청일']:
                dlg.win_request_date.setSelectedDate(self.customer_info['설치 요청일'])

            dlg.exec_()
            if not dlg.normal_shutdown_token: return 0

            # customer infomation backup
            self.customer_info['계약자']  = dlg.customer_info['계약자']
            self.customer_info['연락처'] = dlg.customer_info['연락처']
            self.customer_info['설치주소']  = dlg.customer_info['설치주소']
            self.customer_info['생년월일'] = dlg.customer_info['생년월일']
            self.customer_info['설치 요청일'] = dlg.customer_info['설치 요청일']

            # gift_address is same
            if dlg.customer_info['gift_address_token'] :
                self.win_gift_address.setText(dlg.customer_info['설치주소'])

            # account is same
            if dlg.customer_info['지불 방법'] =="자동 이체" and dlg.customer_info['account_token']:
                self.win_bank_name.setText(dlg.customer_info['은행/카드사'])
                self.win_account_number.setText(dlg.customer_info['계좌/카드번호'])
                self.win_account_owner.setText(dlg.customer_info['유효기간/예금주'])

            model_header = self.find_model_header()
            model_item_list = \
                self.make_model_item_list_for_model_tree(dlg=dlg, model_header=model_header, brand=brand)
            self.win_model_treeview.addTopLevelItem(QTreeWidgetItem(model_item_list))

            self.calculate_fee()

        except Exception as e:
            print("model error", e)

    def model_modify(self):
        try:
            model_header = self.find_model_header()
            currentItem = self.win_model_treeview.currentItem()
            customer_info ={ model_header[i] : currentItem.text(i) \
                                                for i in range(len(model_header))}


            company_name = customer_info['제품사']

            dlg = dlg_modify.model_modify_dlg()
            dlg.QLabel_brand_label.setText(company_name)
            dlg.win_name.setText(customer_info['계약자'])
            dlg.win_phone_number.setText(customer_info['연락처'])
            dlg.win_address.setText(customer_info['설치주소'])
            dlg.win_color.setText(customer_info['색상'])
            dlg.win_birth_date.setText(customer_info['생년월일'])
            dlg.win_card_text.setText(customer_info['은행/카드사'])
            dlg.win_card_account.setText(customer_info['계좌/카드번호'])
            dlg.win_validity.setText(customer_info['유효기간/예금주'])
            dlg.win_current_model.setText(customer_info['모델명'])
            dlg.win_current_price.setText(customer_info['렌탈료'])

            dlg.win_request_date.setSelectedDate(
                QtCore.QDate.fromString(customer_info['설치 요청일'], 'yyyy.MM.dd (ddd)'))
            dlg.win_address_token.setCheckState(0)
            if customer_info['조리수 벨브'] == "yes" : dlg.valve_token.setCheckState(2)
            else : dlg.valve_token.setCheckState(0)
            if customer_info['지불 방법'] == "카드": dlg.win_radio_card.setChecked(True)
            else : dlg.win_radio_account.setChecked(True)
            index = dlg.comboBox.findText(customer_info['주문 상태'], QtCore.Qt.MatchFixedString)
            if index >= 0: dlg.comboBox.setCurrentIndex(index)
            if customer_info['크로스 패키지'] == "yes":
                dlg.cross_package_token.setCheckState(2)
            if customer_info['폐 정수기 수거 요청'] == "yes":
                dlg.old_purifier_token.setCheckState(2)



            dlg.model_init()
            dlg.remark__init()

            dlg.exec_()
            if not dlg.normal_shutdown_token: return 0

            # gift_address is same
            if dlg.customer_info['gift_address_token']:
                self.win_gift_address.setText(dlg.customer_info['설치주소'])

            # remittance is same
            if dlg.customer_info['지불 방법'] == "자동 이체" and dlg.customer_info['account_token']:
                self.win_bank_name.setText(dlg.customer_info['은행/카드사'])
                self.win_account_number.setText(dlg.customer_info['계좌/카드번호'])
                self.win_account_owner.setText(dlg.customer_info['유효기간/예금주'])


            model_header = self.find_model_header()
            model_item_list = \
                self.make_model_item_list_for_model_tree(dlg=dlg, model_header=model_header, brand = customer_info['제품사'])

            index = self.win_model_treeview.indexOfTopLevelItem(self.win_model_treeview.currentItem())
            self.win_model_treeview.takeTopLevelItem(index)
            self.win_model_treeview.insertTopLevelItem(index, QTreeWidgetItem(model_item_list))


            self.calculate_fee()

        except Exception as ex:
            print("Error, model modify : ", ex)

    # delete_model_tree
    def model_delete_btn_click(self):
        if self.win_time.text() != "" : return 0

        index = self.win_model_treeview.indexOfTopLevelItem(self.win_model_treeview.currentItem())  # 현재 선택된 값의 인덱스를 반환
        self.win_model_treeview.takeTopLevelItem(index)
        self.calculate_fee()

    """===================================== gift ====================================="""

    def find_gift_header(self):
        headerItem = self.win_gift_treeview.headerItem()
        num_of_column = headerItem.columnCount()

        gift_header = []
        for i in range(num_of_column):
            x = headerItem.text(i)
            gift_header.append(x)
        return gift_header

    # gift append buttion
    def gift_append_btn_click(self):
        try:
            dlg = dlg_gift.gift_append_dlg()
            dlg.gift_init()

            dlg.exec_()
            if not dlg.normal_shutdown_token: return 0

            gift_header = self.find_gift_header()

            dlg.info["사은품 발송 여부"] = ["접수"] * len(dlg.info['사은품'])
            dlg.info["사은품 비고란"] = [""] * len(dlg.info['사은품'])


            for i in range(len(dlg.info['사은품'])):
                gift_list = [dlg.info.get(gift_header[j], "")[i] for j in range(len(gift_header))]
                self.win_gift_treeview.addTopLevelItem(QTreeWidgetItem(gift_list))

            self.calculate_fee()
        except Exception as e:
            print("gift error", e)

    def gift_force_enroll(self):
        if self.win_gift_force_enroll.text() != "":
            gift = self.win_gift_force_enroll.text() + " (임의 등록)"
            self.win_gift_treeview.addTopLevelItem(QTreeWidgetItem(["", gift]))
            self.win_gift_force_enroll.setText("")

        self.calculate_fee()

    def gift_delete_btn_click(self):
        try:
            currend_index = self.win_gift_treeview.indexOfTopLevelItem(self.win_gift_treeview.currentItem())  # 현재 선택된 값의 인덱스를 반환
            self.win_gift_treeview.takeTopLevelItem(currend_index)                                                # 선택된 익덱스 값 제거

            self.calculate_fee()
        except Exception as ex:
            print(ex)

    """===================================== CRUD google sheet ===================================== """

    def make_list_for_uploading(self, time_stamp, first_row, edit_token = False):

        model_header = self.find_model_header()
        num_of_model_colomn = self.win_model_treeview.columnCount()
        model_info_list = [self.win_model_treeview.topLevelItem(0).text(j) for j in range(num_of_model_colomn)]  # 리스트 형태
        model_info_dict = {model_header[j]: model_info_list[j] for j in range(len(model_info_list))}

        model_info_dict["비고(고객메모)"] = self.win_customer_remark.toPlainText()
        model_info_dict["비고(직원메모)"] = self.win_employee_remark.toPlainText()
        model_info_dict["타임스탬프"] = time_stamp
        model_info_dict["접수경로"] = self.win_order_path.currentText()
        if self.win_call.checkState(): model_info_dict["다이렉트 주문"] = "yes"
        else: model_info_dict["다이렉트 주문"] = "no"
        if edit_token:
            model_info_dict["최초 작성자"] = self.win_first.text()
        else:
            model_info_dict["최초 작성자"] = self.comboBox.currentText()
        model_info_dict["최종 수정인"] = self.comboBox.currentText()
        model_info_dict["주문번호(ESM)"] = self.win_order_number.text()
        model_info_dict["차액"] = self.win_fees.text()

        # gift
        num_of_gift = self.win_gift_treeview.topLevelItemCount()

        gift_header = self.find_gift_header()


        for i in range(num_of_gift):
            for j, header in enumerate(gift_header):
                model_info_dict[header + "_" + str(i+1)] = self.win_gift_treeview.topLevelItem(i).text(j)
                model_info_dict["사은품 주소_" + str(i+1)] = self.win_gift_address.text()


        # remittance
        model_info_dict["송금액"] = self.win_payback_money.text()
        model_info_dict["은행"] =  self.win_bank_name.text()
        model_info_dict["계좌번호"] = self.win_account_number.text().replace("-", "")
        model_info_dict["예금주"] = self.win_account_owner.text()

        if edit_token:
            model_info_dict["최초 작성자"] = self.win_first.text()


        list_for_gs = []
        error_list = []
        for cell in first_row:
            try:
                if model_info_dict[cell] != "":
                        list_for_gs.append(model_info_dict[cell])
                else:
                    list_for_gs.append("")
            except Exception as e:
                print("does not exist ", e)
                list_for_gs.append("")
                error_list.append(cell)
        return list_for_gs

    # append buution
    def google_sheet_append_btn_click(self):
        try:
            # 수정 상태 일때
            if self.win_time.text() != "": return 0
            if self.win_model_treeview.topLevelItemCount() == 0: return 0



            time_stamp = datetime.datetime.now().strftime('%Y. %m. %d %p %I:%M:%S (%f)')
            time_stamp = time_stamp.replace("PM", "오후").replace("AM", "오전")

            list_for_uploading = self.make_list_for_uploading(time_stamp = time_stamp, \
                                                                          first_row = main_sheet_first_row)

            try:
                last_index = len(google_sheet.ms.col_values(1))
            except Exception as ex:
                google_sheet.refresh_data()
                last_index = len(google_sheet.ms.col_values(1))

            google_sheet.ms.insert_row(list_for_uploading, last_index + 1, value_input_option='USER_ENTERED')
            self.defalt_btn_click()


        except Exception as ex:

            text = "Error, cannot append to gs !! \n {}".format(ex)
            self.textEdit.setText(text)

    # modify buttion
    def motify_btn_click(self):

        if not self.win_time.text() :
            self.textEdit.setText("수정 중이 아닌 거 같습니다.")
            return

        try:
            time_stamp = self.win_time.text()

            list_for_update = self.make_list_for_uploading(time_stamp=time_stamp, \
                                                              first_row=main_sheet_first_row,
                                                           edit_token= True)

            try:
                time_stamp_list = google_sheet.ms.col_values(1)
            except Exception as ex:
                google_sheet.refresh_data()
                time_stamp_list = google_sheet.ms.col_values(1)

            if time_stamp in time_stamp_list:
                index = time_stamp_list.index(time_stamp)
            else:
                self.textEdit.setText("일치하는 timestamp가 없습니다.")
                return

            cell_list = google_sheet.ms.range(index + 1, 1, index + 1, len(main_sheet_first_row))

            for cell, data in zip(cell_list, list_for_update):
                cell.value = data

            google_sheet.ms.update_cells(cell_list, value_input_option='USER_ENTERED')

            self.defalt_btn_click()


        except Exception as ex:
            print(ex)
            text = "Error, cannot modify gs : {}".format(ex)
            self.textEdit.setText(text)

    # find buttion
    def find_btn_click(self):
        try:
            dlg = dlg_find.find_dlg()
            dlg.editor = self.comboBox.currentText()
            dlg.exec_()

            if not dlg.normal_shutdown_token : return 0

            self.defalt_btn_click()


            customer_info = dlg.customer_info

            # model
            model_header = self.find_model_header()
            list_for_model_tree = []
            for header in model_header :
                if header in customer_info.keys() : list_for_model_tree.append(customer_info[header])
                else: list_for_model_tree.append("")
            self.win_model_treeview.addTopLevelItem(QTreeWidgetItem(list_for_model_tree))



            # gift
            gift_header = self.find_gift_header()

            for j in range(1, 6):
                line = []
                for header in gift_header:
                    info = customer_info.get(header + "_" + str(j), "")
                    line.append(info)

                if not line[gift_header.index("사은품")]:
                    continue

                self.win_gift_treeview.addTopLevelItem(QTreeWidgetItem(line))



            # etc
            self.win_customer_remark.setText(customer_info["비고(고객메모)"])
            self.win_employee_remark.setText(customer_info["비고(직원메모)"])
            self.win_time.setText(customer_info["타임스탬프"])
            self.win_enroll.setText(customer_info["전산 등록인"])
            self.win_last.setText(customer_info["최종 수정인"])
            self.win_first.setText(customer_info["최초 작성자"])
            self.win_order_number.setText(customer_info["주문번호(ESM)"])
            self.win_order_path.setCurrentText(customer_info['접수경로'].strip())

            self.win_gift_address.setText(customer_info["사은품 주소_1"])
            self.win_payback_money.setText(customer_info["송금액"])
            self.win_bank_name.setText(customer_info["은행"])
            self.win_account_number.setText(customer_info["계좌번호"])
            self.win_account_owner.setText(customer_info["예금주"])
            self.win_remittance_state.setText(customer_info["송금 여부"])
            self.win_remittance_state_2.setText(customer_info["송금 비고란"])
            self.win_fees.setText(customer_info["차액"])


            if "yes" in customer_info['다이렉트 주문']: self.win_call.setCheckState(2)
            else : self.win_call.setCheckState(0)


        except Exception as ex:
            text = "Error during lookup : {}".format(ex)
            self.textEdit.setText(text)

    "===================================== etc ===================================== "

    def defalt_btn_click(self):
        self.win_time.setText("")
        self.win_first.setText("")
        self.win_last.setText("")
        self.win_enroll.setText("")
        self.textEdit.setText("")
        self.win_customer_remark.setText("")
        self.win_employee_remark.setText("")
        self.win_gift_address.setText("")
        self.win_payback_money.setText("")
        self.win_bank_name.setText("")
        self.win_account_number.setText("")
        self.win_account_owner.setText("")
        self.win_remittance_state.setText("")
        self.win_remittance_state_2.setText("")
        self.win_order_number.setText("")
        self.win_fees.setText("")
        self.win_call.setCheckState(0)
        self.win_model_treeview.clear()
        self.win_gift_treeview.clear()
        self.customer_info = {x : "" for x in self.customer_info}

    def order_state_edit(self):

        model_header = self.find_model_header()
        order_state_index = model_header.index("주문 상태")
        current_order_state = self.win_order_state.currentText()
        current_item = self.win_model_treeview.currentItem()
        if not current_item : return 0
        current_item.setText(order_state_index, current_order_state)

    def enroll_complete(self):
        dlg = dlg_enroll_complete.install_complete()
        dlg.win_enroll_man.setText(self.comboBox.currentText())
        dlg.exec_()

    def install_complete(self):
        dlg = dlg_install_complete.install_complete_dlg()
        dlg.exec_()

    def calculate_fee(self):
        try:

            fee_index = self.find_model_header().index("수수료")
            model_fee = int(self.win_model_treeview.topLevelItem(0).text(fee_index).replace(",", ""))


            fee_index = self.find_gift_header().index("단가")
            fee = 0
            for i in range(self.win_gift_treeview.topLevelItemCount()):
                fee += int(self.win_gift_treeview.topLevelItem(i).text(fee_index).replace(",", ""))
            gift_fee = fee

            if self.win_payback_money.text() :
                payback_fee = int(self.win_payback_money.text())
            else:
                payback_fee = 0

            final_fee = model_fee - gift_fee - payback_fee
            self.win_fees.setText(str(final_fee))

        except Exception as ex:
            print("Error, cannot calculate fee", ex)
            self.win_fees.setText("계산불가")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    Window = Main_window()
    Window.show()
    app.exec_()

