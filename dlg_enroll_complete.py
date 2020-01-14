# dlg_enroll_complete.py

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import PyQt5.QtWidgets as QtWidgets
from PyQt5 import uic
from functools import partial
import re
import sys

from credentials import google_sheet, main_sheet_first_row



class install_complete(QDialog):

    def __init__(self):
        QDialog.__init__(self)
        self.ui = uic.loadUi("setting/dialog_enroll_complete.ui", self)

        """================================ size setting===================================="""
        self.treeWidget.setColumnWidth(0, 130)  # 타임스태프
        self.treeWidget.setColumnWidth(1, 45)  # 주문자 번호
        self.treeWidget.setColumnWidth(2, 100)  # 등록 번호
        self.treeWidget.setColumnWidth(3, 50)  # 비고
        self.treeWidget.setColumnWidth(5, 120)  # 제품사 명
        self.treeWidget.setColumnWidth(6, 100)  # 주문 상태
        self.treeWidget.setColumnWidth(7, 130)  # 설치 요청일
        self.treeWidget.setColumnWidth(8, 250)  # 모델명
        self.treeWidget.setColumnWidth(9, 100)  # 렌탈료
        self.treeWidget.setColumnWidth(10, 55)  # 색상
        self.treeWidget.setColumnWidth(11, 65)  # 패키지 or 단품
        self.treeWidget.setColumnWidth(12, 70)  # 계약자
        self.treeWidget.setColumnWidth(13, 130)  # 연락처
        self.treeWidget.setColumnWidth(14, 90)  # 생년월일
        self.treeWidget.setColumnWidth(15, 400)  # 설치 주소
        self.treeWidget.setColumnWidth(16, 30)  # 조리수 벨브
        self.treeWidget.setColumnWidth(17, 200)  # 카드사/은행명
        self.treeWidget.setColumnWidth(18, 300)  # 카드번호/계좌번호


        """==================================================================================="""

        self.win_find_btn.clicked.connect(partial(self.find_model, state = "등록 대기"))
        self.win_reservation_btn.clicked.connect(partial(self.find_model, state = "예약"))
        self.win_modify_standby.clicked.connect(partial(self.edit_state, wanted_state = "등록 대기"))
        self.win_modify_complete.clicked.connect(partial(self.edit_state, wanted_state="등록 완료"))
        self.win_modify_reservation.clicked.connect(partial(self.edit_state, wanted_state="예약"))
        self.win_data_modify.clicked.connect(self.google_sheet_edit)
        self.num_of_find = 0

    # get all data
    def get_data_from_google_sheet(self):
        self.num_of_find = int(self.win_num_of_find_combobox.currentText())
        try : last_num = len(google_sheet.ms.col_values(col = 1, value_render_option='FORMATTED_VALUE')) + 1
        except :
            google_sheet.ms_refresh()
            last_num = len(google_sheet.ms.col_values(col = 1, value_render_option='FORMATTED_VALUE')) + 1

        first_num = last_num - self.num_of_find
        if first_num < 2 : first_num = 2

        header_len = len(main_sheet_first_row)
        num_of_data = last_num - first_num
        g_range = google_sheet.ms.range(first_num, 1, last_num, header_len)
        g_range = [g_range[i*header_len : (i+1)*header_len] for i in range(num_of_data)]
        model_dict_list = []
        for cells in g_range:
            model_dict = {header : cells[i].value for i, header in enumerate(main_sheet_first_row)}
            model_dict_list.append(model_dict)

        return g_range, model_dict_list

    def find_model_header(self):
        headerItem = self.treeWidget.headerItem()
        num_of_column = headerItem.columnCount()

        model_header = []
        for i in range(num_of_column):
            x = headerItem.text(i)
            model_header.append(x)
        return model_header

    def make_list_for_treewidget(self, model_dict):
        model_header = self.find_model_header()
        list = []
        for header in model_header:
            if header in model_dict.keys(): list.append(model_dict[header])
            else : list.append("")
        return list

    def find_model(self, state, company = ""):
        try:
            self.treeWidget.clear()
            _, model_dict_list = self.get_data_from_google_sheet()
            model_header = self.find_model_header()
            order_state_index = model_header.index("주문 상태")
            if company == "" : wanted_company = self.comboBox.currentText()
            else : wanted_company = company
            self.company_backup = wanted_company
            self.order_state_backup = state
            company_index = model_header.index("제품사")
            for model_dict in model_dict_list:
                if not re.findall(r'(\d{6})', model_dict["타임스탬프"]):
                    continue

                list_for_model_tree = self.make_list_for_treewidget(model_dict)
                if list_for_model_tree[order_state_index] != state: continue
                if list_for_model_tree[company_index] != wanted_company and wanted_company != "모두" : continue
                self.treeWidget.addTopLevelItem(QTreeWidgetItem(list_for_model_tree))

        except Exception as ex:
            print(ex)
            text = "Error raise " + str(ex)
            self.textEdit.setText(text)

    def edit_state(self, wanted_state):
        model_header = self.find_model_header()
        order_state_index = model_header.index("주문 상태")
        items = self.treeWidget.selectedItems()
        for item in items:
            item.setText(order_state_index, wanted_state)

    def first_and_last_index(self, g_range):


        for i, x in enumerate(g_range):
            if x[0].value == self.treeWidget.topLevelItem(0).text(0): first_index = i; break;

        index = self.treeWidget.topLevelItemCount()
        g_range.reverse()
        for i, x in enumerate(g_range):
            if x[0].value == self.treeWidget.topLevelItem(index - 1).text(0) : last_index = len(g_range) - i - 1; break;

        g_range.reverse()

        return first_index, last_index

    def make_customer_info_dict_list(self):
        num_model = self.treeWidget.topLevelItemCount()
        model_header = self.find_model_header()
        num_of_column = len(model_header)

        customer_info = []
        for i in range(num_model):
            item = self.treeWidget.topLevelItem(i)
            info = {model_header[j]: item.text(j) for j in range(num_of_column)}
            customer_info.append(info)

        return customer_info

    def google_sheet_edit(self):
        try:
            if self.num_of_find == 0 : return 0
            self.num_of_find += 100  # load more

            g_range, _ = self.get_data_from_google_sheet()
            time_stamp_index = main_sheet_first_row.index("타임스탬프")
            order_state_index = main_sheet_first_row.index("주문 상태")
            enroll_index = main_sheet_first_row.index("전산 등록인")


            first_index, last_index = self.first_and_last_index(g_range = g_range)
            g_range = g_range[first_index : last_index + 1]

            customer_info = self.make_customer_info_dict_list()

            time_stamp_list = []
            order_state_list = []
            for info in customer_info:
                time_stamp_list.append(info["타임스탬프"])
                order_state_list.append(info["주문 상태"])


            edit_cells_state = []
            edit_cells_name = []
            for cells in g_range:

                stamp = cells[time_stamp_index].value
                if stamp in time_stamp_list:
                    index = time_stamp_list.index(stamp)
                    cells[order_state_index].value = order_state_list[index]

                if cells[order_state_index].value == "등록 완료":
                    cells[enroll_index].value = self.win_enroll_man.text()

                edit_cells_state.append(cells[order_state_index])
                edit_cells_name.append(cells[enroll_index])

            google_sheet.ms.update_cells(edit_cells_state, value_input_option='RAW')
            google_sheet.ms.update_cells(edit_cells_name, value_input_option='RAW')
            self.treeWidget.clear()

            self.find_model(state= self.order_state_backup, company= self.company_backup)


        except Exception as ex:
            print(ex)
            text = "오류 발생 ! : " + str(ex)
            self.textEdit.setText(text)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = install_complete()
    myWindow.show()
    app.exec_()
