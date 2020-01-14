# dlg_install_complete.py

from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import PyQt5.QtWidgets as QtWidgets
from PyQt5 import uic
from functools import partial
import copy, re
import sys

from credentials import google_sheet, main_sheet_first_row



class install_complete_dlg(QDialog):

    def __init__(self):

        QDialog.__init__(self)
        self.ui = uic.loadUi("setting/dialog_install_complete.ui", self)

        """================================ size setting ===================================="""
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

        """======================================================================================"""

        today = self.time_start.selectedDate()
        self.time_start.setSelectedDate(today.addDays(-100))
        self.win_modify_complete.clicked.connect(partial(self.state_edit, wanted_state = "등록 완료"))
        self.win_modify_install_complete.clicked.connect(partial(self.state_edit, wanted_state = "설치 완료"))
        self.win_modify_cancel.clicked.connect(partial(self.state_edit, wanted_state = "취소"))
        self.win_reconfirm_btn.clicked.connect(partial(self.state_edit, wanted_state = "재확인 요청"))
        self.win_data_modify.clicked.connect(self.google_sheet_edit)
        self.win_find_btn.clicked.connect(self.find_model)

        self.edit_token = False

    def find_model_header(self):
        headerItem = self.treeWidget.headerItem()
        num_of_column = headerItem.columnCount()

        model_header = []
        for i in range(num_of_column):
            x = headerItem.text(i)
            model_header.append(x)

        return model_header

    # get 10000 data
    def make_g_range(self):

        last_index = len(google_sheet.ms.col_values(1))
        first_index = last_index - 10000
        if first_index <2 : first_index =2
        header_num = len(main_sheet_first_row)

        try:
            g_range = google_sheet.ms.range(first_index, 1, last_index, header_num)
        except:
            google_sheet.refresh_data()
            g_range = google_sheet.ms.range(first_index, 1, last_index, header_num)

        model_num = int(len(g_range) / header_num)
        g_range = [g_range[ i * header_num : (i+1) * header_num] for i in range(model_num)]

        return g_range

    def make_model_dict_list(self, g_range):

        model_dict_list = []
        for cells in g_range:
            model_dict = {header : cells[i].value for i, header in enumerate(main_sheet_first_row)}

            if model_dict["제품사"] != self.company and self.company != "모두" :
                continue

            if model_dict["주문 상태"] != self.order_state and self.order_state != "모두" :
                continue

            model_dict_list.append(model_dict)
        return model_dict_list

    def make_list_froms_dict(self, dict):
        model_header = self.find_model_header()

        list_for_tree = []
        for header in model_header:
            if header in dict.keys():
                list_for_tree.append(dict[header])
            else :
                list_for_tree.append("")

        return list_for_tree

    def find_model(self):
        try:
            self.treeWidget.clear()
            self.company = self.comboBox.currentText()
            self.order_state = self.comboBox_2.currentText()
            self.start_date = int(self.time_start.selectedDate().toString('yyyyMMdd'))
            self.end_date = int(self.time_end.selectedDate().toString('yyyyMMdd'))


            g_range = self.make_g_range()
            model_dict_list = self.make_model_dict_list(g_range = g_range)

            for dict in model_dict_list:
                if not re.findall(r'(\d{6})', dict["타임스탬프"]):
                    continue
                stamp = int(dict["타임스탬프"][:12].replace(" ", "").replace(".", ""))
                if stamp < self.start_date:
                    continue
                if stamp > self.end_date:
                    break

                list_for_tree = self.make_list_froms_dict(dict = dict)
                self.treeWidget.addTopLevelItem(QTreeWidgetItem(list_for_tree))


            self.edit_token = True
        except Exception as ex:
            print(ex)
            text = "Error raise : " + str(ex)
            self.textEdit.setText(text)

    def state_edit(self, wanted_state):
        model_header = self.find_model_header()
        order_state_index = model_header.index("주문 상태")
        complete_date = model_header.index("설치 완료일")
        items = self.treeWidget.selectedItems()

        for item in items:
            item.setText(order_state_index, wanted_state)
            if wanted_state == "등록 완료" :
                item.setText(complete_date, "")

            elif wanted_state == "취소":
                reason = self.win_cancel_reason.text()
                item.setText(complete_date, reason)

            elif wanted_state == "설치 완료" :
                date = self.calendarWidget_3.selectedDate().toString('yyyy.MM.dd (ddd)')
                item.setText(complete_date, date)

            elif wanted_state == "재확인 요청":
                reason = self.win_reconfirm_reason.text()
                item.setText(complete_date, reason)

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

        if customer_info == []:
            customer_info.append([])

        return customer_info

    def google_sheet_edit(self):
        try:
            if not self.edit_token : return 0


            g_range = self.make_g_range()
            customer_info = self.make_customer_info_dict_list()

            order_state_index = main_sheet_first_row.index("주문 상태")
            install_date_index = main_sheet_first_row.index("설치 완료일")
            time_stamp_index = main_sheet_first_row.index("타임스탬프")

            time_stamp_list = []
            order_state_list = []
            install_complete_date = []
            for info in customer_info:
                time_stamp_list.append(info["타임스탬프"])
                order_state_list.append(info["주문 상태"])
                install_complete_date.append(info["설치 완료일"])

            first_index, last_index = self.first_and_last_index(g_range)
            g_range = g_range[first_index: last_index + 1]


            edit_cells_state = []
            edit_cells_install_date = []

            for cells in g_range:
                stamp = cells[time_stamp_index].value
                if stamp in time_stamp_list:
                    index = time_stamp_list.index(cells[time_stamp_index].value)
                    cells[order_state_index].value = order_state_list[index]
                    cells[install_date_index].value = install_complete_date[index]

                edit_cells_state.append(cells[order_state_index])
                edit_cells_install_date.append(cells[install_date_index])

            google_sheet.ms.update_cells(edit_cells_state, value_input_option='RAW')
            google_sheet.ms.update_cells(edit_cells_install_date, value_input_option='RAW')

            self.find_model()


        except Exception as ex:
            print(ex)
            text = "Error raise: " + str(ex)
            self.textEdit.setText(text)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = install_complete_dlg()
    myWindow.show()
    app.exec_()