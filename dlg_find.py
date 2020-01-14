# dlg_find.py

from credentials import google_sheet, main_sheet_first_row
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
import PyQt5.QtWidgets as QtWidgets
import sys
from PyQt5 import uic
import numpy as np
import time, datetime


class find_dlg(QDialog):

    def __init__(self):
        QDialog.__init__(self)
        self.ui = uic.loadUi("setting/dialog_find.ui", self)
        self.ui.show()
        self.find_order_state.clicked.connect(self.order_state_append_btn_click)
        self.find_one_hundred.clicked.connect(self.find_one_hundred_btn_click)
        self.all_find_btn.clicked.connect(self.all_find_btn_click)
        self.treeWidget.setColumnWidth(0, 110)

        self.treeWidget.itemDoubleClicked.connect(self.win_append)
        self.normal_shutdown_token = False
        self.editor = ""

    def find_header(self):
        headerItem = self.treeWidget.headerItem()
        num_of_column = headerItem.columnCount()

        header = []
        for i in range(num_of_column):
            x = headerItem.text(i)
            header.append(x)

        return header

    def make_index_list(self):
        header_list = self.find_header()

        index_list = [main_sheet_first_row.index(x) for x in header_list \
                      if x in main_sheet_first_row]           # header 오류 나면 안됨.
        return index_list

    def make_wanted_range_cells(self, num_of_find = 10000):

        try: last_index = len(google_sheet.ms.col_values(1))
        except Exception as ex:
            google_sheet.refresh_data();
            last_index = len(google_sheet.ms.col_values(1))
        first_index = last_index - num_of_find
        if first_index < 2: first_index = 2
        col_index = len(main_sheet_first_row)

        try:
            wanted_range_cells = google_sheet.ms.range(first_index, 1, \
                                                       last_index, col_index)
        except Exception as ex:
            google_sheet.refresh_data()
            wanted_range_cells = google_sheet.ms.range(first_index, 1, \
                                                       last_index, col_index)

        num_of_model = int(len(wanted_range_cells) / col_index)
        wanted_range_cells = [wanted_range_cells[i * col_index: (i + 1) * col_index] for i in
                              range(num_of_model)]

        return np.array(wanted_range_cells)

    def set_none(self):
        headerItem = self.treeWidget.headerItem()
        num_of_column = headerItem.columnCount()
        self.treeWidget.addTopLevelItem(QTreeWidgetItem(['없음'] + [""] * (num_of_column - 1)))

    def order_state_append_btn_click(self):
        try:

            self.treeWidget.clear()

            index_list = self.make_index_list()
            wanted_value = self.order_state.currentText()
            order_state_index = main_sheet_first_row.index("주문 상태")
            wanted_range_cells = self.make_wanted_range_cells()

            for cells in wanted_range_cells:
                if wanted_value == cells[order_state_index].value:
                    wanted_cells = cells[index_list]
                    list_for_treewidget = [x.value for x in wanted_cells]
                    self.treeWidget.addTopLevelItem(QTreeWidgetItem(list_for_treewidget))
                    continue

            if self.treeWidget.topLevelItemCount() == 0 :
                self.set_none()

        except Exception as ex:
            text = "Error raise : \n" + str(ex)
            self.textEdit.setText(text)

    def all_find_btn_click(self):
        try:
            self.treeWidget.clear()
            index_list = self.make_index_list()
            wanted_value = self.all_text.text()

            if not wanted_value.replace(" ", "") : return 0
            wanted_range_cells = self.make_wanted_range_cells()

            for cells in wanted_range_cells:
                for cell in cells:
                    if wanted_value.replace("-", "").replace(" ", "") \
                                in cell.value.replace("-", "").replace(" ", "") :
                        wanted_cells = cells[index_list]
                        list_for_treewidget = [x.value for x in wanted_cells]
                        self.treeWidget.addTopLevelItem(QTreeWidgetItem(list_for_treewidget))
                        break

            if self.treeWidget.topLevelItemCount() == 0 :
                self.set_none()

        except Exception as ex:
            text = "Error raise \n" + str(ex)
            self.textEdit.setText(text)

    def find_one_hundred_btn_click(self):
        try:
            self.treeWidget.clear()

            index_list = self.make_index_list()
            wanted_range_cells = self.make_wanted_range_cells(num_of_find=100)

            for cells in wanted_range_cells:
                wanted_cells = cells[index_list]
                list_for_treewidget = [x.value for x in wanted_cells]
                self.treeWidget.addTopLevelItem(QTreeWidgetItem(list_for_treewidget))

            if self.treeWidget.topLevelItemCount() == 0 :
                self.set_none()

        except Exception as ex:
            text = "Error raise : \n" + str(ex)
            self.textEdit.setText(text)


    def win_append(self):
        try:
            time_stamp = self.treeWidget.currentItem().text(0)
            time_stamp_index = main_sheet_first_row.index("타임스탬프")

            try:
                time_stamp_list = google_sheet.ms.col_values(col=time_stamp_index + 1, \
                                                             value_render_option='FORMATTED_VALUE')
            except:
                google_sheet.refresh_data()
                time_stamp_list = google_sheet.ms.col_values(col=time_stamp_index + 1, \
                                                             value_render_option='FORMATTED_VALUE')

            if not time_stamp in time_stamp_list:
                raise Exception("No matching information")

            index = time_stamp_list.index(time_stamp)
            g_range = google_sheet.ms.range(index + 1, 1, index + 1, len(main_sheet_first_row))

            self.customer_info = {}
            for cell, name in zip(g_range, main_sheet_first_row):
                self.customer_info[name] = cell.value


            # last access
            time_stamp = datetime.datetime.now().strftime('%m. %d %p %I:%M:%S')
            time_stamp = time_stamp.replace("PM", "오후").replace("AM", "오전")
            text = self.editor + "__({})".format(time_stamp)
            google_sheet.ms.update_cell(index + 1, main_sheet_first_row.index("last_access") + 1, text )

            print("main sheet load complete")
            self.normal_shutdown_token = True
            self.close()

        except Exception as ex:
            text = "Error raise \n" + str(ex)
            self.textEdit.setText(text)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    myWindow = find_dlg()
    myWindow.show()
    app.exec_()