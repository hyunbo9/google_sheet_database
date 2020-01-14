# setting_from_excel.py

from openpyxl import load_workbook
import numpy as np

filename = 'setting.xlsx'
sheetName = 'Setting'

wb = load_workbook(filename)
ws = wb[sheetName]


data_row_list = []
for row in ws:
    data_row = []
    for cell in row:
        data_row.append(cell.value)
    data_row_list.append(data_row)

consultant_index = data_row_list[1].index("상담사")
order_path_index = data_row_list[1].index("접수경로")
google_sheet_name_index = data_row_list[1].index("구글 시트 이름")
main_sheet_name_index = data_row_list[1].index("Main 시트")
model_sheet_name_index = data_row_list[1].index("제품 정보 시트")
gift_sheet_name_index = data_row_list[1].index("사은품 정보 시트")


data_row_list = np.array(data_row_list)

counselor, order_path = [], []
google_sheet_name = data_row_list[3, google_sheet_name_index]
main_sheet_name = data_row_list[3, main_sheet_name_index]
model_sheet_name = data_row_list[3, model_sheet_name_index]
gift_sheet_name = data_row_list[3, gift_sheet_name_index]
consultant_list = [name for name in data_row_list[2:, consultant_index] if name]
order_path_list = [order_path for order_path in data_row_list[2: , order_path_index] if order_path]



if __name__ == "__main__":

    print("consultant_list: ", consultant_list)
    print("order_path_list: ", order_path_list)
    print("google_sheet_name: ", google_sheet_name)
    print("main_sheet_name: ", main_sheet_name)
    print("model_sheet_name: ", model_sheet_name)
    print("gift_sheet_name: ", gift_sheet_name)

