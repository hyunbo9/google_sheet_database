#model_price.py

from credentials import google_sheet
import numpy as np
import time

start_time = time.time()


MODEL_T = np.array(google_sheet.model_sheet.get_all_values())
GIFT_T = google_sheet.gift_sheet.get_all_values()


LG_list, SK_list, HYUNDAI_list, CUCKOO_list, CHUNGHO_list, KYOWON_list, \
        RUHENS_list, WELRIX_list, NOVITA_list, WINIA_list, CHERISH_list = [],[],[],[],[],[],[],[],[],[],[]

for index, text in enumerate(MODEL_T[:, 2]):
    if text == "LG 정수기" or text == "LG 공청기" or text == "LG 기타" or text == "LG 패키지" :
        LG_list.append(index)
    elif text == "SK 정수기" or text == "SK 공청기" or text == "SK 기타" or text == "SK 패키지" :
        SK_list.append(index)
    elif text == "현대 정수기" or text == "현대 공청기" or text == "현대 기타" or text == "현대 패키지":
        HYUNDAI_list.append(index)
    elif text == "쿠쿠 정수기" or text == "쿠쿠 공청기" or text == "쿠쿠 기타" or text == "쿠쿠 패키지":
        CUCKOO_list.append(index)
    elif text == "청호 정수기" or text == "청호 공청기" or text == "청호 기타" or text == "청호 패키지":
        CHUNGHO_list.append(index)
    elif text == "교원 정수기" or text == "교원 공청기" or text == "교원 기타" or text == "교원 패키지":
        KYOWON_list.append(index)
    elif text == "루헨스 정수기" or text == "루헨스 공청기" or text == "루헨스 기타" or text == "루헨스 패키지":
        RUHENS_list.append(index)
    elif text == "웰릭스 정수기" or text == "웰릭스 공청기" or text == "웰릭스 기타" or text == "웰릭스 패키지":
        WELRIX_list.append(index)
    elif text == "노비타 정수기" or text == "노비타 공청기" or text == "노비타 기타" or text == "노비타 패키지":
        NOVITA_list.append(index)
    elif text == "위니아 정수기" or text == "위니아 공청기" or text == "위니아 기타" or text == "위니아 패키지":
        WINIA_list.append(index)
    elif text == "체리쉬 정수기" or text == "체리쉬 공청기" or text == "체리쉬 기타" or text == "체리쉬 패키지":
        CHERISH_list.append(index)


def get_info(MODEL_T, name , index):

    globals()[name + "_name"] = [""]
    globals()[name + "_model"] = [""]
    globals()[name + "_price"] = [""]
    globals()[name + "_package_price"] = [""]
    globals()[name + "_fees"] = [""]
    globals()[name + "_package_fees"] = [""]

    blank_num = 0
    for x in MODEL_T[index + 2:, : 15] :
        if blank_num >= 2: break
        if x[2] ==  "" : blank_num += 1
        else: blank_num = 0

        globals()[name + "_name"].append(x[1])
        globals()[name + "_model"].append(x[2])
        globals()[name + "_price"].append(x[3])
        globals()[name + "_package_price"].append(x[4])
        globals()[name + "_fees"].append(x[5])
        globals()[name + "_package_fees"].append(x[6])





def make_var_from_company_index_list(company, company_list):
    remark_index = company_list[0] + 1
    globals()[company + "_remarks"] = MODEL_T[remark_index][0]

    if not company_list : return

    get_info(MODEL_T= MODEL_T, name = company + "_water_purifier", index = company_list[0])
    get_info(MODEL_T=MODEL_T, name=company + "_air_cleaner", index = company_list[1])
    get_info(MODEL_T=MODEL_T, name=company + "_etc", index = company_list[2])
    get_info(MODEL_T=MODEL_T, name=company + "_package", index = company_list[3])


def make_var():
    make_var_from_company_index_list(company="LG", company_list=LG_list)
    make_var_from_company_index_list(company="SK", company_list=SK_list)
    make_var_from_company_index_list(company="HYUNDAI", company_list=HYUNDAI_list)
    make_var_from_company_index_list(company="CUCKOO", company_list=CUCKOO_list)
    make_var_from_company_index_list(company="CHUNGHO", company_list=CHUNGHO_list)
    make_var_from_company_index_list(company="KYOWON", company_list=KYOWON_list)
    make_var_from_company_index_list(company="RUHENS", company_list=RUHENS_list)
    make_var_from_company_index_list(company="WELRIX", company_list=WELRIX_list)
    make_var_from_company_index_list(company="NOVITA", company_list=NOVITA_list)
    make_var_from_company_index_list(company="SAMSUNG", company_list=WINIA_list)
    make_var_from_company_index_list(company="CHERISH", company_list=CHERISH_list)


make_var()

"""========================================= 사은품 불러오기 ========================================="""

gift1_a = [x[0] for x in GIFT_T]
gift1_b = [x[1] for x in GIFT_T]
gift1_c = [x[2] for x in GIFT_T]
gift1_d = [x[3] for x in GIFT_T]

gift2_a = [x[5] for x in GIFT_T]
gift2_b = [x[6] for x in GIFT_T]
gift2_c = [x[7] for x in GIFT_T]
gift2_d = [x[8] for x in GIFT_T]

gift3_a = [x[10] for x in GIFT_T]
gift3_b = [x[11] for x in GIFT_T]
gift3_c = [x[12] for x in GIFT_T]
gift3_d = [x[13] for x in GIFT_T]


gift1_a[0], gift1_b[0], gift1_c[0], gift1_d[0],\
gift2_a[0], gift2_b[0], gift2_c[0], gift2_d[0],\
gift3_a[0], gift3_b[0], gift3_c[0], gift3_d[0] = "", "", "", "", "", "", "", "", "", "", "", ""


