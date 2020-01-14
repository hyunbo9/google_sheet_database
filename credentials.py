#credentials.py

from setting_from_excel import google_sheet_name, main_sheet_name, model_sheet_name, gift_sheet_name
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import glob, random



class Google_sheet:
    def __init__(self, credentials, google_sheet_name, main_sheet_name, model_sheet_name, gift_sheet_name):
        self.credentials = credentials
        self.google_sheet_name = google_sheet_name
        self.main_sheet_name = main_sheet_name
        self.model_sheet_name = model_sheet_name
        self.gift_sheet_name = gift_sheet_name
        self.get_data()

    def get_data(self):
        try:
            gc = gspread.authorize(self.credentials)
            gs = gc.open(self.google_sheet_name)  # google_sheet
            self.ms = gs.worksheet(self.main_sheet_name)  # main_sheet
            self.model_sheet = gs.worksheet(self.model_sheet_name)
            self.gift_sheet = gs.worksheet(self.gift_sheet_name)
        except Exception as e:
            print("error : ",  e)
            raise Exception("Error raise. While first signing in")


    # Made for re-login when not in use for a long time. or some login error
    def refresh_data(self, number_of_attenpts = 3):
        if number_of_attenpts == 0 :
            raise Exception("Error raise. Try to 3 times. But can not login")

        print(number_of_attenpts, "th attempt")

        try:
            gc = gspread.authorize(self.credentials)
            gs = gc.open(self.google_sheet_name)  # google_sheet
            self.ms = gs.worksheet(self.main_sheet_name)  # main_sheet
        except Exception as e:
            print("error : ", e)
            self.refresh_data(number_of_attenpts -1)



scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']


key_path = glob.glob("google_key/*.json")
key = random.choice(key_path)
credentials = ServiceAccountCredentials.from_json_keyfile_name(key, scope)

google_sheet = Google_sheet(credentials = credentials, google_sheet_name = google_sheet_name, \
                       main_sheet_name = main_sheet_name, model_sheet_name = model_sheet_name, gift_sheet_name = gift_sheet_name)

#pruda_v3 first line
main_sheet_first_row = google_sheet.ms.row_values(row = 1, value_render_option = 'FORMATTED_VALUE')
