import os
import re
import pandas as pd
import pywhatkit as kit
import datetime
import time
from dotenv import load_dotenv

load_dotenv()

ASSETS = './assets/'
EXCEL = os.getenv("EXCEL")
EXCEL_PATH = ASSETS + EXCEL
MESSAGE = os.getenv("MESSAGE")

def get_current_minute():
    current_time = datetime.datetime.now()
    hour = current_time.hour
    minute = current_time.minute
    return hour, minute

def normalize(txt):
    txt = str(txt)
    txt = txt.replace('-', '')
    txt = txt.replace(' ', '')
    txt = txt.strip()
    txt = txt.strip('-')
    txt = txt.lstrip()
    if re.match(r'^\+?56', txt):
        txt = txt.replace('+56', '')
        txt = txt.replace('56', '')
        txt = txt = '+56' + txt
        return txt
    else:
        if re.match(r'^\+?54', txt):
            txt = txt.replace('+549', '')
            txt = txt.replace('+54', '')
            txt = txt.replace('54', '')
            txt = txt.replace('549', '')
        txt = '+549' + txt
        return txt

df = pd.read_excel(EXCEL_PATH)
df = df.fillna('')

for index, row in df.iterrows():
    if row['wsp'] != 'yes' and row['wsp'] != 'fail':
        name = row['Nombre']
        phone_work = row['Teléfono laboral']
        phone_home = row['Teléfono laboral']
        phone_number = phone_home if (phone_home) else phone_work
        if phone_number and phone_number != '':
            phone_number = normalize(phone_number)
            message = f"👋 Hola {name},{MESSAGE}"
            hour, minute = get_current_minute()
            next_minute = minute + 1
            try:
                kit.sendwhatmsg(phone_number, message, hour, next_minute, tab_close=True)
                print(phone_number, message, hour, next_minute)
                # df.loc[index, 'wsp'] = 'yes'
            except Exception as e:
                print(f"An error occurred: {e}")
                # df.loc[index, 'wsp'] = 'fail'
        # df.to_excel(EXCEL_PATH, index=False)
        # time.sleep(15)
