

import time
import pywhatkit
import pyautogui
import numpy as np
import pandas as pd
import openpyxl
from pynput.keyboard import Key, Controller

keyboard = Controller()



def send_whatsapp_message(msg: str, num: str):
    try:
        pywhatkit.sendwhatmsg_instantly(
            phone_no=num,
            message=msg,
            tab_close=True
        )
        time.sleep(2)
        pyautogui.click()
        time.sleep(2)
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        print("Message sent!")
    except Exception as e:
        print(str(e))


dataframe = pd.read_excel('bd.xlsx')
data = pd.DataFrame(dataframe, columns=['Num', 'Fio'])
print(data['Fio'].iloc[0])
if __name__ == "__main__":
    for i in range(5):
        send_whatsapp_message(num = '+' + (str(data['Num'].iloc[i])) ,msg=data['Fio'].iloc[i])







