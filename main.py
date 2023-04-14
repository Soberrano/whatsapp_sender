

import time
import pywhatkit
import pyautogui
import numpy as np
import pandas as pd
import openpyxl
from pynput.keyboard import Key, Controller

keyboard = Controller()



# def send_whatsapp_message(msg: str, num: str):
#     try:
#         pywhatkit.sendwhatmsg_instantly(
#             phone_no=num,
#             message=msg
#             # tab_close=True
#         )
#         #time.sleep(2)
#         #pyautogui.click()
#         # time.sleep(2)
#         # keyboard.press(Key.enter)
#         # keyboard.release(Key.enter)
#         # print("Message sent!")
#     except Exception as e:
#         print(str(e))
#
#
# dataframe = pd.read_excel('bd.xlsx')
# data = pd.DataFrame(dataframe, columns=['Num', 'Fio', 'Recomend'])
# print(data['Fio'].iloc[0])
# if __name__ == "__main__":
#     for i in range(5):
#         send_whatsapp_message(num = '+' + (str(data['Num'].iloc[i])),
#                               #
#                               msg = "Уважаемые родители, "
#                                     "Ваш ребенок " + str(data['Fio'].iloc[i]) +
#                                     " в течении 2022 - 2023 учебного года посещал занятия в Детском технопарке Кванториум"
#                                     "\nУже совсем скоро станут известны итоги его обучения, а пока мы уже открываем запись на следующий учебный год!"
#                                     "\nНаставник вашего ребенка составил персональные рекомендации на следующий учебный год."
#                                     "\nВсе рекомендации составлены исходя из наблюдениий педагога и пожеланий ученика. "
#                                     "\nИзучив их, Вам будет проще сделать правильный выбор."
#                                     "\n\nМы предлагаем Вам и вашему ребенку следующие направления и курсы: \n"  + "*" + str(data['Recomend'].iloc[i])+"*" +
#                                     "\n\nЗаписаться на эти направления, узнать о них подробнее или посмотреть, какие программы есть ещё, Вы можете, пройдя по ссылке:"
#                                     "\nЗапись на направления Кванториума: https://kvantorium.stavdeti.ru/apply_for_training/ "
#                                     "\nЗапись на направления IT-куба: https://itcube.stavdeti.ru/apply_for_training/  "
#                                     "\nПорядок зачисления на 2023 - 2024 учебный год:"
#                                     "\nВы подаете заявку через наш сайт: https://kvantorium.stavdeti.ru/apply_for_training/ "
#                                     "\nЗаявка бронирует место для вашего ребенка в выбранной группе. Бронь действительна 14 календарных дней."
#                                     "\nВ течении этого срока Вам необходимо подать заявление на обучение ЛИЧНО "
#                                     "\nпо адресу г.Михайловск, ул.Привокзальная, д.3. в будни с 9:00 ДО 16:00"
#                                     "\n\nЕсли у Вас остались вопросы, звоните по номеру + 7(8652) 33 - 33 - 83"
#                               )

dataframe = pd.read_excel('bd.xlsx')
data = pd.DataFrame(dataframe, columns=['Num', 'Fio', 'Recomend','link','res'])
for i in range(0, 97):
    if data['Recomend'].iloc[i] != "":
        a = "Уважаемые родители, ваш ребенок " + str(data['Fio'].iloc[i]) +\
        " в течении 2022 - 2023 учебного года посещал занятия в Детском технопарке Кванториум."\
        "\nУже совсем скоро станут известны итоги его обучения, а пока мы уже открываем запись на следующий учебный год!"\
        "\nНаставник вашего ребенка составил персональные рекомендации на следующий учебный год."\
        "\nВсе рекомендации составлены исходя из наблюдениий педагога и пожеланий ученика. "\
        "\nИзучив их, Вам будет проще сделать правильный выбор."\
        "\n\nМы предлагаем Вам и вашему ребенку следующие направления и курсы: \n"  + "*" + str(data['Recomend'].iloc[i])+"*" +\
        "\n\nЗаписаться на эти направления, узнать о них подробнее или посмотреть, какие программы есть ещё, Вы можете, пройдя по ссылке:"\
        "\nЗапись на направления Кванториума: https://kvantorium.stavdeti.ru/apply_for_training/ "\
        "\nЗапись на направления IT-куба: https://itcube.stavdeti.ru/apply_for_training/  "\
        "\nПорядок зачисления на 2023 - 2024 учебный год:"\
        "\nВы подаете заявку через наш сайт: https://kvantorium.stavdeti.ru/apply_for_training/ "\
        "\nЗаявка бронирует место для вашего ребенка в выбранной группе. Бронь действительна 14 календарных дней."\
        "\nВ течении этого срока, Вам необходимо подать заявление на обучение ЛИЧНО "\
        "\nпо адресу г.Михайловск, ул.Привокзальная, д.3. в будни с 9:00 ДО 16:00"\
        "\n\nЕсли у Вас остались вопросы, звоните по номеру + 7(8652) 33 - 33 - 83"
    # https://api.whatsapp.com/send/?phone=79286313998
    data.loc[i,'link'] = "https://api.whatsapp.com/send/?phone="+ str(data['Num'].iloc[i])
    data.loc[i,'res'] = a
data.to_excel('bd.xlsx')





