
import requests
import pathlib
import apimoex
import pandas as pd

board = 'TQBR'

print("Здравствуйте! Пожалуйста, сообщите программе, что вам нужно")
print("Введите \"Таблица\" чтобы получить данные о котировках в EXCEL таблице")
print("Введите \"Акция\" чтобы получить данные о последних ценах на конкретную бумагу")

choice = input()

if choice.casefold() == "акция":
    print(" Введите нужный вам тикер: ")
    ticker = input()

    with requests.Session() as session:
        try:
            data = apimoex.get_board_history(session, ticker)# Выгружаем необхомиые данные
            df = pd.DataFrame(data) #Преобразуем

            print("Укажите,что вас интересует(количество сделок,цена,стоимость всех операций)"
                  " или, если необходимы все данные, укажите 'всё,")
            information = input()

            if information.casefold() == "количество сделок":
                df = df[['TRADEDATE', 'VOLUME']]#Выбираем нужные столбцы с данными
                print(df.tail(), '\n')
                df.info()

            if information.casefold() == "цена":
                df = df[['TRADEDATE', 'CLOSE']]#Выбираем нужные столбцы с данными
                print(df.tail(), '\n')
                df.info()

            if information.casefold() == "стоимость всех операций":
                df = df[['TRADEDATE', 'VALUE']]#Выбираем нужные столбцы с данными
                print(df.tail(), '\n')
                df.info()

            if information.casefold() == "всё":
                df = df[['TRADEDATE', 'CLOSE', 'VOLUME', 'VALUE']]#Выбираем нужные столбцы с данными
                print(df.tail(), '\n')
                df.info()

            else:
                print("Введённый вами параметр не существует, проверьте правильность написания")

        except KeyError:
            print("Тикет не существует")

if choice.casefold() == "таблица":
    with open("C:/PYEX/TICK.txt", "r") as TICKs: # Открвыаем файл в режиме чтения
        TICKs = [line.rstrip() for line in TICKs] # Убираем пробелы в строке
    pathlib.Path("C:/PYEX/Database/{}".format(board)).mkdir(parents=True, exist_ok=True) #Указываем путь и, в случае необходимости, создаем отсутствующие каталоги
    process = 0
    with requests.Session() as session: #Создаем подключение
        for TICK in TICKs:
             process = process + 1
             print((process / len(TICKs)) * 100, ' %')
             data = apimoex.get_board_history(session, TICK, board=board)#История котировок
             if data == []:
                 continue
             df = pd.DataFrame(data)# Преобразуем полученные данные в Dataframe
             df = df[['TRADEDATE', 'CLOSE']] #Выбираем нужные столбцы
             df.to_excel("C:/PYEX/Database/{}/{}.xlsx".format(board, TICK), index=False) #Сохраняем в Excel

else:
    print("Пожалуйста, введите запрос коректно")