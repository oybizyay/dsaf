import openpyxl
import warnings
import datetime
import os
import csv


def ekvr():
    warnings.simplefilter("ignore")

    lst = []
    sm = 0

    book_new = openpyxl.open("ТК/Отчет Эквайринг от 03.07.23.xlsx")
    sheet_new = book_new.active

    with open("пул/e1.csv", 'r', newline='', encoding="UTF-8") as csv_file:
        reader = csv.reader(csv_file, delimiter=';')
        for item in reader:
            if item[7] == "Оплата":
                if item[14] == "Завершена":
                    s = item[8]
                    s = s.split(",")
                    s = s[0] + "." + s[1]
                    s = float(s)
                    sm += s
                    lst.append([item[4], s])

    if os.path.exists("пул/e2.csv"):
        with open("пул/e2.csv", 'r', newline='', encoding="UTF-8") as csv_file:
            reader = csv.reader(csv_file, delimiter=';')
            for item in reader:
                if item[7] == "Оплата":
                    if item[14] == "Завершена":
                        s = item[8]
                        s = s.split(",")
                        s = s[0] + "." + s[1]
                        s = float(s)
                        sm += s
                        lst.append([item[4], s])

    if os.path.exists("пул/e3.csv"):
        with open("пул/e3.csv", 'r', newline='', encoding="UTF-8") as csv_file:
            reader = csv.reader(csv_file, delimiter=';')
            for item in reader:
                if item[7] == "Оплата":
                    if item[14] == "Завершена":
                        s = item[8]
                        s = s.split(",")
                        s = s[0] + "." + s[1]
                        s = float(s)
                        sm += s
                        lst.append([item[4], s])

    if os.path.exists("пул/e4.csv"):
        with open("пул/e4.csv", 'r', newline='', encoding="UTF-8") as csv_file:
            reader = csv.reader(csv_file, delimiter=';')
            for item in reader:
                if item[7] == "Оплата":
                    if item[14] == "Завершена":
                        s = item[8]
                        s = s.split(",")
                        s = s[0] + "." + s[1]
                        s = float(s)
                        sm += s
                        lst.append([item[4], s])

    if os.path.exists("пул/e5.csv"):
        with open("пул/e5.csv", 'r', newline='', encoding="UTF-8") as csv_file:
            reader = csv.reader(csv_file, delimiter=';')
            for item in reader:
                if item[7] == "Оплата":
                    if item[14] == "Завершена":
                        s = item[8]
                        s = s.split(",")
                        s = s[0] + "." + s[1]
                        s = float(s)
                        sm += s
                        lst.append([item[4], s])

    for i in range(1, len(lst) + 1):
        # i принимает значения от 1 до 17, пишем со второй строки I+1, начиная с 0 индекса списка, 0 индекса подсписка I-1
        # 1,len(lst)+1 позваляет перебирать значения от 1 а не от 1, и включительно по число длины списка
        sheet_new[i + 1][1].value = lst[i - 1][0]
        sheet_new[i + 1][7].value = lst[i - 1][1]
        sheet_new[i + 1][8].value = 3

    today = datetime.date.today()
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday != 0:
        yesterday = today - datetime.timedelta(days=1)
    else:
        yesterday = today - datetime.timedelta(days=3)

    book_new.save("res/Отчет Эквайринг от %s %s.xlsx" % (yesterday, round(sm, 2)))
    book_new.close()