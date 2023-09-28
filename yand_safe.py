import openpyxl
import warnings
import datetime
import os
import csv

def yand_safe():

    warnings.simplefilter("ignore")

    lst = []
    sm = 0

    with open("yyyaaasf.txt", "r") as f:
        first_line = f.readline()
        first_list = first_line.split(" ")

        kontrol = int(first_list[1])
        # print(kontrol)

    pp = first_list[2]
    # print(pp)

    book_new = openpyxl.open("ТК/Отчет Яндекс Доставка НП СЭЙФ от 03.07.2023.xlsx")
    sheet_new = book_new.active

    with open("пул/y1.csv", 'r', newline='', encoding="UTF-16 LE") as csv_file:
        reader = csv.reader(csv_file, delimiter='\t')
        for item in reader:
            if item[-4] == pp:
                sm += float(item[-1])
                lst.append([int(item[1]), float(item[-1])])

    if kontrol > 1:
        pp2 = first_list[3]
        with open("пул/y1.csv", 'r', newline='', encoding="UTF-16 LE") as csv_file:
            reader = csv.reader(csv_file, delimiter='\t')
            for item in reader:
                if item[-4] == pp2:
                    sm += float(item[-1])
                    lst.append([int(item[1]), float(item[-1])])

    if kontrol > 2:
        pp3 = first_list[4]
        with open("пул/y1.csv", 'r', newline='', encoding="UTF-16 LE") as csv_file:
            reader = csv.reader(csv_file, delimiter='\t')
            for item in reader:
                if item[-4] == pp3:
                    sm += float(item[-1])
                    lst.append([int(item[1]), float(item[-1])])

    for i in range(1, len(lst) + 1):
        # i принимает значения от 1 до 17, пишем со второй строки I+1, начиная с 0 индекса списка, 0 индекса подсписка I-1
        # 1,len(lst)+1 позваляет перебирать значения от 1 а не от 1, и включительно по число длины списка
        sheet_new[i + 1][0].value = lst[i - 1][0]
        sheet_new[i + 1][7].value = lst[i - 1][1]
        sheet_new[i + 1][8].value = 2

    today = datetime.date.today()
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday != 0:
        yesterday = today - datetime.timedelta(days=1)
    else:
        yesterday = today - datetime.timedelta(days=3)

    book_new.save("res/Отчет Яндекс Доставка НП СЭЙФ от %s %s.xlsx" % (yesterday, round(sm, 2)))
    book_new.close()