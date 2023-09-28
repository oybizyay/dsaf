import csv
import openpyxl
import warnings
import datetime
import os


def post_safe():
    warnings.simplefilter("ignore")

    book_new = openpyxl.open("ТК/Отчет Почта  НП СЭЙФ от 03.07.23.xlsx")
    sheet_new = book_new.active

    lst_np = []
    sm = 0
    my_list = []

    if os.path.exists("пул/ps1.DBF"):
        with open("пул/ps1.DBF", "r", encoding="CP866") as file:
            reader = csv.reader(file, delimiter=" ", quotechar='"', quoting=csv.QUOTE_NONE)
            for row in reader:
                lst = row

        # row это уже список , теперь нужно удалить из него пустоты ""
        for element in lst:
            if element != "":
                my_list.append(element)

        # объявляем в переменную значения из списка с RPO и 2е значение после него это всегда сумма
        for i in range(len(my_list)):
            if "RPO" in my_list[i]:
                a = my_list[i]
                a = a.split("=")
                a = int(a[1])
                b = float(my_list[i + 2])
                sm += b
                c = ""
                lst_np.append([a, b, c])

    # cod post
    if os.path.exists("пул/ps2.xlsx"):
        book = openpyxl.open("пул/ps2.xlsx")
        sheet = book.active
        for row in range(7, sheet.max_row - 3):
            a = int(sheet[row][1].value)
            b = sheet[row][3].value
            b = b.split(",")
            b = b[0] + "." + b[1]
            b = b.split(" ")
            b = b[0] + b[1]
            b = float(b)
            sm += b
            c = 2
            lst_np.append([a, b, c])

    # запись

    for i in range(1, len(lst_np) + 1):
        # i принимает значения от 1 до 17, пишем со второй строки I+1, начиная с 0 индекса списка, 0 индекса подсписка I-1
        # 1,len(lst)+1 позваляет перебирать значения от 1 а не от 1, и включительно по число длины списка

        sheet_new[i + 1][1].value = lst_np[i - 1][0]
        sheet_new[i + 1][7].value = lst_np[i - 1][1]
        sheet_new[i + 1][8].value = lst_np[i - 1][2]

    today = datetime.date.today()
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday != 0:
        yesterday = today - datetime.timedelta(days=1)
    else:
        yesterday = today - datetime.timedelta(days=3)

    book_new.save("res/Отчет Почта  НП СЭЙФ от %s %s.xlsx" % (yesterday, round(sm, 2)))
    book_new.close()