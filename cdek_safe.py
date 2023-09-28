import openpyxl
import warnings
import datetime
import os
import xlrd

def cdek_safe ():
    warnings.simplefilter("ignore")

    workbook = xlrd.open_workbook("пул/ss1.xls", on_demand = True)
    worksheet = workbook.sheet_by_index(0)

    workbook_o = xlrd.open_workbook("пул/ss11.xls", on_demand = True)
    worksheet_o = workbook_o.sheet_by_index(0)

    book_new = openpyxl.open("ТК/Отчет СДЭК НП СЭЙФ от 03.07.2023.xlsx")
    sheet_new = book_new.active

    if os.path.exists("пул/ss2.xls"):

        workbook2 = xlrd.open_workbook("пул/ss2.xls", on_demand = True)
        worksheet2 = workbook2.sheet_by_index(0)

        workbook_o2 = xlrd.open_workbook("пул/ss22.xls", on_demand = True)
        worksheet_o2 = workbook_o2.sheet_by_index(0)


    #список комиссии онлайн отчеты
    list_spo = []

    for i in range(1,worksheet_o.nrows-1):

        a = int(worksheet_o.cell_value(i,0))
        b = worksheet_o.cell_value(i, 1)
        if b != 0.0:
            ab = [a,b]
            list_spo.append(ab)

    if os.path.exists("пул/ss2.xls"):

        for i in range(1, worksheet_o2.nrows - 1):
            a = int(worksheet_o2.cell_value(i, 0))
            b = worksheet_o2.cell_value(i, 1)
            if b != 0.0:
                ab = [a, b]
                list_spo.append(ab)

    # print(list_spo)

    #список заказов
    list_np = []
    sm = 0

    for i in range(1,worksheet.nrows-2):
        if worksheet.cell_value(i,11) != 0:
            a = int(worksheet.cell_value(i,0))
            b = int(worksheet.cell_value(i,1))
            c = worksheet.cell_value(i,11)
            sm += c
            for i in list_spo:
                if a in i:
                    d = 2
                    break
                else:
                    d = ""
            abcd = [b,a,c,d]
            list_np.append(abcd)

    if os.path.exists("пул/ss2.xls"):
        for i in range(1, worksheet2.nrows - 2):
            if worksheet2.cell_value(i,11) != 0:
                a = int(worksheet2.cell_value(i,0))
                b = int(worksheet2.cell_value(i,1))
                c = worksheet2.cell_value(i,11)
                sm += c
                for i in list_spo:
                    if a in i:
                        d = 2
                        break
                    else:
                        d = ""
                abcd = [b,a,c,d]
                list_np.append(abcd)

    #запись
    for i in range(1, len(list_np) + 1):
        # i принимает значения от 1 до 17, пишем со второй строки I+1, начиная с 0 индекса списка, 0 индекса подсписка I-1
        # 1,len(lst)+1 позваляет перебирать значения от 1 а не от 0, и включительно по число длины списка
        sheet_new[i + 1][0].value = list_np[i - 1][0]
        sheet_new[i + 1][1].value = list_np[i - 1][1]
        sheet_new[i + 1][7].value = list_np[i - 1][2]
        sheet_new[i + 1][8].value = list_np[i - 1][3]

    today = datetime.date.today()
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday != 0:
        yesterday = today - datetime.timedelta(days=1)
    else:
        yesterday = today - datetime.timedelta(days=3)

    book_new.save("res/Отчет СДЭК НП СЭЙФ от %s %s.xlsx" % (yesterday, round(sm, 2)))
    book_new.close()

    workbook.release_resources()
    workbook_o.release_resources()

    if os.path.exists("пул/ss2.xls"):
        workbook2.release_resources()

    if os.path.exists("пул/ss22.xls"):
        workbook_o2.release_resources()