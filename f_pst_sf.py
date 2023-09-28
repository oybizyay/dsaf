import openpyxl
import warnings
import datetime

def f_pst_sf ():

    warnings.simplefilter("ignore")

    book = openpyxl.open("пул/fs1.xlsx")
    sheet = book.active

    book_new = openpyxl.open("ТК/Отчет 5ПОСТ НП СЭЙФ от 03.07.23.xlsx")
    sheet_new = book_new.active

    sm = 0
    #очень удобно не делать max_row+1 когда не надо брать поледнюю строку итого
    for row in range(2,sheet.max_row):
        sheet_new[row][0].value = int(sheet[row][0].value)
        sheet_new[row][7].value = sheet[row][5].value
        sm += sheet[row][5].value
        if sheet[row][10].value == "Наличный расчет":
            sheet_new[row][8].value = ""
        else:
            sheet_new[row][8].value = 2

    today = datetime.date.today()
    now = datetime.datetime.now()
    weekday = now.weekday()
    if weekday != 0:
        yesterday = today - datetime.timedelta(days=1)
    else:
        yesterday = today - datetime.timedelta(days=3)

    book_new.save("res/Отчет 5ПОСТ НП СЭЙФ от %s %s.xlsx"%(yesterday, round(sm,2)))
    book_new.close()
    book.close()