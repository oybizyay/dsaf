from f_pst_sf import f_pst_sf
from f_pst_sr import f_pst_sr1, f_pst_sr2
from log_sr import log_sr
from log_safe import log_safe
from cdek_safe import cdek_safe
from cdek import cdek
from yand import yand
from yand_safe import yand_safe
from post import post
from post_safe import post_safe
from box import box
from box_safe import box_safe
from ekvr import ekvr
from dpd import dpd
from dpd_safe import dpd_safe


import os


def reestr():

    resul = []

    if os.path.exists("пул/fs1.xlsx"):
        f_pst_sf()


    if os.path.exists("пул/f1.xlsx"):
        if os.path.exists("пул/f2.xlsx"):
            f_pst_sr2()
        else:
            f_pst_sr1()


    if os.path.exists("пул/o1.xlsx"):
        log_sr_r = log_sr()
        resul.append(log_sr_r)



    if os.path.exists("пул/l1.xlsx"):
        log_safe_r = log_safe()
        resul.append(log_safe_r)


    if os.path.exists("пул/s1.xls"):
        cdek_r = cdek()
        resul.append(cdek_r)



    if os.path.exists("пул/ss1.xls"):
        cdek_safe()


    if os.path.exists("пул/y1.csv"):
        yand()


    if os.path.exists("пул/ys1.csv"):
        yand_safe()



    if os.path.exists("пул/p1.DBF") or os.path.exists("пул/p2.xlsx"):
        post()


    if os.path.exists("пул/ps1.DBF") or os.path.exists("пул/ps2.xlsx"):
        post_safe()

    if os.path.exists("пул/b1.xlsx"):
        box()


    if os.path.exists("пул/bs1.xlsx"):
        box_safe()


    if os.path.exists("пул/e1.csv"):
        ekvr()

    if os.path.exists("пул/d1.xlsx"):
        dpd()

    if os.path.exists("пул/ds1.xlsx"):
        dpd_safe()


    resul.append("*")

    path = "res/"
    for filename in os.listdir(path):
        # Если файл существует и является файлом
        if os.path.isfile(os.path.join(path, filename)):
            resul.append(filename)

    return resul
