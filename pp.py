#!/usr/bin/env python
# -*- coding: utf-8 -*-


import xlrd
from xlwt import *
from pycbrf.toolbox import ExchangeRates
import datetime

excel_data_file = xlrd.open_workbook('./price500f.xls')
sheet = excel_data_file.sheet_by_index(1)
models = (
    5194800000,  # Q5942X - HP 4250
    5123050000,  # 106R02312 - Xerox 3325
    5195560000,  # 106R01412 - Xerox 3300
    5106070000,  # CE390X - HP 600
    5178090000,  # CF281X - HP 605
    5102090000,  # CE505A - HP 2035
    5120390000,  # CE400X - HP 505 color Black
    5130480000,  # CE401A - HP 505 color Blue
    5130500000,  # CE402A - HP 505 color Yellow
    5130490000,  # CE403A - HP 505 color Purple
    5194840000,  # CE255X - Canon 515
    5194870000,  # CC364A - HP 4014
    5109860000,  # 052H - Canon 421
    5220200000,  # 973X L0S07AE - HP PageWide 477 Black
    5220190000,  # 973X F6T81AE - HP PageWide 477 Blue
    5220180000,  # 973X F6T82AE - HP PageWide 477 Red
    5220170000   # 973X F6T83AE - HP PageWide 477 Yellow
)

i = 1
row_number = sheet.nrows
w = Workbook()
ws = w.add_sheet('Price')
for mod in models:  # Итерация по артикулам картриджей
    for row in range(3396, 4800):  # Итерация по строкам файла c 3396 (т.к. катриджи идут с этой строки) до строки 4800
        D_column = str(sheet.row(row)[3])[6:-1]  # Артикул
        N_column = str(sheet.row(row)[13])[7:].replace('.', ',')   # Цена Диллер-500
        H_column = str(sheet.row(row)[7])[6:-1]  # Каталог №
        F_column = str(sheet.row(row)[5])[6:-1]  # Каталог №
        if D_column == str(mod):
            print(mod)
            ws.write(i, 0, F_column)
            ws.write(i, 1, H_column)
            ws.write(i, 2, N_column)
            ws.write(i, 4, Formula("C"+str(i+1)+"*D2"))
            ws.write(i, 6, Formula("E"+str(i+1)+"*F"+str(i+1)))
            i += 1

# Exchange start
now = datetime.datetime.now()
delta = datetime.timedelta(hours=12)
tomorrow = now + delta

if now.hour < 12:
    truedate = now
else:
    truedate = tomorrow
dateforcbr = str(truedate.year) + '-' + str(truedate.month) + '-' + str(truedate.day)

rates = ExchangeRates(dateforcbr)
Ramiskurs = float(rates['USD'].rate) + (float(rates['USD'].rate) * 0.005)

ws.write(1, 3, Ramiskurs)
# Exchange rate finished

# Шапка
ws.write(0, 0, 'Наименование')
ws.write(0, 1, 'Модель')
ws.write(0, 2, 'Цена 1 картриджа в баксах')
ws.write(0, 3, 'Курс бакса + полпроцента')
ws.write(0, 4, 'Цена 1 картриджа в рублях')
ws.write(0, 5, 'Количество картриджей')
ws.write(0, 6, 'Сумма')
ws.write(i+1, 5, 'Итого')
ws.write(i+1, 6, Formula("SUM(G2:G"+str(i)+")"))

w.save('./PPP.xls')