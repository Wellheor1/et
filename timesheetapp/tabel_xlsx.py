import datetime
import calendar

import pytils as pytils
from openpyxl import Workbook


department_table_number = 27
month_name = pytils.dt.ru_strftime(u"%B", inflected=True, date=datetime.datetime.now())
current_year = datetime.date.today().year
current_month = datetime.date.today().month
first_day_month = 1
last_day_month = calendar.monthrange(current_year, current_month)[1]
tabel_type = 'первичный'
hospital_name = 'ОГАУЗ ГИМДКБ'
department_name = 'травмпункт'
date_now = datetime.datetime.now().strftime('%d.%m.%Y')
main_doctor = 'Новожилов В.А.'
head_department = 'Преториус Т.Л.'
old_sestra = 'Тотьямина Д.С.'
hr_specialist = 'Краснова С.А.'

book = Workbook()
active_list = book.active
active_list["N1"] = 'Утв. приказом минфина России'
active_list["N2"] = 'от 30 марта 2015 г. № 52н'

active_list.title = f'{department_name}'

d = active_list.cell(row=4, column=2, value=10)
active_list['A1':'C2'] = 10


book.save('tabel.xlsx')

