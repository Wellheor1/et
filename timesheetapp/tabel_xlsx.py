import datetime
import calendar
import copy

import pytils as pytils

from openpyxl import Workbook
from openpyxl.styles import Alignment, NamedStyle, Font, Border, Side

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
form_okyd = '0504421'

book = Workbook()
active_list = book.active
active_list.page_setup.orientation = active_list.ORIENTATION_LANDSCAPE
active_list.page_setup.paperSize = active_list.PAPERSIZE_A4
active_list.page_setup.fitToPage = 1
active_list.title = f'{department_name}'

styleCenterTitle = NamedStyle(name='styleCenterTitle')
styleCenterTitle.alignment = Alignment(horizontal='center', vertical='center')

styleCenterBorder = NamedStyle(name='styleCenterTitleBorder')
bd = Side(style='thin', color='000000')
styleCenterBorder.border = Border(left=bd, right=bd, top=bd, bottom=bd)
styleCenterBorder.alignment = Alignment(horizontal='center', vertical='center')

styleCenterBorderBottom = NamedStyle(name='styleCenterBorderBottom')
bd = Side(style='thin', color='000000')
styleCenterBorderBottom.border = Border(bottom=bd)
styleCenterBorderBottom.alignment = Alignment(horizontal='center', vertical='center')

styleCenterTitleBold = NamedStyle(name='styleCenterTitleBold')
styleCenterTitleBold.alignment = Alignment(horizontal='center', vertical='center')
styleCenterTitleBold.font = Font(bold=True, color='000000')

styleRightTitle = NamedStyle(name='styleRightTitle')
styleRightTitle.alignment = Alignment(horizontal='right', vertical='center')

styleLeftTitle = NamedStyle(name='styleLeftTitle')
styleLeftTitle.alignment = Alignment(horizontal='left', vertical='center')

styleRightTitleBold = NamedStyle(name='styleRightTitleBold')
styleRightTitleBold.alignment = Alignment(horizontal='right', vertical='center')
styleRightTitleBold.font = Font(bold=True, color='000000')

active_list.merge_cells(start_row=1, start_column=35, end_row=1, end_column=38)
active_list.merge_cells(start_row=2, start_column=35, end_row=2, end_column=38)
d = active_list['AI1']
d.value = 'Утв приказом Минфина России'
d.style = styleRightTitleBold
d = active_list['AI2']
d.value = 'от 30 марта 2015г № 52н'
d.style = styleRightTitleBold

active_list.merge_cells(start_row=3, start_column=14, end_row=3, end_column=16)
d = active_list['N3']
d.value = f'Табель № {department_table_number}'
d.style = styleCenterTitleBold

active_list.merge_cells(start_row=4, start_column=13, end_row=4, end_column=17)
d = active_list['M4']
d.value = 'учета использования рабочего времени'
d.style = styleCenterTitle
active_list.merge_cells(start_row=4, start_column=37, end_row=4, end_column=38)
d = active_list['AK4']
d.value = 'коды'
d.style = styleCenterBorder
d = active_list['AL4']
d.style = styleCenterBorder

active_list.merge_cells(start_row=5, start_column=35, end_row=5, end_column=36)
d = active_list['AI5']
d.value = 'форма ОКУД'
d.style = styleRightTitle
active_list.merge_cells(start_row=5, start_column=37, end_row=5, end_column=38)
d = active_list['AK5']
d.value = f'{form_okyd}'
d.style = styleCenterBorder
d = active_list['AL5']
d.style = styleCenterBorder

active_list.merge_cells(start_row=6, start_column=14, end_row=6, end_column=16)
d = active_list['N6']
d.value = f'За период с {first_day_month} по {last_day_month} {month_name}'
d.style = styleCenterTitle
d = active_list['AJ6']
d.value = 'Дата'
d.style = styleRightTitle
active_list.merge_cells(start_row=6, start_column=37, end_row=6, end_column=38)
d = active_list['AK6']
d.value = f'{date_now}'
d.style = styleCenterBorder
d = active_list['AL6']
d.style = styleCenterBorder

d = active_list['A7']
d.value = 'Учреждение'
d.style = styleLeftTitle
active_list.merge_cells(start_row=7, start_column=14, end_row=7, end_column=16)
d = active_list['N7']
d.value = f'{hospital_name}'
d.style = styleCenterBorderBottom
active_list.merge_cells(start_row=7, start_column=35, end_row=7, end_column=36)
d = active_list['AI7']
d.value = 'по ОКПО'
d.style = styleRightTitle
active_list.merge_cells(start_row=7, start_column=37, end_row=7, end_column=38)
d = active_list['AK7']
d.style = styleCenterBorder
d = active_list['AL7']
d.style = styleCenterBorder
thin = Border(bottom=)
for row in active_list.iter_rows(min_row=7, min_col=3, max_row=7, max_col=30):
    for cell in row:
        cell.border = Border(bottom='thin')


book.save('tabel.xlsx')
