import calendar
import datetime

import pytils as pytils
from openpyxl import Workbook
from openpyxl.styles import Alignment, NamedStyle, Font, Border, Side
from openpyxl.worksheet.page import PageMargins

department_table_number = 27
month_name = pytils.dt.ru_strftime(u"%B", inflected=True, date=datetime.datetime.now())
current_year = datetime.date.today().year
current_month = datetime.date.today().month
first_day_month = 1
last_day_month = calendar.monthrange(current_year, current_month)[1]
tabel_type = 'первичный'
hospital_name = 'ОГАУЗ ГИМДКБ'
department_name = 'Кабинет неотложной травматологии и ортопедии (травмпункт)'
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
active_list.page_margins = PageMargins(top=0.2, bottom=0.2, left=0.2, right=0.2)
active_list.title = 'Табель 52Н'

styleCenterTitle = NamedStyle(name='styleCenterTitle')
styleCenterTitle.alignment = Alignment(horizontal='center', vertical='center')

styleCenterSup = NamedStyle(name='styleCenterSup')
styleCenterSup.font = Font(size=6)
styleCenterSup.alignment = Alignment(horizontal='center', vertical='top')

styleCenterBorder = NamedStyle(name='styleCenterTitleBorder')
bd = Side(style='thin', color='000000')
styleCenterBorder.border = Border(left=bd, right=bd, top=bd, bottom=bd)
styleCenterBorder.font = Font(size=8)
styleCenterBorder.alignment = Alignment(horizontal='center', vertical='center')

styleCenterBorderBottom = NamedStyle(name='styleCenterBorderBottom')
bd = Side(style='thin', color='000000')
styleCenterBorderBottom.border = Border(bottom=bd)
styleCenterBorderBottom.alignment = Alignment(horizontal='center', vertical='center')

styleCenterTitleBold = NamedStyle(name='styleCenterTitleBold')
styleCenterTitleBold.alignment = Alignment(horizontal='center', vertical='center')
styleCenterTitleBold.font = Font(bold=True, color='000000')

styleRightTitle = NamedStyle(name='styleRightTitle')
styleRightTitle.font = Font(size=8)
styleRightTitle.alignment = Alignment(horizontal='right', vertical='center')

styleLeftTitle = NamedStyle(name='styleLeftTitle')
styleLeftTitle.font = Font(size=8)
styleLeftTitle.alignment = Alignment(horizontal='left', vertical='center')

styleRightTitleBold = NamedStyle(name='styleRightTitleBold')
styleRightTitleBold.alignment = Alignment(horizontal='right', vertical='center')
styleRightTitleBold.font = Font(bold=True, color='000000', size=8)

active_list.merge_cells(start_row=1, start_column=35, end_row=1, end_column=38)
active_list['AI1'].value = 'Утв приказом Минфина России'
active_list['AI1'].style = styleRightTitleBold

active_list.merge_cells(start_row=2, start_column=35, end_row=2, end_column=38)
active_list['AI2'].value = 'от 30 марта 2015г № 52н'
active_list['AI2'].style = styleRightTitleBold

active_list.merge_cells(start_row=3, start_column=4, end_row=3, end_column=30)
active_list['D3'].value = f'Табель № {department_table_number}'
active_list['D3'].style = styleCenterTitleBold

active_list.merge_cells(start_row=4, start_column=4, end_row=4, end_column=30)
active_list['D4'].value = 'учета использования рабочего времени'
active_list['D4'].style = styleCenterTitle
active_list.merge_cells(start_row=4, start_column=37, end_row=4, end_column=38)
active_list['AK4'].value = 'коды'

active_list['AJ5'].value = 'форма ОКУД'
active_list['AJ5'].style = styleRightTitle
active_list.merge_cells(start_row=5, start_column=37, end_row=5, end_column=38)
active_list['AK5'].value = f'{form_okyd}'

active_list.merge_cells(start_row=6, start_column=4, end_row=6, end_column=30)
active_list['D6'].value = f'За период с {first_day_month} по {last_day_month} {month_name}'
active_list['D6'].style = styleCenterTitle
active_list['AJ6'].value = 'Дата'
active_list['AJ6'].style = styleRightTitle
active_list.merge_cells(start_row=6, start_column=37, end_row=6, end_column=38)
active_list['AK6'].value = f'{date_now}'

active_list['A7'].value = 'Учреждение'
active_list['A7'].style = styleLeftTitle
active_list.merge_cells(start_row=7, start_column=4, end_row=7, end_column=30)
active_list['D7'].value = f'{hospital_name}'
# active_list.merge_cells(start_row=7, start_column=35, end_row=7, end_column=36)
active_list['AJ7'].value = 'по ОКПО'
active_list['AJ7'].style = styleRightTitle
active_list.merge_cells(start_row=7, start_column=37, end_row=7, end_column=38)

active_list.merge_cells(start_row=8, start_column=1, end_row=8, end_column=3)
active_list['A8'].value = 'Структурное подразделение'
active_list['A8'].style = styleLeftTitle
active_list.merge_cells(start_row=8, start_column=4, end_row=8, end_column=30)
active_list['D8'].value = f'{department_name}'
active_list.merge_cells(start_row=8, start_column=37, end_row=8, end_column=38)

active_list['A9'].value = 'Вид табеля'
active_list['A9'].style = styleLeftTitle
active_list.merge_cells(start_row=9, start_column=4, end_row=9, end_column=30)
active_list['D9'].value = f'{tabel_type}'
active_list.merge_cells(start_row=9, start_column=35, end_row=9, end_column=36)
active_list['AI9'].value = 'Номер корректировки'
active_list['AI9'].style = styleRightTitle
active_list.merge_cells(start_row=9, start_column=37, end_row=9, end_column=38)

active_list.merge_cells(start_row=10, start_column=4, end_row=10, end_column=30)
active_list['D10'].value = '(первичный - 0; корректирующий 1,2 и т.д.)'
active_list.merge_cells(start_row=10, start_column=34, end_row=10, end_column=36)
active_list['AH10'].value = 'Дата формирования документа'
active_list['AH10'].style = styleRightTitle
active_list.merge_cells(start_row=10, start_column=37, end_row=10, end_column=38)
active_list['AK10'].value = f'{date_now}'

for column in active_list['AK4':'AL10']:
    for cell in column:
        cell.style = styleCenterBorder

for column in active_list['D7':'AD9']:
    for cell in column:
        cell.style = styleCenterBorderBottom

for column in active_list['D10':'AD10']:
    for cell in column:
        cell.style = styleCenterSup


book.save('tabel.xlsx')