import copy
import datetime

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Table, TableStyle
from datetime import datetime as dt

pdfmetrics.registerFont(TTFont('PTAstraSerif', 'PT-Astra-Serif_Regular.ttf'))
pdfmetrics.registerFont(TTFont('PTAstraSerifBold', 'PT-Astra-Serif_Bold.ttf'))

def create_document():
    doc = SimpleDocTemplate("tabel.pdf",
                            pagesize=landscape(A4),
                            rightMargin=10 * mm,
                            leftMargin=5 * mm,
                            topMargin=2 * mm,
                            bottomMargin=40 * mm)

    styleSheet = getSampleStyleSheet()
    style = styleSheet["Normal"]
    style.fontName = "PTAstraSerif"
    style.fontSize = 10
    style.alignment = TA_JUSTIFY

    styleCenterBold = copy.deepcopy(style)
    styleCenterBold.fontName = "PTAstraSerifBold"
    styleCenterBold.fontSize = 10
    styleCenterBold.alignment = TA_CENTER

    styleCenter = copy.deepcopy(styleCenterBold)
    styleCenter.fontName = "PTAstraSerif"

    styleLeft = copy.deepcopy(style)
    styleLeft.fontSize = 6
    styleLeft.alignment = TA_LEFT

    styleRight = copy.deepcopy(styleLeft)
    styleRight.alignment = TA_RIGHT

    styleRightBold = copy.deepcopy(styleRight)
    styleRightBold.fontName = "PTAstraSerifBold"

    styleCenterTitle = copy.deepcopy(styleRight)
    styleCenterTitle.alignment = TA_CENTER

    styleCenterSup = copy.deepcopy(styleCenterTitle)
    styleCenterSup.fontSize = 5

    styleCenterData = copy.deepcopy(styleCenterSup)
    styleCenterData.fontSize = 10

    styleCenterDataBold = copy.deepcopy(styleCenterSup)
    styleCenterDataBold.fontSize = 10
    styleCenterDataBold.fontName = "PTAstraSerifBold"

    styleCenterDataTitle = copy.deepcopy(styleCenterSup)
    styleCenterDataTitle.fontSize = 8

    objs = []
    department_number = 27
    start_month = 1
    end_month = 31
    month = 'Марта'
    tabel_type = 'первичный'
    hospital_name = 'ОГАУЗ ГИМДКБ'
    department = 'Кабинет неотложной травматологии и ортопедии (травмпункт)'
    date = datetime.datetime.now().strftime('%d.%m.%Y')
    main_doctor = 'Новожилов В.А.'
    head_department = 'Преториус Т.Л.'
    old_sestra = 'Тотьямина Д.С.'
    hr_specialist = 'Краснова С.А.'




    objs.append(Paragraph('Утв приказом Минфина России <br/> от 30 марта 2015 г. № 52н', styleRightBold))
    opinion = [
        [
            Paragraph('', styleLeft),
            Paragraph(f'Табель № {department_number}', styleCenterBold),
            Paragraph('', styleRight),
            Paragraph('', styleCenterTitle),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph('учета использования рабочего времени', styleCenter),
            Paragraph('', styleRight),
            Paragraph('Коды', styleCenterTitle),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph('', styleCenter),
            Paragraph('Форма ОКУД', styleRight),
            Paragraph('0504421', styleCenterTitle),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph(f'За период с {start_month} по {end_month} {month}', styleCenter),
            Paragraph('Дата', styleRight),
            Paragraph(f'{date}', styleCenterTitle),
        ],
        [
            Paragraph('Учреждение', styleLeft),
            Paragraph(f'{hospital_name}', styleCenter),
            Paragraph('по ОКПО', styleRight),
            Paragraph('', styleCenterTitle),
        ],
        [
            Paragraph('Структурное подразделение', styleLeft),
            Paragraph(f'{department}', styleCenterBold),
            Paragraph('', styleRight),
            Paragraph('', styleCenterTitle),
        ],
        [
            Paragraph('Вид табеля', styleLeft),
            Paragraph(f'{tabel_type}', styleCenter),
            Paragraph('Номер корректировки', styleRight),
            Paragraph('', styleCenterTitle),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph('(первичный - 0, корректирующий 1,2  и т.д)', styleCenterSup),
            Paragraph('Дата формирования документа', styleRight),
            Paragraph(f'{date}', styleCenterTitle),
        ],
    ]
    tbl = Table(opinion, [30 * mm, 180 * mm, 40 * mm, 25 * mm])
    tbl.setStyle(
        TableStyle(
            [
                ('LINEBELOW', (1, 4), (1, 6), 0.75, colors.black),
                ('GRID', (3, 1), (3, -1), 0.75, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 1),
                ('BOTTOMTPADDING', (0, 0), (1, -1), -1),
                ('TOPPADDING', (0, 0), (1, -1), -1),
            ]
        )
    )

    objs.append(tbl)

    objs.append(Spacer(1, 12))
    title = [
            Paragraph('', styleCenterDataTitle),
            Paragraph('', styleCenterDataTitle),
            Paragraph('', styleCenterDataTitle),
        ]
    date_month_start = [Paragraph(f'{x}', styleCenterDataTitle) for x in range(1, 16)]
    summ_15 = [Paragraph('Итого дней (часов) явок (неявок) с 1-15', styleCenterDataTitle)]
    date_month_end = [Paragraph(f'{x}', styleCenterDataTitle) for x in range(16, 32)]
    summ_all = [
        Paragraph('Всего дней (часов) явок (неявок) за месяц', styleCenterDataTitle),
        Paragraph('Всего отработано часов', styleCenterDataTitle),
        Paragraph('Ноч<br/>ные', styleCenterDataTitle),
        Paragraph('Выход<br/>ные', styleCenterDataTitle),
        Paragraph('Празд<br/>ничные', styleCenterDataTitle),
        ]


    title.extend(date_month_start)
    title.extend(summ_15)
    title.extend(date_month_end)
    title.extend(summ_all)

    num = [Paragraph(f'{x}', styleCenterDataTitle) for x in range(1, 41)]

    employees = [
        {"person": "Прет111ориус Т.Л.", "tabel_number": "885", "post": "Заведующий кабинетом врач травматолог-ортопед",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет2222ориус Т.Л.", "tabel_number": "8835", "post": "врач травматолог-ортопед внутр. совместитель",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет3333ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет4444ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет5555ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},
        {"person": "Прет6666ориус Т.Л.", "tabel_number": "8853", "post": "Заведующий",
         "hours15": [f'{x}' for x in range(1, 16)], "summ_15": '324', "hours30": [f'{x}' for x in range(16, 32)],
         "common_days": "31", "common_hours": "555", "night_hours": '33', "weekends_hours": "30",
         "holidays_hours": "10"},

    ]
    employees_data = []
    for employ in employees:
        employee_data = []
        hours15 = []
        hours30 = []
        employee_data.append(Paragraph(employ["person"], styleCenterData))
        employee_data.append(Paragraph(employ["tabel_number"], styleCenterDataBold))
        employee_data.append(Paragraph(employ["post"], styleCenterDataTitle))
        hours15 = [Paragraph(f'{x}', styleCenterData) for x in employ["hours15"]]
        employee_data.extend(hours15)
        employee_data.append(Paragraph(employ["summ_15"], styleCenterData))
        hours30 = [Paragraph(f'{x}', styleCenterData) for x in employ["hours30"]]
        employee_data.extend(hours30)
        employee_data.append(Paragraph(employ["common_days"], styleCenterData))
        employee_data.append(Paragraph(employ["common_hours"], styleCenterData))
        employee_data.append(Paragraph(employ["night_hours"], styleCenterData))
        employee_data.append(Paragraph(employ["weekends_hours"], styleCenterData))
        employee_data.append(Paragraph(employ["holidays_hours"], styleCenterData))
        employees_data.append(employee_data)

    opinion = [
        [
            Paragraph('Фамилия, имя, отчество', styleCenterData),
            Paragraph('Учетный номер', styleCenterData),
            Paragraph('Должность (профессия)', styleCenterData),
            Paragraph('Числа месяца', styleCenterData)
        ],
        title,
        num,
    ]
    opinion.extend(employees_data)

    tbl = Table(opinion)
    tbl.setStyle(
        TableStyle(
            [
                ('GRID', (0, 0), (-1, -1), 0.75, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('SPAN', (0, 0), (0, 1)),
                ('SPAN', (1, 0), (1, 1)),
                ('SPAN', (2, 0), (2, 1)),
                ('SPAN', (3, 0), (-1, 0)),
                ('LEFTPADDING', (3, 0), (-1, -1), 0.2),
                ('RIGHTPADDING', (3, 0), (-1, -1), 0.2),
                ('BOTTOMTPADDING', (3, 0), (-1, -1), 2),
                ('TOPPADDING', (3, 0), (-1, -1), 2),
            ]
        )
    )
    objs.append(tbl)
    objs.append(Spacer(1, 50))
    objs.append(PageBreak())

    def first_pages(canvas, doc):
        canvas.saveState()
        canvas.setFont('PTAstraSerif', 8)

        canvas.drawString(11 * mm, 41 * mm, "Главный врач")
        canvas.drawString(75 * mm, 41 * mm, f"{main_doctor}")

        canvas.line(7 * mm, 40 * mm, 33 * mm, 40 * mm)
        canvas.line(45 * mm, 40 * mm, 101 * mm, 40 * mm)

        canvas.drawString(12 * mm, 37 * mm, "(должность)")
        canvas.drawString(50 * mm, 37 * mm, "(подпись)")

        canvas.drawString(10 * mm, 31 * mm, "Зав. отделением")
        canvas.drawString(75 * mm, 31 * mm, f"{head_department}")

        canvas.line(7 * mm, 30 * mm, 33 * mm, 30 * mm)
        canvas.line(45 * mm, 30 * mm, 101 * mm, 30 * mm)

        canvas.drawString(12 * mm, 27 * mm, "(должность)")
        canvas.drawString(50 * mm, 27 * mm, "(подпись)")

        canvas.drawString(8 * mm, 21 * mm, "Старшая медсестра")
        canvas.drawString(75 * mm, 21 * mm, f"{old_sestra}")

        canvas.line(7 * mm, 20 * mm, 33 * mm, 20 * mm)
        canvas.line(45 * mm, 20 * mm, 101 * mm, 20 * mm)

        canvas.drawString(12 * mm, 17 * mm, "(должность)")
        canvas.drawString(50 * mm, 17 * mm, "(подпись)")

        canvas.drawString(10 * mm, 11 * mm, "Специалист о.к.")
        canvas.drawString(75 * mm, 11 * mm, f"{hr_specialist}")

        canvas.line(7 * mm, 10 * mm, 33 * mm, 10 * mm)
        canvas.line(45 * mm, 10 * mm, 101 * mm, 10 * mm)

        canvas.drawString(12 * mm, 7 * mm, "(должность)")
        canvas.drawString(50 * mm, 7 * mm, "(подпись)")


        canvas.rect(156 * mm, 10 * mm, 112 * mm, 23 * mm, stroke=1, fill=0)
        canvas.setFont('PTAstraSerifBold', 8)
        canvas.drawString(185 * mm, 30 * mm, "Отметка бухгалтерии о принятии настоящего табеля")
        canvas.drawString(160 * mm, 24 * mm, "Исполнитель")
        canvas.setFont('PTAstraSerif', 8)
        canvas.line(180 * mm, 23 * mm, 255 * mm, 23 * mm)
        canvas.drawString(210 * mm, 20 * mm, "(подпись)")

        canvas.line(165 * mm, 13 * mm, 175 * mm, 13 * mm)
        canvas.line(180 * mm, 13 * mm, 210 * mm, 13 * mm)
        canvas.line(215 * mm, 13 * mm, 225 * mm, 13 * mm)

        canvas.restoreState()

    def later_pages(canvas, doc):
        canvas.saveState()
        canvas.setFont('PTAstraSerif', 8)

        canvas.drawString(11 * mm, 41 * mm, "Главный врач")
        canvas.drawString(75 * mm, 41 * mm, f"{main_doctor}")

        canvas.line(7 * mm, 40 * mm, 33 * mm, 40 * mm)
        canvas.line(45 * mm, 40 * mm, 101 * mm, 40 * mm)

        canvas.drawString(12 * mm, 37 * mm, "(должность)")
        canvas.drawString(50 * mm, 37 * mm, "(подпись)")

        canvas.drawString(10 * mm, 31 * mm, "Зав. отделением")
        canvas.drawString(75 * mm, 31 * mm, f"{head_department}")

        canvas.line(7 * mm, 30 * mm, 33 * mm, 30 * mm)
        canvas.line(45 * mm, 30 * mm, 101 * mm, 30 * mm)

        canvas.drawString(12 * mm, 27 * mm, "(должность)")
        canvas.drawString(50 * mm, 27 * mm, "(подпись)")

        canvas.drawString(8 * mm, 21 * mm, "Старшая медсестра")
        canvas.drawString(75 * mm, 21 * mm, f"{old_sestra}")

        canvas.line(7 * mm, 20 * mm, 33 * mm, 20 * mm)
        canvas.line(45 * mm, 20 * mm, 101 * mm, 20 * mm)

        canvas.drawString(12 * mm, 17 * mm, "(должность)")
        canvas.drawString(50 * mm, 17 * mm, "(подпись)")

        canvas.drawString(10 * mm, 11 * mm, "Специалист о.к.")
        canvas.drawString(75 * mm, 11 * mm, f"{hr_specialist}")

        canvas.line(7 * mm, 10 * mm, 33 * mm, 10 * mm)
        canvas.line(45 * mm, 10 * mm, 101 * mm, 10 * mm)

        canvas.drawString(12 * mm, 7 * mm, "(должность)")
        canvas.drawString(50 * mm, 7 * mm, "(подпись)")

        canvas.rect(156 * mm, 10 * mm, 112 * mm, 23 * mm, stroke=1, fill=0)
        canvas.setFont('PTAstraSerifBold', 8)
        canvas.drawString(185 * mm, 30 * mm, "Отметка бухгалтерии о принятии настоящего табеля")
        canvas.drawString(160 * mm, 24 * mm, "Исполнитель")
        canvas.setFont('PTAstraSerif', 8)
        canvas.line(180 * mm, 23 * mm, 255 * mm, 23 * mm)
        canvas.drawString(210 * mm, 20 * mm, "(подпись)")

        canvas.line(165 * mm, 13 * mm, 175 * mm, 13 * mm)
        canvas.line(180 * mm, 13 * mm, 210 * mm, 13 * mm)
        canvas.line(215 * mm, 13 * mm, 225 * mm, 13 * mm)

        canvas.restoreState()

    doc.build(objs, onFirstPage=first_pages, onLaterPages=later_pages)


if __name__ == '__main__':
    create_document()
