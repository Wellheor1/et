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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, HRFlowable, Table, TableStyle
from datetime import datetime as dt

pdfmetrics.registerFont(TTFont('PTAstraSerif', 'PT-Astra-Serif_Regular.ttf'))
pdfmetrics.registerFont(TTFont('PTAstraSerifBold', 'PT-Astra-Serif_Bold.ttf'))

def create_document():
    doc = SimpleDocTemplate("tabel.pdf",
                            pagesize=landscape(A4),
                            rightMargin=30,
                            leftMargin=20,
                            topMargin=5,
                            bottomMargin=20)

    styleSheet = getSampleStyleSheet()
    style = styleSheet["Normal"]
    style.fontName = "PTAstraSerif"
    style.fontSize = 12
    style.alignment = TA_JUSTIFY

    styleCenterBold = copy.deepcopy(style)
    styleCenterBold.fontName = "PTAstraSerifBold"
    styleCenterBold.fontSize = 12
    styleCenterBold.leading = 12
    styleCenterBold.alignment = TA_CENTER

    styleCenter = copy.deepcopy(styleCenterBold)
    styleCenter.fontName = "PTAstraSerif"
    styleCenter.fontSize = 12

    styleLeft = copy.deepcopy(style)
    styleLeft.fontSize = 8
    styleLeft.alignment = TA_LEFT

    styleRight = copy.deepcopy(styleLeft)
    styleRight.alignment = TA_RIGHT

    styleCenter8 = copy.deepcopy(styleRight)
    styleCenter8.alignment = TA_CENTER

    styleCenter6 = copy.deepcopy(styleCenter8)
    styleCenter6.fontSize = 6

    styleCenter5 = copy.deepcopy(styleCenter6)
    styleCenter5.fontSize = 5


    objs = []
    department_number = 27
    start_month = 1
    end_month = 31
    month = 'Марта'
    tabel_type = 'первичный'
    hospital_name = 'ОГАУЗ ГИМДКБ'
    department = 'Кабинет неотложной травматологии и ортопедии (травмпункт)'
    date = datetime.datetime.now().strftime('%d.%m.%Y')

    objs.append(Paragraph('Утв приказом Минфина России <br/> от 30 марта 2015 г. № 52н', styleRight))
    opinion = [
        [
            Paragraph('', styleLeft),
            Paragraph(f'Табель № {department_number}', styleCenterBold),
            Paragraph('', styleRight),
            Paragraph('', styleCenter8),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph('учета использования рабочего времени', styleCenter),
            Paragraph('', styleRight),
            Paragraph('Коды', styleCenter8),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph('', styleCenter),
            Paragraph('Форма ОКУД', styleRight),
            Paragraph('0504421', styleCenter8),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph(f'За период с {start_month} по {end_month} {month}', styleCenter),
            Paragraph('Дата', styleRight),
            Paragraph(f'{date}', styleCenter8),
        ],
        [
            Paragraph('Учреждение', styleLeft),
            Paragraph(f'{hospital_name}', styleCenter),
            Paragraph('по ОКПО', styleRight),
            Paragraph('', styleCenter8),
        ],
        [
            Paragraph('Структурное подразделение', styleLeft),
            Paragraph(f'{department}', styleCenterBold),
            Paragraph('', styleRight),
            Paragraph('', styleCenter8),
        ],
        [
            Paragraph('Вид табеля', styleLeft),
            Paragraph(f'{tabel_type}', styleCenter),
            Paragraph('Номер корректировки', styleRight),
            Paragraph('', styleCenter8),
        ],
        [
            Paragraph('', styleLeft),
            Paragraph('(первичный - 0, корректирующий 1,2  и т.д)', styleCenter6),
            Paragraph('Дата формирования документа', styleRight),
            Paragraph(f'{date}', styleCenter8),
        ],
    ]
    tbl = Table(opinion, [120, 470, 130, 70])
    tbl.setStyle(
        TableStyle(
            [
                ('LINEBELOW', (1, 4), (1, 6), 0.75, colors.black),
                ('GRID', (3, 1), (3, -1), 0.75, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP')
            ]
        )
    )

    objs.append(tbl)

    title = [
            Paragraph('', styleCenter5),
            Paragraph('', styleCenter5),
            Paragraph('', styleCenter5),
        ]
    date_month_start = [Paragraph(f'{x}', styleCenter5) for x in range(1, 16)]
    summ_15 = [Paragraph('Итого дней (часов) явок (неявок) с 1-15', styleCenter5)]
    date_month_end = [Paragraph(f'{x}', styleCenter5) for x in range(16, 32)]
    summ_all = [
        Paragraph('Всего дней (часов) явок (неявок) за месяц', styleCenter5),
        Paragraph('Всего отработано часов', styleCenter5),
        Paragraph('Ноч<br/>ные', styleCenter5),
        Paragraph('Выход<br/>ные', styleCenter5),
        Paragraph('Празд<br/>ничные', styleCenter5),
        ]


    title.extend(date_month_start)
    title.extend(summ_15)
    title.extend(date_month_end)
    title.extend(summ_all)

    num = [Paragraph(f'{x}', styleCenter5) for x in range(1, 41)]

    employees = [
        {"person": "Прет3333ориус", "tabel_number": "885", "post": "Заведующий"},
        {"person": "Прет333fdfориус", "tabel_number": "885", "post": "Заведующий"},
        {"person": "Пре223dsториус", "tabel_number": "883335", "post": "Заведующий"},
    ]


    s = []
    for employ in employees:
        a = []
        a.append(employ["person"])
        a.append(employ["tabel_number"])
        a.append(employ["post"])
        s.append(a)

    opinion = [
        [
            Paragraph('Фамилия, имя, отчество', styleCenter5),
            Paragraph('Учетный номер', styleCenter5),
            Paragraph('Должность (профессия)', styleCenter5),
            Paragraph('Числа месяца', styleCenter5)
        ],
        title,
        num,

    ]
    tbl = Table(opinion)
    tbl.setStyle(
        TableStyle(
            [
                ('GRID', (0, 0), (-1, -1), 0.75, colors.black),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('SPAN', (0, 0), (0, 1)),
                ('SPAN', (1, 0), (1, 1)),
                ('SPAN', (2, 0), (2, 1)),
                ('SPAN', (3, 0), (-1, 0)),

            ]
        )
    )

    objs.append(tbl)
    doc.build(objs)


if __name__ == '__main__':
    create_document()
