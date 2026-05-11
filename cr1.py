from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os
from num2words import num2words


def sum_to_words(amount):
    rub = int(amount)
    kop = int(round((amount - rub) * 100))

    # Переводим рубли (учитываем мужской род)
    rub_text = num2words(rub, lang='ru')
    # Переводим копейки (учитываем женский род: "одна", "две")
    kop_text = num2words(kop, lang='ru')

    # Если нужно просто добавить слова "руб." и "коп." к числам:
    return f"{rub_text} руб. {kop_text} коп."


def add_formatted_paragraph(doc, text, bold=False, size=11, alignment=None, font_name='Times New Roman'):
    """
    Добавление форматированного параграфа
    """
    para = doc.add_paragraph()
    if alignment is not None:
        para.alignment = alignment
    run = para.add_run(text)
    run.font.name = font_name
    run.font.size = Pt(size)
    run.bold = bold
    return para


def create_application_doc(num, date, table1_rows, table1_data, price1, date_work):
    """Создание документа заявки"""
    price1 = float(price1)
    doc = Document()

    # Настройка стилей документа
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Times New Roman'
    font.size = Pt(11)

    # Настройка полей страницы
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

    # Заголовок
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(f'Заявка на работы №{num}')
    title_run.bold = True
    title_run.font.size = Pt(11)
    title_run.font.name = 'Times New Roman'

    # Местоположение и дата
    header_para = doc.add_paragraph()
    header_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run1 = header_para.add_run('г. Новосибирск')
    run1.font.size = Pt(11)
    run1.font.name = 'Times New Roman'

    # Добавляем табуляцию для даты справа
    run2 = header_para.add_run(f'\t\t\t\t\t\t\t\t\t{date}')
    run2.font.size = Pt(11)

    run3 = header_para.add_run('\n\t1. Список объектов:')
    run3.font.size = Pt(12)
    run3.font.name = 'Times New Roman'

    # Создание таблицы объектов
    table1 = doc.add_table(rows=table1_rows, cols=5)
    table1.style = 'Table Grid'
    table1.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Заголовки таблицы
    headers = ['№п/п', 'кол-во ПУ', 'тип ПУ', 'Адрес объекта', 'Примечания']
    for i, header in enumerate(headers):
        cell = table1.rows[0].cells[i]
        cell.text = ''
        para = cell.paragraphs[0]
        run = para.add_run(header)
        run.bold = True
        run.font.size = Pt(10)
        run.font.name = 'Times New Roman'
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Данные таблицы
    data = table1_data

    # Заполнение таблицы данными
    for row_idx, row_data in enumerate(data, start=1):
        for col_idx, cell_data in enumerate(row_data):
            cell = table1.rows[row_idx].cells[col_idx]
            cell.text = ''
            para = cell.paragraphs[0]
            run = para.add_run(str(cell_data))
            run.font.size = Pt(10)
            run.font.name = 'Times New Roman'
            if col_idx <= 2:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Установка ширины столбцов
    widths = [Cm(1.24), Cm(1.41), Cm(1.59), Cm(6.1), Cm(7.58)]
    for row in table1.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    doc.add_paragraph()

    # Раздел 2: Перечень работ
    add_formatted_paragraph(doc, '\t2. Перечень работ:', size=11)

    # Таблица перечня работ
    table2 = doc.add_table(rows=4, cols=6)
    table2.style = 'Table Grid'
    table2.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Заголовки таблицы 2
    headers2 = ['№ п/п', 'Наименование работ', 'Единца измерения', 'Объем выполняемых работ',
                'Цена работ за единицу, руб. без НДС¹', 'Стоимость работ, руб. без НДС¹']

    for i, header in enumerate(headers2):
        cell = table2.rows[0].cells[i]
        cell.text = ''
        para = cell.paragraphs[0]
        run = para.add_run(header)
        run.bold = False
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Объединение ячеек для "Работы:"
    cell_work = table2.rows[1].cells[0]
    merged_cell = cell_work.merge(table2.rows[1].cells[5])
    merged_cell.text = ''
    para = merged_cell.paragraphs[0]
    run = para.add_run('Работы:')
    run.bold = False
    run.font.size = Pt(9)
    run.font.name = 'Times New Roman'

    # Строка с данными работ
    work_data = ['1',
                 'Электромонтажные и пусконаладочные работы по подключению счетчика электрической энергии трехфазного непосредственного (прямого) включения',
                 'шт.', f'{table1_rows}', f'{price1}', f'{price1*table1_rows}']

    for col_idx, cell_data in enumerate(work_data):
        cell = table2.rows[2].cells[col_idx]
        cell.text = ''
        para = cell.paragraphs[0]
        run = para.add_run(str(cell_data))
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        if col_idx in [0, 2, 3, 4, 5]:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Строка "Оборудование:"
    merged_cell2 = table2.rows[3].cells[0].merge(table2.rows[3].cells[5])
    merged_cell2.text = ''
    para = merged_cell2.paragraphs[0]
    run = para.add_run('Оборудование:')
    run.bold = False
    run.font.size = Pt(9)
    run.font.name = 'Times New Roman'

    # Добавление строк для оборудования (Энергомера)
    row_eq = table2.add_row()
    eq_data = ['2', 'Энергомера CE307 R34.749.OG.QYUVLFZ NB02 SPds', 'шт.', f'{table1_rows}', '-', '-']

    for col_idx, cell_data in enumerate(eq_data):
        cell = row_eq.cells[col_idx]
        cell.text = ''
        para = cell.paragraphs[0]
        run = para.add_run(str(cell_data))
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        if col_idx in [0, 2, 3, 4, 5]:
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Итоговая строка
    total_row = table2.add_row()
    merged_total = total_row.cells[0].merge(total_row.cells[4])
    merged_total.text = ''
    para = merged_total.paragraphs[0]
    run = para.add_run('Итоговая стоимость работ по Заявке, руб. без НДС¹')
    run.bold = False
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'

    total_cell = total_row.cells[5]
    total_cell.text = ''
    para = total_cell.paragraphs[0]
    run = para.add_run(f'{price1*table1_rows}')
    run.bold = False
    run.font.size = Pt(11)
    run.font.name = 'Times New Roman'
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    widths = [Cm(0.99), Cm(8), Cm(1.5), Cm(2), Cm(2.75), Cm(2.29)]
    for row in table2.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width
    # Сноска
    doc.add_paragraph()
    footnote = doc.add_paragraph()
    footnote_run = footnote.add_run('¹Размер НДС определяется по ставке, установленной п. 3 ст. 164 НК РФ.\n3. Требования к составу материалов и оборудования: Договором предусмотрено давальческие материалы/оборудование в составе (Заполняется при наличии давальческих материалов/оборудования):')
    footnote_run.font.size = Pt(11)
    footnote_run.font.name = 'Times New Roman'
    doc.add_paragraph()

    # Таблица давальческих материалов
    table3 = doc.add_table(rows=2, cols=4)
    table3.style = 'Table Grid'

    materials_headers = ['№ п/п', 'Наименование давальческих материалов/оборудования', 'Единца измерения', 'Количество']
    for i, header in enumerate(materials_headers):
        cell = table3.rows[0].cells[i]
        cell.text = ''
        para = cell.paragraphs[0]
        run = para.add_run(header)
        run.bold = False
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Строка с прочерками
    dash_data = ['1', '-', '-', '-']
    for i, data in enumerate(dash_data):
        cell = table3.rows[1].cells[i]
        cell.text = ''
        para = cell.paragraphs[0]
        run = para.add_run(data)
        run.font.size = Pt(11)
        run.font.name = 'Times New Roman'
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER

    widths = [Cm(1.43), Cm(8.9), Cm(3.33), Cm(3.37)]
    for row in table3.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width
    # Раздел 4
    add_formatted_paragraph(doc,
                            f'4. Стоимость работ, выполняемых по настоящей Заявке определена в соответствии с условиями Договора и составляет {price1*table1_rows} ({sum_to_words(price1*table1_rows)}), в том числе НДС 22% {price1*table1_rows*0.22} ({sum_to_words(table1_rows*price1*0.22)}).',
                            bold=False, size=11)
    # Раздел 5
    add_formatted_paragraph(doc,
                            '5. Окончательная стоимость работ определяется исходя из фактических объемов, зафиксированных в Акте о приемке выполненных работ, и не может превышать стоимость работ, указанных в п.4 настоящей Заявки.',
                            bold=False, size=11)
    # Раздел 6
    add_formatted_paragraph(doc,
                            '6. В случае, если фактическая стоимость работ оказалась меньше суммы, указанной в п. 4. настоящей Заявки, разница между суммой, указанной в п. 4 настоящей Заявки и стоимостью фактически выполненных работ остается у Заказчика.',
                            bold=False, size=11)
    # Раздел 7
    add_formatted_paragraph(doc,
                            f'7. {date_work}.',
                            bold=False, size=11)
    # Подписи
    add_formatted_paragraph(doc, 'Заказчик', size=11)
    add_formatted_paragraph(doc, 'Должность:', size=11)
    add_formatted_paragraph(doc, '__________________ / ___________ / М.П.', size=10)

    # Сохранение документа
    output_path = 'out1.docx'
    doc.save(output_path)
    print(f"Документ сохранен как: {output_path}")

    return output_path

