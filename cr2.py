from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from num2words import num2words


def sum_to_words(amount):
    rub = int(amount)
    kop = int(round((amount - rub) * 100))

    rub_text = num2words(rub, lang='ru')
    kop_text = num2words(kop, lang='ru')

    return f"{rub_text} руб. {kop_text} коп."



def create_application_doc2(num, date, table1_rows, table1_data, price1):
    doc = Document()

    # Настройка стилей
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(12)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    # Установка узких полей
    sections = doc.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(1)

    # Заголовок акта
    title_para = doc.add_paragraph()
    title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_para.add_run(f'Акт о приемке выполненных работ № {num}')
    title_run.bold = True
    title_run.font.size = Pt(12)
    title_run.font.name = 'Times New Roman'

    # Таблица для даты и города
    header_table = doc.add_table(rows=1, cols=2)
    header_table.alignment = WD_TABLE_ALIGNMENT.LEFT

    cells = header_table.rows[0].cells
    cells[0].text = 'г. Новосибирск'
    cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
    cells[1].text = f'«30» апреля 2026 г.'
    cells[1].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    cells[0].width = Cm(9)
    cells[1].width = Cm(9)

    # Убираем границы таблицы
    for cell in header_table.rows[0].cells:
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for border_name in ['top', 'left', 'bottom', 'right']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'nil')
            tcBorders.append(border)
        tcPr.append(tcBorders)

    main_t1 = 'Акционерное общество «Новосибирскэнергосбыт» (АО «Новосибирскэнергосбыт»), '
    main_t2 = 'именуемое в дальнейшем '
    main_t3 = '«Заказчик»'
    main_t4 = ', в лице Заместителя генерального директора — Технического директора Михайлишина Александра Юрьевича, действующего на основании Доверенности № 24/2024 от 02.02.2024 г., с одной стороны, и '
    main_t5 = 'Публичное акционерное общество «Мобильные ТелеСистемы» (ПАО «МТС») '
    main_t6 = 'именуемое в дальнейшем'
    main_t7 = ' «Подрядчик»'
    main_t8 = (', в лице Директора Департамента по работе с '
               'корпоративными клиентами Ватуля Елены Николаевны, действующего на основании '
               'Доверенности № 77/509-н/77-2024-3-661 от 19.09.2024 г., другой стороны, '
               'составили настоящий Акт к договору подряда №39278 от 28.02.2025 г. в том, что:')

    para = doc.add_paragraph()
    para.paragraph_format.space_after = Pt(6)
    para.add_run(main_t1).bold = True
    para.add_run(main_t2)
    para.add_run(main_t3).bold = True
    para.add_run(main_t4)
    para.add_run(main_t5).bold = True
    para.add_run(main_t6)
    para.add_run(main_t7).bold = True
    para.add_run(main_t8)
    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    # Пункт 1
    para1 = doc.add_paragraph()
    para1.paragraph_format.first_line_indent = Cm(1.25)
    run1 = para1.add_run(
        f'1. В соответствии с заключенным между Подрядчиком и Заказчиком договором, '
        f'Подрядчик выполнил, а Заказчик принял следующие работы по Заявке на работы '
        f'№ {num} от {date}'
    )
    run1.font.name = 'Times New Roman'
    run1.font.size = Pt(12)

    # Пункт 2 - таблица объектов
    para2 = doc.add_paragraph()
    para2.paragraph_format.first_line_indent = Cm(1.25)
    run2 = para2.add_run('2. Объект по адресу:')
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(12)

    # Создание таблицы объектов
    objects_data = table1_data
    table = doc.add_table(rows=table1_rows, cols=5)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Заголовки таблицы
    headers = ['№п/п', 'кол-во ПУ', 'тип ПУ', 'Адрес объекта', 'Примечания']
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(10)

    # Заполнение данных
    for i, row_data in enumerate(objects_data):
        for j, cell_text in enumerate(row_data):
            cell = table.rows[i + 1].cells[j]
            cell.text = ''
            p = cell.paragraphs[0]
            run = p.add_run(str(cell_text))
            run.font.name = 'Times New Roman'
            run.font.size = Pt(10)

    # Настройка ширины колонок
    widths = [Cm(1.24), Cm(1.41), Cm(1.59), Cm(6.1), Cm(7.58)]
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = width

    doc.add_paragraph()

    # Пункт 3 - Таблица работ
    para3 = doc.add_paragraph()
    para3.paragraph_format.first_line_indent = Cm(1.25)
    run3 = para3.add_run('3. Перечень выполненных работ:')
    run3.font.name = 'Times New Roman'
    run3.font.size = Pt(12)

    # Таблица работ
    work_table = doc.add_table(rows=7, cols=6)
    work_table.style = 'Table Grid'
    work_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Количество объектов (без строки заголовка)
    objects_count = table1_rows - 1

    # Заголовки
    work_headers = ['№ п/п', 'Наименование работ', 'Единица измерения',
                    'Объем выполняемых работ', 'Цена работ за единицу, руб. с НДС',
                    'Стоимость работ, руб. с НДС']

    for i, header in enumerate(work_headers):
        cell = work_table.rows[0].cells[i]
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(8)

    # Объединение ячеек для заголовка "Работы:"
    work_table.rows[1].cells[0].merge(work_table.rows[1].cells[5])
    cell_work = work_table.rows[1].cells[0]
    cell_work.text = ''
    p = cell_work.paragraphs[0]
    run = p.add_run('Работы:')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)

    # Данные работ
    work_data = ['1',
                 'Электромонтажные и пусконаладочные работы по подключению счетчика электрической энергии трехфазного непосредственного (прямого) включения',
                 'шт', str(objects_count), f'{price1:.2f}', f'{price1 * objects_count:.2f}']

    for j, cell_text in enumerate(work_data):
        cell = work_table.rows[2].cells[j]
        cell.text = ''
        p = cell.paragraphs[0]
        run = p.add_run(str(cell_text))
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9)
        if j in [0, 2, 3, 4, 5]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Объединение ячеек для заголовка "Оборудование:"
    work_table.rows[3].cells[0].merge(work_table.rows[3].cells[5])
    cell_eq = work_table.rows[3].cells[0]
    cell_eq.text = ''
    p = cell_eq.paragraphs[0]
    run = p.add_run('Оборудование:')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)

    # Данные оборудования
    CONST_EQUIP = 12352.94
    eq_data = ['2',
               'Счетчик электрической энергии трехфазный непосредственного (прямого) включения, соответствующий требованиям ПП РФ № 890 от 19.06.2020 г., Энергомера CE307 R34.749',
               'шт', str(objects_count), f'{CONST_EQUIP:.2f}', f'{CONST_EQUIP * objects_count:.2f}']

    for j, cell_text in enumerate(eq_data):
        cell = work_table.rows[4].cells[j]
        cell.text = ''
        p = cell.paragraphs[0]
        run = p.add_run(str(cell_text))
        run.font.name = 'Times New Roman'
        run.font.size = Pt(9)
        if j in [0, 2, 3, 4, 5]:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Итоговая строка
    total_cost = (price1 + CONST_EQUIP) * objects_count
    nds_amount = total_cost * 0.22

    work_table.rows[5].cells[0].merge(work_table.rows[5].cells[4])
    cell_total_label = work_table.rows[5].cells[0]
    cell_total_label.text = ''
    p = cell_total_label.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run('ИТОГО:')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10)

    cell_total_value = work_table.rows[5].cells[5]
    cell_total_value.text = ''
    p = cell_total_value.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'{total_cost:.2f}')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(10)

    # Строка НДС (исправлены индексы)
    work_table.rows[6].cells[0].merge(work_table.rows[6].cells[4])
    cell_nds_label = work_table.rows[6].cells[0]
    cell_nds_label.text = ''
    p = cell_nds_label.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run('в том числе НДС 22%')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)

    cell_nds_value = work_table.rows[6].cells[5]
    cell_nds_value.text = ''
    p = cell_nds_value.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f'{nds_amount:.2f}')
    run.font.name = 'Times New Roman'
    run.font.size = Pt(9)

    # Настройка ширины колонок
    work_widths = [Cm(0.99), Cm(6), Cm(1.5), Cm(2), Cm(2.75), Cm(2.29)]
    for row in work_table.rows:
        if len(row.cells) == 6:
            for idx, width in enumerate(work_widths):
                row.cells[idx].width = width

    doc.add_paragraph()

    # Пункт 4
    para4 = doc.add_paragraph()
    para4.paragraph_format.first_line_indent = Cm(1.25)
    run4 = para4.add_run(
        f'4. Всего выполнено работ на сумму: {total_cost:.2f} руб. '
        f'({sum_to_words(total_cost)}), '
        f'в том числе НДС 22% - {nds_amount:.2f} руб. '
        f'({sum_to_words(nds_amount)}).'
    )
    run4.font.name = 'Times New Roman'
    run4.font.size = Pt(12)

    # Пункт 5
    para5 = doc.add_paragraph()
    para5.paragraph_format.first_line_indent = Cm(1.25)
    run5 = para5.add_run(
        '5. Вышеперечисленные работы выполнены полностью и в срок. '
        'Заказчик претензий по объему, качеству и срокам выполненных работ не имеет.'
    )
    run5.font.name = 'Times New Roman'
    run5.font.size = Pt(12)

    # Пункт 6
    para6 = doc.add_paragraph()
    para6.paragraph_format.first_line_indent = Cm(1.25)
    run6 = para6.add_run(
        '6. Настоящий Акт составлен в двух идентичных экземплярах, имеющих равную '
        'юридическую силу, по одному для каждой стороны.'
    )
    run6.font.name = 'Times New Roman'
    run6.font.size = Pt(12)

    doc.add_paragraph()

    # Таблица подписей
    sign_table = doc.add_table(rows=2, cols=2)
    sign_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sign_table.style = 'Table Grid'

    for row in sign_table.rows:
        row.cells[0].width = Cm(9)
        row.cells[1].width = Cm(9)

    # Заголовки
    cell_left = sign_table.rows[0].cells[0]
    cell_left.text = ''
    p = cell_left.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Подрядчик:')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

    cell_right = sign_table.rows[0].cells[1]
    cell_right.text = ''
    p = cell_right.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run('Заказчик:')
    run.bold = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(12)

    # Подписи
    cell_left_sign = sign_table.rows[1].cells[0]
    cell_left_sign.text = ''
    for text_line in ['ПАО «МТС»\n', '', 'Директор Департамента по работе',
                      'с корпоративными клиентами\n', '',
                      '____________________/ Ватуля Е.Н. /', '', 'м.п.']:
        if text_line:
            p = cell_left_sign.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text_line)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            if text_line == "ПАО «МТС»\n":
                run.bold = True

    cell_right_sign = sign_table.rows[1].cells[1]
    cell_right_sign.text = ''
    for text_line in ['АО «Новосибирскэнергосбыт»\n', '',
                      'Заместитель генерального директора -',
                      'Технический директор\n', '',
                      '____________________/Михайлишин А.Ю./', '', 'м.п.']:
        if text_line:
            p = cell_right_sign.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(text_line)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            if text_line == "АО «Новосибирскэнергосбыт»\n":
                run.bold = True

    output_path = 'out2.docx'
    doc.save(output_path)
    print(f"Документ сохранен как: {output_path}")

    return output_path
