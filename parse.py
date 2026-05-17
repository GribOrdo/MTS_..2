import tkinter as tk
from tkinter import filedialog
from docx import Document
import re
import datetime


def find_info_in_doc(doc):
    info = {
        "номер заявки": None,
        "дата обращения": None,
        "таблица работ": [],
        "суммарная стоимость": None,
        "дата выполнения": None,
        "адреса объектов": []
    }

    # === Извлечение текста из параграфов ===
    full_text = "\n".join([p.text for p in doc.paragraphs])

    # === Поиск текстовой информации ===

    # Номер заявки
    match = re.search(r'Заявка\s*на\s*работы\s*№\s*(\d+)', full_text, re.IGNORECASE)
    if match:
        info["номер заявки"] = match.group(0).split("№")[1].strip()

    # Дата обращения - ищем дату в первых строках
    match = re.search(r'Приложение\s+1 [^\n\r]*\n([^\n\r]+)', full_text, re.IGNORECASE)
    if match:
        info["дата обращения"] = match.group(0).split(" от ")[1].strip()

    # Дата выполнения работ
    match = re.search(r'Дата\s*выполнения\s*работ:\s*([^\n\r]+)', full_text, re.IGNORECASE)
    if match:
        info["дата выполнения"] = match.group(0).strip()

    # Таблица работ: ищем таблицу с заголовками работ
    for table in doc.tables:
        if table.rows:
            first_row_cells = [cell.text.strip().lower() for cell in table.rows[0].cells]
            first_row_text = " ".join(first_row_cells)

            if ("наименование работ" in first_row_text) and "стоимость работ" in first_row_text:
                for i, row in enumerate(table.rows):
                    row_data = [cell.text.strip() for cell in row.cells]
                    if any(row_data) and not all(cell == "" for cell in row_data):
                        if len(row_data) > 0 and not re.match(r'^(Работы|Оборудование)[:\s]*$', row_data[0]):
                            info["таблица работ"].append(row_data)

                # Ищем итоговую стоимость
                for row in table.rows:
                    row_text = " ".join([cell.text.strip() for cell in row.cells])
                    if "итоговая стоимость" in row_text.lower():
                        # Берем последнюю ячейку
                        last_cell = row.cells[-1].text.strip()
                        if last_cell:
                            info["суммарная стоимость"] = last_cell
                break

    # Таблица адресов объектов
    for table in doc.tables:
        if table.rows:
            first_row_cells = [cell.text.strip().lower() for cell in table.rows[0].cells]
            first_row_text = " ".join(first_row_cells)

            if "адрес объекта" in first_row_text or ("№п/п" in first_row_text and "адрес" in first_row_text):
                for i, row in enumerate(table.rows):
                    row_data = [cell.text.strip() for cell in row.cells]
                    if any(row_data) and not all(cell == "" for cell in row_data):
                        info["адреса объектов"].append(row_data)
                break
    print(info["адреса объектов"])
    return info


def open_file(path):
    file_path = path
    if not file_path:
        return

    doc = Document(file_path)
    info = find_info_in_doc(doc)

    # Вывод информации
    print("\n=== Извлечённая информация ===")
    print("Номер заявки:", info["номер заявки"])
    print("Дата обращения:", info["дата обращения"])
    print("Дата выполнения:", info["дата выполнения"])
    print("Суммарная стоимость:", info["суммарная стоимость"])
    print("Таблица работ:")
    for row in info["таблица работ"]:
        print("  ", row)
    print("Адреса объектов:")
    for row in info["адреса объектов"]:
        print("  ", row)

    return info
