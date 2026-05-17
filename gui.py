import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from docx import Document
from parse import find_info_in_doc
from cr1 import create_application_doc
from cr1 import price_to_show

try:
    from cr2 import create_application_doc2
except ImportError:
    create_application_doc2 = None

WORK_DICT = {
    "1ф": "Электромонтажные и пусконаладочные работы по подключению счетчика электрической энергии однофазного",
    "3ф ПР": "Электромонтажные и пусконаладочные работы по подключению счетчика электрической энергии трехфазного непосредственного (прямого) включения",
    "3ф ПК": "Электромонтажные и пусконаладочные работы по подключению счетчика электрической энергии трехфазного трансформаторного включения"
}


class App:
    def __init__(self, root):
        self.table2_data = None
        self.price_temp = None
        self.root = root
        self.root.title("Обработчик заявок")
        self.root.geometry("1200x700")

        self.current_data = []  # Данные из файла
        self.active_rows = []  # Активные строки (индексы)
        self.file_path = None
        self.cur7 = None
        self.current_template = 1  # 1 или 2

        # Цены для разных типов работ
        self.prices = {
            "1ф": 0.0,
            "3ф ПР": 0.0,
            "3ф ПК": 0.0
        }

        self.setup_ui()

    def setup_ui(self):
        # Верхняя панель
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(top_frame, text="Выбрать Word файл", command=self.load_file,
                  bg="#2196F3", fg="white", font=("Arial", 11, "bold"),
                  height=2, width=20).pack(side=tk.LEFT, padx=5)

        self.lbl_file = tk.Label(top_frame, text="Файл не выбран", fg="gray")
        self.lbl_file.pack(side=tk.LEFT, padx=10)

        # Кнопки шаблонов
        template_frame = tk.Frame(top_frame)
        template_frame.pack(side=tk.LEFT, padx=20)

        self.btn_template1 = tk.Button(template_frame, text="Шаблон 1",
                                       command=lambda: self.switch_template(1),
                                       bg="#4CAF50", fg="white",
                                       font=("Arial", 10, "bold"),
                                       relief="sunken", width=12)
        self.btn_template1.pack(side=tk.LEFT, padx=2)

        self.btn_template2 = tk.Button(template_frame, text="Шаблон 2",
                                       command=lambda: self.switch_template(2),
                                       bg="#607D8B", fg="white",
                                       font=("Arial", 10, "bold"),
                                       relief="raised", width=12)
        self.btn_template2.pack(side=tk.LEFT, padx=2)

        tk.Button(top_frame, text="Создать документ", command=self.create_doc,
                  bg="#4CAF50", fg="white", font=("Arial", 11, "bold"),
                  height=2, width=20).pack(side=tk.RIGHT, padx=5)

        # Параметры
        param_frame = tk.Frame(self.root)
        param_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(param_frame, text="Номер заявки:", font=("Arial", 10)).pack(side=tk.LEFT)
        self.entry_num = tk.Entry(param_frame, width=10, font=("Arial", 10))
        self.entry_num.insert(0, "0")
        self.entry_num.pack(side=tk.LEFT, padx=5)

        tk.Label(param_frame, text="Дата:", font=("Arial", 10)).pack(side=tk.LEFT, padx=(20, 5))
        self.entry_date = tk.Entry(param_frame, width=25, font=("Arial", 10))
        self.entry_date.insert(0, "« 00 » января 2000 г.")
        self.entry_date.pack(side=tk.LEFT, padx=5)

        # Поля для цен
        price_frame = tk.Frame(self.root)
        price_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(price_frame, text="Цены за единицу работы (руб.):",
                 font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))

        # Создаем поля для каждого типа работ
        self.price_entries = {}
        for i, (code, description) in enumerate(WORK_DICT.items()):
            frame = tk.Frame(price_frame)
            frame.pack(side=tk.LEFT, padx=10)

            # Сокращенное описание для метки
            short_desc = code
            tk.Label(frame, text=f"{short_desc}:", font=("Arial", 9)).pack(side=tk.LEFT)

            entry = tk.Entry(frame, width=12, font=("Arial", 10))
            entry.insert(0, "0")
            entry.pack(side=tk.LEFT, padx=2)
            self.price_entries[code] = entry

            # Добавляем подсказку при наведении
            self.create_tooltip(entry, description)

        # Метка текущего шаблона
        self.lbl_template = tk.Label(param_frame, text="Текущий шаблон: 1",
                                     font=("Arial", 10, "bold"), fg="#4CAF50")
        self.lbl_template.pack(side=tk.RIGHT, padx=20)

        # Панель управления строками
        ctrl_frame = tk.Frame(self.root)
        ctrl_frame.pack(fill=tk.X, padx=10, pady=5)

        self.lbl_count = tk.Label(ctrl_frame, text="Записей: 0 | Активных: 0", font=("Arial", 10, "bold"))
        self.lbl_count.pack(side=tk.LEFT)

        tk.Button(ctrl_frame, text="✅ Все", command=lambda: self.toggle_all(True),
                  bg="#607D8B", fg="white", width=10).pack(side=tk.RIGHT, padx=2)
        tk.Button(ctrl_frame, text="❌ Ничего", command=lambda: self.toggle_all(False),
                  bg="#607D8B", fg="white", width=10).pack(side=tk.RIGHT, padx=2)
        tk.Button(ctrl_frame, text="🔄 Инвертировать", command=self.invert_selection,
                  bg="#607D8B", fg="white", width=10).pack(side=tk.RIGHT, padx=2)

        # Таблица с чекбоксами
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Canvas + Scrollbar для прокрутки
        canvas = tk.Canvas(table_frame)
        scrollbar = tk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)

        self.scrollable_frame.bind("<Configure>",
                                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Заголовки таблицы
        headers = ["Акт.", "№п/п", "Кол-во ПУ", "Тип ПУ", "Адрес объекта", "Примечания"]
        widths = [5, 5, 8, 8, 40, 40]

        header_frame = tk.Frame(self.scrollable_frame, bg="#e0e0e0")
        header_frame.pack(fill=tk.X)

        for i, (h, w) in enumerate(zip(headers, widths)):
            tk.Label(header_frame, text=h, font=("Arial", 9, "bold"),
                     bg="#e0e0e0", borderwidth=1, relief="solid",
                     width=w, anchor="w", padx=5).grid(row=0, column=i, sticky="nsew")

        self.table_frame = tk.Frame(self.scrollable_frame)
        self.table_frame.pack(fill=tk.X)

        self.checkboxes = []
        self.labels = []

    def create_tooltip(self, widget, text):
        """Создание всплывающей подсказки для виджета"""

        def enter(event):
            x, y, _, _ = widget.bbox("insert")
            x += widget.winfo_rootx() + 25
            y += widget.winfo_rooty() + 25

            # Создаем окно подсказки
            self.tooltip = tk.Toplevel(widget)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x}+{y}")

            # Разбиваем длинный текст на строки
            words = text.split()
            lines = []
            current_line = []
            for word in words:
                current_line.append(word)
                if len(' '.join(current_line)) > 50:
                    lines.append(' '.join(current_line[:-1]))
                    current_line = [word]
            if current_line:
                lines.append(' '.join(current_line))

            label = tk.Label(self.tooltip, text='\n'.join(lines),
                             justify=tk.LEFT,
                             background="#ffffe0",
                             relief=tk.SOLID,
                             borderwidth=1,
                             font=("Arial", "8", "normal"))
            label.pack()

        def leave(event):
            if hasattr(self, 'tooltip'):
                self.tooltip.destroy()

        widget.bind('<Enter>', enter)
        widget.bind('<Leave>', leave)

    def switch_template(self, template_num):
        """Переключение между шаблонами"""
        self.current_template = template_num

        if template_num == 1:
            self.btn_template1.config(bg="#4CAF50", relief="sunken")
            self.btn_template2.config(bg="#607D8B", relief="raised")
            self.lbl_template.config(text="Текущий шаблон: 1", fg="#4CAF50")

        else:
            self.btn_template1.config(bg="#607D8B", relief="raised")
            self.btn_template2.config(bg="#4CAF50", relief="sunken")
            self.lbl_template.config(text="Текущий шаблон: 2", fg="#4CAF50")

            # Проверяем наличие cr2.py
            if create_application_doc2 is None:
                messagebox.showwarning("Предупреждение",
                                       "Файл cr2.py не найден!\nБудет использован шаблон 1.")
                self.switch_template(1)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="Выберите Word документ",
            filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
        )

        if not file_path:
            return

        try:
            self.file_path = file_path
            self.lbl_file.config(text=os.path.basename(file_path), fg="green")

            # Используем parse.py для извлечения данных
            doc = Document(file_path)
            info = find_info_in_doc(doc)
            self.cur7 = info["дата выполнения"]

            # Заполняем параметры если нашли
            if info["номер заявки"]:
                self.entry_num.delete(0, tk.END)
                self.entry_num.insert(0, info["номер заявки"])

            if info["дата обращения"]:
                self.entry_date.delete(0, tk.END)
                self.entry_date.insert(0, info["дата обращения"][:50])

            if info["таблица работ"]:
                self.table2_data = info["таблица работ"]
                # Извлекаем цены из таблицы работ
                self.extract_prices_from_table(info["таблица работ"])

            # Берем адреса объектов
            self.current_data = []
            for row in info["адреса объектов"]:
                if row[0] != '№п/п':
                    self.current_data.append(row)

            if not self.current_data:
                messagebox.showwarning("Предупреждение", "Не удалось найти данные объектов в файле")
                return

            self.update_table()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

    def extract_prices_from_table(self, table_data):
        """Извлечение цен из таблицы работ для каждого типа"""
        # Сбрасываем все цены в 0
        for code in self.prices:
            self.prices[code] = 0.0
            self.price_entries[code].delete(0, tk.END)
            self.price_entries[code].insert(0, "0")

        # Если данных нет, выходим
        if not table_data or len(table_data) < 4:
            return

        # Пропускаем первые 2 строки (заголовки)
        data_rows = table_data[2:-2]  # Убираем заголовки и итоги

        if len(data_rows) < 2:
            return

        # Разделяем на работы и оборудование
        mid_point = len(data_rows) // 2
        work_rows = data_rows[:mid_point]

        # Сопоставляем работы с типами
        # Проходим по всем строкам работ и ищем соответствия
        for row in work_rows:
            if len(row) < 5:
                continue

            work_name = str(row[1]).lower() if row[1] else ""
            price_str = str(row[4]) if row[4] else "0"

            # Очищаем строку цены
            try:
                price = float(price_str.replace('\xa0', ' ').replace(' ', '').replace(',', '.'))
            except:
                price = 0.0

            # Определяем тип работы по названию
            if "однофазного" in work_name:
                self.prices["1ф"] = price
                self.price_entries["1ф"].delete(0, tk.END)
                self.price_entries["1ф"].insert(0, str(price))
            elif "трансформаторного" in work_name or "полукосвенного" in work_name or "полукосвенного" in work_name:
                self.prices["3ф ПК"] = price
                self.price_entries["3ф ПК"].delete(0, tk.END)
                self.price_entries["3ф ПК"].insert(0, str(price))
            elif "трехфазного" in work_name and "непосредственного" in work_name or "прямого" in work_name:
                self.prices["3ф ПР"] = price
                self.price_entries["3ф ПР"].delete(0, tk.END)
                self.price_entries["3ф ПР"].insert(0, str(price))

    def update_prices_from_entries(self):
        """Обновление цен из полей ввода"""
        for code, entry in self.price_entries.items():
            try:
                value = float(entry.get().replace(',', '.').replace(' ', ''))
                self.prices[code] = value
            except ValueError:
                self.prices[code] = 0.0
                entry.delete(0, tk.END)
                entry.insert(0, "0")

    def update_table(self):
        # Очищаем таблицу
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        self.checkboxes = []
        self.active_rows = list(range(len(self.current_data)))

        # Фиксированные ширины столбцов в пикселях
        col_widths = {
            0: 50,  # Чекбокс
            1: 60,  # №п/п
            2: 70,  # Кол-во ПУ
            3: 80,  # Тип ПУ
            4: 300,  # Адрес объекта
            5: 300  # Примечания
        }

        for row_idx, row_data in enumerate(self.current_data):
            # Чекбокс
            var = tk.BooleanVar(value=True)
            cb = tk.Checkbutton(self.table_frame, variable=var,
                                command=lambda r=row_idx, v=var: self.on_check(r, v))
            cb.grid(row=row_idx, column=0, sticky="w", padx=2)
            cb.var = var
            self.checkboxes.append(cb)

            # Данные
            for col_idx, value in enumerate(row_data):
                if col_idx < 5:
                    # Обрезаем текст для отображения
                    text = str(value)
                    display_text = text[:80] + "..." if len(text) > 100 else text

                    lbl = tk.Label(
                        self.table_frame,
                        text=display_text,
                        font=("Arial", 8),
                        borderwidth=1,
                        relief="solid",
                        width=col_widths[col_idx + 1] // 7,  # Примерная ширина в символах
                        anchor="w",
                        padx=3,
                        wraplength=col_widths[col_idx + 1] - 10  # Перенос текста
                    )
                    lbl.grid(row=row_idx, column=col_idx + 1, sticky="nsew")

            # Цвет фона
            bg_color = "#f5f5f5" if row_idx % 2 == 0 else "white"
            for col in range(6):
                widgets = self.table_frame.grid_slaves(row=row_idx, column=col)
                for widget in widgets:
                    try:
                        widget.config(bg=bg_color)
                    except:
                        pass

        # Фиксируем ширину колонок через grid_columnconfigure
        for col, width in col_widths.items():
            self.table_frame.grid_columnconfigure(col, minsize=width)

        self.lbl_count.config(text=f"Записей: {len(self.current_data)} | Активных: {len(self.active_rows)}")

    def on_check(self, row, var):
        if var.get():
            if row not in self.active_rows:
                self.active_rows.append(row)
        else:
            if row in self.active_rows:
                self.active_rows.remove(row)

        self.active_rows.sort()
        self.lbl_count.config(text=f"Записей: {len(self.current_data)} | Активных: {len(self.active_rows)}")

    def toggle_all(self, state):
        self.active_rows = list(range(len(self.current_data))) if state else []
        for cb in self.checkboxes:
            cb.var.set(state)
        self.lbl_count.config(text=f"Записей: {len(self.current_data)} | Активных: {len(self.active_rows)}")

    def invert_selection(self):
        new_active = []
        for i in range(len(self.current_data)):
            if i not in self.active_rows:
                new_active.append(i)
                self.checkboxes[i].var.set(True)
            else:
                self.checkboxes[i].var.set(False)
        self.active_rows = new_active
        self.lbl_count.config(text=f"Записей: {len(self.current_data)} | Активных: {len(self.active_rows)}")

    def create_doc(self):
        if not self.current_data:
            messagebox.showwarning("Предупреждение", "Нет данных для создания документа")
            return

        if not self.active_rows:
            messagebox.showwarning("Предупреждение", "Нет активных записей. Отметьте хотя бы одну строку.")
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить документ",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")]
        )

        if not file_path:
            return

        try:
            # Обновляем цены из полей ввода
            self.update_prices_from_entries()

            # Получаем данные только активных строк и нормализуем их
            active_data = []
            for i in self.active_rows:
                row = self.current_data[i]
                # Проверяем что строка не пустая
                if not row or all(cell == '' for cell in row):
                    continue

                # Нормализуем строку: должно быть ровно 5 элементов
                normalized_row = list(row)[:5]
                while len(normalized_row) < 5:
                    normalized_row.append("")
                active_data.append(normalized_row)

            if not active_data:
                messagebox.showwarning("Предупреждение", "Нет данных для создания документа")
                return

            num = self.entry_num.get() or "ERROR_IN_NUM"
            date = self.entry_date.get() or "ERROR_IN_DATE"

            # Обновляем table2_data с новыми ценами
            updated_table2_data = self.update_table2_data_with_prices()

            # ВАЖНО: table1_rows должно быть количество строк данных + 1 (заголовок)
            table1_rows = len(active_data) + 1

            # Выбираем функцию создания в зависимости от шаблона
            if self.current_template == 2 and create_application_doc2 is not None:
                output = create_application_doc2(
                    num=num,
                    date=date,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    table2_data=updated_table2_data
                )
                template_name = "Шаблон 2"
            else:
                output = create_application_doc(
                    num=num,
                    date=date,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    table2_data=updated_table2_data,
                    date_work=self.cur7
                )
                template_name = "Шаблон 1"

            # Перемещаем файл
            import shutil
            if output and os.path.exists(output):
                shutil.move(output, file_path)

            messagebox.showinfo("Успех",
                                f"Документ создан!\n\n"
                                f"{template_name}\n"
                                f"Активных записей: {len(active_data)}\n"
                                f"Сохранен в:\n{file_path}")

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось создать документ:\n{str(e)}")

    def update_table2_data_with_prices(self):
        """Обновление table2_data с учетом пользовательских цен"""
        if not self.table2_data:
            return None

        # Создаем копию данных
        updated_data = []
        for row in self.table2_data:
            updated_data.append(list(row))

        # Пропускаем первые 2 строки (заголовки)
        data_start = 2
        data_end = len(updated_data) - 2

        if data_end - data_start < 2:
            return updated_data

        # Разделяем на работы и оборудование
        work_rows_indices = []
        equip_rows_indices = []

        found_equip = False
        for i in range(data_start, data_end):
            cell_text = str(updated_data[i][1]).lower() if len(updated_data[i]) > 1 else ""

            if "оборудование" in cell_text:
                found_equip = True
                continue

            if not found_equip:
                work_rows_indices.append(i)
            else:
                equip_rows_indices.append(i)

        # Обновляем цены в строках работ
        for idx in work_rows_indices:
            if len(updated_data[idx]) < 5:
                continue

            work_name = str(updated_data[idx][1]).lower() if updated_data[idx][1] else ""

            # Определяем тип работы и обновляем цену
            if "однофазного" in work_name:
                updated_data[idx][4] = str(self.prices["1ф"])
            elif "трансформаторного" in work_name or "полукосвенного" in work_name:
                updated_data[idx][4] = str(self.prices["3ф ПК"])
            elif "трехфазного" in work_name and ("непосредственного" in work_name or "прямого" in work_name):
                updated_data[idx][4] = str(self.prices["3ф ПР"])

        return updated_data


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()