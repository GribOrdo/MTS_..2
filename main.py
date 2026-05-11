import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from docx import Document
from parse import find_info_in_doc
from cr1 import create_application_doc

try:
    from cr2 import create_application_doc2
except ImportError:
    create_application_doc2 = None


class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Обработчик заявок")
        self.root.geometry("1200x700")

        self.current_data = []  # Данные из файла
        self.active_rows = []  # Активные строки (индексы)
        self.file_path = None
        self.cur7 = None
        self.current_template = 1  # 1 или 2
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

        tk.Label(param_frame, text="Цена за ед.:", font=("Arial", 10)).pack(side=tk.LEFT, padx=(20, 5))
        self.entry_price = tk.Entry(param_frame, width=10, font=("Arial", 10))
        self.entry_price.insert(0, "0")
        self.entry_price.pack(side=tk.LEFT, padx=5)

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

    def switch_template(self, template_num):
        """Переключение между шаблонами"""
        self.current_template = template_num

        if template_num == 1:
            self.btn_template1.config(bg="#4CAF50", relief="sunken")
            self.btn_template2.config(bg="#607D8B", relief="raised")
            self.lbl_template.config(text="Текущий шаблон: 1", fg="#4CAF50")
            self.entry_price.delete(0, tk.END)
            self.entry_price.insert(0, "1250.00")

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
                self.entry_date.insert(0, info["дата обращения"][:50])  # Обрезаем если длинное

            if info["таблица работ"]:
                self.entry_price.delete(0, tk.END)
                self.entry_price.insert(0, info["таблица работ"][2][4])

            # Берем адреса объектов
            self.current_data = []
            for row in info["адреса объектов"]:
                if row[0] != '№п/п':
                    self.current_data.append(row)  # Берем первые 5 колонок

            if not self.current_data:
                messagebox.showwarning("Предупреждение", "Не удалось найти данные объектов в файле")
                return

            self.update_table()

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")

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
            price = float(self.entry_price.get().replace(',', '.').replace(" ", "")) or "ERROR_IN_PRICE"
            # ВАЖНО: table1_rows должно быть количество строк данных + 1 (заголовок)
            table1_rows = len(active_data) + 1  # +1 для строки заголовка

            # Выбираем функцию создания в зависимости от шаблона
            if self.current_template == 2 and create_application_doc2 is not None:
                output = create_application_doc2(
                    num=num,
                    date=date,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    price1=price,
                )
                template_name = "Шаблон 2"
            else:
                output = create_application_doc(
                    num=num,
                    date=date,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    price1=price,
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
            traceback.print_exc()  # Выведет полную ошибку в консоль
            messagebox.showerror("Ошибка", f"Не удалось создать документ:\n{str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
