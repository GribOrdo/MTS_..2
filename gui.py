import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import os
from docx import Document
from tkinterdnd2 import DND_FILES, TkinterDnD
from parse import find_info_in_doc
from cr1 import create_application_doc
from cr1 import price_to_show
from cr2 import create_application_doc2



class App:
    def __init__(self, root):
        self.table2_data = None
        self.root = root
        self.root.title("Обработчик заявок")
        self.root.geometry("1200x700")

        self.current_data = []  # Данные из файла
        self.active_rows = []  # Активные строки (индексы)
        self.file_path = None
        self.cur7 = None
        self.current_template = 1  # 1 или 2
        self.current_mode = 1   # таблица (1) или строка (1)
        self.setup_ui()
        self.setup_drag_and_drop()

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
                                       relief="sunken", width=10)
        self.btn_template1.pack(side=tk.LEFT, padx=2)

        self.btn_template2 = tk.Button(template_frame, text="Шаблон 2",
                                       command=lambda: self.switch_template(2),
                                       bg="#607D8B", fg="white",
                                       font=("Arial", 10, "bold"),
                                       relief="raised", width=10)
        self.btn_template2.pack(side=tk.LEFT, padx=2)

        # Кнопки режимов
        mode_frame = tk.Frame(top_frame)
        mode_frame.pack(side=tk.LEFT, padx=50)

        self.btn_template3 = tk.Button(template_frame, text="Таблица",
                                       command=lambda: self.switch_mode(1),
                                       bg="#4CAF50", fg="white",
                                       font=("Arial", 10, "bold"),
                                       relief="sunken", width=10)
        self.btn_template3.pack(side=tk.LEFT, padx=2)

        self.btn_template4 = tk.Button(template_frame, text="Текст",
                                       command=lambda: self.switch_mode(2),
                                       bg="#607D8B", fg="white",
                                       font=("Arial", 10, "bold"),
                                       relief="raised", width=10)
        self.btn_template4.pack(side=tk.LEFT, padx=2)

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



        # Метка текущего шаблона
        self.lbl_template = tk.Label(param_frame, text="Текущий шаблон: 1",
                                     font=("Arial", 10, "bold"), fg="#4CAF50")
        self.lbl_template.pack(side=tk.RIGHT, padx=20)

        self.lbl_mode = tk.Label(param_frame, text="Текущий режим: таблица",
                                     font=("Arial", 10, "bold"), fg="#4CAF50")
        self.lbl_mode.pack(side=tk.LEFT, padx=60)

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


        else:
            self.btn_template1.config(bg="#607D8B", relief="raised")
            self.btn_template2.config(bg="#4CAF50", relief="sunken")
            self.lbl_template.config(text="Текущий шаблон: 2", fg="#4CAF50")

    def switch_mode(self, mode_num):
        """Переключение между шаблонами"""
        self.current_mode = mode_num

        if mode_num == 1:
            self.btn_template3.config(bg="#4CAF50", relief="sunken")
            self.btn_template4.config(bg="#607D8B", relief="raised")
            self.lbl_mode.config(text="Текущий режим: таблица", fg="#4CAF50")


        else:
            self.btn_template3.config(bg="#607D8B", relief="raised")
            self.btn_template4.config(bg="#4CAF50", relief="sunken")
            self.lbl_mode.config(text="Текущий режим: текст", fg="#4CAF50")




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
            info = find_info_in_doc(doc, file_path)
            self.cur7 = info["дата выполнения"]
            # Заполняем параметры если нашли
            if info["номер заявки"]:
                self.entry_num.delete(0, tk.END)
                self.entry_num.insert(0, info["номер заявки"])

            if info["дата обращения"]:
                self.entry_date.delete(0, tk.END)
                self.entry_date.insert(0, info["дата обращения"][:50])  # Обрезаем если длинное

            if info["таблица работ"]:
                self.table2_data = info["таблица работ"]

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
            # ВАЖНО: table1_rows должно быть количество строк данных + 1 (заголовок)
            table1_rows = len(active_data) + 1  # +1 для строки заголовка

            # Выбираем функцию создания в зависимости от шаблона
            if self.current_template == 2:
                output = create_application_doc2(
                    num=num,
                    date=date,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    table2_data=self.table2_data,
                    mode=self.current_mode
                )
                template_name = "Шаблон 2"
            else:
                output = create_application_doc(
                    num=num,
                    date=date,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    table2_data=self.table2_data,
                    date_work=self.cur7,
                    mode=self.current_mode
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

    # Inside your App class, add a method to set up drag-and-drop:
    def setup_drag_and_drop(self):
        # Create a label or frame to act as a drag-and-drop target
        self.drop_label = tk.Label(self.root, text="Перетащите файл сюда", bg="#e0e0e0", height=4)
        self.drop_label.pack(fill=tk.X, padx=10, pady=10)

        # Register the label as a drop target
        self.drop_label.drop_target_register(DND_FILES)

        def on_drop(event):
            # event.data contains the file path(s), possibly enclosed in braces if path contains spaces
            files = self.root.tk.splitlist(event.data)
            if files:
                file_path = files[0]
                self.load_file_from_path(file_path)

        # Bind the drop event
        self.drop_label.dnd_bind('<<Drop>>', on_drop)

    def load_file_from_path(self, file_path):
        try:
            self.file_path = file_path
            self.lbl_file.config(text=os.path.basename(file_path), fg="green")
            # Your existing logic to process the file
            doc = Document(file_path)
            info = find_info_in_doc(doc, file_path)
            self.cur7 = info["дата выполнения"]
            # Fill in your data...
            if info["номер заявки"]:
                self.entry_num.delete(0, tk.END)
                self.entry_num.insert(0, info["номер заявки"])

            if info["дата обращения"]:
                self.entry_date.delete(0, tk.END)
                self.entry_date.insert(0, info["дата обращения"][:50])  # Limit length

            if info["таблица работ"]:
                self.table2_data = info["таблица работ"]

            # Обновляем таблицу
            self.current_data = []
            for row in info["адреса объектов"]:
                if row[0] != '№п/п':
                    self.current_data.append(row)  # Берем первые 5 колонок

            self.update_table()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Ошибка", f"Не удалось загрузить файл:\n{str(e)}")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = App(root)
    root.mainloop()
