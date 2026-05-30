import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from docx import Document
from parse import find_info_in_doc
from cr1 import create_application_doc, price_to_show
from cr2 import create_application_doc2
import datetime
from PIL import Image, ImageTk
from config import COLORS, SIZES, FONTS, TABLE_COLUMNS, TEXTS, THEME, TABLE_SETTINGS, BEHAVIOR, RESOURCES, CTK_STYLES
from font_manager import font_manager

# Настройка темы customtkinter
ctk.set_appearance_mode(THEME["appearance_mode"])
ctk.set_default_color_theme("blue")

# Переопределение цветов customtkinter
ctk.ThemeManager.theme["CTkFrame"]["fg_color"] = COLORS["black"]
ctk.ThemeManager.theme["CTkButton"]["fg_color"] = COLORS["red"]
ctk.ThemeManager.theme["CTkButton"]["hover_color"] = COLORS["cyan"]
ctk.ThemeManager.theme["CTkLabel"]["text_color"] = COLORS["white"]


class ModernApp:
    def __init__(self):
        self.date2 = None
        self.current_data = []
        self.active_rows = []
        self.file_path = None
        self.cur7 = None
        self.current_template = 1
        self.current_mode = 1
        self.table2_data = None
        self.checkboxes = []
        self.labels = []
        self.texts = TEXTS["ru"]
        self.is_dark_theme = True
        self.current_theme = "dark"

        # Инициализация менеджера шрифтов
        self.font_manager = font_manager

        # Цвета для строк таблицы в зависимости от темы
        self.table_row_colors = {
            "dark": ["#222222", "#333333"],
            "light": ["#DDDDDD", "#EEEEEE"]
        }

        # Создание главного окна
        self.root = ctk.CTk()
        self.root.title(self.texts["app_title"])
        self.root.geometry(f"{SIZES['window_width']}x{SIZES['window_height']}")
        self.root.minsize(1000, 600)

        # Установка начальной темы
        self.apply_theme()

        # Загрузка и настройка шрифтов
        self.setup_fonts()

        # Настройка интерфейса
        self.setup_ui()

        # Привязка обработчика закрытия окна
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_fonts(self):
        """Настройка кастомных шрифтов"""
        if FONTS["body"][0] and os.path.exists(FONTS["body"][0]):
            self.main_font = self.font_manager.create_ctk_font(
                FONTS["body"][0],
                FONTS["body"][1]
            )
        else:
            self.main_font = ("Arial", FONTS["body"][1])

        if FONTS["title"][0] and os.path.exists(FONTS["title"][0]):
            self.bold_font = self.font_manager.create_ctk_font(
                FONTS["title"][0],
                FONTS["title"][1]
            )
        else:
            self.bold_font = ("Arial", FONTS["title"][1], "bold")

    def apply_theme(self):
        """Применение текущей темы"""
        if self.is_dark_theme:
            ctk.set_appearance_mode("dark")
            self.root.configure(fg_color=COLORS["black"])
            self.current_colors = {
                "bg": COLORS["black"],
                "fg": COLORS["white"],
                "accent": COLORS["red"],
                "hover": COLORS["cyan"],
                "text": COLORS["white"],
                "entry_bg": COLORS["black"],
                "entry_text": COLORS["white"],
                "entry_border": COLORS["white"],
                "table_row1": self.table_row_colors["dark"][0],  # #222222
                "table_row2": self.table_row_colors["dark"][1]  # #333333
            }
        else:
            ctk.set_appearance_mode("light")
            self.root.configure(fg_color=COLORS["white"])
            self.current_colors = {
                "bg": COLORS["white"],
                "fg": COLORS["black"],
                "accent": COLORS["blue"],
                "hover": COLORS["cyan"],
                "text": COLORS["black"],
                "entry_bg": COLORS["white"],
                "entry_text": COLORS["black"],
                "entry_border": COLORS["black"],
                "table_row1": self.table_row_colors["light"][0],  # #DDDDDD
                "table_row2": self.table_row_colors["light"][1]  # #EEEEEE
            }

    def toggle_theme(self):
        """Переключение между темной и светлой темой"""
        self.is_dark_theme = not self.is_dark_theme

        # Сохраняем текущие данные
        current_data_backup = self.current_data
        active_rows_backup = self.active_rows
        current_template_backup = self.current_template
        current_mode_backup = self.current_mode
        file_path_backup = self.file_path

        # Применяем новую тему
        self.apply_theme()

        # Пересоздаем интерфейс
        self.rebuild_ui()

        # Восстанавливаем данные
        self.current_data = current_data_backup
        self.active_rows = active_rows_backup
        self.current_template = current_template_backup
        self.current_mode = current_mode_backup
        self.file_path = file_path_backup

        # Обновляем отображение
        self.update_ui_state()

    def rebuild_ui(self):
        """Перестроение интерфейса при смене темы"""
        for widget in self.root.winfo_children():
            widget.destroy()
        self.setup_ui()

    def update_ui_state(self):
        """Обновление состояния интерфейса после перезагрузки"""
        if hasattr(self, 'btn_theme'):
            if self.is_dark_theme:
                self.btn_theme.configure(text="🌙 Тема")
            else:
                self.btn_theme.configure(text="☀️ Тема")

        if hasattr(self, 'switch_template'):
            self.switch_template(self.current_template)
            self.switch_mode(self.current_mode)

        if self.file_path and hasattr(self, 'load_file_from_path'):
            self.load_file_from_path(self.file_path)

        # Обновляем таблицу с новыми цветами строк
        if self.current_data:
            self.update_table()

    def setup_logo(self):
        """Настройка логотипа в правом верхнем углу"""
        self.logo_frame = ctk.CTkFrame(
            self.root,
            fg_color="transparent",
            height=80
        )
        self.logo_frame.place(relx=1.0, x=-20, y=20, anchor="ne")

        logo_path = RESOURCES["logo"]
        if os.path.exists(logo_path):
            try:
                pil_image = Image.open(logo_path)
                max_size = (150, 80)
                pil_image.thumbnail(max_size, Image.Resampling.LANCZOS)
                self.logo_image = ctk.CTkImage(
                    light_image=pil_image,
                    dark_image=pil_image,
                    size=pil_image.size
                )
                self.logo_label = ctk.CTkLabel(
                    self.logo_frame,
                    image=self.logo_image,
                    text=""
                )
                self.logo_label.pack(expand=True)
            except Exception as e:
                print(f"Ошибка загрузки логотипа: {e}")
                self._show_logo_text()
        else:
            self._show_logo_text()

    def _show_logo_text(self):
        """Отображение текста вместо логотипа"""
        self.logo_label = ctk.CTkLabel(
            self.logo_frame,
            text="Логотип",
            font=self.main_font,
            text_color=self.current_colors["accent"]
        )
        self.logo_label.pack(expand=True)

    def setup_ui(self):
        """Настройка интерфейса"""
        # Основной контейнер
        self.main_container = ctk.CTkFrame(
            self.root,
            fg_color=self.current_colors["bg"],
            corner_radius=0
        )
        self.main_container.pack(fill="both", expand=True, padx=SIZES["padding"]["large"],
                                 pady=SIZES["padding"]["large"])

        # Настройка grid
        self.main_container.grid_rowconfigure(4, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        # Размещаем логотип
        self.setup_logo()

        # Верхняя панель с кнопками
        self.top_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        self.top_frame.grid(row=0, column=0, sticky="ew", pady=(0, SIZES["padding"]["medium"]))
        self.top_frame.grid_columnconfigure(0, weight=1)

        # Левая часть (кнопка выбора файла)
        left_top_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        left_top_frame.grid(row=0, column=0, sticky="w")

        self.btn_select = ctk.CTkButton(
            left_top_frame,
            text=self.texts["select_file"],
            command=self.load_file_from_path,
            fg_color=self.current_colors["accent"],
            hover_color=self.current_colors["hover"],
            font=self.main_font,
            height=40,
            width=200,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_select.pack(side=ctk.LEFT, padx=(0, SIZES["padding"]["medium"]))

        self.lbl_file = ctk.CTkLabel(
            left_top_frame,
            text=self.texts["file_not_selected"],
            font=self.main_font,
            text_color=self.current_colors["text"]
        )
        self.lbl_file.pack(side=ctk.LEFT)

        # Правая часть (кнопки)
        right_top_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        right_top_frame.grid(row=0, column=1, sticky="e")

        # Группа кнопок
        button_frame = ctk.CTkFrame(right_top_frame, fg_color="transparent")
        button_frame.pack(side=ctk.RIGHT, padx=90)

        # Кнопка переключения темы
        theme_text = "🌙 Тема" if self.is_dark_theme else "☀️ Тема"
        self.btn_theme = ctk.CTkButton(
            button_frame,
            text=theme_text,
            command=self.toggle_theme,
            fg_color=self.current_colors["accent"],
            hover_color=self.current_colors["hover"],
            font=self.main_font,
            height=35,
            width=80,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_theme.pack(side=ctk.RIGHT)

        # Кнопка создания документа с отступом справа 40px
        self.btn_create = ctk.CTkButton(
            button_frame,
            text=self.texts["create_doc"],
            command=self.create_doc,
            fg_color=self.current_colors["accent"],
            hover_color=self.current_colors["hover"],
            font=self.bold_font,
            height=40,
            width=200,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_create.pack(side=ctk.RIGHT)  # Отступ справа 40px

        # Панель с шаблонами и режимами
        settings_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        settings_frame.grid(row=1, column=0, sticky="ew", pady=(0, SIZES["padding"]["medium"]))
        settings_frame.grid_columnconfigure(2, weight=1)

        # Группа шаблонов
        template_group = ctk.CTkFrame(settings_frame, fg_color="transparent")
        template_group.grid(row=0, column=0, sticky="w", padx=(0, SIZES["padding"]["large"]))

        ctk.CTkLabel(
            template_group,
            text="Шаблоны:",
            font=self.bold_font,
            text_color=self.current_colors["text"]
        ).pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        self.btn_template1 = ctk.CTkButton(
            template_group,
            text=self.texts["template_1"],
            command=lambda: self.switch_template(1),
            fg_color=self.current_colors["accent"] if self.current_template == 1 else self.current_colors["bg"],
            hover_color=self.current_colors["hover"],
            font=self.main_font,
            width=100,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_template1.pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        self.btn_template2 = ctk.CTkButton(
            template_group,
            text=self.texts["template_2"],
            command=lambda: self.switch_template(2),
            fg_color=self.current_colors["accent"] if self.current_template == 2 else self.current_colors["bg"],
            hover_color=self.current_colors["hover"],
            font=self.main_font,
            width=100,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_template2.pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        # Группа режимов
        mode_group = ctk.CTkFrame(settings_frame, fg_color="transparent")
        mode_group.grid(row=0, column=1, sticky="w", padx=(0, SIZES["padding"]["large"]))

        ctk.CTkLabel(
            mode_group,
            text="Режимы:",
            font=self.bold_font,
            text_color=self.current_colors["text"]
        ).pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        self.btn_mode_table = ctk.CTkButton(
            mode_group,
            text=self.texts["mode_table"],
            command=lambda: self.switch_mode(1),
            fg_color=self.current_colors["accent"] if self.current_mode == 1 else self.current_colors["bg"],
            hover_color=self.current_colors["hover"],
            font=self.main_font,
            width=100,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_mode_table.pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        self.btn_mode_text = ctk.CTkButton(
            mode_group,
            text=self.texts["mode_text"],
            command=lambda: self.switch_mode(2),
            fg_color=self.current_colors["accent"] if self.current_mode == 2 else self.current_colors["bg"],
            hover_color=self.current_colors["hover"],
            font=self.main_font,
            width=100,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_mode_text.pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        # Панель параметров
        param_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        param_frame.grid(row=2, column=0, sticky="ew", pady=(0, SIZES["padding"]["medium"]))

        ctk.CTkLabel(
            param_frame,
            text=self.texts["request_number"],
            font=self.main_font,
            text_color=self.current_colors["text"]
        ).pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        self.entry_num = ctk.CTkEntry(
            param_frame,
            width=120,
            font=self.main_font,
            fg_color=self.current_colors["entry_bg"],
            text_color=self.current_colors["entry_text"],
            border_color=self.current_colors["entry_border"]
        )
        self.entry_num.insert(0, "0")
        self.entry_num.pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        ctk.CTkLabel(
            param_frame,
            text=self.texts["date"],
            font=self.main_font,
            text_color=self.current_colors["text"]
        ).pack(side=ctk.LEFT, padx=(SIZES["padding"]["large"], SIZES["padding"]["small"]))

        self.entry_date = ctk.CTkEntry(
            param_frame,
            width=200,
            font=self.main_font,
            fg_color=self.current_colors["entry_bg"],
            text_color=self.current_colors["entry_text"],
            border_color=self.current_colors["entry_border"]
        )
        entr = f"{str(datetime.date.today().day).rjust(2, '0')}.{str(datetime.date.today().month).rjust(2, '0')}.{datetime.date.today().year} г."
        self.entry_date.insert(0, entr)
        self.entry_date.pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        info_frame = ctk.CTkFrame(param_frame, fg_color="transparent")
        info_frame.pack(side=ctk.RIGHT)

        self.lbl_template = ctk.CTkLabel(
            info_frame,
            text=f"{self.texts['current_template']} {self.current_template}",
            font=self.bold_font,
            text_color=self.current_colors["accent"]
        )
        self.lbl_template.pack(side=ctk.LEFT, padx=SIZES["padding"]["large"])

        self.lbl_mode = ctk.CTkLabel(
            info_frame,
            text=f"{self.texts['current_mode']} {self.texts['mode_table']}",
            font=self.bold_font,
            text_color=self.current_colors["accent"]
        )
        self.lbl_mode.pack(side=ctk.LEFT)

        # Панель управления строками
        ctrl_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        ctrl_frame.grid(row=3, column=0, sticky="ew", pady=(0, SIZES["padding"]["small"]))

        self.lbl_count = ctk.CTkLabel(
            ctrl_frame,
            text=self.texts["records_count"].format(0, 0),
            font=self.bold_font,
            text_color=self.current_colors["text"]
        )
        self.lbl_count.pack(side=ctk.LEFT)

        btn_frame = ctk.CTkFrame(ctrl_frame, fg_color="transparent")
        btn_frame.pack(side=ctk.RIGHT)

        for btn_text, btn_cmd in [
            (self.texts["select_all"], lambda: self.toggle_all(True)),
            (self.texts["select_none"], lambda: self.toggle_all(False)),
            (self.texts["invert"], self.invert_selection)
        ]:
            ctk.CTkButton(
                btn_frame,
                text=btn_text,
                command=btn_cmd,
                fg_color=self.current_colors["accent"],
                hover_color=self.current_colors["hover"],
                font=self.main_font,
                width=100,
                corner_radius=CTK_STYLES["corner_radius"],
                border_width=2,
                border_color=self.current_colors["entry_border"],
                text_color=self.current_colors["text"]
            ).pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        # Таблица
        self.table_container = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        self.table_container.grid(row=4, column=0, sticky="nsew", pady=(0, SIZES["padding"]["medium"]))
        self.table_container.grid_rowconfigure(0, weight=1)
        self.table_container.grid_columnconfigure(0, weight=1)

        # Создание таблицы
        self.create_table()

    def create_table(self):
        """Создание таблицы"""
        if hasattr(self, 'table_frame'):
            self.table_frame.destroy()

        self.table_frame = ctk.CTkScrollableFrame(
            self.table_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        self.table_frame.pack(fill=ctk.BOTH, expand=True)

        headers = ["Акт.", "№п/п", "Кол-во ПУ", "Тип ПУ", "Адрес объекта", "Примечания"]
        col_widths = [
            TABLE_COLUMNS["checkbox"],
            TABLE_COLUMNS["number"],
            TABLE_COLUMNS["quantity"],
            TABLE_COLUMNS["type"],
            TABLE_COLUMNS["address"],
            TABLE_COLUMNS["notes"]
        ]

        header_frame = ctk.CTkFrame(
            self.table_frame,
            fg_color=TABLE_SETTINGS["header_bg"],
            height=TABLE_SETTINGS["header_height"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        header_frame.pack(fill=ctk.X, pady=(0, 2))

        for i, (h, w) in enumerate(zip(headers, col_widths)):
            label = ctk.CTkLabel(
                header_frame,
                text=h,
                font=self.bold_font,
                width=w,
                anchor="w",
                text_color=COLORS["white"]
            )
            label.grid(row=0, column=i, padx=5, pady=5, sticky="w")

        self.data_frame = ctk.CTkFrame(self.table_frame, fg_color="transparent")
        self.data_frame.pack(fill=ctk.BOTH, expand=True)

    def update_table(self):
        """Обновление таблицы с использованием цветов текущей темы"""
        # Очищаем таблицу
        for widget in self.data_frame.winfo_children():
            widget.destroy()

        self.checkboxes = []
        self.active_rows = list(range(len(self.current_data)))

        col_widths = [
            TABLE_COLUMNS["checkbox"],
            TABLE_COLUMNS["number"],
            TABLE_COLUMNS["quantity"],
            TABLE_COLUMNS["type"],
            TABLE_COLUMNS["address"],
            TABLE_COLUMNS["notes"]
        ]

        for row_idx, row_data in enumerate(self.current_data):
            # Определяем цвет фона строки в зависимости от темы
            if TABLE_SETTINGS["alternating_row_colors"]:
                # Используем цвета из current_colors в зависимости от четности строки
                bg_color = self.current_colors["table_row1"] if row_idx % 2 == 0 else self.current_colors["table_row2"]
            else:
                bg_color = "transparent"

            row_frame = ctk.CTkFrame(self.data_frame, fg_color=bg_color, corner_radius=0)
            row_frame.pack(fill=ctk.X)

            # Чекбокс
            var = ctk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(
                row_frame,
                variable=var,
                text="",
                width=col_widths[0],
                onvalue=True,
                offvalue=False,
                command=lambda r=row_idx, v=var: self.on_check(r, v),
                fg_color=COLORS["blue"],
                hover_color=COLORS["cyan"],
                border_color=self.current_colors["entry_border"],
                checkmark_color=self.current_colors["text"]
            )
            cb.grid(row=0, column=0, padx=2, pady=2)
            self.checkboxes.append((cb, var))

            # Данные
            for col_idx, value in enumerate(row_data):
                if col_idx < 5:
                    text = str(value)
                    lbl = ctk.CTkLabel(
                        row_frame,
                        text=text,
                        font=self.main_font,
                        width=col_widths[col_idx + 1],
                        anchor="w",
                        wraplength=col_widths[col_idx + 1] - 10,
                        text_color=self.current_colors["text"]
                    )
                    lbl.grid(row=0, column=col_idx + 1, padx=5, pady=2, sticky="w")

        self.lbl_count.configure(text=self.texts["records_count"].format(
            len(self.current_data), len(self.active_rows)
        ))

    def on_check(self, row, var):
        """Обработчик изменения состояния чекбокса"""
        if var.get():
            if row not in self.active_rows:
                self.active_rows.append(row)
        else:
            if row in self.active_rows:
                self.active_rows.remove(row)

        self.active_rows.sort()
        self.lbl_count.configure(text=self.texts["records_count"].format(
            len(self.current_data), len(self.active_rows)
        ))

    def switch_template(self, template_num):
        """Переключение шаблона"""
        self.current_template = template_num
        self.lbl_template.configure(text=f"{self.texts['current_template']} {template_num}")

        # Обновление стиля кнопок
        if template_num == 1:
            self.btn_template1.configure(fg_color=self.current_colors["accent"])
            self.btn_template2.configure(fg_color=self.current_colors["bg"])
        else:
            self.btn_template1.configure(fg_color=self.current_colors["bg"])
            self.btn_template2.configure(fg_color=self.current_colors["accent"])

    def switch_mode(self, mode_num):
        """Переключение режима"""
        self.current_mode = mode_num
        mode_text = self.texts["mode_table"] if mode_num == 1 else self.texts["mode_text"]
        self.lbl_mode.configure(text=f"{self.texts['current_mode']} {mode_text}")

        # Обновление стиля кнопок
        if mode_num == 1:
            self.btn_mode_table.configure(fg_color=self.current_colors["accent"])
            self.btn_mode_text.configure(fg_color=self.current_colors["bg"])
        else:
            self.btn_mode_table.configure(fg_color=self.current_colors["bg"])
            self.btn_mode_text.configure(fg_color=self.current_colors["accent"])

    def toggle_all(self, state):
        """Выбрать все/снять выделение"""
        if state:
            self.active_rows = list(range(len(self.current_data)))
        else:
            self.active_rows = []

        for cb, var in self.checkboxes:
            if var.get() != state:
                var.set(state)

        self.lbl_count.configure(text=self.texts["records_count"].format(
            len(self.current_data), len(self.active_rows)
        ))

    def invert_selection(self):
        """Инвертировать выделение"""
        for cb, var in self.checkboxes:
            var.set(not var.get())

        self.active_rows = [i for i, (_, var) in enumerate(self.checkboxes) if var.get()]
        self.lbl_count.configure(text=self.texts["records_count"].format(
            len(self.current_data), len(self.active_rows)
        ))

    def load_file_from_path(self, file_path=None):
        """Загрузка файла"""
        if not file_path:
            file_path = filedialog.askopenfilename(
                title="Выберите Word документ",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
            )

        if not file_path:
            return

        try:
            self.file_path = file_path
            sys_name = os.path.basename(file_path)
            for_name = sys_name if len(sys_name) < 30 else sys_name[:27] + '...'
            self.lbl_file.configure(text=for_name, text_color=self.current_colors["accent"])

            # Обработка файла
            doc = Document(file_path)
            info = find_info_in_doc(doc, file_path)

            self.cur7 = info.get("дата выполнения")
            self.date2 = info.get("дата обращения")
            self.table2_data = info.get("таблица работ")

            if info.get("номер заявки"):
                self.entry_num.delete(0, ctk.END)
                self.entry_num.insert(0, info["номер заявки"])

            self.switch_mode(info.get("режим", 1))

            # Обновление таблицы
            self.current_data = []
            for row in info.get("адреса объектов", []):
                if row and row[0] != '№п/п':
                    self.current_data.append(row[:5])

            self.update_table()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror(self.texts["error"], self.texts["load_error"].format(str(e)))

    def create_doc(self):
        """Создание документа"""
        if not self.current_data:
            messagebox.showwarning(self.texts["warning"], self.texts["no_data"])
            return

        if not self.active_rows:
            messagebox.showwarning(self.texts["warning"], self.texts["no_active"])
            return

        file_path = filedialog.asksaveasfilename(
            title="Сохранить документ",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")]
        )

        if not file_path:
            return

        try:
            active_data = []
            for i in self.active_rows:
                row = self.current_data[i]
                if not row or all(cell == '' for cell in row):
                    continue

                normalized_row = list(row)[:5]
                while len(normalized_row) < 5:
                    normalized_row.append("")
                active_data.append(normalized_row)

            if not active_data:
                messagebox.showwarning(self.texts["warning"], self.texts["no_data"])
                return

            num = self.entry_num.get() or "ERROR_IN_NUM"
            date = self.entry_date.get() or "ERROR_IN_DATE"
            table1_rows = len(active_data) + 1

            if self.current_template == 2:
                output = create_application_doc2(
                    num=num,
                    date1=date,
                    date2=self.date2,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    table2_data=self.table2_data,
                    mode=self.current_mode
                )
                template_name = "Шаблон 2"
            else:
                output = create_application_doc(
                    num=num,
                    date=self.date2,
                    table1_rows=table1_rows,
                    table1_data=active_data,
                    table2_data=self.table2_data,
                    date_work=self.cur7,
                    mode=self.current_mode
                )
                template_name = "Шаблон 1"

            import shutil
            if output and os.path.exists(output):
                shutil.move(output, file_path)

            messagebox.showinfo(
                self.texts["red"],
                self.texts["doc_created"].format(template_name, len(active_data), file_path)
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror(self.texts["error"], self.texts["create_error"].format(str(e)))

    def on_closing(self):
        """Обработчик закрытия окна"""
        self.font_manager.cleanup_fonts()
        self.root.destroy()

    def run(self):
        """Запуск приложения"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ModernApp()
    app.run()