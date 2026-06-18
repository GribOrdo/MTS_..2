# config.py - Настройки графического интерфейса
import os
from pathlib import Path

# Цветовая схема (в соответствии с требованиями)
COLORS = {
    "red": "#FF0032",      # Основной красный
    "cyan": "#74BAFF",       # Акцентный светло-синий
    "black": "#1e1f22",         # Темно-серый для фона
    "white": "#FFFFFF",        # Светлый для текста на темном фоне
    "blue": "#1F87FF",      # Успех (синий)
    "gray": "#3F3F3F",         # Серый для неактивных элементов
    "table_row_even": "#888888",  # Четные строки таблицы
    "table_row_odd": "#999999",   # Нечетные строки таблицы
}

# Тема оформления
THEME = {
    "appearance_mode": "dark",     # Принудительно темная тема
    "color_theme": "custom"        # Использование кастомных цветов
}

# Путь к файлу шрифта
FONT_PATH = Path(__file__).parent / "fonts" / "MTSCompact-Medium.ttf"
FONT_BOLD_PATH = Path(__file__).parent / "fonts" / "MTSCompact-Bold.ttf"
FONT_ITALIC_PATH = Path(__file__).parent / "fonts" / "MTSCompact-Black.ttf"

# Если файлы шрифтов не найдены, используем стандартные
if not FONT_PATH.exists():
    FONT_PATH = None
    FONT_BOLD_PATH = None
    FONT_ITALIC_PATH = None

# Настройки шрифтов
FONTS = {
    "title": (FONT_BOLD_PATH, 18),
    "subtitle": (FONT_BOLD_PATH, 14),
    "body": (FONT_PATH, 14),
    "small": (FONT_PATH, 12),
    "button": (FONT_BOLD_PATH, 13),
    "table_header": (FONT_BOLD_PATH, 14),
    "table_cell": (FONT_PATH, 12)
}

# Размеры и геометрия
SIZES = {
    "window_width": 1280,
    "window_height": 720,
    "logo_size": 80,  # Размер логотипа
    "padding": {
        "small": 5,
        "medium": 10,
        "large": 20,
        "xlarge": 30
    }
}

# Ширина колонок таблицы (в пикселях)
TABLE_COLUMNS = {
    "checkbox": 100,      # Чекбокс
    "number": 70,        # №п/п
    "quantity": 100,     # Кол-во ПУ
    "type": 150,         # Тип ПУ
    "address": 420,      # Адрес объекта
    "notes": 380         # Примечания
}

# Настройки прокрутки
SCROLL_SETTINGS = {
    "mousewheel_speed": 4,
    "scrollbar_width": 12
}

# Тексты для интерфейса (поддержка локализации)
TEXTS = {
    "ru": {
        "app_title": "Обработчик заявок",
        "select_file": "Выбрать Word файл",
        "file_not_selected": "Файл не выбран",
        "template_1": "Шаблон 1",
        "template_2": "Шаблон 2",
        "mode_table": "Таблица",
        "mode_text": "Текст",
        "create_doc": "Создать документ",
        "request_number": "Номер заявки:",
        "date": "Дата:",
        "current_template": "Текущий шаблон:",
        "current_mode": "Текущий режим:",
        "records_count": "Записей: {} | Активных: {}",
        "select_all": "✅ Все",
        "select_none": "❌ Ничего",
        "invert": "🔄 Инверт.",
        "drag_drop": "📁 Перетащите файл Word сюда",
        "warning": "Предупреждение",
        "error": "Ошибка",
        "success": "Успех",
        "no_data": "Нет данных для создания документа",
        "no_active": "Нет активных записей. Отметьте хотя бы одну строку.",
        "doc_created": "Документ создан!\n\n{}\nАктивных записей: {}\nСохранен в:\n{}",
        "load_error": "Не удалось загрузить файл:\n{}",
        "create_error": "Не удалось создать документ:\n{}"
    }
}

# Настройки отображения таблицы
TABLE_SETTINGS = {
    "header_height": 40,
    "row_height": 35,
    "alternating_row_colors": True,
    "alternating_colors": [COLORS["table_row_even"], COLORS["table_row_odd"]],
    "header_bg": COLORS["blue"],
    "border_width": 1,
    "border_color": COLORS["white"]
}

# Поведение приложения
BEHAVIOR = {
    "auto_load_last_file": False,
    "confirm_on_exit": False,
    "max_recent_files": 5,
    "auto_save_selection": True
}

# Пути к ресурсам
RESOURCES = {
    "logo": "MTC_Logo_CMYK.png",
    "icon": "app_icon.ico",
    "temp_dir": "temp",
    "fonts_dir": "fonts"
}

# Стили для customtkinter
CTK_STYLES = {
    "corner_radius": 8,
    "border_width": 1,
    "border_color": COLORS["white"],
    "fg_color": COLORS["black"],
    "text_color": COLORS["white"]
}
