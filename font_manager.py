# font_manager.py - Управление шрифтами
import tkinter as tk
from tkinter import font as tkfont
from pathlib import Path
import os
import sys
from typing import Union, Tuple, Optional


def resource_path(relative_path: str) -> str:
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class FontManager:
    """Менеджер для загрузки и управления шрифтами"""

    def __init__(self):
        self.fonts_cache = {}
        self.custom_fonts = []
        self.fonts_dir = Path(resource_path("fonts"))

    def load_custom_font(self, font_path: Union[str, Path]) -> Optional[str]:
        """Загрузка кастомного шрифта из TTF файла"""
        try:
            # Проверяем существует ли файл
            font_path_obj = Path(font_path)
            if not font_path or not font_path_obj.exists():
                return None

            # Для customtkinter и tkinter используем стандартный механизм
            # регистрируем шрифт в системе (только для Windows)
            if os.name == 'nt':  # Windows
                import ctypes
                ctypes.windll.gdi32.AddFontResourceW(str(font_path_obj))

            # Создаем объект шрифта tkinter
            font_name = font_path_obj.stem
            custom_font = tkfont.Font(family=font_name, size=10)
            self.custom_fonts.append(str(font_path_obj))  # Для последующей очистки
            self.fonts_cache[str(font_path_obj)] = font_name
            return font_name
        except Exception as e:
            print(f"Ошибка загрузки шрифта {font_path}: {e}")
            return None

    def get_font(self, font_path: Union[str, Path], size: int, weight: str = "normal", slant: str = "roman") -> Tuple[str, int, str, str]:
        """Получение шрифта с заданными параметрами"""
        try:
            # Если передан относительный путь, ищем в папке fonts
            font_path_obj = Path(font_path)
            if font_path and not font_path_obj.is_absolute():
                font_path_obj = self.fonts_dir / font_path

            if font_path_obj.exists():
                font_name = self.load_custom_font(str(font_path_obj))
                if font_name:
                    return (font_name, size, weight, slant)

            # Возвращаем стандартный шрифт если кастомный не загрузился
            return self.get_default_font(size, weight, slant)
        except:
            return self.get_default_font(size, weight, slant)

    def get_default_font(self, size: int, weight: str = "normal", slant: str = "roman") -> Tuple[str, int, str, str]:
        """Получение стандартного шрифта"""
        default_fonts = [
            "Comic Sans"
        ]
        return (default_fonts[0], size, weight, slant)

    def create_ctk_font(self, font_path: Union[str, Path], size: int) -> Tuple[str, int]:
        """Создание шрифта для customtkinter"""
        try:
            # Если передан относительный путь, ищем в папке fonts
            font_path_obj = Path(font_path)
            if font_path and not font_path_obj.is_absolute():
                font_path_obj = self.fonts_dir / font_path

            if font_path_obj.exists():
                font_name = self.load_custom_font(str(font_path_obj))
                if font_name:
                    return (font_name, size)
            return ("Roboto", size)
        except:
            return ("Roboto", size)

    def cleanup_fonts(self) -> None:
        """Очистка загруженных шрифтов (для Windows)"""
        if os.name == 'nt':
            import ctypes
            for font_path in self.custom_fonts:
                try:
                    ctypes.windll.gdi32.RemoveFontResourceW(str(font_path))
                except:
                    pass
        self.custom_fonts.clear()
        self.fonts_cache.clear()


# Глобальный экземпляр менеджера шрифтов
font_manager = FontManager()
