# font_manager.py - Управление шрифтами
import tkinter as tk
from tkinter import font as tkfont
from pathlib import Path
import os


class FontManager:
    """Менеджер для загрузки и управления шрифтами"""

    def __init__(self):
        self.fonts_cache = {}
        self.custom_fonts = []

    def load_custom_font(self, font_path):
        """Загрузка кастомного шрифта из TTF файла"""
        try:
            # Проверяем существует ли файл
            if not font_path or not Path(font_path).exists():
                return None

            # Для customtkinter и tkinter используем стандартный механизм
            # Регистрируем шрифт в системе (только для Windows)
            if os.name == 'nt':  # Windows
                import ctypes
                ctypes.windll.gdi32.AddFontResourceW(str(font_path))

            # Создаем объект шрифта tkinter
            font_name = Path(font_path).stem
            custom_font = tkfont.Font(family=font_name, size=10)
            self.custom_fonts.append(font_path)  # Для последующей очистки
            self.fonts_cache[font_path] = font_name
            return font_name
        except Exception as e:
            print(f"Ошибка загрузки шрифта {font_path}: {e}")
            return None

    def get_font(self, font_path, size, weight="normal", slant="roman"):
        """Получение шрифта с заданными параметрами"""
        try:
            if font_path and Path(font_path).exists():
                font_name = self.load_custom_font(font_path)
                if font_name:
                    return (font_name, size, weight, slant)

            # Возвращаем стандартный шрифт если кастомный не загрузился
            return self.get_default_font(size, weight, slant)
        except:
            return self.get_default_font(size, weight, slant)

    def get_default_font(self, size, weight="normal", slant="roman"):
        """Получение стандартного шрифта"""
        default_fonts = [
            "Comic Sans"
        ]
        return (default_fonts[0], size, weight, slant)

    def create_ctk_font(self, font_path, size):
        """Создание шрифта для customtkinter"""
        try:
            if font_path and Path(font_path).exists():
                font_name = self.load_custom_font(font_path)
                if font_name:
                    return (font_name, size)
            return ("Roboto", size)
        except:
            return ("Roboto", size)

    def cleanup_fonts(self):
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