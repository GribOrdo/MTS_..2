
try:
    from CTkScrollableDropdown import *
except:
    pass
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
import sys
from tkinterdnd2 import DND_FILES, TkinterDnD

import tkinter as tk   # Import tkinter for TclError

# Setup customtkinter theme
ctk.set_appearance_mode(THEME["appearance_mode"])
ctk.set_default_color_theme("blue")

# Override customtkinter colors
ctk.ThemeManager.theme["CTkFrame"]["fg_color"] = COLORS["black"]
ctk.ThemeManager.theme["CTkButton"]["fg_color"] = COLORS["red"]
ctk.ThemeManager.theme["CTkButton"]["hover_color"] = COLORS["cyan"]
ctk.ThemeManager.theme["CTkLabel"]["text_color"] = COLORS["white"]

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class ModernApp:
    def __init__(self):
        # Initialize attributes to prevent NoneType errors
        self.btn_frame = None
        self.lbl_count = None
        self.ctrl_frame = None
        self.lbl_mode = None
        self.lbl_template = None
        self.entry_date = None
        self.entry_num = None
        self.btn_mode_text = None
        self.btn_mode_table = None
        self.btn_template2 = None
        self.btn_template1 = None
        self.btn_create = None
        self.btn_theme = None
        self.lbl_file = None
        self.btn_select = None
        self.top_frame = None
        self.main_container = None
        self.logo_label = None
        self.logo_image = None
        self.logo_frame = None
        self.current_colors = None
        self.bold_font = None
        self.main_font = None
        self.drag_label = None
        self.drag_overlay = None
        self._resize_job = None  # For debouncing resize events

        self.date2 = None
        self.current_data = []
        self.active_rows = []
        self.file_path = None
        self.cur7 = None
        self.current_template = 1
        self.current_mode = 1  # 1 = table mode, 0 = text mode
        self.table2_data = None
        self.checkboxes = []
        self.labels = []
        self.texts = TEXTS["ru"]
        self.is_dark_theme = True
        self.drag_hover = False

        # New attributes for text mode
        self.aggregated_data = []  # Aggregated data by equipment types
        self.quantity_comboboxes = []  # Dropdown widgets

        # Initialize font manager
        self.font_manager = font_manager

        # Row colors for table based on theme
        self.table_row_colors = {
            "dark": ["#222222", "#333333"],
            "light": ["#DDDDDD", "#EEEEEE"]
        }

        self.root = TkinterDnD.Tk()

        self.root.title(self.texts["app_title"])
        self.root.geometry(f"{SIZES['window_width']}x{SIZES['window_height']}")
        self.root.minsize(1000, 600)

        # --- CHANGED: Do not configure root bg here ---
        # self.root.configure(bg=COLORS["black"])
        # --- ---
        self.root.bind('<Configure>', self._on_window_resize)
        # Setup drag-n-drop for entire window (if available)
        self.setup_drag_and_drop()
        # Apply initial theme
        self.apply_theme()
        # Load and setup fonts
        self.setup_fonts()
        # Setup UI
        self.setup_ui()
        # Bind window close handler
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def _on_window_resize(self, event):
        """Handle window resize with debouncing to prevent lag"""
        if event.widget == self.root:
            if self._resize_job is not None:
                self.root.after_cancel(self._resize_job)
            self._resize_job = self.root.after(100, self._do_resize)

    def _do_resize(self):
        """Actually perform resize operations"""
        self._resize_job = None
        # Update table layout if needed
        if hasattr(self, 'table_frame'):
            self.root.update_idletasks()

    def setup_drag_and_drop(self):
        """Setup drag-n-drop functionality"""
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind('<<Drop>>', self.on_drop)
        self.root.dnd_bind('<<DragEnter>>', self.on_drag_enter)
        self.root.dnd_bind('<<DragLeave>>', self.on_drag_leave)

    def on_drag_enter(self, event):
        """Handler for cursor entering window with file"""
        if not self.drag_hover:
            self.drag_hover = True
            if not self.drag_overlay:
                self.drag_overlay = ctk.CTkFrame(
                    self.root,
                    fg_color="#000000",
                    corner_radius=0
                )
            self.drag_overlay.place(relx=0, rely=0, relwidth=1, relheight=1)

            if not self.drag_label:
                self.drag_label = ctk.CTkLabel(
                    self.drag_overlay,
                    text="📁 Файл сюда\n(.docx)",
                    font=("Arial", 24, "bold"),
                    text_color=COLORS["white"],
                    fg_color="transparent"
                )
                self.drag_label.pack(expand=True)

            self.drag_overlay.configure(fg_color="#000000CC")

    def on_drag_leave(self, event):
        """Handler for cursor leaving window"""
        self.drag_hover = False
        if self.drag_overlay:
            self.drag_overlay.place_forget()

    def on_drop(self, event):
        """Handler for dropping file in window"""
        self.drag_hover = False
        if self.drag_overlay:
            self.drag_overlay.place_forget()

        file_path = event.data

        if file_path.startswith('{') and file_path.endswith('}'):
            file_path = file_path[1:-1]

        file_path = file_path.replace('\\', '')

        if not file_path.lower().endswith('.docx'):
            messagebox.showwarning(
                self.texts["warning"],
                "Please drop a file with .docx extension"
            )
            return

        self.load_file_from_path(file_path)

    def setup_fonts(self):
        """Setup custom fonts"""
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
        """Apply current theme"""
        if self.is_dark_theme:
            ctk.set_appearance_mode("dark")
            # --- CHANGED: Remove root configure bg ---
            # try:
            #     self.root.configure(bg=COLORS["black"])
            # except:
            #     pass
            # --- ---
            self.current_colors = {
                "bg": COLORS["black"],
                "fg": COLORS["white"],
                "accent": COLORS["red"],
                "hover": COLORS["cyan"],
                "text": COLORS["white"],
                "entry_bg": COLORS["black"],
                "entry_text": COLORS["white"],
                "entry_border": COLORS["white"],
                "table_row1": self.table_row_colors["dark"][0],
                "table_row2": self.table_row_colors["dark"][1]
            }
        else:
            ctk.set_appearance_mode("light")
            # --- CHANGED: Remove root configure bg ---
            # try:
            #     self.root.configure(bg=COLORS["white"])
            # except:
            #     pass
            # --- ---
            self.current_colors = {
                "bg": COLORS["white"],
                "fg": COLORS["black"],
                "accent": COLORS["blue"],
                "hover": COLORS["cyan"],
                "text": COLORS["black"],
                "entry_bg": COLORS["white"],
                "entry_text": COLORS["black"],
                "entry_border": COLORS["black"],
                "table_row1": self.table_row_colors["light"][0],
                "table_row2": self.table_row_colors["light"][1]
            }

    def toggle_theme(self):
        """Toggle between dark and light theme"""
        self.is_dark_theme = not self.is_dark_theme

        # Backup current state
        current_data_backup = self.current_data
        active_rows_backup = self.active_rows.copy()
        current_template_backup = self.current_template
        current_mode_backup = self.current_mode
        file_path_backup = self.file_path
        aggregated_data_backup = self.aggregated_data.copy()

        # Apply new theme
        self.apply_theme()

        # Rebuild UI with new theme
        self.rebuild_ui()

        # Restore state
        self.current_data = current_data_backup
        self.active_rows = active_rows_backup
        self.current_template = current_template_backup
        self.current_mode = current_mode_backup
        self.file_path = file_path_backup
        self.aggregated_data = aggregated_data_backup

        # Update UI state
        self.update_ui_state()

    def rebuild_ui(self):
        """Rebuild UI when theme changes"""
        # Clear all widgets properly
        if self.drag_overlay:
            try:
                self.drag_overlay.destroy()
            except tk.TclError:
                pass  # Already destroyed or invalid
            self.drag_overlay = None
        if self.drag_label:
            try:
                self.drag_label.destroy()
            except tk.TclError:
                pass  # Already destroyed or invalid
            self.drag_label = None
        self.drag_hover = False  # Reset drag state

        # --- Clear all widgets properly ---
        children = list(self.root.winfo_children())
        for widget in children:
            try:
                widget.destroy()
            except tk.TclError:
                pass

        # Reset widget references
        self.btn_frame = None
        self.lbl_count = None
        self.ctrl_frame = None
        self.lbl_mode = None
        self.lbl_template = None
        self.entry_date = None
        self.entry_num = None
        self.btn_mode_text = None
        self.btn_mode_table = None
        self.btn_template2 = None
        self.btn_template1 = None
        self.btn_create = None
        self.btn_theme = None
        self.lbl_file = None
        self.btn_select = None
        self.top_frame = None
        self.main_container = None
        self.logo_label = None
        self.logo_image = None
        self.logo_frame = None
        self.drag_label = None
        self.drag_overlay = None

        # Reset lists
        self.checkboxes = []
        self.quantity_comboboxes = []
        self.labels = []

        self.setup_ui()

    def update_ui_state(self):
        """Update UI state after rebuild"""
        if hasattr(self, 'btn_theme'):
            theme_text = "🌙 Ночь" if self.is_dark_theme else "☀️ День"
            self.btn_theme.configure(text=theme_text)

        if hasattr(self, 'switch_template'):
            self.switch_template(self.current_template, update_ui=False)
            self.switch_mode(self.current_mode, update_ui=False)

        if self.file_path:
            self.load_file_from_path(self.file_path, update_ui=False)

        if self.current_data or self.aggregated_data:
            self.update_table()

    def setup_logo(self):
        """Setup logo in top right corner"""
        if self.logo_frame:
            self.logo_frame.destroy()

        self.logo_frame = ctk.CTkFrame(
            self.root,
            fg_color="transparent",
            height=80
        )
        self.logo_frame.place(relx=1.0, x=0, y=0, anchor="ne")
        logo_path = resource_path(RESOURCES["logo"])

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
                if self.logo_label:
                    self.logo_label.destroy()
                self.logo_label = ctk.CTkLabel(
                    self.logo_frame,
                    image=self.logo_image,
                    text=""
                )
                self.logo_label.pack(expand=True)
            except Exception as e:
                print(f"Error loading logo: {e}")
                self._show_logo_text()
        else:
            self._show_logo_text()

    def _show_logo_text(self):
        """Show text instead of logo"""
        if self.logo_label:
            self.logo_label.destroy()
        self.logo_label = ctk.CTkLabel(
            self.logo_frame,
            text="Logo",
            font=self.main_font,
            text_color=self.current_colors["accent"]
        )
        self.logo_label.pack(expand=True)

    def setup_ui(self):
        """Setup UI"""
        # Clear existing container if it exists
        if hasattr(self, 'main_container') and self.main_container:
            self.main_container.destroy()

        self.main_container = ctk.CTkFrame(
            self.root,
            fg_color=self.current_colors["bg"],
            corner_radius=0
        )
        # --- CHANGED: Set padx and pady to 0 ---
        # self.main_container.pack(fill="both", expand=True, padx=SIZES["padding"]["large"], pady=SIZES["padding"]["large"])
        self.main_container.pack(fill="both", expand=True, padx=0, pady=0)
        # --- ---

        self.main_container.grid_rowconfigure(4, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        self.setup_logo()

        # Top panel with buttons
        self.top_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        self.top_frame.grid(row=0, column=0, sticky="ew", pady=(0, SIZES["padding"]["medium"]))
        self.top_frame.grid_columnconfigure(0, weight=1)

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

        right_top_frame = ctk.CTkFrame(self.top_frame, fg_color="transparent")
        right_top_frame.grid(row=0, column=1, sticky="e")

        button_frame = ctk.CTkFrame(right_top_frame, fg_color="transparent")
        button_frame.pack(side=ctk.RIGHT, padx=90)

        theme_text = "🌙 Ночь" if self.is_dark_theme else "☀️ День"
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
        self.btn_create.pack(side=ctk.RIGHT)

        # Template and mode panel
        settings_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        settings_frame.grid(row=1, column=0, sticky="ew", pady=(0, SIZES["padding"]["medium"]))
        settings_frame.grid_columnconfigure(2, weight=1)

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
            command=lambda: self.switch_mode(0),
            fg_color=self.current_colors["accent"] if self.current_mode == 0 else self.current_colors["bg"],
            hover_color=self.current_colors["hover"],
            font=self.main_font,
            width=100,
            corner_radius=CTK_STYLES["corner_radius"],
            border_width=2,
            border_color=self.current_colors["entry_border"],
            text_color=self.current_colors["text"]
        )
        self.btn_mode_text.pack(side=ctk.LEFT, padx=SIZES["padding"]["small"])

        # Parameters panel
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
            text=f"{self.texts['current_mode']} {self.texts['mode_table'] if self.current_mode == 1 else self.texts['mode_text']}",
            font=self.bold_font,
            text_color=self.current_colors["accent"]
        )
        self.lbl_mode.pack(side=ctk.LEFT)

        # Row control panel (only for table mode)
        self.ctrl_frame = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        self.ctrl_frame.grid(row=3, column=0, sticky="ew", pady=(0, SIZES["padding"]["small"]))

        self.lbl_count = ctk.CTkLabel(
            self.ctrl_frame,
            text=self.texts["records_count"].format(0, 0),
            font=self.bold_font,
            text_color=self.current_colors["text"]
        )
        self.lbl_count.pack(side=ctk.LEFT)

        self.btn_frame = ctk.CTkFrame(self.ctrl_frame, fg_color="transparent")
        self.btn_frame.pack(side=ctk.RIGHT)

        for btn_text, btn_cmd in [
            (self.texts["select_all"], lambda: self.toggle_all(True)),
            (self.texts["select_none"], lambda: self.toggle_all(False)),
            (self.texts["invert"], self.invert_selection)
        ]:
            ctk.CTkButton(
                self.btn_frame,
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

        # Table
        if hasattr(self, 'table_container') and self.table_container:
            self.table_container.destroy()

        self.table_container = ctk.CTkFrame(
            self.main_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        self.table_container.grid(row=4, column=0, sticky="nsew", pady=(0, SIZES["padding"]["medium"]))
        self.table_container.grid_rowconfigure(0, weight=1)
        self.table_container.grid_columnconfigure(0, weight=1)

        self.create_table()

    def aggregate_data(self, data):
        """Aggregate data by equipment types and addresses"""
        aggregated = {}

        for row in data:
            if len(row) >= 3:
                # row format: [number, quantity, type, address, notes]
                equipment_type = str(row[2]) if row[2] else "Unknown"
                quantity = int(row[1]) if row[1] and str(row[1]).isdigit() else 1
                address = str(row[3]) if len(row) > 3 and row[3] else "Unknown"
                notes = str(row[4]) if len(row) > 4 and row[4] else ""

                # Создаем уникальный ключ для комбинации тип + адрес
                unique_key = f"{equipment_type}||{address}"

                if unique_key not in aggregated:
                    aggregated[unique_key] = {
                        "type": equipment_type,
                        "total_quantity": 0,
                        "max_quantity": 0,
                        "addresses": [],
                        "notes": [],
                        "address": address
                    }

                aggregated[unique_key]["total_quantity"] += quantity
                aggregated[unique_key]["max_quantity"] = max(
                    aggregated[unique_key]["max_quantity"],
                    quantity
                )
                aggregated[unique_key]["addresses"].append({
                    "address": address,
                    "quantity": quantity
                })
                if notes:
                    aggregated[unique_key]["notes"].append(notes)

        return list(aggregated.values())

    def create_table(self):
        """Create table"""
        if hasattr(self, 'table_frame') and self.table_frame:
            self.table_frame.destroy()

        self.table_frame = ctk.CTkScrollableFrame(
            self.table_container,
            fg_color=self.current_colors["bg"],
            corner_radius=CTK_STYLES["corner_radius"]
        )
        self.table_frame.pack(fill=ctk.BOTH, expand=True)

        # Define headers based on mode
        if self.current_mode == 0:  # Text mode
            headers = ["Выб.", "№П/П", "Кол-во ПУ", "Тип ПУ", "Адрес", "Инфо."]
            col_widths = [
                TABLE_COLUMNS["checkbox"],
                TABLE_COLUMNS["number"],
                TABLE_COLUMNS["quantity"],
                TABLE_COLUMNS["type"],
                TABLE_COLUMNS["address"],
                TABLE_COLUMNS["notes"]
            ]
        else:  # Table mode
            headers = ["Акт.", "№П/П", "Кол-во ПУ", "Тип ПУ", "Адрес", "Инфо."]
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

        # Configure header columns to be centered
        for i in range(len(headers)):
            header_frame.grid_columnconfigure(i, weight=1)

        for i, (h, w) in enumerate(zip(headers, col_widths)):
            label = ctk.CTkLabel(
                header_frame,
                text=h,
                font=self.bold_font,
                width=w,
                anchor="center",
                text_color=COLORS["white"]
            )
            label.grid(row=0, column=i, padx=5, pady=5, sticky="")

        self.data_frame = ctk.CTkFrame(self.table_frame, fg_color="transparent")
        self.data_frame.pack(fill=ctk.BOTH, expand=True)

    def update_table(self):
        """Update table with current theme colors"""
        # Clear existing widgets properly
        if hasattr(self, 'data_frame') and self.data_frame:
            for widget in self.data_frame.winfo_children():
                widget.destroy()

        self.checkboxes = []
        self.quantity_comboboxes = []

        if self.current_mode == 0:
            self.update_table_text_mode()
            # Hide control panel for text mode
            if self.ctrl_frame:
                self.ctrl_frame.grid_remove()
        else:
            self.update_table_table_mode()
            # Show control panel for table mode
            if self.ctrl_frame:
                self.ctrl_frame.grid()

    def update_table_text_mode(self):
        """Update table in 'text' mode - dropdowns with quantity selection"""
        if not self.aggregated_data:
            empty_label = ctk.CTkLabel(
                self.data_frame,
                text="No data to display. Load a file.",
                font=self.main_font,
                text_color=self.current_colors["text"]
            )
            empty_label.pack(pady=50)
            return

        col_widths = [
            TABLE_COLUMNS["checkbox"],
            TABLE_COLUMNS["number"],
            TABLE_COLUMNS["quantity"],
            TABLE_COLUMNS["type"],
            TABLE_COLUMNS["address"],
            TABLE_COLUMNS["notes"]
        ]

        # Counter for sequential numbering
        counter = 1

        for row_idx, eq_data in enumerate(self.aggregated_data):
            # Determine row background color
            if TABLE_SETTINGS["alternating_row_colors"]:
                bg_color = self.current_colors["table_row1"] if row_idx % 2 == 0 else self.current_colors["table_row2"]
            else:
                bg_color = "transparent"

            row_frame = ctk.CTkFrame(self.data_frame, fg_color=bg_color, corner_radius=0)
            row_frame.pack(fill=ctk.X)

            # Configure grid columns to be centered
            for i in range(6):
                row_frame.grid_columnconfigure(i, weight=1)

            # Dropdown for quantity selection (from 0 to max_quantity)
            # Добавляем "0" в список значений
            quantity_values = [str(i) for i in range(0, eq_data["max_quantity"] + 1)]
            quantity_var = ctk.StringVar(value=str(eq_data["max_quantity"]))
            quantity_combobox = ctk.CTkComboBox(
                row_frame,
                variable=quantity_var,
                width=col_widths[0] - 20,
                fg_color=self.current_colors["entry_bg"],
                text_color=self.current_colors["entry_text"],
                button_color=self.current_colors["accent"],
                button_hover_color=self.current_colors["hover"],
                font=self.main_font,
                dropdown_font=self.main_font,
                state="readonly"
            )
            quantity_combobox.set(str(eq_data["max_quantity"]))  # Default to max quantity
            quantity_combobox.grid(row=0, column=0, padx=5, pady=5)
            CTkScrollableDropdown(quantity_combobox, values=quantity_values, button_color="transparent")

            self.quantity_comboboxes.append(quantity_combobox)

            # Sequential number
            number_label = ctk.CTkLabel(
                row_frame,
                text=str(counter),
                font=self.main_font,
                width=col_widths[1],
                anchor="center",
                text_color=self.current_colors["text"]
            )
            number_label.grid(row=0, column=1, padx=5, pady=5)
            counter += 1

            # Total Quantity
            total_label = ctk.CTkLabel(
                row_frame,
                text=str(eq_data["total_quantity"]),
                font=self.main_font,
                width=col_widths[2],
                anchor="center",
                text_color=self.current_colors["text"]
            )
            total_label.grid(row=0, column=2, padx=5, pady=5)

            # Equipment Type
            type_label = ctk.CTkLabel(
                row_frame,
                text=eq_data["type"],
                font=self.main_font,
                width=col_widths[3],
                anchor="center",
                text_color=self.current_colors["text"]
            )
            type_label.grid(row=0, column=3, padx=5, pady=5)

            # Object address
            address_label = ctk.CTkLabel(
                row_frame,
                text=eq_data["address"],  # Показываем конкретный адрес
                font=self.main_font,
                width=col_widths[4],
                anchor="center",
                wraplength=col_widths[4] - 20,
                text_color=self.current_colors["text"]
            )
            address_label.grid(row=0, column=4, padx=5, pady=5)

            # Additional info
            info_label = ctk.CTkLabel(
                row_frame,
                text=f"{eq_data['notes']}"[1:-1],
                font=self.main_font,
                width=col_widths[5],
                anchor="center",
                text_color=self.current_colors["text"]
            )
            info_label.grid(row=0, column=5, padx=5, pady=5)

        # Update count label
        if self.lbl_count:
            self.lbl_count.configure(
                text=f"Equipment types: {len(self.aggregated_data)} | Select quantity for each type"
            )

    def update_table_table_mode(self):
        """Update table in 'table' mode"""
        if not self.current_data:
            empty_label = ctk.CTkLabel(
                self.data_frame,
                text="No data to display. Load a file.",
                font=self.main_font,
                text_color=self.current_colors["text"]
            )
            empty_label.pack(pady=50)
            return

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
            if TABLE_SETTINGS["alternating_row_colors"]:
                bg_color = self.current_colors["table_row1"] if row_idx % 2 == 0 else self.current_colors["table_row2"]
            else:
                bg_color = "transparent"

            row_frame = ctk.CTkFrame(self.data_frame, fg_color=bg_color, corner_radius=0)
            row_frame.pack(fill=ctk.X)

            # Configure grid columns to be centered
            for i in range(6):
                row_frame.grid_columnconfigure(i, weight=1)

            # Checkbox
            var = ctk.BooleanVar(value=True)
            cb = ctk.CTkCheckBox(
                row_frame,
                variable=var,
                text="",
                width=col_widths[0],
                onvalue=True,
                offvalue=False,
                command=lambda r=row_idx, v=var: self.on_check_table_mode(r, v),
                fg_color=COLORS["blue"],
                hover_color=COLORS["cyan"],
                border_color=self.current_colors["entry_border"],
                checkmark_color=self.current_colors["text"]
            )
            cb.grid(row=0, column=0, padx=2, pady=2)
            self.checkboxes.append((cb, var))

            for col_idx, value in enumerate(row_data):
                if col_idx < 5:
                    text = str(value)
                    lbl = ctk.CTkLabel(
                        row_frame,
                        text=text,
                        font=self.main_font,
                        width=col_widths[col_idx + 1],
                        anchor="center",
                        wraplength=col_widths[col_idx + 1] - 10,
                        text_color=self.current_colors["text"]
                    )
                    lbl.grid(row=0, column=col_idx + 1, padx=5, pady=2, sticky="")

        if self.lbl_count:
            self.lbl_count.configure(text=self.texts["records_count"].format(
                len(self.current_data), len(self.active_rows)
            ))

    def on_check_table_mode(self, row, var):
        """Handler for checkbox state change in table mode"""
        if var.get():
            if row not in self.active_rows:
                self.active_rows.append(row)
        else:
            if row in self.active_rows:
                self.active_rows.remove(row)

        self.active_rows.sort()
        if self.lbl_count:
            self.lbl_count.configure(text=self.texts["records_count"].format(
                len(self.current_data), len(self.active_rows)
            ))

    def switch_template(self, template_num, update_ui=True):
        """Switch template"""
        self.current_template = template_num
        if self.lbl_template:
            self.lbl_template.configure(text=f"{self.texts['current_template']} {template_num}")

        if template_num == 1:
            if self.btn_template1:
                self.btn_template1.configure(fg_color=self.current_colors["accent"])
            if self.btn_template2:
                self.btn_template2.configure(fg_color=self.current_colors["bg"])
        else:
            if self.btn_template1:
                self.btn_template1.configure(fg_color=self.current_colors["bg"])
            if self.btn_template2:
                self.btn_template2.configure(fg_color=self.current_colors["accent"])

        if update_ui and hasattr(self, 'current_data') and (self.current_data or self.aggregated_data):
            self.update_table()

    def switch_mode(self, mode_num, update_ui=True):
        """Switch mode"""
        self.current_mode = mode_num
        mode_text = self.texts["mode_table"] if mode_num == 1 else self.texts["mode_text"]
        if self.lbl_mode:
            self.lbl_mode.configure(text=f"{self.texts['current_mode']} {mode_text}")

        if mode_num == 1:
            if self.btn_mode_table:
                self.btn_mode_table.configure(fg_color=self.current_colors["accent"])
            if self.btn_mode_text:
                self.btn_mode_text.configure(fg_color=self.current_colors["bg"])
        else:
            if self.btn_mode_table:
                self.btn_mode_table.configure(fg_color=self.current_colors["bg"])
            if self.btn_mode_text:
                self.btn_mode_text.configure(fg_color=self.current_colors["accent"])

        if update_ui and hasattr(self, 'current_data') and (self.current_data or self.aggregated_data):
            if mode_num == 0 and self.current_data:
                self.aggregated_data = self.aggregate_data(self.current_data)
            self.update_table()

    def toggle_all(self, state):
        """Select all/deselect all (table mode only)"""
        if self.current_mode == 1:  # Table mode only
            if state:
                self.active_rows = list(range(len(self.current_data)))
            else:
                self.active_rows = []

            for cb, var in self.checkboxes:
                if var.get() != state:
                    var.set(state)

            if self.lbl_count:
                self.lbl_count.configure(text=self.texts["records_count"].format(
                    len(self.current_data), len(self.active_rows)
                ))

    def invert_selection(self):
        """Invert selection (table mode only)"""
        if self.current_mode == 1:  # Table mode only
            for cb, var in self.checkboxes:
                var.set(not var.get())

            self.active_rows = [i for i, (_, var) in enumerate(self.checkboxes) if var.get()]
            if self.lbl_count:
                self.lbl_count.configure(text=self.texts["records_count"].format(
                    len(self.current_data), len(self.active_rows)
                ))

    def load_file_from_path(self, file_path=None, update_ui=True):
        """Load file"""

        if not file_path:
            file_path = filedialog.askopenfilename(
                title="Select Word document",
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")]
            )

        if not file_path:
            return

        try:
            self.file_path = file_path
            sys_name = os.path.basename(file_path)
            for_name = sys_name if len(sys_name) < 30 else sys_name[:27] + '...'
            if self.lbl_file:
                self.lbl_file.configure(text=for_name, text_color=self.current_colors["accent"])

            doc = Document(file_path)
            info = find_info_in_doc(doc, file_path)

            self.cur7 = info.get("дата выполнения")
            self.date2 = info.get("дата обращения")
            self.table2_data = info.get("таблица работ")

            if info.get("номер заявки") and self.entry_num:
                self.entry_num.delete(0, ctk.END)
                self.entry_num.insert(0, info["номер заявки"])

            loaded_mode = info.get("режим", 1)
            self.switch_mode(loaded_mode, update_ui=False)

            self.current_data = []
            for row in info.get("адреса объектов", []):
                if row and row[0] != '№п/п':
                    self.current_data.append(row[:5])

            # ВАЖНО: Добавляем агрегацию данных и обновление таблицы для текстового режима
            if self.current_mode == 0:
                self.aggregated_data = self.aggregate_data(self.current_data)
            if update_ui:
                self.update_table()
            else:
                if update_ui:
                    self.update_table()

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror(self.texts["error"], self.texts["load_error"].format(str(e)))

    def create_doc(self):
        """Create document"""
        if self.current_mode == 0:
            # Text mode - collect data from dropdowns
            if not self.aggregated_data:
                messagebox.showwarning(self.texts["warning"], "No data to create document")
                return

            # Collect data for equipment types with quantity > 0
            selected_equipment = []
            counter = 1  # Счетчик для №П/П
            for idx, eq_data in enumerate(self.aggregated_data):
                if idx < len(self.quantity_comboboxes):
                    selected_quantity_str = self.quantity_comboboxes[idx].get()

                    # Проверяем, если выбрано "0", пропускаем эту строку
                    if selected_quantity_str == "0":
                        continue

                    selected_quantity = int(selected_quantity_str)

                    if selected_quantity > 0:
                        # Собираем оригинальные примечания из addresses
                        original_notes = []
                        for addr_info in eq_data["addresses"]:
                            # Ищем оригинальные примечания в данных
                            for original_row in self.current_data:
                                if len(original_row) >= 5 and str(original_row[2]) == eq_data["type"] and str(
                                        original_row[3]) == addr_info["address"]:
                                    if original_row[4]:  # Если есть примечание
                                        original_notes.append(str(original_row[4]))
                                    break

                        selected_equipment.append({
                            "type": eq_data["type"],
                            "quantity": selected_quantity,
                            "addresses": eq_data["addresses"],
                            "total_quantity": eq_data["total_quantity"],
                            "address": eq_data["address"],
                            "number": counter,
                            "notes": original_notes  # Сохраняем оригинальные примечания
                        })
                        counter += 1

            if not selected_equipment:
                messagebox.showwarning(self.texts["warning"],
                                       "Select equipment quantity for installation (must be > 0)")
                return

            # Convert selected data to format for document creation functions
            active_data = []
            for eq in selected_equipment:
                # Используем конкретный адрес из агрегированных данных
                addresses_str = eq["address"]

                # Формируем примечания
                if eq["notes"]:
                    notes_str = "; ".join(eq["notes"])
                else:
                    notes_str = ""

                active_data.append([
                    str(eq["number"]),  # №П/П - добавляем порядковый номер
                    str(eq["quantity"]),  # Qty
                    eq["type"],  # Type
                    addresses_str,  # Address
                    notes_str  # Original notes instead of quantity info
                ])
        else:
            # Table mode - use existing logic
            if not self.current_data:
                messagebox.showwarning(self.texts["warning"], self.texts["no_data"])
                return

            if not self.active_rows:
                messagebox.showwarning(self.texts["warning"], self.texts["no_active"])
                return

            active_data = []
            for i in self.active_rows:
                row = self.current_data[i]
                if not row or all(cell == '' for cell in row):
                    continue

                normalized_row = list(row)[:5]
                while len(normalized_row) < 5:
                    normalized_row.append("")
                active_data.append(normalized_row)

        file_path = filedialog.asksaveasfilename(
            title="Save document",
            defaultextension=".docx",
            filetypes=[("Word documents", "*.docx")]
        )

        if not file_path:
            return

        try:
            if not active_data:
                messagebox.showwarning(self.texts["warning"], "No data to create document")
                return

            num = self.entry_num.get() if self.entry_num else "ERROR_IN_NUM"
            date = self.entry_date.get() if self.entry_date else "ERROR_IN_DATE"
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
                template_name = "Template 2"
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
                template_name = "Template 1"

            import shutil
            if output and os.path.exists(output):
                shutil.move(output, file_path)

            messagebox.showinfo(
                "Документ создан", f"{template_name} создан с {len(active_data)} записями. Сохранен в: {file_path}"
            )

        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror(self.texts["error"], self.texts["create_error"].format(str(e)))

    def on_closing(self):
        """Window close handler"""
        self.font_manager.cleanup_fonts()
        self.root.destroy()

    def run(self):
        """Run application"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ModernApp()
    app.run()
