import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Any
from datetime import datetime
from ..utils.config import Config

class SettingsDialog(tk.Toplevel):
    """Діалог налаштувань програми."""
    
    def __init__(self, parent: tk.Tk, config: Config):
        """
        Ініціалізація діалогу налаштувань.
        
        Args:
            parent: Батьківське вікно
            config: Об'єкт конфігурації
        """
        super().__init__(parent)
        self.config = config
        self.result = None
        
        self.title("Налаштування")
        self.geometry("600x400")
        self.resizable(False, False)
        
        # Встановлюємо модальність
        self.transient(parent)
        self.grab_set()
        
        self._create_widgets()
        self._load_settings()
        
        # Центруємо вікно
        self.update_idletasks()
        width = self.winfo_width()
        height = self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (width // 2)
        y = (self.winfo_screenheight() // 2) - (height // 2)
        self.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
    def _create_widgets(self) -> None:
        """Створення віджетів."""
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Вкладка загальних налаштувань
        general_frame = ttk.Frame(notebook)
        notebook.add(general_frame, text="Загальні")
        
        # Налаштування користувача
        user_frame = ttk.LabelFrame(general_frame, text="Користувач")
        user_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(user_frame, text="Логін:").grid(row=0, column=0, padx=5, pady=5)
        self.login_var = tk.StringVar(value=self.config.get("app.user.login"))
        ttk.Entry(
            user_frame,
            textvariable=self.login_var,
            state='readonly'
        ).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(user_frame, text="Час входу:").grid(row=1, column=0, padx=5, pady=5)
        self.time_var = tk.StringVar(value=self.config.get("app.current_time"))
        ttk.Entry(
            user_frame,
            textvariable=self.time_var,
            state='readonly'
        ).grid(row=1, column=1, padx=5, pady=5)
        
        # Налаштування карти
        map_frame = ttk.LabelFrame(general_frame, text="Карта")
        map_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(map_frame, text="Стиль за замовчуванням:").grid(
            row=0, column=0, padx=5, pady=5
        )
        self.map_style_var = tk.StringVar()
        self.map_style_combo = ttk.Combobox(
            map_frame,
            textvariable=self.map_style_var,
            values=self.config.get("map.styles")
        )
        self.map_style_combo.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(map_frame, text="Кут сектора:").grid(
            row=1, column=0, padx=5, pady=5
        )
        self.sector_angle_var = tk.StringVar()
        ttk.Entry(
            map_frame,
            textvariable=self.sector_angle_var
        ).grid(row=1, column=1, padx=5, pady=5)
        
        ttk.Label(map_frame, text="Радіус сектора:").grid(
            row=2, column=0, padx=5, pady=5
        )
        self.sector_radius_var = tk.StringVar()
        ttk.Entry(
            map_frame,
            textvariable=self.sector_radius_var
        ).grid(row=2, column=1, padx=5, pady=5)
        
        # Налаштування фільтрів
        filter_frame = ttk.LabelFrame(general_frame, text="Фільтри")
        filter_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(filter_frame, text="Початок дня:").grid(
            row=0, column=0, padx=5, pady=5
        )
        self.day_start_var = tk.StringVar()
        ttk.Entry(
            filter_frame,
            textvariable=self.day_start_var
        ).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(filter_frame, text="Кінець дня:").grid(
            row=1, column=0, padx=5, pady=5
        )
        self.day_end_var = tk.StringVar()
        ttk.Entry(
            filter_frame,
            textvariable=self.day_end_var
        ).grid(row=1, column=1, padx=5, pady=5)
        
        # Кнопки
        buttons_frame = ttk.Frame(self)
        buttons_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(
            buttons_frame,
            text="Зберегти",
            command=self._save_settings
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            buttons_frame,
            text="Скасувати",
            command=self.destroy
        ).pack(side=tk.RIGHT, padx=5)
        
        ttk.Button(
            buttons_frame,
            text="За замовчуванням",
            command=self._load_defaults
        ).pack(side=tk.RIGHT, padx=5)
        
    def _load_settings(self) -> None:
        """Завантаження поточних налаштувань."""
        self.map_style_var.set(self.config.get("map.default_style"))
        self.sector_angle_var.set(str(self.config.get("map.sector.angle")))
        self.sector_radius_var.set(str(self.config.get("map.sector.radius")))
        self.day_start_var.set(self.config.get("filters.day_start"))
        self.day_end_var.set(self.config.get("filters.day_end"))
        
    def _load_defaults(self) -> None:
        """Завантаження налаштувань за замовчуванням."""
        if messagebox.askyesno(
            "Підтвердження",
            "Відновити налаштування за замовчуванням?"
        ):
            self.map_style_var.set("OpenStreetMap")
            self.sector_angle_var.set("120")
            self.sector_radius_var.set("500")
            self.day_start_var.set("07:00")
            self.day_end_var.set("20:00")
            
    def _validate_settings(self) -> bool:
        """
        Валідація налаштувань.
        
        Returns:
            bool: True якщо всі налаштування валідні
        """
        try:
            # Перевіряємо кут сектора
            angle = float(self.sector_angle_var.get())
            if not 0 < angle <= 360:
                raise ValueError("Кут сектора має бути від 0 до 360 градусів")
                
            # Перевіряємо радіус сектора
            radius = float(self.sector_radius_var.get())
            if radius <= 0:
                raise ValueError("Радіус сектора має бути більше 0")
                
            # Перевіряємо час
            for time_str in [self.day_start_var.get(), self.day_end_var.get()]:
                hours, minutes = map(int, time_str.split(':'))
                if not (0 <= hours < 24 and 0 <= minutes < 60):
                    raise ValueError("Неправильний формат часу")
                    
            return True
            
        except ValueError as e:
            messagebox.showerror("Помилка", str(e))
            return False
            
    def _save_settings(self) -> None:
        """Збереження налаштувань."""
        if not self._validate_settings():
            return
            
        try:
            # Формуємо нові налаштування
            new_config = {
                "map": {
                    "default_style": self.map_style_var.get(),
                    "sector": {
                        "angle": float(self.sector_angle_var.get()),
                        "radius": float(self.sector_radius_var.get())
                    }
                },
                "filters": {
                    "day_start": self.day_start_var.get(),
                    "day_end": self.day_end_var.get()
                }
            }
            
            # Зберігаємо конфігурацію
            self.config._save_config(new_config)
            
            messagebox.showinfo(
                "Успіх",
                "Налаштування збережено"
            )
            self.destroy()
            
        except Exception as e:
            messagebox.showerror(
                "Помилка",
                f"Помилка збереження налаштувань: {e}"
            )