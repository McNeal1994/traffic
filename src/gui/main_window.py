import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
from typing import Optional
from pathlib import Path
from datetime import datetime
from ..utils.config import Config
from ..core.data_processor import DataProcessor
from ..core.geo_processor import GeoProcessor
from .traffic_tab import TrafficTab
from .movement_tab import MovementTab
from .address_tab import AddressTab
from .settings_dialog import SettingsDialog

class MainWindow:
    """Головне вікно програми."""
    
    def __init__(self, root: tk.Tk, config: Config):
        """
        Ініціалізація головного вікна.
        
        Args:
            root: Кореневий віджет
            config: Об'єкт конфігурації
        """
        self.root = root
        self.config = config
        self.data_processor = DataProcessor(config)
        self.geo_processor = GeoProcessor(config)
        
        self._setup_window()
        self._create_menu()
        self._create_notebook()
        self._create_status_bar()
        
        # Логуємо запуск програми
        logging.info(
            f"Програму запущено користувачем {config.get('app.user.login')} "
            f"о {config.get('app.current_time')}"
        )
        
    def _setup_window(self) -> None:
        """Налаштування параметрів вікна."""
        self.root.title("Аналізатор трафіку")
        self.root.geometry("1200x800")
        
        # Встановлюємо іконку
        try:
            self.root.iconbitmap("assets/icon.ico")
        except:
            logging.warning("Не вдалося завантажити іконку програми")
        
        # Налаштовуємо стилі
        style = ttk.Style()
        style.theme_use('clam')
        
        style.configure(
            "Header.TLabel",
            font=('Helvetica', 12, 'bold'),
            padding=5
        )
        
        style.configure(
            "Status.TLabel",
            padding=2
        )
        
    def _create_menu(self) -> None:
        """Створення головного меню."""
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)
        
        # Меню файлу
        file_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Файл", menu=file_menu)
        file_menu.add_command(label="Відкрити файли...", command=self._open_files)
        file_menu.add_command(label="Зберегти результати...", command=self._save_results)
        file_menu.add_separator()
        file_menu.add_command(label="Вихід", command=self.root.quit)
        
        # Меню інструментів
        tools_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Інструменти", menu=tools_menu)
        tools_menu.add_command(label="Налаштування...", command=self._show_settings)
        tools_menu.add_command(label="Очистити кеш", command=self._clear_cache)
        
        # Меню довідки
        help_menu = tk.Menu(self.menu_bar, tearoff=0)
        self.menu_bar.add_cascade(label="Довідка", menu=help_menu)
        help_menu.add_command(label="Документація", command=self._show_docs)
        help_menu.add_command(label="Про програму", command=self._show_about)
        
    def _create_notebook(self) -> None:
        """Створення вкладок."""
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Створюємо вкладки
        self.traffic_tab = TrafficTab(
            self.notebook,
            self.config,
            self.data_processor
        )
        self.movement_tab = MovementTab(
            self.notebook,
            self.config,
            self.data_processor,
            self.geo_processor
        )
        self.address_tab = AddressTab(
            self.notebook,
            self.config,
            self.data_processor
        )
        
        # Додаємо вкладки до ноутбука
        self.notebook.add(self.traffic_tab, text="Аналіз трафіку")
        self.notebook.add(self.movement_tab, text="Переміщення")
        self.notebook.add(self.address_tab, text="Пошук адрес")
        
        # Додаємо обробник зміни вкладки
        self.notebook.bind('<<NotebookTabChanged>>', self._on_tab_changed)
        
    def _create_status_bar(self) -> None:
        """Створення рядка статусу."""
        self.status_bar = ttk.Frame(self.root)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
        
        self.status_label = ttk.Label(
            self.status_bar,
            text="Готово",
            style="Status.TLabel"
        )
        self.status_label.pack(side=tk.LEFT)
        
        self.progress_bar = ttk.Progressbar(
            self.status_bar,
            mode='determinate',
            length=200
        )
        self.progress_bar.pack(side=tk.RIGHT, padx=5)
        
    def _open_files(self) -> None:
        """Відкриття файлів."""
        try:
            files = filedialog.askopenfilenames(
                title="Виберіть файли трафіку",
                filetypes=[
                    ("Excel файли", "*.xlsx *.xls"),
                    ("Всі файли", "*.*")
                ]
            )
            
            if files:
                # Передаємо файли активній вкладці
                current_tab = self.notebook.select()
                if current_tab == str(self.traffic_tab):
                    self.traffic_tab.load_files(files)
                elif current_tab == str(self.movement_tab):
                    self.movement_tab.load_files(files)
                elif current_tab == str(self.address_tab):
                    self.address_tab.load_files(files)
                    
        except Exception as e:
            self._show_error("Помилка відкриття файлів", str(e))
            
    def _save_results(self) -> None:
        """Збереження результатів."""
        try:
            current_tab = self.notebook.select()
            if current_tab == str(self.traffic_tab):
                self.traffic_tab.save_results()
            elif current_tab == str(self.movement_tab):
                self.movement_tab.save_results()
            elif current_tab == str(self.address_tab):
                self.address_tab.save_results()
                
        except Exception as e:
            self._show_error("Помилка збереження", str(e))
            
    def _show_settings(self) -> None:
        """Показ діалогу налаштувань."""
        SettingsDialog(self.root, self.config)
        
    def _clear_cache(self) -> None:
        """Очищення кешу програми."""
        try:
            cache_dir = Path("cache")
            if cache_dir.exists():
                for file in cache_dir.glob("*"):
                    file.unlink()
                cache_dir.rmdir()
            self.update_status("Кеш очищено")
            
        except Exception as e:
            self._show_error("Помилка очищення кешу", str(e))
            
    def _show_docs(self) -> None:
        """Відкриття документації."""
        docs_path = Path("docs/index.html")
        if docs_path.exists():
            import webbrowser
            webbrowser.open(str(docs_path))
        else:
            messagebox.showwarning(
                "Попередження",
                "Документація недоступна"
            )
            
    def _show_about(self) -> None:
        """Показ інформації про програму."""
        messagebox.showinfo(
            "Про програму",
            "Аналізатор трафіку\n"
            "Версія 1.0\n\n"
            f"Поточний користувач: {self.config.get('app.user.login')}\n"
            f"Час запуску: {self.config.get('app.current_time')}\n\n"
            "© 2024 Всі права захищено"
        )
        
    def _show_error(self, title: str, message: str) -> None:
        """
        Показ повідомлення про помилку.
        
        Args:
            title: Заголовок
            message: Повідомлення
        """
        logging.error(message)
        messagebox.showerror(title, message)
        
    def _on_tab_changed(self, event) -> None:
        """
        Обробник зміни активної вкладки.
        
        Args:
            event: Подія зміни вкладки
        """
        current_tab = self.notebook.select()
        tab_text = self.notebook.tab(current_tab, "text")
        self.update_status(f"Активна вкладка: {tab_text}")
        
    def update_status(self, message: str) -> None:
        """
        Оновлення повідомлення в рядку статусу.
        
        Args:
            message: Повідомлення
        """
        self.status_label.config(text=message)
        
    def update_progress(self, value: int) -> None:
        """
        Оновлення індикатора прогресу.
        
        Args:
            value: Значення прогресу (0-100)
        """
        self.progress_bar['value'] = value
        self.root.update_idletasks()

