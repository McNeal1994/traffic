import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from typing import List, Optional
from datetime import datetime
from pathlib import Path
from ..utils.config import Config
from ..core.data_processor import DataProcessor

class AddressTab(ttk.Frame):
    """Вкладка пошуку адрес без координат."""
    
    def __init__(
        self,
        parent: ttk.Notebook,
        config: Config,
        data_processor: DataProcessor
    ):
        """
        Ініціалізація вкладки пошуку адрес.
        
        Args:
            parent: Батьківський віджет
            config: Об'єкт конфігурації
            data_processor: Обробник даних
        """
        super().__init__(parent)
        self.config = config
        self.data_processor = data_processor
        self.df = None
        self.current_time = datetime.strptime(
            "2025-07-18 13:40:36",
            "%Y-%m-%d %H:%M:%S"
        )
        
        self._create_widgets()

    def _create_widgets(self) -> None:
        """Створення віджетів."""
        # Фрейм фільтрів
        filter_frame = ttk.LabelFrame(self, text="Фільтри")
        filter_frame.pack(fill=tk.X, padx=5, pady=5)

        # Фільтр за датою
        ttk.Label(filter_frame, text="Дата:").grid(
            row=0, column=0, padx=5, pady=5
        )
        self.date_var = tk.StringVar(
            value=self.current_time.strftime("%d.%m.%Y")
        )
        self.date_entry = ttk.Entry(
            filter_frame,
            textvariable=self.date_var
        )
        self.date_entry.grid(row=0, column=1, padx=5, pady=5)

        # Фільтр за адресою
        ttk.Label(filter_frame, text="Адреса:").grid(
            row=0, column=2, padx=5, pady=5
        )
        self.address_var = tk.StringVar()
        self.address_entry = ttk.Entry(
            filter_frame,
            textvariable=self.address_var,
            width=40
        )
        self.address_entry.grid(row=0, column=3, padx=5, pady=5)

        # Кнопки
        buttons_frame = ttk.Frame(filter_frame)
        buttons_frame.grid(row=0, column=4, padx=5, pady=5)

        ttk.Button(
            buttons_frame,
            text="Знайти адреси",
            command=self._find_addresses
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            buttons_frame,
            text="Скинути",
            command=self._reset_filters
        ).pack(side=tk.LEFT, padx=2)

        # Статистика
        stats_frame = ttk.LabelFrame(self, text="Статистика")
        stats_frame.pack(fill=tk.X, padx=5, pady=5)

        self.stats_text = tk.Text(
            stats_frame,
            height=3,
            wrap=tk.WORD
        )
        self.stats_text.pack(fill=tk.X, padx=5, pady=5)

        # Таблиця результатів
        results_frame = ttk.LabelFrame(self, text="Результати")
        results_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Створюємо таблицю
        columns = ("Адреса БС", "Кількість записів", "Останній запис")
        self.tree = ttk.Treeview(
            results_frame,
            columns=columns,
            show="headings"
        )

        # Налаштовуємо заголовки
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        self.tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        # Додаємо прокрутку
        scrollbar = ttk.Scrollbar(
            results_frame,
            orient=tk.VERTICAL,
            command=self.tree.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Контекстне меню
        self.context_menu = tk.Menu(self, tearoff=0)
        self.context_menu.add_command(
            label="Копіювати адресу",
            command=self._copy_address
        )
        self.context_menu.add_command(
            label="Показати деталі",
            command=self._show_details
        )

        self.tree.bind('<Button-3>', self._show_context_menu)
        
    def load_files(self, files: List[str]) -> None:
        """
        Завантаження файлів.
        
        Args:
            files: Список шляхів до файлів
        """
        try:
            self.df = self.data_processor.process_traffic_files(files)
            self._update_statistics()
            
        except Exception as e:
            messagebox.showerror("Помилка", str(e))
            
    def _find_addresses(self) -> None:
        """Пошук адрес без координат."""
        if self.df is None:
            messagebox.showwarning(
                "Попередження",
                "Немає даних для аналізу"
            )
            return
            
        try:
            # Фільтруємо дані
            filtered_df = self.df.copy()
            
            if self.date_var.get():
                date = pd.to_datetime(
                    self.date_var.get(),
                    format='%d.%m.%Y'
                )
                filtered_df = filtered_df[
                    filtered_df['Дата'].dt.date == date.date()
                ]
                
            if self.address_var.get():
                filtered_df = filtered_df[
                    filtered_df['Адреса БС'].str.contains(
                        self.address_var.get(),
                        case=False,
                        na=False
                    )
                ]
            
            # Знаходимо адреси без координат
            no_coords = filtered_df[
                filtered_df['Широта'].isna() |
                filtered_df['Долгота'].isna()
            ]
            
            # Групуємо за адресою
            grouped = no_coords.groupby('Адреса БС').agg({
                'Дата': ['count', 'max']
            }).reset_index()
            
            grouped.columns = ['Адреса БС', 'Кількість записів', 'Останній запис']
            grouped = grouped.sort_values('Кількість записів', ascending=False)
            
            # Оновлюємо таблицю
            self._update_table(grouped)
            
        except Exception as e:
            messagebox.showerror("Помилка", str(e))
            
    def _update_table(self, df: pd.DataFrame) -> None:
        """
        Оновлення таблиці результатів.
        
        Args:
            df: DataFrame з результатами
        """
        # Очищаємо таблицю
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Додаємо нові дані
        for _, row in df.iterrows():
            values = (
                row['Адреса БС'],
                row['Кількість записів'],
                row['Останній запис'].strftime('%d.%m.%Y %H:%M:%S')
            )
            self.tree.insert("", tk.END, values=values)
            
    def _update_statistics(self) -> None:
        """Оновлення статистики."""
        if self.df is None:
            return
            
        # Підраховуємо статистику
        total_addresses = self.df['Адреса БС'].nunique()
        no_coords = self.df[
            self.df['Широта'].isna() |
            self.df['Долгота'].isna()
        ]
        addresses_no_coords = no_coords['Адреса БС'].nunique()
        
        stats = (
            f"Всього унікальних адрес: {total_addresses}\n"
            f"Адрес без координат: {addresses_no_coords}\n"
            f"Відсоток адрес без координат: "
            f"{addresses_no_coords/total_addresses*100:.1f}%"
        )
        
        self.stats_text.delete(1.0, tk.END)
        self.stats_text.insert(tk.END, stats)
        
    def _reset_filters(self) -> None:
        """Скидання фільтрів."""
        self.date_var.set(self.current_time.strftime("%d.%m.%Y"))
        self.address_var.set("")
        
    def _show_context_menu(self, event) -> None:
        """
        Показ контекстного меню.
        
        Args:
            event: Подія кліку правою кнопкою миші
        """
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
            
    def _copy_address(self) -> None:
        """Копіювання адреси в буфер обміну."""
        selection = self.tree.selection()
        if selection:
            address = self.tree.item(selection[0])['values'][0]
            self.clipboard_clear()
            self.clipboard_append(address)
            
    def _show_details(self) -> None:
        """Показ деталей про адресу."""
        selection = self.tree.selection()
        if selection:
            address = self.tree.item(selection[0])['values'][0]
            details = self.df[self.df['Адреса БС'] == address].copy()
            
            details_window = tk.Toplevel(self)
            details_window.title(f"Деталі адреси: {address}")
            details_window.geometry("600x400")
            
            text = tk.Text(details_window, wrap=tk.WORD)
            text.pack(fill=tk.BOTH, expand=True)
            
            details_str = (
                f"Адреса: {address}\n"
                f"Кількість записів: {len(details)}\n"
                f"Перший запис: {details['Дата'].min()}\n"
                f"Останній запис: {details['Дата'].max()}\n"
                f"Унікальних абонентів: {details['Абонент А'].nunique()}\n\n"
                "Останні 10 записів:\n"
            )
            
            text.insert(tk.END, details_str)
            
            # Додаємо останні 10 записів
            last_records = details.sort_values('Дата', ascending=False).head(10)
            for _, row in last_records.iterrows():
                record = (
                    f"Дата: {row['Дата']} "
                    f"Абонент: {row['Абонент А']}\n"
                )
                text.insert(tk.END, record)
                
            text.configure(state='disabled')
            
    def save_results(self) -> None:
        """Збереження результатів."""
        try:
            # Збираємо дані з таблиці
            data = []
            for item in self.tree.get_children():
                data.append(self.tree.item(item)['values'])
                
            if not data:
                messagebox.showwarning(
                    "Попередження",
                    "Немає даних для збереження"
                )
                return
                
            # Створюємо DataFrame
            results_df = pd.DataFrame(
                data,
                columns=['Адреса БС', 'Кількість записів', 'Останній запис']
            )
            
            # Зберігаємо файл
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel файли", "*.xlsx"),
                    ("CSV файли", "*.csv"),
                    ("Всі файли", "*.*")
                ]
            )
            
            if file_path:
                if file_path.endswith('.xlsx'):
                    results_df.to_excel(file_path, index=False)
                elif file_path.endswith('.csv'):
                    results_df.to_csv(file_path, index=False)
                    
                messagebox.showinfo(
                    "Успіх",
                    f"Результати збережено в {file_path}"
                )
                
        except Exception as e:
            messagebox.showerror("Помилка збереження", str(e))