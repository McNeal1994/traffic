"""
Модуль для аналізу переміщень та створення візуалізацій.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from datetime import datetime, time
import folium
from folium.plugins import HeatMap
import json
import os
from math import radians, sin, cos, sqrt, atan2
from collections import defaultdict
from shapely.geometry import Point, shape
from typing import List, Dict, Optional, Tuple
import logging

class MovementTab(ttk.Frame):
    """Вкладка для аналізу переміщень."""

    def __init__(
        self,
        parent: ttk.Notebook,
        config: 'Config',
        data_processor: 'DataProcessor',
        geo_processor: 'GeoProcessor'
    ):
        """
        Ініціалізація вкладки переміщень.

        Args:
            parent: Батьківський віджет
            config: Конфігурація програми
            data_processor: Обробник даних
            geo_processor: Обробник геоданих
        """
        super().__init__(parent)
        self.config = config
        self.data_processor = data_processor
        self.geo_processor = geo_processor

        # Ініціалізація змінних
        self.traffic_files: List[str] = []
        self.geojson_file: Optional[str] = None
        self.polygon = None
        self.current_time = datetime.strptime(
            "2025-07-21 12:07:55",
            "%Y-%m-%d %H:%M:%S"
        )
        self.current_user = "McNeal1994"

        # Створення віджетів
        self._create_widgets()

    def _create_widgets(self) -> None:
        """Створення елементів інтерфейсу."""
        # Створюємо основний фрейм з двома колонками
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Ліва колонка для операцій
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=2)

        # Права колонка для логу
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=2)

        # Створюємо всі фрейми в лівій колонці
        self._create_file_operations_frame(left_frame)
        self._create_parameters_frame(left_frame)
        self._create_date_frame(left_frame)
        self._create_similar_routes_frame(left_frame)
        self._create_common_movements_frame(left_frame)

        # Створюємо лог в правій колонці
        self._create_log_frame(right_frame)

        # Прогрес-бар внизу вікна
        self.progress_bar = ttk.Progressbar(self, mode='determinate')
        self.progress_bar.pack(fill=tk.X, padx=5, pady=5)

    def _create_file_operations_frame(self, parent) -> None:
        """Створення фрейму операцій з файлами."""
        frame = ttk.LabelFrame(parent, text="Операції з файлами")
        frame.pack(fill=tk.X, pady=2)

        ttk.Button(
            frame,
            text="Вибрати файли трафіку",
            command=self._select_traffic_files
        ).pack(side=tk.LEFT, padx=5, pady=2)

        ttk.Button(
            frame,
            text="Вибрати GeoJSON",
            command=self._select_geojson
        ).pack(side=tk.LEFT, padx=5, pady=2)

        ttk.Button(
            frame,
            text="Обробити",
            command=self._process_files
        ).pack(side=tk.LEFT, padx=5, pady=2)

    def _create_parameters_frame(self, parent) -> None:
        """Створення фрейму параметрів."""
        frame = ttk.LabelFrame(parent, text="Параметри")
        frame.pack(fill=tk.X, pady=2)

        # Використовуємо grid для кращого вирівнювання
        frame.grid_columnconfigure(1, weight=1)

        # Максимальна відстань
        ttk.Label(
            frame,
            text="Макс. відстань (м):"
        ).grid(row=0, column=0, padx=5, pady=2, sticky='e')

        self.max_distance = tk.StringVar(value="400")
        ttk.Entry(
            frame,
            textvariable=self.max_distance,
            width=10
        ).grid(row=0, column=1, padx=5, pady=2, sticky='w')

        # Параметри секторів
        ttk.Label(
            frame,
            text="Радіус сектору (м):"
        ).grid(row=1, column=0, padx=5, pady=2, sticky='e')

        self.sector_radius = tk.StringVar(value="500")
        ttk.Entry(
            frame,
            textvariable=self.sector_radius,
            width=10
        ).grid(row=1, column=1, padx=5, pady=2, sticky='w')

        ttk.Label(
            frame,
            text="Кут сектору (°):"
        ).grid(row=2, column=0, padx=5, pady=2, sticky='e')

        self.sector_angle = tk.StringVar(value="120")
        ttk.Entry(
            frame,
            textvariable=self.sector_angle,
            width=10
        ).grid(row=2, column=1, padx=5, pady=2, sticky='w')

        # Мінімальний час (день)
        ttk.Label(
            frame,
            text="Мін. час (день, хв.):"
        ).grid(row=1, column=0, padx=5, pady=2, sticky='e')

        self.day_min_duration = tk.StringVar(value="30")
        ttk.Entry(
            frame,
            textvariable=self.day_min_duration,
            width=10
        ).grid(row=1, column=1, padx=5, pady=2, sticky='w')

        # Мінімальний час (ніч)
        ttk.Label(
            frame,
            text="Мін. час (ніч, хв.):"
        ).grid(row=2, column=0, padx=5, pady=2, sticky='e')

        self.night_min_duration = tk.StringVar(value="60")
        ttk.Entry(
            frame,
            textvariable=self.night_min_duration,
            width=10
        ).grid(row=2, column=1, padx=5, pady=2, sticky='w')

        # Створення мап
        self.create_daily_maps = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frame,
            text="Створювати мапи за кожен день",
            variable=self.create_daily_maps
        ).grid(row=3, column=0, columnspan=2, padx=5, pady=2)

    def _create_date_frame(self, parent) -> None:
        """Створення фрейму роботи з датами."""
        frame = ttk.LabelFrame(parent, text="Карта за обрану дату")
        frame.pack(fill=tk.X, pady=2)

        ttk.Label(
            frame,
            text="Дата (DD.MM.YYYY):"
        ).pack(side=tk.LEFT, padx=5, pady=2)

        self.selected_date = tk.StringVar()
        ttk.Entry(
            frame,
            textvariable=self.selected_date,
            width=10
        ).pack(side=tk.LEFT, padx=5, pady=2)

        ttk.Button(
            frame,
            text="Створити карту",
            command=self._create_map_for_date
        ).pack(side=tk.LEFT, padx=5, pady=2)

    def _create_similar_routes_frame(self, parent) -> None:
        """Створення фрейму пошуку схожих маршрутів."""
        frame = ttk.LabelFrame(parent, text="Пошук схожих маршрутів")
        frame.pack(fill=tk.X, pady=2)

        ttk.Label(
            frame,
            text="Дата (DD.MM.YYYY):"
        ).pack(side=tk.LEFT, padx=5, pady=2)

        self.similar_routes_date = tk.StringVar()
        ttk.Entry(
            frame,
            textvariable=self.similar_routes_date,
            width=10
        ).pack(side=tk.LEFT, padx=5, pady=2)

        ttk.Label(
            frame,
            text="Схожість (%):"
        ).pack(side=tk.LEFT, padx=5, pady=2)

        self.similarity_threshold = tk.StringVar(value="70")
        ttk.Entry(
            frame,
            textvariable=self.similarity_threshold,
            width=5
        ).pack(side=tk.LEFT, padx=5, pady=2)

        ttk.Button(
            frame,
            text="Знайти схожі",
            command=self._find_similar_routes
        ).pack(side=tk.LEFT, padx=5, pady=2)

    def _create_log_frame(self, parent) -> None:
        """Створення фрейму логування."""
        frame = ttk.LabelFrame(parent, text="Лог операцій")
        frame.pack(fill=tk.BOTH, expand=True, pady=2)

        # Кнопки керування логом
        buttons_frame = ttk.Frame(frame)
        buttons_frame.pack(fill=tk.X, pady=2)

        ttk.Button(
            buttons_frame,
            text="Очистити лог",
            command=self._clear_log
        ).pack(side=tk.LEFT, padx=5)

        ttk.Button(
            buttons_frame,
            text="Зберегти лог",
            command=self._save_log
        ).pack(side=tk.LEFT, padx=5)

        # Текстове поле з прокруткою
        self.log_text = tk.Text(
            frame,
            wrap=tk.WORD,
            height=30  # Збільшуємо висоту для кращої видимості
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=2)

        scrollbar = ttk.Scrollbar(
            frame,
            orient=tk.VERTICAL,
            command=self.log_text.yview
        )
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _select_traffic_files(self) -> None:
        """Вибір файлів трафіку."""
        files = filedialog.askopenfilenames(
            title="Виберіть файли трафіку",
            filetypes=[
                ("Excel файли", "*.xlsx"),
                ("Всі файли", "*.*")
            ]
        )
        if files:
            self.traffic_files = files
            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Вибрано {len(files)} файлів трафіку\n\n"
            )
            self.log_text.see(tk.END)

    def _select_geojson(self) -> None:
        """Вибір файлу GeoJSON з полігонами."""
        file = filedialog.askopenfilename(
            title="Виберіть файл GeoJSON",
            filetypes=[
                ("GeoJSON файли", "*.geojson"),
                ("Всі файли", "*.*")
            ]
        )
        if file:
            self.geojson_file = file
            self.load_polygon()
            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): {datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Вибрано файл полігонів: {file}\n\n"
            )
            self.log_text.see(tk.END)

    def load_polygon(self) -> Optional[shape]:
        """
        Завантаження полігону з GeoJSON файлу.

        Returns:
            Optional[shape]: Об'єкт полігону або None
        """
        try:
            if self.geojson_file and os.path.exists(self.geojson_file):
                with open(self.geojson_file, 'r', encoding='utf-8') as f:
                    geojson = json.load(f)
                self.polygon = shape(geojson['features'][0]['geometry'])
                return self.polygon
        except Exception as e:
            logging.error(f"Помилка завантаження полігону: {e}")
            return None

    def parse_time(self, time_str: str) -> Optional[time]:
        """
        Перетворення часу з рядка в об'єкт time.

        Args:
            time_str: Час у форматі рядка

        Returns:
            Optional[time]: Об'єкт time або None при помилці
        """
        if isinstance(time_str, time):
            return time_str
        try:
            return datetime.strptime(str(time_str), '%H:%M:%S').time()
        except ValueError:
            try:
                return datetime.strptime(str(time_str), '%H:%M').time()
            except ValueError:
                logging.error(f"Неможливо розпізнати формат часу: {time_str}")
                return None

    def calculate_distance(
            self,
            lat1: float,
            lon1: float,
            lat2: float,
            lon2: float
    ) -> float:
        """
        Розрахунок відстані між двома точками за формулою гаверсинусів.

        Args:
            lat1: Широта першої точки
            lon1: Довгота першої точки
            lat2: Широта другої точки
            lon2: Довгота другої точки

        Returns:
            float: Відстань у метрах
        """
        R = 6371000  # радіус Землі в метрах

        lat1, lon1, lat2, lon2 = map(radians, [lat1, lon1, lat2, lon2])
        dlat = lat2 - lat1
        dlon = lon2 - lon1

        a = sin(dlat / 2) ** 2 + cos(lat1) * cos(lat2) * sin(dlon / 2) ** 2
        c = 2 * atan2(sqrt(a), sqrt(1 - a))

        return R * c

    def analyze_locations(
            self,
            df: pd.DataFrame,
            min_day_duration: int = 30,
            min_night_duration: int = 60
    ) -> Tuple[Dict, Dict]:
        """
        Аналіз місць перебування для визначення дому та роботи.

        Args:
            df: DataFrame з даними
            min_day_duration: Мінімальна тривалість денного перебування
            min_night_duration: Мінімальна тривалість нічного перебування

        Returns:
            Tuple[Dict, Dict]: (дім, робота) з їх характеристиками
        """
        locations = defaultdict(
            lambda: {
                'day_count': 0,
                'night_count': 0,
                'total_duration': 0,
                'address': None
            }
        )

        for date in df['Дата'].dt.date.unique():
            day_data = df[df['Дата'].dt.date == date].sort_values('Час')

            i = 0
            while i < len(day_data):
                row = day_data.iloc[i]
                coord = (
                    round(row['Широта'], 4),
                    round(row['Долгота'], 4)
                )
                current_time = row['Час']

                # Знаходимо тривалість перебування
                j = i + 1
                while j < len(day_data):
                    next_row = day_data.iloc[j]
                    next_coord = (
                        round(next_row['Широта'], 4),
                        round(next_row['Долгота'], 4)
                    )
                    if next_coord != coord:
                        break
                    j += 1

                if j > i:
                    end_time = day_data.iloc[j - 1]['Час']
                    if isinstance(end_time, str):
                        end_time = self.parse_time(end_time)
                    if isinstance(current_time, str):
                        current_time = self.parse_time(current_time)

                    duration = (
                                       datetime.combine(date, end_time) -
                                       datetime.combine(date, current_time)
                               ).total_seconds() / 60

                    min_duration = (
                        min_night_duration
                        if current_time.hour >= 22 or current_time.hour <= 6
                        else min_day_duration
                    )

                    if duration >= min_duration:
                        locations[coord]['total_duration'] += duration
                        locations[coord]['address'] = row['Адреса БС']

                        hour = current_time.hour
                        if 9 <= hour <= 18:  # Робочий час
                            locations[coord]['day_count'] += 1
                        elif hour >= 23 or hour <= 6:  # Нічний час
                            locations[coord]['night_count'] += 1

                i = j if j > i else i + 1

        # Визначаємо дім та роботу
        home = max(
            locations.items(),
            key=lambda x: (x[1]['night_count'], x[1]['total_duration']),
            default=(None, {'address': None, 'night_count': 0, 'total_duration': 0})
        )

        # Виключаємо домашню адресу з пошуку роботи
        work_locations = {
            k: v for k, v in locations.items()
            if k != (home[0] if home[0] is not None else None)
        }

        work = max(
            work_locations.items(),
            key=lambda x: (x[1]['day_count'], x[1]['total_duration']),
            default=(None, {'address': None, 'day_count': 0, 'total_duration': 0})
        )

        return home, work

    def create_map(
            self,
            df: pd.DataFrame,
            date: datetime,
            filename: str,
            output_dir: Optional[str] = None
    ) -> None:
        """
        Створення карти для конкретної дати.

        Args:
            df: DataFrame з даними
            date: Дата для відображення
            filename: Ім'я файлу для збереження карти
            output_dir: Директорія для збереження результатів
        """
        # Фільтруємо дані для вибраної дати
        day_data = df[df['Дата'].dt.date == date.date()]

        if day_data.empty:
            logging.warning(f"Немає даних для дати {date}")
            return

        # Створюємо базову карту
        m = folium.Map(
            location=[
                day_data['Широта'].mean(),
                day_data['Долгота'].mean()
            ],
            zoom_start=12
        )

        # Отримуємо параметри секторів
        try:
            sector_radius = float(self.sector_radius.get())
            sector_angle = float(self.sector_angle.get())
        except (ValueError, AttributeError):
            sector_radius = 500
            sector_angle = 120

        # Додаємо маркери та сектори для кожної точки
        for _, row in day_data.iterrows():
            # Додаємо маркер
            folium.Marker(
                [row['Широта'], row['Довгота']],
                popup=f"Час: {row['Час']}<br>Адреса: {row['Адреса БС']}"
            ).add_to(m)

            # Додаємо сектор
            if 'Азимут' in row:
                try:
                    azimuth = float(row['Азимут'])
                    # Створюємо сектор використовуючи geo_processor
                    sector = self.geo_processor.create_sector(
                        row['Широта'],
                        row['Довгота'],
                        azimuth,
                        sector_angle,
                        sector_radius
                    )

                    # Додаємо сектор на карту
                    folium.GeoJson(
                        sector,
                        style_function=lambda x: {
                            'fillColor': '#3388ff',
                            'color': '#3388ff',
                            'fillOpacity': 0.2,
                            'weight': 1
                        }
                    ).add_to(m)
                except (ValueError, TypeError):
                    logging.warning(f"Неможливо створити сектор для запису: {row}")

        # Додаємо лінії між послідовними точками
        points = day_data[['Широта', 'Долгота']].values.tolist()
        if len(points) > 1:
            folium.PolyLine(
                points,
                weight=2,
                color='blue',
                opacity=0.8
            ).add_to(m)

        # Додаємо тепловую мапу
        heat_data = day_data[['Широта', 'Долгота']].values.tolist()
        HeatMap(heat_data).add_to(m)

        # Якщо є полігон, додаємо його
        if self.polygon:
            folium.GeoJson(
                self.geojson_file,
                name='Полігон'
            ).add_to(m)

        # Додаємо контроль шарів
        folium.LayerControl().add_to(m)

        # Зберігаємо карту
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
            filename = os.path.join(output_dir, filename)

        m.save(filename)

    def _process_files(self) -> None:
        """Обробка файлів для аналізу переміщень."""
        try:
            if not self.traffic_files:
                raise ValueError("Не вибрано файли трафіку")

            # Отримуємо параметри
            try:
                max_distance = float(self.max_distance.get())
                day_min_duration = int(self.day_min_duration.get())
                night_min_duration = int(self.night_min_duration.get())
            except ValueError as e:
                raise ValueError("Неправильний формат параметрів") from e

            # Створюємо директорію для результатів
            output_dir = os.path.join(
                os.path.dirname(self.traffic_files[0]),
                "results"
            )
            os.makedirs(output_dir, exist_ok=True)

            # Обробляємо кожен файл
            total_files = len(self.traffic_files)
            for idx, file in enumerate(self.traffic_files, 1):
                try:
                    # Оновлюємо прогрес
                    self.progress_bar['value'] = (idx / total_files) * 100
                    self.update_idletasks()

                    # Читаємо файл та виводимо наявні колонки
                    df = pd.read_excel(file)
                    logging.info(f"Наявні колонки у файлі: {', '.join(df.columns)}")

                    # Розширений список можливих назв колонок
                    column_mapping = {
                        # Широта
                        'Latitude': 'Широта',
                        'lat': 'Широта',
                        'latitude': 'Широта',
                        'LAT': 'Широта',
                        'широта': 'Широта',
                        'Lat': 'Широта',
                        'latitude_degrees': 'Широта',

                        # Довгота
                        'Longitude': 'Довгота',
                        'lon': 'Довгота',
                        'longitude': 'Довгота',
                        'LON': 'Довгота',
                        'довгота': 'Довгота',
                        'Long': 'Довгота',
                        'lng': 'Довгота',
                        'Долгота': 'Довгота',
                        'longitude_degrees': 'Довгота',

                        # Дата
                        'Date': 'Дата',
                        'date': 'Дата',
                        'DATE': 'Дата',
                        'дата': 'Дата',

                        # Час
                        'Time': 'Час',
                        'time': 'Час',
                        'TIME': 'Час',
                        'час': 'Час',

                        # Адреса
                        'BS Address': 'Адреса БС',
                        'Address': 'Адреса БС',
                        'address': 'Адреса БС',
                        'BS_Address': 'Адреса БС',
                        'адреса': 'Адреса БС',
                        'адреса бс': 'Адреса БС',

                        # Азимут
                        'Azimuth': 'Азимут',
                        'azimuth': 'Азимут',
                        'AZIMUTH': 'Азимут',
                        'азимут': 'Азимут'
                    }

                    # Друкуємо поточні назви колонок для діагностики
                    self.log_text.insert(
                        tk.END,
                        f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                        f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                        f"Current User's Login: {self.current_user}\n"
                        f"Поточні назви колонок у файлі: {', '.join(df.columns)}\n\n"
                    )
                    self.log_text.see(tk.END)

                    # Створюємо словник для перейменування, враховуючи тільки існуючі колонки
                    rename_dict = {}
                    for old_name, new_name in column_mapping.items():
                        matching_cols = [col for col in df.columns if col.strip().lower() == old_name.strip().lower()]
                        if matching_cols:
                            rename_dict[matching_cols[0]] = new_name

                    # Перейменовуємо колонки
                    df = df.rename(columns=rename_dict)

                    # Друкуємо назви колонок після перейменування
                    self.log_text.insert(
                        tk.END,
                        f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                        f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                        f"Current User's Login: {self.current_user}\n"
                        f"Назви колонок після перейменування: {', '.join(df.columns)}\n\n"
                    )
                    self.log_text.see(tk.END)

                    # Перевіряємо наявність необхідних колонок
                    required_columns = ['Дата', 'Час', 'Широта', 'Довгота', 'Адреса БС']
                    missing_columns = [col for col in required_columns if col not in df.columns]

                    if missing_columns:
                        raise ValueError(
                            f"У файлі відсутні необхідні колонки: {', '.join(missing_columns)}\n"
                            f"Наявні колонки: {', '.join(df.columns)}\n"
                            f"Очікувані назви колонок: {', '.join(required_columns)}"
                        )

                    # Конвертуємо дані
                    try:
                        df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', dayfirst=True)
                    except Exception as e:
                        logging.error(f"Помилка конвертації дати: {e}")
                        df['Дата'] = pd.to_datetime(df['Дата'], errors='coerce')

                    df['Час'] = df['Час'].apply(self.parse_time)

                    # Конвертуємо координати у числовий формат
                    try:
                        df['Широта'] = pd.to_numeric(df['Широта'], errors='coerce')
                        df['Довгота'] = pd.to_numeric(df['Довгота'], errors='coerce')

                        if 'Азимут' in df.columns:
                            df['Азимут'] = pd.to_numeric(df['Азимут'], errors='coerce')
                    except Exception as e:
                        raise ValueError(
                            f"Помилка конвертації координат: {str(e)}\n"
                            f"Приклад даних:\nШирота: {df['Широта'].iloc[0]}\n"
                            f"Довгота: {df['Довгота'].iloc[0]}"
                        )

                    # Перевіряємо валідність координат
                    invalid_coords = df[df['Широта'].isna() | df['Довгота'].isna()]
                    if not invalid_coords.empty:
                        self.log_text.insert(
                            tk.END,
                            f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                            f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                            f"Current User's Login: {self.current_user}\n"
                            f"Знайдено {len(invalid_coords)} рядків з невалідними координатами\n\n"
                        )
                        self.log_text.see(tk.END)
                        # Видаляємо рядки з невалідними координатами
                        df = df.dropna(subset=['Широта', 'Довгота'])

                    if df.empty:
                        raise ValueError("Після обробки даних не залишилось валідних записів")

                    # Аналіз місць перебування
                    home, work = self.analyze_locations(
                        df,
                        min_day_duration=day_min_duration,
                        min_night_duration=night_min_duration
                    )

                    # Зберігаємо результати
                    results_filename = os.path.join(
                        output_dir,
                        f"results_{os.path.splitext(os.path.basename(file))[0]}.xlsx"
                    )

                    with pd.ExcelWriter(results_filename) as writer:
                        # Зберігаємо основні дані
                        self._save_location_data(
                            writer,
                            home,
                            work,
                            df,
                            max_distance
                        )

                        # Зберігаємо дані про переміщення поза полігоном
                        if self.polygon:
                            self._save_outside_polygon_data(
                                writer,
                                df
                            )

                    # Створюємо карти
                    if self.create_daily_maps.get():
                        self._create_daily_maps(
                            df,
                            output_dir
                        )

                    # Логуємо результати
                    self._log_processing_results(
                        file,
                        home,
                        work,
                        output_dir
                    )

                    # Додаємо успішне повідомлення в лог
                    self.log_text.insert(
                        tk.END,
                        f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                        f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                        f"Current User's Login: {self.current_user}\n"
                        f"Файл успішно оброблено: {os.path.basename(file)}\n"
                        f"Результати збережено в: {results_filename}\n\n"
                    )
                    self.log_text.see(tk.END)

                except Exception as e:
                    error_msg = (
                        f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                        f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                        f"Current User's Login: {self.current_user}\n"
                        f"Помилка обробки файлу {file}: {str(e)}\n"
                        f"Наявні колонки: {', '.join(df.columns) if 'df' in locals() else 'немає даних'}\n\n"
                    )
                    logging.error(error_msg)
                    self.log_text.insert(tk.END, error_msg)
                    self.log_text.see(tk.END)
                    continue

            self.progress_bar['value'] = 0
            self.update_idletasks()

            # Показуємо повідомлення про успішне завершення
            messagebox.showinfo(
                "Успіх",
                "Обробка файлів завершена успішно"
            )

        except Exception as e:
            error_msg = (
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Помилка: {str(e)}\n\n"
            )
            messagebox.showerror("Помилка", str(e))
            logging.error(error_msg)
            self.log_text.insert(tk.END, error_msg)
            self.log_text.see(tk.END)
            self.progress_bar['value'] = 0
            self.update_idletasks()

    def _save_location_data(
            self,
            writer: pd.ExcelWriter,
            home: tuple,
            work: tuple,
            df: pd.DataFrame,
            max_distance: float
    ) -> None:
        """
        Збереження даних про місця перебування.

        Args:
            writer: ExcelWriter для запису
            home: Дані про місце проживання
            work: Дані про місце роботи
            df: DataFrame з даними
            max_distance: Максимальна відстань
        """
        # Зберігаємо дані про дім
        home_df = pd.DataFrame([{
            'Широта': home[0][0] if home[0] else None,
            'Довгота': home[0][1] if home[0] else None,
            'Кількість нічних відвідувань': home[1]['night_count'],
            'Загальна тривалість (хв)': home[1]['total_duration'],
            'Адреса БС': home[1]['address']
        }])
        home_df.to_excel(writer, sheet_name='Місце проживання', index=False)

        # Зберігаємо дані про роботу
        work_df = pd.DataFrame([{
            'Широта': work[0][0] if work[0] else None,
            'Довгота': work[0][1] if work[0] else None,
            'Кількість денних відвідувань': work[1]['day_count'],
            'Загальна тривалість (хв)': work[1]['total_duration'],
            'Адреса БС': work[1]['address']
        }])
        work_df.to_excel(writer, sheet_name='Місце роботи', index=False)

    def _save_outside_polygon_data(
            self,
            writer: pd.ExcelWriter,
            df: pd.DataFrame
    ) -> None:
        """
        Збереження даних про переміщення поза полігоном.

        Args:
            writer: ExcelWriter для запису
            df: DataFrame з даними
        """
        outside_locations = []
        for _, row in df.iterrows():
            point = Point(row['Довгота'], row['Широта'])
            if not self.polygon.contains(point):
                outside_locations.append({
                    'Дата': row['Дата'],
                    'Час': row['Час'],
                    'Широта': row['Широта'],
                    'Довгота': row['Долгота'],
                    'Адреса БС': row['Адреса БС']
                })

        if outside_locations:
            outside_df = pd.DataFrame(outside_locations)
            outside_df.to_excel(writer, sheet_name='Поза полігоном', index=False)
        else:
            pd.DataFrame({
                'Інформація': ['Немає локацій поза вказаним полігоном']
            }).to_excel(writer, sheet_name='Поза полігоном', index=False)

    def _create_daily_maps(
            self,
            df: pd.DataFrame,
            output_dir: str
    ) -> None:
        """
        Створення карт за кожен день.

        Args:
            df: DataFrame з даними
            output_dir: Директорія для збереження
        """
        for date in df['Дата'].dt.date.unique():
            map_filename = f"map_{date.strftime('%Y%m%d')}.html"
            self.create_map(
                df,
                datetime.combine(date, datetime.min.time()),
                map_filename,
                output_dir=output_dir
            )

    def _log_processing_results(
            self,
            file: str,
            home: tuple,
            work: tuple,
            output_dir: str
    ) -> None:
        """
        Логування результатів обробки.

        Args:
            file: Шлях до файлу
            home: Дані про місце проживання
            work: Дані про місце роботи
            output_dir: Директорія з результатами
        """
        log_message = (
            f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
            f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Current User's Login: {self.current_user}\n"
            f"Оброблено файл: {file}\n"
        )

        if home[0] is not None:
            log_message += (
                f"\nЙмовірне місце проживання:\n"
                f"Координати: {home[0]}\n"
                f"Адреса: {home[1]['address']}\n"
                f"Кількість нічних відвідувань: {home[1]['night_count']}\n"
            )

        if work[0] is not None:
            log_message += (
                f"\nЙмовірне місце роботи:\n"
                f"Координати: {work[0]}\n"
                f"Адреса: {work[1]['address']}\n"
                f"Кількість денних відвідувань: {work[1]['day_count']}\n"
            )

        log_message += f"\nРезультати збережено в: {output_dir}\n\n"

        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)

    def calculate_route_similarity(
            self,
            route1: pd.DataFrame,
            route2: pd.DataFrame,
            max_distance: float = 400
    ) -> float:
        """
        Розрахунок схожості двох маршрутів.

        Args:
            route1: Перший маршрут (DataFrame з координатами)
            route2: Другий маршрут (DataFrame з координатами)
            max_distance: Максимальна відстань між точками (в метрах)

        Returns:
            float: Відсоток схожості (0-100)
        """
        # Мінімальна кількість схожих точок
        min_points = min(len(route1), len(route2)) * 0.5
        similar_points = 0

        for _, point1 in route1.iterrows():
            for _, point2 in route2.iterrows():
                distance = self.calculate_distance(
                    point1['Широта'],
                    point1['Довгота'],
                    point2['Широта'],
                    point2['Довгота']
                )
                if distance <= max_distance:
                    similar_points += 1
                    break

        return (similar_points / len(route1)) * 100

    def _find_similar_routes(self) -> None:
        """Пошук схожих маршрутів."""
        try:
            if not self.traffic_files:
                raise ValueError("Не вибрано файли трафіку")

            # Перевіряємо вхідні дані
            date_str = self.similar_routes_date.get().strip()
            if not date_str:
                raise ValueError("Введіть дату")

            try:
                selected_date = datetime.strptime(date_str, '%d.%m.%Y')
                similarity_threshold = float(self.similarity_threshold.get())
                if not 0 <= similarity_threshold <= 100:
                    raise ValueError(
                        "Відсоток схожості має бути від 0 до 100"
                    )
            except ValueError as e:
                if "time data" in str(e):
                    raise ValueError(
                        "Неправильний формат дати. Використовуйте DD.MM.YYYY"
                    )
                raise ValueError("Неправильний формат відсотка схожості")

            # Створюємо директорію для результатів
            output_dir = os.path.join(
                os.path.dirname(self.traffic_files[0]),
                "results"
            )
            os.makedirs(output_dir, exist_ok=True)

            similar_routes_found = False

            # Оновлюємо прогрес-бар
            self.progress_bar['value'] = 0
            total_files = len(self.traffic_files)

            for file_idx, file in enumerate(self.traffic_files, 1):
                try:
                    # Оновлюємо прогрес
                    self.progress_bar['value'] = (file_idx / total_files) * 100
                    self.update_idletasks()

                    df = pd.read_excel(file)
                    df['Дата'] = pd.to_datetime(
                        df['Дата'],
                        format='%d.%m.%Y',
                        dayfirst=True
                    )
                    df['Час'] = df['Час'].apply(self.parse_time)

                    # Отримуємо маршрут за вказану дату
                    base_route = df[
                        df['Дата'].dt.date == selected_date.date()
                        ].sort_values('Час')

                    if len(base_route) == 0:
                        continue

                    # Шукаємо схожі маршрути
                    similar_routes = []

                    for date in df['Дата'].dt.date.unique():
                        if date == selected_date.date():
                            continue

                        compare_route = df[
                            df['Дата'].dt.date == date
                            ].sort_values('Час')

                        if len(compare_route) == 0:
                            continue

                        similarity = self.calculate_route_similarity(
                            base_route,
                            compare_route,
                            max_distance=float(self.max_distance.get())
                        )

                        if similarity >= similarity_threshold:
                            similar_routes.append({
                                'date': date,
                                'similarity': similarity,
                                'route': compare_route
                            })

                    # Створюємо звіт, якщо знайдено схожі маршрути
                    if similar_routes:
                        similar_routes_found = True
                        self._create_similar_routes_report(
                            similar_routes,
                            base_route,
                            selected_date,
                            output_dir
                        )

                except Exception as e:
                    logging.error(f"Помилка обробки файлу {file}: {e}")
                    continue

            # Скидаємо прогрес-бар
            self.progress_bar['value'] = 0

            if not similar_routes_found:
                self.log_text.insert(
                    tk.END,
                    f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                    f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"Current User's Login: {self.current_user}\n"
                    f"Не знайдено схожих маршрутів для {date_str}\n\n"
                )
                self.log_text.see(tk.END)

        except Exception as e:
            messagebox.showerror("Помилка", str(e))
            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Помилка: {str(e)}\n\n"
            )
            self.log_text.see(tk.END)

    def _create_similar_routes_report(
            self,
            similar_routes: List[Dict],
            base_route: pd.DataFrame,
            selected_date: datetime,
            output_dir: str
    ) -> None:
        """
        Створення звіту про схожі маршрути.

        Args:
            similar_routes: Список схожих маршрутів
            base_route: Базовий маршрут
            selected_date: Вибрана дата
            output_dir: Директорія для збереження
        """
        # Сортуємо за схожістю
        similar_routes.sort(key=lambda x: x['similarity'], reverse=True)

        # Створюємо Excel звіт
        results_filename = os.path.join(
            output_dir,
            f"similar_routes_{selected_date.strftime('%Y%m%d')}.xlsx"
        )

        with pd.ExcelWriter(results_filename) as writer:
            # Зберігаємо базовий маршрут
            base_route.to_excel(
                writer,
                sheet_name=f"Базовий маршрут {selected_date.strftime('%d.%m.%Y')}",
                index=False
            )

            # Зберігаємо схожі маршрути
            for i, route_data in enumerate(similar_routes, 1):
                route_data['route'].to_excel(
                    writer,
                    sheet_name=(
                        f"Схожий маршрут {i} "
                        f"({route_data['date'].strftime('%d.%m.%Y')})"
                    ),
                    index=False
                )

            # Зберігаємо загальну інформацію
            summary_df = pd.DataFrame([
                {
                    'Дата': route_data['date'].strftime('%d.%m.%Y'),
                    'Схожість (%)': round(route_data['similarity'], 2),
                    'Кількість точок': len(route_data['route'])
                }
                for route_data in similar_routes
            ])
            summary_df.to_excel(
                writer,
                sheet_name='Загальна інформація',
                index=False
            )

        # Створюємо карту з маршрутами
        self._create_similar_routes_map(
            similar_routes,
            base_route,
            selected_date,
            output_dir
        )

        # Логуємо результати
        self.log_text.insert(
            tk.END,
            f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
            f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Current User's Login: {self.current_user}\n"
            f"Знайдено {len(similar_routes)} схожих маршрутів "
            f"для {selected_date.strftime('%d.%m.%Y')}\n"
            f"Створено звіт: {results_filename}\n\n"
        )
        self.log_text.see(tk.END)

    def _create_similar_routes_map(
            self,
            similar_routes: List[Dict],
            base_route: pd.DataFrame,
            selected_date: datetime,
            output_dir: str
    ) -> None:
        """
        Створення карти зі схожими маршрутами.

        Args:
            similar_routes: Список схожих маршрутів
            base_route: Базовий маршрут
            selected_date: Вибрана дата
            output_dir: Директорія для збереження
        """
        map_filename = os.path.join(
            output_dir,
            f"similar_routes_map_{selected_date.strftime('%Y%m%d')}.html"
        )

        # Створюємо базову карту
        m = folium.Map(
            location=[
                base_route['Широта'].mean(),
                base_route['Довгота'].mean()
            ],
            zoom_start=12
        )

        # Додаємо базовий маршрут (червоним)
        points = base_route[['Широта', 'Довгота']].values.tolist()
        if len(points) > 1:
            folium.PolyLine(
                points,
                weight=3,
                color='red',
                opacity=0.8,
                popup=f"Базовий маршрут {selected_date.strftime('%d.%m.%Y')}"
            ).add_to(m)

        # Додаємо схожі маршрути (різними кольорами)
        colors = [
            'blue', 'green', 'purple', 'orange', 'darkred',
            'lightred', 'beige', 'darkblue', 'darkgreen', 'cadetblue'
        ]

        for i, route_data in enumerate(similar_routes):
            points = route_data['route'][['Широта', 'Довгота']].values.tolist()
            if len(points) > 1:
                folium.PolyLine(
                    points,
                    weight=2,
                    color=colors[i % len(colors)],
                    opacity=0.6,
                    popup=(
                        f"Схожий маршрут "
                        f"{route_data['date'].strftime('%d.%m.%Y')} "
                        f"(схожість: {round(route_data['similarity'], 2)}%)"
                    )
                ).add_to(m)

        # Додаємо полігон території
        if self.polygon:
            folium.GeoJson(
                self.geojson_file,
                name='Полігон'
            ).add_to(m)

        # Додаємо контроль шарів
        folium.LayerControl().add_to(m)

        # Зберігаємо карту
        m.save(map_filename)

    def _create_map_for_date(self) -> None:
        """Створення карти для обраної дати."""
        try:
            if not self.traffic_files:
                raise ValueError("Не вибрано файли трафіку")

            date_str = self.selected_date.get().strip()
            if not date_str:
                raise ValueError("Введіть дату")

            try:
                selected_date = datetime.strptime(date_str, '%d.%m.%Y')
            except ValueError:
                raise ValueError(
                    "Неправильний формат дати. Використовуйте DD.MM.YYYY"
                )

            # Створюємо директорію для результатів
            output_dir = os.path.join(
                os.path.dirname(self.traffic_files[0]),
                "results"
            )
            os.makedirs(output_dir, exist_ok=True)

            found_data = False
            total_files = len(self.traffic_files)

            for idx, file in enumerate(self.traffic_files, 1):
                try:
                    # Оновлюємо прогрес
                    self.progress_bar['value'] = (idx / total_files) * 100
                    self.update_idletasks()

                    # Читаємо файл
                    df = pd.read_excel(file)
                    df['Дата'] = pd.to_datetime(
                        df['Дата'],
                        format='%d.%m.%Y',
                        dayfirst=True
                    )
                    df['Час'] = df['Час'].apply(self.parse_time)

                    if selected_date.date() in df['Дата'].dt.date.unique():
                        found_data = True

                        # Створюємо карту
                        map_filename = os.path.join(
                            output_dir,
                            f"map_{selected_date.strftime('%Y%m%d')}.html"
                        )

                        # Фільтруємо дані для вибраної дати
                        day_data = df[
                            df['Дата'].dt.date == selected_date.date()
                            ]

                        # Створюємо базову карту
                        m = folium.Map(
                            location=[
                                day_data['Широта'].mean(),
                                day_data['Довгота'].mean()
                            ],
                            zoom_start=12
                        )

                        # Додаємо маркери для кожної точки
                        for _, row in day_data.iterrows():
                            folium.Marker(
                                [row['Широта'], row['Довгота']],
                                popup=(
                                    f"Час: {row['Час']}<br>"
                                    f"Адреса: {row['Адреса БС']}"
                                )
                            ).add_to(m)

                        # Додаємо лінії між послідовними точками
                        points = day_data[
                            ['Широта', 'Довгота']
                        ].values.tolist()

                        if len(points) > 1:
                            folium.PolyLine(
                                points,
                                weight=2,
                                color='blue',
                                opacity=0.8
                            ).add_to(m)

                        # Додаємо теплову карту
                        HeatMap(points).add_to(m)

                        # Якщо є полігон, додаємо його
                        if self.polygon:
                            folium.GeoJson(
                                self.geojson_file,
                                name='Полігон'
                            ).add_to(m)

                        # Додаємо контроль шарів
                        folium.LayerControl().add_to(m)

                        # Зберігаємо карту
                        m.save(map_filename)

                        self.log_text.insert(
                            tk.END,
                            f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                            f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                            f"Current User's Login: {self.current_user}\n"
                            f"Створено карту для {selected_date.strftime('%d.%m.%Y')}\n"
                            f"Файл: {map_filename}\n\n"
                        )
                        self.log_text.see(tk.END)

                except Exception as e:
                    logging.error(f"Помилка обробки файлу {file}: {e}")
                    continue

            # Скидаємо прогрес-бар
            self.progress_bar['value'] = 0

            if not found_data:
                raise ValueError(
                    f"Дані за {date_str} відсутні у вибраних файлах"
                )

        except Exception as e:
            messagebox.showerror("Помилка", str(e))
            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Помилка: {str(e)}\n\n"
            )
            self.log_text.see(tk.END)

    def _create_common_movements_frame(self, parent) -> None:
        """Створення фрейму для аналізу спільних переміщень."""
        frame = ttk.LabelFrame(parent, text="Аналіз спільних переміщень")
        frame.pack(fill=tk.X, pady=2)

        # Перший номер
        number1_frame = ttk.Frame(frame)
        number1_frame.pack(fill=tk.X, pady=2)

        ttk.Label(
            number1_frame,
            text="Номер 1:",
            width=15,
            anchor='e'
        ).pack(side=tk.LEFT, padx=5)

        self.number1_var = tk.StringVar()
        self.number1_combo = ttk.Combobox(
            number1_frame,
            textvariable=self.number1_var,
            width=30,
            state='readonly'
        )
        self.number1_combo.pack(side=tk.LEFT, padx=5)

        # Другий номер
        number2_frame = ttk.Frame(frame)
        number2_frame.pack(fill=tk.X, pady=2)

        ttk.Label(
            number2_frame,
            text="Номер 2:",
            width=15,
            anchor='e'
        ).pack(side=tk.LEFT, padx=5)

        self.number2_var = tk.StringVar()
        self.number2_combo = ttk.Combobox(
            number2_frame,
            textvariable=self.number2_var,
            width=30,
            state='readonly'
        )
        self.number2_combo.pack(side=tk.LEFT, padx=5)

        # Параметри відстані
        distance_frame = ttk.Frame(frame)
        distance_frame.pack(fill=tk.X, pady=2)

        ttk.Label(
            distance_frame,
            text="Макс. відстань (м):",
            width=15,
            anchor='e'
        ).pack(side=tk.LEFT, padx=5)

        self.common_max_distance = ttk.Entry(
            distance_frame,
            width=10
        )
        self.common_max_distance.insert(0, "100")
        self.common_max_distance.pack(side=tk.LEFT, padx=5)

        # Параметри часу
        time_frame = ttk.Frame(frame)
        time_frame.pack(fill=tk.X, pady=2)

        ttk.Label(
            time_frame,
            text="Часове вікно (хв):",
            width=15,
            anchor='e'
        ).pack(side=tk.LEFT, padx=5)

        self.time_window = ttk.Entry(
            time_frame,
            width=10
        )
        self.time_window.insert(0, "30")
        self.time_window.pack(side=tk.LEFT, padx=5)

        # Кнопка аналізу в окремому фреймі для центрування
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X, pady=2)

        ttk.Button(
            button_frame,
            text="Знайти спільні переміщення",
            command=self._analyze_common_movements,
            width=30  # Фіксована ширина кнопки
        ).pack(pady=2)

    def _select_traffic_files(self) -> None:
        """Вибір файлів трафіку та оновлення списків номерів."""
        files = filedialog.askopenfilenames(
            title="Виберіть файли трафіку",
            filetypes=[
                ("Excel файли", "*.xlsx"),
                ("Всі файли", "*.*")
            ]
        )
        if files:
            self.traffic_files = files
            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Вибрано {len(files)} файлів трафіку\n\n"
            )
            self.log_text.see(tk.END)

            # Оновлюємо списки номерів
            self._update_phone_numbers()

    def _update_phone_numbers(self) -> None:
        """Оновлення списків номерів з файлів трафіку."""
        try:
            all_numbers = set()

            # Читаємо всі файли та збираємо унікальні номери
            for file in self.traffic_files:
                try:
                    df = pd.read_excel(file)
                    # Перевіряємо наявність колонки 'Абонент А'
                    if 'Абонент А' in df.columns:
                        # Конвертуємо номери в рядки та видаляємо пробіли
                        numbers = df['Абонент А'].astype(str).str.strip()
                        # Додаємо тільки непусті значення
                        all_numbers.update(
                            num for num in numbers
                            if num and num.lower() != 'nan'
                        )
                    else:
                        logging.warning(f"Не знайдено колонку 'Абонент А' у файлі {file}")

                except Exception as e:
                    logging.error(f"Помилка читання файлу {file}: {e}")
                    continue

            # Перевіряємо чи є номери
            if not all_numbers:
                raise ValueError("Не знайдено жодного номера в файлах")

            # Сортуємо номери
            sorted_numbers = sorted(list(all_numbers))

            # Оновлюємо комбобокси
            if hasattr(self, 'number1_combo') and hasattr(self, 'number2_combo'):
                self.number1_combo['values'] = sorted_numbers
                self.number2_combo['values'] = sorted_numbers

                # Встановлюємо значення за замовчуванням, якщо є хоча б два номери
                if len(sorted_numbers) >= 2:
                    self.number1_var.set(sorted_numbers[0])
                    self.number2_var.set(sorted_numbers[1])
                elif len(sorted_numbers) == 1:
                    self.number1_var.set(sorted_numbers[0])
                    self.number2_var.set("")
                else:
                    self.number1_var.set("")
                    self.number2_var.set("")

                self.log_text.insert(
                    tk.END,
                    f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                    f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"Current User's Login: {self.current_user}\n"
                    f"Знайдено {len(sorted_numbers)} унікальних номерів\n\n"
                )
                self.log_text.see(tk.END)
            else:
                logging.error("Комбобокси не ініціалізовані")

        except Exception as e:
            error_msg = f"Помилка оновлення списку номерів: {str(e)}"
            messagebox.showerror("Помилка", error_msg)
            logging.error(error_msg)

            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Помилка: {error_msg}\n\n"
            )
            self.log_text.see(tk.END)

    def _analyze_common_movements(self) -> None:
        """Аналіз спільних переміщень двох номерів."""
        try:
            # Перевірка наявності файлів
            if not self.traffic_files:
                raise ValueError("Не вибрано файли трафіку")

            # Отримання вибраних номерів
            number1 = self.number1_var.get()
            number2 = self.number2_var.get()

            if not number1 or not number2:
                raise ValueError("Виберіть обидва номери")

            if number1 == number2:
                raise ValueError("Виберіть різні номери")

            try:
                max_distance = float(self.common_max_distance.get())
                time_window = int(self.time_window.get())
            except ValueError:
                raise ValueError(
                    "Неправильний формат відстані або часового вікна"
                )

            # Створюємо директорію для результатів
            output_dir = os.path.join(
                os.path.dirname(self.traffic_files[0]),
                "results"
            )
            os.makedirs(output_dir, exist_ok=True)

            # Зчитуємо та обробляємо дані
            all_data = []
            for file in self.traffic_files:
                df = pd.read_excel(file)
                all_data.append(df)

            # Об'єднуємо всі дані
            combined_df = pd.concat(all_data, ignore_index=True)
            combined_df['Дата'] = pd.to_datetime(
                combined_df['Дата'],
                format='%d.%m.%Y',
                dayfirst=True
            )
            combined_df['Час'] = combined_df['Час'].apply(self.parse_time)

            # Знаходимо дані для кожного номера
            data1 = combined_df[combined_df['Абонент А'] == number1]
            data2 = combined_df[combined_df['Абонент А'] == number2]

            if data1.empty or data2.empty:
                raise ValueError("Дані для одного або обох номерів відсутні")

            # Знаходимо спільні переміщення
            common_movements = self._find_common_movements(
                data1,
                data2,
                max_distance,
                time_window
            )

            if not common_movements:
                messagebox.showinfo(
                    "Результат",
                    "Спільних переміщень не знайдено"
                )
                return

            # Зберігаємо результати
            self._save_common_movements_results(
                common_movements,
                output_dir,
                number1,
                number2
            )

        except Exception as e:
            messagebox.showerror("Помилка", str(e))
            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Помилка: {str(e)}\n\n"
            )
            self.log_text.see(tk.END)

    def _find_common_movements(
            self,
            data1: pd.DataFrame,
            data2: pd.DataFrame,
            max_distance: float,
            time_window: int
    ) -> List[Dict]:
        """
        Пошук спільних переміщень двох номерів.

        Args:
            data1: DataFrame з даними першого номера
            data2: DataFrame з даними другого номера
            max_distance: Максимальна відстань між точками (метри)
            time_window: Часове вікно для порівняння (хвилини)

        Returns:
            List[Dict]: Список спільних переміщень
        """
        common_movements = []

        # Групуємо дані за датою
        for date in data1['Дата'].dt.date.unique():
            day_data1 = data1[data1['Дата'].dt.date == date].sort_values('Час')
            day_data2 = data2[data2['Дата'].dt.date == date].sort_values('Час')

            if day_data2.empty:
                continue

            # Перевіряємо кожну точку першого номера
            for _, point1 in day_data1.iterrows():
                point1_time = datetime.combine(
                    date,
                    point1['Час']
                )

                # Знаходимо відповідні точки другого номера
                for _, point2 in day_data2.iterrows():
                    point2_time = datetime.combine(
                        date,
                        point2['Час']
                    )

                    # Перевіряємо часову різницю
                    time_diff = abs(
                        (point2_time - point1_time).total_seconds() / 60
                    )

                    if time_diff > time_window:
                        continue

                    # Перевіряємо відстань
                    distance = self.calculate_distance(
                        point1['Широта'],
                        point1['Довгота'],
                        point2['Широта'],
                        point2['Долгота']
                    )

                    if distance <= max_distance:
                        common_movements.append({
                            'Дата': date,
                            'Час1': point1['Час'],
                            'Адреса1': point1['Адреса БС'],
                            'Час2': point2['Час'],
                            'Адреса2': point2['Адреса БС'],
                            'Відстань': round(distance, 2),
                            'Часова_різниця': round(time_diff, 2),
                            'Координати1': (point1['Широта'], point1['Довгота']),
                            'Координати2': (point2['Широта'], point2['Довгота'])
                        })

        return common_movements

    def _save_common_movements_results(
            self,
            common_movements: List[Dict],
            output_dir: str,
            number1: str,
            number2: str
    ) -> None:
        """
        Збереження результатів аналізу спільних переміщень.

        Args:
            common_movements: Список спільних переміщень
            output_dir: Директорія для збереження
            number1: Перший номер
            number2: Другий номер
        """
        # Створюємо DataFrame з результатами
        results_df = pd.DataFrame(common_movements)

        # Формуємо ім'я файлу
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_file = os.path.join(
            output_dir,
            f"common_movements_{timestamp}.xlsx"
        )

        # Зберігаємо Excel файл
        with pd.ExcelWriter(excel_file) as writer:
            # Основний лист з даними
            results_df.to_excel(
                writer,
                sheet_name='Спільні переміщення',
                index=False
            )

            # Лист зі статистикою
            stats = {
                'Показник': [
                    'Кількість спільних подій',
                    'Унікальні дати',
                    'Середня відстань (м)',
                    'Середня часова різниця (хв)',
                    'Максимальна відстань (м)',
                    'Максимальна часова різниця (хв)'
                ],
                'Значення': [
                    len(results_df),
                    len(results_df['Дата'].unique()),
                    results_df['Відстань'].mean(),
                    results_df['Часова_різниця'].mean(),
                    results_df['Відстань'].max(),
                    results_df['Часова_різниця'].max()
                ]
            }
            pd.DataFrame(stats).to_excel(
                writer,
                sheet_name='Статистика',
                index=False
            )

        # Створюємо карту
        map_file = os.path.join(
            output_dir,
            f"common_movements_{timestamp}.html"
        )

        # Створюємо базову карту
        center_lat = results_df['Координати1'].apply(lambda x: x[0]).mean()
        center_lon = results_df['Координати1'].apply(lambda x: x[1]).mean()

        m = folium.Map(
            location=[center_lat, center_lon],
            zoom_start=12
        )

        # Додаємо точки та лінії
        for movement in common_movements:
            # Додаємо маркери
            folium.Marker(
                movement['Координати1'],
                popup=(
                    f"Номер 1<br>"
                    f"Час: {movement['Час1']}<br>"
                    f"Адреса: {movement['Адреса1']}"
                ),
                icon=folium.Icon(color='red')
            ).add_to(m)

            folium.Marker(
                movement['Координати2'],
                popup=(
                    f"Номер 2<br>"
                    f"Час: {movement['Час2']}<br>"
                    f"Адреса: {movement['Адреса2']}"
                ),
                icon=folium.Icon(color='blue')
            ).add_to(m)

            # Додаємо лінію між точками
            folium.PolyLine(
                [movement['Координати1'], movement['Координати2']],
                weight=2,
                color='green',
                opacity=0.8,
                popup=(
                    f"Відстань: {movement['Відстань']}м<br>"
                    f"Часова різниця: {movement['Часова_різниця']}хв"
                )
            ).add_to(m)

        # Зберігаємо карту
        m.save(map_file)

        # Логуємо результати
        self.log_text.insert(
            tk.END,
            f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
            f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"Current User's Login: {self.current_user}\n"
            f"Знайдено {len(common_movements)} спільних переміщень\n"
            f"для номерів {number1} та {number2}\n"
            f"Результати збережено:\n"
            f"Excel: {excel_file}\n"
            f"Карта: {map_file}\n\n"
        )
        self.log_text.see(tk.END)

    def _clear_log(self) -> None:
        """Очищення логу операцій."""
        if messagebox.askyesno(
                "Підтвердження",
                "Ви дійсно хочете очистити лог операцій?"
        ):
            self.log_text.delete(1.0, tk.END)
            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Лог очищено\n\n"
            )
            self.log_text.see(tk.END)

    def _save_log(self) -> None:
        """Збереження логу в файл."""
        try:
            file = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[
                    ("Text files", "*.txt"),
                    ("All files", "*.*")
                ],
                title="Зберегти лог",
                initialdir=os.path.dirname(self.traffic_files[0]) if self.traffic_files else None
            )

            if file:
                with open(file, 'w', encoding='utf-8') as f:
                    f.write(self.log_text.get(1.0, tk.END))

                self.log_text.insert(
                    tk.END,
                    f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                    f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"Current User's Login: {self.current_user}\n"
                    f"Лог збережено в файл: {file}\n\n"
                )
                self.log_text.see(tk.END)

                messagebox.showinfo(
                    "Успіх",
                    f"Лог збережено в файл:\n{file}"
                )

        except Exception as e:
            error_msg = f"Помилка збереження логу: {str(e)}"
            messagebox.showerror("Помилка", error_msg)

            self.log_text.insert(
                tk.END,
                f"Current Date and Time (UTC - YYYY-MM-DD HH:MM:SS formatted): "
                f"{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S')}\n"
                f"Current User's Login: {self.current_user}\n"
                f"Помилка: {error_msg}\n\n"
            )
            self.log_text.see(tk.END)