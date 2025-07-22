"""
Модуль обробки даних для аналізатора трафіку.
"""
import tkinter as tk
from datetime import datetime, timedelta, time
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import pandas as pd
import re
import logging
import os
from typing import Optional, List, Dict, Any, Tuple
from pathlib import Path


class DataProcessor:
    """Клас для обробки даних трафіку."""

    def __init__(self, config: 'Config'):
        """
        Ініціалізація обробника даних.

        Args:
            config: Об'єкт конфігурації
        """
        self.config = config
        self._lock = False
        self.current_time = datetime.strptime(
            "2025-07-21 15:21:03",
            "%Y-%m-%d %H:%M:%S"
        )
        self.current_user = "McNeal1994"
        logging.info(f"Запуск програми користувачем {self.current_user}")

    def merge_traffic_files(self) -> Tuple[bool, str]:
        """
        Об'єднання файлів трафіку.
        Створює два аркуші:
        1. Всі дані
        2. Дані без GPRS з'єднань

        Returns:
            Tuple[bool, str]: (успіх, повідомлення)
        """
        try:
            # Отримуємо поточний час та користувача
            current_time = "2025-07-21 11:56:37"
            current_user = "McNeal1994"

            # Вибір файлів для об'єднання
            files = filedialog.askopenfilenames(
                title="Виберіть файли трафіку для об'єднання",
                filetypes=[
                    ("Excel файли", "*.xlsx"),
                    ("Всі файли", "*.*")
                ]
            )

            if not files:
                return False, "Файли не вибрано"

            # Створюємо директорію для результатів
            output_dir = os.path.dirname(files[0])
            results_dir = os.path.join(output_dir, "results")
            os.makedirs(results_dir, exist_ok=True)

            # Читаємо та об'єднуємо файли
            all_data = []
            file_stats = []

            for file in files:
                try:
                    df = pd.read_excel(file)
                    if not df.empty:
                        # Додаємо метадані
                        df['_source_file'] = os.path.basename(file)
                        df['_merge_time'] = current_time
                        df['_merge_user'] = current_user

                        all_data.append(df)
                        file_stats.append({
                            'Файл': os.path.basename(file),
                            'Записів': len(df),
                            'Розмір (байт)': os.path.getsize(file),
                            'Дата об\'єднання': current_time,
                            'Користувач': current_user
                        })

                        logging.info(f"Прочитано {file}: {len(df)} записів")

                except Exception as e:
                    logging.error(f"Помилка читання {file}: {str(e)}")
                    continue

            if not all_data:
                return False, "Не вдалося прочитати жоден файл"

            # Об'єднуємо всі дані
            merged_df = pd.concat(all_data, ignore_index=True)

            # Створюємо DataFrame без GPRS
            no_gprs_df = merged_df[
                ~merged_df['Тип'].str.contains('GPRS', case=False, na=False)
            ].copy()

            # Формуємо ім'я вихідного файлу
            timestamp = current_time.replace(":", "").replace(" ", "_")
            output_file = os.path.join(
                results_dir,
                f"merged_traffic_{timestamp}.xlsx"
            )

            # Створюємо Excel файл з трьома аркушами
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Аркуш 1: Всі дані
                merged_df.to_excel(
                    writer,
                    sheet_name='Всі дані',
                    index=False
                )

                # Аркуш 2: Дані без GPRS
                no_gprs_df.to_excel(
                    writer,
                    sheet_name='Дані без GPRS',
                    index=False
                )

                # Аркуш 3: Статистика
                stats_df = pd.DataFrame(file_stats)

                # Додаємо підсумкову статистику
                summary_stats = pd.DataFrame([{
                    'Файл': 'ВСЬОГО',
                    'Записів': len(merged_df),
                    'Записів без GPRS': len(no_gprs_df),
                    'Дата об\'єднання': current_time,
                    'Користувач': current_user
                }])

                stats_df = pd.concat([stats_df, summary_stats], ignore_index=True)
                stats_df.to_excel(
                    writer,
                    sheet_name='Статистика',
                    index=False
                )

            # Формуємо повідомлення про успіх
            success_msg = (
                f"Файли успішно об'єднано!\n\n"
                f"Час: {current_time}\n"
                f"Користувач: {current_user}\n\n"
                f"Оброблено файлів: {len(files)}\n"
                f"Загальна кількість записів: {len(merged_df)}\n"
                f"Записів без GPRS: {len(no_gprs_df)}\n"
                f"Результат збережено:\n{output_file}"
            )

            messagebox.showinfo("Успіх", success_msg)
            logging.info(
                f"Файли об'єднано успішно: {output_file}\n"
                f"Всього записів: {len(merged_df)}\n"
                f"Записів без GPRS: {len(no_gprs_df)}"
            )

            return True, success_msg

        except Exception as e:
            error_msg = f"Помилка об'єднання файлів: {str(e)}"
            logging.error(error_msg)
            messagebox.showerror("Помилка", error_msg)
            return False, error_msg

    def preprocess_traffic_file(self, df: pd.DataFrame) -> pd.DataFrame:
        """Обробка файлу трафіку для приведення до стандартного формату."""
        try:
            logging.info(f"Початкові стовпці файлу: {df.columns.tolist()}")

            # Якщо стовпці 'Дата' і 'Час' вже існують, використовуємо їх
            if 'Дата' in df.columns and 'Час' in df.columns:
                logging.info("Знайдено існуючі стовпці 'Дата' та 'Час'")
                return df

            # Перевіряємо чи всі дані в одному стовпці
            if len(df.columns) == 1 or (
                    len(df.columns) > 1 and all(col.startswith('Unnamed:') for col in df.columns[1:])
            ):
                logging.info("Виявлено дані в одному стовпці, виконується розділення")

                # Беремо перший стовпець
                first_col = df.iloc[:, 0].astype(str)

                # Замінюємо множинні табуляції на одну
                first_col = first_col.apply(lambda x: re.sub(r'\t+', '\t', x))

                # Розділяємо дані по табуляції
                split_data = [row.split('\t') for row in first_col]

                # Визначаємо максимальну кількість стовпців
                max_cols = max(len(row) for row in split_data)

                # Доповнюємо рядки, які мають менше стовпців
                split_data = [row + [''] * (max_cols - len(row)) for row in split_data]

                # Створюємо новий DataFrame
                df = pd.DataFrame(split_data)

                # Видаляємо порожні стовпці
                df = df.replace(['', 'nan', 'None'], pd.NA).dropna(axis=1, how='all')

                logging.info(f"Після розділення отримано стовпців: {len(df.columns)}")

            # Видаляємо порожні рядки
            df = df.replace(['', 'nan', 'None'], pd.NA).dropna(how='all')

            # Шукаємо стовпці з датами та часом
            date_col = None
            time_col = None

            # Спочатку шукаємо за назвами стовпців
            for col in df.columns:
                col_name = str(col).lower()
                if 'дата' in col_name:
                    date_col = col
                elif 'час' in col_name:
                    time_col = col

            # Якщо не знайдено, шукаємо за вмістом
            if date_col is None or time_col is None:
                for col in df.columns:
                    sample_values = df[col].dropna().astype(str).head(10)

                    # Шукаємо дату
                    if date_col is None:
                        for val in sample_values:
                            try:
                                if pd.to_datetime(val, format='%d.%m.%Y', errors='coerce') is not None:
                                    date_col = col
                                    break
                            except:
                                continue

                    # Шукаємо час
                    if time_col is None:
                        for val in sample_values:
                            try:
                                if ':' in val and len(val.split(':')) in [2, 3]:
                                    time_col = col
                                    break
                            except:
                                continue

            if date_col is None or time_col is None:
                raise ValueError(f"Не знайдено стовпці з датою та часом. Наявні стовпці: {df.columns.tolist()}")

            # Перейменовуємо знайдені стовпці
            if date_col != 'Дата':
                df = df.rename(columns={date_col: 'Дата'})
            if time_col != 'Час':
                df = df.rename(columns={time_col: 'Час'})

            # Видаляємо рядки з невалідними датами або часом
            df = df.dropna(subset=['Дата', 'Час'])

            if df.empty:
                raise ValueError("Після обробки не залишилось валідних даних")

            logging.info(f"Успішно оброблено файл. Знайдено {len(df)} рядків з датами та часом")
            return df

        except Exception as e:
            logging.error(f"Помилка обробки файлу: {str(e)}")
            raise

    def filter_traffic_by_datetime(self, traffic_files: List[str], filter_file: str,
                                   time_window_before: int, time_window_after: int,
                                   progress_bar: ttk.Progressbar,
                                   root: tk.Tk, output_dir: str) -> Tuple[Optional[str], Optional[pd.DataFrame]]:
        """
        Фільтрація трафіку за датою і часом з гнучким пошуком та асиметричним вікном.

        Args:
            traffic_files: Список файлів з трафіком
            filter_file: Файл з фільтрами
            time_window_before: Часове вікно до події (хвилини)
            time_window_after: Часове вікно після події (хвилини)
            progress_bar: Віджет прогрес-бару
            root: Кореневий віджет
            output_dir: Директорія для збереження результату

        Returns:
            Tuple[str, pd.DataFrame]: Шлях до збереженого файлу та DataFrame з результатами
        """
        try:
            if not traffic_files:
                logging.error("Не вибрано файли трафіку")
                raise ValueError("Не вибрано файли трафіку")
            if not filter_file or not os.path.exists(filter_file):
                logging.error("Файл фільтрів не вказано")
                raise ValueError("Файл фільтрів не вказано")

            # Читаємо файл фільтрів
            filter_df = pd.read_excel(filter_file)
            logging.info(f"Завантажено файл фільтрів: {len(filter_df)} рядків")
            logging.info(f"Стовпці у файлі фільтрів: {filter_df.columns.tolist()}")

            # Конвертуємо дати та час у файлі фільтрів
            filter_df['Дата'] = pd.to_datetime(filter_df['Дата'], format='%d.%m.%Y', errors='coerce')
            filter_df['Час'] = pd.to_datetime(filter_df['Час'].astype(str).apply(lambda x: x.strip()),
                                              format='%H:%M:%S', errors='coerce').dt.time
            filter_df = filter_df.dropna(subset=['Дата', 'Час'])

            if filter_df.empty:
                logging.error("Файл фільтрів не містить валідних даних")
                raise ValueError("Файл фільтрів не містить валідних даних")

            # Створюємо datetime для фільтрів
            filter_df['datetime'] = pd.to_datetime(
                filter_df['Дата'].dt.strftime('%Y-%m-%d') + ' ' +
                filter_df['Час'].apply(lambda x: x.strftime('%H:%M:%S'))
            )

            progress_bar['maximum'] = len(traffic_files)
            progress_bar['value'] = 0
            filtered_dfs = []

            # Статистика для звіту
            stats = {
                'total_matches': 0,
                'in_window_matches': 0,
                'nearest_matches': 0,
                'files_processed': 0,
                'files_with_matches': 0,
                'before_window_matches': 0,
                'after_window_matches': 0
            }

            for idx, traffic_file in enumerate(traffic_files):
                try:
                    if not os.path.exists(traffic_file):
                        logging.warning(f"Файл {traffic_file} не знайдено")
                        continue

                    df = pd.read_excel(traffic_file)
                    logging.info(f"Читання файлу {traffic_file}. Стовпці: {df.columns.tolist()}")

                    # Конвертуємо дати та час у файлі трафіку
                    df['Дата'] = pd.to_datetime(df['Дата'], format='%d.%m.%Y', errors='coerce')
                    df['Час'] = pd.to_datetime(df['Час'].astype(str).apply(lambda x: x.strip()),
                                               format='%H:%M:%S', errors='coerce').dt.time
                    df = df.dropna(subset=['Дата', 'Час'])

                    if df.empty:
                        logging.warning(f"Файл {traffic_file} не містить валідних даних")
                        continue

                    # Створюємо datetime для порівняння
                    df['datetime'] = pd.to_datetime(
                        df['Дата'].dt.strftime('%Y-%m-%d') + ' ' +
                        df['Час'].apply(lambda x: x.strftime('%H:%M:%S'))
                    )

                    stats['files_processed'] += 1
                    file_has_matches = False

                    # Шукаємо співпадіння для кожного фільтру
                    matched_dfs = []
                    for _, filter_row in filter_df.iterrows():
                        filter_dt = filter_row['datetime']

                        # Створюємо часові межі
                        time_before = filter_dt - timedelta(minutes=time_window_before)
                        time_after = filter_dt + timedelta(minutes=time_window_after)

                        # Шукаємо записи в межах вікна
                        mask = (df['datetime'] >= time_before) & (df['datetime'] <= time_after)
                        matches_in_window = df[mask].copy()

                        if not matches_in_window.empty:
                            # Якщо є кілька співпадінь у вікні, знаходимо найближче
                            time_diffs = (matches_in_window['datetime'] - filter_dt).abs()
                            closest_idx = time_diffs.idxmin()
                            match = matches_in_window.loc[[closest_idx]].copy()

                            # Визначаємо тип співпадіння (до чи після)
                            time_diff = (match['datetime'].iloc[0] - filter_dt).total_seconds() / 60
                            if time_diff < 0:
                                match_type = 'До події'
                                stats['before_window_matches'] += 1
                            else:
                                match_type = 'Після події'
                                stats['after_window_matches'] += 1

                            match['ID'] = filter_row['ID']
                            match['Відхилення (хв)'] = abs(time_diff)
                            match['Тип співпадіння'] = f'В межах вікна ({match_type})'
                            matched_dfs.append(match)

                            stats['in_window_matches'] += 1
                            file_has_matches = True

                        else:
                            # Якщо нема співпадінь у вікні, шукаємо найближче значення
                            time_diffs = (df['datetime'] - filter_dt).abs()
                            closest_idx = time_diffs.idxmin()
                            closest_match = df.loc[[closest_idx]].copy()

                            time_diff = (closest_match['datetime'].iloc[0] - filter_dt).total_seconds() / 60
                            if time_diff < 0:
                                match_type = 'До події'
                            else:
                                match_type = 'Після події'

                            closest_match['ID'] = filter_row['ID']
                            closest_match['Відхилення (хв)'] = abs(time_diff)
                            closest_match['Тип співпадіння'] = f'Поза вікном ({match_type})'
                            matched_dfs.append(closest_match)

                            stats['nearest_matches'] += 1
                            file_has_matches = True

                            logging.info(
                                f"Для фільтра ID={filter_row['ID']} використано найближче з'єднання. "
                                f"Відхилення: {abs(time_diff):.2f} хв. {match_type}"
                            )

                    if matched_dfs:
                        filtered_df = pd.concat(matched_dfs, ignore_index=True)
                        # Сортуємо стовпці
                        cols = ['ID', 'Тип співпадіння', 'Відхилення (хв)'] + [
                            col for col in filtered_df.columns
                            if col not in ['ID', 'Тип співпадіння', 'Відхилення (хв)']
                        ]
                        filtered_df = filtered_df[cols]
                        filtered_dfs.append(filtered_df)

                        if file_has_matches:
                            stats['files_with_matches'] += 1
                            stats['total_matches'] += len(filtered_df)

                    progress_bar['value'] = idx + 1
                    root.update_idletasks()

                except Exception as e:
                    logging.error(f"Помилка обробки файлу {traffic_file}: {str(e)}")
                    continue

            if not filtered_dfs:
                logging.info("Не знайдено рядків за фільтром")
                return None, None

            # Об'єднуємо всі результати
            result_df = pd.concat(filtered_dfs, ignore_index=True)

            # Створюємо звіт зі статистикою
            stats_df = pd.DataFrame([{
                'Параметр': 'Часове вікно до події (хв)',
                'Значення': time_window_before
            }, {
                'Параметр': 'Часове вікно після події (хв)',
                'Значення': time_window_after
            }, {
                'Параметр': 'Оброблено файлів',
                'Значення': stats['files_processed']
            }, {
                'Параметр': 'Файлів зі співпадіннями',
                'Значення': stats['files_with_matches']
            }, {
                'Параметр': 'Всього співпадінь',
                'Значення': stats['total_matches']
            }, {
                'Параметр': 'Співпадінь в межах вікна',
                'Значення': stats['in_window_matches']
            }, {
                'Параметр': 'Співпадінь до події',
                'Значення': stats['before_window_matches']
            }, {
                'Параметр': 'Співпадінь після події',
                'Значення': stats['after_window_matches']
            }, {
                'Параметр': 'Найближчих співпадінь поза вікном',
                'Значення': stats['nearest_matches']
            }])

            # Зберігаємо результати
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_file = os.path.join(
                output_dir,
                f"filtered_traffic_{time_window_before}min_before_{time_window_after}min_after_{timestamp}.xlsx"
            )

            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                result_df.to_excel(writer, sheet_name='Результати', index=False)
                stats_df.to_excel(writer, sheet_name='Статистика', index=False)

            logging.info(f"Збережено відфільтрований файл: {output_file}")
            logging.info(
                f"Статистика фільтрації:\n"
                f"- Часове вікно до події: {time_window_before} хв\n"
                f"- Часове вікно після події: {time_window_after} хв\n"
                f"- Оброблено файлів: {stats['files_processed']}\n"
                f"- Файлів зі співпадіннями: {stats['files_with_matches']}\n"
                f"- Всього співпадінь: {stats['total_matches']}\n"
                f"- В межах вікна: {stats['in_window_matches']}\n"
                f"  - До події: {stats['before_window_matches']}\n"
                f"  - Після події: {stats['after_window_matches']}\n"
                f"- Поза вікном: {stats['nearest_matches']}"
            )

            return output_file, result_df

        except Exception as e:
            logging.error(f"Помилка фільтрації: {str(e)}")
            raise ValueError(f"Помилка фільтрації: {str(e)}")
