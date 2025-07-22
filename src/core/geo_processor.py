import folium
from folium import plugins
import pandas as pd
from typing import List, Dict, Optional, Tuple
from pathlib import Path
import logging
from datetime import datetime
from ..utils.config import Config


class GeoProcessor:
    """Клас для обробки геоданих та створення карт."""

    def __init__(self, config: Config):
        """
        Ініціалізація обробника геоданих.

        Args:
            config: Об'єкт конфігурації
        """
        self.config = config
        self.current_time = datetime.strptime(
            "2025-07-18 13:46:47",
            "%Y-%m-%d %H:%M:%S"
        )

    def create_movement_map(
            self,
            df: pd.DataFrame,
            date: datetime,
            show_sectors: bool = False,
            sector_angle: float = 120,
            sector_radius: float = 500,
            map_style: str = "OpenStreetMap"
    ) -> str:
        """
        Створення карти переміщень.

        Args:
            df: DataFrame з даними
            date: Дата для фільтрації
            show_sectors: Показувати сектори
            sector_angle: Кут сектора
            sector_radius: Радіус сектора
            map_style: Стиль карти

        Returns:
            str: Шлях до створеного HTML файлу
        """
        try:
            # Фільтруємо дані за датою
            day_data = df[df['Дата'].dt.date == date.date()].copy()
            if day_data.empty:
                raise ValueError(f"Немає даних для {date.date()}")

            # Створюємо карту
            center_lat = day_data['Широта'].mean()
            center_lon = day_data['Долгота'].mean()
            m = folium.Map(
                location=[center_lat, center_lon],
                zoom_start=12,
                tiles=map_style
            )

            # Додаємо маркери та лінії переміщення
            self._add_movement_markers(m, day_data)

            if show_sectors:
                self._add_sectors(
                    m,
                    day_data,
                    sector_angle,
                    sector_radius
                )

            # Зберігаємо карту
            return self._save_map(m, f"movement_{date.strftime('%Y%m%d')}")

        except Exception as e:
            logging.error(f"Помилка створення карти переміщень: {e}")
            raise

    def create_heatmap(
            self,
            df: pd.DataFrame,
            radius: int = 15,
            map_style: str = "OpenStreetMap"
    ) -> str:
        """
        Створення теплової карти.

        Args:
            df: DataFrame з даними
            radius: Радіус точок теплової карти
            map_style: Стиль карти

        Returns:
            str: Шлях до створеного HTML файлу
        """
        try:
            # Створюємо базову карту
            center_lat = df['Широта'].mean()
            center_lon = df['Долгота'].mean()
            m = folium.Map(
                location=[center_lat, center_lon],
                zoom_start=12,
                tiles=map_style
            )

            # Готуємо дані для теплової карти
            heat_data = [[row['Широта'], row['Долгота']]
                         for _, row in df.iterrows()]

            # Додаємо тепловий шар
            plugins.HeatMap(
                heat_data,
                radius=radius,
                blur=15,
                max_zoom=13
            ).add_to(m)

            # Зберігаємо карту
            return self._save_map(m, "heatmap")

        except Exception as e:
            logging.error(f"Помилка створення теплової карти: {e}")
            raise

    def _add_movement_markers(self, m: folium.Map, df: pd.DataFrame) -> None:
        """
        Додавання маркерів та ліній переміщення на карту.

        Args:
            m: Об'єкт карти
            df: DataFrame з даними
        """
        try:
            # Сортуємо дані за часом
            sorted_df = df.sort_values(['Абонент А', 'Час'])

            # Створюємо групи за абонентами
            for _, group in sorted_df.groupby('Абонент А'):
                coordinates = []

                for _, row in group.iterrows():
                    # Додаємо маркер
                    popup_text = (
                        f"Час: {row['Час']}<br>"
                        f"Адреса: {row['Адреса БС']}<br>"
                        f"Абонент: {row['Абонент А']}"
                    )

                    folium.Marker(
                        [row['Широта'], row['Долгота']],
                        popup=popup_text,
                        icon=folium.Icon(color='red')
                    ).add_to(m)

                    coordinates.append([row['Широта'], row['Долгота']])

                # Додаємо лінію переміщення
                if len(coordinates) > 1:
                    folium.PolyLine(
                        coordinates,
                        weight=2,
                        color='blue',
                        opacity=0.8
                    ).add_to(m)

        except Exception as e:
            logging.error(f"Помилка додавання маркерів: {e}")
            raise

    def _add_sectors(
            self,
            m: folium.Map,
            df: pd.DataFrame,
            angle: float,
            radius: float
    ) -> None:
        """
        Додавання секторів на карту.

        Args:
            m: Об'єкт карти
            df: DataFrame з даними
            angle: Кут сектора
            radius: Радіус сектора
        """
        try:
            for _, row in df.iterrows():
                if 'Аз.' in row:
                    # Створюємо сектор
                    folium.Sector(
                        location=[row['Широта'], row['Долгота']],
                        radius=radius,
                        start_angle=row['Аз.'] - angle / 2,
                        end_angle=row['Аз.'] + angle / 2,
                        color='red',
                        fill=True,
                        opacity=0.4
                    ).add_to(m)

        except Exception as e:
            logging.error(f"Помилка додавання секторів: {e}")
            raise

    def _save_map(self, m: folium.Map, prefix: str) -> str:
        """
        Збереження карти у файл.

        Args:
            m: Об'єкт карти
            prefix: Префікс імені файлу

        Returns:
            str: Шлях до збереженого файлу
        """
        try:
            output_dir = Path("output/maps")
            output_dir.mkdir(parents=True, exist_ok=True)

            timestamp = self.current_time.strftime("%Y%m%d_%H%M%S")
            output_file = output_dir / f"{prefix}_{timestamp}.html"
            m.save(str(output_file))

            return str(output_file)

        except Exception as e:
            logging.error(f"Помилка збереження карти: {e}")
            raise