import re
from typing import Optional, List, Dict, Any
from datetime import datetime, time
import pandas as pd

class DataValidator:
    """Клас для валідації даних."""
    
    @staticmethod
    def validate_date(date_str: str) -> Optional[datetime]:
        """
        Валідація дати.
        
        Args:
            date_str: Рядок з датою
            
        Returns:
            Optional[datetime]: Об'єкт datetime або None
        """
        try:
            return pd.to_datetime(date_str, format='%d.%m.%Y')
        except:
            return None
            
    @staticmethod
    def validate_time(time_str: str) -> Optional[time]:
        """
        Валідація часу.
        
        Args:
            time_str: Рядок з часом
            
        Returns:
            Optional[time]: Об'єкт time або None
        """
        try:
            dt = pd.to_datetime(time_str, format='%H:%M:%S')
            return dt.time()
        except:
            return None
            
    @staticmethod
    def validate_coordinates(lat: float, lon: float) -> bool:
        """
        Валідація координат.
        
        Args:
            lat: Широта
            lon: Довгота
            
        Returns:
            bool: True якщо координати валідні
        """
        try:
            return -90 <= float(lat) <= 90 and -180 <= float(lon) <= 180
        except:
            return False
            
    @staticmethod
    def validate_phone(phone: str) -> bool:
        """
        Валідація номера телефону.
        
        Args:
            phone: Номер телефону
            
        Returns:
            bool: True якщо номер валідний
        """
        pattern = r'^\+?\d{10,12}$'
        return bool(re.match(pattern, str(phone)))
        
    @staticmethod
    def validate_required_columns(
        df: pd.DataFrame,
        required_columns: List[str]
    ) -> List[str]:
        """
        Перевірка наявності обов'язкових колонок.
        
        Args:
            df: DataFrame для перевірки
            required_columns: Список обов'язкових колонок
            
        Returns:
            List[str]: Список відсутніх колонок
        """
        return [col for col in required_columns if col not in df.columns]
        
    @staticmethod
    def validate_file_extension(filename: str, allowed_extensions: List[str]) -> bool:
        """
        Перевірка розширення файлу.
        
        Args:
            filename: Ім'я файлу
            allowed_extensions: Дозволені розширення
            
        Returns:
            bool: True якщо розширення дозволене
        """
        return filename.lower().endswith(tuple(allowed_extensions))

class ConfigValidator:
    """Клас для валідації конфігурації."""
    
    @staticmethod
    def validate_config(config: Dict[str, Any]) -> List[str]:
        """
        Валідація конфігурації.
        
        Args:
            config: Словник з конфігурацією
            
        Returns:
            List[str]: Список помилок валідації
        """
        errors = []
        
        # Перевірка обов'язкових секцій
        required_sections = ['app', 'map', 'traffic', 'filters', 'logging']
        for section in required_sections:
            if section not in config:
                errors.append(f"Відсутня обов'язкова секція: {section}")
                
        # Перевірка параметрів додатку
        if 'app' in config:
            app_config = config['app']
            if 'expiration_date' not in app_config:
                errors.append("Відсутня дата закінчення терміну дії")
            if 'encoding' not in app_config:
                errors.append("Відсутнє кодування")
                
        # Перевірка параметрів карти
        if 'map' in config:
            map_config = config['map']
            if 'styles' not in map_config or not map_config['styles']:
                errors.append("Відсутні стилі карти")
            if 'default_style' not in map_config:
                errors.append("Відсутній стиль за замовчуванням")
                
        return errors