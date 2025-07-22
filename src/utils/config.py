import yaml
import logging
from datetime import datetime
from typing import Any, Optional


class Config:
    def __init__(self, config_file: str = "config.yaml"):
        """
        Ініціалізація конфігурації.

        Args:
            config_file: Шлях до файлу конфігурації
        """
        self.config_file = config_file
        self.config = self._load_config()

        # Встановлюємо значення за замовчуванням, якщо їх немає
        self._set_defaults()

    def _load_config(self) -> dict:
        """
        Завантаження конфігурації з файлу.

        Returns:
            dict: Конфігурація
        """
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f) or {}
        except Exception as e:
            logging.error(f"Помилка завантаження конфігурації: {e}")
            return {}

    def _set_defaults(self):
        """Встановлення значень за замовчуванням."""
        defaults = {
            "app": {
                "name": "AnalyzeR",
                "version": "2.0.0",
                "expiration_date": "2026-12-31",  # Термін дії за замовчуванням
                "logging": {
                    "level": "INFO",
                    "format": "%(asctime)s - %(levelname)s - %(message)s"
                }
            },
            "traffic": {
                "required_columns": [
                    "Адреса БС",
                    "Абонент А",
                    "Дата",
                    "Час"
                ],
                "similarity_threshold": 90,
                "max_distance": 400,
                "time_window": 30
            },
            "database": {
                "path": "addresses.db",
                "backup_path": "backups/"
            }
        }

        # Рекурсивне оновлення конфігурації
        self._update_recursive(self.config, defaults)

        # Зберігаємо оновлену конфігурацію
        self._save_config()

    def _update_recursive(self, current: dict, default: dict):
        """
        Рекурсивне оновлення словника конфігурації.

        Args:
            current: Поточний словник
            default: Словник зі значеннями за замовчуванням
        """
        for key, value in default.items():
            if key not in current:
                current[key] = value
            elif isinstance(value, dict) and isinstance(current[key], dict):
                self._update_recursive(current[key], value)

    def _save_config(self):
        """Збереження конфігурації у файл."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                yaml.safe_dump(self.config, f, allow_unicode=True)
        except Exception as e:
            logging.error(f"Помилка збереження конфігурації: {e}")

    def get(self, path: str, default: Any = None) -> Any:
        """
        Отримання значення з конфігурації за шляхом.

        Args:
            path: Шлях до значення (наприклад "app.name")
            default: Значення за замовчуванням

        Returns:
            Any: Значення з конфігурації
        """
        current = self.config
        try:
            for key in path.split('.'):
                current = current[key]
            return current
        except (KeyError, TypeError):
            return default

    def set(self, path: str, value: Any):
        """
        Встановлення значення в конфігурації.

        Args:
            path: Шлях до значення (наприклад "app.name")
            value: Нове значення
        """
        current = self.config
        keys = path.split('.')
        for key in keys[:-1]:
            current = current.setdefault(key, {})
        current[keys[-1]] = value
        self._save_config()

    def check_expiration(self) -> bool:
        """
        Перевірка терміну дії програми.

        Returns:
            bool: True якщо термін дії не закінчився
        """
        try:
            expiration_date = datetime.strptime(
                self.get("app.expiration_date"),
                "%Y-%m-%d"
            )
            current_date = datetime.now()

            if current_date > expiration_date:
                logging.error("Термін дії програми закінчився")
                return False

            days_left = (expiration_date - current_date).days
            if days_left <= 30:
                logging.warning(
                    f"Залишилось {days_left} днів до закінчення терміну дії"
                )

            return True

        except Exception as e:
            logging.error(f"Помилка перевірки терміну дії: {e}")
            return False