#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import logging
from pathlib import Path
from datetime import datetime
from src.utils.config import Config
from src.utils.logger import setup_logging
from src.gui.main_window import MainWindow
import tkinter as tk

def main():
    """Головна точка входу програми."""
    try:
        # Ініціалізуємо конфігурацію
        config = Config()
        
        # Перевіряємо термін дії
        config.check_expiration()
        
        # Налаштовуємо логування
        setup_logging(config)
        
        # Створюємо головне вікно
        root = tk.Tk()
        app = MainWindow(root, config)
        
        # Запускаємо програму
        logging.info(f"Запуск програми користувачем {config.get('user.login', 'Unknown')}")
        root.mainloop()
        
    except Exception as e:
        logging.critical(f"Критична помилка: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()