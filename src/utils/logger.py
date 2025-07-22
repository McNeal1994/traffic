import logging
from pathlib import Path
from typing import Dict
from datetime import datetime
from .config import Config

def setup_logging(config: Config) -> None:
    """
    Налаштування системи логування.
    
    Args:
        config: Об'єкт конфігурації
    """
    log_config = config.get("logging", {})
    log_file = Path(log_config.get("file", "logs/app.log"))
    
    # Створюємо директорію для логів
    log_file.parent.mkdir(parents=True, exist_ok=True)
    
    # Форматуємо ім'я файлу з датою
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = log_file.parent / f"log_{timestamp}.log"
    
    # Налаштовуємо логування
    logging.basicConfig(
        level=getattr(logging, log_config.get("level", "INFO")),
        format=log_config.get(
            "format", 
            "%(asctime)s - %(levelname)s - %(message)s"
        ),
        handlers=[
            logging.FileHandler(
                log_file,
                encoding=config.get("app.encoding", "utf-8")
            ),
            logging.StreamHandler()
        ]
    )
    
    logging.info("Логування налаштовано")