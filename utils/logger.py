"""
Модуль конфигурации логирования.

Предоставляет логирование одновременно в файл и консоль с настраиваемой детальностью.
Поддерживает разные уровни логирования и форматированный вывод.
"""

import logging
import sys
from pathlib import Path
from typing import Optional

from config import DEFAULT_CONFIG


class ConsoleFormatter(logging.Formatter):
    """
    Цветной форматтер для вывода в консоль.

    Использует ANSI коды для цветного вывода сообщений.
    """

    # ANSI коды цветов
    COLORS = {
        logging.DEBUG: "\033[36m",     # Голубой
        logging.INFO: "\033[32m",      # Зелёный
        logging.WARNING: "\033[33m",   # Жёлтый
        logging.ERROR: "\033[31m",     # Красный
        logging.CRITICAL: "\033[35m",  # Пурпурный
    }
    RESET = "\033[0m"

    def format(self, record: logging.LogRecord) -> str:
        """
        Форматирует запись лога с цветами.

        Аргументы:
            record: Запись лога для форматирования

        Возвращает:
            Отформатированная строка с цветами
        """
        color = self.COLORS.get(record.levelno, "")
        message = super().format(record)
        return f"{color}{message}{self.RESET}"


class FileFormatter(logging.Formatter):
    """
    Детальный форматтер для вывода в файл.

    Добавляет информацию о процессе и потоке.
    """

    def format(self, record: logging.LogRecord) -> str:
        """
        Форматирует запись лога с детальной информацией.

        Аргументы:
            record: Запись лога для форматирования

        Возвращает:
            Отформатированная строка с деталями
        """
        # Добавляем информацию о процессе
        record.process_info = f"PID:{record.process}"
        return super().format(record)


def setup_logger(
    log_file: Optional[str] = None,
    verbose: bool = False,
    console_output: bool = True
) -> logging.Logger:
    """
    Настраивает конфигурацию логирования.

    Аргументы:
        log_file: Путь к файлу лога (по умолчанию из конфига)
        verbose: Включить детальное (DEBUG) логирование
        console_output: Также выводить в консоль

    Возвращает:
        Настроенный экземпляр логгера
    """
    log_file = log_file or DEFAULT_CONFIG.log_file

    # Создаём логгер
    logger = logging.getLogger()
    logger.handlers.clear()
    logger.setLevel(logging.DEBUG if verbose else logging.INFO)

    # Файловый обработчик (всегда детальный)
    file_handler = logging.FileHandler(
        log_file,
        encoding="utf-8",
        mode="a"  # Режим добавления
    )
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(FileFormatter(
        fmt="%(asctime)s | %(levelname)-8s | %(process_info)s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    ))
    logger.addHandler(file_handler)

    # Консольный обработчик (опционально)
    if console_output:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.DEBUG if verbose else logging.WARNING)
        console_handler.setFormatter(ConsoleFormatter(
            fmt="%(levelname)s: %(message)s"
        ))
        logger.addHandler(console_handler)

    return logger


def get_logger(name: str) -> logging.Logger:
    """
    Получает экземпляр логгера для конкретного модуля.

    Аргументы:
        name: Имя модуля (обычно __name__)

    Возвращает:
        Экземпляр логгера
    """
    return logging.getLogger(name)


class ProgressLogger:
    """
    Контекстный менеджер для логирования прогресса операции.

    Пример использования:
        with ProgressLogger("Конвертация файлов", total=100) as progress:
            for i, file in enumerate(files):
                convert(file)
                progress.update(i + 1)
    """

    def __init__(
        self,
        operation: str,
        total: int,
        logger: Optional[logging.Logger] = None
    ):
        """
        Инициализация логгера прогресса.

        Аргументы:
            operation: Название операции
            total: Общее количество элементов
            logger: Логгер для использования (по умолчанию корневой)
        """
        self.operation = operation
        self.total = total
        self.current = 0
        self.logger = logger or logging.getLogger(__name__)

    def __enter__(self) -> "ProgressLogger":
        """Вход в контекстный менеджер."""
        self.logger.info(f"{self.operation}: начало ({self.total} элементов)")
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Выход из контекстного менеджера."""
        if exc_type is None:
            self.logger.info(f"{self.operation}: завершено ({self.current}/{self.total})")
        else:
            self.logger.error(f"{self.operation}: ошибка на {self.current}/{self.total}")
        return False

    def update(self, current: int, message: Optional[str] = None):
        """
        Обновляет счётчик прогресса.

        Аргументы:
            current: Текущее количество обработанных элементов
            message: Опциональное сообщение для логирования
        """
        self.current = current
        if message:
            self.logger.debug(f"{self.operation}: {current}/{self.total} - {message}")

    def increment(self, message: Optional[str] = None):
        """
        Увеличивает счётчик прогресса на 1.

        Аргументы:
            message: Опциональное сообщение для логирования
        """
        self.update(self.current + 1, message)