"""
Пакет конвертеров документов.

Содержит базовый класс и реализации конвертеров для различных бэкендов.

Примеры использования:
    from converters import WordConverter, LibreOfficeConverter
    from converters import create_converter, get_available_backends
"""

from converters.base import BaseConverter, ConversionError, ConversionStatus, ConversionResult
from converters.factory import (
    ConverterBackend,
    create_converter,
    get_available_backends,
    get_worker_function,
    convert_file
)

__all__ = [
    # Базовый класс и типы
    "BaseConverter",
    "ConversionError",
    "ConversionStatus",
    "ConversionResult",
    # Фабрика
    "ConverterBackend",
    "create_converter",
    "get_available_backends",
    "get_worker_function",
    "convert_file",
]