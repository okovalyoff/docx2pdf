"""
Фабрика конвертеров.

Автоматически выбирает лучший доступный бэкенд конвертера
на основе операционной системы и установленного программного обеспечения.
"""

import platform
from enum import Enum
from pathlib import Path
from typing import Callable, Tuple

from converters.base import BaseConverter


class ConverterBackend(Enum):
    """Доступные бэкенды конвертера."""
    WORD_COM = "word"           # Windows Word COM интерфейс
    LIBREOFFICE = "libreoffice"  # LibreOffice headless (кроссплатформенный)
    AUTO = "auto"               # Автоматический выбор лучшего


def get_available_backends() -> list[ConverterBackend]:
    """
    Получает список доступных бэкендов конвертера в системе.

    Возвращает бэкенды в порядке предпочтения.

    Возвращает:
        Список доступных бэкендов
    """
    available = []

    # Проверяем LibreOffice (кроссплатформенный, более надёжный)
    try:
        from converters.libreoffice import LibreOfficeConverter
        test_converter = LibreOfficeConverter(
            input_root=Path("."),
            output_root=Path(".")
        )
        if test_converter.is_available:
            available.append(ConverterBackend.LIBREOFFICE)
    except ImportError:
        pass

    # Проверяем Windows Word COM (только Windows, быстрее на Windows)
    if platform.system() == "Windows":
        try:
            from converters.word import WordConverter
            test_converter = WordConverter(
                input_root=Path("."),
                output_root=Path(".")
            )
            if test_converter.is_available:
                # Предпочитаем Word на Windows (быстрее запускается)
                available.insert(0, ConverterBackend.WORD_COM)
        except ImportError:
            pass

    return available


def create_converter(
    backend: ConverterBackend,
    input_root: Path,
    output_root: Path,
    retry_count: int = 3,
    retry_base_delay: float = 2.0,
    retry_max_delay: float = 30.0,
    timeout: int = 120,
    resume: bool = True,
    overwrite: bool = False
) -> BaseConverter:
    """
    Создаёт экземпляр конвертера для указанного бэкенда.

    Аргументы:
        backend: Какой бэкенд конвертера использовать
        input_root: Корневая директория для входных файлов
        output_root: Корневая директория для выходных файлов
        retry_count: Количество повторных попыток при ошибке
        retry_base_delay: Базовая задержка для экспоненциального отступа
        retry_max_delay: Максимальная задержка между попытками
        timeout: Таймаут для одной конвертации
        resume: Пропускать существующие файлы
        overwrite: Перезаписывать существующие файлы

    Возвращает:
        Настроенный экземпляр конвертера

    Исключения:
        RuntimeError: Если запрошенный бэкенд недоступен
    """
    if backend == ConverterBackend.AUTO:
        # Автоматический выбор лучшего доступного
        available = get_available_backends()
        if not available:
            raise RuntimeError(
                "Нет доступного бэкенда конвертера. "
                "Установите LibreOffice или (на Windows) Microsoft Word."
            )
        backend = available[0]

    if backend == ConverterBackend.WORD_COM:
        try:
            from converters.word import WordConverter
            return WordConverter(
                input_root=input_root,
                output_root=output_root,
                retry_count=retry_count,
                retry_base_delay=retry_base_delay,
                retry_max_delay=retry_max_delay,
                timeout=timeout,
                resume=resume,
                overwrite=overwrite
            )
        except ImportError as e:
            raise RuntimeError(
                f"Бэкенд Word COM недоступен: {e}. "
                "Этот бэкенд требует Windows с установленным Microsoft Word."
            )

    elif backend == ConverterBackend.LIBREOFFICE:
        try:
            from converters.libreoffice import LibreOfficeConverter
            return LibreOfficeConverter(
                input_root=input_root,
                output_root=output_root,
                retry_count=retry_count,
                retry_base_delay=retry_base_delay,
                retry_max_delay=retry_max_delay,
                timeout=timeout,
                resume=resume,
                overwrite=overwrite
            )
        except ImportError as e:
            raise RuntimeError(
                f"Бэкенд LibreOffice недоступен: {e}. "
                "Пожалуйста, установите LibreOffice."
            )

    else:
        raise ValueError(f"Неизвестный бэкенд: {backend}")


def get_worker_function(
    backend: ConverterBackend
) -> Callable[
    [Path, Path, Path, int, float, float, int, bool, bool],
    Tuple[str, str, float]
]:
    """
    Получает соответствующую рабочую функцию для ProcessPoolExecutor.

    Рабочие функции — это функции уровня модуля, которые могут быть
    сериализованы для мультипроцессинга.

    Аргументы:
        backend: Какой бэкенд конвертера использовать

    Возвращает:
        Рабочую функцию, принимающую параметры конвертации и возвращающую
        кортеж (статус, сообщение_об_ошибке, время_выполнения).
    """
    if backend == ConverterBackend.AUTO:
        available = get_available_backends()
        if not available:
            raise RuntimeError("Нет доступного бэкенда конвертера")
        backend = available[0]

    if backend == ConverterBackend.WORD_COM:
        from converters.word import convert_with_word
        return convert_with_word

    elif backend == ConverterBackend.LIBREOFFICE:
        from converters.libreoffice import convert_with_libreoffice
        return convert_with_libreoffice

    else:
        raise ValueError(f"Неизвестный бэкенд: {backend}")


def convert_file(
    docx_path: Path,
    input_root: Path,
    output_root: Path,
    retries: int,
    base_delay: int,
    resume: bool,
    overwrite: bool
) -> str:
    """
    Функция конвертации для обратной совместимости.

    Использует автоматически определённый лучший бэкенд.

    Аргументы:
        docx_path: Путь к файлу для конвертации
        input_root: Корневая директория входных файлов
        output_root: Корневая директория выходных файлов
        retries: Количество повторных попыток
        base_delay: Базовая задержка между попытками
        resume: Пропускать существующие файлы
        overwrite: Перезаписывать существующие файлы

    Возвращает:
        Статус конвертации (строка)
    """
    from config import DEFAULT_CONFIG

    converter = create_converter(
        backend=ConverterBackend.AUTO,
        input_root=input_root,
        output_root=output_root,
        retry_count=retries,
        retry_base_delay=float(base_delay),
        retry_max_delay=DEFAULT_CONFIG.retry_max_delay,
        timeout=DEFAULT_CONFIG.conversion_timeout,
        resume=resume,
        overwrite=overwrite
    )

    result = converter.convert(docx_path)
    return result.status.value