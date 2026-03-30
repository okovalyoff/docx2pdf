"""
Модуль сканирования файлов.

Сканирует директории на наличие DOCX/DOC файлов для конвертации.
Поддерживает фильтрацию по расширению и сортировку.
"""

import logging
from pathlib import Path
from typing import Iterator, Optional

from config import DEFAULT_CONFIG


def scan_docx_files(
    root: Path,
    extensions: Optional[tuple] = None,
    include_hidden: bool = False,
    sort_by_size: bool = False
) -> list[Path]:
    """
    Сканирует директорию на наличие документов Word.

    Аргументы:
        root: Корневая директория для сканирования
        extensions: Расширения файлов для включения (по умолчанию: .docx, .doc)
        include_hidden: Включать скрытые файлы и директории
        sort_by_size: Сортировать по размеру (сначала маленькие для быстрой обратной связи)

    Возвращает:
        Список объектов Path для найденных файлов

    Исключения:
        FileNotFoundError: Если корневая директория не существует
        NotADirectoryError: Если корневой путь не является директорией
    """
    logger = logging.getLogger(__name__)

    # Проверяем корневую директорию
    if not root.exists():
        raise FileNotFoundError(f"Директория не найдена: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"Не является директорией: {root}")

    # Используем расширения по умолчанию
    extensions = extensions or DEFAULT_CONFIG.input_extensions

    # Нормализуем расширения (убеждаемся, что они начинаются с .)
    extensions = tuple(
        ext if ext.startswith(".") else f".{ext}"
        for ext in extensions
    )

    # Сканируем файлы
    files = []
    for ext in extensions:
        pattern = f"*{ext}"

        if include_hidden:
            # Включаем скрытые файлы
            for file_path in root.rglob(pattern):
                files.append(file_path)
        else:
            # Пропускаем скрытые файлы и директории
            for file_path in root.rglob(pattern):
                # Проверяем, что ни одна часть пути не скрыта
                if not any(part.startswith(".") for part in file_path.parts):
                    files.append(file_path)

    # Удаляем дубликаты (на случай перекрывающихся шаблонов)
    files = list(set(files))

    # Сортируем
    if sort_by_size:
        # Сортируем по размеру файла (сначала маленькие) для быстрой обратной связи
        files.sort(key=lambda f: f.stat().st_size if f.exists() else 0)
    else:
        # Сортируем по имени для консистентного порядка
        files.sort()

    logger.debug(f"Найдено {len(files)} файлов в {root}")

    return files


def scan_docx_files_iter(
    root: Path,
    extensions: Optional[tuple] = None,
    include_hidden: bool = False
) -> Iterator[Path]:
    """
    Сканирует директорию на наличие документов Word (версия-генератор).

    Экономичная по памяти версия для очень больших директорий.

    Аргументы:
        root: Корневая директория для сканирования
        extensions: Расширения файлов для включения
        include_hidden: Включать скрытые файлы и директории

    Возвращает:
        Генератор объектов Path для найденных файлов
    """
    # Проверяем корневую директорию
    if not root.exists():
        raise FileNotFoundError(f"Директория не найдена: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"Не является директорией: {root}")

    extensions = extensions or DEFAULT_CONFIG.input_extensions
    extensions = tuple(
        ext if ext.startswith(".") else f".{ext}"
        for ext in extensions
    )

    seen = set()  # Для отслеживания дубликатов

    for ext in extensions:
        for file_path in root.rglob(f"*{ext}"):
            if include_hidden or not any(
                part.startswith(".") for part in file_path.parts
            ):
                # Избегаем дубликатов
                if file_path not in seen:
                    seen.add(file_path)
                    yield file_path


def get_file_stats(files: list[Path]) -> dict:
    """
    Получает статистику о файлах для конвертации.

    Аргументы:
        files: Список путей к файлам

    Возвращает:
        Словарь со статистикой (количество, общий_размер, средний_размер и т.д.)
    """
    if not files:
        return {
            "count": 0,
            "total_size": 0,
            "avg_size": 0,
            "min_size": 0,
            "max_size": 0,
            "total_size_mb": 0.0
        }

    sizes = []
    for f in files:
        try:
            sizes.append(f.stat().st_size)
        except OSError:
            sizes.append(0)

    total = sum(sizes)

    return {
        "count": len(files),
        "total_size": total,
        "total_size_mb": total / (1024 * 1024),
        "avg_size": total // len(sizes) if sizes else 0,
        "min_size": min(sizes) if sizes else 0,
        "max_size": max(sizes) if sizes else 0
    }


def print_file_stats(files: list[Path]) -> None:
    """
    Выводит статистику о файлах в консоль.

    Аргументы:
        files: Список путей к файлам
    """
    stats = get_file_stats(files)

    print(f"\n📊 Статистика файлов:")
    print(f"  Количество: {stats['count']}")
    print(f"  Общий размер: {stats['total_size_mb']:.2f} MB")
    print(f"  Средний размер: {stats['avg_size'] / 1024:.1f} KB")
    print(f"  Мин. размер: {stats['min_size'] / 1024:.1f} KB")
    print(f"  Макс. размер: {stats['max_size'] / 1024:.1f} KB")