"""
Модуль для работы с метаданными PDF.

Предоставляет функции для чтения и записи метаданных в PDF файлы.
Использует библиотеку pypdf для манипуляции метаданными.
"""

import logging
import os
import time
import gc
from pathlib import Path
from typing import Optional
from datetime import datetime

logger = logging.getLogger(__name__)


def is_pypdf_available() -> bool:
    """
    Проверяет доступность библиотеки pypdf.

    Возвращает:
        True если pypdf или PyPDF2 установлены
    """
    try:
        import pypdf
        return True
    except ImportError:
        try:
            import PyPDF2
            return True
        except ImportError:
            return False


def _get_pdf_library():
    """
    Получает доступную библиотеку для работы с PDF.

    Возвращает:
        Модуль pypdf или PyPDF2

    Исключения:
        ImportError: Если ни одна библиотека не установлена
    """
    try:
        import pypdf
        return pypdf
    except ImportError:
        try:
            import PyPDF2
            return PyPDF2
        except ImportError:
            raise ImportError(
                "Для записи метаданных необходима библиотека pypdf или PyPDF2. "
                "Установите: pip install pypdf"
            )


def _replace_file_with_retry(
    source: Path,
    target: Path,
    max_retries: int = 5,
    delay: float = 0.5
) -> bool:
    """
    Заменяет файл с повторными попытками при ошибке доступа.

    Windows может блокировать файл после записи, поэтому
    пробуем несколько раз с задержкой.

    Аргументы:
        source: Путь к исходному файлу (временный)
        target: Путь к целевому файлу
        max_retries: Максимальное количество попыток
        delay: Задержка между попытками в секундах

    Возвращает:
        True если замена прошла успешно
    """
    for attempt in range(max_retries):
        try:
            # Принудительный сбор мусора для освобождения дескрипторов
            gc.collect()

            # Пробуем заменить файл
            os.replace(str(source), str(target))
            return True

        except PermissionError:
            if attempt < max_retries - 1:
                logger.debug(
                    f"Файл заблокирован, попытка {attempt + 1}/{max_retries}: {target}"
                )
                time.sleep(delay)
            else:
                raise
        except OSError as e:
            if attempt < max_retries - 1:
                logger.debug(f"Ошибка OS при замене файла: {e}, повторяем...")
                time.sleep(delay)
            else:
                raise

    return False


def write_metadata(
    pdf_path: Path,
    author: Optional[str] = None,
    title: Optional[str] = None,
    subject: Optional[str] = None,
    keywords: Optional[str] = None,
    creator: Optional[str] = None,
    producer: Optional[str] = None,
    organization: Optional[str] = None,
    source_file: Optional[Path] = None
) -> bool:
    """
    Записывает метаданные в PDF файл.

    Аргументы:
        pdf_path: Путь к PDF файлу
        author: Автор документа
        title: Название документа
        subject: Тема документа
        keywords: Ключевые слова
        creator: Программа-создатель
        producer: Программа-производитель
        organization: Организация
        source_file: Путь к исходному DOCX файлу

    Возвращает:
        True если метаданные записаны успешно, False при ошибке
    """
    if not is_pypdf_available():
        logger.warning(
            "Библиотека pypdf/PyPDF2 не установлена. "
            "Метаданные не будут записаны. Установите: pip install pypdf"
        )
        return False

    pdf_lib = _get_pdf_library()
    temp_path = pdf_path.with_suffix(".pdf.tmp")

    try:
        # Читаем существующий PDF и сразу закрываем файл
        reader = None
        writer = None

        try:
            with open(pdf_path, "rb") as f:
                reader = pdf_lib.PdfReader(f)
                writer = pdf_lib.PdfWriter()

                # Копируем все страницы
                for page in reader.pages:
                    writer.add_page(page)

                # Получаем существующие метаданные
                existing_metadata = dict(reader.metadata or {})

            # Файл закрыт, теперь работаем с метаданными
            # Создаём новые метаданные на основе существующих
            new_metadata = existing_metadata.copy()

            # Обновляем/добавляем новые значения
            if author:
                new_metadata["/Author"] = author

            if title:
                new_metadata["/Title"] = title
            elif source_file:
                # Используем имя файла как заголовок
                new_metadata["/Title"] = source_file.stem

            if subject:
                new_metadata["/Subject"] = subject

            if keywords:
                new_metadata["/Keywords"] = keywords

            if creator:
                new_metadata["/Creator"] = creator

            if producer:
                new_metadata["/Producer"] = producer

            # Добавляем дату модификации
            new_metadata["/ModDate"] = datetime.now().strftime("D:%Y%m%d%H%M%S+00'00'")

            # Добавляем информацию об организации (кастомное поле)
            if organization:
                new_metadata["/Company"] = organization

            # Записываем метаданные
            writer.add_metadata(new_metadata)

            # Сохраняем во временный файл
            with open(temp_path, "wb") as out:
                writer.write(out)

        finally:
            # Явно освобождаем ресурсы
            del reader
            del writer
            gc.collect()

        # Небольшая пауза перед заменой файла
        time.sleep(0.3)

        # Заменяем оригинальный файл с повторными попытками
        _replace_file_with_retry(temp_path, pdf_path)

        logger.debug(f"Метаданные записаны в {pdf_path}")
        return True

    except PermissionError as e:
        logger.warning(
            f"Не удалось записать метаданные в {pdf_path}: файл заблокирован. {e}"
        )
        # Удаляем временный файл
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass
        return False

    except Exception as e:
        logger.error(f"Ошибка записи метаданных в {pdf_path}: {e}")
        # Удаляем временный файл
        try:
            if temp_path.exists():
                temp_path.unlink()
        except Exception:
            pass
        return False


def write_metadata_from_config(
    pdf_path: Path,
    source_file: Optional[Path] = None
) -> bool:
    """
    Записывает метаданные в PDF файл, используя конфигурацию.

    Аргументы:
        pdf_path: Путь к PDF файлу
        source_file: Путь к исходному DOCX файлу

    Возвращает:
        True если метаданные записаны успешно
    """
    from config import DEFAULT_CONFIG

    if not DEFAULT_CONFIG.write_metadata:
        return False

    metadata_config = DEFAULT_CONFIG.metadata

    return write_metadata(
        pdf_path=pdf_path,
        author=metadata_config.author,
        title=source_file.stem if source_file else None,
        subject=metadata_config.subject or None,
        keywords=metadata_config.keywords or None,
        creator=metadata_config.creator,
        producer=metadata_config.creator,
        organization=metadata_config.organization or None,
        source_file=source_file
    )


def read_metadata(pdf_path: Path) -> dict:
    """
    Читает метаданные из PDF файла.

    Аргументы:
        pdf_path: Путь к PDF файлу

    Возвращает:
        Словарь с метаданными
    """
    if not is_pypdf_available():
        logger.warning("Библиотека pypdf/PyPDF2 не установлена")
        return {}

    try:
        pdf_lib = _get_pdf_library()

        with open(pdf_path, "rb") as f:
            reader = pdf_lib.PdfReader(f)
            metadata = dict(reader.metadata or {})

            # Преобразуем в обычный словарь с понятными ключами
            result = {}
            key_mapping = {
                "/Author": "author",
                "/Title": "title",
                "/Subject": "subject",
                "/Keywords": "keywords",
                "/Creator": "creator",
                "/Producer": "producer",
                "/ModDate": "modified",
                "/CreationDate": "created",
                "/Company": "organization"
            }

            for key, value in metadata.items():
                readable_key = key_mapping.get(key, key.lstrip("/"))
                result[readable_key] = str(value) if value else ""

            return result

    except Exception as e:
        logger.error(f"Ошибка чтения метаданных из {pdf_path}: {e}")
        return {}


def print_metadata(pdf_path: Path) -> None:
    """
    Выводит метаданные PDF файла в консоль.

    Аргументы:
        pdf_path: Путь к PDF файлу
    """
    metadata = read_metadata(pdf_path)

    if not metadata:
        print(f"Метаданные не найдены или не удалось прочитать: {pdf_path}")
        return

    print(f"\n📄 Метаданные: {pdf_path.name}")
    print("-" * 50)

    labels = {
        "author": "Автор",
        "title": "Название",
        "subject": "Тема",
        "keywords": "Ключевые слова",
        "creator": "Создатель",
        "producer": "Производитель",
        "modified": "Изменён",
        "created": "Создан",
        "organization": "Организация"
    }

    for key, value in metadata.items():
        label = labels.get(key, key)
        if value:
            print(f"  {label}: {value}")
