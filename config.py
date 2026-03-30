"""
Модуль конфигурации конвертера DOCX → PDF.

Содержит все настраиваемые параметры, пути и настройки.
Использует dataclass для типобезопасности и лучшей поддержки IDE.
"""

from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
import getpass
from datetime import datetime


@dataclass
class MetadataConfig:
    """
    Конфигурация метаданных PDF.

    Атрибуты:
        author: Автор документа
        organization: Название организации
        creator: Программа-создатель (название конвертера)
        subject: Тема документа
        keywords: Ключевые слова (через запятую)
    """
    author: str = field(default_factory=lambda: getpass.getuser())
    organization: str = ""
    creator: str = "DOCX2PDF Converter PRO"
    subject: str = ""
    keywords: str = ""

    def to_dict(self, source_file: Optional[Path] = None) -> dict:
        """
        Преобразует конфигурацию метаданных в словарь для pypdf.

        Аргументы:
            source_file: Путь к исходному файлу (для добавления в метаданные)

        Возвращает:
            Словарь с метаданными
        """
        metadata = {
            "/Author": self.author,
            "/Creator": self.creator,
            "/Producer": self.creator,
        }

        if self.organization:
            metadata["/Company"] = self.organization

        if self.subject:
            metadata["/Subject"] = self.subject

        if self.keywords:
            metadata["/Keywords"] = self.keywords

        # Добавляем дату модификации
        metadata["/ModDate"] = datetime.now().strftime("D:%Y%m%d%H%M%S+00'00'")

        if source_file:
            metadata["/Title"] = source_file.stem

        return metadata


# ============================================================================
# КОНФИГУРАЦИЯ МЕТАДАННЫХ (настройте под себя)
# ============================================================================

METADATA_CONFIG = MetadataConfig(
    author="Безкровная О.А.",
    organization="ГБОУ ЛНР «Успенская СШ № 1»",
    creator="DOCX2PDF Converter PRO",
    subject="Меню-требование",
    keywords="меню, столовая, питание"
)


# ============================================================================
# ОСНОВНЫЕ НАСТРОЙКИ КОНВЕРТЕРА
# ============================================================================

# Входная директория с DOCX файлами
INPUT_DIR = Path(r"C:\Users\Владелец\Documents\Столовая\docx")

# Выходная директория для PDF файлов
OUTPUT_DIR = Path(r"C:\Users\Владелец\Documents\Столовая\pdf")

# Количество параллельных процессов
# Примечание: Для Word COM используйте 1-2 воркера во избежание исчерпания ресурсов.
# Для LibreOffice можно больше.
MAX_WORKERS = 1

# Настройки повторных попыток с экспоненциальной задержкой
RETRY_COUNT = 3
RETRY_BASE_DELAY = 2.0  # секунды
RETRY_MAX_DELAY = 30.0  # максимальная задержка между попытками

# Таймаут для конвертации одного файла (предотвращает зависания)
CONVERSION_TIMEOUT = 120  # секунды

# Выходные файлы
LOG_FILE = "errors.log"
CSV_REPORT = "report.csv"

# Расширения входных файлов
INPUT_EXTENSIONS = (".docx", ".doc")

# Бэкенд по умолчанию: "word", "libreoffice", "auto"
DEFAULT_BACKEND = "word"

# Записывать метаданные в PDF
WRITE_METADATA = True


# ============================================================================
# КЛАСС КОНФИГУРАЦИИ (для программного использования)
# ============================================================================

class Config:
    """
    Класс-обёртка для доступа к настройкам.
    
    Предоставляет удобный доступ к конфигурации через атрибуты.
    """

    @property
    def input_dir(self) -> Path:
        """Входная директория с DOCX файлами."""
        return INPUT_DIR

    @property
    def output_dir(self) -> Path:
        """Выходная директория для PDF файлов."""
        return OUTPUT_DIR

    @property
    def max_workers(self) -> int:
        """Количество параллельных процессов."""
        return MAX_WORKERS

    @property
    def retry_count(self) -> int:
        """Количество повторных попыток при ошибке."""
        return RETRY_COUNT

    @property
    def retry_base_delay(self) -> float:
        """Базовая задержка между попытками (секунды)."""
        return RETRY_BASE_DELAY

    @property
    def retry_max_delay(self) -> float:
        """Максимальная задержка между попытками (секунды)."""
        return RETRY_MAX_DELAY

    @property
    def conversion_timeout(self) -> int:
        """Таймаут на конвертацию одного файла (секунды)."""
        return CONVERSION_TIMEOUT

    @property
    def log_file(self) -> str:
        """Имя файла для логирования ошибок."""
        return LOG_FILE

    @property
    def csv_report(self) -> str:
        """Имя файла CSV-отчёта."""
        return CSV_REPORT

    @property
    def input_extensions(self) -> tuple:
        """Допустимые расширения входных файлов."""
        return INPUT_EXTENSIONS

    @property
    def default_backend(self) -> str:
        """Бэкенд конвертера по умолчанию."""
        return DEFAULT_BACKEND

    @property
    def metadata(self) -> MetadataConfig:
        """Конфигурация метаданных PDF."""
        return METADATA_CONFIG

    @property
    def write_metadata(self) -> bool:
        """Записывать ли метаданные в PDF."""
        return WRITE_METADATA

    @property
    def output_dir_absolute(self) -> Path:
        """Абсолютный путь к выходной директории."""
        return OUTPUT_DIR.resolve()

    @property
    def input_dir_absolute(self) -> Path:
        """Абсолютный путь к входной директории."""
        return INPUT_DIR.resolve()

    def validate_paths(self) -> list[str]:
        """
        Проверка существования входной директории.

        Возвращает:
            Список предупреждений (пустой, если всё в порядке)
        """
        warnings = []
        if not INPUT_DIR.exists():
            warnings.append(f"Входная директория не существует: {INPUT_DIR}")
        return warnings


# Глобальный экземпляр конфигурации
DEFAULT_CONFIG = Config()

# Константы для обратной совместимости
DEFAULT_INPUT = INPUT_DIR
DEFAULT_OUTPUT = OUTPUT_DIR
DEFAULT_WORKERS = MAX_WORKERS