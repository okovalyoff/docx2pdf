"""
Базовый модуль конвертера.

Определяет абстрактный интерфейс для всех реализаций конвертеров.
Поддерживает несколько бэкендов: Windows Word COM, LibreOffice и др.
"""

from abc import ABC, abstractmethod
from dataclasses import dataclass
from enum import Enum
from pathlib import Path
from typing import Optional
import logging


class ConversionStatus(Enum):
    """Статус операции конвертации."""
    SUCCESS = "success"    # Успешно сконвертировано
    FAILED = "failed"      # Ошибка конвертации
    SKIPPED = "skipped"    # Пропущено (уже существует)
    TIMEOUT = "timeout"    # Превышено время ожидания


@dataclass
class ConversionResult:
    """
    Результат конвертации одного файла.

    Атрибуты:
        input_path: Путь к исходному файлу
        output_path: Путь к выходному файлу
        status: Статус конвертации
        error_message: Сообщение об ошибке (если есть)
        attempts: Количество попыток
        duration_seconds: Время выполнения в секундах
    """
    input_path: Path
    output_path: Path
    status: ConversionStatus
    error_message: Optional[str] = None
    attempts: int = 1
    duration_seconds: float = 0.0

    def to_csv_row(self) -> list[str]:
        """Преобразует результат в строку CSV."""
        return [
            str(self.input_path),
            self.status.value,
            self.error_message or "",
            f"{self.duration_seconds:.2f}"
        ]


class ConversionError(Exception):
    """Исключение, возникающее при ошибке конвертации."""
    pass


class BaseConverter(ABC):
    """
    Абстрактный базовый класс для конвертеров документов.

    Реализации:
    - WordConverter: Использует Windows Word COM интерфейс
    - LibreOfficeConverter: Использует LibreOffice в headless режиме

    Атрибуты:
        input_root: Корневая директория с исходными файлами
        output_root: Корневая директория для выходных файлов
        retry_count: Количество повторных попыток
        retry_base_delay: Базовая задержка между попытками
        retry_max_delay: Максимальная задержка между попытками
        timeout: Таймаут на одну конвертацию
        resume: Пропускать уже сконвертированные файлы
        overwrite: Перезаписывать существующие файлы
    """

    def __init__(
        self,
        input_root: Path,
        output_root: Path,
        retry_count: int = 3,
        retry_base_delay: float = 2.0,
        retry_max_delay: float = 30.0,
        timeout: int = 120,
        resume: bool = True,
        overwrite: bool = False
    ):
        self.input_root = input_root
        self.output_root = output_root
        self.retry_count = retry_count
        self.retry_base_delay = retry_base_delay
        self.retry_max_delay = retry_max_delay
        self.timeout = timeout
        self.resume = resume
        self.overwrite = overwrite
        self.logger = logging.getLogger(self.__class__.__name__)

    @property
    @abstractmethod
    def name(self) -> str:
        """Возвращает название бэкенда конвертера."""
        pass

    @property
    @abstractmethod
    def is_available(self) -> bool:
        """Проверяет доступность бэкенда в системе."""
        pass

    @abstractmethod
    def _convert_single(self, docx_path: Path, output_pdf: Path) -> None:
        """
        Конвертирует один DOCX файл в PDF.

        Должен быть реализован в подклассах.
        При ошибке должен вызывать ConversionError.

        Аргументы:
            docx_path: Путь к исходному DOCX файлу
            output_pdf: Путь для сохранения PDF файла
        """
        pass

    def get_output_path(self, docx_path: Path) -> Path:
        """
        Вычисляет выходной путь PDF для заданного входного файла.

        Аргументы:
            docx_path: Путь к исходному файлу

        Возвращает:
            Путь к выходному PDF файлу
        """
        relative = docx_path.relative_to(self.input_root)
        output_pdf = self.output_root / relative.with_suffix(".pdf")
        return output_pdf

    def should_skip(self, output_pdf: Path) -> bool:
        """
        Проверяет, нужно ли пропустить файл (уже существует и режим resume).

        Аргументы:
            output_pdf: Путь к выходному файлу

        Возвращает:
            True если файл нужно пропустить
        """
        if not output_pdf.exists():
            return False

        # В режиме перезаписи удаляем существующий файл
        if self.overwrite:
            output_pdf.unlink(missing_ok=True)
            return False

        # В режиме продолжения пропускаем существующие
        if self.resume:
            return True

        return False

    def convert(self, docx_path: Path) -> ConversionResult:
        """
        Конвертирует DOCX файл в PDF с логикой повторных попыток.

        Это основная точка входа для конвертации.
        Обрабатывает повторные попытки, таймауты и логирование ошибок.

        Аргументы:
            docx_path: Путь к исходному DOCX файлу

        Возвращает:
            Объект ConversionResult с результатом конвертации
        """
        import time

        start_time = time.monotonic()
        output_pdf = self.get_output_path(docx_path)
        output_pdf.parent.mkdir(parents=True, exist_ok=True)

        # Проверка на пропуск
        if self.should_skip(output_pdf):
            return ConversionResult(
                input_path=docx_path,
                output_path=output_pdf,
                status=ConversionStatus.SKIPPED
            )

        last_error: Optional[str] = None

        for attempt in range(1, self.retry_count + 1):
            try:
                # Выполняем конвертацию
                self._convert_single(docx_path, output_pdf)

                # Проверяем, что выходной файл создан
                if output_pdf.exists():
                    # Записываем метаданные в PDF
                    self._write_metadata(output_pdf, docx_path)

                    duration = time.monotonic() - start_time
                    return ConversionResult(
                        input_path=docx_path,
                        output_path=output_pdf,
                        status=ConversionStatus.SUCCESS,
                        attempts=attempt,
                        duration_seconds=duration
                    )
                else:
                    raise ConversionError(
                        f"Конвертация завершена, но выходной файл не создан: {output_pdf}"
                    )

            except Exception as e:
                last_error = str(e)
                self.logger.error(
                    f"[Попытка {attempt}/{self.retry_count}] "
                    f"Ошибка конвертации {docx_path}: {e}"
                )

                # Очистка при ошибке
                self._cleanup_on_failure(output_pdf)

                # Ожидание перед повторной попыткой (экспоненциальная задержка)
                if attempt < self.retry_count:
                    delay = min(
                        self.retry_base_delay * (2 ** (attempt - 1)),
                        self.retry_max_delay
                    )
                    time.sleep(delay)

        # Все попытки исчерпаны
        duration = time.monotonic() - start_time
        return ConversionResult(
            input_path=docx_path,
            output_path=output_pdf,
            status=ConversionStatus.FAILED,
            error_message=last_error,
            attempts=self.retry_count,
            duration_seconds=duration
        )

    def _cleanup_on_failure(self, output_pdf: Path) -> None:
        """
        Удаляет частичный/повреждённый выходной файл после ошибки.

        Аргументы:
            output_pdf: Путь к выходному файлу для удаления
        """
        try:
            if output_pdf.exists():
                output_pdf.unlink(missing_ok=True)
        except Exception as e:
            self.logger.warning(f"Не удалось удалить {output_pdf}: {e}")

    def _write_metadata(self, output_pdf: Path, source_file: Path) -> None:
        """
        Записывает метаданные в PDF файл после успешной конвертации.

        Аргументы:
            output_pdf: Путь к созданному PDF файлу
            source_file: Путь к исходному DOCX файлу
        """
        try:
            from output.pdf_metadata import write_metadata_from_config
            write_metadata_from_config(output_pdf, source_file)
        except Exception as e:
            # Ошибка записи метаданных не должна прерывать конвертацию
            self.logger.warning(f"Не удалось записать метаданные в {output_pdf}: {e}")

    def __repr__(self) -> str:
        return f"{self.__class__.__name__}(backend={self.name})"