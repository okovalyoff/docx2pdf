"""
Модуль командного интерфейса.

Парсит и валидирует аргументы командной строки для конвертера DOCX → PDF.
"""

import argparse
import sys
from pathlib import Path
from dataclasses import dataclass

from config import DEFAULT_CONFIG, INPUT_DIR, OUTPUT_DIR
from converters.factory import ConverterBackend, get_available_backends


@dataclass
class CLIArgs:
    """
    Распарсенные аргументы командной строки.

    Атрибуты:
        input_dir: Входная директория с DOCX файлами
        output_dir: Выходная директория для PDF файлов
        workers: Количество параллельных процессов
        resume: Пропускать уже сконвертированные файлы
        overwrite: Перезаписывать существующие файлы
        verbose: Показывать детальную информацию
        dry_run: Показать что будет сделано без конвертации
        timeout: Таймаут на один файл
        no_report: Не создавать CSV отчёт
        backend: Выбранный бэкенд конвертера
    """
    input_dir: Path
    output_dir: Path
    workers: int
    resume: bool
    overwrite: bool
    verbose: bool
    dry_run: bool
    timeout: int
    no_report: bool
    backend: ConverterBackend

    def __post_init__(self):
        """Валидация аргументов после инициализации."""
        # Проверяем существование входной директории
        if not self.input_dir.exists():
            raise ValueError(f"Входная директория не существует: {self.input_dir}")

        if not self.input_dir.is_dir():
            raise ValueError(f"Входной путь не является директорией: {self.input_dir}")

        # Проверяем конфликтующие опции
        if self.resume and self.overwrite:
            raise ValueError("Нельзя использовать одновременно --resume и --overwrite")

        # Предупреждение о количестве воркеров
        if self.backend == ConverterBackend.WORD_COM and self.workers > 3:
            print(
                f"⚠️  Предупреждение: Использование {self.workers} воркеров с Word COM "
                "может вызвать проблемы. Рекомендуется: 1-3 воркера.",
                file=sys.stderr
            )


def parse_args() -> CLIArgs:
    """Парсит аргументы командной строки и возвращает экземпляр CLIArgs."""
    parser = argparse.ArgumentParser(
        prog="docx2pdf",
        description="Конвертер DOCX → PDF (PRO) - Быстрая и надёжная конвертация документов",
        epilog=(
            "Примеры использования:\n"
            "  %(prog)s                                    # Использовать пути по умолчанию\n"
            "  %(prog)s --input ./docs --output ./pdf      # Свои пути\n"
            "  %(prog)s --overwrite                        # Перезаписать существующие\n"
            "  %(prog)s --workers 4 --libreoffice          # Параллельно через LibreOffice\n"
            "  %(prog)s --dry-run                          # Предпросмотр без конвертации"
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter
    )

    # Ввод/вывод
    parser.add_argument(
        "--input", "-i",
        type=Path,
        default=INPUT_DIR,
        help=f"Входная директория с DOCX файлами (по умолчанию: {INPUT_DIR})"
    )
    parser.add_argument(
        "--output", "-o",
        type=Path,
        default=OUTPUT_DIR,
        help=f"Выходная директория для PDF файлов (по умолчанию: {OUTPUT_DIR})"
    )

    # Опции обработки
    parser.add_argument(
        "--workers", "-w",
        type=int,
        default=DEFAULT_CONFIG.max_workers,
        help=(
            f"Количество параллельных процессов (по умолчанию: {DEFAULT_CONFIG.max_workers}). "
            "Для Word COM используйте 1-3. Для LibreOffice можно больше."
        )
    )
    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_CONFIG.conversion_timeout,
        help=f"Таймаут на файл в секундах (по умолчанию: {DEFAULT_CONFIG.conversion_timeout})"
    )

    # Режим продолжения/перезаписи
    mode_group = parser.add_mutually_exclusive_group()
    mode_group.add_argument(
        "--resume",
        action="store_true",
        default=True,
        help="Пропускать уже сконвертированные файлы (поведение по умолчанию)"
    )
    mode_group.add_argument(
        "--overwrite",
        action="store_true",
        help="Перезаписывать существующие PDF файлы"
    )
    mode_group.add_argument(
        "--no-resume",
        action="store_true",
        help="Отключить режим продолжения (аналогично отсутствию --resume)"
    )

    # Детальность вывода
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Показывать детальную информацию о процессе"
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Показать что будет сделано без фактической конвертации"
    )
    parser.add_argument(
        "--no-report",
        action="store_true",
        help="Не создавать файл CSV отчёта"
    )

    # Выбор бэкенда
    backend_group = parser.add_mutually_exclusive_group()
    backend_group.add_argument(
        "--word",
        action="store_true",
        help="Принудительно использовать Windows Word COM бэкенд"
    )
    backend_group.add_argument(
        "--libreoffice",
        action="store_true",
        help="Принудительно использовать LibreOffice бэкенд (кроссплатформенный)"
    )

    args = parser.parse_args()

    # Определяем бэкенд
    if args.word:
        backend = ConverterBackend.WORD_COM
    elif args.libreoffice:
        backend = ConverterBackend.LIBREOFFICE
    else:
        # Используем бэкенд из конфига
        backend_map = {
            "word": ConverterBackend.WORD_COM,
            "libreoffice": ConverterBackend.LIBREOFFICE,
            "auto": ConverterBackend.AUTO
        }
        backend = backend_map.get(
            DEFAULT_CONFIG.default_backend.lower(),
            ConverterBackend.AUTO
        )

    # Обрабатываем логику resume (--no-resume отключает resume по умолчанию)
    resume = args.resume and not args.no_resume

    # Создаём экземпляр CLIArgs
    try:
        return CLIArgs(
            input_dir=args.input,
            output_dir=args.output,
            workers=args.workers,
            resume=resume,
            overwrite=args.overwrite,
            verbose=args.verbose,
            dry_run=args.dry_run,
            timeout=args.timeout,
            no_report=args.no_report,
            backend=backend
        )
    except ValueError as e:
        parser.error(str(e))


def print_available_backends() -> None:
    """Выводит доступные бэкенды конвертера в системе."""
    print("Доступные бэкенды конвертера:")
    available = get_available_backends()

    if not available:
        print("  ❌ Нет доступных бэкендов")
        print("\nУстановите один из:")
        print("  • LibreOffice (кроссплатформенный): https://libreoffice.org")
        print("  • Microsoft Word (только Windows)")
        return

    for backend in available:
        if backend == ConverterBackend.WORD_COM:
            print("  ✅ Word COM (Windows) - Быстрый, нативное форматирование")
        elif backend == ConverterBackend.LIBREOFFICE:
            print("  ✅ LibreOffice (кроссплатформенный) - Надёжный, параллельный")


def validate_environment(args: CLIArgs) -> list[str]:
    """
    Проверяет окружение для конвертации.

    Аргументы:
        args: Распарсенные аргументы CLI

    Возвращает:
        Список предупреждений (пустой, если всё в порядке)
    """
    warnings = []

    # Проверяем доступность бэкенда
    available = get_available_backends()
    if not available:
        warnings.append(
            "Нет доступного бэкенда конвертера. "
            "Установите LibreOffice или Microsoft Word."
        )
    elif args.backend != ConverterBackend.AUTO:
        if args.backend not in available:
            warnings.append(
                f"Запрошенный бэкенд '{args.backend.value}' недоступен. "
                f"Доступные: {[b.value for b in available]}"
            )

    # Проверяем количество воркеров vs бэкенд
    if args.backend == ConverterBackend.WORD_COM and args.workers > 3:
        warnings.append(
            f"Использование {args.workers} воркеров с Word COM не рекомендуется. "
            "Рассмотрите --workers 2 или --libreoffice для параллельной обработки."
        )

    return warnings