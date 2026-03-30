#!/usr/bin/env python3
"""
Конвертер DOCX → PDF (PRO)

Высокопроизводительный конвертер документов Microsoft Word в PDF
с поддержкой параллельной обработки и кроссплатформенности.

Использование:
    python main.py                           # Использовать настройки по умолчанию
    python main.py --input ./docs            # Указать входную директорию
    python main.py --output ./pdf            # Указать выходную директорию
    python main.py --workers 4               # 4 параллельных процесса
    python main.py --overwrite               # Перезаписать существующие
    python main.py --libreoffice             # Использовать LibreOffice
    python main.py --dry-run                 # Предпросмотр без конвертации
"""

import signal
import sys
import time
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from tqdm import tqdm

from config import DEFAULT_CONFIG
from converters.factory import (
    ConverterBackend,
    get_available_backends,
    get_worker_function
)
from utils.cli import parse_args, validate_environment, print_available_backends
from utils.logger import setup_logger, get_logger
from utils.scanner import scan_docx_files, print_file_stats
from output.report import write_report


# Глобальный флаг для плавного завершения
_shutdown_requested = False


def signal_handler(signum, frame):
    """
    Обработчик сигналов прерывания для плавного завершения.

    При первом нажатии Ctrl+C дожидается завершения текущих конвертаций.
    """
    global _shutdown_requested
    _shutdown_requested = True
    print("\n\n⚠️  Запрошено завершение. Дожидаемся завершения текущих конвертаций...")
    print("Нажмите Ctrl+C ещё раз для принудительного выхода.")
    signal.signal(signal.SIGINT, force_quit_handler)


def force_quit_handler(signum, frame):
    """Принудительный выход при повторном прерывании."""
    print("\nПринудительный выход...")
    sys.exit(1)


def run_conversion(args, logger) -> tuple[dict, list]:
    """
    Запускает процесс конвертации.

    Аргументы:
        args: Распарсенные аргументы CLI
        logger: Логгер для записи сообщений

    Возвращает:
        Кортеж (словарь статистики, строки отчёта)
    """
    global _shutdown_requested

    # Получаем рабочую функцию для выбранного бэкенда
    worker_func = get_worker_function(args.backend)

    # Инициализируем счётчики
    stats = {
        "success": 0,
        "failed": 0,
        "skipped": 0,
        "timeout": 0
    }
    report_rows = []

    # Получаем файлы для обработки
    files = scan_docx_files(args.input_dir)

    if not files:
        logger.warning("DOCX файлы не найдены во входной директории")
        return stats, report_rows

    logger.info(f"Найдено {len(files)} DOCX файлов для обработки")

    # Выводим статистику файлов
    if args.verbose:
        print_file_stats(files)

    # Режим dry-run - только показать что будет сделано
    if args.dry_run:
        print(f"\n📁 Режим предпросмотра - {len(files)} файлов будет сконвертировано:")
        for f in files[:10]:  # Показываем первые 10
            print(f"  {f}")
        if len(files) > 10:
            print(f"  ... и ещё {len(files) - 10}")
        return stats, report_rows

    # Создаём выходную директорию
    args.output_dir.mkdir(parents=True, exist_ok=True)

    # Обрабатываем файлы параллельно
    start_time = time.monotonic()

    with ProcessPoolExecutor(max_workers=args.workers) as executor:
        # Отправляем все задачи
        futures = {}
        for file_path in files:
            if _shutdown_requested:
                break

            future = executor.submit(
                worker_func,
                file_path,
                args.input_dir,
                args.output_dir,
                DEFAULT_CONFIG.retry_count,
                DEFAULT_CONFIG.retry_base_delay,
                DEFAULT_CONFIG.retry_max_delay,
                args.timeout,
                args.resume,
                args.overwrite
            )
            futures[future] = file_path

        # Обрабатываем результаты с прогресс-баром
        with tqdm(
            total=len(futures),
            desc="📄 DOCX → PDF",
            unit="файл",
            ncols=80
        ) as pbar:
            for future in as_completed(futures):
                if _shutdown_requested:
                    # Отменяем оставшиеся задачи
                    for f in futures:
                        f.cancel()
                    break

                file_path = futures[future]

                try:
                    status, error_msg, duration = future.result(timeout=args.timeout + 10)
                    stats[status] = stats.get(status, 0) + 1
                    report_rows.append([
                        str(file_path),
                        status,
                        error_msg,
                        f"{duration:.2f}"
                    ])

                    if args.verbose:
                        pbar.write(f"  ✓ {file_path.name} ({duration:.1f}с)")

                except Exception as e:
                    stats["failed"] += 1
                    report_rows.append([
                        str(file_path),
                        "failed",
                        str(e),
                        "0.00"
                    ])
                    logger.error(f"Ошибка обработки {file_path}: {e}")

                pbar.update(1)

    elapsed_time = time.monotonic() - start_time
    stats["_elapsed"] = elapsed_time

    return stats, report_rows


def print_summary(stats: dict, elapsed_time: float):
    """
    Выводит сводку результатов конвертации.

    Аргументы:
        stats: Словарь со статистикой конвертации
        elapsed_time: Общее время выполнения в секундах
    """
    print("\n" + "=" * 50)
    print("📊 Сводка конвертации")
    print("=" * 50)

    total = sum(v for k, v in stats.items() if not k.startswith("_"))

    # Количество по статусам
    if stats.get("success", 0) > 0:
        print(f"  ✅ Успешно:  {stats['success']}")
    if stats.get("skipped", 0) > 0:
        print(f"  ⏭️  Пропущено: {stats['skipped']}")
    if stats.get("failed", 0) > 0:
        print(f"  ❌ Ошибок:   {stats['failed']}")
    if stats.get("timeout", 0) > 0:
        print(f"  ⏱️  Таймаут:  {stats['timeout']}")

    # Время
    print(f"\n  📁 Всего файлов: {total}")
    print(f"  ⏱️  Время: {elapsed_time:.1f}с")

    if stats.get("success", 0) > 0:
        avg_time = elapsed_time / stats["success"]
        print(f"  📈 В среднем: {avg_time:.2f}с/файл")


def main():
    """Главная точка входа."""
    global _shutdown_requested

    # Парсим аргументы
    try:
        args = parse_args()
    except ValueError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        sys.exit(1)

    # Настраиваем логирование
    logger = setup_logger(verbose=args.verbose)

    # Проверяем окружение
    warnings = validate_environment(args)
    for warning in warnings:
        logger.warning(warning)
        print(f"⚠️  Предупреждение: {warning}", file=sys.stderr)

    if not get_available_backends():
        print("\n❌ Нет доступного бэкенда конвертера!")
        print_available_backends()
        sys.exit(1)

    # Показываем информацию о бэкенде
    if args.verbose:
        available = get_available_backends()
        backend = args.backend if args.backend != ConverterBackend.AUTO else available[0]
        print(f"\n🔧 Используется бэкенд: {backend.value}")
        print(f"📁 Вход:  {args.input_dir}")
        print(f"📁 Выход: {args.output_dir}")
        print(f"⚙️  Воркеры: {args.workers}")

        # Информация о метаданных
        if DEFAULT_CONFIG.write_metadata:
            print(f"📝 Метаданные: включены")
            print(f"   Автор: {DEFAULT_CONFIG.metadata.author}")
            print(f"   Организация: {DEFAULT_CONFIG.metadata.organization}")
        print()

    # Устанавливаем обработчики сигналов
    signal.signal(signal.SIGINT, signal_handler)

    try:
        # Запускаем конвертацию
        stats, report_rows = run_conversion(args, logger)

        # Записываем отчёт
        if not args.no_report and report_rows:
            write_report(report_rows)
            print(f"\n📄 Отчёт: {DEFAULT_CONFIG.csv_report}")

        # Выводим сводку
        elapsed_time = stats.pop("_elapsed", 0)
        print_summary(stats, elapsed_time)

        # Выводим расположение файла логов
        print(f"📋 Ошибки: {DEFAULT_CONFIG.log_file}")

        # Выходим с соответствующим кодом
        if _shutdown_requested:
            sys.exit(130)  # 128 + SIGINT
        if stats.get("failed", 0) > 0:
            sys.exit(1)

        print("\n✅ Готово!")

    except Exception as e:
        logger.exception("Критическая ошибка при конвертации")
        print(f"\n❌ Критическая ошибка: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()