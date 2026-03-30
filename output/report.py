"""
Модуль генерации отчётов.

Создаёт CSV отчёты с результатами конвертации и метриками времени.
"""

import csv
import logging
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from config import DEFAULT_CONFIG


@dataclass
class ReportRow:
    """
    Одна строка в отчёте конвертации.

    Атрибуты:
        file_path: Путь к файлу
        status: Статус конвертации
        error_message: Сообщение об ошибке
        duration_seconds: Время выполнения в секундах
    """
    file_path: str
    status: str
    error_message: str
    duration_seconds: float

    def to_list(self) -> List[str]:
        """Преобразует в список для записи в CSV."""
        return [
            self.file_path,
            self.status,
            self.error_message,
            f"{self.duration_seconds:.2f}"
        ]


class ConversionReport:
    """
    Управляет генерацией отчёта конвертации.

    Поддерживает как немедленную запись, так и накопление пакета.
    Может использоваться как контекстный менеджер.

    Пример использования:
        with ConversionReport("report.csv") as report:
            for file in files:
                result = convert(file)
                report.add_row(...)
    """

    HEADER = ["file", "status", "error", "duration_sec"]

    def __init__(
        self,
        output_path: Optional[Path] = None,
        append: bool = False
    ):
        """
        Инициализация отчёта.

        Аргументы:
            output_path: Путь к выходному файлу
            append: Добавлять к существующему файлу
        """
        self.output_path = output_path or Path(DEFAULT_CONFIG.csv_report)
        self.append = append
        self.rows: List[ReportRow] = []
        self.logger = logging.getLogger(__name__)
        self._start_time: Optional[datetime] = None

    def __enter__(self) -> "ConversionReport":
        """Вход в контекстный менеджер."""
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """Выход из контекстный менеджер."""
        if exc_type is None:
            self.write()
        return False

    def start(self):
        """Отмечает начало пакета конвертации."""
        self._start_time = datetime.now()

    def add_row(
        self,
        file_path: str,
        status: str,
        error_message: str = "",
        duration_seconds: float = 0.0
    ):
        """
        Добавляет строку в отчёт.

        Аргументы:
            file_path: Путь к файлу
            status: Статус конвертации
            error_message: Сообщение об ошибке (опционально)
            duration_seconds: Время выполнения в секундах
        """
        row = ReportRow(
            file_path=file_path,
            status=status,
            error_message=error_message,
            duration_seconds=duration_seconds
        )
        self.rows.append(row)

    def add_result(self, result):
        """
        Добавляет результат конвертации в отчёт.

        Аргументы:
            result: Объект ConversionResult
        """
        self.add_row(
            file_path=str(result.input_path),
            status=result.status.value,
            error_message=result.error_message or "",
            duration_seconds=result.duration_seconds
        )

    def write(self) -> Path:
        """
        Записывает отчёт в CSV файл.

        Возвращает:
            Путь к записанному файлу
        """
        mode = "a" if self.append else "w"

        with open(
            self.output_path,
            mode,
            newline="",
            encoding="utf-8"
        ) as f:
            writer = csv.writer(f)

            # Записываем заголовок только в режиме записи
            if mode == "w":
                writer.writerow(self.HEADER)

            # Записываем строки
            for row in self.rows:
                writer.writerow(row.to_list())

        self.logger.info(f"Отчёт записан в {self.output_path}")
        return self.output_path

    def get_summary(self) -> dict:
        """
        Получает сводную статистику из собранных строк.

        Возвращает:
            Словарь со сводной статистикой
        """
        stats = {
            "total": len(self.rows),
            "success": 0,
            "failed": 0,
            "skipped": 0,
            "timeout": 0,
            "total_duration": 0.0,
            "avg_duration": 0.0,
            "errors": []
        }

        successful_durations = []

        for row in self.rows:
            stats[row.status] = stats.get(row.status, 0) + 1

            if row.status == "success":
                successful_durations.append(row.duration_seconds)
                stats["total_duration"] += row.duration_seconds

            if row.error_message:
                stats["errors"].append({
                    "file": row.file_path,
                    "error": row.error_message
                })

        if successful_durations:
            stats["avg_duration"] = sum(successful_durations) / len(successful_durations)

        return stats

    def print_summary(self):
        """Выводит сводку в консоль."""
        stats = self.get_summary()

        print("\n📊 Сводка отчёта:")
        print(f"  Всего файлов: {stats['total']}")
        print(f"  ✅ Успешно: {stats['success']}")
        print(f"  ❌ Ошибок: {stats['failed']}")
        print(f"  ⏭️  Пропущено: {stats['skipped']}")

        if stats['avg_duration'] > 0:
            print(f"\n⏱️  Время:")
            print(f"  Всего: {stats['total_duration']:.1f}с")
            print(f"  В среднем: {stats['avg_duration']:.2f}с/файл")


def write_report(
    rows: List[List[str]],
    output_path: Optional[str] = None
) -> str:
    """
    Записывает отчёт конвертации в CSV файл.

    Функция для обратной совместимости.

    Аргументы:
        rows: Список строк (каждая строка: [файл, статус, ошибка, время])
        output_path: Путь к выходному файлу (по умолчанию из конфига)

    Возвращает:
        Путь к записанному файлу
    """
    output_path = output_path or DEFAULT_CONFIG.csv_report

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["file", "status", "error", "duration_sec"])
        writer.writerows(rows)

    return output_path


def append_report(rows: List[List[str]], output_path: Optional[str] = None) -> str:
    """
    Добавляет к существующему отчёту конвертации.

    Аргументы:
        rows: Список строк для добавления
        output_path: Путь к выходному файлу

    Возвращает:
        Путь к записанному файлу
    """
    output_path = output_path or DEFAULT_CONFIG.csv_report

    file_exists = Path(output_path).exists()

    with open(output_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)

        # Записываем заголовок если файл новый
        if not file_exists:
            writer.writerow(["file", "status", "error", "duration_sec"])

        writer.writerows(rows)

    return output_path