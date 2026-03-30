"""
Конвертер на базе Windows Word COM.

Использует приложение Windows Word через COM интерфейс для конвертации DOCX → PDF.
Работает только на Windows с установленным Microsoft Word.

Потокобезопасность: Каждый процесс получает свой экземпляр Word.
Обязательно использовать с ProcessPoolExecutor, не ThreadPoolExecutor.
"""

import platform
import subprocess
from pathlib import Path
from typing import Optional

from converters.base import BaseConverter, ConversionError


class WordConverter(BaseConverter):
    """
    Конвертер на базе Windows Word COM.

    Использует Word COM интерфейс для нативной конвертации документов.
    Каждый процесс должен инициализировать свой COM-экземпляр для потокобезопасности.

    Примечание:
        Word COM интерфейс не является потокобезопасным, поэтому каждый процесс
        должен работать с отдельным экземпляром Word. Именно поэтому используется
        ProcessPoolExecutor вместо ThreadPoolExecutor.
    """

    # Экземпляр Word для текущего процесса (не разделяется между процессами)
    _word_app: Optional[object] = None

    @property
    def name(self) -> str:
        """Возвращает название бэкенда."""
        return "Windows Word COM"

    @property
    def is_available(self) -> bool:
        """Проверяет, что система Windows и Word установлен."""
        if platform.system() != "Windows":
            return False

        try:
            # Проверяем наличие Word через реестр
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                r"Word.Application\CLSID"
            ):
                return True
        except (ImportError, OSError):
            return False

    def _ensure_word_instance(self) -> object:
        """
        Создаёт экземпляр Word COM для текущего процесса.

        Создаёт новый экземпляр, если он не существует.
        Использует модель STA (Single-Threaded Apartment) для потокобезопасности.

        Возвращает:
            Объект Word Application

        Исключения:
            ConversionError: Если не удалось создать экземпляр Word
        """
        if self._word_app is not None:
            return self._word_app

        try:
            # Инициализация COM в режиме STA (требуется для Word)
            import pythoncom
            pythoncom.CoInitialize()

            # Создание экземпляра Word Application
            import win32com.client
            self._word_app = win32com.client.Dispatch("Word.Application")
            self._word_app.Visible = False
            self._word_app.DisplayAlerts = False  # Отключаем диалоговые окна

            self.logger.debug("Создан новый экземпляр Word COM")
            return self._word_app

        except ImportError as e:
            raise ConversionError(
                f"Не установлены необходимые пакеты. "
                f"Выполните: pip install pywin32. Ошибка: {e}"
            )
        except Exception as e:
            raise ConversionError(f"Не удалось создать экземпляр Word: {e}")

    def _convert_single(self, docx_path: Path, output_pdf: Path) -> None:
        """
        Конвертирует один файл с использованием Word COM интерфейса.

        Этот метод вызывается базовым классом convert().

        Аргументы:
            docx_path: Путь к исходному DOCX файлу
            output_pdf: Путь для сохранения PDF файла

        Исключения:
            ConversionError: При ошибке конвертации
        """
        word = self._ensure_word_instance()

        try:
            # Открываем документ
            doc = word.Documents.Open(
                str(docx_path.resolve()),
                ReadOnly=True,
                Visible=False
            )

            try:
                # Экспортируем в PDF (WD_FORMAT_PDF = 17)
                doc.ExportAsFixedFormat(
                    str(output_pdf.resolve()),
                    17,  # wdExportFormatPDF
                    False,  # OpenAfterExport
                    0,  # wdExportAllDocument
                )
            finally:
                # Всегда закрываем документ
                doc.Close(False)  # False = не сохранять изменения

        except Exception as e:
            # При ошибке завершаем Word для предотвращения зомби-процессов
            self._kill_word_process()
            self._word_app = None
            raise ConversionError(f"Ошибка конвертации Word: {e}")

    def _kill_word_process(self) -> None:
        """Завершает процесс Word для восстановления после ошибок."""
        try:
            import psutil
            for proc in psutil.process_iter(["name"]):
                try:
                    if proc.info["name"] and "WINWORD" in proc.info["name"].upper():
                        proc.kill()
                        self.logger.debug(f"Завершён процесс Word: PID {proc.pid}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
        except ImportError:
            # Резервный вариант: используем команду taskkill
            subprocess.run(
                ["taskkill", "/F", "/IM", "WINWORD.EXE"],
                capture_output=True,
                timeout=30
            )
        except Exception as e:
            self.logger.warning(f"Не удалось завершить процесс Word: {e}")

    def cleanup(self) -> None:
        """Очищает экземпляр Word COM."""
        if self._word_app is not None:
            try:
                self._word_app.Quit()
            except Exception:
                self._kill_word_process()
            finally:
                self._word_app = None

        # Деинициализация COM
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:
            pass


def convert_with_word(
    docx_path: Path,
    input_root: Path,
    output_root: Path,
    retry_count: int,
    retry_base_delay: float,
    retry_max_delay: float,
    timeout: int,
    resume: bool,
    overwrite: bool
) -> tuple[str, str, float]:
    """
    Рабочая функция для ProcessPoolExecutor.

    Создаёт новый экземпляр WordConverter для каждого процесса
    (требуется для безопасности COM).
    Возвращает кортеж для удобной сериализации между процессами.

    Аргументы:
        docx_path: Путь к исходному файлу
        input_root: Корневая директория с исходными файлами
        output_root: Корневая директория для выходных файлов
        retry_count: Количество повторных попыток
        retry_base_delay: Базовая задержка между попытками
        retry_max_delay: Максимальная задержка
        timeout: Таймаут на конвертацию
        resume: Пропускать существующие файлы
        overwrite: Перезаписывать существующие файлы

    Возвращает:
        Кортеж (статус, сообщение_об_ошибке, время_выполнения)
    """
    converter = WordConverter(
        input_root=input_root,
        output_root=output_root,
        retry_count=retry_count,
        retry_base_delay=retry_base_delay,
        retry_max_delay=retry_max_delay,
        timeout=timeout,
        resume=resume,
        overwrite=overwrite
    )

    try:
        result = converter.convert(docx_path)
        return (
            result.status.value,
            result.error_message or "",
            result.duration_seconds
        )
    finally:
        converter.cleanup()