"""
Конвертер на базе LibreOffice в headless режиме.

Кроссплатформенная конвертация DOCX → PDF с использованием LibreOffice
в headless режиме. Работает на Windows, Linux и macOS.

Потокобезопасность: LibreOffice может запускаться параллельно в нескольких экземплярах.
"""

import platform
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

from converters.base import BaseConverter, ConversionError


class LibreOfficeConverter(BaseConverter):
    """
    Конвертер на базе LibreOffice в headless режиме.

    Использует команду soffice --headless --convert-to pdf.
    Работает на всех платформах, где установлен LibreOffice.

    Примечание:
        LibreOffice в headless режиме более надёжен для параллельной обработки,
        чем Word COM, так как каждый вызов создает независимый процесс.
    """

    # Распространённые имена исполняемых файлов LibreOffice по платформам
    EXECUTABLE_NAMES = {
        "Windows": ["soffice.exe", "libreoffice.exe"],
        "Linux": ["soffice", "libreoffice", "libreoffice7.x"],
        "Darwin": [  # macOS
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "soffice"
        ]
    }

    @property
    def name(self) -> str:
        """Возвращает название бэкенда."""
        return "LibreOffice Headless"

    @property
    def is_available(self) -> bool:
        """Проверяет, что LibreOffice установлен и доступен."""
        return self._find_executable() is not None

    def _find_executable(self) -> Optional[str]:
        """
        Ищет исполняемый файл LibreOffice в системе.

        Возвращает:
            Путь к исполняемому файлу или None, если не найден
        """
        system = platform.system()
        names = self.EXECUTABLE_NAMES.get(system, ["soffice"])

        for name in names:
            # Проверяем, является ли имя абсолютным путём
            if Path(name).is_absolute():
                if Path(name).exists():
                    return name
            else:
                # Ищем в PATH
                found = shutil.which(name)
                if found:
                    return found

        return None

    def _convert_single(self, docx_path: Path, output_pdf: Path) -> None:
        """
        Конвертирует с использованием LibreOffice в headless режиме.

        LibreOffice требует, чтобы выходной файл был в той же директории,
        что и входной, или использует параметр --outdir (доступно в новых версиях).

        Аргументы:
            docx_path: Путь к исходному DOCX файлу
            output_pdf: Путь для сохранения PDF файла

        Исключения:
            ConversionError: При ошибке конвертации
        """
        executable = self._find_executable()
        if not executable:
            raise ConversionError(
                "LibreOffice не найден. Пожалуйста, установите LibreOffice "
                "и убедитесь, что 'soffice' или 'libreoffice' доступен в PATH."
            )

        # LibreOffice лучше работает с абсолютными путями
        input_file = docx_path.resolve()
        output_dir = output_pdf.parent.resolve()

        # Формируем команду
        cmd = [
            executable,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", str(output_dir),
            str(input_file)
        ]

        try:
            # Запускаем конвертацию с таймаутом
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=self.timeout,
                cwd=str(output_dir)  # Устанавливаем рабочую директорию
            )

            if result.returncode != 0:
                error_msg = result.stderr.strip() or result.stdout.strip()
                raise ConversionError(
                    f"LibreOffice вернул код ошибки {result.returncode}: {error_msg}"
                )

            # LibreOffice создаёт PDF с тем же именем, но расширением .pdf
            expected_output = output_dir / input_file.with_suffix(".pdf").name

            # Переименовываем, если нужно (если выходной путь отличается)
            if expected_output != output_pdf and expected_output.exists():
                expected_output.rename(output_pdf)

            # Проверяем выходной файл
            if not output_pdf.exists():
                raise ConversionError(
                    f"LibreOffice завершил работу, но выходной файл не найден: {output_pdf}"
                )

        except subprocess.TimeoutExpired:
            raise ConversionError(
                f"Превышено время ожидания LibreOffice ({self.timeout} секунд)"
            )
        except FileNotFoundError as e:
            raise ConversionError(f"Исполняемый файл LibreOffice не найден: {e}")
        except Exception as e:
            if isinstance(e, ConversionError):
                raise
            raise ConversionError(f"Ошибка конвертации LibreOffice: {e}")

    def convert_batch(
        self,
        docx_paths: list[Path],
        batch_size: int = 10
    ) -> list[tuple[Path, str]]:
        """
        Конвертирует несколько файлов в одном экземпляре LibreOffice.

        Более эффективно, чем отдельные конвертации, так как время
        запуска LibreOffice распределяется на несколько файлов.

        Аргументы:
            docx_paths: Список путей к файлам для конвертации
            batch_size: Количество файлов в одной партии

        Возвращает:
            Список кортежей (путь, статус)
        """
        results = []
        executable = self._find_executable()

        if not executable:
            for path in docx_paths:
                results.append((path, "failed"))
            return results

        # Обрабатываем партиями
        for i in range(0, len(docx_paths), batch_size):
            batch = docx_paths[i:i + batch_size]

            # Создаём временную директорию для пакетной обработки
            with tempfile.TemporaryDirectory() as temp_dir:
                temp_path = Path(temp_dir)

                # Копируем входные файлы во временную директорию
                for docx_path in batch:
                    temp_input = temp_path / docx_path.name
                    shutil.copy2(docx_path, temp_input)

                # Конвертируем все файлы во временной директории
                cmd = [
                    executable,
                    "--headless",
                    "--convert-to", "pdf",
                    str(temp_path)
                ]

                try:
                    subprocess.run(
                        cmd,
                        capture_output=True,
                        timeout=self.timeout * len(batch)
                    )
                except Exception:
                    pass

                # Перемещаем выходные файлы и записываем результаты
                for docx_path in batch:
                    temp_pdf = temp_path / docx_path.with_suffix(".pdf").name
                    output_pdf = self.get_output_path(docx_path)
                    output_pdf.parent.mkdir(parents=True, exist_ok=True)

                    if temp_pdf.exists():
                        shutil.move(str(temp_pdf), str(output_pdf))
                        results.append((docx_path, "success"))
                    else:
                        results.append((docx_path, "failed"))

        return results


def convert_with_libreoffice(
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
    converter = LibreOfficeConverter(
        input_root=input_root,
        output_root=output_root,
        retry_count=retry_count,
        retry_base_delay=retry_base_delay,
        retry_max_delay=retry_max_delay,
        timeout=timeout,
        resume=resume,
        overwrite=overwrite
    )

    result = converter.convert(docx_path)
    return (
        result.status.value,
        result.error_message or "",
        result.duration_seconds
    )