"""
Утилиты для работы с процессами Word.

Предоставляет функции для управления процессами Word на Windows.
Используется для очистки после ошибок конвертации.
"""

import logging
import platform
import subprocess
from typing import List

logger = logging.getLogger(__name__)


class WordProcessManager:
    """
    Менеджер процессов Word.

    Предоставляет методы для обнаружения и завершения процессов Word.
    Полезен для восстановления после зависаний и ошибок конвертации.
    """

    @staticmethod
    def is_windows() -> bool:
        """Проверяет, что запуск происходит на Windows."""
        return platform.system() == "Windows"

    @staticmethod
    def get_word_processes() -> List[dict]:
        """
        Получает список запущенных процессов Word.

        Возвращает:
            Список словарей с информацией о процессах (pid, name, memory_mb)
        """
        if not WordProcessManager.is_windows():
            return []

        processes = []

        try:
            import psutil

            for proc in psutil.process_iter(["pid", "name", "memory_info"]):
                try:
                    name = proc.info.get("name", "")
                    if name and "WINWORD" in name.upper():
                        processes.append({
                            "pid": proc.info["pid"],
                            "name": name,
                            "memory_mb": proc.info["memory_info"].rss // (1024 * 1024)
                        })
                except (psutil.NoSuchProcess, psutil.AccessDenied, KeyError):
                    pass

        except ImportError:
            logger.debug("psutil не установлен, невозможно получить список процессов")

        return processes

    @staticmethod
    def kill_word_processes(
        timeout: int = 30,
        force: bool = False
    ) -> int:
        """
        Завершает все процессы Word.

        Сначала пытается завершить процессы корректно через terminate(),
        затем, если force=True, принудительно через kill().

        Аргументы:
            timeout: Таймаут в секундах для плавного завершения
            force: Принудительно завершить, если плавное не удалось

        Возвращает:
            Количество завершённых процессов
        """
        if not WordProcessManager.is_windows():
            logger.debug("Не Windows, нет процессов Word для завершения")
            return 0

        killed = 0

        # Сначала пробуем через psutil (более надёжно)
        try:
            import psutil
            import time

            for proc in psutil.process_iter(["pid", "name"]):
                try:
                    if proc.info["name"] and "WINWORD" in proc.info["name"].upper():
                        proc.terminate()
                        killed += 1
                        logger.debug(f"Завершён процесс Word: PID {proc.pid}")
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    pass
                except Exception as e:
                    logger.warning(f"Ошибка завершения процесса: {e}")

            # Ждём плавного завершения
            time.sleep(2)

            # Проверяем оставшиеся процессы
            remaining = WordProcessManager.get_word_processes()
            if remaining and force:
                for proc in remaining:
                    try:
                        p = psutil.Process(proc["pid"])
                        p.kill()
                        logger.debug(f"Принудительно завершён процесс Word: PID {proc['pid']}")
                    except Exception as e:
                        logger.warning(f"Ошибка принудительного завершения: {e}")

        except ImportError:
            # Резервный вариант: команда taskkill
            try:
                result = subprocess.run(
                    ["taskkill", "/F", "/IM", "WINWORD.EXE"],
                    capture_output=True,
                    text=True,
                    timeout=timeout
                )

                if result.returncode == 0:
                    logger.debug("Процессы Word завершены через taskkill")
                    killed = 1  # Невозможно определить точное количество
                elif "not found" not in result.stderr.lower():
                    logger.warning(f"taskkill не удалось: {result.stderr}")

            except subprocess.TimeoutExpired:
                logger.error("taskkill превысил время ожидания")
            except FileNotFoundError:
                logger.error("Команда taskkill не найдена")
            except Exception as e:
                logger.error(f"Ошибка выполнения taskkill: {e}")

        return killed

    @staticmethod
    def is_word_running() -> bool:
        """Проверяет, запущен ли Word в данный момент."""
        return len(WordProcessManager.get_word_processes()) > 0

    @staticmethod
    def is_word_installed() -> bool:
        """
        Проверяет, установлен ли Microsoft Word.

        Возвращает:
            True если Word установлен
        """
        if not WordProcessManager.is_windows():
            return False

        try:
            import winreg
            with winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                r"Word.Application\CLSID"
            ):
                return True
        except (OSError, ImportError):
            pass

        return False


# Функции для обратной совместимости
def kill_word() -> int:
    """
    Завершает все процессы Word.

    Функция для обратной совместимости.

    Возвращает:
        Количество завершённых процессов
    """
    return WordProcessManager.kill_word_processes(force=True)


def get_word_pids() -> List[int]:
    """
    Получает список PID процессов Word.

    Возвращает:
        Список идентификаторов процессов
    """
    return [p["pid"] for p in WordProcessManager.get_word_processes()]


def is_word_installed() -> bool:
    """
    Проверяет, установлен ли Microsoft Word.

    Возвращает:
        True если Word установлен
    """
    return WordProcessManager.is_word_installed()