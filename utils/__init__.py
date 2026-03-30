"""
Пакет утилит для конвертера.

Содержит вспомогательные модули для CLI, логирования,
сканирования файлов и работы с процессами.

Примеры использования:
    from utils import parse_args, setup_logger, scan_docx_files
"""

from utils.cli import parse_args, CLIArgs, validate_environment, print_available_backends
from utils.logger import setup_logger, get_logger, ProgressLogger
from utils.scanner import scan_docx_files, scan_docx_files_iter, get_file_stats

__all__ = [
    # CLI
    "parse_args",
    "CLIArgs",
    "validate_environment",
    "print_available_backends",
    # Logger
    "setup_logger",
    "get_logger",
    "ProgressLogger",
    # Scanner
    "scan_docx_files",
    "scan_docx_files_iter",
    "get_file_stats",
]