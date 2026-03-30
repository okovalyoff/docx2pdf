"""
Пакет для работы с выходными данными.

Содержит модули для генерации отчётов и работы с метаданными PDF.

Примеры использования:
    from output import write_report, read_metadata, print_metadata
"""

from output.report import write_report, append_report, ConversionReport
from output.pdf_metadata import (
    write_metadata,
    write_metadata_from_config,
    read_metadata,
    print_metadata,
    is_pypdf_available
)

__all__ = [
    # Report
    "write_report",
    "append_report",
    "ConversionReport",
    # PDF Metadata
    "write_metadata",
    "write_metadata_from_config",
    "read_metadata",
    "print_metadata",
    "is_pypdf_available",
]