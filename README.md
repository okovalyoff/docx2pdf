# 📄 DOCX → PDF Converter PRO

[![Python](https://badgen.net/badge/Python/3.10%2B/blue)](https://www.python.org/)
[![Platform](https://badgen.net/badge/Platform/Windows%20%7C%20Linux%20%7C%20macOS/grey)](https://github.com)
[![License](https://badgen.net/badge/License/MIT/green)](LICENSE)

> Высокопроизводительный конвертер документов Microsoft Word в PDF с поддержкой параллельной обработки, кроссплатформенности и автоматической записи метаданных.

![Status](https://badgen.net/badge/Status/Production%20Ready/green)

---

## ✨ Возможности

### 🚀 Производительность

- **Параллельная обработка** — используйте несколько процессов для одновременной конвертации файлов
- **ProcessPoolExecutor** — безопасная работа с Word COM через отдельные процессы
- **Экспоненциальный backoff** — интеллектуальные повторные попытки при ошибках
- **Graceful shutdown** — корректное завершение по Ctrl+C без потери данных
- **Прогресс-бар** — наглядное отображение процесса конвертации

### 🌍 Кроссплатформенность

| Платформа | Бэкенд | Статус |
|-----------|--------|--------|
| 🪟 Windows | Microsoft Word COM | ✅ Полная поддержка |
| 🪟 Windows | LibreOffice | ✅ Полная поддержка |
| 🐧 Linux | LibreOffice | ✅ Полная поддержка |
| 🍎 macOS | LibreOffice | ✅ Полная поддержка |

### 📝 Метаданные PDF

Автоматическая запись метаданных в созданные PDF-файлы:

| Поле | Описание |
|------|----------|
| Author | Автор документа |
| Title | Название документа |
| Subject | Тема документа |
| Keywords | Ключевые слова |
| Creator | Программа-создатель |
| Organization | Организация |

### 🛠️ Гибкая конфигурация

- Настройка через `config.py` или аргументы CLI
- Режим **dry-run** для предпросмотра без конвертации
- Режим **resume** для продолжения прерванной конвертации
- Режим **overwrite** для перезаписи существующих файлов
- Настраиваемый таймаут для предотвращения зависаний

---

## 📦 Установка

### Требования

- Python 3.10 или выше
- Microsoft Word (Windows) или LibreOffice (все платформы)

### Шаг 1: Клонирование репозитория

```bash
git clone https://github.com/okovalyoff/docx2pdf.git
cd docx2pdf-converter
```

### Шаг 2: Создание виртуального окружения

```bash
# Windows
python -m venv venv
venv\Scripts\activate

# Linux / macOS
python3 -m venv venv
source venv/bin/activate
```

### Шаг 3: Установка зависимостей

```bash
pip install -r requirements.txt
```

### Зависимости

```
tqdm>=4.65.0          # Прогресс-бары
pywin32>=305          # Windows Word COM (только Windows)
psutil>=5.9.0         # Управление процессами
pypdf>=3.0.0          # Работа с метаданными PDF
```

---

## 🚀 Использование

### Базовый запуск

```bash
python main.py
```

### С указанием директорий

```bash
python main.py --input ./documents --output ./pdf
```

### Ключи командной строки

| Ключ | Сокращение | Описание |
|------|------------|----------|
| `--input` | `-i` | Входная директория с DOCX файлами |
| `--output` | `-o` | Выходная директория для PDF файлов |
| `--workers` | `-w` | Количество параллельных процессов |
| `--timeout` | | Таймаут на файл (секунды) |
| `--verbose` | `-v` | Детальный вывод |
| `--dry-run` | | Предпросмотр без конвертации |
| `--overwrite` | | Перезаписывать существующие файлы |
| `--resume` | | Пропускать уже сконвертированные (по умолчанию) |
| `--no-report` | | Не создавать CSV отчёт |
| `--word` | | Использовать Word COM (Windows) |
| `--libreoffice` | | Использовать LibreOffice |

### Примеры

```bash
# Предпросмотр файлов для конвертации
python main.py --dry-run

# Параллельная конвертация с 4 воркерами через LibreOffice
python main.py --workers 4 --libreoffice

# Перезаписать все PDF файлы с детальным выводом
python main.py --overwrite --verbose

# Конвертация с таймаутом 60 секунд на файл
python main.py --timeout 60
```

---

## ⚙️ Конфигурация

### Файл config.py

```python
# Входная и выходная директории
INPUT_DIR = Path(r"C:\Users\...\Documents\docx")
OUTPUT_DIR = Path(r"C:\Users\...\Documents\pdf")

# Количество параллельных процессов
MAX_WORKERS = 2

# Бэкенд по умолчанию: "word", "libreoffice", "auto"
DEFAULT_BACKEND = "word"

# Метаданные PDF
METADATA_CONFIG = MetadataConfig(
    author="Ваше Имя",
    organization="Название организации",
    subject="Тема документа",
    keywords="ключевые, слова"
)

# Записывать метаданные в PDF
WRITE_METADATA = True
```

### Рекомендации по количеству воркеров

| Конфигурация | Воркеров | Пояснение |
|--------------|----------|-----------|
| 2-4 GB RAM | 1 | Минимальная нагрузка |
| 8 GB RAM | 1-2 | Рекомендуется |
| 16 GB RAM | 2-3 | Оптимально |
| 32+ GB RAM | 3-4 | Максимальная производительность |

> ⚠️ **Важно:** Word COM не является потокобезопасным. Использование более 3 воркеров может привести к нестабильной работе.

---

## 📁 Структура проекта

```
docx2pdf/
├── main.py                    # Точка входа
├── config.py                  # Конфигурация
├── requirements.txt           # Зависимости
│
├── converters/                # Модуль конвертеров
│   ├── __init__.py
│   ├── base.py                # Базовый абстрактный класс
│   ├── factory.py             # Фабрика конвертеров
│   ├── word.py                # Windows Word COM бэкенд
│   └── libreoffice.py         # LibreOffice бэкенд
│
├── utils/                     # Служебные модули
│   ├── __init__.py
│   ├── cli.py                 # Парсер аргументов CLI
│   ├── logger.py              # Настройка логирования
│   ├── scanner.py             # Сканирование файлов
│   └── word_utils.py          # Утилиты для работы с Word
│
└── output/                    # Модули вывода
    ├── __init__.py
    ├── report.py              # Генерация CSV отчётов
    └── pdf_metadata.py        # Работа с метаданными PDF
```

### Архитектура

```
┌─────────────────────────────────────────────────────────┐
│                         main.py                         │
│                    (Точка входа, CLI)                   │
└─────────────────────────────────────────────────────────┘
                              │
          ┌───────────────────┼───────────────────┐
          ▼                   ▼                   ▼
┌─────────────────┐ ┌─────────────────┐ ┌─────────────────┐
│   converters/   │ │     utils/      │ │     output/     │
│                 │ │                 │ │                 │
│ • base.py       │ │ • cli.py        │ │ • report.py     │
│ • factory.py    │ │ • logger.py     │ │ • pdf_metadata  │
│ • word.py       │ │ • scanner.py    │ │                 │
│ • libreoffice   │ │ • word_utils.py │ │                 │
└─────────────────┘ └─────────────────┘ └─────────────────┘
          │
          ▼
┌─────────────────────────────────────────────────────────┐
│                    Бэкенды конвертера                   │
│                                                         │
│   ┌─────────────────┐             ┌─────────────────┐   │
│   │    Word COM     │             │   LibreOffice   │   │
│   │   (Windows)     │             │  (Cross-plat)   │   │
│   └─────────────────┘             └─────────────────┘   │
└─────────────────────────────────────────────────────────┘
```

---

## 🔧 Расширенные возможности

### Проверка метаданных PDF

```python
from output.pdf_metadata import print_metadata
from pathlib import Path

print_metadata(Path("документ.pdf"))
```

**Вывод:**
```
📄 Метаданные: документ.pdf
--------------------------------------------------
  Автор: Пупкин Василий
  Название: документ
  Тема: Меню-требование
  Ключевые слова: меню, столовая, питание
  Организация: ООО «Столовая № 1»
```

### Программное использование

```python
from pathlib import Path
from converters import create_converter, ConverterBackend

# Создание конвертера
converter = create_converter(
    backend=ConverterBackend.WORD_COM,
    input_root=Path("./documents"),
    output_root=Path("./pdf"),
    retry_count=3,
    timeout=120
)

# Конвертация одного файла
result = converter.convert(Path("./documents/file.docx"))

print(f"Статус: {result.status.value}")
print(f"Время: {result.duration_seconds:.2f} сек")
```

### CSV отчёт

После конвертации создаётся файл `report.csv`:

```csv
file,status,error,duration_sec
document1.docx,success,,45.23
document2.docx,success,,38.15
document3.docx,failed,Ошибка доступа,0.00
```

---

## 🎯 Производительность

### Тестовые данные

| Параметр | Значение |
|----------|----------|
| Количество файлов | 6 |
| Средний размер файла | 46.8 KB |
| Содержимое | 5 страниц таблиц |
| Платформа | Windows 11 |

### Результаты

| Конфигурация | Время | Скорость |
|--------------|-------|----------|
| 1 воркер, Word COM | 392 сек | 65 сек/файл |
| 2 воркера, Word COM | ~200 сек | ~33 сек/файл |

---

## 🐛 Устранение неполадок

### Word COM не найден (Windows)

```
❌ Нет доступного бэкенда конвертера!
```

**Решение:** Убедитесь, что Microsoft Word установлен и активирован.

### LibreOffice не найден

**Решение:** Установите LibreOffice и убедитесь, что `soffice` доступен в PATH:

```bash
# Linux
sudo apt install libreoffice

# macOS
brew install --cask libreoffice
```

### Ошибка доступа к файлу

```
[WinError 5] Отказано в доступе
```

**Решение:** Скрипт автоматически повторяет запись с задержкой. Если ошибка сохраняется — закройте все экземпляры Word и повторите.

### Медленная конвертация

- Уменьшите количество воркеров
- Увеличьте таймаут для больших файлов
- Проверьте свободное место на диске

---

## 📋 Сравнение бэкендов

| Характеристика | Word COM | LibreOffice |
|----------------|----------|-------------|
| Платформа | Только Windows | Все платформы |
| Скорость запуска | Быстро | Медленнее |
| Качество PDF | Отличное | Хорошее |
| Параллелизм | 1-3 воркера | 2-4 воркера |
| Форматирование | Идеальное | Хорошее |
| Зависимости | MS Office | LibreOffice |

---

## 📄 Лицензия

Этот проект распространяется под лицензией MIT. Подробнее см. в файле [LICENSE](LICENSE).

---

## 👤 Автор

**Олег Ковалев**

- GitHub: [@okovalyoff](https://github.com/okovalyoff)
- Telegram: [@Oleg_K79](https://t.me/oleg_k79)
- E-mail: [o.kovalyoff@gmail.com](mailto:o.kovalyoff@gmail.com)

---

## 🙏 Благодарности

- [pythoncom/pywin32](https://github.com/mhammond/pywin32) — работа с Windows COM
- [pypdf](https://github.com/py-pdf/pypdf) — манипуляция PDF метаданными
- [tqdm](https://github.com/tqdm/tqdm) — красивые прогресс-бары
- [LibreOffice](https://www.libreoffice.org/) — кроссплатформенная конвертация

---

![Status](https://badgen.net/badge/Maintenance/Active/green)
![Docs](https://badgen.net/badge/Documentation/Complete/blue)

---

<p align="center">
  <b>⭐ Если проект был полезен, поставьте звезду!</b>
</p>
