# Medical Application

Медицинское приложение для анализа данных пациентов с использованием PySide6 GUI.

## Описание проекта

Приложение предназначено для:
- Ввода и анализа данных пациентов
- Расчетов медицинских показателей (СКФ, клиренс креатинина)
- Генетического анализа (CYP2C19, ABCB1)
- Прогнозирования результатов лечения
- Экспорта данных в Excel и Word форматы

## Зависимости

- Python 3.8+
- PySide6 - GUI фреймворк
- openpyxl - работа с Excel файлами
- python-docx - работа с Word документами
- pyinstaller - создание исполняемых файлов

## Установка и настройка

### 1. Клонирование репозитория
```bash
git clone <repository-url>
cd Medical_appication
```

### 2. Создание виртуального окружения
```bash
python3 -m venv venv
```

### 3. Активация виртуального окружения
```bash
# macOS/Linux
source venv/bin/activate

# Windows
venv\Scripts\activate
```

### 4. Установка зависимостей
```bash
pip install -r requirements.txt
```

## Запуск приложения

```bash
python Main.py
```

## Сборка исполняемого файла

```bash
pyinstaller Main.spec
```

Исполняемый файл будет создан в папке `dist/`.

## Структура проекта

- `Main.py` - главный файл приложения
- `logic/` - модули логики приложения
  - `Mod1.py`, `Mod2.py`, `Mod3.py` - модули расчетов
  - `exel_utils.py` - утилиты для работы с Excel
  - `Prognosis.py` - модуль прогнозирования
- `classes/` - классы данных
  - `Patient.py` - класс пациента с данными и перечислениями
- `tests/` - тесты
- `build/` - файлы сборки PyInstaller


## Команды для линтера (Важно запускать при каждом изменении кода!)
```bash
pre-commit run --all-files
```
