import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

DEFAULT_FILENAME = "patients.xlsx"

def create_or_load_workbook(filename=DEFAULT_FILENAME):
    if os.path.exists(filename):
        wb = load_workbook(filename)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.append([
            "Пол", "Возраст", "Вес", "Рост", "Креатинин", "Клиренс креатинина", "MPV", "PLCR",
            "Спонтанная агрегация", "Индуц. агрегация 1 мкМоль АДФ", "Индуц. агрегация 5 мкМоль АДФ",
            "Индуц. агрегация 15 мкл арахидоновой кислоты", "Генотип CYP2C19", "Генотип ABCB1",
            "Препараты", "Состояние агрегации", "Скорость выведения клопидогрела (ABCB1)",
            "Модуль 1", "Модуль 2", "Модуль 3", "Коэффициент прогноза", "Оценка прогноза"
        ])
    return wb, ws

def append_patient_data(
    filename, data_row
):
    wb, ws = create_or_load_workbook(filename)
    ws.append(data_row)
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except Exception:
                pass
        adjusted_width = max_length + 2  # небольшой запас
        ws.column_dimensions[col_letter].width = adjusted_width
    wb.save(filename)
