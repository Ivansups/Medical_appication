import re
import unittest
import sys
import os
import tempfile
from PySide6.QtWidgets import QApplication
from PySide6.QtTest import QTest
from PySide6.QtCore import Qt

# Импортируем ваши модули
from logic.Main import MainWindow
from logic.Mod1 import mod1_first, mod1_first_ABCB1, mod1_second
from logic.Mod2 import mod2_first
from logic.Mod3 import mod3_first
from logic.Prognosis import calculate_prognosis, prognosis_text
from logic.exel_utils import append_patient_data, create_or_load_workbook

class TestMedicalModules(unittest.TestCase):
    def test_mod1_first(self):
        result = mod1_first(25, "CYP 2c19*1")
        self.assertEqual(
            result,
            (
                "Нормальный метаболизм клопидогрела",
                [
                    "Заменить клопидогрел на оригинальный препарат (Плавикс) или препарат другого производителя",
                    "Контроль агрегации тромбоцитов через 7 дней терапии"
                ]
            )
        )
        result = mod1_first(25, "CYP 2c19*2")
        self.assertEqual(
            result,
            (
                "Замедление метаболизма клопидогрела",
                [
                    "Заменить АСК+клопидогрел на комбинацию препаратов АСК+тикагрелор",
                    "Контроль агрегации тромбоцитов через 7 дней терапии"
                ]
            )
        )
    
    def test_prognosis_calculation(self):
        result = calculate_prognosis("Муж", 55, 80, 175, 90, 85, 10.5, 25, 12, 30, 40, 18)
        self.assertAlmostEqual(result, 1.909, places=3)    
    
    def test_mod2_first(self):
        result = mod2_first(2)
        self.assertEqual(
            result,
            (
                "Агрегация тромбоцитов значительно подавлена (T ≤ 10%)",
                "Продолжить прием ацетилсалициловой кислоты",
                "Высокий риск геморрагических осложнений"
            )
        )
        result = mod2_first(15)
        self.assertEqual(
            result,
            (
                "Агрегация тромбоцитов сохранена",
                "Замена на препарат ацетилсалициловой кислоты другого производителя",
                "Контроль агрегации тромбоцитов через 7 дней терапии"
            )
        )
        result = mod2_first(8)
        self.assertEqual(
            result,
            (
                "Агрегация тромбоцитов сохранена",
                "Замена на препарат ацетилсалициловой кислоты другого производителя",
                "Контроль агрегации тромбоцитов через 7 дней терапии"
            )
        )
        result = mod2_first(-10)
        self.assertEqual(
            result,
            (
                "Агрегация тромбоцитов сохранена",
                "Замена на препарат ацетилсалициловой кислоты другого производителя",
                "Контроль агрегации тромбоцитов через 7 дней терапии"
            )
        )

    def test_mod3_first(self):
        result = mod3_first(10)
        self.assertEqual(
            result,
            (
                "Агрегация тромбоцитов значительно подавлена (T ≤ 10%)",
                "Высокий риск геморрагических осложнений",
                "Замена терапии на АСК+клопидогрел",
                "Контроль агрегации тромбоцитов через 7 дней терапии"
            )
        )
        result = mod3_first(100)
        self.assertEqual(
            result,
            (
                "Нет специфических рекомендаций",
                "Возможно, введены некорректные данные",
                "",
                ""
            )
        )



if __name__ == '__main__':
    unittest.main()