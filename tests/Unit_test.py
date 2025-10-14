import os
import sys

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

import unittest  # noqa: E402

# Импортируем ваши модули
from logic.exel_utils import create_or_load_workbook  # noqa: E402
from logic.Mod1 import mod1, mod1_text  # noqa: E402
from logic.Mod2 import mod2  # noqa: E402
from logic.Mod3 import mod3  # noqa: E402
from logic.Mod4 import mod4  # noqa: E402
from logic.Mod5 import mod5  # noqa: E402


class TestMedicalModules(unittest.TestCase):
    def test_mod1(self):
        result = mod1("Муж", 55, 80, 175, 90, 85, 10.5, 25, 12, 30, 40, 18)
        self.assertIsInstance(result, float)

        prognosis, recommendations = mod1_text(1.5)
        self.assertEqual(prognosis, "Благоприятная")
        self.assertIsInstance(recommendations, list)

    def test_prognosis_calculation(self):
        result = mod1("Муж", 55, 80, 175, 90, 85, 10.5, 25, 12, 30, 40, 18)
        self.assertIsInstance(result, float)

    def test_mod2(self):
        result = mod2(2, "CYP 2c19*1")
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)

        result = mod2(15, "CYP 2c19*2")
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)

    def test_mod3(self):
        result = mod3(10, "TT")
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)

        result = mod3(100, "CC")
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)

    def test_mod4(self):
        result = mod4(10)
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 4)

    def test_mod5(self):
        result = mod5(10)
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 4)

    def test_excel_utils(self):
        wb, ws = create_or_load_workbook()
        self.assertIsNotNone(wb)
        self.assertIsNotNone(ws)


if __name__ == "__main__":
    unittest.main()
