import sys
import os

sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '../logic')))

from logic.Main import MainWindow

import unittest
from PySide6.QtWidgets import QApplication
from PySide6.QtTest import QTest
from PySide6.QtCore import Qt

app = QApplication(sys.argv)

class TestMainWindow(unittest.TestCase):
    def setUp(self):
        self.window = MainWindow()
        self.window.show()

    def tearDown(self):
        self.window.close()

    def enter_AgeField(self):
        age_field = self.window.age
        QTest.keyClicks(age_field, "35")
        self.assertEqual(age_field.text(), "35")
    def enter_Weight(self):
        weight = self.window.weight
        QTest.keyClicks(weight, "120")
        self.assertEqual(self.window.weight.text(), "120")

    def test_multiple_drug_choices(self):
        self.window.drug_aspirin.setChecked(True)
        self.window.drug_aspirin_clopidogrel.setChecked(True)
        self.assertTrue(self.window.drug_aspirin.isChecked())
        self.assertTrue(self.window.drug_aspirin_clopidogrel.isChecked())
        self.assertFalse(self.window.drug_clopidogrel.isChecked())
        self.assertFalse(self.window.drug_aspirin_ticagrelor.isChecked())
        drugs = []
        if self.window.drug_aspirin.isChecked():
            drugs.append("АСК")
        if self.window.drug_clopidogrel.isChecked():
            drugs.append("клопидогрел")
        if self.window.drug_aspirin_clopidogrel.isChecked():
            drugs.append("АСК+клопидогрел")
        if self.window.drug_aspirin_ticagrelor.isChecked():
            drugs.append("АСК+тикагрелор")
        self.assertEqual(drugs, ["АСК", "АСК+клопидогрел"])
    def test_vakidate_value(self):
        self.window.age.setText('150')
        # Проверяем, что поле age имеет красный фон при неверном значении
        self.assertIn("background-color: #ffcccc", self.window.age.styleSheet())
        
        # Проверяем, что поле weight не имеет красного фона при пустом значении
        self.window.weight.setText('')
        self.assertNotIn("background-color: #ffcccc", self.window.weight.styleSheet())

if __name__ == '__main__':
    unittest.main()