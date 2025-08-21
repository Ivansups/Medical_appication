from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QFormLayout, QLineEdit, QComboBox, QPushButton, QTextEdit, QCheckBox, QFileDialog, QMessageBox, QGroupBox, QHBoxLayout, QLabel, QScrollArea
)
from PySide6.QtCore import Qt, QDate
import sys

from logic.Class import PatientData, Gender, CYP2C19
from logic.Mod1 import mod1, mod1_text
from logic.Mod2 import mod2
from logic.Mod3 import mod3
from logic.Mod4 import mod4
from logic.Mod5 import mod5
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
        if ws is None:
            ws = wb.create_sheet("Sheet1")
        ws.append([
            "Пол", "Возраст", "Вес", "Рост", "Креатинин", "Клиренс креатинина", "MPV", "PLCR",
            "Спонтанная агрегация", "Индуц. агрегация 1 мкМоль АДФ", "Индуц. агрегация 5 мкМоль АДФ",
            "Индуц. агрегация 15 мкл арахидоновой кислоты", "Генотип CYP2C19", "Генотип ABCB1",
            "Препараты", "Состояние агрегации", "Скорость выведения клопидогрела (ABCB1)",
            "Модуль 1", "Модуль 2", "Модуль 3", "Коэффициент прогноза", "Оценка прогноза"
        ])
    return wb, ws

def autofit_columns(ws):
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column  # номер колонки (1, 2, 3...)
        for cell in column_cells:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        adjusted_width = max_length + 2  # небольшой запас
        ws.column_dimensions[get_column_letter(column)].width = adjusted_width

def append_patient_data(filename, data_row):
    wb, ws = create_or_load_workbook(filename)
    if ws is not None:
        ws.append(data_row)
        autofit_columns(ws)  # <--- вот здесь!
        wb.save(filename)
    else:
        raise ValueError("Не удалось создать или получить рабочий лист Excel")

# Удалена функция create_mpv_chart

class ReportWindow(QWidget):
    def __init__(self, report_text, patient_data=None, excel_filename="patients.xlsx"):
        super().__init__()
        self.setWindowTitle("📋 Полный медицинский отчет по пациенту")
        self.resize(900, 700)
        self.patient_data = patient_data
        self.excel_filename = excel_filename
        
        # Главный layout
        main_layout = QVBoxLayout(self)
        
        # Заголовок
        header_label = QLabel("МЕДИЦИНСКИЙ ОТЧЕТ ПО ПАЦИЕНТУ")
        header_label.setStyleSheet("""
            QLabel {
                font-size: 18px;
                font-weight: bold;
                color: #2c3e50;
                background-color: #ecf0f1;
                padding: 15px;
                border: 2px solid #3498db;
                border-radius: 8px;
                margin: 5px;
            }
        """)
        header_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_layout.addWidget(header_label)
        
        # Пример: добавляем диаграмму MPV, если есть данные
        if patient_data and len(patient_data) > 6 and patient_data[6]:
            try:
                mpv_value = float(patient_data[6])
                # chart_label = create_mpv_chart(mpv_value) # Удалено
                # main_layout.addWidget(chart_label) # Удалено
                pass # Удалено
            except Exception as e:
                print(f'Ошибка построения диаграммы MPV: {e}')
        
        # Область прокрутки для текста
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # Текстовое поле
        self.text = QTextEdit()
        self.text.setReadOnly(True)
        self.text.setStyleSheet("""
            QTextEdit {
                font-family: 'Consolas', 'Monaco', 'Courier New', monospace;
                font-size: 12px;
                line-height: 1.4;
                background-color: #f8f9fa;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                padding: 15px;
                color: #2c3e50;
            }
            QTextEdit:focus {
                border: 2px solid #3498db;
            }
        """)
        
        # Форматируем текст отчета
        formatted_text = self.format_report_text(report_text)
        self.text.setText(formatted_text)
        
        scroll_area.setWidget(self.text)
        main_layout.addWidget(scroll_area)
        
        # Кнопки действий
        button_layout = QHBoxLayout()
        
        # Кнопка копирования
        copy_button = QPushButton("📋 Копировать отчет")
        copy_button.clicked.connect(self.copy_to_clipboard)
        copy_button.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border: 2px solid #27ae60;
                border-radius: 5px;
                background-color: #27ae60;
                color: white;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #229954;
            }
            QPushButton:pressed {
                background-color: #1e8449;
            }
        """)
        
        # Кнопка сохранения в файл
        save_button = QPushButton("💾 Сохранить в файл")
        save_button.clicked.connect(self.save_to_file)
        save_button.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border: 2px solid #3498db;
                border-radius: 5px;
                background-color: #3498db;
                color: white;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
        """)
        
        # Кнопка сохранения в Excel
        excel_button = QPushButton("📊 Сохранить в Excel")
        excel_button.clicked.connect(self.save_to_excel)
        excel_button.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border: 2px solid #f39c12;
                border-radius: 5px;
                background-color: #f39c12;
                color: white;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #e67e22;
            }
            QPushButton:pressed {
                background-color: #d35400;
            }
        """)
        
        
        # Кнопка закрытия
        close_button = QPushButton("❌ Закрыть")
        close_button.clicked.connect(self.close)
        close_button.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border: 2px solid #e74c3c;
                border-radius: 5px;
                background-color: #e74c3c;
                color: white;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #c0392b;
            }
            QPushButton:pressed {
                background-color: #a93226;
            }
        """)
        
        button_layout.addWidget(copy_button)
        button_layout.addWidget(save_button)
        button_layout.addWidget(excel_button)
        button_layout.addStretch()
        button_layout.addWidget(close_button)
        
        main_layout.addLayout(button_layout)
        
        # Стили для окна
        self.setStyleSheet("""
            QWidget {
                background-color: #ecf0f1;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                background-color: #f0f0f0;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;
            }
        """)
    
    def format_report_text(self, text):
        """Форматирует текст отчета для лучшего отображения"""
        # Заменяем разделители на более красивые
        text = text.replace("==============================", "╔══════════════════════════════════════════════════════════════╗")
        text = text.replace("------------------------------", "╟──────────────────────────────────────────────────────────────╢")
        
        # Добавляем цветовое выделение для заголовков
        lines = text.split('\n')
        formatted_lines = []
        
        for line in lines:
            if line.strip().startswith('I.') or line.strip().startswith('II.') or line.strip().startswith('III.') or \
               line.strip().startswith('IV.') or line.strip().startswith('V.') or line.strip().startswith('VI.') or \
               line.strip().startswith('VII.') or line.strip().startswith('VIII.'):
                formatted_lines.append(f"<h3 style='color: #2c3e50; background-color: #ecf0f1; padding: 5px; border-radius: 3px;'>{line}</h3>")
            elif line.strip().startswith('МЕДИЦИНСКИЙ ОТЧЕТ'):
                formatted_lines.append(f"<h2 style='color: #3498db; text-align: center; font-size: 16px;'>{line}</h2>")
            elif line.strip().startswith('Модуль'):
                formatted_lines.append(f"<h4 style='color: #e67e22;'>{line}</h4>")
            elif line.strip().startswith('Коэффициент прогноза:'):
                formatted_lines.append(f"<p style='color: #27ae60; font-weight: bold;'>{line}</p>")
            elif line.strip().startswith('Оценка:'):
                formatted_lines.append(f"<p style='color: #27ae60; font-weight: bold;'>{line}</p>")
            elif line.strip().startswith('╔') or line.strip().startswith('╟'):
                formatted_lines.append(f"<p style='color: #7f8c8d; font-family: monospace;'>{line}</p>")
            else:
                formatted_lines.append(f"<p>{line}</p>")
        
        return '\n'.join(formatted_lines)
    
    def copy_to_clipboard(self):
        """Копирует отчет в буфер обмена"""
        clipboard = QApplication.clipboard()
        clipboard.setText(self.text.toPlainText())
        QMessageBox.information(self, "Копирование", "Отчет скопирован в буфер обмена!")
    
    def save_to_file(self):
        """Сохраняет отчет в текстовый файл"""
        filename, _ = QFileDialog.getSaveFileName(
            self, 
            "Сохранить отчет", 
            f"медицинский_отчет_{QDate.currentDate().toString('yyyy-MM-dd')}.txt",
            "Текстовые файлы (*.txt);;Все файлы (*)"
        )
        if filename:
            try:
                with open(filename, 'w', encoding='utf-8') as f:
                    f.write(self.text.toPlainText())
                QMessageBox.information(self, "Сохранение", f"Отчет сохранен в файл:\n{filename}")
            except Exception as e:
                QMessageBox.critical(self, "Ошибка", f"Не удалось сохранить файл:\n{str(e)}")
    
    def save_to_excel(self):
        """Сохраняет данные пациента в Excel файл"""
        if not self.patient_data:
            QMessageBox.warning(self, "Предупреждение", "Данные пациента недоступны для сохранения в Excel")
            return
        
        try:
            # Используем данные пациента для сохранения в Excel
            data_row = self.patient_data
            
            # Сохраняем в Excel
            append_patient_data(self.excel_filename, data_row)
            
            QMessageBox.information(self, "Сохранение в Excel", 
                                  f"Данные пациента успешно сохранены в файл:\n{self.excel_filename}")
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", 
                               f"Не удалось сохранить данные в Excel:\n{str(e)}")
            print(f"Ошибка сохранения в Excel: {e}")
            import traceback
            traceback.print_exc()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Медицинское приложение")
        self.resize(1000, 800)
        
        # Создаем главный layout
        main_layout = QVBoxLayout(self)
        
        # Создаем область прокрутки
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # Создаем контейнер для содержимого
        content_widget = QWidget()
        layout = QVBoxLayout(content_widget)

        # === ГРУППА 1: ОСНОВНЫЕ ДАННЫЕ ПАЦИЕНТА ===
        basic_group = QGroupBox("Основные данные пациента")
        basic_layout = QFormLayout()
        
        # Поля выбора
        self.gender = QComboBox()
        self.gender.addItem("")  # Для необязательного выбора
        self.gender.addItems([g.value for g in Gender])
        basic_layout.addRow("Пол (выберите):", self.gender)
        
        # Поля ввода
        self.age = QLineEdit()
        self.age.setPlaceholderText("Введите возраст (лет)")
        basic_layout.addRow("Возраст:", self.age)

        self.weight = QLineEdit()
        self.weight.setPlaceholderText("Введите вес (кг)")
        basic_layout.addRow("Вес:", self.weight)

        self.height_field = QLineEdit()
        self.height_field.setPlaceholderText("Введите рост (см)")
        basic_layout.addRow("Рост:", self.height_field)
        
        basic_group.setLayout(basic_layout)
        layout.addWidget(basic_group)

        # === ГРУППА 2: ГЕНОТИПЫ ===
        genotype_group = QGroupBox("Генотипы")
        genotype_layout = QFormLayout()
        
        self.cyp2c19 = QComboBox()
        self.cyp2c19.addItem("")
        self.cyp2c19.addItems([c.value for c in CYP2C19])
        genotype_layout.addRow("Генотип CYP2C19:", self.cyp2c19)

        self.abcb1 = QComboBox()
        self.abcb1.addItem("")
        self.abcb1.addItems(["TT", "TC", "CC"])
        genotype_layout.addRow("Генотип ABCB1:", self.abcb1)
        
        genotype_group.setLayout(genotype_layout)
        layout.addWidget(genotype_group)

        # === ГРУППА 3: БИОХИМИЧЕСКИЕ ПОКАЗАТЕЛИ ===
        bio_group = QGroupBox("Биохимические показатели")
        bio_layout = QFormLayout()
        
        self.creatinine = QLineEdit()
        self.creatinine.setPlaceholderText("Введите креатинин (мкмоль/л)")
        bio_layout.addRow("Креатинин:", self.creatinine)

        self.creatinine_clearance = QLineEdit()
        self.creatinine_clearance.setPlaceholderText("Введите клиренс креатинина (мл/мин)")
        bio_layout.addRow("Клиренс креатинина:", self.creatinine_clearance)
        
        bio_group.setLayout(bio_layout)
        layout.addWidget(bio_group)

        # === ГРУППА 4: ТРОМБОЦИТАРНЫЕ ПОКАЗАТЕЛИ ===
        platelet_group = QGroupBox("Тромбоцитарные показатели")
        platelet_layout = QFormLayout()
        
        self.mpv = QLineEdit()
        self.mpv.setPlaceholderText("Введите MPV (фл)")
        platelet_layout.addRow("Величина тромбоцитов MPV:", self.mpv)

        self.plcr = QLineEdit()
        self.plcr.setPlaceholderText("Введите PLCR (%)")
        platelet_layout.addRow("Отн. кол-во больших тромбоцитов PLCR:", self.plcr)
        
        platelet_group.setLayout(platelet_layout)
        layout.addWidget(platelet_group)

        # === ГРУППА 5: АГРЕГАЦИЯ ТРОМБОЦИТОВ ===
        aggregation_group = QGroupBox("Агрегация тромбоцитов")
        aggregation_layout = QFormLayout()
        
        self.spontaneous_aggregation = QLineEdit()
        self.spontaneous_aggregation.setPlaceholderText("Введите спонтанную агрегацию (усл.ед.)")
        aggregation_layout.addRow("Спонтанная агрегация:", self.spontaneous_aggregation)

        self.induced_aggregation_1_ADP = QLineEdit()
        self.induced_aggregation_1_ADP.setPlaceholderText("Введите % агрегации")
        aggregation_layout.addRow("Индуц. агрегация 1 мкМоль АДФ:", self.induced_aggregation_1_ADP)

        self.induced_aggregation_5_ADP = QLineEdit()
        self.induced_aggregation_5_ADP.setPlaceholderText("Введите % агрегации")
        aggregation_layout.addRow("Индуц. агрегация 5 мкМоль АДФ:", self.induced_aggregation_5_ADP)

        self.induced_aggregation_15_ARA = QLineEdit()
        self.induced_aggregation_15_ARA.setPlaceholderText("Введите % агрегации")
        aggregation_layout.addRow("Индуц. агрегация 15 мкл арахидоновой кислоты:", self.induced_aggregation_15_ARA)
        
        aggregation_group.setLayout(aggregation_layout)
        layout.addWidget(aggregation_group)

        # === ГРУППА 6: ПРЕПАРАТЫ ===
        drugs_group = QGroupBox("Препараты")
        drugs_layout = QVBoxLayout()
        
        drugs_label = QLabel("Выберите принимаемые препараты:")
        drugs_layout.addWidget(drugs_label)
        
        self.drug_aspirin = QCheckBox("АСК")
        self.drug_clopidogrel = QCheckBox("Клопидогрел")
        self.drug_aspirin_clopidogrel = QCheckBox("АСК+клопидогрел")
        self.drug_aspirin_ticagrelor = QCheckBox("АСК+тикагрелор")
        
        drugs_layout.addWidget(self.drug_aspirin)
        drugs_layout.addWidget(self.drug_clopidogrel)
        drugs_layout.addWidget(self.drug_aspirin_clopidogrel)
        drugs_layout.addWidget(self.drug_aspirin_ticagrelor)
        
        drugs_group.setLayout(drugs_layout)
        layout.addWidget(drugs_group)

        # === ГРУППА 7: ДЕЙСТВИЯ ===
        actions_group = QGroupBox("⚙️ Действия")
        actions_layout = QVBoxLayout()
        
        self.report_button = QPushButton("📄 Сформировать полный отчет")
        self.report_button.clicked.connect(self.generate_report)
        actions_layout.addWidget(self.report_button)

        self.save_excel_button = QPushButton("💾 Сохранить в Excel")
        self.save_excel_button.clicked.connect(self.save_to_excel)
        actions_layout.addWidget(self.save_excel_button)

        self.choose_excel_button = QPushButton("📁 Выбрать файл Excel")
        self.choose_excel_button.clicked.connect(self.choose_excel_file)
        actions_layout.addWidget(self.choose_excel_button)
        
        actions_group.setLayout(actions_layout)
        layout.addWidget(actions_group)

        # Устанавливаем контейнер в область прокрутки
        scroll_area.setWidget(content_widget)
        
        # Добавляем область прокрутки в главный layout
        main_layout.addWidget(scroll_area)

        self.excel_filename = DEFAULT_FILENAME

        # Применяем стили для лучшего визуального разделения
        self.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                border: 2px solid #cccccc;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #2c3e50;
            }
            QLineEdit {
                padding: 5px;
                border: 1px solid #bdc3c7;
                border-radius: 3px;
                background-color: #f8f9fa;
            }
            QLineEdit:focus {
                border: 2px solid #3498db;
                background-color: white;
            }
            QComboBox {
                padding: 5px;
                border: 1px solid #bdc3c7;
                border-radius: 3px;
                background-color: #ecf0f1;
            }
            QComboBox:focus {
                border: 2px solid #3498db;
            }
            QPushButton {
                padding: 8px 16px;
                border: 2px solid #3498db;
                border-radius: 5px;
                background-color: #3498db;
                color: white;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #21618c;
            }
            QCheckBox {
                spacing: 8px;
                font-weight: normal;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QScrollBar:vertical {
                background-color: #f0f0f0;
                width: 12px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #c0c0c0;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #a0a0a0;
            }
        """)

        # Подключение валидации к полям
        self.age.textChanged.connect(self.validate_age)
        self.weight.textChanged.connect(self.validate_weight)
        self.height_field.textChanged.connect(self.validate_height)
        self.creatinine.textChanged.connect(self.validate_creatinine)
        self.creatinine_clearance.textChanged.connect(self.validate_creatinine_clearance)
        self.mpv.textChanged.connect(self.validate_mpv)
        self.plcr.textChanged.connect(self.validate_plcr)
        self.spontaneous_aggregation.textChanged.connect(self.validate_spontaneous_aggregation)
        self.induced_aggregation_1_ADP.textChanged.connect(self.validate_induced_aggregation_1_ADP)
        self.induced_aggregation_5_ADP.textChanged.connect(self.validate_induced_aggregation_5_ADP)
        self.induced_aggregation_15_ARA.textChanged.connect(self.validate_induced_aggregation_15_ARA)

    def validate_age(self):
        """Проверка возраста"""
        try:
            age = int(self.age.text())
            if age <= 0 or age > 120:
                self.age.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.age.setStyleSheet("")
                return True
        except ValueError:
            self.age.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
            return False

    def validate_weight(self):
        """Проверка веса"""
        try:
            weight = float(self.weight.text())
            if weight <= 0 or weight > 300:
                self.weight.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.weight.setStyleSheet("")
                return True
        except ValueError:
            self.weight.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
            return False

    def validate_height(self):
        """Проверка роста"""
        try:
            height = float(self.height_field.text())
            if height <= 0 or height > 250:
                self.height_field.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.height_field.setStyleSheet("")
                return True
        except ValueError:
            if self.height_field.text():
                self.height_field.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.height_field.setStyleSheet("")
                return True

    def validate_creatinine(self):
        """Проверка креатинина"""
        try:
            creatinine = float(self.creatinine.text())
            if creatinine <= 0 or creatinine > 1000:
                self.creatinine.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.creatinine.setStyleSheet("")
                return True
        except ValueError:
            if self.creatinine.text():  # Только если поле не пустое
                self.creatinine.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.creatinine.setStyleSheet("")
                return True

    def validate_creatinine_clearance(self):
        """Проверка клиренса креатинина"""
        try:
            clearance = float(self.creatinine_clearance.text())
            if clearance <= 0 or clearance > 200:
                self.creatinine_clearance.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.creatinine_clearance.setStyleSheet("")
                return True
        except ValueError:
            if self.creatinine_clearance.text():
                self.creatinine_clearance.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.creatinine_clearance.setStyleSheet("")
                return True

    def validate_mpv(self):
        """Проверка MPV"""
        try:
            mpv = float(self.mpv.text())
            if mpv <= 0 or mpv > 20:
                self.mpv.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.mpv.setStyleSheet("")
                return True
        except ValueError:
            if self.mpv.text():
                self.mpv.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.mpv.setStyleSheet("")
                return True

    def validate_plcr(self):
        """Проверка PLCR"""
        try:
            plcr = float(self.plcr.text())
            if plcr < 0 or plcr > 100:
                self.plcr.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.plcr.setStyleSheet("")
                return True
        except ValueError:
            if self.plcr.text():
                self.plcr.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.plcr.setStyleSheet("")
                return True

    def validate_spontaneous_aggregation(self):
        """Проверка спонтанной агрегации"""
        try:
            agg = float(self.spontaneous_aggregation.text())
            if agg < 0 or agg > 100:
                self.spontaneous_aggregation.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.spontaneous_aggregation.setStyleSheet("")
                return True
        except ValueError:
            if self.spontaneous_aggregation.text():
                self.spontaneous_aggregation.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.spontaneous_aggregation.setStyleSheet("")
                return True

    def validate_induced_aggregation_1_ADP(self):
        """Проверка индуцированной агрегации 1 мкМоль АДФ"""
        try:
            agg = float(self.induced_aggregation_1_ADP.text())
            if agg < 0 or agg > 100:
                self.induced_aggregation_1_ADP.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.induced_aggregation_1_ADP.setStyleSheet("")
                return True
        except ValueError:
            if self.induced_aggregation_1_ADP.text():
                self.induced_aggregation_1_ADP.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.induced_aggregation_1_ADP.setStyleSheet("")
                return True

    def validate_induced_aggregation_5_ADP(self):
        """Проверка индуцированной агрегации 5 мкМоль АДФ"""
        try:
            agg = float(self.induced_aggregation_5_ADP.text())
            if agg < 0 or agg > 100:
                self.induced_aggregation_5_ADP.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.induced_aggregation_5_ADP.setStyleSheet("")
                return True
        except ValueError:
            if self.induced_aggregation_5_ADP.text():
                self.induced_aggregation_5_ADP.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.induced_aggregation_5_ADP.setStyleSheet("")
                return True

    def validate_induced_aggregation_15_ARA(self):
        """Проверка индуцированной агрегации 15 мкл арахидоновой кислоты"""
        try:
            agg = float(self.induced_aggregation_15_ARA.text())
            if agg < 0 or agg > 100:
                self.induced_aggregation_15_ARA.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.induced_aggregation_15_ARA.setStyleSheet("")
                return True
        except ValueError:
            if self.induced_aggregation_15_ARA.text():
                self.induced_aggregation_15_ARA.setStyleSheet("background-color: #ffcccc; border: 2px solid red;")
                return False
            else:
                self.induced_aggregation_15_ARA.setStyleSheet("")
                return True

    def validate_all_fields(self):
        """Проверка всех полей перед сохранением"""
        validations = [
            self.validate_age(),
            self.validate_weight(),
            self.validate_height(),
            self.validate_creatinine(),
            self.validate_creatinine_clearance(),
            self.validate_mpv(),
            self.validate_plcr(),
            self.validate_spontaneous_aggregation(),
            self.validate_induced_aggregation_1_ADP(),
            self.validate_induced_aggregation_5_ADP(),
            self.validate_induced_aggregation_15_ARA()
        ]

        if not all(validations):
            QMessageBox.warning(self, "Ошибка валидации",
                              "Пожалуйста, исправьте ошибки в полях (выделены красным)")
            return False
        return True

    def generate_report(self):
        try:
            if not self.validate_all_fields():
                return
            
            # Сбор данных
            gender = Gender(self.gender.currentText()) if self.gender.currentText() else None
            age = int(self.age.text()) if self.age.text() else None
            T = float(self.induced_aggregation_5_ADP.text()) if self.induced_aggregation_5_ADP.text() else None
            cyp = self.cyp2c19.currentText() if self.cyp2c19.currentText() else None
            cyp_enum = CYP2C19(cyp) if cyp else None

            T2 = float(self.induced_aggregation_15_ARA.text()) if self.induced_aggregation_15_ARA.text() else None
            T3 = float(self.spontaneous_aggregation.text()) if self.spontaneous_aggregation.text() else None

            patient = PatientData(
                gender=gender,
                age=age,
                T=T,
                cyp2c19=cyp_enum
            )

            # Модуль 1
            mod1_score, mod1_recommendations = mod1_first(patient.T, patient.cyp2c19.value if patient.cyp2c19 else None)
            mod1_text = f"Модуль 1:\nОценка: {mod1_score}\nРекомендации:\n" + "\n".join(mod1_recommendations)

            # Модуль 2
            mod2_res = mod2_first(T2)
            mod2_text = (
                f"Модуль 2:\n"
                f"{mod2_res[0]}\n"
                f"{mod2_res[1]}\n"
                f"{mod2_res[2] if len(mod2_res) > 2 else ''}"
            )

            # Модуль 3
            mod3_res = mod3_first(T3)
            mod3_text = (
                f"Модуль 3:\n"
                f"{mod3_res[0]}\n"
                f"{mod3_res[1]}\n"
                f"{mod3_res[2] if len(mod3_res) > 2 else ''}"
            )

            abcb1 = self.abcb1.currentText() if self.abcb1.currentText() else None
            drugs = []
            if self.drug_aspirin.isChecked():
                drugs.append("АСК")
            if self.drug_clopidogrel.isChecked():
                drugs.append("клопидогрел")
            if self.drug_aspirin_clopidogrel.isChecked():
                drugs.append("АСК+клопидогрел")
            if self.drug_aspirin_ticagrelor.isChecked():
                drugs.append("АСК+тикагрелор")

            abcb1_result = mod1_first_ABCB1(abcb1) if abcb1 else "Нет данных"

            aggregation_state = mod1_second(T)[0] if T is not None else "Нет данных"
            report = (
                "==============================\n"
                "        МЕДИЦИНСКИЙ ОТЧЕТ\n"
                "==============================\n\n"
                "I. ОБЩИЕ ДАННЫЕ ПАЦИЕНТА\n"
                "------------------------------\n"
                f"Пол: {gender.value if gender else ''}\n"
                f"Возраст: {age if age else ''}\n"
                f"Вес: {self.weight.text()}\n"
                f"Рост: {self.height_field.text()}\n"
                f"Креатинин: {self.creatinine.text()}\n"
                f"Клиренс креатинина: {self.creatinine_clearance.text()}\n"
                f"MPV: {self.mpv.text()}\n"
                f"PLCR: {self.plcr.text()}\n"
                f"Спонтанная агрегация: {self.spontaneous_aggregation.text()}\n"
                f"Индуц. агрегация 1 мкМоль АДФ: {self.induced_aggregation_1_ADP.text()}\n"
                f"Индуц. агрегация 5 мкМоль АДФ: {self.induced_aggregation_5_ADP.text()}\n"
                f"Индуц. агрегация 15 мкл арахидоновой кислоты: {self.induced_aggregation_15_ARA.text()}\n"
                "\n"
                "II. ГЕНЕТИЧЕСКИЕ ДАННЫЕ\n"
                "------------------------------\n"
                f"Генотип CYP2C19: {cyp if cyp else ''}\n"
                f"Генотип ABCB1: {abcb1 if abcb1 else ''}\n"
                "\n"
                "III. ФАРМАКОТЕРАПИЯ\n"
                "------------------------------\n"
                f"Препараты: {', '.join(drugs)}\n"
                "\n"
                "IV. СОСТОЯНИЕ АГРЕГАЦИИ ТРОМБОЦИТОВ\n"
                "------------------------------\n"
                f"{aggregation_state}\n"
                "\n"
                "V. КОРРЕКЦИЯ ФАРМАКОТЕРАПИИ КЛОПИДОГРЕЛА\n"
                "------------------------------\n"
                f"Скорость выведения клопидогрела (ABCB1): {abcb1_result}\n"
                f"{mod1_text}\n"
                "\n"
                "VI. РЕКОМЕНДАЦИИ ПО МОДУЛЮ 2 (АСК)\n"
                "------------------------------\n"
                f"{mod2_text}\n"
                "\n"
                "VII. РЕКОМЕНДАЦИИ ПО МОДУЛЮ 3 (ТИКАГРЕЛОР)\n"
                "------------------------------\n"
                f"{mod3_text}\n"
                "\n"
            )
            # Блок ПРОГНОЗ
            try:
                prognosis_value = calculate_prognosis(
                    gender.value if gender else None,
                    age,
                    float(self.weight.text()) if self.weight.text() else None,
                    float(self.height_field.text()) if self.height_field.text() else None,
                    float(self.creatinine.text()) if self.creatinine.text() else None,
                    float(self.creatinine_clearance.text()) if self.creatinine_clearance.text() else None,
                    float(self.mpv.text()) if self.mpv.text() else None,
                    float(self.plcr.text()) if self.plcr.text() else None,
                    float(self.spontaneous_aggregation.text()) if self.spontaneous_aggregation.text() else None,
                    float(self.induced_aggregation_1_ADP.text()) if self.induced_aggregation_1_ADP.text() else None,
                    float(self.induced_aggregation_5_ADP.text()) if self.induced_aggregation_5_ADP.text() else None,
                    float(self.induced_aggregation_15_ARA.text()) if self.induced_aggregation_15_ARA.text() else None,
                )
                prognosis_result = prognosis_text(prognosis_value)
                prognosis_block = (
                    "VIII. ПРОГНОЗ\n"
                    "------------------------------\n"
                    f"Коэффициент прогноза: {prognosis_value:.3f}\n"
                    f"Оценка: {prognosis_result}\n"
                    "==============================\n"
                )
            except Exception as e:
                prognosis_block = (
                    "VIII. ПРОГНОЗ\n"
                    "------------------------------\n"
                    f"Ошибка расчета коэффициента прогноза: {e}\n"
                    "==============================\n"
                )

            report += prognosis_block

            # Подготавливаем данные для Excel
            data_row = [
                gender.value if gender else '',
                age if age else '',
                self.weight.text(),
                self.height_field.text(),
                self.creatinine.text(),
                self.creatinine_clearance.text(),
                self.mpv.text(),
                self.plcr.text(),
                self.spontaneous_aggregation.text(),
                self.induced_aggregation_1_ADP.text(),
                self.induced_aggregation_5_ADP.text(),
                self.induced_aggregation_15_ARA.text(),
                cyp if cyp else '',
                abcb1 if abcb1 else '',
                ', '.join(drugs),
                aggregation_state,
                abcb1_result,
                f"{mod1_score}: {'; '.join(mod1_recommendations)}",
                "; ".join([str(x) for x in mod2_res if x]),
                "; ".join([str(x) for x in mod3_res if x]),
                prognosis_value if 'prognosis_value' in locals() else '',
                prognosis_result if 'prognosis_result' in locals() else ''
            ]

            self.report_window = ReportWindow(report, data_row, self.excel_filename)
            self.report_window.show()
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка генерации отчета", 
                               f"Произошла ошибка при формировании отчета:\n{str(e)}")
            print(f"Ошибка в generate_report: {e}")
            import traceback
            traceback.print_exc()

    def choose_excel_file(self):
        filename, _ = QFileDialog.getSaveFileName(self, "Выберите файл Excel", self.excel_filename, "Excel Files (*.xlsx)")
        if filename:
            self.excel_filename = filename

    def save_to_excel(self):
        if not self.validate_all_fields():
            return
        # Просто сохраняем в self.excel_filename, не спрашивая пользователя!
        # Соберите все данные, как в generate_report
        gender = self.gender.currentText()
        age = self.age.text()
        weight = self.weight.text()
        height = self.height_field.text()
        creatinine = self.creatinine.text()
        creatinine_clearance = self.creatinine_clearance.text()
        mpv = self.mpv.text()
        plcr = self.plcr.text()
        spontaneous_aggregation = self.spontaneous_aggregation.text()
        induced_aggregation_1_ADP = self.induced_aggregation_1_ADP.text()
        induced_aggregation_5_ADP = self.induced_aggregation_5_ADP.text()
        induced_aggregation_15_ARA = self.induced_aggregation_15_ARA.text()
        cyp = self.cyp2c19.currentText()
        abcb1 = self.abcb1.currentText()
        drugs = []
        if self.drug_aspirin.isChecked():
            drugs.append("АСК")
        if self.drug_clopidogrel.isChecked():
            drugs.append("клопидогрел")
        if self.drug_aspirin_clopidogrel.isChecked():
            drugs.append("АСК+клопидогрел")
        if self.drug_aspirin_ticagrelor.isChecked():
            drugs.append("АСК+тикагрелор")
        drugs_str = ", ".join(drugs)

        # Получить все расчеты и рекомендации (как в generate_report)
        T = float(self.induced_aggregation_5_ADP.text()) if self.induced_aggregation_5_ADP.text() else None
        aggregation_state = mod1_second(T)[0] if T is not None else "Нет данных"
        abcb1_result = mod1_first_ABCB1(abcb1) if abcb1 else "Нет данных"
        mod1_score, mod1_recommendations = mod1_first(T, cyp)
        mod1_text = f"{mod1_score}: {'; '.join(mod1_recommendations)}"
        T2 = float(self.induced_aggregation_15_ARA.text()) if self.induced_aggregation_15_ARA.text() else None
        mod2_res = mod2_first(T2)
        mod2_text = "; ".join([str(x) for x in mod2_res if x])
        T3 = float(self.spontaneous_aggregation.text()) if self.spontaneous_aggregation.text() else None
        mod3_res = mod3_first(T3)
        mod3_text = "; ".join([str(x) for x in mod3_res if x])

        try:
            prognosis_value = calculate_prognosis(
                gender,
                int(age) if age else None,
                float(weight) if weight else None,
                float(height) if height else None,
                float(creatinine) if creatinine else None,
                float(creatinine_clearance) if creatinine_clearance else None,
                float(mpv) if mpv else None,
                float(plcr) if plcr else None,
                float(spontaneous_aggregation) if spontaneous_aggregation else None,
                float(induced_aggregation_1_ADP) if induced_aggregation_1_ADP else None,
                float(induced_aggregation_5_ADP) if induced_aggregation_5_ADP else None,
                float(induced_aggregation_15_ARA) if induced_aggregation_15_ARA else None,
            )
            prognosis_result = prognosis_text(prognosis_value)
        except Exception as e:
            prognosis_value = ""
            prognosis_result = f"Ошибка: {e}"

        data_row = [
            gender, age, weight, height, creatinine, creatinine_clearance, mpv, plcr,
            spontaneous_aggregation, induced_aggregation_1_ADP, induced_aggregation_5_ADP,
            induced_aggregation_15_ARA, cyp, abcb1, drugs_str, aggregation_state,
            abcb1_result, mod1_text, mod2_text, mod3_text, prognosis_value, prognosis_result
        ]
        append_patient_data(self.excel_filename, data_row)
        QMessageBox.information(self, "Сохранение", f"Данные успешно сохранены в файл:\n{self.excel_filename}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
