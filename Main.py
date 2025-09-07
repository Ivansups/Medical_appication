from PySide6.QtWidgets import QApplication
import sys
import os

# Добавляем текущую директорию в путь для поиска модулей
sys.path.append(os.path.join(os.path.dirname(__file__), 'classes'))

from classes.MainWindow import MainWindow

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
