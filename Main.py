import os
import sys

# Добавляем текущую директорию в путь для поиска модулей
sys.path.append(os.path.join(os.path.dirname(__file__), "classes"))

from PySide6.QtWidgets import QApplication  # noqa: E402

from classes.MainWindow import MainWindow  # noqa: E402

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
