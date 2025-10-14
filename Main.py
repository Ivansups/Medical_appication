import os
import sys

if getattr(sys, "frozen", False):
    application_path = sys._MEIPASS
    sys.path.append(os.path.join(application_path, "classes"))
else:
    sys.path.append(os.path.join(os.path.dirname(__file__), "classes"))

from PySide6.QtWidgets import QApplication  # noqa: E402

from classes.MainWindow import MainWindow  # noqa: E402

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
