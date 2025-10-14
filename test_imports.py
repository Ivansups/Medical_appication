#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys

print("Python version:", sys.version)
print("Current directory:", os.getcwd())

# Добавляем текущую директорию в путь для поиска модулей
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
print("Python path:", sys.path[:3])

try:
    print("Testing PySide6 import...")
    from PySide6.QtWidgets import QApplication

    print("✓ PySide6 imported successfully")

    print("Testing classes import...")
    from classes.MainWindow import MainWindow

    print("✓ MainWindow imported successfully")

    print("All imports successful! Starting application...")

    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    print("Application started successfully!")
    print("Close the window to exit.")
    sys.exit(app.exec())

except Exception as e:
    print(f"✗ Error: {e}")
    import traceback

    traceback.print_exc()
