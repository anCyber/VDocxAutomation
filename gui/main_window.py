from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QApplication, QWidget, QLabel

import sys

def display_window(text):
    app = QApplication(sys.argv)
    window = QWidget()

    temp_label = QLabel()
    temp_label.setText(f"{text}")
    temp_label.setParent(window)
    
    window.show()
    app.exec()




if (__name__ == "__main__"):
    display_window()