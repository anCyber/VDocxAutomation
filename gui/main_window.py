from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QApplication, QWidget, QLabel

import sys

def display_window(text):
    app = QApplication(sys.argv)
    window = QWidget()

    temp_label = QLabel(window)
    temp_label.setText(f"{text}")
    
    window.show()
    app.exec()




if (__name__ == "__main__"):
    display_window("yoo")