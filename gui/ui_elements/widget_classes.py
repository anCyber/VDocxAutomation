from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QLabel, QLineEdit, QRadioButton, QPushButton
from PyQt6.QtGui import QFontDatabase, QFont

class Q_HeadingLabel(QLabel):
    def __init__(self, text: str, additional_qss: str):
        super().__init__()

        self.style_sheet_string = """
            .Q_HeadingLabel {
                color: #FFFFFF;
        """ + additional_qss + "}"
        self.setText(text)
        self.setProperty("class", "Q_HeadingLabel")
        self.setStyleSheet (f"{self.style_sheet_string}")

        # https://ru.stackoverflow.com/a/1548924
        font_path = r"gui\ui_elements\fonts\Brygada1918-Medium.ttf"
        QFontDatabase.addApplicationFont(font_path)   
        family = QFontDatabase.applicationFontFamilies(0)
        font = QFont(family, 40, 500, False)
        self.setFont(font)





class Q_DefaultTextLabel(QLabel):
    def __init__(self, text: str, additional_qss: str):
        super().__init__()
        self.style_sheet_string = """
            .Q_DefaultTextLabel {
                color: #FFFFFF;
        """ + additional_qss + "}"
        self.setText(text)
        self.setProperty("class", "Q_DefaultTextLabel")
        self.setStyleSheet (f"{self.style_sheet_string}")

        # https://ru.stackoverflow.com/a/1548924
        font_path = r"gui\ui_elements\fonts\Brygada1918-Regular.ttf"
        QFontDatabase.addApplicationFont(font_path)   
        family = QFontDatabase.applicationFontFamilies(0)
        font = QFont(family, 18, 400, False)
        self.setFont(font)



class Q_LineEdit(QLineEdit):
    def __init__(self, additional_qss: str):
        super().__init__()

        self.style_sheet_string = """
            .Q_LineEdit {
                background-color: rgba(216, 248, 255, 0.05);
                border-style: solid;
                border-width: 1px;
                border-color: rgba(86, 104, 119, 1);
                border-radius: 6px;

                color: #FFFFFF;
                padding-left: 12px;
                padding-top: 4px;
                padding-bottom: 4px;
            """ + additional_qss + "}" + """
            .Q_LineEdit:focus {
                background-color: rgba(216, 248, 255, 0.15);
                border-color: rgba(173, 215, 235, 1);
            }"""
        
        self.setProperty("class", "Q_LineEdit")
        self.setStyleSheet (f"{self.style_sheet_string}")


        font = QFont("Arial", 12, 400, False)
        self.setFont(font)



class Q_RadioButton(QRadioButton):
    def __init__(self, additional_qss: str):
        super().__init__()

        self.style_sheet_string = """
            .Q_RadioButton {
                background-color: rgba(216, 248, 255, 0.05);
                border-style: solid;
                border-width: 2px;
                border-color: rgba(92, 112, 126, 1);
                border-radius: 12px;
                outline: none;
            """ + additional_qss + "}" + """
            .Q_RadioButton::indicator {
                background-color: rgba(0, 0, 0, 0);
                border-color: rgba(0, 0, 0, 0);
            }
            .Q_RadioButton:checked {
                background-color: rgba(216, 248, 255, 0.15);
                border-color: rgba(173, 215, 235, 1);
            }
            .Q_RadioButton::indicator:checked {
                margin: 3px;
                background-color: rgba(173, 215, 235, 1);
                border-radius: 7px;
                width: 14px;
                height: 14px;
            }
            """
        self.setFixedSize(24, 24)
        self.setProperty("class", "Q_RadioButton")
        self.setStyleSheet (f"{self.style_sheet_string}")



class Q_RadioButtonLabel(QLabel):
    def __init__(self, text: str, additional_qss: str):
        super().__init__()
        self.style_sheet_string = """
            .Q_RadioButtonLabel {
                color: #FFFFFF;
                margin-left: 8px;
        """ + additional_qss + "}"
        self.setText(text)
        self.setProperty("class", "Q_RadioButtonLabel")
        self.setStyleSheet (f"{self.style_sheet_string}")

        # https://ru.stackoverflow.com/a/1548924
        font_path = r"gui\ui_elements\fonts\Brygada1918-Regular.ttf"
        QFontDatabase.addApplicationFont(font_path)   
        family = QFontDatabase.applicationFontFamilies(0)
        font = QFont(family, 16, 400, False)
        self.setFont(font)



class Q_ExploreButton(QPushButton):
    def __init__(self, text: str, additional_qss: str):
        super().__init__()

        self.style_sheet_string = """
            .Q_ExploreButton {
                background-color: rgba(50, 56, 61, 1);
                border-style: solid;
                border-width: 1px;
                border-color: rgba(92, 112, 126, 1);
                border-radius: 8px;

                color: rgba(244, 252, 255, 1);
            """ + additional_qss + "}" + """
            .Q_ExploreButton:hover {
                background-color: rgba(50, 56, 61, 0.70);
                border-color: rgba(72, 95, 102, 1);
                color: rgba(224, 230, 232, 1);
            }
            .Q_ExploreButton:pressed {
                background-color: rgba(50, 56, 61, 0.50);
                border-color: rgba(55, 71, 80, 1);
                color: rgba(205, 209, 212, 1);
            }
            """
        self.setFixedHeight(44)
        self.setText(text)
        self.setProperty("class", "Q_ExploreButton")
        self.setStyleSheet (f"{self.style_sheet_string}")


        # https://ru.stackoverflow.com/a/1548924
        font_path = r"gui\ui_elements\fonts\Brygada1918-Regular.ttf"
        QFontDatabase.addApplicationFont(font_path)   
        family = QFontDatabase.applicationFontFamilies(0)
        font = QFont(family, 16, 400, False)
        self.setFont(font)




class Q_CheckBox(QRadioButton):
    def __init__(self, additional_qss: str):
        super().__init__()

        self.style_sheet_string = """
            .Q_CheckBox {
                background-color: rgba(29, 35, 41, 1);
                border-style: solid;
                border-width: 2px;
                border-color: rgba(92, 112, 126, 1);
                border-radius: 4px;
                margin-left: 14px;
                outline: none;
            """ + additional_qss + "}" + """
            .Q_CheckBox::indicator {
                background-color: rgba(0, 0, 0, 0);
                border-color: rgba(0, 0, 0, 0);
            }
            .Q_CheckBox:checked {
                background-color: rgba(216, 248, 255, 0.15);
                border-color: rgba(173, 215, 235, 1);
            }
            .Q_CheckBox::indicator:checked {
                margin: 3px;
                background-color: rgba(173, 215, 235, 1);
                border-radius: 2px;
                width: 14px;
                height: 14px;
            }
            """
        self.setFixedSize(38, 24)
        self.setProperty("class", "Q_CheckBox")
        self.setStyleSheet (f"{self.style_sheet_string}")




class Q_RunButton(QPushButton):
    def __init__(self, text: str, additional_qss: str):
        super().__init__()

        self.style_sheet_string = """
            .Q_RunButton {
                background-color: rgba(36, 55, 62, 1);
                border-style: solid;
                border-width: 2px;
                border-color: rgba(144, 173, 188, 1);
                border-radius: 8px;

                color: rgba(251, 254, 255, 1);
            """ + additional_qss + "}" + """
            .Q_RunButton:hover {
                background-color: rgba(30, 48, 54, 1);
                border-color: rgba(124, 150, 161, 1);
                color: rgba(228, 230, 232, 1);
            }
            .Q_RunButton:pressed {
                background-color: rgba(24, 40, 48, 1);
                border-color: rgba(108, 130, 140, 1);
                color: rgba(205, 209, 212, 1);
            }
            """
        self.setFixedHeight(60)
        self.setText(text)
        self.setProperty("class", "Q_RunButton")
        self.setStyleSheet (f"{self.style_sheet_string}")


        # https://ru.stackoverflow.com/a/1548924
        font_path = r"gui\ui_elements\fonts\Brygada1918-SemiBold.ttf"
        QFontDatabase.addApplicationFont(font_path)   
        family = QFontDatabase.applicationFontFamilies(0)
        font = QFont(family, 16, 600, False)
        self.setFont(font)