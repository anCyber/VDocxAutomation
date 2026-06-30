from PyQt6 import QtWidgets
from PyQt6.QtWidgets import (
    QApplication, QMainWindow,
    QWidget, QFrame, QFileDialog,
    QGridLayout
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIcon


from gui.ui_elements.widget_classes import (
    Q_HeadingLabel, Q_DefaultTextLabel, Q_LineEdit, Q_RadioButton,
    Q_RadioButtonLabel, Q_ExploreButton, Q_CheckBox, Q_RunButton
)

from excel_xlsx.popup_windows import (
    ___AttemptedToReadNegOrZeroRowsOrIncludesSymbols_PopUpWindow___,
    ___InsertedRowOrColumnSymbolIsOfUnexpectedTypeOrOutOfRange_PopUpWindow___,
    ___InputDataFileWithSuchDirectoryDoesNotExist_PopUpWindow___,
    ___OutputFileWithSuchDirectoryDoesNotExist_PopUpWindow___,
    ___InsertedTopLeftCoordinatIsIncorrect_PopUpWindow___,
    ___RunWasSuccessfull_PopUpWindow___
)
from main_script import run_program

import sys
import os
from os.path import exists, isdir


class Window(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setGeometry(0, 0, 1200, 800)
        self.setWindowTitle("VDocxAutomation")
        self.setWindowIcon(QIcon(r"gui\WindowIcon.png"))

        self._init_UI_()
    

    def _init_UI_(self):
        ''' All the widgets that are placed on the window are here '''
        self.background_widget = self.create_main_background_widget("#1E242D", "#181B20")
        self.grid_layout = QGridLayout(self.background_widget)
        self.configure_grid_layout()
        self.selected_output_file = None

        self.sidepanel = self.create_sidepanel()
        
        # Word-section
        self.word_heading_label = Q_HeadingLabel("Настройка Word-файла", additional_qss="")
        self.grid_layout.addWidget(self.word_heading_label, 0, 0, 1, 17, alignment=Qt.AlignmentFlag.AlignHCenter)
        self.create_output_format_section()
        self.create_employees_amount_section()
        self.create_output_path_section()

        self.separator = self.create_separator()

        # Excel-section
        self.excel_heading_label = Q_HeadingLabel("Настройка Excel-файла", additional_qss="")
        self.grid_layout.addWidget(self.excel_heading_label, 10, 0, 1, 17, alignment=Qt.AlignmentFlag.AlignHCenter)
        self.create_top_left_coordinate_section()
        self.create_copy_excel_section()
        self.create_input_data_path_section()

        self.run_button = Q_RunButton("Запустить программу", additional_qss="")
        self.grid_layout.addWidget(self.run_button, 19, 8, 1, 4)
        self.run_button.clicked.connect(self._run_button_clicked_)



    def create_output_format_section(self) -> None:
        ''' Creates output format section in Word-section '''
        self.groupbox_radio_buttons = QFrame()
        self.grid_layout.addWidget(self.groupbox_radio_buttons, 2, 4, 3, 6)
        self.radio_buttons_group_layout = QGridLayout()
        self.radio_buttons_group_layout.setRowStretch(0, 24)
        self.radio_buttons_group_layout.setRowMinimumHeight(1, 12)
        self.radio_buttons_group_layout.setRowStretch(1, 0)
        self.radio_buttons_group_layout.setRowStretch(2, 24)
        self.radio_buttons_group_layout.setSpacing(0)
        self.radio_buttons_group_layout.setContentsMargins(0,0,0,0)
        self.groupbox_radio_buttons.setLayout(self.radio_buttons_group_layout)
        
        self.output_format_label = Q_DefaultTextLabel("Формат output’а:", additional_qss="margin-bottom: 12px;")
        self.grid_layout.addWidget(self.output_format_label, 1, 3, 1, 3, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter))

        self.output_format_radio_button_1 = Q_RadioButton(additional_qss="")
        self.radio_buttons_group_layout.addWidget(self.output_format_radio_button_1, 0, 0, 1, 1)
        self.output_format_radio_button_1.setChecked(True)

        self.output_format_label_1 = Q_RadioButtonLabel("Единый файл на N страниц", additional_qss="")
        self.radio_buttons_group_layout.addWidget(self.output_format_label_1, 0, 1, 1, 1, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter))

        self.output_format_radio_button_2 = Q_RadioButton(additional_qss="")
        self.radio_buttons_group_layout.addWidget(self.output_format_radio_button_2, 2, 0, 1, 1)

        self.output_format_label_2 = Q_RadioButtonLabel("N одностраничных файлов", additional_qss="")
        self.radio_buttons_group_layout.addWidget(self.output_format_label_2, 2, 1, 1, 1, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter))

    
    def create_employees_amount_section(self) -> None:
        ''' Creates employees amount section in Word-section '''
        self.employees_amount_label = Q_DefaultTextLabel("Кол-во страниц:", additional_qss="")
        self.grid_layout.addWidget(self.employees_amount_label, 1, 11, 1, 2, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter))

        self.employees_amount_line_edit = Q_LineEdit(additional_qss="margin-left: 24px;")
        self.grid_layout.addWidget(self.employees_amount_line_edit, 1, 13, 1, 2)


    def create_output_path_section(self) -> None:
        ''' Creates output path section in Word-section '''
        self.output_path_label = Q_DefaultTextLabel("Путь к output-файлу:", additional_qss="margin-top: 28px; margin-left: 4px; margin-bottom: 8px;")
        self.grid_layout.addWidget(self.output_path_label, 5, 7, 1, 4, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop))

        self.output_path_button = Q_ExploreButton("Выбрать output-файл", additional_qss="")
        self.grid_layout.addWidget(self.output_path_button, 6, 2, 1, 4)
        self.output_path_button.clicked.connect(self._explore_output_button_clicked_)

        self.output_path_line_edit = Q_LineEdit(additional_qss="height: 36px;")
        self.grid_layout.addWidget(self.output_path_line_edit, 6, 7, 1, 9)
        
        
    def create_top_left_coordinate_section(self) -> None:
        ''' Creates top left coordinate section in Excel-section '''
        self.top_left_coordinate_label = Q_DefaultTextLabel("Координаты левой-верхней\nячейки таблицы с данными:", additional_qss="")
        self.grid_layout.addWidget(self.top_left_coordinate_label, 11, 3, 5, 5, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter))

        self.top_left_coordinate_line_edit = Q_LineEdit(additional_qss="margin-left: 24px;")
        self.grid_layout.addWidget(self.top_left_coordinate_line_edit, 12, 9, 3, 1,alignment=Qt.AlignmentFlag.AlignVCenter)


    def create_copy_excel_section(self) -> None:
        ''' Creates copy excel section in Excel-section '''
        self.copy_excel_label = Q_DefaultTextLabel("Создать копию excel-файла", additional_qss="")
        self.grid_layout.addWidget(self.copy_excel_label, 13, 11, 1, 3, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter))

        self.copy_excel_checkbox = Q_CheckBox(additional_qss="")
        self.grid_layout.addWidget(self.copy_excel_checkbox, 13, 14, 1, 1)


    
    def create_input_data_path_section(self) -> None:
        ''' Creates input data path section in Excel-section '''
        self.input_data_path_label = Q_DefaultTextLabel("Путь к файлу с данными:", additional_qss="margin-top: 28px; margin-left: 4px; margin-bottom: 8px;")
        self.grid_layout.addWidget(self.input_data_path_label, 16, 7, 1, 4, alignment=(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop))

        self.input_data_path_button = Q_ExploreButton("Выбрать файл с данными", additional_qss="")
        self.grid_layout.addWidget(self.input_data_path_button, 17, 2, 1, 4)
        self.input_data_path_button.clicked.connect(self._explore_input_data_button_clicked_)

        self.input_data_path_line_edit = Q_LineEdit(additional_qss="height: 36px;")
        self.grid_layout.addWidget(self.input_data_path_line_edit, 17, 7, 1, 9)



    def create_main_background_widget(self, color1: str, color2: str) -> QWidget:
        ''' Creates widget that automatically takes up all the space in window and is used as background. Copied function from "AudioRedactor" project :) '''
        background_central_widget = QWidget()
        self.setCentralWidget(background_central_widget)
        bg_widget_layout = QGridLayout(background_central_widget)
        bg_widget_layout.setRowStretch(0, 1)
        bg_widget_layout.setColumnStretch(0, 1)

        bg_widget = QWidget()
        bg_widget.setProperty("class", "bg_widget")
        bg_widget.setStyleSheet (
        f"""
            .bg_widget {{
                background: qlineargradient(x1:0.5, y1:0, x2:0.5, y2:1, stop:0 {color1}, stop:1 {color2});
            }}
        """
        )
        bg_widget_layout.setSpacing(0)
        bg_widget_layout.setContentsMargins(0,0,0,0)
        bg_widget_layout.addWidget(bg_widget, 0, 0, 1, 1)
        return bg_widget



    def configure_grid_layout(self) -> None:
        ''' Configures the rows and columns of the grid layout '''
        self.grid_layout.setContentsMargins(0,0,0,0)
        self.grid_layout.setSpacing(0)
        rows_heights = [112, 32, 24, 12, 24, 56, 44, 48, 8, 28, 84, 4, 4, 24, 4, 4, 56, 44, 60, 64, 64]
        columns_widths = [60, 40, 100, 20, 24, 100, 90, 28, 6, 112, 148, 12, 148, 88, 24, 100, 100]

        for i in range(len(rows_heights)):
            self.grid_layout.setRowStretch(i, rows_heights[i])
            self.grid_layout.setRowMinimumHeight(i, rows_heights[i])
        
        for i in range(len(columns_widths)):
            self.grid_layout.setColumnStretch(i, columns_widths[i])
            self.grid_layout.setColumnMinimumWidth(i, columns_widths[i])

        self.grid_layout.setColumnStretch(5, 0)
        self.grid_layout.setColumnStretch(6, 0)
        self.grid_layout.setColumnStretch(7, 0)
        self.grid_layout.setColumnStretch(8, 0)
        self.grid_layout.setColumnStretch(9, 0)
        self.grid_layout.setColumnStretch(11, 0)
        self.grid_layout.setColumnStretch(12, 0)
        self.grid_layout.setColumnStretch(13, 0)

        
    def create_sidepanel(self) -> QWidget:
        ''' Creates the sidepanel widget (purely for aesthetics) '''
        sidepanel = QWidget()
        sidepanel.setProperty("class", "sidepanel")
        sidepanel.setStyleSheet (
        f"""
            .sidepanel {{
                background-color: rgba(217, 217, 217, 0.04);
                border-right-style: solid;
                border-right-width: 1px;
                border-right-color: rgba(173, 215, 235, 0.10);
            }}
        """
        )
        self.grid_layout.addWidget(sidepanel, 0, 0, 21, 1)
        return sidepanel
    

    def create_separator(self) -> QWidget:
        ''' Creates the separator widget (purely for aesthetics) '''
        separator = QWidget()
        separator.setProperty("class", "separator")
        separator.setStyleSheet (
        f"""
            .separator {{
                background-color: rgba(0, 0, 0, 0);
                border-top-style: solid;
                border-top-width: 1px;
                border-top-color: rgba(37, 42, 51, 1);
                border-bottom-style: solid;
                border-bottom-width: 1px;
                border-bottom-color: rgba(37, 42, 51, 1);
            }}
        """
        )
        self.grid_layout.addWidget(separator, 8, 1, 1, 16)
        return separator




    def _explore_output_button_clicked_(self) -> None:
        ''' Event listener for {self.output_path_button}'''
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("*.docx")
        was_file_selected: bool = file_dialog.exec()
        if (was_file_selected):
            self.selected_output_file = file_dialog.selectedFiles()[0]
            self.output_path_line_edit.setText(self.selected_output_file)

    
    def _explore_input_data_button_clicked_(self) -> None:
        ''' Event listener for {self.input_data_path_button}'''
        file_dialog = QFileDialog()
        file_dialog.setNameFilter("*.xlsx")
        was_file_selected: bool = file_dialog.exec()
        if (was_file_selected):
            self.selected_input_data_file = file_dialog.selectedFiles()[0]
            self.input_data_path_line_edit.setText(self.selected_input_data_file)


    def validate_data(self) -> list:
        ''' Returns the list of necessary data for the program to work '''
        if (self.output_format_radio_button_1.isChecked()):
            should_create_single_file: bool = True
        else:
            should_create_single_file: bool = False

        try:
            EMPLOYEES_AMOUNT: int = int(self.employees_amount_line_edit.text())
            if (EMPLOYEES_AMOUNT < 1):
                raise ValueError
        except ValueError:
            ___AttemptedToReadNegOrZeroRowsOrIncludesSymbols_PopUpWindow___()
            self.employees_amount_line_edit.setText("")
            self.employees_amount_line_edit.setFocus()
            return []

        PATH_TO_OUTPUT: str = fr"{self.output_path_line_edit.text()}"
        if (PATH_TO_OUTPUT == ""):
            ___OutputFileWithSuchDirectoryDoesNotExist_PopUpWindow___("Empty")
            self.output_path_line_edit.setText("")
            self.output_path_line_edit.setFocus()
            return []
        output_directory_split: list[str] = PATH_TO_OUTPUT.split(r"/")[:-1:]
        separator: str = r"/"
        output_directory_connected = separator.join(output_directory_split)
        if (not os.path.isdir(output_directory_connected)):
            ___OutputFileWithSuchDirectoryDoesNotExist_PopUpWindow___("")
            self.output_path_line_edit.setText("")
            self.output_path_line_edit.setFocus()
            return []

        try:
            FIRST_CELL_COLUMN_LETTER: str = self.top_left_coordinate_line_edit.text()[0]
        except IndexError:
            ___InsertedRowOrColumnSymbolIsOfUnexpectedTypeOrOutOfRange_PopUpWindow___("Empty")
            self.top_left_coordinate_line_edit.setFocus()
            return []

        if ((ord(FIRST_CELL_COLUMN_LETTER) < 65) or (ord(FIRST_CELL_COLUMN_LETTER) > 84)):
            ___InsertedRowOrColumnSymbolIsOfUnexpectedTypeOrOutOfRange_PopUpWindow___("")
            self.top_left_coordinate_line_edit.setText("")
            self.top_left_coordinate_line_edit.setFocus()
            return []

        try:
            FIRST_CELL_ROW: int = int(self.top_left_coordinate_line_edit.text()[1::])
        except ValueError:
            ___InsertedRowOrColumnSymbolIsOfUnexpectedTypeOrOutOfRange_PopUpWindow___("")
            self.top_left_coordinate_line_edit.setText("")
            self.top_left_coordinate_line_edit.setFocus()
            return []

        should_create_copy_of_excel: bool = self.copy_excel_checkbox.isChecked()

        
        PATH_TO_DATA_FILE: str = fr"{self.input_data_path_line_edit.text()}"
        if (PATH_TO_DATA_FILE == ""):
            ___InputDataFileWithSuchDirectoryDoesNotExist_PopUpWindow___("Empty")
            self.input_data_path_line_edit.setFocus()
            return []
        if (not os.path.exists(PATH_TO_DATA_FILE)):
            ___InputDataFileWithSuchDirectoryDoesNotExist_PopUpWindow___("")
            self.input_data_path_line_edit.setText("")
            self.input_data_path_line_edit.setFocus()
            return []
        return [should_create_single_file, EMPLOYEES_AMOUNT, PATH_TO_OUTPUT, FIRST_CELL_COLUMN_LETTER, FIRST_CELL_ROW, should_create_copy_of_excel, PATH_TO_DATA_FILE]




    def _run_button_clicked_(self) -> None:
        ''' Event listener for {self.run_button}. Initializes the program with the inserted values. '''
        data_list: list = self.validate_data()
        if (data_list != []):
            try:
                was_run_succesfull: bool = run_program(
                    data_list[0], data_list[1], data_list[2], data_list[3], data_list[4], data_list[5], data_list[6]
                )
                ___RunWasSuccessfull_PopUpWindow___(was_run_succesfull)
            except IndexError:
                ___InsertedTopLeftCoordinatIsIncorrect_PopUpWindow___()
                self.top_left_coordinate_line_edit.setText("")
                self.top_left_coordinate_line_edit.setFocus()
                self.top_left_coordinate_line_edit

            






def main():
    app = QApplication([])
    window = Window()
    
    window.show()
    sys.exit(app.exec())




if (__name__ == "__main__"):
    main()