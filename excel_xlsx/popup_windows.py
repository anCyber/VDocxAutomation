from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QMessageBox
import sys


class ___AttemptedToReadTooManyRows_PopUpWindow___(QMessageBox):
    def __init__(self, maximum_rows: int):
        super().__init__()
        self.setIcon(QMessageBox.Icon.Critical)
        self.setWindowTitle("Ошибка!")
        if (maximum_rows > 0):
            self.setText("При попытке считать данные из Excel-файла возникла ошибка!\n" \
            "В указанном файле недостаточно строк с данными, чтобы создать желаемое количество страниц.")
            self.continue_button = self.addButton(f"Считать все существующие строки: ({maximum_rows})", QMessageBox.ButtonRole.YesRole)
        else:
            self.setText("При попытке считать данные из Excel-файла возникла ошибка!\n" \
            "Указанный файл не содержит строк с информацией о работниках, либо указанная координата левого-верхнего угла данных в таблице не является точной.")
            self.continue_button = self.addButton(f"Указать новый путь к файлу", QMessageBox.ButtonRole.YesRole)

        self.abort_button = self.addButton("Выйти", QMessageBox.ButtonRole.NoRole)
        self.setDefaultButton(self.abort_button)
        self.buttonClicked.connect(self.option_clicked)
        if (self.exec() == 3):  # sys.exit if pop up was closed with the "x" button
            sys.exit()




    def option_clicked(self):
        if (self.clickedButton() == self.continue_button):
            self.close()
        else:
            sys.exit()




class ___AttemptedToReadNegOrZeroRowsOrIncludesSymbols_PopUpWindow___(QMessageBox):
    def __init__(self):
        super().__init__()
        self.setIcon(QMessageBox.Icon.Critical)
        self.setWindowTitle("Ошибка!")
        self.setText("Указанное количество строк таблицы, подлежащих считыванию, не является положительным числом или содержит символы, не являющиеся цифрами.\n" \
        "Пожалуйста, вводите только числа больше нуля.")
        self.continue_button = self.addButton(f"Указать другое количество строк", QMessageBox.ButtonRole.YesRole)
        self.abort_button = self.addButton("Выйти", QMessageBox.ButtonRole.NoRole)
        self.setDefaultButton(self.abort_button)
        self.buttonClicked.connect(self.option_clicked)
        if (self.exec() == 3):  # sys.exit if pop up was closed with the "x" button
            sys.exit()



    def option_clicked(self):
        if (self.clickedButton() == self.continue_button):
            self.close()
        else:
            sys.exit()



class ___InsertedRowOrColumnSymbolIsOfUnexpectedTypeOrOutOfRange_PopUpWindow___(QMessageBox):
    def __init__(self, text: str):
        super().__init__()
        
        self.setWindowTitle("Ошибка!")
        if (text == "Empty"):
            self.setIcon(QMessageBox.Icon.Warning)
            self.setText("Введите координату левого-верхнего угла данных таблицы.")
            self.continue_button = self.addButton(f"Указать координату", QMessageBox.ButtonRole.YesRole)
        else:
            self.setIcon(QMessageBox.Icon.Critical)
            self.setText("Проблема с введёнными координатами левого-верхнего угла данных таблицы. Вводите данные в формате:\n" \
            "'{Большая Английская Буква в диапазоне [от A до T]}{Цифра/Число}'. Примеры:\n" \
            "A5, D30, I16.\n" \
            "Буквы, идущие после T вызовут ошибку в связи с системой координат Excel: (X -> Y -> Z -> AA -> AB -> AC -> ...)")
            self.continue_button = self.addButton(f"Указать новую координату", QMessageBox.ButtonRole.YesRole)

        self.abort_button = self.addButton("Выйти", QMessageBox.ButtonRole.NoRole)
        self.setDefaultButton(self.abort_button)
        self.buttonClicked.connect(self.option_clicked)
        if (self.exec() == 3):  # sys.exit if pop up was closed with the "x" button
            sys.exit()

    def option_clicked(self):
        if (self.clickedButton() == self.continue_button):
            self.close()
        else:
            sys.exit()



class ___InsertedTopLeftCoordinatIsIncorrect_PopUpWindow___(QMessageBox):
    def __init__(self):
        super().__init__()
        
        self.setWindowTitle("Ошибка!")
        self.setIcon(QMessageBox.Icon.Critical)
        self.setText("Проблема с введёнными координатами левого-верхнего угла данных таблицы. Введённая координата не является левым верхним углом данных.\n" \
        "Координата должна вести на ячейку 'Структурное подразделение' первого работника.")
        self.continue_button = self.addButton(f"Указать новую координату", QMessageBox.ButtonRole.YesRole)
        self.abort_button = self.addButton("Выйти", QMessageBox.ButtonRole.NoRole)
        self.setDefaultButton(self.abort_button)
        self.buttonClicked.connect(self.option_clicked)
        if (self.exec() == 3):  # sys.exit if pop up was closed with the "x" button
            sys.exit()

    def option_clicked(self):
        if (self.clickedButton() == self.continue_button):
            self.close()
        else:
            sys.exit()



class ___InputDataFileWithSuchDirectoryDoesNotExist_PopUpWindow___(QMessageBox):
    def __init__(self, text: str):
        super().__init__()
        
        self.setWindowTitle("Ошибка!")
        if (text == "Empty"):
            self.setIcon(QMessageBox.Icon.Warning)
            self.setText("Пожалуйста, укажите путь к Excel-файлу с информацией о работниках.")
            self.continue_button = self.addButton(f"Указать путь к файлу", QMessageBox.ButtonRole.YesRole)
        else:
            self.setIcon(QMessageBox.Icon.Critical)
            self.setText("Не удалось найти Excel-файл с введённым расположением. Пожалуйста, выберите существующий файл.")
            self.continue_button = self.addButton(f"Указать новый путь к файлу", QMessageBox.ButtonRole.YesRole)

        self.abort_button = self.addButton("Выйти", QMessageBox.ButtonRole.NoRole)
        self.setDefaultButton(self.abort_button)
        self.buttonClicked.connect(self.option_clicked)
        if (self.exec() == 3):  # sys.exit if pop up was closed with the "x" button
            sys.exit()



    def option_clicked(self):
        if (self.clickedButton() == self.continue_button):
            self.close()
        else:
            sys.exit()



class ___OutputFileWithSuchDirectoryDoesNotExist_PopUpWindow___(QMessageBox):
    def __init__(self, text: str):
        super().__init__()
        
        self.setWindowTitle("Ошибка!")
        if (text == "Empty"):
            self.setIcon(QMessageBox.Icon.Warning)
            self.setText("Пожалуйста, укажите путь к создаваемому Word-файлу")
            self.continue_button = self.addButton(f"Указать путь к файлу", QMessageBox.ButtonRole.YesRole)
        else:
            self.setIcon(QMessageBox.Icon.Critical)
            self.setText("Не удалось создать Word-файл в указанное расположение. Пожалуйста, выберите существующую директорию.\n" \
            "Подсказка: \n" \
            "В меню файл-диалога, открывающегося при нажатии на кнопку:\n 'Выбрать output-файл', вы можете, находясь в папке, где вы хотите получить документ, нажать правой кнопкой мыши по свободному пространству и создать пустой Word-файл.\n" \
            "После этого по необходимости переименуйте созданный файл и выберите его. Путь автоматически скопируется в поле ввода.")
            self.continue_button = self.addButton(f"Указать новый путь к файлу", QMessageBox.ButtonRole.YesRole)

        self.abort_button = self.addButton("Выйти", QMessageBox.ButtonRole.NoRole)
        self.setDefaultButton(self.abort_button)
        self.buttonClicked.connect(self.option_clicked)
        if (self.exec() == 3):  # sys.exit if pop up was closed with the "x" button
            sys.exit()


    def option_clicked(self):
        if (self.clickedButton() == self.continue_button):
            self.close()
        else:
            sys.exit()




class ___RunWasSuccessfull_PopUpWindow___(QMessageBox):
    def __init__(self, was_run_successfull: bool):
        super().__init__()
        if (was_run_successfull):
            self.setIcon(QMessageBox.Icon.Information)
            self.setWindowTitle("Успех!")
            self.setText("Файл был успешно создан в указанной директории.")
            self.setStandardButtons(QMessageBox.StandardButton.Ok)
        else:
            self.setIcon(QMessageBox.Icon.Critical)
            self.setText("Произошла непредвиденная ошибка. Word-файл не был создан.")
            self.continue_button = self.addButton(f"Вернуться к окну программы", QMessageBox.ButtonRole.YesRole)
            self.abort_button = self.addButton("Выйти", QMessageBox.ButtonRole.NoRole)
            self.setDefaultButton(self.abort_button)
            self.buttonClicked.connect(self.option_clicked)

        if (self.exec() == 3):  # sys.exit if pop up was closed with the "x" button
            sys.exit()


    def option_clicked(self):
        if ((self.clickedButton() == self.continue_button) or (self.clickedButton==QMessageBox.StandardButton.Ok)):
            self.close()
        else:
            sys.exit()