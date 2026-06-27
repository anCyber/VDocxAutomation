import openpyxl

CONVERT_LETTERS_TO_NUMBERS = 64


def convert_datetime_object_to_str(datetime_object) -> str:
    ''' Converts the datetime object into formatted string. Strings Format: "%d.%m.%Y" '''
    datetime_object_string = str(datetime_object)
    datetime_object_string = datetime_object_string[:10]
    separated_YMD_date_string = datetime_object_string.split("-")
    formated_DMY_string = f"{separated_YMD_date_string[2]}.{separated_YMD_date_string[1]}.{separated_YMD_date_string[0]}"
    return formated_DMY_string


def iterate_through_table_rows(worksheet: openpyxl.worksheet, read_rows_amount: int, first_cell_row: int, first_cell_column_letter: str):
    ''' Iterates through table rows and returns the list of rows, containing desired information '''
    first_cell_column = ord(f"{first_cell_column_letter}") - CONVERT_LETTERS_TO_NUMBERS
    READ_COLUMNS_AMOUNT = 5 + first_cell_column

    list_of_rows: list = []
    for table_row in worksheet.iter_rows(first_cell_row, (read_rows_amount + first_cell_row - 1), first_cell_column, READ_COLUMNS_AMOUNT):
        table_row_list = []
        for cell in table_row:
            table_row_list.append(cell.value)
        
        table_row_list[-1] = convert_datetime_object_to_str(table_row_list[-1])
        table_row_list.pop(3)  # Removing undesired column from list (tabel_number)
        list_of_rows.append(table_row_list)
    return list_of_rows



def get_list_of_info_about_employees(PATH_TO_DATA_FILE: str, read_rows_amount: int, first_cell_row: int, first_cell_column_letter: str):
    '''
        Reads the Excel-file and returns extracted Dataframe. Contains following columns:
        "Структурное подразделение", "Должность (специальность, профессия) по штатному расписанию", "Фамилия, имя отчество", "количество календарных дней", "запланированная"
    '''
    excel_file = openpyxl.load_workbook(PATH_TO_DATA_FILE)
    first_sheet = excel_file.worksheets[0]

    list_of_table_rows = iterate_through_table_rows(first_sheet, read_rows_amount, first_cell_row, first_cell_column_letter)
    return list_of_table_rows



if (__name__ == "__main__"):
    get_list_of_info_about_employees(r"excel_xlsx\data_text.xlsx", 10, 5, "A")