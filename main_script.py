from excel_xlsx.employee_class import create_class_instances
from word_docx.create_docx import create_word_document

def run_program(
    should_create_single_file: bool, EMPLOYEES_AMOUNT: int,
    PATH_TO_OUTPUT: str,
    FIRST_CELL_COLUMN_LETTER: str, FIRST_CELL_ROW: int, should_create_copy_of_excel: bool,
    PATH_TO_DATA_FILE: str
) -> bool:
    list_of_employee_instances: list = create_class_instances(
        PATH_TO_DATA_FILE = PATH_TO_DATA_FILE,
        read_rows_amount = EMPLOYEES_AMOUNT,
        first_cell_row = FIRST_CELL_ROW,
        first_cell_column_letter = FIRST_CELL_COLUMN_LETTER,
        should_create_copy_of_excel = should_create_copy_of_excel
    )
    create_word_document(
        should_create_single_file = should_create_single_file,
        PATH_TO_OUTPUT = PATH_TO_OUTPUT,
        EMPLOYEES_AMOUNT = len(list_of_employee_instances),
        list_of_employee_instances = list_of_employee_instances
    )
    return True