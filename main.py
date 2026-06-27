from gui.main_window import display_window
from excel_xlsx.employee_class import create_class_instances
from word_docx.create_docx import create_word_document

PATH_TO_MANY_OUTPUTS = r"Output\DocxFile"
PATH_TO_SINGLE_OUTPUT = r"Output\DocxFile.docx"

PATH_TO_DATA_FILE = r"excel_xlsx\data_text_moved.xlsx"
EMPLOYEES_AMOUNT = 8
FIRST_CELL_ROW = 16
FIRST_CELL_COLUMN_LETTER = "I"
SHOULD_CREATE_COPY_OF_EXCEL = True

def main():
    display_window("How's it going mate?")
    list_of_employee_instances: list = create_class_instances(
        PATH_TO_DATA_FILE = PATH_TO_DATA_FILE,
        read_rows_amount = EMPLOYEES_AMOUNT,
        first_cell_row = FIRST_CELL_ROW,
        first_cell_column_letter = FIRST_CELL_COLUMN_LETTER,
        should_create_copy_of_excel = SHOULD_CREATE_COPY_OF_EXCEL
    )
    create_word_document(True, PATH_TO_SINGLE_OUTPUT, EMPLOYEES_AMOUNT, list_of_employee_instances)

if (__name__ == "__main__"):
    main()