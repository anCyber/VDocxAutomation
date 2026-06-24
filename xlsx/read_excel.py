import pandas as pd

COLUMN_A_COORDINATE = "A"
COLUMN_A_COORDINATE_ENCODED = ord(COLUMN_A_COORDINATE)  # chr() to decode back to symbol


def convert_F_datetime_dataframe_to_list(datetime_column: pd.DataFrame) -> list[str]:
        ''' Converts the dataframe containing datetime object into list of formatted strings. Strings Format: "%d.%m.%Y" '''
        list_of_dates = []
        datetime_column_list = datetime_column.to_numpy().tolist()

        for datetime_object in datetime_column_list:
            datetime_object_string = str(datetime_object[0])
            datetime_object_string = datetime_object_string[:10]
            separated_YMD_date_string = datetime_object_string.split("-")
            formated_DMY_string = f"{separated_YMD_date_string[2]}.{separated_YMD_date_string[1]}.{separated_YMD_date_string[0]}"
            list_of_dates.append(formated_DMY_string)

        return list_of_dates

def get_dataframe(PATH_TO_DATA_FILE: str, read_rows_amount: int) -> pd.DataFrame:
    '''
        Reads the Excel-file and returns extracted Dataframe. Contains following columns:
        "Структурное подразделение", "Должность (специальность, профессия) по штатному расписанию", "Фамилия, имя отчество", "количество календарных дней", "запланированная"
    '''
    dataframe_ABC_columns = pd.read_excel(PATH_TO_DATA_FILE, usecols=f"{COLUMN_A_COORDINATE}:{chr(COLUMN_A_COORDINATE_ENCODED + 2)}", skiprows=[1, 2, 3], nrows=read_rows_amount)
    dataframe_E_column = pd.read_excel(PATH_TO_DATA_FILE, usecols=f"{chr(COLUMN_A_COORDINATE_ENCODED + 4)}", skiprows=[0, 2, 3], nrows=read_rows_amount)
    dataframe_F_column = pd.read_excel(PATH_TO_DATA_FILE, usecols=f"{chr(COLUMN_A_COORDINATE_ENCODED + 5)}", skiprows=[0, 1, 3], nrows=read_rows_amount)

    list_of_dates = convert_F_datetime_dataframe_to_list(dataframe_F_column)
    # Source - https://stackoverflow.com/a/42049158
    # Posted by jezrael, modified by community. See post 'Timeline' for change history
    # Retrieved 2026-06-24, License - CC BY-SA 3.0
    dataframe_from_list_of_dates = pd.DataFrame({"запланированная": list_of_dates})

    DataFrame = pd.concat([dataframe_ABC_columns, dataframe_E_column, dataframe_from_list_of_dates], axis=1)
    print(f"\n{'-'*100}\nDataFrame From Excel:\n{DataFrame}\n")
    return DataFrame



if (__name__ == "__main__"):
    get_dataframe(r"xlsx\data_text.xlsx", 10)