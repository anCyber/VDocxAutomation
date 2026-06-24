from read_excel import get_dataframe

def get_every_person_info_dict(PATH_TO_DATA_FILE: str, READ_ROWS: int) -> dict:
    '''
        Returns a dictionary containing info about every person. Dict Format: dict[f"person_{index + 1}"]: corresponding_info_dictionary.
        corresponding_info_dictionary Keys: "workplace unit", "position", "full name", "days amount", "scheduled date".
    '''
    DataFrame = get_dataframe(PATH_TO_DATA_FILE, READ_ROWS)

    every_person_data_list: list[list] = DataFrame.to_numpy().tolist()
    every_person_dict: dict = {}
    for person_info_list in every_person_data_list:
        current_person_info_dict = {
            "workplace unit": person_info_list[0],
            "position": person_info_list[1],
            "full name": person_info_list[2],
            "days amount": person_info_list[3],
            "scheduled date": person_info_list[4],
        }

        index = every_person_data_list.index(person_info_list)
        every_person_dict[f"person_{index + 1}"] = current_person_info_dict

    print("{")
    for (key, value) in every_person_dict.items():
        print(f"\tevery_person_dict[{key}]: {value},")
    print("}")

    return every_person_dict



if (__name__ == "__main__"):
    get_every_person_info_dict(r"xlsx\data_text.xlsx", 10)