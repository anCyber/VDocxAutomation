from extract_data import get_every_person_info_dict
from dataclasses import dataclass

PATH_TO_DATA_FILE = r"xlsx\data_text.xlsx"

@dataclass
class Employee:
    workplace_unit: str     # needs to be put in Rp
    position: str           # needs to be put in Dp
    full_name: str          # needs to be splitted into parts  +  get initials  +  get Dp of surname
    days_amount: int        # does not require formatting
    scheduled_date: str     # already formatted

    def __post_init__(self):
        ''' 
            Sets as class fields variables that require being evaluated.
            Fields that are set here: surname, first_name, patronymic, first_name_initial, patronymic_initial
        '''
        # [surname, first_name, patronymic]
        full_name_list = self.full_name.split(" ")
        self.surname: str = full_name_list[0]
        self.first_name: str = full_name_list[1]
        self.patronymic: str = full_name_list[2]
        
        # [first_name_initial, patronymic_initial]
        self.first_name_initial: str = self.first_name[0]
        self.patronymic_initial: str = self.patronymic[0]









def main():
    every_person_info_dict = get_every_person_info_dict(PATH_TO_DATA_FILE, 10)
    print("\n\n\n\n")
    print(every_person_info_dict)


if (__name__ == "__main__"):
    main()