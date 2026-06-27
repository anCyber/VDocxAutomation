from dataclasses import dataclass

from pytrovich.enums import NamePart, Case
from pytrovich.maker import PetrovichDeclinationMaker
from pytrovich.detector import PetrovichGenderDetector
maker = PetrovichDeclinationMaker()
detector = PetrovichGenderDetector()

import pymorphy2
morph = pymorphy2.MorphAnalyzer()

from pyphrasy.inflect import PhraseInflector
inflector = PhraseInflector(morph)

if (__name__ == "__main__"):
    from read_excel import get_list_of_info_about_employees
else:
    from .read_excel import get_list_of_info_about_employees



@dataclass
class Employee:
    workplace_unit: str     # needs to be put in Gc                                                 
    position: str           # needs to be put in Dc                                                 
    full_name: str          # needs to be splitted into parts  +  get initials  +  get Dc of surname
    days_amount: int        # does not require formatting                                           
    scheduled_date: str     # already formatted                                                     

    def __post_init__(self):
        ''' 
            Sets as class fields variables that require being evaluated.
            Fields that are set here:
            surname, first_name, patronymic, first_name_initial, patronymic_initial,
            surname_dative_case, position_dative_case, workplace_unit_genitive_case,
            days_amount_word_form_day, days_amount_word_form_calendar
        '''
        # [surname, first_name, patronymic]
        full_name_list = self.full_name.split(" ")
        self.surname: str = full_name_list[0]
        self.first_name: str = full_name_list[1]
        self.patronymic: str = full_name_list[2]
        
        # [first_name_initial, patronymic_initial]
        self.first_name_initial: str = self.first_name[0]
        self.patronymic_initial: str = self.patronymic[0]

        # [surname_dative_case]
        gender = detector.detect(firstname=self.first_name, lastname=self.surname, middlename=self.patronymic)
        self.surname_dative_case: str = maker.make(NamePart.LASTNAME, gender, Case.DATIVE, self.surname)       

        # [position_dative_case]
        self.position_dative_case: str = inflector.inflect(self.position, "datv").capitalize()

        # [workplace_unit_genitive_case]
        self.workplace_unit_genitive_case: str = inflector.inflect(self.workplace_unit, "gent").capitalize()

        # [days_amount_word_form_calendar, days_amount_word_form_day]
        word_form_object = morph.parse("каледарный")[0]
        self.days_amount_word_form_calendar = word_form_object.make_agree_with_number(self.days_amount).word
        word_form_object = morph.parse("день")[0]
        self.days_amount_word_form_day = word_form_object.make_agree_with_number(self.days_amount).word


    def __str__(self):
        ''' str(class_instance) → clear display of info about instance (Employee) '''
        return  f"\n\tРаботник ИППИ\n\t- Фамилия, имя отчество: {self.surname} {self.first_name} {self.patronymic}\n\t- Структурное подразделение: {self.workplace_unit}\n\t" \
                f"- Должность: {self.position}\n\t- Количество календарных дней отпуска: {self.days_amount}\n\t- Запланированная дата отпуска: {self.scheduled_date}"



def create_class_instances(PATH_TO_DATA_FILE: str, read_rows_amount: int, first_cell_row: int, first_cell_column_letter: str) -> list[Employee]:
    ''' Creates insatances of Employee class for every list of info read from excel. Also creates a list of all class instances '''
    every_employee_info_list = get_list_of_info_about_employees(PATH_TO_DATA_FILE, read_rows_amount, first_cell_row, first_cell_column_letter)
    employees_list: list = []
    for employee_info in every_employee_info_list:
        new_employee = Employee(
            workplace_unit=employee_info[0],
            position=employee_info[1],
            full_name=employee_info[2],
            days_amount=employee_info[3],
            scheduled_date=employee_info[4]
        )
        employees_list.append(new_employee)

    return employees_list


if (__name__ == "__main__"):
    create_class_instances(r"excel_xlsx\data_text.xlsx", 10, 5, "A")