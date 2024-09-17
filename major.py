from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from transliterate import translit
import re
import pandas as pd
import os

from enum import Enum


class Course(Enum):
    EXCEL = "Excel"
    PowerBI = "PowerBI"
    PYTHON = "Python"
    SQL = "SQL"
    SQL_ = "SQL_"
    PBIWaitList = "PBIWaitList"
    SQLWaitList = "SQLWaitList"
    ExcelWaitList = "ExcelWaitList"
    PythonWaitList = "PythonWaitList"
    PowerBIFinance = "PowerBIFinance"
    PBI = "PBI"
    AIcourse = "AIcourse"
    Test = "Test"


class Matching:

    def __init__(self, path1: str, path2: str, course: Course):
        self.path1 = path1
        self.path2 = path2
        self.course = course

    def translate_names(self, name):
        if isinstance(name, str):
            transliterated = translit(name, 'ru', reversed=True)

            corrections = {
                'Ajslu': 'Aisulu',
                'Ajsl': 'Aisl',
                'Julija': 'Yuliya',
                'Merej': 'Merey',
                'Aleksandr': 'Alexander',
                'Vladimir': 'Vladimir',
                'Olga': 'Olya',
                'Ekaterina': 'Yekaterina',
                'Mikhail': 'Mikhail',
                'Anastasia': 'Anastasiya',
                'Natalija': 'Natalia',
                'Sergej': 'Sergey',
                'Dmitrij': 'Dmitry',
                'Tatjana': 'Tatiana',
                'Stanislav': 'Stanislav',
                'Vladislav': 'Vladislav',
                'Andrej': 'Andrey',
                'Natalia': 'Natalia',
                'Yurij': 'Yuri',
                'Elena': 'Elena',
                'Daria': 'Daria',
                'Boris': 'Boris',
                'Viktor': 'Viktor',
                'Oleg': 'Oleg',
                'Svetlana': 'Svetlana',
                'Larisa': 'Larisa',
                'Pavel': 'Pavel',
                'Yelena': 'Elena',
                'Gennady': 'Gennady',
                'Nikolai': 'Nikolai',
                'Margarita': 'Margarita',
                'Yulia': 'Yulia',
                'Irina': 'Irina',
                'Oksana': 'Oksana',
                'Yekaterina': 'Ekaterina',
                'Maksim': 'Maxim',
                'Aj': 'Ay',
                'Mej': 'Mei',
                'Jasmin': 'Yasmin',
                'Artem': 'Artyom',
                "Lapyt'ko": 'Lapytko',
                'Kөsherbaj': 'Kosherbay',
                'Ljazzat': 'Lyazzat',
                'Lja': 'Lya',
                "Ol'ga": 'Olga',
                "l'eva": 'lyeva',
                'baj': 'bay',
                'Құlmұhammed': 'Kulmukhammed',
                'lja': 'lya',
                'Baj': 'Bai'

            }
            for incorrect, correct in corrections.items():
                transliterated = transliterated.replace(incorrect, correct)
            return transliterated
        return ''

    def check_language(self, name):
        if isinstance(name, str):
            russian_pattern = re.compile(r'[А-Яа-яЁё]')
            # english_pattern = re.compile(r'[A-Za-z]')
            if russian_pattern.search(name):
                return self.translate_names(name=name)
            else:
                return name
        return ''

    def del_lats_valie(self, x):
        if type(x) == str:
            split_value = x.split(' ')
            if len(split_value) == 3:
                new_value = ' '.join(split_value[:-1])
                return new_value
            return x

    def main(self, save=True):
        df_1 = pd.read_excel(self.path1, sheet_name='сделки')
        df_2 = pd.read_excel(self.path2, sheet_name='Карта уроков по ученикам')
        df_1 = df_1[df_1['курс'] == self.course.value]
        df_1['KEY'] = None
        df_2['KEY'] = None

        unique_key = 1

        for idx_1, name_1 in df_1['ФИО'].apply(self.del_lats_valie).items():
            for idx_2, name_2 in df_2['Имя'].items():
                result = fuzz.token_set_ratio(self.check_language(name_1), self.check_language(name_2))
                if result >= 90:
                    df_1.at[idx_1, 'KEY'] = unique_key
                    df_2.at[idx_2, 'KEY'] = unique_key
                    unique_key += 1
                    print(f'{result}  {self.translate_names(name_1)} ===> {self.translate_names(name_2)}')
        if save:
            new_folder_path = 'Export'
            if not os.path.exists(new_folder_path):
                os.makedirs(new_folder_path)

            df_2.to_excel('Export/One.xlsx')
            df_1.to_excel('Export/Two.xlsx')