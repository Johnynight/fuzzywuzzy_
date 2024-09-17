from major import Matching, Course

"""
Модуль выводит индекс Левенштейна, если > 90 то присваивает одинаковый индекс в колонку key


Класс Matching принимает следующие аргументы:
    path1 = НФ Учет сделок 2024 Databoom
    path2 = xlsx выгрузка с платформы 
    course = Фильтр в таблице path1 по виду курса. Выбирать из класса Course что бы не возникли ошибки


Метод main по дефолту ставить save=True (Сохраняет 2 файла с новыми колонками KEY) 
в папку Export с названием One.xlsx Two.xlsx


(Нужен для отладки)
"""

a = Matching(path1='Files/Copy of НФ Учет сделок 2024 Databoom.xlsx',
             path2='Files/PBI82.xlsx',
             course=Course.PowerBI)

a.main(save=True)
