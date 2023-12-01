import pandas as pd
from modules.Table import Table


class AppendixResult(Table):
    @staticmethod
    def job(path_to_load: str, path_to_save: str) -> None:
        Table.concat(path_to_load, path_to_save, 'Запросы')

    @staticmethod
    def create_table(path_to_create: str) -> None:
        # file_object = pd.DataFrame(
        #     {'Вид работ': [], 'Ед. изм': [], 'Месяц': [], 'Год:': [], 'План/Факт': [], 'Значение': [],
        #      'Проект (Паспорт ГРП)': [], 'Участок': [], 'Стадия': [], 'Проект-Участок': [], 'План на год': []})
        file_object = pd.DataFrame({'Вид работ': [], 'Ед. изм': [], 'Месяц': [], 'Год:': [], 'План/Факт': [], 'Значение': [],
                                    'Проект (Паспорт ГРП)': [], 'Участок': [],'Лицензия': [], 'Стадия': [], 'Проект-Участок': [],'Видработ_едимз' : []})
        file_object.to_excel(path_to_create, sheet_name='Запросы', index=False)
