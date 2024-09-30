from modules.Table import Table
import pandas as pd


class VolumesResult(Table):
    @staticmethod
    def job(path_to_load: str, path_to_save: str) -> None:
        Table.concat(path_to_load, path_to_save, 'Объёмы-Затраты')

    @staticmethod
    def create_table(path_to_create: str) -> None:
        file_object = pd.DataFrame({'Вид работ': [], 'План/Факт': [], 'Ед. изм.': [], 'Значение': [], 'Организация': [],
                                    'Проект': [], 'Участок': [], 'Филиал': [], 'Лицензия': [],'Подразделение(1С)':[], 'Месяц': [], 'Год': []})
        file_object.to_excel(path_to_create, sheet_name='Объёмы-Затраты', index=False)
