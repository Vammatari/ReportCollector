import pandas as pd
from modules.Table import Table


class AppendixResult(Table):
    @staticmethod
    def job(path_to_load: str, path_to_save: str) -> None:
        Table.concat(path_to_load, path_to_save, 'Запросы')

    @staticmethod
    def create_table_month(path_to_create: str) -> None:
        # file_object = pd.DataFrame(
        #     {'Вид работ': [], 'Ед. изм': [], 'Месяц': [], 'Год:': [], 'План/Факт': [], 'Значение': [],
        #      'Проект (Паспорт ГРП)': [], 'Участок': [], 'Стадия': [], 'Проект-Участок': [], 'План на год': []})
        month_results = pd.DataFrame({'WORK_NAME': [], 'MEAS_UNITS': [], 'MONTH': [], 'YEAR': [], 'PLAN_FACT': [], 'AMOUNT': [],
                                    'PROJECT': [], 'DISTRICT': [],'LICENSE': [], 'STAGE': [], 'PROJECT_AREA': [],'WORK_TYPE_MU' : []})
        year_results = pd.DataFrame({'WORK_NAME': [], 'MEAS_UNITS': [], 'YEAR': [], 'AMOUNT': [],
                                    'PROJECT': [], 'DISTRICT': [],'LICENSE': [], 'STAGE': [], 'PROJECT_AREA': [],'WORK_TYPE_MU' : []})
        
        with pd.ExcelWriter(path_to_create) as writer:
            month_results.to_excel(writer, sheet_name='PO_2024_MONTH_TEMP', index=False)
            year_results.to_excel(writer, sheet_name='PO_2024_YEAR_TEMP', index=False)
    
