import pandas as pd
import openpyxl


class Table:
    _path_to_save: str
    _path_to_load: str

    def __init__(self, path_to_load: str, path_to_save: str):
        self._path_to_load = path_to_load
        self._path_to_save = path_to_save

    @staticmethod
    def get_length(df: pd.DataFrame) -> int:
        return len(df.columns.values)

    @staticmethod
    def concat(path_to_load: str, path_to_save: str, sheet_name: str) -> None:
        try:
            result = openpyxl.load_workbook(path_to_save)
            sheet = result[sheet_name]
            df = pd.read_excel(path_to_load, sheet_name, engine='openpyxl')
            length = Table.get_length(df)
            for row in range(df.shape[0]):
                new_row = []
                for col in range(length):
                    new_row.append(df.iloc[row, col])
                sheet.append(new_row)
            result.save(path_to_save)
        except PermissionError:
            pass
