from modules.appendix_module.AppendixTable import AppendixTable
from modules.appendix_module.AppendixResult import AppendixResult
from modules.volumes_module.VolumesTable import VolumesTable
from modules.volumes_module.VolumesResult import VolumesResult
from PyQt5.QtWidgets import QFileDialog
from exceptions.ApplicationExceptions import *


files_list = []


def create_file(initialdir: str, code: int) -> str:
    save_path = QFileDialog.getSaveFileName(caption='Сохранить как', directory=initialdir, filter='Excel запросы (*.xlsx)')
    save_path = save_path[0]
    if not save_path:
        raise FileToSaveNotSelectedError
    if code == 1 or code == 2:
        AppendixResult.create_table_month(save_path)
    elif code == 3 or code == 4:
        VolumesResult.create_table(save_path)

    return save_path


def add_tables(initialdir: str) -> list:
    files = QFileDialog.getOpenFileNames(caption='Выберите таблицы', directory=initialdir,
                                         filter='Таблицы (*.xlsm *.xlsx)')
    files = files[0]
    files_list.extend(files)
    return files


def make_table(initialdir: str, code: int):
    if files_list:
        saving_path = create_file(initialdir, code)
        for book in range(len(files_list)):
            __fill_with_book(path_to_load=files_list[book], path_to_save=saving_path, code=code)
        files_list.clear()
    else:
        raise FileToReadNotSelectedError


def __fill_with_book(path_to_load: str, path_to_save: str, code: int) -> None:
    if code == 1:
        AppendixTable.job(path_to_load, path_to_save)
    elif code == 2:
        AppendixResult.job(path_to_load, path_to_save)
    elif code == 3:
        VolumesTable.job(path_to_load, path_to_save)
    elif code == 4:
        VolumesResult.job(path_to_load, path_to_save)

