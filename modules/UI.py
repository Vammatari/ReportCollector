import typing
from PyQt5.QtWidgets import QPushButton, QMessageBox, QMainWindow, QRadioButton, QAction, QListWidget,\
    QListWidgetItem, QButtonGroup,QWidget,QLabel,QVBoxLayout, QLineEdit
from PyQt5 import QtCore, QtGui
from PyQt5.QtCore import pyqtSlot
from modules.AppLogic import *
from modules.Settings import Settings
import os


class UI(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.width = 350
        self.height = 102
        self.dy = 25
        self.code = 0
        self.list = QListWidget(self)
        self.settings = Settings('settings.ini')
        self.init_main_ui()
        self.list.itemDoubleClicked.connect(self.remove)
    def init_main_ui(self) -> None:
        self.setWindowTitle('ReportCollector')
        self.setFixedSize(self.width, self.height)
        self.move(800, 400)
        self.setWindowIcon(QtGui.QIcon('excel-ico.ico'))

        # Прогресс бар

# progress_bar = QProgressBar(self)
# progress_bar.setGeometry(0, 100 + self.dy, self.width + 33, 20)
# timer = QBasicTimer()

        # Экшены

        initaldir_action = QAction('Начальная директория', self)
        initaldir_action.setStatusTip('Установить начальную директорию')
        initaldir_action.triggered.connect(self.set_initialdir)
        connection_action = QAction('Подключение к БД',self)
        connection_action.triggered.connect(self.open_connection)
        # Меню бар

        menu_bar = self.menuBar()
        initialdir_menu = menu_bar.addMenu('Настройки')
        initialdir_menu.addAction(initaldir_action)
        initialdir_menu.addAction(connection_action)
        
        # Список

        self.list.setGeometry(15, 130, 320, 300)
        self.list.hide()

        # Ниже блок с кнопками ПО-1 -> Запросы

        self.app_button = QPushButton('ПО-1', self)
        self.app_button.setToolTip('Собрать таблицу запросов из таблиц приложений')
        self.app_button.setGeometry(25, 15 + self.dy, 150, 50)
        self.app_button.clicked.connect(self.app_request_on_click)

        # Ниже блок с кнопками Объёмы -> Результат

        self.volumes_button = QPushButton('Объёмы-Затраты', self)
        self.volumes_button.setToolTip('Собрать таблицу из таблиц объёмов')
        self.volumes_button.setGeometry(175, 15 + self.dy, 150, 50)
        self.volumes_button.clicked.connect(self.volumes_on_click)
        # self.volumes_button.setEnabled(False)

        self.show()

    def init_work_ui(self) -> None:
        
        self.setFixedSize(350, 540)
        self.list.show()
        
        # self.make_radio = QRadioButton("Собрать", self)
        # self.make_radio.setGeometry(50, 75 + self.dy, 100, 20)
        # self.make_radio.setChecked(True)
        # self.make_radio.selected = True
        # self.make_radio.clicked.connect(self.make_radio_on_click)
        # self.make_radio.show()
        self.back_to_main_button = QPushButton('Вернуться к выбору задачи', self)
        self.back_to_main_button.setGeometry(25, 65 + self.dy,300,25)
        self.back_to_main_button.show()
        self.back_to_main_button.clicked.connect(self.init_main_ui)
        # self.concat_radio = QRadioButton("Сложить", self)
        # self.concat_radio.setGeometry(200, 75 + self.dy, 100, 20)
        # self.concat_radio.selected = False
        # self.concat_radio.clicked.connect(self.concat_radio_on_click)
        # self.concat_radio.show()

        # self.radio_group = QButtonGroup()
        # self.radio_group.addButton(self.make_radio)
        # self.radio_group.addButton(self.concat_radio)

        self.add_button = QPushButton('Добавить', self)
        self.add_button.setGeometry(25, 410 + self.dy, 150, 50)
        self.add_button.clicked.connect(self.add_button_on_click)
        self.add_button.show()

        self.start_button = QPushButton('Создать Excel-файл', self)
        self.start_button.setGeometry(175, 410 + self.dy, 150, 50)
        self.start_button.clicked.connect(self.start_button_on_click)
        self.start_button.show()
        self.db_button = QPushButton('Загрузить в БД', self)
        self.db_button.setGeometry(25,460 + self.dy, 300, 50)
        self.db_button.show()
        self.db_button.setEnabled(False)

    # def init_third_ui(self) -> None:
    #     self.setFixedSize(350, 540)

    #     self.list.show()

    #     self.make_radio = QRadioButton("Собрать", self)
    #     self.make_radio.setGeometry(50, 75 + self.dy, 100, 20)
    #     self.make_radio.setChecked(True)
    #     self.make_radio.selected = True
    #     self.make_radio.clicked.connect(self.make_radio_on_click)
    #     self.make_radio.show()
    #     self.make_radio.setHidden(True)

    #     self.concat_radio = QRadioButton("Сложить", self)
    #     self.concat_radio.setGeometry(200, 75 + self.dy, 100, 20)
    #     self.concat_radio.selected = False
    #     self.concat_radio.clicked.connect(self.concat_radio_on_click)
    #     self.concat_radio.show()
    #     self.concat_radio.setHidden(True)

    #     self.radio_group = QButtonGroup()
    #     self.radio_group.addButton(self.make_radio)
    #     self.radio_group.addButton(self.concat_radio)

    #     self.add_button = QPushButton('Добавить', self)
    #     self.add_button.setGeometry(25, 410 + self.dy, 150, 50)
    #     self.add_button.clicked.connect(self.add_button_on_click)
    #     self.add_button.show()

    #     self.start_button = QPushButton('Создать Excel-файл', self)
    #     self.start_button.setGeometry(175, 410 + self.dy, 150, 50)
    #     self.start_button.clicked.connect(self.start_button_on_click)
    #     self.start_button.show()
    #     self.db_button = QPushButton('Загрузить в БД', self)
    #     self.db_button.setGeometry(25,460 + self.dy, 300, 50)
    #     self.db_button.show()
    #     self.db_button.setEnabled(False)

    def set_initialdir(self):
        self.settings.set_initialdir(QFileDialog.getExistingDirectory(self, 'Установить директорию',directory='C:\\'))
    def open_connection(self):
        self.w = DbWindow()
        self.w.show()

    # @pyqtSlot()
    # def make_radio_on_click(self):
    #     if not self.make_radio.selected:
    #         self.code -= 1
    #         self.make_radio.selected = True
    #         self.concat_radio.selected = False

    # @pyqtSlot()
    # def concat_radio_on_click(self):
    #     if not self.concat_radio.selected:
    #         self.code += 1
    #         self.concat_radio.selected = True
    #         self.make_radio.selected = False

    @pyqtSlot()
    def add_button_on_click(self):
        self.fill_list(add_tables(initialdir=self.settings.get_initialdir()))

    @pyqtSlot()
    def start_button_on_click(self):
        try:
            make_table(code=self.code, initialdir=self.settings.get_initialdir())
            self.msg_success()
            self.list.clear()
        except PermissionError:
            self.msg_save_error()
        except FileToReadNotSelectedError:
            self.msg_read_warning()
        except FileToSaveNotSelectedError:
            self.msg_save_warning()
        except Exception as e:
            self.msg_exception(e)

    @pyqtSlot()
    def app_request_on_click(self):
        self.code = 1
        self.volumes_button.setEnabled(False)
        self.init_work_ui()

    @pyqtSlot()
    def volumes_on_click(self):
        self.code = 3
        self.app_button.setEnabled(False)
        self.init_work_ui()

    @staticmethod
    def msg_success():
        msg = QMessageBox()
        msg.setText('Перенос завершён')
        msg.setInformativeText('Таблица успешно собрана')
        msg.setWindowTitle('Готово')
        msg.setIcon(QMessageBox.Information)
        msg.exec_()
        

    @staticmethod
    def msg_read_warning():
        msg = QMessageBox()
        msg.setText('Файл не выбран')
        msg.setInformativeText('Вы не выбрали ни одного файла для чтения')
        msg.setWindowTitle('Предупреждение')
        msg.setIcon(QMessageBox.Warning)
        msg.exec_()

    @staticmethod
    def msg_save_warning():
        msg = QMessageBox()
        msg.setText('Файл не выбран')
        msg.setInformativeText('Вы не выбрали ни одного файла для записи')
        msg.setWindowTitle('Предупреждение')
        msg.setIcon(QMessageBox.Warning)
        msg.exec_()

    @staticmethod
    def msg_save_error():
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText('Ошибка доступа к файлу')
        msg.setInformativeText('Проверьте, возможно файл был открыт в другом окне')
        msg.setWindowTitle('Ошибка')
        msg.exec_()

    @staticmethod
    def msg_exception(text):
        msg = QMessageBox()
        msg.setText(
            'Если вы это читаете, пожалуйста, отправьте мне сообщение на PortynkinDV@polymetal.ru с текстом ниже.')
        msg.setInformativeText(text)
        msg.setWindowTitle('Исключение')
        msg.setIcon(QMessageBox.Warning)
        msg.exec_()

    def fill_list(self, arr: list) -> None:
        try:
            for item in arr:
                list_item = QListWidgetItem()
                list_item.setIcon(QtGui.QIcon('list-icon.png'))
                list_item.setText(os.path.basename(item))
                self.list.addItem(list_item)
            print(len(self.list))
            print(files_list)
        except Exception as e:
            print(str(e))
    
    def remove(self,current_item):
        current_row = self.list.currentRow()
        if current_row >= 0:
            current_item = self.list.takeItem(current_row)
            print(len(self.list))
            print(files_list)
            print(current_item.text())
            for _ in files_list:
                if current_item.text() in _:
                    files_list.remove(_)
                else:
                    print('ffgfgf')
            print(files_list)



class DbWindow(QWidget):
    
    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.width = 350
        self.height = 415
        self.dx = 20
        self.dy = 25

        self.settings = Settings('settings.ini')
        self.init_db_connection()


    def init_db_connection(self):
        self.setWindowTitle('DB Connection')
        self.setFixedSize(self.width,self.height)
        self.setWindowIcon(QtGui.QIcon('database.ico'))
# ------------------------------------------------------------------
        self.inst_label = QLabel('Имя сервера:',self)
        self.inst_label.setGeometry(10, 10, self.width - self.dx, 25)
        self.insttext = QLineEdit(self)
        self.insttext.setGeometry(10, self.inst_label.y() + self.dy, self.width - self.dx, 25)
        self.insttext.setText(self.settings.get_dbinstance())
# ------------------------------------------------------------------
        self.db_label = QLabel('Имя базы данных:',self)
        self.db_label.setGeometry(10, 60, self.width - self.dx, 25)
        self.dbtext = QLineEdit(self)
        self.dbtext.setGeometry(10,85, self.width - self.dx, 25)
        self.dbtext.setText(self.settings.get_dbname())
# ------------------------------------------------------------------
        self.po_month_plan_table_label = QLabel('Имя таблицы(ПО-1 план на месяц):', self)
        self.po_month_plan_table_label.setGeometry(10,110,self.width - self.dx,25)
        self.po_month_plan_tabletext = QLineEdit(self)
        self.po_month_plan_tabletext.setGeometry(10,self.po_month_plan_table_label.y() + self.dy,self.width - self.dx,25)
        self.po_month_plan_tabletext.setText(self.settings.get_po_month_plan_table_name())

        self.po_month_fact_table_label = QLabel('Имя таблицы(ПО-1 факт за месяц):', self)
        self.po_month_fact_table_label.setGeometry(10, self.po_month_plan_tabletext.y() + self.dy, self.width - self.dx, 25)
        self.po_month_fact_tabletext = QLineEdit(self)
        self.po_month_fact_tabletext.setGeometry(10,self.po_month_fact_table_label.y() + self.dy, self.width - self.dx,25)
        self.po_month_fact_tabletext.setText(self.settings.get_po_month_fact_table_name())

        self.po_year_plan_table_label = QLabel('Имя таблицы(ПО-1 план на год):', self)
        self.po_year_plan_table_label.setGeometry(10,self.po_month_fact_tabletext.y() + self.dy, self.width - self.dx, 25)
        self.po_year_plan_tabletext = QLineEdit(self)
        self.po_year_plan_tabletext.setGeometry(10,self.po_year_plan_table_label.y() + self.dy, self.width - self.dx, 25)
        self.po_year_plan_tabletext.setText(self.settings.get_po_year_plan_table_name())

        self.ve_plan_table_label = QLabel('Имя таблицы(ОЗ план):',self)
        self.ve_plan_table_label.setGeometry(10, self.po_year_plan_tabletext.y() + self.dy, self.width - self.dx, 25)
        self.ve_plan_tabletext = QLineEdit(self)
        self.ve_plan_tabletext.setGeometry(10, self.ve_plan_table_label.y() + self.dy, self.width - self.dx, 25)
        self.ve_plan_tabletext.setText(self.settings.get_ve_plan_table_name())

        self.ve_fact_table_label = QLabel('Имя таблицы(ОЗ факт):',self)
        self.ve_fact_table_label.setGeometry(10,self.ve_plan_tabletext.y() + self.dy, self.width - self.dx, 25)
        self.ve_fact_tabletext = QLineEdit(self)
        self.ve_fact_tabletext.setGeometry(10,self.ve_fact_table_label.y() + self.dy, self.width - self.dx, 25)
        self.ve_fact_tabletext.setText(self.settings.get_ve_fact_table_name())
        self.inst_button = QPushButton('Сохранить',self)
        self.inst_button.setGeometry(10,self.ve_fact_tabletext.y() + self.dy + 10,self.width - self.dx,25)
        self.inst_button.clicked.connect(self.set_connection)
        self.inst_button.clicked.connect(self.close)

    def set_connection(self):
        self.settings.set_dbinstance(self.insttext.text())
        self.settings.set_dbname(self.dbtext.text())
        self.settings.set_po_month_plan_table_name(self.po_month_plan_tabletext.text())
        self.settings.set_po_month_fact_table_name(self.po_month_fact_tabletext.text())
        self.settings.set_po_year_plan_table_name(self.po_year_plan_tabletext.text())
        self.settings.set_ve_plan_table_name(self.ve_plan_tabletext.text())
        self.settings.set_ve_fact_table_name(self.ve_fact_tabletext.text())

class POWindow(QWidget):

    def __init__(self, parent = None):
        super().__init__(parent)
        