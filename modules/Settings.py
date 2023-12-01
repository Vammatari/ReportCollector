import configparser
import os

class Settings:

    def __init__(self, config_path: str):
        # self._path = '/'.join((os.path.abspath(__file__).replace('\\', '/')).split('/')[:-1])
        # self._config_path = os.path.join(self._path, config_path)
        self._config_path = os.path.join(os.path.dirname(__file__), 'settings.ini')
        self._config_file = self.__init_config()

    def __init_config(self) -> configparser.ConfigParser:
        config = configparser.ConfigParser()
        config.read(self._config_path, encoding="utf-8")
        return config
    # =======================================================
    # начальная директория
    def set_initialdir(self, new_dir: str) -> None:
        self._config_file.set('FILE', 'InitialDir', new_dir)
        self.__save()

    def get_initialdir(self) -> str:
        return self._config_file['FILE']['InitialDir']
    # =======================================================
    # инстанс
    def set_dbinstance(self,new_inst):
        self._config_file.set('DBINSTANCE','dbinst',new_inst)
        self.__save()

    def get_dbinstance(self) -> str:
        return  self._config_file['DBINSTANCE']['dbinst']
    # =======================================================
    # имя базы
    def set_dbname(self,new_db):
        self._config_file.set('DBNAME','db_name',new_db)
        self.__save()

    def get_dbname(self):
        return self._config_file['DBNAME']['db_name']
    # =======================================================
    # ПО-1 план на месяц
    def set_po_month_plan_table_name(self,new_po_month_plan):
        self._config_file.set('TABLENAME', 'po_month_plan_table_name', new_po_month_plan)
        self.__save()

    def get_po_month_plan_table_name(self):
        return self._config_file['TABLENAME']['po_month_plan_table_name']
    # =======================================================
    # ПО-1 факт за месяц
    def set_po_month_fact_table_name(self,new_po_month_fact):
        self._config_file.set('TABLENAME', 'po_month_fact_table_name', new_po_month_fact)
        self.__save()

    def get_po_month_fact_table_name(self):
        return self._config_file['TABLENAME']['po_month_fact_table_name']
    # =======================================================
    # ПО-1 план на год 
    def set_po_year_plan_table_name(self,new_po_year_plan):
        self._config_file.set('TABLENAME', 'po_year_plan_table_name', new_po_year_plan)
        self.__save()
    
    def get_po_year_plan_table_name(self):
        return self._config_file['TABLENAME']['po_year_plan_table_name']
    # =======================================================
    # ОЗ план
    def set_ve_plan_table_name(self,new_ve_plan):
        self._config_file.set('TABLENAME', 've_plan_table_name', new_ve_plan)
        self.__save()

    def get_ve_plan_table_name(self):
        return self._config_file['TABLENAME']['ve_plan_table_name']
    # =======================================================
    # ОЗ факт
    def set_ve_fact_table_name(self,new_ve_fact):
        self._config_file.set('TABLENAME', 've_fact_table_name', new_ve_fact)
        self.__save()

    def get_ve_fact_table_name(self):
        return self._config_file['TABLENAME']['ve_fact_table_name']
    def get_job_list(self):
        return self._config_file['JOB']['job_list']

    # def set_tablename(self,new_table):
    #     self._config_file.set('TABLENAME','table_name',new_table)
    #     self.__save()

    # def get_tablename(self):
    #     return self._config_file['TABLENAME']['table_name']

    def __save(self) -> None:
        with open(self._config_path, 'w+', encoding='utf-8') as new_config:
            self._config_file.write(new_config)



    