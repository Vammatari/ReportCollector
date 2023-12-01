import pyodbc
from modules.Settings import Settings

class DBConnection:
    def __init__(self) -> None:

        self.settings = Settings('settings.ini')
        self.instance_name = self.settings.get_dbinstance()
        self.db_name = self.settings.get_dbname()
        self.po_month_plan_table_name = self.settings.get_po_month_plan_table_name()
        self.po_month_fact_table_name = self.settings.get_po_month_fact_table_name()
        self.po_year_plan_table_name = self.settings.get_po_year_plan_table_name()
        self.ve_plan_table_name = self.settings.get_ve_plan_table_name()
        self.ve_fact_table_name = self.settings.get_ve_fact_table_name()
        self.conn = self.init_connection()
    
    def init_connection(self):

        self.conn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=' 
            + self.instance_name 
            + ';DATABASE=' 
            + self.db_name 
            + ';Trusted_Connection=True'
            , autocommit=True
            )
        self.cursor = self.conn.cursor()
        return self.cursor
    
    def insert_query_execute(self,table_name, col_list, values) -> object:
        if isinstance(col_list,list):
            col_list = ', '.join(col_list)
        self.query = self.conn.execute(f"INSERT INTO {table_name} ({col_list}) VALUES ( ? {', ? '*(len(col_list) - 1)})",values)

        return self.query

    def select_query_execute(self, table_name, col_list) -> object:
        
        self.query = self.conn.execute(f"SELECT {table_name} FROM {col_list} ")
        return self.query
    
    def select_query_with_condition(self, table_name, col_list, condition):
        if isinstance(col_list,list):
            col_list = ', '.join(col_list)
        self.query = self.conn.execute(f"SELECT {table_name} FROM {col_list} WHERE {condition}")
        return self.query
    




    