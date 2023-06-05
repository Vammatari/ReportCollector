import pyodbc
instance_name = 'INDSQL\TESTDB'
db_name = 'PM_GEO_KAR'
table_name = 'po.PLM_GRR_PLAN_YEAR'
conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+instance_name+';DATABASE='+db_name+';Trusted_Connection=True',autocommit=True)

cursor = conn.cursor()

# # query = cursor.execute("SELECT [Категория видов работ (ПО-1)] FROM " + table_name + " WHERE [Категория видов работ (ПО-1)] = 'Анализ шлиховых проб'")

# # listw = list(query.fetchall())


# # print(*listw,sep='\n')
# def query_ex(curs,table_name, values = ""):
#     a = curs.execute("SELECT * FROM " + table_name)
#     return print(*list(a.fetchall()),sep='\n')

# query_ex(cursor,table_name)

# # query2 = cursor.execute(
# #     "INSERT INTO (@s)" + 
# #     table_name + 
# #     " ([PROJECT_ID],[AREA_ID],[YEAR],[WORK_TYPE_PO_ID],[MEASUREMENT_UNITS_ID],[AMOUNT]) VALUES (?,?,?,?,?,?)", 
# #     [12,12,2023,12,12,121212])
col_list = ['[PROJECT_ID]','[AREA_ID]','[YEAR]','[WORK_TYPE_PO_ID]','[MEASUREMENT_UNITS_ID]','[AMOUNT]']
values = [14,14,2023,14,14,141414]
# query = cursor.execute("INSERT INTO " + table_name + "(" + ', '.join(col_list) + ") VALUES ( ?" + ", ? "*(len(col_list) - 1)  + ")",values)
query = cursor.execute(f"INSERT INTO {table_name} ({', '.join(col_list)}) VALUES ( ? {', ? '*(len(col_list) - 1)})",values)
# # list2 = list(query2.fetchall())

# # print(*list2, sep='\n')

# list = ['q','w','e','r','t']
# print(list)
# a = ', '.join(list)
# print(a)