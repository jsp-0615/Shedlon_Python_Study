"""
@author:jsp
@file:Q3.py
@time:2023/4/10 18:53
@desc:
"""

import pandas as pd
from Sql_connection import database_connection
server = '172.22.2.34'
database = 'ORAPS_SCSS_BS_2.0_PURE_1114'
username = 'sa'
password = 'Oraps123'
##############################################################
# dir_path=os.path.dirname(__file__)
dir_path='J:\\Training_Period_Document\\培训考核\\python_test\\'
##############################################################
connection=database_connection(username,password,server,database)
engine=connection.get_engine()
session=connection.create_session()
try:
    result=session.execute("select * from process")
    excel_process = pd.DataFrame(result.fetchall())
    excel_process.to_excel(dir_path + "导出process.xlsx", sheet_name='sheet1', index=False)
except Exception as e:
    print(e)
finally:
    str = connection.db_close(engine)
    print(str)

