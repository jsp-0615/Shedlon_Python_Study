"""
@author:jsp
@file:Q1.py
@time:2023/4/10 13:10
@desc:
"""
import os.path

import pandas as pd
import time
from Sql_connection import database_connection

def get_excel(dir_str,fileName):
    dir_path=dir_str
    excel_name=fileName
    file_path=dir_path+excel_name
    excel=pd.read_excel(file_path)
    return excel
def data_to_sql(data,table_name,engine):
    # temp_data=data[columns]
    t1=time.time()
    print('数据开始写入数据库的时间:{0}'.format(t1))
    # 因为一开始已经做了删库的操作，如果if_exists是replace会卡住，所以改成append
    data.to_sql(name=table_name,con=engine,index=False,if_exists='append')
    t2=time.time()
    print('数据结束写入数据库的时间:{0}'.format(t2))
    print('成功插入数据%d条，'%len(data), '耗费时间：%.5f秒。'%(t2-t1))
server = '172.22.2.34'
database = 'ORAPS_SCSS_BS_2.0_PURE_1114'
username = 'sa'
password = 'Oraps123'
connection=database_connection(username,password,server,database)
engine=connection.get_engine()
session=connection.create_session()
#读取数据表A
##############################################################
# dir_path=os.path.dirname(__file__)
dir_path='J:\\Training_Period_Document\\培训考核\\python_test\\'
##############################################################
excel_A_name='数据表A.xlsx'
excel_A=get_excel(dir_path,excel_A_name)
delete_sql='delete from bom_item'
try:
    #开启事务
    session.begin(subtransactions=True)
    # 先删除原bom_item表数据
    delete_result=session.execute(delete_sql)
    print(delete_result.rowcount)
    data_to_sql(excel_A,'bom_item',engine)
    session.commit()
except Exception as e:
    print(e)
finally:
    str = connection.db_close(engine)
    print(str)