"""
@author:jsp
@file:python_datetime_demo.py
@time:2023/3/11 15:40
@desc:
"""
# from datetime import datetime
#
# #获取当前时间
# v1=datetime.now()
#
#
# string_date=v1.strftime("%Y-%m-%d %H:%M:%S") #2023-03-11 15:43:48
# string_date=v1.strftime("%Y年%m月%d日") #2023年03月11日
# print(string_date)
import os
from datetime import datetime,timedelta
# BASE_PATH = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# log_folder_path= os.path.join(BASE_PATH,'log')
# #在当前根目录下创建log目录用来存记录信息
# if not os.path.exists(log_folder_path):
#     os.mkdir(log_folder_path)
#
# #让用户去注册
# user=input("用户名：")
# pwd=input("密码：")
# record='用户名：'+user+','+'密码：'+pwd+'\n'
#
#
# #写入内容
# date_string=datetime.now().strftime("%Y%m%d")
# # file_name=date_string+".txt"
# file_name="{}.txt".format(date_string)
# file_abs_path=os.path.join(log_folder_path,file_name)
# f=open(file_abs_path,mode='a',encoding='utf8')
# f.write(record)
# f.close()

# t1=datetime.now()
# t2=t1+timedelta(days=10)
# data_string=t2.strftime("%Y-%m-%d")
# print(data_string)


data_string="2023-3-11"
#字符串转换成datetime
datetime_value=datetime.strptime(data_string,"%Y-%m-%d")
print(type(datetime_value))
