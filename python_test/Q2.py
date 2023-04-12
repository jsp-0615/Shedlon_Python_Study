"""
@author:jsp
@file:Q2.py
@time:2023/4/10 18:48
@desc:
"""
import pandas as pd

def get_excel(dir_str,fileName):
    dir_path=dir_str
    excel_name=fileName
    file_path=dir_path+excel_name
    excel=pd.read_excel(file_path)
    return excel
#读取数据表A
##############################################################
# dir_path=os.path.dirname(__file__)
dir_path='J:\\Training_Period_Document\\培训考核\\python_test\\'
##############################################################
excel_A_name='数据表A.xlsx'
excel_A=get_excel(dir_path,excel_A_name)
excel_B_name="数据表B.xlsx"
excel_B=get_excel(dir_path,excel_B_name)
temp_excel=excel_B
columns=excel_B.columns
temp=excel_A[temp_excel.iloc[0,:]]
temp.columns=excel_B.columns
excel_B=pd.concat([excel_B,temp],ignore_index=True)
excel_B.to_excel(dir_path+'合并后的数据表B.xlsx',sheet_name='sheet1', index=False)