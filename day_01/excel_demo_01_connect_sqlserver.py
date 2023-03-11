import pyodbc
server = 'DESKTOP-DH171PS\SQLSERVERSHELDON' # SQL Server服务器名称
database = 'DB_ORAPS' # 数据库名称
username = 'sa' # 数据库用户名
password = '0615..jsp.' # 数据库密码
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+ password)
if cnxn:
    print("连接成功")
cursor = cnxn.cursor()
cursor.execute('SELECT * FROM t_emp')
for row in cursor:
    print(row)