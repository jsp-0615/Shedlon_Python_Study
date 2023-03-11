import pyodbc

def main():
    server = 'DESKTOP-DH171PS\SQLSERVERSHELDON'  # SQL Server服务器名称
    database = 'DB_ORAPS'  # 数据库名称
    username = 'sa'  # 数据库用户名
    password = '0615..jsp.'  # 数据库密码
    conn = pyodbc.connect(
        'DRIVER={SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
    with conn.cursor() as cursor:#上下文语法
        result=cursor.execute('SELECT * FROM t_emp')
        for row in cursor:
            print(row)
        # for row in result:
        #     print(row)

    try:
        with conn.cursor() as cursor:  # 上下文语法
            result = cursor.execute('SELECT * FROM t_emp')
            for row in cursor:
                print(row)
            #cursor.execute('insert into t_emp(id,name) values(15,"蒋天生")')
            conn.commit()
    except pyodbc.Error as error:
        print(error)
        conn.rollback()
    finally:
        conn.close()

if __name__=='__main__':
    main()