"""
@author:jsp
@file:Sql_connection.py
@time:2023/4/11 9:38
@desc:
"""
from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import pymssql

class database_connection:
    def __init__(self,username,password,server,database):
        self._dbserver=server
        self._dbdatabase=database
        self._dbusername=username
        self._dbpassword=password
        self._engine=self.get_engine()
        self._session=self.create_session()

    def get_engine(self):
        try:
            engine = create_engine("mssql+pymssql://" + self._dbusername + ":" + self._dbpassword + "@" + self._dbserver + "/" + self._dbdatabase,echo=True)
        except Exception as e:
            raise
            engine=False
        finally:
            return engine
    def create_session(self):
        if (self._engine):
            Session=sessionmaker(bind=self._engine)
            session=Session()
            return session
        else:
            raise
    def db_close(self,engine):
        if (engine):
            try:
                if(self._engine!=None):
                    self._engine.dispose()
                    return '关闭连接成功'
            except Exception as e:
                return e.__context__

