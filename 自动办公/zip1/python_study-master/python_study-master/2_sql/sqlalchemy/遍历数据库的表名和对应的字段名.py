# -*- coding: utf-8 -*-
# @Time : 2020/3/13 15:43
# @Author : lzf
# @File : 遍历数据库的表名和对应的字段名.py
# @Software: PyCharm
# @Description:



from sqlalchemy import create_engine,inspect
from sqlalchemy.ext.declarative import declarative_base
Base = declarative_base()
engine = create_engine('mssql+pymssql://sa:123@127.0.0.1:1433/dbname?charset=utf8')
inspector = inspect(engine)

for table_name in inspector.get_table_names():
    print(table_name)
    for column in inspector.get_columns(table_name):
       print("Column: %s" % column['name'])


