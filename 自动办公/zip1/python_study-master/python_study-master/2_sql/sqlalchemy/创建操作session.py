# -*- coding: utf-8 -*-
# @Time : 2020/3/13 15:46
# @Author : lzf
# @File : 创建操作session.py
# @Software: PyCharm
# @Description:

from sqlalchemy import create_engine,inspect
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker

Base = declarative_base()
engine = create_engine('mssql+pymssql://sa:123@127.0.0.1:1433/dbname?charset=utf8')
Session = sessionmaker(bind=engine)
session = Session()
all = session.query(models.table_name).all()
print(len(all))