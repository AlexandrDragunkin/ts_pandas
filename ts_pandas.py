# -*- coding: utf-8 -*-

import pandas as pd
import pyodbc

def list_tables(tab_place):
    """
    :param tab_place: Файл базы данных MS Access.
    :return: Список таблиц в данной базе данных.
    """
    with pyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};Dbq='
    +tab_place+';Uid=;Pwd=;') as con:
        crs = con.cursor()
        tables = [table_info.table_name for table_info in crs.tables(tableType='TABLE')]
    return tables
    

def read_table(tab_place, table_name, *args, **kwargs):
    """
    Чтение базы данных MS Access в виде фрейма данных Pandas.

    :param tab_place: Файл базы данных MS Access.
    :param table_name: Имя обрабатываемой таблицы.
    :param args: позиционные аргументы pandas.read_sql
    :param kwargs: аргументы ключевых слов pandas.read_sql

    """
    try:
        with pyodbc.connect('Driver={Microsoft Access Driver (*.mdb)};Dbq='
        +tab_place+';Uid=;Pwd=;') as con:
            if 'sql' in kwargs:
                sqlquery = kwargs['sql']
            else:
                sqlquery = 'SELECT * FROM {};'.format(table_name)
            df = pd.read_sql(sqlquery, con, *args, **kwargs)
            return df
    except Exception as e:
        print('Error Table {0}: {1}'.format(table_name, e.args))

        
if __name__=='__main__':
    mdb_path_ARL = r'c:\ARL7\Data\PKM\Proto\DoorSys63.mdb'
    lt=list_tables(mdb_path_ARL)
    print(lt)
    df=read_table(mdb_path_ARL,  'DSColorProfile')
    print(df)
