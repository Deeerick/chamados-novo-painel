import os
import pyodbc
import jaydebeapi

from utils.config import SERVER, DATABASE, TRUSTED_CONNECTION, USER, PASSWORD_TDV, URL, TDV


def conexao_sql():
    st_server = SERVER
    st_database = DATABASE
    st_trusted_connection = TRUSTED_CONNECTION

    conn_str = f'DRIVER={{SQL Server}};SERVER={st_server};DATABASE={st_database};Trusted_Connection={st_trusted_connection}'
    conn_sql = pyodbc.connect(conn_str)

    # print('Conexão com o SQL Server realizada com sucesso!')

    return conn_sql


def conexao_tdv():
    os.environ["CLASSPATH"] = TDV

    conn_tdv = jaydebeapi.connect(
        "cs.jdbc.driver.CompositeDriver", URL, [USER, PASSWORD_TDV])
    
    print('Conexão com o TDV realizada com sucesso!')

    return conn_tdv
