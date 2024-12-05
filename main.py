import re
import pandas as pd

from utils.config import QUERY
from utils.conexao_database import conexao_sql, conexao_tdv
from datetime import timedelta


def main():
    
    # Conectar ao banco de dados e carregar os dados
    conn_sql = conexao_sql()
    df = pd.read_sql_query(QUERY, conn_sql)
    print('Dados carregados.')
    
    # Converter as colunas de data para o formato dd/mm/aaaa
    df['Início'] = pd.to_datetime(df['Início']).dt.strftime('%d/%m/%Y')
    df['Término'] = pd.to_datetime(df['Término']).dt.strftime('%d/%m/%Y')
    df['Próxima_Atualização'] = pd.to_datetime(df['Próxima_Atualização']).dt.strftime('%d/%m/%Y')
    
    # Aplicar a função para calcular as datas de término e próxima atualização
    df['Término'] = df.apply(lambda row: adicionar_dias(row['Início']) if pd.isna(row['Término']) else row['Término'], axis=1)
    df['Próxima_Atualização'] = df['Término']

    # Tratar valores NaN na coluna 'Plataforma'
    df['Plataforma'] = df['Plataforma'].fillna('')

    # Aplicar a função tratar_plataforma
    df['Plataforma'] = df['Plataforma'].apply(tratar_plataforma)

    # Salvar o DataFrame tratado de volta no arquivo Excel
    output_file = 'chamados_epm.xlsx'
    df.to_excel(output_file, index=False)

    print(f'Arquivo Excel criado.')
    

# Função para calcular a data adicionando 90 dias e ajustando para evitar finais de semana
def adicionar_dias(data_inicio):
    
    # Se data_inicio for NaN, retornar NaN
    # if pd.isna(data_inicio):
    #     return data_inicio
    
    # Adicionar 90 dias
    data_inicio_dt = pd.to_datetime(data_inicio, format='%d/%m/%Y')
    data_termino = data_inicio_dt + timedelta(days=90)
    
    # Se cair no sábado (5) ou domingo (6)
    if data_termino.weekday() > 4:
        data_termino -= timedelta(days=(data_termino.weekday() - 4))
        
    data_termino = pd.to_datetime(data_termino).strftime('%d/%m/%Y')
        
    return data_termino


# Função para tratar a coluna 'Plataforma'
def tratar_plataforma(plataforma):
    
    # Se plataforma for NaN, retornar string vazia
    if plataforma is None or pd.isna(plataforma):
        return ''
    
    # Se plataforma for uma string de 3 caracteres, retornar a string com o hífen no meio
    if re.match(r'^[A-Z]\d{2}$', plataforma):
        return f'{plataforma[0]}-{plataforma[1:]}'
    
    # Se plataforma for uma string de 2 caracteres e um dígito, retornar a string com "P" no começo e o hífen no meio
    elif re.match(r'^[A-Z]{2}\d$', plataforma):
        return f'P{plataforma[:2]}-{plataforma[2]}'
    
    return plataforma


if __name__ == '__main__':
    main()
