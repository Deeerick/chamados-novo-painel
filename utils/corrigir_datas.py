import pandas as pd
from datetime import timedelta, datetime


# Leitura do Excel com os dados
df = pd.read_excel('chamados_epm_atualizado.xlsx')


# Função para ajustar e formatar as datas
def ajustar_e_formatar_data(coluna) -> None:
    df[coluna] = pd.to_datetime(df[coluna], format='%d/%m/%Y')
    df[coluna] = df[coluna].apply(ajustar_data)
    df[coluna] = df[coluna].dt.strftime('%d/%m/%Y')


def ajustar_data(data):
    if data.weekday() == 5:  # Se for sábado
        return data - timedelta(days=1)
    elif data.weekday() == 6:  # Se for domingo
        return data - timedelta(days=2)
    else:
        return data


def main():
    colunas_para_ajustar = ['Início',
                            'Término', 'Próxima_Atualização']
    for coluna in colunas_para_ajustar:
        ajustar_e_formatar_data(coluna)


    # Salvando os valores convertidos
    data_atual = datetime.now().strftime('%d.%m.%Y')
    df.to_excel(f'{data_atual} - chamados_epm_final.xlsx', index=False)
    print('Processo concluído!')


if __name__ == '__main__':
    main()
