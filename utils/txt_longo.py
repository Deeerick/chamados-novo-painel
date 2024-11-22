import pandas as pd

from login_sap import login_sap
from config import USER, PASSWORD


def texto_longo():
    """
    Retorna o texto longo de cada nota de manutenção e insere no campo Escopo do Excel.
    """
    # Conectar ao SAP
    session = login_sap()
    try:
        session.findById("wnd[0]/usr/txtRSYST-BNAME").text = USER
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = PASSWORD
        session.findById("wnd[0]").sendVKey(0)
        session.findById("wnd[0]/tbar[0]/okcd").text = "IW22"
        session.findById("wnd[0]").sendVKey(0)
    except:
        session.findById("wnd[0]/tbar[0]/okcd").text = "IW22"
        session.findById("wnd[0]").sendVKey(0)
    
    # Carregar as notas de manutenção
    df_notas = pd.read_excel('chamados_epm.xlsx')
    notas = df_notas['Nota'].tolist()
    
    for i, nota in enumerate(notas):
        session.findById("wnd[0]/usr/ctxtRIWO00-QMNUM").text = nota
        session.findById("wnd[0]").sendVKey(0)
        
        texto_longo = session.findById("wnd[0]/usr/tabsTAB_GROUP_10/tabp10\\TAB01/ssubSUB_GROUP_10:SAPLIQS0:7235/subCUSTOM_SCREEN:SAPLIQS0:7212/subSUBSCREEN_4:SAPLIQS0:7715/cntlTEXT/shellcont/shell").text
        
        # Inserir o texto longo no campo Escopo do DataFrame
        df_notas.at[i, 'Escopo'] = texto_longo
        
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
    
    # Salvar o DataFrame atualizado de volta no arquivo Excel
    df_notas.to_excel('chamados_epm_atualizado.xlsx', index=False)
    print('Arquivo Excel atualizado com sucesso.')
    
    session.findById("wnd[0]").close()
    session.findById("wnd[1]/usr/btnSPOP-OPTION1").press()


if __name__ == '__main__':
    texto_longo()
