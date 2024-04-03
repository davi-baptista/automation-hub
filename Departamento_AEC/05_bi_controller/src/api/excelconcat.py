import pandas as pd
from api import main


def excelconcat(filepath_actual_month, filepath_bd_general):
    # Ler os arquivos
    dados_mes_atual = pd.read_excel(filepath_actual_month)
    dados_acumulados = pd.read_excel(filepath_bd_general)

    # Criar uma chave única combinando CNPJ e MES
    dados_mes_atual['chave_unica'] = dados_mes_atual['CNPJ'].astype(str) + '-' + dados_mes_atual['MES'].astype(str)
    dados_acumulados['chave_unica'] = dados_acumulados['CNPJ'].astype(str) + '-' + dados_acumulados['MES'].astype(str)

    # Identificar registros novos ou para atualização
    novos_ou_para_atualizar = dados_mes_atual['chave_unica'].isin(dados_acumulados['chave_unica'])

    # Atualizar os registros existentes
    for chave in dados_mes_atual[novos_ou_para_atualizar]['chave_unica']:
        dados_acumulados.loc[dados_acumulados['chave_unica'] == chave, dados_mes_atual.columns] = dados_mes_atual.loc[dados_mes_atual['chave_unica'] == chave].values

    # Adicionar novos registros
    dados_novos = dados_mes_atual[~novos_ou_para_atualizar]
    dados_atualizados = pd.concat([dados_acumulados, dados_novos], ignore_index=True)

    dados_atualizados = dados_atualizados.sort_values(['CNPJ', 'MES'])

    # Remover a coluna chave_unica antes de salvar
    dados_atualizados = dados_atualizados.drop(columns=['chave_unica'])
    
    save_excel_with_formatting(dados_atualizados, filepath_bd_general)
    print("Dados atualizados salvos com sucesso.")


def save_excel_with_formatting(df, filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    worksheet = writer.sheets['Sheet1']
    
    # Adicionando uma tabela com estilo
    (max_row, max_col) = df.shape
    column_settings = [{'header': column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 6'})

    # Ajustando o espaçamento das colunas
    for i, col in enumerate(df.columns):
        worksheet.set_column(i, i, 15)

    # Fechar o writer e salvar o arquivo Excel
    writer.close()
    
    
if __name__ == "__main__":
    bd_general = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\05_power_bi\data\empresas.xlsx'
    month = '202401202402'
    
    filepath, status = main(month)
    
    if status == 'OK':
        excelconcat(filepath, bd_general)
    else:
        print(status)