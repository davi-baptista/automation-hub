import pandas as pd
from API.api import main

def merge_excel(filepath_bd_roberta, filepath_bd_athenas):
    # Carregar as planilhas, especificando que a coluna 'CNPJ' deve ser tratada como string
    planilha1 = pd.read_excel(filepath_bd_roberta, sheet_name='ATIVA', dtype=str, skiprows=1)
    planilha2 = pd.read_excel(filepath_bd_athenas, dtype=str)

    # Preparação e limpeza de dados
    planilha1['CNPJ'] = limpar_cnpj(planilha1['CNPJ'])
    planilha2['CNPJ'] = limpar_cnpj(planilha2['CNPJ'])
    planilha1['GRUPO'] = planilha1['GRUPO'].str.strip()
    planilha2['GRUPO'] = planilha2['GRUPO'].str.strip()
    planilha2['NOMESTATUS'] = planilha2['NOMESTATUS'].str.strip()

    # Filtrando planilha pelas empresas com status igual a ativo
    planilha2_ativos = planilha2[planilha2['NOMESTATUS'] == 'ATIVO']

    # Pegando apenas as colunas que eu preciso usar e criando um novo df com elas para facilitar as comparações
    cnpj_planilha1 = planilha1[['CNPJ', 'RAZÃO SOCIAL','GRUPO']]
    cnpj_planilha2 = planilha2_ativos[['CNPJ', 'NOME','GRUPO']]

    # Juntar as colunas 'CNPJ' das duas planilhas com base na coluna 'CNPJ'
    merged = cnpj_planilha1.merge(cnpj_planilha2, on='CNPJ', how='outer', indicator=True, suffixes=('_planilha1', '_planilha2'))

    # Adicionar uma nova coluna 'Status' com base na coluna '_merge'
    # Adicionar uma nova coluna 'Status_Grupo' com base na comparação da coluna 'GRUPO'
    def determine_status_and_group(row):
        if row['_merge'] == 'both':
            status = 'MANTER'
            # Verifica se os valores de 'GRUPO' batem
            if row['GRUPO_planilha1'] == row['GRUPO_planilha2']:
                status_grupo = 'OK'
            else:
                status_grupo = f'Modificar para {row["GRUPO_planilha1"]}'
            return row['RAZÃO SOCIAL'] if pd.notnull(row['RAZÃO SOCIAL']) else row['NOME'], status, status_grupo
        elif row['_merge'] == 'left_only':
            status = 'ATIVAR'
            status_grupo = F'Adicionar: {row['GRUPO_planilha1']}'  # Não aplicável pois não há comparação de grupo
            return row['RAZÃO SOCIAL'], status, status_grupo
        else:  # right_only
            status = 'INATIVAR'
            status_grupo = 'N/A'  # Não aplicável pois não há comparação de grupo
            return row['NOME'], status, status_grupo

    merged['NAME'], merged['Status'], merged['Status_Grupo'] = zip(*merged.apply(determine_status_and_group, axis=1))

    # Agora remova as colunas que não são mais necessárias
    merged.drop(columns=['RAZÃO SOCIAL', 'NOME', '_merge', 'GRUPO_planilha1', 'GRUPO_planilha2'], inplace=True)

    # Remover duplicatas, mantendo a primeira ocorrência
    merged = merged.drop_duplicates(subset=['CNPJ', 'NAME'])

    # Salvar a nova planilha
    print('Planilha Feita')
    save_excel_with_formatting(merged, r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\02_bater_grupos\output\result_spreadsheet.xlsx')

def limpar_cnpj(cnpj):
    return cnpj.str.replace(r'[.,/ -]', '', regex=True).str.strip()

def limpar_grupo(grupo):
    # Primeiro, remove espaços no início e no fim
    grupo_limpo = grupo.str.strip()
    return grupo_limpo

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
        worksheet.set_column(i, i, 25)

    # Fechar o writer e salvar o arquivo Excel
    writer.close()
    
if __name__ == "__main__":
    filepath, status = main()
    
    if status == 'OK':
        merge_excel(r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\02_bater_grupos\data\BASE DE DADOS DE CLIENTES ATIVOS.xlsx', filepath)
    else:
        print(status)