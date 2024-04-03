import pandas as pd

# Função para limpar a coluna CNPJ
def limpar_cnpj(cnpj):
    return cnpj.str.replace(r'[.,/ -]', '', regex=True).str.strip()

# Carregar as planilhas, especificando que a coluna 'CNPJ' deve ser tratada como string
planilha1 = pd.read_excel(r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\01_bater_status\data\BASE DE DADOS DE CLIENTES ATIVOS V.11_01_2024.xlsx', sheet_name='ATIVA', dtype={'CNPJ': str}, skiprows=1)
planilha2 = pd.read_excel(r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\01_bater_status\data\EMPRESAS ATHENAS.xlsx', dtype={'CNPJ': str})

# Converter a coluna 'CNPJ' para string em ambas as planilhas
planilha1['CNPJ'] = limpar_cnpj(planilha1['CNPJ'].astype(str))
planilha2['CNPJ'] = limpar_cnpj(planilha2['CNPJ'].astype(str))
planilha2['NOMESTATUS'] = planilha2['NOMESTATUS'].str.strip()

planilha2_ativos = planilha2[planilha2['NOMESTATUS'] == 'ATIVO']

cnpj_planilha1 = planilha1[['CNPJ', 'RAZÃO SOCIAL']]
cnpj_planilha2 = planilha2_ativos[['CNPJ', 'NOME']]

# Juntar as colunas 'CNPJ' das duas planilhas com base na coluna 'CNPJ'
merged = cnpj_planilha1.merge(cnpj_planilha2, on='CNPJ', how='outer', indicator=True)

# Adicionar uma nova coluna 'Status' com base na coluna '_merge'
def determine_status(row):
    if row['_merge'] == 'both':
        print('MANTER')
        return row['RAZÃO SOCIAL'], 'MANTER'
    elif row['_merge'] == 'left_only':
        print('ATIVAR')
        return row['RAZÃO SOCIAL'], 'ATIVAR'
    else:  # right_only
        print('INATIVAR')
        return row['NOME'], 'INATIVAR'

merged['NAME'], merged['Status'] = zip(*merged.apply(determine_status, axis=1))

# Remover a coluna auxiliar '_merge'
merged.drop(columns=['RAZÃO SOCIAL', 'NOME', '_merge'], inplace=True)

# Remover duplicatas, mantendo a primeira ocorrência
merged = merged.drop_duplicates(subset=['CNPJ', 'NAME'])

# Salvar a nova planilha
print('Planilha Feita')
merged.to_excel(r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AE\01_bater_status\output\planilha_resultado.xlsx', index=False)
