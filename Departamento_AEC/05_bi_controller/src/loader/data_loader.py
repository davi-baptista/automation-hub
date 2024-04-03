# data_loader.py
import pandas as pd
import streamlit as st

@st.cache_data
def carregar_dados(status=1):
    df = pd.read_excel(r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\05_bi_controller\data\empresas.xlsx', dtype={'CNPJ': str, 'NOME': str, 'GRUPO': str})
    
    def formatar_cnpj(cnpj):
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"

    # Aplicando a formatação à coluna CNPJ
    df['CNPJ'] = df['CNPJ'].apply(lambda x: formatar_cnpj(x))
    
    # Limpeza e preparação dos dados
    df['NOMESTATUS'] = df['NOMESTATUS'].str.strip()
    df['MES'] = pd.to_datetime(df['MES'], format='%Y%m').dt.to_period('M')
    df['MES_NOME'] = df['MES'].dt.strftime('%B %Y')
    
    # # Identificar nomes duplicados com CNPJs diferentes
    # df['identificador'] = df.groupby(['NOME', 'CNPJ']).ngroup()
    # duplicadas = df[df.duplicated('NOME', keep=False)].sort_values(by=['NOME', 'identificador'])
    # duplicadas['indice'] = duplicadas.groupby(['NOME'])['CNPJ'].transform(lambda x: pd.factorize(x)[0] + 1)
    # df = df.merge(duplicadas[['NOME', 'CNPJ', 'MES', 'indice']], on=['NOME', 'CNPJ', 'MES'], how='left')
    
    # # Remover o índice 1 das empresas "principais"
    # df['NOME'] = df.apply(lambda x: f"{x['NOME']}" if x['indice'] == 1 else f"{x['NOME']} ({x['indice']})", axis=1)
    # df.drop(columns=['identificador', 'indice'], inplace=True)
    
    # def custom_status_aggregation(series):
    #     if 'ATIVO' in series.values:
    #         return 'ATIVO'
    #     else:
    #         return series.iloc[0]

    # aggregations = {
    #     'NOME': 'first',
    #     'NOMESTATUS': 'first',  # Usa a função customizada para 'NOMESTATUS'
    #     'GRUPO': 'first',
    #     'QTDEFUNCIONARIOS': 'sum',
    #     'QTDESOCIOS': 'sum',
    #     'ADMISSOES': 'sum',
    #     'DEMISSOES': 'sum',
    #     'REGISTROSCONTABEIS': 'sum',
    #     'NOTASENTRADAS': 'sum',
    #     'NOTASSAIDAS': 'sum',
    #     'NOTASSERVTOMADOS': 'sum',
    #     'NOTASSERVPRESTADOS': 'sum',
    #     'RECEITABRUTA': 'sum',
    #     'HONORARIOS': 'sum',
    #     'MES_NOME': 'first'
    # }

    # df = df.groupby(['NOME', 'CNPJ', 'MES'], as_index=False).agg(aggregations)
    df['chave_unica'] = df['NOME'] + ' - ' + df['CNPJ']
    
    if status == 1:
        df = df[df['NOMESTATUS'] == 'ATIVO']
    
    df = carregar_honorarios(df)
    
    return df


def carregar_honorarios(df):
    path_honorarios = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\05_bi_controller\src\excel\BASE DE DADOS DE CLIENTES ATIVOS.xlsx'
    df_honorarios = pd.read_excel(path_honorarios, dtype={'CNPJ': str, 'HONORÁRIO': float})
    
    # Fazendo o merge/join entre as planilhas baseado na coluna CNPJ
    # Isso adicionará a coluna HONORARIOS_PAGOS ao df principal
    df_final = pd.merge(df, df_honorarios[['CNPJ', 'HONORÁRIO']], on='CNPJ', how='left')
    # Preencher valores NaN com 0 se não houver honorários correspondentes
    df_final['HONORÁRIO'].fillna(0, inplace=True)
    
    df_final = df_final.rename(columns={'HONORÁRIO': 'HONORARIOS_PAGOS'})
    
    def selecionar_linha(grupo):
        if grupo['HONORARIOS_PAGOS'].max() > 0:
            # Se houver alguma linha com HONORARIOS_PAGOS > 0, mantenha a primeira dessas linhas
            return grupo[grupo['HONORARIOS_PAGOS'] > 0].head(1)
        else:
            # Caso contrário, mantenha apenas a primeira linha do grupo
            return grupo.head(1)

    df_final_cleaned = df_final.groupby(['NOME', 'CNPJ', 'MES'], as_index=False).apply(selecionar_linha).reset_index(drop=True)
    
    return df_final_cleaned