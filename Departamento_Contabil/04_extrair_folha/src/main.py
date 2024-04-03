import pandas as pd
from dateutil.relativedelta import relativedelta
from datetime import datetime
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import fitz


def tirar_colunas(df, colunas):
    # Removendo colunas, ignorando erros se algumas colunas não forem encontradas
    df = df.drop(colunas, axis=1, errors='ignore')
    return df
    
    
def adicionar_mes_ano_anterior(df, mes_ano, coluna='Historico'):
    # Adicionando ' <MESANTERIOR/ANO>' ao final de cada valor na coluna especificada
    df[coluna] = df[coluna].astype(str) + ' <' + mes_ano + '>'
    
    return df
    
    
def adicionar_coluna_ultimo_dia_util(df, mes_ano):
    # Converter a string 'mes_ano' para um objeto datetime
    # É necessário especificar um dia para a conversão; vamos usar o dia 1 por padrão
    primeiro_dia_mes_atual = datetime.strptime(mes_ano, "%m/%Y").replace(day=1)
    
    primeiro_dia_proximo_mes = primeiro_dia_mes_atual + relativedelta(months=1)
    
    # Usando BMonthEnd para encontrar o último dia útil do mês anterior
    ultimo_dia_util = pd.offsets.BMonthEnd(1).rollback(primeiro_dia_proximo_mes)
    
    # Formatando a data para dia/mês/ano
    data_formatada = ultimo_dia_util.strftime('%d/%m/%Y')
    
    # Adicionando a data formatada como uma nova coluna no início do DataFrame
    df.insert(0, 'UltimoDiaUtilMesAnterior', data_formatada)
    
    return df


def reorganizar_colunas(df):
    colunas_reordenadas = ['UltimoDiaUtilMesAnterior', 'Deb.', 'Cred.', 'Valor', 'Historico', 'Empresa-Mes-Ano']
        
    # Aplicando a nova ordem ao DataFrame
    df = df[colunas_reordenadas]
    return df


def padronizar_11_digitos_valor(df):
    # Definindo uma função auxiliar para aplicar as transformações a cada valor
    def transformar_valor(valor):
        valor_str = str(valor).replace(',', '.')
        valor_limpo = re.sub(r'[^\d.-]', '', valor_str)
        valor_float = float(valor_limpo)

        valor_float = f'{valor_float:.2f}'
        valor_str = str(valor_float)
        
        # Substituir vírgula por ponto
        valor_str = valor_str.replace(',', '').replace('.', ',')
        
        valor_limpo = re.sub(r'[^\d,\.]', '', valor_str)
        
        valor_final = valor_limpo.zfill(12)
        
        return valor_final
    
    # Aplicando a função auxiliar a toda a coluna "valores"
    df['Valor'] = df['Valor'].apply(transformar_valor)
    
    return df


def padronizar_6_digitos(df):
    # Definindo uma função auxiliar para aplicar zfill nas colunas 'Deb.' e 'Cred.'
    def aplicar_zfill(valor):
        if pd.isna(valor) or valor in ['', ' ']:
            return valor
        valor_final = str(valor).zfill(6)
        
        return valor_final
    
    # Aplicando zfill às colunas 'Deb.' e 'Cred.' com 7 dígitos
    df['Deb.'] = df['Deb.'].apply(aplicar_zfill)
    df['Cred.'] = df['Cred.'].apply(aplicar_zfill)

    return df
    
    
def limpeza_de_dados(df, mes_ano):
    # Tirando colunas indesejadas
    df = tirar_colunas(df, ['Coluna1', 'Nome', 'Tipo'])
    df = adicionar_mes_ano_anterior(df, mes_ano, 'Historico')
    df = adicionar_coluna_ultimo_dia_util(df, mes_ano)
    df = reorganizar_colunas(df)
    df = padronizar_11_digitos_valor(df)
    df = padronizar_6_digitos(df)
    return df


def buscar_correspondencias_no_banco(df, chave, tipo):
    excel_path = r"X:\INOV - RPA's\Departamento-Contabil\Script - GerarFolha\BASE\BASE.xlsx"
    base_prov = pd.read_excel(excel_path, sheet_name='BASE_PROV', dtype=str)
    base_desc = pd.read_excel(excel_path, sheet_name='BASE_DESC', dtype=str)
    base = pd.read_excel(excel_path, sheet_name='BASE', dtype=str)

    # Concatenando todas as bases para o mesmo df
    base_final = pd.concat([base_prov, base_desc, base], ignore_index=True)
    
    # Converter 'Coluna1' para minúsculas para comparação insensível a maiúsculas
    df['Coluna1_lower'] = df['Coluna1'].str.lower()
    base_final['Coluna1_lower'] = base_final['Coluna1'].str.lower()
    
    df_merged_outer = pd.merge(df, base_final, on='Coluna1_lower', how='outer', indicator=True)
    nao_bateu_df = df_merged_outer[df_merged_outer['_merge'] == 'left_only'].drop(columns=['_merge', 'Coluna1_lower'])
    bateu_nos_dois = df_merged_outer[df_merged_outer['_merge'] == 'both'].drop(columns=['_merge', 'Coluna1_lower'])
    
    nao_bateu_df['Empresa-Mes-Ano'] = f'{tipo}:{chave}'
    bateu_nos_dois['Empresa-Mes-Ano'] = f'{tipo}:{chave}'
    return bateu_nos_dois, nao_bateu_df


def ler_pdf(file, tipo):
    linhas_df = []
    with fitz.open(file) as pdf:
        texto_total = ""
        for numero_pagina in range(len(pdf)):
            pagina = pdf.load_page(numero_pagina)
            texto = pagina.get_text()
            texto_total += f"{texto}\n"

        padrao = r"((?:\S+\s+){7})Total Geral"
        resultado = re.search(padrao, texto_total)

        if resultado:
            palavras = resultado.group(1).split()[:-1]
            for palavra in palavras:
                linhas_df.append({'Coluna1': tipo, 'Valor': palavra})

            df = pd.DataFrame(linhas_df)
            return df
        else:
            print("A frase 'Total Geral' não foi encontrada.")
            return None


def trocar_linhas(df):
    df_copia = df.copy()
    trocas = [(1, 0), (2, 4), (3, 4), (4, 5)]
    for indice1, indice2 in trocas:
        df_copia.iloc[[indice1, indice2]] = df_copia.iloc[[indice2, indice1]].values

    return df_copia


def limpar_valor_e_converter_float(valor):
    valor_str = str(valor).replace('.', '').replace(',', '.')
    valor_limpo = re.sub(r'[^\d.-]', '', valor_str)
    return float(valor_limpo)


def somar_valor_provisoes(df_valor):
    novas_linhas = []
    for i in range(0, len(df_valor), 2):
        if i + 1 < len(df_valor):
            valor1 = limpar_valor_e_converter_float(df_valor.iloc[i]['Valor'])
            valor2 = limpar_valor_e_converter_float(df_valor.iloc[i + 1]['Valor'])
            soma = valor1 + valor2
            novas_linhas.append({'Coluna1': df_valor.iloc[i]['Coluna1'], 'Valor': soma})

    df_resultado = pd.DataFrame(novas_linhas)
    df_resultado['Valor'] = df_resultado['Valor'].apply(lambda x: f"{x:.2f}".replace('.', ','))
    return df_resultado


def adicionar_tipo_coluna1(df):
    textos = ['SALARIO', 'INSS', 'FGTS']
    if len(textos) != len(df):
        raise ValueError("Quantidade de textos e linhas do DataFrame não correspondem.")
    for i, texto in enumerate(textos):
        df.at[i, 'Coluna1'] = f"{texto}:{df.at[i, 'Coluna1']}"

    return df


def adicionar_colunas_nome_mes(df, nome, mes):
    # Adicionando a nova coluna com o valor_nome. A nova coluna será 'Nome'
    df['Nome'] = nome
    df['Mes'] = mes
    
    # Adicionando ' <MESANTERIOR/ANO>' ao final de cada valor na coluna especificada
    df['Coluna1'] = 'REVERSAO' + df['Coluna1'].astype(str)

    # Movendo a coluna 'Nome' para ser a primeira coluna do DataFrame
    colunas = ['Nome'] + ['Mes'] + [col for col in df.columns if col != 'Nome' and col != 'Mes']
    df = df[colunas]
    
    return df


def adicionar_ao_excel_base_prov(excel_path, df_existente, df_novo):
    # Concatenar df_existente e df_novo
    df_combinado = pd.concat([df_existente, df_novo], ignore_index=True)

    chaves = ['Nome', 'Mes', 'Coluna1']

    # Remover duplicatas, mantendo a última ocorrência
    df_atualizado = df_combinado.drop_duplicates(subset=chaves, keep='last').reset_index(drop=True)
    df_atualizado.to_excel(excel_path, index=False)
    

def extrair_provisao(filepath, excel_path):
    dicionario_provisoes = {}
    dfs_nao_bateu = pd.DataFrame()
    df_provisoes = pd.DataFrame()
    
    df_prov_base = pd.read_excel(excel_path)
    chaves = []

    for file in filepath:
        nome_arquivo_com_extensao = os.path.basename(file)
        nome_arquivo, _ = os.path.splitext(nome_arquivo_com_extensao)
        
        palavras = re.split(r'[\s-]+', nome_arquivo)
        
        tipo = palavras[1].upper()
        mes_ano = f'{palavras[-1]}-{palavras[-2]}'
        mes_ano_barras = f'{palavras[-2]}/{palavras[-1]}'
        nome_empresa = palavras[-3]
        
        if 'provisão' in nome_arquivo.lower():
            df_valores_provisoes = ler_pdf(file, tipo)
            chave = f'{nome_empresa}-{mes_ano}'
            chaves.append(chave)
            
            if not df_valores_provisoes.empty:
                df_valores_provisoes = trocar_linhas(df_valores_provisoes)
                df_provisoes_somado = somar_valor_provisoes(df_valores_provisoes)
                df_provisoes_final = adicionar_tipo_coluna1(df_provisoes_somado)
                df_resumo_geral, df_nao_bateu = buscar_correspondencias_no_banco(df_provisoes_final, chave, 'PROVISAO')
                df_limpo = limpeza_de_dados(df_resumo_geral, mes_ano_barras)
                
                if not df_nao_bateu.empty:
                    df_nao_bateu = df_nao_bateu.drop(['Valor'] ,axis=1, errors='ignore')
                    dfs_nao_bateu = pd.concat([dfs_nao_bateu, df_nao_bateu], ignore_index=True)
                
                if not df_limpo.empty:
                    if chave in dicionario_provisoes:
                        dicionario_provisoes[chave] = pd.concat([dicionario_provisoes[chave], df_limpo], ignore_index=True)
                    else:
                        dicionario_provisoes[chave] = df_limpo
                        
                        
                df_provisoes_final_empresa = adicionar_colunas_nome_mes(df_provisoes_final, nome_empresa, mes_ano)
                df_provisoes = pd.concat([df_provisoes, df_provisoes_final_empresa], ignore_index=True)
                if not df_provisoes.empty:
                    adicionar_ao_excel_base_prov(excel_path, df_prov_base, df_provisoes)
        
    return dicionario_provisoes, dfs_nao_bateu, chaves

    
def tratar_valor_celula_acima(valor, valores_corretos):
    """
    Substitui o valor pela correspondência correta, se encontrada.
    """
    for valor_correto in valores_corretos:
        if valor_correto.lower() in valor.lower():
            valor = valor_correto
    return valor


def verificar_fim(df, sub_index, coluna_procura):
    """
    Verifica se o dataframe chegou ao final de uma seção baseado no valor da célula.
    Retorna o índice da coluna e um booleano indicando se é o fim.
    """
    if pd.isnull(df.iloc[sub_index][coluna_procura]):
        return sub_index, True

    primeira_palavra = df.iloc[sub_index, coluna_procura]
    primeira_letra = str(primeira_palavra)[0]

    if primeira_letra not in '0123456789':
        return sub_index, True
    return None, False

    
def adicionar_ao_df(nome, valor, temp_dfs):
    """
    Adiciona um novo DataFrame temporário à lista, contendo nome e valor, se nenhum deles for 'nan'.
    """
    if 'nan' not in nome and 'nan' not in str(valor):
        temp_df = pd.DataFrame({'Coluna1': [nome], 'Valor': [valor]})
        temp_dfs.append(temp_df)
    
    return temp_dfs

    
def extrair_dados(excel_path, mes_ano):
    """
    Extrai dados de um arquivo Excel especificado, consolidando-os em um único DataFrame.
    """
    df = pd.read_excel(excel_path, header=None)
    coluna_procura = 0
    
    consolidado_df = pd.DataFrame(columns=['Coluna1', 'Valor'])

    temp_dfs = []
    for index, row in df.iterrows():
        if row[coluna_procura] == 'Proventos':
            headers = df.iloc[index]
            coluna_descontos = headers[headers == 'Descontos'].index[0]
            colunas_valor = [idx for idx, val in enumerate(headers) if val == 'Valor']
            
            valor_celula_acima = tratar_valor_celula_acima(df.iloc[index - 1, coluna_procura], ['Folha de Pagamento', 'Férias', 'Rescisão', '13º Salário'])
            
            coluna_final_proventos = 0
            coluna_final_descontos = 0
            
            
            if headers[coluna_procura] == 'Proventos':
                for sub_index in range(index + 1, len(df)):
                    coluna, status = verificar_fim(df, sub_index, coluna_procura)
                    if status:
                        coluna_final_proventos = coluna
                        break
                    
                    nome_provento = str(df.iloc[sub_index][coluna_procura]).strip()
                    nome_provento = re.sub(' +', ' ', nome_provento) + str(valor_celula_acima).strip()
                    valor_proventos = df.iloc[sub_index, colunas_valor[0]]
                    temp_dfs = adicionar_ao_df(nome_provento, valor_proventos, temp_dfs)
            
            if headers[coluna_descontos] == 'Descontos':
                for sub_index in range(index + 1, len(df)):
                    coluna, status = verificar_fim(df, sub_index, coluna_descontos)
                    if status:
                        coluna_final_descontos = coluna
                        break
                    
                    nome_desconto = str(df.iloc[sub_index][coluna_descontos]).strip()
                    nome_desconto = re.sub(' +', ' ', nome_desconto) + str(valor_celula_acima).strip()
                    valor_descontos = df.iloc[sub_index, colunas_valor[1]]
                    temp_dfs = adicionar_ao_df(nome_desconto, valor_descontos, temp_dfs)
            
            
            limite = max(coluna_final_proventos, coluna_final_descontos)
            linha_liquido = df.iloc[limite + 1]
            coluna_liquido_a_receber = linha_liquido[linha_liquido == 'Líquido a Receber:'].index
            
            if len(coluna_liquido_a_receber) > 0:
                for coluna in range(coluna_liquido_a_receber[0] + 1, len(linha_liquido)):
                    valor = linha_liquido[coluna]
                    if pd.notna(valor):
                        nome_liquido = f'Líquido a Receber:{valor_celula_acima}'
                        temp_dfs = adicionar_ao_df(nome_liquido, valor, temp_dfs)
                        
                        
            linha_fgts = df.iloc[limite + 3]
            indices_fgts = [idx for idx, coluna in enumerate(linha_fgts) if 'FGTS' in str(coluna)]
            
            if len(indices_fgts) == 1:
                linha_headers = df.iloc[limite + 4]
                coluna_valor_fgts = None
                for coluna in range(indices_fgts[0], len(linha_headers)):
                    if linha_headers[coluna] == 'Valor':
                        coluna_valor_fgts = coluna
                        
                if coluna_valor_fgts != None:
                    for coluna in range(0, len(linha_headers)):
                        if linha_headers[coluna] == 'Competência':
                            indice = 5
                            nao_repetir = False
                            coluna_correta = None
                            valor = 0
                            
                            while True:
                                linha_valores = df.iloc[limite + indice]
                                if not nao_repetir:
                                    for coluna_valores in range(0, len(linha_valores)):
                                        if str(linha_valores[coluna_valores])[0] in '0123456789':
                                            coluna_correta = coluna_valores
                                            nao_repetir = True
                                            break
                                    
                                if coluna_correta != None:
                                    if linha_valores[coluna_correta] == mes_ano:
                                        valor += linha_valores[coluna_valor_fgts]
                                    elif linha_valores[coluna_correta][0] not in '0123456789':
                                        break
                                indice += 1
                    
                    nome = linha_fgts[indices_fgts[0]]
                    nome_fgts = f'{nome}:{valor_celula_acima}'
                    
                    temp_dfs = adicionar_ao_df(nome_fgts, valor, temp_dfs)
                    
                else:
                    print('Erro ao achar valor do FGTS. Favor conferir.')
                    
            elif len(indices_fgts) > 1:
                for i in indices_fgts:
                    fgts = linha_fgts[i]
                    if pd.notna(fgts):
                        match = re.search(r'^FGTS: \d+', str(fgts))
                        if match:
                            nome = fgts[0:4]
                            nome_fgts = f'{nome}:{valor_celula_acima}'
                            
                            valor = fgts[6:]
                            valor = float(valor.replace('.', '').replace(',', '.'))
                        
                            temp_dfs = adicionar_ao_df(nome_fgts, valor, temp_dfs)
            
    if temp_dfs:
        consolidado_df = pd.concat([consolidado_df] + temp_dfs, ignore_index=True)
    
    return consolidado_df


def extrair_resumo_geral(filepath):
    dicionario_resumo_geral = {}
    dfs_nao_bateu = pd.DataFrame()
    chaves = []
    
    for file in filepath:
        nome_arquivo_com_extensao = os.path.basename(file)
        nome_arquivo, _ = os.path.splitext(nome_arquivo_com_extensao)
        
        if 'resumo geral do mêsperíodo' in nome_arquivo.lower():
            nomes = re.split(r'[\s-]+', nome_arquivo)
            
            mes_ano = f'{nomes[-1]}-{nomes[-2]}'
            mes_ano_barra = f'{nomes[-2]}/{nomes[-1]}'
            nome_empresa = nomes[-3]
            chave = f'{nome_empresa}-{mes_ano}'
            chaves.append(chave)
            
            df = extrair_dados(file, mes_ano_barra)
            df_mesclado, df_nao_bateu = buscar_correspondencias_no_banco(df, chave, 'RESUMO_GERAL')
            
            if not df_nao_bateu.empty:
                df_nao_bateu = df_nao_bateu.drop(['Valor'], axis=1, errors='ignore')
                dfs_nao_bateu = pd.concat([dfs_nao_bateu, df_nao_bateu], ignore_index=True)
            
            # Fazendo limpeza de dados no df
            df_limpo = limpeza_de_dados(df_mesclado, mes_ano_barra)
            
            if not df_limpo.empty:
                dicionario_resumo_geral[chave] = df_limpo
    
    return dicionario_resumo_geral, dfs_nao_bateu, chaves
    
    
def filtrar_por_chave(df, chave):
    palavras = chave.split('-')
    nome_empresa = palavras[0]  # Tudo exceto os últimos 7 caracteres (assumindo o formato 'Ano-Mes' como '2022-01')
    mes_ano = f'{palavras[1]}-{palavras[2]}'
    
    # Criando a máscara booleana para identificar as linhas que correspondem a ambos, nome_empresa e mes_ano
    mask = (df['Nome'] == nome_empresa) & (df['Mes'] == mes_ano)
    
    # Filtrando o DataFrame com a máscara
    df_filtrado = df[mask]
    
    return df_filtrado


def chave_mes_anterior(chaves):
    chaves_ajustadas = []
    for chave in chaves:
        palavras = chave.split('-')
        mes_ano = f'{palavras[1]}-{palavras[2]}'
        
        # Converter a string para um objeto datetime (assumindo o dia 1 para facilitar a operação)
        data = datetime.strptime(mes_ano, '%Y-%m')

        # Subtrair um mês usando relativedelta
        data_subtraida = data - relativedelta(months=1)

        # Converter de volta para o formato "ano-mes"
        nome_mes_ano_ajustado = f'{palavras[0]}-{data_subtraida.strftime('%Y-%m')}'
        chaves_ajustadas.append(nome_mes_ano_ajustado)
        
    return chaves_ajustadas


def extrair_reversao(filepath, chaves_anterior, chaves):
    df_prov_base = pd.read_excel(filepath)
    dfs_nao_bateu = pd.DataFrame()
    dicionario = {}
    
    for chave_anterior, chave in zip(chaves_anterior, chaves):
        palavras = chave.split('-')
        mes_ano_barra = f'{palavras[2]}/{palavras[1]}'
        
        df_resultado_filtro_nome_mes = filtrar_por_chave(df_prov_base, chave_anterior)
        df_prov_base_final, df_nao_bateu = buscar_correspondencias_no_banco(df_resultado_filtro_nome_mes, chave, 'REVERSAO')
        df_prov_limpo = limpeza_de_dados(df_prov_base_final, mes_ano_barra)
        
        if not df_nao_bateu.empty:
            df_nao_bateu = df_nao_bateu.drop(['Valor'], axis=1, errors='ignore')
            dfs_nao_bateu = pd.concat([dfs_nao_bateu, df_nao_bateu], ignore_index=True)
                
        if not df_prov_limpo.empty:
            if chave in dicionario:
                dicionario[chave] = pd.concat([dicionario[chave], df_prov_limpo], ignore_index=True)
            else:
                dicionario[chave] = df_prov_limpo
    
    return dicionario, dfs_nao_bateu


def concatenar_dicionarios(dicionario_resumo_geral, dicionario_provisoes):
    dicionario_concatenado = {}
    for chave in dicionario_resumo_geral:
        if chave in dicionario_provisoes:
            # Chave existe em ambos os dicionários, concatenar os DataFrames
            dicionario_concatenado[chave] = pd.concat([dicionario_resumo_geral[chave], dicionario_provisoes[chave]], ignore_index=True)
        else:
            # Chave existe apenas no dicionario_resumo_geral
            dicionario_concatenado[chave] = dicionario_resumo_geral[chave]

    # Agora, verificar as chaves exclusivas do dicionario2 e adicionar ao dicionario_concatenado
    for chave in dicionario_provisoes:
        if chave not in dicionario_concatenado:
            dicionario_concatenado[chave] = dicionario_provisoes[chave]
            
    return dicionario_concatenado
            
            
def verificar_colunas_com_nan(dicionario_dfs):
    df_com_nan = pd.DataFrame()
    
    for chave, df in dicionario_dfs.items():
        mask = df.isna().any(axis=1)
        
        # Selecionar as linhas com NaN e adicioná-las ao df_coluna_nan
        linhas_com_na = df.loc[mask]
        df_com_nan = pd.concat([df_com_nan, linhas_com_na], ignore_index=True)

        # Remover as linhas com NaN do DataFrame atual
        df_limpo = df.loc[~mask]

        dicionario_dfs[chave] = df_limpo
        
    df_com_nan.fillna('nan', inplace=True)

    return dicionario_dfs, df_com_nan


def remover_nulos(dicionario_dfs):
    # Iterar sobre o dicionário e filtrar as linhas onde 'Valor' não é '000000000,00'
    for chave, df in dicionario_dfs.items():
        dicionario_dfs[chave] = df[df['Valor'] != '000000000,00']
    
    return dicionario_dfs


def criar_arquivos_txt(diretorio, dicionario_dfs, df_nao_bateu, df_com_nan):
    for chave, df in dicionario_dfs.items():
        palavras = chave.split('-')
        
        nome_arquivo = f'FOL{palavras[2]}{palavras[1]} {palavras[0]}'
        
        df = df.drop(['Empresa-Mes-Ano'], axis=1, errors='ignore')

        caminho_completo = f'{diretorio}/{nome_arquivo}.txt'
        df.to_csv(caminho_completo, sep='\t', index=False, header=None)

        print(f'DataFrame {chave} salvo como {caminho_completo}')
        log_text.insert(tk.END, f'DataFrame {chave} salvo como {caminho_completo}\n\n', 'info')
        app.update()
        
    if not df_nao_bateu.empty:
        caminho_completo = f'{diretorio}/NAO_ENCONTRADO.txt'
        df_nao_bateu.to_csv(caminho_completo, sep='\t', index=False, header=None)

        print(f'DataFrame "NAO_ENCONTRADO" salvo como {caminho_completo}')
        log_text.insert(tk.END, f'DataFrame {chave} salvo como {caminho_completo}\n\n', 'info')
        app.update()
        
        messagebox.showinfo("ATENÇÃO", "Dados encontrados que não bateram com a base de dados")
        
    if not df_com_nan.empty:
        caminho_completo = f'{diretorio}/COLUNA_NAN.txt'
        df_com_nan.to_csv(caminho_completo, sep='\t', index=False)

        print(f'DataFrame "COLUNA_NAN" salvo como {caminho_completo}')
        log_text.insert(tk.END, f'DataFrame {chave} salvo como {caminho_completo}\n\n', 'info')
        app.update()
    
        messagebox.showinfo("ATENÇÃO", "Foi encontrado chaves que ainda não estão cadastradas no banco de dados ou celulas vazias.\nConfira no txt NAO_ENCONTRADO para mais detalhes!")


def selecionar_diretorio():
    diretorio = filedialog.askdirectory()
    caminho_entry.delete(0, tk.END)
    caminho_entry.insert(0, diretorio)


def iniciar_processamento():
    diretorio = caminho_entry.get()

    if diretorio:
        try:
            log_text.insert(tk.END, "Iniciando processamento de arquivos...\n", 'info')
            app.update()
            arquivos = os.listdir(diretorio)
            excel_path = r"C:\Users\davi.inov\Desktop\Projetos\Departamento_Contabil\04_extrair_folha\data\database\database.xlsx"

            filepath_resumo_geral = [os.path.join(diretorio, arquivo) for arquivo in arquivos if arquivo.lower().endswith('.xls')]
            filepath_provisao = [os.path.join(diretorio, arquivo) for arquivo in arquivos if arquivo.lower().endswith('.pdf')]
            
            if not filepath_resumo_geral and not filepath_provisao:
                log_text.insert(tk.END, "Nenhum arquivo encontrado.\n\n", 'info')
                app.update()
                return
            
            dicionario_resumo_geral, df_nao_bateu_resumo_geral, chaves_1 = extrair_resumo_geral(filepath_resumo_geral)
            dicionario_provisoes, df_nao_bate_provisoes, chaves_2 = extrair_provisao(filepath_provisao, excel_path)
            
            chaves_unicas = set(chaves_1 + chaves_2)
            chaves_anterior = chave_mes_anterior(chaves_unicas)
            chaves_unicas_anterior = list(chaves_anterior)
            
            dicionario_reversao, df_nao_bateu_reversao = extrair_reversao(excel_path, chaves_unicas_anterior, chaves_unicas)
            
            df_nao_bateu = pd.concat([df_nao_bateu_resumo_geral, df_nao_bate_provisoes, df_nao_bateu_reversao], ignore_index=True)
            
            # Inicializando o dicionário para armazenar os DataFrames concatenados
            dicionario_concatenado = concatenar_dicionarios(dicionario_resumo_geral, dicionario_provisoes)
            dicionario_concatenado = concatenar_dicionarios(dicionario_concatenado, dicionario_reversao)
            
            dicionario_concatenado_sem_nan, df_com_nan = verificar_colunas_com_nan(dicionario_concatenado)
            dicionario_concatenado_final = remover_nulos(dicionario_concatenado_sem_nan)
            
            nome_subpasta = "Resultados_processados"
            caminho_subpasta = os.path.join(diretorio, nome_subpasta)
            os.makedirs(caminho_subpasta, exist_ok=True)
            
            criar_arquivos_txt(caminho_subpasta, dicionario_concatenado_final, df_nao_bateu, df_com_nan)
                
            messagebox.showinfo("Programa Concluído", "O programa terminou de compilar todos os arquivos que encontrou!")
            
        except Exception as e:
            log_text.insert(tk.END, f"Erro encontrado: {e}.\n\n", 'info')
            app.update()
            
            
if __name__ == '__main__':
    app = tk.Tk()
    app.title("Selecionador de Diretório")

    frame = tk.Frame(app)
    frame.pack(padx=10, pady=10)

    caminho_entry = tk.Entry(frame, width=100)
    caminho_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    botao_buscar = tk.Button(frame, text="Buscar", command=selecionar_diretorio)
    botao_buscar.pack(side=tk.LEFT, padx=(10, 0))

    botao_enviar = tk.Button(frame, text="Enviar", command=iniciar_processamento)
    botao_enviar.pack(side=tk.LEFT, padx=(10, 0))

    log_text = scrolledtext.ScrolledText(app, height=10)
    log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    log_text.tag_configure('info', foreground='blue')
    log_text.tag_configure('error', foreground='red')
        
    app.mainloop()