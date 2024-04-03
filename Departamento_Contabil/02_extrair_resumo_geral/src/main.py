import pandas as pd
from datetime import datetime, timedelta
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext


def tratar_valor_celula_acima(valor, valores_corretos):
    for valor_correto in valores_corretos:
        if valor_correto.lower() in valor.lower():
            print(valor)
            valor = valor_correto
            
    return valor


def verificar_fim(df, sub_index, coluna_procura):
    if pd.isnull(df.iloc[sub_index][coluna_procura]):
        coluna = sub_index
        return coluna, True

    primera_palavra = df.iloc[sub_index, coluna_procura]  # Pegando o valor
    primeira_letra = str(primera_palavra)[0]  # Convertendo o valor para string

    # Agora, você pode verificar o primeiro caractere da string
    if primeira_letra not in '0123456789':
        coluna = sub_index
        return coluna, True
    return None, False


def adicionar_ao_df(nome, valor, temp_dfs):
    if 'nan' not in nome and 'nan' not in str(valor):
        valor_final = f'{valor:.2f}'
        valor_final = str(valor_final).zfill(12)
        valor_final = valor_final.replace('.', ',')
        temp_df = pd.DataFrame({'Coluna1': [nome], 'Valor': [valor_final]})
        temp_dfs.append(temp_df)
        return temp_dfs
    
    return temp_dfs


def pegar_mes_ano_anterior():
    # Obtendo a data atual
    hoje = datetime.now()
    
    # Calculando o primeiro dia do mês atual
    primeiro_dia_mes_atual = hoje.replace(day=1)
    
    # Obtendo o último dia do mês anterior subtraindo um dia do primeiro dia do mês atual
    ultimo_dia_mes_anterior = primeiro_dia_mes_atual - timedelta(days=1)
    
    # Formatando o mês e o ano do mês anterior
    mes_ano_anterior = ultimo_dia_mes_anterior.strftime('%m/%Y')
    
    return mes_ano_anterior


def extrair_dados(excel_path):
    df = pd.read_excel(excel_path, header=None)
    
    temp_dfs = []  # Lista para coletar os DataFrames temporários
    
    # Inicializando um DataFrame para consolidar tanto proventos quanto descontos
    consolidado_df = pd.DataFrame(columns=['Coluna1', 'Valor'])
    
    # Atualizar para referenciar colunas pelo índice gerado automaticamente
    coluna_procura = 0  # Coluna onde os termos "Folha de Pagamento", etc., são buscados

    # Identificar as células-alvo e extrair os dados, incluindo as novas colunas
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
            # Pegar liquido a receber
            linha_liquido = df.iloc[limite + 1]
            coluna_liquido_a_receber = linha_liquido[linha_liquido == 'Líquido a Receber:'].index
            
            if len(coluna_liquido_a_receber) > 0:
                for coluna in range(coluna_liquido_a_receber[0] + 1, len(linha_liquido)):
                    valor = linha_liquido[coluna]
                    if pd.notna(valor):
                        nome_liquido = f'Líquido a Receber:{valor_celula_acima}'
                        temp_dfs = adicionar_ao_df(nome_liquido, valor, temp_dfs)
                        
            # Pegar FGTS
            linha_fgts = df.iloc[limite + 3]
            
            indices_fgts = [idx for idx, coluna in enumerate(linha_fgts) if 'FGTS' in str(coluna)]
            
            if len(indices_fgts) == 1:
                linha_headers = df.iloc[limite + 4]
                coluna_valor_fgts = None
                
                for coluna in range(indices_fgts[0], len(linha_headers)):
                    if linha_headers[coluna] == 'Valor': # Pegar o valor embaixo de FGTS
                        coluna_valor_fgts = coluna
                
                if coluna_valor_fgts != None:
                    for coluna in range(0, len(linha_headers)):
                        if linha_headers[coluna] == 'Competência':
                            indice = 5
                            nao_repetir = False
                            coluna_correta = None
                            valor = 0
                            mes_ano_anterior = pegar_mes_ano_anterior()
                            
                            while True:
                                linha_valores = df.iloc[limite + indice]
                                
                                if not nao_repetir:
                                    for coluna_valores in range(0, len(linha_valores)):
                                        if str(linha_valores[coluna_valores])[0] in '0123456789':
                                            coluna_correta = coluna_valores
                                            nao_repetir = True
                                            print(f'coluna correta {coluna_correta}')
                                            print(f'linha {linha_valores[coluna_correta]}')
                                            print(mes_ano_anterior)
                                            break
                                    
                                if coluna_correta != None:
                                    if linha_valores[coluna_correta] == mes_ano_anterior:
                                        print(f'{linha_valores[coluna_correta]}')
                                        print(coluna_valor_fgts)
                                        valor += linha_valores[coluna_valor_fgts]
                                    elif linha_valores[coluna_correta][0] not in '0123456789':
                                        break
                                indice += 1
                            break
                    print(valor)
                    input()
                    nome = linha_fgts[indices_fgts[0]]
                    nome_fgts = f'{nome}:{valor_celula_acima}'
                    
                    temp_dfs = adicionar_ao_df(nome_fgts, valor, temp_dfs)
                    
                else:
                    print('Erro ao achar valor do FGTS. Favor conferir.')
                        
                    
            elif len(indices_fgts) > 1:
                for i in indices_fgts: # Trocar pelos valores de indices_fgts
                    fgts = linha_fgts[i]
                    if pd.notna(fgts):
                        match = re.search(r'^FGTS: \d+', str(fgts))
                        if match:
                            nome = fgts[0:5]
                            nome_fgts = f'{nome}{valor_celula_acima}'
                            
                            valor = fgts[6:]
                            valor = float(valor.replace('.', '').replace(',', '.'))
                            
                            valor_fgts = f'{valor:.2f}'
                            valor_fgts = str(valor_fgts).zfill(12)
                            valor_fgts = valor_fgts.replace('.', ',')
                            
                            if 'nan' not in nome_fgts and 'nan' not in valor_fgts:
                                temp_df = pd.DataFrame({'Coluna1': [nome_fgts], 'Valor': [valor_fgts]})
                                temp_dfs.append(temp_df)
                            break
                
            
    if temp_dfs:  # Verificar se a lista não está vazia
        consolidado_df = pd.concat([consolidado_df] + temp_dfs, ignore_index=True)
    
    return consolidado_df


def buscar_correspondencias(df, excel_path):
    # Carregar as folhas da outra planilha
    base_prov = pd.read_excel(excel_path, sheet_name='BASE_PROV', dtype=str)
    base_desc = pd.read_excel(excel_path, sheet_name='BASE_DESC', dtype=str)
    base = pd.read_excel(excel_path, sheet_name='BASE', dtype=str)
    
    base_final = pd.concat([base_prov, base_desc, base], ignore_index=True)
    
    # Realizar o merge com how='outer' para incluir todas as linhas de ambos os DataFrames
    df_merged_outer = pd.merge(df, base_final, on='Coluna1', how='outer', indicator=True)

    # Filtrar as linhas que só existem em df (não bateram com base_final) e remover a coluna desnecessaria
    nao_bateu_df = df_merged_outer[df_merged_outer['_merge'] == 'left_only']
    nao_bateu_df = nao_bateu_df.drop(columns=['_merge'])
    
    # Filtrar as linhas que bateram nos dois DataFrames e remover a coluna desnecessaria
    bateu_nos_dois = df_merged_outer[df_merged_outer['_merge'] == 'both']
    bateu_nos_dois = bateu_nos_dois.drop(columns=['_merge'])
    
    return bateu_nos_dois, nao_bateu_df


def tirar_colunas(df, colunas):
    # Removendo colunas, ignorando erros se algumas colunas não forem encontradas
    df = df.drop(colunas, axis=1, errors='ignore')
    return df


def adicionar_mes_ano_anterior(df, coluna='Historico'):
    mes_ano_anterior = pegar_mes_ano_anterior()
    # Adicionando ' <MESANTERIOR/ANO>' ao final de cada valor na coluna especificada
    df[coluna] = df[coluna].astype(str) + ' <' + mes_ano_anterior + '>'
    
    return df


def adicionar_coluna_ultimo_dia_util(df):
    # Obtendo a data atual
    hoje = datetime.now()
    
    # Calculando o primeiro dia do mês atual
    primeiro_dia_mes_atual = hoje.replace(day=1)
    
    # Usando BMonthEnd para encontrar o último dia útil do mês anterior
    ultimo_dia_util = pd.offsets.BMonthEnd(1).rollback(primeiro_dia_mes_atual)
    
    # Formatando a data para dia/mês/ano
    data_formatada = ultimo_dia_util.strftime('%d/%m/%Y')
    
    # Adicionando a data formatada como uma nova coluna no início do DataFrame
    df.insert(0, 'UltimoDiaUtilMesAnterior', data_formatada)
    
    return df


def reorganizar_colunas(df):
    colunas_reordenadas = ['UltimoDiaUtilMesAnterior', 'Deb.', 'Cred.', 'Valor', 'Historico']

    # Aplicando a nova ordem ao DataFrame
    df = df[colunas_reordenadas]
    return df


def limpeza_de_dados(df):
    # Tirando colunas indesejadas
    df = tirar_colunas(df, ['Coluna1', 'Nome', 'Tipo'])
    df = adicionar_mes_ano_anterior(df, 'Historico')
    df = adicionar_coluna_ultimo_dia_util(df)
    df = reorganizar_colunas(df)
    return df


def selecionar_diretorio():
    diretorio = filedialog.askdirectory()
    caminho_entry.delete(0, tk.END)
    caminho_entry.insert(0, diretorio)


def iniciar_processamento():
    diretorio = caminho_entry.get()

    if diretorio:
        log_text.insert(tk.END, "Iniciando processamento de arquivos...\n", 'info')
        app.update()
        arquivos = os.listdir(diretorio)

        # Filtra apenas os arquivos que são PDFs
        filepath = [os.path.join(diretorio, arquivo) for arquivo in arquivos if arquivo.lower().endswith('.xls')]
        
        dfs_nao_bateu = []
        for file in filepath:
            nome_arquivo_com_extensao = os.path.basename(file)
            nome_arquivo, _ = os.path.splitext(nome_arquivo_com_extensao)
            
            if 'resumo geral do mêsperíodo' in nome_arquivo.lower():
                nomes = re.split(r'[\s-]+', nome_arquivo)
                
                mes_ano = f'{nomes[-1]}-{nomes[-2]}'
                nome_empresa = nomes[-3]
                
                df = extrair_dados(file)
                df_mesclado, df_nao_bateu = buscar_correspondencias(df, r"X:\INOV - RPA's\Departamento-Contabil\Script - ExtrairProventos\BASE\BASE.xlsx")
                
                if not df_nao_bateu.empty:
                    df_nao_bateu = df_nao_bateu.drop(['Valor'] ,axis=1, errors='ignore')
                    dfs_nao_bateu.append(df_nao_bateu)
                
                # Fazendo limpeza de dados no df
                df_limpo = limpeza_de_dados(df_mesclado)
                
                pasta = os.path.dirname(file)
                
                caminho_saida = pasta
                caminho_saida += f'/{nome_empresa}-{mes_ano}.txt'
                df_limpo.to_csv(caminho_saida, sep='\t', index=False, header=False)
                
                caminho_saida_excel = pasta + f'/{nome_empresa}-{mes_ano}.xlsx'
                df_limpo.to_excel(caminho_saida_excel, index=False)
                
                caminho_erros = pasta
                caminho_erros += f'/NAO_ENCONTRADOS-{mes_ano}.txt'
                df_nao_bateu.to_csv(caminho_erros, sep='\t', index=False, header=False)
                
                log_text.insert(tk.END, f"{nome_empresa} foi processada com sucesso.\n", 'info')
                log_text.insert(tk.END, f"Caminho adicionado na pasta {caminho_saida}.\n\n", 'info')
                app.update()
            
        if dfs_nao_bateu:
            log_text.insert(tk.END, f"Foi encontrado linhas que ainda não constam na base de dados. Confira na pasta {caminho_erros}.\n\n", 'info')
            
        messagebox.showinfo("Programa Concluído", "O programa terminou de compilar todos os arquivos que encontrou!")


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