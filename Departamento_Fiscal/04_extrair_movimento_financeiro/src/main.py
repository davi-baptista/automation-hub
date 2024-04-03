import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os


def clean_nan_rows(df):
    df = df.dropna(subset=['Histórico', 'Documento'], how='all')
    return df


def adjust_column_date(df):
    data_atual = None
    for index, row in df.iterrows():
        if pd.notnull(row['Data']): 
            data_atual = row['Data']
        else:
            df.at[index, 'Data'] = data_atual
    return df


def concat_columns(df):
    df['Título/Parcela'] = df['Título/Parcela'].fillna('')
    df['Concatenado'] = "VR REF " + df['Histórico'].astype(str) + " " + df['Documento'].astype(str) + " " + df['Título/Parcela'].astype(str)
    return df
    
    
def create_column_saldo(df):
    df['Crédito'] = df['Crédito'].fillna(0)
    df['Débito'] = df['Débito'].fillna(0)
    
    df['Pagamento'] = ''
    df['Recebimento'] = ''
    
    # Identificar linhas com NaN e atribuir identificadores únicos
    nan_indices = df[df['Título/Parcela'] == ''].index
    for i, index in enumerate(nan_indices):
        df.at[index, 'Título/Parcela'] = f"UniqueNaNIdentifier_{i}"

    indices_primeiros_elementos = []
    for titulo in df['Título/Parcela'].unique():
        grupo = df[df['Título/Parcela'] == titulo]
        
        if grupo['Crédito'].sum() == 0:
            # Para todos os débitos com o mesmo título, tratar individualmente
            for idx, row in grupo.iterrows():
                # Atualiza 'Pagamento' com o valor de 'Débito' para cada linha do grupo
                df.at[idx, 'Pagamento'] = row['Débito']
                
            indices_primeiros_elementos.extend(grupo.index.tolist())
        else:
            saldo_grupo = grupo['Crédito'] - grupo['Débito']
            saldo_total_grupo = saldo_grupo.sum()
            
            indice_primeiro_elemento = df[df['Título/Parcela'] == titulo].index[0]
            if saldo_total_grupo >= 0:
                df.at[indice_primeiro_elemento, 'Recebimento'] = saldo_total_grupo
            else:
                df.at[indice_primeiro_elemento, 'Pagamento'] = abs(saldo_total_grupo)
                
            indices_primeiros_elementos.append(indice_primeiro_elemento)
    
    # Filtra o DataFrame para manter apenas os registros indicados
    df = df.loc[indices_primeiros_elementos].copy()
    
    for index in nan_indices:
        df.at[index, 'Título/Parcela'] = ''
    return df

    
def create_df_mov_financeiro(df):
    df_mov_financeiro = pd.DataFrame()
    df_mov_financeiro['DIA'] = df['Data']
    df_mov_financeiro['PAGAMENTO'] = df['Pagamento']
    df_mov_financeiro['RECEBIMENTO'] = df['Recebimento']
    df_mov_financeiro['HISTORICO'] = df['Concatenado']
    df_mov_financeiro['CONTA BANCO'] = '11528'
    return df_mov_financeiro


def save_excel_with_formatting(df, filename):
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['Sheet1']
    money_format = workbook.add_format({'num_format': '#,##0.00'})  # Formato de moeda com 2 casas decimais

    (max_row, max_col) = df.shape
    column_settings = [{'header': column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 6'})
    
    # Aplicando formatação de duas casas decimais para colunas de pagamento e recebimento
    worksheet.set_column('B:B', 18, money_format)  # Coluna de Pagamento
    worksheet.set_column('C:C', 18, money_format)  # Coluna de Recebimento

    # Ajustando o espaçamento das demais colunas
    worksheet.set_column('A:A', 15)  # Coluna DIA
    worksheet.set_column('D:D', 30)  # Coluna HISTORICO
    worksheet.set_column('E:E', 15)  # Coluna CONTA BANCO

    # Fechar o writer e salvar o arquivo Excel
    writer.close()


def selecionar_diretorio():
    diretorio = filedialog.askopenfilename()
    caminho_entry.delete(0, tk.END)
    caminho_entry.insert(0, diretorio)


def iniciar_processamento():
    diretorio = caminho_entry.get()
    
    if diretorio:
        try:
            log_text.insert(tk.END, "Iniciando processamento do extrato...\n", 'info')
            app.update()
            
            df = pd.read_excel(diretorio, header=8)
            log_text.insert(tk.END, "Extrato lido com sucesso, aplicando tratamento nos dados...\n", 'info')
            app.update()
            
            df_clean = clean_nan_rows(df)
            df_date = adjust_column_date(df_clean)
            df_concat = concat_columns(df_date)
            df_saldo = create_column_saldo(df_concat)
            df_mov_financeiro = create_df_mov_financeiro(df_saldo) 
            log_text.insert(tk.END, "Dados tratados, salvando em excel...\n", 'info')
            app.update()
            
            exit_path = os.path.dirname(diretorio) + '/mov_financeiro.xlsx'
            save_excel_with_formatting(df_mov_financeiro, exit_path)
            
            log_text.insert(tk.END, f"Arquivo salvo no caminho {exit_path}...\n\n", 'info')
            app.update()
            messagebox.showinfo("Script Finalizado", "O script foi concluído com sucesso!")
            
        except Exception as e:
            log_text.insert(tk.END, f"Falha ao rodar o script. Verifique seus arquivos de entrada.\n\n", 'info')
            app.update()
            messagebox.showinfo("Erro!", "Falha ao rodar o script. Verifique seus arquivos de entrada.")
            
        
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