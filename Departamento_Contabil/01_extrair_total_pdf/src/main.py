import fitz  # PyMuPDF
import re  # Regular expressions
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os


def ler_pdf(filepath, valores_aninhados):
    # Separando o nome do arquivo para pegar o local da filial e o tipo do dado
    nome_arquivo = os.path.basename(filepath)
    local = nome_arquivo.split()[2].upper()
    tipo = nome_arquivo.split()[1].upper()
    
    # Abrir o arquivo PDF
    with fitz.open(filepath) as pdf:
        
        texto_total = ""
        # Extrair o texto de todas as páginas
        for numero_pagina in range(len(pdf)):
            pagina = pdf.load_page(numero_pagina)
            texto = pagina.get_text()
            texto_total += texto + "\n"  # Adiciona quebra de linha entre as páginas
        
        # Procurar quaisquer 6 palavras que precedem "Total Geral"
        # Cada palavra é seguida de espaços ou pontuações (\S+ representa qualquer sequência de caracteres não-espaço)
        padrao = r"((?:\S+\s+){7})Total Geral"
        resultado = re.search(padrao, texto_total)
        
        if resultado:
            palavras = resultado.group(1)  # Captura apenas as palavras, sem "Total Geral"
            
            # Separar e imprimir as palavras
            palavras_separadas = palavras.split()[:-1]
                
            if local not in valores_aninhados:
                valores_aninhados[local] = {}
            valores_aninhados[local][tipo] = palavras_separadas
            return valores_aninhados
        else:
            print("A frase 'Total Geral' não foi encontrada ou não há 6 palavras antes dela.")
            return
            
            
def criar_pdf(filepath, valores_aninhados):
    # Lista para armazenar todas as linhas de dados
    dados = []
    
    # Iterar sobre os valores aninhados para construir os dados
    for local, tipos in valores_aninhados.items():
        for tipo, valores in tipos.items():
            # Cada linha de dados terá a estrutura: [local, tipo, valor1, valor2, ..., valor6]
            valores_reorganizados = [valores[1], valores[0], valores[4], valores[2], valores[5], valores[3]]
            linha = [local, tipo] + valores_reorganizados
            dados.append(linha)
    
    # Criar um DataFrame com os dados
    df = pd.DataFrame(dados, columns=['Local', 'Tipo', 'Salario Provisionado', 'Salario Provisionar', 'INSS Provisionado', 'INSS Provisionar', 'FGTS Provisionado', 'FGTS Provisionar'])
    
    # Salvar o DataFrame em um arquivo Excel
    caminho_completo = os.path.join(filepath, 'dados_saida.xlsx')
    save_excel_with_formatting(df, caminho_completo)
    print(f"Arquivo Excel salvo
          em: {caminho_completo}")
        
        
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
        # Aqui você pode ajustar a largura como preferir
        worksheet.set_column(i, i, 20)

    # Fechar o writer e salvar o arquivo Excel
    writer.close()
    
    
def selecionar_diretorio():
    diretorio = filedialog.askdirectory()
    caminho_entry.delete(0, tk.END)
    caminho_entry.insert(0, diretorio)
    
    
def iniciar_processamento():
    diretorio = caminho_entry.get()

    if diretorio:
        log_text.insert(tk.END, "Iniciando processamento de pdf's...\n", 'info')
        app.update()

        # Lista todos os arquivos no diretório
        arquivos = os.listdir(diretorio)

        # Filtra apenas os arquivos que são PDFs
        filepath = [os.path.join(diretorio, arquivo) for arquivo in arquivos if arquivo.lower().endswith('.pdf')]

        valores_aninhados = {}
        for file in filepath:
            valores_aninhados = ler_pdf(file, valores_aninhados)

        criar_pdf(diretorio, valores_aninhados)
        for local, tipos in valores_aninhados.items():
            print(f"Local: {local}")
            for tipo, valores in tipos.items():
                print(f"  Tipo: {tipo}")
                print("    Valores:")
                for valor in valores:
                    print(f"      {valor}")
                    
        log_text.insert(tk.END, "Processamento finalizado.\n", 'info')
        log_text.insert(tk.END, f"Arquivo em excel salvo na pasta {diretorio}.\n\n", 'info')
        app.update()
        messagebox.showinfo("Programa Concluído", "O programa terminou de compilar todos os pdf's e retornou o arquivo desejado!")
            
            
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