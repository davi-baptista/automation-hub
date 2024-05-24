import pdfplumber
import re
import pandas as pd
import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
import os

# Função para extrair nome, PIS e valor
def extract_info(lines, index):
    nome = ''
    pis = ''
    valor = '0'
    
    if index < len(lines):
        next_line = lines[index]
        match = re.search(r'(\D+?)([\d.-]+)', next_line)
        if match:
            nome = match.group(1).strip()
            pis = match.group(2).strip()
            
            if index + 1 < len(lines):
                next_line = lines[index + 1]
                if next_line[0].isdigit():
                    value_matches = re.findall(r'[\d.,]+', next_line)
                    if value_matches:
                        valor = value_matches[4].strip()    
    
    return nome, pis, valor

def extract_company_and_registration(lines):
    company_name = ''
    registration_number = ''
    for line in lines:
        match = re.match(r'^EMPRESA:(.*?)\s+INSCRIÇÃO:\s+([\d./-]+)', line)
        if match:
            company_name = match.group(1).strip()
            registration_number = match.group(2).strip()
            return company_name, registration_number
    return company_name, registration_number

def extract_comp_value(lines):
    for line in lines:
        match = re.match(r'^COMP:(\d{2}/\d{4})', line)
        if match:
            return match.group(1)
    return ''

# Função para criar o Excel com as colunas NOME, PIS e VALOR
def save_to_excel(data, filename):
    df = pd.DataFrame(data, columns=['NOME', 'PIS', 'VALOR'])
    df.to_excel(filename, index=False)

    # Estilizar como tabela
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # Definir o intervalo da tabela
    tab = Table(displayName="Tabela1", ref=f"A1:C{len(data) + 1}")

    # Adicionar estilo à tabela
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    tab.tableStyleInfo = style
    ws.add_table(tab)

    column_widths = {'A': 40, 'B': 30, 'C': 20}
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width

    wb.save(filename)

# Abrir o arquivo PDF e extrair informações
def main(filepath):
    try:
        data = []
        verificador = False
        company_name, registration_number, comp_value = 'Error', 'Error', 'Error'
        with pdfplumber.open(filepath) as pdf:
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                
                lines = text.splitlines()
                if verificador == False:
                    comp_value = extract_comp_value(lines)
                    company_name, registration_number = extract_company_and_registration(lines)
                    comp_value = comp_value.replace('/', '-')
                    verificador = True

                for i, line in enumerate(lines):    
                    if line == 'BASE CÁL PREV SOCIAL':
                        for j in range(i + 1, len(lines)):
                            if lines[j][0].isdigit():
                                continue
                            nome, pis, valor = extract_info(lines, j)
                            if nome and pis and valor:
                                data.append([nome, pis, valor])
                        break

        # Salvar os dados no arquivo Excel
        output_path = f'{os.path.dirname(filepath)}/{company_name} {comp_value}.xlsx'
        save_to_excel(data, output_path)
        return 'Dados processados.'
    
    except Exception as e:
        return f'Erro no processamento dos dados: {e}'