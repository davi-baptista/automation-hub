import requests
import pandas as pd


def clean_cnpj(cnpj):
    """Remove common punctuation from CNPJ: ".", "/", and "-"."""
    return cnpj.replace(".", "").replace("/", "").replace("-", "")


def fetch_cnpj_info(cnpj, count):
    """Fetches and processes CNPJ information from the API."""
    limit = 400
    if count[0] >= limit:
        raise Exception(f"Limite de {limit} chamadas à API alcançado.")
    
    cleaned_cnpj = clean_cnpj(cnpj)
    
    url = f'https://brasilapi.com.br/api/cnpj/v1/{cleaned_cnpj}'

    response = requests.get(url)
    
    if response.status_code == 200:
        count[0] += 1
        return response.json()
    else:
        return None


def capture_cnpj_data(cnpj, count):
    try:
        cnpj_data = fetch_cnpj_info(cnpj, count)
    except Exception as e:
        print(e)
        cnpj_data = None
        
    return cnpj_data


if __name__ == '__main__':
    filepath = r"C:\Users\davi.inov\Desktop\Projetos\Departamento_Contabil\12_consulta_cnpj\data\Empresas Decon.xlsx"   
    df = pd.read_excel(filepath, dtype=str)
    
    count = [0] # Usando uma lista para persistir o valor entre chamadas de função
    for column in ['Razão Social', 'Data de Início de Atividade', 'Situação Cadastral', 'Sócios']:
        if column not in df.columns:
            df[column] = None
    
    for index, row in df.iterrows():
        if pd.notna(row['Razão Social']) and pd.notna(row['Data de Início de Atividade']) and pd.notna(row['Situação Cadastral']) and pd.notna(row['Sócios']):
            continue
        cnpj = row['CNPJ']
        cnpj_data = capture_cnpj_data(cnpj, count)
        if cnpj_data:
            row['Razão Social'] = cnpj_data.get('razao_social', 'Não disponível')
            row['Data de Início de Atividade'] = cnpj_data.get('data_inicio_atividade', 'Não disponível')
            row['Situação Cadastral'] = cnpj_data.get('descricao_situacao_cadastral', 'Não disponível')
            socios = cnpj_data.get('qsa', [])
            row['Sócios'] = ', '.join([socio['nome_socio'] for socio in socios]) if socios else 'Não há sócios listados'
            df.to_excel(filepath, index=False)