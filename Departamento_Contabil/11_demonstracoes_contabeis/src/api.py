import requests
import os

token_filepath = r"X:\INOV - RPA's\Departamento-Contabil\Script - ValidadorDeEstoque\api_token\token.txt"

def users_auth():
    """
    Authenticate the user and obtain a token, or use an existing token.
    """
    try:
        with open(token_filepath, 'r') as file:
            stored_token = file.read().strip()
            
        if stored_token:
            print('Token found')
            return stored_token
    except FileNotFoundError:
        print("Token file not found. Proceeding to authenticate.")
    
    url = "https://api.athenas.online/v2/usuarios/auth"

    # Cabeçalho com o identificador
    headers = {'sub': ' '}

    # Dados de credenciais no corpo da solicitação
    credentials_data  = {
        'login': ' ',
        'password': ' '
    }
    
    # Envia a solicitação POST com os dados de credenciais e cabeçalho
    response = requests.post(url, json=credentials_data , headers=headers)

    # Verificar se o login foi bem-sucedido (código de status 2xx)
    if response.status_code // 100 == 2:
        print("Login successful!")
        authentication_token = response.json().get("token")
        
        with open(token_filepath, 'w') as arquivo:
            arquivo.write(authentication_token)
        print('token guardado')
        
        return authentication_token

    else:
        print(f"Falha no login. Código de status: {response.status_code}")
        return None
    
    
def check_excel_existence(account_plan):
    """
    Check if the Excel file for the given account plan exists.
    """
    folder_path = r"X:\INOV - RPA's\Departamento-Contabil\Script - DemonstracaoContabil\data"
    file_name = f"plano_{account_plan}.xlsx"
    full_path = os.path.join(folder_path, file_name)
    
    if os.path.isfile(full_path):
        return full_path, 'File found'
    else:
        return None, 'File not found'
    

def handle_401_error():
    """
    Handle error 401 by refreshing the authentication token.
    """
    with open(token_filepath, 'r+') as arquivo:
        arquivo.seek(0)
        arquivo.truncate()
        
    new_auth_token = users_auth()
    headers = {"Authorization": f"Bearer {new_auth_token}" if new_auth_token else ""}
    return headers
    
    
def check_account_plan(cnpj, url, headers):
    """
    Check the account plan status for the given CNPJ.
    """
    json_data = {
        "sql": f"SELECT e.nome, codigoplanocontas FROM tabempresas e JOIN tabfilial F ON F.codigoempresa = e.codigo WHERE F.CNPJ = {cnpj}"
    }
    
    response_sql = requests.post(url, json=json_data, headers=headers)

    if response_sql.status_code // 100 == 2:
        print("Account plan consultation carried out successfully!")
        dados = response_sql.json()
        print(dados)
        
        if not dados:
            print('Invalid CNPJ')
            return None, None, 'Invalid CNPJ'
        
        excel_path, existence_status = check_excel_existence(dados[0]['CODIGOPLANOCONTAS'])
        return dados[0]['NOME'], excel_path, existence_status
        
    else:
        print(f"SQL request failed. Status code: {response_sql.status_code}")
        return None, None, f"SQL request failed. Status code: {response_sql.status_code}"
    

def sql_tools(auth_token, cnpj, start_date, end_date):
    """
    Perform SQL tools operations.
    """
    print('Performing SQL query')
    
    sql_query_url  = " "
    headers = {"Authorization": f"Bearer {auth_token}" if auth_token else ""}
    
    company_name, excel_path, account_plan_status = check_account_plan(cnpj, sql_query_url, headers)
    
    if account_plan_status == "SQL request failed. Status code: 401":
        headers = handle_401_error()
        company_name, excel_path, account_plan_status = check_account_plan(cnpj, sql_query_url, headers)
    
    if account_plan_status == 'File found':
        try:
            json_data = {
                "sql": f"SELECT CODIGOCONTAMAE, CODIGOCONTACONTABIL AS CODIGO, NOME, SUM(CREDITO) AS CREDITO, SUM(DEBITO) AS DEBITO, SUM(SALDO_ANTERIOR_CREDITO) AS SALDO_ANTERIOR_CREDITO, SUM(SALDO_ANTERIOR_DEBITO) AS SALDO_ANTERIOR_DEBITO, ABS(SUM(SALDO_ANTERIOR_CREDITO - SALDO_ANTERIOR_DEBITO + CREDITO - DEBITO)) AS SALDO_FINAL FROM ( SELECT P.CODIGOCONTAMAE, S.CODIGOCONTACONTABIL, P.NOME, CAST(COALESCE(VALORCREDITO, 0) AS NUMERIC(15,2)) AS CREDITO, CAST(COALESCE(VALORDEBITO, 0) AS NUMERIC(15,2)) AS DEBITO, 0 AS SALDO_ANTERIOR_CREDITO, 0 AS SALDO_ANTERIOR_DEBITO FROM TABSALDOCONTABIL S JOIN tabfilial F ON F.codigoempresa = S.codigoempresa AND F.codigo = S.codigofilial LEFT JOIN TABEMPRESAS E ON E.CODIGO = F.codigoempresa LEFT JOIN TABPLANOCONTAS P ON P.CODIGOPLANOCONTAS = E.CODIGOPLANOCONTAS AND P.CODIGO = S.CODIGOCONTACONTABIL WHERE F.CNPJ = {cnpj} AND DATA BETWEEN '{start_date}' AND '{end_date}' AND INICIAL NOT IN (1, 2, 4, 5, 6) UNION ALL SELECT P.CODIGOCONTAMAE, S.CODIGOCONTACONTABIL, P.NOME, 0 AS CREDITO, 0 AS DEBITO, CAST(VALORCREDITO AS NUMERIC(15,2)) AS SALDO_ANTERIOR_CREDITO, CAST(VALORDEBITO AS NUMERIC(15,2)) AS SALDO_ANTERIOR_DEBITO FROM TABSALDOCONTABIL S JOIN tabfilial F ON F.codigoempresa = S.codigoempresa AND F.codigo = S.codigofilial LEFT JOIN TABEMPRESAS E ON E.CODIGO = F.codigoempresa LEFT JOIN TABPLANOCONTAS P ON P.CODIGOPLANOCONTAS = E.CODIGOPLANOCONTAS AND P.CODIGO = S.CODIGOCONTACONTABIL WHERE F.CNPJ = {cnpj} AND inicial = 1 AND DATA BETWEEN '{start_date}' AND '{end_date}' ) AS combined GROUP BY CODIGOCONTAMAE, CODIGOCONTACONTABIL, NOME HAVING ABS(SUM(SALDO_ANTERIOR_CREDITO - SALDO_ANTERIOR_DEBITO + CREDITO - DEBITO)) <> 0 ORDER BY CODIGOCONTACONTABIL;"
            }
            json_code = {
                "sql": f"SELECT DISTINCT p.CODIGOCONTAMAE, p.CODIGO, p.NOME FROM TABPLANOCONTAS p JOIN TABFILIAL f ON f.CNPJ = '{cnpj}' LEFT JOIN TABEMPRESAS e ON e.CODIGO = f.codigoempresa WHERE p.CODIGOPLANOCONTAS = e.codigoplanocontas AND ( p.tipo = 1 OR p.codigo IN ( SELECT DISTINCT s.codigocontacontabil FROM TABSALDOCONTABIL s JOIN TABPLANOCONTAS px ON s.CODIGOCONTACONTABIL = px.CODIGO WHERE s.CODIGOEMPRESA = e.CODIGO AND s.data BETWEEN '{start_date}' AND '{end_date}' AND p.codigo = s.codigocontacontabil ) ) ORDER BY p.CODIGO;"
            }

            response_data = requests.post(sql_query_url, json=json_data, headers=headers)
            json_data, status_response_data = check_response(response_data)
            
            status = 'Empty company'
            if status_response_data == 'Found company data':
                response_code = requests.post(sql_query_url, json=json_code, headers=headers)
                json_code, status_response_code = check_response(response_code)
                
                if status_response_code == 'Found company data':
                    mescled_data, status = mesclar_dados(json_data, json_code)
            
            return company_name, excel_path, mescled_data, status
        
        except Exception as e:
            print(e)
            return None, None, None, e
        
    return None, None, None, account_plan_status
    
    
def calcular_saldo_recursivo(codigo, json_data, json_code, saldos):
    if saldos[codigo] is not None:
        return saldos[codigo]

    # Obter todos os filhos deste código
    filhos = [data['CODIGO'] for data in json_code if data.get('CODIGOCONTAMAE') == codigo]
    soma_filhos = sum(calcular_saldo_recursivo(filho, json_data, json_code, saldos) for filho in filhos)
    
    # Encontrar o elemento no json_code para verificar o nome
    nome_pai = next((item['NOME'] for item in json_code if item['CODIGO'] == codigo), "")
    if '(-)' in nome_pai or '( - )' in nome_pai:
        soma_filhos *= -1
    
    saldos[codigo] = soma_filhos
    return soma_filhos

    
def mesclar_dados(json_data, json_code):
    # Inicializar saldos com None para todos os códigos em json_code
    saldos = {code['CODIGO']: None for code in json_code}
    
    # Preencher os saldos iniciais usando json_data
    for data in json_data:
        if data['CODIGO'] in saldos:
            saldo_final = float(data['SALDO_FINAL'])
            if '(-)' in data['NOME'] or '( - )' in data ['NOME']:
                saldo_final *= -1
            saldo_formatado = round(saldo_final, 2)
            saldos[data['CODIGO']] = saldo_formatado

    # Processar saldos para todos os códigos, especialmente para aqueles sem valores definidos
    for code in json_code:
        if saldos[code['CODIGO']] is None:
            calcular_saldo_recursivo(code['CODIGO'], json_data, json_code, saldos)

    # Atualizar json_code com os saldos calculados
    for code in json_code:
        if saldos[code['CODIGO']] is not None:
            code['SALDO_FINAL'] = round(saldos[code['CODIGO']], 2)
        else:
            code['SALDO_FINAL'] = 0.00

    for item in json_code:
        print("Código:", item['CODIGO'], "Saldo Final:", item['SALDO_FINAL'])
        
    return json_code, 'Found company data'
    
    
def check_response(response):
    if response.status_code // 100 == 2:
        print("Company data query was successful!")
        json_data = response.json()
        
        status_return = 'Found company data'
        if not json_data:
            raise Exception("Empty company")
        
        return json_data, status_return
    else:
        print(f"SQL request failed. Status code: {response.status_code}")
        raise Exception(f"SQL request failed. Status code: {response.status_code}")


def api_main(cnpj, start_date, end_date):
    token = users_auth()
    
    company_name, excel_path, data, status = sql_tools(token, cnpj, start_date, end_date)
    return company_name, excel_path, data, status

if __name__ == "__main__":
    api_main('13311185000105', '01/01/2023', '12/31/2023')