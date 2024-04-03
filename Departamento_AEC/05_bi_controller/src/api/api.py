import requests
import pandas as pd

token_filepath = r'c:\Users\davi.inov\Desktop\token.txt'


def usuarios_auth():
    with open(token_filepath, 'r') as arquivo:
        token_armazenado = arquivo.read().strip()
        
    if token_armazenado != '':
        print('Token encontrado')
        return token_armazenado
    
    url = " "

    # Cabeçalho com o identificador
    headers = {
        'sub': ' ',
    }

    # Dados de credenciais no corpo da solicitação
    dados_credenciais = {
        'login': ' ',
        'password': ' '
    }
    
    # Envia a solicitação POST com os dados de credenciais e cabeçalho
    response = requests.post(url, json=dados_credenciais, headers=headers)

    # Verificar se o login foi bem-sucedido (código de status 2xx)
    if response.status_code // 100 == 2:
        print("Login bem-sucedido!")

        # Obter o token de autenticação
        token_autenticacao = response.json().get("token")
        
        with open(token_filepath, 'w') as arquivo:
            arquivo.write(token_autenticacao)
        print('token guardado')
        
        return token_autenticacao

    else:
        print(f"Falha no login. Código de status: {response.status_code}")
        print(response.text)  # Imprimir o corpo da resposta em caso de falha
        return None
    

def ferramentas_sql(token_autenticacao, filepath, month):
    print('Realizando consulta sql')
    # URL do endpoint que executa a consulta SQL
    url_consulta_sql = " "
    
    # Criar um objeto JSON com a chave "sql" e a consulta como valor
    json_data = {
        "sql": f"SELECT e.nome, e.cnpj, f.NOME GRUPO, x.MES, x.QTDEFUNCIONARIOS, x.QTDESOCIOS, x.ADMISSOES, x.DEMISSOES, X.REGISTROSCONTABEIS, x.NOTASENTRADAS, x.NOTASSAIDAS, x.NOTASSERVTOMADOS, x.NOTASSERVPRESTADOS, ES.descricao NOMESTATUS, r.RECEITATOTAL - r.DEVOLUCOESVENDAS as RECEITABRUTA FROM PRC_MOVIMENTO_PESSOAS_EMPRESA('where 1=1 and f.cnpj > 10000000000', {month}) x LEFT JOIN tabfilial e ON x.empresa = e.codigoempresa AND x.filial = e.codigo LEFT JOIN tabfilial_status ES ON ES.CODIGO=E.STATUSEMPRESA LEFT JOIN tabempresas te ON te.CODIGO = e.CODIGOEMPRESA LEFT JOIN tabgrupoassociado f ON te.CODIGOGRUPO = f.CODIGO LEFT JOIN tabcontroletributos t ON t.codigoempresa = x.empresa and t.anomes = x.mes LEFT JOIN tabcontroletributosreceitas r on t.idmaster = r.idmaster AND r.codigofilial = e.codigo"
    }

    # Adiciona o token de autenticação aos headers
    headers = {
        "Authorization": f"Bearer {token_autenticacao}" if token_autenticacao else ""
    }

    # Fazer uma solicitação POST para executar a consulta SQL
    response_sql = requests.post(url_consulta_sql, json=json_data, headers=headers)

    # Verificar se a solicitação SQL foi bem-sucedida (código de status 2xx)
    if response_sql.status_code // 100 == 2:
        print("Consulta realizada com sucesso!")
        dados_resposta = response_sql.json()

        # Criar DataFrame com os novos dados
        df = pd.DataFrame(dados_resposta)
        
        df['GRUPO'] = df['GRUPO'].str.strip()
        
        df.fillna(0, inplace=True)
        
        df.loc[df['GRUPO'].isin(['0', 0, 'NÃO SE APLICA']), 'GRUPO'] = 'SEM GRUPO'
        
        df['HONORARIOS'] = df.apply(lambda row: row['QTDEFUNCIONARIOS'] * 28 + (row['REGISTROSCONTABEIS'] + row['NOTASENTRADAS'] + row['NOTASSAIDAS'] + row['NOTASSERVTOMADOS'] + row['NOTASSERVPRESTADOS']) * 2.8, axis=1)

        # Salvar o DataFrame combinado de volta no arquivo Excel
        save_excel_with_formatting(df, filepath)
        return filepath, 'OK'
        
    else:
        print(f"Falha na solicitação SQL. Código de status: {response_sql.status_code}")
        print(response_sql.text)  # Imprimir o corpo da resposta em caso de falha
        return None, f"Falha na solicitação SQL. Código de status: {response_sql.status_code}"
        

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
        worksheet.set_column(i, i, 15)

    # Fechar o writer e salvar o arquivo Excel
    writer.close()
    

def main(month):
    filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\05_power_bi\src\excel\mes_empresas.xlsx'
    
    token = usuarios_auth()
    return_filepath, status = ferramentas_sql(token, filepath, month)
    
    return return_filepath, status


if __name__ == "__main__":
    month = '202401202412'
    main(month)