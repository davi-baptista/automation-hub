"""
# Projeto de Automação de Validação de Estoque

Este projeto contém scripts automatizados para autenticação de usuários, verificação e atualização de arquivos Excel com base nos dados obtidos da API da Athenas.

## Requisitos

- Python 3.6+
- requests
- openpyxl

## Instalação

1. Clone o repositório:
    ```bash
   git clone https://github.com/seu-usuario/seu-repositorio.git
    ```
2. Navegue até o diretório do projeto:
    ```bash
   cd seu-repositorio
    ```
3. Instale as dependências executando o comando abaixo:
    ```bash
   pip install requests openpyxl
    ```

## Arquivos

### api.py

Este módulo lida com a autenticação de usuários na API da Athenas e verifica a existência de arquivos Excel para planos de contas específicos.

#### Funções

- `users_auth()`: Autentica o usuário e obtém um token, ou usa um token existente.
- `check_excel_existence(account_plan)`: Verifica se o arquivo Excel para o plano de contas fornecido existe.
- `api_main(cnpj, start_date, end_date)`: Função principal que autentica o usuário, verifica a existência do arquivo Excel e obtém os dados da API.

### interface.py

Este módulo lida com a atualização dos arquivos Excel com base nos dados obtidos da API.

#### Funções

- `update_excel_balance(company_name, excel_file, data, fiscal_year)`: Atualiza o arquivo Excel com os dados fornecidos.
- `main(cnpj, start_date, end_date)`: Função principal que chama a função `api_main` do módulo `api.py` e atualiza o arquivo Excel.

### main.py

Este script executa a função principal do módulo `interface.py` com parâmetros de exemplo.

## Execução

Para executar o robô, use o comando:
```bash
python main.py
```
