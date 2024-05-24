# Projeto de Automação de Validação de Demonstrativo Contábil

Este projeto contém scripts automatizados para autenticação de usuários, verificação e atualização de arquivos Excel com base nos dados obtidos da API da Athenas.

## Requisitos

- Python 3.6+
- requests
- openpyxl
- pandas

## Instalação

1. Clone o repositório:
   ```bash
   /git clone https://github.com/seu-usuario/seu-repositorio.git/
2. Navegue até o diretório do projeto:
   ```bash
   /cd seu-repositorio/
   ```
3. Instale as dependências executando o comando abaixo:
   ```bash
   /pip install requests openpyxl pandas/
   ```

## Estrutura do Código

### `api.py`

Este módulo lida com a autenticação de usuários na API da Athenas e verifica a existência de arquivos Excel para planos de contas específicos.

#### Funções

- `users_auth()`: Autentica o usuário e obtém um token, ou usa um token existente.
- `check_excel_existence(account_plan)`: Verifica se o arquivo Excel para o plano de contas fornecido existe.
- `handle_401_error()`: Lida com o erro 401 atualizando o token de autenticação.
- `check_account_plan(cnpj, url, headers)`: Verifica o status do plano de contas para o CNPJ fornecido.
- `api_main(cnpj, start_date, end_date)`: Função principal que autentica o usuário, verifica a existência do arquivo Excel e obtém os dados da API.

### `interface.py`

Este módulo lida com a atualização dos arquivos Excel com base nos dados obtidos da API.

#### Funções

- `update_excel_balance(company_name, cnpj_formatado, fiscal_year, sheet_name, excel_file, data)`: Atualiza o arquivo Excel com os dados fornecidos.
- `formatar_cnpj(cnpj)`: Formata o CNPJ para o padrão desejado.
- `first_last_day_of_year_formatted(year)`: Retorna o primeiro e último dia do ano no formato MM-DD-YYYY.

### `main.py`

Este script executa a função principal do módulo `interface.py` com parâmetros de exemplo.

#### Funções

- `main(cnpj, year)`: Função principal que chama a função `api_main` do módulo `api.py` e atualiza o arquivo Excel.

## Execução

Para executar o script principal, use o comando:
```bash
python main.py
```

## Observações

- Certifique-se de que os arquivos Excel estejam no formato esperado para que o processamento ocorra corretamente.
- O log na interface fornecerá feedback sobre o status do processamento.