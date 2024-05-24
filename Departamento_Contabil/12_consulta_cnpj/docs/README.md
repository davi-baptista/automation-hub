# Consulta de CNPJ

Este projeto contém scripts automatizados para consultar informações de CNPJ de empresas e processar arquivos Excel com essas informações, utilizando uma interface gráfica construída com PyQt5.

## Requisitos

- Python 3.6+
- requests
- pandas
- openpyxl
- PyQt5

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
    pip install requests pandas openpyxl PyQt5
    ```

## Estrutura do Código

### `interface.py`

Este módulo contém a interface gráfica do aplicativo, que permite aos usuários digitar o CNPJ de uma empresa, selecionar um arquivo Excel e enviar essas informações para serem processadas.

#### Funções Principais

- `__init__()`: Inicializa a janela principal e configura a interface do usuário.
- `setup_ui()`: Configura os layouts e widgets da interface.
- `toggle_theme()`: Alterna entre o modo claro e escuro.
- `apply_theme()`: Aplica o tema selecionado (claro ou escuro).
- `select_file()`: Abre o diálogo para selecionar um arquivo.
- `send_company_info()`: Envia as informações da empresa para processamento.

### `main.py`

Este módulo lida com a lógica de captura e processamento das informações de CNPJ de empresas, utilizando a API BrasilAPI para obter os dados.

#### Funções

- `clean_cnpj(cnpj)`: Limpa o CNPJ removendo caracteres não numéricos.
- `fetch_cnpj_info(cnpj, count)`: Obtém e processa as informações do CNPJ a partir da API.
- `capture_cnpj_data(cnpj, count)`: Captura os dados do CNPJ e lida com possíveis exceções.

## Execução

Para executar a interface gráfica do aplicativo, use o comando:
```bash
python interface.py
```

## Utilização

1. Abra o aplicativo e digite o CNPJ da empresa.
2. Clique em "Selecionar Arquivo" para escolher o arquivo Excel.
3. Clique em "Enviar" para processar os dados.
4. Use o botão "Alterar Tema" para alternar entre os modos claro e escuro.

O aplicativo processará as informações do CNPJ e atualizará o arquivo Excel com os dados obtidos.

## Observações

- Certifique-se de que o arquivo Excel esteja no formato esperado para que o processamento ocorra corretamente.
- O log na interface fornecerá feedback sobre o status do processamento.