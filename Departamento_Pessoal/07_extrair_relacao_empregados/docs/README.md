# Capturador de Funcionários

Este projeto contém scripts automatizados para extrair informações de funcionários de um arquivo PDF e salvar essas informações em um arquivo Excel, utilizando uma interface gráfica construída com PyQt5.

## Requisitos

- Python 3.6+
- PyQt5
- pdfplumber
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
    pip install PyQt5 pdfplumber openpyxl
    ```

## Estrutura do Código

### `interface.py`

Este módulo contém a interface gráfica do aplicativo, que permite aos usuários digitar o caminho de um arquivo PDF, selecionar o arquivo e enviar essas informações para serem processadas.

#### Funções Principais

- `__init__()`: Inicializa a janela principal e configura a interface do usuário.
- `setup_ui()`: Configura os layouts e widgets da interface.
- `toggle_theme()`: Alterna entre o modo claro e escuro.
- `apply_theme()`: Aplica o tema selecionado (claro ou escuro).
- `select_file()`: Abre o diálogo para selecionar um arquivo.
- `send_info()`: Envia as informações do arquivo PDF para processamento.

### `main.py`

Este módulo lida com a lógica de extração das informações dos funcionários a partir de um arquivo PDF e salva essas informações em um arquivo Excel.

#### Funções

- `extract_info(lines, start_index)`: Extrai as informações do nome, PIS e valor das linhas do PDF.
- `extract_company_and_registration(lines)`: Extrai o nome da empresa e o número de registro das linhas do PDF.
- `extract_comp_value(lines)`: Extrai o valor da compensação das linhas do PDF.
- `save_to_excel(data, filename)`: Salva os dados extraídos em um arquivo Excel.
- `main(filepath)`: Função principal que processa o arquivo PDF e salva os dados no arquivo Excel.

## Execução

Para executar a interface gráfica do aplicativo, use o comando:
```bash
python interface.py
```

## Utilização

1. Abra o aplicativo e digite o caminho do arquivo PDF.
2. Clique em "Selecionar Arquivo" para escolher o arquivo PDF.
3. Clique em "Enviar" para processar os dados.
4. Use o botão "Alterar Tema" para alternar entre os modos claro e escuro.

O aplicativo processará as informações do arquivo PDF e salvará os dados extraídos em um arquivo Excel.

## Observações

- Certifique-se de que o arquivo PDF esteja no formato esperado para que o processamento ocorra corretamente.
- O log na interface fornecerá feedback sobre o status do processamento.
