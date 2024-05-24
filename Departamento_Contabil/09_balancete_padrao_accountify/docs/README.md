# Balancete Excel para Padrão Accountify

Este projeto contém um aplicativo de desktop para processar arquivos Excel de balancetes e convertê-los para o padrão Accountify, utilizando a biblioteca PyQt5 para a interface gráfica.

## Requisitos

- Python 3.6+
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
   pip install pandas openpyxl PyQt5
    ```

## Estrutura do Código

### `main.py`

Este script contém a interface gráfica do aplicativo, que permite aos usuários selecionar diretórios com arquivos Excel, processar esses arquivos e convertê-los para o padrão Accountify.

#### Funções Principais

- `__init__()`: Inicializa a janela principal e configura a interface do usuário.
- `setup_ui()`: Configura os layouts e widgets da interface.
- `toggle_theme()`: Alterna entre o modo claro e escuro.
- `apply_theme()`: Aplica o tema selecionado (claro ou escuro).
- `log(message)`: Adiciona mensagens ao log de interface.
- `send_directory()`: Processa o diretório selecionado e salva os arquivos no formato Accountify.
- `processar_arquivo(caminho_arquivo)`: Processa um único arquivo Excel.
- `choose_directory()`: Abre o diálogo para selecionar um diretório.

### Temas

- `apply_dark_mode(app)`: Aplica o tema escuro ao aplicativo.
- `apply_light_mode(app)`: Aplica o tema claro ao aplicativo.

## Execução

Para executar o aplicativo, use o comando:
   ```bash
python main.py
   ```

## Utilização

1. Abra o aplicativo e selecione o diretório que contém os arquivos Excel.
2. Clique em "Buscar" para selecionar o diretório.
3. Clique em "Enviar" para processar os arquivos.
4. Use o botão "Alterar Tema" para alternar entre os modos claro e escuro.

O aplicativo processará os arquivos no diretório selecionado e os salvará em uma pasta chamada `accountify`.

## Observações

- Certifique-se de que os arquivos Excel estejam no formato esperado para que o processamento ocorra corretamente.
- O log na interface fornecerá feedback sobre o status do processamento.