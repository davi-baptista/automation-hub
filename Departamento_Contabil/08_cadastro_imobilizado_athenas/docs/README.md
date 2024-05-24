# Bot de Cadastro Imobilizado Athenas

Este projeto contém um robô automatizado utilizando a biblioteca BotCity para interagir com o sistema Athenas. O objetivo do robô é realizar o cadastro de imobilizados nas empresas listadas em uma planilha Excel.

## Requisitos

- Python 3.6+
- BotCity Core
- PyGetWindow
- PyAutoGUI
- Pandas

## Instalação

1. Clone o repositório:
   ```bash
   git clone https://github.com/seu-usuario/seu-repositorio.git
   ```
2. Navegue até o diretório do projeto:
   ```bash
   cd seu-repositorio
   ```
3. Instale as dependências:
   ```bash
   pip install botcity-core pygetwindow pyautogui pandas
   ```

## Configuração

Antes de executar o script, você precisa configurar alguns parâmetros no código:

- Caminho do arquivo Excel com os dados:
   ```bash
  filepath = r'C:\\caminho\\para\\seu\\arquivo\\IMOB BASE.xlsx'
  ```

## Execução

Para executar o robô, use o comando:
```bash
python main.py
```

## Estrutura do Código

### Classe Bot

A classe principal do robô, `Bot`, herda de `DesktopBot` e contém os métodos principais para a automação:

- `action(execution)`: Método principal que inicia a execução do robô, carregando os dados do Excel e iterando pelas empresas para realizar o cadastro de imobilizados.
- `open_athenas()`: Foca e abre o aplicativo Athenas.
- `abrindo_empresa(companie)`: Abre a empresa especificada no sistema.
- `abrindo_cadastro_imobilizado()`: Navega até a seção de cadastro de imobilizados.
- `cadastro_imobilizado(df_codigos, codigo_antigo)`: Realiza o cadastro de imobilizados utilizando os códigos fornecidos.
- `salvar_e_proximo()`: Salva o cadastro atual e avança para o próximo.
- `verificar_ativo()`: Verifica se o imobilizado está ativo.
- `verificar_e_trocar(df)`: Verifica e troca a conta contábil se necessário.
- `bater_codigo(df, codigo_antes)`: Compara e obtém o código correto a ser utilizado.
- `not_found(label)`: Método auxiliar para tratar elementos não encontrados.