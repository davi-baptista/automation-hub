# Envio de balancete para plataforma

Este projeto contém um robô automatizado utilizando a biblioteca BotCity para interagir com o site Accountify. O objetivo do robô é fazer login no sistema, acessar grupos específicos, e enviar balancetes conforme valores obtidos de arquivos Excel em um diretório específico.

## Requisitos

- Python 3.6+
- BotCity Web
- WebDriver Manager
- PyAutoGUI
- Pynput

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
   pip install botcity-web webdriver-manager pyautogui pynput
    ```

## Configuração

Antes de executar o script, você precisa configurar alguns parâmetros no código:

- Email e senha para login no sistema Accountfy:
  ```bash
  email = 'seu-email@example.com'
  password = 'sua-senha'
  ```
- Caminho do diretório onde estão os arquivos Excel:
  ```bash
  filepath = r'C:\\caminho\\para\\seus\\arquivos'
  ```

## Execução

Para executar o robô, use o comando:
```bash
python main.py
```

## Estrutura do Código

### Classe MyWebBot

A classe principal do robô, `MyWebBot`, herda de `WebBot` e contém os métodos principais para a automação:

- `action()`: Método principal que inicia a execução do robô, fazendo login, acessando grupos e enviando balancetes.
- `configure()`: Configura o ChromeDriver e as opções do navegador.
- `login(email, password)`: Realiza o login no sistema Accountfy.
- `join_group()`: Acessa o grupo específico.
- `obter_valores(diretorio)`: Obtém os valores dos balancetes a partir dos arquivos Excel no diretório especificado.
- `sending_balancete(valor, index)`: Envia os balancetes para o sistema Accountfy.
- `not_found(label)`: Método auxiliar para tratar elementos não encontrados.