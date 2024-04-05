# Recolhedor de Dados dos Funcionarios

## Introdução

O MyWebBot é um bot automatizado desenvolvido para interagir com o portal eSocial do governo brasileiro. Seu principal objetivo é navegar pelo portal, realizar buscas específicas de CPFs e nomes de empregados, e baixar as informações cadastrais desses empregados em formato PDF. Esse processo automatiza tarefas repetitivas relacionadas à gestão de empregados e ao cumprimento de obrigações legais junto ao eSocial.

## Índice

- [Introdução](#introdução)
- [Instalação](#instalação)
- [Uso](#uso)
- [Funcionalidades](#funcionalidades)
- [Dependências](#dependências)
- [Configuração](#configuração)
- [Documentação](#documentação)
- [Exemplos](#exemplos)
- [Solução de Problemas](#solução-de-problemas)

## Instalação

Para utilizar o MyWebBot, é necessário ter Python instalado na máquina. Além disso, o bot depende de bibliotecas específicas, que podem ser instaladas via pip:

```
pip install botcity.web pandas webdriver_manager
```

## Uso

1. Configure o ambiente virtual Python e instale as dependências necessárias.
2. Modifique o caminho do arquivo `Esocial.xlsx` no método `action` para corresponder ao local onde o arquivo Excel está armazenado em sua máquina.
3. Execute o script Python:

```
python bot.py
```

## Funcionalidades

- **Login Automatizado:** Acessa automaticamente o portal eSocial utilizando certificado digital.
- **Busca de Empregados:** Realiza a busca de empregados pelo CPF e nome.
- **Download de PDFs:** Baixa os dados cadastrais dos empregados em formato PDF.
- **Gestão de Status:** Atualiza o status no arquivo Excel para evitar downloads repetidos.

## Dependências

- Python 3
- botcity.web
- pandas
- webdriver_manager
- Google Chrome e ChromeDriver

## Configuração

Altere as configurações do ChromeDriver e os caminhos de download dentro do método `configure` para corresponder ao seu ambiente de trabalho.

## Documentação

Para mais informações sobre as bibliotecas utilizadas, consulte:

- [botcity.web documentation](https://docs.botcity.dev/web/getting-started/)
- [pandas documentation](https://pandas.pydata.org/pandas-docs/stable/)
- [ChromeDriver documentation](https://sites.google.com/a/chromium.org/chromedriver/)

## Exemplos

O script `bot.py` já é um exemplo completo de como utilizar o 'Recolhedor de Dados dos Funcionarios' para automatizar o download de informações cadastrais de empregados do eSocial.

## Solução de Problemas

Para qualquer problema relacionado ao funcionamento do bot, verifique se todas as dependências estão corretamente instaladas e se os caminhos para arquivos e diretórios estão configurados corretamente.