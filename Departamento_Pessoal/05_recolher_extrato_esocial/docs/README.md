# Recolhedor de Extrato eSocial

## Introdução

Este projeto desenvolve um bot de web scraping e automação de tarefas na web, focado principalmente na interação com o site da Caixa Econômica Federal para a obtenção de extratos analíticos do FGTS de funcionários. O bot navega pelo site, faz login, realiza consultas específicas e baixa os extratos em formato PDF.

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

Para executar este bot, você precisa ter Python instalado na sua máquina, além das seguintes bibliotecas:

- pandas
- botcity.web
- selenium
- webdriver_manager

Instale as dependências com o seguinte comando:

```bash
pip install pandas botcity.web selenium webdriver_manager
```

## Uso

Para iniciar o bot, execute o script `bot.py` no terminal ou prompt de comando:

```bash
python bot.py
```

## Funcionalidades

- **Login Automatizado:** Realiza o login no site da Caixa Econômica Federal.
- **Download de Extratos:** Baixa o extrato analítico e o extrato por trabalhador em formato PDF.
- **Atualização de Status:** Atualiza o status da tarefa no arquivo Excel de entrada após a conclusão de cada operação.

## Dependências

- Python 3.x
- Pandas
- Selenium
- Webdriver_Manager
- botcity.web

## Configuração

O script é configurado para operar com o navegador Chrome. As opções de inicialização e o diretório de download dos arquivos podem ser ajustados no método `configure`.

## Documentação

Para mais informações sobre as bibliotecas utilizadas, consulte:

- [Pandas](https://pandas.pydata.org/pandas-docs/stable/index.html)
- [Selenium](https://www.selenium.dev/documentation/)
- [BotCity Web](https://docs.botcity.dev/web/)

## Exemplos

Os exemplos de uso específicos estão codificados dentro dos métodos `extrato_analitico`, `extrato_trabalhador`, e `navigate_and_download_pdf`.

## Solução de Problemas

Para qualquer problema relacionado ao funcionamento do bot, verifique se todas as dependências estão corretamente instaladas e se os caminhos para arquivos e diretórios estão configurados corretamente.