# Bot de Processamento de Notas Fiscais SEFAZ CE

## Introdução

Este projeto é uma aplicação de raspagem de dados e processamento de dados projetada para automatizar a extração de dados de notas fiscais do site da SEFAZ CE (Secretaria da Fazenda do Estado do Ceará). Utiliza o Framework de Automação Web BotCity para navegar e extrair informações especificadas de notas fiscais, processando-as e gerando os resultados em um arquivo Excel. A aplicação conta com uma GUI amigável construída com Tkinter, permitindo aos usuários selecionar facilmente o arquivo fonte para processamento e visualizar logs em tempo real do processo de raspagem.

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

Para executar esta aplicação, certifique-se de ter o Python instalado no seu sistema. Em seguida, instale os pacotes Python necessários executando:

```bash
pip install botcity-web pandas webdriver-manager tkinter
```

## Uso

Para iniciar a aplicação, execute o script \`bot.py\`. Isso abrirá a GUI, onde você pode selecionar o arquivo Excel contendo as chaves de acesso e números das notas fiscais para processamento. Clique em "Enviar" para iniciar o procedimento de raspagem e processamento de dados. O resultado será salvo em um arquivo Excel no mesmo diretório do arquivo de entrada.

## Funcionalidades

- Raspagem automatizada de dados de notas fiscais da SEFAZ CE.
- Processamento de dados e exportação para formato Excel.
- GUI amigável para fácil operação.
- Logs em tempo real do processo de raspagem.

## Dependências

- botcity-web
- pandas
- webdriver-manager
- tkinter

## Configuração

Nenhuma configuração adicional é necessária para executar esta aplicação além dos passos de instalação fornecidos.

## Documentação

Para mais informações sobre as bibliotecas utilizadas neste projeto, por favor consulte a documentação oficial:
- [BotCity Framework de Automação Web](https://botcity.dev/documentation/web)
- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [Webdriver Manager for Python](https://github.com/SergeyPirogov/webdriver_manager)
- [Tkinter](https://docs.python.org/3/library/tkinter.html)

## Exemplos

Um exemplo de formato de arquivo Excel para entrada:
- O arquivo deve conter pelo menos duas colunas: "Chave acesso" e "N° NF", sem nenhum cabeçalho.

## Solução de Problemas

Se você encontrar problemas com o processo de raspagem de dados, certifique-se de que o site da SEFAZ CE não mudou seu layout ou esquema de URL. Além disso, verifique se o seu arquivo de entrada corresponde ao formato requerido.