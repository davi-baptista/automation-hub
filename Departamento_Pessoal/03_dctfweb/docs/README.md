# Bot de Automação DCTFWeb

## Introdução

Este projeto consiste em um bot de automação desenvolvido para facilitar o processo de transmissão e consulta de declarações na DCTFWeb, um sistema da Receita Federal do Brasil. Utilizando técnicas de automação de desktop e web, o bot automatiza tarefas repetitivas, como preenchimento de datas, seleção de certificados digitais, navegação entre páginas e download de documentos necessários.

## Índice

- [Introdução](#introdução)
- [Funcionalidades](#funcionalidades)
- [Instalação](#instalação)
- [Uso](#uso)
- [Dependências](#dependências)
- [Configuração](#configuração)
- [Documentação](#documentação)
- [Exemplos](#exemplos)
- [Solução de Problemas](#solução-de-problemas)

## Funcionalidades

- Abertura automática do navegador e navegação até o e-CAC da Receita Federal.
- Seleção automática de certificados digitais.
- Preenchimento automático de datas com base no mês atual.
- Download de recibos de entrega e DARFs.
- Preenchimento e atualização de planilhas Excel com o status de cada empresa processada.

## Instalação

Para instalar as dependências necessárias para este projeto, é recomendável usar um ambiente virtual Python. Você pode instalar as dependências utilizando o comando:

```bash
pip install -r requirements.txt
```

## Uso

Para executar o bot, utilize o comando:

```bash
python bot.py
```

Certifique-se de que todas as configurações e caminhos de arquivos estejam corretos antes de executar o script.

## Dependências

Este projeto depende das seguintes bibliotecas Python:

- botcity.core
- botcity.web
- pandas
- pyautogui
- pywinauto
- datetime
- dateutil

## Configuração

Antes de utilizar o bot, é necessário configurar os seguintes itens:

1. Caminho do navegador e do certificado digital.
2. Diretórios de download e saída para os arquivos baixados.
3. Planilha Excel com as empresas a serem processadas.

## Documentação

A documentação de cada função está disponível no código-fonte, explicando o propósito e o funcionamento de cada uma.

## Exemplos

Exemplos de uso do bot incluem a automação do processo de envio de declarações fiscais e a consulta de saldo devedor para empresas específicas.

## Solução de Problemas

Para qualquer problema relacionado ao funcionamento do bot, verifique se todas as dependências estão corretamente instaladas e se os caminhos para arquivos e diretórios estão configurados corretamente.