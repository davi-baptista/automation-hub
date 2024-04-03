# API de Automação de Extração de Dados

## Introdução

Este projeto utiliza uma API para automatizar a extração de dados, processando esses dados para limpeza, comparação e análise entre planilhas Excel distintas. Destina-se a facilitar a manutenção de uma base de dados atualizada de clientes ou empresas, identificando registros ativos, inativos, e novos de forma eficiente.

## Índice

- [Instalação](#instalação)
- [Configuração](#configuração)
- [Uso](#uso)
- [Funcionalidades](#funcionalidades)
- [Dependências](#dependências)
- [Documentação](#documentação)
- [Exemplos](#exemplos)
- [Resolução de Problemas](#resolução-de-problemas)

## Instalação

Para utilizar este script, é necessário ter Python instalado na sua máquina, bem como o gerenciador de pacotes pip para instalar as seguintes dependências:

```bash
requests
pandas
xlsxwriter
```

Instale as dependências com o comando:

```bash
pip install requests pandas xlsxwriter
```

## Configuração

Antes de usar o script, é necessário configurar o acesso à API:

1. Obtenha um token de autenticação e salve-o em um arquivo de texto (`token.txt`) no diretório especificado no script `api.py`.
2. Certifique-se de que as credenciais de login para a API estão corretas e seguras.

## Uso

Para usar este projeto:

1. Execute o script `main.py` para iniciar o processo de extração e processamento de dados.
2. O script irá utilizar o token de autenticação para acessar a API, extrair os dados necessários, e processá-los conforme especificado.

## Funcionalidades

- Autenticação segura na API para extração de dados.
- Limpeza e formatação de dados de CNPJ em planilhas Excel.
- Comparação de dados entre planilhas para identificar o status dos registros.
- Geração de uma nova planilha com os resultados da análise.

## Dependências

- requests: Para realizar solicitações HTTP à API.
- pandas: Usado na manipulação e análise dos dados.
- xlsxwriter: Necessário para a criação e formatação de arquivos Excel.

## Documentação

Consulte a documentação oficial das bibliotecas utilizadas para mais detalhes:

- [Requests](https://docs.python-requests.org/en/master/)
- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [XlsxWriter](https://xlsxwriter.readthedocs.io/)

## Exemplos

Exemplos específicos de uso não são aplicáveis devido à natureza personalizada deste script, que depende das planilhas de dados do usuário e da configuração da API.

## Resolução de Problemas

Para problemas com a instalação de dependências, verifique a documentação oficial das bibliotecas. Para problemas de autenticação ou extração de dados, certifique-se de que as credenciais e o token de autenticação estejam corretos e válidos.