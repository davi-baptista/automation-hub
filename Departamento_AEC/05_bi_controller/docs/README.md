# Dashboard de Sucesso do Cliente com Atualização Automática de Dados

## Introdução

Este projeto consiste em um dashboard interativo desenvolvido com Streamlit, que apresenta várias páginas com visualizações de dados distintas, como dados gerais, comparativo de faturamento, comparativo de funcionários e comparativo de notas. Além disso, incorpora uma API para automatização da extração de dados, permitindo atualizações automáticas diretamente no banco de dados.

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

Para instalar e executar o dashboard, siga estas etapas:

```bash
pip install streamlit pandas requests xlsxwriter
```

## Configuração

1. Adicione todos os endpoints para utilização da API dentro do código.
2. Configure o `token_filepath` no arquivo `api.py` com o caminho para o arquivo contendo o token de autenticação da API.
3. Atualize as credenciais de login da API (`login` e `password`) em `api.py`.

## Uso

Para iniciar o dashboard, execute:

```bash
streamlit run app.py
```

Acesse o dashboard através do navegador no endereço indicado pelo Streamlit.

## Funcionalidades

- **Visualização de Dados Gerais:** Apresenta uma visão geral dos dados relevantes.
- **Comparativo de Faturamento:** Exibe uma comparação do faturamento entre diferentes períodos.
- **Comparativo de Funcionários:** Mostra a evolução do quadro de funcionários.
- **Comparativo de Notas:** Compara as notas fiscais emitidas em diferentes períodos.
- **Atualização Automática de Dados:** Utiliza uma API para extrair e atualizar os dados diretamente no banco de dados.

## Dependências

- `streamlit`: Para criação do dashboard.
- `pandas`: Para manipulação de dados.
- `requests`: Para fazer solicitações à API.
- `xlsxwriter`: Para escrever dados em arquivos Excel com formatação.

## Documentação

Consulte a documentação das bibliotecas utilizadas para mais detalhes:

- [Streamlit](https://docs.streamlit.io)
- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [Requests](https://docs.python-requests.org/en/master/)
- [XlsxWriter](https://xlsxwriter.readthedocs.io/)

## Exemplos

Devido à natureza específica deste projeto, que depende de informações confidenciais e sistemas externos, exemplos detalhados de uso não são fornecidos.

## Resolução de Problemas

Verifique se todas as dependências estão corretamente instaladas e se as credenciais de acesso à API estão atualizadas. Para problemas com visualizações específicas, assegure-se de que os dados estão sendo corretamente extraídos e formatados.