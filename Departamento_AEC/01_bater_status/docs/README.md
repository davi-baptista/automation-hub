# Manipulação de Dados de CNPJ

## Introdução

Este projeto visa a automatização do processo de limpeza, comparação e análise de dados de CNPJ entre duas planilhas Excel distintas. Ele facilita a identificação de registros ativos, inativos e novos, agilizando o processo de manutenção de uma base de dados atualizada de clientes ou empresas.

## Índice

- [Instalação](#instalação)
- [Uso](#uso)
- [Funcionalidades](#funcionalidades)
- [Dependências](#dependências)
- [Documentação](#documentação)
- [Exemplos](#exemplos)
- [Resolução de Problemas](#resolução-de-problemas)

## Instalação

Para rodar este script, é necessário ter o Python instalado em sua máquina, além do gerenciador de pacotes pip para instalar as seguintes bibliotecas:

```bash
pandas
openpyxl
```

Instale as dependências usando o comando:

```bash
pip install pandas openpyxl
```

## Uso

1. Certifique-se de que as planilhas Excel de origem estejam no formato correto e nos caminhos especificados no script.
2. Execute o script \`main.py\` para realizar a limpeza dos dados, a comparação dos CNPJs entre as planilhas e a geração de uma nova planilha com o resultado da comparação.

## Funcionalidades

- Limpeza dos dados de CNPJ para remover caracteres especiais.
- Comparação dos CNPJs entre duas planilhas para identificar quais estão ativos, inativos ou precisam ser ativados.
- Geração de uma nova planilha Excel com os resultados da comparação, incluindo a indicação do status de cada CNPJ.

## Dependências

- pandas: Usado para manipulação e análise dos dados.
- openpyxl: Necessário para a leitura e escrita de arquivos Excel.

## Documentação

A documentação das bibliotecas utilizadas pode ser encontrada nos seguintes links:

- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)

## Exemplos

Não aplicável, pois a execução do script depende das planilhas específicas do usuário.

## Resolução de Problemas

Para problemas relacionados à instalação das dependências, verifique a documentação oficial das bibliotecas pandas e openpyxl. Para outras questões, certifique-se de que os caminhos das planilhas e os formatos dos dados estejam corretos conforme esperado pelo script.