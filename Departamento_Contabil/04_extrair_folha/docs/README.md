# Processador de Dados Contábeis

## Introdução
Este script Python é projetado para automatizar o processamento de dados contábeis, incluindo a leitura de provisões e resumos gerais de arquivos PDF e Excel, tratamento e formatação dos dados, e a geração de arquivos de texto formatados para uso contábil.

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
Para usar este script, é necessário ter Python 3 instalado em seu ambiente, além das seguintes bibliotecas: pandas, python-dateutil, PyMuPDF (fitz), tkinter e openpyxl. A instalação pode ser feita via pip:

```bash
pip install pandas python-dateutil PyMuPDF tkinter openpyxl
```

## Uso
O script é executado através de uma interface gráfica Tkinter, onde o usuário pode selecionar o diretório contendo os arquivos a serem processados e iniciar o processamento com um clique.

1. Execute o script:
```bash
python main.py
```
2. Use o botão "Buscar" para selecionar o diretório contendo os arquivos PDF e Excel.
3. Clique em "Enviar" para iniciar o processamento dos dados.

## Funcionalidades
- Leitura e processamento de dados de arquivos PDF e Excel.
- Tratamento e formatação de dados contábeis.
- Geração de arquivos de texto formatados para importação em sistemas contábeis.

## Dependências
- pandas
- python-dateutil
- PyMuPDF (fitz)
- tkinter
- openpyxl

## Configuração
Nenhuma configuração adicional é necessária além da instalação das dependências.

## Documentação
A documentação das bibliotecas utilizadas pode ser encontrada nos seguintes links:

- Pandas: https://pandas.pydata.org/pandas-docs/stable/
- Python-dateutil: https://dateutil.readthedocs.io/en/stable/
- PyMuPDF (fitz): https://pymupdf.readthedocs.io/en/latest/
- Tkinter: https://docs.python.org/3/library/tkinter.html
- Openpyxl: https://openpyxl.readthedocs.io/en/stable/

## Exemplos
Exemplos específicos de uso podem ser encontrados comentados no código-fonte, demonstrando como as funções podem ser aplicadas para diferentes cenários de processamento de dados.

## Solução de Problemas
Para problemas comuns, verifique a seção de solução de problemas no código ou entre em contato com os contribuidores.