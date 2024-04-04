# NiboExtractDownloader

## Introdução

`NiboExtractDownloader` é uma ferramenta automatizada construída para facilitar o download de extratos bancarios de cada empresa dentro da plataforma Nibo. Utilizando tecnologias de automação web, o `NiboExtractDownloader` permite aos usuários automatizar o processo de login e download de arquivos com base em critérios específicos como ano e mês. Ideal para contadores, gestores financeiros e qualquer usuário que necessite regularmente de extrair informações contábeis ou financeiras do Nibo.

## Tabela de Conteúdos

- [Instalação](#instalação)
- [Uso](#uso)
  - [Interface Gráfica](#interface-gráfica)
  - [Preenchimento do Excel](#preenchimento-do-excel)
- [Características](#características)
- [Dependências](#dependências)
- [Configuração](#configuração)
- [Exemplos](#exemplos)
- [Solução de Problemas](#solução-de-problemas)
- [Contribuidores](#contribuidores)
- [Licença](#licença)

## Instalação

Para instalar o `NiboExtractDownloader`, siga estes passos:

1. Certifique-se de ter Python instalado em sua máquina. É recomendado usar Python 3.8 ou superior.
2. Baixe o projeto para o seu ambiente de trabalho.
3. Abra um terminal ou prompt de comando na pasta do projeto.
4. Instale as dependências necessárias diretamente com os seguintes comandos:

```bash
pip install botcity-web==0.0.1 pandas tk webdriver-manager python-dateutil shutil
```

## Uso

### Interface Gráfica

`NiboExtractDownloader` oferece uma interface gráfica amigável para facilitar o uso:

1. Execute o script `main.py`.
2. Na interface que aparece, selecione o arquivo Excel com os critérios de download.
3. Especifique o ano e os meses desejados para download.
4. Clique em "Iniciar" para proceder com o download.

### Preenchimento do Excel

O arquivo Excel deve ser preenchido com as seguintes informações:

- **Nome da Empresa**: Nome da empresa dentro do Nibo

## Características

- Automatiza o processo de login no Nibo.
- Permite o download de documentos com base em critérios de data específicos.
- Interface gráfica intuitiva para facilitar a operação.

## Dependências

O `NiboExtractDownloader` depende das seguintes bibliotecas Python:

- `botcity.web`
- `pandas`
- `tkinter`
- `webdriver_manager`
- `shutil`
- `datetime`
- `dateutil`

## Configuração

Antes de utilizar o `NiboExtractDownloader`, certifique-se de ter o ChromeDriver compatível com a versão do seu navegador Chrome instalado. O script `main.py` tentará gerenciar isso automaticamente, mas problemas de compatibilidade podem requerer instalação manual.

## Exemplos

Para iniciar o download de documentos de Março de 2021, por exemplo:

1. Preencha o arquivo Excel com os nomes das empresas desejadas na coluna 'Nome_empresas'.
2. Execute o `main.py` e siga as instruções na interface gráfica.

## Solução de Problemas

Se encontrar erros durante o uso do `NiboExtractDownloader`, verifique se:

- Certifique de colocar login e senha dentro do código-fonte.
- Todas as dependências Python estão corretamente instaladas.
- Todos os dados necessarios estão corretos.
