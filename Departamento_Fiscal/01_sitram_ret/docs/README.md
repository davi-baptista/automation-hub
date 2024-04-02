# Ferramenta de extração e manipulação de dados de notas fiscais

## Introdução

Este projeto é uma automação web desenvolvida para extrair informações de notas fiscais do site da SEFAZ do Ceará. A automação navega pelo site, insere chaves de acesso de notas fiscais, extrai informações detalhadas de cada item da nota e salva esses dados em um arquivo Excel formatado.

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

Para executar este projeto, você precisa ter Python instalado em sua máquina, além das seguintes bibliotecas:

- pandas
- tkinter
- selenium
- webdriver-manager
- botcity.web

É recomendável criar um ambiente virtual Python para instalar essas dependências. Use o seguinte comando para instalar todas as dependências necessárias:

```bash
pip install pandas selenium webdriver-manager botcity.web
```

## Uso

Para usar este script, siga estes passos:

1. Inicie o script com `python main.py`.
2. Utilize a interface gráfica para selecionar o arquivo CSV contendo as chaves de acesso e os números das notas fiscais.
3. Clique em "Enviar" para iniciar o processamento.
4. Os resultados serão salvos em um arquivo Excel no mesmo diretório do arquivo CSV original.

## Funcionalidades

- Extração automática de dados de notas fiscais do site da SEFAZ-CE.
- Interface gráfica para fácil seleção de arquivos e acompanhamento do processo.
- Salvamento dos dados extraídos em um arquivo Excel formatado.

## Dependências

Este projeto depende das seguintes bibliotecas Python:

- pandas: Para manipulação de dados e salvamento em arquivo Excel.
- tkinter: Para criação da interface gráfica.
- selenium: Para navegação e interação com páginas web.
- webdriver-manager: Para gerenciamento automático dos drivers de navegadores.
- botcity.web: Framework para automação web.

## Configuração

Nenhuma configuração adicional é necessária após a instalação das dependências.

## Documentação

Para mais informações sobre as bibliotecas utilizadas, consulte a documentação oficial:

- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [Selenium](https://selenium-python.readthedocs.io/)
- [Webdriver Manager](https://github.com/SergeyPirogov/webdriver_manager)
- [BotCity](https://docs.botcity.dev/web/getting-started/)

## Exemplos

Um exemplo de arquivo CSV de entrada deve ter a seguinte estrutura:

```bash
Chave acesso,N° NF
<chave_de_acesso>,<numero_da_nota>
```

## Solução de Problemas

Caso encontre problemas durante o uso do script, verifique se as dependências estão corretamente instaladas e atualizadas. Certifique-se também de que está utilizando a versão mais recente do navegador compatível com o Selenium.