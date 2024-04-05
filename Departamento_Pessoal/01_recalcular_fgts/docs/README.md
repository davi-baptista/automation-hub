# Bot de Recálculo do FGTS

## Introdução

Este projeto implementa um bot de automação de desktop projetado para automatizar o processo de recálculo do FGTS (Fundo de Garantia do Tempo de Serviço) e a transmissão das informações para o sistema da Caixa Econômica Federal. Utilizando a biblioteca `botcity.core`, junto com outras ferramentas como `pandas`, `pyautogui`, `pywinauto`, e `pygetwindow`, o bot executa uma série de ações automáticas no sistema operacional Windows para processar as informações necessárias e realizar o recálculo e a transmissão do FGTS.

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

Para executar este bot, é necessário ter Python instalado em seu sistema. Recomenda-se o uso do Python 3.8 ou superior. As dependências do projeto podem ser instaladas manualmente utilizando o pip, o gerenciador de pacotes do Python. As principais bibliotecas necessárias incluem:

```bash
pip install botcity.core pandas pyautogui pywinauto pygetwindow
```

## Uso

Para iniciar o bot, navegue até o diretório do projeto no terminal e execute:

```bash
python main.py
```

## Funcionalidades

- **Leitura de Dados**: Lê dados de uma planilha Excel especificada para processar informações de competências, CNPJ, datas de pagamento, entre outros.
- **Manipulação de Arquivos**: Abre, valida e processa arquivos `.re`, `.bkp`, e `.gbk` relacionados às obrigações do FGTS.
- **Interação com Aplicativos**: Utiliza `pyautogui` e `pywinauto` para interagir com aplicativos do Windows, como o SEFIP e o navegador, para a execução de tarefas específicas.
- **Transmissão de Dados**: Realiza a transmissão de dados recálculados para o sistema da Caixa através da Conectividade Social.

## Dependências

- botcity.core
- pandas
- pyautogui
- pywinauto
- pygetwindow
- re
- time
- os
- shutil

## Configuração

Antes de executar o bot, é necessário configurar os caminhos dos arquivos de entrada, saída e os diretórios dos aplicativos utilizados pelo bot conforme especificado nas variáveis no início do script. Além disso, é essencial garantir que todos os aplicativos necessários estejam instalados no sistema.

## Documentação

A documentação detalhada das bibliotecas utilizadas pode ser encontrada nos respectivos sites:

- [BotCity](https://botcity.dev/)
- [Pandas](https://pandas.pydata.org/)
- [PyAutoGUI](https://pyautogui.readthedocs.io/)
- [Pywinauto](https://pywinauto.readthedocs.io/)
- [PyGetWindow](https://pypi.org/project/PyGetWindow/)

## Exemplos

Um exemplo detalhado de uso é fornecido no script `main.py`, demonstrando a automação do processo de recálculo do FGTS.

## Solução de Problemas

Caso encontre problemas durante a execução do bot, verifique se todas as dependências foram instaladas corretamente e se os caminhos dos arquivos e diretórios estão configurados corretamente. Além disso, certifique-se de que os aplicativos necessários estão instalados e funcionando no seu sistema.