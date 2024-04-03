# Script de Preenchimento Automático para Athenas

## Introdução

Este script foi desenvolvido para automatizar o processo de preenchimento de informações no software Athenas. Ele lê os dados de um arquivo Excel, que deve conter as colunas NCM, CST e CFOP, e utiliza esses dados para preencher automaticamente no Athenas através de simulações de pressionamentos de teclas e cliques do mouse.

## Conteúdo

- [Instalação](#instalação)
- [Uso](#uso)
- [Funcionalidades](#funcionalidades)
- [Dependências](#dependências)
- [Configuração](#configuração)
- [Documentação](#documentação)
- [Exemplos](#exemplos)
- [Solução de Problemas](#solução-de-problemas)

## Instalação

Para utilizar este script, é necessário ter o Python instalado em sua máquina. Além disso, algumas bibliotecas são essenciais:

```bash
pip install pandas pyautogui keyboard tkinter
```

## Uso

1. Execute o script `bot.py`.
2. Utilize a interface gráfica para selecionar o arquivo Excel contendo os dados de NCM, CST e CFOP.
3. Clique em "Enviar" para iniciar o processamento.
4. O script aguardará que a tecla F6 seja pressionada para iniciar o preenchimento no Athenas.
5. Certifique-se de que o Athenas esteja aberto e na tela correta para o preenchimento dos dados.

## Funcionalidades

- Leitura de dados a partir de um arquivo Excel.
- Interface gráfica para seleção do arquivo e início do processo.
- Preenchimento automático no software Athenas utilizando simulação de interações do usuário.

## Dependências

- pandas
- tkinter
- pyautogui
- keyboard

## Configuração

Nenhuma configuração adicional é necessária além da instalação das dependências.

## Documentação

A documentação das bibliotecas utilizadas pode ser encontrada nos seguintes links:

- Pandas: https://pandas.pydata.org/
- Tkinter: https://docs.python.org/3/library/tkinter.html
- PyAutoGUI: https://pyautogui.readthedocs.io/
- Keyboard: https://github.com/boppreh/keyboard

## Exemplos

O arquivo Excel deve seguir o formato abaixo:

| NCM  | CST | CFOP |
|------|-----|------|
| 1234 | 100 | 5405 |
| ...  | ... | ...  |

## Solução de Problemas

Para problemas relacionados à execução do script, certifique-se de que todas as dependências estejam corretamente instaladas e que o Python esteja atualizado.