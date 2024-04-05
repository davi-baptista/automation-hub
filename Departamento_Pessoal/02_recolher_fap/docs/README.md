# Recolhedor de FAP

## Introdução

Este projeto é composto por dois scripts Python principais, `main_outdated.py` e `baixar_certificados.py`, desenvolvidos para automatizar tarefas específicas relacionadas à gestão de certificados digitais e à interação com um website específico para recolher informações de FAP (Fator Acidentário de Prevenção).

## Índice

- [Instalação](#instalação)
- [Uso](#uso)
  - [Baixar Certificados](#baixar-certificados)
  - [Script Principal](#script-principal)
- [Funcionalidades](#funcionalidades)
- [Dependências](#dependências)
- [Configuração](#configuração)
- [Documentação](#documentação)
- [Exemplos](#exemplos)
- [Solução de Problemas](#solução-de-problemas)

## Instalação

Para utilizar o 'Recolhedor de FAP', é necessário ter Python instalado na máquina. Além disso, o bot depende de bibliotecas específicas, que podem ser instaladas via pip:

```
pip install botcity-core-desktop pywinauto pyautogui pandas openpyxl
```

## Uso

### Baixar Certificados

O script `baixar_certificados.py` é usado para baixar todos os certificados necessários para o funcionamento do script principal. É importante garantir que os caminhos para os certificados e as dependências estejam corretos.

```
python baixar_certificados.py
```

### Script Principal

O `main_outdated.py` executa a lógica principal do projeto, utilizando os certificados baixados pelo script auxiliar.

```
python main_outdated.py
```

## Funcionalidades

- **Baixar Certificados**: Automatiza o processo de download de certificados digitais.
- **Automatização de Tarefas Web**: Realiza tarefas automatizadas em um website específico para coletar informações de FAP.

## Dependências

Os scripts dependem de várias bibliotecas externas, incluindo:

- `botcity.core`
- `pywinauto`
- `pyautogui`
- `pandas`
- `openpyxl`

## Configuração

Certifique-se de ajustar os caminhos para os certificados, pastas de downloads e destinos dentro dos scripts para corresponder ao seu ambiente de execução.

## Exemplos

Suponha que você queira automatizar a coleta de informações de FAP para uma lista de empresas. O script principal navegará até o website desejado, fará login usando um certificado digital e coletará as informações necessárias, salvando os resultados em um arquivo numa pasta e salvando o status em Excel.

## Solução de Problemas

Para qualquer problema relacionado ao funcionamento do bot, verifique se todas as dependências estão corretamente instaladas e se os caminhos para arquivos e diretórios estão configurados corretamente.