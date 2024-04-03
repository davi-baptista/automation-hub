# Automação de Cadastro no SIEG

## Introdução

Este projeto implementa uma solução de automação web para efetuar cadastros no sistema SIEG, utilizando informações de clientes contidas em uma planilha Excel. O objetivo é automatizar o processo de cadastro para aumentar a eficiência e reduzir a possibilidade de erros manuais.

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

Este projeto exige a instalação do Python e algumas bibliotecas essenciais:

```bash
pip install botcity-web webdriver-manager pandas
```

## Configuração

1. Atualize as variáveis `email_gestta`, `password_gestta`, `email_sieg`, e `password_sieg` com as credenciais de acesso aos sistemas Gestta e SIEG, respectivamente.
2. Certifique-se de que o caminho do arquivo Excel (`filepath`) esteja correto e acessível.

## Uso

Execute o script `main.py` para iniciar o processo de automação. O script fará o seguinte:

1. Lê a planilha Excel com os dados dos clientes.
2. Realiza login nos sistemas Gestta e SIEG.
3. Navega até as seções apropriadas em ambos os sistemas.
4. Efetua o cadastro dos clientes no sistema SIEG, utilizando as informações obtidas.

## Funcionalidades

- Leitura de dados de clientes a partir de uma planilha Excel.
- Login automático em sistemas web.
- Navegação e interação automática com elementos web.
- Cadastro automático de clientes no sistema SIEG.

## Dependências

- `botcity-web`: Para a automação de tarefas web.
- `webdriver-manager`: Para gerenciar automaticamente os drivers dos navegadores.
- `pandas`: Para manipulação e análise de dados em Python.

## Documentação

A documentação para as bibliotecas utilizadas pode ser encontrada nos seguintes links:

- [BotCity Web](https://botcity.dev/docs/libraries/web/)
- [WebDriver Manager](https://github.com/SergeyPirogov/webdriver_manager)
- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)

## Exemplos

Devido à natureza específica deste projeto, que depende de informações confidenciais e sistemas externos, exemplos detalhados de uso não são fornecidos.

## Resolução de Problemas

Certifique-se de que todas as dependências estão corretamente instaladas e que as credenciais de acesso aos sistemas estejam atualizadas. Para problemas relacionados à automação web, verifique se a versão do navegador é compatível com o driver utilizado.