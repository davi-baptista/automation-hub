# Automação de Alertas de Certificados SIEG

## Introdução

Este projeto implementa uma solução automatizada para monitorar o status de certificados digitais no portal SIEG, extrair essas informações, e enviar alertas por e-mail para os responsáveis. O objetivo é prevenir a expiração de certificados essenciais para os processos da empresa.

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

Este projeto requer Python e algumas bibliotecas específicas para a automação web e manipulação de dados:

```bash
pip install botcity-web selenium pandas webdriver-manager
```

## Configuração

Para a execução correta do bot, siga estes passos:

1. Configure as credenciais de e-mail (remetente e senha) diretamente no script ou em um arquivo de configuração seguro.
2. Ajuste os diretórios de entrada e saída para os arquivos Excel de acordo com o ambiente de execução.

## Uso

Execute o script `main.py` para iniciar o processo de automação. O bot realizará as seguintes ações:

1. Acessa o portal SIEG e realiza o login.
2. Navega até a seção de gerenciamento de certificados.
3. Baixa a planilha de certificados.
4. Processa a planilha, identificando certificados vencidos e próximos do vencimento.
5. Envia um e-mail com as planilhas anexadas para os responsáveis.

## Funcionalidades

- Login automático no portal SIEG.
- Extração automática de dados sobre os certificados digitais.
- Identificação de certificados vencidos e com vencimento iminente.
- Envio de e-mails automatizados com alertas sobre o status dos certificados.

## Dependências

- `botcity-web` para a automação de tarefas web.
- `selenium` para controle do navegador.
- `pandas` para manipulação de dados.
- `webdriver-manager` para gerenciamento automático dos drivers de navegador.

## Documentação

A documentação para as bibliotecas usadas pode ser encontrada nos seguintes links:

- [BotCity Web](https://botcity.dev/docs/libraries/web/)
- [Selenium](https://selenium-python.readthedocs.io/)
- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [Webdriver Manager](https://github.com/SergeyPirogov/webdriver_manager)

## Exemplos

A aplicação é bastante específica e depende de informações sensíveis, portanto, exemplos detalhados de uso não são fornecidos.

## Resolução de Problemas

Verifique se todas as dependências estão corretamente instaladas e que as credenciais de acesso ao portal SIEG e ao servidor SMTP estão atualizadas. Para erros relacionados ao navegador ou ao Selenium, assegure-se de que a versão do driver corresponde à versão do navegador instalado.