# Extrator de mensagens importantes

## Introdução
Script Python desenvolvido para automatizar a extração de mensagens importantes do eCAC, a criação de relatórios específicos em Excel com base nessas mensagens e o envio desses relatórios para endereços de e-mail pré-determinados. Utiliza a biblioteca `botcity.web` para navegação e interação web, além de outras bibliotecas para manipulação de dados e envio de e-mails.

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
Para instalar as dependências necessárias, execute o seguinte comando:

```bash
pip install botcity.web pandas smtplib datetime email webdriver_manager
```

## Uso
Para usar o AutoComplianceBot, siga estes passos:

1. Certifique-se de ter Python instalado em seu ambiente.
2. Instale todas as dependências necessárias.
3. Configure as variáveis de ambiente e/ou modifique o script para incluir seus dados de autenticação e caminhos de diretório.
4. Execute o script:

```bash
python bot.py
```

## Funcionalidades
- Extração automática de mensagens importantes do eCAC.
- Geração de relatórios em Excel com base nos assuntos das mensagens.
- Envio automático de e-mails com os relatórios anexados para destinatários específicos.

## Dependências
- botcity.web
- pandas
- smtplib
- datetime
- email
- webdriver_manager

## Configuração
Antes de executar o script, é necessário configurar algumas variáveis dentro do script, como credenciais de e-mail e caminhos de diretórios para downloads e saída de arquivos.

## Documentação
Para mais informações sobre as bibliotecas utilizadas, consulte a documentação oficial:

- [botcity.web](https://botcity.dev/docs/web)
- [pandas](https://pandas.pydata.org/docs/)
- [smtplib](https://docs.python.org/3/library/smtplib.html)

## Exemplos
O script já está configurado para executar as tarefas de navegação web, download de arquivos, geração de relatórios em Excel e envio de e-mails. Modificações podem ser necessárias para atender a especificidades do ambiente de cada usuário.

## Solução de Problemas
Para problemas com a execução do script, verifique se todas as dependências foram corretamente instaladas e se as configurações do script estão corretas e compatíveis com seu sistema.