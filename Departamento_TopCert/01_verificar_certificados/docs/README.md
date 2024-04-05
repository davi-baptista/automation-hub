# Validador de certificado digital em lote

## Introdução

O projeto é uma aplicação Python que incorpora uma série de funcionalidades relacionadas à criptografia, manipulação de arquivos e diretórios, processamento de datas, uso de expressões regulares, manipulação de dados com pandas, envio de e-mails com anexos, e manipulação de arquivos Excel com openpyxl. Este conjunto de ferramentas oferece uma solução abrangente para automatizar o processamento de dados e comunicação em contextos que requerem segurança, organização e eficiência.

## Índice

- [Introdução](#introdução)
- [Funcionalidades](#funcionalidades)
- [Instalação](#instalação)
- [Uso](#uso)
- [Dependências](#dependências)
- [Configuração](#configuração)
- [Documentação](#documentação)
- [Exemplos](#exemplos)
- [Solução de Problemas](#solução-de-problemas)

## Funcionalidades

- **Criptografia de Dados**: Utiliza a biblioteca \`cryptography\` para garantir a segurança dos dados processados.
- **Manipulação de Arquivos e Diretórios**: Cria, lê e escreve em arquivos e diretórios, permitindo a organização eficiente dos dados.
- **Processamento de Datas**: Manipula informações de data e hora para rastreamento e agendamento de tarefas.
- **Expressões Regulares**: Filtra e processa textos com expressões regulares para uma análise de dados mais precisa.
- **Manipulação de Dados com Pandas**: Usa a biblioteca \`pandas\` para leitura, manipulação e análise de dados estruturados.
- **Envio de E-mails com Anexos**: Automatiza o envio de e-mails com anexos, facilitando a comunicação e a distribuição de informações.
- **Manipulação de Arquivos Excel**: Utiliza \`openpyxl\` para ler e escrever em arquivos Excel, permitindo uma integração suave com ferramentas de escritório.

## Instalação

Para instalar as dependências necessárias para executar este projeto, você precisará de um ambiente Python. É recomendável usar um ambiente virtual para evitar conflitos de dependências. As dependências específicas podem ser instaladas via pip:

```bash
pip install cryptography pandas openpyxl smtplib email
```

## Uso

O arquivo `main.py` é o ponto de entrada do projeto. Para executar o projeto, navegue até o diretório contendo `main.py` e execute o seguinte comando no terminal:

```bash
python main.py
```

## Dependências

As principais dependências do projeto incluem:

- `cryptography`: Para a criptografia de dados.
- `pandas`: Para manipulação e análise de dados.
- `openpyxl`: Para leitura e escrita de arquivos Excel.
- `smtplib` e `email`: Para o envio de e-mails com anexos.

## Configuração

Certifique-se de ajustar os caminhos para os certificados, emails e senhas padrões dentro do script para corresponder ao seu ambiente de execução.

## Documentação

É recomendado revisar o código fonte para entender melhor cada função e sua finalidade.

## Exemplos

Suponha que você queira automatizar a coleta de informações sobre a validade de certificados digitais dentro de uma pasta, navegando para suas subpastas, e retorne se esta valido e quantos dias para acabar a validade. O script principal navegará até a pasta desejada, entrará em cada certificado digital, colocará a senha e coletará as informações necessárias, salvando os resultados em um arquivo excel na mesma pasta e criando uma nova pasta 'Vencidos' com todos os certificado digitais encontrados que estão vencidos. Caso ele não entre no certificado com as senhas padrões, ele procurará dentro da pasta se existe alguma senha informada e tentará usar ela, caso não consiga, avisará falha.

## Solução de Problemas

Em caso de erros durante a execução, consulte as mensagens de erro fornecidas pelo programa e verifique os logs de erro dentro do excel.
Lembre-se de alterar os caminhos, senhas e emails dentro do codigo fonte