# Extrator de dados importantes DCTFWEB

## Introdução
Este projeto contém um script Python denominado `main.py`, desenvolvido para automatizar a coleta de informações financeiras de empresas a partir da plataforma Sieg IRIS. O foco está nos dados relacionados à DCTFWeb. O script utiliza a biblioteca `botcity.web` para navegar na web, coletar dados necessários e organizar essas informações em um DataFrame, utilizando a biblioteca `pandas`. Finalmente, os dados são salvos em um arquivo Excel com formatação específica.

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
Para utilizar o script, é necessário instalar as dependências listadas na seção [Dependências](#dependências). Isso pode ser feito executando o seguinte comando no terminal:

```bash
pip install botcity-web pandas webdriver-manager python-dateutil
```

## Uso
Para executar o script, navegue até o diretório onde o arquivo `bot.py` está localizado e execute o comando:

```bash
python bot.py
```

O script iniciará o processo de coleta de dados automaticamente. Certifique-se de ter configurado o email e a senha utilizados para o login, bem como os caminhos dos arquivos de banco de dados e saída, dentro do script.

## Funcionalidades
- Configuração automática do WebDriver do Chrome.
- Login na plataforma Sieg IRIS.
- Navegação até a seção de DCTFWeb.
- Coleta de dados financeiros do último mês para várias empresas.
- Organização dos dados em um DataFrame do pandas.
- Salvamento dos dados em um arquivo Excel com formatação.

## Dependências
- `botcity-web`
- `pandas`
- `webdriver-manager`
- `python-dateutil`

## Configuração
Antes de executar o script, certifique-se de configurar as seguintes variáveis no método `action`:

- `email`: E-mail utilizado para o login na plataforma.
- `password`: Senha utilizada para o login.
- `database_path`: Caminho do arquivo Excel contendo os CNPJs das empresas.
- `exit_path`: Caminho onde o arquivo Excel com os dados coletados será salvo.

## Documentação
Para mais informações sobre as bibliotecas utilizadas, consulte:

- [BotCity Web](https://botcity.dev/documentation/botcity-web/)
- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [WebDriver Manager for Python](https://github.com/SergeyPirogov/webdriver_manager)
- [Python Dateutil](https://dateutil.readthedocs.io/en/stable/)

## Exemplos
Não aplicável, pois o script é executado conforme configurado e não requer interação adicional do usuário além do comando de execução.

## Solução de Problemas
Para solucionar problemas comuns, verifique se todas as dependências foram instaladas corretamente e se os caminhos dos arquivos estão corretos e acessíveis.