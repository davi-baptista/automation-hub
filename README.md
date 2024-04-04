# Automation Hub

## Introdução
Este é o epicentro do Hub de Automação, um espaço dedicado onde abrigo meus projetos de automação, cada qual em seu próprio nicho de eficiência. Aqui, a paixão por simplificar o complexo dá vida a soluções inovadoras. Desenvolvidos com o poder do Python, estes projetos são mais do que meros códigos; são a chave para desbloquear uma nova dimensão de produtividade e otimização de tempo. Este documento não apenas te guia pelos recônditos deste universo de automação, mas também serve como o primeiro passo para transformar a complexidade em simplicidade. Desde a configuração inicial até a conversão de scripts em aplicativos autônomos.

## Índice

- [Introdução](#introdução)
- [Instalação](#instalação)
- [Uso](#uso)
- [Solução de Problemas](#solução-de-problemas)
- [Transformando Script em Executável](#transformando-script-em-executável)
- [Contribuidores](#contribuidores)
- [Licença](#licença)

## Instalação

Para utilizar qualquer um dos projetos contidos neste repositório, siga os passos abaixo:

1. Navegue até o README do projeto específico localizado dentro de seu respectivo departamento.
2. Siga as instruções de instalação específicas para aquele projeto, incluindo a instalação de dependências.

## Uso

O uso específico para cada projeto pode ser encontrado dentro do README individual de cada projeto. É recomendado seguir as diretrizes específicas fornecidas para evitar quaisquer problemas.

## Solução de Problemas

Ao enfrentar problemas ao utilizar os projetos, siga estas etapas para solução de problemas:

1. Garanta que todas as dependências foram instaladas corretamente e estão na versão correta conforme especificado no README de cada projeto.
2. Verifique se os caminhos dentro dos scripts estão corretamente configurados para o seu ambiente específico.
3. Consulte a seção de Solução de Problemas no README de cada projeto para soluções específicas.

## Transformando Script em Executável

Para converter os scripts Python dos projetos em arquivos executáveis (.exe), siga os passos abaixo:

1. Entre em um ambiente virtual utilizando o comando `python -m venv venv` seguido por `source venv/bin/activate` (Linux/macOS) ou `.\venv\Scripts\activate` (Windows).
2. Navegue até o diretório do projeto específico para o qual deseja criar um executável.
3. Instale as dependências necessárias conforme listadas no README do projeto.
4. Instale o PyInstaller com o comando `pip install pyinstaller`.
5. Execute o comando `pyinstaller --onefile main.py` para gerar o executável do arquivo `main.py`. Ajuste o nome do arquivo `.py` conforme necessário para outros scripts e adicione '--noconsole' caso não queira que a janela do terminal seja aberta junto do executavel.

# Licença

Este projeto está licenciado sob a Licença Pública Geral GNU versão 3 (GPL-3.0).