# Automação de Certificados e Procurações SPE

## Introdução

Este projeto contém dois scripts Python destinados a automatizar o processo de manipulação de certificados digitais e gestão de procurações no Sistema de Procurações Eletrônicas (SPE). Utiliza-se de bibliotecas como `botcity`, `pyautogui`, `pandas`, `openpyxl` e `pywinauto` para facilitar a busca, download, validação de certificados digitais e o gerenciamento automático de procurações.

## Índice

- [Introdução](#introdução)
- [Preparação](#preparação)
- [Instalação](#instalação)
- [Uso](#uso)
  - [Baixar Certificados](#baixar-certificados)
  - [Gerenciar Procurações](#gerenciar-procurações)
- [Funcionalidades](#funcionalidades)
- [Dependências](#dependências)
- [Configuração](#configuração)
- [Solução de Problemas](#solução-de-problemas)

## Preparação

Antes de executar os scripts, é essencial que não existam certificados previamente instalados na máquina. Para remover quaisquer certificados existentes, siga os passos abaixo:

1. Pressione `Win + R` e digite `certmgr.msc` para abrir o gerenciador de certificados.
2. Navegue pelas pastas `Pessoal -> Certificados`.
3. Selecione os certificados listados, clique com o botão direito e escolha `Excluir`.
4. Confirme a remoção quando solicitado.

Esta etapa garante que o script de download de certificados possa operar sem interferências de certificados preexistentes.

## Instalação

Certifique-se de ter Python instalado em seu sistema. As dependências podem ser instaladas utilizando o seguinte comando pip:

```bash
pip install botcity pyautogui pandas openpyxl pywinauto
```

## Uso

### Baixar Certificados

O script `main.py` gerencia a busca, download e validação de certificados digitais. Ele deve ser executado primeiro:

```bash
python main.py
```

Durante sua execução, ele chama automaticamente o script `procuracao_spe.py` (referido aqui como main) para cada certificado processado, iniciando o fluxo de criação e gestão de procurações no SPE.

### Gerenciar Procurações

Embora o script `procuracao_spe.py` seja invocado automaticamente pelo `baixar-certificado.py`, ele também pode ser executado de forma independente para gerenciamento de procurações:

```bash
python procuracao_spe.py
```

Nota: Se atente que para rodar este script é necessario ter o certificado instalado dentro da maquina e caso tenha mais de um o script não sabera qual escolher para assinar.

## Funcionalidades

- **Baixar Certificados**: Realiza a busca, download e validação de certificados digitais.
- **Gerar Procurações**: Automatiza a criação de procurações no SPE.
- **Assinar Procurações**: Automatiza a assinatura eletrônica das procurações.
- **Gerenciamento de Procurações**: Facilita a exclusão e download de procurações existentes.

## Dependências

- botcity
- pyautogui
- pandas
- openpyxl
- pywinauto

## Configuração

Revise e configure as variáveis relacionadas a caminhos de diretórios, CNPJ, e-mail, etc., diretamente nos scripts, conforme as necessidades específicas do seu ambiente e dos processos que deseja automatizar.

## Solução de Problemas

Para qualquer problema relacionado ao funcionamento do bot, verifique se todas as dependências estão corretamente instaladas e se os caminhos para arquivos e diretórios estão configurados corretamente.