# Extrator de Dados de PDF e Gerador de Excel

## Introdução
Este projeto fornece uma aplicação GUI que extrai dados específicos de arquivos PDF e gera um relatório detalhado em Excel. Ele é projetado para ler documentos PDF, identificar e extrair pontos de dados predeterminados e, em seguida, organizar essas informações em uma planilha Excel. A aplicação é particularmente útil para processar documentos financeiros, extrair figuras específicas relacionadas a salários, INSS e FGTS, e compilá-los em um arquivo Excel estruturado para análise adicional.

## Índice
1. [Instalação](#instalação)
2. [Uso](#uso)
3. [Recursos](#recursos)
4. [Dependências](#dependências)
5. [Configuração](#configuração)
6. [Documentação](#documentação)
7. [Exemplos](#exemplos)
8. [Resolução de Problemas](#resolução-de-problemas)

## Instalação
Para instalar as bibliotecas necessárias para este projeto, execute o seguinte comando:

```bash
pip install PyMuPDF pandas tkinter
```

## Uso
Para usar a aplicação, siga estes passos:
1. Execute o script \`main.py\`.
2. Use a GUI para selecionar o diretório contendo os arquivos PDF que deseja processar.
3. Clique em "Enviar" para iniciar o processo de extração e geração do relatório.
4. Uma vez concluído, a aplicação salvará um arquivo Excel no diretório selecionado contendo os dados extraídos.

## Recursos
- GUI para fácil seleção de diretório.
- Detecção e extração automáticas de pontos de dados específicos de arquivos PDF.
- Geração de um relatório Excel bem estruturado.
- Formatação de Excel personalizável.

## Dependências
- PyMuPDF
- pandas
- tkinter

## Configuração
Nenhuma configuração adicional é necessária para executar esta aplicação além da instalação de suas dependências.

## Documentação
Atualmente, a documentação se limita a este arquivo README. Mais documentação será adicionada conforme o projeto evolui.

## Exemplos
Para exemplos de uso e instruções mais detalhadas, consulte a seção [Uso](#uso).

## Resolução de Problemas
Se você encontrar problemas com arquivos PDF que não estão sendo processados corretamente, certifique-se de que eles não estão criptografados e estão em um formato que o PyMuPDF possa ler. Para outros problemas, verificar a saída do console para mensagens de erro pode fornecer orientações sobre como resolvê-los.