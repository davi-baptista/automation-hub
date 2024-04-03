# Processador de Extratos Bancários

## Introdução

Este script em Python é projetado para processar extratos bancários fornecidos em formato Excel, realizar tratamentos de dados específicos e salvar os resultados em um novo arquivo Excel com formatações personalizadas. Ele lida com a limpeza de linhas com valores NaN, ajuste de datas, concatenação de colunas para formar um campo histórico, cálculo de saldo por título/parcela, e finalmente, geração de um DataFrame contendo movimentações financeiras formatadas para exportação.

## Índice

- [Instalação](#instalação)
- [Uso](#uso)
- [Características](#características)
- [Dependências](#dependências)
- [Documentação](#documentação)

## Instalação

Para usar este script, você precisa ter o Python instalado em sua máquina. É recomendável usar o Python 3.8 ou superior. Além disso, você precisará instalar algumas bibliotecas Python:

1. Pandas
2. Tkinter
3. XlsxWriter

Você pode instalar essas dependências usando pip:

```bash
pip install pandas xlsxwriter
```

### Usando `requirements.txt`
Um arquivo `requirements.txt` é fornecido para facilitar a instalação das dependências. Para instalar todas as dependências de uma só vez, execute o seguinte comando no terminal:

```bash
pip install -r requirements.txt
```

Isso instalará automaticamente as bibliotecas necessárias, incluindo Pandas, Tkinter (se necessário), e XlsxWriter.

Nota: Tkinter geralmente vem pré-instalado com Python. Caso contrário, consulte a documentação específica para sua plataforma.

## Uso

Para utilizar este script, siga os passos abaixo:

1. Abra o terminal ou prompt de comando.
2. Navegue até o diretório onde o script `main.py` está localizado.
3. Execute o script com o comando:

```bash
python main.py
```

4. Uma interface gráfica será aberta. Clique em "Buscar" para selecionar o arquivo Excel contendo o extrato bancário.
5. Após selecionar o arquivo, clique em "Enviar" para iniciar o processamento.
6. O script irá processar o arquivo e salvar um novo arquivo Excel no mesmo diretório do arquivo original, com o nome `mov_financeiro.xlsx`.

## Características

- **Limpeza de Dados:** Remove linhas com valores NaN nas colunas especificadas.
- **Ajuste de Datas:** Garante a continuidade das datas entre transações.
- **Concatenação de Colunas:** Forma um campo histórico único a partir de várias colunas.
- **Cálculo de Saldo:** Identifica pagamentos e recebimentos baseando-se no saldo de títulos/parcelas.
- **Exportação Formatada:** Gera um arquivo Excel com as movimentações financeiras devidamente formatadas.

## Dependências

- Python 3.8+
- Pandas
- Tkinter
- XlsxWriter

## Documentação

Para mais informações sobre as bibliotecas utilizadas, consulte:

- [Pandas](https://pandas.pydata.org/pandas-docs/stable/)
- [Tkinter](https://docs.python.org/3/library/tkinter.html)
- [XlsxWriter](https://xlsxwriter.readthedocs.io/)