from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.serialization.pkcs12 import load_key_and_certificates
import os
from datetime import datetime, timedelta
import re
import pandas as pd
import shutil
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

contador_validos = 0
contador_validos_proximo = 0
contador_invalidos = 0

def mapear_arquivos_txt(pasta):
    txt_map = {}
    for root, dirs, files in os.walk(pasta):
        if 'vencido' in root.lower():
            continue
        for file in files:
            if file.endswith('.txt') or not re.match(r'.*\.[a-zA-Z]{3}$', file):
                if root not in txt_map:
                    txt_map[root] = []
                txt_map[root].append(file)
    return txt_map
    
    
def search_password(texto):
    senhas_fixas = ['1234', '123456', '12345678']
    # Inicializando o dicionário de resultados
    resultado_senha = []
    
    # Busca por padrões que incluam "senha", seguidos por caracteres como : - ou espaços, e então captura a senha
    padrao = re.compile(r'senha[s]?[\s:_-]*\s*(\S+)|(\S+)\s*senha[s]?[\s:_-]*\s*(\S+)', re.IGNORECASE)
    resultados = padrao.findall(texto)
    
    for resultado in resultados:
        for possivel_senha in filter(None, resultado):
            if possivel_senha not in senhas_fixas and possivel_senha not in resultado_senha:
                resultado_senha.append(possivel_senha)
            
    # Capturando a primeira e a última palavra do texto
    palavras = texto.split()
    if palavras:
        primeira_palavra, ultima_palavra = palavras[0], palavras[-1]
        if primeira_palavra not in senhas_fixas:
            resultado_senha.append(primeira_palavra)
        if ultima_palavra not in senhas_fixas and ultima_palavra != primeira_palavra:
            resultado_senha.append(ultima_palavra)
            
    return resultado_senha
    
    
def inserir_senhas(caminho_completo, pasta_vencidos, conteudo_certificado, senhas):
    resultados = []
    senha_encontrada = False
    
    global contador_validos
    global contador_validos_proximo
    global contador_invalidos
    
    for senha in senhas:
        try:
            chave_privada, certificado, cadeia_certificados = load_key_and_certificates(
                conteudo_certificado, senha.encode(), default_backend()
            )
            if certificado:
                data_criacao = certificado.not_valid_before
                um_ano_depois_criacao = data_criacao + timedelta(days=365)
                hoje = datetime.now()
                
                dias_validade = um_ano_depois_criacao - hoje
                validade = hoje - data_criacao < timedelta(days=365)
                
                if not validade:
                    diretorio, arquivo = os.path.split(caminho_completo)
                    _, pasta_imediata = os.path.split(diretorio)
                    
                    destino_vencido_pasta = os.path.join(pasta_vencidos, pasta_imediata)
                    os.makedirs(destino_vencido_pasta, exist_ok=True)
                    
                    destino_vencido = os.path.join(destino_vencido_pasta, arquivo)
                    
                    shutil.move(caminho_completo, destino_vencido)
                    # caminho_completo = destino_vencido
                    
                    status = 'Vencido'
                    mensagem = 'Movido para pasta "Vencidos"'
                    contador_invalidos += 1
                    
                else:
                    mensagem = f'{dias_validade.days} dias de validade'
                    status = 'Valido'
                    contador_validos += 1
                    
                    if dias_validade.days <= 30:
                        mensagem = f'{dias_validade.days} dias de validade restantes. **ATENÇÃO**'
                        contador_validos_proximo += 1
                    
                resultados.append((caminho_completo, status, mensagem))
                senha_encontrada = True
                break
        except ValueError:
            continue
        
    return resultados, senha_encontrada
                                

def verificar_certificados_com_senhas_recursivo(caminho_pasta, senhas):
    pasta_vencidos = os.path.join(caminho_pasta, 'Vencidos')
    os.makedirs(pasta_vencidos, exist_ok=True)
        
    resultados = []

    for diretorio, subpastas, arquivos in os.walk(caminho_pasta):
        if 'vencido' in diretorio.lower():
            continue
        arquivos_certificados = [arq for arq in arquivos if arq.endswith('.p12') or arq.endswith('.pfx')]

        for arquivo in arquivos_certificados:
            caminho_completo = os.path.join(diretorio, arquivo)
            senha_encontrada = False

            with open(caminho_completo, 'rb') as f:
                conteudo_certificado = f.read()
            
            res, senha_encontrada = inserir_senhas(caminho_completo, pasta_vencidos, conteudo_certificado, senhas)
            resultados.extend(res)

            if not senha_encontrada:
                
                nome_cert_sem_extensao, _ = os.path.splitext(arquivo)
                senhas_potenciais = search_password(nome_cert_sem_extensao)
                
                res, senha_encontrada = inserir_senhas(caminho_completo, pasta_vencidos, conteudo_certificado, senhas_potenciais)
                resultados.extend(res)
                
                if not senha_encontrada and diretorio in txt_map:
                    for file_txt in txt_map[diretorio]:
                        nome_txt_sem_extensao, _ = os.path.splitext(file_txt)
                        
                        senhas_potenciais_txt = search_password(nome_txt_sem_extensao)
                        
                        res, senha_encontrada = inserir_senhas(caminho_completo, pasta_vencidos, conteudo_certificado, senhas_potenciais_txt)
                        resultados.extend(res)
                
                if not senha_encontrada:
                    resultados.append((caminho_completo, 'Falha', 'Senha não encontrada'))

    return resultados
    
    
def extrair_numero(texto):
    """
    Tenta extrair um número do início da string.
    Retorna um valor numérico muito alto se não houver número, para que strings sem números
    fiquem ordenadas após as strings com números.
    """
    match = re.match(r'(\d+)', texto)
    if match:
        return int(match.group(1))
    else:
        return float('inf') 
    
    
def save_excel_with_formatting(resultados, filepath):
    df = pd.DataFrame(resultados, columns=['Caminho do Certificado', 'Status', 'Mensagem'])
    
    df['NumeroOrdenacao'] = df['Mensagem'].apply(extrair_numero)
    df = df.sort_values(by=['NumeroOrdenacao', 'Status'], ascending=[True, False])
    df = df.drop('NumeroOrdenacao', axis=1)
    
    writer = pd.ExcelWriter(filepath, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    worksheet = writer.sheets['Sheet1']

    # Adicionando uma tabela com estilo
    (max_row, max_col) = df.shape
    column_settings = [{'header': column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 6'})

    # Ajustando o espaçamento das colunas
    for i, col in enumerate(df.columns):
        worksheet.set_column(i, i, 120)
        if col in ['Status']:
            worksheet.set_column(i, i, 10)
        if col in ['Mensagem']:
            worksheet.set_column(i, i, 45)

    # Fechar o writer e salvar o arquivo Excel
    writer.close()
            
    
def enviar_email_com_anexo(email_remetente, senha, email_destinatario, arquivo_anexo):
    print('Enviando email.')
    global contador_validos
    global contador_validos_proximo
    global contador_invalidos
    
    # Criar a mensagem de e-mail
    msg = MIMEMultipart('alternative')
    msg['From'] = email_remetente
    msg['To'] = email_destinatario
    msg['Subject'] = "Atualização Quinzenal do Status dos Certificados Digitais"

    corpo_mensagem = f"""
    Prezada TopCert,<br><br>

    Este e-mail é parte da nossa verificação regular quinzenal para manter a segurança e a eficiência das operações digitais. Estou enviando esta atualização para informar sobre o status atual dos certificados digitais.<br><br>

    Como parte da rotina de manutenção, verificamos todos os certificados digitais e compilamos as informações relevantes no anexo deste e-mail. É importante revisar este documento para garantir que todos os certificados estejam atualizados e válidos.<br><br>

    <strong>Resumo do Status:</strong><br>
    - Certificados Válidos: {contador_validos}<br>
    - Certificados Próximos do Vencimento: {contador_validos_proximo}<br>
    - Certificados Vencidos: {contador_invalidos}<br><br>

    Agradeço sua atenção a este assunto e estou à disposição para qualquer suporte adicional.<br><br>

    Atenciosamente,<br>
    Davi - Inovação
    """
    
    # Anexando o corpo do e-mail à mensagem
    msg.attach(MIMEText(corpo_mensagem, 'html'))
    
    # Anexar o arquivo
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(arquivo_anexo, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(arquivo_anexo)}"')
    msg.attach(part)

    # Conectar ao servidor do Gmail e enviar o e-mail
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(email_remetente, senha)
    server.sendmail(email_remetente, email_destinatario, msg.as_string())
    server.quit()
    print('Email enviado!')
    

# Preenche os campos vazios
senhas = [' ']

caminho_pasta = r'G:\Certificados\CERTIFICADOS NOVOS'

# Primeiro, mapeia todos os arquivos .txt relevantes em cada diretório
txt_map = mapear_arquivos_txt(caminho_pasta)

resultados_validade = verificar_certificados_com_senhas_recursivo(caminho_pasta, senhas)

results_path = os.path.join(caminho_pasta, 'resultados_certificados.xlsx')
save_excel_with_formatting(resultados_validade, results_path)

# Preenche os campos vazios
enviar_email_com_anexo(' ', ' ', ' ', results_path)
