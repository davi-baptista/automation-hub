from botcity.web import WebBot, Browser, By
from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import smtplib
from datetime import datetime
from dateutil.relativedelta import relativedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
import os


class MyWebBot(WebBot):


    def action(self, execution=None):
        key_defis = ["DCTF"]
        key_diretoras = [
            "ECF",
            "Termo de Intimação",
            "Ciência do Processo/Procedimento",
            "Comunicado Cadin",
            "DTE",
            "Redarf Net",
            "PER/DCOMP",
            "SIMPLES NACIONAL",
            "Notificação prévia visando à autorregularização",
            "Lei 12.996/2014",
            "Parcelamento",
            "Comunicação para Compensação de Ofício",
            "MALHA FISCAL"
        ]
        key_fernanda = [
            "GFIP",
            "Termo de Intimação",
            "PARCELAMENTOS PREVIDENCIÁRIOS PASSÍVEIS DE RESCISÃO"
        ]
        key_deleg = ["AVISO PARA REGULARIZAÇÃO DE OBRA"]
        
        filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Compliance\01_ecac_mensagens_importantes\data'
        
        email = ' '
        password = ' '
        
        email_defis = ' '
        email_diretoras = ' '
        email_fernanda = ' '
        email_deleg = ' '
        email_restos = ' '
    
        corpo_defis = "Prezados, bom dia!<br><br><strong>ATENÇÃO!</strong><br>Segue em anexo relação das empresas cujas DCTF´s foram recepcionadas pela RFB.<br><br><br>Observação: Este e-mail será disparado semanalmente e enviará todas as mensagens importantes que encontrar até um mês antes."
        
        self.configure(filepath)
        self.browse("https://hub.sieg.com/IriS/#/")
        self.maximize_window()
        self.login(email, password)
        download_file = self.download_file(filepath)
        files_defis, files_director, files_fernanda, files_deleg, arquivos_restos = self.create_excels(download_file, key_defis, key_diretoras, key_fernanda, key_deleg)
        
        self.send_email(email_defis, files_defis, corpo_defis)
        self.send_email(email_diretoras, files_director)
        self.send_email(email_fernanda, files_fernanda)
        self.send_email(email_deleg, files_deleg)
        self.send_email(email_restos, arquivos_restos)
        
        
    def configure(self, filepath):
        # Configurando o ChromeDriver
        self.driver_path = ChromeDriverManager().install()
        self.options = default_options(headless=False, download_folder_path=filepath, user_data_dir=None, page_load_strategy='normal')
    
    
    def login(self, email, password):
        print('Fazendo login')
        campo_email = self.find_element('txtEmail', By.ID, waiting_time=10000)
        campo_email.send_keys(email)
        
        campo_senha = self.find_element('txtPassword', By.ID, waiting_time=10000)
        campo_senha.send_keys(password)
        self.wait(200)
        self.enter()


    def download_file(self, filepath):
        print('Baixando excel')
        caixa_postal = self.find_element('//*[@id="iriS"]/div[2]/div/div[3]/div/div/div[2]/div[40]/a/a/span[2]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(caixa_postal, visible=True, waiting_time=20000)
        caixa_postal.click()
        
        self.find_element('//*[@id="tableMessage"]/tbody/tr[1]/td[3]', By.XPATH, waiting_time=60000)
        self.wait(1000)
        
        exportar_mensagens = self.find_element('//*[@id="btnExportExcel"]/span', By.XPATH, waiting_time=30000)
        exportar_mensagens.click()
        
        self.wait_for_new_file(path=filepath, file_extension=".crdownload", timeout=300000)
        while True:
            last_file = self.get_last_created_file(path=filepath, file_extension='.xlsx')
            current_date = datetime.now().strftime("%d-%m-%Y")
            
            if current_date in last_file:
                print(last_file)
                break
            self.wait(1000)
            
        print('Excel baixado')
        return last_file


    def safe_file_name(self, nome):
        for char in ["/", "\\", ":", "*", "?", "\"", "<", ">", "|"]:
            nome = nome.replace(char, "_")
        return nome
    
    
    def save_excel_with_formatting(self, df, filename):
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Criando um formato de célula com alinhamento centralizado
        center_format = workbook.add_format({'align': 'center'})
        
        # Adicionando uma tabela com estilo
        (max_row, max_col) = df.shape
        column_settings = [{'header': column} for column in df.columns]
        worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 6'})

        # Ajustando o espaçamento das colunas
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, 60)
            if col in ['Destinatário', 'Enviada em']:
                # Aplicando o formato centralizado a colunas específicas
                worksheet.set_column(i, i, 30, center_format)

        # Fechar o writer e salvar o arquivo Excel
        writer.close()
        

    def create_excels(self, filepath, key_defis, key_diretoras, key_fernanda, key_deleg):
        print('Checando excel')
        files_defis, files_director, files_fernanda, files_deleg, files_remains = [], [], [], [], []

        # Ler o arquivo Excel
        df = pd.read_excel(filepath, skiprows=1, dtype={'Destinatário': str})
        df = df[['Empresa', 'Assunto da Mensagem', 'Destinatário', 'Enviada em']]
        df['Enviada em'] = pd.to_datetime(df['Enviada em'], format='%d/%m/%Y')
        
        data_atual = datetime.now()
        um_mes_atras = data_atual - relativedelta(months=1)
        
        # Filtrando o DataFrame para datas no último mês
        df = df[(df['Enviada em'] >= um_mes_atras) & (df['Enviada em'] <= data_atual)]
        
        df['Enviada em'] = df['Enviada em'].dt.strftime('%d/%m/%Y')
        
        exit_filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Compliance\01_ecac_mensagens_importantes\output'
        os.makedirs(exit_filepath, exist_ok=True)
        
        df['Assunto da Mensagem'] = df['Assunto da Mensagem'].str.replace(r'\(s\)', '', regex=True)
        key_words = key_defis + key_diretoras + key_fernanda + key_deleg
        
        for keyword in key_words:
            group = df[df['Assunto da Mensagem'].str.contains(keyword, na=False, case=False)]
            if not group.empty:
                group = group.sort_values('Empresa')
                filename = self.safe_file_name(keyword) + ".xlsx"
                full_path = os.path.join(exit_filepath, filename)
                self.save_excel_with_formatting(group, full_path)
                
                if keyword in key_defis:
                    files_defis.append(full_path)
                if keyword in key_diretoras:
                    files_director.append(full_path)
                if keyword in key_fernanda:
                    files_fernanda.append(full_path)
                if keyword in key_deleg:
                    files_deleg.append(full_path)
                
                # Remover as linhas processadas do DataFrame original
                df = df.drop(group.index)
                
        remains = df
        if not group.empty:
            remains.to_excel(exit_filepath + r'\Restos.xlsx', index=False)
            files_remains.append(remains)
            
        print('Excel checado')
        return files_defis, files_director, files_fernanda, files_deleg, files_remains
    
    
    def send_email(self, destinatario, excel_files, corpo_email="Prezados, bom dia!<br><br><strong>ATENÇÃO!</strong><br>Segue em anexo relação das empresas que possuem mensagens importantes na Caixa Postal do ECAC – RFB para análise. <br><br><strong>AÇÃO</strong>: Planejar a resolução e <strong>observar os prazos</strong> a serem cumpridos conforme cada assunto!<br><br><br>Observação: Este e-mail será disparado semanalmente e enviará todas as mensagens importantes que encontrar até um mês antes."):
        print('Mandando email')
        
        if not excel_files:
            return
        
        sender = ' '
        password = ' '

        servidor_smtp = 'smtp.gmail.com'
        porta_smtp = 587

        subject = '⚠️ Mensagens importantes ECAC'

        message = MIMEMultipart('alternative')
        message.attach(MIMEText(corpo_email, 'html'))

        message['From'] = sender
        message['To'] = destinatario
        message['Subject'] = subject

        # Anexar os arquivos Excel
        for filepath in excel_files:
            with open(filepath, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)
            filename = os.path.basename(filepath)
            header_filename = Header(filename, 'utf-8').encode()
            part.add_header("Content-Disposition", f"attachment; filename={header_filename}")
            message.attach(part)

        with smtplib.SMTP(servidor_smtp, porta_smtp) as servidor:
            servidor.starttls()
            servidor.login(sender, password)
            servidor.send_message(message)
            
        print('Email enviado!')
            
            
    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    MyWebBot.main()