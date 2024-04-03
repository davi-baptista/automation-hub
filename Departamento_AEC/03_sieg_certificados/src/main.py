from botcity.web import WebBot, Browser, By
from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
import pandas as pd
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os


class MyWebBot(WebBot):


    def action(self, execution=None):   
        email = ' '
        password = ' '
        
        entrada = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\03_sieg_certificados\data'
        os.makedirs(entrada, exist_ok=True)
        
        saida = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\03_sieg_certificados\output'
        os.makedirs(saida, exist_ok=True)
        
        self.configure(entrada)
        self.browse("https://hub.sieg.com/")
        self.maximize_window()
        self.login(email, password)
        self.navigate_to_service()
        filepath = self.download_file(entrada)
        excel_files = self.create_excel(filepath, saida)
        self.send_email(excel_files) #sucessodocliente@controller-rnc.com.br
        
        
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


    def navigate_to_service(self):
        print('Abrindo serviços')
        todos_servicos = self.find_element('//*[@id="sieg-header"]/div/div[1]/a[2]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(todos_servicos, visible=True, waiting_time=20000)
        todos_servicos.click()
        gerenciar_certificados = self.find_element('//*[@id="rowHubServices"]/div[8]/a/div[2]', By.XPATH, waiting_time=10000)
        gerenciar_certificados.click()


    def download_file(self, entrada):
        print('Baixando excel')
        exportar_excel = self.find_element('//*[@id="ctl00"]/div[11]/div[2]/div/div[2]/div/div/div/div[1]/div/div/div[1]/div[1]/div[2]/a[3]', By.XPATH, waiting_time=60000)
        exportar_excel.click()
        self.wait_for_new_file(path=entrada, file_extension=".crdownload", timeout=300000)
        
        while True:
            condition = False
            for arquivo in os.listdir(entrada):
                path_completo = os.path.join(entrada, arquivo)
                if os.path.isfile(path_completo) and arquivo.endswith('.xlsx'):
                    condition = True
            
            if condition == False:
                continue
            
            last_file = self.get_last_created_file(path=entrada, file_extension='.xlsx')
            data_atual = datetime.now().strftime("%d-%m-%Y")
            
            if data_atual in last_file:
                print(last_file)
                break
            self.wait(1000)
        return last_file


    def create_excel(self, filepath, saida):
        print('Checando excel')
        planilha = pd.read_excel(filepath, dtype={'CPF_CNPJ': str})
        arquivos_criados = []
        colunas_selecionadas = ['Nome', 'CPF_CNPJ', 'Vencimento', 'Status']

        def save_excel_with_formatting(df, filename):
            writer = pd.ExcelWriter(filename, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            
            worksheet = writer.sheets['Sheet1']

            # Adicionando uma tabela com estilo
            (max_row, max_col) = df.shape
            column_settings = [{'header': column} for column in df.columns]
            worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 6'})

            # Ajustando o espaçamento das colunas
            for i, col in enumerate(df.columns):
                # Aqui você pode ajustar a largura como preferir
                worksheet.set_column(i, i, 45)

            # Fechar o writer e salvar o arquivo Excel
            writer.close()
        
        linhas_vencidas = planilha[planilha['Status'] == 'Vencido'].copy()
        if not linhas_vencidas.empty:
            filename = f'{saida}\\Certificados-Vencidos.xlsx'
            linhas_vencidas = linhas_vencidas[colunas_selecionadas]
            save_excel_with_formatting(linhas_vencidas, filename)
            arquivos_criados.append(filename)

        linhas_a_vencer = planilha[planilha['Status'] == 'À Vencer'].copy()
        if not linhas_a_vencer.empty:
            data_futura = datetime.now() + pd.Timedelta(days=10)
            linhas_a_vencer = linhas_a_vencer[colunas_selecionadas]
            linhas_a_vencer['Vencimento'] = pd.to_datetime(linhas_a_vencer['Vencimento'], dayfirst=True)
            linhas_a_vencer = linhas_a_vencer[linhas_a_vencer['Vencimento'] <= data_futura]

            if not linhas_a_vencer.empty:
                filename = f'{saida}\\Vencer-Em-Breve.xlsx'
                linhas_a_vencer['Vencimento'] = linhas_a_vencer['Vencimento'].dt.strftime('%d/%m/%Y')
                save_excel_with_formatting(linhas_a_vencer, filename)
                arquivos_criados.append(filename)

        print('Excel checado')
        print(arquivos_criados)
        return arquivos_criados
    
    
    def send_email(self, excel_files):
        print('Mandando email')
        destinatarios = ' '
        remetente = ' '
        senha = ' '

        servidor_smtp = 'smtp.gmail.com'
        porta_smtp = 587

        assunto = 'Alerta de certificados SIEG'
        corpo_email = 'Prezados, bom dia<br><br><strong>ATENÇÃO</strong>: Segue relação de empresas com certificados vencidos e/ou próximos ao vencimento no SIEG.<br> Áreas: <strong>TOPCERT</strong>, <strong>DETEC</strong> e <strong>AEC</strong> sigam com as ações necessárias para evitar a interrupção dos processos da plataforma.<br><br> Observação: Este e-mail será disparado semanalmente à medida que os certificados estejam vencidos e/ou próximos do vencimento na plataforma.<br><br><br> Atenciosamente.'

        mensagem = MIMEMultipart('alternative')
        mensagem.attach(MIMEText(corpo_email, 'html'))

        mensagem['From'] = remetente
        mensagem['To'] = destinatarios
        mensagem['Subject'] = assunto

        # Anexar os arquivos Excel
        for filename in excel_files:
            with open(filename, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())

            encoders.encode_base64(part)
            partes = filename.split('\\')
            part.add_header("Content-Disposition", f"attachment; filename={partes[-1]}")
            mensagem.attach(part)

        with smtplib.SMTP(servidor_smtp, porta_smtp) as servidor:
            servidor.starttls()
            servidor.login(remetente, senha)
            servidor.send_message(mensagem)

        print('Email enviado!')
            
            
    def not_found(self, label):
        print(f"Element not found: {label}")


def main():
    MyWebBot.main()


if __name__ == '__main__':
    main()