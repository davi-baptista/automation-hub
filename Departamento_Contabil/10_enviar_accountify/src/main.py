from botcity.web import WebBot, By
from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
import os
import pyautogui
from pynput.keyboard import Controller


class MyWebBot(WebBot):
    
    
    def action(self, execution=None):
        print('Iniciando robo, aguarde um momento...')
        email = ' '
        password = ' '
        
        filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Contabil\10_enviar_accountify\data'
        
        self.configure()
        self.browse("https://cloud.accountfy.com/#")
        self.maximize_window()
        self.login(email, password)
        self.join_group()
        valores = self.obter_valores(filepath)
        index = 0
        for valor in valores:
            index = self.sending_balancete(valor, index)
            self.wait(1000)
        
        print('Programa finalizado.')
        
        
    def configure(self):
        # Configurando o ChromeDriver
        self.driver_path = ChromeDriverManager().install()
        self.options = default_options(headless=False, user_data_dir=None, page_load_strategy='normal')
    
    
    def login(self, email, password):
        print('Abrindo Sieg IRIS e indo até a aba de DCTFWEB.')
        email_field = self.find_element('/html/body/app-root/app-login/div/div/div[1]/div/app-accy-input/div[1]/input', By.XPATH, waiting_time=60000)
        email_field.send_keys(email)
        avancar = self.find_element('/html/body/app-root/app-login/div/div/div[1]/div/div[2]/button', By.XPATH, waiting_time=10000)
        avancar.click()
        email_field = self.find_element('//*[@id="identifierId"]', By.XPATH, waiting_time=10000)
        email_field.send_keys(email)
        avancar = self.find_element('//*[@id="identifierNext"]/div/button', By.XPATH, waiting_time=10000)
        avancar.click()
        password_field = self.find_element('//*[@id="password"]/div[1]/div/div[1]/input', By.XPATH, waiting_time=30000)
        self.wait_for_element_visibility(password_field, visible=True, waiting_time=15000)
        password_field.send_keys(password)
        avancar = self.find_element('//*[@id="passwordNext"]/div/button', By.XPATH, waiting_time=10000)
        avancar.click()
    
    
    def join_group(self):
        group = self.find_element('//*[@id="groups"]/div/ul/li[1]/div/div[2]/h4', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(group, visible=True, waiting_time=15000)
        group.click()
        menu_upload = self.find_element('//*[@id="sidenav"]/ul/li[2]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(menu_upload, visible=True, waiting_time=15000)
        menu_upload.click()
    
    
    def sending_balancete(self, valores, index):
        upload = self.find_element('/html/body/app-root/app-main-layout/div/div/app-balance-sheet/div[2]/div[1]/div[1]/div[2]/div[2]/button', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(upload, visible=True, waiting_time=15000)
        upload.click()
        send_file = self.find_element('//*[@id="tab1"]/div/form/div[1]/div/div', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(send_file, visible=True, waiting_time=15000)
        send_file.click()
        self.wait(5000)
        keyboard = Controller()
        keyboard.type(valores['caminho_arquivo'])
        pyautogui.press('enter')
        send_companie = self.find_element('//*[@id="tab1"]/div/form/div[2]/div[1]/app-accy-selecao-simples-multipla/div', By.XPATH, waiting_time=10000)
        send_companie.click()
        
        companies = self.find_elements('dropdown-item', By.CLASS_NAME)
        
        achou = False
        for companie in companies:
            print(companie.get_attribute('textContent'))
            print(valores['nome_empresa'].upper())
            if valores['nome_empresa'].upper() in companie.get_attribute('textContent').upper():
                companie.click()
                achou = True
                break
        
        if not achou:
            print('Não achou empresa com este nome')
            input()
        
        calendar = self.find_element('//*[@id="tab1"]/div/form/div[2]/div[2]/div/app-accy-datepicker/dp-date-picker/div/div/input', By.XPATH, waiting_time=10000)
        calendar.click()
        while True:
            print('ano')
            self.wait(500)
            year = self.find_element(f'//*[@id="cdk-overlay-{index}"]/div/div/dp-month-calendar/div/dp-calendar-nav/div/div[1]/span', By.XPATH, waiting_time=15000)
            year_text = year.get_attribute('textContent')
            if year_text == valores['ano']:
                break
            elif year_text < valores['ano']:
                print('Ano impossivel')
                input()
            else:
                previous_year = self.find_element(f'//*[@id="cdk-overlay-{index}"]/div/div/dp-month-calendar/div/dp-calendar-nav/div/div[2]/div[1]/button', By.XPATH, waiting_time=10000)
                previous_year.click()
        
        months = self.find_elements('dp-calendar-month', By.CLASS_NAME)
        achou = False
        for month in months:
            if month.get_attribute('textContent').upper() == valores['mes'][0:3].upper():
                month.click()
                achou = True
                break
        
        if not achou:
            print('Não achou empresa com este nome')
            input()
            
        interpretacao_automatica = self.find_element('//*[@id="tab1"]/div/form/div[5]/div[3]/label/span', By.XPATH, waiting_time=10000)
        interpretacao_automatica.click()
        
        importar_periodo_unico = self.find_element('//*[@id="tab1"]/div/form/div[6]/div[2]/button/span', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(importar_periodo_unico, visible=True, waiting_time=15000)
        importar_periodo_unico.click()
        index += 1
        return index
        
        
    def obter_valores(self, diretorio):
        valores = []
        for pasta in os.listdir(diretorio):
            if os.path.isdir(os.path.join(diretorio, pasta)):
                mes_nome, ano = pasta.split()
                for arquivo in os.listdir(os.path.join(diretorio, pasta)):
                    if arquivo.endswith('.xlsx'):
                        nome_empresa = arquivo.split('.')[0]
                        caminho_arquivo = os.path.join(diretorio, pasta, arquivo)
                        valores.append({'nome_empresa': nome_empresa, 'mes': mes_nome, 'ano': ano, 'caminho_arquivo': caminho_arquivo})
        return valores
        
        
    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    MyWebBot.main()