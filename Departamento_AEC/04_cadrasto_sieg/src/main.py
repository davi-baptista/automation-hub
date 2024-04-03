from botcity.web import WebBot, Browser, By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

class MyWebBot(WebBot):
    
    
    def action(self, execution=None):
        email_gestta = ' '
        password_gestta = ' '
        email_sieg = ' '
        password_sieg = ' '
        filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_AEC\04_cadrasto_sieg\data\BASE DE DADOS DE CLIENTES ATIVOS.xlsx'
        
        self.configure()

        df = pd.read_excel(filepath)
        print('Arquivo excel lido com sucesso')
        
        if 'PROCESSADO' not in df.columns:
            df['PROCESSADO'] = df.duplicated('CNPJ', keep='first')
            df.to_excel(filepath, index=False)
            
        contador_gestta = 0
        contador_sieg = 0
        for index, row in df.iterrows():
            try:
                process = row['PROCESSADO']
                if process or process == 'Concluido':
                    if process == 'VERDADEIRO':
                        df.at[index, 'PROCESSADO'] = 'Conferir. CNPJ repetido'
                    continue
                self.browse("https://app.gestta.com.br")
                self.maximize_window()
                
                if contador_gestta == 0:
                    self.login(email_gestta, password_gestta, 'email', 'password')
                    contador_gestta = 1
                    
                cnpj = row['CNPJ']
                name = row['RAZÃO SOCIAL']
                state = row['ESTADO']
                county = row['MUNICIPIO']
                self.navegate_to_tarefas()
                login, senha = self.get_login(cnpj)
                self.create_tab('https://hub.sieg.com/')
                if contador_sieg == 0:
                    self.login(email_sieg, password_sieg, 'txtEmail', 'txtPassword')
                    contador_sieg = 1
                self.navigate_to_service()
                self.register_automation(cnpj, name, state, county, login, senha)
                
                df.at[index, 'PROCESSADO'] = 'Concluido'
                
                df.to_excel(filepath, index=False)
                
                if len(self.get_tabs()) > 1:
                    self.close_page()
                
            except AttributeError as e:
                print(f'Erro de atributo: {e}')
                if 'get_attribute' in str(e):
                    df.at[index, 'PROCESSADO'] = 'Login e senha não encontrados no gestta'
                elif "click" in str(e):
                    df.at[index, 'PROCESSADO'] = 'Empresa não encontrada no gestta'
                else:
                    df.at[index, 'PROCESSADO'] = f'Outro erro de atributo: {e}'
                    
                df.to_excel(filepath, index=False)
                
                if len(self.get_tabs()) > 1:
                    self.close_page()
                
            except Exception as e:
                print(e)
                df.at[index, 'PROCESSADO'] = e
                
                df.to_excel(filepath, index=False)
                
                if len(self.get_tabs()) > 1:
                    self.close_page()
            
            
    def configure(self):
        # Configurando o ChromeDriver
        self.driver_path = ChromeDriverManager().install()
        self.headless = False
        
        
    def login(self, email, password, id_email, id_password):
        print('Fazendo login')
        
        campo_email = self.find_element(id_email, By.ID, waiting_time=20000)
        campo_email.send_keys(email)
        
        campo_senha = self.find_element(id_password, By.ID, waiting_time=10000)
        campo_senha.send_keys(password)
        self.wait(500)
        self.enter()
        
        
    def navegate_to_tarefas(self):
        print('Indo até tarefas')
        menu_tarefas = self.find_element('//*[@id="gestta-menu"]/div/div[5]/div/div/div[2]/div', By.XPATH, waiting_time=20000)
        menu_tarefas.click()
        
        todos_usuarios = self.find_element('//*[@id="gestta-multiselect-dropdown-4"]/div/div/p/i', By.XPATH, waiting_time=20000)
        todos_usuarios.click()
        
        filtro_avancado = self.find_element('//*[@id="page-wrapper"]/div[1]/div/nav/form[1]/div[4]/button[1]', By.XPATH, waiting_time=20000)
        filtro_avancado.click()
        
        marcar_concluida = self.find_element('/html/body/div[1]/div/div/div[2]/form[2]/div[2]/div/label/span', By.XPATH, waiting_time=20000)
        marcar_concluida.click()
        
        filtrar = self.find_element('/html/body/div[1]/div/div/div[2]/div[7]/div/button', By.XPATH, waiting_time=10000)
        filtrar.click()
    
    
    def get_login(self, cnpj):
        print('Pegando login e senha')
        todas_clientes = self.find_element('//*[@id="gestta-multiselect-dropdown-8"]/div/div/p', By.XPATH, waiting_time=10000)
        todas_clientes.click()
        campo_empresas = self.find_element('/html/body/ul[1]/li[1]/input', By.XPATH, waiting_time=20000)
        campo_empresas.send_keys(cnpj)
        selecionar_empresa = self.find_element('/html/body/ul[1]/li[4]/a', By.XPATH, waiting_time=10000)
        selecionar_empresa.click()
        marcar_concluida = self.find_element('//*[@id="gestta-multiselect-dropdown-8"]/div/div/p', By.XPATH, waiting_time=10000)
        marcar_concluida.click()
        primeira_tarefa = self.find_element('//*[@id="mixed-task-structure-screen"]/div/div/div[1]/div/div/div/div/div[3]/div/ul/li[1]', By.XPATH, waiting_time=10000)
        primeira_tarefa.click()
        dados_cliente = self.find_element('//*[@ng-click="taskDetails.actions.openCustomerDetailsModal()"]', By.XPATH, waiting_time=10000)
        dados_cliente.click()
        campo_texto = self.find_element('//*[@ng-bind-html="dataCtrl.details.observation"]', By.XPATH, waiting_time=10000)
        campo_texto = campo_texto.get_attribute('textContent')
        print(campo_texto)
        
        def tratar_login(texto):
            palavras = texto.split()
            login = None
            senha = None
            for i, palavra in enumerate(palavras):
                if login != None and senha != None:
                    break
                
                if palavra.lower() == 'login:':
                    login = palavras[i + 1]
                    
                elif palavra.lower() == 'senha:':
                    senha = palavras[i + 1]
                    
            if login == None or senha == None or login.lower() == 'aguardando' or senha.lower() == 'aguardando':
                raise Exception(f'Empresa sem login e senha')
            return login, senha
        
        login, senha = tratar_login(campo_texto)
        return login, senha


    def navigate_to_service(self):
        print('Abrindo serviços')
        
        todos_servicos = self.find_element('//*[@id="sieg-header"]/div/div[1]/a[2]', By.XPATH, waiting_time=60000)
        todos_servicos.click()
        
        autodocs = self.find_element('//*[@id="rowHubServices"]/div[4]/a/div[2]', By.XPATH, waiting_time=10000)
        autodocs.click()
        
        adicionar_automacao = self.find_element('//*[@id="btnAddAutomation"]', By.XPATH, waiting_time=10000)
        adicionar_automacao.click()
        
        
    def register_automation(self, cnpj, name, state, county, login, password):
        print('Registrando automação')
        
        selecionar = self.find_element('//*[@id="modalAutomationType"]/div/div/div[2]/a', By.XPATH, waiting_time=20000)
        self.wait_for_element_visibility(selecionar, visible=True, waiting_time=10000)
        selecionar.click()
        
        estado = self.find_element('//*[@id="ddlConsultAutomationUf"]', By.XPATH, waiting_time=20000)
        self.wait_for_element_visibility(estado, visible=True, waiting_time=10000)
        estado.click()
        
        estado.send_keys(state)
        self.enter()
        
        municipio = self.find_element('//*[@id="ddlConsultAutomationMunicipio"]', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(municipio, visible=True, waiting_time=10000)
        municipio.click()
        municipio.send_keys(county)
        
        proximo = self.find_element('//*[@id="location-tab"]/div[3]/div/a', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(proximo, visible=True, waiting_time=10000)
        proximo.click()
        proximo.click()
        
        razao_social = self.find_element('//*[@id="txtNameCompany"]', By.XPATH, waiting_time=20000)
        self.wait_for_element_visibility(razao_social, visible=True, waiting_time=10000)
        razao_social.click()
        razao_social.send_keys(name)
        
        cnpj_cpf = self.find_element('//*[@id="txtCnpjCompany"]', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(cnpj_cpf, visible=True, waiting_time=10000)
        cnpj_cpf.click()
        cnpj_cpf.send_keys(cnpj)
        
        campo_login = self.find_element('//*[@id="txtAutomationLoginName"]', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(campo_login, visible=True, waiting_time=10000)
        campo_login.click()
        campo_login.send_keys(login)
        
        senha = self.find_element('//*[@id="txtAutomationLoginPassword"]', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(senha, visible=True, waiting_time=10000)
        senha.click()
        senha.send_keys(password)
        
        proximo = self.find_element('//*[@id="login-tab"]/div[5]/div/a', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(proximo, visible=True, waiting_time=10000)
        proximo.click()
        
        marcar_download = self.find_element('//*[@id="cbkDownloadDanfe"]', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(marcar_download, visible=True, waiting_time=10000)
        marcar_download.click()
        
        proximo = self.find_element('//*[@id="config-tab"]/div[4]/div/a', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(proximo, visible=True, waiting_time=10000)
        proximo.click()
        
        salvar = self.find_element('//*[@id="btnAddAutomationNfse"]', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(salvar, visible=True, waiting_time=10000)
        salvar.click()
        
        print('Cadastro finalizado com sucesso')
    
    def not_found(self, label):
        print(f"Element not found: {label}")

if __name__ == '__main__':
    MyWebBot.main()
    