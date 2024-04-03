from botcity.web import WebBot, Browser, By
from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
import locale


class MyWebBot(WebBot):
    
    
    def action(self, execution=None):
        print('Iniciando robo, aguarde um momento...')
        email = 'inovacao@controller-rnc.com.br'
        password = 1234
        database_path = r"C:\Users\davi.inov\Desktop\Projetos\Departamento_Contabil\02_extrair_dctfweb_Iris\data\database.xlsx"
        exit_path = r"C:\Users\davi.inov\Desktop\Projetos\Departamento_Contabil\02_extrair_dctfweb_Iris\output\output_sieg.xlsx"
        
        self.configure()
        self.browse("https://hub.sieg.com/IriS")
        self.maximize_window()
        self.login(email, password)
        self.navigate_to_dctf()
        data, mes_str = self.get_last_month()
        emps, nomes, cnpjs = self.get_dados_from_excel(database_path)
        
        # Defina as colunas do seu DataFrame
        colunas = ["EMP", "Cod.N", "PREVIDENCIÁRIA SEGURADOS", "PREVIDENCIÁRIA PATRONAL", "OUTRAS ENTIDADES E FUNDOS", "IRRF", "Erro"]
        df = pd.DataFrame(columns=colunas)

        for emp, nome, cnpj in zip(emps, nomes, cnpjs):
            try:
                print(f'->Pegando dados da GREENLIFE {nome}.')
                totais = self.get_totais(cnpj, data, mes_str)
                df = self.create_excel(df, emp, nome, totais, 'Concluido')
                print('Dados capturados do sieg com sucesso\n')
                
            except Exception as e:
                print(f'Erro {e}.\n')
                df = self.create_excel(df, emp, nome, {key: None for key in colunas}, str(f'Erro {e}'))
                continue
        
        print('Formatando excel...')
        
        pd.set_option('future.no_silent_downcasting', True)
        df.fillna(0, inplace=True)
        
        self.save_excel_with_formatting(df, exit_path)
        print(f'Excel salvo no caminho: {exit_path}\n')
        print('Programa finalizado.')
        input()
        
        
    def configure(self):
        # Configurando o ChromeDriver
        self.driver_path = ChromeDriverManager().install()
        self.options = default_options(headless=False, download_folder_path=self.filepath, user_data_dir=None, page_load_strategy='normal')
    
    
    def login(self, email, password):
        print('Abrindo Sieg IRIS e indo até a aba de DCTFWEB.')
        campo_email = self.find_element('txtEmail', By.ID, waiting_time=10000)
        campo_email.send_keys(email)
        campo_senha = self.find_element('txtPassword', By.ID, waiting_time=10000)
        campo_senha.send_keys(password)
        self.wait(200)
        self.enter()
        
        
    def navigate_to_dctf(self):
        dctfweb = self.find_element('//*[@id="iriS"]/div[2]/div/div[3]/div/div/div[2]/div[43]/a/a/span[2]', By.XPATH, waiting_time=60000)
        dctfweb.click()
        
        self.find_element('//*[@id="tableDCTFWeb"]/tbody/tr[2]/td[12]/a', By.XPATH, waiting_time=60000)
    
    
    def get_last_month(self):
        locale.setlocale(locale.LC_TIME, 'pt_BR.utf8')
        now = datetime.now()

        # Subtrair um mês da data atual
        last_month = now - relativedelta(months=1)

        # Formatar a data no formato 'MM/YYYY'
        data = last_month.strftime('%m/%Y')
        mes_abreviado = last_month.strftime('%b').capitalize()

        return data, mes_abreviado
    
    
    def get_dados_from_excel(self, filepath):
        print('->Extraindo dados do excel\n')
        df = pd.read_excel(filepath, dtype={'CNPJ':str})
        df_filtrado = df.dropna(subset=['CNPJ'])
        cnpjs = df_filtrado['CNPJ']
        emp = df_filtrado['EMP']
        nomes = df_filtrado['Cod.N']
        
        print('Dados capturados do excel com sucesso')
        return emp, nomes, cnpjs
    
    
    def get_totais(self, cnpj, data, mes_str):
        self.key_esc()
        todas_empresas = self.find_element('//*[@id="select2-ddlCompanyIris-container"]', By.XPATH, waiting_time=30000)
        todas_empresas.click()
        
        pesquisar_empresa = self.find_element('/html/body/span/span/span[1]/input', By.XPATH, waiting_time=30000)
        pesquisar_empresa.send_keys(cnpj)
        self.enter()
        
        carregando = self.find_element('//*[@id="tableDCTFWeb_processing"]', By.XPATH, waiting_time=10000)
        self.wait_for_element_visibility(carregando, visible=True, waiting_time=10000)
        self.wait_for_element_visibility(carregando, visible=False, waiting_time=60000)
        
        tabela = self.find_element('//*[@id="tableDCTFWeb"]/tbody/tr[2]/td[12]/a', By.XPATH, waiting_time=3000)
        if tabela is None:
            raise Exception(f'Não encontrado nada no mês inserido')
        
        selecionar_data = self.find_element('//*[@id="txtDatePicker"]', By.XPATH, waiting_time=30000)
        selecionar_data.click()
        self.control_a()
        self.control_c()
        if self.get_clipboard() == '':
            selecionar_data.send_keys(data)
            meses = self.find_elements(f'body > div.datepicker.datepicker-dropdown.dropdown-menu.datepicker-orient-left.datepicker-orient-top > div.datepicker-months > table > tbody > tr > td > span', By.CSS_SELECTOR, waiting_time=30000)
            for mes in meses:
                if mes_str in mes.get_attribute('textContent'):
                    mes.click()
                    break
                
            carregando = self.find_element('//*[@id="tableDCTFWeb_processing"]', By.XPATH, waiting_time=10000)
            self.wait_for_element_visibility(carregando, visible=True, waiting_time=30000)
            self.wait_for_element_visibility(carregando, visible=False, waiting_time=60000)
            
        tabela = self.find_element('//*[@id="tableDCTFWeb"]/tbody/tr[2]/td[12]/a', By.XPATH, waiting_time=3000)
        if tabela is None:
            raise Exception(f'Não encontrado nada no mês inserido')
        
        detalhes = self.find_element('//*[@id="tableDCTFWeb"]/tbody/tr[1]/td/div/a', By.XPATH, waiting_time=3000)
        if detalhes is not None:
            detalhes.click()
        
        i = 1
        while True:
            i += 1
            situacao = self.find_element(f'//*[@id="tableDCTFWeb"]/tbody/tr[{i}]/td[3]/span', By.XPATH, waiting_time=30000)
            
            if 'Ativa' in situacao.get_attribute('textContent'):
                mes_ativo = self.find_element(f'//*[@id="tableDCTFWeb"]/tbody/tr[{i}]/td[12]/a', By.XPATH, waiting_time=30000)
                mes_ativo.click()
                break
            
            if 'Em andamento' in situacao.get_attribute('textContent'):
                raise Exception(f'Situação: Em andamento.')
        
        totais_tipo = self.find_elements(f'#tableModalDctfDebts > tbody > tr > td > div > span > b:nth-of-type(1)', By.CSS_SELECTOR, waiting_time=30000)
        totais_valor = self.find_elements(f'#tableModalDctfDebts > tbody > tr > td > div > span > b:nth-of-type(2)', By.CSS_SELECTOR, waiting_time=30000)
        resultados = {}
        
        for tipo, valor in zip(totais_tipo, totais_valor):
            # Processar o texto do tipo e do valor
            chave = tipo.get_attribute('textContent').replace('Total ', '').replace('CONTRIBUIÇÃO ', '').replace('PARA ', '')
            valor = valor.get_attribute('textContent').replace('Total Calculado ', '')
            
            # Adicionar ao dicionário
            resultados[chave] = valor

        
        fechar = self.find_element('//*[@id="modalDctfDebt"]/div/div/div[3]/a', By.XPATH, waiting_time=30000)
        fechar.click()
        
        return resultados


    def create_excel(self, df, emp, nome, totais, error_msg):
        totais['EMP'] = emp
        totais['Cod.N'] = nome
        totais['Erro'] = error_msg
        novo_registro_df = pd.DataFrame([totais])
        df = pd.concat([df, novo_registro_df], ignore_index=True)
        return df
    
    
    def save_excel_with_formatting(self, df, filename):
        writer = pd.ExcelWriter(filename, engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        
        worksheet = writer.sheets['Sheet1']
        
        # Adicionando uma tabela com estilo
        (max_row, max_col) = df.shape
        column_settings = [{'header': column} for column in df.columns]
        worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings, 'style': 'Table Style Medium 6'})

        # Ajustando o espaçamento das colunas
        for i, col in enumerate(df.columns):
            worksheet.set_column(i, i, 15)

        # Fechar o writer e salvar o arquivo Excel
        writer.close()
    
    
    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    MyWebBot.main()