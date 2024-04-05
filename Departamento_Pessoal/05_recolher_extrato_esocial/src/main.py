from botcity.web import WebBot, By
from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
import pandas as pd
import os


class MyWebBot(WebBot):


    def action(self, execution=None):
        self.configure()
        self.login()
        
        filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\05_recolher_extrato_esocial\data\RESCISÃO.xlsx'
        df = pd.read_excel(filepath, dtype={'CPF': str})
        
        if 'Status' not in df.columns:
            df['Status'] = ''
        
        for index, row in df.iterrows():
            try:
                status_atual = str(row['Status'])
                if status_atual != 'nan' and status_atual:
                    print(f'Pulando ({row["NOME"]}) pois o status já está definido como {status_atual}.')
                    continue
                
                nome, pis = row['NOME'], row['PIS']
                self.extrato_analitico(pis)
                self.extrato_trabalhador(pis, nome)
                
                resultado = 'Concluido com sucesso'
                df.at[index, 'Status'] = resultado
                df.to_excel(filepath, index=False)
            
            except Exception as e:
                print(f'Erro {e}')
                resultado = f'Erro {e}'
                df.at[index, 'Status'] = resultado
                df.to_excel(filepath, index=False)
            
            
    def configure(self):
        # Configurando o ChromeDriver
        self.driver_path = ChromeDriverManager().install()
        self.options = default_options(headless=False, download_folder_path=r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\05_recolher_extrato_esocial\output', page_load_strategy='normal')
    
    
    def login(self):
        self.browse("https://conectividadesocialv2.caixa.gov.br/sicns/")
        self.maximize_window()
        empregador = self.find_element('//*[@id="btnEmpregador"]', By.XPATH, waiting_time=10000)
        empregador.click()
        input()
    
    
    def extrato_analitico(self, pis):
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td[1]/select/option[11]', By.XPATH, waiting_time=10000)
        empregador.click()
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[2]/td[3]/table[3]/tbody/tr[3]/td/select/option[2]', By.XPATH, waiting_time=10000)
        empregador.click()
        empregador = self.find_element('//*[@id="txtPIS"]', By.XPATH, waiting_time=10000)
        empregador.send_keys(pis)
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[2]/td[3]/table[3]/tbody/tr[9]/td/a[1]', By.XPATH, waiting_time=10000)
        empregador.click()
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[2]/td[3]/table[3]/tbody/tr[2]/td', By.XPATH, waiting_time=10000)
        texto = empregador.get_attribute('textContent')
        if 'Solicitação efetuada com sucesso' in texto:
            return f'Analitico Ok'
        return f'Analitico Erro'
    
    
    def extrato_trabalhador(self, pis, nome):
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[1]/td[3]/table/tbody/tr/td[1]/select/option[11]', By.XPATH, waiting_time=10000)
        empregador.click()
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[3]/td[3]/p/table/tbody/tr[3]/td[2]/select/option[2]', By.XPATH, waiting_time=10000)
        empregador.click()
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[3]/td[3]/p/table/tbody/tr[11]/td[2]/input', By.XPATH, waiting_time=10000)
        empregador.send_keys(pis)
        empregador = self.find_element('/html/body/form/table[2]/tbody/tr[3]/td[3]/p/table/tbody/tr[23]/td/a[1]', By.XPATH, waiting_time=10000)
        empregador.click()
        
        linhas = self.find_elements('tr', By.TAG_NAME)
        numero_linhas = len(linhas)-37
        
        empregador = self.find_element(f'/html/body/form/table[2]/tbody/tr[2]/td[3]/table[3]/tbody/tr[2]/td/table[2]/tbody/tr[{numero_linhas}]/td/a[2]', By.XPATH, waiting_time=10000)
        empregador.click()
        
        abas = self.get_tabs()
        while(len(abas)<1):
            abas = self.get_tabs()
        self.activate_tab(abas[1])
        self.execute_javascript("window.print();")

        path = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\05_recolher_extrato_esocial\output\CSE - Conectividade Social _ Empregador.pdf'

        if self.wait_for_file(path=path, timeout=60000):
            novo_nome = os.path.join(os.path.dirname(path), f'{nome}.pdf')
            
            if os.path.exists(novo_nome):
                print(f'O arquivo {novo_nome} já existe. Deletando o PDF baixado.')
                os.remove(path)
                self.close_page()
                return 'ERRO: Arquivo já existe'
            
            os.rename(path, novo_nome)
            
        self.close_page()
        
        
    def not_found(self, label):
        print(f"Element not found: {label}")


def main():
    MyWebBot.main()


if __name__ == '__main__':
    main()
