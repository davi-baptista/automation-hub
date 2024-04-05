from botcity.web import WebBot, By
from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import os


class MyWebBot(WebBot):


    def action(self, execution=None):
        self.configure()
        self.login()
        self.new_aba_navigate()
        
        filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\04_recolher_dados_funcionarios_esocial\data\Esocial.xlsx'
        df = pd.read_excel(filepath, dtype=str)
        
        if 'Status' not in df.columns:
            df['Status'] = ''
        
        for index, row in df.iterrows():
            
            status_atual = str(row['Status']).strip()
            if status_atual != 'nan' and status_atual:
                print(f'Pulando CPF {row["CPF"]} ({row["NOME"]}) pois o status já está definido como {status_atual}.')
                continue
            
            cpf, nome = row['CPF'].zfill(11), row['NOME']
            
            resultado = self.navigate_and_download_pdf(cpf, nome)
            df.at[index, 'Status'] = resultado
            
            df.to_excel(filepath, index=False)
            print(f'Status do CPF {cpf} ({nome}) atualizado como {resultado} e salvo no arquivo.')
        
        
    def configure(self):
        # Configurando o ChromeDriver
        self.driver_path = ChromeDriverManager().install()
        self.options = default_options(headless=False, download_folder_path=r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\04_recolher_dados_funcionarios_esocial\output\eSocial', user_data_dir=r'c:\Users\davi.inov\AppData\Local\Google\Chrome for Testing\User Data', page_load_strategy='normal')
    
    
    def login(self):
        self.browse("https://login.esocial.gov.br/login.aspx")
        self.maximize_window()
        entrar_gov = self.find_element('//*[@id="login-acoes"]/div[2]/p/button', By.XPATH, waiting_time=10000)
        entrar_gov.click()
        certificado_digital = self.find_element('//*[@id="login-certificate"]', By.XPATH, waiting_time=10000)
        certificado_digital.click()
        input()
    
    
    def new_aba_navigate(self):
        abas = self.get_tabs()
        self.activate_tab(abas[0])
        self.browse('https://www.esocial.gov.br/portal/Home/Inicial')
        empregado = self.find_element('//*[@id="menuEmpregado"]', By.XPATH, waiting_time=60000)
        empregado.click()
        gestao_empregados = self.find_element('/html/body/div[3]/div[3]/div/div/ul/li[2]/ul/li[2]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(gestao_empregados, visible=True, waiting_time=60000)
        gestao_empregados.click()
    
    
    def navigate_and_download_pdf(self, cpf, nome):
        try:
            pesquisar = self.find_element('filtro', By.ID, waiting_time=60000)
            pesquisar.click()
            self.control_a()
            pesquisar.send_keys(cpf)
            
            contrato_atual = self.find_element('//*[@id="ui-id-3"]', By.XPATH, waiting_time=60000)
            self.wait_for_element_visibility(contrato_atual, visible=True, waiting_time=60000)
            contrato_atual.click()
            
            dados_cadastrais = self.find_element('//*[@id="conteudo-pagina"]/form/div[3]/div[1]/div[2]/ul/li/a', By.XPATH, waiting_time=60000)
            dados_cadastrais.click()
            
            consultar_dados = self.find_element('//*[@id="conteudo-pagina"]/form/div[3]/div[1]/div[2]/ul/li/ul/li[1]/a', By.XPATH, waiting_time=60000)
            self.wait_for_element_visibility(consultar_dados, visible=True, waiting_time=60000)
            consultar_dados.click()
            
            carregando = self.find_element('//*[@id="loading"]', By.XPATH, waiting_time=60000)
            self.wait_for_element_visibility(carregando, visible=True, waiting_time=60000)
            self.wait_for_element_visibility(carregando, visible=False, waiting_time=60000)
            
            self.execute_javascript("window.print();")
            path = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\04_recolher_dados_funcionarios_esocial\output\eSocial\eSocial.pdf'

            if self.wait_for_file(path=path, timeout=60000):
                novo_nome = os.path.join(os.path.dirname(path), f'{nome}.pdf')
                if os.path.exists(novo_nome):
                    print(f'O arquivo {novo_nome} já existe. Deletando o PDF baixado.')
                    os.remove(path)
                    return 'ERRO: Arquivo já existe'
                
                os.rename(path, novo_nome)
            
            print('PDF baixado com sucesso')
            voltar = self.find_element('//*[@id="btnVoltar"]', By.XPATH, waiting_time=60000)
            voltar.click()
            return f'Arquivo baixado com sucesso'

        except Exception as e:
            print(f'Erro {e}')
            self.new_aba_navigate()
            return f'Erro ao baixar o arquivo {e}'
        
        
    def not_found(self, label):
        print(f"Element not found: {label}")


def main():
    MyWebBot.main()


if __name__ == '__main__':
    main()
