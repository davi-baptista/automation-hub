from botcity.core import DesktopBot
import os
import shutil
import pyautogui
from pywinauto import Application
from datetime import datetime
import calendar
import pandas as pd
from dateutil.relativedelta import relativedelta

class Bot(DesktopBot):
    
    #-Função que implementa a lógica toda
    def action(self, execution=None):
        filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\03. DCTFWEB\Entrada\EMPRESAS DCTFWEB.xlsx'
        df = pd.read_excel(filepath, dtype=str)
        df = self.criando_colunas_status(df)
        
        self.open_ecac()
        self.select_certificate('controller consultores')
        self.navigate_to_dctf()
        self.human_check()
        days, month, year = self.get_dates()
        self.fill_dates(days, month, year)
        self.site_configures()
        
        pasta_downloads = r'C:\Users\davi.inov\Downloads'
        pasta_saida_darfs = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\03. DCTFWEB\Saida\DARFs'
        
        codigos_fernanda = ['0561', '1082', '1099', '1138', '1141', '1162', '1170', '1176', '1181', '1184', '1191',
                            '1196', '1200', '1213', '1218', '1221', '1646', '1657']
        codigos_berna = ['0422', '0588', '1708', '3208', '3280', '3426', '5952', '5960', '5979', '5987', '8045']
        
        for index, row in df.iterrows():
            try:
                companie = row['Nome']
                cnpj = row['Código']
                print(f'Abrindo empresa {companie}')
                
                status_pessoal = str(row['Status Pessoal'])
                status_fiscal = str(row['Status Fiscal'])
                status_transmissao = str(row['Status Transmissao'])
                
                pular_visualizacao, pular_transmissao = self.verificar_status(status_pessoal, status_fiscal, status_transmissao)
                if pular_visualizacao == True and pular_transmissao == True:
                    continue
                
                status_transmissao = 'Não há transmissão'
                self.search_outorgante(cnpj)
                self.carregando()
    
                if self.find( "nenhuma_declaracao", matching=0.97, waiting_time=2000):
                    raise Exception(f'Nenhuma declaração encontrada na empresa {companie}')
                
                if pular_transmissao == False:
                    if self.find( "transmitir", matching=0.97, waiting_time=2000):
                        self.click()
                        status_transmissao = self.transmitir()
                        
                        if self.find( "nenhuma_declaracao", matching=0.97, waiting_time=2000):
                            raise Exception(f'Nenhuma declaração encontrada na empresa {companie}')
            
                        if not self.find( "pesquisar", matching=0.97, waiting_time=30000):
                            self.not_found("pesquisar")
                        self.click()
                        self.wait(1500)
                    
                    if not self.find( "visualizar_recibo", matching=0.97, waiting_time=5000):
                        self.not_found("visualizar_recibo")
                    self.click()
                    if not self.find( "recibo_de_entrega", matching=0.97, waiting_time=30000):
                        self.not_found("recibo_de_entrega")
                    pyautogui.hotkey('ctrl', 'f4')
                    self.recent_download(companie, pasta_downloads, pasta_saida_darfs, 3)
                    status_transmissao += ', recibo baixado'
                
                if pular_visualizacao == False:
                    if not self.find( "saldo_zerado", matching=0.95, waiting_time=3000):
                        if self.find( "visualizar", matching=0.97, waiting_time=2000):
                            self.click()
                            status_pessoal = self.visualizar(codigos_fernanda, companie, pasta_downloads, pasta_saida_darfs, 0)
                            status_fiscal = self.visualizar(codigos_berna, companie, pasta_downloads, pasta_saida_darfs, 1)
                            self.voltar()
                            self.human_check()
                    else:
                        status_pessoal = 'Saldo zerado'
                        status_fiscal = 'Saldo zerado'
                    
                df = self.preencher_excel(df, index, filepath, status_pessoal, status_fiscal, status_transmissao)
                
            except Exception as e:
                print(f'Erro {e}')
                status_pessoal = f'{e}'
                status_fiscal = f'{e}'
                    
                df = self.preencher_excel(df, index, filepath, status_pessoal, status_fiscal, status_transmissao)
                
                if self.find( "erro_efetuar_operacao", matching=0.97, waiting_time=3000):
                    print('Erro inesperado encontrado, tratando-o')
                    self.key_f5()
                    self.tratar_erro_inesperado(days, month, year)
                
                self.voltar_inicio(days, month, year)
                continue
            
            finally:
                print('\n')
        
        if not self.find( "sair_com_seguranca", matching=0.97, waiting_time=10000):
            self.not_found("sair_com_seguranca")
        self.click()
        
        
    def criando_colunas_status(self, df):
        if 'Status Pessoal' not in df.columns:
            df['Status Pessoal'] = ''
        
        if 'Status Fiscal' not in df.columns:
            df['Status Fiscal'] = ''
        
        if 'Status Transmissao' not in df.columns:
            df['Status Transmissao'] = ''
        
        return df
    
            
    def open_ecac(self):
        self.execute(r'C:\Program Files\Mozilla Firefox\private_browsing.exe')
        if not self.find( "firefoxlogo", matching=0.97, waiting_time=30000):
            self.not_found("firefoxlogo")
        self.kb_type("https://cav.receita.fazenda.gov.br/autenticacao/login")
        self.enter()
        
        if not self.find( "entrar_gov", matching=0.97, waiting_time=30000):
            self.not_found("entrar_gov")
        self.wait(1000)
        self.double_click()
        
        
    def select_certificate(self, certificate):
        if not self.find( "certificadodigital", matching=0.97, waiting_time=30000): #Abre os certificados
            self.not_found("certificadodigital")
        self.wait(1000)
        self.click_relative(x=50, y=0)
        if not self.find( "detalhes_certificado", matching=0.97, waiting_time=30000): #Abre os certificados
            self.not_found("detalhes_certificado")
        self.kb_type(certificate)
        self.enter()
    
    
    def navigate_to_dctf(self):
        if not self.find( "declaracoes", matching=0.97, waiting_time=30000):
            self.not_found("declaracoes")
        self.click()
        
        if not self.find( "transmitir_dctf", matching=0.97, waiting_time=30000):
            self.not_found("transmitir_dctf")
        self.click()
    
    
    def human_check(self, check18h=0):
        if check18h == 1:
            return
        if not self.find( "sou_humano", matching=0.97, waiting_time=30000):
            self.not_found("sou_humano")
        self.click_relative(x=-30, y=0)
        
        contador = 0
        while True:
            contador += 1
            if self.find( "humano_verificado", matching=0.90, waiting_time=1000):
                self.click()
                break
            else:
                if self.find( "sou_humano", matching=0.97, waiting_time=1000):
                    self.click_relative(x=-30, y=0)
            if contador == 10:
                raise Exception('Captcha falhou')
            
        if not self.find( "prosseguir", matching=0.97, waiting_time=30000):
            self.not_found("prosseguir")
        self.click()
        
        return 'OK'
    
    
    def get_dates(self):
        data_hora_atual = datetime.now()

        # Subtraindo um mês da data atual
        data_hora_atual -= relativedelta(months=1)

        year = data_hora_atual.year
        month = data_hora_atual.month

        _, last_day = calendar.monthrange(year, month)
        
        # Formatando o último dia, o mês e o ano como strings
        last_day_str = str(last_day).zfill(2)
        month_str = str(month).zfill(2)
        year_str = str(year)
        
        return last_day_str, month_str, year_str
    
    
    def fill_dates(self, dia, mes, ano):
        if not self.find( "periodo_inicial", matching=0.97, waiting_time=30000):
            self.not_found("periodo_inicial")
        self.click_relative(x=0, y=30)
        self.control_a()
        
        self.kb_type('01')
        self.kb_type(mes)
        self.kb_type(ano)
        self.wait(500)
        
        if not self.find( "periodo_final", matching=0.97, waiting_time=30000):
            self.not_found("periodo_final")
        self.click_relative(x=30, y=30)
        self.control_a()
            
        self.kb_type(dia)
        self.kb_type(mes)
        self.kb_type(ano)
        
        return 'OK'
    
    
    def site_configures(self):
        # if not self.find( "exibir_situacao", matching=0.97, waiting_time=30000):
        #     self.not_found("exibir_situacao")
        # self.click_relative(x=-15, y=0)
        
        # if not self.find( "saldo_a_pagar", matching=0.97, waiting_time=30000):
        #     self.not_found("saldo_a_pagar")
        # self.click_relative(x=-10, y=0)
        
        # if not self.find( "minimizar_situacao", matching=0.97, waiting_time=30000):
        #     self.not_found("minimizar_situacao")
        # self.click_relative(x=-15, y=0)
        
        if not self.find( "sou_procurador", matching=0.97, waiting_time=30000):
            self.not_found("sou_procurador")
        self.click()
        
        
    def verificar_status(self, status_pessoal, status_fiscal, status_transmissao):
                pular_visualizacao = False
                if (status_pessoal != 'nan' and status_pessoal) and (status_fiscal != 'nan' and status_fiscal):
                    pular_visualizacao = True
                
                pular_transmissao = False
                if status_transmissao != 'nan' and status_transmissao:
                    pular_transmissao = True
                
                return pular_visualizacao, pular_transmissao
    
    
    def search_outorgante(self, companie):
        if not self.find( "outorgante", matching=0.90, waiting_time=30000):
            self.not_found("outorgante")
        self.click_relative(x=30, y=30)
        
        if not self.find( "cancelar_selecao", matching=0.93, waiting_time=30000):
            self.not_found("cancelar_selecao")
        self.click()
        self.paste(companie)
        self.enter()
        self.key_esc()
        
        if self.find( "sem_selecao", matching=0.90, waiting_time=2000):
            print('Empresa não encontrada')
            raise Exception(f'Empresa {companie} não encontrada')
        
        if not self.find( "pesquisar", matching=0.97, waiting_time=30000):
            self.not_found("pesquisar")
        self.click()
    
    
    def carregando(self):
        self.wait(500)
        while True:
            if self.find( "outorgante", matching=0.90, waiting_time=1000):
                self.click_relative(x=30, y=30)
                if self.find( "cancelar_selecao", matching=0.90, waiting_time=1000):
                    self.key_esc()
                    break
            self.wait(500)
        
        
    def transmitir(self):
        print('Fazendo transmissão')
        try:
            if self.find( "transmitir_2", matching=0.97, waiting_time=5000):
                self.click()
            else:
                self.enter()
        
            if not self.find( "pasta_assinador_digital", matching=0.97, waiting_time=30000):
                self.not_found("pasta_assinador_digital")
            self.wait(1000)
            self.type_up()
            self.enter()
            
            if not self.find( "abrir_executavel", matching=0.97, waiting_time=30000):
                self.not_found("abrir_executavel")
            self.wait(500)
            self.enter()
            
            if not self.find( "executar_java", matching=0.97, waiting_time=30000):
                self.not_found("executar_java")
            self.wait(500)
            self.enter()
            
            if not self.find( "documento_assinado_sucesso", matching=0.97, waiting_time=30000):
                self.not_found("documento_assinado_sucesso")
            self.wait(500)
            self.enter()
            
            if not self.find( "transmissao_sucesso", matching=0.97, waiting_time=30000):
                self.not_found("transmissao_sucesso")
            if not self.find( "transmissao_ok", matching=0.97, waiting_time=30000):
                self.not_found("transmissao_ok")
            self.click()
            
            return 'Transmitido com sucesso'
        
        except Exception as e:
            raise Exception(f'Erro durante transmissao: {e}')
        
        
    def visualizar(self, codigos, nome, pasta_downloads, pasta_saida, contador=0):
        print('Fazendo visualização')
        if contador == 0:
            pyautogui.hotkey('ctrl', 'shift', 'i')
            
        if not self.find( "filtros", matching=0.90, waiting_time=30000):
            raise Exception(f'Não encontrado nada na empresa')
        self.wait(1000)
        
        if not self.find( "inspecionar", matching=0.97, waiting_time=30000):
            self.not_found("inspecionar")
        self.click()
        self.wait(1000)
        
        if not self.find( "filtros", matching=0.90, waiting_time=30000):
            self.not_found("filtros")
        self.click()
        self.wait(500)
        pyautogui.hotkey('ctrl', 'shift', 'k')
        
        self.paste("document.querySelector('#cphConteudo_TabelaVinculacoesDARF_GridViewVinculacoesimgExpandirTodos').click();")
        self.enter()
        
        if contador == 0:
            self.paste("document.querySelector('#cphConteudo_TabelaVinculacoesDARF_GridViewVinculacoes_chkEmitirGuiaPagamento_0').click();")
            self.enter()
            if self.find( "removera_selecoes", matching=0.95, waiting_time=2000):
                self.enter()
            self.enter()
        elif contador == 1:
            self.paste("document.querySelector('#cphConteudo_TabelaVinculacoesDARF_GridViewVinculacoes_chkEmitirGuiaPagamento_0').click();")
            self.enter()
            if self.find( "removera_selecoes", matching=0.95, waiting_time=2000):
                self.enter()
            self.enter()
            self.paste("document.querySelector('#cphConteudo_TabelaVinculacoesDARF_GridViewVinculacoes_chkEmitirGuiaPagamento_0').click();")
            self.enter()
            if self.find( "removera_selecoes", matching=0.95, waiting_time=2000):
                self.enter()
            self.enter()
            
        for codigo in codigos:
            codigo_javascript = "{"
            codigo_javascript += f"const codigoProcurado = '{codigo}';"
            codigo_javascript += "const todosOsCodigos = document.querySelectorAll('.descricao-codigo-receita'); todosOsCodigos.forEach((elemento, indice) => { if (elemento.textContent.includes(codigoProcurado)) { const idDoBotao = `cphConteudo_TabelaVinculacoesDARF_GridViewVinculacoes_chkEmitirGuiaPagamento_${indice + 1}`; document.querySelector(`#${idDoBotao}`).click(); } }); if (todosOsCodigos.length === 0) { console.log('Código não encontrado.'); }}"

            self.paste(codigo_javascript)
            self.enter()
            
        self.paste("document.querySelector('#LinkEmitirDARFIntegral').click();")
        self.enter()
        
        if self.find( "receita_federal", matching=0.90, waiting_time=8000):
            pyautogui.hotkey('ctrl', 'f4')
        
        if self.find( "nenhum_credito", matching=0.97, waiting_time=1000):
            print('Nenhum credito encontrado')
            return 'Nenhum credito encontrado'
        
        self.recent_download(nome, pasta_downloads, pasta_saida, contador)
        
        if not self.find( "guia_gerada_ok", matching=0.97, waiting_time=60000):
            self.not_found("guia_gerada_ok")
        self.click()
        
        return 'Guia gerada com sucesso!'
    
    
    def voltar(self):
        pyautogui.hotkey('ctrl', 'shift', 'k')
        self.wait(500)
        self.paste("document.querySelector('#cphConteudo_LinkRetornar').click();")
        self.enter()
        pyautogui.hotkey('ctrl', 'shift', 'i')
        
        if self.find( "voltar_guia", matching=0.97, waiting_time=2000):
            self.click()
            
            
    def tratar_erro_inesperado(self, days, month, year):
        self.human_check()
        self.fill_dates(days, month, year)
        self.site_configures()
    
    
    def voltar_inicio(self, days, month, year):
        pyautogui.hotkey('ctrl', 'l')
        self.wait(500)
        self.paste('https://cav.receita.fazenda.gov.br/ecac/Aplicacao.aspx?id=10015&origem=menu')
        self.enter()
        if self.find( "inspecionar", matching=0.97, waiting_time=3000):
            pyautogui.hotkey('ctrl', 'shift', 'i')
        
        contador = 0
        while True:
            contador += 1
            self.key_f5()
            if not self.find( "sou_humano", matching=0.97, waiting_time=3000):
                continue
            
            status_captcha = self.human_check()
            if status_captcha == 'OK':
                break
            
            if contador == 10:
                break
        
        status_fill_dates = self.fill_dates(days, month, year)
        if status_fill_dates != 'OK':
            os._exit()
            
        self.site_configures()
                    
                    
    def recent_download(self, company_name, pasta_downloads, caminho_saida_darf, contador):
        arquivos = [os.path.join(pasta_downloads, arquivo) for arquivo in os.listdir(pasta_downloads)]
        
        departamento = None
        if contador == 0:
            departamento = 'Pessoal'
        elif contador == 1:
            departamento = 'Fiscal'
        elif contador == 3:
            departamento = 'Recibo'
            
        if arquivos:
            arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
            caminho_destino = os.path.join(caminho_saida_darf, company_name, departamento, os.path.basename(arquivo_mais_recente))
            
            os.makedirs(os.path.join(caminho_saida_darf, company_name, departamento), exist_ok=True)
            shutil.move(arquivo_mais_recente, caminho_destino)

            print(f'O arquivo mais recente foi movido para: {caminho_destino}')
            return caminho_destino
        
        else:
            print('Não foi encontrado nenhum arquivo na pasta de downloads')
            raise Exception(f'Não foi encontrado nenhum arquivo na pasta de downloads')
    
    
    def preencher_excel(self, df, index, filepath, status_pessoal, status_fiscal, status_transmissao):
        print(f'Status Pessoal {status_pessoal}')
        print(f'Status Fiscal {status_fiscal}')
        print(f'Status Transmissao {status_transmissao}')
        df.at[index, 'Status Pessoal'] = status_pessoal
        df.at[index, 'Status Fiscal'] = status_fiscal
        df.at[index, 'Status Transmissao'] = status_transmissao
        
        df.to_excel(filepath, index=False)
        return df
    
    
    def fecharprocesso(self, nome_do_processo):
        app = Application(backend="uia").connect(path='nome_do_processo')
        janela = app.window()
        janela.close()
        
        
    def not_found(self, label):
        print(f'Element not found: {label}')
        raise Exception(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()