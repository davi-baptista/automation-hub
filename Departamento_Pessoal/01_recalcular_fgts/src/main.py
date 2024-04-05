from botcity.core import DesktopBot
import os
import shutil
import pyautogui
from pywinauto import Application
import pandas as pd
import re
import time
import pygetwindow as gw

class Bot(DesktopBot):
    
    
    def action(self, execution=None):
        
        # Modifique os caminhos conforme o caminho da sua maquina
        path = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\01_recalcular_fgts\data\Planilha Recalculo.xlsx'
        arquivo_govbox = r'C:\Program Files (x86)\GOVBOX\Arquivos'
        pasta_download = r'C:\Users\davi.inov\Downloads'
        
        df = pd.read_excel(path, dtype=str)
        if 'Status' not in df.columns:
            df['Status'] = ''
        
        for indice, row in df.iterrows():
            try:
                status_atual = str(row['Status']).strip()
                
                if status_atual != 'nan' and status_atual:
                    print(f'Pulando competencia {row["Competencia"]} da empresa {row["CNPJ"]} pois o status já está definido como {status_atual}.')
                    continue
                
                companie = self.limpar_caracteres(str(row['CNPJ'])).zfill(14)
                competencia = self.limpar_caracteres(str(row['Competencia'])).zfill(6)
                data_pagamento = self.limpar_caracteres(str(row['Data de pagamento'])).zfill(8)
                workers = self.limpar_caracteres(str(row['PIS']))
                if workers != 'nan':
                    print(workers)
                    workers = [worker.zfill(11) for worker in workers.split(';')]
                
                filepath = row['Caminho do arquivo']
                certificate = row['Certificado']
                pis_type = row['Todos os pis']
                
                print(f'CNPJ: {companie}')
                print(f'Competencia: {competencia}')
                print(f'Data de pagamento: {data_pagamento}')
                print(f'PIS: {workers}')
                
                self.open_programs()
                
                if filepath.lower().endswith('.re'):
                    self.sefip_file(filepath, competencia, data_pagamento)
                elif filepath.lower().endswith('.bkp'):
                    self.bkp_file(filepath, competencia, data_pagamento)
                else:
                    self.gbk_file(filepath, competencia, data_pagamento)
                    
                self.mark_modality(pis_type, companie, workers)
                self.simulating_and_executating()
                new_path, arquivo_recente = self.previous_path(filepath, competencia, arquivo_govbox)
                caminho_xml = self.open_conectividade(competencia, certificate, new_path, arquivo_recente, pasta_download)
                
                # Abrindo sefip
                self.open_window('Sefip.exe')
                self.wait(200)
                
                pasta_relatorios = r'C:\!!Recalculo'
                self.criando_pasta(pasta_relatorios)
                
                # Abrindo relatorios e baixando
                self.gerando_relatorios(caminho_xml, arquivo_recente, new_path, pasta_relatorios, arquivo_govbox)
                print('Concluído com sucesso')
                df.at[indice, 'Status'] = 'Concluído com sucesso'
                df.to_excel(path, index=False)
                
            except Exception as e:
                print(f'Erro: {e}')
                df.at[indice, 'Status'] = f'Erro {e}'
                df.to_excel(path, index=False)
                continue
                
                
    def esperar_janela_abrir(self, nome_janela, timeout=30):
        start_time = time.time()

        while time.time() - start_time < timeout:
            janelas = gw.getWindowsWithTitle(nome_janela)

            if janelas:
                # A janela foi encontrada
                janelas[0].activate()
                return janelas[0]

            time.sleep(1)  # Aguarda por 1 segundo antes de verificar novamente

        # Se o timeout for atingido
        return None
    
    
    def limpar_caracteres(self, texto):
        """Remove '.', '/', '-', e ' ' de uma string."""
        if not isinstance(texto, str):
            return texto  # Retorna o input sem modificação se não for uma string
        return re.sub(r'[./\- ]', '', texto)
            
            
    # Função para achar o programa e abri-lo
    def open_programs(self):
        #1-Abrindo govbox do zero
        if not self.find_process('govbox.exe') and not self.find_process('Sefip.exe'):
            self.execute(r'C:\Program Files (x86)\GOVBOX\starter.exe')
            if not self.find( 'sefip', matching=0.97, waiting_time=30000):
                self.not_found('sefip')
            self.open_sefip()
        
        #2-Abrindo govbox de segundo plano
        elif not self.find_process('Sefip.exe'):
            self.open_window('govbox.exe')
            if not self.find( 'sefip', matching=0.97, waiting_time=30000):
                self.not_found('sefip')
            self.open_sefip()
        
        self.esperar_janela_abrir("sefip")
        while True:
            pyautogui.hotkey('alt', 'tab')
            if self.find( 'sefipaba', matching=0.97, waiting_time=500):
                break
        self.sefip_primary('cancelarmodalidades')
            
            
    # App govbox
    def open_sefip(self):
        if not self.find( 'sefip', matching=0.97, waiting_time=10000):
            self.not_found('sefip')
        self.click()
        
        govbox_ok = self.esperar_janela_abrir("govbox", 3)
        if govbox_ok:
            self.tab()
            self.tab()
            self.enter()
        
        data_recolhimento = self.esperar_janela_abrir("sefip", 3)
        if data_recolhimento:
            self.enter()
        
        atualizacao = self.esperar_janela_abrir("atualização", 5)
        if atualizacao:
            self.enter()
            
            
    # Achar processo govbox
    def open_window(self, processo):
        try:
            app = Application(backend='uia').connect(path=processo)
            janela = app.top_window()
            janela.set_focus()
        
        except Exception as e:
            print('Janela não encontrada.')
            if processo == 'govbox.exe':
                self.execute(r'C:\Program Files (x86)\GOVBOX\govbox.exe')
            
            elif processo == 'sefip.exe':
                self.execute(r'C:\Program Files (x86)\GOVBOX\govbox.exe')
                if not self.find( 'sefip', matching=0.97, waiting_time=30000):
                    self.not_found('sefip')
                self.click()
        
        
    def sefip_primary(self, image):
        while True:
            if self.find( image, matching=0.97, waiting_time=1000):
                self.click()
                
            pyautogui.hotkey('alt', 'x')
            self.kb_type('i')
            if self.find( 'tabelasinss', matching=0.97, waiting_time=1000):
                self.alt_f4()
                return
            
            self.key_esc()
            self.enter()
            self.wait(500)
            
            pyautogui.hotkey('alt', 'x')
            self.kb_type('i')
            if self.find( 'tabelasinss', matching=0.97, waiting_time=1000):
                self.alt_f4()
                return
            
            self.alt_f4()
            self.wait(500)
                
                
    # Abrindo movimentos
    def open_movement(self, competencies, data_pagamento):
        pyautogui.hotkey('alt', 'x')
        self.enter()
        self.kb_type('m')
        self.wait(200)
        pyautogui.hotkey('alt', 'n')
        
        if self.find( 'movimentoaberto', matching=0.97, waiting_time=3000):
            self.enter()
        self.kb_type(competencies)
        self.tab()
        self.kb_type('115')
        self.enter()
        self.tab()
        self.type_down()
        self.tab()
        self.kb_type(data_pagamento)
        self.enter()
        
        if self.find( 'confirmardados', matching=0.97, waiting_time=10000):
            self.enter()
        self.wait(500)
            
            
    # Limpando dados
    def clear_data(self):
        self.wait(1000)
        self.alt_f()
        self.kb_type('l')
        self.wait(200)
        self.enter()
        self.wait(200)
        
        if not self.find( 'processofinalizado', matching=0.97, waiting_time=10000):
            self.not_found('processofinalizado')
        self.enter()
        self.wait(200)
        
        
    # Pegando pelo .re
    def sefip_file(self, file, competencies, data_pagamento):
        self.clear_data()
        pyautogui.hotkey('alt', 'a')
        self.kb_type('i')
        
        if not self.find( 'abrirfolha', matching=0.97, waiting_time=10000):
            self.not_found('abrirfolha')
        self.wait(1000)
        self.kb_type(file)
        self.wait(200)
        self.enter()
        
        if not self.find( 'validacaofinalizada', matching=0.97, waiting_time=10000):
            self.not_found('validacaofinalizada')    
        self.enter()
        
        if not self.find( 'alocacao', matching=0.97, waiting_time=10000):  
            self.not_found('alocacao')
        self.click()
        self.open_movement(competencies, data_pagamento)
        
        if self.find( 'modalidadetrabalhador', matching=0.97, waiting_time=10000):
            self.enter()
            
            
    # Pegando pelo .bkp
    def bkp_file(self, file, competencies, data_pagamento):
        self.clear_data()
        self.alt_f()
        self.kb_type('r')
        
        if not self.find( 'restaurarok', matching=0.97, waiting_time=10000):
            self.not_found('restaurarok')
        self.tab()
        self.enter()    
        
        if not self.find( 'pasta2', matching=0.97, waiting_time=10000):
            self.not_found('pasta2')
        self.wait(1000)
        self.kb_type(file)
        self.enter()
        self.tab()
        self.type_right()
        self.type_right()
        self.enter()
        
        if not self.find( 'informacaoaba', matching=0.97, waiting_time=20000):
            self.not_found('informacaoaba')
        self.wait(500)
        self.enter()
        self.wait(200)
        self.open_movement(competencies, data_pagamento)
        
        if self.find( 'indiceserro', matching=0.97, waiting_time=5000):
            self.enter()
        else:
            self.enter()
            
        app = Application(backend='uia').connect(path='Sefip.exe')
        janela = app.window()
        janela.close()
        self.wait(3000)
        self.open_window('govbox.exe')
        
        if not self.find( 'sefip', matching=0.97, waiting_time=30000):
            self.not_found('sefip')
        self.open_sefip()
        self.enter()
        
        if self.find( 'processoconcluido', matching=0.97, waiting_time=60000):
            self.enter()
        self.wait(200)
        self.open_movement(competencies, data_pagamento)
        
        if self.find( 'informacaoaviso', matching=0.97, waiting_time=10000):
            self.enter()
        
        
    # Pegando pelo .gbk
    def gbk_file(self, file, competencies, data_pagamento):
        self.clear_data()
        self.alt_f()
        self.kb_type('r')
        
        if not self.find( 'restaurarok', matching=0.97, waiting_time=10000):
            self.not_found('restaurarok')
        self.tab()
        self.enter()    
        
        if not self.find( 'pasta2', matching=0.97, waiting_time=10000):
            self.not_found('pasta2')
        self.wait(1000)
        self.kb_type(file)
        self.enter()
        
        if not self.find( 'processofinalizado', matching=0.97, waiting_time=180000):
            self.not_found('processofinalizado')
        self.enter()
        self.open_movement(competencies, data_pagamento)
        
        if self.find( 'indiceserro', matching=0.97, waiting_time=10000):
            self.enter()
        app = Application(backend='uia').connect(path='Sefip.exe')
        janela = app.window()
        janela.close()
        self.wait(3000)
        self.open_window('govbox.exe')
        
        if not self.find( 'sefip', matching=0.97, waiting_time=30000):
            self.not_found('sefip')
        self.open_sefip()
        
        if not self.find( 'processoconcluido', matching=0.97, waiting_time=300000):  
            self.not_found('processoconcluido')
        self.enter()
        self.wait(200)
        self.open_movement(competencies, data_pagamento)
        
        if self.find( 'informacaoaviso', matching=0.97, waiting_time=10000):
            self.enter()
            
            
    def mark_modality(self, pis_type, companie, workers):
        if pis_type == '1':
            self.locate_companies(companie)
            self.alt_e()
            self.kb_type('m')
        else:
            self.locate_companies(companie)
            
            # Repetindo 3 vezes para pegar todas as 3 modalidades e garantir que nenhum funcionario fique para tras
            for contador in range(1, 4):
                # Abrindo modalidades por atalho
                pyautogui.hotkey('alt', 'a')
                self.kb_type('m')
                self.wait(500)
                self.tab()
                self.tab()
                for _ in range(contador):
                    self.type_down()
                
                # Indo ate a modalidade destino
                self.shift_tab()
                self.shift_tab()
                self.shift_tab()
                self.shift_tab()
                
                # Colocando a modalidade 9confirmação
                self.type_down()
                self.type_down()
                
                # Se contador não é igual a 3
                if contador != 3:
                    if not self.find( 'passartudo', matching=0.97, waiting_time=10000):  
                        self.not_found('passartudo')
                    self.click()
                    
                    if not self.find( 'salvarmodalidade', matching=0.97, waiting_time=2000):
                        self.alt_f4()
                        continue
                    self.click()
                    self.wait(500)
                    self.enter()
                    self.wait(500)
                    
                # Se é igual a 3, ou seja, tenho que fazer outra rotina para colocar os trabalhores e não mais tirar
                else:
                    # Escolhendo recolhimento
                    self.type_up()
                    
                    # Indo ate parte para buscar funcionario
                    self.tab()
                    self.tab()
                    self.tab()
                    # For para adicionar todos os trabalhores que precisa ao recolhimento
                    for worker in workers:
                        self.paste(worker)
                        if not self.find( 'associar_selecionado', matching=0.97, waiting_time=2000):  
                            self.not_found('associar_selecionado')
                        self.click()
                        
                        # Voltando para parte de busca
                        self.shift_tab()
                        self.shift_tab()
                        self.shift_tab()
                        self.shift_tab()
                    self.enter()
                    if not self.find( 'alteracoes_efetuadas', matching=0.97, waiting_time=2000):  
                        self.not_found('alteracoes_efetuadas')
                    self.enter()
                # Marcando participação das empresas
                    self.locate_companies(companie)
                    self.alt_e()
                    self.kb_type('m')
                    if self.find( 'trabalhadores_condicoes', matching=0.97, waiting_time=2000):
                        self.enter()
            
            
    # Função para localizar as companies
    def locate_companies(self, companie):
        self.alt_e()
        self.wait(200)
        self.kb_type('l')
        self.shift_tab()
        self.type_down()
        self.paste(companie)
        self.enter()
        if not self.find( 'play', matching=0.97, waiting_time=30000):
            self.not_found('play')
        self.double_click()
        self.wait(200)
        self.alt_f4()
        self.wait(500)
        
        
    def simulating_and_executating(self):
        # Simulando e executando
        if not self.find( 'cod.rec', matching=0.97, waiting_time=10000):  
            self.not_found('cod.rec')
        self.click()
        
        if self.find( 'simular', matching=0.97, waiting_time=10000):  
            self.wait(500)
        pyautogui.hotkey('alt', 's')
        
        # Verificando se há inconsistencia
        if self.find( 'inconsistencia', matching=0.97, waiting_time=5000):
            # Se sim, feche-a e lance uma exceção para o sistema
            self.alt_f4()
            raise Exception(f'Inconsistência no arquivo')
        
        if self.find( 'relatoriook', matching=0.97, waiting_time=10000):
            self.enter()
        pyautogui.hotkey('alt', 'c')
        
        if self.find( 'preservacao_arquivo', matching=0.97, waiting_time=3000):  
            self.enter()
            
        if self.find( 'infook', matching=0.97, waiting_time=10000):  
            self.enter()
        
        if not self.find( 'salvarsaida', matching=0.97, waiting_time=30000):  
            self.not_found('salvarsaida')
        self.enter()
        
        if self.find( 'transmissao', matching=0.97, waiting_time=2000):  
            self.enter()
        
        if self.find( 'processofinalizado2', matching=0.97, waiting_time=2000):  
            self.enter()
        
        if self.find( 'declaracaodados', matching=0.97, waiting_time=2000):  
            self.enter()
    
    
    def previous_path(self, filepath, competencia, arquivo_govbox):
        # Voltando uma pasta nos arquivos
        previous_path = os.path.dirname(filepath)

        new_path = f'{previous_path}\\Recalculo {competencia}'
        
        if os.path.exists(new_path):
            contador = 1
            while os.path.exists(f'{previous_path}\\Recalculo {competencia} {contador}'):
                contador += 1
            new_path = f'{previous_path}\\Recalculo {competencia} {contador}'
        
        os.mkdir(new_path)
        
        arquivo_recente = self.arquivos_recente(arquivo_govbox)
        arquivo_recente += '\\' + os.path.basename(arquivo_recente) + '.SFP'
        
        return new_path, arquivo_recente
    
    
    def arquivos_recente(self, atual_path):
        path = None
        arquivos = [os.path.join(atual_path, arquivo) for arquivo in os.listdir(atual_path)]
        
        if arquivos:
            # Encontra o arquivo mais recente modificado com base no tempo de modificação
            path = max(arquivos, key=os.path.getmtime)
 
        else:
            print('Não foi encontrado nenhum arquivo na pasta de downloads')
            raise Exception(f'Não foi encontrado nenhum arquivo na pasta de downloads')
        
        return path
    
    
    def open_conectividade(self, competencia, certificate, new_path, arquivo_recente, pasta_download):
        pyautogui.hotkey('alt', 'tab')
        self.wait(200)
        self.browse('https://conectividadesocialv2.caixa.gov.br/cx-postal/#/home')
        # Procurando e achando o certificado
        self.select_certificate(certificate)
        
        # Preenchendo formulario
        if not self.find( 'novamsg', matching=0.97, waiting_time=30000):
            self.not_found('novamsg')
        self.click()
        self.wait(500)
        
        if not self.find( 'selectservic', matching=0.97, waiting_time=10000):
            self.not_found('selectservic')
        self.click()
        self.type_down()
        self.enter()
        self.wait(500)
        self.tab()
        self.paste(f'Recalculo {competencia}')
        self.tab()
        self.enter()
        self.kb_type('Ceará')
        self.enter()
        self.tab()
        self.enter()
        self.kb_type('Fortaleza')
        self.enter()
        self.tab()
        self.enter() #abrindo enviar arquivos do formulario da conectividade
        if not self.find( 'pasta2', matching=0.97, waiting_time=2000):
            self.not_found('pasta2')
        self.paste(arquivo_recente)
        self.enter()
        self.wait(500)
        self.tab()
        self.tab()
        self.tab()  
        self.enter() #enviando formulario
        
        if not self.find( 'envioarquivos', matching=0.97, waiting_time=10000):
            self.not_found('envioarquivos')
        self.enter() #baixando xml
        self.tab()
        self.enter() #baixando pdf
        
        if self.find( 'permitir', matching=0.97, waiting_time=2000):
            self.double_click() #apertando em permitir downloads caso apareça
        self.wait(3000)
        
        # Mandando arquivos baixados recentemente em pasta_download para new_path
        caminho_xml = self.recent_download(pasta_download, new_path, 2)
        pyautogui.hotkey('ctrl', 'f4')
        
        return caminho_xml
                
                
    def recent_download(self, pasta_download, pasta_destino, quantity):
        try:
            # Lista todos os arquivos na pasta de downloads
            arquivos = [os.path.join(pasta_download, arquivo) for arquivo in os.listdir(pasta_download)]
            # Verifica se tem arquivos encontrados
            caminho_return = None
            for _ in range(quantity):
                if arquivos:
                    # Encontra o arquivo mais recente modificado com base no tempo de modificação
                    arquivo_mais_recente = max(arquivos, key=os.path.getmtime)
                    # Construa o caminho de destino
                    caminho_destino = os.path.join(pasta_destino, os.path.basename(arquivo_mais_recente))
                    # Pegando o caminho que preciso retornar no codigo
                    if os.path.basename(arquivo_mais_recente).endswith('xml'):
                        caminho_return = caminho_destino
                    # Move o arquivo para a pasta de destino
                    shutil.move(arquivo_mais_recente, caminho_destino)
                    # Removo da lista de arquivos
                    arquivos.remove(arquivo_mais_recente)

                    print(f'O arquivo mais recente foi movido para: {caminho_destino}')

                else:
                    print('Não foi encontrado nenhum arquivo na pasta de downloads')
                    raise Exception(f'Não foi encontrado nenhum arquivo na pasta de downloads')
            return caminho_return

        except Exception as e:
            print(f'Ocorreu um erro ao listar os arquivos: {e}')
            
            
    # Função para selecionar o certificates:
    def select_certificate(self, certificate):
        if self.find( 'selecionecertificado', matching=0.97, waiting_time=8000):
            while True:
                self.tab()
                self.enter()
                self.wait(200)
                self.tab()
                self.type_right()
                self.tab()
                self.tab()
                self.type_down()
                self.tab()
                self.control_c()
                self.alt_f4()
                self.wait(200)
                pyautogui.hotkey('shift', 'tab')
                if self.get_clipboard() == certificate:
                    self.enter()
                    break
                self.type_down()
                
                
    def criando_pasta(self, pasta_relatorios):
        # Criando pasta e entrando no arquivo com o os
        if not os.path.exists(pasta_relatorios):
            os.mkdir(pasta_relatorios)
        else:
            shutil.rmtree(pasta_relatorios)
            os.mkdir(pasta_relatorios)
                        
                        
    def gerando_relatorios(self, caminho_xml, arquivo_recente, new_path, pasta_relatorios, arquivo_govbox):
        self.alt_r()
        self.kb_type('g')
        self.kb_type('i')
        if not self.find( 'pasta2', matching=0.97, waiting_time=2000):
            self.not_found('pasta2')
        self.paste(caminho_xml)
        self.enter()
        
        if not self.find( '1156', matching=0.97, waiting_time=10000):
            self.not_found('1156')
        self.wait(100)
        self.enter()
        
        if not self.find( 'pasta', matching=0.97, waiting_time=10000):
            self.not_found('pasta')
        self.enter()
        
        if not self.find( 'selecaoarquivosaida', matching=0.97, waiting_time=10000):
            self.not_found('selecaoarquivosaida')
        self.wait(100)
        self.paste(arquivo_recente)
        self.enter()
        self.kb_type('v')
        
        if not self.find( 'grfpdf', matching=0.97, waiting_time=10000):
            self.not_found('grfpdf')
        self.kb_type('p')
        self.wait(500)
        self.control_c()
        self.type_down()
        self.enter()
        self.wait(1000)
        self.enter()
        self.alt_f4()
        self.wait(200)
        
        # Gerando analitico grf
        self.alt_r()
        self.kb_type('m')
        self.kb_type('n')
        if not self.find( 'analiticogrf', matching=0.97, waiting_time=10000):
            self.not_found('analiticogrf')
        self.kb_type('p')
        
        if not self.find( 'caminhopdf', matching=0.97, waiting_time=10000):
            self.not_found('caminhopdf')
        self.enter()
        
        if not self.find( 'processofinalizado', matching=0.97, waiting_time=10000):
            self.not_found('processofinalizado')
        self.enter()
        self.wait(200)
        self.alt_f4()
        self.wait(200)
        
        # Gerar RE
        self.kb_type('r')
        if not self.find( 'relatoriorepdf', matching=0.97, waiting_time=10000):
            self.not_found('relatoriorepdf')
        self.kb_type('g')
        
        if not self.find( 'caminhopdf', matching=0.97, waiting_time=10000):
            self.not_found('caminhopdf')
        self.enter()
        
        if not self.find( 'processofinalizado', matching=0.97, waiting_time=10000):
            self.not_found('processofinalizado')
        self.enter()
        self.wait(200)
        self.alt_f4()
        self.wait(200)
        
        # Gerar comprovante de previdencia
        self.kb_type('m')
        if not self.find( 'comprovantepdf', matching=0.97, waiting_time=10000):
            self.not_found('comprovantepdf')
        self.kb_type('g')
        
        if not self.find( 'caminhopdf', matching=0.97, waiting_time=10000):
            self.not_found('caminhopdf')
        self.enter()
        
        if not self.find( 'processofinalizado', matching=0.97, waiting_time=10000):
            self.not_found('processofinalizado')
        self.enter()
        self.wait(200)
        self.alt_f4()
        self.wait(200)
        
        # Gerar analitico gps
        self.kb_type('a')
        if not self.find( 'movimentopdf', matching=0.97, waiting_time=10000):
            self.not_found('movimentopdf')
        self.kb_type('g')
        
        if not self.find( 'caminhopdf', matching=0.97, waiting_time=10000):
            self.not_found('caminhopdf')
        self.enter()
        
        if not self.find( 'processofinalizado', matching=0.97, waiting_time=10000):
            self.not_found('processofinalizado')
        self.enter()
        self.wait(200)
        self.key_esc()
        self.key_esc()
        self.wait(100)
        
        self.recent_download(arquivo_govbox, new_path, 1)
        # Mandando relatorios baixados para a pasta paths
        self.recent_download(pasta_relatorios, new_path, 5)
        self.wait(1000)
        
        
    def not_found(self, label):
        raise Exception(f'Element not found: {label}')


if __name__ == '__main__':
    Bot.main()