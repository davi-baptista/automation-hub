from botcity.core import DesktopBot
import pyautogui
import os
import shutil
import subprocess
from datetime import datetime, timedelta

class Bot(DesktopBot):
    
    
    def __init__(self, nome_certificado):
        super().__init__()  # Chama o construtor da classe base se necessário
        self.nome_certificado = nome_certificado
        
        
    def action(self, execution=None):
        
        cnpj_controller = '05.494.895/0001-81'
        email = 'depes@controller-rnc.com.br'
        pasta_downloads = r'C:\Users\davi.inov\Downloads'
        pasta_saida_procuracoes = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\06_procuracao_spe\output\Procuracoes'
        
        try:
            self.login(self.nome_certificado)
            status_login = self.rotinas_login()
            if status_login == 'Verificacao_duas_etapas':
                self.alt_f4()
                return 'Verificacao em duas etapas', ''
            status_procuracao = self.conferir_cnpj(cnpj_controller)
            print(status_procuracao)
            
            if status_procuracao == 'Nenhuma procuracao ativa':
                self.nova_procuracao(status_procuracao)
                self.procuracao_outorgante(email)
                self.procuracao_outorgado(cnpj_controller, email)
                self.procuracao_servicos()
                self.procuracao_vigencia()
                status = self.procuracao_gerar_procuracao()
                print(status)
                if status == 'Procuracao salva, fazer assinatura novamente':
                    self.filtrar_por_cnpj_e_status(cnpj_controller, 'pendente')
                    status_procuracao = self.fazer_assinatura_completa()
                    
                elif status == 'Autorizar navegador':
                    self.autorizar_navegador()
                    status = self.assinar()
                    if status == 'Assinada':
                        status_procuracao = self.conferir_procuracao()
                    elif status == 'Assinada e baixada':
                        status_procuracao = 'Procuracao feita e assinada, mover pasta'
                    else:
                        status_procuracao = 'Erro inesperado ao assinar procuracao'
                        
                else:
                    status_procuracao = status
                    
            elif status_procuracao == 'Procuracao feita e sem assinatura, refaze-lo':
                    status_procuracao = self.excluir_procuracao()
            
            caminho_procuracao = ''
            if status_procuracao == 'Procuracao feita e assinada, mover pasta':
                status_procuracao, caminho_procuracao = self.mover_ultima_procuracao(pasta_downloads, pasta_saida_procuracoes, self.nome_certificado)
                
            elif status_procuracao == 'Procuracao feita e assinada':
                self.abrir_nova_aba()
                status_download = self.baixar_procuracao(cnpj_controller)
                if status_download == 'Procuracao baixada com sucesso':
                    status_procuracao, caminho_procuracao = self.mover_ultima_procuracao(pasta_downloads, pasta_saida_procuracoes, self.nome_certificado)
                    
                elif status_download == 'Erro ao baixar procuracao':
                    status_procuracao = 'Procuracao feita, assinada e nao baixada'
                    
                else:
                    status_procuracao = status_download
                
            self.alt_f4()
            return status_procuracao, caminho_procuracao
               
        except Exception as e:
            print(f'Erro no site {e}')
            fechar_processo('firefox.exe')
            return 'Erro desconhecido ao rodar o site', ''


    def login(self, certificate):
        self.execute(r'C:\Program Files\Mozilla Firefox\private_browsing.exe')
        if not self.find( "firefox_logo", matching=0.97, waiting_time=30000):
            self.not_found("firefox_logo")
        self.kb_type("https://spe.sistema.gov.br/")
        self.enter()
        
        if not self.find( "entrar_com_govbr", matching=0.92, waiting_time=30000):
            self.not_found("entrar_com_govbr")
        self.click()
        
        if not self.find( "seu_certificado_digital", matching=0.97, waiting_time=30000):
            self.not_found("seu_certificado_digital")
        self.click()
        
        if not self.find( "detalhes_certificado", matching=0.97, waiting_time=30000):
            self.not_found("detalhes_certificado")
        self.kb_type(certificate)
        self.enter()
        
        
    def rotinas_login(self):
        nao_repetir = []
        # Logica para garantir que vou autorizar e concordar com os termos, independente da ordem em que aparecam
        while True:
            if 'autorizar' not in nao_repetir:
                if self.find( "autorizar", matching=0.97, waiting_time=5000):
                    self.click()
                    nao_repetir.append('autorizar')
                    continue
            if 'termos_de_uso' not in nao_repetir:
                if self.find( "termo_de_uso_concordo", matching=0.97, waiting_time=5000):
                    self.click()
                    nao_repetir.append('termos_de_uso')
                    continue
            break
        
        if self.find( "verificacao_duas_etapas", matching=0.97, waiting_time=5000):
            return "Verificacao_duas_etapas"
        
        
    def filtrar_por_cnpj_e_status(self, cnpj, status):
        print('Filtrando por cnpj e status')
        while True:
            if not self.find( "cnpj_outorgado", matching=0.97, waiting_time=30000):
                self.not_found("cnpj_outorgado")
            self.wait(500)
            self.click_relative(x=20, y=50)
            self.paste(cnpj)
            self.control_a()
            self.control_c()
            if self.get_clipboard() == cnpj:
                break
            
        self.enter()
        self.tab()
        self.tab()
        self.tab()
        self.tab()
        self.tab()
        
        if status == 'ativa':
            self.type_down()
            
        elif status == 'pendente':
            self.type_down()
            self.type_down()
            
        self.wait(500)
        self.enter()
        if not self.find( "filtrar", matching=0.97, waiting_time=30000):
            self.not_found("filtrar")
        self.click()
        
        carregando = self.find( "carregando", matching=0.97, waiting_time=2000)
        while carregando:
            if not self.find( "carregando", matching=0.97, waiting_time=500):
                break
        
        if self.find( "nenhuma_empresa_encontrada", matching=0.97, waiting_time=3000):
            return 'Nenhuma empresa encontrada'
        return 'Empresa encontrada'
        
        
    def conferir_cnpj(self, cnpj):
        print('Conferindo cnpj')
        if self.find( "nenhuma_empresa_encontrada", matching=0.97, waiting_time=5000):
            return 'Nenhuma procuracao ativa'
        
        status_filtro = self.filtrar_por_cnpj_e_status(cnpj, 'ativa')
        if status_filtro == 'Nenhuma empresa encontrada':
            self.shift_tab()
            self.shift_tab()
            self.shift_tab()
            self.type_down()
            self.type_down()
            self.enter()
            
            if not self.find( "filtrar", matching=0.97, waiting_time=30000):
                self.not_found("filtrar")
            self.click()
            
            carregando = self.find( "carregando", matching=0.97, waiting_time=2000)
            while carregando:
                if not self.find( "carregando", matching=0.97, waiting_time=500):
                    break
                
            if self.find( "nenhuma_empresa_encontrada", matching=0.97, waiting_time=1000):
                return 'Nenhuma procuracao ativa'
            else:
                return 'Procuracao feita e sem assinatura, refaze-lo'
        else:
            return 'Procuracao feita e assinada'
    
    
    def nova_procuracao(self, status):
        print('Fazendo nova procuracao')
        if status == 'Procuracao ja feita':
            print('Procuracao ja feita')
            return
        
        if not self.find( "nova_procuracao", matching=0.97, waiting_time=30000):
            self.not_found("nova_procuracao")
        self.click()
        
        
    def procuracao_outorgante(self, email):
        print('1/5')
        if not self.find( "email", matching=0.97, waiting_time=30000):
            self.not_found("email")
        self.click_relative(x=20, y=25)
        self.control_a()
        self.paste(email)
        self.tab()
        self.kb_type(email)
        
        if not self.find( "avancar", matching=0.97, waiting_time=30000):
            self.not_found("avancar")
        self.click()
        
        
    def procuracao_outorgado(self, cnpj, email):
        print('2/5')
        if not self.find( "marcar_cnpj", matching=0.97, waiting_time=30000):
            self.not_found("marcar_cnpj")
        self.click_relative(x=-20, y=0)
        
        self.tab()
        self.paste(cnpj)
        
        self.tab()
        self.control_a()
        self.paste(email)
        
        self.tab()
        self.kb_type(email)
        
        self.tab()
        self.space()    
        
        if not self.find( "avancar", matching=0.97, waiting_time=30000):
            self.not_found("avancar")
        self.click()
        
        
    def procuracao_servicos(self):
        print('3/5')
        if not self.find( "domicilio_eletronico_trabalhista", matching=0.97, waiting_time=30000):
            self.not_found("domicilio_eletronico_trabalhista")
        self.click_relative(x=20, y=55)
        
        if not self.find( "domicilio_eletronico_trabalhista", matching=0.97, waiting_time=30000):
            self.not_found("domicilio_eletronico_trabalhista")
        self.click()
        self.page_down()
        
        if not self.find( "fgts_digital", matching=0.97, waiting_time=30000):
            self.not_found("fgts_digital")
        self.click_relative(x=20, y=55)
        
        if not self.find( "marcar_todos", matching=0.97, waiting_time=30000):
            self.not_found("marcar_todos")
        self.click()
        
        if not self.find( "fgts_digital", matching=0.97, waiting_time=30000):
            self.not_found("fgts_digital")
        self.click()
        
        if not self.find( "avancar", matching=0.97, waiting_time=30000):
            self.not_found("avancar")
        self.click()
        
        
    def procuracao_vigencia(self):
        print('4/5')
        if not self.find( "avancar", matching=0.97, waiting_time=30000):
            self.not_found("avancar")
        self.click()
    
    
    def procuracao_gerar_procuracao(self):
        try:
            print('5/5')
            if not self.find( "spe", matching=0.97, waiting_time=30000):
                self.not_found("spe")
            self.page_down()
            
            if not self.find( "assinar", matching=0.97, waiting_time=5000):
                self.not_found("assinar")
            self.click()
            
            if self.find( "navegador_nao_autorizado", matching=0.97, waiting_time=5000):
                return 'Autorizar navegador'
            
            status = self.conferir_procuracao()
            return status
        except:
            return 'Erro inesperado ao fazer procuracao'
        
        
    def conferir_procuracao(self):
        if self.find( "procuracao_salva", matching=0.97, waiting_time=10000):
            if self.find( "nao_foi_possivel_assinar", matching=0.97, waiting_time=1000):
                return 'Procuracao salva, fazer assinatura novamente'
            return 'Procuracao salva, conferir assinatura'
                
        if self.find( "procuracao_salva_assinada", matching=0.97, waiting_time=30000):
            return 'Procuracao feita e assinada'
        
        return 'Erro na procuracao'
        
        
    def aplicar_assinatura(self, contador=1):
        if self.find( "acoes", matching=0.97, waiting_time=5000):
            print('Aplicando assinatura')
            if contador == 2:
                while True:
                    if not self.find( "navegador_nao_autorizado", matching=0.97, waiting_time=1000):
                        break
                    
            if self.find( "acoes", matching=0.97, waiting_time=3000):
                self.click()
            self.page_down()
            self.page_down()
            
            if not self.find( "acoes_assinar", matching=0.97, waiting_time=30000):
                self.not_found("acoes_assinar")
            self.click()
        
        
    def autorizar_navegador(self):
        print('Autorizando navegador')
        if self.find( "navegador_nao_autorizado", matching=0.97, waiting_time=10000):
            if not self.find( "aqui", matching=0.97, waiting_time=30000):
                self.not_found("aqui")
            self.click()
            
            if not self.find( "endereco_do_assinador", matching=0.97, waiting_time=30000):
                self.not_found("endereco_do_assinador")
            self.click()
            
            if not self.find( "avancado", matching=0.97, waiting_time=30000):
                self.not_found("avancado")
            self.click()
            
            if not self.find( "aceitar_risco", matching=0.97, waiting_time=30000):
                self.not_found("aceitar_risco")
            self.click()
            
            if not self.find( "procedimento_sucesso", matching=0.97, waiting_time=30000):
                self.not_found("procedimento_sucesso")
            pyautogui.hotkey('ctrl', 'f4')
        
        
    def assinar(self):
        print('Assinando')
        if not self.find( "spe", matching=0.97, waiting_time=30000):
            self.not_found("spe")
        self.page_down()
        self.page_down()
        
        if self.find( "assinar", matching=0.95, waiting_time=2000):
            self.click()
        else:
            return 'Erro ao assinar procuracao'
        self.wait(1000)
        
        if self.find( "procuracao_baixada", matching=0.97, waiting_time=15000):
            return 'Assinada e baixada'
        
        if self.find( "procuracao_assinada", matching=0.97, waiting_time=15000):
            return 'Assinada'
        return 'Nao assinada'
    
    
    def fazer_assinatura_completa(self):
        try:
            self.aplicar_assinatura()
            self.autorizar_navegador()
            self.aplicar_assinatura(2)
            status_assinatura = self.assinar()
            if status_assinatura == 'Assinada':
                return 'Procuracao feita e assinada'
            elif status_assinatura == 'Assinada e baixada':
                return 'Procuracao feita e assinada, mover pasta'
            return 'Procuracao feita e nao assinada'
        except:
            return 'Erro inesperado ao assinar procuracao'
    
    
    def excluir_procuracao(self):
        try:
            if self.find( "acoes", matching=0.97, waiting_time=5000):
                print('Excluindo procuracao')
                self.click()
                self.page_down()
                
                if not self.find( "acoes_excluir", matching=0.97, waiting_time=30000):
                    self.not_found("acoes_excluir")
                self.click()
                
                if not self.find( "excluir_procuracao", matching=0.97, waiting_time=30000):
                    self.not_found("excluir_procuracao")
                self.click()
                
                if self.find( "procuracao_excluida", matching=0.97, waiting_time=15000):
                    return 'Procuracao excluida com sucesso'
                return 'Erro ao excluir procuracao'
        except:
            return 'Erro desconhecido ao excluir procuracao'
    
    
    def abrir_nova_aba(self):
        print('Abrindo nova aba')
        self.control_t()
        
        if not self.find( "firefox_logo", matching=0.97, waiting_time=30000):
            self.not_found("firefox_logo")
        self.paste('https://spe.sistema.gov.br/procuracao')
        self.enter()
    
    
    def baixar_procuracao(self, cnpj):
        print('Baixando procuracao')
        status_filtro = self.filtrar_por_cnpj_e_status(cnpj, 'ativa')
        if status_filtro == 'Nenhuma empresa encontrada':
            return 'Procuracao nao foi feita'
        
        if not self.find( "acoes", matching=0.97, waiting_time=30000):
            self.not_found("acoes")
        self.click()
        self.page_down()
        self.page_down()
        self.page_down()
        
        if not self.find( "acoes_download", matching=0.97, waiting_time=30000):
            self.not_found("acoes_download")
        self.click()
        
        if not self.find( "procuracao_baixada", matching=0.97, waiting_time=30000):
            return 'Erro ao baixar procuracao'
        return 'Procuracao baixada com sucesso'
    
    
    def mover_ultima_procuracao(self, download_path, destino_path, novo_nome):
        print('Movendo pastas')
        agora = datetime.now()
        arquivos_procuracao = []

        for arquivo in os.listdir(download_path):
            if 'procuracao' in arquivo.lower() and arquivo.lower().endswith('.pdf'):
                caminho_completo = os.path.join(download_path, arquivo)
                data_modificacao = datetime.fromtimestamp(os.path.getmtime(caminho_completo))
                
                # Verifica se a data de modificação está dentro do intervalo de 1 minuto da hora atual
                if agora - timedelta(minutes=1) <= data_modificacao <= agora + timedelta(minutes=1):
                    arquivos_procuracao.append(arquivo)
                    
        if not arquivos_procuracao:
            print('Nenhum arquivo de procuração encontrado ou modificado recentemente.')
            return 'Procuracao feita, assinada e nao baixada', ''
        
        # Ordena a lista de arquivos pela data de modificação, do mais recente ao mais antigo
        arquivos_procuracao.sort(key=lambda arquivo: os.path.getmtime(os.path.join(download_path, arquivo)), reverse=True)
        
        # Pega o caminho completo do arquivo mais recente
        ultimo_arquivo = os.path.join(download_path, arquivos_procuracao[0])
        
        # Define o novo caminho com o novo nome
        nome_seguro = novo_nome.translate({ord(c): "-" for c in '<>:"/\\|?*'})
        novo_caminho = os.path.join(destino_path, nome_seguro + ".pdf")
        
        if novo_caminho and os.path.exists(novo_caminho):
            print('Arquivo ja existe e esta na pasta')
            return 'Procuracao feita, assinada e baixada', novo_caminho
        
        # Move e renomeia o arquivo
        shutil.move(ultimo_arquivo, novo_caminho)
        print(f"Arquivo '{arquivos_procuracao[0]}' movido e renomeado para '{novo_caminho}'.")
        return 'Procuracao feita, assinada e baixada', novo_caminho
    
    def not_found(self, label):
        print(f"Element not found: {label}")
        raise Exception(f'Element not found {label}')


def fechar_processo(nome_processo):
    try:
        # Constrói o comando para forçar o encerramento do processo pelo nome
        comando = f"taskkill /im {nome_processo} /f"
        subprocess.run(comando, check=True, shell=True)
        print(f"Processo '{nome_processo}' encerrado com sucesso.")
    except subprocess.CalledProcessError as e:
        print(f"Erro ao tentar encerrar o processo '{nome_processo}'. Detalhes do erro: {e}")
        
        
def main(nome_certificado=''):
    bot_procuracao = Bot(nome_certificado)
    status_procuracao, caminho_procuracao = bot_procuracao.action()
    
    return status_procuracao, caminho_procuracao


if __name__ == '__main__':
    nome_certificado = 'controller'
    main(nome_certificado)