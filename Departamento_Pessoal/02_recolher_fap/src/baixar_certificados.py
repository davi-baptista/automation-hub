from botcity.core import DesktopBot
from pywinauto import Application
import pyautogui
import pandas as pd
import os
import re


class Bot(DesktopBot):
    
    
    def action(self, execution=None):
        senhas_fixas = ['1234', '123456', '12345678']
        pasta_certificados = r'c:\Users\davi.inov\Desktop\CERTIFICADOS NOVOS'
        caminho_excel_retorno = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Pessoal\06. Procuracao_SPE\Saida\relatorio.xlsx'
        
        # Primeiro, mapeia todos os arquivos .txt relevantes em cada diretório
        txt_map = self.mapear_arquivos_txt(pasta_certificados)
        
        if os.path.exists(caminho_excel_retorno):
            df = pd.read_excel(caminho_excel_retorno)
        else:
            df = pd.DataFrame(columns=['Nome do Certificado', 'Status da Procuração', 'Caminho do certificado'])
            
        # Percorre o diretório e subdiretórios em busca de certificados
        for root, dirs, files in os.walk(pasta_certificados):
            # Verifica se 'vencido' não faz parte do caminho atual
            if 'vencido' in root.lower():
                continue
            
            for file in files:
                # Verifica se o arquivo tem uma das extensões desejadas (.p12, .pfx)
                if file.endswith(('.p12', '.pfx')):
                    caminho_completo = os.path.join(root, file)
                    
                    # Verifica se o caminho_completo ja existe no excel, se ja existir, pula pro proximo
                    
                    
                    if df['Caminho do certificado'].isin([caminho_completo]).any():
                        continue
                        
                    # Chama capturar_arquivos para processar o certificado encontrado
                    status = self.capturar_arquivos(root, file, caminho_completo, senhas_fixas, txt_map)
                    
                    if status != 'Nao achou nenhuma senha':
                        self.execute(r'C:\Windows\System32\certmgr.msc')
                        
                        if not self.find( "certmgr", matching=0.97, waiting_time=30000):
                            self.not_found("certmgr")
                        self.type_right()
                        self.type_right()
                        self.type_right()
                        
                        if not self.find( "emitido_para", matching=0.97, waiting_time=30000):
                            self.not_found("emitido_para")
                        self.tab()
                        
                        pyautogui.hotkey('alt', 'o', 'a')
                        if not self.find( "informacoes_certificado", matching=0.97, waiting_time=30000):
                            self.not_found("informacoes_certificado")
                            
                        status_procuracao = 'Certificado invalido'
                        if not self.find( "expirou_ou_nao_valido", matching=0.97, waiting_time=500):
                            status_procuracao = 'Certificado valido'
                        
                        nome_certificado = self.pegar_certificado()
                        self.alt_f4()
                        self.delete()
                        if not self.find( "excluir_certificado", matching=0.97, waiting_time=5000):
                            self.not_found("excluir_certificado")
                        self.enter()
                        self.wait(500)
                        self.alt_f4()
                        
                        df = self.adicionar_ao_excel(df, nome_certificado, status_procuracao, caminho_completo, caminho_excel_retorno)
                              
        
    def mapear_arquivos_txt(self, pasta):
        txt_map = {}
        for root, dirs, files in os.walk(pasta):
            if 'vencido' in root.lower():
                continue
            for file in files:
                if file.endswith('.txt') or not re.match(r'.*\.[a-zA-Z]{3}$', file):
                    if root not in txt_map:
                        txt_map[root] = []
                    txt_map[root].append(file)
        return txt_map
        
        
    def capturar_arquivos(self, root, file, caminho_completo, senhas_fixas, txt_map):
        self.execute(caminho_completo)
        if not self.find( "certificados", matching=0.97, waiting_time=30000):
            self.not_found("certificados")
        self.enter()
        
        # if not self.find( "arquivo_importado", matching=0.97, waiting_time=30000):
        #     self.not_found("arquivo_importado")
        self.enter()
        
        # if not self.find( "protecao_cha
        # e")
        
        achou = None
        for senha in senhas_fixas:
            self.control_a()
            self.paste(senha)
            self.enter()
            if self.find( "senha_incorreta", matching=0.97, waiting_time=1000):
                self.enter()
            else:
                achou = f'Senha {senha}'
                self.enter()
                self.enter()
                if not self.find( "importacao_exito", matching=0.97, waiting_time=30000):
                    self.not_found("importacao_exito")
                self.enter()
                return achou
        
        # Remove a extensão do arquivo antes de chamar search_password
        nome_arquivo_sem_extensao, _ = os.path.splitext(file)
        senhas_potenciais = self.search_password(nome_arquivo_sem_extensao, senhas_fixas)
        
        # Usa o mapeamento de arquivos .txt para buscar mais senhas
        if root in txt_map:
            for file_txt in txt_map[root]:
                nome_txt_sem_extensao, _ = os.path.splitext(file_txt)
                senhas_potenciais.extend(self.search_password(nome_txt_sem_extensao, senhas_fixas))
                
        # Remove possíveis duplicatas mantendo a ordem
        senhas_potenciais = list(dict.fromkeys(senhas_potenciais))
        
        print(f"Certificado: {file}, Senhas Potenciais: {senhas_potenciais}")
        
        achou = None
        for senha in senhas_potenciais:
            self.control_a()
            self.paste(senha)
            self.enter()
            if self.find( "senha_incorreta", matching=0.97, waiting_time=1000):
                self.enter()
            else:
                achou = f'Senha {senha}'
                self.enter()
                self.enter()
                if not self.find( "importacao_exito", matching=0.97, waiting_time=30000):
                    self.not_found("importacao_exito")
                self.enter()
                return achou
        
        self.alt_f4()
        return 'Não achou nenhuma senha'
        
        
    def search_password(self, texto, senhas_fixas):
        # Inicializando o dicionário de resultados
        resultado_senha = []
        
        # Busca por padrões que incluam "senha", seguidos por caracteres como : - ou espaços, e então captura a senha
        padrao = re.compile(r'senha[s]?[\s:_-]*\s*(\S+)|(\S+)\s*senha[s]?[\s:_-]*\s*(\S+)', re.IGNORECASE)
        resultados = padrao.findall(texto)
        
        for resultado in resultados:
            for possivel_senha in filter(None, resultado):
                if possivel_senha not in senhas_fixas and possivel_senha not in resultado_senha:
                    resultado_senha.append(possivel_senha)
                
        # Capturando a primeira e a última palavra do texto
        palavras = texto.split()
        if palavras:
            resultado_senha.append(palavras[0])
            resultado_senha.append(palavras[-1])
                
        return resultado_senha
    
    
    def pegar_certificado(self):
        self.tab()
        self.type_right()
        self.tab()
        self.tab()
        self.type_down()
        self.type_down()
        self.type_down()
        self.type_down()
        self.type_down()
        self.type_down()
        self.type_down()
        self.tab()
        self.type_up()
        pyautogui.hotkey('ctrl', 'right')
        pyautogui.hotkey('ctrl', 'right')
        self.hold_shift()
        self.key_end()
        self.type_left()
        self.release_shift()
        self.control_c()
        
        return self.get_clipboard()
        
        
    def adicionar_ao_excel(self, df, nome_certificado, status_procuracao, caminho_certificado, caminho_excel):
        # Cria um DataFrame para a nova linha
        nova_linha_df = pd.DataFrame({
            'Nome do Certificado': [nome_certificado],
            'Status da Procuração': [status_procuracao],
            'Caminho do certificado': [caminho_certificado]
        })
        
        # Usa pd.concat para adicionar a nova linha ao DataFrame existente
        df = pd.concat([df, nova_linha_df], ignore_index=True)

        # Salva o DataFrame atualizado em um arquivo Excel
        df.to_excel(caminho_excel, index=False)
        return df
        
        
    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()