from botcity.core import DesktopBot
import pygetwindow as gw
import pyautogui
import pandas as pd
import sys


class Bot(DesktopBot):
    
    def action(self, execution=None):
        filepath = r'C:\Users\davi.inov\Desktop\Projetos\Departamento_Contabil\08_cadastro_imobilizado_athenas\data\IMOB BASE.xlsx'
        df_codigos = pd.read_excel(filepath, sheet_name='Planilha1', dtype=str)
        df_empresas = pd.read_excel(filepath, sheet_name='Planilha2', dtype=str)
        if 'Status' not in df_empresas.columns:
            df_empresas['Status'] = ''
            
        self.open_athenas()
        codigo_antigo = None
        
        for index, row in df_empresas.iterrows():
            status_atual = str(row['Status']).strip()
            if status_atual != 'nan' and status_atual:
                continue
            
            print(f'Abrindo empresa {row['Empresa']}')
            self.abrindo_empresa(row['Codigo'])
            self.abrindo_cadastro_imobilizado()
            
            while True:
                codigo_antigo = self.cadastro_imobilizado(df_codigos, codigo_antigo)
                if codigo_antigo == 'Proxima empresa':
                    break
                
    
    def open_athenas(self):
        status_app = self.focar_aplicativo_por_nome('Athenas')
        if status_app == 'Nao achou':
            print('Encerrando robo. Abra o athenas e tente novamente.')
        
        if not self.find( 'athenas_logo', matching=0.97, waiting_time=10000):
            self.not_found('athenas_logo')
        pyautogui.hotkey('alt', 'y', '3')
        self.wait(500)
        self.kb_type('y11')
        self.wait(500)
        self.kb_type('2024')
        self.enter()
        self.wait(500)
                 
                 
    def abrindo_empresa(self, companie):
        self.key_f8()
        self.paste(companie)
        self.enter()
        self.wait(3000)
    
    
    def abrindo_cadastro_imobilizado(self):
        pyautogui.hotkey('alt', 'c', 'i')
        if not self.find( 'cadastro_imobilizado', matching=0.97, waiting_time=10000):
            self.not_found('cadastro_imobilizado')
        self.wait(500)
        pyautogui.hotkey('alt', 'a')
        self.wait(500)
        pyautogui.hotkey('alt', 'o')
        
    def focar_aplicativo_por_nome(self, nome_janela):
        janelas = gw.getWindowsWithTitle(nome_janela)

        if len(janelas) > 0:
            janela = janelas[0]
            janela.activate()
            return 'Achou'
        else:
            return 'Nao achou'
        
    
    def cadastro_imobilizado(self, df, codigo_antigo):
        self.verificar_conta_invalida()
        codigo_atual = self.verificar_codigo_atual(codigo_antigo)
        
        if codigo_atual != 'Proxima empresa':
            status_codigo = self.verificar_ativo()
            if status_codigo == 'Ativo':
                self.verificar_e_trocar(df)
            self.salvar_e_proximo()
            
        return codigo_atual
    
    
    def verificar_conta_invalida(self):
        while True:
            if self.find( 'conta_invalida', matching=0.97, waiting_time=500):
                self.enter()
                continue
            break
        
        
    def verificar_codigo_atual(self, codigo_antigo):    
        pyautogui.hotkey('alt', 'a')
        if self.find( 'conta_invalida', matching=0.97, waiting_time=500):
            self.enter()
        pyautogui.press('esc')
        if self.find( 'conta_invalida', matching=0.97, waiting_time=500):
            self.enter()
        pyautogui.hotkey('alt', 'c')
        pyautogui.hotkey('alt', 'a')
        pyautogui.hotkey('ctrl', 'c')
        codigo_atual = self.get_clipboard()
        if codigo_atual == codigo_antigo:
            return 'Proxima empresa'
        return codigo_atual
    
        
    def salvar_e_proximo(self):
        pyautogui.hotkey('alt', 'o')
        if self.find( 'mudanca_taxa_depreciacao', matching=0.97, waiting_time=800):
            pyautogui.hotkey('alt', 'c')
        if not self.find( 'proximo_codigo', matching=0.97, waiting_time=10000):
            self.not_found('proximo_codigo')
        self.click()
        
        
    def verificar_ativo(self):
        if not self.find( 'ativo', matching=0.97, waiting_time=500):
            return 'Nao ativo'
        return 'Ativo'
        
        
    def verificar_e_trocar(self, df):
        if not self.find( 'conta_contabil', matching=0.97, waiting_time=10000):
            self.not_found('conta_contabil')
        self.double_click_relative(x=10, y=20)
        if self.find( 'conta_invalida', matching=0.97, waiting_time=500):
            self.enter()
            
        for i in range(3):
            if i == 1:
                self.paste('410101050000003')
                pyautogui.press('tab')
                if self.find( 'conta_invalida', matching=0.97, waiting_time=500):
                    self.enter()
                continue
                
            pyautogui.hotkey('ctrl', 'c')
            codigo_antes = self.get_clipboard()
            codigo_depois = self.bater_codigo(df, codigo_antes)
            self.paste(codigo_depois)
            pyautogui.press('tab')
            if self.find( 'conta_invalida', matching=0.97, waiting_time=500):
                self.enter()
            
        
    def bater_codigo(self, df, codigo_antes):
        linha = df[df['Antes'] == codigo_antes].index
        if not linha.empty:
            codigo_depois = df.loc[linha[0], 'Depois']
        else:
            print('Não bateu')
            codigo_depois = codigo_antes
            
        return codigo_depois
    
        
    def not_found(self, label):
        raise Exception(f'Element not found: {label}')


if __name__ == '__main__':
    Bot.main()