from botcity.web import WebBot, By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import threading


class MyWebBot(WebBot):
    
    
    def __init__(self, filepath, log_callback, *args, **kwargs):
        super(MyWebBot, self).__init__(*args, **kwargs)
        self.filepath = filepath
        self.saida_excel = f'{os.path.dirname(self.filepath)}/saida.xlsx'
        self.log_callback = log_callback
        
        
    def action(self, execution=None):
        self.configure()
        
        self.log_callback('Capturando chaves de acesso no excel')
        chaves_acesso, numero_notas = self.extract_chave_acesso(self.filepath)
        if chaves_acesso.empty:
            print("Não há chaves de acesso válidas para processar.")
            return
        
        df_saida, quantidade_colunas = self.create_excel(self.saida_excel)
        
        self.browse("http://www2.sefaz.ce.gov.br/sitram-internet/masterDetailNotaFiscal.do?method=prepareSearch")
        self.set_screen_resolution(1600, 1800)
        self.maximize_window()
        
        for chave_acesso, numero_nota in zip(chaves_acesso, numero_notas):
            try:
                self.browse("http://www2.sefaz.ce.gov.br/sitram-internet/masterDetailNotaFiscal.do?method=prepareSearch")
                self.preencher_chave_acesso(chave_acesso)
                quantidade_registros = self.abrir_nota()
                
                if quantidade_registros == 'PAGA ou PARCELADA ou DEB. AUTUADO':
                    dados = []
                    dados.append([numero_nota] + ['0'] * (quantidade_colunas-2) + ['Paga ou Parcelada ou Deb. Autuado'])
                    df_saida = self.preencher_excel(self.saida_excel, df_saida, dados)
                    self.log_callback(f'Nota fiscal paga ou parcelada ou deb. autuada.')
                    continue
                
                contador = 1
                dados = []
                
                for i in range(int(quantidade_registros)):
                    try:
                        dados_linha, item_texto = self.get_calculos(i+1, contador)
                        dados_texto = self.get_valores(item_texto)

                        linha = [numero_nota] + dados_linha + dados_texto + ['Item Capturado']
                        
                        if len(linha) != quantidade_colunas:
                            raise Exception(f'Faltou dados no item')
                            
                        dados.append(linha)
                        contador += 2
                        
                    except Exception as e:
                        linha = [numero_nota] + dados_linha + dados_texto + [f'Erro ao Capturar Item: {e}']
                        dados.append(linha)
                        self.log_callback(f'Erro ao Capturar Item: {e}')
                        continue
                    
                self.log_callback('Nota capturada com sucesso!')
                
            except Exception as e:
                dados = []
                dados.append([numero_nota] + ['0'] * (quantidade_colunas-2) + [f'Erro ao Capturar Nota {e}'])
                print(f'Erro encontrado {e}')
                self.log_callback(f'Erro ao Capturar Nota {e}')
                continue
            
            finally:
                df_saida = self.preencher_excel(self.saida_excel, df_saida, dados)
                
        print('Formatando')
        self.save_excel_with_formatting(df_saida, self.saida_excel)
    
    
    def configure(self):
        self.driver_path = ChromeDriverManager().install()
        self.headless = True
    
    
    def extract_chave_acesso(self, filepath):
        print('Extraindo dados do excel')
        df = pd.read_csv(filepath)
        df_filtrado = df.dropna(subset=['Chave acesso', 'N° NF'])
        chaves_acesso = df_filtrado['Chave acesso']
        numeros_nf = df_filtrado['N° NF']
        print(chaves_acesso)
        self.log_callback(f'{len(chaves_acesso)} chaves de acesso encontrada(s)')
        
        return chaves_acesso, numeros_nf
    
    
    def create_excel(self, filepath):
        print('Criando Excel de saida')
        colunas = ['N° NF', 'Número', 'Código', 'NCM', 'Produto', 'CFOP', 'CST', 'Qtde', 'Vlr Unit', 'Total', 'BC ICMS', 'IPI',
                   'ICMS Destacado', 'Alíq ICMS', 'ICMS Calc. FECOP - Receita 2020',
                   'ICMS Calc. - Receitas 1023/1031/1090/1120', 'Frete Rateado', 'Crédito de Origem', 'Situação']
        df = pd.DataFrame(columns=colunas)
        df.to_excel(filepath, index=False)
        
        return df, len(colunas)
    
    
    def preencher_chave_acesso(self, chave_acesso):
        self.log_callback(f'Preenchendo chave de acesso: "{chave_acesso}"')
        print(f'Preenchendo chave de acesso: "{chave_acesso}"')
        campo_chave_acesso = self.find_element('//*[@id="chaveAcesso"]', By.XPATH, waiting_time=60000)
        campo_chave_acesso.send_keys(chave_acesso)
        
        pesquisar = self.find_element('//*[@id="pesquisar"]', By.XPATH, waiting_time=60000)
        pesquisar.click()
        
        
    def abrir_nota(self):
        linhas = self.find_elements('#NotaFiscalForm tbody tr', By.CSS_SELECTOR, waiting_time=60000)
        numero_linhas = len(linhas)
        print(numero_linhas)
        
        if numero_linhas == 1:
            nota = self.find_element('//*[@id="NotaFiscalForm"]/div[3]/table/tbody/tr/td[7]', By.XPATH, waiting_time=60000)
            notatexto = nota.get_attribute('textContent')
            if 'A PAGAR' in notatexto or 'SEM COBRANÇA' in notatexto:
                nota.click()
                
            elif 'PAGA' in notatexto or 'PARCELADA' in notatexto or 'AUTUADO' in notatexto:
                return 'PAGA ou PARCELADA ou DEB. AUTUADO'
            
        else:
            for i in range(numero_linhas):
                i += 1
                nota = self.find_element(f'//*[@id="NotaFiscalForm"]/div[3]/table/tbody/tr[{i}]/td[7]', By.XPATH, waiting_time=60000)
                notatexto = nota.get_attribute('textContent')
                print(notatexto)
                if 'A PAGAR' in notatexto or 'SEM COBRANÇA' in notatexto:
                    nota.click()
                    break
                
                elif 'PAGA' in notatexto or 'PARCELADA' in notatexto or 'AUTUADO' in notatexto:
                    return 'PAGA ou PARCELADA ou DEB. AUTUADO'
        
        registros = self.find_element('//*[@id="tableItens"]/thead/tr[2]/td/nav/ul/li[3]/span/span[3]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(registros, visible=True, waiting_time=20000)
        self.wait(500)
        registros = self.find_element('//*[@id="tableItens"]/thead/tr[2]/td/nav/ul/li[3]/span/span[3]', By.XPATH, waiting_time=60000)
        
        quantidade_registros = registros.get_attribute('textContent')
        print(quantidade_registros)
        self.log_callback(f'Nota fiscal aberta, {quantidade_registros} itens encontrados.')
        return quantidade_registros
    
    
    def get_calculos(self, item, contador):
        self.log_callback(f'Pegando dados do item {item}')
        print(f'Pegando dados do item {item}')
        resto = item%100
        
        if resto == 1 and item != 1:
            self.proxima_pagina()
            
        dados = []
        for i in range(15):
            i += 1
            linha = self.find_element(f'//*[@id="{item}"]/td[{i}]/span', By.XPATH, waiting_time=10000)
            dados.append(linha.get_attribute('textContent'))
        
        if resto == 0:
            resto = item
        calculadora = self.find_element(f'/html/body/div[1]/div[2]/div/section/form[2]/span/div/div[2]/table/tbody/tr[{resto}]/td[16]', By.XPATH, waiting_time=60000)
        calculadora.click()
        
        item_texto = self.find_element(f'//*[@id="ui-id-{contador}"]/pre', By.XPATH, waiting_time=60000)
        item_texto = item_texto.get_attribute('textContent')
        
        fechar = self.find_element('/html/body/div[7]/div[2]/div/button', By.XPATH, waiting_time=60000)
        fechar.click()
        
        return dados, item_texto
        
        
    def proxima_pagina(self):
        prox_pag = self.find_element(f'//*[@id="tableItens"]/tfoot/tr/td/nav/ul/li[4]/a', By.XPATH, waiting_time=60000)
        prox_pag.click()
        tabela = self.find_element(f'/html/body/div[1]/div[2]/div/section/form[2]/span/div/div[2]/table/tbody/tr[1]/td[1]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(tabela, visible=True, waiting_time=20000)
        self.wait(500)
        
        
    def get_valores(self, item_texto):
        palavras = item_texto.split()
        frete_rateado = 0
        credito_origem = 0
        
        for i, palavra in enumerate(palavras):
            
            if palavra == '(rateado):':
                frete_rateado = palavras[i + 2]
                frete_rateado = frete_rateado.replace('.', '').replace(',', '.')
                frete_rateado = round(float(frete_rateado), 2)
                
            elif palavra == 'Crédito':
                if palavras[i+1] == 'de':
                    if palavras[i+2] == 'origem:':
                        credito_origem = palavras[i + 4]
                        credito_origem = credito_origem.replace('%', '').replace('.', '').replace(',', '.')
                        credito_origem = round(float(credito_origem), 2)
        
        return [frete_rateado, credito_origem]
    
    
    def preencher_excel(self, filepath, df_saida, dados):
        for linha in dados:
            linha_df = pd.DataFrame([linha], columns=df_saida.columns)

            df_saida = pd.concat([df_saida, linha_df], ignore_index=True)

        df_saida.to_excel(filepath, index=False)

        return df_saida
    
    
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
            # Aqui você pode ajustar a largura como preferir
            worksheet.set_column(i, i, 10)

        # Fechar o writer e salvar o arquivo Excel
        writer.close()
            
            
    def not_found(self, label):
        print(f"Element not found: {label}")
        
        
def selecionar_diretorio():
    diretorio = filedialog.askopenfilename()
    caminho_entry.delete(0, tk.END)
    caminho_entry.insert(0, diretorio)


def iniciar_processamento():
        diretorio = caminho_entry.get()

        if diretorio:
            log_text.insert(tk.END, "Iniciando processamento...\n", 'info')
            app.update()

            # Crie uma thread para executar o processamento
            threading.Thread(target=processamento_thread, args=(diretorio,)).start()


def processamento_thread(diretorio):
    try:
        bot = MyWebBot(filepath=diretorio, log_callback=lambda mensagem: app.after(0, atualizar_interface, mensagem))
        bot.action()
        # Agende uma chamada à thread principal para atualizar a interface
        app.after(0, atualizar_interface, "Processamento concluído.")
        messagebox.showinfo("Script Finalizado", "O script foi concluído com sucesso!")
    except Exception as e:
        # Agende uma chamada à thread principal para lidar com erros
        app.after(0, atualizar_interface, f"Erro durante o processamento: {str(e)}")
        messagebox.showinfo("Erro!", "Falha ao rodar o script. Verifique seus arquivos de entrada.")


def atualizar_interface(mensagem):
    log_text.insert(tk.END, f"{mensagem}\n", 'info')
    log_text.see(tk.END)  # Auto-scroll
    app.update()


if __name__ == '__main__':
    app = tk.Tk()
    app.title("Selecionador de Diretório")

    frame = tk.Frame(app)
    frame.pack(padx=10, pady=10)

    caminho_entry = tk.Entry(frame, width=100)
    caminho_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    botao_buscar = tk.Button(frame, text="Buscar", command=selecionar_diretorio)
    botao_buscar.pack(side=tk.LEFT, padx=(10, 0))

    botao_enviar = tk.Button(frame, text="Enviar", command=iniciar_processamento)
    botao_enviar.pack(side=tk.LEFT, padx=(10, 0))

    log_text = scrolledtext.ScrolledText(app, height=10)
    log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    log_text.tag_configure('info', foreground='blue')
    log_text.tag_configure('error', foreground='red')
    app.mainloop()