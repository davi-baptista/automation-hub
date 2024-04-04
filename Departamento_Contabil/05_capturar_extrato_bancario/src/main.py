from botcity.web import WebBot, By
from botcity.web.browsers.chrome import default_options
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta
import shutil


class NiboDownloader(WebBot):
    
    
    def __init__(self, filepath, year_marked, months_marked, log_callback, *args, **kwargs):
        super(NiboDownloader, self).__init__(*args, **kwargs)
        self.filepath = filepath
        self.year_marked = year_marked
        self.months_marked = months_marked
        self.log_callback = log_callback
        
    
    def action(self, execution=None):
        self.log_callback('Iniciando robo, aguarde um momento...')
        print('Iniciando robo, aguarde um momento...')
        
        email = ' '
        password = ' '
        
        results_path = os.path.join(os.path.dirname(self.filepath), 'downloads')
        results_path = results_path.replace("/", "\\")
        
        os.makedirs(results_path, exist_ok=True)
        
        df = pd.read_excel(self.filepath)
        
        df = self.create_columns_status(df, self.year_marked, self.months_marked)
        self.configure(results_path)
        self.login(email, password)
        
        for index, row in df.iterrows():
            self.browse("https://contador.nibo.com.br/Management/Index")
            self.log_callback(f' -> Abrindo empresa {row['Nome_empresa']}.')
            print(f'Abrindo empresa {row['Nome_empresa']}.')
            
            company = row['Nome_empresa']   
            self.navigate_to_documents(company)
            
            for month in self.months_marked:
                try:
                    status_month = f'Status_{month}/{self.year_marked}'
                    status_atual = str(row[status_month]).strip()
                    
                    if status_atual != 'nan' and status_atual:
                        print('a')
                        continue
                    
                    index_month = self.calcular_indice_mes(self.year_marked, month)
                    self.navigate_to_month(index_month)
                    status, downloads_number = self.get_total_titles(results_path)
                    
                    if status != 'Extratos Bancarios não encontrados':
                        files_number = self.move_files(results_path, company, self.year_marked, month)

                    if downloads_number != files_number:
                        status = 'VERIFICAR, ERRO NO DOWNLOAD DE ARQUIVOS'
                        
                    df.at[index, status_month] = status
                    df.to_excel(self.filepath, index=False)
                    
                    voltar = self.find_element('//*[@id="detail6"]/span', By.XPATH, waiting_time=60000)
                    voltar.click()
                    
                    self.log_callback(f' -> {status}.\n -> Mes {month} concluído e adicionada ao excel.\n')
                    print(f'{status}.\nMes {month} e adicionada ao excel.')
                    self.wait(2000)

                    
                except Exception as e:
                    df.at[index, status_month] = f'Erro encontrado {e}'
                    df.to_excel(self.filepath, index=False)
                    
                    print(f'Erro encontrado {e}')
                    self.log_callback(f' -> Erro encontrado {e}\n')
            
            self.log_callback(' -> Empresa finalizada.\n')
            print('Empresa finalizada.')
        
        self.save_excel_with_formatting(df, self.filepath)
        self.stop_browser()
        
                
    def create_columns_status(self, df, year_marked, months_marked):
        for month in months_marked:
            status_month = f'Status_{month}/{year_marked}'
            if status_month not in df.columns:
                df[status_month] = ''
        return df
        
        
    def configure(self, filepath):
        # Configurando o ChromeDriver
        self.driver_path = ChromeDriverManager().install()
        self.options = default_options(headless=True, download_folder_path=filepath, user_data_dir=None, page_load_strategy='normal')
    
    
    def login(self, email, password):
        self.log_callback('Abrindo Nibo e indo até a aba de documentos recebidos.')
        print('Abrindo Nibo e indo até a aba de documentos recebidos.')
        
        self.browse("https://contador.nibo.com.br/Management/Index")
        self.maximize_window()
        campo_email = self.find_element('emailAddress', By.ID, waiting_time=10000)
        campo_email.send_keys(email)
        campo_senha = self.find_element('password', By.ID, waiting_time=10000)
        campo_senha.send_keys(password)
        self.wait(200)
        self.enter()
        
        
    def navigate_to_documents(self, company):
        buscar_por = self.find_element('//*[@id="config2"]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(buscar_por, visible=True, waiting_time=60000)
        buscar_por.click()
        
        self.control_a()
        buscar_por.send_keys(company)
        self.wait(1000)
        
        filtrar = self.find_element('//*[@id="documentmovimentmanagement"]/div[2]/div/div[2]/div[4]/button', By.XPATH, waiting_time=60000)
        filtrar.click()
        
    
    def calcular_indice_mes(self, year_marked, month_marked):
        curr_date = datetime.now()
        
        months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
        
        month_index = months.index(month_marked)
        
        # Corrige o índice pois list.index() é baseado em 0, mas os meses são baseados em 1
        month_numeric = month_index + 1
        
        # Calcula a diferença em meses entre a data atual e a data escolhida
        diferenca_meses = (curr_date.year - int(year_marked)) * 12 + curr_date.month - month_numeric
        
        # Inicializa o índice para o intervalo de 6 meses (0 a 5)
        indice = -1
        
        # Verifica se a data escolhida está dentro do intervalo de 6 meses
        while diferenca_meses >= 0:
            if diferenca_meses < 6:
                indice = 5 - diferenca_meses
                break
            else:
                # Simula o avanço para a próxima página (ou seja, recua 6 meses a partir da data atual)
                departamento_contabil = self.find_element(f'//*[@id="documentmovimentmanagement"]/div[2]/div/div[3]/div[2]/div/div[3]/div[1]/i', By.XPATH, waiting_time=60000)
                departamento_contabil.click()
                
                month = self.find_element('//*[@id="config3-3"]', By.XPATH, waiting_time=60000)
                self.wait_for_element_visibility(month, visible=True, waiting_time=60000)
                
                curr_date -= relativedelta(months=6)
                diferenca_meses = (curr_date.year - int(year_marked)) * 12 + curr_date.month - month_numeric
                
        return indice
                
    
    def navigate_to_month(self, indice):
        month = self.find_element('//*[@id="config3-3"]', By.XPATH, waiting_time=60000)
        self.wait_for_element_visibility(month, visible=True, waiting_time=60000)
        month.click()

        departamento_contabil = self.find_element(f'//*[@id="config6-{indice}"]/span[1]', By.XPATH, waiting_time=60000)
        departamento_contabil.click()
    
    
    def get_total_titles(self, results_path):
        titulos = self.find_elements('#detail1 > h4', By.CSS_SELECTOR, waiting_time=60000)
        if not titulos:
            print('Extratos Bancarios não encontrados')
            return 'Extratos Bancarios não encontrados', 0
        
        index = None
        for i, titulo in enumerate(titulos):
            if 'Extratos bancários' in titulo.get_attribute('textContent'):
                index = i + 1
        
        table = self.find_element(f'/html/body/div[4]/div/div[3]/div[2]/div/div[2]/div/div/div[3]/div[1]/div[2]/div/div[{index}]/div/table', By.XPATH, waiting_time=60000)
        
        buttons = table.find_elements(By.CLASS_NAME, 'bank-statement-files')
        if buttons:
            for button in buttons:
                button.click()
        
        downloads = table.find_elements(By.CLASS_NAME, 'fa-cloud-download')    
        if downloads:
            for download in downloads:
                download.click()
            
        if len(downloads) == 0:
            return 'Nenhum arquivo encontrado para download', 0
        
        else:
            timer = len(downloads)*500
            self.wait(timer)
            self.wait_files(results_path)
            
            while True:
                tabs = self.get_tabs()
                len_tabs = len(tabs)
                
                if len_tabs == 1:
                    break
                else:
                    self.activate_tab(tabs[-1])
                    self.get_screenshot(f'{results_path}/screenshot{len(tabs)}.png')
                    self.close_page()
                    self.wait(1000)
                    
                self.wait_files(results_path)
                
        return f'Total de {len(downloads)} arquivos baixados', len(downloads)
    
    
    def nome_pasta_seguro(self, nome):
        # Define uma lista de caracteres que não são permitidos em nomes de arquivos e diretórios
        caracteres_proibidos = "/<>:\"|?*\\"
        
        # Substitui cada caractere proibido por um sublinhado (ou outro caractere de sua escolha)
        for char in caracteres_proibidos:
            nome = nome.replace(char, "_")
        
        return nome


    def wait_files(self, results_path):
        # Loop para esperar até que não haja arquivos .crdownload
        while True:
            # Lista todos os arquivos em results_path
            arquivos = os.listdir(results_path)
            
            # Verifica se há arquivos .crdownload
            crdownload_existente = any(arq.endswith('.crdownload') for arq in arquivos)
            
            if not crdownload_existente:
                break  # Sai do loop se não houver arquivos .crdownload
            
            self.wait(1000)
        
        
    def move_files(self, results_path, company, year, month):
        target_dir = os.path.join(results_path, self.nome_pasta_seguro(company), year, month)
        os.makedirs(target_dir, exist_ok=True)
        
        cont = 0
        # Lista todos os arquivos em results_path
        arquivos = os.listdir(results_path)
        for arquivo in arquivos:
            path_completo = os.path.join(results_path, arquivo)
            destino_completo = os.path.join(target_dir, arquivo)
            
            if os.path.isfile(path_completo):  # Verifica se é um arquivo
                if os.path.exists(destino_completo):
                    # Remove o arquivo de destino para permitir a sobrescrita
                    os.remove(destino_completo)
                cont += 1
                shutil.move(path_completo, target_dir)

        print("Todos os arquivos foram movidos.")
        return cont
    
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


def select_directory():
    directory = filedialog.askopenfilename()
    path_entry.delete(0, tk.END)
    path_entry.insert(0, directory)


def iniciar_processamento(year_marked, months_marked):
    directory = path_entry.get()
    
    if not directory:
        messagebox.showinfo("Erro", "Selecione um arquivo excel antes de enviar.")
        return
    
    if not months_marked:
        messagebox.showinfo("Erro", "Selecione os meses antes de enviar.")
        return
    
    ano_atual = datetime.now().year
    mes_atual = datetime.now().month
    
    months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
            "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    
    if year_marked == str(ano_atual) and any(months.index(mes) >= mes_atual for mes in months_marked):
        messagebox.showinfo("Erro", "Não é possível selecionar um mês futuro no ano atual.")
        return

    if directory and len(months_marked) != 0:
        try:
            bot = NiboDownloader(filepath=directory, year_marked=year_marked, months_marked=months_marked, log_callback=lambda mensagem: atualizar_interface(mensagem))
            bot.action()
            
            log_text.insert(tk.END, "Script finalizado com sucesso.\n\n", 'info')
            app.update()
            messagebox.showinfo("Script Finalizado", "O script foi concluído com sucesso!")
        except Exception as e:
            log_text.insert(tk.END, f"Erro no script {e}.\n Verifique o excel inserido.\n\n", 'info')
            app.update()
            messagebox.showinfo("Script Finalizado", "O script encontrou um erro!")
    

def atualizar_interface(mensagem):
    log_text.insert(tk.END, f"{mensagem}\n", 'info')
    log_text.see(tk.END)  # Auto-scroll
    app.update()
   
   
def create_checkboxes(app):
    months = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho", "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
    # Novo frame para as checkboxes dos meses
    frame = tk.Frame(app)
    frame.pack(padx=10, pady=10)

    checkboxes_vars = {}
    coluna = 0  # Inicializa a contagem de colunas para organizar os checkboxes
    for month in months:
        checkbox_var = tk.IntVar()
        checkbox = tk.Checkbutton(frame, text=month, variable=checkbox_var)
        checkbox.grid(row=0, column=coluna, sticky="w")
        coluna += 1  # Incrementa a coluna para o próximo checkbox ficar ao lado

        checkboxes_vars[month] = checkbox_var
        
    return checkboxes_vars


def verificar_checkboxes_marcadas(checkboxes_vars, year_selected):
    months = [mes for mes, var in checkboxes_vars.items() if var.get() == 1]
    year = year_selected.get()
    print(f'{months} - {year}')
    
    iniciar_processamento(year, months)


def criar_combobox_anos(app):
    # Novo frame para as checkboxes dos meses
    frame = tk.Frame(app)
    frame.pack(padx=10, pady=10)
    
    # Gerar a lista de anos de 2021 até o ano atual
    ano_atual = datetime.now().year
    anos = [str(ano) for ano in range(2021, ano_atual + 1)]
    
    # Criar a variável que vai armazenar o valor selecionado na combobox
    ano_selecionado = tk.StringVar()
    
    # Criar a combobox e configurá-la com a lista de anos
    combobox_anos = ttk.Combobox(frame, textvariable=ano_selecionado, values=anos, state="readonly")
    combobox_anos.set(str(ano_atual))  # Definir o valor padrão como o ano atual
    combobox_anos.pack()  # Posicionar a combobox no frame
    
    return ano_selecionado


if __name__ == '__main__':
    app = tk.Tk()
    app.title("NiboExctact Downloader")

    frame_principal = tk.Frame(app)
    frame_principal.pack(padx=10, pady=10)

    path_entry = tk.Entry(frame_principal, width=100)
    path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

    botao_buscar = tk.Button(frame_principal, text="Buscar", command=select_directory)
    botao_buscar.pack(side=tk.LEFT, padx=(10, 0))

    botao_enviar = tk.Button(frame_principal, text="Iniciar", command=lambda: verificar_checkboxes_marcadas(checkboxes_vars, year_selected))
    botao_enviar.pack(side=tk.LEFT, padx=(10, 0))

    checkboxes_vars = create_checkboxes(app)
    
    year_selected = criar_combobox_anos(app)

    log_text = scrolledtext.ScrolledText(app, height=10)
    log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    log_text.tag_configure('info', foreground='blue')
    log_text.tag_configure('error', foreground='red')

    app.mainloop()