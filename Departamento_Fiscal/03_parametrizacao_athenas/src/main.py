import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pyautogui
import keyboard


def preencher_athenas(ncm, cst_icms, cfop):
    pyautogui.hold('shift')
    pyautogui.write('OT')
    pyautogui.press('tab', presses=4)
    pyautogui.press('backspace')
    pyautogui.press('tab')
    pyautogui.write(ncm)
    pyautogui.press('tab', presses=3)
    pyautogui.write(str(cst_icms), interval=0)
    pyautogui.press('tab', presses=3)
    pyautogui.press('backspace')
    pyautogui.press('tab')
    pyautogui.write(cfop)   
    pyautogui.press('tab', presses=2)
    pyautogui.write(str(int(cfop)+1000))
    pyautogui.press('down')
    pyautogui.press('home')


def selecionar_diretorio():
    diretorio = filedialog.askopenfilename()
    caminho_entry.delete(0, tk.END)
    caminho_entry.insert(0, diretorio)


def iniciar_processamento():
    diretorio = caminho_entry.get()

    if diretorio:
        log_text.insert(tk.END, "Aguardando pressionamento da tecla f6 para iniciar.\n", 'info')
        app.update()
        
        df = pd.read_excel(diretorio, dtype={'NCM': str, 'CST': str, 'CFOP': str})
        if 'Status' not in df.columns:
            df['Status'] = ''
        
        while True:
            app.update()
            if keyboard.is_pressed('f6'):
                print('Tecla f6 pressionada. Iniciando processamento...\n')
                log_text.insert(tk.END, "Tecla f6 pressionada.\n", 'info')
                log_text.insert(tk.END, "Iniciando processamento...\n", 'info')
                break
                
        pyautogui.press('tab')
        pyautogui.hotkey('ctrl', 'end')
        pyautogui.press('down')
        pyautogui.press('home')
        
        for indice, row in df.iterrows():
            status_atual = str(row['Status']).strip()
            if status_atual != 'nan' and status_atual:
                print(f'Pulando pois o status já está definido como {status_atual}.')
                continue
            
            if keyboard.is_pressed('f6'):
                log_text.insert(tk.END, "Tecla 'f6' pressionada. Parando o loop...\n", 'info')
                app.update()
                break
            
            ncm = row['NCM']
            cst = row['CST'][-2:]
            cfop = row['CFOP']
            
            preencher_athenas(ncm, cst, cfop)  
            df.at[indice, 'Status'] = 'Concluído'
            
        df.to_excel(diretorio, index=False)
        pyautogui.press('up')
        log_text.insert(tk.END, "Script finalizado com sucesso.\n\n", 'info')
        messagebox.showinfo("Script Finalizado", "O script foi concluído com sucesso!")
    
    
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