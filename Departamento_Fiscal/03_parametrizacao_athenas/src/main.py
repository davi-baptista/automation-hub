import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QFileDialog, QMessageBox
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt
import pyautogui
import pandas as pd
import keyboard
import os
from pynput.keyboard import Controller


class MainWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Parametrizador Athenas")
        self.setGeometry(100, 100, 600, 400)

        self.dark_mode = True
        self.setup_ui()
        self.apply_theme()
        
    def setup_ui(self):
        layout = QVBoxLayout()
        self.setLayout(layout)

        # Input layout for the CNPJ input and buttons
        input_layout = QHBoxLayout()
        layout.addLayout(input_layout)

        self.file_path_input = QLineEdit()
        self.file_path_input.setPlaceholderText("Escolha o arquivo excel")
        self.file_path_input.setReadOnly(False)
        input_layout.addWidget(self.file_path_input)

        self.select_file_button = QPushButton("Selecionar Arquivo")
        self.select_file_button.clicked.connect(self.select_file)
        input_layout.addWidget(self.select_file_button)
        
        self.send_button = QPushButton("Enviar")
        self.send_button.clicked.connect(self.send_athenas_data)
        input_layout.addWidget(self.send_button)

        self.toggle_theme_button = QPushButton("Alterar Tema")
        self.toggle_theme_button.clicked.connect(self.toggle_theme)
        layout.addWidget(self.toggle_theme_button)
        
        # Log area
        self.log_area = QTextEdit(readOnly=True)
        layout.addWidget(self.log_area)

    def select_file(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(self, "Selecionar Arquivo", "", "XLSX Files (*.xlsx)")
        if file_path:
            self.file_path_input.setText(file_path)
            
            
    def preencher_athenas(self, cfop_origem, ncm, cst_icms, cfop):
        pyautogui.write('OT')
        pyautogui.press('tab', presses=3)
        pyautogui.write(str(cfop_origem))
        pyautogui.press('tab', presses=2)
        if str(ncm) != 'nan':
            pyautogui.write(str(ncm))
        pyautogui.press('tab', presses=3)
        pyautogui.write(str(cst_icms), interval=0)
        pyautogui.press('tab', presses=3)
        pyautogui.press('backspace')
        pyautogui.press('tab')
        pyautogui.write(cfop)
        pyautogui.press('tab', presses=2)
        if str(cfop_origem) == 'nan':
            pyautogui.write(str(int(cfop)+1000))
        else:
            pyautogui.write(str(cfop))
        pyautogui.press('down')
        pyautogui.press('home')
            
            
    def preencher_athenas_federal(self, ncm, cst, cfop, aliquota, cod_empresa):
        pyautogui.write('XX')
        pyautogui.press('tab')
        pyautogui.write(str(ncm), interval=0.1)
        pyautogui.press('tab')
        if str(cfop) != 'nan':
            pyautogui.write(str(cfop))
        else:
            pyautogui.write('XXXX')
        pyautogui.press('tab', presses=2)
        pyautogui.write(str(cst))
        pyautogui.press('tab')
        if str(aliquota) != 'nan':
            keyboard = Controller()
            keyboard.type(str(aliquota).replace('.', ','))
        pyautogui.press('tab')
        pyautogui.write(str(cod_empresa))
        pyautogui.press('down')
        pyautogui.press('home')
        
        
    def send_athenas_data(self):
        file_path = self.file_path_input.text()
        if not os.path.exists(file_path):
            self.log(f'Verifique se o caminho do arquivo esta correto.\n')
            QMessageBox.warning(self, "Erro", 'Caminho invalido! Procure a pasta novamente.')
            return
        else:
            try:
                self.log("Pressione F6 para iniciar o robô.")
                self.wait_for_f6()
                
            except Exception as e:
                log_msg = f"Erro inesperado no script {e}."
                self.log(log_msg)
                QMessageBox.warning(self, "Erro", log_msg)
                
    def wait_for_f6(self):
        while True:
            if keyboard.is_pressed('f6'):
                self.log("Tecla 'F6' pressionada. Iniciando processamento...")
                self.process_data()
                break
            
    def process_data(self):
        try:
            file_path = self.file_path_input.text()
            df = pd.read_excel(file_path, dtype=str)
            print(df.dtypes)
            if 'Status_CFOP' not in df.columns:
                df['Status_CFOP'] = ''
            if 'Status_PIS' not in df.columns:
                df['Status_PIS'] = ''
            if 'Status_COFINS' not in df.columns:
                df['Status_COFINS'] = ''
                
            self.indo_para_fim()
            
            for indice, row in df.iterrows():
                status_atual = str(row['Status_CFOP']).strip()
                if status_atual != 'nan' and status_atual:
                    self.log(f"Pulando pois o status já está definido como {status_atual}.")
                    continue
                
                if keyboard.is_pressed('f6'):
                    self.log("Tecla 'F6' pressionada. Parando o loop...")
                    break
                
                ncm = row['NCM']
                cst = row['CST'][-2:]
                cfop = row['CFOP']
                cfop_origem = row['CFOP_ORIGEM']
                
                self.preencher_athenas(cfop_origem, ncm, cst, cfop)  
                df.at[indice, 'Status_CFOP'] = 'Concluído'
                
                df.to_excel(file_path, index=False)
                    
            pyautogui.hotkey('shift', 'tab')
            pyautogui.press('right', presses=4)
            pyautogui.press('enter')
            self.indo_para_fim()
                
            for indice, row in df.iterrows():
                status_atual = str(row['Status_PIS']).strip()
                if status_atual != 'nan' and status_atual:
                    self.log(f"Pulando pois o status já está definido como {status_atual}.")
                    continue
                
                if keyboard.is_pressed('f6'):
                    self.log("Tecla 'F6' pressionada. Parando o loop...")
                    break
                
                ncm = row['NCM']
                cst = row['CST_FEDERAL'][-2:]
                cfop = row['CFOP_FEDERAL']
                aliquota = row['ALIQ_PIS']
                codigo_empresa = row['COD_EMPRESA']
                
                self.preencher_athenas_federal(ncm, cst, cfop, aliquota, codigo_empresa)  
                df.at[indice, 'Status_PIS'] = 'Concluído'
                
                df.to_excel(file_path, index=False)
                
            pyautogui.press('end')
            pyautogui.press('tab', presses=3)
            self.indo_para_fim()
            
            for indice, row in df.iterrows():
                status_atual = str(row['Status_COFINS']).strip()
                if status_atual != 'nan' and status_atual:
                    self.log(f"Pulando pois o status já está definido como {status_atual}.")
                    continue
                
                if keyboard.is_pressed('f6'):
                    self.log("Tecla 'F6' pressionada. Parando o loop...")
                    break
                
                ncm = row['NCM']
                cst = row['CST_FEDERAL'][-2:]
                cfop = row['CFOP_FEDERAL']
                aliquota = row['ALIQ_COFINS']
                codigo_empresa = row['COD_EMPRESA']
                print(aliquota)
                
                self.preencher_athenas_federal(ncm, cst, cfop, aliquota, codigo_empresa)
                df.at[indice, 'Status_COFINS'] = 'Concluído'
                
                df.to_excel(file_path, index=False)
            
            log_msg = "Script finalizado com sucesso."
            self.log(log_msg)
            QMessageBox.information(self, "Sucesso", log_msg)
            
        except Exception as e:
            log_msg = f"Erro inesperado encontrado. '{e}'"
            self.log(log_msg)
            QMessageBox.warning(self, "Erro", log_msg)
            
            
    def indo_para_fim(self):
        pyautogui.press('tab')
        pyautogui.hotkey('ctrl', 'end')
        pyautogui.press('down')
        pyautogui.press('home')
    
    
    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()


    def apply_theme(self):
        if self.dark_mode:
            apply_dark_mode(app)
        else:
            apply_light_mode(app)
            

    def log(self, message):
        self.log_area.append(message)
        QApplication.processEvents()


def apply_dark_mode(app):
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.WindowText, Qt.white)
    palette.setColor(QPalette.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ToolTipBase, Qt.white)
    palette.setColor(QPalette.ToolTipText, Qt.white)
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.HighlightedText, Qt.black)
    app.setPalette(palette)


def apply_light_mode(app):
    app.setPalette(app.style().standardPalette())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())