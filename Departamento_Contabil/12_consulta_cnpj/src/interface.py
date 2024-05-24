import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QFileDialog, QMessageBox
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt
import os
import pandas as pd

from main import capture_cnpj_data


class MainWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Consulta CNPJ")
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

        self.company_cnpj_input = QLineEdit()
        self.company_cnpj_input.setPlaceholderText("Digite o cnpj da empresa aqui, apenas números")
        input_layout.addWidget(self.company_cnpj_input)

        self.select_file_button = QPushButton("Selecionar Arquivo")
        self.select_file_button.clicked.connect(self.select_file)
        input_layout.addWidget(self.select_file_button)

        self.send_button = QPushButton("Enviar")
        self.send_button.clicked.connect(self.send_company_info)
        input_layout.addWidget(self.send_button)

        self.toggle_theme_button = QPushButton("Alterar Tema")
        self.toggle_theme_button.clicked.connect(self.toggle_theme)
        layout.addWidget(self.toggle_theme_button)

        # Log area
        self.log_area = QTextEdit(readOnly=True)
        layout.addWidget(self.log_area)

    
    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        self.apply_theme()


    def apply_theme(self):
        if self.dark_mode:
            apply_dark_mode(app)
        else:
            apply_light_mode(app)
                

    def select_file(self):
        file_dialog = QFileDialog(self)
        file_path, _ = file_dialog.getOpenFileName(self, "Selecionar Arquivo", "", "XLSX Files (*.xlsx)")
        if file_path:
            self.company_cnpj_input.setText(file_path)
    
    
    def clean_cnpj(self, cnpj):
        """Remove common punctuation from CNPJ: ".", "/", and "-"."""
        return cnpj.replace(".", "").replace("/", "").replace("-", "")
            

    def check_fields(self, company_cnpj):
        status_cnpj = None
        if len(self.clean_cnpj(company_cnpj)) != 14 or not self.clean_cnpj(company_cnpj).isdigit():
            status_cnpj = "CNPJ deve conter exatamente 14 dígitos numéricos."
        
        status_path = None
        if not os.path.exists(company_cnpj):
            status_path = 'Caminho invalido! Procure a pasta novamente.'
        
        elif os.path.exists(company_cnpj) and not company_cnpj.endswith('.xlsx'):
            status_path = 'Arquivo invalido! Procure a pasta novamente.'
        
        return status_cnpj, status_path

                
    def send_company_info(self):
        company_cnpj = self.company_cnpj_input.text()
        
        cnpj_error, path_error = self.check_fields(company_cnpj)
        if cnpj_error and path_error:
            QMessageBox.warning(self, "Field Verification", 'Envie um caminho de um arquivo ou um CNPJ valido para iniciar')
            return

        elif not cnpj_error:
            try:
                self.log(f"> CNPJ: {company_cnpj}")
                
                count = [0]
                cnpj_data = capture_cnpj_data(company_cnpj, count)
                if cnpj_data:
                    razao_social = cnpj_data.get('razao_social', 'Não disponível')
                    data_inicio = cnpj_data.get('data_inicio_atividade', 'Não disponível')
                    situacao_cadastral = cnpj_data.get('descricao_situacao_cadastral', 'Não disponível')
                    socios = cnpj_data.get('qsa', [])
                    nomes_socios = [socio['nome_socio'] for socio in socios] if socios else ['Não há sócios listados']
                    self.log(f"> Razão Social: {razao_social}")
                    self.log(f"> Data de Início de Atividade: {data_inicio}")
                    self.log(f"> Situação Cadastral: {situacao_cadastral}")
                    self.log("> Sócios:")
                    for nome in nomes_socios:
                        self.log(f"  - {nome}")
                else:
                    self.log('Nada encontrado no CNPJ informado. \nConfira se foi inserido corretamente')
            
            except Exception as e:
                self.log(f'Erro inesperado. {e}')
                
        elif not path_error:
            try:
                self.log(f"> Caminho excel capturado")
                filepath = company_cnpj
                df = pd.read_excel(filepath, dtype=str)
                
                for column in ['Razão Social', 'Data de Início de Atividade', 'Situação Cadastral', 'Sócios']:
                    if column not in df.columns:
                        df[column] = None
                df.to_excel(filepath, index=False)
                df = pd.read_excel(filepath, dtype=str)
                        
                count = [0]
                total_rows = df.shape[0]
                cnpj_data_count, not_cnpj_data_count = 0, 0
                for index, row in df.iterrows():
                    if pd.notna(row['Razão Social']) and pd.notna(row['Data de Início de Atividade']) and pd.notna(row['Situação Cadastral']) and pd.notna(row['Sócios']):
                        continue
                    
                    self.log(f'Carregando... {index+1}/{total_rows}')
                    cnpj = row['CNPJ']
                    cnpj_data = capture_cnpj_data(cnpj, count)
                    if cnpj_data:
                        cnpj_data_count += 1
                        row['Razão Social'] = cnpj_data.get('razao_social', 'Não disponível')
                        row['Data de Início de Atividade'] = cnpj_data.get('data_inicio_atividade', 'Não disponível')
                        row['Situação Cadastral'] = cnpj_data.get('descricao_situacao_cadastral', 'Não disponível')
                        socios = cnpj_data.get('qsa', [])
                        row['Sócios'] = ', '.join([socio['nome_socio'] for socio in socios]) if socios else 'Não há sócios listados'
                        df.to_excel(filepath, index=False)
                    else:
                        not_cnpj_data_count += 1
                        row['Razão Social'] = 'NADA ENCONTRADO'
                        row['Data de Início de Atividade'] = 'NADA ENCONTRADO'
                        row['Situação Cadastral'] = 'NADA ENCONTRADO'
                        row['Sócios'] = 'NADA ENCONTRADO'
                        df.to_excel(filepath, index=False)
                
                self.log('Processamento finalizado.')
                self.log(f"{cnpj_data_count} CNPJ's compilados com sucesso!")
                if not_cnpj_data_count != 0:
                    self.log(f"{not_cnpj_data_count} CNPJ's sem nenhuma informação encontrada.")
                    
            except Exception as e:
                self.log(f'Erro inesperado. {e}')
                
        self.log("\n")
        

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