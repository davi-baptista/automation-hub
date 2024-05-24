import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QSpinBox, QMessageBox
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt, QDate
import os

from main import main


class MainWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Demonstração")
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

        year_layout = QHBoxLayout()
        layout.addLayout(year_layout)

        self.year_input = QSpinBox()
        self.year_input.setRange(2000, QDate.currentDate().year())
        self.year_input.setValue(QDate.currentDate().year())
        self.year_input.setFixedWidth(100)
        year_layout.addWidget(self.year_input)

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
            
            
    def clean_cnpj(self, cnpj):
        """Remove common punctuation from CNPJ: ".", "/", and "-"."""
        return cnpj.replace(".", "").replace("/", "").replace("-", "")
                
                
    def send_company_info(self):
        company_cnpj = self.company_cnpj_input.text()
        selected_year = self.year_input.value()
        
        validation_msg = self.check_fields(company_cnpj)
        if validation_msg:
            QMessageBox.warning(self, "Field Verification", validation_msg)
            return

        self.log(f"Enviando: {company_cnpj}, com o ano {selected_year}")
        
        excel_path, main_status = main(self.clean_cnpj(company_cnpj), selected_year)
        self.review_main_status(excel_path, main_status)
            

    def check_fields(self, company_cnpj):
        if len(self.clean_cnpj(company_cnpj)) != 14 or not self.clean_cnpj(company_cnpj).isdigit():
            return "CNPJ deve conter exatamente 14 dígitos numéricos."
        return None
        
        
    def review_main_status(self, excel_path, status):
        if status == 'Excel file saved successfully':
            log_msg = f"Robô finalizado com sucesso. Dados salvos na pasta.\n{os.path.dirname(excel_path)}"
            self.log(log_msg + '\n')
            QMessageBox.information(self, "Sucesso", log_msg)
        elif status == 'File not found':
            log_msg = "Esta empresa ainda não está padronizada no sistema. Por favor, chame o suporte."
            self.log(log_msg + '\n')
            QMessageBox.warning(self, "Aviso", log_msg)
        elif status == 'Invalid CNPJ':
            log_msg = "Verifique se o CNPJ informado esta correto."
            self.log(log_msg + '\n')
            QMessageBox.warning(self, "Aviso", log_msg)
        elif status == 'Empty company':
            log_msg = "Nada encontrado no CPNJ colocado. Verifique as datas inseridas."
            self.log(log_msg + '\n')
            QMessageBox.warning(self, "Aviso", log_msg)
        else:
            log_msg = f"Ocorreu um erro inesperado.\n{status}.\nPor favor, tente novamente ou contate o suporte."
            self.log(log_msg + '\n')
            QMessageBox.critical(self, "Erro", log_msg)
            

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