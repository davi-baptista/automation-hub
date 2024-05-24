import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QFileDialog, QDateEdit, QMessageBox
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt, QDate

from main import main


class MainWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Validador de Estoque")
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

        # Layout for start and end date inputs
        date_layout = QHBoxLayout()
        layout.addLayout(date_layout)

        self.start_date_input = QDateEdit(calendarPopup=True)
        self.start_date_input.setDate(QDate.currentDate())
        date_layout.addWidget(self.start_date_input)

        self.end_date_input = QDateEdit(calendarPopup=True)
        self.end_date_input.setDate(QDate.currentDate())
        date_layout.addWidget(self.end_date_input)

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
        start_date, end_date = self.start_date_input.date(), self.end_date_input.date()
        
        validation_msg = self.check_fields(company_cnpj, start_date, end_date)
        if validation_msg:
            QMessageBox.warning(self, "Field Verification", validation_msg)
            return

        self.log(f"Sending: {company_cnpj}, from {start_date.toString('dd/MM/yyyy')} to {end_date.toString('dd/MM/yyyy')}")
        start_date_str, end_date_str = start_date.toString("MM-dd-yyyy"), end_date.toString("MM-dd-yyyy")
        
        excel_path, main_status = main(self.clean_cnpj(company_cnpj), start_date_str, end_date_str)
        self.review_main_status(excel_path, main_status)


    def check_fields(self, company_cnpj, start_date, end_date):
        if len(self.clean_cnpj(company_cnpj)) != 14 or not self.clean_cnpj(company_cnpj).isdigit():
            return "CNPJ deve conter exatamente 14 dígitos numéricos."
        if start_date > end_date:
            return "A data de início deve ser anterior à data de término."
        if start_date > QDate.currentDate() or end_date > QDate.currentDate():
            return "As datas não podem ser futuras em relação à data atual."
        return None
        
        
    def review_main_status(self, excel_path, status):
        if status == 'Excel file saved successfully':
            log_msg = f"Robô finalizado com sucesso. Dados salvos na pasta.\n{excel_path}"
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