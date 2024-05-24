import sys
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLineEdit, QTextEdit, QFileDialog, QMessageBox
from PyQt5.QtGui import QPalette, QColor
from PyQt5.QtCore import Qt
import os

from main import main


class MainWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Capturador de Funcionarios")
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

        self.filepath_input = QLineEdit()
        self.filepath_input.setPlaceholderText("Coloque o caminho do PDF aqui")
        input_layout.addWidget(self.filepath_input)

        self.select_file_button = QPushButton("Selecionar Arquivo")
        self.select_file_button.clicked.connect(self.select_file)
        input_layout.addWidget(self.select_file_button)

        self.send_button = QPushButton("Enviar")
        self.send_button.clicked.connect(self.send_info)
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
        file_path, _ = file_dialog.getOpenFileName(self, "Selecionar Arquivo", "", "PDF Files (*.pdf)")
        if file_path:
            self.filepath_input.setText(file_path)
    
    
    def clean_cnpj(self, cnpj):
        """Remove common punctuation from CNPJ: ".", "/", and "-"."""
        return cnpj.replace(".", "").replace("/", "").replace("-", "")
            

    def check_fields(self, filepath):
        if not os.path.exists(filepath):
            status_path = 'Caminho invalido! Procure a pasta novamente.'
        
        elif os.path.exists(filepath) and not filepath.endswith('.pdf') and not filepath.endswith('.PDF'):
            status_path = 'Arquivo invalido! Procure a pasta novamente.'
        
        else:
            status_path = 'Caminho valido'

        return status_path

                
    def send_info(self):
        filepath = self.filepath_input.text()
        
        path_error = self.check_fields(filepath)
        if path_error != 'Caminho valido':
            QMessageBox.warning(self, "Field Verification", 'Envie um caminho de um arquivo PDF valido para iniciar.')
            return

        else:
            try:
                self.log('Caminho valido. Iniciando processamento dos dados...')
                main_return = main(filepath)
                self.log(f'{main_return}')
                if 'Erro' not in main_return:
                    self.log('Processamento finalizado.')
                    self.log('Excel criado na mesma pasta de envio!')
                    QMessageBox.information(self, "Sucesso", 'Processamento completo.')
                    
                else:
                    QMessageBox.critical(self, "Erro", 'Erro inesperado no processamento encontrado.')

            except Exception as e:
                self.log(f'Erro inesperado.{e}')
                self.log(e)
                QMessageBox.critical(self, "Erro", 'Erro inesperado encontrado.')
                
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