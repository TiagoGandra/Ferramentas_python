import sys
import pandas as pd
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt

class CSVToExcelApp(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        # Configuração da janela
        self.setWindowTitle('Conversor CSV para Excel')
        self.setGeometry(100, 100, 450, 350)
        self.setStyleSheet("""
            QWidget {
                background-color: #2C3E50;
                font-family: Arial;
            }
            QLabel {
                color: #ECF0F1;
                font-size: 14px;
                padding: 10px;
            }
            QPushButton {
                background-color: #1ABC9C;
                color: white;
                font-size: 16px;
                padding: 10px;
                border: none;
                border-radius: 10px;
            }
            QPushButton:hover {
                background-color: #16A085;
            }
            QPushButton:disabled {
                background-color: #7F8C8D;
                color: #BDC3C7;
            }
        """)

        # Layout principal
        layout = QVBoxLayout()

        # Label de instruções
        self.instructions_label = QLabel('Carregue um arquivo CSV para converter em Excel', self)
        self.instructions_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.instructions_label)

        # Botão de upload de CSV
        self.upload_button = QPushButton('Upload CSV', self)
        self.upload_button.clicked.connect(self.process_csv)
        layout.addWidget(self.upload_button)

        # Label para mostrar o status da conversão
        self.status_label = QLabel('', self)
        self.status_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.status_label)

        # Botão de salvar como Excel
        self.save_button = QPushButton('Salvar como Excel', self)
        self.save_button.setEnabled(False)  # Desativado inicialmente
        self.save_button.clicked.connect(self.save_excel)
        layout.addWidget(self.save_button)

        # Define o layout
        self.setLayout(layout)

    def process_csv(self):
        # Abre o dialogo para selecionar o arquivo CSV
        csv_path, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo CSV", "", "CSV Files (*.csv)")
        if not csv_path:
            return

        # Exibe o caminho do arquivo carregado
        self.status_label.setText(f"Arquivo carregado: {csv_path}")
        self.save_button.setEnabled(True)
        self.csv_path = csv_path  # Guarda o caminho do arquivo CSV

    def save_excel(self):
        try:
            # Carrega o CSV usando pandas
            df = pd.read_csv(self.csv_path)

            # Abre a janela para salvar o arquivo Excel usando o método correto
            excel_path, _ = QFileDialog.getSaveFileName(self, "Salvar como Excel", "", "Excel Files (*.xlsx)")
            if not excel_path:
                return

            # Salva o arquivo como Excel
            df.to_excel(excel_path, index=False)
            self.status_label.setText(f"Arquivo salvo como: {excel_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar o arquivo: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CSVToExcelApp()
    window.show()
    sys.exit(app.exec_())
