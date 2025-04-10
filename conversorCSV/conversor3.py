import sys
import pandas as pd
import pdfplumber
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QFileDialog, QMessageBox
from PyQt5.QtCore import Qt

class PDFToExcelApp(QWidget):
    def __init__(self):
        super().__init__()

        self.init_ui()

    def init_ui(self):
        # Configuração da janela
        self.setWindowTitle('Conversor PDF para Excel')
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
        self.instructions_label = QLabel('Carregue um arquivo PDF para converter em Excel', self)
        self.instructions_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.instructions_label)

        # Botão de upload de PDF
        self.upload_button = QPushButton('Upload PDF', self)
        self.upload_button.clicked.connect(self.process_pdf)
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

    def process_pdf(self):
        # Abre o diálogo para selecionar o arquivo PDF
        pdf_path, _ = QFileDialog.getOpenFileName(self, "Selecione o arquivo PDF", "", "PDF Files (*.pdf)")
        if not pdf_path:
            return

        try:
            # Extrai a tabela do PDF usando pdfplumber
            with pdfplumber.open(pdf_path) as pdf:
                tables = []
                for page in pdf.pages:
                    table = page.extract_table()
                    if table:
                        tables.append(pd.DataFrame(table[1:], columns=table[0]))  # Ignora o cabeçalho

            if not tables:
                QMessageBox.warning(self, "Aviso", "Nenhuma tabela encontrada no arquivo PDF.")
                return

            # Combina todas as tabelas extraídas em um único DataFrame
            self.df = pd.concat(tables, ignore_index=True)
            self.status_label.setText(f"Arquivo carregado: {pdf_path}")
            self.save_button.setEnabled(True)
            self.pdf_path = pdf_path  # Guarda o caminho do arquivo PDF
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao carregar o arquivo PDF: {str(e)}")

    def save_excel(self):
        try:
            # Abre a janela para salvar o arquivo Excel
            excel_path, _ = QFileDialog.getSaveFileName(self, "Salvar como Excel", "", "Excel Files (*.xlsx)")
            if not excel_path:
                return

            # Salva o arquivo como Excel
            self.df.to_excel(excel_path, index=False)
            self.status_label.setText(f"Arquivo salvo como: {excel_path}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro ao salvar o arquivo: {str(e)}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = PDFToExcelApp()
    window.show()
    sys.exit(app.exec_())
