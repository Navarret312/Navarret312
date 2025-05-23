import sys
import pytesseract
from PIL import Image
import re
import pandas as pd
import json
import os
from pdf2image import convert_from_path
from docx import Document
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QTextEdit, QLineEdit, QPushButton, QInputDialog, QFileDialog, QCheckBox, QComboBox

# Defina o caminho para o executável do Tesseract
caminho = r"C:\Users\GOWTX\AppData\Local\Programs\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r"\tesseract.exe"

def extrair_texto_de_imagem(imagem_path):
    texto_completo = ""
    try:
        with Image.open(imagem_path) as img:
            for i in range(img.n_frames):  # Itera sobre todas as páginas
                img.seek(i)
                texto_completo += pytesseract.image_to_string(img)
    except Exception as e:
        texto_completo = f"Erro ao processar imagem: {e}"
    return texto_completo

def extrair_texto_de_pdf(pdf_path):
    texto_completo = ""
    try:
        # Converte cada página do PDF em uma imagem
        imagens = convert_from_path(pdf_path)
        for imagem in imagens:
            # Extrai texto de cada imagem usando pytesseract
            texto_completo += pytesseract.image_to_string(imagem)
    except Exception as e:
        texto_completo = f"Erro ao processar PDF: {e}"
    return texto_completo

def extrair_texto_de_docx(docx_path):
    doc = Document(docx_path)
    texto_completo = "\n".join(paragraph.text for paragraph in doc.paragraphs)
    return texto_completo

def extrair_texto_arquivo(file_path):
    if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.tiff', '.bmp', '.tif')):
        return extrair_texto_de_imagem(file_path)
    elif file_path.lower().endswith('.pdf'):
        return extrair_texto_de_pdf(file_path)
    elif file_path.lower().endswith('.docx'):
        return extrair_texto_de_docx(file_path)
    else:
        raise ValueError("Formato de arquivo não suportado")

class TemplateWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.dataframe = pd.DataFrame()  # Inicializa um DataFrame vazio
        self.configuracoes_registradas = []  # Lista para armazenar configurações registradas

    def initUI(self):
        layout = QVBoxLayout()

        self.texto_extraido = QTextEdit(self)
        self.texto_extraido.setReadOnly(True)
        layout.addWidget(QLabel("Texto extraído do arquivo:"))
        layout.addWidget(self.texto_extraido)

        self.chave_input = QLineEdit(self)
        layout.addWidget(QLabel("Digite as chaves que deseja buscar (separadas por vírgula):"))
        layout.addWidget(self.chave_input)

        self.num_linhas_input = QLineEdit(self)
        self.num_linhas_input.setText("1")  # Valor padrão
        layout.addWidget(QLabel("Número de linhas a serem extraídas:"))
        layout.addWidget(self.num_linhas_input)

        self.max_caracteres_input = QLineEdit(self)
        layout.addWidget(QLabel("Máximo de caracteres a serem coletados (opcional):"))
        layout.addWidget(self.max_caracteres_input)

        self.extrair_todas_checkbox = QCheckBox("Extrair todas as correspondências", self)
        layout.addWidget(self.extrair_todas_checkbox)

        self.usar_prefixo_checkbox = QCheckBox("Usar o nome do campo como prefixo para todos os resultados", self)
        layout.addWidget(self.usar_prefixo_checkbox)

        self.resultado_label = QLabel(self)
        layout.addWidget(self.resultado_label)

        self.buscar_button = QPushButton("Buscar Valor", self)
        self.buscar_button.clicked.connect(self.buscar_valor)
        layout.addWidget(self.buscar_button)

        self.registrar_button = QPushButton("Registrar campo", self)
        self.registrar_button.clicked.connect(self.registrar_campo)
        layout.addWidget(self.registrar_button)

        self.salvar_chaves_button = QPushButton("Salvar Configurações", self)
        self.salvar_chaves_button.clicked.connect(self.salvar_configuracoes)
        layout.addWidget(self.salvar_chaves_button)

        self.selecionar_arquivo_button = QPushButton("Selecionar Arquivo", self)
        self.selecionar_arquivo_button.clicked.connect(self.selecionar_arquivo)
        layout.addWidget(self.selecionar_arquivo_button)

        self.setLayout(layout)
        self.setWindowTitle('Definir Template')
        self.setGeometry(100, 100, 600, 400)

    def selecionar_arquivo(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getOpenFileName(self, "Selecionar Arquivo", "", "Todos os Arquivos (*);;Imagens (*.png;*.jpg;*.jpeg;*.bmp;*.tiff;*.tif);;PDFs (*.pdf);;Documentos Word (*.docx)", options=options)
        if file_path:
            self.carregar_texto(file_path)

    def carregar_texto(self, file_path):
        try:
            texto = extrair_texto_arquivo(file_path)
            self.texto_extraido.setPlainText(texto)
        except ValueError as e:
            self.resultado_label.setText(str(e))

    def buscar_valor(self):
        texto = self.texto_extraido.toPlainText()
        chaves = self.chave_input.text().strip().split(",")  # Divide as chaves por vírgula

        try:
            num_linhas = int(self.num_linhas_input.text().strip())
        except ValueError:
            num_linhas = 1

        try:
            max_caracteres = int(self.max_caracteres_input.text().strip())
        except ValueError:
            max_caracteres = None

        self.valores_encontrados = {}

        extrair_todas = self.extrair_todas_checkbox.isChecked()

        for chave in chaves:
            chave = chave.strip()
            chave = chave.replace(":", "").strip()

            padrao = rf"{re.escape(chave)}\s*:\s*((?:[^\n]*\n){{0,{num_linhas}}})"

            if extrair_todas:
                resultados = re.findall(padrao, texto, re.IGNORECASE | re.MULTILINE)
                if resultados:
                    for i, resultado in enumerate(resultados):
                        valor_encontrado = resultado.strip()
                        if max_caracteres is not None:
                            valor_encontrado = valor_encontrado[:max_caracteres]
                        self.valores_encontrados[f"{chave}_{i+1}"] = valor_encontrado
                    break
            else:
                resultado = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE)
                if resultado:
                    valor_encontrado = resultado.group(1).strip()
                    if max_caracteres is not None:
                        valor_encontrado = valor_encontrado[:max_caracteres]
                    self.valores_encontrados[chave] = valor_encontrado
                    break

        if self.valores_encontrados:
            resultados_formatados = "\n".join(f"{k}: {v}" for k, v in self.valores_encontrados.items())
            self.resultado_label.setText(f"Valores encontrados:\n{resultados_formatados}")
        else:
            self.resultado_label.setText("Nenhum valor encontrado para as chaves fornecidas.")

    def registrar_campo(self):
        if self.valores_encontrados:
            usar_prefixo = self.usar_prefixo_checkbox.isChecked()
            extrair_todas = self.extrair_todas_checkbox.isChecked()

            nome_campo = ""
            if usar_prefixo:
                nome_campo, ok = QInputDialog.getText(self, 'Registrar campo', 'Nome do campo para os valores encontrados:')
                if not ok or not nome_campo:
                    return

            # Registrar a configuração usando apenas a chave original
            chaves = self.chave_input.text().strip().split(",")
            chave_original = chaves[0].strip()  # Usar a primeira chave para registro
            self.configuracoes_registradas.append({
                "nome_coluna": nome_campo if usar_prefixo else chave_original,
                "chaves": chaves,  # Salvar todas as chaves
                "extrair_todas": "Sim" if extrair_todas else "Não",
                "num_linhas": self.num_linhas_input.text().strip(),
                "usar_prefixo": "Sim" if usar_prefixo else "Não",
                "max_caracteres": self.max_caracteres_input.text().strip() or None
            })

            for chave, valor in self.valores_encontrados.items():
                if usar_prefixo and extrair_todas:
                    sufixo = chave.split('_')[-1]
                    nome_coluna = f"{nome_campo}_{sufixo}"
                elif usar_prefixo:
                    nome_coluna = f"{nome_campo}"
                else:
                    nome_coluna, ok = QInputDialog.getText(self, 'Nome da coluna', f'Defina o nome da coluna para o valor encontrado ({chave}):')
                    if not ok or not nome_coluna:
                        continue

                self.dataframe.at[0, nome_coluna] = valor

            print(self.dataframe)
        else:
            self.resultado_label.setText("Nenhum valor encontrado para registrar.")

    def salvar_configuracoes(self):
        options = QFileDialog.Options()
        file_path, _ = QFileDialog.getSaveFileName(self, "Salvar Configurações", "", "JSON Files (*.json)", options=options)
        if file_path:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.configuracoes_registradas, f, ensure_ascii=False, indent=4)
            print(f"Configurações registradas salvas em '{file_path}'.")

class ExtractDataWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.template_combobox = QComboBox(self)
        self.atualizar_lista_templates()
        layout.addWidget(QLabel("Selecione um template:"))
        layout.addWidget(self.template_combobox)

        self.selecionar_pasta_button = QPushButton("Selecionar Pasta de Arquivos", self)
        self.selecionar_pasta_button.clicked.connect(self.selecionar_pasta)
        layout.addWidget(self.selecionar_pasta_button)

        self.setLayout(layout)
        self.setWindowTitle('Extrair Dados')
        self.setGeometry(100, 100, 400, 200)

        self.pasta_path = None

    def atualizar_lista_templates(self):
        self.template_combobox.clear()
        json_files = [f for f in os.listdir('.') if f.endswith('.json')]
        self.template_combobox.addItems(json_files)

    def selecionar_pasta(self):
        self.pasta_path = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Arquivos")
        if self.pasta_path:
            print(f"Pasta selecionada: {self.pasta_path}")
            self.extrair_dados()

    def extrair_dados(self):
        template_name = self.template_combobox.currentText()
        if not template_name or not self.pasta_path:
            print("Template ou pasta não selecionados.")
            return

        with open(template_name, 'r', encoding='utf-8') as f:
            configuracoes = json.load(f)

        dataframe = pd.DataFrame()

        for root, dirs, files in os.walk(self.pasta_path):
            for file in files:
                file_path = os.path.join(root, file)
                try:
                    texto = extrair_texto_arquivo(file_path)
                    print(f"Processando arquivo: {file_path}")
                    print(f"Texto extraído: {texto[:500]}...")  # Imprime os primeiros 500 caracteres do texto extraído
                    for config in configuracoes:
                        chaves = config['chaves']
                        num_linhas = int(config['num_linhas'])
                        usar_prefixo = config.get('usar_prefixo', "Não") == "Sim"
                        extrair_todas = config['extrair_todas'] == "Sim"
                        max_caracteres = int(config.get('max_caracteres', 0)) if config.get('max_caracteres') else None

                        for chave in chaves:
                            padrao = rf"{re.escape(chave)}\s*:\s*((?:[^\n]*\n){{0,{num_linhas}}})"
                            print(f"Procurando por chave: {chave} com padrão: {padrao}")

                            if extrair_todas:
                                resultados = re.findall(padrao, texto, re.IGNORECASE | re.MULTILINE)
                                if resultados:
                                    print(f"Encontrados {len(resultados)} resultados para a chave '{chave}'")
                                    for i, resultado in enumerate(resultados):
                                        valor_encontrado = resultado.strip()
                                        if max_caracteres is not None:
                                            valor_encontrado = valor_encontrado[:max_caracteres]
                                        if usar_prefixo:
                                            nome_coluna = f"{config['nome_coluna']}_{i+1}"
                                        else:
                                            nome_coluna = config['nome_coluna']
                                        dataframe.at[file, nome_coluna] = valor_encontrado
                                    break  # Interrompe após encontrar correspondências para a primeira chave válida
                            else:
                                resultado = re.search(padrao, texto, re.IGNORECASE | re.MULTILINE)
                                if resultado:
                                    valor_encontrado = resultado.group(1).strip()
                                    if max_caracteres is not None:
                                        valor_encontrado = valor_encontrado[:max_caracteres]
                                    nome_coluna = config['nome_coluna']
                                    dataframe.at[file, nome_coluna] = valor_encontrado
                                    break  # Interrompe após encontrar a primeira correspondência válida
                except Exception as e:
                    print(f"Erro ao processar arquivo {file_path}: {e}")

        print(dataframe)

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        self.definir_template_button = QPushButton("Definir Template", self)
        self.definir_template_button.clicked.connect(self.abrir_template_window)
        layout.addWidget(self.definir_template_button)

        self.extrair_dados_button = QPushButton("Extrair Dados", self)
        self.extrair_dados_button.clicked.connect(self.abrir_extract_data_window)
        layout.addWidget(self.extrair_dados_button)

        self.setLayout(layout)
        self.setWindowTitle('Menu Principal')
        self.setGeometry(100, 100, 300, 150)

    def abrir_template_window(self):
        self.template_window = TemplateWindow()
        self.template_window.show()

    def abrir_extract_data_window(self):
        self.extract_data_window = ExtractDataWindow()
        self.extract_data_window.show()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())