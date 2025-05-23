import sys
import json
import os
import pandas as pd
from PyQt5.QtWidgets import QApplication, QGraphicsView, QGraphicsScene, QFileDialog, QPushButton, QVBoxLayout, QWidget, QGraphicsRectItem, QInputDialog, QLabel, QDialog, QDialogButtonBox, QComboBox, QHBoxLayout, QToolTip
from PyQt5.QtGui import QPixmap, QPainter, QPen, QColor
from PyQt5.QtCore import Qt, QRectF
import pytesseract
from PIL import Image

# Defina o caminho para o Tesseract
caminho = r"C:\Users\GOWTX\AppData\Local\Programs\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r"\tesseract.exe"

class ImageViewer(QGraphicsView):
    def __init__(self, df, coordinates):
        super().__init__()
        self.setScene(QGraphicsScene(self))
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)

        self.pixmap_item = None
        self.scale_factor = 1.0
        self.selection_mode = False
        self.rect_start_pos = None
        self.rect_item = None
        self.rectangles = []
        self.df = df  # DataFrame para armazenar os campos
        self.coordinates = coordinates  # Lista para armazenar coordenadas

    def abrir_imagem(self, caminho_imagem):
        pixmap = QPixmap(caminho_imagem)
        if self.pixmap_item:
            self.scene().removeItem(self.pixmap_item)
        self.pixmap_item = self.scene().addPixmap(pixmap)
        self.setSceneRect(self.pixmap_item.boundingRect())
        self.resetTransform()
        self.scale(1.0, 1.0)

    def wheelEvent(self, event):
        if event.angleDelta().y() > 0:
            self.scale(1.2, 1.2)
        else:
            self.scale(1 / 1.2, 1 / 1.2)

    def mousePressEvent(self, event):
        if self.selection_mode and event.button() == Qt.LeftButton:
            self.rect_start_pos = self.mapToScene(event.pos())
            self.rect_item = QGraphicsRectItem()
            self.rect_item.setPen(QPen(Qt.red, 2))
            self.scene().addItem(self.rect_item)

    def mouseMoveEvent(self, event):
        if self.selection_mode and self.rect_start_pos and self.rect_item:
            rect_end_pos = self.mapToScene(event.pos())
            rect = QRectF(self.rect_start_pos, rect_end_pos).normalized()
            self.rect_item.setRect(rect)

    def mouseReleaseEvent(self, event):
        if self.selection_mode and self.rect_item:
            # Solicitar nome do campo
            campo_nome, ok = QInputDialog.getText(self, 'Nome do Campo', 'Digite o nome do campo:')
            if ok and campo_nome:  # Verifica se o usuário clicou em OK e forneceu um nome
                x1, y1, x2, y2 = self.rect_item.rect().getCoords()
                
                # Extração de texto
                image = self.pixmap_item.pixmap().toImage()
                width = image.width()
                height = image.height()
                buffer = image.bits().asstring(width * height * 4)  # Assuming RGBA format

                pil_image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "RGBA", 0, 1)
                cropped_image = pil_image.crop((x1, y1, x2, y2))
                text = pytesseract.image_to_string(cropped_image)

                # Adicionar os dados ao DataFrame usando pd.concat
                new_row = pd.DataFrame({'Campo': [campo_nome], 'Valor': [text], 'X1': [int(x1)], 'Y1': [int(y1)], 'X2': [int(x2)], 'Y2': [int(y2)]})
                self.df = pd.concat([self.df, new_row], ignore_index=True)

                # Adicionar coordenadas à lista
                self.coordinates.append({'Campo': campo_nome, 'X1': int(x1), 'Y1': int(y1), 'X2': int(x2), 'Y2': int(y2)})

                self.rectangles.append((campo_nome, self.rect_item))  # Armazena o nome e o retângulo
            self.rect_item = None

    def toggle_selection_mode(self):
        self.selection_mode = not self.selection_mode

    def get_selected_areas(self):
        areas = []
        for campo_nome, rect in self.rectangles:
            rect_coords = rect.rect()
            areas.append((campo_nome, rect_coords.left(), rect_coords.top(), rect_coords.right(), rect_coords.bottom()))
        return areas

class ExtractDialog(QDialog):
    def __init__(self, extracted_data, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Dados Extraídos")

        layout = QVBoxLayout(self)
        for i, (coords, text) in enumerate(extracted_data):
            label = QLabel(f"Área {i + 1} ({coords}):\n{text.strip()}")
            layout.addWidget(label)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok, self)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)

class TemplateExtractor(QWidget):
    def __init__(self):
        super().__init__()
        self.btn_style = """
        QPushButton {
            background-color: #4CAF50; /* Verde */
            color: white;
            padding: 10px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
        }
        QPushButton:hover {
            background-color: #45a049; /* Verde mais escuro */
        }
        """
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Template Extractor')
        self.setStyleSheet("background-color: #084B8A;")  # Cor de fundo da janela
        layout = QVBoxLayout()

        btn_define_template = QPushButton('Definir Template', self)
        btn_define_template.setStyleSheet(self.btn_style)
        btn_define_template.clicked.connect(self.open_template_screen)

        btn_extract_data = QPushButton('Extrair Dados', self)
        btn_extract_data.setStyleSheet(self.btn_style)
        btn_extract_data.clicked.connect(self.open_extraction_screen)

        # Layout horizontal para centralizar os botões
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn_define_template)
        btn_layout.addWidget(btn_extract_data)
        btn_layout.setAlignment(Qt.AlignCenter)  # Centraliza os botões

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def open_template_screen(self):
        # Solicitar nome do template
        template_name, ok = QInputDialog.getText(self, 'Nome do Template', 'Digite o nome do template:')
        if ok and template_name:
            self.template_screen = TemplateScreen(template_name, self.btn_style)
            self.template_screen.show()

    def open_extraction_screen(self):
        self.extraction_screen = ExtractionScreen()
        self.extraction_screen.show()

class TemplateScreen(QWidget):
    def __init__(self, template_name, btn_style):
        super().__init__()
        self.template_name = template_name  # Armazena o nome do template
        self.df = pd.DataFrame(columns=['Campo', 'Valor', 'X1', 'Y1', 'X2', 'Y2'])  # DataFrame para armazenar os campos
        self.coordinates = []  # Lista para armazenar coordenadas
        self.btn_style = btn_style  # Armazena o estilo do botão
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        self.viewer = ImageViewer(self.df, self.coordinates)  # Passa o DataFrame e a lista de coordenadas para o ImageViewer
        layout.addWidget(self.viewer)

        # Função para criar um botão com tooltip
        def create_button_with_tooltip(text, tooltip, callback):
            btn = QPushButton(text, self)
            btn.setStyleSheet(self.btn_style)
            btn.setToolTip(tooltip)
            btn.clicked.connect(callback)
            return btn

        btn_open = create_button_with_tooltip('Abrir Imagem', 'Clique para abrir uma imagem.', self.abrir_imagem)
        btn_select = create_button_with_tooltip('Selecionar Campos', 'Clique para selecionar campos na imagem.', self.viewer.toggle_selection_mode)
        btn_extract = create_button_with_tooltip('Extrair Dados', 'Clique para extrair dados da imagem.', self.extrair_dados)
        btn_prev = create_button_with_tooltip('Página Anterior', 'Clique para voltar à página anterior.', self.pagina_anterior)
        btn_next = create_button_with_tooltip('Próxima Página', 'Clique para avançar para a próxima página.', self.proxima_pagina)
        btn_save_coordinates = create_button_with_tooltip('Salvar Coordenadas', 'Clique para salvar as coordenadas dos campos.', self.save_coordinates)

        # Layout horizontal para centralizar os botões
        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn_open)
        btn_layout.addWidget(btn_select)
        btn_layout.addWidget(btn_extract)
        btn_layout.addWidget(btn_prev)
        btn_layout.addWidget(btn_next)
        btn_layout.addWidget(btn_save_coordinates)
        btn_layout.setAlignment(Qt.AlignCenter)  # Centraliza os botões

        layout.addLayout(btn_layout)
        self.setLayout(layout)

    def abrir_imagem(self):
        # Permitir selecionar um arquivo TIFF com várias páginas
        caminho_imagem, _ = QFileDialog.getOpenFileName(self, 'Abrir Imagem', '', 'Images (*.tif *.tiff *.png *.jpg *.jpeg *.bmp *.gif)')
        if caminho_imagem:
            self.image_files = []  # Limpa a lista de arquivos de imagem
            self.indice_atual = 0  # Reseta o índice para a primeira imagem
            self.load_images(caminho_imagem)  # Carrega as imagens do arquivo

    def load_images(self, caminho_imagem):
        # Carrega todas as páginas de um arquivo TIFF
        with Image.open(caminho_imagem) as img:
            for i in range(img.n_frames):
                img.seek(i)
                img_data = img.convert("RGBA")  # Converte para um formato compatível
                buffer = img_data.tobytes()  # Obtém os dados da imagem
                self.image_files.append(buffer)  # Armazena os dados da imagem
            self.viewer.abrir_imagem(caminho_imagem)  # Abre a primeira imagem

    def extrair_dados(self):
        selected_areas = self.viewer.get_selected_areas()
        if not selected_areas:
            return

        extracted_data = []
        rows_to_add = []  # Lista para armazenar as linhas a serem adicionadas ao DataFrame

        for campo_nome, x1, y1, x2, y2 in selected_areas:
            # Extração de texto
            image = self.viewer.pixmap_item.pixmap().toImage()
            width = image.width()
            height = image.height()
            buffer = image.bits().asstring(width * height * 4)  # Assuming RGBA format

            pil_image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "RGBA", 0, 1)
            cropped_image = pil_image.crop((x1, y1, x2, y2))
            text = pytesseract.image_to_string(cropped_image)

            # Adicionar os dados à lista
            rows_to_add.append({'Campo': campo_nome, 'Valor': text, 'X1': x1, 'Y1': y1, 'X2': x2, 'Y2': y2})

            extracted_data.append((campo_nome, text))

        # Criar um DataFrame apenas uma vez
        if rows_to_add:
            new_rows_df = pd.DataFrame(rows_to_add)
            self.df = pd.concat([self.df, new_rows_df], ignore_index=True)

        dialog = ExtractDialog(extracted_data, self)
        dialog.exec_()

        # Exibir o DataFrame no console
        print(self.df)

    def save_coordinates(self):
        with open(f'{self.template_name}_coordenadas.json', 'w') as json_file:
            json.dump(self.coordinates, json_file, indent=4)
        print(f"Coordenadas salvas em {self.template_name}_coordenadas.json")

    def pagina_anterior(self):
        if self.indice_atual > 0:
            self.indice_atual -= 1
            self.show_image()  # Mostra a imagem anterior

    def proxima_pagina(self):
        if self.indice_atual < len(self.image_files) - 1:
            self.indice_atual += 1
            self.show_image()  # Mostra a próxima imagem

    def show_image(self):
        # Exibe a imagem correspondente ao índice atual
        if self.image_files:
            img_data = self.image_files[self.indice_atual]
            img = Image.frombytes("RGBA", (self.viewer.pixmap_item.pixmap().width(), self.viewer.pixmap_item.pixmap().height()), img_data)
            img.save("temp_image.png")  # Salva a imagem temporariamente para visualização
            self.viewer.abrir_imagem("temp_image.png")

class ExtractionScreen(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()
        
        # Dropdown para selecionar o arquivo JSON
        self.json_dropdown = QComboBox(self)
        self.json_dropdown.setStyleSheet("""
        QComboBox {
            background-color: #f0f0f0; /* Cor de fundo */
            border: 1px solid #ccc; /* Borda */
            border-radius: 5px; /* Bordas arredondadas */
            padding: 5px; /* Espaçamento interno */
            font-size: 16px; /* Tamanho da fonte */
        }
        QComboBox::drop-down {
            border: none; /* Remove a borda do dropdown */
        }
        """)
        self.load_json_files()
        layout.addWidget(self.json_dropdown)

        # Botão para extrair dados de todos os arquivos em uma pasta
        btn_extract_folder = QPushButton('Extrair Dados de Pasta', self)
        btn_extract_folder.setStyleSheet("""
        QPushButton {
            background-color: #4CAF50; /* Verde */
            color: white;
            padding: 10px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
        }
        QPushButton:hover {
            background-color: #45a049; /* Verde mais escuro */
        }
        """)
        btn_extract_folder.clicked.connect(self.extract_data_from_folder)
        layout.addWidget(btn_extract_folder)

        # Botão para extrair dados de arquivos específicos
        btn_extract_files = QPushButton('Extrair Dados de Arquivos Específicos', self)
        btn_extract_files.setStyleSheet("""
        QPushButton {
            background-color: #4CAF50; /* Verde */
            color: white;
            padding: 10px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
        }
        QPushButton:hover {
            background-color: #45a049; /* Verde mais escuro */
        }
        """)
        btn_extract_files.clicked.connect(self.extract_data_from_files)
        layout.addWidget(btn_extract_files)

        self.setLayout(layout)

    def load_json_files(self):
        # Carrega todos os arquivos JSON do diretório atual para o dropdown
        json_files = [f for f in os.listdir('.') if f.endswith('.json')]
        self.json_dropdown.addItems(json_files)

    def extract_data_from_folder(self):
        # Obter o arquivo JSON selecionado
        json_file_path = self.json_dropdown.currentText()
        if not json_file_path:
            print("Nenhum arquivo JSON selecionado.")
            return

        try:
            with open(json_file_path, 'r') as json_file:
                coordinates = json.load(json_file)
        except FileNotFoundError:
            print("Arquivo coordenadas.json não encontrado.")
            return

        # Permitir ao usuário selecionar uma pasta
        pasta = QFileDialog.getExistingDirectory(self, 'Selecionar Pasta', '')
        if not pasta:
            return

        # Criar um DataFrame para armazenar todos os dados extraídos
        all_extracted_data = pd.DataFrame(columns=[coord['Campo'] for coord in coordinates])  # Inicializa o DataFrame com as colunas baseadas no JSON

        # Iterar sobre todos os arquivos na pasta
        for arquivo in os.listdir(pasta):
            if arquivo.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff')):
                caminho_imagem = os.path.join(pasta, arquivo)

                # Criar um dicionário para armazenar os dados extraídos deste arquivo
                extracted_data = {'Arquivo': arquivo}  # Adiciona o nome do arquivo

                for coord in coordinates:
                    campo_nome = coord['Campo']
                    x1, y1, x2, y2 = coord['X1'], coord['Y1'], coord['X2'], coord['Y2']

                    # Extração de texto
                    image = Image.open(caminho_imagem)
                    cropped_image = image.crop((x1, y1, x2, y2))
                    text = pytesseract.image_to_string(cropped_image)

                    # Adicionar o valor extraído ao dicionário
                    extracted_data[campo_nome] = text

                # Adicionar os dados extraídos ao DataFrame
                all_extracted_data = pd.concat([all_extracted_data, pd.DataFrame([extracted_data])], ignore_index=True)

        # Salvar os dados extraídos em um arquivo Excel
        all_extracted_data.to_excel('dados_extraidos.xlsx', index=False)
        print("Dados extraídos salvos em dados_extraidos.xlsx")

    def extract_data_from_files(self):
        # Obter o arquivo JSON selecionado
        json_file_path = self.json_dropdown.currentText()
        if not json_file_path:
            print("Nenhum arquivo JSON selecionado.")
            return

        try:
            with open(json_file_path, 'r') as json_file:
                coordinates = json.load(json_file)
        except FileNotFoundError:
            print("Arquivo coordenadas.json não encontrado.")
            return

        # Permitir ao usuário selecionar arquivos específicos
        arquivos, _ = QFileDialog.getOpenFileNames(self, 'Selecionar Arquivos', '', 'Images (*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff)')
        if not arquivos:
            return

        # Criar um DataFrame para armazenar todos os dados extraídos
        all_extracted_data = pd.DataFrame(columns=[coord['Campo'] for coord in coordinates])  # Inicializa o DataFrame com as colunas baseadas no JSON

        # Iterar sobre os arquivos selecionados
        for caminho_imagem in arquivos:
            # Criar um dicionário para armazenar os dados extraídos deste arquivo
            extracted_data = {'Arquivo': os.path.basename(caminho_imagem)}  # Adiciona o nome do arquivo

            for coord in coordinates:
                campo_nome = coord['Campo']
                x1, y1, x2, y2 = coord['X1'], coord['Y1'], coord['X2'], coord['Y2']

                # Extração de texto
                image = Image.open(caminho_imagem)
                cropped_image = image.crop((x1, y1, x2, y2))
                text = pytesseract.image_to_string(cropped_image)

                # Adicionar o valor extraído ao dicionário
                extracted_data[campo_nome] = text

            # Adicionar os dados extraídos ao DataFrame
            all_extracted_data = pd.concat([all_extracted_data, pd.DataFrame([extracted_data])], ignore_index=True)

        # Salvar os dados extraídos em um arquivo Excel
        all_extracted_data.to_excel('dados_extraidos.xlsx', index=False)
        print("Dados extraídos salvos em dados_extraidos.xlsx")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Define um estilo moderno
    main_window = TemplateExtractor()
    main_window.show()
    sys.exit(app.exec_())