
import sys
from PyQt5.QtWidgets import QApplication, QGraphicsView, QGraphicsScene, QFileDialog, QPushButton, QVBoxLayout, QWidget, QGraphicsRectItem, QInputDialog, QLabel, QDialog, QDialogButtonBox
from PyQt5.QtGui import QPixmap, QPainter, QPen
from PyQt5.QtCore import Qt, QRectF
import pytesseract
from PIL import Image
import numpy as np

# Defina o caminho para o Tesseract
caminho = r"C:\Users\GOWTX\AppData\Local\Programs\Tesseract-OCR"
pytesseract.pytesseract.tesseract_cmd = caminho + r"\tesseract.exe"

class ImageViewer(QGraphicsView):
    def __init__(self):
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
            self.rectangles.append(self.rect_item)
            self.rect_item = None

    def toggle_selection_mode(self):
        self.selection_mode = not self.selection_mode

    def get_selected_areas(self):
        areas = []
        for rect in self.rectangles:
            rect_coords = rect.rect()
            areas.append((rect_coords.left(), rect_coords.top(), rect_coords.right(), rect_coords.bottom()))
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

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.imagens = []  # Lista para armazenar as imagens
        self.indice_atual = -1  # Índice da imagem atual
        self.image_files = []  # Lista para armazenar os caminhos dos arquivos de imagem

    def initUI(self):
        self.setWindowTitle('Template Extractor')
        layout = QVBoxLayout()

        self.viewer = ImageViewer()
        layout.addWidget(self.viewer)

        btn_open = QPushButton('Abrir Imagem', self)
        btn_open.clicked.connect(self.abrir_imagem)
        layout.addWidget(btn_open)

        btn_select = QPushButton('Selecionar Campos', self)
        btn_select.clicked.connect(self.viewer.toggle_selection_mode)
        layout.addWidget(btn_select)

        btn_extract = QPushButton('Extrair Dados', self)
        btn_extract.clicked.connect(self.extrair_dados)
        layout.addWidget(btn_extract)

        btn_prev = QPushButton('Página Anterior', self)
        btn_prev.clicked.connect(self.pagina_anterior)
        layout.addWidget(btn_prev)

        btn_next = QPushButton('Próxima Página', self)
        btn_next.clicked.connect(self.proxima_pagina)
        layout.addWidget(btn_next)

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
        for coords in selected_areas:
            x1, y1, x2, y2 = map(int, coords)
            image = self.viewer.pixmap_item.pixmap().toImage()
            width = image.width()
            height = image.height()
            buffer = image.bits().asstring(width * height * 4)  # Assuming RGBA format

            pil_image = Image.frombuffer("RGBA", (width, height), buffer, "raw", "RGBA", 0, 1)
            cropped_image = pil_image.crop((x1, y1, x2, y2))
            text = pytesseract.image_to_string(cropped_image)
            extracted_data.append((coords, text))

        dialog = ExtractDialog(extracted_data, self)
        dialog.exec_()

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

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    ex.show()
    sys.exit(app.exec_())

