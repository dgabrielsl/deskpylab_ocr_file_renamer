import os
import sys
from pathlib import Path
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.QtCore import Qt

from PyPDF2 import PdfReader
import fitz
import pytesseract
from PIL import Image
import re
from datetime import datetime

os.system('cls')

class Main(QMainWindow, QWidget):
    def __init__(self):
        super().__init__()
        self.init()
        self.site()
        self.show()

    def init(self):
        self.setWindowIcon(QIcon('icon.ico'))
        self.setWindowTitle('DeskPyL - OCR Renamer')
        self.setMinimumWidth(900)
        self.setMinimumHeight(500)

    def site(self):
        widget = QWidget()
        layer = QVBoxLayout()
        ww = self.width()

        h1 = QLabel('OCR Renamer')
        h1.setObjectName('h1-main-title')
        h1.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        h1.setMinimumWidth(ww)
        layer.addWidget(h1)

        h2 = QLabel('File renaming wizard by image reading - PyTesseract')
        h2.setObjectName('h2-main-title')
        h2.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        h2.setMinimumWidth(ww)
        layer.addWidget(h2)

        self.path_1 = QLineEdit('')
        self.path_1.setPlaceholderText('Sin configurar*')
        self.bttn_1 = QPushButton('Guardar documentos procesados en')
        self.path_2 = QLineEdit('')
        self.path_2.setPlaceholderText('Sin configurar*')
        self.bttn_2 = QPushButton('Buscar los documentos a procesar')
        self.bttn_3 = QPushButton('Procesar')

        self.path_1.setReadOnly(True)
        self.path_2.setReadOnly(True)

        self.bttn_1.setCursor(Qt.CursorShape.PointingHandCursor)
        self.bttn_2.setCursor(Qt.CursorShape.PointingHandCursor)
        self.bttn_3.setCursor(Qt.CursorShape.PointingHandCursor)

        self.bttn_1.clicked.connect(self.fd)
        self.bttn_2.clicked.connect(self.fd)
        self.bttn_3.clicked.connect(self.run_wizzard)

        self.bttn_2.setDisabled(True)
        self.bttn_3.setDisabled(True)

        layer.addWidget(self.path_1)
        layer.addWidget(self.bttn_1)
        layer.addWidget(self.path_2)
        layer.addWidget(self.bttn_2)
        layer.addWidget(self.bttn_3)

        layer.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop)

        widget.setLayout(layer)

        self.setCentralWidget(widget)

    def fd(self):
        sender = self.sender().text()
        record_path = QFileDialog.getExistingDirectory()

        if record_path != '':
            if sender == 'Guardar documentos procesados en':
                self.path_1.setText(record_path)
                self.bttn_2.setDisabled(False)

            elif sender == 'Buscar los documentos a procesar':
                self.path_2.setText(record_path)
                self.bttn_3.setDisabled(False)

    def run_wizzard(self):

        path_1 = self.path_1.text()
        path_2 = self.path_2.text()

        scan_directory = os.listdir(path_2)

        self.tree = []

        for sd in scan_directory:
            self.tree.append(f'{path_2}/{sd}')

        zoom = 4.0

        _deskpyl_temp = f'{path_2}/_deskpyl_temp'

        try: os.makedirs(_deskpyl_temp)
        except FileExistsError: pass
        except Exception as e: print(e)

        mat = fitz.Matrix(zoom,zoom)
        pattern = r'^(\D{5,6})(\s*)(\#*)(\D\-\d{4,10})'

        self.processed_documents = 0

        dt = datetime.now()
        print(f'Start time: {dt.hour}:{dt.minute}:{dt.second}')

        for leaf in self.tree:

            doc = fitz.open(leaf)
            self.group = []
            counter = 1
            document = ''
            document_name_found = False

            for d in doc:
                pix = d.get_pixmap(matrix=mat)
                res = f'{_deskpyl_temp}/img #{counter}.jpg'
                self.group.append(res)
                pix.save(res)

                raw_text = str(((pytesseract.image_to_string(Image.open(res)))))
                raw_text = raw_text.lower().replace('é','e')
                text_to_lines = raw_text.split('\n')
                counter += 1

                for ttl in text_to_lines:
                    if re.search(pattern,ttl):
                        doc.close()
                        document = ttl.replace('pagare','').replace('.','').replace('#','').replace(' ','')
                        document = document.upper()
                        document_name_found = True
                        os.rename(leaf,f'{path_1}/{document}.pdf')
                        print(f'({self.processed_documents + 1}/{len(self.tree)}) {path_2}/{document}.pdf successfully renamed...')
                        break

                if document_name_found: break

            self.processed_documents += 1

        try:
            for g in self.group:
                try: os.remove(g)
                except Exception as e: print(e)
            self.group.clear()
        except Exception as e: print(e)

        try: os.removedirs(_deskpyl_temp)
        except Exception as e: print(e)

        counter = 1

        dt = datetime.now()
        print(f'End time: {dt.hour}:{dt.minute}:{dt.second}')

        QMessageBox.information(
            self,
            'DeskPyL - OCR Renamer Widget',
            '\nThe wizzard has been ended the process...\t\n',
            QMessageBox.StandardButton.Ok,
            QMessageBox.StandardButton.Ok
        )

if __name__ == '__main__':
    app = QApplication(sys.argv)
    # app.setStyleSheet(Path('style.qss').read_text())
    app.setStyleSheet("""
        QWidget{
            background: #001a20;
            color: #0fd;
            font-size: 16px;
        }
        QLineEdit{
            margin: 5px 0;
            padding: 10px 20px;
            background: #fff;
            color: #024;
            border-radius: 20px;
        }
        QPushButton{
            margin-bottom: 40px;
            padding: 10px;
            background: #aff;
            color: #035;
            border-radius: 20px;
        }
        QPushButton:hover{
            background: #5dd;
        }
        QPushButton:focus{
            background: #0bf;
        }
        #h1-main-title{
            padding: 12px 0;
            background: #002a35;
            font-size: 20px;
            border-radius: 20px;
        }
        #h2-main-title{
            margin-bottom: 30px;
            font-style: italic;
        }
    """)
    win = Main()
    sys.exit(app.exec())