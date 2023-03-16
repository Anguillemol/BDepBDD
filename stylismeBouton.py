import sys
import io
import tempfile
import os
import shutil
import openpyxl as wb
import pandas as pd
import numpy as np

from pathlib import Path
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.client_credential import ClientCredential
from office365.sharepoint.files.file import File 
from datetime import datetime
from  cryptography.fernet import Fernet


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        # Configuration de la fenêtre principale
        self.setWindowTitle("Fenêtre principale")
        self.setFixedSize(600,400)

        button_style = '''
        QPushButton {
            outline: 0;
            border: none;
            padding: 0 56px;
            height: 45px;
            line-height: 45px;
            border-radius: 7px;
            font-weight: 400;
            font-size: 16px;
            background: #fff;
            color: #696969;
        }
        QPushButton:hover {
            background: rgba(255,255,255,0.9);
            box-shadow: 0 6px 20px rgb(93 93 93 / 23%);
            transition: background 0.2s ease,color 0.2s ease,box-shadow 0.2s ease;
        }
        '''
        self.button = QPushButton("Mon Bouton")
        self.button.setStyleSheet(button_style)

        self.button.setMaximumSize(QSize(300, 100))
        iconCreer = QIcon()
        iconCreer.addPixmap(QPixmap("Icons/create.png"), QIcon.Mode.Normal, QIcon.State.Off)
        self.button.setIcon(iconCreer)


        self.layoutV1 = QHBoxLayout()
        self.layoutV1.addWidget(self.button)
        self.setLayout(self.layoutV1)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
