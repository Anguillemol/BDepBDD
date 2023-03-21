from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon, QPainter, QColor, QPixmap, QBrush
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QLabel

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Créer un label rouge
        label = QLabel("15")
        label.setPixmap(self.create_pixmap(QColor("red"), QSize(160, 160)))

        # Créer un bouton avec l'icône du label rouge
        button = QPushButton()
        button.setIcon(QIcon(label.pixmap()))
        button.setIconSize(label.size())

        # Ajouter le bouton à la fenêtre principale
        self.setCentralWidget(button)

    def create_pixmap(self, color, size):
        # Créer un pixmap avec un fond transparent
        pixmap = QPixmap(size)
        pixmap.fill(QColor(0,0,0,0))

        # Dessiner un cercle rouge sur le pixmap
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QColor(0,0,0,0))
        painter.setBrush(QBrush(color))
        painter.drawEllipse(pixmap.rect())
        painter.end()

        return pixmap

if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()
