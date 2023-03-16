import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Configuration de la fenêtre principale
        self.setWindowTitle("Fenêtre principale")
        self.setGeometry(100, 100, 400, 300)

        # Configuration du bouton
        self.button = QPushButton("Cliquez-moi", self)
        self.button.setGeometry(50, 50, 100, 50)
        
        # Connexion du signal "clicked" du bouton à la méthode correspondante
        self.button.clicked.connect(self.on_button_clicked)

    def on_button_clicked(self):
        print("Le bouton a été cliqué !")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
