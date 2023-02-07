# First window
from PyQt6.QtWidgets import QMainWindow, QPushButton

class FirstWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.button = QPushButton("Open Second Window", self)
        self.button.clicked.connect(self.open_second_window)

    def open_second_window(self):
        SecondWindow.data = "Hello from First Window"
        self.second_window = SecondWindow()
        self.second_window.show()
        self.close()

# Second window
from PyQt6.QtWidgets import QMainWindow, QLabel

class SecondWindow(QMainWindow):
    data = ""

    def __init__(self):
        super().__init__()
        self.label = QLabel(SecondWindow.data, self)

# Main
from PyQt6.QtWidgets import QApplication

app = QApplication([])
first_window = FirstWindow()
first_window.show()
app.exec()
