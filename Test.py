from PyQt6.QtWidgets import *
import sys

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Main Window")
        self.setFixedSize(300, 200)

        # Create a scroll area widget
        scroll_area = QScrollArea(self)
        scroll_area.setWidgetResizable(True)
        scroll_area.setFixedSize(280, 180)

        # Create a widget to hold the contents
        container_widget = QWidget(self)
        container_layout = QVBoxLayout(container_widget)

        # Add some widgets to the layout
        for i in range(20):
            container_layout.addWidget(QLabel("Label {}".format(i)))

        # Set the container widget as the contents of the scroll area
        scroll_area.setWidget(container_widget)

app = QApplication(sys.argv)
window = MainWindow()
window.show()
sys.exit(app.exec())
