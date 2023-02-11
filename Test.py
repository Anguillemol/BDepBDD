import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableView, QComboBox
from PyQt6.QtGui import QStandardItemModel, QStandardItem

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Create a table view
        self.tableView = QTableView(self)

        # Create a combobox
        self.comboBox = QComboBox(self)
        self.comboBox.addItem("Data 1")
        self.comboBox.addItem("Data 2")

        # Connect the currentIndexChanged signal to the updateData slot function
        self.comboBox.currentIndexChanged.connect(self.updateData)

        # Set the layout
        self.setCentralWidget(self.tableView)
        self.comboBox.move(0, 0)

    def updateData(self, index):
        # Get the selected data from the combobox
        data = ["Data 1", "Data 2"][index]

        # Update the data in the table view
        model = QStandardItemModel(4, 4)
        for row in range(4):
            for column in range(4):
                item = QStandardItem(f"{data} - ({row}, {column})")
                model.setItem(row, column, item)
        self.tableView.setModel(model)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
