import sys
from PyQt6.QtWidgets import QApplication, QMainWindow, QTableView, QComboBox
from PyQt6.QtGui import QStandardItemModel, QStandardItem
from PyQt6.QtCore import *

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Create a table model
        model = QStandardItemModel(4, 2)
        model.setHorizontalHeaderLabels(['Column 1', 'Column 2'])

        for row in range(4):
            for column in range(2):
                item = QStandardItem("Row {0}, Column {1}".format(row, column))
                model.setItem(row, column, item)

        
        # Create a sort filter proxy model
        proxyModel = QSortFilterProxyModel()
        proxyModel.setSourceModel(model)

        # Create a table view
        self.tableView = QTableView()
        self.tableView.setModel(proxyModel)

        # Enable sorting and filtering on the table view
        self.tableView.setSortingEnabled(True)
        self.tableView.setAlternatingRowColors(True)

        # Set the layout
        self.setCentralWidget(self.tableView)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
