import sys
from PyQt6.QtCore import *
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from PyQt6.uic import loadUi

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        layout = QVBoxLayout()
        self.edit = QLineEdit()
        self.edit.textChanged.connect(self.filter)
        layout.addWidget(self.edit)
        
        data = [
            ('France', 'Paris'),
            ('United Kingdom', 'London'),
            ('Italy', 'Rome'),
            ('Germany', 'Berlin'),
            ('QWERTY', 'AZERTY')
        ]
        
        self.tableView = QTableView()
        self.model = QStandardItemModel(4, 2)
        for row in range(5):
            for column in range(2):
                item = QStandardItem(data[row][column])
                self.model.setItem(row, column, item)
                
        self.proxyModel = QSortFilterProxyModel(self.tableView)
        self.proxyModel.setFilterCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.proxyModel.setSourceModel(self.model)
        self.tableView.setModel(self.proxyModel)
        
    
        layout.addWidget(self.tableView)
        self.setLayout(layout)

    def filter(self, filter_text):
        self.proxyModel.setFilterFixedString(filter_text)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
