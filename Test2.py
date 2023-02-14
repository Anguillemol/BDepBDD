from PyQt5.QtWidgets import QApplication, QMainWindow, QTableView, QMenu, QLineEdit, QVBoxLayout, QWidget, QAction, QFileDialog
from PyQt5.QtCore import Qt, QSortFilterProxyModel
from PyQt5.QtGui import QStandardItemModel, QStandardItem
import sys

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set up the window
        self.setWindowTitle('Table with Filters')
        self.setGeometry(100, 100, 800, 600)

        # Create the table view
        self.table_view = QTableView()
        self.model = QStandardItemModel()
        self.table_view.setModel(self.model)

        # Set up the table view
        self.table_view.horizontalHeader().setSectionsClickable(True)
        self.table_view.horizontalHeader().setContextMenuPolicy(Qt.CustomContextMenu) # Set the context menu policy for the header
        self.table_view.horizontalHeader().customContextMenuRequested.connect(self.show_filter_menu)

        # Add some data to the table
        self.model.setHorizontalHeaderLabels(['Name', 'Age', 'Gender'])
        for row in range(10):
            name_item = QStandardItem(f'Person {row}')
            age_item = QStandardItem(str(20 + row))
            gender_item = QStandardItem('Male' if row % 2 == 0 else 'Female')
            self.model.appendRow([name_item, age_item, gender_item])

        # Set the layout
        central_widget = QWidget()
        layout = QVBoxLayout()
        layout.addWidget(self.table_view)
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

    # Show the filter menu for a specific column
    def show_filter_menu(self, position):
        column = self.table_view.horizontalHeader().logicalIndexAt(position)
        filter_menu = QMenu()
        filter_widget = QLineEdit()
        filter_widget.textChanged.connect(self.model.layoutChanged.emit)
        filter_widget.textChanged.connect(lambda text: setattr(filter_widget, 'filter', text))
        filter_widget.returnPressed.connect(filter_menu.close)
        filter_widget.installEventFilter(self)
        filter_action = QAction('Filter', self)
        #filter_action.setDefaultWidget(filter_widget)
        filter_widget.setWindowRole(QAction.ActionPosition)
        filter_action.setShortcut('Ctrl+F')
        filter_menu.addAction(filter_action)
        filter_menu.addSeparator()
        filter_model = QSortFilterProxyModel()
        filter_model.setSourceModel(self.model)
        filter_model.setFilterKeyColumn(column)
        filter_model.setFilterCaseSensitivity(Qt.CaseInsensitive)
        self.table_view.setModel(filter_model)
        self.table_view.horizontalHeader().setSectionResizeMode(QTableView.ResizeToContents)
        filter_widget.setFocus()
        filter_widget.selectAll()
        filter_widget.filter = ''
        self.table_view.horizontalHeader().viewport().update()
        filter_menu.exec_(self.table_view.horizontalHeader().mapToGlobal(position))

    # Create the filter menu for a specific column
    def create_filter_menu(self, filter_model, filter_column):
        filter_menu = QMenu()
        filter_widget = QLineEdit()
        filter_widget.textChanged.connect(filter_model.setFilterFixedString)
        filter_action = QAction('Filter', self)
        filter_action.triggered.connect(filter_widget.setFocus)
        filter_menu.addAction(filter_action)

        filter_model.setFilterKeyColumn(filter_column)
        filter_model.setFilterCaseSensitivity(Qt.CaseInsensitive)
        filter_model.setFilterRole(Qt.DisplayRole)

        return filter_menu

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
