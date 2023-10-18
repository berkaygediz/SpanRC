import sys
import csv
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.table = QTableWidget(self)
        self.setCentralWidget(self.table)
        self.table.clear()
        self.table.setRowCount(30)
        self.table.setColumnCount(15)
        self.undoStack = QUndoStack(self)
        self.createActions()
        self.createMenu()
        self.setWindowTitle('SpanRC')
        self.setGeometry(100, 100, 800, 600)
        self.statusBar().showMessage('Loaded.', 1000)
        self.show()

    def createActions(self):
        self.newAction = QAction('New', self)
        self.newAction.setShortcut('Ctrl+N')
        self.newAction.triggered.connect(self.table.clear)
        self.openAction = QAction('Open', self)
        self.openAction.setShortcut('Ctrl+O')
        self.openAction.triggered.connect(self.tableLoad)
        self.saveAction = QAction('Save', self)
        self.saveAction.setShortcut('Ctrl+S')
        self.saveAction.triggered.connect(self.tableSave)
        self.printAction = QAction('Print', self)
        self.printAction.setShortcut('Ctrl+P')
        self.printAction.triggered.connect(self.printTable)
        self.exitAction = QAction('Exit', self)
        self.exitAction.setShortcut('Ctrl+Q')
        self.exitAction.triggered.connect(self.closeApp)
        self.undoAction = self.undoStack.createUndoAction(self, 'Undo')
        self.undoAction.setShortcut('Ctrl+Z')
        self.redoAction = self.undoStack.createRedoAction(self, 'Redo')
        self.redoAction.setShortcut('Ctrl+Y')
        self.deleteAction = QAction('Delete', self)
        self.deleteAction.setShortcut('Del')
        self.deleteAction.triggered.connect(self.deleteCell)

    def deleteCell(self):
        for item in self.table.selectedItems():
            item.setText('')

    def createMenu(self):
        menubar = self.menuBar()
        fileMenu = menubar.addMenu('File')
        fileMenu.addAction(self.newAction)
        fileMenu.addAction(self.openAction)
        fileMenu.addAction(self.saveAction)
        fileMenu.addAction(self.printAction)
        fileMenu.addAction(self.exitAction)
        editMenu = menubar.addMenu('Edit')
        editMenu.addAction(self.deleteAction)
        editMenu.addAction(self.undoAction)
        editMenu.addAction(self.redoAction)


    def tableLoad(self):
        filePath, _ = QFileDialog.getOpenFileName(self, 'Open', '', 'CSV Files (*.csv);;All Files (*)')

        if filePath:
            with open(filePath, 'r', newline='') as file:
                data = list(csv.reader(file))

            self.table.setRowCount(len(data))
            self.table.setColumnCount(len(data[0]))
            for rowid, row in enumerate(data):
                for colid, cell in enumerate(row):
                    item = QTableWidgetItem(cell)
                    self.table.setItem(rowid, colid, item)

    def tableSave(self):
        filePath, _ = QFileDialog.getSaveFileName(self, 'Save', '', 'CSV Files (*.csv);;All Files (*)')

        if filePath:
            with open(filePath, 'w', newline='') as file:
                writer = csv.writer(file)
                for rowid in range(self.table.rowCount()):
                    row = []
                    for colid in range(self.table.columnCount()):
                        item = self.table.item(rowid, colid)
                        if item is not None:
                            row.append(item.text())
                        else:
                            row.append('')
                    writer.writerow(row)

    def printTable(self):
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec_() == QPrintDialog.Accepted:
            self.table.render(printer)

    def closeApp(self, event):
        reply = QMessageBox.question(self, 'Exit', 'Are you sure you want to exit?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setOrganizationName("berkaygediz")
    app.setApplicationName("SpanRC")
    app.setApplicationVersion("1.0.0")
    main_window = MainWindow()
    main_window.show()
    sys.exit(app.exec_())
