import os
import sys
import csv
import datetime
import time
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtTest import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *

# NOT STABLE - DO NOT USE - NOT FINISHED - ONLY FOR COMMIT - NOT WORKING - NOT TESTED

class SRC_Threading(QThread):
    update_signal = pyqtSignal()

    def __init__(self, parent=None):
        super(SRC_Threading, self).__init__(parent)
        self.running = False

    def run(self):
        if not self.running:
            self.running = True
            time.sleep(0.15)
            self.update_signal.emit()
            self.running = False

class SRC_Workbook(QMainWindow):
    def __init__(self, parent=None):
        super(SRC_Workbook, self).__init__(parent)
        starttime = datetime.datetime.now()
        self.src_thread = SRC_Threading()
        self.src_thread.update_signal.connect(self.update_statistics)
        self.default_values = {
            'row': 1,
            'column': 1,
            'rowspan': 1,
            'columnspan': 1,
            'text': '',
            'theme': 'light',
            'is_saved': None,
            'file_name': None,
            'default_directory': None,
        }
        self.selected_file = None
        self.file_name = None
        self.is_saved = None
        self.default_directory = QDir().homePath()
        self.directory = self.default_directory
        self.src_table = QTableWidget(self)
        self.setCentralWidget(self.src_table)
        self.status_bar = self.statusBar()
        self.setup_dock()
        self.dock_widget.hide()
        self.update_statistics()
        self.src_table.itemSelectionChanged.connect(self.src_thread.start)
        self.setup_actions()
        self.setup_toolbar()
        self.light_theme = QPalette()
        self.dark_theme = QPalette()
        self.light_theme.setColor(QPalette.Window, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.WindowText, QColor(37, 38, 39))
        self.dark_theme.setColor(QPalette.Window, QColor(37, 38, 39))
        self.dark_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Base, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Text, QColor(42, 43, 44))
        self.dark_theme.setColor(QPalette.Base, QColor(42, 43, 44))
        self.dark_theme.setColor(QPalette.Text, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Highlight, QColor(218, 221, 177))
        self.light_theme.setColor(QPalette.HighlightedText, QColor(42, 43, 44))
        self.dark_theme.setColor(QPalette.Highlight, QColor(218, 221, 177))
        self.dark_theme.setColor(QPalette.HighlightedText, QColor(42, 43, 44))
        self.setPalette(self.light_theme)
        self.src_restorestate()
        self.src_restoretheme()
        self.update_title()
        self.setUnifiedTitleAndToolBarOnMac(True)
        self.setWindowIcon(QIcon('icon.png'))
        self.showMaximized()
        self.src_table.setFocus()
        self.src_table.resizeColumnsToContents()
        self.src_table.resizeRowsToContents()
        if self.src_table.columnCount() == 0 and self.src_table.rowCount() == 0:
            self.src_table.setColumnCount(100)
            self.src_table.setRowCount(50)
            self.src_table.clearSpans()
            self.src_table.setItem(0, 0, QTableWidgetItem(''))
        endtime = datetime.datetime.now()
        self.statusBar().showMessage(str((endtime - starttime).total_seconds()) + " ms", 2000)

    def update_title(self):
        file = self.file_name if self.file_name else 'Untitled'
        if self.is_saved == True: asterisk = ""
        else: asterisk = "*"
        self.setWindowTitle(f"{file}{asterisk} - SpanRC")

    def src_restorestate(self):
        settings = QSettings()
        self.restoreGeometry(settings.value('geometry', self.saveGeometry()))
        self.restoreState(settings.value('windowState', self.saveState()))
        self.src_table.setColumnCount(int(settings.value('columnCount', self.src_table.columnCount())))
        self.src_table.setRowCount(int(settings.value('rowCount', self.src_table.rowCount())))
        self.src_table.clearSpans()
        for row in range(self.src_table.rowCount()):
            for column in range(self.src_table.columnCount()):
                if settings.value(f'row{row}column{column}rowspan', None) and settings.value(f'row{row}column{column}columnspan', None):
                    self.src_table.setSpan(row, column, int(settings.value(f'row{row}column{column}rowspan')), int(settings.value(f'row{row}column{column}columnspan')))
        for row in range(self.src_table.rowCount()):
            for column in range(self.src_table.columnCount()):
                if settings.value(f'row{row}column{column}text', None):
                    self.src_table.setItem(row, column, QTableWidgetItem(settings.value(f'row{row}column{column}text')))
        self.src_table.resizeColumnsToContents()
        self.src_table.resizeRowsToContents()
        self.src_table.setCurrentCell(int(settings.value('currentRow', 0)), int(settings.value('currentColumn', 0)))
        self.src_table.scrollToItem(self.src_table.item(int(settings.value('currentRow', 0)), int(settings.value('currentColumn', 0))))
        self.src_table.setFocus()

    def src_restoretheme(self):
        settings = QSettings()
        if settings.value('theme', 'light') == 'light':
            self.setPalette(self.light_theme)
        elif settings.value('theme', 'light') == 'dark':
            self.setPalette(self.dark_theme)

    def src_savestate(self):
        settings = QSettings("berkaygediz", "SpanRC")
        settings.setValue('geometry', self.saveGeometry())
        settings.setValue('windowState', self.saveState())
        settings.setValue('columnCount', self.src_table.columnCount())
        settings.setValue('rowCount', self.src_table.rowCount())
        settings.setValue('currentRow', self.src_table.currentRow())
        settings.setValue('currentColumn', self.src_table.currentColumn())
        for row in range(self.src_table.rowCount()):
            for column in range(self.src_table.columnCount()):
                if self.src_table.span(row, column) != (1, 1):
                    settings.setValue(f'row{row}column{column}rowspan', self.src_table.span(row, column)[0])
                    settings.setValue(f'row{row}column{column}columnspan', self.src_table.span(row, column)[1])
                if self.src_table.item(row, column) != None:
                    settings.setValue(f'row{row}column{column}text', self.src_table.item(row, column).text())
        settings.setValue('theme', self.default_values['theme'])

    def setup_actions(self):
        self.newAction = self.create_action('New', 'Clear', self.src_table.clear, QKeySequence.New)
        self.openAction = self.create_action('Open', 'Open', self.open, QKeySequence.Open)
        self.saveAction = self.create_action('Save', 'Save', self.save, QKeySequence.Save)
        self.saveasAction = self.create_action('Save As', 'Save As', self.save, QKeySequence.SaveAs)
        self.printAction = self.create_action('Print', 'Print', self.print, QKeySequence.Print)
        self.exitAction = self.create_action('Exit', 'Exit', self.closeEvent, QKeySequence.Quit)
        self.deleteAction = self.create_action('Delete', 'Delete', self.deletecell, QKeySequence.Delete)
        self.aboutAction = self.create_action('About', 'About', self.about)
        self.aboutQtAction = self.create_action('About Qt', 'About Qt', qApp.aboutQt)

    def about(self):
        QMessageBox.about(self, "About SpanRC",
                "<b>SpanRC</b> v1.1.0<br><br>"
                "A row & column spanned table editor.<br><br>")

    def create_action(self, text, status_tip, function, shortcut=None):
        action = QAction(text, self)
        action.setStatusTip(status_tip)
        action.triggered.connect(function)
        if shortcut:
            action.setShortcut(shortcut)
        return action
    
    def toolbarlabel(self, toolbar, text):
        label = QLabel(f"<b>{text}</b>")
        toolbar.addWidget(label)

    def setup_toolbar(self):
        self.toolbar = self.addToolBar('File')
        self.toolbarlabel(self.toolbar, 'File: ')
        self.toolbar.addAction(self.newAction)
        self.toolbar.addAction(self.openAction)
        self.toolbar.addAction(self.saveAction)
        self.toolbar.addAction(self.saveasAction)
        self.toolbar.addAction(self.printAction)
        self.toolbar.addAction(self.exitAction)
        self.toolbar = self.addToolBar('Interface')
        self.toolbarlabel(self.toolbar, 'UI: ')
        self.toolbar.addAction(self.aboutAction)
        self.toolbar.addAction(self.aboutQtAction)
        self.addToolBarBreak()
        self.toolbar = self.addToolBar('Edit')
        self.toolbarlabel(self.toolbar, 'Edit: ')
        self.toolbar.addAction(self.deleteAction)

    def setup_dock(self):
        self.dock_widget = QDockWidget('Help', self)
        self.statistics_label = QLabel()
        self.help_label = QLabel()
        self.help_label.setWordWrap(True)
        self.help_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.help_label.setTextFormat(Qt.RichText)
        self.help_label.setText("<html><head><style>"   
                                "table {border-collapse: collapse; width: 100%;}"
                                "th, td {text-align: left; padding: 8px;}"
                                "tr:nth-child(even) {background-color: #f2f2f2;}"
                                "tr:hover {background-color: #ddd;}"
                                "th {background-color: #4CAF50; color: white;}"
                                "</style></head><body>"
                                "<table><tr>"
                                "<th>Shortcut</th>"
                                "<th>Function</th>"
                                "</tr>"
                                "<tr><td>Ctrl + N</td><td>New</td></tr>"
                                "<tr><td>Ctrl + O</td><td>Open</td></tr>"
                                "<tr><td>Ctrl + S</td><td>Save</td></tr>"
                                "<tr><td>Ctrl + Shift + S</td><td>Save As</td></tr>"
                                "<tr><td>Ctrl + P</td><td>Print</td></tr>"
                                "<tr><td>Ctrl + Q</td><td>Exit</td></tr>"
                                "<tr><td>Ctrl + D</td><td>Delete</td></tr>"
                                "<tr><td>Ctrl + A</td><td>About</td></tr>"
                                "<tr><td>Ctrl + Shift + A</td><td>About Qt</td></tr>"
                                "</table></body></html>")
        self.dock_widget.setWidget(self.help_label)
        self.dock_widget.setObjectName("Help")
        self.dock_widget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

    def create_dock_widget(self):
        widget = QWidget()
        layout = QVBoxLayout()
        widget.setLayout(layout)
        self.statistics = QLabel()
        self.statistics.setText('Statistics')
        layout.addWidget(self.statistics)
        return widget
    
    def update_statistics(self):
        row = self.src_table.rowCount()
        column = self.src_table.columnCount()
        selected_cell = (self.src_table.currentRow() + 1, self.src_table.currentColumn() + 1)

        statistics = (
            "<html><head><style>"
            "table {border-collapse: collapse; width: 100%;}"
            "th, td {text-align: left; padding: 8px;}"
            "tr:nth-child(even) {background-color: #f2f2f2;}"
            ".highlight {background-color: #E2E3E1; color: #000000}"
            "tr:hover {background-color: #ddd;}"
            "th {background-color: #4CAF50; color: white;}"
            "#sr-text { background-color: #E2E3E1; color: #000000; }"
            "</style></head><body>"
            "<table>"
            "<tr><th>Statistics</th>"
        )
        statistics += f"<td>Rows: </td><td>{row}</td><td>Cols: </td><td>{column}</td>"
        statistics += f"<td>Selected Cell: </td><td>{selected_cell[0]}:{selected_cell[1]}</td>"
        if self.src_table.selectedRanges():
            statistics += "<td>Selected Range: </td><td>"
            for selected_range in self.src_table.selectedRanges():
                statistics += (
                    f"{selected_range.topRow() + 1}:{selected_range.leftColumn() + 1} - "
                    f"{selected_range.bottomRow() + 1}:{selected_range.rightColumn() + 1}<br>"
                )
        else:
            statistics += "<td>Selected Range: </td><td>None</td>"

        statistics += "</td><td id='sr-text'>SpanRC</td></tr></table></body></html>"
        self.update_title()
        self.statistics_label.setText(statistics)
        self.statusBar().addPermanentWidget(self.statistics_label)

    def add_toolbar_button(self, toolbar, text, status_tip, function, shortcut=None):
        button = QToolButton()
        button.setText(text)
        button.setStatusTip(status_tip)
        button.clicked.connect(function)
        if shortcut:
            button.setShortcut(shortcut)
        toolbar.addWidget(button)

    def add_toolbar_label(self, toolbar, text):
        label = QLabel(f"<b>{text}</b>")
        toolbar.addWidget(label)

    def deletecell(self):
        for item in self.src_table.selectedItems():
            item.setText('')

    def open(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_filter = "CSV Files (*.csv)"
        selected_file, _ = QFileDialog.getOpenFileName(self, "Open", self.directory, file_filter, options=options)
        if selected_file:
            self.selected_file = selected_file
            self.openfile()
            self.directory = os.path.dirname(self.selected_file)
            return True
        else:
            return False
        
    def openfile(self):
        if not self.selected_file:
            return self.open()
        else:
            with open(self.selected_file, 'r') as file:
                self.src_table.clearSpans()
                self.src_table.setRowCount(0)
                self.src_table.setColumnCount(0)
                for rowdata in csv.reader(file):
                    row = self.src_table.rowCount()
                    self.src_table.insertRow(row)
                    if len(rowdata) > self.src_table.columnCount():
                        self.src_table.setColumnCount(len(rowdata))
                    for column, data in enumerate(rowdata):
                        item = QTableWidgetItem(data)
                        self.src_table.setItem(row, column, item)
                self.is_saved = True
                self.file_name = os.path.basename(self.selected_file)
                self.setWindowTitle(self.file_name)

    def save(self):
        if not self.selected_file:
            self.saveas()
            self.openfile()
            return True
        else:
            self.savefile()
            return True

    def saveas(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_filter = "CSV Files (*.csv)"
        selected_file, _ = QFileDialog.getSaveFileName(self, "Save As", self.directory, file_filter, options=options)
        if selected_file:
            self.selected_file = selected_file
            self.savefile()
            self.directory = os.path.dirname(self.selected_file)
            return True
        else:
            return False

    def savefile(self):
        if not self.selected_file:
            self.saveas()
            self.openfile()
        else:
            with open(self.selected_file, 'w', newline='') as file:
                writer = csv.writer(file)
                for rowid in range(self.src_table.rowCount()):
                    row = []
                    for colid in range(self.src_table.columnCount()):
                        item = self.src_table.item(rowid, colid)
                        if item is not None:
                            row.append(item.text())
                        else:
                            row.append('')
                    writer.writerow(row)
                    
    def print(self):
        printer = QPrinter(QPrinter.HighResolution)
        dialog = QPrintDialog(printer, self)
        if dialog.exec_() == QPrintDialog.Accepted:
            self.src_table.render(printer)

    def closeEvent(self, event):
        if self.is_saved == False:
            reply = QMessageBox.question(self, 'SpanRC',
                                         "Are you sure to save SpanRC?", QMessageBox.Yes |
                                         QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.Yes:
                self.src_savestate()
                event.accept()
            else:
                event.ignore()
        else:
            self.src_savestate()
            event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    core = SRC_Workbook()
    app.setOrganizationName('berkaygediz')
    app.setApplicationName('SpanRC')
    app.setApplicationDisplayName('SpanRC')
    app.setApplicationVersion('1.1.0')
    core.show()
    sys.exit(app.exec_())
