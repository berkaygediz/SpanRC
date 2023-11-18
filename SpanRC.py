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
from modules.translations import *


class SRC_UndoCommand(QUndoCommand):
    def __init__(self, table, old_data, new_data, row, col):
        super().__init__()
        self.table = table
        self.row = row
        self.col = col
        self.old_data = old_data
        self.new_data = new_data

    def redo(self):
        item = self.table.item(self.row, self.col)
        self.table.blockSignals(True)
        item.setText(self.new_data)
        self.table.blockSignals(False)

    def undo(self):
        item = self.table.item(self.row, self.col)
        self.table.blockSignals(True)
        item.setText(self.old_data)
        self.table.blockSignals(False)


class SRC_About(QMainWindow):
    def __init__(self, parent=None):
        super(SRC_About, self).__init__(parent)
        self.setWindowFlags(Qt.WindowCloseButtonHint)
        self.setWindowModality(Qt.ApplicationModal)
        self.setGeometry(QStyle.alignedRect(Qt.LeftToRight, Qt.AlignCenter, self.size(
        ), QApplication.desktop().availableGeometry()))
        self.about_label = QLabel()
        self.about_label.setWordWrap(True)
        self.about_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.about_label.setTextFormat(Qt.RichText)
        self.about_label.setText("<center>"
                                 "<b>SpanRC</b><br>"
                                 "A powerful spreadsheet application<br><br>"
                                 "Made by Berkay Gediz<br>"
                                 "Apache License 2.0</center>")
        self.setCentralWidget(self.about_label)


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
        settings = QSettings("berkaygediz", "SpanRC")
        self.undo_stack = QUndoStack(self)
        self.undo_stack.setUndoLimit(100)
        self.src_thread = SRC_Threading()
        self.src_thread.update_signal.connect(self.SRC_updateStatistics)
        self.default_values = {
            'row': 1,
            'column': 1,
            'rowspan': 1,
            'columnspan': 1,
            'text': '',
            'window_geometry': None,
            'is_saved': None,
            'file_name': None,
            'default_directory': None,
            'current_theme': 'light',
            'current_language': 'English',
            'file_type': 'xsrc',
            'currentRow': 0,
            'currentColumn': 0,
            'rowCount': 50,
            'columnCount': 100,
            'windowState': None
        }
        self.light_theme = QPalette()
        self.dark_theme = QPalette()
        self.light_theme.setColor(QPalette.Window, QColor(239, 213, 195))
        self.light_theme.setColor(QPalette.WindowText, QColor(37, 38, 39))
        self.light_theme.setColor(QPalette.Base, QColor(230, 230, 230))
        self.light_theme.setColor(QPalette.Text, QColor(0, 0, 0))
        self.light_theme.setColor(QPalette.Highlight, QColor(221, 216, 184))
        self.light_theme.setColor(QPalette.HighlightedText, QColor(0, 0, 0))
        self.dark_theme.setColor(QPalette.Window, QColor(58, 68, 93))
        self.dark_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Base, QColor(94, 87, 104))
        self.dark_theme.setColor(QPalette.Text, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Highlight, QColor(221, 216, 184))
        self.setWindowIcon(QIcon('icon.png'))
        self.selected_file = None
        self.file_name = None
        self.is_saved = None
        self.default_directory = QDir().homePath()
        self.directory = self.default_directory
        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()
        self.language = settings.value("current_language")
        self.SRC_setupDock()
        self.status_bar = self.statusBar()
        self.src_table = QTableWidget(self)
        self.setCentralWidget(self.src_table)
        self.SRC_setupActions()
        self.SRC_setupToolbar()
        self.setPalette(self.light_theme)
        self.src_table.itemSelectionChanged.connect(self.src_thread.start)
        self.showMaximized()
        QTimer.singleShot(50, self.SRC_restoreTheme)
        QTimer.singleShot(150, self.SRC_restoreState)
        self.dock_widget.hide()
        self.src_table.setFocus()
        self.last_opened_file = self.SRC_lastOpened()
        if self.last_opened_file and os.path.exists(self.last_opened_file):
            self.loadFile(self.last_opened_file)

        if self.file_name != None:
            self.src_table.resizeColumnsToContents()
            self.src_table.resizeRowsToContents()

        if self.src_table.columnCount() == 0 and self.src_table.rowCount() == 0:
            self.src_table.setColumnCount(100)
            self.src_table.setRowCount(50)
            self.src_table.clearSpans()
            self.src_table.setItem(0, 0, QTableWidgetItem(''))

        self.SRC_updateTitle()
        endtime = datetime.datetime.now()
        self.statusBar().showMessage(str((endtime - starttime).total_seconds()) + " ms", 2000)

    def closeEvent(self, event):
        settings = QSettings("berkaygediz", "SpanRC")
        reply = QMessageBox.question(self, 'SpanRC',
                                     translations[settings.value(
                                         "current_language")]["exit_message"], QMessageBox.Yes |
                                     QMessageBox.No, QMessageBox.No)
        if self.is_saved is False:
            if reply == QMessageBox.Yes:
                self.SRC_saveState()
                event.accept()
            else:
                self.SRC_saveState()
                event.ignore()
        else:
            if reply == QMessageBox.Yes:
                self.SRC_saveState()
                event.accept()
            else:
                self.SRC_saveState()
                event.ignore()

    def SRC_changeLanguage(self):
        language = self.language_combobox.currentText()
        settings = QSettings("berkaygediz", "SpanRC")
        settings.setValue("current_language", language)
        settings.sync()

    def SRC_lastOpened(self):
        settings = QSettings("berkaygediz", "SpanRC")
        return settings.value('last_opened_file', '')

    def SRC_updateTitle(self):
        settings = QSettings("berkaygediz", "SpanRC")
        file = self.file_name if self.file_name else translations[settings.value(
            "current_language")]["new_title"]
        if self.is_saved == True:
            asterisk = ""
        else:
            asterisk = "*"
        self.setWindowTitle(f"{file}{asterisk} — SpanRC")

    def SRC_updateStatistics(self):
        settings = QSettings("berkaygediz", "SpanRC")
        row = self.src_table.rowCount()
        column = self.src_table.columnCount()
        selected_cell = (self.src_table.currentRow() + 1,
                         self.src_table.currentColumn() + 1)

        statistics = (
            "<html><head><style>"
            "table { width: 100%; border-spacing: 0;}"
            "th, td {text-align: left; padding: 8px;}"
            "tr:nth-child(even) {background-color: #f2f2f2;}"
            ".highlight {background-color: #E2E3E1; color: #000000}"
            "tr:hover {background-color: #ddd;}"
            "th {background-color: #4CAF50; color: white;}"
            "#sr-text { background-color: #E2E3E1; color: #000000; font-weight: bold;}"
            "</style></head><body>"
            "<table><tr>"
            f"<th>{translations[settings.value('current_language')]['statistics_title']}</th>"
        )
        statistics += f"<td>{translations[settings.value('current_language')]['statistics_message1']}</td><td>{row}</td><td>Cols: </td><td>{column}</td>"
        statistics += f"<td>{translations[settings.value('current_language')]['statistics_message2']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"
        if self.src_table.selectedRanges():
            statistics += f"<td>{translations[settings.value('current_language')]['statistics_message3']}</td><td>"
            for selected_range in self.src_table.selectedRanges():
                statistics += (
                    f"{selected_range.topRow() + 1}:{selected_range.leftColumn() + 1} - "
                    f"{selected_range.bottomRow() + 1}:{selected_range.rightColumn() + 1}</td>"
                )
        else:
            statistics += f"<td>{translations[settings.value('current_language')]['statistics_message4']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"

        statistics += "</td><td id='sr-text'>SpanRC</td></tr></table></body></html>"
        self.SRC_updateTitle()
        self.statistics_label.setText(statistics)
        self.statusBar().addPermanentWidget(self.statistics_label)

    def SRC_saveState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        settings.setValue('window_geometry', self.saveGeometry())
        self.is_saved = settings.value('is_saved', False)
        self.file_name = settings.value('file_name', '')
        self.default_directory = settings.value(
            'default_directory', QDir().homePath())
        settings.setValue(
            'current_theme', 'dark' if self.palette() == self.dark_theme else 'light')
        settings.setValue("current_language",
                          self.language_combobox.currentText())
        settings.setValue('file_type', self.default_values['file_type'])

        if self.selected_file:
            settings.setValue('last_opened_file', self.selected_file)

        settings.sync()

    def SRC_restoreState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        self.language_combobox.setCurrentText(
            settings.value("current_language"))
        geometry = settings.value('window_geometry')
        if geometry is not None:
            self.restoreGeometry(geometry)

        self.src_table.setColumnCount(
            int(settings.value('columnCount', self.src_table.columnCount())))
        self.src_table.setRowCount(
            int(settings.value('rowCount', self.src_table.rowCount())))
        self.restoreState(settings.value('windowState', self.saveState()))

        self.last_opened_file = settings.value('last_opened_file', '')
        if self.last_opened_file and os.path.exists(self.last_opened_file):
            self.loadFile(self.last_opened_file)

        for row in range(self.src_table.rowCount()):
            for column in range(self.src_table.columnCount()):
                if settings.value(f'row{row}column{column}rowspan', None) and settings.value(f'row{row}column{column}columnspan', None):
                    self.src_table.setSpan(row, column, int(settings.value(f'row{row}column{column}rowspan')), int(
                        settings.value(f'row{row}column{column}columnspan')))
        for row in range(self.src_table.rowCount()):
            for column in range(self.src_table.columnCount()):
                if settings.value(f'row{row}column{column}text', None):
                    self.src_table.setItem(row, column, QTableWidgetItem(
                        settings.value(f'row{row}column{column}text')))

        if self.file_name != None:
            self.src_table.resizeColumnsToContents()
            self.src_table.resizeRowsToContents()

        self.is_saved = settings.value('is_saved', self.is_saved)
        self.file_name = settings.value('file_name', self.file_name)
        self.default_directory = settings.value(
            'default_directory', self.default_directory)
        self.selected_file = settings.value('last_opened_file', '')
        self.src_table.setCurrentCell(
            int(settings.value('currentRow', 0)), int(settings.value('currentColumn', 0)))
        self.src_table.scrollToItem(self.src_table.item(
            int(settings.value('currentRow', 0)), int(settings.value('currentColumn', 0))))
        self.src_table.setFocus()

        self.SRC_restoreTheme(),
        self.SRC_updateTitle()

    def SRC_restoreTheme(self):
        theme = QSettings("berkaygediz", "SpanRC")
        if theme.value('current_theme') == 'dark':
            self.setPalette(self.dark_theme)
        else:
            self.setPalette(self.light_theme)
        self.SRC_toolbarTheme()

    def SRC_toolbarTheme(self):
        palette = self.palette()
        if palette == self.light_theme:
            text_color = QColor(37, 38, 39)
        else:
            text_color = QColor(255, 255, 255)

        for toolbar in self.findChildren(QToolBar):
            for action in toolbar.actions():
                if action.text():
                    action_color = QPalette()
                    action_color.setColor(QPalette.ButtonText, text_color)
                    action_color.setColor(QPalette.WindowText, text_color)
                    toolbar.setPalette(action_color)

    def SRC_themeAction(self):
        if self.palette() == self.light_theme:
            self.setPalette(self.dark_theme)
        else:
            self.setPalette(self.light_theme)

        self.SRC_toolbarTheme()

    def SRC_setupActions(self):
        settings = QSettings("berkaygediz", "SpanRC")
        self.newAction = self.SRC_createAction(
            translations[settings.value("current_language")]["new"], translations[settings.value("current_language")]["new_title"], self.new, QKeySequence.New)
        self.openAction = self.SRC_createAction(
            translations[settings.value("current_language")]["open"], translations[settings.value("current_language")]["open_title"], self.open, QKeySequence.Open)
        self.saveAction = self.SRC_createAction(
            translations[settings.value("current_language")]["save"], translations[settings.value("current_language")]["save_title"], self.save, QKeySequence.Save)
        self.saveasAction = self.SRC_createAction(
            translations[settings.value("current_language")]["save_as"], translations[settings.value("current_language")]["save_as_title"], self.saveAs, QKeySequence.SaveAs)
        self.printAction = self.SRC_createAction(
            translations[settings.value("current_language")]["print"], translations[settings.value("current_language")]["print_title"], self.print, QKeySequence.Print)
        self.exitAction = self.SRC_createAction(
            translations[settings.value("current_language")]["exit"], translations[settings.value("current_language")]["exit_message"], self.close, QKeySequence.Quit)
        self.deleteAction = self.SRC_createAction(
            translations[settings.value("current_language")]["delete"], translations[settings.value("current_language")]["delete_title"], self.deletecell, QKeySequence.Delete)
        self.aboutAction = self.SRC_createAction(
            translations[settings.value("current_language")]["about"], translations[settings.value("current_language")]["about_title"], self.showAbout, QKeySequence.HelpContents)
        self.undoAction = self.SRC_createAction(
            translations[settings.value("current_language")]["undo"], translations[settings.value("current_language")]["undo_title"], self.undo_stack.undo, QKeySequence.Undo)
        self.redoAction = self.SRC_createAction(
            translations[settings.value("current_language")]["redo"], translations[settings.value("current_language")]["redo_title"], self.undo_stack.redo, QKeySequence.Redo)
        self.darklightAction = self.SRC_createAction(
            "Dark/Light", "Change theme", self.SRC_themeAction)

    def SRC_createAction(self, text, status_tip, function, shortcut=None):
        action = QAction(text, self)
        action.setStatusTip(status_tip)
        action.triggered.connect(function)
        if shortcut:
            action.setShortcut(shortcut)
        return action

    def SRC_setupToolbar(self):
        settings = QSettings("berkaygediz", "SpanRC")

        self.file_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["file"])
        self.file_toolbar.setObjectName('File')
        self.SRC_toolbarLabel(self.file_toolbar, translations[settings.value(
            "current_language")]["file"] + ": ")
        self.file_toolbar.addActions([self.newAction, self.openAction, self.saveAction,
                                      self.saveasAction, self.printAction, self.exitAction])

        self.edit_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["edit"])
        self.edit_toolbar.setObjectName('Edit')
        self.SRC_toolbarLabel(self.edit_toolbar, translations[settings.value(
            "current_language")]["edit"] + ": ")
        self.edit_toolbar.addActions(
            [self.undoAction, self.redoAction, self.deleteAction])

        self.addToolBarBreak()
        self.formula_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["formula"])
        self.formula_toolbar.setObjectName('Formula')
        self.SRC_toolbarLabel(self.formula_toolbar, translations[settings.value(
            "current_language")]["formula"] + ": ")
        self.formula_edit = QLineEdit(self)
        self.formula_edit.returnPressed.connect(self.calculateFormula)
        if self.palette() == self.light_theme:
            self.formula_edit.setStyleSheet(
                "background-color: rgb(239, 213, 195); color: #000000")
            self.formula_toolbar.setStyleSheet(
                "background-color: rgb(239, 213, 195); color: #000000")
        else:
            self.formula_edit.setStyleSheet(
                "background-color: rgb(58, 68, 93); color: #FFFFFF")
            self.formula_toolbar.setStyleSheet(
                "background-color: rgb(58, 68, 93); color: #FFFFFF")

        self.formula_toolbar.addWidget(self.formula_edit)

        self.interface_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["interface"])
        self.interface_toolbar.setObjectName('Interface')
        self.SRC_toolbarLabel(self.interface_toolbar, translations[settings.value(
            "current_language")]["interface"] + ": ")
        self.interface_toolbar.addAction(self.darklightAction)
        self.interface_toolbar.addAction(self.dock_widget.toggleViewAction())
        self.interface_toolbar.addAction(self.aboutAction)
        self.language_combobox = QComboBox(self)
        self.language_combobox.addItems(
            ["English", "Türkçe", "Azərbaycanca", "Deutsch", "Español"])
        self.language_combobox.currentIndexChanged.connect(
            self.SRC_changeLanguage)
        self.interface_toolbar.addWidget(self.language_combobox)

    def SRC_setupDock(self):
        settings = QSettings("berkaygediz", "SpanRC")
        self.dock_widget = QDockWidget(translations[settings.value(
            "current_language")]["help"], self)
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
                                "<table><tr><th>Shortcut</th><th>Function</th></tr>"
                                f"<tr><td>Ctrl + N</td><td>{translations[settings.value('current_language')]['new_title']}</td></tr>"
                                f"<tr><td>Ctrl + O</td><td>{translations[settings.value('current_language')]['open_title']}</td></tr>"
                                f"<tr><td>Ctrl + S</td><td>{translations[settings.value('current_language')]['save_title']}</td></tr>"
                                f"<tr><td>Ctrl + Shift + S</td><td>{translations[settings.value('current_language')]['save_as_title']}</td></tr>"
                                f"<tr><td>Ctrl + P</td><td>{translations[settings.value('current_language')]['print_title']}</td></tr>"
                                f"<tr><td>Ctrl + Q</td><td>{translations[settings.value('current_language')]['exit_title']}</td></tr>"
                                f"<tr><td>Ctrl + D</td><td>{translations[settings.value('current_language')]['delete_title']}</td></tr>"
                                f"<tr><td>Ctrl + A</td><td>{translations[settings.value('current_language')]['about_title']}</td></tr>"
                                "</table></body></html>")
        self.dock_widget.setWidget(self.help_label)
        self.dock_widget.setObjectName('Help')
        self.dock_widget.setAllowedAreas(
            Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

    def SRC_createDockwidget(self):
        widget = QWidget()
        layout = QVBoxLayout()
        widget.setLayout(layout)
        self.statistics = QLabel()
        self.statistics.setText('Statistics')
        layout.addWidget(self.statistics)
        return widget

    def SRC_toolbarLabel(self, toolbar, text):
        label = QLabel(f"<b>{text}</b>")
        toolbar.addWidget(label)

    def new(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.is_saved == False:
            reply = QMessageBox.question(self, 'SpanRC',
                                         translations[settings.value(
                                             "current_language")]["new_message"],
                                         QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.Yes:
                self.SRC_saveState()
                self.src_table.clearSpans()
                self.src_table.setRowCount(50)
                self.src_table.setColumnCount(100)
                self.is_saved = False
                self.file_name = None
                self.setWindowTitle(translations[settings.value(
                    "current_language")]["new_title"] + " — SpanRC")
                self.directory = self.default_directory
                return True
            else:
                return False
        else:
            self.src_table.clearSpans()
            self.src_table.setRowCount(50)
            self.src_table.setColumnCount(100)
            self.is_saved = False
            self.file_name = None
            self.setWindowTitle(translations[settings.value(
                "current_language")]["new_title"] + " — SpanRC")
            self.directory = self.default_directory
            return True

    def open(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.is_saved is False:
            reply = QMessageBox.question(self, 'SpanRC',
                                         translations[settings.value(
                                             "current_language")]["open_message"],
                                         QMessageBox.No, QMessageBox.No)

            if reply == QMessageBox.Yes:
                self.SRC_saveState()
                return self.openfile()
        else:
            return self.openfile()

    def openfile(self):
        settings = QSettings("berkaygediz", "SpanRC")
        options = QFileDialog.Options()
        file_filter = "SpanRC Workbook (*.xsrc);;Comma Separated Values (*.csv)"
        selected_file, _ = QFileDialog.getOpenFileName(
            self, translations[settings.value("current_language")]["open_title"] + " — SpanRC", self.directory, file_filter, options=options)

        if selected_file:
            self.loadFile(selected_file)
            return True
        return False

    def loadFile(self, file_path):
        self.selected_file = file_path
        self.file_name = os.path.basename(self.selected_file)
        self.directory = os.path.dirname(self.selected_file)
        self.setWindowTitle(self.file_name)

        if file_path.endswith('.xsrc') or file_path.endswith('.csv'):
            self.loadTable(file_path)

        self.is_saved = True
        self.file_name = os.path.basename(self.selected_file)
        self.directory = os.path.dirname(self.selected_file)
        self.SRC_updateTitle()
        return True

    def loadTable(self, file_path):
        with open(file_path, 'r') as file:
            reader = csv.reader(file)
            self.src_table.clearSpans()
            self.src_table.setRowCount(0)
            self.src_table.setColumnCount(0)
            for row in reader:
                rowPosition = self.src_table.rowCount()
                self.src_table.insertRow(rowPosition)
                self.src_table.setColumnCount(len(row))
                for column in range(len(row)):
                    item = QTableWidgetItem(row[column])
                    self.src_table.setItem(rowPosition, column, item)
        self.src_table.resizeColumnsToContents()
        self.src_table.resizeRowsToContents()

    def save(self):
        if not self.selected_file:
            if not self.saveAs():
                return False

        self.saveFile()
        self.directory = os.path.dirname(self.selected_file)
        self.SRC_updateTitle()
        self.SRC_saveState()
        return True

    def saveAs(self):
        options = QFileDialog.Options()
        settings = QSettings("berkaygediz", "SpanRC")
        options |= QFileDialog.ReadOnly
        file_filter = f"{translations[settings.value('current_language')]['xsrc']} (*.xsrc);;Comma Separated Values (*.csv)"
        selected_file, _ = QFileDialog.getSaveFileName(
            self, translations[settings.value("current_language")]["save_as_title"] + " — SpanRC", self.directory, file_filter, options=options)
        if selected_file:
            self.directory = os.path.dirname(self.selected_file)
            self.selected_file = selected_file
            self.saveFile()
            return True
        else:
            return False

    def saveFile(self):
        if not self.file_name:
            self.saveAs()
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

        self.file_name = os.path.basename(self.selected_file)
        self.directory = os.path.dirname(self.selected_file)
        self.status_bar.showMessage("Saved.", 2000)
        self.is_saved = True
        self.SRC_updateTitle()

    def print(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.Letter)
        printer.setOrientation(QPrinter.Landscape)
        printer.setFullPage(True)
        printer.setPageMargins(0, 0, 0, 0, QPrinter.Millimeter)
        printer.setDocName(self.file_name)

        preview_dialog = QPrintPreviewDialog(printer, self)
        preview_dialog.paintRequested.connect(self.printPreview)
        preview_dialog.exec_()

    def printPreview(self, printer):
        self.src_table.render(printer)

    def showAbout(self):
        self.showAbout = SRC_About()
        self.showAbout.show()

    def selectedCells(self):
        selected_cells = self.src_table.selectedRanges()

        if not selected_cells:
            raise ValueError("ERROR: 0 cells.")

        values = []

        for selected_range in selected_cells:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                for col in range(selected_range.leftColumn(), selected_range.rightColumn() + 1):
                    item = self.src_table.item(row, col)
                    if item and item.text().strip().isdigit():
                        values.append(int(item.text()))

        return values

    def calculateFormula(self):
        try:
            formulavalue = self.formula_edit.text().strip()

            formulas = ['sum', 'avg', 'count', 'max', 'min']
            formula = formulavalue.split()[0]

            if formula not in formulas:
                raise ValueError(f"ERROR: {formula}")

            operands = self.selectedCells()

            if not operands:
                raise ValueError(f"ERROR: {operands}")

            if formula == 'sum':
                result = sum(operands)
            elif formula == 'avg':
                result = sum(operands) / len(operands)
            elif formula == 'count':
                result = len(operands)
            elif formula == 'max':
                result = max(operands)
            elif formula == 'min':
                result = min(operands)

            QMessageBox.information(self, "Formula", f"{formula} = {result}")

        except ValueError as valuerror:
            err_msg = f"ERROR: {str(valuerror)}"
            QErrorMessage().showMessage(err_msg)
        except Exception as exception:
            err_msg = f"ERROR: {str(exception)}"
            QErrorMessage().showMessage(err_msg)

    def deletecell(self):
        for item in self.src_table.selectedItems():
            item.setText('')


if __name__ == '__main__':
    app = QApplication(sys.argv)
    core = SRC_Workbook()
    app.setOrganizationName('berkaygediz')
    app.setApplicationName('SpanRC')
    app.setApplicationDisplayName('SpanRC')
    app.setApplicationVersion('1.2.0')
    core.show()
    sys.exit(app.exec_())
