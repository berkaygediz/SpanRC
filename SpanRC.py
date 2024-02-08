import os
import sys
import csv
import datetime
import time
import qtawesome as qta
import matplotlib.pyplot as plt
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
from modules.translations import *

default_values = {
    "row": 1,
    "column": 1,
    "rowspan": 1,
    "columnspan": 1,
    "text": "",
    "window_geometry": None,
    "is_saved": None,
    "file_name": None,
    "default_directory": None,
    "current_theme": "light",
    "current_language": "English",
    "file_type": "xsrc",
    "last_opened_file": None,
    "currentRow": 0,
    "currentColumn": 0,
    "rowCount": 50,
    "columnCount": 100,
    "windowState": None,
}


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


class SRC_About(QMainWindow):
    def __init__(self, parent=None):
        super(SRC_About, self).__init__(parent)
        self.setWindowFlags(Qt.Dialog)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setGeometry(
            QStyle.alignedRect(
                Qt.LeftToRight,
                Qt.AlignCenter,
                self.size(),
                QApplication.desktop().availableGeometry(),
            )
        )
        self.about_label = QLabel()
        self.about_label.setWordWrap(True)
        self.about_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.about_label.setTextFormat(Qt.RichText)
        self.about_label.setText(
            "<center>"
            "<b>SpanRC</b><br>"
            "A powerful spreadsheet application<br><br>"
            "Made by Berkay Gediz<br>"
            "Apache License 2.0</center>"
        )
        self.setCentralWidget(self.about_label)


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


class SRC_Workbook(QMainWindow):
    def __init__(self, parent=None):
        super().__init__(parent)
        starttime = datetime.datetime.now()
        settings = QSettings("berkaygediz", "SpanRC")
        self.setWindowIcon(QIcon("icon.png"))
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.src_thread = SRC_Threading()
        self.src_thread.update_signal.connect(self.SRC_updateStatistics)
        self.SRC_themePalette()
        self.undo_stack = QUndoStack(self)
        self.undo_stack.setUndoLimit(100)
        self.selected_file = None
        self.file_name = None
        self.is_saved = None
        self.default_directory = QDir().homePath()
        self.directory = self.default_directory
        self.SRC_setupDock()
        self.dock_widget.hide()
        self.status_bar = self.statusBar()
        self.src_table = QTableWidget(self)
        self.setCentralWidget(self.src_table)
        self.SRC_setupActions()
        self.SRC_setupToolbar()
        self.setPalette(self.light_theme)
        self.src_table.itemSelectionChanged.connect(self.src_thread.start)
        self.src_table.setCursor(Qt.CursorShape.SizeAllCursor)
        self.showMaximized()
        self.setFocus()
        QTimer.singleShot(50, self.SRC_restoreTheme)
        QTimer.singleShot(150, self.SRC_restoreState)
        if self.src_table.columnCount() == 0 and self.src_table.rowCount() == 0:
            self.src_table.setColumnCount(100)
            self.src_table.setRowCount(50)
            self.src_table.clearSpans()
            self.src_table.setItem(0, 0, QTableWidgetItem(""))

            self.SRC_updateTitle()
            self.setFocus()
            endtime = datetime.datetime.now()
            self.status_bar.showMessage(
                str((endtime - starttime).total_seconds()) + " ms", 2500
            )

    def closeEvent(self, event):
        settings = QSettings("berkaygediz", "SpanRC")

        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()

        if self.is_saved == False:
            reply = QMessageBox.question(
                self,
                "SpanRC",
                translations[settings.value("current_language")]["exit_message"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.SRC_saveState()
                event.accept()
            else:
                self.SRC_saveState()
                event.ignore()
        else:
            self.SRC_saveState()
            event.accept()

    def SRC_changeLanguage(self):
        language = self.language_combobox.currentText()
        settings = QSettings("berkaygediz", "SpanRC")
        settings.setValue("current_language", language)
        settings.sync()
        self.SRC_toolbarTranslate()
        self.SRC_updateStatistics()
        self.SRC_updateTitle()

    def SRC_updateTitle(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()
        file = (
            self.file_name
            if self.file_name
            else translations[settings.value("current_language")]["new_title"]
        )
        if self.is_saved == True:
            asterisk = ""
        else:
            asterisk = "*"
        self.setWindowTitle(f"{file}{asterisk} — SpanRC")

    def SRC_updateStatistics(self):
        settings = QSettings("berkaygediz", "SpanRC")
        row = self.src_table.rowCount()
        column = self.src_table.columnCount()
        selected_cell = (
            self.src_table.currentRow() + 1,
            self.src_table.currentColumn() + 1,
        )

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
        statistics += f"<td>{translations[settings.value('current_language')]['statistics_message1']}</td><td>{row}</td><td>{translations[settings.value('current_language')]['statistics_message2']}</td><td>{column}</td>"
        statistics += f"<td>{translations[settings.value('current_language')]['statistics_message2']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"
        if self.src_table.selectedRanges():
            statistics += f"<td>{translations[settings.value('current_language')]['statistics_message3']}</td><td>"
            for selected_range in self.src_table.selectedRanges():
                statistics += f"{selected_range.topRow() + 1}:{selected_range.leftColumn() + 1} - {selected_range.bottomRow() + 1}:{selected_range.rightColumn() + 1}</td>"
        else:
            statistics += f"<td>{translations[settings.value('current_language')]['statistics_message4']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"

        statistics += "</td><td id='sr-text'>SpanRC</td></tr></table></body></html>"
        self.statistics_label.setText(statistics)
        self.statusBar().addPermanentWidget(self.statistics_label)
        self.SRC_updateTitle()

    def SRC_saveState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        settings.setValue("window_geometry", self.saveGeometry())
        settings.setValue("default_directory", self.default_directory)
        self.file_name = settings.value("file_name", self.file_name)
        self.is_saved = settings.value("is_saved", self.is_saved)
        if self.selected_file:
            settings.setValue("last_opened_file", self.selected_file)
        settings.setValue(
            "current_theme", "dark" if self.palette() == self.dark_theme else "light"
        )
        settings.setValue("current_language", self.language_combobox.currentText())
        settings.sync()

    def SRC_restoreState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        self.geometry = settings.value("window_geometry")
        self.directory = settings.value("default_directory", self.default_directory)
        self.file_name = settings.value("file_name", self.file_name)
        self.is_saved = settings.value("is_saved", self.is_saved)
        self.language_combobox.setCurrentText(settings.value("current_language"))

        if self.geometry is not None:
            self.restoreGeometry(self.geometry)

        self.last_opened_file = settings.value("last_opened_file", "")
        if self.last_opened_file and os.path.exists(self.last_opened_file):
            self.loadFile(self.last_opened_file)

        self.src_table.setColumnCount(
            int(settings.value("columnCount", self.src_table.columnCount()))
        )
        self.src_table.setRowCount(
            int(settings.value("rowCount", self.src_table.rowCount()))
        )

        for row in range(self.src_table.rowCount()):
            for column in range(self.src_table.columnCount()):
                if settings.value(
                    f"row{row}column{column}rowspan", None
                ) and settings.value(f"row{row}column{column}columnspan", None):
                    self.src_table.setSpan(
                        row,
                        column,
                        int(settings.value(f"row{row}column{column}rowspan")),
                        int(settings.value(f"row{row}column{column}columnspan")),
                    )
        for row in range(self.src_table.rowCount()):
            for column in range(self.src_table.columnCount()):
                if settings.value(f"row{row}column{column}text", None):
                    self.src_table.setItem(
                        row,
                        column,
                        QTableWidgetItem(settings.value(f"row{row}column{column}text")),
                    )

        if self.file_name != None:
            self.src_table.resizeColumnsToContents()
            self.src_table.resizeRowsToContents()

        self.src_table.setCurrentCell(
            int(settings.value("currentRow", 0)),
            int(settings.value("currentColumn", 0)),
        )
        self.src_table.scrollToItem(
            self.src_table.item(
                int(settings.value("currentRow", 0)),
                int(settings.value("currentColumn", 0)),
            )
        )

        if self.file_name:
            self.is_saved = True
        else:
            self.is_saved = False

        self.SRC_restoreTheme(),
        self.SRC_updateTitle()

    def SRC_restoreTheme(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if settings.value("current_theme") == "dark":
            self.setPalette(self.dark_theme)
        else:
            self.setPalette(self.light_theme)
        self.SRC_toolbarTheme()

    def SRC_themePalette(self):
        self.light_theme = QPalette()
        self.dark_theme = QPalette()

        self.light_theme.setColor(QPalette.Window, QColor(89, 111, 183))
        self.light_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Base, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Text, QColor(0, 0, 0))
        self.light_theme.setColor(QPalette.Highlight, QColor(221, 216, 184))
        self.light_theme.setColor(QPalette.HighlightedText, QColor(0, 0, 0))

        self.dark_theme.setColor(QPalette.Window, QColor(58, 68, 93))
        self.dark_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Base, QColor(94, 87, 104))
        self.dark_theme.setColor(QPalette.Text, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Highlight, QColor(221, 216, 184))
        self.dark_theme.setColor(QPalette.HighlightedText, QColor(0, 0, 0))

    def SRC_themeAction(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.palette() == self.light_theme:
            self.setPalette(self.dark_theme)
            settings.setValue("current_theme", "dark")
        else:
            self.setPalette(self.light_theme)
            settings.setValue("current_theme", "light")
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

    def SRC_toolbarTranslate(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()
        self.language = settings.value("current_language")
        self.newAction.setText(translations[settings.value("current_language")]["new"])
        self.newAction.setStatusTip(
            translations[settings.value("current_language")]["new_title"]
        )
        self.openAction.setText(
            translations[settings.value("current_language")]["open"]
        )
        self.openAction.setStatusTip(
            translations[settings.value("current_language")]["open_title"]
        )
        self.saveAction.setText(
            translations[settings.value("current_language")]["save"]
        )
        self.saveAction.setStatusTip(
            translations[settings.value("current_language")]["save_title"]
        )
        self.saveasAction.setText(
            translations[settings.value("current_language")]["save_as"]
        )
        self.saveasAction.setStatusTip(
            translations[settings.value("current_language")]["save_as_title"]
        )
        self.printAction.setText(
            translations[settings.value("current_language")]["print"]
        )
        self.printAction.setStatusTip(
            translations[settings.value("current_language")]["print_title"]
        )
        self.exitAction.setText(
            translations[settings.value("current_language")]["exit"]
        )
        self.exitAction.setStatusTip(
            translations[settings.value("current_language")]["exit_title"]
        )
        self.deleteAction.setText(
            translations[settings.value("current_language")]["delete"]
        )
        self.deleteAction.setStatusTip(
            translations[settings.value("current_language")]["delete_title"]
        )
        self.aboutAction.setText(
            translations[settings.value("current_language")]["about"]
        )
        self.aboutAction.setStatusTip(
            translations[settings.value("current_language")]["about_title"]
        )
        self.undoAction.setText(
            translations[settings.value("current_language")]["undo"]
        )
        self.undoAction.setStatusTip(
            translations[settings.value("current_language")]["undo_title"]
        )
        self.redoAction.setText(
            translations[settings.value("current_language")]["redo"]
        )
        self.redoAction.setStatusTip(
            translations[settings.value("current_language")]["redo_title"]
        )
        self.darklightAction.setText(
            translations[settings.value("current_language")]["darklight"]
        )
        self.darklightAction.setStatusTip(
            translations[settings.value("current_language")]["darklight_message"]
        )
        self.help_label.setText(
            "<html><head><style>"
            "table {border-collapse: collapse; width: 100%;}"
            "th, td {text-align: left; padding: 8px;}"
            "tr:nth-child(even) {background-color: #f2f2f2;}"
            "tr:hover {background-color: #ddd;}"
            "th {background-color: #4CAF50; color: white;}"
            "</style></head><body>"
            "<table><tr><th>Shortcut</th><th>Function</th></tr>"
            f"<tr><td>Ctrl + O</td><td>{translations[settings.value('current_language')]['open_title']}</td></tr>"
            f"<tr><td>Ctrl + S</td><td>{translations[settings.value('current_language')]['save_title']}</td></tr>"
            f"<tr><td>Ctrl + N</td><td>{translations[settings.value('current_language')]['new_title']}</td></tr>"
            f"<tr><td>Ctrl + Shift + S</td><td>{translations[settings.value('current_language')]['save_as_title']}</td></tr>"
            f"<tr><td>Ctrl + P</td><td>{translations[settings.value('current_language')]['print_title']}</td></tr>"
            f"<tr><td>Ctrl + Q</td><td>{translations[settings.value('current_language')]['exit_title']}</td></tr>"
            f"<tr><td>Ctrl + D</td><td>{translations[settings.value('current_language')]['delete_title']}</td></tr>"
            f"<tr><td>Ctrl + A</td><td>{translations[settings.value('current_language')]['about_title']}</td></tr>"
            f"<tr><td>Ctrl + Z</td><td>{translations[settings.value('current_language')]['undo_title']}</td></tr>"
            f"<tr><td>Ctrl + Y</td><td>{translations[settings.value('current_language')]['redo_title']}</td></tr>"
            f"<tr><td>Ctrl + L</td><td>{translations[settings.value('current_language')]['darklight_message']}</td></tr>"
            "</table></body></html>"
        )

    def SRC_setupDock(self):
        settings = QSettings("berkaygediz", "SpanRC")
        self.dock_widget = QDockWidget(
            translations[settings.value("current_language")]["help"], self
        )
        self.statistics_label = QLabel()
        self.help_label = QLabel()
        self.help_label.setWordWrap(True)
        self.help_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.help_label.setTextFormat(Qt.RichText)
        self.help_label.setText(
            "<html><head><style>"
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
            f"<tr><td>Ctrl + Z</td><td>{translations[settings.value('current_language')]['undo_title']}</td></tr>"
            f"<tr><td>Ctrl + Y</td><td>{translations[settings.value('current_language')]['redo_title']}</td></tr>"
            f"<tr><td>Ctrl + L</td><td>{translations[settings.value('current_language')]['darklight_message']}</td></tr>"
            "</table></body></html>"
        )
        self.dock_widget.setWidget(self.help_label)
        self.dock_widget.setObjectName("Help")
        self.dock_widget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

    def SRC_toolbarLabel(self, toolbar, text):
        label = QLabel(f"<b>{text}</b>")
        toolbar.addWidget(label)

    def SRC_createAction(self, text, status_tip, function, shortcut=None, icon=None):
        action = QAction(text, self)
        action.setStatusTip(status_tip)
        action.triggered.connect(function)
        if shortcut:
            action.setShortcut(shortcut)
        if icon:
            action.setIcon(QIcon(icon))
        return action

    def SRC_setupActions(self):
        settings = QSettings("berkaygediz", "SpanRC")
        current_language = settings.value("current_language")
        icon_theme = "white"
        if self.palette() == self.light_theme:
            icon_theme = "black"
        actionicon = qta.icon("fa5s.file", color=icon_theme)
        self.newAction = self.SRC_createAction(
            translations[current_language]["new"],
            translations[current_language]["new_title"],
            self.new,
            QKeySequence.New,
            actionicon,
        )
        actionicon = qta.icon("fa5s.folder-open", color=icon_theme)
        self.openAction = self.SRC_createAction(
            translations[current_language]["open"],
            translations[current_language]["open_title"],
            self.open,
            QKeySequence.Open,
            actionicon,
        )
        actionicon = qta.icon("fa5s.save", color=icon_theme)
        self.saveAction = self.SRC_createAction(
            translations[current_language]["save"],
            translations[current_language]["save_title"],
            self.save,
            QKeySequence.Save,
            actionicon,
        )
        actionicon = qta.icon("fa5.save", color=icon_theme)
        self.saveasAction = self.SRC_createAction(
            translations[current_language]["save_as"],
            translations[current_language]["save_as_title"],
            self.saveAs,
            QKeySequence.SaveAs,
            actionicon,
        )
        actionicon = qta.icon("fa5s.print", color=icon_theme)
        self.printAction = self.SRC_createAction(
            translations[current_language]["print"],
            translations[current_language]["print_title"],
            self.print,
            QKeySequence.Print,
            actionicon,
        )
        actionicon = qta.icon("fa5s.sign-out-alt", color=icon_theme)
        self.exitAction = self.SRC_createAction(
            translations[current_language]["exit"],
            translations[current_language]["exit_message"],
            self.close,
            QKeySequence.Quit,
            actionicon,
        )
        actionicon = qta.icon("fa5s.trash-alt", color=icon_theme)
        self.deleteAction = self.SRC_createAction(
            translations[current_language]["delete"],
            translations[current_language]["delete_title"],
            self.deletecell,
            QKeySequence.Delete,
            actionicon,
        )
        actionicon = qta.icon("fa5s.plus", color=icon_theme)
        self.addrowAction = self.SRC_createAction(
            translations[current_language]["add_row"],
            translations[current_language]["add_row_title"],
            self.addrow,
            QKeySequence("Ctrl+Shift+R"),
            actionicon,
        )
        actionicon = qta.icon("fa5s.plus", color=icon_theme)
        self.addcolumnAction = self.SRC_createAction(
            translations[current_language]["add_column"],
            translations[current_language]["add_column_title"],
            self.addcolumn,
            QKeySequence("Ctrl+Shift+C"),
            actionicon,
        )
        actionicon = qta.icon("fa5s.plus", color=icon_theme)
        self.addrowaboveAction = self.SRC_createAction(
            translations[current_language]["add_row_above"],
            translations[current_language]["add_row_above_title"],
            self.addrowabove,
            QKeySequence("Ctrl+Shift+T"),
            actionicon,
        )
        actionicon = qta.icon("fa5s.plus", color=icon_theme)
        self.addcolumnleftAction = self.SRC_createAction(
            translations[current_language]["add_column_left"],
            translations[current_language]["add_column_left_title"],
            self.addcolumnleft,
            QKeySequence("Ctrl+Shift+L"),
            actionicon,
        )
        actionicon = qta.icon("fa5s.question-circle", color=icon_theme)
        self.hide_dock_widget_action = self.SRC_createAction(
            translations[current_language]["help"],
            translations[current_language]["help"],
            self.SRC_toggleDock,
            QKeySequence("Ctrl+H"),
            actionicon,
        )
        actionicon = qta.icon("fa5s.info-circle", color=icon_theme)
        self.aboutAction = self.SRC_createAction(
            translations[current_language]["about"],
            translations[current_language]["about_title"],
            self.showAbout,
            QKeySequence.HelpContents,
            actionicon,
        )
        actionicon = qta.icon("fa5s.undo", color=icon_theme)
        self.undoAction = self.SRC_createAction(
            translations[current_language]["undo"],
            translations[current_language]["undo_title"],
            self.undo_stack.undo,
            QKeySequence.Undo,
            actionicon,
        )
        actionicon = qta.icon("fa5s.redo", color=icon_theme)
        self.redoAction = self.SRC_createAction(
            translations[current_language]["redo"],
            translations[current_language]["redo_title"],
            self.undo_stack.redo,
            QKeySequence.Redo,
            actionicon,
        )
        actionicon = qta.icon("fa5s.moon", color=icon_theme)
        self.darklightAction = self.SRC_createAction(
            translations[current_language]["darklight"],
            translations[current_language]["darklight_message"],
            self.SRC_themeAction,
            QKeySequence("Ctrl+D"),
            actionicon,
        )

    def SRC_setupToolbar(self):
        settings = QSettings("berkaygediz", "SpanRC")

        self.file_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["file"]
        )
        self.file_toolbar.setObjectName("File")
        self.SRC_toolbarLabel(
            self.file_toolbar,
            translations[settings.value("current_language")]["file"] + ": ",
        )
        self.file_toolbar.addActions(
            [
                self.newAction,
                self.openAction,
                self.saveAction,
                self.saveasAction,
                self.printAction,
                self.exitAction,
            ]
        )

        self.edit_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["edit"]
        )
        self.edit_toolbar.setObjectName("Edit")
        self.SRC_toolbarLabel(
            self.edit_toolbar,
            translations[settings.value("current_language")]["edit"] + ": ",
        )
        self.edit_toolbar.addActions(
            [
                self.undoAction,
                self.redoAction,
                self.deleteAction,
                self.addrowAction,
                self.addcolumnAction,
                self.addrowaboveAction,
                self.addcolumnleftAction,
            ]
        )

        self.interface_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["interface"]
        )
        self.interface_toolbar.setObjectName("Interface")
        self.SRC_toolbarLabel(
            self.interface_toolbar,
            translations[settings.value("current_language")]["interface"] + ": ",
        )
        self.interface_toolbar.addAction(self.darklightAction)
        self.interface_toolbar.addAction(self.hide_dock_widget_action)
        self.interface_toolbar.addAction(self.aboutAction)
        self.language_combobox = QComboBox(self)
        self.language_combobox.addItems(
            ["English", "Türkçe", "Azərbaycanca", "Deutsch", "Español"]
        )
        self.language_combobox.currentIndexChanged.connect(self.SRC_changeLanguage)
        self.interface_toolbar.addWidget(self.language_combobox)
        self.addToolBarBreak()
        self.formula_toolbar = self.addToolBar(
            translations[settings.value("current_language")]["formula"]
        )
        self.formula_toolbar.setObjectName("Formula")
        self.SRC_toolbarLabel(
            self.formula_toolbar,
            translations[settings.value("current_language")]["formula"] + ": ",
        )
        self.formula_edit = QLineEdit()
        self.formula_edit.setPlaceholderText(
            translations[settings.value("current_language")]["formula"]
        )
        self.formula_edit.returnPressed.connect(self.calculateFormula)
        self.formula_button = QPushButton(
            translations[settings.value("current_language")]["compute"]
        )
        self.formula_button.setStyleSheet(
            "background-color: #A72461; color: #FFFFFF; font-weight: bold; padding: 10px; border-radius: 10px; border: 1px solid #000000; margin-left: 10px;"
        )
        self.formula_button.setCursor(Qt.PointingHandCursor)
        self.formula_button.clicked.connect(self.calculateFormula)
        self.formula_toolbar.addWidget(self.formula_edit)
        self.formula_toolbar.addWidget(self.formula_button)

        if self.palette() == self.light_theme:
            self.formula_toolbar.setStyleSheet(
                "background-color: #FFFFFF; color: #000000; font-weight: bold; padding: 10px; border-radius: 10px;"
            )
            self.formula_edit.setStyleSheet(
                "background-color: #FFFFFF; color: #000000; font-weight: bold; padding: 10px; border-radius: 10px; border: 3px solid #000000;"
            )
        else:
            self.formula_toolbar.setStyleSheet(
                "background-color: #000000; color: #FFFFFF; font-weight: bold; padding: 10px; border-radius: 10px;"
            )
            self.formula_edit.setStyleSheet(
                "background-color: #000000; color: #FFFFFF; font-weight: bold; padding: 10px; border-radius: 10px; border: 1px solid #FFFFFF;"
            )

    def SRC_createDockwidget(self):
        widget = QWidget()
        layout = QVBoxLayout()
        widget.setLayout(layout)
        self.statistics = QLabel()
        self.statistics.setText("Statistics")
        layout.addWidget(self.statistics)
        return widget

    def SRC_toggleDock(self):
        if self.dock_widget.isHidden():
            self.dock_widget.show()
        else:
            self.dock_widget.hide()

    def new(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.is_saved == False:
            reply = QMessageBox.question(
                self,
                "SpanRC",
                translations[settings.value("current_language")]["new_title"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.SRC_saveState()
                self.src_table.clearContents()
                self.src_table.setRowCount(50)
                self.src_table.setColumnCount(100)
                self.is_saved = False
                self.file_name = None
                self.setWindowTitle(
                    translations[settings.value("current_language")]["new_title"]
                    + " — SpanRC"
                )
                self.directory = self.default_directory
                return True
            else:
                return False
        else:
            self.src_table.clearContents()
            self.src_table.setRowCount(50)
            self.src_table.setColumnCount(100)
            self.is_saved = False
            self.file_name = None
            self.setWindowTitle(
                translations[settings.value("current_language")]["new_title"]
                + " — SpanRC"
            )
            self.directory = self.default_directory
            return True

    def open(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.is_saved is False:
            reply = QMessageBox.question(
                self,
                "SpanRC",
                translations[settings.value("current_language")]["open_message"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

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
            self,
            translations[settings.value("current_language")]["open_title"]
            + " — SpanRC",
            self.directory,
            file_filter,
            options=options,
        )

        if selected_file:
            self.loadFile(selected_file)
            return True
        return False

    def loadFile(self, file_path):
        self.selected_file = file_path
        self.file_name = os.path.basename(self.selected_file)
        self.directory = os.path.dirname(self.selected_file)
        self.setWindowTitle(self.file_name)

        if file_path.endswith(".xsrc") or file_path.endswith(".csv"):
            self.loadTable(file_path)

        self.is_saved = True
        self.file_name = os.path.basename(self.selected_file)
        self.directory = os.path.dirname(self.selected_file)
        self.SRC_updateTitle()
        return True

    def loadTable(self, file_path):
        with open(file_path, "r") as file:
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
            self,
            translations[settings.value("current_language")]["save_as_title"]
            + " — SpanRC",
            self.directory,
            file_filter,
            options=options,
        )
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
            with open(self.selected_file, "w", newline="") as file:
                writer = csv.writer(file)
                for rowid in range(self.src_table.rowCount()):
                    row = []
                    for colid in range(self.src_table.columnCount()):
                        item = self.src_table.item(rowid, colid)
                        if item is not None:
                            row.append(item.text())
                        else:
                            row.append("")
                    writer.writerow(row)

        self.file_name = os.path.basename(self.selected_file)
        self.directory = os.path.dirname(self.selected_file)
        self.status_bar.showMessage("Saved.", 2000)
        self.is_saved = True
        self.SRC_updateTitle()

    def print(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOrientation(QPrinter.Landscape)
        printer.setPageMargins(0, 0, 0, 0, QPrinter.Millimeter)
        printer.setFullPage(True)
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
                for col in range(
                    selected_range.leftColumn(), selected_range.rightColumn() + 1
                ):
                    item = self.src_table.item(row, col)
                    if item and item.text().strip().isdigit():
                        values.append(int(item.text()))

        return values

    def calculateFormula(self):
        try:
            formulavalue = self.formula_edit.text().strip()

            formulas = [
                "sum",
                "avg",
                "count",
                "max",
                "similargraph",
                "pointgraph",
                "bargraph",
                "piegraph",
                "histogram",
            ]
            formula = formulavalue.split()[0]

            if formula not in formulas:
                raise ValueError(f"ERROR: {formula}")

            operands = self.selectedCells()

            if not operands:
                raise ValueError(f"ERROR: {operands}")

            if formula == "sum":
                result = sum(operands)
            elif formula == "avg":
                result = sum(operands) / len(operands)
            elif formula == "count":
                result = len(operands)
            elif formula == "max":
                result = max(operands)
            elif formula == "min":
                result = min(operands)
            elif formula == "similargraph":
                pltgraph = plt.figure()
                pltgraph.suptitle("Similar Graph")
                plt.plot(operands)
                plt.xlabel("X")
                plt.ylabel("Y")
                plt.grid(True)
                plt.show()
                result = "Graph"
            elif formula == "pointgraph":
                pltgraph = plt.figure()
                pltgraph.suptitle("Point Graph")
                plt.plot(operands, "o")
                plt.xlabel("X")
                plt.ylabel("Y")
                plt.grid(True)
                plt.show()
                result = "Graph"
            elif formula == "bargraph":
                pltgraph = plt.figure()
                pltgraph.suptitle("Bar Graph")
                plt.bar(range(len(operands)), operands)
                plt.xlabel("X")
                plt.ylabel("Y")
                plt.grid(True)
                plt.show()
                result = "Graph"
            elif formula == "piegraph":
                pltgraph = plt.figure()
                pltgraph.suptitle("Pie Graph")
                plt.pie(operands)
                plt.xlabel("X")
                plt.ylabel("Y")
                plt.grid(True)
                plt.show()
                result = "Graph"
            elif formula == "histogram":
                pltgraph = plt.figure()
                pltgraph.suptitle("Histogram")
                plt.hist(operands)
                plt.xlabel("X")
                plt.ylabel("Y")
                plt.grid(True)
                plt.show()
                result = "Graph"

            if result == "Graph":
                pass
            else:
                QMessageBox.information(self, "Formula", f"{formula} : {result}")

        except ValueError as valueerror:
            QMessageBox.critical(self, "Formula", str(valueerror))
        except Exception as exception:
            QMessageBox.critical(self, "Formula", str(exception))

    def deletecell(self):
        for item in self.src_table.selectedItems():
            item.setText("")

    def addrow(self):
        self.src_table.insertRow(self.src_table.rowCount())

    def addcolumn(self):
        self.src_table.insertColumn(self.src_table.columnCount())

    def addrowabove(self):
        self.src_table.insertRow(self.src_table.currentRow())

    def addcolumnleft(self):
        self.src_table.insertColumn(self.src_table.currentColumn())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setOrganizationName("berkaygediz")
    app.setApplicationName("SpanRC")
    app.setApplicationDisplayName("SpanRC")
    app.setApplicationVersion("1.3.18")
    wb = SRC_Workbook()
    wb.show()
    sys.exit(app.exec_())
