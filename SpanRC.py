import csv
import datetime
import os
import sys
import time

import matplotlib.pyplot as plt
import psutil
import qtawesome as qta
from PySide6.QtCore import *
from PySide6.QtGui import *
from PySide6.QtOpenGL import *
from PySide6.QtOpenGLWidgets import *
from PySide6.QtPrintSupport import *
from PySide6.QtWidgets import *

from modules.translations import *

fallbackValues = {
    "currentRow": 0,
    "currentColumn": 0,
    "rowCount": 50,
    "columnCount": 100,
    "windowScale": None,
    "isSaved": None,
    "fileName": None,
    "defaultDirectory": None,
    "appTheme": "light",
    "appLanguage": "English",
    "adaptiveResponse": 1,
}


class SRC_Threading(QThread):
    update_signal = Signal()

    def __init__(self, adaptiveResponse, parent=None):
        super(SRC_Threading, self).__init__(parent)
        self.adaptiveResponse = float(adaptiveResponse)
        self.running = False

    def run(self):
        if not self.running:
            self.running = True
            time.sleep(0.15 * self.adaptiveResponse)
            self.update_signal.emit()
            self.running = False


class SRC_About(QMainWindow):
    def __init__(self, parent=None):
        super(SRC_About, self).__init__(parent)
        self.setWindowFlags(Qt.Dialog)
        self.setWindowIcon(QIcon("spanrc_icon.ico"))
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setGeometry(
            QStyle.alignedRect(
                Qt.LeftToRight,
                Qt.AlignCenter,
                self.size(),
                QApplication.primaryScreen().availableGeometry(),
            )
        )
        self.about_label = QLabel()
        self.about_label.setWordWrap(True)
        self.about_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.about_label.setTextFormat(Qt.RichText)
        self.about_label.setText(
            "<center>"
            f"<b>{app.applicationDisplayName()}</b><br><br>"
            "A powerful spreadsheet application<br>"
            "Made by Berkay Gediz<br><br>"
            "GNU General Public License v3.0<br>GNU LESSER GENERAL PUBLIC LICENSE v3.0<br>Mozilla Public License Version 2.0<br><br><br>"
            "OpenGL: <b>ON</b></center>"
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
        if settings.value("appLanguage") == None:
            settings.setValue("appLanguage", "English")
            settings.sync()
        if settings.value("adaptiveResponse") == None:
            settings.setValue("adaptiveResponse", 1)
            settings.sync()
        self.setWindowIcon(QIcon("spanrc_icon.ico"))
        self.setWindowModality(Qt.ApplicationModal)

        centralWidget = QOpenGLWidget(self)

        layout = QVBoxLayout(centralWidget)
        self.hardwareAcceleration = QOpenGLWidget()
        layout.addWidget(self.hardwareAcceleration)
        self.setCentralWidget(centralWidget)

        self.src_thread = SRC_Threading(
            adaptiveResponse=settings.value("adaptiveResponse")
        )
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

        self.adaptiveResponse = settings.value("adaptiveResponse")

        QTimer.singleShot(50 * self.adaptiveResponse, self.SRC_restoreTheme)
        QTimer.singleShot(150 * self.adaptiveResponse, self.SRC_restoreState)
        if self.src_table.columnCount() == 0 and self.src_table.rowCount() == 0:
            self.src_table.setColumnCount(100)
            self.src_table.setRowCount(50)
            self.src_table.clearSpans()
            self.src_table.setItem(0, 0, QTableWidgetItem(""))

            self.SRC_updateTitle()
            self.setFocus()
            endtime = datetime.datetime.now()
            self.status_bar.showMessage(
                str((endtime - starttime).total_seconds()) + " ms",
                2500 * self.adaptiveResponse,
            )

    def closeEvent(self, event):
        settings = QSettings("berkaygediz", "SpanRC")

        if settings.value("appLanguage") == None:
            settings.setValue("appLanguage", "English")
            settings.sync()

        if self.is_saved == False:
            reply = QMessageBox.question(
                self,
                app.applicationDisplayName(),
                translations[settings.value("appLanguage")]["exit_message"],
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
        settings.setValue("appLanguage", language)
        settings.sync()
        self.SRC_toolbarTranslate()
        self.SRC_updateStatistics()
        self.SRC_updateTitle()

    def SRC_updateTitle(self):
        settings = QSettings("berkaygediz", "SpanRC")
        file = (
            self.file_name
            if self.file_name
            else translations[settings.value("appLanguage")]["new_title"]
        )
        if self.is_saved == True:
            asterisk = ""
        else:
            asterisk = "*"
        self.setWindowTitle(f"{file}{asterisk} — {app.applicationDisplayName()}")

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
            f"<th>{translations[settings.value('appLanguage')]['statistics_title']}</th>"
        )
        statistics += f"<td>{translations[settings.value('appLanguage')]['statistics_message1']}</td><td>{row}</td><td>{translations[settings.value('appLanguage')]['statistics_message2']}</td><td>{column}</td>"
        statistics += f"<td>{translations[settings.value('appLanguage')]['statistics_message2']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"
        if self.src_table.selectedRanges():
            statistics += f"<td>{translations[settings.value('appLanguage')]['statistics_message3']}</td><td>"
            for selected_range in self.src_table.selectedRanges():
                statistics += f"{selected_range.topRow() + 1}:{selected_range.leftColumn() + 1} - {selected_range.bottomRow() + 1}:{selected_range.rightColumn() + 1}</td>"
        else:
            statistics += f"<td>{translations[settings.value('appLanguage')]['statistics_message4']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"

        statistics += f"</td><td id='sr-text'>{app.applicationDisplayName()}</td></tr></table></body></html>"
        self.statistics_label.setText(statistics)
        self.statusBar().addPermanentWidget(self.statistics_label)
        self.SRC_updateTitle()

    def SRC_saveState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        settings.setValue("windowScale", self.saveGeometry())
        settings.setValue("defaultDirectory", self.default_directory)
        self.file_name = settings.value("fileName", self.file_name)
        self.is_saved = settings.value("isSaved", self.is_saved)
        if self.selected_file:
            settings.setValue("fileName", self.selected_file)
        settings.setValue(
            "appTheme", "dark" if self.palette() == self.dark_theme else "light"
        )
        settings.setValue("appLanguage", self.language_combobox.currentText())
        settings.sync()

    def SRC_restoreState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        self.geometry = settings.value("windowScale")
        self.directory = settings.value("defaultDirectory", self.default_directory)
        self.file_name = settings.value("fileName", self.file_name)
        self.is_saved = settings.value("isSaved", self.is_saved)
        self.language_combobox.setCurrentText(settings.value("appLanguage"))

        if self.geometry is not None:
            self.restoreGeometry(self.geometry)

        self.last_opened_file = settings.value("fileName", "")
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
        if settings.value("appTheme") == "dark":
            self.setPalette(self.dark_theme)
        else:
            self.setPalette(self.light_theme)
        self.SRC_toolbarTheme()

    def SRC_themePalette(self):
        self.light_theme = QPalette()
        self.dark_theme = QPalette()

        self.light_theme.setColor(QPalette.Window, QColor(3, 65, 135))
        self.light_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Base, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Text, QColor(0, 0, 0))
        self.light_theme.setColor(QPalette.Highlight, QColor(105, 117, 156))
        self.light_theme.setColor(QPalette.ButtonText, QColor(0, 0, 0))

        self.dark_theme.setColor(QPalette.Window, QColor(35, 39, 52))
        self.dark_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Base, QColor(80, 85, 122))
        self.dark_theme.setColor(QPalette.Text, QColor(0, 0, 0))
        self.dark_theme.setColor(QPalette.Highlight, QColor(105, 117, 156))
        self.dark_theme.setColor(QPalette.ButtonText, QColor(0, 0, 0))

    def SRC_themeAction(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.palette() == self.light_theme:
            self.setPalette(self.dark_theme)
            settings.setValue("appTheme", "dark")
        else:
            self.setPalette(self.light_theme)
            settings.setValue("appTheme", "light")
        self.SRC_toolbarTheme()

    def SRC_toolbarTheme(self):
        palette = self.palette()
        if palette == self.light_theme:
            text_color = QColor(255, 255, 255)
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
        if settings.value("appLanguage") == None:
            settings.setValue("appLanguage", "English")
            settings.sync()
        self.language = settings.value("appLanguage")
        self.newAction.setText(translations[settings.value("appLanguage")]["new"])
        self.newAction.setStatusTip(
            translations[settings.value("appLanguage")]["new_title"]
        )
        self.openAction.setText(translations[settings.value("appLanguage")]["open"])
        self.openAction.setStatusTip(
            translations[settings.value("appLanguage")]["open_title"]
        )
        self.saveAction.setText(translations[settings.value("appLanguage")]["save"])
        self.saveAction.setStatusTip(
            translations[settings.value("appLanguage")]["save_title"]
        )
        self.saveasAction.setText(
            translations[settings.value("appLanguage")]["save_as"]
        )
        self.saveasAction.setStatusTip(
            translations[settings.value("appLanguage")]["save_as_title"]
        )
        self.printAction.setText(translations[settings.value("appLanguage")]["print"])
        self.printAction.setStatusTip(
            translations[settings.value("appLanguage")]["print_title"]
        )
        self.exitAction.setText(translations[settings.value("appLanguage")]["exit"])
        self.exitAction.setStatusTip(
            translations[settings.value("appLanguage")]["exit_title"]
        )
        self.deleteAction.setText(translations[settings.value("appLanguage")]["delete"])
        self.deleteAction.setStatusTip(
            translations[settings.value("appLanguage")]["delete_title"]
        )
        self.aboutAction.setText(translations[settings.value("appLanguage")]["about"])
        self.aboutAction.setStatusTip(
            translations[settings.value("appLanguage")]["about_title"]
        )
        self.undoAction.setText(translations[settings.value("appLanguage")]["undo"])
        self.undoAction.setStatusTip(
            translations[settings.value("appLanguage")]["undo_title"]
        )
        self.redoAction.setText(translations[settings.value("appLanguage")]["redo"])
        self.redoAction.setStatusTip(
            translations[settings.value("appLanguage")]["redo_title"]
        )
        self.darklightAction.setText(
            translations[settings.value("appLanguage")]["darklight"]
        )
        self.darklightAction.setStatusTip(
            translations[settings.value("appLanguage")]["darklight_message"]
        )
        self.helpText.setText(
            "<html><head><style>"
            "table {border-collapse: collapse; width: 100%;}"
            "th, td {text-align: left; padding: 8px;}"
            "tr:nth-child(even) {background-color: #f2f2f2;}"
            "tr:hover {background-color: #ddd;}"
            "th {background-color: #4CAF50; color: white;}"
            "</style></head><body>"
            "<table><tr><th>Shortcut</th><th>Function</th></tr>"
            f"<tr><td>Ctrl + O</td><td>{translations[settings.value('appLanguage')]['open_title']}</td></tr>"
            f"<tr><td>Ctrl + S</td><td>{translations[settings.value('appLanguage')]['save_title']}</td></tr>"
            f"<tr><td>Ctrl + N</td><td>{translations[settings.value('appLanguage')]['new_title']}</td></tr>"
            f"<tr><td>Ctrl + Shift + S</td><td>{translations[settings.value('appLanguage')]['save_as_title']}</td></tr>"
            f"<tr><td>Ctrl + P</td><td>{translations[settings.value('appLanguage')]['print_title']}</td></tr>"
            f"<tr><td>Ctrl + Q</td><td>{translations[settings.value('appLanguage')]['exit_title']}</td></tr>"
            f"<tr><td>Ctrl + D</td><td>{translations[settings.value('appLanguage')]['delete_title']}</td></tr>"
            f"<tr><td>Ctrl + A</td><td>{translations[settings.value('appLanguage')]['about_title']}</td></tr>"
            f"<tr><td>Ctrl + Z</td><td>{translations[settings.value('appLanguage')]['undo_title']}</td></tr>"
            f"<tr><td>Ctrl + Y</td><td>{translations[settings.value('appLanguage')]['redo_title']}</td></tr>"
            f"<tr><td>Ctrl + L</td><td>{translations[settings.value('appLanguage')]['darklight_message']}</td></tr>"
            "</table></body></html>"
        )

    def SRC_setupDock(self):
        settings = QSettings("berkaygediz", "SpanRC")
        settings.sync()
        self.dock_widget = QDockWidget(
            translations[settings.value("appLanguage")]["help"] + " && Graph Log", self
        )
        self.dock_widget.setObjectName("Help & Graph Log")
        self.dock_widget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

        self.scrollableArea = QScrollArea()
        self.GraphLog_QVBox = QVBoxLayout()
        self.statistics_label = QLabel()
        self.helpText = QLabel()
        self.helpText.setWordWrap(True)
        self.helpText.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.helpText.setTextFormat(Qt.RichText)
        self.helpText.setText(
            "<html><head><style>"
            "table {border-collapse: collapse; width: 100%;}"
            "th, td {text-align: left; padding: 8px;}"
            "tr:nth-child(even) {background-color: #f2f2f2;}"
            "tr:hover {background-color: #ddd;}"
            "th {background-color: #4CAF50; color: white;}"
            "</style></head><body>"
            "<table><tr><th>Shortcut</th><th>Function</th></tr>"
            f"<tr><td>Ctrl + N</td><td>{translations[settings.value('appLanguage')]['new_title']}</td></tr>"
            f"<tr><td>Ctrl + O</td><td>{translations[settings.value('appLanguage')]['open_title']}</td></tr>"
            f"<tr><td>Ctrl + S</td><td>{translations[settings.value('appLanguage')]['save_title']}</td></tr>"
            f"<tr><td>Ctrl + Shift + S</td><td>{translations[settings.value('appLanguage')]['save_as_title']}</td></tr>"
            f"<tr><td>Ctrl + P</td><td>{translations[settings.value('appLanguage')]['print_title']}</td></tr>"
            f"<tr><td>Ctrl + Q</td><td>{translations[settings.value('appLanguage')]['exit_title']}</td></tr>"
            f"<tr><td>Ctrl + D</td><td>{translations[settings.value('appLanguage')]['delete_title']}</td></tr>"
            f"<tr><td>Ctrl + A</td><td>{translations[settings.value('appLanguage')]['about_title']}</td></tr>"
            f"<tr><td>Ctrl + Z</td><td>{translations[settings.value('appLanguage')]['undo_title']}</td></tr>"
            f"<tr><td>Ctrl + Y</td><td>{translations[settings.value('appLanguage')]['redo_title']}</td></tr>"
            f"<tr><td>Ctrl + L</td><td>{translations[settings.value('appLanguage')]['darklight_message']}</td></tr>"
            "</table><p>NOTE: <b>Graph Log</b> support planned.</p></body></html>"
        )
        self.GraphLog_QVBox.addWidget(self.helpText)

        self.dock_widget.setObjectName("Help")
        self.dock_widget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

        self.dock_widget.setWidget(self.scrollableArea)

        self.dock_widget.setFeatures(
            QDockWidget.NoDockWidgetFeatures | QDockWidget.DockWidgetClosable
        )
        self.dock_widget.setWidget(self.scrollableArea)
        self.scrollableArea.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scrollableArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scrollableArea.setWidgetResizable(True)
        scroll_contents = QWidget()
        scroll_contents.setLayout(self.GraphLog_QVBox)
        self.scrollableArea.setWidget(scroll_contents)

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
            action.setIcon(
                QIcon("")
            )  # qtawesome is library based on qt5 --> icon = Qt5.QtGui.QIcon
        return action

    def SRC_setupActions(self):
        settings = QSettings("berkaygediz", "SpanRC")
        current_language = settings.value("appLanguage")
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
        icon_theme = "white"

        self.file_toolbar = self.addToolBar(
            translations[settings.value("appLanguage")]["file"]
        )
        self.file_toolbar.setObjectName("File")
        self.SRC_toolbarLabel(
            self.file_toolbar,
            translations[settings.value("appLanguage")]["file"] + ": ",
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
            translations[settings.value("appLanguage")]["edit"]
        )
        self.edit_toolbar.setObjectName("Edit")
        self.SRC_toolbarLabel(
            self.edit_toolbar,
            translations[settings.value("appLanguage")]["edit"] + ": ",
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
            translations[settings.value("appLanguage")]["interface"]
        )
        self.interface_toolbar.setObjectName("Interface")
        self.SRC_toolbarLabel(
            self.interface_toolbar,
            translations[settings.value("appLanguage")]["interface"] + ": ",
        )
        actionicon = qta.icon("fa5b.affiliatetheme", color="white")
        self.theme_action = self.SRC_createAction(
            translations[settings.value("appLanguage")]["darklight"],
            translations[settings.value("appLanguage")]["darklight_message"],
            self.SRC_themeAction,
            QKeySequence("Ctrl+Shift+T"),
            actionicon,
        )
        self.theme_action.setCheckable(True)
        self.theme_action.setChecked(settings.value("appTheme") == "dark")

        self.interface_toolbar.addAction(self.theme_action)
        actionicon = qta.icon("fa5s.leaf", color=icon_theme)
        self.powersaveraction = QAction("Power Saver", self, checkable=True)
        # self.powersaveraction.setIcon(QIcon(actionicon))
        self.powersaveraction.setStatusTip(
            "Experimental power saver function. Restart required."
        )
        self.powersaveraction.toggled.connect(self.SRC_hybridSaver)

        self.interface_toolbar.addAction(self.powersaveraction)
        response_exponential = settings.value(
            "adaptiveResponse", fallbackValues["adaptiveResponse"]
        )
        self.powersaveraction.setChecked(response_exponential == 12)
        self.interface_toolbar.addAction(self.powersaveraction)
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
            translations[settings.value("appLanguage")]["formula"]
        )
        self.formula_toolbar.setObjectName("Formula")
        self.SRC_toolbarLabel(
            self.formula_toolbar,
            translations[settings.value("appLanguage")]["formula"] + ": ",
        )
        self.formula_edit = QLineEdit()
        self.formula_edit.setPlaceholderText(
            translations[settings.value("appLanguage")]["formula"]
        )
        self.formula_edit.returnPressed.connect(self.calculateFormula)
        self.formula_button = QPushButton(
            translations[settings.value("appLanguage")]["compute"]
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

    def SRC_hybridSaver(self, checked):
        settings = QSettings("berkaygediz", "SpanRC")
        if checked:
            battery = psutil.sensors_battery()
            if battery:
                if battery.percent <= 35 and not battery.power_plugged:
                    # Ultra
                    self.adaptiveResponse = 12
                else:
                    # Standard
                    self.adaptiveResponse = 6
            else:
                # Global Standard
                self.adaptiveResponse = 3
        else:
            self.adaptiveResponse = fallbackValues["adaptiveResponse"]

        settings.setValue("adaptiveResponse", self.adaptiveResponse)
        settings.sync()

    def new(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.is_saved == True:
            self.src_table.clearContents()
            self.src_table.setRowCount(50)
            self.src_table.setColumnCount(100)
            self.directory = self.default_directory
            self.file_name = None
            self.is_saved = False
            self.SRC_updateTitle()
        else:
            reply = QMessageBox.question(
                self,
                app.applicationDisplayName(),
                translations[settings.value("appLanguage")]["new_title"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.src_table.clearContents()
                self.src_table.setRowCount(50)
                self.src_table.setColumnCount(100)
                self.setWindowTitle(
                    translations[settings.value("appLanguage")]["new_title"]
                    + f" — {app.applicationDisplayName()}"
                )
                self.directory = self.default_directory
                self.file_name = None
                self.is_saved = False
                self.SRC_updateTitle()
                return True
            else:
                pass

    def open(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if self.is_saved is False:
            reply = QMessageBox.question(
                self,
                app.applicationDisplayName(),
                translations[settings.value("appLanguage")]["open"],
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
            translations[settings.value("appLanguage")]["open_title"]
            + f" — {app.applicationDisplayName()}",
            self.directory,
            file_filter,
            options=options,
        )

        if selected_file:
            self.loadFile(selected_file)
            return True
        else:
            return False

    def loadFile(self, file_path):
        self.selected_file = file_path
        self.file_name = os.path.basename(self.selected_file)
        self.directory = os.path.dirname(self.selected_file)

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
        if self.is_saved == False:
            self.saveFile()
        elif self.file_name == None:
            self.saveAs()
        else:
            self.saveFile()

    def saveAs(self):
        options = QFileDialog.Options()
        settings = QSettings("berkaygediz", "SpanRC")
        options |= QFileDialog.ReadOnly
        file_filter = f"{translations[settings.value('appLanguage')]['xsrc']} (*.xsrc);;Comma Separated Values (*.csv)"
        selected_file, _ = QFileDialog.getSaveFileName(
            self,
            translations[settings.value("appLanguage")]["save_as_title"]
            + f" — {app.applicationDisplayName()}",
            self.directory,
            file_filter,
            options=options,
        )
        if selected_file:
            self.file_name = selected_file
            self.directory = os.path.dirname(self.file_name)
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
    if getattr(sys, "frozen", False):
        applicationPath = sys._MEIPASS
    elif __file__:
        applicationPath = os.path.dirname(__file__)
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(os.path.join(applicationPath, "spanrc_icon.ico")))
    app.setOrganizationName("berkaygediz")
    app.setApplicationName("SpanRC")
    app.setApplicationDisplayName("SpanRC 2024.08")
    app.setApplicationVersion("1.4.2024.08-1")
    wb = SRC_Workbook()
    wb.show()
    sys.exit(app.exec())
