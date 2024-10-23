import csv
import datetime
import locale
import os
import sys

import matplotlib.pyplot as plt
import pandas as pd
import psutil
from PySide6.QtCore import *
from PySide6.QtGui import *
from PySide6.QtOpenGL import *
from PySide6.QtOpenGLWidgets import *
from PySide6.QtPrintSupport import *
from PySide6.QtWidgets import *

from modules.globals import *
from modules.threading import *


class SS_About(QMainWindow):
    def __init__(self, parent=None):
        super(SS_About, self).__init__(parent)
        self.setWindowFlags(Qt.Dialog)
        self.setWindowIcon(QIcon("solidsheets_icon.ico"))
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
            "GNU General Public License v3.0<br>GNU LESSER GENERAL PUBLIC LICENSE v3.0<br>Mozilla Public License Version 2.0<br><br><b>Libraries: </b> pandas-dev/pandas, matplotlib/matplotlib, openpyxl/openpyxl, PySide6, psutil<br><br>"
            "OpenGL: <b>ON</b></center>"
        )
        self.setCentralWidget(self.about_label)


class SS_UndoCommand(QUndoCommand):
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


class SS_Workbook(QMainWindow):
    def __init__(self, parent=None):
        super(SS_Workbook, self).__init__(parent)
        QTimer.singleShot(0, self.initUI)

    def initUI(self):
        starttime = datetime.datetime.now()
        self.setWindowIcon(QIcon("solidsheets_icon.ico"))
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumSize(768, 540)
        system_language = locale.getlocale()[1]
        settings = QSettings("berkaygediz", "SolidSheets")
        if system_language not in languages.items():
            settings.setValue("appLanguage", "1252")
            settings.sync()
        if settings.value("adaptiveResponse") == None:
            settings.setValue("adaptiveResponse", 1)
            settings.sync()

        centralWidget = QOpenGLWidget(self)

        layout = QVBoxLayout(centralWidget)
        self.hardwareAcceleration = QOpenGLWidget()
        layout.addWidget(self.hardwareAcceleration)
        self.setCentralWidget(centralWidget)

        self.SS_thread = ThreadingEngine(
            adaptiveResponse=settings.value("adaptiveResponse")
        )
        self.SS_thread.update.connect(self.updateStatistics)

        self.themePalette()
        self.selected_file = None
        self.file_name = None
        self.is_saved = None
        self.default_directory = QDir().homePath()
        self.directory = self.default_directory

        self.undo_stack = QUndoStack(self)
        self.undo_stack.setUndoLimit(100)

        self.initDock()
        self.dock_widget.hide()

        self.status_bar = self.statusBar()
        self.SpreadsheetArea = QTableWidget(self)
        self.setCentralWidget(self.SpreadsheetArea)

        self.SpreadsheetArea.setDisabled(True)
        self.initActions()
        self.initToolbar()
        self.adaptiveResponse = settings.value("adaptiveResponse")

        self.setPalette(self.light_theme)
        self.SpreadsheetArea.itemSelectionChanged.connect(self.SS_thread.start)
        self.SpreadsheetArea.setCursor(Qt.CursorShape.SizeAllCursor)
        self.showMaximized()
        self.setFocus()

        QTimer.singleShot(50 * self.adaptiveResponse, self.restoreTheme)
        QTimer.singleShot(150 * self.adaptiveResponse, self.restoreState)
        if (
            self.SpreadsheetArea.columnCount() == 0
            and self.SpreadsheetArea.rowCount() == 0
        ):
            self.SpreadsheetArea.setColumnCount(100)
            self.SpreadsheetArea.setRowCount(50)
            self.SpreadsheetArea.clearSpans()
            self.SpreadsheetArea.setItem(0, 0, QTableWidgetItem(""))

        self.SpreadsheetArea.setDisabled(False)
        self.updateTitle()

        endtime = datetime.datetime.now()
        self.status_bar.showMessage(
            str((endtime - starttime).total_seconds()) + " ms", 2500
        )

    def closeEvent(self, event):
        settings = QSettings("berkaygediz", "SolidSheets")
        if self.is_saved == False:
            reply = QMessageBox.question(
                self,
                f"{app.applicationDisplayName()}",
                translations[settings.value("appLanguage")]["exit_message"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.saveState()
                event.accept()
            else:
                self.saveState()
                event.ignore()
        else:
            self.saveState()
            event.accept()

    def changeLanguage(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        settings.setValue("appLanguage", self.language_combobox.currentData())
        settings.sync()
        self.updateTitle()
        self.updateStatistics()
        self.toolbarTranslate()

    def updateTitle(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        file = (
            self.file_name
            if self.file_name
            else translations[settings.value("appLanguage")]["new_title"]
        )
        if file.endswith(".xlsx") or file.endswith(".xsrc"):
            textMode = " - Read Only"
        else:
            textMode = ""
        if self.is_saved == True:
            asterisk = ""
        else:
            asterisk = "*"
        self.setWindowTitle(
            f"{file}{asterisk}{textMode} — {app.applicationDisplayName()}"
        )

    def updateStatistics(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        row = self.SpreadsheetArea.rowCount()
        column = self.SpreadsheetArea.columnCount()
        selected_cell = (
            self.SpreadsheetArea.currentRow() + 1,
            self.SpreadsheetArea.currentColumn() + 1,
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
        if self.SpreadsheetArea.selectedRanges():
            statistics += f"<td>{translations[settings.value('appLanguage')]['statistics_message3']}</td><td>"
            for selected_range in self.SpreadsheetArea.selectedRanges():
                statistics += f"{selected_range.topRow() + 1}:{selected_range.leftColumn() + 1} - {selected_range.bottomRow() + 1}:{selected_range.rightColumn() + 1}</td>"
        else:
            statistics += f"<td>{translations[settings.value('appLanguage')]['statistics_message4']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"

        statistics += f"</td><td id='sr-text'>{app.applicationDisplayName()}</td></tr></table></body></html>"
        self.statistics_label.setText(statistics)
        self.statusBar().addPermanentWidget(self.statistics_label)
        self.updateTitle()

    def saveState(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        settings.setValue("windowScale", self.saveGeometry())
        settings.setValue("defaultDirectory", self.default_directory)
        settings.setValue("fileName", self.file_name)
        settings.setValue("isSaved", self.is_saved)
        settings.setValue(
            "scrollPosition", self.SpreadsheetArea.verticalScrollBar().value()
        )
        settings.setValue(
            "appTheme", "dark" if self.palette() == self.dark_theme else "light"
        )
        settings.setValue("appLanguage", self.language_combobox.currentData())
        settings.setValue("adaptiveResponse", self.adaptiveResponse)
        settings.sync()

    def restoreState(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        self.geometry = settings.value("windowScale")
        self.directory = settings.value("defaultDirectory", self.default_directory)
        self.file_name = settings.value("fileName")
        self.is_saved = settings.value("isSaved")
        index = self.language_combobox.findData(settings.value("appLanguage"))
        self.language_combobox.setCurrentIndex(index)

        if self.geometry is not None:
            self.restoreGeometry(self.geometry)

        if self.file_name and os.path.exists(self.file_name):
            if self.file_name.endswith(".xlsx"):
                read_file = pd.read_excel(self.file_name)
                read_file.to_csv(
                    f"{self.file_name}-transformed.ssfs", index=None, header=True
                )
                # df = pd.DataFrame(pd.read_csv(f"{self.file_name}-transformed.ssfs"))
                self.LoadSpreadsheet(f"{self.file_name}-transformed.ssfs")
            else:
                if (
                    self.file_name.endswith(".ssfs")
                    or self.file_name.endswith(".xsrc")
                    or self.file_name.endswith(".csv")
                ):
                    self.LoadSpreadsheet(self.file_name)

        self.SpreadsheetArea.setColumnCount(
            int(settings.value("columnCount", self.SpreadsheetArea.columnCount()))
        )
        self.SpreadsheetArea.setRowCount(
            int(settings.value("rowCount", self.SpreadsheetArea.rowCount()))
        )

        for row in range(self.SpreadsheetArea.rowCount()):
            for column in range(self.SpreadsheetArea.columnCount()):
                if settings.value(
                    f"row{row}column{column}rowspan", None
                ) and settings.value(f"row{row}column{column}columnspan", None):
                    self.SpreadsheetArea.setSpan(
                        row,
                        column,
                        int(settings.value(f"row{row}column{column}rowspan")),
                        int(settings.value(f"row{row}column{column}columnspan")),
                    )
        for row in range(self.SpreadsheetArea.rowCount()):
            for column in range(self.SpreadsheetArea.columnCount()):
                if settings.value(f"row{row}column{column}text", None):
                    self.SpreadsheetArea.setItem(
                        row,
                        column,
                        QTableWidgetItem(settings.value(f"row{row}column{column}text")),
                    )

        if self.file_name != None:
            self.SpreadsheetArea.resizeColumnsToContents()
            self.SpreadsheetArea.resizeRowsToContents()

        self.SpreadsheetArea.setCurrentCell(
            int(settings.value("currentRow", 0)),
            int(settings.value("currentColumn", 0)),
        )
        self.SpreadsheetArea.scrollToItem(
            self.SpreadsheetArea.item(
                int(settings.value("currentRow", 0)),
                int(settings.value("currentColumn", 0)),
            )
        )

        if self.file_name:
            self.is_saved = True
        else:
            self.is_saved = False

        self.adaptiveResponse = settings.value("adaptiveResponse")
        self.restoreTheme(),
        self.updateTitle()

    def restoreTheme(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        if settings.value("appTheme") == "dark":
            self.setPalette(self.dark_theme)
        else:
            self.setPalette(self.light_theme)
        self.toolbarTheme()

    def themePalette(self):
        self.light_theme = QPalette()
        self.dark_theme = QPalette()

        self.light_theme.setColor(QPalette.Window, QColor(3, 65, 135))
        self.light_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Base, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Text, QColor(0, 0, 0))
        self.light_theme.setColor(QPalette.Highlight, QColor(105, 117, 156))
        self.light_theme.setColor(QPalette.Button, QColor(230, 230, 230))
        self.light_theme.setColor(QPalette.ButtonText, QColor(0, 0, 0))

        self.dark_theme.setColor(QPalette.Window, QColor(35, 39, 52))
        self.dark_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Base, QColor(80, 85, 122))
        self.dark_theme.setColor(QPalette.Text, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Highlight, QColor(105, 117, 156))
        self.dark_theme.setColor(QPalette.Button, QColor(0, 0, 0))
        self.dark_theme.setColor(QPalette.ButtonText, QColor(255, 255, 255))

    def themeAction(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        if self.palette() == self.light_theme:
            self.setPalette(self.dark_theme)
            settings.setValue("appTheme", "dark")
        else:
            self.setPalette(self.light_theme)
            settings.setValue("appTheme", "light")
        self.toolbarTheme()

    def toolbarTheme(self):
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

    def toolbarTranslate(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        system_language = locale.getlocale()[1]
        if system_language not in languages.keys():
            settings.setValue("appLanguage", "1252")
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
        self.addrowAction.setText(
            translations[settings.value("appLanguage")]["add_row"]
        )
        self.addrowAction.setStatusTip(
            translations[settings.value("appLanguage")]["add_row"]
        )
        self.addcolumnAction.setText(
            translations[settings.value("appLanguage")]["add_column"]
        )
        self.addcolumnAction.setStatusTip(
            translations[settings.value("appLanguage")]["add_column"]
        )
        self.addrowaboveAction.setText(
            translations[settings.value("appLanguage")]["add_row_above"]
        )
        self.addrowaboveAction.setStatusTip(
            translations[settings.value("appLanguage")]["add_row_above"]
        )
        self.addcolumnleftAction.setText(
            translations[settings.value("appLanguage")]["add_column_left"]
        )
        self.addcolumnleftAction.setStatusTip(
            translations[settings.value("appLanguage")]["add_column_left"]
        )
        self.darklightAction.setText(
            translations[settings.value("appLanguage")]["darklight"]
        )
        self.darklightAction.setStatusTip(
            translations[settings.value("appLanguage")]["darklight_message"]
        )
        self.powersaveraction.setText(
            translations[settings.value("appLanguage")]["powersaver"]
        )
        self.powersaveraction.setStatusTip(
            translations[settings.value("appLanguage")]["powersaver_message"]
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

    def initDock(self):
        settings = QSettings("berkaygediz", "SolidSheets")
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

    def toolbarLabel(self, toolbar, text):
        label = QLabel(f"<b>{text}</b>")
        toolbar.addWidget(label)

    def createAction(self, text, status_tip, function, shortcut=None, icon=None):
        action = QAction(text, self)
        action.setStatusTip(status_tip)
        action.triggered.connect(function)
        if shortcut:
            action.setShortcut(shortcut)
        if icon:
            action.setIcon(QIcon(""))
        return action

    def initActions(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        self.newAction = self.createAction(
            translations[settings.value("appLanguage")]["new"],
            translations[settings.value("appLanguage")]["new_title"],
            self.new,
            QKeySequence.New,
            "",
        )
        self.openAction = self.createAction(
            translations[settings.value("appLanguage")]["open"],
            translations[settings.value("appLanguage")]["open_title"],
            self.Open,
            QKeySequence.Open,
            "",
        )
        self.saveAction = self.createAction(
            translations[settings.value("appLanguage")]["save"],
            translations[settings.value("appLanguage")]["save_title"],
            self.Save,
            QKeySequence.Save,
            "",
        )
        self.saveasAction = self.createAction(
            translations[settings.value("appLanguage")]["save_as"],
            translations[settings.value("appLanguage")]["save_as_title"],
            self.SaveAs,
            QKeySequence.SaveAs,
            "",
        )
        self.printAction = self.createAction(
            translations[settings.value("appLanguage")]["print"],
            translations[settings.value("appLanguage")]["print_title"],
            self.PrintSpreadsheet,
            QKeySequence.Print,
            "",
        )
        self.exitAction = self.createAction(
            translations[settings.value("appLanguage")]["exit"],
            translations[settings.value("appLanguage")]["exit_message"],
            self.close,
            QKeySequence.Quit,
            "",
        )
        self.deleteAction = self.createAction(
            translations[settings.value("appLanguage")]["delete"],
            translations[settings.value("appLanguage")]["delete_title"],
            self.cellDelete,
            QKeySequence.Delete,
            "",
        )
        self.addrowAction = self.createAction(
            translations[settings.value("appLanguage")]["add_row"],
            translations[settings.value("appLanguage")]["add_row_title"],
            self.rowAdd,
            QKeySequence("Ctrl+Shift+R"),
            "",
        )
        self.addcolumnAction = self.createAction(
            translations[settings.value("appLanguage")]["add_column"],
            translations[settings.value("appLanguage")]["add_column_title"],
            self.columnAdd,
            QKeySequence("Ctrl+Shift+C"),
            "",
        )
        self.addrowaboveAction = self.createAction(
            translations[settings.value("appLanguage")]["add_row_above"],
            translations[settings.value("appLanguage")]["add_row_above_title"],
            self.rowAddAbove,
            QKeySequence("Ctrl+Shift+T"),
            "",
        )
        self.addcolumnleftAction = self.createAction(
            translations[settings.value("appLanguage")]["add_column_left"],
            translations[settings.value("appLanguage")]["add_column_left_title"],
            self.columnAddLeft,
            QKeySequence("Ctrl+Shift+L"),
            "",
        )
        self.hide_dock_widget_action = self.createAction(
            translations[settings.value("appLanguage")]["help"],
            translations[settings.value("appLanguage")]["help"],
            self.SS_toggleDock,
            QKeySequence("Ctrl+H"),
            "",
        )
        self.aboutAction = self.createAction(
            translations[settings.value("appLanguage")]["about"],
            translations[settings.value("appLanguage")]["about_title"],
            self.viewAbout,
            QKeySequence.HelpContents,
            "",
        )
        self.undoAction = self.createAction(
            translations[settings.value("appLanguage")]["undo"],
            translations[settings.value("appLanguage")]["undo_title"],
            self.undo_stack.undo,
            QKeySequence.Undo,
            "",
        )
        self.redoAction = self.createAction(
            translations[settings.value("appLanguage")]["redo"],
            translations[settings.value("appLanguage")]["redo_title"],
            self.undo_stack.redo,
            QKeySequence.Redo,
            "",
        )
        self.darklightAction = self.createAction(
            translations[settings.value("appLanguage")]["darklight"],
            translations[settings.value("appLanguage")]["darklight_message"],
            self.themeAction,
            QKeySequence("Ctrl+D"),
            "",
        )

    def initToolbar(self):
        settings = QSettings("berkaygediz", "SolidSheets")

        self.file_toolbar = self.addToolBar(
            translations[settings.value("appLanguage")]["file"]
        )
        self.file_toolbar.setObjectName("File")
        self.toolbarLabel(
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
        self.toolbarLabel(
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
        self.toolbarLabel(
            self.interface_toolbar,
            translations[settings.value("appLanguage")]["interface"] + ": ",
        )
        self.theme_action = self.createAction(
            translations[settings.value("appLanguage")]["darklight"],
            translations[settings.value("appLanguage")]["darklight_message"],
            self.themeAction,
            QKeySequence("Ctrl+Shift+T"),
            "",
        )
        self.theme_action.setCheckable(True)
        self.theme_action.setChecked(settings.value("appTheme") == "dark")

        self.interface_toolbar.addAction(self.theme_action)
        self.powersaveraction = QAction(
            translations[settings.value("appLanguage")]["powersaver"],
            self,
            checkable=True,
        )
        self.powersaveraction.setStatusTip(
            translations[settings.value("appLanguage")]["powersaver_message"]
        )
        self.powersaveraction.toggled.connect(self.hybridSaver)

        self.interface_toolbar.addAction(self.powersaveraction)
        if settings.value("adaptiveResponse") == None:
            response_exponential = settings.setValue(
                "adaptiveResponse", fallbackValues["adaptiveResponse"]
            )
        else:
            response_exponential = settings.value(
                "adaptiveResponse",
            )

        self.powersaveraction.setChecked(response_exponential > 1)
        self.interface_toolbar.addAction(self.powersaveraction)
        self.interface_toolbar.addAction(self.hide_dock_widget_action)
        self.interface_toolbar.addAction(self.aboutAction)
        self.language_combobox = QComboBox(self)
        self.language_combobox.setStyleSheet("background-color:#000000; color:#FFFFFF;")
        for lcid, name in languages.items():
            self.language_combobox.addItem(name, lcid)

        self.language_combobox.currentIndexChanged.connect(self.changeLanguage)
        self.interface_toolbar.addWidget(self.language_combobox)
        self.addToolBarBreak()
        self.formula_toolbar = self.addToolBar(
            translations[settings.value("appLanguage")]["formula"]
        )
        self.formula_toolbar.setObjectName("Formula")
        self.toolbarLabel(
            self.formula_toolbar,
            translations[settings.value("appLanguage")]["formula"] + ": ",
        )
        self.formula_edit = QLineEdit()
        self.formula_edit.setPlaceholderText(
            translations[settings.value("appLanguage")]["formula"]
        )
        self.formula_edit.returnPressed.connect(self.computeFormula)
        self.formula_button = QPushButton(
            translations[settings.value("appLanguage")]["compute"]
        )
        self.formula_button.setStyleSheet(
            "background-color: #A72461; color: #FFFFFF; font-weight: bold; padding: 10px; border-radius: 10px; border: 1px solid #000000; margin-left: 10px;"
        )
        self.formula_button.setCursor(Qt.PointingHandCursor)
        self.formula_button.clicked.connect(self.computeFormula)
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

    def SS_createDockwidget(self):
        widget = QWidget()
        layout = QVBoxLayout()
        widget.setLayout(layout)
        self.statistics = QLabel()
        self.statistics.setText("Statistics")
        layout.addWidget(self.statistics)
        return widget

    def SS_toggleDock(self):
        if self.dock_widget.isHidden():
            self.dock_widget.show()
        else:
            self.dock_widget.hide()

    def hybridSaver(self, checked):
        settings = QSettings("berkaygediz", "SolidSheets")
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
        settings = QSettings("berkaygediz", "SolidSheets")
        if self.is_saved == True:
            self.SpreadsheetArea.clearContents()
            self.SpreadsheetArea.setRowCount(50)
            self.SpreadsheetArea.setColumnCount(100)
            self.directory = self.default_directory
            self.file_name = None
            self.is_saved = False
            self.updateTitle()
        else:
            reply = QMessageBox.question(
                self,
                app.applicationDisplayName(),
                translations[settings.value("appLanguage")]["new_title"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.SpreadsheetArea.clearContents()
                self.SpreadsheetArea.setRowCount(50)
                self.SpreadsheetArea.setColumnCount(100)
                self.setWindowTitle(
                    translations[settings.value("appLanguage")]["new_title"]
                    + f" — {app.applicationDisplayName()}"
                )
                self.directory = self.default_directory
                self.file_name = None
                self.is_saved = False
                self.updateTitle()
                return True
            else:
                pass

    def Open(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        if self.is_saved is False:
            reply = QMessageBox.question(
                self,
                app.applicationDisplayName(),
                translations[settings.value("appLanguage")]["open"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.saveState()
                return self.OpenProcess()
        else:
            return self.OpenProcess()

    def OpenProcess(self):
        settings = QSettings("berkaygediz", "SolidSheets")
        options = QFileDialog.Options()
        selected_file, _ = QFileDialog.getOpenFileName(
            self,
            translations[settings.value("appLanguage")]["open_title"]
            + f" — {app.applicationDisplayName()}",
            self.directory,
            fallbackValues["readFilter"],
            options=options,
        )

        if selected_file:
            self.file_name = selected_file

            if self.file_name.endswith(".xlsx"):
                try:
                    read_file = pd.read_excel(self.file_name)
                    read_file.to_csv(
                        f"{self.file_name}-transformed.ssfs", index=None, header=True
                    )
                    # df = pd.DataFrame(pd.read_csv(f"{self.file_name}-transformed.ssfs"))
                    self.LoadSpreadsheet(f"{self.file_name}-transformed.ssfs")
                except:
                    QMessageBox.warning(self, None, "Conversion failed.")
            else:
                if (
                    self.file_name.endswith(".ssfs")
                    or self.file_name.endswith(".xsrc")
                    or self.file_name.endswith(".csv")
                ):
                    self.LoadSpreadsheet(self.file_name)

            self.directory = os.path.dirname(self.file_name)
            self.is_saved = True
            self.updateTitle()
        else:
            pass

    def LoadSpreadsheet(self, file_path):
        with open(file_path, "r") as file:
            reader = csv.reader(file)
            self.SpreadsheetArea.clearSpans()
            self.SpreadsheetArea.setRowCount(0)
            self.SpreadsheetArea.setColumnCount(0)
            for row in reader:
                rowPosition = self.SpreadsheetArea.rowCount()
                self.SpreadsheetArea.insertRow(rowPosition)
                self.SpreadsheetArea.setColumnCount(len(row))
                for column in range(len(row)):
                    item = QTableWidgetItem(row[column])
                    self.SpreadsheetArea.setItem(rowPosition, column, item)
        self.SpreadsheetArea.resizeColumnsToContents()
        self.SpreadsheetArea.resizeRowsToContents()

    def Save(self):
        if self.is_saved == False:
            self.SaveProcess()
        elif self.file_name == None:
            self.SaveAs()
        else:
            self.SaveProcess()

    def SaveAs(self):
        options = QFileDialog.Options()
        settings = QSettings("berkaygediz", "SolidSheets")
        options |= QFileDialog.ReadOnly
        selected_file, _ = QFileDialog.getSaveFileName(
            self,
            translations[settings.value("appLanguage")]["save_as_title"]
            + f" — {app.applicationDisplayName()}",
            self.directory,
            fallbackValues["writeFilter"],
            options=options,
        )
        if selected_file:
            self.file_name = selected_file
            self.directory = os.path.dirname(self.file_name)
            self.SaveProcess()
            return True
        else:
            return False

    def SaveProcess(self):
        if not self.file_name:
            self.SaveAs()
        else:
            if self.file_name.lower().endswith(".xlsx"):
                None
            else:
                with open(self.file_name, "w", newline="") as file:
                    writer = csv.writer(file)
                    for rowid in range(self.SpreadsheetArea.rowCount()):
                        row = []
                        for colid in range(self.SpreadsheetArea.columnCount()):
                            item = self.SpreadsheetArea.item(rowid, colid)
                            if item is not None:
                                row.append(item.text())
                            else:
                                row.append("")
                        writer.writerow(row)

        self.status_bar.showMessage("Saved.", 2000)
        self.is_saved = True
        self.updateTitle()

    def PrintSpreadsheet(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOrientation(QPrinter.Landscape)
        printer.setPageMargins(0, 0, 0, 0, QPrinter.Millimeter)
        printer.setFullPage(True)
        printer.setDocName(self.file_name)

        preview_dialog = QPrintPreviewDialog(printer, self)
        preview_dialog.paintRequested.connect(self.SpreadsheetArea.print_)
        preview_dialog.exec_()

    def viewAbout(self):
        self.viewAbout = SS_About()
        self.viewAbout.show()

    def selectedCells(self):
        selected_cells = self.SpreadsheetArea.selectedRanges()

        if not selected_cells:
            raise ValueError("ERROR: 0 cells.")

        values = []

        for selected_range in selected_cells:
            for row in range(selected_range.topRow(), selected_range.bottomRow() + 1):
                for col in range(
                    selected_range.leftColumn(), selected_range.rightColumn() + 1
                ):
                    item = self.SpreadsheetArea.item(row, col)
                    if item and item.text().strip().isdigit():
                        values.append(int(item.text()))

        return values

    def computeFormula(self):
        try:
            formulavalue = self.formula_edit.text().strip()
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

    def cellDelete(self):
        for item in self.SpreadsheetArea.selectedItems():
            item.setText("")

    def rowAdd(self):
        self.SpreadsheetArea.insertRow(self.SpreadsheetArea.rowCount())

    def rowAddAbove(self):
        self.SpreadsheetArea.insertRow(self.SpreadsheetArea.currentRow())

    def columnAdd(self):
        self.SpreadsheetArea.insertColumn(self.SpreadsheetArea.columnCount())

    def columnAddLeft(self):
        self.SpreadsheetArea.insertColumn(self.SpreadsheetArea.currentColumn())


if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        applicationPath = sys._MEIPASS
    elif __file__:
        applicationPath = os.path.dirname(__file__)
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(os.path.join(applicationPath, "solidsheets_icon.ico")))
    app.setOrganizationName("berkaygediz")
    app.setApplicationName("SolidSheets")
    app.setApplicationDisplayName("SolidSheets 2024.10")
    app.setApplicationVersion("1.4.2024.10-1")
    wb = SS_Workbook()
    wb.show()
    sys.exit(app.exec())
