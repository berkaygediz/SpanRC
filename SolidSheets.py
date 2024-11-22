import csv
import datetime
import locale
import os
import platform
import sys
from datetime import timezone

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

try:
    from ctypes import windll

    windll.shell32.SetCurrentProcessExplicitAppUserModelID(
        "berkaygediz.SolidSheets.1.5"
    )
except ImportError:
    pass

try:
    settings = QSettings("berkaygediz", "SolidSheets")
    lang = settings.value("appLanguage")
except:
    pass


class SS_ControlInfo(QMainWindow):
    def __init__(self, parent=None):
        super(SS_ControlInfo, self).__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setWindowIcon(QIcon(fallbackValues["icon"]))
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)

        screen = QGuiApplication.primaryScreen().availableGeometry()
        self.setGeometry(
            QStyle.alignedRect(
                Qt.LayoutDirection.LeftToRight,
                Qt.AlignmentFlag.AlignCenter,
                QSize(int(screen.width() * 0.2), int(screen.height() * 0.2)),
                screen,
            )
        )
        self.setStyleSheet("background-color: transparent;")
        self.setWindowOpacity(0.75)

        self.widget_central = QWidget(self)
        self.layout_central = QVBoxLayout(self.widget_central)

        self.title = QLabel("SolidSheets", self)
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title.setFont(QFont("Roboto", 30))
        self.title.setStyleSheet(
            "background-color: #4CAF50; color: #FFFFFF; font-weight: bold; font-size: 30px; border-radius: 25px; border: 1px solid #388E3C;"
        )

        self.layout_central.addWidget(self.title)
        self.setCentralWidget(self.widget_central)

        QTimer.singleShot(500, self.showWB)

    def showWB(self):
        self.hide()
        self.wb_window = SS_Workbook()
        self.wb_window.show()


class SS_About(QMainWindow):
    def __init__(self, parent=None):
        super(SS_About, self).__init__(parent)
        self.setWindowFlags(Qt.Dialog)
        self.setWindowIcon(QIcon(fallbackValues["icon"]))
        self.setWindowModality(Qt.WindowModality.ApplicationModal),
        self.setMinimumSize(540, 300)

        lang = settings.value("appLanguage")

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
            f"<b>{QApplication.applicationName()}</b><br><br>"
            "Real-time calculation and formula supported spreadsheet editor.<br>"
            "Made by Berkay Gediz<br><br>"
            "Licenses:<br>"
            "- GNU General Public License v3.0<br>"
            "- GNU LESSER GENERAL PUBLIC LICENSE v3.0<br>"
            "- Mozilla Public License Version 2.0<br><br>"
            "<b>Libraries:</b> pandas-dev/pandas, matplotlib/matplotlib, "
            "openpyxl/openpyxl, PySide6, psutil<br><br>"
            "<b>OpenGL:</b> <b>ON</b></center>"
        )
        self.setCentralWidget(self.about_label)


class SS_Help(QMainWindow):
    def __init__(self, parent=None):
        super(SS_Help, self).__init__(parent)
        self.setWindowFlags(Qt.Dialog)
        self.setWindowIcon(QIcon(fallbackValues["icon"]))
        self.setWindowModality(Qt.WindowModality.WindowModal)
        self.setMinimumSize(540, 460)

        lang = settings.value("appLanguage")

        self.setGeometry(
            QStyle.alignedRect(
                Qt.LeftToRight,
                Qt.AlignCenter,
                self.size(),
                QApplication.primaryScreen().availableGeometry(),
            )
        )

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        main_widget = QWidget()
        layout = QVBoxLayout()
        layout.setContentsMargins(0, 20, 0, 0)

        self.help_label = QLabel()
        self.help_label.setWordWrap(True)
        self.help_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.help_label.setTextFormat(Qt.RichText)

        self.help_label.setText(
            "<html><head><style>"
            "table {border-collapse: collapse; width: 80%; margin: auto;}"
            "th, td {text-align: left; padding: 8px;}"
            "tr:nth-child(even) {background-color: #f2f2f2;}"
            "tr:hover {background-color: #ddd;}"
            "th {background-color: #4CAF50; color: white;}"
            "body {text-align: center;}"
            "</style></head><body>"
            "<h1>Help</h1>"
            "<table><tr><th>Shortcut</th><th>Function</th></tr>"
            f"<tr><td>Ctrl + O</td><td>{translations[lang]['open_title']}</td></tr>"
            f"<tr><td>Ctrl + S</td><td>{translations[lang]['save_title']}</td></tr>"
            f"<tr><td>Ctrl + N</td><td>{translations[lang]['new_title']}</td></tr>"
            f"<tr><td>Ctrl + Shift + S</td><td>{translations[lang]['save_as_title']}</td></tr>"
            f"<tr><td>Ctrl + P</td><td>{translations[lang]['print_title']}</td></tr>"
            f"<tr><td>Ctrl + Q</td><td>{translations[lang]['exit_title']}</td></tr>"
            f"<tr><td>Ctrl + D</td><td>{translations[lang]['delete_title']}</td></tr>"
            f"<tr><td>Ctrl + A</td><td>{translations[lang]['about_title']}</td></tr>"
            f"<tr><td>Ctrl + Z</td><td>{translations[lang]['undo_title']}</td></tr>"
            f"<tr><td>Ctrl + Y</td><td>{translations[lang]['redo_title']}</td></tr>"
            f"<tr><td>Ctrl + L</td><td>{translations[lang]['darklight_message']}</td></tr>"
            f"<tr><td>Ctrl + Shift + R</td><td>{translations[lang]['add_row_title']}</td></tr>"
            f"<tr><td>Ctrl + Shift + C</td><td>{translations[lang]['add_column_title']}</td></tr>"
            f"<tr><td>Ctrl + Shift + T</td><td>{translations[lang]['add_row_above_title']}</td></tr>"
            f"<tr><td>Ctrl + Shift + L</td><td>{translations[lang]['add_column_left_title']}</td></tr>"
            "</table></body></html>"
        )

        layout.addWidget(self.help_label)
        main_widget.setLayout(layout)
        scroll_area.setWidget(main_widget)

        self.setCentralWidget(scroll_area)


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
        if item is None:
            return
        self.table.blockSignals(True)
        try:
            item.setText(self.new_data)
        finally:
            self.table.blockSignals(False)

    def undo(self):
        item = self.table.item(self.row, self.col)
        if item is None:
            return
        self.table.blockSignals(True)
        try:
            item.setText(self.old_data)
        finally:
            self.table.blockSignals(False)


class SS_Workbook(QMainWindow):
    def __init__(self, parent=None):
        super(SS_Workbook, self).__init__(parent)

        QTimer.singleShot(0, self.initUI)

    def initUI(self):
        starttime = datetime.datetime.now()
        self.setWindowIcon(QIcon(fallbackValues["icon"]))
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setMinimumSize(768, 540)
        system_language = locale.getlocale()[1]
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
        self.SpreadsheetArea.horizontalHeader().sectionClicked.connect(
            self.changeColumnName
        )
        self.SpreadsheetArea.itemChanged.connect(self.onItemChanged)
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
            self.SpreadsheetArea.setItem(0, 0, None)

        self.SpreadsheetArea.setDisabled(False)
        self.updateTitle()

        endtime = datetime.datetime.now()
        self.status_bar.showMessage(
            str((endtime - starttime).total_seconds()) + " sec", 2500
        )

    def onItemChanged(self, item):
        if item.row() == 0:
            header_text = item.text()
            self.SpreadsheetArea.setHorizontalHeaderItem(
                item.column(), QTableWidgetItem(header_text)
            )

    def closeEvent(self, event):
        if self.is_saved == False:
            reply = QMessageBox.question(
                self,
                f"{app.applicationDisplayName()}",
                translations[lang]["exit_message"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.saveState()
                self.cleanupCache()
                event.accept()
            else:
                self.saveState()
                event.ignore()
        else:
            self.saveState()
            self.cleanupCache()
            event.accept()

    def cleanupCache(self):
        cache_dir = self.controlCacheDir()
        if os.path.exists(cache_dir):
            for filename in os.listdir(cache_dir):
                if filename.startswith("solidsheets_G") and filename.endswith(
                    (".png", ".jpg", ".jpeg")
                ):
                    file_path = os.path.join(cache_dir, filename)
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        QMessageBox.critical(
                            self, "Cache Cleanup", f"Error deleting file: {e}"
                        )

    def changeLanguage(self):
        settings.setValue("appLanguage", self.language_combobox.currentData())
        settings.sync()
        self.toolbarTranslate()
        self.updateStatistics()
        self.updateTitle()

    def updateTitle(self):
        lang = settings.value("appLanguage")
        file = self.file_name or translations[lang]["new_title"]

        textMode = (
            translations[lang]["readonly"] if file.endswith((".xlsx", ".xsrc")) else ""
        )
        asterisk = "*" if not self.is_saved else ""

        self.setWindowTitle(
            f"{file}{asterisk}{textMode} — {app.applicationDisplayName()}"
        )

    def updateStatistics(self):
        row = self.SpreadsheetArea.rowCount()
        column = self.SpreadsheetArea.columnCount()
        selected_cell = (
            self.SpreadsheetArea.currentRow() + 1,
            self.SpreadsheetArea.currentColumn() + 1,
        )

        lang = settings.value("appLanguage")

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
            f"<th>{translations[lang]['statistics_title']}</th>"
        )

        statistics += (
            f"<td>{translations[lang]['statistics_message1']}</td><td>{row}</td>"
        )
        statistics += (
            f"<td>{translations[lang]['statistics_message2']}</td><td>{column}</td>"
        )
        statistics += (
            f"<td>{translations[lang]['statistics_message3']}</td><td>{row*column}</td>"
        )

        if selected_cell:
            statistics += f"<td>{translations[lang]['statistics_message3']}</td><td>"
            statistics += f"{selected_cell[0]}:{selected_cell[1]}</td>"
        else:
            statistics += f"<td>{translations[lang]['statistics_message4']}</td><td>{selected_cell[0]}:{selected_cell[1]}</td>"

        statistics += f"</td><td id='sr-text'>{app.applicationDisplayName()}</td></tr></table></body></html>"

        self.statistics_label.setText(statistics)
        self.statusBar().addPermanentWidget(self.statistics_label)
        self.updateTitle()

        self.SpreadsheetArea.resizeColumnsToContents()
        self.SpreadsheetArea.resizeRowsToContents()

    def saveState(self):
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
        self.geometry = settings.value("windowScale")
        self.directory = settings.value("defaultDirectory", self.default_directory)
        self.file_name = settings.value("fileName")
        self.is_saved = settings.value("isSaved")
        index = self.language_combobox.findData(lang)
        self.language_combobox.setCurrentIndex(index)

        if self.geometry is not None:
            self.restoreGeometry(self.geometry)

        if self.file_name and os.path.exists(self.file_name):
            if self.file_name.endswith(".xlsx"):
                self.loadSpreadsheetFromExcel(self.file_name)
            else:
                self.loadSpreadsheetFromOrigin(self.file_name)

        self.setSpreadsheetSize()
        self.restoreCellProperties()
        if self.file_name:
            self.SpreadsheetArea.resizeColumnsToContents()
            self.SpreadsheetArea.resizeRowsToContents()

        self.restoreCurrentCell()
        self.adaptiveResponse = settings.value("adaptiveResponse")
        self.restoreTheme()
        self.updateTitle()

    def setSpreadsheetSize(self):
        self.SpreadsheetArea.setColumnCount(
            int(settings.value("columnCount", self.SpreadsheetArea.columnCount()))
        )
        self.SpreadsheetArea.setRowCount(
            int(settings.value("rowCount", self.SpreadsheetArea.rowCount()))
        )

    def restoreCellProperties(self):
        for row in range(self.SpreadsheetArea.rowCount()):
            for column in range(self.SpreadsheetArea.columnCount()):
                rowspan = settings.value(f"row{row}column{column}rowspan", None)
                columnspan = settings.value(f"row{row}column{column}columnspan", None)
                if rowspan and columnspan:
                    self.SpreadsheetArea.setSpan(
                        row,
                        column,
                        int(rowspan),
                        int(columnspan),
                    )

        for row in range(self.SpreadsheetArea.rowCount()):
            for column in range(self.SpreadsheetArea.columnCount()):
                cell_text = settings.value(f"row{row}column{column}text", None)
                if cell_text is not None:
                    self.SpreadsheetArea.setItem(
                        row,
                        column,
                        QTableWidgetItem(cell_text),
                    )

    def restoreCurrentCell(self):
        current_row = int(settings.value("currentRow", 0))
        current_column = int(settings.value("currentColumn", 0))
        self.SpreadsheetArea.setCurrentCell(current_row, current_column)
        self.SpreadsheetArea.scrollToItem(
            self.SpreadsheetArea.item(current_row, current_column)
        )

    def restoreTheme(self):
        if settings.value("appTheme") == "dark":
            self.setPalette(self.dark_theme)
        else:
            self.setPalette(self.light_theme)
        self.toolbarTheme()

    def themePalette(self):
        self.light_theme = QPalette()
        self.dark_theme = QPalette()

        self.light_theme.setColor(QPalette.Window, QColor(50, 50, 50))
        self.light_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Base, QColor(255, 255, 255))
        self.light_theme.setColor(QPalette.Text, QColor(0, 0, 0))
        self.light_theme.setColor(QPalette.Highlight, QColor(173, 216, 230))
        self.light_theme.setColor(QPalette.Button, QColor(220, 220, 220))
        self.light_theme.setColor(QPalette.ButtonText, QColor(0, 0, 0))

        self.dark_theme.setColor(QPalette.Window, QColor(50, 50, 50))
        self.dark_theme.setColor(QPalette.WindowText, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Base, QColor(70, 70, 70))
        self.dark_theme.setColor(QPalette.Text, QColor(255, 255, 255))
        self.dark_theme.setColor(QPalette.Highlight, QColor(100, 100, 255))
        self.dark_theme.setColor(QPalette.Button, QColor(60, 60, 60))
        self.dark_theme.setColor(QPalette.ButtonText, QColor(255, 255, 255))

    def themeAction(self):
        if self.palette() == self.light_theme:
            self.setPalette(self.dark_theme)
            settings.setValue("appTheme", "dark")
        else:
            self.setPalette(self.light_theme)
            settings.setValue("appTheme", "light")
        self.toolbarTheme()

    def toolbarTheme(self):
        text_color = QColor(255, 255, 255)

        for toolbar in self.findChildren(QToolBar):
            for action in toolbar.actions():
                if action.text():
                    action_color = QPalette()
                    action_color.setColor(QPalette.ButtonText, text_color)
                    action_color.setColor(QPalette.WindowText, text_color)
                    toolbar.setPalette(action_color)

    def toolbarTranslate(self):
        system_language = locale.getlocale()[1]
        if system_language not in languages.keys():
            settings.setValue("appLanguage", "1252")
            settings.sync()

        self.updateTitle()
        lang = settings.value("appLanguage")

        actions = [
            ("newAction", "new"),
            ("openAction", "open"),
            ("saveAction", "save"),
            ("saveasAction", "save_as"),
            ("printAction", "print"),
            ("deleteAction", "delete"),
            ("undoAction", "undo"),
            ("redoAction", "redo"),
            ("addrowAction", "add_row"),
            ("addcolumnAction", "add_column"),
            ("addrowaboveAction", "add_row_above"),
            ("addcolumnleftAction", "add_column_left"),
            ("darklightAction", "darklight"),
            ("powersaveraction", "powersaver"),
            ("helpAction", "help"),
            ("aboutAction", "about"),
        ]

        for action_name, translation_key in actions:
            action = getattr(self, action_name)
            action.setText(translations[lang][translation_key])
            action.setStatusTip(
                translations[lang][f"{translation_key}_title"]
                if f"{translation_key}_title" in translations[lang]
                else translations[lang][translation_key]
            )

        self.translateToolbarLabel(self.file_toolbar, translations[lang]["file"])
        self.translateToolbarLabel(self.edit_toolbar, translations[lang]["edit"])
        self.translateToolbarLabel(
            self.interface_toolbar, translations[lang]["interface"]
        )
        self.translateToolbarLabel(self.formula_toolbar, translations[lang]["formula"])

        self.formula_button.setText(translations[lang]["compute"])

    def translateToolbarLabel(self, toolbar, label_key):
        self.updateToolbarLabel(
            toolbar, translations[lang].get(label_key, label_key) + ": "
        )

    def updateToolbarLabel(self, toolbar, new_label):
        for widget in toolbar.children():
            if isinstance(widget, QLabel):
                widget.setText(f"<b>{new_label}</b>")
                return

    def initDock(self):
        self.statistics_label = QLabel()
        self.dock_widget = QDockWidget("Graph Log", self)
        self.dock_widget.setObjectName("Graph Log")
        self.dock_widget.setAllowedAreas(Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)

        self.scrollableArea = QScrollArea()
        self.scrollableArea.setWidgetResizable(True)
        self.scrollableArea.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.scrollableArea.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.GraphLog_QVBox = QVBoxLayout()

        scroll_contents = QWidget()
        scroll_contents.setLayout(self.GraphLog_QVBox)
        self.scrollableArea.setWidget(scroll_contents)

        self.dock_widget.setWidget(self.scrollableArea)
        self.dock_widget.setFeatures(
            QDockWidget.NoDockWidgetFeatures | QDockWidget.DockWidgetClosable
        )

        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

    def toolbarCustomLabel(
        self,
        toolbar,
        text,
        font_size=None,
        color="#FFFFFF",
        background_color=None,
        icon_path=None,
    ):
        label = QLabel(f"<b>{text}</b>")
        label.setStyleSheet(f"font-size: {font_size}px; color: {color};")
        if background_color:
            label.setStyleSheet(
                label.styleSheet() + f"background-color: {background_color};"
            )

        if icon_path:
            icon = QIcon(icon_path)
            icon_label = QLabel()
            icon_label.setPixmap(icon.pixmap(24, 24))
            toolbar.addWidget(icon_label)

        toolbar.addWidget(label)

    def createAction(self, text, status_tip, function, shortcut=None, icon=None):
        action = QAction(text, self)
        action.setStatusTip(status_tip)
        action.triggered.connect(function)
        if shortcut:
            action.setShortcut(shortcut)
        if icon:
            action.setIcon(QIcon(icon))
        return action

    def initActions(self):
        action_definitions = [
            {
                "name": "newAction",
                "text": translations[lang]["new"],
                "status_tip": translations[lang]["new_title"],
                "function": self.new,
                "shortcut": QKeySequence.New,
            },
            {
                "name": "openAction",
                "text": translations[lang]["open"],
                "status_tip": translations[lang]["open_title"],
                "function": self.openFile,
                "shortcut": QKeySequence.Open,
            },
            {
                "name": "saveAction",
                "text": translations[lang]["save"],
                "status_tip": translations[lang]["save_title"],
                "function": self.saveFile,
                "shortcut": QKeySequence.Save,
            },
            {
                "name": "saveasAction",
                "text": translations[lang]["save_as"],
                "status_tip": translations[lang]["save_as_title"],
                "function": self.saveFileAs,
                "shortcut": QKeySequence.SaveAs,
            },
            {
                "name": "printAction",
                "text": translations[lang]["print"],
                "status_tip": translations[lang]["print_title"],
                "function": self.printSpreadsheet,
                "shortcut": QKeySequence.Print,
            },
            {
                "name": "deleteAction",
                "text": translations[lang]["delete"],
                "status_tip": translations[lang]["delete_title"],
                "function": self.cellDelete,
                "shortcut": QKeySequence.Delete,
            },
            {
                "name": "addrowAction",
                "text": translations[lang]["add_row"],
                "status_tip": translations[lang]["add_row_title"],
                "function": self.rowAdd,
                "shortcut": QKeySequence("Ctrl+Shift+R"),
            },
            {
                "name": "addcolumnAction",
                "text": translations[lang]["add_column"],
                "status_tip": translations[lang]["add_column_title"],
                "function": self.columnAdd,
                "shortcut": QKeySequence("Ctrl+Shift+C"),
            },
            {
                "name": "addrowaboveAction",
                "text": translations[lang]["add_row_above"],
                "status_tip": translations[lang]["add_row_above_title"],
                "function": self.rowAddAbove,
                "shortcut": QKeySequence("Ctrl+Shift+T"),
            },
            {
                "name": "addcolumnleftAction",
                "text": translations[lang]["add_column_left"],
                "status_tip": translations[lang]["add_column_left_title"],
                "function": self.columnAddLeft,
                "shortcut": QKeySequence("Ctrl+Shift+L"),
            },
            {
                "name": "hide_dock_widget_action",
                "text": "Graph Log",
                "status_tip": "Graph Log",
                "function": self.toggleDock,
                "shortcut": QKeySequence("Ctrl+H"),
            },
            {
                "name": "helpAction",
                "text": translations[lang]["help"],
                "status_tip": translations[lang]["help"],
                "function": self.viewHelp,
                "shortcut": QKeySequence("Ctrl+H"),
            },
            {
                "name": "aboutAction",
                "text": translations[lang]["about"],
                "status_tip": translations[lang]["about_title"],
                "function": self.viewAbout,
                "shortcut": QKeySequence.HelpContents,
            },
            {
                "name": "undoAction",
                "text": translations[lang]["undo"],
                "status_tip": translations[lang]["undo_title"],
                "function": self.undo_stack.undo,
                "shortcut": QKeySequence.Undo,
            },
            {
                "name": "redoAction",
                "text": translations[lang]["redo"],
                "status_tip": translations[lang]["redo_title"],
                "function": self.undo_stack.redo,
                "shortcut": QKeySequence.Redo,
            },
            {
                "name": "darklightAction",
                "text": translations[lang]["darklight"],
                "status_tip": translations[lang]["darklight_message"],
                "function": self.themeAction,
                "shortcut": QKeySequence("Ctrl+D"),
            },
        ]

        for action_def in action_definitions:
            setattr(
                self,
                action_def["name"],
                self.createAction(
                    action_def["text"],
                    action_def["status_tip"],
                    action_def["function"],
                    action_def["shortcut"],
                ),
            )

    def initToolbar(self):
        self.file_toolbar = self.addToolBar(translations[lang]["file"])
        self.file_toolbar.setObjectName("File")
        self.toolbarCustomLabel(
            self.file_toolbar,
            translations[lang]["file"] + ": ",
        )
        self.file_toolbar.addActions(
            [
                self.newAction,
                self.openAction,
                self.saveAction,
                self.saveasAction,
                self.printAction,
            ]
        )

        self.edit_toolbar = self.addToolBar(translations[lang]["edit"])
        self.edit_toolbar.setObjectName("Edit")
        self.toolbarCustomLabel(
            self.edit_toolbar,
            translations[lang]["edit"] + ": ",
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

        self.interface_toolbar = self.addToolBar(translations[lang]["interface"])
        self.interface_toolbar.setObjectName("Interface")
        self.toolbarCustomLabel(
            self.interface_toolbar,
            translations[lang]["interface"] + ": ",
        )

        self.theme_action = self.createAction(
            translations[lang]["darklight"],
            translations[lang]["darklight_message"],
            self.themeAction,
            QKeySequence("Ctrl+Shift+T"),
            "",
        )
        self.theme_action.setCheckable(True)
        self.theme_action.setChecked(settings.value("appTheme") == "dark")
        self.interface_toolbar.addAction(self.theme_action)

        self.powersaveraction = QAction(
            translations[lang]["powersaver"],
            self,
            checkable=True,
        )
        self.powersaveraction.setStatusTip(translations[lang]["powersaver_message"])
        self.powersaveraction.toggled.connect(self.hybridSaver)

        adaptiveResponse = settings.value(
            "adaptiveResponse", fallbackValues["adaptiveResponse"]
        )
        self.powersaveraction.setChecked(adaptiveResponse > 1)
        self.interface_toolbar.addAction(self.powersaveraction)

        self.language_combobox = QComboBox(self)
        self.language_combobox.setStyleSheet("background-color:#000000; color:#FFFFFF;")
        for lcid, name in languages.items():
            self.language_combobox.addItem(name, lcid)
        self.language_combobox.currentIndexChanged.connect(self.changeLanguage)
        self.interface_toolbar.addWidget(self.language_combobox)

        self.addToolBarBreak()

        self.formula_toolbar = self.addToolBar(translations[lang]["formula"])
        self.formula_toolbar.setObjectName("Formula")
        self.toolbarCustomLabel(
            self.formula_toolbar,
            translations[lang]["formula"] + ": ",
        )

        self.formula_edit = QLineEdit()
        self.formula_edit.setPlaceholderText(translations[lang]["formula"])
        self.formula_edit.returnPressed.connect(self.computeFormula)

        self.formula_button = QPushButton(translations[lang]["compute"])
        self.formula_button.setStyleSheet(
            """
            QPushButton {
                background-color: #A72461; 
                color: #FFFFFF; 
                font-weight: bold; 
                padding: 10px; 
                border-radius: 10px; 
                border: 1px solid #000000; 
                margin-left: 10px;
            }
            QPushButton:hover {
                background-color: #C7357A; 
                border: 1px solid #FFDDDD; 
            }
            """
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

        self.interface_toolbar.addAction(self.hide_dock_widget_action)
        self.interface_toolbar.addAction(self.helpAction)
        self.interface_toolbar.addAction(self.aboutAction)

    def changeColumnName(self, column):
        current_name = (
            self.SpreadsheetArea.horizontalHeaderItem(column).text()
            if self.SpreadsheetArea.horizontalHeaderItem(column)
            else ""
        )

        if not current_name or current_name.lower() == "unnamed":
            msg_text = "This column is unnamed. Would you like to give it a new name?"
        else:
            msg_text = f"What would you like to do with the '{current_name}' column?"

        ccn_msg = QMessageBox(self)
        ccn_msg.setText(msg_text)

        rename_button = ccn_msg.addButton("Rename", QMessageBox.ActionRole)
        select_button = ccn_msg.addButton("Select All", QMessageBox.ActionRole)
        cancel_button = ccn_msg.addButton(QMessageBox.Cancel)
        ccn_msg.exec()

        first_cell_item = self.SpreadsheetArea.item(0, column)
        if first_cell_item is None or first_cell_item.text() == "":
            self.SpreadsheetArea.setItem(0, column, QTableWidgetItem(f"{column + 1}"))
            QTimer.singleShot(0, self.resizeTable)

        if ccn_msg.clickedButton() == cancel_button:
            return

        if ccn_msg.clickedButton() == rename_button:
            new_name, ok = QInputDialog.getText(
                self,
                "Change Column Name",
                "New column name:",
                text=current_name,
            )

            if not ok or not new_name or len(new_name.strip()) == 0:
                QMessageBox.warning(
                    self, "Invalid Input", "Please enter a valid column name."
                )
                return

            max_length = 100
            if len(new_name) > max_length:
                QMessageBox.warning(
                    self,
                    "Invalid Input",
                    f"Column name cannot be longer than {max_length} characters.",
                )
                return

            if not new_name.isalnum() and " " not in new_name:
                QMessageBox.warning(
                    self,
                    "Invalid Input",
                    "Column name can only contain letters, numbers, and spaces.",
                )
                return

            self.SpreadsheetArea.setHorizontalHeaderItem(
                column, QTableWidgetItem(new_name)
            )

            first_cell_item = self.SpreadsheetArea.item(0, column)
            if first_cell_item is None or first_cell_item.text() == "":
                self.SpreadsheetArea.setItem(0, column, QTableWidgetItem(f"{new_name}"))

            elif first_cell_item:
                first_cell_item.setText(new_name)

            QTimer.singleShot(0, self.resizeTable)

        elif ccn_msg.clickedButton() == select_button:
            row_count = self.SpreadsheetArea.rowCount()

            for row in range(row_count):
                item = self.SpreadsheetArea.item(row, column)
                if item is None:
                    item = QTableWidgetItem()
                    self.SpreadsheetArea.setItem(row, column, item)

                item.setSelected(True)

    def createDockwidget(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        self.statistics = QLabel("Statistics")
        layout.addWidget(self.statistics)
        return widget

    def toggleDock(self):
        self.dock_widget.setVisible(not self.dock_widget.isVisible())

    def hybridSaver(self, checked):
        if checked:
            battery = psutil.sensors_battery()
            if battery:
                self.adaptiveResponse = (
                    12 if battery.percent <= 35 and not battery.power_plugged else 6
                )
            else:
                self.adaptiveResponse = 3  # Global Standard
        else:
            self.adaptiveResponse = fallbackValues["adaptiveResponse"]

        settings.setValue("adaptiveResponse", self.adaptiveResponse)
        settings.sync()

    def new(self):
        if self.is_saved:
            self.resetSpreadsheet()
        else:
            reply = QMessageBox.question(
                self,
                app.applicationDisplayName(),
                translations[lang]["new_title"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )

            if reply == QMessageBox.Yes:
                self.resetSpreadsheet()

    def resetSpreadsheet(self):
        self.SpreadsheetArea.clearContents()

        self.SpreadsheetArea.setRowCount(0)
        self.SpreadsheetArea.setColumnCount(0)

        self.SpreadsheetArea.setHorizontalHeaderLabels([])
        self.SpreadsheetArea.setVerticalHeaderLabels([])

        self.SpreadsheetArea.setRowCount(50)
        self.SpreadsheetArea.setColumnCount(100)

        default_headers = [
            str(i + 1) for i in range(self.SpreadsheetArea.columnCount())
        ]
        self.SpreadsheetArea.setHorizontalHeaderLabels(default_headers)

        QTimer.singleShot(0, self.resizeTable)

        self.directory = self.default_directory
        self.file_name = None
        self.is_saved = False

        self.updateStatistics()
        self.updateTitle()

    def resizeTable(self):
        self.SpreadsheetArea.resizeColumnsToContents()
        self.SpreadsheetArea.resizeRowsToContents()

    def openFile(self):
        if not self.is_saved:
            reply = QMessageBox.question(
                self,
                app.applicationDisplayName(),
                translations[lang]["open"],
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self.saveState()

        return self.openFileProcess()

    def openFileProcess(self):
        options = QFileDialog.Options()
        selected_file, _ = QFileDialog.getOpenFileName(
            self,
            f"{translations[lang]['open_title']} — {app.applicationDisplayName()}",
            self.directory,
            fallbackValues["readFilter"],
            options=options,
        )

        if selected_file:
            self.resetSpreadsheet()
            self.file_name = selected_file

            if self.file_name.endswith(".xlsx"):
                try:
                    self.loadSpreadsheetFromExcel(self.file_name)
                except:
                    QMessageBox.warning(self, None, "Conversion failed.")
                    raise
            else:
                if (
                    self.file_name.endswith(".ssfs")
                    or self.file_name.endswith(".xsrc")
                    or self.file_name.endswith(".csv")
                ):
                    self.loadSpreadsheet(self.file_name)
            self.directory = os.path.dirname(self.file_name)
            self.is_saved = True
            self.updateStatistics()
            self.updateTitle()

    def loadSpreadsheetFromExcel(self, file_name):
        try:
            read_file = pd.read_excel(file_name)

            new_headers = []
            for i, col in enumerate(read_file.columns):
                if col.startswith("Unnamed") or pd.isnull(col):
                    new_headers.append(str(i + 1))
                else:
                    new_headers.append(col)

            read_file.columns = new_headers
            read_file.to_csv(f"{file_name}.ssfs", index=None, header=True)
            self.loadSpreadsheet(f"{file_name}.ssfs")
            os.remove(f"{file_name}.ssfs")
        except Exception as e:
            QMessageBox.warning(self, None, f"Conversion failed: {e}")

    def loadSpreadsheetFromOrigin(self, file_name):
        if file_name.endswith((".ssfs", ".xsrc", ".csv")):
            self.loadSpreadsheet(file_name)

    def loadSpreadsheet(self, file_path):
        with open(file_path, "r", encoding="utf-8") as file:
            reader = csv.reader(file)
            data = list(reader)

            if data:
                column_headers = data[0]

                for i in range(len(column_headers)):
                    if not column_headers[i].strip():
                        column_headers[i] = f"{i + 1}"

                self.SpreadsheetArea.setHorizontalHeaderLabels(column_headers)

                self.SpreadsheetArea.setRowCount(len(data))
                self.SpreadsheetArea.setColumnCount(len(column_headers))

                for row in range(len(data)):
                    for column in range(len(data[row])):
                        item = QTableWidgetItem(data[row][column])
                        self.SpreadsheetArea.setItem(row, column, item)

        self.SpreadsheetArea.resizeColumnsToContents()
        self.SpreadsheetArea.resizeRowsToContents()

    def saveFile(self):
        if self.file_name is None:
            self.saveFileAs()
        elif not self.is_saved:
            self.saveFileProcess()
        else:
            self.status_bar.showMessage("No changes to save.", 2000)

    def saveFileAs(self):
        options = QFileDialog.Options()
        selected_file, _ = QFileDialog.getSaveFileName(
            self,
            f"{translations[lang]['save_as_title']} — {app.applicationDisplayName()}",
            self.directory,
            fallbackValues["writeFilter"],
            options=options,
        )
        if selected_file:
            self.file_name = selected_file
            self.directory = os.path.dirname(self.file_name)
            self.saveFileProcess()

    def saveFileProcess(self):
        if not self.file_name:
            QMessageBox.warning(self, "Warning", "No file name specified.")
            return

        file_extension = os.path.splitext(self.file_name)[1]
        save_method = {
            ".xlsx": self.saveAsExcel,
            ".ssfs": self.saveAsSSFS,
            ".csv": self.saveAsCSV,
        }.get(file_extension, self.saveAsCSV)

        try:
            save_method()
            self.is_saved = True
            self.updateTitle()
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save file: {e}")

    def cellData(self):
        return [
            [
                (
                    self.SpreadsheetArea.item(rowid, colid).text()
                    if self.SpreadsheetArea.item(rowid, colid)
                    else ""
                )
                for colid in range(self.SpreadsheetArea.columnCount())
            ]
            for rowid in range(self.SpreadsheetArea.rowCount())
        ]

    def saveAsExcel(self):
        try:
            pd.DataFrame(self.cellData()).to_excel(
                self.file_name, index=False, header=False
            )
            self.status_bar.showMessage("Saved as Excel file.", 2000)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save as Excel: {e}")

    def saveAsSSFS(self):
        try:
            with open(self.file_name, "w", newline="", encoding="utf-8") as file:
                csv.writer(file).writerows(self.cellData())
            self.status_bar.showMessage("Saved as SSFS file.", 2000)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save as SSFS: {e}")

    def saveAsCSV(self):
        try:
            with open(self.file_name, "w", newline="", encoding="utf-8") as file:
                csv.writer(file).writerows(self.cellData())
            self.status_bar.showMessage("Saved as CSV file.", 2000)
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to save as CSV: {e}")

    def printSpreadsheet(self):
        selected_ranges = self.SpreadsheetArea.selectedRanges()

        if not selected_ranges:
            QMessageBox.warning(self, None, translations[lang]["print_warning"])
            return

        printer = self.setupPrinter()
        preview_dialog = QPrintPreviewDialog(printer, self)
        preview_dialog.paintRequested.connect(self.printSelectedCells)
        preview_dialog.exec()

    def setupPrinter(self):
        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageOrientation(QPageLayout.Orientation.Landscape)
        printer.setPageMargins(QMargins(0, 0, 0, 0), QPageLayout.Millimeter)
        printer.setFullPage(True)
        printer.setDocName(self.file_name)
        return printer

    def printSelectedCells(self, printer):
        painter = QPainter(printer)
        rect = painter.viewport().adjusted(20, 20, -20, -20)

        for range_ in self.SpreadsheetArea.selectedRanges():
            self.paintSelectedRange(painter, rect, range_)

        painter.end()

    def paintSelectedRange(self, painter, rect, range_):
        for row in range(range_.topRow(), range_.bottomRow() + 1):
            for column in range(range_.leftColumn(), range_.rightColumn() + 1):
                item = self.SpreadsheetArea.item(row, column)
                if item:
                    self.drawCell(painter, rect, row, column, item.text())

    def drawCell(self, painter, rect, row, column, text):
        cell_width = self.SpreadsheetArea.columnWidth(column)
        cell_height = self.SpreadsheetArea.rowHeight(row)

        x = rect.x() + (column - rect.left()) * cell_width
        y = rect.y() + (row - rect.top()) * cell_height

        painter.fillRect(QRect(x, y, cell_width, cell_height), Qt.lightGray)
        painter.drawRect(QRect(x, y, cell_width, cell_height))

        painter.setPen(Qt.black)
        painter.setFont(QFont("Arial", 10))
        painter.drawText(QRect(x, y, cell_width, cell_height), Qt.AlignCenter, text)

    def viewAbout(self):
        viewAbout = SS_About(self)
        viewAbout.show()

    def viewHelp(self):
        help_window = SS_Help(self)
        help_window.show()

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
                    if item:
                        text = item.text().strip()
                        if text.replace(".", "", 1).isdigit() and text.count(".") < 2:
                            try:
                                value = float(text)
                                values.append(value)
                            except ValueError:
                                continue

        return values

    def controlCacheDir(self):
        if platform.system() == "Windows":
            cache_dir = os.path.join(self.directory, ".sscache")
        else:
            cache_dir = os.path.join(self.directory, ".sscache")

        if not os.path.exists(cache_dir):
            os.makedirs(cache_dir)

            if platform.system() == "Windows":
                import ctypes

                ctypes.windll.kernel32.SetFileAttributesW(cache_dir, 2)  # Hidden Folder

        return cache_dir

    def computeFormula(self):
        QTimer.singleShot(25, self.processFormula)

    def processFormula(self):
        try:
            formulavalue = self.formula_edit.text().strip()
            if not formulavalue:
                QMessageBox.warning(self, "Input Error", "ERROR: No formula provided.")
                return

            formula = formulavalue.split()[0]

            if formula not in formulas:
                QMessageBox.warning(self, "Input Error", f"ERROR: {formula}")
                return

            operands = self.selectedCells()

            if not operands:
                QMessageBox.warning(self, "Input Error", "ERROR: No operands selected.")
                return

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
            elif formula in graphformulas:
                start_elapsed = datetime.datetime.now()
                pltgraph = plt.figure()
                pltgraph.suptitle(graphformulas[formula])
                x_label, y_label = self.getGraphLabels(formula, operands)

                if formula == "similargraph":
                    plt.plot(operands)
                elif formula == "pointgraph":
                    plt.plot(operands, "o")
                elif formula == "bargraph":
                    plt.bar(range(len(operands)), operands)
                elif formula == "piegraph":
                    plt.pie(operands)
                    y_label = "Percentage"
                elif formula == "histogram":
                    plt.hist(operands)
                    y_label = "Frequency"

                plt.xlabel(x_label)
                plt.ylabel(y_label)
                plt.grid(True)
                datetime_string = QDateTime.currentDateTimeUtc().toString(
                    "yyyy-MM-dd HH:mm:ss"
                )
                utc_timestamp = (
                    datetime.datetime.now(timezone.utc)
                    .replace(tzinfo=timezone.utc)
                    .timestamp()
                )
                cache_dir = self.controlCacheDir()
                graph_path = os.path.join(
                    cache_dir, f"solidsheets_G{utc_timestamp}.png"
                )
                plt.savefig(graph_path)
                plt.close()
                result = "Graph"
                end_elasped = datetime.datetime.now()

                graph_container = QWidget()
                graph_layout = QVBoxLayout(graph_container)
                graph_layout.setContentsMargins(10, 10, 10, 10)

                graph_label = QLabel()
                graph_label.setPixmap(QPixmap(graph_path))
                graph_label.setScaledContents(True)
                graph_layout.addWidget(graph_label)

                date_label = QLabel(
                    f"{datetime_string} ({str((end_elasped - start_elapsed).total_seconds())} sec)"
                )
                date_label.setAlignment(Qt.AlignCenter)
                graph_layout.addWidget(date_label)

                button_layout = QGridLayout()

                save_button = QPushButton("Save")
                save_button.clicked.connect(lambda: self.saveGraph(graph_path))
                save_button.setStyleSheet(
                    """
                    QPushButton {
                        background-color: #6200EE; 
                        color: white; 
                        font-weight: bold; 
                        padding: 10px; 
                        border-radius: 5px; 
                        border: none;
                    }
                    QPushButton:hover {
                        background-color: #3700B3;
                    }
                """
                )
                save_button.setCursor(QCursor(Qt.PointingHandCursor))

                button_layout.addWidget(save_button, 0, 0)

                delete_button = QPushButton("Delete")
                delete_button.clicked.connect(
                    lambda: self.deleteGraph(graph_container, graph_path)
                )
                delete_button.setStyleSheet(
                    """
                    QPushButton {
                        background-color: #D50000; 
                        color: white; 
                        font-weight: bold; 
                        padding: 10px; 
                        border-radius: 5px; 
                        border: none;
                    }
                    QPushButton:hover {
                        background-color: #9B0000;
                    }
                """
                )
                delete_button.setCursor(QCursor(Qt.PointingHandCursor))

                button_layout.addWidget(delete_button, 0, 1)

                graph_layout.addLayout(button_layout)

                self.GraphLog_QVBox.insertWidget(0, graph_container)

                graph_count = self.GraphLog_QVBox.count()
                self.dock_widget.setWindowTitle(f"Graph Log ({graph_count})")
                self.dock_widget.setVisible(True)

            if result != "Graph":
                QMessageBox.information(self, "Formula", f"{formula} : {result}")

        except ValueError as valueerror:
            QMessageBox.critical(self, "Formula", str(valueerror))
        except Exception as exception:
            QMessageBox.critical(self, "Formula", str(exception))

    def getGraphLabels(self, formula, operands):
        column_headers = self.getColumnHeadersForSelectedCells()

        x_label = "Index"
        y_label = "Value"

        if formula in ["similargraph", "pointgraph", "bargraph"]:
            if len(column_headers) >= 2:
                x_label = f"X"
                y_label = f"Y"
            else:
                x_label = "X"
                y_label = "Y"

        elif formula == "piegraph":
            if column_headers:
                x_label = f"X"
                y_label = "Percentage"
            else:
                x_label = "Categories"
                y_label = "Percentage"

        elif formula == "histogram":
            if column_headers:
                x_label = f"X"
                y_label = "Frequency"
            else:
                x_label = "Categories"
                y_label = "Frequency"

        if operands:
            if len(operands) > 1:
                x_label = f"X"
                y_label = f"Y"

        return x_label, y_label

    def getColumnHeadersForSelectedCells(self):
        column_headers = []

        for col in range(self.SpreadsheetArea.columnCount()):
            header = self.SpreadsheetArea.horizontalHeaderItem(col)
            if header:
                column_headers.append(header.text())
        return column_headers

    def saveGraph(self, filepath):
        options = QFileDialog.Options()
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save Graph",
            self.directory + "/" + filepath,
            fallbackValues["graphSaveFilter"],
            options=options,
        )
        if filename:
            QPixmap(filepath).save(filename)

    def deleteGraph(self, graph_container, filepath):
        try:
            self.GraphLog_QVBox.removeWidget(graph_container)
            graph_container.deleteLater()

            self.dock_widget.update()
            graph_count = self.GraphLog_QVBox.count()
            self.dock_widget.setWindowTitle(f"Graph Log ({graph_count})")
            if os.path.exists(filepath):
                os.remove(filepath)
            else:
                QMessageBox.warning(self, None, "Graph cache not found.")
        except Exception as e:
            QMessageBox.critical(self, "Delete Graph", str(e))

    def cellDelete(self):
        for item in self.SpreadsheetArea.selectedItems():
            item.setText("")

    def rowAdd(self):
        self.SpreadsheetArea.insertRow(self.SpreadsheetArea.rowCount())

    def rowAddAbove(self):
        current_row = self.SpreadsheetArea.currentRow()

        if current_row == 0:
            QMessageBox.warning(
                self, "Warning", "Cannot add a row above the first row."
            )
        else:
            self.SpreadsheetArea.insertRow(current_row)

    def columnAdd(self):
        self.SpreadsheetArea.insertColumn(self.SpreadsheetArea.columnCount())

    def columnAddLeft(self):
        self.SpreadsheetArea.insertColumn(self.SpreadsheetArea.currentColumn())


if __name__ == "__main__":
    if getattr(sys, "frozen", False):
        applicationPath = sys._MEIPASS
    else:
        applicationPath = os.path.dirname(__file__)
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(os.path.join(applicationPath, fallbackValues["icon"])))
    app.setOrganizationName("berkaygediz")
    app.setApplicationName("SolidSheets")
    app.setApplicationDisplayName("SolidSheets 2024.11")
    app.setApplicationVersion("1.5.2024.11-1")
    wb = SS_ControlInfo()
    wb.show()
    sys.exit(app.exec())
