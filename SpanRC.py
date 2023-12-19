import os
import sys
import csv
import datetime
import time
import platform
import sqlite3
import pdb
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtTest import *
from PyQt5.QtCore import *
from PyQt5.QtPrintSupport import *
from modules.translations import *
from modules.secrets import *

# DEBUG ONLY VERSION - DO NOT USE IN PRODUCTION

default_values = {
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
    'last_opened_file': None,
    'currentRow': 0,
    'currentColumn': 0,
    'rowCount': 50,
    'columnCount': 100,
    'windowState': None,
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


class SRC_ControlInfo(QMainWindow, QThread):
    def __init__(self, parent=None):
        super(SRC_ControlInfo, self).__init__(parent)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        settings = QSettings("berkaygediz", "SpanRC")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint)
        screen = QGuiApplication.primaryScreen().availableGeometry()
        self.setGeometry(
            QStyle.alignedRect(
                Qt.LayoutDirection.LeftToRight,
                Qt.AlignmentFlag.AlignCenter,
                QSize(int(screen.width() * 0.2),
                      int(screen.height() * 0.2)),
                screen
            )
        )
        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()
        self.language = settings.value("current_language")
        self.setStyleSheet("background-color: transparent;")
        self.setWindowOpacity(0.75)
        self.widget_central = QWidget()
        self.layout_central = QVBoxLayout(self.widget_central)
        self.widget_central.setStyleSheet(
            "background-color: #6F61C0; border-radius: 50px; border: 1px solid #A084E8; border-radius: 10px;")
        self.title = QLabel("SpanRC", self)
        self.title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.title.setFont(QFont('Roboto', 30))
        self.title.setStyleSheet(
            "background-color: #A084E8; color: #FFFFFF; font-weight: bold; font-size: 30px; border-radius: 25px; border: 1px solid #000000;")
        self.layout_central.addWidget(self.title)
        self.label_status = QLabel("...", self)
        self.label_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.label_status.setFont(QFont('Roboto', 10))
        self.layout_central.addWidget(self.label_status)
        self.setCentralWidget(self.widget_central)
        if serverconnect():
            sqlite_file = os.path.join(os.path.dirname(
                os.path.abspath(__file__)), 'src.db')
            sqlitedb = sqlite3.connect(sqlite_file)
            sqlitecursor = sqlitedb.cursor()
            sqlitecursor.execute(
                "CREATE TABLE IF NOT EXISTS profile (email TEXT, password TEXT)")
            sqlitecursor.execute(
                "CREATE TABLE IF NOT EXISTS apps (spanrc INTEGER DEFAULT 0, email TEXT)")
            sqlitecursor.execute(
                "CREATE TABLE IF NOT EXISTS log (email TEXT, devicename TEXT, log TEXT, logdate TIMESTAMP DEFAULT CURRENT_TIMESTAMP)")

            profile_email = sqlitecursor.execute(
                "SELECT email FROM profile").fetchone()
            profile_pw = sqlitecursor.execute(
                "SELECT password FROM profile").fetchone()

            if profile_email is not None and profile_pw is not None:
                local_email = profile_email[0]
                local_pw = profile_pw[0]
                mysql_connection = serverconnect()
                if mysql_connection:
                    mysqlcursor = mysql_connection.cursor()
                    mysqlcursor.execute(
                        "SELECT * FROM profile WHERE email = %s", (local_email,))
                    user_result = mysqlcursor.fetchone()

                    if user_result is not None:
                        if local_email == user_result[1] and local_pw == user_result[2]:
                            SRC_Workbook().show()
                            sqlitecursor.execute(
                                "INSERT INTO log (email, devicename, log) VALUES (?, ?, ?)", (local_email, platform.node(), "Login Successful!"))
                            mysqlcursor.execute(
                                "INSERT INTO log (email, devicename, log) VALUES (%s, %s, %s)", (local_email, platform.node(), "Login Successful!"))
                            sqlitecursor.connection.commit()
                            mysql_connection.commit()
                            if settings.value("current_language") is None:
                                settings.setValue(
                                    "current_language", "English")
                                settings.sync()
                            self.label_status.setStyleSheet(
                                "background-color: #7900FF; color: #FFFFFF; font-weight: bold; font-size: 16px; border-radius: 30px; border: 1px solid #000000;")
                            self.label_status.setText(
                                "💡" + local_email.split('@')[0])
                            QTimer.singleShot(2500, self.hide)

                        else:
                            self.label_status.setText(
                                translations[settings.value("current_language")]["wrong_password"])
                            sqlitecursor.execute(
                                "INSERT INTO log (email, devicename, log) VALUES (?, ?, ?)", (local_email, platform.node(), "Wrong Password!"))
                            mysqlcursor.execute(
                                "INSERT INTO log (email, devicename, log) VALUES (%s, %s, %s)", (local_email, platform.node(), "Wrong Password!"))
                            mysql_connection.commit()
                            sqlitecursor.execute(
                                "DROP TABLE IF EXISTS profile")
                            sqlitecursor.execute("DROP TABLE IF EXISTS apps")
                            sqlitecursor.execute("DROP TABLE IF EXISTS log")
                            sqlitecursor.connection.commit()
                            SRC_ControlInfo().show()
                            QTimer.singleShot(2500, self.hide)
                    else:
                        self.label_status.setStyleSheet(
                            "background-color: #252525; color: #FFFFFF; font-weight: bold; font-size: 16px; border-radius: 30px; border: 1px solid #000000;")
                        self.label_status.setText(
                            translations[settings.value("current_language")]["no_account"])
                        logout_register = QPushButton(translations[settings.value(
                            "current_language")]["register"])
                        logout_register.setStyleSheet(
                            "background-color: #7900FF; color: #FFFFFF; font-weight: bold; font-size: 16px; border-radius: 30px; border: 1px solid #000000;")
                        logout_register.clicked.connect(SRC_Welcome().show)
                        self.layout_central.addWidget(logout_register)
                        self.close_button = QPushButton(translations[settings.value(
                            "current_language")]["exit"])
                        self.close_button.setStyleSheet(
                            "background-color: #7900FF; color: #FFFFFF; font-weight: bold; font-size: 16px; border-radius: 30px; border: 1px solid #000000;")
                        self.close_button.clicked.connect(sys.exit)
                        self.layout_central.addWidget(self.close_button)

                else:
                    self.label_status.setText(
                        translations[settings.value("current_language")]["connection_denied"])
            else:
                SRC_Welcome().show()
                QTimer.singleShot(0, self.hide)
        else:
            self.label_status.setText(translations[settings.value(
                "current_language")]["connection_denied"])
            QTimer.singleShot(10000, sys.exit)


class SRC_Welcome(QMainWindow):
    def __init__(self, parent=None):
        super(SRC_Welcome, self).__init__(parent)
        starttime = datetime.datetime.now()
        settings = QSettings("berkaygediz", "SpanRC")
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        screen = QGuiApplication.primaryScreen().availableGeometry()
        self.setGeometry(
            QStyle.alignedRect(
                Qt.LayoutDirection.LeftToRight,
                Qt.AlignmentFlag.AlignCenter,
                QSize(int(screen.width() * 0.4),
                      int(screen.height() * 0.5)),
                screen
            )
        )

        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()
        self.language = settings.value("current_language")

        self.setStyleSheet("background-color: #F9E090;")
        self.setWindowTitle(translations[settings.value(
            "current_language")]["welcome-title"])
        introduction = QVBoxLayout()
        self.statusBar().setStyleSheet(
            "background-color: #F9E090; color: #000000; font-weight: bold; font-size: 12px; border-radius: 10px; border: 1px solid #000000;")
        product_label = QLabel("SpanRC 🎉")
        product_label.setStyleSheet(
            "background-color: #DFBB9D; color: #000000; font-size: 28px; font-weight: bold; border-radius: 10px; border: 1px solid #000000;")
        product_label.setAlignment(Qt.AlignmentFlag.AlignCenter)

        intro_label = QLabel(translations[settings.value(
            "current_language")]["intro"])
        intro_label.setStyleSheet(
            "background-color: #F7E2D6; color: #000000; font-size: 12px; border-radius: 10px; border: 1px solid #000000;")
        intro_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        intro_label.setFixedHeight(50)

        def SRC_changeLanguage():
            language = self.language_combobox.currentText()
            settings = QSettings("berkaygediz", "SpanRC")
            settings.setValue("current_language", language)
            settings.sync()

        self.language_combobox = QComboBox()
        self.language_combobox.setStyleSheet(
            "background-color: #B3E8E5; color: #000000; border-radius: 10px; border: 1px solid #000000; margin: 10px; margin-left: 0px; margin-right: 0px; padding: 0px; font-size: 16px;")
        self.language_combobox.addItems(
            ["English", "Türkçe", "Azərbaycanca", "Deutsch", "Español"])
        self.language_combobox.currentTextChanged.connect(SRC_changeLanguage)
        self.language_combobox.setCurrentText(
            settings.value("current_language"))

        introduction.addWidget(product_label)
        introduction.addWidget(intro_label)
        introduction.addWidget(self.language_combobox)
        label_email = QLabel(
            translations[settings.value("current_language")]["email"] + ":")
        label_email.setStyleSheet(
            "color: #A72461; font-weight: bold; font-size: 16px; margin: 10px; margin-left: 0px; margin-right: 0px; padding: 0px;")
        textbox_email = QLineEdit()
        textbox_email.setStyleSheet(
            "background-color: #B3E8E5; color: #000000; border-radius: 10px; border: 1px solid #000000; margin: 10px; margin-left: 0px; margin-right: 0px; padding: 0px; font-size: 16px;")

        label_password = QLabel(
            translations[settings.value("current_language")]["password"] + ":")
        label_password.setStyleSheet(
            "color: #A72461; font-weight: bold; font-size: 16px; margin: 10px; margin-left: 0px; margin-right: 0px; padding: 0px;")
        textbox_password = QLineEdit()
        textbox_password.setStyleSheet(
            "background-color: #B3E8E5; color: #000000; border-radius: 10px; border: 1px solid #000000; margin: 10px; margin-left: 0px; margin-right: 0px; padding: 0px; font-size: 16px;")
        textbox_password.setEchoMode(QLineEdit.EchoMode.Password)

        label_exception = QLabel()
        label_exception.setStyleSheet(
            "color: #A72461; font-weight: bold; font-size: 16px; margin: 10px; margin-left: 0px; margin-right: 0px; padding: 0px;")
        label_exception.setAlignment(Qt.AlignmentFlag.AlignCenter)

        bottom_layout = QVBoxLayout()
        bottom_layout.addWidget(label_email)
        bottom_layout.addWidget(textbox_email)
        bottom_layout.addWidget(label_password)
        bottom_layout.addWidget(textbox_password)
        bottom_layout.addWidget(label_exception)

        button_layout = QHBoxLayout()

        button_login = QPushButton(
            translations[settings.value("current_language")]["login"])
        button_login.setStyleSheet(
            "background-color: #A72461; color: #FFFFFF; font-weight: bold; padding: 10px; border-radius: 10px; border: 1px solid #000000; margin: 10px;")

        def register():
            if textbox_email.text() == "" or textbox_password.text() == "":
                error = QMessageBox()
                error.setWindowTitle(translations[settings.value(
                    "current_language")]["error"])
                error.setText(translations[settings.value(
                    "current_language")]["fill_all"])
                error.setIcon(QMessageBox.Warning)
                error.setStandardButtons(QMessageBox.Ok)
                error.exec_()
            else:
                mysqlserver = serverconnect()
                mysqlcursor = mysqlserver.cursor()
                sqlite_file = os.path.join(os.path.dirname(
                    os.path.abspath(__file__)), 'src.db')
                sqliteConnection = sqlite3.connect(sqlite_file)
                mysqlcursor.execute(
                    "SELECT * FROM profile WHERE email = %s", (textbox_email.text(),))
                user_result = mysqlcursor.fetchone()

                if user_result == None:
                    mysqlcursor.execute("INSERT INTO profile (email, password) VALUES (%s, %s)", (
                        textbox_email.text(), textbox_password.text()))
                    mysqlserver.commit()
                    mysqlcursor.execute(
                        "INSERT INTO log (devicename, log) VALUES (%s, %s)", (platform.node(), "Register Successful!"))
                    mysqlserver.commit()
                    mysqlcursor.execute(
                        "INSERT INTO apps (spanrc, email) VALUES (%s, %s)", (0, textbox_email.text()))
                    sqliteConnection.execute("DELETE FROM profile")
                    sqliteConnection.execute("INSERT INTO profile (email, password) VALUES (?, ?)", (
                        textbox_email.text(), textbox_password.text()))
                    sqliteConnection.connection.commit()
                    label_exception.setText(translations[settings.value(
                        "current_language")]["register_success"])
                else:
                    label_exception.setText(
                        translations[settings.value("current_language")]["already_registered"])

        def login():
            if textbox_email.text() == "" or textbox_password.text() == "":
                if settings.value("current_language") == None:
                    settings.setValue("current_language", "English")
                    settings.sync()
                QMessageBox.warning(
                    self, translations[settings.value("current_language")]["error"], translations[settings.value("current_language")]["fill_all"])
            else:
                mysqlserver = serverconnect()
                mysqlcursor = mysqlserver.cursor()
                sqlite_file = os.path.join(os.path.dirname(
                    os.path.abspath(__file__)), 'src.db')
                sqliteConnection = sqlite3.connect(sqlite_file).cursor()
                mysqlcursor.execute(
                    "SELECT * FROM profile WHERE email = %s", (textbox_email.text(),))
                user_result = mysqlcursor.fetchone()
                mysqlcursor.execute(
                    "SELECT * FROM profile WHERE password = %s", (textbox_password.text(),))
                pw_result = mysqlcursor.fetchone()

                if user_result and pw_result:
                    mysqlcursor.execute(
                        "INSERT INTO log (email, devicename, log) VALUES (%s, %s, %s)", (textbox_email.text(), platform.node(), "Login Successful!"))
                    mysqlserver.commit()
                    sqliteConnection.execute("DROP TABLE IF EXISTS profile")
                    sqliteConnection.connection.commit()
                    sqliteConnection.execute(
                        "CREATE TABLE IF NOT EXISTS profile (email TEXT, password TEXT)")
                    sqliteConnection.connection.commit()
                    sqliteConnection.execute("INSERT INTO profile (email, password) VALUES (?, ?)", (
                        textbox_email.text(), textbox_password.text()))
                    sqliteConnection.connection.commit()
                    time.sleep(1)
                    textbox_email.setText("")
                    textbox_password.setText("")
                    label_exception.setText("")
                    SRC_Workbook().show()
                    QTimer.singleShot(1000, self.hide)
                else:
                    label_exception.setText(
                        translations[settings.value("current_language")]["wrong_credentials"])

        button_login.clicked.connect(login)
        button_register = QPushButton(translations[settings.value(
            "current_language")]["register"])
        button_register.setStyleSheet(
            "background-color: #A72461; color: #FFFFFF; font-weight: bold; padding: 10px; border-radius: 10px; border: 1px solid #000000; margin: 10px;")
        button_register.clicked.connect(register)

        button_layout.addWidget(button_login)
        button_layout.addWidget(button_register)

        introduction.addLayout(bottom_layout)
        introduction.addLayout(button_layout)

        central_widget = QWidget()
        central_widget.setLayout(introduction)

        self.setCentralWidget(central_widget)
        endtime = datetime.datetime.now()
        self.statusBar().showMessage(
            str((endtime - starttime).total_seconds()) + " ms", 2500)


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
        # press c to continue (initialization bug)
        pdb.set_trace()
        print("1. SRC_Workbook init")
        starttime = datetime.datetime.now()
        settings = QSettings("berkaygediz", "SpanRC")
        self.setWindowIcon(QIcon("icon.png"))
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        sqlite_file = os.path.join(os.path.dirname(
            os.path.abspath(__file__)), 'src.db')
        sqlitedb = sqlite3.connect(sqlite_file)
        sqlitecursor = sqlitedb.cursor()
        user_email = sqlitecursor.execute(
            "SELECT email FROM profile").fetchone()[0]
        mysqlserver = serverconnect()
        mysqlcursor = mysqlserver.cursor()
        print("2. SRC_Workbook - database process")
        auth = mysqlcursor.execute(
            "SELECT spanrc FROM apps WHERE spanrc = '1' AND email = %s", (user_email,))
        auth = mysqlcursor.fetchone()
        sqlitedb.close()
        mysqlserver.close()
        if auth is not None:
            self.src_thread = SRC_Threading()
            self.src_thread.update_signal.connect(
                self.SRC_updateStatistics)
            self.SRC_themePalette()
            self.undo_stack = QUndoStack(self)
            self.undo_stack.setUndoLimit(100)
            self.selected_file = None
            self.file_name = None
            self.is_saved = None
            self.default_directory = QDir().homePath()
            self.directory = self.default_directory
            if settings.value("current_language") == None:
                settings.setValue("current_language", "English")
                settings.sync()
            self.language = settings.value("current_language")
            if settings.value("current_theme") == None:
                settings.setValue("current_theme", "light")
                settings.sync()
            self.theme = settings.value("current_theme")
            self.SRC_setupDock()
            self.dock_widget.hide()
            self.status_bar = self.statusBar()
            self.src_table = QTableWidget(self)
            self.setCentralWidget(self.src_table)
            self.SRC_setupActions()
            self.SRC_setupToolbar()
            self.setPalette(self.light_theme)
            self.src_table.itemSelectionChanged.connect(
                self.src_thread.start)
            self.src_table.setCursor(Qt.CursorShape.SizeAllCursor)
            self.showMaximized()
            self.setFocus()
            QTimer.singleShot(50, self.SRC_restoreTheme)
            QTimer.singleShot(150, self.SRC_restoreState)
            if self.src_table.columnCount() == 0 and self.src_table.rowCount() == 0:
                self.src_table.setColumnCount(100)
                self.src_table.setRowCount(50)
                self.src_table.clearSpans()
                self.src_table.setItem(0, 0, QTableWidgetItem(''))

            self.SRC_updateTitle()
            endtime = datetime.datetime.now()
            self.status_bar.showMessage(
                str((endtime - starttime).total_seconds()) + " ms", 2500)
            print("3. SRC_Workbook loaded")
        else:
            QMessageBox.warning(
                self, "Hata", "Ürünü kullanmak için satın almanız gerekmektedir!")
            sys.exit()

    def closeEvent(self, event):
        settings = QSettings("berkaygediz", "SpanRC")

        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()

        if self.is_saved == False:
            reply = QMessageBox.question(self, 'SpanRC',
                                         translations[settings.value("current_language")]["exit_message"], QMessageBox.Yes |
                                         QMessageBox.No, QMessageBox.No)

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
        self.SRC_updateTitle()

    def SRC_updateTitle(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()
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
            f"<th>{translations[settings.value(
                'current_language')]['statistics_title']}</th>"
        )
        statistics += f"<td>{translations[settings.value('current_language')]['statistics_message1']}</td><td>{
            row}</td><td>Cols: </td><td>{column}</td>"
        statistics += f"<td>{translations[settings.value('current_language')]['statistics_message2']}</td><td>{
            selected_cell[0]}:{selected_cell[1]}</td>"
        if self.src_table.selectedRanges():
            statistics += f"<td>{translations[settings.value(
                'current_language')]['statistics_message3']}</td><td>"
            for selected_range in self.src_table.selectedRanges():
                statistics += (
                    f"{selected_range.topRow() + 1}:{selected_range.leftColumn() + 1} - "
                    f"{selected_range.bottomRow(
                    ) + 1}:{selected_range.rightColumn() + 1}</td>"
                )
        else:
            statistics += f"<td>{translations[settings.value('current_language')]['statistics_message4']}</td><td>{
                selected_cell[0]}:{selected_cell[1]}</td>"

        statistics += "</td><td id='sr-text'>SpanRC</td></tr></table></body></html>"
        self.statistics_label.setText(statistics)
        self.statusBar().addPermanentWidget(self.statistics_label)
        self.SRC_updateTitle()

    def SRC_saveState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        settings.setValue('window_geometry', self.saveGeometry())
        settings.setValue('default_directory', self.default_directory)
        self.file_name = settings.value('file_name', self.file_name)
        self.is_saved = settings.value('is_saved', self.is_saved)
        if self.selected_file:
            settings.setValue('last_opened_file', self.selected_file)
        settings.setValue(
            'current_theme', 'dark' if self.palette() == self.dark_theme else 'light')
        settings.setValue("current_language",
                          self.language_combobox.currentText())
        settings.sync()

    def SRC_restoreState(self):
        settings = QSettings("berkaygediz", "SpanRC")
        self.geometry = settings.value('window_geometry')
        self.directory = settings.value(
            'default_directory', self.default_directory)
        self.file_name = settings.value('file_name', self.file_name)
        self.is_saved = settings.value('is_saved', self.is_saved)
        self.language_combobox.setCurrentText(
            settings.value("current_language"))

        if self.geometry is not None:
            self.restoreGeometry(self.geometry)

        self.last_opened_file = settings.value('last_opened_file', '')
        if self.last_opened_file and os.path.exists(self.last_opened_file):
            self.loadFile(self.last_opened_file)

        self.src_table.setColumnCount(
            int(settings.value('columnCount', self.src_table.columnCount())))
        self.src_table.setRowCount(
            int(settings.value('rowCount', self.src_table.rowCount())))

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

        self.src_table.setCurrentCell(
            int(settings.value('currentRow', 0)), int(settings.value('currentColumn', 0)))
        self.src_table.scrollToItem(self.src_table.item(
            int(settings.value('currentRow', 0)), int(settings.value('currentColumn', 0))))

        if self.file_name:
            self.is_saved = True
        else:
            self.is_saved = False

        self.SRC_restoreTheme(),
        self.SRC_updateTitle()

    def SRC_restoreTheme(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if settings.value('current_theme') == 'dark':
            self.setPalette(self.dark_theme)
        else:
            self.setPalette(self.light_theme)
        self.SRC_toolbarTheme()

    def SRC_themePalette(self):
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

    def SRC_themeAction(self):
        if self.palette() == self.light_theme:
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

    def SRC_toolbarTranslate(self):
        settings = QSettings("berkaygediz", "SpanRC")
        if settings.value("current_language") == None:
            settings.setValue("current_language", "English")
            settings.sync()
        self.language = settings.value("current_language")
        self.language = settings.value("current_language")
        self.file_toolbar.setWindowTitle(
            translations[settings.value("current_language")]["file"])
        self.edit_toolbar.setWindowTitle(
            translations[settings.value("current_language")]["edit"])
        self.formula_toolbar.setWindowTitle(
            translations[settings.value("current_language")]["formula"])
        self.interface_toolbar.setWindowTitle(
            translations[settings.value("current_language")]["interface"])
        self.newAction.setText(
            translations[settings.value("current_language")]["new"])
        self.newAction.setStatusTip(
            translations[settings.value("current_language")]["new_title"])
        self.openAction.setText(
            translations[settings.value("current_language")]["open"])
        self.openAction.setStatusTip(
            translations[settings.value("current_language")]["open_title"])
        self.saveAction.setText(
            translations[settings.value("current_language")]["save"])
        self.saveAction.setStatusTip(
            translations[settings.value("current_language")]["save_title"])
        self.saveasAction.setText(
            translations[settings.value("current_language")]["save_as"])
        self.saveasAction.setStatusTip(
            translations[settings.value("current_language")]["save_as_title"])
        self.printAction.setText(
            translations[settings.value("current_language")]["print"])
        self.printAction.setStatusTip(
            translations[settings.value("current_language")]["print_title"])
        self.exitAction.setText(
            translations[settings.value("current_language")]["exit"])
        self.exitAction.setStatusTip(
            translations[settings.value("current_language")]["exit_message"])
        self.deleteAction.setText(
            translations[settings.value("current_language")]["delete"])
        self.deleteAction.setStatusTip(
            translations[settings.value("current_language")]["delete_title"])
        self.aboutAction.setText(
            translations[settings.value("current_language")]["about"])
        self.aboutAction.setStatusTip(
            translations[settings.value("current_language")]["about_title"])
        self.undoAction.setText(
            translations[settings.value("current_language")]["undo"])
        self.undoAction.setStatusTip(
            translations[settings.value("current_language")]["undo_title"])
        self.redoAction.setText(
            translations[settings.value("current_language")]["redo"])
        self.redoAction.setStatusTip(
            translations[settings.value("current_language")]["redo_title"])
        self.darklightAction.setText(
            translations[settings.value("current_language")]["darklight"])
        self.darklightAction.setStatusTip(
            translations[settings.value("current_language")]["darklight_message"])

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
                                f"<tr><td>Ctrl + N</td><td>{translations[settings.value(
                                    'current_language')]['new_title']}</td></tr>"
                                f"<tr><td>Ctrl + O</td><td>{translations[settings.value(
                                    'current_language')]['open_title']}</td></tr>"
                                f"<tr><td>Ctrl + S</td><td>{translations[settings.value(
                                    'current_language')]['save_title']}</td></tr>"
                                f"<tr><td>Ctrl + Shift + S</td><td>{
                                    translations[settings.value('current_language')]['save_as_title']}</td></tr>"
                                f"<tr><td>Ctrl + P</td><td>{translations[settings.value(
                                    'current_language')]['print_title']}</td></tr>"
                                f"<tr><td>Ctrl + Q</td><td>{translations[settings.value(
                                    'current_language')]['exit_title']}</td></tr>"
                                f"<tr><td>Ctrl + D</td><td>{translations[settings.value(
                                    'current_language')]['delete_title']}</td></tr>"
                                f"<tr><td>Ctrl + A</td><td>{translations[settings.value(
                                    'current_language')]['about_title']}</td></tr>"
                                "</table></body></html>")
        self.dock_widget.setWidget(self.help_label)
        self.dock_widget.setObjectName('Help')
        self.dock_widget.setAllowedAreas(
            Qt.LeftDockWidgetArea | Qt.RightDockWidgetArea)
        self.addDockWidget(Qt.RightDockWidgetArea, self.dock_widget)

    def SRC_toolbarLabel(self, toolbar, text):
        label = QLabel(f"<b>{text}</b>")
        toolbar.addWidget(label)

    def SRC_createAction(self, text, status_tip, function, shortcut=None):
        action = QAction(text, self)
        action.setStatusTip(status_tip)
        action.triggered.connect(function)
        if shortcut:
            action.setShortcut(shortcut)
        return action

    def SRC_setupActions(self):
        settings = QSettings("berkaygediz", "SpanRC")
        current_language = settings.value("current_language")

        self.newAction = self.SRC_createAction(
            translations[current_language]["new"], translations[current_language]["new_title"], self.new, QKeySequence.New)
        self.openAction = self.SRC_createAction(
            translations[current_language]["open"], translations[current_language]["open_title"], self.open, QKeySequence.Open)
        self.saveAction = self.SRC_createAction(
            translations[current_language]["save"], translations[current_language]["save_title"], self.save, QKeySequence.Save)
        self.saveasAction = self.SRC_createAction(
            translations[current_language]["save_as"], translations[current_language]["save_as_title"], self.saveAs, QKeySequence.SaveAs)
        self.printAction = self.SRC_createAction(
            translations[current_language]["print"], translations[current_language]["print_title"], self.print, QKeySequence.Print)
        self.exitAction = self.SRC_createAction(
            translations[current_language]["exit"], translations[current_language]["exit_message"], self.close, QKeySequence.Quit)
        self.deleteAction = self.SRC_createAction(
            translations[current_language]["delete"], translations[current_language]["delete_title"], self.deletecell, QKeySequence.Delete)
        self.aboutAction = self.SRC_createAction(
            translations[current_language]["about"], translations[current_language]["about_title"], self.showAbout, QKeySequence.HelpContents)
        self.undoAction = self.SRC_createAction(
            translations[current_language]["undo"], translations[current_language]["undo_title"], self.undo_stack.undo, QKeySequence.Undo)
        self.redoAction = self.SRC_createAction(
            translations[current_language]["redo"], translations[current_language]["redo_title"], self.undo_stack.redo, QKeySequence.Redo)
        self.darklightAction = self.SRC_createAction(
            translations[current_language]["darklight"], translations[current_language]["darklight_message"], self.SRC_themeAction, QKeySequence("Ctrl+D"))
        self.logoutaction = self.SRC_createAction(
            translations[settings.value("current_language")]["logout"], "", self.logout)
        self.syncsettingsaction = self.SRC_createAction(translations[settings.value(
            "current_language")]["syncsettings"], "", self.syncsettings)

    def syncsettings(self):
        settings = QSettings("berkaygediz", "SpanRC")
        sqlite_file = os.path.join(os.path.dirname(
            os.path.abspath(__file__)), 'src.db')
        sqliteConnection = sqlite3.connect(sqlite_file).cursor()
        user_email = sqliteConnection.execute(
            "SELECT email FROM profile").fetchone()[0]
        mysqlserver = serverconnect()
        mysqlcursor = mysqlserver.cursor()
        mysqlcursor.execute(
            "SELECT * FROM user_settings WHERE email = %s", (user_email,))
        user_settings = mysqlcursor.fetchone()
        if user_settings == None:
            sqliteConnection.execute(
                "SELECT * FROM user_settings WHERE email = ?", (user_email,))
            user_settings = sqliteConnection.fetchone()
            if user_settings == None:
                pass
            else:
                settings.setValue("current_theme", user_settings[1])
                settings.setValue("current_language", user_settings[2])
                settings.sync()
            mysqlcursor.execute(
                "INSERT INTO user_settings (email, theme, language) VALUES (%s, %s, %s)", (user_email, settings.value("current_theme"), settings.value("current_language")))
            mysqlserver.commit()
        else:
            mysqlcursor.execute(
                "UPDATE user_settings SET theme = %s, language = %s WHERE email = %s", (settings.value("current_theme"), settings.value("current_language"), user_email))
            mysqlserver.commit()

    def logout(self):
        sqlite_file = os.path.join(os.path.dirname(
            os.path.abspath(__file__)), 'src.db')
        sqliteConnection = sqlite3.connect(sqlite_file).cursor()
        sqliteConnection.execute("DROP TABLE profile")
        sqliteConnection.connection.commit()
        sqliteConnection.execute("DROP TABLE apps")
        sqliteConnection.connection.commit()
        sqliteConnection.execute("DROP TABLE log")
        sqliteConnection.connection.commit()
        try:
            settings = QSettings("berkaygediz", "SpanRC")
            settings.clear()
            settings.sync()
        except:
            pass
        QTimer.singleShot(0, self.hide)
        QTimer.singleShot(750, SRC_ControlInfo().show)

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

        self.toolbar = self.addToolBar("Account")
        self.SRC_toolbarLabel(self.toolbar, "Account: ")
        sqlite_file = os.path.join(os.path.dirname(
            os.path.abspath(__file__)), 'src.db')
        sqliteConnection = sqlite3.connect(sqlite_file).cursor()
        user_email = sqliteConnection.execute(
            "SELECT email FROM profile").fetchone()[0]
        self.user_name = QLabel(user_email)
        self.user_name.setStyleSheet("color: #000000;")
        self.toolbar.addWidget(self.user_name)
        self.toolbar.addAction(self.logoutaction)
        self.toolbar.addAction(self.syncsettingsaction)
        self.addToolBarBreak()

    def SRC_createDockwidget(self):
        widget = QWidget()
        layout = QVBoxLayout()
        widget.setLayout(layout)
        self.statistics = QLabel()
        self.statistics.setText('Statistics')
        layout.addWidget(self.statistics)
        return widget

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
        file_filter = f"{translations[settings.value(
            'current_language')]['xsrc']} (*.xsrc);;Comma Separated Values (*.csv)"
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

            QMessageBox.information(self, "Formula", f"{formula} : {result}")

        except ValueError as valueerror:
            QMessageBox.critical(self, "Formula", str(valueerror))
        except Exception as exception:
            QMessageBox.critical(self, "Formula", str(exception))

    def deletecell(self):
        for item in self.src_table.selectedItems():
            item.setText('')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setOrganizationName('berkaygediz')
    app.setApplicationName('SpanRC')
    app.setApplicationDisplayName('SpanRC')
    app.setApplicationVersion('1.3.0')
    ci = SRC_ControlInfo()
    ci.show()
    sys.exit(app.exec_())
