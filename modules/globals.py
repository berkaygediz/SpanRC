from PySide6.QtCore import *

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
    "appLanguage": "1252",
    "adaptiveResponse": 1,
    "readFilter": "General File (*.ssfs *.xsrc *.csv *.xlsx)",
    "writeFilter": "SolidSheets Workbook (*.ssfs);;Comma Separated Values (*.csv)",
}

# Locale ID (LCID)
languages = {
    "1252": "English",
    "1031": "Deutsch",
    "1034": "Español",
    "1055": "Türkçe",
    "1068": "Azərbaycanca",
    "1091": "Uzbek",
    "2052": "中文",  # Chinese
    "1042": "한국어",  # Korean
    "1041": "日本語",  # Japanese
    "1025": "العربية",  # Saudi Arabia
    "1049": "Русский",  # Russia
    "1036": "Français",
    "1032": "Ελληνικά",  # Greek
}

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

translations = {
    "1252": {
        "new": "New",
        "new_title": "New file",
        "open": "Open",
        "open_title": "Open file",
        "save": "Save",
        "save_title": "Save file",
        "save_as": "Save as",
        "save_as_title": "Save file as",
        "delete": "Delete",
        "delete_title": "Delete file",
        "print": "Print",
        "print_title": "Print file",
        "undo": "Undo",
        "undo_title": "Undo action",
        "redo": "Redo",
        "redo_title": "Redo action",
        "about": "About",
        "about_title": "About this program",
        "formula": "Formula",
        "file": "File",
        "edit": "Edit",
        "interface": "Interface",
        "darklight": "Dark/Light",
        "darklight_message": "Switch between dark and light mode",
        "help": "Help",
        "exit": "Exit",
        "exit_message": "Are you sure you want to exit?",
        "exit_title": "Exit confirmation",
        "statistics_title": "Statistics",
        "statistics_message1": "Row: ",
        "statistics_message2": "Cols: ",
        "statistics_message3": "Selected Cells:",
        "statistics_message4": "Selected Range:",
        "xsrc": "SolidSheets Workbook",
        "connection_verified": "Connection verified",
        "connection_denied": "Connection denied",
        "logout": "Logout",
        "syncsettings": "Sync Settings",
        "welcome-title": "Welcome — SolidSheets",
        "intro": "Login or register with <b>BG Ecosystem</b> account to continue.",
        "wrong_password": "Wrong password",
        "no_account": "You don't have an account.",
        "email": "Email",
        "password": "Password",
        "login": "Login",
        "register": "Register",
        "error": "Error",
        "fill_all": "Please fill all fields.",
        "login_success": "Login successful, redirecting...",
        "wrong_credentials": "Wrong credentials",
        "register_success": "Registration successful, please login.",
        "already_registered": "Already registered",
        "exit": "Exit",
        "compute": "Compute",
        "account": "Account",
        "add_row": "Add row",
        "add_row_title": "Add row",
        "add_column": "Add column",
        "add_column_title": "Add column",
        "add_row_above": "Add row above",
        "add_row_above_title": "Add row above",
        "add_column_left": "Add column left",
        "add_column_left_title": "Add column left",
        "powersaver": "Power Saver",
        "powersaver_message": "Hybrid (Ultra/Standard) power saver",
    },
    "1055": {
        "new": "Yeni",
        "new_title": "Yeni dosya",
        "open": "Aç",
        "open_title": "Dosya aç",
        "save": "Kaydet",
        "save_title": "Dosya kaydet",
        "save_as": "Farklı kaydet",
        "save_as_title": "Dosya farklı kaydet",
        "delete": "Sil",
        "delete_title": "Dosya sil",
        "print": "Yazdır",
        "print_title": "Dosya yazdır",
        "undo": "Geri al",
        "undo_title": "İşlemi geri al",
        "redo": "İleri al",
        "redo_title": "İşlemi ileri al",
        "about": "Hakkında",
        "about_title": "Bu program hakkında",
        "formula": "Formül",
        "file": "Dosya",
        "edit": "Düzenle",
        "interface": "Arayüz",
        "darklight": "Karanlık/Aydınlık",
        "darklight_message": "Karanlık ve aydınlık mod arasında geçiş yapın",
        "help": "Yardım",
        "exit": "Çıkış",
        "exit_message": "Çıkmak istediğinize emin misiniz?",
        "exit_title": "Çıkış onayı",
        "statistics_title": "İstatistik",
        "statistics_message1": "Satır: ",
        "statistics_message2": "Sütünlar: ",
        "statistics_message3": "Seçilmiş Hücreler:",
        "statistics_message4": "Seçilmiş Aralık:",
        "xsrc": "SolidSheets Çalışma Kitabı",
        "connection_verified": "Bağlantı doğrulandı",
        "connection_denied": "Bağlantı reddedildi",
        "logout": "Çıkış Yap",
        "syncsettings": "Ayarları Senkronize Et",
        "welcome-title": "Hoşgeldiniz — SolidSheets",
        "intro": "Devam etmek için <b>BG Ecosystem</b> hesabınızla giriş yapın veya kaydolun.",
        "wrong_password": "Yanlış şifre",
        "no_account": "Hesabınız bulunmamaktadır.",
        "email": "E-posta",
        "password": "Şifre",
        "login": "Giriş Yap",
        "register": "Kayıt Ol",
        "error": "Hata",
        "fill_all": "Lütfen tüm alanları doldurun.",
        "login_success": "Giriş başarılı, yönlendiriliyor...",
        "wrong_credentials": "Yanlış kimlik bilgileri",
        "register_success": "Kayıt başarılı, lütfen giriş yapın.",
        "already_registered": "Zaten kayıtlı",
        "exit": "Çıkış",
        "compute": "Hesapla",
        "account": "Hesap",
        "add_row": "Satır ekle",
        "add_row_title": "Satır ekle",
        "add_column": "Sütun ekle",
        "add_column_title": "Sütun ekle",
        "add_row_above": "Üst satır ekle",
        "add_row_above_title": "Üst satır ekle",
        "add_column_left": "Sol sütun ekle",
        "add_column_left_title": "Sol sütun ekle",
        "powersaver": "Güç Tasarrufu",
        "powersaver_message": "Hibrit (Ultra/Standart) güç tasarrufu.",
    },
    "1068": {
        "new": "Yeni",
        "new_title": "Yeni fayl",
        "open": "Aç",
        "open_title": "Fayl aç",
        "save": "Yadda saxla",
        "save_title": "Fayl yadda saxla",
        "save_as": "Fərqli yadda saxla",
        "save_as_title": "Fayl fərqli yadda saxla",
        "delete": "Sil",
        "delete_title": "Fayl sil",
        "print": "Çap et",
        "print_title": "Fayl çap et",
        "undo": "Geri al",
        "undo_title": "Əməliyyatı geri al",
        "redo": "İrəli al",
        "redo_title": "Əməliyyatı irəli al",
        "about": "Haqqında",
        "about_title": "Bu proqram haqqında",
        "formula": "Formula",
        "file": "Fayl",
        "edit": "Düzənlə",
        "interface": "İnterfeys",
        "darklight": "Tünd/Aydın",
        "darklight_message": "Tünd və aydın rejimləri arasında keçid edin",
        "help": "Kömək",
        "exit": "Çıxış",
        "exit_message": "Çıxmaq istədiyinizə əminsiniz?",
        "exit_title": "Çıxış təsdiqi",
        "statistics_title": "Statistika",
        "statistics_message1": "Sətir: ",
        "statistics_message2": "Sütunlar: ",
        "statistics_message3": "Seçilmiş Xanalar:",
        "statistics_message4": "Seçilmiş Aralıq:",
        "xsrc": "SolidSheets İş Kitabı",
        "connection_verified": "Bağlantı doğrulandı",
        "connection_denied": "Bağlantı reddedildi",
        "logout": "Çıxış Et",
        "syncsettings": "Ayarları Senkronize Et",
        "welcome-title": "Xoş Gəlmisiniz — SolidSheets",
        "intro": "Davam etmək üçün <b>BG Ecosystem</b> hesabınızla daxil olun və ya qeydiyyatdan keçin.",
        "wrong_password": "Yanlış şifrə",
        "no_account": "Hesabınız yoxdur.",
        "email": "E-poçt",
        "password": "Şifrə",
        "login": "Giriş Et",
        "register": "Qeydiyyatdan Keç",
        "error": "Xəta",
        "fill_all": "Xahiş olunur bütün xanaları doldurun.",
        "login_success": "Giriş uğurlu, yönləndirilir...",
        "wrong_credentials": "Yanlış kimlik məlumatları",
        "register_success": "Qeydiyyat uğurlu, xahiş olunur giriş edin.",
        "already_registered": "Artıq qeydiyyatdan keçilib",
        "exit": "Çıxış",
        "compute": "Hesabla",
        "account": "Hesab",
        "add_row": "Sətir əlavə et",
        "add_row_title": "Sətir əlavə et",
        "add_column": "Sütun əlavə et",
        "add_column_title": "Sütun əlavə et",
        "add_row_above": "Üst sətir əlavə et",
        "add_row_above_title": "Üst sətir əlavə et",
        "add_column_left": "Sol sütun əlavə et",
        "add_column_left_title": "Sol sütun əlavə et",
        "powersaver": "Enerji qənaəti",
        "powersaver_message": "Hibrid (Ultra/Standart) enerji qənaəti",
    },
    "1031": {
        "new": "Neu",
        "new_title": "Neue Datei",
        "open": "Öffnen",
        "open_title": "Datei öffnen",
        "save": "Speichern",
        "save_title": "Datei speichern",
        "save_as": "Speichern als",
        "save_as_title": "Datei speichern als",
        "delete": "Löschen",
        "delete_title": "Datei löschen",
        "print": "Drucken",
        "print_title": "Datei drucken",
        "about": "Über",
        "about_title": "Über dieses Programm",
        "formula": "Formel",
        "undo": "Rückgängig",
        "undo_title": "Aktion rückgängig machen",
        "redo": "Wiederholen",
        "redo_title": "Aktion wiederholen",
        "file": "Datei",
        "edit": "Bearbeiten",
        "interface": "Interface",
        "darklight": "Dunkel/Hell",
        "darklight_message": "Wechseln Sie zwischen Dunkel- und Hellmodus",
        "help": "Hilfe",
        "exit": "Beenden",
        "exit_message": "Wollen Sie wirklich beenden?",
        "exit_title": "Beenden Bestätigung",
        "statistics_title": "Statistik",
        "statistics_message1": "Zeile: ",
        "statistics_message2": "Spalten: ",
        "statistics_message3": "Ausgewählte Zellen:",
        "statistics_message4": "Ausgewählter Bereich:",
        "xsrc": "SolidSheets Arbeitsmappe",
        "connection_verified": "Verbindung verifiziert",
        "connection_denied": "Verbindung verweigert",
        "logout": "Ausloggen",
        "syncsettings": "Einstellungen synchronisieren",
        "welcome-title": "Willkommen — SolidSheets",
        "intro": "Melden Sie sich mit Ihrem <b>BG Ecosystem</b>-Konto an oder registrieren Sie sich, um fortzufahren.",
        "wrong_password": "Falsches Passwort",
        "no_account": "Sie haben kein Konto.",
        "email": "E-Mail",
        "password": "Passwort",
        "login": "Anmelden",
        "register": "Registrieren",
        "error": "Fehler",
        "fill_all": "Bitte füllen Sie alle Felder aus.",
        "login_success": "Anmeldung erfolgreich, Weiterleitung...",
        "wrong_credentials": "Falsche Anmeldedaten",
        "register_success": "Registrierung erfolgreich, bitte anmelden.",
        "already_registered": "Bereits registriert",
        "exit": "Beenden",
        "compute": "Berechnen",
        "account": "Konto",
        "add_row": "Zeile hinzufügen",
        "add_row_title": "Zeile hinzufügen",
        "add_column": "Spalte hinzufügen",
        "add_column_title": "Spalte hinzufügen",
        "add_row_above": "Zeile oben hinzufügen",
        "add_row_above_title": "Zeile oben hinzufügen",
        "add_column_left": "Spalte links hinzufügen",
        "add_column_left_title": "Spalte links hinzufügen",
        "powersaver": "Energiesparmodus",
        "powersaver_message": "Hybrid (Ultra/Standard) Energiesparmodus",
    },
    "1034": {
        "new": "Nuevo",
        "new_title": "Nuevo archivo",
        "open": "Abrir",
        "open_title": "Abrir archivo",
        "save": "Guardar",
        "save_title": "Guardar archivo",
        "save_as": "Guardar como",
        "save_as_title": "Guardar archivo como",
        "delete": "Eliminar",
        "delete_title": "Eliminar archivo",
        "print": "Imprimir",
        "print_title": "Imprimir archivo",
        "about": "Acerca de",
        "about_title": "Acerca de este programa",
        "formula": "Fórmula",
        "undo": "Deshacer",
        "undo_title": "Deshacer acción",
        "redo": "Rehacer",
        "redo_title": "Rehacer acción",
        "file": "Archivo",
        "edit": "Editar",
        "interface": "Interfaz",
        "darklight": "Oscuro/Claro",
        "darklight_message": "Cambiar entre modo oscuro y claro",
        "help": "Ayuda",
        "exit": "Salir",
        "exit_message": "¿Está seguro de que desea salir?",
        "exit_title": "Confirmación de salida",
        "statistics_title": "Estadísticas",
        "statistics_message1": "Fila: ",
        "statistics_message2": "Columnas: ",
        "statistics_message3": "Celdas seleccionadas:",
        "statistics_message4": "Rango seleccionado:",
        "xsrc": "SolidSheets Libro de trabajo",
        "connection_verified": "Conexión verificada",
        "connection_denied": "Conexión denegada",
        "welcome-title": "Bienvenido — SolidSheets",
        "intro": "Inicia sesión o regístrate con una cuenta de <b>BG Ecosystem</b> para continuar.",
        "logout": "Cerrar sesión",
        "syncsettings": "Sincronizar ajustes",
        "wrong_password": "Contraseña incorrecta",
        "no_account": "No tienes una cuenta.",
        "email": "Correo electrónico",
        "password": "Contraseña",
        "login": "Iniciar sesión",
        "register": "Registrarse",
        "error": "Error",
        "fill_all": "Por favor, rellene todos los campos.",
        "login_success": "Inicio de sesión correcto, redireccionando...",
        "wrong_credentials": "Credenciales incorrectas",
        "register_success": "Registro correcto, por favor inicie sesión.",
        "already_registered": "Ya registrado",
        "exit": "Salir",
        "compute": "Calcular",
        "account": "Cuenta",
        "add_row": "Añadir fila",
        "add_row_title": "Añadir fila",
        "add_column": "Añadir columna",
        "add_column_title": "Añadir columna",
        "add_row_above": "Añadir fila arriba",
        "add_row_above_title": "Añadir fila arriba",
        "add_column_left": "Añadir columna izquierda",
        "add_column_left_title": "Añadir columna izquierda",
        "powersaver": "Ahorro de energía",
        "powersaver_message": "Ahorro de energía híbrido (Ultra/Estándar).",
    },
    "1091": {
        "new": "Yangi",
        "new_title": "Yangi fayl",
        "open": "Ochish",
        "open_title": "Faylni ochish",
        "save": "Saqlash",
        "save_title": "Faylni saqlash",
        "save_as": "Boshqacha saqlash",
        "save_as_title": "Faylni boshqacha saqlash",
        "delete": "O'chirish",
        "delete_title": "Faylni o'chirish",
        "print": "Chop etish",
        "print_title": "Faylni chop etish",
        "undo": "Bekor qilish",
        "undo_title": "Amalni bekor qilish",
        "redo": "Qayta bajarish",
        "redo_title": "Amalni qayta bajarish",
        "about": "To'g'risida",
        "about_title": "Ushbu dastur to'g'risida",
        "formula": "Formulalar",
        "file": "Fayl",
        "edit": "Tahrirlash",
        "interface": "Interfeys",
        "darklight": "Qorong'i/Yorqin",
        "darklight_message": "Qorong'i va yorqin rejimlar o'rtasida o'ting",
        "help": "Yordam",
        "exit": "Chiqish",
        "exit_message": "Chiqmoqchi ekanligingizga ishonchingiz komilmi?",
        "exit_title": "Chiqish tasdiqlash",
        "statistics_title": "Statistika",
        "statistics_message1": "Qator: ",
        "statistics_message2": "Ustunlar: ",
        "statistics_message3": "Tanlangan Hujayralar:",
        "statistics_message4": "Tanlangan Oralq:",
        "xsrc": "SolidSheets Ishchi Kitobi",
        "connection_verified": "Aloqa tasdiqlandi",
        "connection_denied": "Aloqa rad etildi",
        "logout": "Chiqish",
        "syncsettings": "Sozlamalarni Sinxronlashtirish",
        "welcome-title": "Xush kelibsiz — SolidSheets",
        "intro": "Davom etish uchun <b>BG Ecosystem</b> hisobingiz bilan kiring yoki ro'yxatdan o'ting.",
        "wrong_password": "Noto'g'ri parol",
        "no_account": "Hisobingiz yo'q.",
        "email": "Elektron pochta",
        "password": "Parol",
        "login": "Kirish",
        "register": "Ro'yxatdan o'tish",
        "error": "Xato",
        "fill_all": "Iltimos, barcha maydonlarni to'ldiring.",
        "login_success": "Kirish muvaffaqiyatli, yo'naltirilmoqda...",
        "wrong_credentials": "Noto'g'ri ma'lumotlar",
        "register_success": "Ro'yxatdan o'tish muvaffaqiyatli, iltimos, kiring.",
        "already_registered": "Allaqachon ro'yxatdan o'tilgan",
        "exit": "Chiqish",
        "compute": "Hisoblash",
        "account": "Hisob",
        "add_row": "Qator qo'shish",
        "add_row_title": "Qator qo'shish",
        "add_column": "Ustun qo'shish",
        "add_column_title": "Ustun qo'shish",
        "add_row_above": "Yuqoriga qator qo'shish",
        "add_row_above_title": "Yuqoriga qator qo'shish",
        "add_column_left": "Chapga ustun qo'shish",
        "add_column_left_title": "Chapga ustun qo'shish",
        "powersaver": "Quvvatni tejash",
        "powersaver_message": "Gibrid (Ultra/Standart) quvvatni tejash.",
    },
    "2052": {
        "new": "新建",
        "new_title": "新文件",
        "open": "打开",
        "open_title": "打开文件",
        "save": "保存",
        "save_title": "保存文件",
        "save_as": "另存为",
        "save_as_title": "另存文件",
        "delete": "删除",
        "delete_title": "删除文件",
        "print": "打印",
        "print_title": "打印文件",
        "undo": "撤销",
        "undo_title": "撤销操作",
        "redo": "重做",
        "redo_title": "重做操作",
        "about": "关于",
        "about_title": "关于这个程序",
        "formula": "公式",
        "file": "文件",
        "edit": "编辑",
        "interface": "界面",
        "darklight": "黑暗/明亮",
        "darklight_message": "在黑暗模式和明亮模式之间切换",
        "help": "帮助",
        "exit": "退出",
        "exit_message": "您确定要退出吗？",
        "exit_title": "退出确认",
        "statistics_title": "统计信息",
        "statistics_message1": "行：",
        "statistics_message2": "列：",
        "statistics_message3": "选定单元格：",
        "statistics_message4": "选定范围：",
        "xsrc": "SolidSheets 工作簿",
        "connection_verified": "连接已验证",
        "connection_denied": "连接被拒绝",
        "logout": "登出",
        "syncsettings": "同步设置",
        "welcome-title": "欢迎 — SolidSheets",
        "intro": "请使用您的 <b>BG Ecosystem</b> 账户登录或注册以继续。",
        "wrong_password": "密码错误",
        "no_account": "没有账户。",
        "email": "电子邮件",
        "password": "密码",
        "login": "登录",
        "register": "注册",
        "error": "错误",
        "fill_all": "请填写所有字段。",
        "login_success": "登录成功，正在重定向...",
        "wrong_credentials": "凭据错误",
        "register_success": "注册成功，请登录。",
        "already_registered": "已注册",
        "exit": "退出",
        "compute": "计算",
        "account": "账户",
        "add_row": "添加行",
        "add_row_title": "添加行",
        "add_column": "添加列",
        "add_column_title": "添加列",
        "add_row_above": "添加上方行",
        "add_row_above_title": "添加上方行",
        "add_column_left": "添加左侧列",
        "add_column_left_title": "添加左侧列",
        "powersaver": "省电",
        "powersaver_message": "混合（超/标准）省电模式。",
    },
    "1042": {
        "new": "새로 만들기",
        "new_title": "새 파일",
        "open": "열기",
        "open_title": "파일 열기",
        "save": "저장",
        "save_title": "파일 저장",
        "save_as": "다른 이름으로 저장",
        "save_as_title": "파일을 다른 이름으로 저장",
        "delete": "삭제",
        "delete_title": "파일 삭제",
        "print": "인쇄",
        "print_title": "파일 인쇄",
        "undo": "실행 취소",
        "undo_title": "작업 취소",
        "redo": "다시 실행",
        "redo_title": "작업 다시 실행",
        "about": "정보",
        "about_title": "이 프로그램에 대하여",
        "formula": "수식",
        "file": "파일",
        "edit": "편집",
        "interface": "인터페이스",
        "darklight": "어두운/밝은 모드",
        "darklight_message": "어두운 모드와 밝은 모드 간 전환",
        "help": "도움말",
        "exit": "종료",
        "exit_message": "정말로 종료하시겠습니까?",
        "exit_title": "종료 확인",
        "statistics_title": "통계",
        "statistics_message1": "행: ",
        "statistics_message2": "열: ",
        "statistics_message3": "선택한 셀:",
        "statistics_message4": "선택한 범위:",
        "xsrc": "SolidSheets 작업북",
        "connection_verified": "연결 확인됨",
        "connection_denied": "연결 거부됨",
        "logout": "로그아웃",
        "syncsettings": "설정 동기화",
        "welcome-title": "환영합니다 — SolidSheets",
        "intro": "<b>BG Ecosystem</b> 계정으로 로그인하거나 등록하여 계속 진행하세요.",
        "wrong_password": "잘못된 비밀번호",
        "no_account": "계정이 없습니다.",
        "email": "이메일",
        "password": "비밀번호",
        "login": "로그인",
        "register": "가입하기",
        "error": "오류",
        "fill_all": "모든 필드를 입력해주세요.",
        "login_success": "로그인 성공, 리디렉션 중...",
        "wrong_credentials": "잘못된 인증 정보",
        "register_success": "가입 성공, 로그인하세요.",
        "already_registered": "이미 등록됨",
        "exit": "종료",
        "compute": "계산",
        "account": "계정",
        "add_row": "행 추가",
        "add_row_title": "행 추가",
        "add_column": "열 추가",
        "add_column_title": "열 추가",
        "add_row_above": "위에 행 추가",
        "add_row_above_title": "위에 행 추가",
        "add_column_left": "왼쪽에 열 추가",
        "add_column_left_title": "왼쪽에 열 추가",
        "powersaver": "전원 절약",
        "powersaver_message": "혼합(초기/표준) 전원 절약.",
    },
    "1041": {
        "new": "新規作成",
        "new_title": "新しいファイル",
        "open": "開く",
        "open_title": "ファイルを開く",
        "save": "保存",
        "save_title": "ファイルを保存",
        "save_as": "名前を付けて保存",
        "save_as_title": "ファイルを名前を付けて保存",
        "delete": "削除",
        "delete_title": "ファイルを削除",
        "print": "印刷",
        "print_title": "ファイルを印刷",
        "undo": "元に戻す",
        "undo_title": "操作を元に戻す",
        "redo": "やり直す",
        "redo_title": "操作をやり直す",
        "about": "情報",
        "about_title": "このプログラムについて",
        "formula": "数式",
        "file": "ファイル",
        "edit": "編集",
        "interface": "インターフェース",
        "darklight": "ダーク/ライト",
        "darklight_message": "ダークモードとライトモードの切り替え",
        "help": "ヘルプ",
        "exit": "終了",
        "exit_message": "本当に終了しますか？",
        "exit_title": "終了確認",
        "statistics_title": "統計",
        "statistics_message1": "行: ",
        "statistics_message2": "列: ",
        "statistics_message3": "選択したセル:",
        "statistics_message4": "選択した範囲:",
        "xsrc": "SolidSheets ワークブック",
        "connection_verified": "接続が確認されました",
        "connection_denied": "接続が拒否されました",
        "logout": "ログアウト",
        "syncsettings": "設定を同期",
        "welcome-title": "ようこそ — SolidSheets",
        "intro": "<b>BG Ecosystem</b> アカウントでログインするか、登録して続行してください。",
        "wrong_password": "パスワードが間違っています",
        "no_account": "アカウントがありません。",
        "email": "メールアドレス",
        "password": "パスワード",
        "login": "ログイン",
        "register": "登録",
        "error": "エラー",
        "fill_all": "すべてのフィールドを入力してください。",
        "login_success": "ログイン成功、リダイレクト中...",
        "wrong_credentials": "認証情報が間違っています",
        "register_success": "登録成功、ログインしてください。",
        "already_registered": "すでに登録済み",
        "exit": "終了",
        "compute": "計算",
        "account": "アカウント",
        "add_row": "行を追加",
        "add_row_title": "行を追加",
        "add_column": "列を追加",
        "add_column_title": "列を追加",
        "add_row_above": "上に行を追加",
        "add_row_above_title": "上に行を追加",
        "add_column_left": "左に列を追加",
        "add_column_left_title": "左に列を追加",
        "powersaver": "省電力",
        "powersaver_message": "ハイブリッド（ウルトラ/スタンダード）省電力。",
    },
    "1025": {
        "new": "جديد",
        "new_title": "ملف جديد",
        "open": "فتح",
        "open_title": "فتح ملف",
        "save": "حفظ",
        "save_title": "حفظ الملف",
        "save_as": "حفظ باسم",
        "save_as_title": "حفظ الملف باسم",
        "delete": "حذف",
        "delete_title": "حذف الملف",
        "print": "طباعة",
        "print_title": "طباعة الملف",
        "undo": "تراجع",
        "undo_title": "تراجع عن العملية",
        "redo": "إعادة",
        "redo_title": "إعادة العملية",
        "about": "حول",
        "about_title": "حول هذا البرنامج",
        "formula": "صيغة",
        "file": "ملف",
        "edit": "تحرير",
        "interface": "واجهة",
        "darklight": "داكن/فاتح",
        "darklight_message": "التبديل بين الوضع الداكن والوضع الفاتح",
        "help": "مساعدة",
        "exit": "خروج",
        "exit_message": "هل أنت متأكد أنك تريد الخروج؟",
        "exit_title": "تأكيد الخروج",
        "statistics_title": "إحصائيات",
        "statistics_message1": "صف: ",
        "statistics_message2": "أعمدة: ",
        "statistics_message3": "الخلايا المختارة:",
        "statistics_message4": "النطاق المختار:",
        "xsrc": "دفتر عمل SolidSheets",
        "connection_verified": "تم التحقق من الاتصال",
        "connection_denied": "تم رفض الاتصال",
        "logout": "تسجيل الخروج",
        "syncsettings": "مزامنة الإعدادات",
        "welcome-title": "مرحبًا — SolidSheets",
        "intro": "يرجى تسجيل الدخول أو التسجيل باستخدام حساب <b>BG Ecosystem</b> للمتابعة.",
        "wrong_password": "كلمة مرور خاطئة",
        "no_account": "لا يوجد لديك حساب.",
        "email": "البريد الإلكتروني",
        "password": "كلمة المرور",
        "login": "تسجيل الدخول",
        "register": "تسجيل",
        "error": "خطأ",
        "fill_all": "يرجى ملء جميع الحقول.",
        "login_success": "تم تسجيل الدخول بنجاح، جاري التوجيه...",
        "wrong_credentials": "معلومات الاعتماد خاطئة",
        "register_success": "تم التسجيل بنجاح، يرجى تسجيل الدخول.",
        "already_registered": "مُسجل بالفعل",
        "exit": "خروج",
        "compute": "حساب",
        "account": "حساب",
        "add_row": "إضافة صف",
        "add_row_title": "إضافة صف",
        "add_column": "إضافة عمود",
        "add_column_title": "إضافة عمود",
        "add_row_above": "إضافة صف أعلى",
        "add_row_above_title": "إضافة صف أعلى",
        "add_column_left": "إضافة عمود يسار",
        "add_column_left_title": "إضافة عمود يسار",
        "powersaver": "توفير الطاقة",
        "powersaver_message": "توفير الطاقة الهجين (فائق/قياسي).",
    },
    "1049": {
        "new": "Новый",
        "new_title": "Новый файл",
        "open": "Открыть",
        "open_title": "Открыть файл",
        "save": "Сохранить",
        "save_title": "Сохранить файл",
        "save_as": "Сохранить как",
        "save_as_title": "Сохранить файл как",
        "delete": "Удалить",
        "delete_title": "Удалить файл",
        "print": "Печать",
        "print_title": "Напечатать файл",
        "undo": "Отменить",
        "undo_title": "Отменить действие",
        "redo": "Повторить",
        "redo_title": "Повторить действие",
        "about": "О программе",
        "about_title": "Информация о программе",
        "formula": "Формула",
        "file": "Файл",
        "edit": "Редактировать",
        "interface": "Интерфейс",
        "darklight": "Темный/Светлый",
        "darklight_message": "Переключение между темным и светлым режимами",
        "help": "Помощь",
        "exit": "Выход",
        "exit_message": "Вы уверены, что хотите выйти?",
        "exit_title": "Подтверждение выхода",
        "statistics_title": "Статистика",
        "statistics_message1": "Строка: ",
        "statistics_message2": "Столбцы: ",
        "statistics_message3": "Выбранные ячейки:",
        "statistics_message4": "Выбранный диапазон:",
        "xsrc": "Рабочая книга SolidSheets",
        "connection_verified": "Подключение подтверждено",
        "connection_denied": "Подключение отклонено",
        "logout": "Выйти",
        "syncsettings": "Синхронизировать настройки",
        "welcome-title": "Добро пожаловать — SolidSheets",
        "intro": "Чтобы продолжить, войдите в свою учетную запись <b>BG Ecosystem</b> или зарегистрируйтесь.",
        "wrong_password": "Неверный пароль",
        "no_account": "У вас нет аккаунта.",
        "email": "Электронная почта",
        "password": "Пароль",
        "login": "Войти",
        "register": "Зарегистрироваться",
        "error": "Ошибка",
        "fill_all": "Пожалуйста, заполните все поля.",
        "login_success": "Успешный вход, перенаправление...",
        "wrong_credentials": "Неверные учетные данные",
        "register_success": "Регистрация прошла успешно, пожалуйста, войдите.",
        "already_registered": "Уже зарегистрирован",
        "exit": "Выход",
        "compute": "Вычислить",
        "account": "Аккаунт",
        "add_row": "Добавить строку",
        "add_row_title": "Добавить строку",
        "add_column": "Добавить столбец",
        "add_column_title": "Добавить столбец",
        "add_row_above": "Добавить строку выше",
        "add_row_above_title": "Добавить строку выше",
        "add_column_left": "Добавить столбец слева",
        "add_column_left_title": "Добавить столбец слева",
        "powersaver": "Энергосбережение",
        "powersaver_message": "Гибридный (Ультра/Стандартный) режим энергосбережения.",
    },
    "1036": {
        "new": "Nouveau",
        "new_title": "Nouveau fichier",
        "open": "Ouvrir",
        "open_title": "Ouvrir un fichier",
        "save": "Enregistrer",
        "save_title": "Enregistrer le fichier",
        "save_as": "Enregistrer sous",
        "save_as_title": "Enregistrer le fichier sous",
        "delete": "Supprimer",
        "delete_title": "Supprimer le fichier",
        "print": "Imprimer",
        "print_title": "Imprimer le fichier",
        "undo": "Annuler",
        "undo_title": "Annuler l'opération",
        "redo": "Rétablir",
        "redo_title": "Rétablir l'opération",
        "about": "À propos",
        "about_title": "À propos de ce programme",
        "formula": "Formule",
        "file": "Fichier",
        "edit": "Modifier",
        "interface": "Interface",
        "darklight": "Sombre/Clair",
        "darklight_message": "Basculer entre le mode sombre et le mode clair",
        "help": "Aide",
        "exit": "Quitter",
        "exit_message": "Êtes-vous sûr de vouloir quitter ?",
        "exit_title": "Confirmation de sortie",
        "statistics_title": "Statistiques",
        "statistics_message1": "Ligne : ",
        "statistics_message2": "Colonnes : ",
        "statistics_message3": "Cellules sélectionnées :",
        "statistics_message4": "Plage sélectionnée :",
        "xsrc": "Classeur SolidSheets",
        "connection_verified": "Connexion vérifiée",
        "connection_denied": "Connexion refusée",
        "logout": "Déconnexion",
        "syncsettings": "Synchroniser les paramètres",
        "welcome-title": "Bienvenue — SolidSheets",
        "intro": "Pour continuer, connectez-vous avec votre compte <b>BG Ecosystem</b> ou inscrivez-vous.",
        "wrong_password": "Mot de passe incorrect",
        "no_account": "Vous n'avez pas de compte.",
        "email": "E-mail",
        "password": "Mot de passe",
        "login": "Se connecter",
        "register": "S'inscrire",
        "error": "Erreur",
        "fill_all": "Veuillez remplir tous les champs.",
        "login_success": "Connexion réussie, redirection...",
        "wrong_credentials": "Informations d'identification incorrectes",
        "register_success": "Inscription réussie, veuillez vous connecter.",
        "already_registered": "Déjà inscrit",
        "exit": "Quitter",
        "compute": "Calculer",
        "account": "Compte",
        "add_row": "Ajouter une ligne",
        "add_row_title": "Ajouter une ligne",
        "add_column": "Ajouter une colonne",
        "add_column_title": "Ajouter une colonne",
        "add_row_above": "Ajouter une ligne au-dessus",
        "add_row_above_title": "Ajouter une ligne au-dessus",
        "add_column_left": "Ajouter une colonne à gauche",
        "add_column_left_title": "Ajouter une colonne à gauche",
        "powersaver": "Économie d'énergie",
        "powersaver_message": "Mode d'économie d'énergie hybride (Ultra/Standard).",
    },
    "1032": {
        "new": "Νέο",
        "new_title": "Νέο αρχείο",
        "open": "Άνοιγμα",
        "open_title": "Άνοιγμα αρχείου",
        "save": "Αποθήκευση",
        "save_title": "Αποθήκευση αρχείου",
        "save_as": "Αποθήκευση ως",
        "save_as_title": "Αποθήκευση αρχείου ως",
        "delete": "Διαγραφή",
        "delete_title": "Διαγραφή αρχείου",
        "print": "Εκτύπωση",
        "print_title": "Εκτύπωση αρχείου",
        "undo": "Αναίρεση",
        "undo_title": "Αναίρεση ενέργειας",
        "redo": "Επαναφορά",
        "redo_title": "Επαναφορά ενέργειας",
        "about": "Σχετικά",
        "about_title": "Σχετικά με αυτό το πρόγραμμα",
        "formula": "Τύπος",
        "file": "Αρχείο",
        "edit": "Επεξεργασία",
        "interface": "Διεπαφή",
        "darklight": "Σκοτεινός/Φωτεινός",
        "darklight_message": "Εναλλαγή μεταξύ σκοτεινής και φωτεινής λειτουργίας",
        "help": "Βοήθεια",
        "exit": "Έξοδος",
        "exit_message": "Είστε σίγουροι ότι θέλετε να αποχωρήσετε;",
        "exit_title": "Επιβεβαίωση εξόδου",
        "statistics_title": "Στατιστικά",
        "statistics_message1": "Γραμμή: ",
        "statistics_message2": "Στήλες: ",
        "statistics_message3": "Επιλεγμένα κελιά:",
        "statistics_message4": "Επιλεγμένη περιοχή:",
        "xsrc": "Βιβλίο εργασίας SolidSheets",
        "connection_verified": "Η σύνδεση επιβεβαιώθηκε",
        "connection_denied": "Η σύνδεση απορρίφθηκε",
        "logout": "Αποσύνδεση",
        "syncsettings": "Συγχρονισμός ρυθμίσεων",
        "welcome-title": "Καλώς ήρθατε — SolidSheets",
        "intro": "Για να συνεχίσετε, συνδεθείτε με τον λογαριασμό σας <b>BG Ecosystem</b> ή εγγραφείτε.",
        "wrong_password": "Λάθος κωδικός πρόσβασης",
        "no_account": "Δεν έχετε λογαριασμό.",
        "email": "Ηλεκτρονικό ταχυδρομείο",
        "password": "Κωδικός πρόσβασης",
        "login": "Σύνδεση",
        "register": "Εγγραφή",
        "error": "Σφάλμα",
        "fill_all": "Παρακαλώ συμπληρώστε όλα τα πεδία.",
        "login_success": "Η σύνδεση ήταν επιτυχής, ανακατεύθυνση...",
        "wrong_credentials": "Λάθος στοιχεία σύνδεσης",
        "register_success": "Η εγγραφή ήταν επιτυχής, παρακαλώ συνδεθείτε.",
        "already_registered": "Ήδη εγγεγραμμένος",
        "exit": "Έξοδος",
        "compute": "Υπολογισμός",
        "account": "Λογαριασμός",
        "add_row": "Προσθήκη γραμμής",
        "add_row_title": "Προσθήκη γραμμής",
        "add_column": "Προσθήκη στήλης",
        "add_column_title": "Προσθήκη στήλης",
        "add_row_above": "Προσθήκη γραμμής από πάνω",
        "add_row_above_title": "Προσθήκη γραμμής από πάνω",
        "add_column_left": "Προσθήκη στήλης αριστερά",
        "add_column_left_title": "Προσθήκη στήλης αριστερά",
        "powersaver": "Εξοικονόμηση ενέργειας",
        "powersaver_message": "Υβριδική (Ultra/Standard) λειτουργία εξοικονόμησης ενέργειας.",
    },
}
