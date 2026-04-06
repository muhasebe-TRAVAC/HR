#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# main.py

import sys
import atexit
import logging
import bcrypt
from datetime import datetime, timedelta

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QDialog, QVBoxLayout,
    QLabel, QLineEdit, QPushButton, QMessageBox, QFormLayout, QFrame
)
from PyQt5.QtCore import Qt, QTimer, QEvent, pyqtSignal, QObject
from PyQt5.QtGui import QFont
from PyQt5.QtNetwork import QLocalServer, QLocalSocket

import constants
from constants import STYLE, APP_VERSION, setup_directories
from database import DatabaseManager
from utils import btn

from tabs.dashboard_tab  import DashboardTab
from tabs.employees_tab  import EmployeesTab
from tabs.attendance_tab import AttendanceTab
from tabs.leaves_tab     import LeavesTab
from tabs.loans_tab      import LoansTab
from tabs.payroll_tab    import PayrollTab
from tabs.reports_tab    import ReportsTab
from tabs.audit_tab      import AuditTab
from tabs.settings_tab   import SettingsTab
from tabs.receipts_tab   import ReceiptsTab


# ==================== إعداد نظام السجلات ====================
def _setup_logging() -> None:
    """
    تهيئة نظام السجلات مرة واحدة عند بدء التشغيل.
    يكتب إلى hr_system.log وإلى الـ console معاً.
    """
    log_format = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"
    logging.basicConfig(
        level=logging.INFO,
        format=log_format,
        handlers=[
            logging.FileHandler("hr_system.log", encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
    )
    logging.getLogger("PyQt5").setLevel(logging.WARNING)


logger = logging.getLogger(__name__)


# ==================== منع التشغيل المزدوج ====================
SERVER_NAME = "hr_system_server"


def is_already_running() -> bool:
    """التحقق من أن نسخة أخرى من البرنامج تعمل بالفعل."""
    socket = QLocalSocket()
    socket.connectToServer(SERVER_NAME)
    if socket.waitForConnected(1000):
        socket.disconnectFromServer()
        return True
    return False


def create_server() -> QLocalServer:
    """إنشاء خادم محلي لمنع التشغيل المزدوج."""
    server = QLocalServer()
    QLocalServer.removeServer(SERVER_NAME)
    if server.listen(SERVER_NAME):
        return server
    return None


# ==================== كلاس التواصل بين التبويبات ====================
class Communicator(QObject):
    """إشارة موحدة لإخطار التبويبات بأي تغيير في البيانات."""
    dataChanged = pyqtSignal(str, object)


# ==================== نافذة تسجيل الدخول ====================
class LoginDialog(QDialog):
    """
    نافذة تسجيل الدخول مع:
    - حماية brute force: قفل مؤقت بعد MAX_ATTEMPTS محاولة فاشلة.
    - تحذير كلمة المرور الافتراضية.
    """
    MAX_ATTEMPTS    = 5
    LOCKOUT_SECONDS = 30

    def __init__(self, db: DatabaseManager):
        super().__init__()
        self.db            = db
        self.user          = None
        self._failed       = 0
        self._locked_until = None
        self.setWindowTitle("تسجيل الدخول - نظام الموارد البشرية")
        self.setFixedSize(400, 320)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(15)
        layout.setContentsMargins(40, 30, 40, 30)

        title = QLabel("🏢 نظام الموارد البشرية")
        title.setAlignment(Qt.AlignCenter)
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setStyleSheet("color:#1976D2; margin-bottom:10px;")
        layout.addWidget(title)

        company = self.db.get_setting('company_name', 'الشركة')
        co_lbl  = QLabel(company)
        co_lbl.setAlignment(Qt.AlignCenter)
        co_lbl.setStyleSheet("color:#555; font-size:13px;")
        layout.addWidget(co_lbl)

        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setStyleSheet("color:#ddd;")
        layout.addWidget(line)

        form = QFormLayout()
        form.setSpacing(10)

        self.username = QLineEdit()
        self.username.setPlaceholderText("اسم المستخدم")
        self.username.setText("admin")
        form.addRow("المستخدم:", self.username)

        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.Password)
        self.password.setPlaceholderText("كلمة المرور")
        self.password.returnPressed.connect(self._login)
        form.addRow("كلمة المرور:", self.password)

        layout.addLayout(form)

        self._lockout_lbl = QLabel("")
        self._lockout_lbl.setAlignment(Qt.AlignCenter)
        self._lockout_lbl.setStyleSheet("color:#D32F2F; font-size:11px;")
        self._lockout_lbl.hide()
        layout.addWidget(self._lockout_lbl)

        btn_login = QPushButton("دخول")
        btn_login.setStyleSheet(
            "background:#1976D2;color:white;padding:10px;font-size:14px;")
        btn_login.clicked.connect(self._login)
        layout.addWidget(btn_login)

    def _login(self):
        # التحقق من حالة القفل
        if self._locked_until and datetime.now() < self._locked_until:
            remaining = int((self._locked_until - datetime.now()).total_seconds())
            self._lockout_lbl.setText(f"⛔ تم قفل الدخول. حاول بعد {remaining} ثانية")
            self._lockout_lbl.show()
            logger.warning("محاولة دخول خلال فترة القفل: %s",
                           self.username.text().strip())
            return

        username = self.username.text().strip()
        password = self.password.text()

        if not username or not password:
            QMessageBox.warning(self, "خطأ", "أدخل اسم المستخدم وكلمة المرور")
            return

        user_row = self.db.fetch_one(
            "SELECT id, username, full_name, role, password_hash "
            "FROM users WHERE username=? AND is_active=1",
            (username,)
        )

        if not user_row:
            self._handle_failed_attempt(username)
            return

        user_id, db_username, full_name, role, stored_hash = user_row

        try:
            if bcrypt.checkpw(password.encode('utf-8'), stored_hash.encode('utf-8')):
                self._failed       = 0
                self._locked_until = None
                self._lockout_lbl.hide()

                self.user = {
                    'id':       user_id,
                    'username': db_username,
                    'name':     full_name,
                    'role':     role,
                }
                self.db.set_current_user(user_id)
                logger.info("تسجيل دخول ناجح: %s (%s)", db_username, role)

                # تحذير كلمة المرور الافتراضية
                if username == "admin" and bcrypt.checkpw(
                        b"admin123", stored_hash.encode('utf-8')):
                    QMessageBox.warning(
                        self, "⚠️ تنبيه أمني",
                        "أنت تستخدم كلمة المرور الافتراضية (admin123).\n"
                        "يُرجى تغييرها فوراً من الإعدادات ← إدارة المستخدمين."
                    )

                self.accept()
            else:
                self._handle_failed_attempt(username)

        except Exception as e:
            logger.error("خطأ في التحقق من كلمة المرور: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ", f"حدث خطأ في التحقق: {str(e)}")
            self.password.clear()

    def _handle_failed_attempt(self, username: str):
        self._failed += 1
        self.password.clear()
        logger.warning("محاولة دخول فاشلة #%d للمستخدم: %s",
                       self._failed, username)

        if self._failed >= self.MAX_ATTEMPTS:
            self._locked_until = datetime.now() + timedelta(
                seconds=self.LOCKOUT_SECONDS)
            self._failed = 0
            self._lockout_lbl.setText(
                f"⛔ تم قفل الدخول لمدة {self.LOCKOUT_SECONDS} ثانية "
                f"بسبب كثرة المحاولات")
            self._lockout_lbl.show()
            logger.warning("تم قفل تسجيل الدخول لمدة %d ثانية (مستخدم: %s)",
                           self.LOCKOUT_SECONDS, username)
        else:
            remaining_attempts = self.MAX_ATTEMPTS - self._failed
            self._lockout_lbl.setText(
                f"⚠️ كلمة مرور خاطئة. محاولات متبقية: {remaining_attempts}")
            self._lockout_lbl.show()
            QMessageBox.critical(
                self, "خطأ", "اسم المستخدم أو كلمة المرور غير صحيحة")


# ==================== النافذة الرئيسية ====================
class MainWindow(QMainWindow):
    """
    النافذة الرئيسية مع:
    - Session Timeout صحيح: يعتمد على eventFilter على مستوى QApplication
      لاصطياد كل حركة مؤشر أو ضغطة مفتاح في أي مكان من الواجهة.
      (الإصلاح: mousePressEvent/keyPressEvent على MainWindow كانا يفوِّتان
       الأحداث داخل QTableWidget وغيرها من الـ child widgets)
    - إغلاق نظيف: closeEvent يوقف المؤقتات ويسجِّل الخروج.
    """
    IDLE_TIMEOUT_MS = 15 * 60 * 1000   # 15 دقيقة

    def __init__(self, db: DatabaseManager, user: dict):
        super().__init__()
        self.db   = db
        self.user = user
        self.comm = Communicator()
        self._build()
        self._start_session_timer()

    # ---------- بناء الواجهة ----------
    def _build(self):
        company = self.db.get_setting('company_name', 'الشركة')
        self.setWindowTitle(
            f"نظام الموارد البشرية v{APP_VERSION} | {company} | {self.user['name']}")
        self.setGeometry(50, 50, 1400, 800)
        self.setLayoutDirection(Qt.RightToLeft)
        self.statusBar().showMessage(self._status_text())

        self.tabs = QTabWidget()
        self.tabs.addTab(DashboardTab(self.db,  self.user, self.comm), "🏠 لوحة التحكم")
        self.tabs.addTab(EmployeesTab(self.db,  self.user, self.comm), "👥 الموظفون")
        self.tabs.addTab(AttendanceTab(self.db, self.user, self.comm), "📅 الحضور")
        self.tabs.addTab(LeavesTab(self.db,     self.user, self.comm), "🏖️ الإجازات")
        self.tabs.addTab(PayrollTab(self.db,    self.user, self.comm), "💰 الرواتب")
        self.tabs.addTab(LoansTab(self.db,      self.user, self.comm), "💳 السلف")
        self.tabs.addTab(ReportsTab(self.db,    self.user, self.comm), "📊 التقارير")
        self.tabs.addTab(AuditTab(self.db,      self.user, self.comm), "📋 سجل الأحداث")
        self.tabs.addTab(ReceiptsTab(self.db,   self.user, self.comm), "🧾 إيصالات الدفع")
        self.tabs.addTab(SettingsTab(self.db,   self.user, self.comm), "⚙️ الإعدادات")
        self.setCentralWidget(self.tabs)

        # ساعة شريط الحالة — مرجع صريح لمنع garbage collection
        self._clock_timer = QTimer(self)
        self._clock_timer.timeout.connect(self._update_time)
        self._clock_timer.start(60_000)

    # ---------- Session Timeout (مُصلَح) ----------
    def _start_session_timer(self):
        """
        بدء مؤقت الخمول وتثبيت eventFilter على مستوى التطبيق.

        الإصلاح الجوهري:
        - النسخة القديمة كانت تعتمد على mousePressEvent/keyPressEvent
          في MainWindow فقط، وهذا يفوِّت كل الأحداث داخل الـ child widgets
          (QTableWidget، QComboBox، QLineEdit، ...) لأنها تستهلك الحدث
          قبل وصوله للـ MainWindow.
        - الحل الصحيح: تثبيت eventFilter على QApplication نفسه،
          فيصطاد كل حدث في أي widget بالتطبيق بأكمله.
        """
        self._idle_timer = QTimer(self)
        self._idle_timer.setSingleShot(True)
        self._idle_timer.setInterval(self.IDLE_TIMEOUT_MS)
        self._idle_timer.timeout.connect(self._lock_session)
        self._idle_timer.start()

        # تثبيت الفلتر على مستوى التطبيق
        QApplication.instance().installEventFilter(self)

    def eventFilter(self, obj, event) -> bool:
        """
        اصطياد أحداث النشاط من أي مكان في التطبيق لإعادة ضبط مؤقت الخمول.
        يعيد False دائماً لعدم التدخل في سلسلة معالجة الأحداث.
        """
        if event.type() in (
            QEvent.MouseButtonPress,
            QEvent.MouseButtonRelease,
            QEvent.MouseMove,
            QEvent.KeyPress,
            QEvent.Wheel,
        ):
            self._reset_idle_timer()
        return False   # لا تُوقف انتشار الحدث

    def _reset_idle_timer(self):
        """إعادة تشغيل مؤقت الخمول."""
        if hasattr(self, '_idle_timer'):
            self._idle_timer.start()

    def _lock_session(self):
        """قفل الجلسة وإظهار نافذة إعادة تسجيل الدخول."""
        logger.info("قفل الجلسة تلقائياً بسبب الخمول — المستخدم: %s",
                    self.user['name'])
        self.hide()
        dlg = LoginDialog(self.db)
        dlg.username.setText(self.user['username'])
        dlg.username.setEnabled(False)   # المستخدم ثابت — فقط كلمة المرور

        if dlg.exec_() == QDialog.Accepted:
            self.user = dlg.user
            self.show()
            self._reset_idle_timer()
        else:
            QApplication.quit()

    # ---------- الإغلاق النظيف ----------
    def closeEvent(self, event):
        """
        إيقاف المؤقتات وإزالة eventFilter عند إغلاق النافذة.
        يضمن عدم بقاء أي مرجع لـ self بعد التدمير.
        """
        try:
            # إيقاف المؤقتات أولاً
            if hasattr(self, '_idle_timer'):
                self._idle_timer.stop()
            if hasattr(self, '_clock_timer'):
                self._clock_timer.stop()
            # إزالة فلتر الأحداث
            QApplication.instance().removeEventFilter(self)
            logger.info("تم إغلاق النافذة الرئيسية بشكل نظيف")
        except Exception as e:
            logger.error("خطأ أثناء إغلاق النافذة: %s", e)
        finally:
            event.accept()

    # ---------- مساعدات ----------
    def _status_text(self) -> str:
        return (f"مرحباً {self.user['name']} ({self.user['role']})  |  "
                f"{datetime.now().strftime('%A %d/%m/%Y %H:%M')}")

    def _update_time(self):
        self.statusBar().showMessage(self._status_text())


# ==================== نقطة البداية ====================
if __name__ == "__main__":
    # 1. تهيئة نظام السجلات أولاً — قبل أي شيء آخر
    _setup_logging()
    logger.info("بدء تشغيل نظام الموارد البشرية v%s", APP_VERSION)

    app = QApplication(sys.argv)
    app.setLayoutDirection(Qt.RightToLeft)
    app.setFont(QFont("Arial", 10))
    app.setStyleSheet(STYLE)

    # 2. إنشاء المجلدات المطلوبة
    setup_directories()

    # 3. منع التشغيل المزدوج
    if is_already_running():
        QMessageBox.critical(
            None, "خطأ",
            "البرنامج يعمل بالفعل.\nيرجى إغلاق النسخة الأخرى أولاً.")
        sys.exit(0)

    # الاحتفاظ بمرجع للخادم داخل app لحمايته من garbage collection
    app._lock_server = create_server()
    if app._lock_server is None:
        QMessageBox.critical(
            None, "خطأ",
            "فشل في إنشاء قفل التطبيق.")
        sys.exit(0)

    # 4. تهيئة قاعدة البيانات
    db = DatabaseManager()

    # تسجيل إغلاق نظيف لقاعدة البيانات عند أي خروج
    def _cleanup():
        try:
            db.close()
            QLocalServer.removeServer(SERVER_NAME)
            logger.info("تم تنظيف الموارد بنجاح")
        except Exception as e:
            logger.error("خطأ في التنظيف: %s", e)

    atexit.register(_cleanup)

    # 5. إعداد اللغة — استخدام constants المودول مباشرة (لا القيمة)
    lang = db.get_setting('language', 'ar')
    if lang in ('ar', 'tr'):
        constants.CURRENT_LANG = lang
        logger.info("اللغة المُحمَّلة: %s", lang)

    # 6. تسجيل الدخول
    login = LoginDialog(db)
    if login.exec_() != QDialog.Accepted:
        sys.exit(0)

    user   = login.user
    window = MainWindow(db, user)
    window.show()

    exit_code = app.exec_()
    logger.info("تم إغلاق التطبيق (exit code: %d)", exit_code)
    sys.exit(exit_code)
