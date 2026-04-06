#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# database.py

import sqlite3
import bcrypt
import os
import json
import logging
from contextlib import contextmanager
from datetime import datetime
from typing import Optional, Any, List, Tuple

from constants import DATABASE_NAME, BACKUP_FOLDER

logger = logging.getLogger(__name__)

# ==================== جدول الـ Migrations ====================
# كل إدخال: (رقم_الإصدار, SQL)
# أضف هنا فقط — لا تعدّل ما هو موجود — الترقيم تصاعدي دائماً
MIGRATIONS: List[Tuple[int, str]] = [
    (1,  "ALTER TABLE leave_types ADD COLUMN is_annual INTEGER DEFAULT 0"),
    (2,  "ALTER TABLE leave_types ADD COLUMN max_requests INTEGER"),
    (3,  "ALTER TABLE employees ADD COLUMN bank_salary REAL DEFAULT 0"),
    (4,  "ALTER TABLE employees ADD COLUMN cash_salary REAL DEFAULT 0"),
    (5,  "ALTER TABLE employees ADD COLUMN is_exempt_from_fingerprint INTEGER DEFAULT 0"),
    (6,  "ALTER TABLE employees ADD COLUMN social_security_date TEXT"),
    (7,  "ALTER TABLE payroll ADD COLUMN bank_salary REAL DEFAULT 0"),
    (8,  "ALTER TABLE payroll ADD COLUMN cash_salary REAL DEFAULT 0"),
    (9,  "ALTER TABLE payroll ADD COLUMN is_archived INTEGER DEFAULT 0"),
    (10, "ALTER TABLE payroll ADD COLUMN approved_at TEXT"),
    (11, "ALTER TABLE payroll ADD COLUMN unpaid_leave_deduction REAL DEFAULT 0"),
    (12, "ALTER TABLE installments ADD COLUMN notes TEXT"),
    (13, "ALTER TABLE attendance ADD COLUMN is_approved INTEGER DEFAULT 0"),
    (14, "ALTER TABLE attendance ADD COLUMN approved_at TEXT"),
    (15, "ALTER TABLE attendance ADD COLUMN approved_by INTEGER"),
]


class DatabaseManager:
    """
    مدير قاعدة البيانات الرئيسي للنظام.
    يتعامل مع جميع عمليات قاعدة البيانات، إنشاء الجداول، الفهارس،
    وإدارة السجلات مع تسجيل الحركات.

    الإصلاحات المطبقة في هذه النسخة:
    - إصلاح خطأ (if params) ← (if params is not None) في execute_query/fetch_all/fetch_one
    - إضافة فلاج _in_transaction لمنع الـ commit المبكر داخل المعاملات
    - إصلاح transaction() context manager ليعمل بشكل صحيح مع rollback
    - إصلاح _insert_defaults: bcrypt لا يُشغَّل إلا عند الحاجة الفعلية
    - إضافة execute_many للإدراج الجماعي بكفاءة
    - إضافة close() لإغلاق نظيف عند إنهاء التطبيق
    - إضافة migration رقم 15 لعمود approved_by في attendance
    """

    def __init__(self, db_name: str = DATABASE_NAME):
        self.db_name             = db_name
        self.conn                = None
        self.cursor              = None
        self.current_user_id: Optional[int] = None
        # فلاج يمنع execute_query من الـ commit أثناء معاملة صريحة
        self._in_transaction: bool = False
        self._connect()
        self.create_tables()
        self._run_migrations()
        self._create_indexes()
        self._insert_defaults()
        self._safe_remove_department_text_column()

    # ==================== الاتصال ====================
    def _connect(self) -> None:
        """فتح الاتصال بقاعدة البيانات مع تفعيل المفاتيح الخارجية و WAL."""
        try:
            self.conn = sqlite3.connect(self.db_name, check_same_thread=False)
            self.conn.execute("PRAGMA foreign_keys = ON")
            self.conn.execute("PRAGMA journal_mode = WAL")
            self.conn.row_factory = sqlite3.Row   # يتيح الوصول بالاسم لاحقاً
            self.cursor = self.conn.cursor()
            logger.debug("تم الاتصال بقاعدة البيانات: %s", self.db_name)
        except sqlite3.Error as e:
            logger.critical("فشل الاتصال بقاعدة البيانات: %s", e, exc_info=True)
            raise

    # alias للتوافق مع الكود القديم
    def connect(self) -> None:
        self._connect()

    def close(self) -> None:
        """إغلاق الاتصال بقاعدة البيانات بشكل نظيف."""
        try:
            if self.conn:
                self.conn.close()
                self.conn   = None
                self.cursor = None
                logger.info("تم إغلاق الاتصال بقاعدة البيانات")
        except Exception as e:
            logger.error("خطأ عند إغلاق قاعدة البيانات: %s", e)

    # ==================== استعلامات أساسية ====================
    def execute_query(self, query: str, params: Optional[tuple] = None) -> bool:
        """
        تنفيذ استعلام (إدراج، تحديث، حذف).
        - يُرجع True عند النجاح، False عند الفشل.
        - لا يُنفِّذ commit إذا كنا داخل transaction() صريح،
          بل يترك ذلك لـ transaction() نفسه.

        الإصلاح: استبدال (if params) بـ (if params is not None)
        لمعالجة حالة تمرير tuple فارغ () بشكل صحيح.
        """
        try:
            if params is not None:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            # لا تُنفِّذ commit إذا كنا داخل معاملة صريحة
            if not self._in_transaction:
                self.conn.commit()
            return True
        except sqlite3.Error as e:
            logger.error("خطأ في تنفيذ الاستعلام: %s\n%s", e, query, exc_info=True)
            if not self._in_transaction:
                self.conn.rollback()
            return False

    def execute_many(self, query: str, params_list: List[tuple]) -> bool:
        """
        تنفيذ استعلام بشكل جماعي (مثل إدراج قائمة من السجلات دفعة واحدة).
        أكفأ بكثير من استدعاء execute_query في حلقة.

        مثال:
            self.db.execute_many(
                "INSERT INTO payroll (employee_id, month) VALUES (?, ?)",
                [(1, 3), (2, 3), (3, 3)]
            )
        """
        try:
            self.cursor.executemany(query, params_list)
            if not self._in_transaction:
                self.conn.commit()
            return True
        except sqlite3.Error as e:
            logger.error("خطأ في execute_many: %s\n%s", e, query, exc_info=True)
            if not self._in_transaction:
                self.conn.rollback()
            return False

    def fetch_all(self, query: str, params: Optional[tuple] = None) -> List[tuple]:
        """
        جلب جميع النتائج لاستعلام SELECT.
        الإصلاح: (if params is not None) بدلاً من (if params).
        """
        try:
            if params is not None:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            logger.error("خطأ في جلب البيانات: %s", e, exc_info=True)
            return []

    def fetch_one(self, query: str, params: Optional[tuple] = None) -> Optional[tuple]:
        """
        جلب نتيجة واحدة لاستعلام SELECT.
        الإصلاح: (if params is not None) بدلاً من (if params).
        """
        try:
            if params is not None:
                self.cursor.execute(query, params)
            else:
                self.cursor.execute(query)
            return self.cursor.fetchone()
        except sqlite3.Error as e:
            logger.error("خطأ في جلب البيانات: %s", e, exc_info=True)
            return None

    def last_id(self) -> Optional[int]:
        """
        إرجاع آخر ID تم إدراجه.
        تنبيه: استخدم هذه الدالة مباشرة بعد execute_query INSERT
        وقبل أي استعلام آخر.
        """
        return self.cursor.lastrowid

    def table_exists(self, table_name: str) -> bool:
        """التحقق من وجود جدول في قاعدة البيانات."""
        row = self.fetch_one(
            "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
            (table_name,)
        )
        return row is not None

    def column_exists(self, table_name: str, column_name: str) -> bool:
        """التحقق من وجود عمود معين في جدول."""
        cols = [c[1] for c in self.fetch_all(f"PRAGMA table_info({table_name})")]
        return column_name in cols

    # ==================== Transaction Context Manager ====================
    @contextmanager
    def transaction(self):
        """
        Context manager للعمليات الجماعية الذرية.

        - يضبط _in_transaction=True لمنع execute_query من الـ commit المبكر.
        - يُنفِّذ commit واحداً عند نهاية الـ block ناجحاً.
        - يُنفِّذ rollback كاملاً عند أي استثناء.
        - يُعيد _in_transaction=False دائماً في نهاية الـ block.

        الاستخدام:
            with self.db.transaction():
                self.db.execute_query("INSERT INTO ...")
                self.db.execute_query("UPDATE ...")
                # إذا رُمي استثناء هنا → rollback كامل لكلا الاستعلامين
        """
        self._in_transaction = True
        try:
            yield self
            self.conn.commit()
            logger.debug("تم تأكيد الـ Transaction بنجاح")
        except Exception as e:
            self.conn.rollback()
            logger.error("تم التراجع عن الـ Transaction بسبب: %s", e, exc_info=True)
            raise
        finally:
            # يُعيد الفلاج دائماً بغض النظر عن النتيجة
            self._in_transaction = False

    # ==================== إنشاء الجداول ====================
    def create_tables(self) -> None:
        """إنشاء جميع جداول قاعدة البيانات إذا لم تكن موجودة."""
        tables = [
            # جدول إصدارات الـ Schema (يُنشأ أولاً)
            '''CREATE TABLE IF NOT EXISTS schema_version (
                version     INTEGER PRIMARY KEY,
                applied_at  TEXT DEFAULT CURRENT_TIMESTAMP
            )''',

            # جدول المستخدمين
            '''CREATE TABLE IF NOT EXISTS users (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                username        TEXT UNIQUE NOT NULL,
                password_hash   TEXT NOT NULL,
                full_name       TEXT NOT NULL,
                role            TEXT NOT NULL DEFAULT 'hr',
                is_active       INTEGER DEFAULT 1,
                created_at      TEXT DEFAULT CURRENT_TIMESTAMP
            )''',

            # جدول الأقسام
            '''CREATE TABLE IF NOT EXISTS departments (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                name        TEXT UNIQUE NOT NULL,
                manager_id  INTEGER,
                parent_id   INTEGER,
                notes       TEXT
            )''',

            # جدول الموظفين
            '''CREATE TABLE IF NOT EXISTS employees (
                id                          INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_code               TEXT UNIQUE NOT NULL,
                first_name                  TEXT NOT NULL,
                last_name                   TEXT NOT NULL,
                national_id                 TEXT,
                nationality                 TEXT DEFAULT 'تركي',
                birth_date                  TEXT,
                gender                      TEXT DEFAULT 'ذكر',
                position                    TEXT,
                department_id               INTEGER,
                hire_date                   TEXT NOT NULL,
                contract_type               TEXT DEFAULT 'دوام كامل',
                basic_salary                REAL NOT NULL DEFAULT 0,
                housing_allowance           REAL DEFAULT 0,
                transportation_allowance    REAL DEFAULT 0,
                food_allowance              REAL DEFAULT 0,
                phone_allowance             REAL DEFAULT 0,
                other_allowances            REAL DEFAULT 0,
                bank_salary                 REAL DEFAULT 0,
                cash_salary                 REAL DEFAULT 0,
                phone                       TEXT,
                email                       TEXT,
                address                     TEXT,
                bank_name                   TEXT,
                bank_account                TEXT,
                iban                        TEXT,
                fingerprint_id              TEXT UNIQUE NOT NULL,
                social_security_number      TEXT,
                social_security_registered  INTEGER DEFAULT 0,
                social_security_percent     REAL DEFAULT 9.75,
                social_security_date        TEXT,
                iqama_number                TEXT,
                iqama_expiry                TEXT,
                passport_number             TEXT,
                passport_expiry             TEXT,
                work_permit_number          TEXT,
                work_permit_expiry          TEXT,
                health_insurance_number     TEXT,
                health_insurance_expiry     TEXT,
                status                      TEXT DEFAULT 'نشط',
                notes                       TEXT,
                created_at                  TEXT DEFAULT CURRENT_TIMESTAMP,
                is_exempt_from_fingerprint  INTEGER DEFAULT 0,
                FOREIGN KEY (department_id) REFERENCES departments(id) ON DELETE SET NULL
            )''',

            # جدول الوثائق
            '''CREATE TABLE IF NOT EXISTS documents (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id     INTEGER NOT NULL,
                document_name   TEXT NOT NULL,
                document_path   TEXT NOT NULL,
                document_type   TEXT,
                expiry_date     TEXT,
                upload_date     TEXT NOT NULL,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
            )''',

            # جدول الحضور
            # is_approved: 0=مسودة، 1=معتمد
            # approved_at: تاريخ ووقت الاعتماد
            # approved_by: معرف المستخدم الذي اعتمد السجل
            '''CREATE TABLE IF NOT EXISTS attendance (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id         INTEGER,
                fingerprint_id      TEXT,
                punch_date          TEXT NOT NULL,
                check_in            TEXT,
                check_out           TEXT,
                work_hours          REAL DEFAULT 0,
                overtime_hours      REAL DEFAULT 0,
                late_minutes        INTEGER DEFAULT 0,
                early_leave_minutes INTEGER DEFAULT 0,
                status              TEXT DEFAULT 'حاضر',
                notes               TEXT,
                is_approved         INTEGER DEFAULT 0,
                approved_at         TEXT,
                approved_by         INTEGER,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE SET NULL,
                FOREIGN KEY (approved_by) REFERENCES users(id)     ON DELETE SET NULL
            )''',

            # جدول البصمات الخام
            '''CREATE TABLE IF NOT EXISTS fingerprint_raw (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                fingerprint_id  TEXT NOT NULL,
                punch_datetime  TEXT NOT NULL,
                punch_type      TEXT,
                source_file     TEXT,
                processed       INTEGER DEFAULT 0,
                imported_at     TEXT DEFAULT CURRENT_TIMESTAMP
            )''',

            # جدول أنواع الإجازات
            '''CREATE TABLE IF NOT EXISTS leave_types (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                name            TEXT UNIQUE NOT NULL,
                days_per_year   INTEGER DEFAULT 0,
                paid            INTEGER DEFAULT 1,
                carry_over      INTEGER DEFAULT 0,
                is_annual       INTEGER DEFAULT 0,
                max_requests    INTEGER,
                notes           TEXT
            )''',

            # جدول أرصدة الإجازات
            '''CREATE TABLE IF NOT EXISTS leave_balance (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id     INTEGER NOT NULL,
                leave_type_id   INTEGER NOT NULL,
                year            INTEGER NOT NULL,
                total_days      REAL DEFAULT 0,
                used_days       REAL DEFAULT 0,
                pending_days    REAL DEFAULT 0,
                UNIQUE(employee_id, leave_type_id, year),
                FOREIGN KEY (employee_id)   REFERENCES employees(id)   ON DELETE CASCADE,
                FOREIGN KEY (leave_type_id) REFERENCES leave_types(id) ON DELETE CASCADE
            )''',

            # جدول طلبات الإجازات
            '''CREATE TABLE IF NOT EXISTS leave_requests (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id     INTEGER NOT NULL,
                leave_type_id   INTEGER NOT NULL,
                start_date      TEXT NOT NULL,
                end_date        TEXT NOT NULL,
                days_count      REAL NOT NULL,
                reason          TEXT,
                status          TEXT DEFAULT 'قيد المراجعة',
                approved_by     INTEGER,
                approved_at     TEXT,
                notes           TEXT,
                created_at      TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (employee_id)   REFERENCES employees(id)   ON DELETE CASCADE,
                FOREIGN KEY (leave_type_id) REFERENCES leave_types(id) ON DELETE CASCADE
            )''',

            # جدول السلف
            '''CREATE TABLE IF NOT EXISTS loans (
                id                  INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id         INTEGER NOT NULL,
                loan_type           TEXT DEFAULT 'سلفة',
                amount              REAL NOT NULL,
                monthly_installment REAL NOT NULL,
                remaining_amount    REAL NOT NULL,
                start_month         INTEGER NOT NULL,
                start_year          INTEGER NOT NULL,
                total_installments  INTEGER,
                status              TEXT DEFAULT 'نشط',
                payment_method      TEXT DEFAULT 'cash',
                notes               TEXT,
                created_at          TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
            )''',

            # جدول أقساط السلف
            '''CREATE TABLE IF NOT EXISTS installments (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                loan_id     INTEGER NOT NULL,
                due_date    TEXT NOT NULL,
                amount      REAL NOT NULL,
                paid_amount REAL DEFAULT 0,
                status      TEXT DEFAULT 'pending',
                paid_date   TEXT,
                payroll_id  INTEGER,
                notes       TEXT,
                created_at  TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (loan_id)    REFERENCES loans(id)   ON DELETE CASCADE,
                FOREIGN KEY (payroll_id) REFERENCES payroll(id) ON DELETE SET NULL
            )''',

            # جدول الرواتب
            '''CREATE TABLE IF NOT EXISTS payroll (
                id                       INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id              INTEGER NOT NULL,
                month                    INTEGER NOT NULL,
                year                     INTEGER NOT NULL,
                work_days                INTEGER DEFAULT 0,
                actual_days              REAL DEFAULT 0,
                basic_salary             REAL DEFAULT 0,
                housing_allowance        REAL DEFAULT 0,
                transportation_allowance REAL DEFAULT 0,
                food_allowance           REAL DEFAULT 0,
                phone_allowance          REAL DEFAULT 0,
                other_allowances         REAL DEFAULT 0,
                overtime_hours           REAL DEFAULT 0,
                overtime_amount          REAL DEFAULT 0,
                bonus                    REAL DEFAULT 0,
                total_earnings           REAL DEFAULT 0,
                absence_deduction        REAL DEFAULT 0,
                late_deduction           REAL DEFAULT 0,
                loan_deduction           REAL DEFAULT 0,
                loan_deduction_bank      REAL DEFAULT 0,
                loan_deduction_cash      REAL DEFAULT 0,
                social_security          REAL DEFAULT 0,
                other_deductions         REAL DEFAULT 0,
                unpaid_leave_deduction   REAL DEFAULT 0,
                total_deductions         REAL DEFAULT 0,
                net_salary               REAL DEFAULT 0,
                bank_salary              REAL DEFAULT 0,
                cash_salary              REAL DEFAULT 0,
                status                   TEXT DEFAULT 'مسودة',
                notes                    TEXT,
                is_archived              INTEGER DEFAULT 0,
                approved_at              TEXT,
                created_at               TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(employee_id, month, year),
                FOREIGN KEY (employee_id) REFERENCES employees(id) ON DELETE CASCADE
            )''',

            # جدول الإعدادات
            '''CREATE TABLE IF NOT EXISTS settings (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                setting_name    TEXT UNIQUE NOT NULL,
                setting_value   TEXT
            )''',

            # جدول سجل الأحداث
            '''CREATE TABLE IF NOT EXISTS audit_log (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id     INTEGER,
                action      TEXT NOT NULL,
                table_name  TEXT,
                record_id   INTEGER,
                old_value   TEXT,
                new_value   TEXT,
                created_at  TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users(id) ON DELETE SET NULL
            )''',

            # جدول قواعد الإجازات السنوية
            '''CREATE TABLE IF NOT EXISTS leave_rules (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                from_year   INTEGER NOT NULL,
                to_year     INTEGER NOT NULL,
                days        INTEGER NOT NULL,
                notes       TEXT
            )''',

            # جدول العطل الرسمية
            '''CREATE TABLE IF NOT EXISTS holidays (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                name            TEXT NOT NULL,
                holiday_date    TEXT NOT NULL,
                type            TEXT NOT NULL,
                year            INTEGER,
                notes           TEXT
            )''',

            # جدول ملحقات الرواتب
            '''CREATE TABLE IF NOT EXISTS payroll_attachments (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                payroll_id  INTEGER NOT NULL,
                amount      REAL NOT NULL,
                reason      TEXT,
                notes       TEXT,
                created_by  INTEGER,
                created_at  TEXT,
                FOREIGN KEY (payroll_id) REFERENCES payroll(id) ON DELETE CASCADE
            )''',

            # جدول الإيصالات
            '''CREATE TABLE IF NOT EXISTS receipts (
                id              INTEGER PRIMARY KEY AUTOINCREMENT,
                receipt_number  TEXT UNIQUE,
                receiver_name   TEXT NOT NULL,
                amount          REAL NOT NULL,
                amount_words    TEXT,
                payment_date    TEXT NOT NULL,
                payment_type    TEXT,
                description     TEXT,
                notes           TEXT,
                created_by      INTEGER,
                created_at      TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (created_by) REFERENCES users(id) ON DELETE SET NULL
            )''',
        ]

        for sql in tables:
            try:
                self.cursor.execute(sql)
            except sqlite3.Error as e:
                logger.error("خطأ في إنشاء الجدول: %s", e, exc_info=True)
        self.conn.commit()

    # ==================== Schema Versioning ====================
    def _run_migrations(self) -> None:
        """
        تطبيق Migrations المعلقة فقط.
        كل migration ترقّم وتُنفَّذ مرة واحدة — لا تكرار.
        """
        row = self.fetch_one("SELECT MAX(version) FROM schema_version")
        current_version = row[0] if row and row[0] else 0

        pending = [(v, sql) for v, sql in MIGRATIONS if v > current_version]
        if not pending:
            logger.debug("لا توجد migrations معلقة (الإصدار الحالي: %d)", current_version)
            return

        for version, sql in sorted(pending):
            try:
                self.cursor.execute(sql)
                self.cursor.execute(
                    "INSERT INTO schema_version (version) VALUES (?)", (version,))
                self.conn.commit()
                logger.info("✅ تم تطبيق Migration رقم %d", version)
            except sqlite3.OperationalError as e:
                # إذا كان العمود موجوداً مسبقاً → تجاوز وتسجيل الإصدار
                err_msg = str(e).lower()
                if "duplicate column" in err_msg or "already exists" in err_msg:
                    self.cursor.execute(
                        "INSERT OR IGNORE INTO schema_version (version) VALUES (?)",
                        (version,))
                    self.conn.commit()
                    logger.debug("Migration %d: العمود موجود مسبقاً — تجاوز", version)
                else:
                    logger.error("❌ فشل Migration %d: %s", version, e, exc_info=True)

    # ==================== حذف عمود department القديم ====================
    def _safe_remove_department_text_column(self) -> None:
        """إزالة العمود النصي department إذا كان موجوداً."""
        try:
            cols = [c[1] for c in self.fetch_all("PRAGMA table_info(employees)")]
            if 'department' not in cols:
                return
            if sqlite3.sqlite_version_info >= (3, 35, 0):
                self.cursor.execute("ALTER TABLE employees DROP COLUMN department")
                self.conn.commit()
                logger.info("تم حذف العمود القديم department من employees")
            else:
                logger.warning(
                    "إصدار SQLite لا يدعم DROP COLUMN — عمود department سيُتجاهل")
        except Exception as e:
            logger.error("خطأ في حذف عمود department: %s", e, exc_info=True)

    # ==================== الفهارس ====================
    def _create_indexes(self) -> None:
        """إنشاء الفهارس لتحسين الأداء."""
        indexes = [
            "CREATE INDEX IF NOT EXISTS idx_attendance_emp_date    ON attendance(employee_id, punch_date)",
            "CREATE INDEX IF NOT EXISTS idx_attendance_status       ON attendance(status)",
            "CREATE INDEX IF NOT EXISTS idx_attendance_approved     ON attendance(is_approved)",
            "CREATE INDEX IF NOT EXISTS idx_fingerprint_processed   ON fingerprint_raw(processed)",
            "CREATE INDEX IF NOT EXISTS idx_fingerprint_id          ON fingerprint_raw(fingerprint_id)",
            "CREATE INDEX IF NOT EXISTS idx_payroll_month_year      ON payroll(month, year)",
            "CREATE INDEX IF NOT EXISTS idx_payroll_archived        ON payroll(is_archived)",
            "CREATE INDEX IF NOT EXISTS idx_leave_requests_status   ON leave_requests(status)",
            "CREATE INDEX IF NOT EXISTS idx_leave_balance_year      ON leave_balance(year)",
            "CREATE INDEX IF NOT EXISTS idx_installments_loan       ON installments(loan_id)",
            "CREATE INDEX IF NOT EXISTS idx_installments_due        ON installments(due_date)",
            "CREATE INDEX IF NOT EXISTS idx_installments_payroll    ON installments(payroll_id)",
            "CREATE INDEX IF NOT EXISTS idx_attachments_payroll     ON payroll_attachments(payroll_id)",
            "CREATE INDEX IF NOT EXISTS idx_employees_exempt        ON employees(is_exempt_from_fingerprint)",
            "CREATE INDEX IF NOT EXISTS idx_employees_status        ON employees(status)",
            "CREATE INDEX IF NOT EXISTS idx_leave_requests_dates    ON leave_requests(start_date, end_date)",
            "CREATE INDEX IF NOT EXISTS idx_fingerprint_raw_date    ON fingerprint_raw(punch_datetime)",
            "CREATE INDEX IF NOT EXISTS idx_audit_log_created       ON audit_log(created_at)",
            "CREATE INDEX IF NOT EXISTS idx_audit_log_user          ON audit_log(user_id)",
        ]
        for sql in indexes:
            try:
                self.cursor.execute(sql)
            except Exception as e:
                logger.error("خطأ في إنشاء الفهرس: %s", e)
        self.conn.commit()

    # ==================== القيم الافتراضية ====================
    def _insert_defaults(self) -> None:
        """
        إدراج القيم الافتراضية (مستخدم، إعدادات، أنواع الإجازات).

        الإصلاح: bcrypt.hashpw() بطيء بتصميمه — لا يُستدعى إلا إذا
        لم يكن مستخدم admin موجوداً فعلاً في قاعدة البيانات.
        """
        # المستخدم الافتراضي admin — فقط عند الإنشاء الأول
        existing = self.fetch_one(
            "SELECT id FROM users WHERE username = ?", ("admin",))
        if not existing:
            salt    = bcrypt.gensalt()
            pw_hash = bcrypt.hashpw(b"admin123", salt).decode()
            self.cursor.execute(
                "INSERT OR IGNORE INTO users "
                "(username, password_hash, full_name, role) VALUES (?,?,?,?)",
                ("admin", pw_hash, "مدير النظام", "admin")
            )
            logger.info("تم إنشاء المستخدم الافتراضي admin")

        # الإعدادات الافتراضية
        defaults = [
            ('company_name',           'الشركة'),
            ('company_address',        ''),
            ('company_phone',          ''),
            ('working_hours',          '8'),
            ('work_start_time',        '08:00'),
            ('work_end_time',          '17:00'),
            ('overtime_rate',          '1.5'),
            ('late_tolerance_minutes', '10'),
            ('work_days_month',        '26'),
            ('gosi_employee_percent',  '9.75'),
            ('gosi_company_percent',   '12.0'),
            ('absence_daily_calc',     'salary_div_days'),
            ('currency',               'TRY'),
            ('absence_deduction_rate', '1.0'),
            ('late_tolerance_type',    '0'),
            ('lunch_break',            '30'),
            ('num_breaks',             '2'),
            ('break_duration',         '15'),
            ('include_breaks',         '0'),
            ('work_days',              '0,1,2,3,4'),
            ('gosi_max_salary',        '45000'),
            ('rounding',               '1'),
            ('language',               'ar'),
        ]
        self.cursor.executemany(
            "INSERT OR IGNORE INTO settings (setting_name, setting_value) VALUES (?,?)",
            defaults
        )

        # أنواع الإجازات الافتراضية
        leave_types = [
            ('سنوية',     21, 1, 1, 1, None),
            ('مرضية',     30, 1, 0, 0, None),
            ('طارئة',      5, 1, 0, 0, None),
            ('بدون راتب',  0, 0, 0, 0, None),
            ('أمومة',     70, 1, 0, 0, None),
            ('رسمية',      0, 1, 0, 0, None),
            ('زواج',       5, 1, 0, 0, 1),
        ]
        self.cursor.executemany(
            "INSERT OR IGNORE INTO leave_types "
            "(name, days_per_year, paid, carry_over, is_annual, max_requests) "
            "VALUES (?,?,?,?,?,?)",
            leave_types
        )

        self.conn.commit()

    # ==================== الإعدادات ====================
    def get_setting(self, name: str, default: str = '') -> str:
        """جلب قيمة إعداد معين."""
        row = self.fetch_one(
            "SELECT setting_value FROM settings WHERE setting_name=?", (name,))
        return row[0] if row else default

    def set_setting(self, name: str, value: Any) -> None:
        """تعيين قيمة إعداد (إدراج أو تحديث)."""
        self.execute_query(
            "INSERT OR REPLACE INTO settings (setting_name, setting_value) VALUES (?,?)",
            (name, str(value))
        )

    # ==================== النسخ الاحتياطي الآمن ====================
    def backup(self) -> Optional[str]:
        """
        إنشاء نسخة احتياطية آمنة باستخدام SQLite Online Backup API.
        يتعامل بشكل صحيح مع WAL mode ويضمن اتساق البيانات.
        """
        try:
            ts   = datetime.now().strftime('%Y%m%d_%H%M%S')
            path = os.path.join(BACKUP_FOLDER, f"backup_{ts}.db")
            dest = sqlite3.connect(path)
            # pages=100: نسخ 100 صفحة في كل خطوة (أداء جيد + غير تعطيلي)
            self.conn.backup(dest, pages=100)
            dest.close()
            logger.info("✅ تم إنشاء النسخة الاحتياطية: %s", path)
            return path
        except Exception as e:
            logger.error("❌ فشل النسخ الاحتياطي: %s", e, exc_info=True)
            return None

    # ==================== تسجيل الأحداث ====================
    def log_action(
        self,
        action:     str,
        table_name: Optional[str] = None,
        record_id:  Optional[int] = None,
        old_value:  Any = None,
        new_value:  Any = None
    ) -> None:
        """
        تسجيل حركة في سجل الأحداث.
        لا تُسجِّل إذا لم يكن هناك مستخدم محدد (قبل تسجيل الدخول).
        """
        if self.current_user_id is None:
            return

        def _serialize(val: Any) -> Optional[str]:
            if val is None:
                return None
            if isinstance(val, str):
                return val
            try:
                return json.dumps(val, ensure_ascii=False)
            except Exception:
                return str(val)

        self.execute_query(
            """INSERT INTO audit_log
               (user_id, action, table_name, record_id, old_value, new_value, created_at)
               VALUES (?, ?, ?, ?, ?, ?, ?)""",
            (self.current_user_id, action, table_name, record_id,
             _serialize(old_value), _serialize(new_value),
             datetime.now().isoformat())
        )

    def log_insert(self, table_name: str, record_id: int, values: Any) -> None:
        """تسجيل عملية إضافة."""
        self.log_action("إضافة", table_name, record_id, None, values)

    def log_update(self, table_name: str, record_id: int,
                   old_values: Any, new_values: Any) -> None:
        """تسجيل عملية تعديل."""
        self.log_action("تعديل", table_name, record_id, old_values, new_values)

    def log_delete(self, table_name: str, record_id: int, old_values: Any) -> None:
        """تسجيل عملية حذف."""
        self.log_action("حذف", table_name, record_id, old_values, None)

    def log_custom(self, action: str, table_name: Optional[str] = None,
                   record_id: Optional[int] = None, details: Any = None) -> None:
        """تسجيل حركة مخصصة."""
        self.log_action(action, table_name, record_id, None, details)

    def set_current_user(self, user_id: int) -> None:
        """تعيين المستخدم الحالي لتسجيل الحركات."""
        self.current_user_id = user_id
