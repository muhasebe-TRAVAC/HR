#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/payroll_tab.py

import os
import logging
from datetime import datetime, date, timedelta

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QSpinBox,
    QPushButton, QMessageBox, QTableWidgetItem, QGroupBox, QHeaderView,
    QFileDialog, QDialog, QFormLayout, QDoubleSpinBox, QDialogButtonBox,
    QLineEdit, QTextEdit, QCompleter, QTabWidget, QAbstractItemView,
    QMenu, QAction, QProgressDialog, QTableWidget
)
from PyQt5.QtPrintSupport import QPrintDialog, QPrinter, QPrintPreviewDialog
from PyQt5.QtGui import QTextDocument, QColor
from PyQt5.QtCore import Qt, QDate

from database import DatabaseManager
from utils import (make_table, fill_table, btn,
                   can_edit, can_process_payroll,
                   number_to_words_tr)          # ← استيراد من utils (لا تكرار)
from constants import (BTN_SUCCESS, BTN_PRIMARY, BTN_DANGER,
                       BTN_GRAY, BTN_PURPLE, BTN_TEAL, BTN_WARNING)

logger = logging.getLogger(__name__)

# أعمدة الجدول الثابتة
PAYROLL_COLUMNS = [
    "#",                    # 0  (مخفي - payroll_id)
    "م",                    # 1  (رقم تسلسلي)
    "اسم الموظف",           # 2
    "الراتب الأساسي",       # 3
    "البدلات",              # 4
    "أوفرتايم",             # 5
    "مكافآت",               # 6
    "إجمالي الإضافات",      # 7
    "غياب",                 # 8
    "إجازات بدون راتب",     # 9
    "تأخير",                # 10
    "إجمالي الخصومات",      # 11
    "إجمالي الاستحقاق",     # 12
    "سلفة بنكية",           # 13
    "سلفة نقدية",           # 14
    "صافي الراتب للدفع",    # 15
    "راتب بنكي",            # 16
    "راتب نقدي",            # 17
    "ملاحظات",              # 18
    "الحالة"                # 19
]


class PayrollTab(QWidget):
    """
    تبويب الرواتب.

    الإصلاحات في هذه النسخة:
    - حُذفت نسخة number_to_words_tr المكررة — تُستورَد من utils الآن.
    - Guard في _calculate(): لا يُسمَح بحساب رواتب شهر إلا إذا وُجد
      حضور معتمد (is_approved=1) لذلك الشهر.
    - _get_attendance_data() تجلب فقط السجلات المعتمدة (is_approved=1).
    - المعاملات تستخدم with self.db.transaction() بدلاً من BEGIN/COMMIT اليدوي.
    - إصلاح total_deductions: يُحسَب بشكل صريح ومتسق في كل مكان.
    """

    def __init__(self, db: DatabaseManager, user: dict, comm=None):
        super().__init__()
        self.db           = db
        self.user         = user
        self.comm         = comm
        self._current_ids  = []
        self._archived_ids = []
        self._reload_settings()
        self._build()
        self._load_current()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    # ==================== الإعدادات ====================
    def _reload_settings(self):
        self.work_days_month  = int(self.db.get_setting('work_days_month', '26'))
        self.work_hours_daily = float(self.db.get_setting('working_hours', '8'))
        self.ot_rate          = float(self.db.get_setting('overtime_rate', '1.5'))
        self.absence_rate     = float(self.db.get_setting('absence_deduction_rate', '1.0'))
        self.late_tol         = int(self.db.get_setting('late_tolerance_minutes', '10'))
        self.late_tol_type    = int(self.db.get_setting('late_tolerance_type', '0'))
        self.rounding         = int(self.db.get_setting('rounding', '1'))

    def _on_data_changed(self, data_type: str, data):
        if data_type == 'employee':
            self._refresh_dept_filter()
            self._refresh_dept_filter_archived()
        elif data_type == 'payroll':
            self._load_current()
            self._load_archived()
        elif data_type == 'settings':
            self._reload_settings()
            self._load_current()
        elif data_type in ('installment', 'loan'):
            self._load_current()

    # ==================== بناء الواجهة ====================
    def _build(self):
        main_layout = QVBoxLayout(self)
        self.tabs = QTabWidget()
        self.tabs.currentChanged.connect(self._on_tab_changed)

        self.current_tab  = QWidget()
        self.archived_tab = QWidget()
        self._build_current_tab()
        self._build_archived_tab()
        self.tabs.addTab(self.current_tab,  "📋 الرواتب الحالية")
        self.tabs.addTab(self.archived_tab, "📁 الرواتب المعتمدة")

        main_layout.addWidget(self.tabs)
        self._apply_permissions()

    # ---------- تبويب الرواتب الحالية ----------
    def _build_current_tab(self):
        layout = QVBoxLayout(self.current_tab)

        tools = QHBoxLayout()
        self.btn_calculate         = btn("⚙️ حساب الرواتب",         BTN_SUCCESS, self._calculate)
        self.btn_approve_single    = btn("✅ اعتماد المحدد",         BTN_PRIMARY, self._approve_selected)
        self.btn_approve_all       = btn("✅ اعتماد الكل",           BTN_PRIMARY, self._approve_all_current)
        self.btn_edit_installments = btn("💰 تعديل القسط المحسوب",  BTN_WARNING, self._edit_selected)
        self.btn_add_bonus         = btn("🎁 إضافة مكافأة",          BTN_TEAL,   self._add_bonus)
        self.btn_refresh_current   = btn("🔄 تحديث",                 BTN_GRAY,   self._load_current)
        self.btn_print_current     = btn("🖨️ طباعة",               BTN_TEAL,   self._print_current_payroll)
        self.btn_export_current    = btn("📊 تصدير Excel",           BTN_PURPLE, self._export_current_excel)
        self.btn_payslip           = btn("🧾 قصاصات الراتب",        BTN_PURPLE, self._print_payslip)
        self.btn_print_all_receipts= btn("🖨️ طباعة جميع الإيصالات",BTN_TEAL,   self._print_all_receipts)

        for b in (self.btn_calculate, self.btn_approve_single, self.btn_approve_all,
                  self.btn_edit_installments, self.btn_add_bonus,
                  self.btn_refresh_current, self.btn_print_current,
                  self.btn_export_current, self.btn_payslip,
                  self.btn_print_all_receipts):
            tools.addWidget(b)
        tools.addStretch()
        layout.addLayout(tools)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("الشهر:"))
        self.current_month = QComboBox()
        self.current_month.addItems([
            'يناير','فبراير','مارس','أبريل','مايو','يونيو',
            'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'])
        self.current_month.setCurrentIndex(date.today().month - 1)
        filter_layout.addWidget(self.current_month)

        filter_layout.addWidget(QLabel("السنة:"))
        self.current_year = QSpinBox()
        self.current_year.setRange(2020, 2050)
        self.current_year.setValue(date.today().year)
        filter_layout.addWidget(self.current_year)

        filter_layout.addWidget(QLabel("القسم:"))
        self.current_dept = QComboBox()
        self.current_dept.addItem("جميع الأقسام", None)
        self._refresh_dept_filter()
        filter_layout.addWidget(self.current_dept)

        filter_layout.addWidget(QLabel("الحالة:"))
        self.current_status = QComboBox()
        self.current_status.addItems(["مسودة", "الكل"])
        filter_layout.addWidget(self.current_status)

        filter_layout.addWidget(QLabel("طريقة الدفع:"))
        self.current_payment = QComboBox()
        self.current_payment.addItems(["الكل", "بنكي", "نقدي"])
        filter_layout.addWidget(self.current_payment)

        filter_layout.addWidget(btn("تصفية", BTN_PRIMARY, self._load_current))
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        self.current_table = make_table(PAYROLL_COLUMNS)
        self.current_table.setColumnHidden(0, True)
        self.current_table.setSortingEnabled(False)
        self.current_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.current_table.customContextMenuRequested.connect(self._show_context_menu)
        self.current_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.current_table.setSelectionMode(QAbstractItemView.MultiSelection)
        self._set_column_widths(self.current_table)
        layout.addWidget(self.current_table)

        self._lbl_cur_count = QLabel("عدد الموظفين: 0")
        self._lbl_cur_basic  = QLabel("إجمالي الأساسي: 0")
        self._lbl_cur_net    = QLabel("إجمالي الصافي: 0")
        layout.addWidget(self._create_summary_group(
            self._lbl_cur_count, self._lbl_cur_basic, self._lbl_cur_net))

    # ---------- تبويب الرواتب المعتمدة ----------
    def _build_archived_tab(self):
        layout = QVBoxLayout(self.archived_tab)

        tools = QHBoxLayout()
        self.btn_unapprove_single = btn("⚠️ إلغاء اعتماد المحدد", BTN_DANGER, self._unapprove_selected)
        self.btn_unapprove_all    = btn("⚠️ إلغاء اعتماد الكل",   BTN_DANGER, self._unapprove_all_archived)
        self.btn_refresh_archived = btn("🔄 تحديث",                BTN_GRAY,   self._load_archived)
        self.btn_print_archived   = btn("🖨️ طباعة",               BTN_TEAL,   self._print_archived_payroll)
        self.btn_export_archived  = btn("📊 تصدير Excel",          BTN_PURPLE, self._export_archived_excel)
        self.btn_maintenance      = btn("🔧 صيانة الرواتب",        BTN_WARNING, self._run_maintenance)

        for b in (self.btn_unapprove_single, self.btn_unapprove_all,
                  self.btn_refresh_archived, self.btn_print_archived,
                  self.btn_export_archived, self.btn_maintenance):
            tools.addWidget(b)
        tools.addStretch()
        layout.addLayout(tools)

        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("الشهر:"))
        self.archived_month = QComboBox()
        self.archived_month.addItems([
            'يناير','فبراير','مارس','أبريل','مايو','يونيو',
            'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'])
        self.archived_month.setCurrentIndex(date.today().month - 1)
        filter_layout.addWidget(self.archived_month)

        filter_layout.addWidget(QLabel("السنة:"))
        self.archived_year = QSpinBox()
        self.archived_year.setRange(2020, 2050)
        self.archived_year.setValue(date.today().year)
        filter_layout.addWidget(self.archived_year)

        filter_layout.addWidget(QLabel("القسم:"))
        self.archived_dept = QComboBox()
        self.archived_dept.addItem("جميع الأقسام", None)
        self._refresh_dept_filter_archived()
        filter_layout.addWidget(self.archived_dept)

        filter_layout.addWidget(QLabel("طريقة الدفع:"))
        self.archived_payment = QComboBox()
        self.archived_payment.addItems(["الكل", "بنكي", "نقدي"])
        filter_layout.addWidget(self.archived_payment)

        filter_layout.addWidget(btn("تصفية", BTN_PRIMARY, self._load_archived))
        filter_layout.addStretch()
        layout.addLayout(filter_layout)

        self.archived_table = make_table(PAYROLL_COLUMNS)
        self.archived_table.setColumnHidden(0, True)
        self.archived_table.setSortingEnabled(True)
        self.archived_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.archived_table.customContextMenuRequested.connect(self._show_context_menu)
        self.archived_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.archived_table.setSelectionMode(QAbstractItemView.MultiSelection)
        self._set_column_widths(self.archived_table)
        layout.addWidget(self.archived_table)

        self._lbl_arc_count = QLabel("عدد الموظفين: 0")
        self._lbl_arc_basic  = QLabel("إجمالي الأساسي: 0")
        self._lbl_arc_net    = QLabel("إجمالي الصافي: 0")
        layout.addWidget(self._create_summary_group(
            self._lbl_arc_count, self._lbl_arc_basic, self._lbl_arc_net))

    # ==================== مساعدات الواجهة ====================
    def _set_column_widths(self, table):
        widths = [30, 160, 90, 85, 80, 70, 100, 75, 100, 70,
                  110, 110, 90, 90, 120, 90, 90, 140, 70]
        for i, w in enumerate(widths):
            table.setColumnWidth(i + 1, w)

    def _refresh_dept_filter(self):
        self.current_dept.clear()
        self.current_dept.addItem("جميع الأقسام", None)
        for did, name in self.db.fetch_all(
                "SELECT id, name FROM departments ORDER BY name"):
            self.current_dept.addItem(name, did)

    def _refresh_dept_filter_archived(self):
        self.archived_dept.clear()
        self.archived_dept.addItem("جميع الأقسام", None)
        for did, name in self.db.fetch_all(
                "SELECT id, name FROM departments ORDER BY name"):
            self.archived_dept.addItem(name, did)

    def _create_summary_group(self, lbl_count, lbl_basic, lbl_net):
        group  = QGroupBox("الملخص المالي")
        layout = QHBoxLayout()
        for lbl in (lbl_count, lbl_basic, lbl_net):
            lbl.setStyleSheet(
                "font-weight:bold; font-size:13px; padding:8px; "
                "background:#e8f5e9; border-radius:4px;")
            layout.addWidget(lbl)
        group.setLayout(layout)
        return group

    def _apply_permissions(self):
        role  = self.user['role']
        can_p = can_process_payroll(role)
        can_e = can_edit(role)
        self.btn_calculate.setVisible(can_p)
        self.btn_approve_single.setVisible(can_p)
        self.btn_approve_all.setVisible(can_p)
        self.btn_unapprove_single.setVisible(can_p)
        self.btn_unapprove_all.setVisible(can_p)
        self.btn_edit_installments.setVisible(can_e)
        self.btn_add_bonus.setVisible(can_e)
        self.btn_maintenance.setVisible(can_p)

    def _on_tab_changed(self, index: int):
        if index == 0:
            self._load_current()
        else:
            self._load_archived()

    # ==================== تحميل البيانات ====================
    def _build_payroll_query(self) -> str:
        return """
            SELECT p.id,
                   e.first_name || ' ' || e.last_name,
                   p.basic_salary,
                   p.housing_allowance + p.transportation_allowance +
                   p.food_allowance + p.phone_allowance + p.other_allowances,
                   p.overtime_amount,
                   p.bonus,
                   p.housing_allowance + p.transportation_allowance +
                   p.food_allowance + p.phone_allowance + p.other_allowances +
                   p.overtime_amount + p.bonus,
                   p.absence_deduction,
                   COALESCE(p.unpaid_leave_deduction, 0),
                   p.late_deduction,
                   p.absence_deduction + COALESCE(p.unpaid_leave_deduction,0) +
                   p.late_deduction,
                   p.total_earnings,
                   p.loan_deduction_bank,
                   p.loan_deduction_cash,
                   p.net_salary,
                   p.bank_salary,
                   p.cash_salary,
                   p.notes,
                   p.status
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
        """

    def _rows_to_display(self, data: list) -> list:
        result = []
        for seq, row in enumerate(data, start=1):
            result.append([
                row[0],             # 0: id (مخفي)
                seq,                # 1: رقم تسلسلي
                row[1],             # 2: اسم الموظف
                f"{row[2]:,.2f}",   # 3: الراتب الأساسي
                f"{row[3]:,.2f}",   # 4: البدلات
                f"{row[4]:,.2f}",   # 5: أوفرتايم
                f"{row[5]:,.2f}",   # 6: مكافآت
                f"{row[6]:,.2f}",   # 7: إجمالي الإضافات
                f"{row[7]:,.2f}",   # 8: غياب
                f"{row[8]:,.2f}",   # 9: إجازات بدون راتب
                f"{row[9]:,.2f}",   # 10: تأخير
                f"{row[10]:,.2f}",  # 11: إجمالي الخصومات
                f"{row[11]:,.2f}",  # 12: إجمالي الاستحقاق
                f"{row[12]:,.2f}",  # 13: سلفة بنكية
                f"{row[13]:,.2f}",  # 14: سلفة نقدية
                f"{row[14]:,.2f}",  # 15: صافي الراتب للدفع
                f"{row[15]:,.2f}",  # 16: راتب بنكي
                f"{row[16]:,.2f}",  # 17: راتب نقدي
                row[17] or "",      # 18: ملاحظات
                row[18]             # 19: الحالة
            ])
        return result

    def _load_current(self):
        m   = self.current_month.currentIndex() + 1
        y   = self.current_year.value()
        did = self.current_dept.currentData()
        pay = self.current_payment.currentText()
        sts = self.current_status.currentText()

        q = self._build_payroll_query()
        q += " WHERE p.month=? AND p.year=?"
        params = [m, y]

        if sts == "مسودة":
            q += " AND p.status='مسودة'"
        if did:
            q += " AND e.department_id=?"
            params.append(did)
        if pay == "بنكي":
            q += " AND p.bank_salary > 0"
        elif pay == "نقدي":
            q += " AND p.cash_salary > 0"
        q += " ORDER BY e.first_name"

        data = self.db.fetch_all(q, params)
        self._current_ids = [r[0] for r in data]
        fill_table(self.current_table, self._rows_to_display(data),
                   colors={19: lambda v: "#F57C00" if v == "مسودة" else "#388E3C"})
        self._update_summary(data, self._lbl_cur_count,
                             self._lbl_cur_basic, self._lbl_cur_net)

    def _load_archived(self):
        m   = self.archived_month.currentIndex() + 1
        y   = self.archived_year.value()
        did = self.archived_dept.currentData()
        pay = self.archived_payment.currentText()

        q = self._build_payroll_query()
        q += " WHERE p.month=? AND p.year=? AND p.status='معتمد'"
        params = [m, y]

        if did:
            q += " AND e.department_id=?"
            params.append(did)
        if pay == "بنكي":
            q += " AND p.bank_salary > 0"
        elif pay == "نقدي":
            q += " AND p.cash_salary > 0"
        q += " ORDER BY e.first_name"

        data = self.db.fetch_all(q, params)
        self._archived_ids = [r[0] for r in data]
        fill_table(self.archived_table, self._rows_to_display(data),
                   colors={19: lambda v: "#388E3C"})
        self._update_summary(data, self._lbl_arc_count,
                             self._lbl_arc_basic, self._lbl_arc_net)

    def _update_summary(self, data, lbl_count, lbl_basic, lbl_net):
        currency  = self.db.get_setting('currency', 'ريال')
        n         = len(data)
        tot_basic = sum(r[2]  for r in data if r[2])
        tot_net   = sum(r[14] for r in data if r[14])
        lbl_count.setText(f"عدد الموظفين: {n}")
        lbl_basic.setText(f"إجمالي الأساسي: {tot_basic:,.0f} {currency}")
        lbl_net.setText(f"إجمالي الصافي: {tot_net:,.0f} {currency}")

    # ==================== حساب الرواتب ====================
    def _calculate(self):
        m = self.current_month.currentIndex() + 1
        y = self.current_year.value()

        # ============================================================
        # Guard: التحقق من وجود حضور معتمد للشهر المحدد
        # ============================================================
        approved_att = self.db.fetch_one(
            """SELECT COUNT(*)
               FROM attendance
               WHERE is_approved = 1
                 AND strftime('%m', punch_date) = ?
                 AND strftime('%Y', punch_date) = ?""",
            (f"{m:02d}", str(y)))
        approved_att_count = approved_att[0] if approved_att else 0

        if approved_att_count == 0:
            QMessageBox.warning(
                self, "لا يمكن الحساب",
                f"لا توجد سجلات حضور معتمدة لشهر "
                f"{self.current_month.currentText()} {y}.\n\n"
                "يجب اعتماد سجلات الحضور أولاً من تبويب الحضور "
                "← مسودة الحضور ← زر 'اعتماد الفترة'.")
            return
        # ============================================================

        # التعامل مع الرواتب الموجودة
        approved_count = (self.db.fetch_one(
            "SELECT COUNT(*) FROM payroll "
            "WHERE month=? AND year=? AND status='معتمد'",
            (m, y)) or [0])[0]

        if approved_count > 0:
            if QMessageBox.question(
                    self, "تأكيد",
                    f"يوجد {approved_count} راتب معتمد لهذا الشهر.\n"
                    "سيتم حذف المسودات فقط والإبقاء على المعتمدة.\nمتابعة؟",
                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                return
            self.db.execute_query("""
                UPDATE installments SET payroll_id=NULL, paid_amount=0, notes=NULL
                WHERE payroll_id IN (
                    SELECT id FROM payroll WHERE month=? AND year=? AND status='مسودة'
                )""", (m, y))
            self.db.execute_query(
                "DELETE FROM payroll WHERE month=? AND year=? AND status='مسودة'",
                (m, y))
        else:
            existing = (self.db.fetch_one(
                "SELECT COUNT(*) FROM payroll WHERE month=? AND year=?",
                (m, y)) or [0])[0]
            if existing > 0:
                if QMessageBox.question(
                        self, "تأكيد",
                        "يوجد رواتب لهذا الشهر. إعادة الحساب ستحذفها.\nمتابعة؟",
                        QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                    return
                self.db.execute_query("""
                    UPDATE installments SET payroll_id=NULL, paid_amount=0, notes=NULL
                    WHERE payroll_id IN (
                        SELECT id FROM payroll WHERE month=? AND year=?
                    )""", (m, y))
                self.db.execute_query(
                    "DELETE FROM payroll WHERE month=? AND year=?", (m, y))

        employees = self.db.fetch_all("""
            SELECT id,
                   basic_salary,
                   housing_allowance, transportation_allowance,
                   food_allowance, phone_allowance, other_allowances,
                   COALESCE(bank_salary, 0),
                   COALESCE(cash_salary, 0),
                   is_exempt_from_fingerprint,
                   social_security_registered,
                   social_security_percent
            FROM employees WHERE status='نشط'
        """)

        if not employees:
            QMessageBox.warning(self, "تنبيه", "لا يوجد موظفون نشطون")
            return

        # جلب بيانات الحضور المعتمد فقط
        attendance_data   = self._get_attendance_data(m, y)
        unpaid_leave_data = self._get_unpaid_leave_data(m, y)
        installments_data = self._get_installments_data(m, y)

        progress = QProgressDialog(
            "جاري حساب الرواتب...", "إلغاء", 0, len(employees), self)
        progress.setWindowModality(Qt.WindowModal)

        ok = 0
        for idx, emp in enumerate(employees):
            progress.setValue(idx)
            if progress.wasCanceled():
                break

            (emp_id, basic, housing, transport, food, phone, other,
             bank_salary_orig, cash_salary_orig,
             is_exempt, ss_registered, ss_percent) = emp

            basic            = float(basic   or 0)
            housing          = float(housing or 0)
            transport        = float(transport or 0)
            food             = float(food    or 0)
            phone            = float(phone   or 0)
            other            = float(other   or 0)
            bank_salary_orig = float(bank_salary_orig or 0)

            wdm         = self.work_days_month  if self.work_days_month  > 0 else 26
            wdh         = self.work_hours_daily if self.work_hours_daily > 0 else 8
            daily_rate  = basic / wdm
            hourly_rate = basic / (wdm * wdh)

            att = attendance_data.get(emp_id, {
                'ot_hours': 0.0, 'absent_days': 0,
                'late_minutes': 0, 'early_minutes': 0})

            if is_exempt:
                att = {'ot_hours': 0.0, 'absent_days': 0,
                       'late_minutes': 0, 'early_minutes': 0}

            ot_hours     = float(att['ot_hours']   or 0)
            absent_hours = float(att['absent_days'] or 0) * wdh
            late_hours   = (float(att['late_minutes']  or 0) +
                            float(att['early_minutes'] or 0)) / 60.0

            net_ot, net_absence, net_late = self._calc_netting(
                ot_hours, absent_hours, late_hours)

            ot_amount         = net_ot      * hourly_rate * self.ot_rate
            absence_deduction = net_absence * hourly_rate * self.absence_rate
            late_deduction    = net_late    * hourly_rate * self.absence_rate

            unpaid = unpaid_leave_data.get(emp_id, {'days': 0.0, 'hours': 0.0})
            unpaid_deduction = (float(unpaid['days'])  * daily_rate +
                                float(unpaid['hours']) * hourly_rate)

            bonus         = 0.0
            allowances    = housing + transport + food + phone + other
            total_earnings = (basic + allowances + ot_amount + bonus
                              - absence_deduction - unpaid_deduction - late_deduction)

            inst_list = installments_data.get(emp_id, [])
            loan_bank = sum(i['due'] for i in inst_list if i['method'] == 'bank')
            loan_cash = sum(i['due'] for i in inst_list if i['method'] == 'cash')

            # الخصومات الكلية (بدون الأقساط — تُعرَض منفصلة)
            total_deductions = absence_deduction + unpaid_deduction + late_deduction

            net_payable = total_earnings - loan_bank - loan_cash

            bank_paid, cash_paid = self._calc_bank_cash(
                ss_registered, net_payable, bank_salary_orig, loan_bank)
            bank_paid = max(0.0, bank_paid)
            cash_paid = max(0.0, cash_paid)

            # تقريب الراتب النقدي لأقرب 100
            cash_rounding_note = ""
            if cash_paid > 0:
                cash_paid_orig = cash_paid
                cash_paid      = float(round(cash_paid / 100) * 100)
                if abs(cash_paid - cash_paid_orig) >= 0.5:
                    cash_rounding_note = (
                        f" | الراتب النقدي دُوِّر من {cash_paid_orig:,.0f}"
                        f" إلى {cash_paid:,.0f} (أقرب 100)")

            month_name = self.current_month.currentText()
            notes = f"تم حساب راتب شهر {month_name} {y}."
            if loan_bank + loan_cash > 0:
                notes += " تضمين أقساط شهرية."
            notes += cash_rounding_note

            self.db.execute_query("""
                INSERT INTO payroll (
                    employee_id, month, year,
                    basic_salary, housing_allowance, transportation_allowance,
                    food_allowance, phone_allowance, other_allowances,
                    overtime_hours, overtime_amount, bonus,
                    total_earnings,
                    absence_deduction, late_deduction, unpaid_leave_deduction,
                    loan_deduction_bank, loan_deduction_cash,
                    total_deductions, net_salary,
                    bank_salary, cash_salary,
                    notes, status, approved_at
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                emp_id, m, y,
                basic, housing, transport, food, phone, other,
                ot_hours, ot_amount, bonus,
                total_earnings,
                absence_deduction, late_deduction, unpaid_deduction,
                loan_bank, loan_cash,
                total_deductions, net_payable,
                bank_paid, cash_paid,
                notes, 'مسودة', None
            ))

            payroll_id = self.db.last_id()

            for inst in inst_list:
                self.db.execute_query("""
                    UPDATE installments
                    SET payroll_id = ?,
                        notes = COALESCE(notes, '') ||
                                ' | مدرج في راتب شهر ' || ? || '/' || ?
                    WHERE id = ?
                """, (payroll_id, str(m), str(y), inst['inst_id']))

            ok += 1

        progress.setValue(len(employees))

        self.db.log_action("حساب رواتب", "payroll", None, None,
                           {"month": m, "year": y, "count": ok})
        if self.comm:
            self.comm.dataChanged.emit(
                'payroll', {'action': 'calculate', 'month': m, 'year': y})

        QMessageBox.information(
            self, "نجاح",
            f"تم حساب رواتب {ok} موظف بنجاح.\n"
            f"(استُخدمت {approved_att_count} سجل حضور معتمد)")
        self._load_current()

    # ==================== منطق المقاصة ====================
    @staticmethod
    def _calc_netting(ot_hours: float, absent_hours: float,
                      late_hours: float) -> tuple:
        """
        مقاصة بين الإضافي والغياب والتأخير.
        يُرجع (net_ot_hours, net_absence_hours, net_late_hours).
        """
        total_deficit = absent_hours + late_hours

        if ot_hours <= 0:
            return 0.0, absent_hours, late_hours
        if total_deficit <= 0:
            return ot_hours, 0.0, 0.0
        if abs(ot_hours - total_deficit) < 0.001:
            return 0.0, 0.0, 0.0
        if ot_hours > total_deficit:
            return ot_hours - total_deficit, 0.0, 0.0

        remaining = total_deficit - ot_hours
        if total_deficit > 0:
            abs_ratio  = absent_hours / total_deficit
            late_ratio = late_hours   / total_deficit
            return 0.0, remaining * abs_ratio, remaining * late_ratio
        return 0.0, 0.0, 0.0

    # ==================== توزيع الراتب ====================
    @staticmethod
    def _calc_bank_cash(ss_registered, net_payable: float,
                        bank_salary_orig: float, loan_bank: float) -> tuple:
        if not ss_registered:
            return 0.0, net_payable
        available_bank = bank_salary_orig - loan_bank
        if available_bank <= 0:
            return 0.0, net_payable
        if net_payable <= available_bank:
            return net_payable, 0.0
        return available_bank, net_payable - available_bank

    # ==================== جلب بيانات الحضور ====================
    def _get_attendance_data(self, month: int, year: int) -> dict:
        """
        جلب ملخص الحضور من السجلات المعتمدة فقط (is_approved=1).

        الإصلاح: الشرط AND is_approved=1 يضمن أن حساب الرواتب
        يعتمد فقط على البيانات التي مرّت بمرحلة الاعتماد.
        """
        rows = self.db.fetch_all("""
            SELECT employee_id,
                   SUM(overtime_hours)                             AS ot_hours,
                   SUM(CASE WHEN status='غائب' THEN 1 ELSE 0 END) AS absent_days,
                   SUM(CASE WHEN status NOT IN ('غائب','إجازة')
                            THEN late_minutes ELSE 0 END)          AS late_minutes,
                   SUM(CASE WHEN status NOT IN ('غائب','إجازة')
                            THEN early_leave_minutes ELSE 0 END)   AS early_minutes
            FROM attendance
            WHERE is_approved = 1
              AND strftime('%m', punch_date) = ?
              AND strftime('%Y', punch_date) = ?
            GROUP BY employee_id
        """, (f"{month:02d}", str(year)))

        data = {}
        for emp_id, ot, absent, late, early in rows:
            data[emp_id] = {
                'ot_hours':      float(ot     or 0),
                'absent_days':   int(absent   or 0),
                'late_minutes':  int(late     or 0),
                'early_minutes': int(early    or 0),
            }
        return data

    def _get_unpaid_leave_data(self, month: int, year: int) -> dict:
        start = date(year, month, 1)
        end   = (date(year, month, 1) + timedelta(days=32)).replace(day=1) \
                - timedelta(days=1)

        rows = self.db.fetch_all("""
            SELECT lr.employee_id,
                   SUM(
                       julianday(MIN(lr.end_date, ?)) -
                       julianday(MAX(lr.start_date, ?)) + 1
                   ) AS days,
                   0 AS hours
            FROM leave_requests lr
            JOIN leave_types lt ON lr.leave_type_id = lt.id
            WHERE lt.paid   = 0
              AND lr.status = 'موافق'
              AND lr.start_date <= ?
              AND lr.end_date   >= ?
            GROUP BY lr.employee_id
        """, (end.isoformat(), start.isoformat(),
              end.isoformat(), start.isoformat()))

        data = {}
        for emp_id, days, hours in rows:
            if days and days > 0:
                data[emp_id] = {'days': days, 'hours': hours or 0}
        return data

    def _get_installments_data(self, month: int, year: int) -> dict:
        rows = self.db.fetch_all("""
            SELECT i.id, i.loan_id,
                   (i.amount - COALESCE(i.paid_amount, 0)) AS due,
                   l.employee_id, l.payment_method
            FROM installments i
            JOIN loans l ON i.loan_id = l.id
            WHERE strftime('%Y-%m', i.due_date) = ?
              AND i.status     != 'paid'
              AND i.payroll_id IS NULL
        """, (f"{year}-{month:02d}",))

        data = {}
        for inst_id, loan_id, due, emp_id, method in rows:
            if not due or due <= 0:
                continue
            data.setdefault(emp_id, []).append({
                'inst_id': inst_id,
                'loan_id': loan_id,
                'due':     due,
                'method':  method or 'cash'
            })
        return data

    # ==================== الاعتماد ====================
    def _approve_selected(self):
        rows = set(item.row() for item in self.current_table.selectedItems())
        if not rows:
            QMessageBox.warning(self, "تنبيه", "يرجى تحديد صفوف للاعتماد")
            return
        ids = [self._current_ids[r] for r in rows
               if r < len(self._current_ids)]
        if ids:
            self._approve_payrolls(
                ids, self.current_month.currentIndex() + 1,
                self.current_year.value())

    def _approve_all_current(self):
        ids = self._current_ids[:]
        if not ids:
            QMessageBox.warning(self, "تنبيه", "لا توجد رواتب غير معتمدة")
            return
        m = self.current_month.currentIndex() + 1
        y = self.current_year.value()
        if QMessageBox.question(
                self, "تأكيد",
                f"اعتماد {len(ids)} راتب لشهر "
                f"{self.current_month.currentText()} {y}؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        self._approve_payrolls(ids, m, y)

    def _approve_payrolls(self, payroll_ids: list, month: int, year: int):
        """
        اعتماد الرواتب وتسوية الأقساط — في معاملة واحدة ذرية.
        الإصلاح: يستخدم with self.db.transaction() بدلاً من BEGIN يدوي.
        """
        try:
            with self.db.transaction():
                cur = self.db.conn.cursor()

                for pid in payroll_ids:
                    cur.execute("""
                        UPDATE payroll
                        SET status='معتمد', approved_at=?
                        WHERE id=?
                    """, (datetime.now().isoformat(), pid))

                    installs = cur.execute("""
                        SELECT i.id, i.loan_id, i.amount, i.paid_amount
                        FROM installments i
                        WHERE i.payroll_id = ?
                    """, (pid,)).fetchall()

                    for inst_id, loan_id, amount, paid_amount in installs:
                        pay_amount = paid_amount or amount
                        if not pay_amount or pay_amount <= 0:
                            continue
                        is_fully_paid = (pay_amount >= amount)
                        cur.execute("""
                            UPDATE installments
                            SET status    = ?,
                                paid_date = ?,
                                notes     = COALESCE(notes,'') ||
                                            ' | مسدد عبر راتب ' || ? || '/' || ?
                            WHERE id = ?
                        """, ('paid' if is_fully_paid else 'partial',
                              date.today().isoformat(),
                              str(month), str(year), inst_id))

                        cur.execute("""
                            UPDATE loans
                            SET remaining_amount = MAX(0, remaining_amount - ?)
                            WHERE id = ?
                        """, (pay_amount, loan_id))

                        new_rem = cur.execute(
                            "SELECT remaining_amount FROM loans WHERE id=?",
                            (loan_id,)).fetchone()
                        new_status = 'مكتمل' if (new_rem and new_rem[0] <= 0) else 'نشط'
                        cur.execute(
                            "UPDATE loans SET status=? WHERE id=?",
                            (new_status, loan_id))

            self.db.log_action("اعتماد رواتب", "payroll", None, None,
                               {"month": month, "year": year,
                                "count": len(payroll_ids)})
            if self.comm:
                self.comm.dataChanged.emit(
                    'payroll', {'action': 'approve',
                                'month': month, 'year': year})
                self.comm.dataChanged.emit(
                    'installment', {'action': 'approve',
                                    'month': month, 'year': year})

            QMessageBox.information(
                self, "نجاح", f"تم اعتماد {len(payroll_ids)} راتب بنجاح")
            self._load_current()
            self._load_archived()

        except Exception as e:
            logger.error("خطأ في اعتماد الرواتب: %s", e, exc_info=True)
            QMessageBox.critical(
                self, "خطأ", f"حدث خطأ أثناء الاعتماد:\n{e}")

    # ==================== إلغاء الاعتماد ====================
    def _unapprove_selected(self):
        rows = set(item.row() for item in self.archived_table.selectedItems())
        if not rows:
            QMessageBox.warning(self, "تنبيه", "يرجى تحديد صفوف")
            return
        ids = [self._archived_ids[r] for r in rows
               if r < len(self._archived_ids)]
        if not ids:
            return
        if QMessageBox.question(
                self, "تأكيد إلغاء الاعتماد",
                f"إلغاء اعتماد {len(ids)} راتب؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        for pid in ids:
            self._unapprove_single(pid)
        self._load_current()
        self._load_archived()

    def _unapprove_all_archived(self):
        ids = self._archived_ids[:]
        if not ids:
            QMessageBox.warning(self, "تنبيه", "لا توجد رواتب معتمدة")
            return
        if QMessageBox.question(
                self, "تأكيد",
                f"إلغاء اعتماد جميع رواتب الشهر ({len(ids)} راتب)؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        for pid in ids:
            self._unapprove_single(pid)
        self._load_current()
        self._load_archived()

    def _unapprove_single(self, payroll_id: int):
        """
        إلغاء اعتماد راتب واحد وعكس تسوية الأقساط.
        الإصلاح: يستخدم with self.db.transaction().
        """
        try:
            with self.db.transaction():
                cur = self.db.conn.cursor()

                installs = cur.execute("""
                    SELECT id, loan_id, amount, paid_amount
                    FROM installments WHERE payroll_id=?
                """, (payroll_id,)).fetchall()

                for inst_id, loan_id, amount, paid_amount in installs:
                    reversed_amount = paid_amount or amount or 0
                    cur.execute("""
                        UPDATE installments
                        SET paid_amount=0, status='pending',
                            paid_date=NULL, payroll_id=NULL, notes=NULL
                        WHERE id=?
                    """, (inst_id,))
                    if reversed_amount > 0:
                        cur.execute("""
                            UPDATE loans
                            SET remaining_amount = remaining_amount + ?,
                                status = 'نشط'
                            WHERE id=?
                        """, (reversed_amount, loan_id))

                cur.execute("""
                    UPDATE payroll SET status='مسودة', approved_at=NULL
                    WHERE id=?
                """, (payroll_id,))

            self.db.log_action("إلغاء اعتماد راتب", "payroll",
                               payroll_id, None, None)

        except Exception as e:
            logger.error("خطأ في إلغاء اعتماد الراتب: %s", e, exc_info=True)
            QMessageBox.critical(
                self, "خطأ", f"خطأ أثناء إلغاء الاعتماد:\n{e}")

    # ==================== تعديل الأقساط وإضافة المكافأة ====================
    def _edit_selected(self):
        row = self.current_table.currentRow()
        if row < 0 or row >= len(self._current_ids):
            QMessageBox.warning(self, "تنبيه", "اختر راتباً أولاً")
            return
        pid = self._current_ids[row]
        st  = self.db.fetch_one("SELECT status FROM payroll WHERE id=?", (pid,))
        if not st or st[0] != 'مسودة':
            QMessageBox.warning(self, "تنبيه", "لا يمكن تعديل راتب معتمد.")
            return
        dlg = EditPayrollInstallmentsDialog(self, self.db, self.user, pid, self.comm)
        if dlg.exec_() == QDialog.Accepted:
            self._load_current()
            if self.comm:
                self.comm.dataChanged.emit('installment', {'payroll_id': pid})

    def _add_bonus(self):
        row = self.current_table.currentRow()
        if row < 0 or row >= len(self._current_ids):
            QMessageBox.warning(self, "تنبيه", "اختر راتباً أولاً")
            return
        pid = self._current_ids[row]
        st  = self.db.fetch_one("SELECT status FROM payroll WHERE id=?", (pid,))
        if not st or st[0] != 'مسودة':
            QMessageBox.warning(self, "تنبيه", "لا يمكن تعديل راتب معتمد.")
            return
        dlg = ManageBonusDialog(self, self.db, self.user, pid)
        if dlg.exec_() == QDialog.Accepted:
            self._load_current()
            if self.comm:
                self.comm.dataChanged.emit('payroll', {'id': pid})

    # ==================== طباعة قصاصات الراتب ====================
    def _print_payslip(self):
        m = self.current_month.currentIndex() + 1
        y = self.current_year.value()

        data = self._fetch_payslip_data(m, y, 'معتمد')
        if not data:
            reply = QMessageBox.question(
                self, "تنبيه",
                "لا توجد رواتب معتمدة.\n"
                "هل تريد طباعة قصاصات المسودة؟",
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
            data = self._fetch_payslip_data(m, y, 'مسودة')

        if not data:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات رواتب")
            return

        self._render_payslips(data, m, y)

    def _fetch_payslip_data(self, m: int, y: int, status: str) -> list:
        return self.db.fetch_all("""
            SELECT p.id,
                   p.basic_salary,
                   p.housing_allowance, p.transportation_allowance,
                   p.food_allowance, p.phone_allowance, p.other_allowances,
                   p.overtime_amount, p.bonus,
                   p.total_earnings,
                   p.absence_deduction, p.late_deduction,
                   COALESCE(p.unpaid_leave_deduction, 0),
                   p.loan_deduction_bank, p.loan_deduction_cash,
                   p.total_deductions, p.net_salary,
                   p.bank_salary, p.cash_salary,
                   p.notes, p.status,
                   e.first_name || ' ' || e.last_name,
                   e.employee_code,
                   COALESCE(d.name, ''),
                   COALESCE(e.iban, ''),
                   COALESCE(e.bank_name, '')
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE p.month=? AND p.year=? AND p.status=?
            ORDER BY e.first_name
        """, (m, y, status))

    def _render_payslips(self, data: list, m: int, y: int):
        months_tr   = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
                        'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
        month_name  = months_tr[m - 1]
        company_name    = self.db.get_setting('company_name', 'Şirket')
        company_address = self.db.get_setting('company_address', '')
        company_phone   = self.db.get_setting('company_phone', '')
        logo_path       = self.db.get_setting('company_logo', '')

        logo_html = (f'<img src="{logo_path}" width="120" '
                     f'style="float:left; margin-right:15px;">'
                     if logo_path and os.path.exists(logo_path) else "")

        all_html = ""
        for row in data:
            try:
                (pay_id, basic, housing, transport, food, phone, other,
                 ot, bonus, total_earn, absence, late, unpaid,
                 loan_bank, loan_cash, total_ded, net,
                 bank_sal, cash_sal, notes, status,
                 emp_name, emp_code, dept, iban, bank_name) = row

                all_allow = ((housing or 0) + (transport or 0) + (food or 0)
                             + (phone or 0) + (other or 0))

                html = f"""
                <!DOCTYPE html><html dir="ltr">
                <head><meta charset="UTF-8">
                <style>
                body{{font-family:Arial;margin:0;padding:8px;font-size:9pt;}}
                .header{{display:flex;align-items:center;border-bottom:2px solid #1976D2;
                          padding-bottom:6px;margin-bottom:8px;}}
                .co-info{{flex:1;text-align:right;}}
                .co-info h2{{margin:0;font-size:13pt;color:#1976D2;}}
                .co-info p{{margin:2px 0;font-size:8pt;color:#555;}}
                .title{{text-align:center;font-size:12pt;font-weight:bold;color:#1976D2;
                         margin:6px 0;border-bottom:1px dashed #1976D2;padding-bottom:4px;}}
                .grid{{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin:6px 0;}}
                .section{{background:#f5f9ff;border:1px solid #cce0ff;
                           border-radius:4px;padding:6px;}}
                .section h4{{margin:0 0 4px 0;color:#1976D2;font-size:9pt;
                              border-bottom:1px solid #cce0ff;padding-bottom:2px;}}
                .row{{display:flex;justify-content:space-between;padding:2px 0;font-size:8.5pt;}}
                .row .lbl{{color:#555;}}.row .val{{font-weight:bold;}}
                .total-row{{display:flex;justify-content:space-between;
                             background:#1976D2;color:white;padding:4px 6px;
                             border-radius:3px;margin-top:4px;font-weight:bold;}}
                .net-box{{background:#e8f5e9;border:1px solid #66bb6a;
                           border-radius:4px;padding:6px;margin:6px 0;text-align:center;}}
                .net-box .net-val{{font-size:14pt;font-weight:bold;color:#2e7d32;}}
                .sig{{display:flex;justify-content:space-around;margin-top:16px;font-size:8pt;}}
                .sig-box{{text-align:center;width:140px;}}
                .sig-line{{border-top:1px solid #333;margin-top:25px;padding-top:3px;}}
                .footer{{margin-top:8px;font-size:7pt;color:#999;text-align:center;
                          border-top:1px solid #eee;padding-top:4px;}}
                </style></head><body>
                <div class="header">{logo_html}
                    <div class="co-info">
                        <h2>{company_name}</h2>
                        <p>{company_address}</p><p>Tel: {company_phone}</p>
                    </div>
                </div>
                <div class="title">MAAŞ BORDROSU — {month_name} {y}</div>
                <div class="section" style="margin-bottom:6px;">
                    <div class="row"><span class="lbl">Personel:</span>
                        <span class="val">{emp_name} ({emp_code})</span></div>
                    <div class="row"><span class="lbl">Departman:</span>
                        <span class="val">{dept}</span></div>
                    <div class="row"><span class="lbl">Banka / IBAN:</span>
                        <span class="val">{bank_name} — {iban}</span></div>
                </div>
                <div class="grid">
                    <div class="section"><h4>Kazançlar</h4>
                        <div class="row"><span class="lbl">Temel Maaş</span>
                            <span class="val">{basic or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Yardımlar</span>
                            <span class="val">{all_allow:,.2f}</span></div>
                        <div class="row"><span class="lbl">Fazla Mesai</span>
                            <span class="val">{ot or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Prim</span>
                            <span class="val">{bonus or 0:,.2f}</span></div>
                        <div class="total-row"><span>Toplam Kazanç</span>
                            <span>{total_earn or 0:,.2f}</span></div>
                    </div>
                    <div class="section"><h4>Kesintiler</h4>
                        <div class="row"><span class="lbl">Devamsızlık</span>
                            <span class="val">{absence or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Geç Kalma / Erken Çıkış</span>
                            <span class="val">{late or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Ücretsiz İzin</span>
                            <span class="val">{unpaid or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Avans (Banka)</span>
                            <span class="val">{loan_bank or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Avans (Nakit)</span>
                            <span class="val">{loan_cash or 0:,.2f}</span></div>
                        <div class="total-row"><span>Toplam Kesinti</span>
                            <span>{total_ded or 0:,.2f}</span></div>
                    </div>
                </div>
                <div class="net-box">
                    <div style="font-size:9pt;color:#555;margin-bottom:2px;">Net Ödeme</div>
                    <div class="net-val">{net or 0:,.2f}</div>
                    <div style="font-size:8pt;color:#388e3c;margin-top:2px;">
                        Banka: {bank_sal or 0:,.2f} &nbsp;|&nbsp; Nakit: {cash_sal or 0:,.2f}
                    </div>
                </div>
                <div class="sig">
                    <div class="sig-box"><div class="sig-line">Personel İmzası</div></div>
                    <div class="sig-box"><div class="sig-line">Muhasebe İmzası</div></div>
                    <div class="sig-box"><div class="sig-line">Yetkili İmzası</div></div>
                </div>
                <div class="footer">Bu belge İnsan Kaynakları Sistemi tarafından oluşturulmuştur —
                    {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
                </body></html>"""
                all_html += html + "<div style='page-break-after:always;'></div>"
            except Exception as e:
                logger.error("خطأ في إنشاء قصاصة: %s", e)

        if not all_html:
            QMessageBox.warning(self, "خطأ", "لم يتم إنشاء أي قصاصة.")
            return

        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        printer.setOrientation(QPrinter.Portrait)
        printer.setPageMargins(5, 5, 5, 5, QPrinter.Millimeter)
        dlg = QPrintDialog(printer, self)
        if dlg.exec_() == QDialog.Accepted:
            doc = QTextDocument()
            doc.setHtml(all_html)
            doc.print_(printer)

    # ==================== طباعة كشف الرواتب ====================
    def _print_payroll_sheet(self, table, month_combo, year_spin,
                              dept_filter=None, payment_filter=None):
        try:
            m          = month_combo.currentIndex() + 1
            y          = year_spin.value()
            months_ar  = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                           'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
            month_name = months_ar[m - 1]
            company_name    = self.db.get_setting('company_name', 'الشركة')
            company_address = self.db.get_setting('company_address', '')
            company_phone   = self.db.get_setting('company_phone', '')
            logo_path       = self.db.get_setting('company_logo', '')

            logo_html    = (f'<img src="{logo_path}" width="80" style="float:left;">'
                            if logo_path and os.path.exists(logo_path) else "")
            filter_info  = f"الشهر: {month_name} {y}"
            if dept_filter and dept_filter.currentData():
                filter_info += f" | القسم: {dept_filter.currentText()}"
            if payment_filter and payment_filter.currentText() != "الكل":
                filter_info += f" | طريقة الدفع: {payment_filter.currentText()}"

            headers     = [table.horizontalHeaderItem(c).text()
                           for c in range(1, table.columnCount())
                           if not table.isColumnHidden(c)]
            visible_cols = [c for c in range(1, table.columnCount())
                            if not table.isColumnHidden(c)]
            rows_data, totals = [], {}
            for r in range(table.rowCount()):
                vals = []
                for c in visible_cols:
                    item = table.item(r, c)
                    text = item.text() if item else ""
                    vals.append(text)
                    try:
                        totals[c] = totals.get(c, 0.0) + float(
                            text.replace(',', '').replace(' ', ''))
                    except Exception:
                        pass
                rows_data.append(vals)

            total_row = []
            for i, c in enumerate(visible_cols):
                if i == 0:
                    total_row.append("الإجمالي")
                elif c in totals:
                    total_row.append(f"{totals[c]:,.2f}")
                else:
                    total_row.append("")

            html = f"""
            <!DOCTYPE html><html dir="rtl">
            <head><meta charset="UTF-8"><style>
            body{{font-family:Arial;margin:2px;padding:0;}}
            .header{{display:flex;align-items:center;margin-bottom:4px;
                      border-bottom:1px solid #1976D2;}}
            .co-info{{flex:1;text-align:right;}}
            .co-info h2{{margin:0;font-size:11pt;color:#1976D2;}}
            .co-info p{{margin:1px 0;font-size:7pt;color:#555;}}
            .title{{font-size:10pt;font-weight:bold;text-align:center;
                     margin:3px 0;color:#1976D2;}}
            .fi{{background:#f0f0f0;padding:2px 4px;font-size:7pt;margin:2px 0;}}
            table{{border-collapse:collapse;width:100%;font-size:6pt;}}
            th,td{{border:1px solid #aaa;padding:1px 2px;text-align:center;}}
            th{{background:#1976D2;color:white;}}
            .tr{{background:#e0e0e0;font-weight:bold;}}
            .footer{{margin-top:3px;font-size:5pt;color:#888;text-align:center;}}
            </style></head><body>
            <div class="header">{logo_html}
                <div class="co-info"><h2>{company_name}</h2>
                    <p>{company_address} — هاتف: {company_phone}</p></div>
            </div>
            <div class="title">كشف الرواتب</div>
            <div class="fi">{filter_info}</div>
            <table><thead><tr>{"".join(f"<th>{h}</th>" for h in headers)}</tr></thead>
            <tbody>"""

            for row in rows_data:
                html += "<tr>" + "".join(f"<td>{c}</td>" for c in row) + "</tr>"
            html += ('<tr class="tr">' +
                     "".join(f"<td>{c}</td>" for c in total_row) + "</tr>")
            html += f"""</tbody></table>
            <div class="footer">تم الإنشاء: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
            </body></html>"""

            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            printer.setOrientation(QPrinter.Landscape)
            printer.setPageMargins(3, 3, 3, 3, QPrinter.Millimeter)
            dlg = QPrintDialog(printer, self)
            if dlg.exec_() == QDialog.Accepted:
                doc = QTextDocument()
                doc.setHtml(html)
                doc.print_(printer)

        except Exception as e:
            QMessageBox.critical(self, "خطأ في الطباعة", str(e))

    def _print_current_payroll(self):
        if self.current_table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للطباعة")
            return
        self._print_payroll_sheet(
            self.current_table, self.current_month, self.current_year,
            self.current_dept, self.current_payment)

    def _print_archived_payroll(self):
        if self.archived_table.rowCount() == 0:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للطباعة")
            return
        self._print_payroll_sheet(
            self.archived_table, self.archived_month, self.archived_year,
            self.archived_dept, self.archived_payment)

    # ==================== تصدير Excel ====================
    def _export_current_excel(self):
        self._export_excel(
            self.current_month, self.current_year,
            self.current_dept, self.current_payment, is_current=True)

    def _export_archived_excel(self):
        self._export_excel(
            self.archived_month, self.archived_year,
            self.archived_dept, self.archived_payment, is_current=False)

    def _export_excel(self, month_combo, year_spin,
                      dept_filter=None, payment_filter=None,
                      is_current: bool = True):
        try:
            import pandas as pd
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")
            return

        try:
            m      = month_combo.currentIndex() + 1
            y      = year_spin.value()
            did    = dept_filter.currentData() if dept_filter else None
            pay    = payment_filter.currentText() if payment_filter else "الكل"
            status = "مسودة" if is_current else "معتمد"

            q = f"""
                SELECT e.employee_code,
                       e.first_name || ' ' || e.last_name,
                       COALESCE(d.name, '') AS department,
                       p.basic_salary,
                       p.housing_allowance + p.transportation_allowance +
                       p.food_allowance + p.phone_allowance + p.other_allowances,
                       p.overtime_amount, p.bonus,
                       p.housing_allowance + p.transportation_allowance +
                       p.food_allowance + p.phone_allowance + p.other_allowances +
                       p.overtime_amount + p.bonus,
                       p.absence_deduction,
                       COALESCE(p.unpaid_leave_deduction, 0),
                       p.late_deduction,
                       p.absence_deduction + COALESCE(p.unpaid_leave_deduction,0) +
                       p.late_deduction,
                       p.total_earnings,
                       p.loan_deduction_bank, p.loan_deduction_cash,
                       p.net_salary, p.bank_salary, p.cash_salary,
                       p.notes, p.status
                FROM payroll p
                JOIN employees e ON p.employee_id = e.id
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE p.month=? AND p.year=? AND p.status='{status}'
            """
            params = [m, y]
            if did:
                q += " AND e.department_id=?"
                params.append(did)
            if pay == "بنكي":
                q += " AND p.bank_salary > 0"
            elif pay == "نقدي":
                q += " AND p.cash_salary > 0"
            q += " ORDER BY e.first_name"

            data = self.db.fetch_all(q, params)
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return

            cols = ["الرقم الوظيفي", "الاسم", "القسم",
                    "الراتب الأساسي", "البدلات", "أوفرتايم", "مكافآت",
                    "إجمالي الإضافات", "غياب", "إجازات بدون راتب", "تأخير",
                    "إجمالي الخصومات", "إجمالي الاستحقاق",
                    "سلفة بنكية", "سلفة نقدية", "صافي الراتب للدفع",
                    "راتب بنكي", "راتب نقدي", "ملاحظات", "الحالة"]
            df = pd.DataFrame(data, columns=cols)

            num_cols = ["الراتب الأساسي", "البدلات", "أوفرتايم", "مكافآت",
                        "إجمالي الإضافات", "غياب", "إجازات بدون راتب", "تأخير",
                        "إجمالي الخصومات", "إجمالي الاستحقاق",
                        "سلفة بنكية", "سلفة نقدية", "صافي الراتب للدفع",
                        "راتب بنكي", "راتب نقدي"]
            totals  = df[num_cols].sum()
            tot_row = ["الإجمالي", "", ""] + list(totals) + ["", ""]
            df = pd.concat(
                [df, pd.DataFrame([tot_row], columns=cols)],
                ignore_index=True)

            months_ar = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                         'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
            fname    = f"payroll_{months_ar[m-1]}_{y}.xlsx"
            path, _  = QFileDialog.getSaveFileName(
                self, "حفظ", fname, "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", "تم تصدير الملف بنجاح")

        except Exception as e:
            QMessageBox.critical(self, "خطأ في التصدير", str(e))

    # ==================== القائمة المنبثقة + إيصال الدفع ====================
    def _show_context_menu(self, pos):
        tab_idx = self.tabs.currentIndex()
        table   = self.current_table  if tab_idx == 0 else self.archived_table
        ids     = self._current_ids   if tab_idx == 0 else self._archived_ids
        row     = table.currentRow()
        if row < 0 or row >= len(ids):
            return

        pid      = ids[row]
        row_data = self.db.fetch_one("""
            SELECT e.first_name || ' ' || e.last_name,
                   p.net_salary, p.month, p.year
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            WHERE p.id = ?
        """, (pid,))
        if not row_data:
            return

        emp_name, net_salary, month, year = row_data
        menu   = QMenu()
        action = QAction("🧾 Ödeme Makbuzu Yazdır", self)
        action.triggered.connect(
            lambda: self._print_receipt_tr(
                pid, emp_name, net_salary, month, year))
        menu.addAction(action)
        menu.exec_(table.viewport().mapToGlobal(pos))

    def _print_receipt_tr(self, payroll_id: int, emp_name: str,
                           amount: float, month: int, year: int):
        try:
            currency   = self.db.get_setting('currency', 'TL')
            months_tr  = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
                           'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
            month_name = months_tr[month - 1] if 1 <= month <= 12 else str(month)
            company_name    = self.db.get_setting('company_name', 'Şirket')
            company_address = self.db.get_setting('company_address', '')
            company_phone   = self.db.get_setting('company_phone', '')
            logo_path       = self.db.get_setting('company_logo', '')

            logo_html = (f'<img src="{logo_path}" width="130" '
                         f'style="float:left; margin-right:15px;">'
                         if logo_path and os.path.exists(logo_path) else "")

            cash_row = self.db.fetch_one(
                "SELECT cash_salary, notes FROM payroll WHERE id=?",
                (payroll_id,))
            cash      = float(cash_row[0] or 0) if cash_row else 0.0
            pay_notes = cash_row[1] or "" if cash_row else ""

            bank_row = self.db.fetch_one(
                "SELECT bank_salary FROM payroll WHERE id=?", (payroll_id,))
            bank = float(bank_row[0] or 0) if bank_row else 0.0

            rounding_note_html = ""
            if "دُوِّر" in pay_notes:
                import re
                m_rnd = re.search(r'الراتب النقدي دُوِّر[^|]+', pay_notes)
                if m_rnd:
                    rounding_note_html = (
                        f'<div style="font-size:8pt;color:#e65100;margin-top:4px;">'
                        f'⚠ {m_rnd.group(0).strip()}</div>')

            words_cash = number_to_words_tr(cash, currency) if cash > 0 else "—"
            words_bank = number_to_words_tr(bank, currency) if bank > 0 else "—"

            html = f"""
            <!DOCTYPE html><html dir="ltr">
            <head><meta charset="UTF-8"><style>
            body{{font-family:Arial;margin:0;padding:14px;font-size:11pt;line-height:1.8;}}
            .header{{display:flex;align-items:center;border-bottom:2px solid #1976D2;
                      padding-bottom:10px;margin-bottom:14px;}}
            .co-info{{flex:1;text-align:right;}}
            .co-info h2{{margin:0;font-size:14pt;color:#1976D2;}}
            .co-info p{{margin:3px 0;font-size:9.5pt;color:#555;}}
            .title{{font-size:15pt;font-weight:bold;text-align:center;color:#1976D2;
                     margin:12px 0;border:2px solid #1976D2;padding:8px;
                     border-radius:4px;letter-spacing:1px;}}
            .box{{background:#f5f9ff;border:1px solid #cce0ff;
                   border-radius:6px;padding:14px 18px;margin:10px 0;}}
            .row{{display:flex;justify-content:space-between;padding:7px 0;
                   border-bottom:1px dotted #ddd;font-size:11pt;}}
            .row:last-child{{border-bottom:none;}}
            .lbl{{color:#555;}}.val{{font-weight:bold;color:#1a237e;}}
            .words{{font-size:9.5pt;color:#388e3c;margin:3px 0 6px 0;
                     padding-right:4px;font-style:italic;}}
            .sig-area{{display:flex;justify-content:space-between;margin-top:40px;}}
            .sig-box{{text-align:center;width:43%;}}
            .sig-line{{border-top:1px solid #333;margin-top:36px;
                        padding-top:6px;font-size:10pt;}}
            .footer{{margin-top:20px;font-size:7.5pt;color:#aaa;text-align:center;
                      border-top:1px solid #eee;padding-top:6px;}}
            </style></head><body>
            <div class="header">{logo_html}
                <div class="co-info"><h2>{company_name}</h2>
                    <p>{company_address}</p><p>Tel: {company_phone}</p></div>
            </div>
            <div class="title">ÖDEME MAKBUZU</div>
            <div class="box">
                <div class="row"><span class="lbl">Tarih:</span>
                    <span class="val">{datetime.now().strftime('%d/%m/%Y')}</span></div>
                <div class="row"><span class="lbl">Alıcı:</span>
                    <span class="val">{emp_name}</span></div>
                <div class="row"><span class="lbl">Dönem:</span>
                    <span class="val">{month_name} {year}</span></div>
                <div class="row"><span class="lbl">Net Ödeme:</span>
                    <span class="val">{amount or 0:,.2f} {currency}</span></div>
                <div class="row"><span class="lbl">Nakit Ödeme:</span>
                    <span class="val">{cash:,.2f} {currency}</span></div>
                {rounding_note_html}
                <div class="words">{words_cash}</div>
                <div class="row"><span class="lbl">Banka Ödemesi:</span>
                    <span class="val">{bank:,.2f} {currency}</span></div>
                <div class="words">{words_bank}</div>
            </div>
            <div class="sig-area">
                <div class="sig-box">
                    <div class="sig-line">Alıcı Adı Soyadı / İmzası</div></div>
                <div class="sig-box">
                    <div class="sig-line">Yetkili Adı Soyadı / İmzası</div></div>
            </div>
            <div class="footer">Bu belge İnsan Kaynakları Sistemi tarafından oluşturulmuştur —
                {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
            </body></html>"""

            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            printer.setOrientation(QPrinter.Portrait)
            printer.setPageMargins(8, 8, 8, 8, QPrinter.Millimeter)
            dlg = QPrintDialog(printer, self)
            if dlg.exec_() == QDialog.Accepted:
                doc = QTextDocument()
                doc.setHtml(html)
                doc.print_(printer)

            self.db.execute_query("""
                UPDATE payroll
                SET notes = COALESCE(notes,'') || ' | طباعة إيصال: ' || ?
                WHERE id = ?
            """, (datetime.now().strftime('%Y-%m-%d %H:%M'), payroll_id))
            self._load_current()
            self._load_archived()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Yazdırma hatası: {e}")

    # ==================== طباعة جميع الإيصالات ====================
    def _print_all_receipts(self):
        tab_idx = self.tabs.currentIndex()
        if tab_idx == 0:
            ids, month_combo, year_spin = (
                self._current_ids[:], self.current_month, self.current_year)
        else:
            ids, month_combo, year_spin = (
                self._archived_ids[:], self.archived_month, self.archived_year)

        if not ids:
            QMessageBox.warning(self, "تنبيه", "لا توجد رواتب لطباعة إيصالاتها")
            return
        if QMessageBox.question(
                self, "تأكيد",
                f"سيتم طباعة {len(ids)} إيصال دفع. متابعة؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return

        try:
            currency     = self.db.get_setting('currency', 'TL')
            company_name = self.db.get_setting('company_name', 'Şirket')
            company_addr = self.db.get_setting('company_address', '')
            company_tel  = self.db.get_setting('company_phone', '')
            logo_path    = self.db.get_setting('company_logo', '')
            months_tr    = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
                             'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
            m            = month_combo.currentIndex() + 1
            y            = year_spin.value()
            month_name   = months_tr[m - 1]
            logo_html    = (f'<img src="{logo_path}" width="130" '
                            f'style="float:left; margin-right:15px;">'
                            if logo_path and os.path.exists(logo_path) else "")

            import re
            all_html = ""
            for pid in ids:
                row = self.db.fetch_one("""
                    SELECT e.first_name || ' ' || e.last_name,
                           p.net_salary, p.cash_salary,
                           p.bank_salary, p.notes
                    FROM payroll p
                    JOIN employees e ON p.employee_id = e.id
                    WHERE p.id = ?
                """, (pid,))
                if not row:
                    continue

                emp_name, net, cash, bank, pay_notes = row
                net       = float(net  or 0)
                cash      = float(cash or 0)
                bank      = float(bank or 0)
                pay_notes = pay_notes or ""

                words_cash = number_to_words_tr(cash, currency) if cash > 0 else "—"
                words_bank = number_to_words_tr(bank, currency) if bank > 0 else "—"

                rnd_note = ""
                m_rnd = re.search(r'الراتب النقدي دُوِّر[^|]+', pay_notes)
                if m_rnd:
                    rnd_note = (
                        f'<div style="font-size:8pt;color:#e65100;">'
                        f'⚠ {m_rnd.group(0).strip()}</div>')

                html = f"""
                <!DOCTYPE html><html dir="ltr">
                <head><meta charset="UTF-8"><style>
                body{{font-family:Arial;margin:0;padding:14px;font-size:11pt;}}
                .header{{display:flex;align-items:center;border-bottom:2px solid #1976D2;
                          padding-bottom:10px;margin-bottom:14px;}}
                .co-info{{flex:1;text-align:right;}}
                .co-info h2{{margin:0;font-size:14pt;color:#1976D2;}}
                .title{{font-size:15pt;font-weight:bold;text-align:center;color:#1976D2;
                         margin:12px 0;border:2px solid #1976D2;padding:8px;}}
                .box{{background:#f5f9ff;border:1px solid #cce0ff;
                       border-radius:6px;padding:14px 18px;margin:10px 0;}}
                .row{{display:flex;justify-content:space-between;padding:7px 0;
                       border-bottom:1px dotted #ddd;}}
                .lbl{{color:#555;}}.val{{font-weight:bold;color:#1a237e;}}
                .words{{font-size:9.5pt;color:#388e3c;font-style:italic;}}
                .sig-area{{display:flex;justify-content:space-between;margin-top:40px;}}
                .sig-box{{text-align:center;width:43%;}}
                .sig-line{{border-top:1px solid #333;margin-top:36px;padding-top:6px;}}
                .footer{{margin-top:20px;font-size:7.5pt;color:#aaa;text-align:center;}}
                </style></head><body>
                <div class="header">{logo_html}
                    <div class="co-info"><h2>{company_name}</h2>
                        <p>{company_addr}</p><p>Tel: {company_tel}</p></div>
                </div>
                <div class="title">ÖDEME MAKBUZU</div>
                <div class="box">
                    <div class="row"><span class="lbl">Tarih:</span>
                        <span class="val">{datetime.now().strftime('%d/%m/%Y')}</span></div>
                    <div class="row"><span class="lbl">Alıcı:</span>
                        <span class="val">{emp_name}</span></div>
                    <div class="row"><span class="lbl">Dönem:</span>
                        <span class="val">{month_name} {y}</span></div>
                    <div class="row"><span class="lbl">Net Ödeme:</span>
                        <span class="val">{net:,.2f} {currency}</span></div>
                    <div class="row"><span class="lbl">Nakit Ödeme:</span>
                        <span class="val">{cash:,.2f} {currency}</span></div>
                    {rnd_note}
                    <div class="words">{words_cash}</div>
                    <div class="row"><span class="lbl">Banka Ödemesi:</span>
                        <span class="val">{bank:,.2f} {currency}</span></div>
                    <div class="words">{words_bank}</div>
                </div>
                <div class="sig-area">
                    <div class="sig-box">
                        <div class="sig-line">Alıcı Adı Soyadı / İmzası</div></div>
                    <div class="sig-box">
                        <div class="sig-line">Yetkili Adı Soyadı / İmzası</div></div>
                </div>
                <div class="footer">Bu belge İnsan Kaynakları Sistemi tarafından oluşturulmuştur —
                    {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
                </body></html>"""
                all_html += html + "<div style='page-break-after:always;'></div>"

            if not all_html:
                QMessageBox.warning(self, "خطأ", "لم يتم إنشاء أي إيصال")
                return

            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            printer.setOrientation(QPrinter.Portrait)
            printer.setPageMargins(8, 8, 8, 8, QPrinter.Millimeter)
            dlg = QPrintDialog(printer, self)
            if dlg.exec_() == QDialog.Accepted:
                doc = QTextDocument()
                doc.setHtml(all_html)
                doc.print_(printer)

        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"خطأ في الطباعة الجماعية:\n{e}")

    # ==================== صيانة الرواتب ====================
    def _run_maintenance(self):
        m = self.archived_month.currentIndex() + 1
        y = self.archived_year.value()
        dlg = PayrollMaintenanceDialog(self, self.db, m, y)
        if dlg.exec_() == QDialog.Accepted:
            self._load_current()
            self._load_archived()


# ===================================================================
# نافذة تعديل الأقساط المحسوبة
# ===================================================================
class EditPayrollInstallmentsDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, user: dict,
                 payroll_id: int, comm=None):
        super().__init__(parent)
        self.db         = db
        self.user       = user
        self.payroll_id = payroll_id
        self.comm       = comm
        self.setWindowTitle("تعديل خصم السلف من الراتب")
        self.setMinimumSize(860, 560)
        self.setLayoutDirection(Qt.RightToLeft)
        if self._load_data():
            self._build()

    def _load_data(self) -> bool:
        self.payroll = self.db.fetch_one("""
            SELECT id, employee_id, month, year,
                   basic_salary, total_earnings,
                   absence_deduction, late_deduction,
                   COALESCE(unpaid_leave_deduction, 0),
                   loan_deduction_bank, loan_deduction_cash, net_salary
            FROM payroll WHERE id=?
        """, (self.payroll_id,))

        if not self.payroll:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على الراتب")
            self.reject()
            return False

        self.employee_id = self.payroll[1]
        emp = self.db.fetch_one(
            "SELECT first_name || ' ' || last_name FROM employees WHERE id=?",
            (self.employee_id,))
        self.emp_name = emp[0] if emp else "غير معروف"

        self.installments = self.db.fetch_all("""
            SELECT i.id, i.loan_id, i.due_date,
                   i.amount,
                   COALESCE(i.paid_amount, 0),
                   (i.amount - COALESCE(i.paid_amount, 0)) AS remaining,
                   l.loan_type, l.payment_method,
                   CASE WHEN i.payroll_id = ? THEN 1 ELSE 0 END AS linked
            FROM installments i
            JOIN loans l ON i.loan_id = l.id
            WHERE l.employee_id = ?
              AND l.status      = 'نشط'
              AND i.status     != 'paid'
              AND (i.payroll_id IS NULL OR i.payroll_id = ?)
            ORDER BY i.due_date
        """, (self.payroll_id, self.employee_id, self.payroll_id))
        return True

    def _build(self):
        layout = QVBoxLayout(self)

        info = QLabel(
            f"الموظف: {self.emp_name}  |  "
            f"الراتب الأساسي: {self.payroll[4]:,.2f}  |  "
            f"إجمالي الاستحقاق: {self.payroll[5]:,.2f}")
        info.setStyleSheet("font-weight:bold; padding:8px; background:#e3f2fd;")
        layout.addWidget(info)

        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "تحديد", "نوع السلفة", "تاريخ الاستحقاق", "المبلغ الأصلي",
            "المدفوع", "المتبقي", "طريقة الدفع",
            "المبلغ المراد خصمه", "ملاحظات"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._populate_table()
        layout.addWidget(self.table)

        total_layout = QHBoxLayout()
        total_layout.addWidget(QLabel("إجمالي الخصم المحدد:"))
        self._total_lbl = QLabel("0.00")
        self._total_lbl.setStyleSheet(
            "font-weight:bold; color:#1976D2; font-size:14px;")
        total_layout.addWidget(self._total_lbl)
        total_layout.addStretch()
        layout.addLayout(total_layout)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn("💾 حفظ التغييرات", BTN_SUCCESS, self._save))
        btn_layout.addWidget(btn("❌ إلغاء",          BTN_DANGER,  self.reject))
        layout.addLayout(btn_layout)

        self._update_total()
        self.table.itemChanged.connect(self._on_check_changed)

    def _populate_table(self):
        self.table.setRowCount(len(self.installments))
        for row_idx, inst in enumerate(self.installments):
            (inst_id, loan_id, due_date, amount,
             paid, remaining, loan_type, method, linked) = inst

            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Checked if linked else Qt.Unchecked)
            self.table.setItem(row_idx, 0, chk)

            for col, val in [(1, loan_type or ""), (2, due_date or ""),
                              (3, f"{amount:,.2f}"), (4, f"{paid:,.2f}"),
                              (5, f"{remaining:,.2f}"),
                              (6, "بنكي" if method == 'bank' else "نقدي")]:
                self.table.setItem(row_idx, col, QTableWidgetItem(val))

            spin = QDoubleSpinBox()
            spin.setRange(0.0, max(remaining, 0.0))
            spin.setDecimals(2)
            spin.setValue(remaining if linked else 0.0)
            spin.valueChanged.connect(self._update_total)
            self.table.setCellWidget(row_idx, 7, spin)

            notes_item = QTableWidgetItem("")
            notes_item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.table.setItem(row_idx, 8, notes_item)

    def _on_check_changed(self, item):
        if item.column() != 0:
            return
        row     = item.row()
        checked = item.checkState() == Qt.Checked
        spin    = self.table.cellWidget(row, 7)
        if spin:
            rem_item = self.table.item(row, 5)
            if rem_item:
                try:
                    rem = float(rem_item.text().replace(',', ''))
                    spin.setValue(rem if checked else 0.0)
                except Exception:
                    pass
        self._update_total()

    def _update_total(self):
        total = sum(
            self.table.cellWidget(r, 7).value()
            for r in range(self.table.rowCount())
            if self.table.cellWidget(r, 7))
        self._total_lbl.setText(f"{total:,.2f}")

    def _save(self):
        try:
            with self.db.transaction():
                cur   = self.db.conn.cursor()
                month = self.payroll[2]
                year  = self.payroll[3]

                cur.execute("""
                    UPDATE installments
                    SET payroll_id=NULL, paid_amount=0, notes=NULL
                    WHERE payroll_id=?
                """, (self.payroll_id,))

                loan_bank = loan_cash = 0.0

                for row_idx in range(self.table.rowCount()):
                    chk = self.table.item(row_idx, 0)
                    if not chk or chk.checkState() != Qt.Checked:
                        continue
                    inst_id = self.installments[row_idx][0]
                    method  = self.installments[row_idx][7]
                    spin    = self.table.cellWidget(row_idx, 7)
                    amount  = spin.value() if spin else 0.0
                    if amount <= 0:
                        continue

                    note_item = self.table.item(row_idx, 8)
                    note_txt  = note_item.text().strip() if note_item else ""
                    note_full = (f"مدرج في راتب شهر {month}/{year}"
                                 + (f" — {note_txt}" if note_txt else ""))

                    cur.execute("""
                        UPDATE installments
                        SET payroll_id=?, paid_amount=?, notes=?
                        WHERE id=?
                    """, (self.payroll_id, amount, note_full, inst_id))

                    if method == 'bank':
                        loan_bank += amount
                    else:
                        loan_cash += amount

                total_earnings   = self.payroll[5] or 0
                absence          = self.payroll[6] or 0
                late             = self.payroll[7] or 0
                unpaid           = self.payroll[8] or 0
                total_deductions = absence + late + unpaid + loan_bank + loan_cash
                net_salary       = total_earnings - loan_bank - loan_cash

                cur.execute("""
                    UPDATE payroll
                    SET loan_deduction_bank=?, loan_deduction_cash=?,
                        total_deductions=?, net_salary=?
                    WHERE id=?
                """, (loan_bank, loan_cash, total_deductions,
                      net_salary, self.payroll_id))

            QMessageBox.information(self, "نجاح", "تم تحديث خصم السلف بنجاح")
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ:\n{e}")


# ===================================================================
# نافذة إدارة المكافأة
# ===================================================================
class ManageBonusDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, user: dict, payroll_id: int):
        super().__init__(parent)
        self.db         = db
        self.user       = user
        self.payroll_id = payroll_id
        self._load_current_bonus()
        self.setWindowTitle("إدارة المكافأة")
        self.setFixedSize(400, 240)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()

    def _load_current_bonus(self):
        row = self.db.fetch_one(
            "SELECT bonus FROM payroll WHERE id=?", (self.payroll_id,))
        self.current_bonus = float(row[0] or 0) if row else 0.0

    def _build(self):
        layout   = QVBoxLayout(self)
        currency = self.db.get_setting('currency', 'ريال')

        current_lbl = QLabel(
            f"المكافأة الحالية: {self.current_bonus:,.2f} {currency}")
        current_lbl.setStyleSheet(
            "font-weight:bold; padding:8px; background:#fff9c4; "
            "border-radius:4px; font-size:13px;")
        layout.addWidget(current_lbl)

        form = QFormLayout()
        self.bonus_spin = QDoubleSpinBox()
        self.bonus_spin.setRange(0.0, 9_999_999.99)
        self.bonus_spin.setDecimals(2)
        self.bonus_spin.setSuffix(f" {currency}")
        self.bonus_spin.setValue(self.current_bonus)
        form.addRow("القيمة الجديدة للمكافأة:", self.bonus_spin)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        btn_save   = QPushButton("💾 حفظ")
        btn_save.setStyleSheet(BTN_SUCCESS)
        btn_save.clicked.connect(self._save)

        btn_delete = QPushButton("🗑️ حذف المكافأة")
        btn_delete.setStyleSheet(BTN_DANGER)
        btn_delete.clicked.connect(self._delete)
        btn_delete.setEnabled(self.current_bonus > 0)

        btn_cancel = QPushButton("❌ إلغاء")
        btn_cancel.setStyleSheet(BTN_GRAY)
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_delete)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def _save(self):
        self._apply(self.bonus_spin.value(), "تعديل مكافأة")

    def _delete(self):
        if QMessageBox.question(
                self, "تأكيد", "هل تريد حذف المكافأة وضبطها على صفر؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        self._apply(0.0, "حذف مكافأة")

    def _apply(self, new_bonus: float, action: str):
        row = self.db.fetch_one("""
            SELECT total_earnings, net_salary, bonus
            FROM payroll WHERE id=?
        """, (self.payroll_id,))
        if not row:
            return
        total_earn_old, net_old, bonus_old = row
        bonus_old = float(bonus_old  or 0)
        diff      = new_bonus - bonus_old
        new_total = float(total_earn_old or 0) + diff
        new_net   = float(net_old        or 0) + diff

        self.db.execute_query("""
            UPDATE payroll
            SET bonus=?, total_earnings=?, net_salary=?
            WHERE id=?
        """, (new_bonus, new_total, new_net, self.payroll_id))

        self.db.log_action(action, "payroll", self.payroll_id,
                           {"bonus": bonus_old}, {"bonus": new_bonus})
        self.accept()


# الاسم القديم للتوافق
AddBonusDialog = ManageBonusDialog


# ===================================================================
# نافذة صيانة الرواتب
# ===================================================================
class PayrollMaintenanceDialog(QDialog):
    CHECKS = [
        {
            "id":        "neg_net",
            "title":     "صافي الراتب سالب",
            "desc":      "صافي الراتب للدفع أقل من صفر",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, p.net_salary "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? AND p.net_salary < 0"),
            "fixable":   True,
            "fix_label": "ضبط صافي الراتب على 0",
            "fix_sql":   "UPDATE payroll SET net_salary=0, bank_salary=0, cash_salary=0 WHERE id=?",
        },
        {
            "id":        "bank_cash_mismatch",
            "title":     "مجموع بنكي + نقدي ≠ صافي الراتب",
            "desc":      "فارق > 1 بين net_salary و bank+cash",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, "
                          "p.net_salary, p.bank_salary, p.cash_salary "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND ABS(p.net_salary-(p.bank_salary+p.cash_salary)) > 1"),
            "fixable":   True,
            "fix_label": "تعديل نقدي = صافي − بنكي",
            "fix_sql":   "UPDATE payroll SET cash_salary=MAX(0,net_salary-bank_salary) WHERE id=?",
        },
        {
            "id":        "wrong_total_earn",
            "title":     "إجمالي الاستحقاق خاطئ",
            "desc":      "total_earnings ≠ أساسي+بدلات+OT+مكافأة−خصومات (فارق > 1)",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, p.total_earnings, "
                          "(p.basic_salary+p.housing_allowance+p.transportation_allowance+"
                          "p.food_allowance+p.phone_allowance+p.other_allowances+"
                          "p.overtime_amount+p.bonus-p.absence_deduction-p.late_deduction-"
                          "COALESCE(p.unpaid_leave_deduction,0)) AS exp "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND ABS(p.total_earnings-(p.basic_salary+p.housing_allowance+"
                          "p.transportation_allowance+p.food_allowance+p.phone_allowance+"
                          "p.other_allowances+p.overtime_amount+p.bonus-p.absence_deduction-"
                          "p.late_deduction-COALESCE(p.unpaid_leave_deduction,0)))>1"),
            "fixable":   True,
            "fix_label": "إعادة حساب إجمالي الاستحقاق",
            "fix_sql":   ("UPDATE payroll SET total_earnings="
                          "basic_salary+housing_allowance+transportation_allowance+"
                          "food_allowance+phone_allowance+other_allowances+"
                          "overtime_amount+bonus-absence_deduction-late_deduction-"
                          "COALESCE(unpaid_leave_deduction,0) WHERE id=?"),
        },
        {
            "id":        "wrong_total_ded",
            "title":     "إجمالي الخصومات خاطئ",
            "desc":      "total_deductions ≠ غياب+تأخير+بدون راتب (فارق > 1)",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, "
                          "p.total_deductions, "
                          "(p.absence_deduction+p.late_deduction+"
                          "COALESCE(p.unpaid_leave_deduction,0)) AS exp "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND ABS(p.total_deductions-(p.absence_deduction+p.late_deduction+"
                          "COALESCE(p.unpaid_leave_deduction,0)))>1"),
            "fixable":   True,
            "fix_label": "إعادة حساب إجمالي الخصومات",
            "fix_sql":   ("UPDATE payroll SET total_deductions="
                          "absence_deduction+late_deduction+"
                          "COALESCE(unpaid_leave_deduction,0) WHERE id=?"),
        },
        {
            "id":        "unregistered_has_bank",
            "title":     "موظف غير مسجل بالتأمينات لديه راتب بنكي",
            "desc":      "الموظف غير المسجل يجب أن يكون راتبه نقدياً كاملاً",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, p.bank_salary "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND (e.social_security_registered=0 "
                          "OR e.social_security_registered IS NULL) "
                          "AND p.bank_salary > 0"),
            "fixable":   True,
            "fix_label": "نقل البنكي إلى النقدي",
            "fix_sql":   "UPDATE payroll SET cash_salary=net_salary, bank_salary=0 WHERE id=?",
        },
        {
            "id":        "active_no_payroll",
            "title":     "موظفون نشطون بدون راتب",
            "desc":      "موظفون بحالة نشط لم يُحسب لهم راتب هذا الشهر",
            "query":     ("SELECT e.id, e.first_name||' '||e.last_name, e.basic_salary "
                          "FROM employees e WHERE e.status='نشط' "
                          "AND e.id NOT IN "
                          "(SELECT employee_id FROM payroll WHERE month=? AND year=?)"),
            "fixable":   False,
        },
        {
            "id":        "orphan_inst",
            "title":     "أقساط مرتبطة برواتب غير موجودة",
            "desc":      "installments.payroll_id يشير لراتب محذوف",
            "query":     None,
            "fixable":   True,
            "fix_label": "فك ربط الأقساط اليتيمة",
            "fix_sql":   ("UPDATE installments SET payroll_id=NULL, paid_amount=0, "
                          "status='pending', notes=NULL WHERE id=?"),
            "is_inst":   True,
        },
    ]

    def __init__(self, parent, db: DatabaseManager, month: int, year: int):
        super().__init__(parent)
        self.db    = db
        self.month = month
        self.year  = year
        months = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                  'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
        self.setWindowTitle(f"🔧 صيانة الرواتب — {months[month-1]} {year}")
        self.setMinimumSize(940, 640)
        self.setLayoutDirection(Qt.RightToLeft)
        self._results = {}
        self._build()
        self._run_all_checks()

    def _build(self):
        layout = QVBoxLayout(self)

        hdr = QLabel(
            f"فحص رواتب شهر {self.month}/{self.year}  "
            f"— {len(self.CHECKS)} قاعدة فحص")
        hdr.setStyleSheet(
            "font-weight:bold; font-size:13px; padding:8px; "
            "background:#e3f2fd; border-radius:4px;")
        layout.addWidget(hdr)

        self.tbl = QTableWidget()
        self.tbl.setColumnCount(5)
        self.tbl.setHorizontalHeaderLabels(
            ["الحالة", "القاعدة", "بيانات", "عدد", "إجراء"])
        hh = self.tbl.horizontalHeader()
        for col, mode in [(0, QHeaderView.ResizeToContents),
                          (1, QHeaderView.Stretch),
                          (2, QHeaderView.Stretch),
                          (3, QHeaderView.ResizeToContents),
                          (4, QHeaderView.ResizeToContents)]:
            hh.setSectionResizeMode(col, mode)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setWordWrap(True)
        layout.addWidget(self.tbl)

        btn_row      = QHBoxLayout()
        self.btn_recheck = btn("🔄 إعادة الفحص",      BTN_PRIMARY, self._run_all_checks)
        self.btn_fix_all = btn("🔧 إصلاح كل المشاكل", BTN_SUCCESS, self._fix_all)
        self.btn_close   = btn("✅ إغلاق",              BTN_GRAY,   self.accept)
        self.lbl_sum     = QLabel("")
        self.lbl_sum.setStyleSheet("font-weight:bold; padding:4px; font-size:12px;")
        btn_row.addWidget(self.btn_recheck)
        btn_row.addWidget(self.btn_fix_all)
        btn_row.addWidget(self.lbl_sum)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_close)
        layout.addLayout(btn_row)

    def _run_all_checks(self):
        self._results.clear()
        self.tbl.setRowCount(0)
        total = fixable_cnt = 0

        for check in self.CHECKS:
            rows = self._exec_check(check)
            self._results[check['id']] = rows
            self._add_row(check, rows)
            total += len(rows)
            if rows and check.get('fixable'):
                fixable_cnt += len(rows)

        if total == 0:
            self.lbl_sum.setText("✅ لا توجد مشاكل")
            self.lbl_sum.setStyleSheet(
                "font-weight:bold; color:#388E3C; padding:4px;")
        else:
            self.lbl_sum.setText(
                f"⚠️ {total} مشكلة — {fixable_cnt} قابلة للإصلاح التلقائي")
            self.lbl_sum.setStyleSheet(
                "font-weight:bold; color:#D32F2F; padding:4px;")

    def _exec_check(self, check: dict) -> list:
        try:
            if check['id'] == 'orphan_inst':
                return self.db.fetch_all(
                    "SELECT i.id, COALESCE(e.first_name||' '||e.last_name,'?'), i.amount "
                    "FROM installments i "
                    "JOIN loans l ON i.loan_id=l.id "
                    "JOIN employees e ON l.employee_id=e.id "
                    "WHERE i.payroll_id IS NOT NULL "
                    "AND i.payroll_id NOT IN (SELECT id FROM payroll)") or []
            return self.db.fetch_all(
                check['query'], (self.month, self.year)) or []
        except Exception as e:
            return [("ERR", f"خطأ في الفحص: {e}", "—")]

    def _add_row(self, check: dict, rows: list):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        if not rows:
            si = QTableWidgetItem("✅ سليم")
            si.setForeground(QColor("#388E3C"))
        elif not check.get('fixable'):
            si = QTableWidgetItem(f"ℹ️ {len(rows)}")
            si.setForeground(QColor("#1565C0"))
        else:
            si = QTableWidgetItem(f"❌ {len(rows)}")
            si.setForeground(QColor("#C62828"))
        self.tbl.setItem(r, 0, si)

        self.tbl.setItem(r, 1, QTableWidgetItem(
            f"{check['title']}\n{check.get('desc','')}"))

        if rows:
            names = [str(row[1]) for row in rows[:6]]
            if len(rows) > 6:
                names.append(f"...و{len(rows)-6} آخرين")
            self.tbl.setItem(r, 2, QTableWidgetItem("\n".join(names)))
        else:
            self.tbl.setItem(r, 2, QTableWidgetItem("—"))

        self.tbl.setItem(r, 3, QTableWidgetItem(str(len(rows))))

        if rows and check.get('fixable'):
            fb = QPushButton(f"🔧 {check.get('fix_label','إصلاح')}")
            fb.setStyleSheet(BTN_WARNING)
            fb.clicked.connect(lambda _, c=check: self._fix_one(c))
            self.tbl.setCellWidget(r, 4, fb)
        else:
            self.tbl.setItem(r, 4, QTableWidgetItem("—"))

        self.tbl.setRowHeight(r, max(48, 18 * min(len(rows), 6) + 12))

    def _fix_one(self, check: dict):
        rows = self._results.get(check['id'], [])
        if not rows:
            return
        if QMessageBox.question(
                self, "تأكيد",
                f"إصلاح {len(rows)} سجل بـ «{check['fix_label']}»؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        self._apply_fix(check, rows)

    def _fix_all(self):
        fixable = [c for c in self.CHECKS
                   if c.get('fixable') and self._results.get(c['id'])]
        if not fixable:
            QMessageBox.information(self, "تنبيه", "لا توجد مشاكل قابلة للإصلاح")
            return
        total = sum(len(self._results[c['id']]) for c in fixable)
        if QMessageBox.question(
                self, "تأكيد",
                f"إصلاح {total} مشكلة دفعة واحدة؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        try:
            with self.db.transaction():
                fixed = 0
                for check in fixable:
                    for row in self._results[check['id']]:
                        self.db.conn.execute(check['fix_sql'], (row[0],))
                        fixed += 1
            self.db.log_action(
                "صيانة رواتب شاملة", "payroll", None, None,
                {"month": self.month, "year": self.year, "fixed": fixed})
            QMessageBox.information(self, "نجاح", f"تم إصلاح {fixed} مشكلة")
            self._run_all_checks()
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"فشل:\n{e}")

    def _apply_fix(self, check: dict, rows: list):
        try:
            with self.db.transaction():
                for row in rows:
                    self.db.conn.execute(check['fix_sql'], (row[0],))
            self.db.log_action(
                f"صيانة: {check['title']}", "payroll", None, None,
                {"month": self.month, "year": self.year, "fixed": len(rows)})
            QMessageBox.information(self, "نجاح", f"تم إصلاح {len(rows)} سجل")
            self._run_all_checks()
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"فشل:\n{e}")