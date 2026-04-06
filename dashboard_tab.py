#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/dashboard_tab.py

import logging
from datetime import datetime, date, timedelta

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QGroupBox, QFrame
)
from PyQt5.QtCore import Qt

from database import DatabaseManager
from utils import make_table, fill_table, btn
from constants import BTN_PRIMARY

logger = logging.getLogger(__name__)


class DashboardTab(QWidget):
    """
    لوحة التحكم الرئيسية — تعرض مؤشرات الأداء الفورية.

    الإصلاحات في هذه النسخة:
    - إصلاح حساب الغائبين: المنطق القديم كان يعتبر الموظفين الذين
      لا يوجد لهم أي سجل حضور اليوم "حاضرين"، والآن:
      غائب = نشط + غير معفى من البصمة + لا سجل حضور اليوم + ليس في إجازة معتمدة
    - دمج استعلامات التنبيهات الثلاثة في UNION واحد
    - تحديث العملة من الإعدادات عند كل تحميل
    - إضافة عمود القسم للموظفين الغائبين الأكثر
    """

    def __init__(self, db: DatabaseManager, user: dict, comm=None):
        super().__init__()
        self.db   = db
        self.user = user
        self.comm = comm
        self._build_ui()
        self._load_data()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    def _on_data_changed(self, data_type: str, data):
        if data_type in ('employee', 'attendance', 'leave_request',
                         'payroll', 'loan', 'settings'):
            self._load_data()

    # ==================== بناء الواجهة ====================
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # --- رسالة الترحيب ---
        welcome = QLabel(
            f"أهلاً {self.user['name']} | "
            f"{datetime.now().strftime('%A %d/%m/%Y')}")
        welcome.setStyleSheet(
            "font-size:14px; color:#1976D2; font-weight:bold; padding:8px;")
        layout.addWidget(welcome)

        # --- بطاقات المؤشرات ---
        cards = QHBoxLayout()
        self.card_employees = self._create_card("👥 الموظفون النشطون", "0",       "#1976D2")
        self.card_absent    = self._create_card("❌ غائبون اليوم",      "0",       "#D32F2F")
        self.card_leaves    = self._create_card("🏖️ في إجازة",        "0",       "#F57C00")
        self.card_payroll   = self._create_card("💰 رواتب هذا الشهر",  "0",       "#388E3C")
        for card in (self.card_employees, self.card_absent,
                     self.card_leaves, self.card_payroll):
            cards.addWidget(card)
        layout.addLayout(cards)

        # --- السطر الأول: التنبيهات + الإجازات المعلقة ---
        row1 = QHBoxLayout()

        alerts_group = QGroupBox("⚠️ وثائق تنتهي خلال 30 يوماً")
        alerts_lay   = QVBoxLayout()
        self.alerts_table = make_table(
            ["الموظف", "نوع الوثيقة", "تاريخ الانتهاء", "الأيام المتبقية"])
        alerts_lay.addWidget(self.alerts_table)
        alerts_group.setLayout(alerts_lay)
        row1.addWidget(alerts_group, 1)

        leaves_group = QGroupBox("📋 طلبات الإجازة قيد المراجعة")
        leaves_lay   = QVBoxLayout()
        self.pending_leaves = make_table(
            ["الموظف", "نوع الإجازة", "من", "إلى", "الأيام"])
        leaves_lay.addWidget(self.pending_leaves)
        leaves_group.setLayout(leaves_lay)
        row1.addWidget(leaves_group, 1)

        layout.addLayout(row1)

        # --- السطر الثاني: أكثر غياباً + أكثر أوفرتايم ---
        row2 = QHBoxLayout()

        absent_group = QGroupBox("📊 أكثر الموظفين غياباً (آخر 12 شهراً)")
        absent_lay   = QVBoxLayout()
        self.most_absent_table = make_table(["الموظف", "القسم", "أيام الغياب"])
        absent_lay.addWidget(self.most_absent_table)
        absent_group.setLayout(absent_lay)
        row2.addWidget(absent_group, 1)

        ot_group = QGroupBox("⏰ أكثر الموظفين ساعات إضافية (آخر 12 شهراً)")
        ot_lay   = QVBoxLayout()
        self.most_overtime_table = make_table(["الموظف", "القسم", "الساعات الإضافية"])
        ot_lay.addWidget(self.most_overtime_table)
        ot_group.setLayout(ot_lay)
        row2.addWidget(ot_group, 1)

        layout.addLayout(row2)

        # --- السطر الثالث: أقساط مستحقة + رواتب معلقة ---
        row3 = QHBoxLayout()

        loans_group = QGroupBox("💳 أقساط مستحقة هذا الشهر")
        loans_lay   = QVBoxLayout()
        self.due_loans_table = make_table(["الموظف", "المبلغ المستحق"])
        loans_lay.addWidget(self.due_loans_table)
        loans_group.setLayout(loans_lay)
        row3.addWidget(loans_group, 1)

        payroll_group = QGroupBox("⚠️ رواتب غير معتمدة")
        payroll_lay   = QVBoxLayout()
        self.pending_payroll_table = make_table(["الشهر", "السنة", "عدد الموظفين"])
        payroll_lay.addWidget(self.pending_payroll_table)
        payroll_group.setLayout(payroll_lay)
        row3.addWidget(payroll_group, 1)

        layout.addLayout(row3)

        refresh_btn = btn("🔄 تحديث", BTN_PRIMARY, self._load_data)
        layout.addWidget(refresh_btn, alignment=Qt.AlignLeft)

    def _create_card(self, title: str, value: str, color: str) -> QFrame:
        frame = QFrame()
        frame.setStyleSheet(f"""
            QFrame {{
                background: {color}; border-radius: 8px;
                padding: 12px; min-height: 80px;
            }}
        """)
        lay       = QVBoxLayout(frame)
        lbl_title = QLabel(title)
        lbl_title.setStyleSheet("color:white; font-size:12px; font-weight:bold;")
        lbl_value = QLabel(value)
        lbl_value.setStyleSheet("color:white; font-size:22px; font-weight:bold;")
        lbl_value.setObjectName("value")
        lay.addWidget(lbl_title)
        lay.addWidget(lbl_value)
        return frame

    def _update_card_value(self, card: QFrame, new_value) -> None:
        card.findChild(QLabel, "value").setText(str(new_value))

    # ==================== تحميل البيانات ====================
    def _load_data(self):
        try:
            self._load_cards()
            self._load_alerts()
            self._load_pending_leaves()
            self._load_most_absent()
            self._load_most_overtime()
            self._load_due_loans()
            self._load_pending_payroll()
        except Exception as e:
            logger.error("خطأ في تحميل لوحة التحكم: %s", e, exc_info=True)

    # ---------- البطاقات ----------
    def _load_cards(self):
        today    = date.today().isoformat()
        today_dt = date.today()
        m, y     = today_dt.month, today_dt.year

        # عدد الموظفين النشطين
        r = self.db.fetch_one(
            "SELECT COUNT(*) FROM employees WHERE status = 'نشط'")
        active = r[0] if r else 0
        self._update_card_value(self.card_employees, active)

        # الموظفون في إجازة معتمدة اليوم
        r_leave = self.db.fetch_one(
            """SELECT COUNT(DISTINCT employee_id)
               FROM leave_requests
               WHERE status    = 'موافق'
                 AND start_date <= ?
                 AND end_date   >= ?""",
            (today, today))
        on_leave = r_leave[0] if r_leave else 0
        self._update_card_value(self.card_leaves, on_leave)

        # ================================================================
        # إصلاح حساب الغائبين
        # ----------------------------------------------------------------
        # المنطق القديم الخاطئ:
        #   present = من لديه سجل حضور بحالة غير (غائب/إجازة)
        #   absent  = active - present - on_leave
        # المشكلة: من لا يوجد له أي سجل اليوم يُحسَب ضمن "present"
        # لأن الطرح يفترض أن الكل موجود في الجدول.
        #
        # المنطق الصحيح:
        #   غائب = نشط + غير معفى من البصمة
        #          + لا يوجد له سجل حضور اليوم بحالة إيجابية
        #          + ليس في إجازة معتمدة
        # ================================================================
        r_absent = self.db.fetch_one(
            """SELECT COUNT(*)
               FROM employees e
               WHERE e.status = 'نشط'
                 AND e.is_exempt_from_fingerprint = 0
                 AND NOT EXISTS (
                     -- له سجل حضور اليوم بحالة غير غياب
                     SELECT 1 FROM attendance a
                     WHERE a.employee_id = e.id
                       AND a.punch_date  = ?
                       AND a.status NOT IN ('غائب')
                 )
                 AND NOT EXISTS (
                     -- هو في إجازة معتمدة اليوم
                     SELECT 1 FROM leave_requests lr
                     WHERE lr.employee_id = e.id
                       AND lr.status      = 'موافق'
                       AND lr.start_date <= ?
                       AND lr.end_date   >= ?
                 )""",
            (today, today, today))
        absent = r_absent[0] if r_absent else 0
        self._update_card_value(self.card_absent, absent)

        # إجمالي صافي الرواتب للشهر الحالي
        r_pay = self.db.fetch_one(
            "SELECT COALESCE(SUM(net_salary), 0) FROM payroll "
            "WHERE month = ? AND year = ?",
            (m, y))
        total    = r_pay[0] if r_pay else 0
        currency = self.db.get_setting('currency', 'TRY')
        self._update_card_value(self.card_payroll, f"{total:,.0f} {currency}")

    # ---------- تنبيهات الوثائق ----------
    def _load_alerts(self):
        """
        الإصلاح: استعلام UNION واحد بدلاً من 3 استعلامات منفصلة في loop.
        أسرع بكثير ويُقلِّل round trips لقاعدة البيانات.
        """
        today     = date.today()
        today_str = today.isoformat()
        limit_str = (today + timedelta(days=30)).isoformat()

        data = self.db.fetch_all(
            """
            SELECT first_name || ' ' || last_name AS emp_name,
                   'الإقامة'                       AS doc_type,
                   iqama_expiry                    AS expiry
            FROM employees
            WHERE status = 'نشط'
              AND iqama_expiry IS NOT NULL AND iqama_expiry != ''
              AND iqama_expiry BETWEEN ? AND ?

            UNION ALL

            SELECT first_name || ' ' || last_name,
                   'جواز السفر',
                   passport_expiry
            FROM employees
            WHERE status = 'نشط'
              AND passport_expiry IS NOT NULL AND passport_expiry != ''
              AND passport_expiry BETWEEN ? AND ?

            UNION ALL

            SELECT first_name || ' ' || last_name,
                   'التأمين الصحي',
                   health_insurance_expiry
            FROM employees
            WHERE status = 'نشط'
              AND health_insurance_expiry IS NOT NULL
              AND health_insurance_expiry != ''
              AND health_insurance_expiry BETWEEN ? AND ?

            ORDER BY expiry ASC
            """,
            (today_str, limit_str,
             today_str, limit_str,
             today_str, limit_str)
        )

        rows = []
        for emp_name, doc_type, expiry in data:
            try:
                days = (date.fromisoformat(expiry) - today).days
                rows.append((emp_name, doc_type, expiry, days))
            except (ValueError, TypeError):
                pass

        fill_table(self.alerts_table, rows, colors={
            3: lambda v: "#D32F2F" if int(v) <= 7
               else ("#F57C00" if int(v) <= 14 else None)
        })

    # ---------- إجازات قيد المراجعة ----------
    def _load_pending_leaves(self):
        data = self.db.fetch_all(
            """SELECT e.first_name || ' ' || e.last_name,
                      lt.name,
                      lr.start_date, lr.end_date, lr.days_count
               FROM leave_requests lr
               JOIN employees   e  ON lr.employee_id   = e.id
               JOIN leave_types lt ON lr.leave_type_id = lt.id
               WHERE lr.status = 'قيد المراجعة'
               ORDER BY lr.created_at DESC
               LIMIT 20""")
        fill_table(self.pending_leaves, data)

    # ---------- أكثر الموظفين غياباً ----------
    def _load_most_absent(self):
        today = date.today()
        start = date(today.year - 1, today.month, 1)
        end   = date(today.year,     today.month, 1) - timedelta(days=1)

        data = self.db.fetch_all(
            """SELECT e.first_name || ' ' || e.last_name,
                      COALESCE(d.name, '—'),
                      COUNT(*) AS absent_days
               FROM attendance a
               JOIN employees  e ON a.employee_id  = e.id
               LEFT JOIN departments d ON e.department_id = d.id
               WHERE a.status     = 'غائب'
                 AND a.punch_date BETWEEN ? AND ?
               GROUP BY a.employee_id
               ORDER BY absent_days DESC
               LIMIT 10""",
            (start.isoformat(), end.isoformat()))
        fill_table(self.most_absent_table, data)

    # ---------- أكثر الموظفين أوفرتايم ----------
    def _load_most_overtime(self):
        today = date.today()
        start = date(today.year - 1, today.month, 1)
        end   = date(today.year,     today.month, 1) - timedelta(days=1)

        data = self.db.fetch_all(
            """SELECT e.first_name || ' ' || e.last_name,
                      COALESCE(d.name, '—'),
                      ROUND(SUM(a.overtime_hours), 1) AS total_ot
               FROM attendance a
               JOIN employees  e ON a.employee_id  = e.id
               LEFT JOIN departments d ON e.department_id = d.id
               WHERE a.overtime_hours > 0
                 AND a.punch_date BETWEEN ? AND ?
               GROUP BY a.employee_id
               ORDER BY total_ot DESC
               LIMIT 10""",
            (start.isoformat(), end.isoformat()))
        fill_table(self.most_overtime_table, data)

    # ---------- أقساط مستحقة ----------
    def _load_due_loans(self):
        today = date.today()
        month_start = date(today.year, today.month, 1).isoformat()
        if today.month == 12:
            next_month = date(today.year + 1, 1, 1).isoformat()
        else:
            next_month = date(today.year, today.month + 1, 1).isoformat()

        currency = self.db.get_setting('currency', 'TRY')

        raw = self.db.fetch_all(
            """SELECT e.first_name || ' ' || e.last_name,
                      SUM(i.amount - COALESCE(i.paid_amount, 0)) AS due
               FROM installments i
               JOIN loans    l ON i.loan_id     = l.id
               JOIN employees e ON l.employee_id = e.id
               WHERE i.due_date >= ? AND i.due_date < ?
                 AND i.status   != 'paid'
               GROUP BY l.employee_id
               ORDER BY due DESC""",
            (month_start, next_month))

        # إضافة رمز العملة للعرض
        data = [(name, f"{due:,.2f} {currency}") for name, due in raw]
        fill_table(self.due_loans_table, data)

    # ---------- رواتب غير معتمدة ----------
    def _load_pending_payroll(self):
        data = self.db.fetch_all(
            """SELECT month, year, COUNT(*) AS emp_count
               FROM payroll
               WHERE status = 'مسودة'
               GROUP BY month, year
               ORDER BY year DESC, month DESC""")
        fill_table(self.pending_payroll_table, data)