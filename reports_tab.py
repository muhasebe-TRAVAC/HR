#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/reports_tab.py

import logging
from datetime import datetime, date

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QSpinBox,
    QDateEdit, QPushButton, QMessageBox, QFileDialog, QTabWidget,
    QGroupBox, QCompleter, QCheckBox
)
from PyQt5.QtCore import Qt, QDate

from database import DatabaseManager
from utils import make_table, fill_table, btn
from constants import BTN_PRIMARY, BTN_SUCCESS, BTN_TEAL, BTN_PURPLE

logger = logging.getLogger(__name__)

_MONTHS_AR = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
               'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
_MONTHS_TR = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
               'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']


class ReportsTab(QWidget):
    """
    تبويب التقارير.

    الإصلاحات:
    - حُذفت الإشارات إلى أعمدة غير موجودة في جدول payroll:
      (actual_days, social_security, loan_deduction)
      والاستعاضة عنها بالأعمدة الصحيحة.
    - استعلامات SARIE: ترتيب المعاملات مُصحَّح.
    - إضافة logging لأخطاء التصدير.
    - _load_payroll_report: يجمع الخصومات من الأعمدة الصحيحة.
    """

    def __init__(self, db: DatabaseManager, user: dict, comm=None):
        super().__init__()
        self.db   = db
        self.user = user
        self.comm = comm
        self._build()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    def _on_data_changed(self, data_type: str, data):
        if data_type == 'employee':
            self._refresh_employee_filters()
        elif data_type == 'department':
            self._refresh_department_filters()

    def _refresh_employee_filters(self):
        emps = self.db.fetch_all(
            "SELECT id, first_name||' '||last_name "
            "FROM employees WHERE status='نشط'")
        for combo in (self.ar_emp_filter, self.gr_emp_filter):
            combo.clear()
            combo.addItem("جميع الموظفين", None)
            for eid, name in emps:
                combo.addItem(name, eid)

    def _refresh_department_filters(self):
        depts = self.db.fetch_all("SELECT id, name FROM departments ORDER BY name")
        self.pr_dept_filter.clear()
        self.pr_dept_filter.addItem("جميع الأقسام", None)
        for did, name in depts:
            self.pr_dept_filter.addItem(name, did)

    def _build(self):
        layout = QVBoxLayout(self)
        tabs   = QTabWidget()
        tabs.addTab(self._build_payroll_report(),   "📊 تقرير الرواتب")
        tabs.addTab(self._build_attendance_report(),"📅 تقرير الحضور")
        tabs.addTab(self._build_gosi_report(),      "🏥 تقرير التأمينات")
        tabs.addTab(self._build_sarie_report(),     "🏦 تحويل بنكي (SARIE)")
        tabs.addTab(self._build_employees_report(), "👥 تقرير الموظفين")
        layout.addWidget(tabs)

    # ==================== تقرير الرواتب ====================
    def _build_payroll_report(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        ctrl = QHBoxLayout()
        self.pr_month = QComboBox()
        self.pr_month.addItems(_MONTHS_AR)
        self.pr_month.setCurrentIndex(date.today().month - 1)
        self.pr_year = QSpinBox()
        self.pr_year.setRange(2020, 2050)
        self.pr_year.setValue(date.today().year)

        self.pr_dept_filter = QComboBox()
        self.pr_dept_filter.addItem("جميع الأقسام", None)
        for did, name in self.db.fetch_all(
                "SELECT id, name FROM departments ORDER BY name"):
            self.pr_dept_filter.addItem(name, did)
        self.pr_dept_filter.setEditable(True)
        self.pr_dept_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.pr_dept_filter.completer().setFilterMode(Qt.MatchContains)

        self.pr_status_filter = QComboBox()
        self.pr_status_filter.addItems(["جميع الحالات", "مسودة", "معتمد"])

        for label, widget in [
            ("الشهر:", self.pr_month), ("السنة:", self.pr_year),
            ("القسم:", self.pr_dept_filter), ("الحالة:", self.pr_status_filter),
        ]:
            ctrl.addWidget(QLabel(label))
            ctrl.addWidget(widget)

        ctrl.addWidget(btn("عرض",     BTN_PRIMARY, self._load_payroll_report))
        ctrl.addWidget(btn("📊 Excel", BTN_SUCCESS, self._export_payroll_excel))
        ctrl.addStretch()
        lay.addLayout(ctrl)

        # الأعمدة: بدون actual_days (غير موجود في payroll)
        self.pr_table = make_table([
            "الموظف", "القسم", "الأساسي", "البدلات",
            "أوفرتايم", "إجمالي الاستحقاق", "الخصومات", "الصافي", "الحالة"
        ])
        lay.addWidget(self.pr_table)

        self.pr_summary = QLabel("")
        self.pr_summary.setStyleSheet(
            "font-weight:bold; font-size:13px; padding:8px; background:#f5f5f5;")
        lay.addWidget(self.pr_summary)
        return w

    def _load_payroll_report(self):
        m       = self.pr_month.currentIndex() + 1
        y       = self.pr_year.value()
        dept_id = self.pr_dept_filter.currentData()
        status  = self.pr_status_filter.currentText()

        # الخصومات = غياب + تأخير + إجازة بدون راتب (بدون الأقساط)
        q = """
            SELECT e.first_name||' '||e.last_name,
                   COALESCE(d.name,'') AS department,
                   p.basic_salary,
                   p.housing_allowance + p.transportation_allowance +
                   p.food_allowance + p.phone_allowance + p.other_allowances,
                   p.overtime_amount,
                   p.total_earnings,
                   p.absence_deduction + p.late_deduction +
                   COALESCE(p.unpaid_leave_deduction, 0),
                   p.net_salary,
                   p.status
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE p.month = ? AND p.year = ?
        """
        params = [m, y]
        if dept_id is not None:
            q += " AND e.department_id = ?"
            params.append(dept_id)
        if status != "جميع الحالات":
            q += " AND p.status = ?"
            params.append(status)
        q += " ORDER BY d.name, e.first_name"

        data = self.db.fetch_all(q, params) or []
        fill_table(self.pr_table, data)

        if data:
            currency   = self.db.get_setting('currency', 'ريال')
            tot_earn   = sum(r[5] for r in data if r[5])
            tot_ded    = sum(r[6] for r in data if r[6])
            tot_net    = sum(r[7] for r in data if r[7])
            self.pr_summary.setText(
                f"إجمالي الاستحقاق: {tot_earn:,.0f} {currency}  |  "
                f"إجمالي الخصومات: {tot_ded:,.0f} {currency}  |  "
                f"إجمالي الصافي: {tot_net:,.0f} {currency}  |  "
                f"عدد الموظفين: {len(data)}")
        else:
            self.pr_summary.setText("لا توجد بيانات تطابق المعايير المحددة")

    def _export_payroll_excel(self):
        try:
            import pandas as pd
            m       = self.pr_month.currentIndex() + 1
            y       = self.pr_year.value()
            dept_id = self.pr_dept_filter.currentData()
            status  = self.pr_status_filter.currentText()

            q = """
                SELECT e.employee_code, e.first_name||' '||e.last_name,
                       COALESCE(d.name,'') AS department,
                       p.basic_salary,
                       p.housing_allowance, p.transportation_allowance,
                       p.food_allowance, p.phone_allowance, p.other_allowances,
                       p.overtime_amount, p.bonus, p.total_earnings,
                       p.absence_deduction, p.late_deduction,
                       COALESCE(p.unpaid_leave_deduction, 0),
                       p.loan_deduction_bank + COALESCE(p.loan_deduction_cash, 0),
                       p.total_deductions, p.net_salary, p.status
                FROM payroll p
                JOIN employees e ON p.employee_id = e.id
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE p.month = ? AND p.year = ?
            """
            params = [m, y]
            if dept_id is not None:
                q += " AND e.department_id = ?"
                params.append(dept_id)
            if status != "جميع الحالات":
                q += " AND p.status = ?"
                params.append(status)
            q += " ORDER BY d.name, e.first_name"

            data = self.db.fetch_all(q, params)
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return

            df = pd.DataFrame(data, columns=[
                "الرقم الوظيفي", "الاسم", "القسم", "الأساسي",
                "بدل سكن", "بدل نقل", "بدل غذاء", "بدل هاتف", "بدلات أخرى",
                "أوفرتايم", "مكافأة", "إجمالي الاستحقاق",
                "خصم غياب", "خصم تأخير", "إجازة بدون راتب", "خصم سلف",
                "إجمالي الخصومات", "الصافي", "الحالة"
            ])
            fname = f"payroll_report_{_MONTHS_AR[m-1]}_{y}.xlsx"
            path, _ = QFileDialog.getSaveFileName(
                self, "حفظ", fname, "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", "تم تصدير التقرير")
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")
        except Exception as e:
            logger.error("خطأ في تصدير تقرير الرواتب: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))

    # ==================== تقرير الحضور ====================
    def _build_attendance_report(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        ctrl = QHBoxLayout()
        self.ar_from = QDateEdit()
        self.ar_from.setCalendarPopup(True)
        self.ar_from.setDate(QDate.currentDate().addDays(-30))
        self.ar_to = QDateEdit()
        self.ar_to.setCalendarPopup(True)
        self.ar_to.setDate(QDate.currentDate())

        self.ar_emp_filter = QComboBox()
        self.ar_emp_filter.addItem("جميع الموظفين", None)
        for eid, name in self.db.fetch_all(
                "SELECT id, first_name||' '||last_name "
                "FROM employees WHERE status='نشط'"):
            self.ar_emp_filter.addItem(name, eid)
        self.ar_emp_filter.setEditable(True)
        self.ar_emp_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.ar_emp_filter.completer().setFilterMode(Qt.MatchContains)

        self.ar_dept_filter = QComboBox()
        self.ar_dept_filter.addItem("جميع الأقسام", None)
        for did, name in self.db.fetch_all(
                "SELECT id, name FROM departments ORDER BY name"):
            self.ar_dept_filter.addItem(name, did)
        self.ar_dept_filter.setEditable(True)
        self.ar_dept_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.ar_dept_filter.completer().setFilterMode(Qt.MatchContains)

        for label, widget in [
            ("من:", self.ar_from), ("إلى:", self.ar_to),
            ("الموظف:", self.ar_emp_filter), ("القسم:", self.ar_dept_filter),
        ]:
            ctrl.addWidget(QLabel(label))
            ctrl.addWidget(widget)
        ctrl.addWidget(btn("عرض",     BTN_PRIMARY, self._load_attendance_report))
        ctrl.addWidget(btn("📊 Excel", BTN_SUCCESS, self._export_attendance_excel))
        ctrl.addStretch()
        lay.addLayout(ctrl)

        self.ar_table = make_table([
            "الموظف", "القسم", "أيام الحضور", "إجمالي الساعات",
            "أوفرتايم", "تأخير (د)", "غياب"
        ])
        lay.addWidget(self.ar_table)
        return w

    def _load_attendance_report(self):
        d1      = self.ar_from.date().toString(Qt.ISODate)
        d2      = self.ar_to.date().toString(Qt.ISODate)
        emp_id  = self.ar_emp_filter.currentData()
        dept_id = self.ar_dept_filter.currentData()

        q = """
            SELECT e.first_name||' '||e.last_name,
                   COALESCE(d.name,'') AS department,
                   COUNT(CASE WHEN a.status='حاضر' THEN 1 END),
                   ROUND(SUM(a.work_hours), 1),
                   ROUND(SUM(a.overtime_hours), 1),
                   SUM(a.late_minutes),
                   COUNT(CASE WHEN a.status='غائب' THEN 1 END)
            FROM attendance a
            JOIN employees e ON a.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE a.punch_date BETWEEN ? AND ?
        """
        params = [d1, d2]
        if emp_id is not None:
            q += " AND a.employee_id = ?"
            params.append(emp_id)
        if dept_id is not None:
            q += " AND e.department_id = ?"
            params.append(dept_id)
        q += " GROUP BY a.employee_id ORDER BY e.first_name"

        fill_table(self.ar_table, self.db.fetch_all(q, params) or [])

    def _export_attendance_excel(self):
        try:
            import pandas as pd
            d1      = self.ar_from.date().toString(Qt.ISODate)
            d2      = self.ar_to.date().toString(Qt.ISODate)
            emp_id  = self.ar_emp_filter.currentData()
            dept_id = self.ar_dept_filter.currentData()

            q = """
                SELECT e.employee_code, e.first_name||' '||e.last_name,
                       COALESCE(d.name,''),
                       COUNT(CASE WHEN a.status='حاضر' THEN 1 END),
                       ROUND(SUM(a.work_hours), 1),
                       ROUND(SUM(a.overtime_hours), 1),
                       SUM(a.late_minutes),
                       COUNT(CASE WHEN a.status='غائب' THEN 1 END)
                FROM attendance a
                JOIN employees e ON a.employee_id = e.id
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE a.punch_date BETWEEN ? AND ?
            """
            params = [d1, d2]
            if emp_id is not None:
                q += " AND a.employee_id = ?"
                params.append(emp_id)
            if dept_id is not None:
                q += " AND e.department_id = ?"
                params.append(dept_id)
            q += " GROUP BY a.employee_id ORDER BY e.first_name"

            data = self.db.fetch_all(q, params)
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return
            df = pd.DataFrame(data, columns=[
                "الرقم", "الاسم", "القسم", "أيام الحضور",
                "إجمالي الساعات", "أوفرتايم", "تأخير(د)", "غياب"])
            path, _ = QFileDialog.getSaveFileName(
                self, "حفظ",
                f"attendance_report_{d1}_{d2}.xlsx", "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", "تم التصدير")
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")
        except Exception as e:
            logger.error("خطأ في تصدير تقرير الحضور: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))

    # ==================== تقرير التأمينات (GOSI) ====================
    def _build_gosi_report(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        ctrl = QHBoxLayout()
        self.gr_month = QComboBox()
        self.gr_month.addItems(_MONTHS_AR)
        self.gr_month.setCurrentIndex(date.today().month - 1)
        self.gr_year = QSpinBox()
        self.gr_year.setRange(2020, 2050)
        self.gr_year.setValue(date.today().year)

        self.gr_emp_filter = QComboBox()
        self.gr_emp_filter.addItem("جميع الموظفين", None)
        for eid, name in self.db.fetch_all(
                "SELECT id, first_name||' '||last_name "
                "FROM employees WHERE status='نشط' "
                "AND social_security_registered=1"):
            self.gr_emp_filter.addItem(name, eid)
        self.gr_emp_filter.setEditable(True)
        self.gr_emp_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.gr_emp_filter.completer().setFilterMode(Qt.MatchContains)

        self.gr_dept_filter = QComboBox()
        self.gr_dept_filter.addItem("جميع الأقسام", None)
        for did, name in self.db.fetch_all(
                "SELECT id, name FROM departments ORDER BY name"):
            self.gr_dept_filter.addItem(name, did)
        self.gr_dept_filter.setEditable(True)
        self.gr_dept_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.gr_dept_filter.completer().setFilterMode(Qt.MatchContains)

        for label, widget in [
            ("الشهر:", self.gr_month), ("السنة:", self.gr_year),
            ("الموظف:", self.gr_emp_filter), ("القسم:", self.gr_dept_filter),
        ]:
            ctrl.addWidget(QLabel(label))
            ctrl.addWidget(widget)
        ctrl.addWidget(btn("عرض",     BTN_PRIMARY, self._load_gosi))
        ctrl.addWidget(btn("📊 Excel", BTN_SUCCESS, self._export_gosi))
        ctrl.addStretch()
        lay.addLayout(ctrl)

        self.gr_table = make_table([
            "الموظف", "القسم", "رقم التأمينات", "الراتب الأساسي",
            "اشتراك الموظف", "اشتراك الشركة", "الإجمالي"
        ])
        lay.addWidget(self.gr_table)

        self.gr_summary = QLabel("")
        self.gr_summary.setStyleSheet(
            "font-weight:bold; padding:8px; background:#f5f5f5;")
        lay.addWidget(self.gr_summary)
        return w

    def _load_gosi(self):
        m       = self.gr_month.currentIndex() + 1
        y       = self.gr_year.value()
        emp_id  = self.gr_emp_filter.currentData()
        dept_id = self.gr_dept_filter.currentData()
        emp_pct = float(self.db.get_setting('gosi_employee_percent', '9.75'))
        co_pct  = float(self.db.get_setting('gosi_company_percent',  '12.0'))
        max_sal = float(self.db.get_setting('gosi_max_salary', '45000'))

        q = """
            SELECT e.first_name||' '||e.last_name,
                   COALESCE(d.name,'') AS department,
                   COALESCE(e.social_security_number,''),
                   p.basic_salary,
                   ROUND(MIN(p.basic_salary, ?) * ? / 100, 2) AS emp_share,
                   ROUND(MIN(p.basic_salary, ?) * ? / 100, 2) AS co_share,
                   ROUND(MIN(p.basic_salary, ?) * (? + ?) / 100, 2) AS total
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE p.month = ? AND p.year = ?
              AND e.social_security_registered = 1
        """
        params = [max_sal, emp_pct, max_sal, co_pct, max_sal, emp_pct, co_pct, m, y]
        if emp_id is not None:
            q += " AND e.id = ?"
            params.append(emp_id)
        if dept_id is not None:
            q += " AND e.department_id = ?"
            params.append(dept_id)

        data = self.db.fetch_all(q, params) or []
        fill_table(self.gr_table, data)

        if data:
            currency = self.db.get_setting('currency', 'ريال')
            tot_emp  = sum(r[4] for r in data if r[4])
            tot_co   = sum(r[5] for r in data if r[5])
            tot_all  = sum(r[6] for r in data if r[6])
            self.gr_summary.setText(
                f"إجمالي اشتراك الموظفين: {tot_emp:,.0f} {currency}  |  "
                f"إجمالي اشتراك الشركة: {tot_co:,.0f} {currency}  |  "
                f"إجمالي التأمينات: {tot_all:,.0f} {currency}")
        else:
            self.gr_summary.setText("لا توجد بيانات لهذا الشهر")

    def _export_gosi(self):
        try:
            import pandas as pd
            m       = self.gr_month.currentIndex() + 1
            y       = self.gr_year.value()
            emp_id  = self.gr_emp_filter.currentData()
            dept_id = self.gr_dept_filter.currentData()
            emp_pct = float(self.db.get_setting('gosi_employee_percent', '9.75'))
            co_pct  = float(self.db.get_setting('gosi_company_percent',  '12.0'))
            max_sal = float(self.db.get_setting('gosi_max_salary', '45000'))

            q = """
                SELECT e.employee_code,
                       e.first_name||' '||e.last_name,
                       COALESCE(d.name,''),
                       e.national_id,
                       COALESCE(e.social_security_number,''),
                       e.nationality,
                       p.basic_salary,
                       ROUND(MIN(p.basic_salary, ?) * ? / 100, 2),
                       ROUND(MIN(p.basic_salary, ?) * ? / 100, 2)
                FROM payroll p
                JOIN employees e ON p.employee_id = e.id
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE p.month = ? AND p.year = ?
                  AND e.social_security_registered = 1
            """
            params = [max_sal, emp_pct, max_sal, co_pct, m, y]
            if emp_id is not None:
                q += " AND e.id = ?"
                params.append(emp_id)
            if dept_id is not None:
                q += " AND e.department_id = ?"
                params.append(dept_id)

            data = self.db.fetch_all(q, params)
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return
            df = pd.DataFrame(data, columns=[
                "رقم الموظف", "الاسم", "القسم", "رقم الهوية",
                "رقم التأمينات", "الجنسية", "الراتب الأساسي",
                "اشتراك الموظف", "اشتراك الشركة"])
            path, _ = QFileDialog.getSaveFileName(
                self, "حفظ",
                f"GOSI_{_MONTHS_AR[m-1]}_{y}.xlsx", "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", "تم تصدير تقرير التأمينات")
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")
        except Exception as e:
            logger.error("خطأ في تصدير تقرير التأمينات: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))

    # ==================== تقرير الموظفين ====================
    def _build_employees_report(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        ctrl = QHBoxLayout()
        self.er_dept_filter = QComboBox()
        self.er_dept_filter.addItem("جميع الأقسام", None)
        for did, name in self.db.fetch_all(
                "SELECT id, name FROM departments ORDER BY name"):
            self.er_dept_filter.addItem(name, did)
        self.er_dept_filter.setEditable(True)
        self.er_dept_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.er_dept_filter.completer().setFilterMode(Qt.MatchContains)

        self.er_status_filter = QComboBox()
        self.er_status_filter.addItems(
            ["جميع الحالات", "نشط", "غير نشط", "إجازة", "منتهي الخدمة"])

        self.er_social_filter = QComboBox()
        self.er_social_filter.addItems(
            ["الكل", "مسجل في التأمينات", "غير مسجل"])

        for label, widget in [
            ("القسم:", self.er_dept_filter),
            ("الحالة:", self.er_status_filter),
            ("التأمينات:", self.er_social_filter),
        ]:
            ctrl.addWidget(QLabel(label))
            ctrl.addWidget(widget)
        ctrl.addWidget(btn("عرض",     BTN_PRIMARY, self._load_employees_report))
        ctrl.addWidget(btn("📊 Excel", BTN_SUCCESS, self._export_employees_excel))
        ctrl.addStretch()
        lay.addLayout(ctrl)

        self.er_table = make_table([
            "الرقم", "الاسم", "القسم", "الوظيفة",
            "الحالة", "الراتب الأساسي", "مسجل في التأمينات"
        ])
        lay.addWidget(self.er_table)

        self.er_summary = QLabel("")
        self.er_summary.setStyleSheet(
            "font-weight:bold; padding:8px; background:#f5f5f5;")
        lay.addWidget(self.er_summary)
        return w

    def _load_employees_report(self):
        dept_id = self.er_dept_filter.currentData()
        status  = self.er_status_filter.currentText()
        social  = self.er_social_filter.currentText()

        q = """
            SELECT e.employee_code,
                   e.first_name||' '||e.last_name,
                   COALESCE(d.name,''),
                   e.position, e.status,
                   e.basic_salary,
                   CASE WHEN e.social_security_registered=1 THEN 'نعم' ELSE 'لا' END
            FROM employees e
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE 1=1
        """
        params = []
        if dept_id is not None:
            q += " AND e.department_id = ?"
            params.append(dept_id)
        if status != "جميع الحالات":
            q += " AND e.status = ?"
            params.append(status)
        if social == "مسجل في التأمينات":
            q += " AND e.social_security_registered = 1"
        elif social == "غير مسجل":
            q += " AND (e.social_security_registered=0 OR e.social_security_registered IS NULL)"
        q += " ORDER BY e.first_name"

        data = self.db.fetch_all(q, params) or []
        fill_table(self.er_table, data)

        currency     = self.db.get_setting('currency', 'ريال')
        total_salary = sum(r[5] for r in data if r[5])
        self.er_summary.setText(
            f"عدد الموظفين: {len(data)} | "
            f"إجمالي الرواتب: {total_salary:,.0f} {currency}")

    def _export_employees_excel(self):
        try:
            import pandas as pd
            dept_id = self.er_dept_filter.currentData()
            status  = self.er_status_filter.currentText()
            social  = self.er_social_filter.currentText()

            q = """
                SELECT e.employee_code, e.first_name||' '||e.last_name,
                       COALESCE(d.name,''), e.position, e.status,
                       e.basic_salary, e.housing_allowance,
                       e.transportation_allowance, e.food_allowance,
                       e.phone_allowance, e.other_allowances,
                       e.bank_salary, e.cash_salary,
                       e.phone, e.email, e.address,
                       e.bank_name, e.bank_account, e.iban,
                       e.fingerprint_id, e.social_security_number,
                       CASE WHEN e.social_security_registered=1 THEN 'نعم' ELSE 'لا' END,
                       e.hire_date, e.birth_date, e.nationality, e.gender,
                       e.iqama_number, e.iqama_expiry,
                       e.passport_number, e.passport_expiry,
                       e.health_insurance_number, e.health_insurance_expiry,
                       e.notes
                FROM employees e
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE 1=1
            """
            params = []
            if dept_id is not None:
                q += " AND e.department_id = ?"
                params.append(dept_id)
            if status != "جميع الحالات":
                q += " AND e.status = ?"
                params.append(status)
            if social == "مسجل في التأمينات":
                q += " AND e.social_security_registered = 1"
            elif social == "غير مسجل":
                q += " AND (e.social_security_registered=0 OR e.social_security_registered IS NULL)"
            q += " ORDER BY e.first_name"

            data = self.db.fetch_all(q, params)
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return
            cols = [
                "الرقم الوظيفي", "الاسم", "القسم", "الوظيفة", "الحالة",
                "الراتب الأساسي", "بدل سكن", "بدل نقل", "بدل غذاء",
                "بدل هاتف", "بدلات أخرى", "راتب بنكي", "راتب نقدي",
                "الهاتف", "البريد", "العنوان", "اسم البنك",
                "رقم الحساب", "IBAN", "رقم البصمة", "رقم التأمينات",
                "مسجل في التأمينات", "تاريخ التوظيف", "تاريخ الميلاد",
                "الجنسية", "الجنس", "رقم الإقامة", "انتهاء الإقامة",
                "رقم الجواز", "انتهاء الجواز",
                "رقم التأمين الصحي", "انتهاء التأمين الصحي", "ملاحظات"
            ]
            df = pd.DataFrame(data, columns=cols)
            path, _ = QFileDialog.getSaveFileName(
                self, "حفظ", "employees_full_report.xlsx", "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", "تم تصدير التقرير")
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")
        except Exception as e:
            logger.error("خطأ في تصدير تقرير الموظفين: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))

    # ==================== تقرير SARIE ====================
    def _build_sarie_report(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        info = QLabel(
            "ملف SARIE هو ملف نصي يُرفع للبنك لتحويل الرواتب دفعة واحدة.\n"
            "يحتوي على: IBAN الموظف، الاسم، المبلغ، والمرجع.")
        info.setStyleSheet(
            "background:#e3f2fd; padding:10px; border-radius:6px; color:#333;")
        info.setWordWrap(True)
        lay.addWidget(info)

        ctrl = QHBoxLayout()
        self.sr_month = QComboBox()
        self.sr_month.addItems(_MONTHS_AR)
        self.sr_month.setCurrentIndex(date.today().month - 1)
        self.sr_year = QSpinBox()
        self.sr_year.setRange(2020, 2050)
        self.sr_year.setValue(date.today().year)

        self.sr_bank_filter = QComboBox()
        self.sr_bank_filter.addItem("جميع البنوك", None)
        for (bk,) in self.db.fetch_all(
                "SELECT DISTINCT bank_name FROM employees "
                "WHERE bank_name IS NOT NULL AND bank_name != ''"):
            if bk:
                self.sr_bank_filter.addItem(bk, bk)
        self.sr_bank_filter.setEditable(True)
        self.sr_bank_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.sr_bank_filter.completer().setFilterMode(Qt.MatchContains)

        self.sr_include_cash = QCheckBox("تضمين الرواتب النقدية (بدون IBAN)")

        for label, widget in [
            ("الشهر:", self.sr_month), ("السنة:", self.sr_year),
            ("البنك:", self.sr_bank_filter),
        ]:
            ctrl.addWidget(QLabel(label))
            ctrl.addWidget(widget)
        ctrl.addWidget(self.sr_include_cash)
        ctrl.addWidget(btn("عرض",         BTN_PRIMARY, self._load_sarie))
        ctrl.addWidget(btn("💾 تصدير SARIE", BTN_SUCCESS, self._export_sarie))
        ctrl.addWidget(btn("📊 Excel",     BTN_TEAL,    self._export_sarie_excel))
        ctrl.addStretch()
        lay.addLayout(ctrl)

        self.sr_table = make_table(
            ["الموظف", "IBAN", "البنك", "صافي الراتب", "المرجع"])
        lay.addWidget(self.sr_table)
        return w

    def _sarie_query(self, m: int, y: int,
                     include_cash: bool, bank_filter) -> tuple:
        """
        بناء استعلام SARIE مع ترتيب المعاملات الصحيح.

        المعاملات في الاستعلام:
        1. م (للمرجع)   2. س (للمرجع)
        3. م (WHERE)    4. س (WHERE)
        [5. اسم البنك إن وُجد]
        """
        ref_expr = "'SAL-'||e.employee_code||'-'||?||'-'||?"
        base_params = [m, y, m, y]

        if include_cash:
            q = f"""
                SELECT e.first_name||' '||e.last_name,
                       COALESCE(e.iban,'نقدي'),
                       COALESCE(e.bank_name,'نقدي'),
                       p.bank_salary,
                       {ref_expr}
                FROM payroll p
                JOIN employees e ON p.employee_id = e.id
                WHERE p.month=? AND p.year=? AND p.status='معتمد'
            """
        else:
            q = f"""
                SELECT e.first_name||' '||e.last_name,
                       e.iban, e.bank_name,
                       p.bank_salary,
                       {ref_expr}
                FROM payroll p
                JOIN employees e ON p.employee_id = e.id
                WHERE p.month=? AND p.year=? AND p.status='معتمد'
                  AND e.iban IS NOT NULL AND e.iban != ''
                  AND p.bank_salary > 0
            """

        params = list(base_params)
        if bank_filter is not None:
            q += " AND e.bank_name = ?"
            params.append(bank_filter)
        q += " ORDER BY e.first_name"
        return q, params

    def _load_sarie(self):
        m   = self.sr_month.currentIndex() + 1
        y   = self.sr_year.value()
        bk  = self.sr_bank_filter.currentData()
        inc = self.sr_include_cash.isChecked()
        q, params = self._sarie_query(m, y, inc, bk)
        fill_table(self.sr_table, self.db.fetch_all(q, params) or [])

    def _export_sarie(self):
        m   = self.sr_month.currentIndex() + 1
        y   = self.sr_year.value()
        bk  = self.sr_bank_filter.currentData()
        inc = self.sr_include_cash.isChecked()
        q, params = self._sarie_query(m, y, inc, bk)

        data = self.db.fetch_all(q, params)
        if not data:
            QMessageBox.warning(
                self, "تنبيه",
                "لا يوجد رواتب معتمدة تطابق المعايير المحددة.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "حفظ ملف SARIE",
            f"SARIE_{_MONTHS_AR[m-1]}_{y}.txt", "Text Files (*.txt)")
        if not path:
            return

        lines = []
        total = 0.0
        for name, iban, bank, amount, ref in data:
            if iban and iban != 'نقدي' and amount and amount > 0:
                lines.append(
                    f"{iban}|{name}|{amount:.2f}|{ref}|SAR|"
                    f"{date.today().strftime('%d%m%Y')}")
                total += amount

        with open(path, 'w', encoding='utf-8') as f:
            f.write("\n".join(lines))

        currency = self.db.get_setting('currency', 'ريال')
        QMessageBox.information(
            self, "نجاح",
            f"تم تصدير ملف SARIE\n"
            f"عدد التحويلات: {len(lines)}\n"
            f"المجموع: {total:,.2f} {currency}")

    def _export_sarie_excel(self):
        try:
            import pandas as pd
            m   = self.sr_month.currentIndex() + 1
            y   = self.sr_year.value()
            bk  = self.sr_bank_filter.currentData()
            inc = self.sr_include_cash.isChecked()

            # استعلام موسَّع للإكسل
            ref_expr = "'SAL-'||e.employee_code||'-'||?||'-'||?"
            base_params = [m, y, m, y]

            if inc:
                q = f"""
                    SELECT e.employee_code,
                           e.first_name||' '||e.last_name,
                           COALESCE(e.iban,'نقدي'),
                           COALESCE(e.bank_name,'نقدي'),
                           p.bank_salary,
                           {ref_expr}
                    FROM payroll p
                    JOIN employees e ON p.employee_id = e.id
                    WHERE p.month=? AND p.year=? AND p.status='معتمد'
                """
            else:
                q = f"""
                    SELECT e.employee_code,
                           e.first_name||' '||e.last_name,
                           e.iban, e.bank_name, p.bank_salary,
                           {ref_expr}
                    FROM payroll p
                    JOIN employees e ON p.employee_id = e.id
                    WHERE p.month=? AND p.year=? AND p.status='معتمد'
                      AND e.iban IS NOT NULL AND e.iban != ''
                      AND p.bank_salary > 0
                """

            params = list(base_params)
            if bk is not None:
                q += " AND e.bank_name = ?"
                params.append(bk)
            q += " ORDER BY e.first_name"

            data = self.db.fetch_all(q, params)
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return
            df = pd.DataFrame(data, columns=[
                "رقم الموظف", "الاسم", "IBAN", "البنك",
                "صافي الراتب", "المرجع"])
            path, _ = QFileDialog.getSaveFileName(
                self, "حفظ",
                f"transfers_{_MONTHS_AR[m-1]}_{y}.xlsx", "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", "تم التصدير")
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")
        except Exception as e:
            logger.error("خطأ في تصدير SARIE Excel: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))
