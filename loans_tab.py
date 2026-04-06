#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/loans_tab.py

import math
import logging
from datetime import datetime, date, timedelta
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton,
    QMessageBox, QDialog, QFormLayout, QLineEdit, QSpinBox, QDoubleSpinBox,
    QDialogButtonBox, QGroupBox, QInputDialog, QDateEdit, QFileDialog,
    QHeaderView, QAbstractItemView, QCompleter, QSplitter, QTableWidgetItem,
    QCheckBox, QButtonGroup, QApplication, QTextEdit
)
from PyQt5.QtCore import Qt, QDate, QTimer
from PyQt5.QtGui import QColor

from database import DatabaseManager
from utils import make_table, fill_table, btn, can_add, can_edit, can_delete
from constants import BTN_SUCCESS, BTN_PRIMARY, BTN_DANGER, BTN_GRAY, BTN_WARNING, BTN_PURPLE, BTN_TEAL

logger = logging.getLogger(__name__)


class LoansTab(QWidget):
    def __init__(self, db: DatabaseManager, user: dict, comm=None):
        super().__init__()
        self.db             = db
        self.user           = user
        self.comm           = comm
        self.current_loan_id = None
        self._build()
        QTimer.singleShot(100, self._initial_load)
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    def _initial_load(self):
        """
        التحقق من وجود جدول installments يتم في database.py.
        هنا نبدأ التحميل مباشرة — لا تكرار لكود الـ migration.
        """
        self._refresh_employee_filters()
        self._load_loans()
        self._load_installments()

    def _on_data_changed(self, data_type: str, data):
        if data_type == 'employee':
            self._refresh_employee_filters()
            self._load_installments()
        elif data_type == 'loan':
            self._load_loans()
            self._load_installments()
        elif data_type == 'payroll':
            if data and data.get('action') == 'approve_all':
                self._update_installments_status_from_payroll(
                    data.get('month'), data.get('year'))
            self._load_installments()
        elif data_type == 'settings':
            self._load_loans()
            self._load_installments()

    def _update_installments_status_from_payroll(self, month, year):
        """تحديث حالة الأقساط إلى 'paid' بعد اعتماد الراتب."""
        try:
            payroll_ids = self.db.fetch_all(
                "SELECT id FROM payroll WHERE month=? AND year=? AND status='معتمد'",
                (month, year)
            )
            for (pid,) in payroll_ids:
                self.db.execute_query(
                    "UPDATE installments SET status='paid' "
                    "WHERE payroll_id=? AND status!='paid'",
                    (pid,)
                )
        except Exception as e:
            logger.error("خطأ في تحديث حالة الأقساط: %s", e, exc_info=True)

    # ==================== فلاتر الموظفين والسلف ====================
    def _refresh_employee_filters(self):
        current_loan_emp = self.loan_emp_filter.currentData() if hasattr(self, 'loan_emp_filter') else None
        current_inst_emp = self.inst_emp_filter.currentData() if hasattr(self, 'inst_emp_filter') else None

        self.loan_emp_filter.clear()
        self.loan_emp_filter.addItem("جميع الموظفين", None)
        self.inst_emp_filter.clear()
        self.inst_emp_filter.addItem("جميع الموظفين", None)

        emps = self.db.fetch_all(
            "SELECT id, first_name||' '||last_name FROM employees "
            "WHERE status='نشط' ORDER BY first_name"
        )
        for eid, name in emps:
            self.loan_emp_filter.addItem(name, eid)
            self.inst_emp_filter.addItem(name, eid)

        if current_loan_emp:
            idx = self.loan_emp_filter.findData(current_loan_emp)
            if idx >= 0:
                self.loan_emp_filter.setCurrentIndex(idx)
        if current_inst_emp:
            idx = self.inst_emp_filter.findData(current_inst_emp)
            if idx >= 0:
                self.inst_emp_filter.setCurrentIndex(idx)

        self._refresh_loan_filter()

    def _refresh_loan_filter(self):
        current_loan = self.inst_loan_filter.currentData() if hasattr(self, 'inst_loan_filter') else None
        self.inst_loan_filter.clear()
        self.inst_loan_filter.addItem("جميع السلف", None)

        emp_id = self.inst_emp_filter.currentData()
        try:
            if emp_id:
                loans = self.db.fetch_all(
                    "SELECT id, loan_type, amount FROM loans "
                    "WHERE employee_id=? ORDER BY start_year DESC, start_month DESC",
                    (emp_id,)
                )
            else:
                loans = self.db.fetch_all(
                    "SELECT id, loan_type, amount FROM loans "
                    "ORDER BY start_year DESC, start_month DESC"
                )
            for lid, ltype, amount in loans:
                self.inst_loan_filter.addItem(f"{ltype} - {amount:,.0f}", lid)
            if current_loan:
                idx = self.inst_loan_filter.findData(current_loan)
                if idx >= 0:
                    self.inst_loan_filter.setCurrentIndex(idx)
        except Exception as e:
            logger.error("خطأ في تحميل السلف: %s", e, exc_info=True)

    def _refresh_year_filter(self):
        current_year = self.inst_year_filter.currentData() if hasattr(self, 'inst_year_filter') else None
        self.inst_year_filter.clear()
        self.inst_year_filter.addItem("جميع السنوات", None)

        years = self.db.fetch_all(
            "SELECT DISTINCT strftime('%Y', due_date) FROM installments ORDER BY 1 DESC"
        )
        for (y,) in years:
            if y:
                self.inst_year_filter.addItem(y, int(y))
        if current_year:
            idx = self.inst_year_filter.findData(current_year)
            if idx >= 0:
                self.inst_year_filter.setCurrentIndex(idx)

    # ==================== بناء الواجهة ====================
    def _build(self):
        main_layout = QVBoxLayout(self)

        # ---- جدول السلف ----
        loans_group  = QGroupBox("💳 قائمة السلف")
        loans_layout = QVBoxLayout()

        loans_tools = QHBoxLayout()
        self.btn_new          = btn("➕ سلفة جديدة",    BTN_SUCCESS, self._new_loan)
        self.btn_edit_loan    = btn("✏️ تعديل سلفة",    BTN_PRIMARY, self._edit_loan)
        self.btn_delete_loan  = btn("🗑️ حذف سلفة",     BTN_DANGER,  self._delete_loan)
        self.btn_refresh_loans= btn("🔄 تحديث",          BTN_GRAY,    self._refresh_all)
        self.btn_export_loans = btn("📥 تصدير السلف",   BTN_PURPLE,  self._export_loans_excel)
        for b in [self.btn_new, self.btn_edit_loan, self.btn_delete_loan,
                  self.btn_refresh_loans, self.btn_export_loans]:
            loans_tools.addWidget(b)
        loans_tools.addStretch()
        loans_layout.addLayout(loans_tools)

        loans_filter = QHBoxLayout()
        loans_filter.addWidget(QLabel("الحالة:"))
        self.status_filter = QComboBox()
        self.status_filter.addItems(["جميع السلف", "نشط", "مكتمل", "ملغي"])
        self.status_filter.currentIndexChanged.connect(self._load_loans)
        loans_filter.addWidget(self.status_filter)

        loans_filter.addWidget(QLabel("الموظف:"))
        self.loan_emp_filter = QComboBox()
        self.loan_emp_filter.setMinimumWidth(200)
        self.loan_emp_filter.setEditable(True)
        self.loan_emp_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.loan_emp_filter.completer().setFilterMode(Qt.MatchContains)
        self.loan_emp_filter.currentIndexChanged.connect(self._load_loans)
        loans_filter.addWidget(self.loan_emp_filter)

        loans_filter.addWidget(QLabel("من تاريخ:"))
        self.loan_date_from = QDateEdit()
        self.loan_date_from.setCalendarPopup(True)
        self.loan_date_from.setDate(QDate.currentDate().addMonths(-12))
        self.loan_date_from.dateChanged.connect(self._load_loans)
        loans_filter.addWidget(self.loan_date_from)

        loans_filter.addWidget(QLabel("إلى تاريخ:"))
        self.loan_date_to = QDateEdit()
        self.loan_date_to.setCalendarPopup(True)
        self.loan_date_to.setDate(QDate.currentDate())
        self.loan_date_to.dateChanged.connect(self._load_loans)
        loans_filter.addWidget(self.loan_date_to)
        loans_filter.addStretch()
        loans_layout.addLayout(loans_filter)

        self.loans_table = make_table([
            "id", "الموظف", "النوع", "المبلغ", "القسط الشهري",
            "المتبقي", "شهر البداية", "إجمالي الأقساط", "المسددة", "المتبقية",
            "طريقة الدفع", "الحالة"
        ])
        self.loans_table.setColumnHidden(0, True)
        self.loans_table.setSortingEnabled(False)
        self.loans_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.loans_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.loans_table.itemSelectionChanged.connect(self._on_loan_selected)
        loans_layout.addWidget(self.loans_table)
        loans_group.setLayout(loans_layout)
        main_layout.addWidget(loans_group)

        # ---- جدول الأقساط ----
        inst_group  = QGroupBox("📊 جميع الأقساط")
        inst_layout = QVBoxLayout()

        inst_tools = QHBoxLayout()
        self.btn_pay              = btn("💰 تسديد قسط",    BTN_WARNING, self._pay_installment)
        self.btn_edit_installment = btn("✏️ تعديل قسط",    BTN_PRIMARY, self._edit_installment)
        self.btn_export_inst      = btn("📥 تصدير Excel",  BTN_PURPLE,  self._export_installments_excel)
        self.btn_refresh_inst     = btn("🔄 تحديث",         BTN_GRAY,    self._load_installments)
        self.btn_clear_filters    = btn("🗑️ مسح الفلاتر", BTN_GRAY,    self._clear_installment_filters)
        for b in [self.btn_pay, self.btn_edit_installment, self.btn_export_inst,
                  self.btn_refresh_inst, self.btn_clear_filters]:
            inst_tools.addWidget(b)
        inst_tools.addStretch()
        inst_layout.addLayout(inst_tools)

        inst_filter = QHBoxLayout()
        inst_filter.addWidget(QLabel("الموظف:"))
        self.inst_emp_filter = QComboBox()
        self.inst_emp_filter.setMinimumWidth(150)
        self.inst_emp_filter.setEditable(True)
        self.inst_emp_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.inst_emp_filter.completer().setFilterMode(Qt.MatchContains)
        self.inst_emp_filter.currentIndexChanged.connect(self._on_inst_emp_changed)
        inst_filter.addWidget(self.inst_emp_filter)

        inst_filter.addWidget(QLabel("السلفة:"))
        self.inst_loan_filter = QComboBox()
        self.inst_loan_filter.setMinimumWidth(150)
        self.inst_loan_filter.currentIndexChanged.connect(self._load_installments)
        inst_filter.addWidget(self.inst_loan_filter)

        inst_filter.addWidget(QLabel("السنة:"))
        self.inst_year_filter = QComboBox()
        self.inst_year_filter.setMinimumWidth(80)
        self.inst_year_filter.currentIndexChanged.connect(self._load_installments)
        inst_filter.addWidget(self.inst_year_filter)

        inst_filter.addWidget(QLabel("الشهر:"))
        self.inst_month_filter = QComboBox()
        self.inst_month_filter.addItem("جميع الأشهر", None)
        months_ar = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                     'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
        for i, m in enumerate(months_ar, 1):
            self.inst_month_filter.addItem(m, i)
        self.inst_month_filter.currentIndexChanged.connect(self._load_installments)
        inst_filter.addWidget(self.inst_month_filter)

        inst_filter.addWidget(QLabel("الحالة:"))
        self.status_group = QButtonGroup(self)
        self.status_group.setExclusive(False)
        self.chk_pending = QCheckBox("غير مسدد")
        self.chk_partial = QCheckBox("مسدد جزئياً")
        self.chk_paid    = QCheckBox("مسدد")
        for chk in [self.chk_pending, self.chk_partial, self.chk_paid]:
            self.status_group.addButton(chk)
            chk.stateChanged.connect(self._load_installments)
            inst_filter.addWidget(chk)

        inst_filter.addWidget(QLabel("من تاريخ:"))
        self.inst_due_from = QDateEdit()
        self.inst_due_from.setCalendarPopup(True)
        self.inst_due_from.setDate(QDate(2020, 1, 1))
        self.inst_due_from.dateChanged.connect(self._load_installments)
        inst_filter.addWidget(self.inst_due_from)

        inst_filter.addWidget(QLabel("إلى:"))
        self.inst_due_to = QDateEdit()
        self.inst_due_to.setCalendarPopup(True)
        self.inst_due_to.setDate(QDate.currentDate().addYears(5))
        self.inst_due_to.dateChanged.connect(self._load_installments)
        inst_filter.addWidget(self.inst_due_to)
        inst_filter.addStretch()
        inst_layout.addLayout(inst_filter)

        self.inst_table = make_table([
            "id", "الموظف", "السلفة", "شهر الاستحقاق", "المبلغ الأصلي",
            "المبلغ المدفوع", "الرصيد المتبقي", "تاريخ الدفع",
            "الحالة", "مصدر الدفع", "ملاحظات"
        ])
        self.inst_table.setColumnHidden(0, True)
        self.inst_table.setSortingEnabled(True)
        self.inst_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.inst_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.inst_table.setSelectionMode(QAbstractItemView.MultiSelection)
        inst_layout.addWidget(self.inst_table)

        inst_group.setLayout(inst_layout)
        main_layout.addWidget(inst_group)

        self._apply_permissions()
        self._refresh_year_filter()

    def _on_inst_emp_changed(self):
        self._refresh_loan_filter()
        self._load_installments()

    def _refresh_all(self):
        self._refresh_employee_filters()
        self._refresh_year_filter()
        self._load_loans()
        self._load_installments()

    def _apply_permissions(self):
        role = self.user['role']
        self.btn_new.setVisible(can_add(role))
        self.btn_edit_loan.setVisible(can_edit(role))
        self.btn_delete_loan.setVisible(can_delete(role))
        self.btn_pay.setVisible(can_edit(role))
        self.btn_edit_installment.setVisible(can_edit(role))
        self.btn_export_inst.setVisible(True)
        self.btn_export_loans.setVisible(True)

    # ==================== تحميل البيانات ====================
    def _load_loans(self):
        try:
            sf        = self.status_filter.currentText()
            emp_id    = self.loan_emp_filter.currentData()
            date_from = self.loan_date_from.date().toPyDate()
            date_to   = self.loan_date_to.date().toPyDate()

            q = """
                SELECT l.id,
                       e.first_name||' '||e.last_name,
                       l.loan_type,
                       l.amount,
                       l.monthly_installment,
                       l.remaining_amount,
                       l.start_month||'/'||l.start_year,
                       l.total_installments,
                       COALESCE((SELECT COUNT(*) FROM installments
                                 WHERE loan_id=l.id AND status='paid'), 0),
                       l.total_installments - COALESCE(
                           (SELECT COUNT(*) FROM installments
                            WHERE loan_id=l.id AND status='paid'), 0),
                       CASE WHEN l.payment_method='bank' THEN 'بنكي' ELSE 'نقدي' END,
                       l.status
                FROM loans l
                JOIN employees e ON l.employee_id = e.id
                WHERE 1=1
            """
            params = []

            if sf != "جميع السلف":
                q += " AND l.status=?"
                params.append(sf)
            if emp_id:
                q += " AND l.employee_id=?"
                params.append(emp_id)

            q += """
                AND (
                    (l.start_year > ?) OR
                    (l.start_year = ? AND l.start_month >= ?)
                ) AND (
                    (l.start_year < ?) OR
                    (l.start_year = ? AND l.start_month <= ?)
                )
            """
            params.extend([
                date_from.year, date_from.year, date_from.month,
                date_to.year,   date_to.year,   date_to.month,
            ])
            q += " ORDER BY l.created_at DESC"

            data = self.db.fetch_all(q, params) or []
            fill_table(
                self.loans_table,
                [list(r) for r in data],
                colors={11: lambda v: (
                    "#388E3C" if v == "مكتمل" else
                    "#D32F2F" if v == "ملغي" else
                    "#F57C00"
                )}
            )
            self._refresh_loan_filter()

        except Exception as e:
            logger.error("خطأ في تحميل السلف: %s", e, exc_info=True)
            QMessageBox.warning(self, "خطأ", f"حدث خطأ أثناء تحميل السلف:\n{str(e)}")

    def _on_loan_selected(self):
        selected = self.loans_table.selectedItems()
        if not selected:
            return
        item = self.loans_table.item(selected[0].row(), 0)
        if not item or not item.text():
            return
        self.current_loan_id = int(item.text())

        emp = self.db.fetch_one(
            "SELECT employee_id FROM loans WHERE id=?", (self.current_loan_id,))
        if emp:
            emp_id = emp[0]
            idx = self.inst_emp_filter.findData(emp_id)
            if idx >= 0:
                self.inst_emp_filter.setCurrentIndex(idx)
            self._refresh_loan_filter()
            idx2 = self.inst_loan_filter.findData(self.current_loan_id)
            if idx2 >= 0:
                self.inst_loan_filter.setCurrentIndex(idx2)

    def _load_installments(self):
        try:
            emp_id   = self.inst_emp_filter.currentData()
            loan_id  = self.inst_loan_filter.currentData()
            year     = self.inst_year_filter.currentData()
            month    = self.inst_month_filter.currentData()
            due_from = self.inst_due_from.date().toPyDate()
            due_to   = self.inst_due_to.date().toPyDate()

            status_conds = []
            if self.chk_pending.isChecked(): status_conds.append("i.status='pending'")
            if self.chk_partial.isChecked(): status_conds.append("i.status='partial'")
            if self.chk_paid.isChecked():    status_conds.append("i.status='paid'")
            status_sql = (" AND (" + " OR ".join(status_conds) + ")"
                          if status_conds else "")

            q = f"""
                SELECT i.id,
                       e.first_name||' '||e.last_name,
                       l.loan_type||' - '||printf('%d', l.amount)||' '||
                           COALESCE((SELECT setting_value FROM settings
                                     WHERE setting_name='currency'), 'ريال'),
                       i.due_date,
                       i.amount,
                       i.paid_amount,
                       i.paid_date,
                       CASE i.status
                           WHEN 'pending' THEN 'غير مسدد'
                           WHEN 'partial' THEN 'مسدد جزئياً'
                           WHEN 'paid'    THEN 'مسدد'
                       END,
                       CASE
                           WHEN i.payroll_id IS NOT NULL THEN
                               CASE WHEN p.status='معتمد' THEN 'راتب معتمد'
                                    ELSE 'راتب مسودة' END
                           ELSE 'يدوي'
                       END,
                       i.notes
                FROM installments i
                JOIN loans l     ON i.loan_id    = l.id
                JOIN employees e ON l.employee_id = e.id
                LEFT JOIN payroll p ON i.payroll_id = p.id
                WHERE i.due_date BETWEEN ? AND ?
                {status_sql}
            """
            params = [due_from.isoformat(), due_to.isoformat()]

            if emp_id:
                q += " AND l.employee_id=?"
                params.append(emp_id)
            if loan_id:
                q += " AND i.loan_id=?"
                params.append(loan_id)
            if year:
                q += " AND strftime('%Y', i.due_date)=?"
                params.append(str(year))
            if month:
                q += " AND strftime('%m', i.due_date)=?"
                params.append(f"{month:02d}")

            q += " ORDER BY i.due_date DESC"

            data = self.db.fetch_all(q, params) or []

            display = []
            for row in data:
                try:
                    due = date.fromisoformat(row[3])
                    month_year = due.strftime("%m/%Y")
                except Exception:
                    month_year = row[3]

                amount    = row[4] or 0
                paid      = row[5] or 0
                remaining = amount - paid
                display.append([
                    row[0],
                    row[1] or "غير معروف",
                    row[2] or "غير معروف",
                    month_year,
                    f"{amount:,.2f}",
                    f"{paid:,.2f}",
                    f"{remaining:,.2f}",
                    row[6] or "--",
                    row[7] or "غير معروف",
                    row[8],
                    row[9] or "",
                ])

            self.inst_table.setRowCount(0)
            for ri, row_data in enumerate(display):
                self.inst_table.insertRow(ri)
                for ci, val in enumerate(row_data):
                    item = QTableWidgetItem(str(val))
                    item.setTextAlignment(Qt.AlignCenter)
                    if ci == 8:   # عمود الحالة
                        clr = ("#388E3C" if val == "مسدد" else
                               "#F57C00" if val == "مسدد جزئياً" else
                               "#D32F2F")
                        item.setForeground(QColor(clr))
                    self.inst_table.setItem(ri, ci, item)

        except Exception as e:
            logger.error("خطأ في تحميل الأقساط: %s", e, exc_info=True)
            QMessageBox.warning(self, "خطأ", f"حدث خطأ أثناء تحميل الأقساط:\n{str(e)}")

    def _clear_installment_filters(self):
        self.inst_emp_filter.setCurrentIndex(0)
        self.inst_loan_filter.setCurrentIndex(0)
        self.inst_year_filter.setCurrentIndex(0)
        self.inst_month_filter.setCurrentIndex(0)
        self.chk_pending.setChecked(False)
        self.chk_partial.setChecked(False)
        self.chk_paid.setChecked(False)
        self.inst_due_from.setDate(QDate(2020, 1, 1))
        self.inst_due_to.setDate(QDate.currentDate().addYears(5))
        self._load_installments()

    # ==================== إدارة السلف ====================
    def _new_loan(self):
        dlg = LoanDialog(self, self.db, self.comm)
        if dlg.exec_() == QDialog.Accepted:
            self._load_loans()
            self._load_installments()
            self._refresh_year_filter()
            if self.comm:
                self.comm.dataChanged.emit('loan', {'action': 'add'})

    def _edit_loan(self):
        if not self.current_loan_id:
            QMessageBox.warning(self, "تنبيه", "الرجاء تحديد سلفة أولاً")
            return
        paid_count = (self.db.fetch_one(
            "SELECT COUNT(*) FROM installments WHERE loan_id=? AND paid_amount>0",
            (self.current_loan_id,)) or [0])[0]
        if paid_count > 0:
            QMessageBox.warning(self, "تنبيه", "لا يمكن تعديل سلفة بعد بدء دفع أقساطها.")
            return
        dlg = EditLoanDialog(self, self.db, self.user, self.current_loan_id, self.comm)
        if dlg.exec_() == QDialog.Accepted:
            self._load_loans()
            self._load_installments()
            if self.comm:
                self.comm.dataChanged.emit('loan', {'action': 'edit', 'id': self.current_loan_id})

    def _delete_loan(self):
        if not self.current_loan_id:
            QMessageBox.warning(self, "تنبيه", "الرجاء تحديد سلفة أولاً")
            return

        loan_info = self.db.fetch_one(
            "SELECT e.first_name||' '||e.last_name, l.loan_type "
            "FROM loans l JOIN employees e ON l.employee_id=e.id WHERE l.id=?",
            (self.current_loan_id,)
        )
        if not loan_info:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على السلفة")
            return
        emp_name, loan_type = loan_info

        # أقساط مرتبطة برواتب معتمدة
        approved = self.db.fetch_all(
            "SELECT DISTINCT p.id, p.month, p.year "
            "FROM installments i JOIN payroll p ON i.payroll_id=p.id "
            "WHERE i.loan_id=? AND p.status='معتمد'",
            (self.current_loan_id,)
        )
        if approved:
            msg = "هذه السلفة مرتبطة برواتب معتمدة. لا يمكن حذفها.\nالرواتب:"
            for _, m, y in approved:
                msg += f"\n- {m}/{y}"
            QMessageBox.warning(self, "تنبيه", msg)
            return

        # أقساط مرتبطة برواتب مسودة
        drafts = self.db.fetch_all(
            "SELECT DISTINCT p.id FROM installments i "
            "JOIN payroll p ON i.payroll_id=p.id "
            "WHERE i.loan_id=? AND p.status='مسودة'",
            (self.current_loan_id,)
        )
        if drafts:
            reply = QMessageBox.question(
                self, "تأكيد",
                f"هذه السلفة مرتبطة بـ {len(drafts)} راتب مسودة. "
                "سيتم حذف هذه الرواتب.\nمتابعة؟",
                QMessageBox.Yes | QMessageBox.No
            )
            if reply != QMessageBox.Yes:
                return

        try:
            conn = self.db.conn
            conn.execute("BEGIN TRANSACTION")
            for (pid,) in drafts:
                conn.execute("DELETE FROM payroll WHERE id=? AND status='مسودة'", (pid,))
            conn.execute(
                "UPDATE installments SET payroll_id=NULL WHERE loan_id=?",
                (self.current_loan_id,))
            conn.execute("DELETE FROM installments WHERE loan_id=?", (self.current_loan_id,))
            conn.execute("DELETE FROM loans WHERE id=?", (self.current_loan_id,))
            conn.commit()

            self.db.log_action("حذف سلفة", "loans", self.current_loan_id,
                               {"employee": emp_name, "type": loan_type}, None)
            QMessageBox.information(self, "نجاح", "تم حذف السلفة والبيانات المرتبطة بنجاح")
            self.current_loan_id = None
            self._load_loans()
            self._load_installments()
            if self.comm:
                self.comm.dataChanged.emit('loan', {'action': 'delete'})

        except Exception as e:
            self.db.conn.rollback()
            logger.error("خطأ في حذف السلفة: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء الحذف:\n{str(e)}")

    # ==================== تسديد وتعديل الأقساط ====================
    def _pay_installment(self):
        selected = self.inst_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "تنبيه", "الرجاء اختيار قسط من الجدول")
            return

        row    = selected[0].row()
        id_itm = self.inst_table.item(row, 0)
        if not id_itm:
            return
        inst_id = int(id_itm.text())

        inst = self.db.fetch_one("""
            SELECT i.id, i.loan_id, i.amount, i.paid_amount, i.status, i.payroll_id,
                   e.first_name||' '||e.last_name,
                   l.loan_type, l.remaining_amount, l.payment_method
            FROM installments i
            JOIN loans l     ON i.loan_id    = l.id
            JOIN employees e ON l.employee_id = e.id
            WHERE i.id=?
        """, (inst_id,))
        if not inst:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على القسط")
            return

        (inst_id, loan_id, amount, paid_amount, status,
         payroll_id, emp_name, loan_type, remaining, pay_method) = inst

        if payroll_id:
            p_status = self.db.fetch_one(
                "SELECT status FROM payroll WHERE id=?", (payroll_id,))
            if p_status and p_status[0] == 'معتمد':
                QMessageBox.warning(self, "تنبيه",
                                    "لا يمكن تعديل قسط مرتبط براتب معتمد.")
                return
            if QMessageBox.question(
                    self, "تأكيد",
                    "هذا القسط مرتبط براتب مسودة. هل تريد متابعة الدفع اليدوي؟",
                    QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                return

        currency         = self.db.get_setting('currency', 'ريال')
        remaining_amount = amount - (paid_amount or 0)
        dlg = PayInstallmentDialog(self, remaining_amount, currency, amount, paid_amount)
        if dlg.exec_() != QDialog.Accepted:
            return

        pay_amount = dlg.amount
        pay_date   = dlg.paid_date
        new_paid   = (paid_amount or 0) + pay_amount
        new_status = 'paid' if abs(new_paid - amount) < 0.01 else 'partial'

        try:
            conn = self.db.conn
            conn.execute("BEGIN TRANSACTION")

            conn.execute("""
                UPDATE installments
                SET paid_amount=?, status=?, paid_date=?, notes=?
                WHERE id=?
            """, (new_paid, new_status, pay_date,
                  f"دفع يدوي في {pay_date}", inst_id))

            new_loan_remaining = remaining - pay_amount
            conn.execute("UPDATE loans SET remaining_amount=? WHERE id=?",
                         (new_loan_remaining, loan_id))
            if new_loan_remaining <= 0:
                conn.execute("UPDATE loans SET status='مكتمل' WHERE id=?", (loan_id,))

            if payroll_id:
                total_ded = (conn.execute(
                    "SELECT SUM(paid_amount) FROM installments WHERE payroll_id=?",
                    (payroll_id,)).fetchone()[0] or 0)
                if pay_method == 'bank':
                    conn.execute(
                        "UPDATE payroll SET loan_deduction=?, loan_deduction_bank=?,"
                        " loan_deduction_cash=0 WHERE id=?",
                        (total_ded, total_ded, payroll_id))
                else:
                    conn.execute(
                        "UPDATE payroll SET loan_deduction=?, loan_deduction_bank=0,"
                        " loan_deduction_cash=? WHERE id=?",
                        (total_ded, total_ded, payroll_id))

            conn.commit()
            self.db.log_action("تسديد قسط يدوي", "installments", inst_id,
                               None, {"paid_amount": pay_amount, "status": new_status})
            QMessageBox.information(self, "نجاح", f"تم تسديد {pay_amount:.2f} بنجاح")
            self._load_loans()
            self._load_installments()
            if self.comm:
                self.comm.dataChanged.emit('loan', {'action': 'pay', 'id': loan_id})
                if payroll_id:
                    self.comm.dataChanged.emit('payroll', {'action': 'refresh', 'id': payroll_id})

        except Exception as e:
            self.db.conn.rollback()
            logger.error("خطأ في تسديد القسط: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء الحفظ:\n{str(e)}")

    def _edit_installment(self):
        selected = self.inst_table.selectedItems()
        if not selected:
            QMessageBox.warning(self, "تنبيه", "الرجاء تحديد قسط")
            return
        id_itm = self.inst_table.item(selected[0].row(), 0)
        if not id_itm:
            return
        inst_id = int(id_itm.text())

        p_info = self.db.fetch_one(
            "SELECT p.status FROM installments i "
            "JOIN payroll p ON i.payroll_id=p.id WHERE i.id=?",
            (inst_id,)
        )
        if p_info and p_info[0] == 'معتمد':
            QMessageBox.warning(self, "تنبيه",
                                "لا يمكن تعديل قسط مرتبط براتب معتمد.")
            return
        if p_info and p_info[0] == 'مسودة':
            if QMessageBox.question(
                    self, "تأكيد",
                    "هذا القسط مرتبط براتب مسودة. أي تعديل سيؤثر على الراتب.\n"
                    "هل تريد المتابعة؟",
                    QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                return

        dlg = EditInstallmentDialog(self, self.db, self.user, inst_id, self.comm)
        if dlg.exec_() == QDialog.Accepted:
            self._load_installments()
            self._load_loans()
            if self.comm:
                self.comm.dataChanged.emit('installment', {'action': 'edit', 'id': inst_id})

    # ==================== التصدير ====================
    def _export_installments_excel(self):
        try:
            import pandas as pd
            data = []
            for r in range(self.inst_table.rowCount()):
                data.append([
                    self.inst_table.item(r, c).text()
                    if self.inst_table.item(r, c) else ""
                    for c in range(1, self.inst_table.columnCount())
                ])
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return
            cols = ["الموظف", "السلفة", "شهر الاستحقاق", "المبلغ الأصلي",
                    "المبلغ المدفوع", "الرصيد المتبقي", "تاريخ الدفع",
                    "الحالة", "مصدر الدفع", "ملاحظات"]
            df   = pd.DataFrame(data, columns=cols)
            fname = f"installments_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            path, _ = QFileDialog.getSaveFileName(self, "حفظ", fname, "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", f"تم تصدير {len(data)} قسط بنجاح")
        except ImportError:
            QMessageBox.critical(self, "خطأ",
                                 "الرجاء تثبيت pandas و openpyxl:\npip install pandas openpyxl")
        except Exception as e:
            logger.error("خطأ في تصدير الأقساط: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))

    def _export_loans_excel(self):
        try:
            import pandas as pd
            sf        = self.status_filter.currentText()
            emp_id    = self.loan_emp_filter.currentData()
            date_from = self.loan_date_from.date().toPyDate()
            date_to   = self.loan_date_to.date().toPyDate()

            q = """
                SELECT e.employee_code, e.first_name||' '||e.last_name,
                       COALESCE(d.name,'') AS department,
                       l.loan_type, l.amount, l.monthly_installment,
                       l.remaining_amount, l.start_month||'/'||l.start_year,
                       l.total_installments,
                       COALESCE((SELECT COUNT(*) FROM installments
                                 WHERE loan_id=l.id AND status='paid'), 0),
                       l.total_installments - COALESCE(
                           (SELECT COUNT(*) FROM installments
                            WHERE loan_id=l.id AND status='paid'), 0),
                       CASE WHEN l.payment_method='bank' THEN 'بنكي' ELSE 'نقدي' END,
                       l.status, l.notes
                FROM loans l
                JOIN employees e ON l.employee_id = e.id
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE 1=1
            """
            params = []
            if sf != "جميع السلف":
                q += " AND l.status=?"; params.append(sf)
            if emp_id:
                q += " AND l.employee_id=?"; params.append(emp_id)
            q += """
                AND (
                    (l.start_year > ?) OR
                    (l.start_year = ? AND l.start_month >= ?)
                ) AND (
                    (l.start_year < ?) OR
                    (l.start_year = ? AND l.start_month <= ?)
                ) ORDER BY l.created_at DESC
            """
            params.extend([
                date_from.year, date_from.year, date_from.month,
                date_to.year,   date_to.year,   date_to.month,
            ])

            data = self.db.fetch_all(q, params)
            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return

            cols = ["الرقم الوظيفي", "الموظف", "القسم", "نوع السلفة",
                    "المبلغ", "القسط الشهري", "المتبقي", "شهر البداية",
                    "إجمالي الأقساط", "المسددة", "المتبقية",
                    "طريقة الدفع", "الحالة", "ملاحظات"]
            df    = pd.DataFrame(data, columns=cols)
            fname = f"loans_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            path, _ = QFileDialog.getSaveFileName(self, "حفظ", fname, "Excel (*.xlsx)")
            if path:
                df.to_excel(path, index=False)
                QMessageBox.information(self, "نجاح", f"تم تصدير {len(data)} سلفة بنجاح")

        except ImportError:
            QMessageBox.critical(self, "خطأ", "الرجاء تثبيت pandas و openpyxl")
        except Exception as e:
            logger.error("خطأ في تصدير السلف: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))


# ===================================================================
# نافذة إضافة سلفة جديدة
# ===================================================================
class LoanDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, comm=None):
        super().__init__(parent)
        self.db   = db
        self.comm = comm
        self.setWindowTitle("سلفة / قرض جديد")
        self.setFixedSize(450, 500)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()

    def _build(self):
        lay  = QVBoxLayout(self)
        form = QFormLayout()
        form.setSpacing(10)

        self.emp = QComboBox()
        for eid, name in self.db.fetch_all(
                "SELECT id, first_name||' '||last_name FROM employees "
                "WHERE status='نشط' ORDER BY first_name"):
            self.emp.addItem(name, eid)
        self.emp.setEditable(True)
        self.emp.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.emp.completer().setFilterMode(Qt.MatchContains)

        self.ltype = QComboBox()
        self.ltype.addItems(["سلفة", "قرض", "عهدة", "أخرى"])

        self.payment_method = QComboBox()
        self.payment_method.addItems(["نقدي", "بنكي"])

        currency = self.db.get_setting('currency', 'ريال')

        self.amount = QDoubleSpinBox()
        self.amount.setRange(1, 999999)
        self.amount.setSuffix(f" {currency}")
        self.amount.valueChanged.connect(self._calc)

        self.installment = QDoubleSpinBox()
        self.installment.setRange(1, 999999)
        self.installment.setSuffix(f" {currency}")
        self.installment.valueChanged.connect(self._calc)

        self.months_lbl = QLabel("0 شهر")
        self.months_lbl.setStyleSheet("font-weight:bold; color:#1976D2;")

        self.start_month = QComboBox()
        self.start_month.addItems(['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                                   'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'])
        self.start_month.setCurrentIndex(datetime.now().month - 1)

        self.start_year = QSpinBox()
        self.start_year.setRange(2020, 2050)
        self.start_year.setValue(datetime.now().year)

        self.notes = QLineEdit()

        form.addRow("الموظف:",        self.emp)
        form.addRow("نوع السلفة:",    self.ltype)
        form.addRow("طريقة الدفع:",  self.payment_method)
        form.addRow("إجمالي المبلغ:", self.amount)
        form.addRow("القسط الشهري:", self.installment)
        form.addRow("عدد الأشهر:",   self.months_lbl)
        form.addRow("شهر البداية:",  self.start_month)
        form.addRow("سنة البداية:",  self.start_year)
        form.addRow("ملاحظات:",      self.notes)
        lay.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _calc(self):
        if self.installment.value() > 0:
            months = math.ceil(self.amount.value() / self.installment.value())
            self.months_lbl.setText(f"{months} شهر")
        else:
            self.months_lbl.setText("0 شهر")

    def _save(self):
        if self.amount.value() <= 0 or self.installment.value() <= 0:
            QMessageBox.warning(self, "خطأ", "المبلغ والقسط يجب أن يكونا أكبر من صفر")
            return

        months = math.ceil(self.amount.value() / self.installment.value())
        pm     = 'bank' if self.payment_method.currentText() == "بنكي" else 'cash'

        ok = self.db.execute_query("""
            INSERT INTO loans (
                employee_id, loan_type, amount, monthly_installment,
                remaining_amount, start_month, start_year, total_installments,
                payment_method, notes, status
            ) VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """, (self.emp.currentData(), self.ltype.currentText(),
              self.amount.value(), self.installment.value(),
              self.amount.value(),
              self.start_month.currentIndex() + 1, self.start_year.value(),
              months, pm, self.notes.text().strip() or None, 'نشط'))

        if not ok:
            QMessageBox.critical(self, "خطأ", "فشل في إضافة السلفة")
            return

        loan_id = self.db.last_id()
        amt     = self.installment.value()
        last    = self.amount.value() - (amt * (months - 1))
        sm      = self.start_month.currentIndex() + 1
        sy      = self.start_year.value()

        for i in range(months):
            y = sy + (sm + i - 1) // 12
            m = (sm + i - 1) % 12 + 1
            self.db.execute_query(
                "INSERT INTO installments (loan_id, due_date, amount, status) VALUES (?,?,?,?)",
                (loan_id, date(y, m, 1).isoformat(),
                 amt if i < months - 1 else last, 'pending'))

        self.db.log_action("إضافة سلفة", "loans", loan_id, None,
                           {"employee_id": self.emp.currentData(),
                            "amount": self.amount.value(), "installments": months})
        QMessageBox.information(self, "نجاح", f"تم إضافة السلفة بنجاح بعدد {months} قسط")
        self.accept()


# ===================================================================
# نافذة تعديل السلفة
# ===================================================================
class EditLoanDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, user, loan_id, comm=None):
        super().__init__(parent)
        self.db      = db
        self.user    = user
        self.loan_id = loan_id
        self.comm    = comm
        self.setWindowTitle("تعديل السلفة")
        self.setFixedSize(450, 500)
        self.setLayoutDirection(Qt.RightToLeft)
        self._load_data()
        self._build()

    def _load_data(self):
        data = self.db.fetch_one("""
            SELECT employee_id, loan_type, amount, monthly_installment,
                   remaining_amount, start_month, start_year, total_installments,
                   payment_method, notes
            FROM loans WHERE id=?
        """, (self.loan_id,))
        if not data:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على السلفة")
            self.reject()
            return
        (self.emp_id, self.loan_type, self.amount_val, self.monthly_val,
         self.remaining, self.start_m, self.start_y,
         self.total_inst, self.pay_method, self.notes_val) = data

    def _build(self):
        lay  = QVBoxLayout(self)
        form = QFormLayout()
        form.setSpacing(10)

        emp_name = self.db.fetch_one(
            "SELECT first_name||' '||last_name FROM employees WHERE id=?",
            (self.emp_id,))[0]
        form.addRow("الموظف:", QLabel(emp_name))

        self.ltype = QComboBox()
        self.ltype.addItems(["سلفة", "قرض", "عهدة", "أخرى"])
        self.ltype.setCurrentText(self.loan_type)
        form.addRow("نوع السلفة:", self.ltype)

        self.payment_method = QComboBox()
        self.payment_method.addItems(["نقدي", "بنكي"])
        self.payment_method.setCurrentText("بنكي" if self.pay_method == 'bank' else "نقدي")
        form.addRow("طريقة الدفع:", self.payment_method)

        currency = self.db.get_setting('currency', 'ريال')

        self.amount = QDoubleSpinBox()
        self.amount.setRange(1, 999999)
        self.amount.setValue(self.amount_val)
        self.amount.setSuffix(f" {currency}")
        self.amount.valueChanged.connect(self._calc)
        form.addRow("إجمالي المبلغ:", self.amount)

        self.installment = QDoubleSpinBox()
        self.installment.setRange(1, 999999)
        self.installment.setValue(self.monthly_val)
        self.installment.setSuffix(f" {currency}")
        self.installment.valueChanged.connect(self._calc)
        form.addRow("القسط الشهري:", self.installment)

        self.months_lbl = QLabel(f"{self.total_inst} شهر")
        self.months_lbl.setStyleSheet("font-weight:bold; color:#1976D2;")
        form.addRow("عدد الأشهر:", self.months_lbl)

        self.start_month = QComboBox()
        self.start_month.addItems(['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                                   'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر'])
        self.start_month.setCurrentIndex(self.start_m - 1)
        form.addRow("شهر البداية:", self.start_month)

        self.start_year = QSpinBox()
        self.start_year.setRange(2020, 2050)
        self.start_year.setValue(self.start_y)
        form.addRow("سنة البداية:", self.start_year)

        self.notes = QLineEdit()
        self.notes.setText(self.notes_val or "")
        form.addRow("ملاحظات:", self.notes)

        lay.addLayout(form)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _calc(self):
        if self.installment.value() > 0:
            months = math.ceil(self.amount.value() / self.installment.value())
            self.months_lbl.setText(f"{months} شهر")
        else:
            self.months_lbl.setText("0 شهر")

    def _save(self):
        new_amount = self.amount.value()
        new_inst   = self.installment.value()
        if new_amount <= 0 or new_inst <= 0:
            QMessageBox.warning(self, "خطأ", "المبلغ والقسط يجب أن يكونا أكبر من صفر")
            return

        new_months = math.ceil(new_amount / new_inst)
        pm = 'bank' if self.payment_method.currentText() == "بنكي" else 'cash'

        self.db.execute_query("""
            UPDATE loans SET loan_type=?, amount=?, monthly_installment=?,
                remaining_amount=?, start_month=?, start_year=?,
                total_installments=?, payment_method=?, notes=?
            WHERE id=?
        """, (self.ltype.currentText(), new_amount, new_inst, new_amount,
              self.start_month.currentIndex() + 1, self.start_year.value(),
              new_months, pm, self.notes.text().strip() or None, self.loan_id))

        self.db.execute_query(
            "DELETE FROM installments WHERE loan_id=?", (self.loan_id,))

        sm   = self.start_month.currentIndex() + 1
        sy   = self.start_year.value()
        last = new_amount - (new_inst * (new_months - 1))

        for i in range(new_months):
            y = sy + (sm + i - 1) // 12
            m = (sm + i - 1) % 12 + 1
            self.db.execute_query(
                "INSERT INTO installments (loan_id, due_date, amount, status) VALUES (?,?,?,?)",
                (self.loan_id, date(y, m, 1).isoformat(),
                 new_inst if i < new_months - 1 else last, 'pending'))

        self.db.log_action("تعديل سلفة", "loans", self.loan_id, None,
                           {"amount": new_amount, "installments": new_months})
        QMessageBox.information(self, "نجاح", "تم تعديل السلفة وإعادة إنشاء الأقساط")
        self.accept()


# ===================================================================
# نافذة تسديد قسط
# ===================================================================
class PayInstallmentDialog(QDialog):
    def __init__(self, parent, max_amount, currency, original_amount=0, paid_so_far=0):
        super().__init__(parent)
        self.setWindowTitle("تسديد قسط")
        self.setFixedSize(400, 250)
        self.setLayoutDirection(Qt.RightToLeft)
        self.amount    = 0
        self.paid_date = None

        layout = QVBoxLayout(self)
        info   = QLabel(
            f"المبلغ الأصلي للقسط: {original_amount:.2f} {currency}\n"
            f"المدفوع سابقاً: {paid_so_far:.2f} {currency}")
        info.setStyleSheet("color:#1976D2; font-weight:bold;")
        layout.addWidget(info)

        form = QFormLayout()
        self.amount_spin = QDoubleSpinBox()
        self.amount_spin.setRange(0.01, 999999)
        self.amount_spin.setSuffix(f" {currency}")
        self.amount_spin.setValue(max_amount if max_amount > 0 else original_amount - paid_so_far)
        form.addRow("المبلغ المراد دفعه:", self.amount_spin)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        form.addRow("تاريخ الدفع:", self.date_edit)
        layout.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def _accept(self):
        self.amount    = self.amount_spin.value()
        self.paid_date = self.date_edit.date().toString(Qt.ISODate)
        self.accept()


# ===================================================================
# نافذة تعديل القسط
# ===================================================================
class EditInstallmentDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, user, installment_id, comm=None):
        super().__init__(parent)
        self.db             = db
        self.user           = user
        self.installment_id = installment_id
        self.comm           = comm
        self.setWindowTitle("تعديل القسط")
        self.setFixedSize(450, 500)
        self.setLayoutDirection(Qt.RightToLeft)
        self._load_data()
        self._build()

    def _load_data(self):
        data = self.db.fetch_one("""
            SELECT i.amount, i.paid_amount, i.status, i.paid_date, i.loan_id, i.notes,
                   l.loan_type, e.first_name||' '||e.last_name,
                   l.amount, l.remaining_amount,
                   i.payroll_id, l.payment_method
            FROM installments i
            JOIN loans l     ON i.loan_id    = l.id
            JOIN employees e ON l.employee_id = e.id
            WHERE i.id=?
        """, (self.installment_id,))
        if not data:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على القسط")
            self.reject()
            return
        (self.inst_amount, self.paid_amount, self.status, self.paid_date,
         self.loan_id, self.notes, self.loan_type, self.emp_name,
         self.loan_total, self.loan_remaining,
         self.payroll_id, self.loan_pay_method) = data

    def _build(self):
        layout = QVBoxLayout(self)
        form   = QFormLayout()

        form.addRow("الموظف:",                 QLabel(self.emp_name))
        form.addRow("السلفة:",                  QLabel(self.loan_type))
        form.addRow("إجمالي السلفة:",           QLabel(f"{self.loan_total:.2f}"))
        form.addRow("المبلغ الأصلي لهذا القسط:", QLabel(f"{self.inst_amount:.2f}"))

        if self.payroll_id:
            p = self.db.fetch_one("SELECT status FROM payroll WHERE id=?", (self.payroll_id,))
            source = ("راتب معتمد" if p and p[0] == 'معتمد' else "راتب مسودة") if p else "يدوي"
        else:
            source = "يدوي"
        form.addRow("مصدر الدفع:", QLabel(source))

        self.paid_spin = QDoubleSpinBox()
        self.paid_spin.setRange(0, self.inst_amount * 10)
        self.paid_spin.setValue(self.paid_amount or 0)
        self.paid_spin.valueChanged.connect(self._on_paid_changed)
        form.addRow("المبلغ المدفوع (إجمالي):", self.paid_spin)

        self.remaining_lbl = QLabel(f"{self.inst_amount - (self.paid_amount or 0):.2f}")
        self.remaining_lbl.setStyleSheet("font-weight:bold; color:#1976D2;")
        form.addRow("الرصيد المتبقي:", self.remaining_lbl)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(
            QDate.fromString(self.paid_date, Qt.ISODate)
            if self.paid_date else QDate.currentDate()
        )
        form.addRow("تاريخ الدفع:", self.date_edit)

        self.status_combo = QComboBox()
        self.status_combo.addItems(["غير مسدد", "مسدد جزئياً", "مسدد"])
        self.status_combo.setCurrentIndex({'pending': 0, 'partial': 1, 'paid': 2}
                                          .get(self.status, 0))
        self.status_combo.currentIndexChanged.connect(self._on_status_changed)
        form.addRow("الحالة:", self.status_combo)

        self.notes_edit = QTextEdit()
        self.notes_edit.setText(self.notes or "")
        self.notes_edit.setMaximumHeight(80)
        form.addRow("ملاحظات:", self.notes_edit)

        self.clear_date_check = QCheckBox("مسح تاريخ الدفع")
        form.addRow("", self.clear_date_check)

        layout.addLayout(form)

        self.info_lbl = QLabel("")
        self.info_lbl.setStyleSheet("color:#1976D2; font-weight:bold;")
        layout.addWidget(self.info_lbl)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)
        self._update_info()

    def _on_paid_changed(self):
        new_paid  = self.paid_spin.value()
        remaining = max(0, self.inst_amount - new_paid)
        self.remaining_lbl.setText(f"{remaining:.2f}")
        self._update_info()
        if remaining <= 0:
            self.status_combo.setCurrentIndex(2)
        elif new_paid > 0:
            self.status_combo.setCurrentIndex(1)
        else:
            self.status_combo.setCurrentIndex(0)

    def _on_status_changed(self):
        idx = self.status_combo.currentIndex()
        if idx == 2: self.paid_spin.setValue(self.inst_amount)
        elif idx == 0: self.paid_spin.setValue(0)
        self._update_info()

    def _update_info(self):
        diff = self.paid_spin.value() - (self.paid_amount or 0)
        if abs(diff) < 0.01:
            self.info_lbl.setText("لا يوجد تغيير في المبلغ المدفوع")
        elif diff > 0:
            self.info_lbl.setText(f"سيتم زيادة المبلغ المدفوع بمقدار {diff:.2f}")
        else:
            self.info_lbl.setText(f"سيتم تخفيض المبلغ المدفوع بمقدار {-diff:.2f}")

    def _save(self):
        if self.payroll_id:
            p = self.db.fetch_one("SELECT status FROM payroll WHERE id=?", (self.payroll_id,))
            if p and p[0] == 'معتمد':
                QMessageBox.warning(self, "تحذير",
                                    "لا يمكن تعديل قسط مرتبط براتب معتمد.")
                self.reject()
                return

        new_paid   = self.paid_spin.value()
        new_date   = self.date_edit.date().toString(Qt.ISODate)
        new_status = ('paid'    if new_paid >= self.inst_amount - 0.01 else
                      'partial' if new_paid > 0 else 'pending')

        if new_status in ('pending', 'partial') and (
                new_paid == 0 or self.clear_date_check.isChecked()):
            new_date = None

        new_notes = self.notes_edit.toPlainText().strip()

        try:
            conn = self.db.conn
            conn.execute("BEGIN TRANSACTION")

            conn.execute("""
                UPDATE installments
                SET paid_amount=?, paid_date=?, status=?, notes=?
                WHERE id=?
            """, (new_paid, new_date, new_status, new_notes, self.installment_id))

            total_paid   = (conn.execute(
                "SELECT SUM(paid_amount) FROM installments WHERE loan_id=?",
                (self.loan_id,)).fetchone()[0] or 0)
            new_remaining = max(0, self.loan_total - total_paid)
            conn.execute("UPDATE loans SET remaining_amount=? WHERE id=?",
                         (new_remaining, self.loan_id))
            if new_remaining <= 0:
                conn.execute("UPDATE loans SET status='مكتمل' WHERE id=?", (self.loan_id,))
            else:
                ls = conn.execute("SELECT status FROM loans WHERE id=?",
                                  (self.loan_id,)).fetchone()
                if ls and ls[0] == 'مكتمل':
                    conn.execute("UPDATE loans SET status='نشط' WHERE id=?", (self.loan_id,))

            if self.payroll_id:
                total_ded = (conn.execute(
                    "SELECT SUM(paid_amount) FROM installments WHERE payroll_id=?",
                    (self.payroll_id,)).fetchone()[0] or 0)
                if self.loan_pay_method == 'bank':
                    conn.execute("""
                        UPDATE payroll
                        SET loan_deduction=?, loan_deduction_bank=?, loan_deduction_cash=0
                        WHERE id=?
                    """, (total_ded, total_ded, self.payroll_id))
                else:
                    conn.execute("""
                        UPDATE payroll
                        SET loan_deduction=?, loan_deduction_bank=0, loan_deduction_cash=?
                        WHERE id=?
                    """, (total_ded, total_ded, self.payroll_id))

            conn.commit()
            self.db.log_action("تعديل قسط", "installments", self.installment_id,
                               None, {"paid_amount": new_paid, "status": new_status})
            QMessageBox.information(self, "نجاح", "تم تعديل القسط وتحديث الرصيد")
            if self.comm and self.payroll_id:
                self.comm.dataChanged.emit('payroll', {'action': 'refresh', 'id': self.payroll_id})
                self.comm.dataChanged.emit('installment',
                                           {'action': 'edit', 'id': self.installment_id,
                                            'payroll_id': self.payroll_id})
            self.accept()

        except Exception as e:
            self.db.conn.rollback()
            logger.error("خطأ في تعديل القسط: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء الحفظ:\n{str(e)}")
