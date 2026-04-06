#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/attendance_tab.py

import os
import csv
import sqlite3
import logging
from datetime import datetime, date, timedelta

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QDateEdit,
    QPushButton, QMessageBox, QFileDialog, QDialog, QTextEdit, QGroupBox,
    QFormLayout, QSpinBox, QLineEdit, QProgressBar, QTableWidgetItem,
    QDialogButtonBox, QGridLayout, QTimeEdit, QAbstractItemView, QCompleter,
    QCheckBox, QScrollArea, QApplication, QFrame, QTabWidget, QProgressDialog
)
from PyQt5.QtCore import Qt, QDate, QTime, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont, QColor, QBrush

from database import DatabaseManager
from utils import make_table, fill_table, btn, can_add, can_edit, can_delete, can_approve
from constants import (BTN_SUCCESS, BTN_PRIMARY, BTN_WARNING,
                       BTN_TEAL, BTN_GRAY, BTN_PURPLE, BTN_DANGER)

logger = logging.getLogger(__name__)

def safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default

class AttendanceTab(QWidget):
    """
    تبويب الحضور — مُقسَّم إلى تبويبين فرعيين:

    1. مسودة الحضور  (is_approved = 0):
       - قابل للتعديل والحذف والإدخال اليدوي.
       - زر "اعتماد الفترة" يعتمد جميع سجلات المسودة لفترة محددة.
       - لا يُسمح باعتماد نفس الموظف + نفس التاريخ مرتين.

    2. الحضور المعتمد (is_approved = 1):
       - للقراءة فقط — يُستخدَم عند حساب الرواتب.
       - زر "إلغاء اعتماد" للمدير فقط في حالات الضرورة.

    التغييرات الجوهرية:
    - حُذف زر "تعبئة الغياب اليدوي" — الغياب يُضاف تلقائياً أثناء
      معالجة البصمات لكل يوم عمل ليس فيه سجل للموظف.
    - معالجة البصمات تستدعي _auto_fill_absent() بعد الانتهاء.
    - FingerprintImportDialog يستخدم اتصال db.conn المشترك (لا اتصال منفصل).
    """

    def __init__(self, db: DatabaseManager, user: dict, comm=None, parent=None):
        super().__init__(parent)
        self.db   = db
        self.user = user
        self.comm = comm
        self._build()
        self._set_initial_dates()
        self._load_draft()
        self._load_approved()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    def _on_data_changed(self, data_type: str, data):
        if data_type == 'employee':
            self._load_emp_combo(self.draft_emp_filter)
            self._load_emp_combo(self.appr_emp_filter)
        elif data_type == 'settings':
            self._update_work_settings()
        elif data_type == 'leave_request':
            self._load_draft()

    # ==================== إعدادات العمل ====================
    def _update_work_settings(self):
        self.work_start = QTime.fromString(
            self.db.get_setting('work_start_time', '08:00'), "HH:mm")
        self.work_end = QTime.fromString(
            self.db.get_setting('work_end_time', '17:00'), "HH:mm")
        self.work_hours = float(self.db.get_setting('working_hours', '8'))
        self.late_tol   = int(self.db.get_setting('late_tolerance_minutes', '10'))
        work_days_str   = self.db.get_setting('work_days', '0,1,2,3,4')
        self.work_days_indices = [
            int(x) for x in work_days_str.split(',') if x.strip().isdigit()]

    def _get_work_settings_dict(self) -> dict:
        return {
            'work_start': self.db.get_setting('work_start_time', '08:00'),
            'work_end':   self.db.get_setting('work_end_time',   '17:00'),
            'work_hours': float(self.db.get_setting('working_hours',          '8')),
            'late_tol':   int(self.db.get_setting('late_tolerance_minutes',  '10')),
        }

    def _get_break_settings_dict(self) -> dict:
        return {
            'lunch_break':    int(self.db.get_setting('lunch_break',    '30')),
            'num_breaks':     int(self.db.get_setting('num_breaks',     '2')),
            'break_duration': int(self.db.get_setting('break_duration', '15')),
            'include_breaks': int(self.db.get_setting('include_breaks', '0')),
        }

    # ==================== بناء الواجهة ====================
    def _build(self):
        layout = QVBoxLayout(self)

        self.inner_tabs = QTabWidget()
        self.inner_tabs.addTab(self._build_draft_tab(),    "📋 مسودة الحضور")
        self.inner_tabs.addTab(self._build_approved_tab(), "✅ الحضور المعتمد")
        layout.addWidget(self.inner_tabs)

        self._update_work_settings()

    # ----------------------------------------------------------
    # تبويب المسودة
    # ----------------------------------------------------------
    def _build_draft_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        # --- شريط الأزرار ---
        toolbar = QHBoxLayout()
        self.btn_import       = btn("📥 استيراد بصمات",    BTN_SUCCESS, self._import_fp)
        self.btn_process      = btn("⚙️ معالجة البصمات",   BTN_PRIMARY, self._process_fp)
        self.btn_manual       = btn("➕ إدخال يدوي",        BTN_WARNING, self._manual_entry)
        self.btn_edit         = btn("✏️ تعديل",             BTN_PRIMARY, self._edit_attendance)
        self.btn_bulk_edit    = btn("👥 تعديل جماعي",       BTN_PURPLE,  self._bulk_edit)
        self.btn_delete       = btn("🗑️ حذف",              BTN_DANGER,  self._delete_attendance)
        self.btn_delete_sel   = btn("🗑️ حذف المحدد",       BTN_DANGER,  self._delete_selected)
        self.btn_approve_per  = btn("✅ اعتماد الفترة",     BTN_SUCCESS, self._approve_period)
        self.btn_reset_fp     = btn("🔄 إعادة تعيين البصمات", BTN_WARNING, self._reset_fingerprints)
        self.btn_refresh_d    = btn("🔄 تحديث",             BTN_GRAY,    self._load_draft)
        self.btn_export_d     = btn("📊 تقرير Excel",       BTN_PURPLE,  self._export_excel_draft)
        self.btn_fp_template  = btn("📤 قالب بصمات",        BTN_TEAL,    self._export_fp_template)

        for b in (self.btn_import, self.btn_process, self.btn_manual,
                  self.btn_edit, self.btn_bulk_edit, self.btn_delete,
                  self.btn_delete_sel, self.btn_approve_per,
                  self.btn_reset_fp, self.btn_refresh_d,
                  self.btn_export_d, self.btn_fp_template):
            toolbar.addWidget(b)
        toolbar.addStretch()
        lay.addLayout(toolbar)

        # --- شريط الفلاتر ---
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("من:"))
        self.draft_date_from = QDateEdit()
        self.draft_date_from.setCalendarPopup(True)
        filter_row.addWidget(self.draft_date_from)
        filter_row.addWidget(QLabel("إلى:"))
        self.draft_date_to = QDateEdit()
        self.draft_date_to.setCalendarPopup(True)
        filter_row.addWidget(self.draft_date_to)

        self.draft_emp_filter = QComboBox()
        self.draft_emp_filter.setMinimumWidth(150)
        self._load_emp_combo(self.draft_emp_filter)
        self.draft_emp_filter.setEditable(True)
        self.draft_emp_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.draft_emp_filter.completer().setFilterMode(Qt.MatchContains)
        filter_row.addWidget(QLabel("الموظف:"))
        filter_row.addWidget(self.draft_emp_filter)

        # --- فلتر الحالة المتقدم ---
        self.status_group = QGroupBox("الحالة")
        status_layout = QVBoxLayout(self.status_group)

        self.chk_all_status = QCheckBox("جميع الحالات")
        self.chk_present    = QCheckBox("حاضر")
        self.chk_half_day   = QCheckBox("نصف يوم")
        self.chk_absent     = QCheckBox("غائب")
        self.chk_leave      = QCheckBox("إجازة")

        # الحالة الافتراضية: جميع الحالات محددة
        for chk in (
            self.chk_all_status,
            self.chk_present,
            self.chk_half_day,
            self.chk_absent,
            self.chk_leave
        ):
            chk.setChecked(True)

        status_layout.addWidget(self.chk_all_status)
        status_layout.addWidget(self.chk_present)
        status_layout.addWidget(self.chk_half_day)
        status_layout.addWidget(self.chk_absent)
        status_layout.addWidget(self.chk_leave)

        filter_row.addWidget(self.status_group)         
        lay.addLayout(filter_row)


        # معلومات أيام العمل
        self.lbl_work_days = QLabel("أيام العمل في الفترة: 0")
        self.lbl_work_days.setStyleSheet(
            "font-weight:bold; padding:4px; background:#e0f7fa;")
        lay.addWidget(self.lbl_work_days)

        # الجدول
        self.draft_table = make_table([
            "id", "الموظف", "التاريخ", "دخول", "خروج",
            "ساعات العمل", "أوفرتايم", "تأخير(د)", "خروج مبكر(د)",
            "الحالة", "ملاحظات"])
        self.draft_table.setColumnHidden(0, True)
        self.draft_table.setSortingEnabled(False)
        self.draft_table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.draft_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        lay.addWidget(self.draft_table)

        # ملخص
        sum_row = QHBoxLayout()
        self.lbl_total_d   = QLabel("السجلات: 0")
        self.lbl_hours_d   = QLabel("الساعات: 0")
        self.lbl_ot_d      = QLabel("أوفرتايم: 0")
        self.lbl_absent_d  = QLabel("غياب: 0")
        self.lbl_late_d    = QLabel("تأخير: 0.0 ساعة")
        self.lbl_early_d   = QLabel("خروج مبكر: 0.0 ساعة")
        for lbl in (self.lbl_total_d, self.lbl_hours_d, self.lbl_ot_d,
                    self.lbl_absent_d, self.lbl_late_d, self.lbl_early_d):
            lbl.setStyleSheet(
                "font-weight:bold; padding:6px; background:#e3f2fd; border-radius:4px;")
            sum_row.addWidget(lbl)
        sum_row.addStretch()
        lay.addLayout(sum_row)

        self._apply_draft_permissions()
        return w

    def _apply_draft_permissions(self):
        role = self.user['role']
        self.btn_import.setVisible(can_add(role))
        self.btn_process.setVisible(can_add(role))
        self.btn_manual.setVisible(can_add(role))
        self.btn_edit.setVisible(can_edit(role))
        self.btn_bulk_edit.setVisible(can_edit(role))
        self.btn_delete.setVisible(can_delete(role))
        self.btn_delete_sel.setVisible(can_delete(role))
        self.btn_approve_per.setVisible(can_approve(role))
        self.btn_reset_fp.setVisible(can_add(role))

    # ----------------------------------------------------------
    # تبويب المعتمد
    # ----------------------------------------------------------
    def _build_approved_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        toolbar = QHBoxLayout()
        self.btn_unapprove = btn("↩️ إلغاء اعتماد", BTN_WARNING, self._unapprove_record)
        self.btn_refresh_a = btn("🔄 تحديث",         BTN_GRAY,    self._load_approved)
        self.btn_export_a  = btn("📊 تقرير Excel",   BTN_PURPLE,  self._export_excel_approved)
        for b in (self.btn_unapprove, self.btn_refresh_a, self.btn_export_a):
            toolbar.addWidget(b)
        toolbar.addStretch()

        # تحذير توضيحي
        lbl_info = QLabel(
            "⚠️ هذه السجلات معتمدة وتُستخدَم في حساب الرواتب. "
            "تعديلها يستلزم إلغاء الاعتماد أولاً.")
        lbl_info.setStyleSheet(
            "color:#E65100; font-weight:bold; padding:5px; background:#FFF3E0;")
        toolbar.addWidget(lbl_info)
        lay.addLayout(toolbar)

        # فلاتر
        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("من:"))
        self.appr_date_from = QDateEdit()
        self.appr_date_from.setCalendarPopup(True)
        filter_row.addWidget(self.appr_date_from)
        filter_row.addWidget(QLabel("إلى:"))
        self.appr_date_to = QDateEdit()
        self.appr_date_to.setCalendarPopup(True)
        filter_row.addWidget(self.appr_date_to)

        self.appr_emp_filter = QComboBox()
        self.appr_emp_filter.setMinimumWidth(150)
        self._load_emp_combo(self.appr_emp_filter)
        self.appr_emp_filter.setEditable(True)
        self.appr_emp_filter.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.appr_emp_filter.completer().setFilterMode(Qt.MatchContains)
        filter_row.addWidget(QLabel("الموظف:"))
        filter_row.addWidget(self.appr_emp_filter)

        filter_row.addWidget(btn("بحث", BTN_PRIMARY, self._load_approved))
        filter_row.addStretch()
        lay.addLayout(filter_row)

        self.appr_table = make_table([
            "id", "الموظف", "التاريخ", "دخول", "خروج",
              "ساعات العمل", "أوفرتايم", "تأخير(د)", "خروج مبكر(د)",
              "الحالة", "تاريخ الاعتماد"
        ])

        self.appr_table.setColumnHidden(0, True)
        self.appr_table.setSortingEnabled(False)

        self.appr_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.appr_table.setToolTip(
            "هذه السجلات معتمدة وتُستخدم في حساب الرواتب.\n"
            "للتعديل، يجب إلغاء الاعتماد أولاً."
        )

        lay.addWidget(self.appr_table)

        # ملخص
        sum_row = QHBoxLayout()
        self.lbl_total_a  = QLabel("السجلات: 0")
        self.lbl_hours_a  = QLabel("الساعات: 0")
        self.lbl_ot_a     = QLabel("أوفرتايم: 0")
        self.lbl_absent_a = QLabel("غياب: 0")
        for lbl in (self.lbl_total_a, self.lbl_hours_a,
                    self.lbl_ot_a, self.lbl_absent_a):
            lbl.setStyleSheet(
                "font-weight:bold; padding:6px; background:#E8F5E9; border-radius:4px;")
            sum_row.addWidget(lbl)
        sum_row.addStretch()
        lay.addLayout(sum_row)

        self.btn_unapprove.setVisible(self.user['role'] == 'admin')
        return w

    # ==================== دوال مساعدة ====================
    def _load_emp_combo(self, combo: QComboBox):
        combo.clear()
        combo.addItem("-- جميع الموظفين --", None)
        for eid, name in self.db.fetch_all(
                "SELECT id, first_name||' '||last_name "
                "FROM employees WHERE status='نشط'"):
            combo.addItem(name, eid)

    def _get_min_max_dates(self):
        r1 = self.db.fetch_one(
            "SELECT MIN(punch_date), MAX(punch_date) FROM attendance")
        r2 = self.db.fetch_one(
            "SELECT MIN(DATE(punch_datetime)), MAX(DATE(punch_datetime)) "
            "FROM fingerprint_raw")
        dates = []
        for r in (r1, r2):
            if r and r[0]:
                dates.append(QDate.fromString(r[0], Qt.ISODate))
            if r and r[1]:
                dates.append(QDate.fromString(r[1], Qt.ISODate))
        valid = [d for d in dates if d.isValid()]
        if not valid:
            return None, None
        return min(valid), max(valid)

    def _set_initial_dates(self):
        min_d, max_d = self._get_min_max_dates()
        default_from = QDate.currentDate().addDays(-365)
        default_to   = QDate.currentDate()
        from_d = max(min_d, default_from) if min_d else default_from
        to_d   = max_d if max_d else default_to
        for combo in (self.draft_date_from, self.appr_date_from):
            combo.setDate(from_d)
        for combo in (self.draft_date_to, self.appr_date_to):
            combo.setDate(to_d)

    def _set_all_dates(self):
        min_d, max_d = self._get_min_max_dates()
        if min_d and max_d:
            self.draft_date_from.setDate(min_d)
            self.draft_date_to.setDate(max_d)
            self._load_draft()
        else:
            QMessageBox.information(self, "معلومات",
                                    "لا توجد بيانات حضور في النظام بعد.")

    def _update_work_days_label(self):
        from_d = self.draft_date_from.date().toPyDate()
        to_d   = self.draft_date_to.date().toPyDate()
        count  = sum(
            1 for d in _date_range(from_d, to_d)
            if d.weekday() in self.work_days_indices)
        self.lbl_work_days.setText(f"أيام العمل في الفترة: {count}")

    # ==================== تحميل المسودة ====================

    def _load_draft(self):
        d_from = self.draft_date_from.date().toString(Qt.ISODate)
        d_to   = self.draft_date_to.date().toString(Qt.ISODate)

        emp_id     = self.draft_emp_filter.currentData()

        # تنبيه في حال وجود سجلات معتمدة ضمن نفس الفترة
        q_check = """
            SELECT COUNT(*) FROM attendance
            WHERE is_approved = 1
              AND punch_date BETWEEN ? AND ?
        """
        params_check = [d_from, d_to]

        if emp_id is not None:
            q_check += " AND employee_id = ?"
            params_check.append(emp_id)

        q = """
            SELECT a.id, e.first_name || ' ' || e.last_name,
                   a.punch_date, a.check_in, a.check_out,
                   ROUND(a.work_hours, 2), ROUND(a.overtime_hours, 2),
                   a.late_minutes, a.early_leave_minutes,
                   a.status, a.notes
            FROM attendance a
            JOIN employees e ON a.employee_id = e.id
            WHERE a.is_approved = 0
              AND a.punch_date BETWEEN ? AND ?
        """
        params = [d_from, d_to]

        if emp_id is not None:
            q += " AND a.employee_id = ?"
            params.append(emp_id)

        # تطبيق فلتر الحالة المتقدم
selected_statuses = self._get_selected_statuses()
q = _apply_status_filter(q, params, selected_statuses)

q += " ORDER BY a.punch_date DESC, e.first_name"


        data = self.db.fetch_all(q, params)
        self._draft_ids = [row[0] for row in data]

        fill_table(self.draft_table, data)

        _colorize_attendance(
            self.draft_table,
            data,
            status_col=9,
            late_col=7,
            early_col=8,
            ot_col=6
        )

        _update_summary(
            data,
            self.lbl_total_d,
            self.lbl_hours_d,
            self.lbl_ot_d,
            self.lbl_absent_d,
            self.lbl_late_d,
            self.lbl_early_d,
            hours_col=5,
            ot_col=6,
            late_col=7,
            early_col=8,
            status_col=9
        )

        self._update_work_days_label()
    def _get_selected_statuses(self):
        statuses = []
        if self.chk_present.isChecked():
            statuses.append("حاضر")
        if self.chk_half_day.isChecked():
            statuses.append("نصف يوم")
        if self.chk_absent.isChecked():
        statuses.append("غائب")
        if self.chk_leave.isChecked():
        statuses.append("إجازة")
        return statuses

    def _warn_if_approved_overlap(self) -> None:
        d_from = self.draft_date_from.date().toString(Qt.ISODate)
        d_to   = self.draft_date_to.date().toString(Qt.ISODate)

        emp_id = self.draft_emp_filter.currentData()

        q_check = """
            SELECT COUNT(*) FROM attendance
            WHERE is_approved = 1
              AND punch_date BETWEEN ? AND ?
        """
        params_check = [d_from, d_to]

        if emp_id is not None:
            q_check += " AND employee_id = ?"
            params_check.append(emp_id)

        approved_count = self.db.fetch_one(q_check, params_check)
        if approved_count and approved_count[0] > 0:
            QMessageBox.information(
                self,
                "تنبيه",
                "⚠️ توجد سجلات حضور معتمدة ضمن هذه الفترة.\n"
                "لن تظهر في المسودة، ويمكنك مراجعتها من تبويب الحضور المعتمد."
            )

    # ==================== تحميل المعتمد ====================
    def _load_approved(self):
        d_from = self.appr_date_from.date().toString(Qt.ISODate)
        d_to   = self.appr_date_to.date().toString(Qt.ISODate)
        emp_id = self.appr_emp_filter.currentData()

        q = """
            SELECT a.id, e.first_name||' '||e.last_name,
                   a.punch_date, a.check_in, a.check_out,
                   ROUND(a.work_hours,2), ROUND(a.overtime_hours,2),
                   a.late_minutes, a.early_leave_minutes, a.status,
                   a.approved_at
            FROM attendance a
            JOIN employees e ON a.employee_id = e.id
            WHERE a.is_approved = 1
              AND a.punch_date BETWEEN ? AND ?
        """
        params = [d_from, d_to]
        if emp_id is not None:
            q += " AND a.employee_id = ?"
            params.append(emp_id)
        q += " ORDER BY a.punch_date DESC, e.first_name"

        data = self.db.fetch_all(q, params)
        self._appr_ids = [row[0] for row in data]
        fill_table(self.appr_table, data)
        _colorize_attendance(self.appr_table, data, status_col=9,
                             late_col=7, early_col=8, ot_col=6)
        total   = len(data)
        hours   = sum(row[5] for row in data if row[5])
        ot      = sum(row[6] for row in data if row[6])
        absent  = sum(1 for row in data if row[9] == "غائب")
        self.lbl_total_a.setText(f"السجلات: {total}")
        self.lbl_hours_a.setText(f"الساعات: {hours:.1f}")
        self.lbl_ot_a.setText(f"أوفرتايم: {ot:.1f}")
        self.lbl_absent_a.setText(f"غياب: {absent}")

    # ==================== اعتماد الفترة ====================
    def _approve_period(self):
        """
        اعتماد جميع سجلات المسودة ضمن الفترة المحددة.

        الضمانات:
        - لا يُعتمَد أي سجل إذا كان هناك سجل معتمد للموظف نفسه
          في نفس التاريخ (منع التضارب).
        - يُسجَّل approved_by و approved_at لكل سجل.
        """
        d_from  = self.draft_date_from.date().toString(Qt.ISODate)
        d_to    = self.draft_date_to.date().toString(Qt.ISODate)
        emp_id  = self.draft_emp_filter.currentData()

        # التحقق من وجود سجلات مسودة
        q_count = """
            SELECT COUNT(*) FROM attendance
            WHERE is_approved = 0
              AND punch_date BETWEEN ? AND ?
        """
        params_count = [d_from, d_to]
        if emp_id is not None:
            q_count += " AND employee_id = ?"
            params_count.append(emp_id)

        cnt = self.db.fetch_one(q_count, params_count)
        draft_count = cnt[0] if cnt else 0

        if draft_count == 0:
            QMessageBox.information(self, "معلومات",
                                    "لا توجد سجلات مسودة لاعتمادها في هذه الفترة.")
            return

        # التحقق من تضارب مع معتمد سابق
        q_conflict = """
            SELECT COUNT(*) FROM attendance d
            WHERE d.is_approved = 0
              AND d.punch_date BETWEEN ? AND ?
              AND EXISTS (
                  SELECT 1 FROM attendance a
                  WHERE a.employee_id = d.employee_id
                    AND a.punch_date  = d.punch_date
                    AND a.is_approved = 1
              )
        """
        params_cf = [d_from, d_to]
        if emp_id is not None:
            q_conflict += " AND d.employee_id = ?"
            params_cf.append(emp_id)

        cf = self.db.fetch_one(q_conflict, params_cf)
        conflicts = cf[0] if cf else 0

        emp_label = "الفترة المحددة"
        if emp_id is not None:
            emp_row = self.db.fetch_one(
                "SELECT first_name||' '||last_name FROM employees WHERE id=?",
                (emp_id,))
            if emp_row:
                emp_label = emp_row[0]

        msg = (f"سيتم اعتماد {draft_count} سجل حضور للفترة "
               f"من {d_from} إلى {d_to}\nللموظف: {emp_label}")
        if conflicts > 0:
            msg += (f"\n\n⚠️ تحذير: {conflicts} سجل لديها تضارب مع سجلات "
                    f"معتمدة مسبقاً وسيتم تخطيها.")

        if QMessageBox.question(
                self, "تأكيد الاعتماد", msg,
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return

        now = datetime.now().isoformat()
        q_update = """
            UPDATE attendance
            SET is_approved = 1,
                approved_at = ?,
                approved_by = ?
            WHERE is_approved = 0
              AND punch_date BETWEEN ? AND ?
              AND NOT EXISTS (
                  SELECT 1 FROM attendance a2
                  WHERE a2.employee_id = attendance.employee_id
                    AND a2.punch_date  = attendance.punch_date
                    AND a2.is_approved = 1
              )
        """
        params_upd = [now, self.user['id'], d_from, d_to]
        if emp_id is not None:
            q_update += " AND employee_id = ?"
            params_upd.append(emp_id)

        self.db.execute_query(q_update, params_upd)
        self.db.log_custom(
            "اعتماد الحضور",
            "attendance",
            details={"from": d_from, "to": d_to,
                     "employee_id": emp_id, "count": draft_count})

        QMessageBox.information(
            self, "نجاح",
            f"تم اعتماد سجلات الحضور بنجاح.\n"
            f"يمكنك الآن حساب الرواتب لهذه الفترة.")

        self._load_draft()
        self._load_approved()
        if self.comm:
            self.comm.dataChanged.emit('attendance', {'action': 'approve'})

    # ==================== إلغاء الاعتماد ====================
    def _unapprove_record(self):
        """إلغاء اعتماد سجل — للمدير فقط — في حالات الضرورة."""
        row = self.appr_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "خطأ", "اختر سجلاً أولاً")
            return
        if not hasattr(self, '_appr_ids') or row >= len(self._appr_ids):
            return
        rec_id = self._appr_ids[row]

        if QMessageBox.question(
                self, "تأكيد إلغاء الاعتماد",
                "سيُنقَل هذا السجل إلى المسودة وسيتوقف احتسابه في الرواتب.\n"
                "هل أنت متأكد؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.execute_query(
                "UPDATE attendance SET is_approved=0, approved_at=NULL, "
                "approved_by=NULL WHERE id=?",
                (rec_id,))
            self._load_draft()
            self._load_approved()

    # ==================== عمليات المسودة ====================
    def _manual_entry(self, record=None):
        dlg = ManualAttendanceDialog(self, self.db, record)
        if dlg.exec_() == QDialog.Accepted:
            if record is None:
                nd = dlg.punch_date.date()
                if nd < self.draft_date_from.date():
                    self.draft_date_from.setDate(nd)
                if nd > self.draft_date_to.date():
                    self.draft_date_to.setDate(nd)
            self._load_draft()
            if self.comm:
                self.comm.dataChanged.emit('attendance', {'action': 'add'})

    def _edit_attendance(self):
        row = self.draft_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "خطأ", "اختر سجلاً أولاً")
            return
        if not hasattr(self, '_draft_ids') or row >= len(self._draft_ids):
            return
        rec_id = self._draft_ids[row]
        record = self.db.fetch_one(
            "SELECT * FROM attendance WHERE id=?", (rec_id,))
        if not record:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على السجل")
            return
        self._manual_entry(record)

    def _delete_attendance(self):
        row = self.draft_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "خطأ", "اختر سجلاً أولاً")
            return
        if not hasattr(self, '_draft_ids') or row >= len(self._draft_ids):
            return
        rec_id = self._draft_ids[row]
        if QMessageBox.question(
                self, "تأكيد", "هل أنت متأكد من حذف هذا السجل؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.execute_query(
                "DELETE FROM attendance WHERE id=?", (rec_id,))
            self._load_draft()

    def _delete_selected(self):
        selected = set(item.row() for item in self.draft_table.selectedItems())
        if not selected or not hasattr(self, '_draft_ids'):
            QMessageBox.warning(self, "خطأ", "لم يتم تحديد أي سجلات")
            return
        ids = [self._draft_ids[r] for r in selected if r < len(self._draft_ids)]
        if not ids:
            return
        if QMessageBox.question(
                self, "تأكيد", f"هل أنت متأكد من حذف {len(ids)} سجل؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            for rid in ids:
                self.db.execute_query(
                    "DELETE FROM attendance WHERE id=?", (rid,))
            self._load_draft()

    def _bulk_edit(self):
        selected = set(item.row() for item in self.draft_table.selectedItems())
        if not selected:
            QMessageBox.warning(self, "تنبيه", "يرجى تحديد صفوف للتعديل")
            return
        ids = [self._draft_ids[r] for r in selected
               if r < len(self._draft_ids)]
        if not ids:
            return
        dlg = BulkEditAttendanceDialog(
            self, self.db, ids, self._get_work_settings_dict())
        if dlg.exec_() == QDialog.Accepted:
            self._load_draft()
            if self.comm:
                self.comm.dataChanged.emit('attendance', {'action': 'bulk_edit'})

    # ==================== استيراد ومعالجة البصمات ====================
    def _import_fp(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "اختر ملف البصمة", "",
            "Text Files (*.txt);;CSV Files (*.csv);;"
            "DAT Files (*.dat);;All Files (*)")
        if not path:
            return
        dlg = FingerprintImportDialog(self, path, self.db)
        if dlg.exec_() == QDialog.Accepted:
            QMessageBox.information(
                self, "تم الاستيراد",
                "تم استيراد البصمات بنجاح.\n"
                "اضغط 'معالجة البصمات' لتحويلها إلى سجلات حضور.")
            self._load_draft()

    def _export_fp_template(self):
        try:
            import pandas as pd
            path, _ = QFileDialog.getSaveFileName(
                self, "حفظ قالب البصمات",
                "fingerprint_template.xlsx", "Excel (*.xlsx)")
            if not path:
                return
            df = pd.DataFrame({
                'EnNo': [1001, 1002],
                'DateTime': ['2025-01-01 08:00:00', '2025-01-01 17:00:00'],
                'Name': ['موظف 1', 'موظف 2']
            })
            df.to_excel(path, index=False)
            QMessageBox.information(
                self, "نجاح",
                "تم حفظ قالب البصمات.\n"
                "يجب أن يحتوي الملف على عمودي EnNo و DateTime على الأقل.")
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")

    def _process_fp(self):
        unprocessed = self.db.fetch_one(
            "SELECT COUNT(*) FROM fingerprint_raw WHERE processed=0")
        count = unprocessed[0] if unprocessed else 0
        if count == 0:
            QMessageBox.information(
                self, "معلومة",
                "لا توجد بصمات جديدة للمعالجة.\n"
                "إذا كنت قد استوردت بصمات سابقاً، فقد تكون جميعها معالجة.")
            return

        if QMessageBox.question(
                self, "معالجة البصمات",
                f"يوجد {count} سجل غير معالج.\n"
                f"بعد المعالجة سيتم إضافة أيام الغياب تلقائياً.\n"
                f"هل تريد المتابعة؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return

        self._prog = QProgressBar()
        self._prog.setWindowTitle("جارٍ المعالجة...")
        self._prog.setRange(0, count)
        self._prog.setMinimumWidth(400)
        self._prog.setWindowFlags(Qt.Dialog | Qt.WindowTitleHint)
        self._prog.setAlignment(Qt.AlignCenter)
        self._prog.setFormat("معالجة البصمات... %v / %m")
        self._prog.show()

        self._worker = FingerprintWorker(
            self.db.db_name,
            self._get_work_settings_dict(),
            self._get_break_settings_dict())
        self._worker.progress.connect(lambda done, tot: self._prog.setValue(done))
        self._worker.finished.connect(self._on_process_done)
        self._worker.error.connect(self._on_process_error)
        self._worker.start()

    def _on_process_done(self, ok, fail, no_emp, unmatched, leave_conflicts):
        self._prog.close()

        # ← إضافة الغياب تلقائياً بعد المعالجة
        absent_added = self._auto_fill_absent()

        msg = (f"✅ تمت معالجة البصمات: {ok} سجل\n"
               f"📅 غياب تلقائي مضاف: {absent_added} يوم\n"
               f"❌ فشل تقني: {fail}\n"
               f"👤 بدون موظف مطابق: {no_emp}")
        if unmatched:
            sample = sorted(unmatched)[:10]
            msg += ("\n\n─── أرقام بصمة غير موجودة ───\n"
                    + "\n".join(f"  • {fp}" for fp in sample))
            if len(unmatched) > 10:
                msg += f"\n  ... و{len(unmatched)-10} أخرى"
            msg += "\n\n💡 تأكد من إضافة الموظفين بأرقام البصمة الصحيحة."
        if leave_conflicts:
            msg += "\n\n⚠️ أيام حُوِّلت إلى إجازة:"
            for emp, days in leave_conflicts.items():
                msg += f"\n   {emp}: {', '.join(days)}"

        (QMessageBox.information if ok > 0 else QMessageBox.warning)(
            self, "نتيجة المعالجة", msg)
        self._load_draft()

    def _on_process_error(self, err: str):
        self._prog.close()
        QMessageBox.critical(self, "خطأ في المعالجة", err)

    def _auto_fill_absent(self) -> int:
        """
        يُضيف أيام غياب تلقائياً لكل موظف نشط غير معفى من البصمة
        في كل يوم عمل ليس فيه سجل حضور ولا إجازة معتمدة.

        يُستدعى تلقائياً بعد معالجة البصمات.
        يُرجع عدد أيام الغياب المضافة.
        """
        work_days_str    = self.db.get_setting('work_days', '0,1,2,3,4')
        work_days_idx    = [int(x) for x in work_days_str.split(',')
                            if x.strip().isdigit()]

        # تحديد نطاق التواريخ من fingerprint_raw المعالجة
        r = self.db.fetch_one(
            "SELECT MIN(DATE(punch_datetime)), MAX(DATE(punch_datetime)) "
            "FROM fingerprint_raw WHERE processed = 1")
        if not r or not r[0]:
            return 0

        from_d = date.fromisoformat(r[0])
        to_d   = date.fromisoformat(r[1])

        # جلب الموظفين النشطين غير المعفيين
        employees = self.db.fetch_all(
    "SELECT id, hire_date, termination_date FROM employees "
    "WHERE status='نشط' AND is_exempt_from_fingerprint=0")
        if not employees:
            return 0

        # جلب الإجازات المعتمدة دفعة واحدة
        leave_days: dict[tuple, bool] = {}
        for emp_id, s_str, e_str in self.db.fetch_all(
                """SELECT employee_id, start_date, end_date
                   FROM leave_requests
                   WHERE status='موافق'
                     AND start_date <= ? AND end_date >= ?""",
                (to_d.isoformat(), from_d.isoformat())):
            for d in _date_range(date.fromisoformat(s_str),
                                 date.fromisoformat(e_str)):
                leave_days[(emp_id, d.isoformat())] = True

        # جلب سجلات الحضور الموجودة دفعة واحدة
        existing: dict[tuple, bool] = {}
        for emp_id, punch_d in self.db.fetch_all(
                "SELECT employee_id, punch_date FROM attendance "
                "WHERE punch_date BETWEEN ? AND ?",
                (from_d.isoformat(), to_d.isoformat())):
            existing[(emp_id, punch_d)] = True

        # أيام العمل في النطاق
        work_days = [d for d in _date_range(from_d, to_d)
                     if d.weekday() in work_days_idx]

        added = 0
        for emp_id, hire_d, term_d in employees:
            hire_date = date.fromisoformat(hire_d) if hire_d else None
            term_date = date.fromisoformat(term_d) if term_d else None

            for day in work_days:
                if hire_date and day < hire_date:
                    continue
                if term_date and day > term_date:
                    continue
                key = (emp_id, day.isoformat())
                if existing.get(key):
                    continue
                if leave_days.get(key):
                    continue
                self.db.execute_query(
                    """INSERT INTO attendance
                       (employee_id, punch_date, status,
                        work_hours, overtime_hours,
                        late_minutes, early_leave_minutes, notes)
                       VALUES (?,?,?,?,?,?,?,?)""",
                    (emp_id, day.isoformat(), "غائب",
                     0, 0, 0, 0, "غياب تلقائي"))
                added += 1

        return added

    def _reset_fingerprints(self):
        if QMessageBox.question(
                self, "تأكيد إعادة التعيين",
                "سيتم إعادة تعيين جميع البصمات إلى حالة غير معالجة.\n"
                "هل أنت متأكد؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            if self.db.execute_query(
                    "UPDATE fingerprint_raw SET processed = 0"):
                QMessageBox.information(
                    self, "تم",
                    "تم إعادة تعيين البصمات.\n"
                    "يمكنك الآن معالجتها من جديد.")
                self._load_draft()
            else:
                QMessageBox.critical(self, "خطأ", "فشل إعادة التعيين.")

    # ==================== تصدير Excel ====================
    def _export_excel_draft(self):
        self._export_excel(approved=False)

    def _export_excel_approved(self):
        self._export_excel(approved=True)

    def _export_excel(self, approved: bool = False):
        try:
            import pandas as pd
        except ImportError:
            QMessageBox.critical(self, "خطأ", "pip install pandas openpyxl")
            return

        if approved:
            d_from = self.appr_date_from.date().toString(Qt.ISODate)
            d_to   = self.appr_date_to.date().toString(Qt.ISODate)
            emp_id = self.appr_emp_filter.currentData()
            is_app = 1
        else:
            d_from = self.draft_date_from.date().toString(Qt.ISODate)
            d_to   = self.draft_date_to.date().toString(Qt.ISODate)
            emp_id = self.draft_emp_filter.currentData()
            is_app = 0

        q = """
            SELECT e.employee_code, e.first_name||' '||e.last_name,
                   d.name, a.punch_date, a.check_in, a.check_out,
                   a.work_hours, a.overtime_hours,
                   a.late_minutes, a.early_leave_minutes, a.status, a.notes
            FROM attendance a
            JOIN employees e ON a.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE a.is_approved = ?
              AND a.punch_date BETWEEN ? AND ?
        """
        params = [is_app, d_from, d_to]
        if emp_id is not None:
            q += " AND a.employee_id = ?"
            params.append(emp_id)
        q += " ORDER BY a.punch_date, e.first_name"

        data = self.db.fetch_all(q, params)
        if not data:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
            return

        df = pd.DataFrame(data, columns=[
            "الرقم", "الاسم", "القسم", "التاريخ", "دخول", "خروج",
            "ساعات العمل", "أوفرتايم", "تأخير(د)", "خروج مبكر(د)",
            "الحالة", "ملاحظات"])

        label = "معتمد" if approved else "مسودة"
        path, _ = QFileDialog.getSaveFileName(
            self, "حفظ التقرير",
            f"attendance_{label}_{d_from}.xlsx", "Excel (*.xlsx)")
        if path:
            df.to_excel(path, index=False)
            QMessageBox.information(self, "نجاح", "تم تصدير التقرير")


# ==================== دوال مساعدة على مستوى الـ module ====================

def _date_range(start: date, end: date):
    """مولِّد للتواريخ من start إلى end شاملاً."""
    current = start
    while current <= end:
        yield current
        current += timedelta(days=1)


def _apply_status_filter(q: str, params: list, status_filter) -> str:
    """
    يدعم:
    - نص واحد: "حاضر"
    - قائمة: ["حاضر", "غائب"]
    - None / [] => جميع الحالات
    """
    if not status_filter:
        return q

    # إذا جاء نص واحد
    if isinstance(status_filter, str):
        if status_filter == "جميع الحالات":
            return q
        q += " AND a.status = ?"
        params.append(status_filter)
        return q

    # إذا جاءت قائمة
    status_list = [s for s in status_filter if s and s != "جميع الحالات"]
    if not status_list:
        return q

    placeholders = ",".join("?" for _ in status_list)
    q += f" AND a.status IN ({placeholders})"
    params.extend(status_list)
    return q

def _colorize_attendance(table, data, status_col, late_col, early_col, ot_col):
    """تلوين خلايا الجدول بناءً على الحالة والتأخير."""
    for row_idx, row_data in enumerate(data):
        status = row_data[status_col + 1]   # +1 لأن العمود الأول هو id
        late   = row_data[late_col + 1]
        early  = row_data[early_col + 1]
        ot     = row_data[ot_col + 1]

        status_item = table.item(row_idx, status_col)
        if status_item:
            if status == "غائب":
                status_item.setBackground(QColor("#D32F2F"))
                status_item.setForeground(QColor("white"))
            elif status == "نصف يوم":
                status_item.setBackground(QColor("#FFA500"))
            elif status == "إجازة":
                status_item.setBackground(QColor("#4CAF50"))
                status_item.setForeground(QColor("white"))

        late_val = safe_float(late)
        if late_val > 0:
            item = table.item(row_idx, late_col)
            if item:
                item.setBackground(QColor("#FFCDD2"))
                item.setToolTip(f"تأخير: {late_val} دقيقة")

        early_val = safe_float(early)
        if early_val > 0:
            item = table.item(row_idx, early_col)
            if item:
                item.setBackground(QBrush(QColor("#FFF9C4")))
                item.setToolTip(f"خروج مبكر: {early_val} دقيقة")

        ot_val = safe_float(ot)
        if ot_val > 0:
            item = table.item(row_idx, ot_col)
            if item:
                item.setForeground(QColor("#2E7D32"))
                item.setToolTip(f"أوفرتايم: {ot_val} ساعة")

def _update_summary(data, lbl_total, lbl_hours, lbl_ot, lbl_absent,
                    lbl_late=None, lbl_early=None,
                    hours_col=5, ot_col=6, late_col=7, early_col=8, status_col=9):

    summary = calculate_attendance_summary(
        data,
        hours_col=hours_col,
        ot_col=ot_col,
        late_col=late_col,
        early_col=early_col,
        status_col=status_col
    )

    lbl_total.setText(f"السجلات: {summary['records']}")
    lbl_hours.setText(f"الساعات: {summary['hours']:.1f}")
    lbl_ot.setText(f"أوفرتايم: {summary['overtime']:.1f}")
    lbl_absent.setText(f"غياب: {summary['absent']}")

    if lbl_late:
        lbl_late.setText(f"تأخير: {summary['late_hours']:.1f} ساعة")

    if lbl_early:
        lbl_early.setText(f"خروج مبكر: {summary['early_hours']:.1f} ساعة")

def calculate_attendance_summary(data,
                                 hours_col=5,
                                 ot_col=6,
                                 late_col=7,
                                 early_col=8,
                                 status_col=9):
    """
    حساب ملخص الحضور بشكل موحّد.
    يُرجع القيم بالساعات حيثما ينطبق.
    """
    # +1 لأن العمود الأول هو id
    h_col  = hours_col + 1
    o_col  = ot_col    + 1
    la_col = late_col  + 1
    ea_col = early_col + 1
    st_col = status_col + 1

    total_records = len(data)
    total_hours   = sum(safe_float(row[h_col]) for row in data)
    total_ot      = sum(safe_float(row[o_col]) for row in data)
    absent_days   = sum(1 for row in data if row[st_col] == "غائب")

    late_minutes  = sum(safe_float(row[la_col]) for row in data)
    early_minutes = sum(safe_float(row[ea_col]) for row in data)

    return {
        "records": total_records,
        "hours": round(total_hours, 2),
        "overtime": round(total_ot, 2),
        "absent": absent_days,
        "late_hours": round(late_minutes / 60.0, 2),
        "early_hours": round(early_minutes / 60.0, 2),
    }


# ==================== نافذة استيراد البصمات ====================
class FingerprintImportDialog(QDialog):
    """
    استيراد ملفات البصمات إلى fingerprint_raw.

    الإصلاح: يستخدم db.conn المشترك بدلاً من فتح اتصال SQLite جديد،
    مما يمنع تعارض القفل في WAL mode.
    """

    def __init__(self, parent, file_path: str, db: DatabaseManager):
        super().__init__(parent)
        self.file_path = file_path
        self.db        = db
        self.setWindowTitle("استيراد بصمات")
        self.setMinimumSize(900, 700)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()
        self._detect_all()

    def _build(self):
        lay = QVBoxLayout(self)

        lbl = QLabel(f"الملف: {os.path.basename(self.file_path)}")
        lbl.setStyleSheet("font-weight:bold; color:#1976D2; font-size:13px;")
        lay.addWidget(lbl)

        sg = QGroupBox("إعدادات الاستيراد")
        sl = QGridLayout()

        sl.addWidget(QLabel("الترميز:"), 0, 0)
        self.cmb_encoding = QComboBox()
        self.cmb_encoding.addItems(["utf-8","windows-1254","latin-1","utf-16","cp1256"])
        self.cmb_encoding.setEditable(True)
        self.cmb_encoding.currentTextChanged.connect(self._update_preview)
        sl.addWidget(self.cmb_encoding, 0, 1)

        sl.addWidget(QLabel("الفاصل:"), 1, 0)
        self.cmb_delimiter = QComboBox()
        self.cmb_delimiter.addItems(["فاصلة (,)", "تاب (\\t)", "فاصلة منقوطة (;)", "مسافة ( )"])
        self.cmb_delimiter.setCurrentIndex(1)
        self.cmb_delimiter.currentIndexChanged.connect(self._update_preview)
        sl.addWidget(self.cmb_delimiter, 1, 1)

        sl.addWidget(QLabel("تخطي أول N سطراً:"), 2, 0)
        self.spin_skip = QSpinBox()
        self.spin_skip.setRange(0, 100)
        self.spin_skip.setValue(1)
        self.spin_skip.valueChanged.connect(self._update_preview)
        sl.addWidget(self.spin_skip, 2, 1)

        sl.addWidget(QLabel("رقم عمود البصمة:"), 3, 0)
        self.spin_fp_col = QSpinBox()
        self.spin_fp_col.setRange(0, 50)
        self.spin_fp_col.setValue(2)
        self.spin_fp_col.valueChanged.connect(self._update_preview)
        sl.addWidget(self.spin_fp_col, 3, 1)

        sl.addWidget(QLabel("رقم عمود التاريخ:"), 4, 0)
        self.spin_dt_col = QSpinBox()
        self.spin_dt_col.setRange(0, 50)
        self.spin_dt_col.setValue(9)
        self.spin_dt_col.valueChanged.connect(self._update_preview)
        sl.addWidget(self.spin_dt_col, 4, 1)

        sl.addWidget(QLabel("صيغة التاريخ:"), 5, 0)
        self.edit_dt_fmt = QLineEdit("%Y-%m-%d %H:%M:%S")
        self.edit_dt_fmt.textChanged.connect(self._update_preview)
        sl.addWidget(self.edit_dt_fmt, 5, 1)

        sl.addWidget(QLabel("من تاريخ:"), 6, 0)
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addMonths(-1))
        sl.addWidget(self.date_from, 6, 1)

        sl.addWidget(QLabel("إلى تاريخ:"), 7, 0)
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        sl.addWidget(self.date_to, 7, 1)

        sl.addWidget(btn("🔄 كشف تلقائي", BTN_TEAL, self._detect_all), 8, 0, 1, 2)
        sg.setLayout(sl)
        lay.addWidget(sg)

        pg = QGroupBox("معاينة البيانات (أول 20 سطر)")
        pl = QVBoxLayout()
        self.preview_text = QTextEdit()
        self.preview_text.setReadOnly(True)
        self.preview_text.setFont(QFont("Courier New", 9))
        pl.addWidget(self.preview_text)
        pg.setLayout(pl)
        lay.addWidget(pg, stretch=1)

        self.result_label = QLabel("")
        self.result_label.setStyleSheet("padding:6px; background:#f5f5f5;")
        lay.addWidget(self.result_label)

        bl = QHBoxLayout()
        bl.addWidget(btn("📥 استيراد", BTN_SUCCESS, self._import))
        bl.addWidget(btn("❌ إلغاء",   BTN_DANGER,  self.reject))
        lay.addLayout(bl)

    def _get_delimiter(self) -> str:
        return [',', '\t', ';', ' '][self.cmb_delimiter.currentIndex()]

    def _read_lines(self, encoding: str):
        try:
            with open(self.file_path, 'r',
                      encoding=encoding, errors='replace') as f:
                return f.readlines()
        except Exception:
            return None

    def _detect_all(self):
        try:
            import chardet
            with open(self.file_path, 'rb') as f:
                res = chardet.detect(f.read(10000))
            enc = res.get('encoding', 'utf-8') or 'utf-8'
            idx = self.cmb_encoding.findText(enc)
            if idx >= 0:
                self.cmb_encoding.setCurrentIndex(idx)
            else:
                self.cmb_encoding.setCurrentText(enc)
        except ImportError:
            pass
        except Exception:
            pass

        lines = self._read_lines(self.cmb_encoding.currentText())
        if not lines:
            return

        # كشف الفاصل
        try:
            dialect = csv.Sniffer().sniff("\n".join(lines[:5]))
            d = dialect.delimiter
            for i, c in enumerate([',', '\t', ';', ' ']):
                if d == c:
                    self.cmb_delimiter.setCurrentIndex(i)
                    break
        except Exception:
            pass

        # تحليل الأعمدة
        import re
        delim   = self._get_delimiter()
        skip    = self.spin_skip.value()
        rows    = [l.strip().split(delim) for l in lines[skip:skip+20]]
        n_cols  = max((len(r) for r in rows), default=0)
        fp_sc   = [0] * n_cols
        dt_sc   = [0] * n_cols
        dt_pat  = re.compile(r'\d{4}[-/]\d{2}[-/]\d{2}[\sT]\d{2}:\d{2}')
        for row in rows:
            for i, v in enumerate(row):
                v = v.strip()
                if v.isdigit() and 4 <= len(v) <= 10:
                    fp_sc[i] += 1
                if dt_pat.search(v):
                    dt_sc[i] += 1
        if fp_sc and max(fp_sc) > 0:
            self.spin_fp_col.setValue(fp_sc.index(max(fp_sc)))
        if dt_sc and max(dt_sc) > 0:
            self.spin_dt_col.setValue(dt_sc.index(max(dt_sc)))

        self._update_preview()

    def _update_preview(self):
        lines = self._read_lines(self.cmb_encoding.currentText())
        if not lines:
            self.preview_text.setText("تعذر قراءة الملف.")
            return
        delim  = self._get_delimiter()
        skip   = self.spin_skip.value()
        fp_c   = self.spin_fp_col.value()
        dt_c   = self.spin_dt_col.value()
        out    = []
        for i, line in enumerate(lines[skip:skip+20]):
            parts = line.strip().split(delim)
            fp = parts[fp_c] if fp_c < len(parts) else "N/A"
            dt = parts[dt_c] if dt_c < len(parts) else "N/A"
            out.append(f"[{i+1}] البصمة: {fp:15} التاريخ: {dt}")
        self.preview_text.setText("\n".join(out))

    def _import(self):
        try:
            delim    = self._get_delimiter()
            skip     = self.spin_skip.value()
            fp_col   = self.spin_fp_col.value()
            dt_col   = self.spin_dt_col.value()
            dt_fmt   = self.edit_dt_fmt.text().strip()
            src      = os.path.basename(self.file_path)
            from_d   = self.date_from.date().toPyDate()
            to_d     = self.date_to.date().toPyDate()

            lines = self._read_lines(self.cmb_encoding.currentText())
            if not lines:
                QMessageBox.critical(self, "خطأ", "فشل قراءة الملف.")
                return

            ok = dup = 0
            errors = []

            # ← الإصلاح: استخدام db.conn المشترك لا اتصال جديد
            cur = self.db.conn.cursor()

            for i, line in enumerate(lines[skip:]):
                line = line.strip()
                if not line:
                    continue
                parts = line.split(delim)
                if len(parts) <= max(fp_col, dt_col):
                    errors.append(f"سطر {i+skip+1}: عدد أعمدة غير كافٍ")
                    continue
                fp_id  = parts[fp_col].strip()
                dt_str = parts[dt_col].strip()
                if not fp_id or not dt_str:
                    errors.append(f"سطر {i+skip+1}: قيمة فارغة")
                    continue
                try:
                    dt_obj = datetime.strptime(dt_str, dt_fmt)
                except Exception as ex:
                    errors.append(f"سطر {i+skip+1}: صيغة تاريخ خاطئة ({dt_str})")
                    continue

                if not (from_d <= dt_obj.date() <= to_d):
                    continue

                dt_iso = dt_obj.strftime("%Y-%m-%d %H:%M:%S")
                cur.execute(
                    "SELECT 1 FROM fingerprint_raw "
                    "WHERE fingerprint_id=? AND punch_datetime=?",
                    (fp_id, dt_iso))
                if cur.fetchone():
                    dup += 1
                    continue
                cur.execute(
                    "INSERT INTO fingerprint_raw "
                    "(fingerprint_id, punch_datetime, source_file) "
                    "VALUES (?,?,?)",
                    (fp_id, dt_iso, src))
                ok += 1
                if ok % 500 == 0:
                    self.db.conn.commit()

            self.db.conn.commit()

            msg = (f"✅ استُورد: {ok} سجل\n"
                   f"🔁 مكرر: {dup}\n"
                   f"⚠️ أخطاء: {len(errors)}")
            if errors:
                msg += "\n\nأول 10 أخطاء:\n" + "\n".join(errors[:10])
            QMessageBox.information(self, "نتيجة الاستيراد", msg)
            self.accept()

        except Exception as e:
            logger.error("خطأ في استيراد البصمات: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ غير متوقع", str(e))


# ==================== عامل معالجة البصمات ====================
class FingerprintWorker(QThread):
    progress = pyqtSignal(int, int)
    finished = pyqtSignal(int, int, int, set, dict)
    error    = pyqtSignal(str)

    def __init__(self, db_path: str, settings: dict, break_settings: dict):
        super().__init__()
        self.db_path       = db_path
        self.settings      = settings
        self.break_settings = break_settings

    def run(self):
        try:
            result = self._process()
            self.finished.emit(*result)
        except Exception as e:
            self.error.emit(str(e))

    def _process(self):
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON")
        conn.execute("PRAGMA journal_mode = WAL")
        cur  = conn.cursor()

        work_start  = self.settings['work_start']
        work_end    = self.settings['work_end']
        work_hours  = self.settings['work_hours']

        lunch_break    = self.break_settings.get('lunch_break', 30)
        num_breaks     = self.break_settings.get('num_breaks', 2)
        break_duration = self.break_settings.get('break_duration', 15)
        include_breaks = self.break_settings.get('include_breaks', 0)
        total_break_h  = (lunch_break + num_breaks * break_duration) / 60.0

        ws_time = datetime.strptime(work_start, "%H:%M").time()
        we_time = datetime.strptime(work_end,   "%H:%M").time()
        CLOSE   = 10   # دقائق — عتبة البصمتين المتقاربتين

        cur.execute("""
            SELECT fingerprint_id, DATE(punch_datetime),
                   GROUP_CONCAT(TIME(punch_datetime), ' ')
            FROM (
                SELECT fingerprint_id, punch_datetime
                FROM fingerprint_raw
                WHERE processed = 0
                ORDER BY fingerprint_id, punch_datetime
            )
            GROUP BY fingerprint_id, DATE(punch_datetime)
        """)
        raw = cur.fetchall()

        cur.execute(
            "SELECT id, fingerprint_id FROM employees WHERE status='نشط'")
        emp_map = {}
        emp_map_stripped = {}
        for eid, fid in cur.fetchall():
            if fid:
                emp_map[fid.strip()] = eid
                s = fid.strip().lstrip('0') or '0'
                emp_map_stripped[s] = eid

        exempt = {r[0] for r in cur.execute(
            "SELECT id FROM employees "
            "WHERE is_exempt_from_fingerprint=1").fetchall()}

        def find_emp(fp):
            fp = fp.strip()
            if fp in emp_map:
                return emp_map[fp]
            s = fp.lstrip('0') or '0'
            return emp_map_stripped.get(s)

        cur.execute(
            "SELECT employee_id, start_date, end_date "
            "FROM leave_requests WHERE status='موافق'")
        # جلب الإجازات الساعية المعتمدة
        cur.execute(
            "SELECT employee_id, leave_date, from_time, to_time "
            "FROM hourly_leaves WHERE status='موافق'")
        hourly_leaves = {}
        for emp_id, d, f_t, t_t in cur.fetchall():
            hourly_leaves.setdefault((emp_id, d), []).append((f_t, t_t))

        leave_days = {}
        for emp_id, s_str, e_str in cur.fetchall():
            s, e = date.fromisoformat(s_str), date.fromisoformat(e_str)
            for d in _date_range(s, e):
                leave_days[(emp_id, d.isoformat())] = True

        def mins(t1, t2):
            def tm(t): return t.hour * 60 + t.minute + t.second / 60.0
            return abs(tm(t1) - tm(t2))

        total = len(raw)
        ok = fail = no_emp = 0
        unmatched      = set()
        leave_conflicts = {}

        for idx, (fp_id, punch_date, times_str) in enumerate(raw):
            if idx % 50 == 0:
                self.progress.emit(idx, total)

            try:
                emp_id = find_emp(fp_id)
                if not emp_id:
                    no_emp += 1
                    unmatched.add(fp_id)
                    continue
                if emp_id in exempt:
                    continue

                times = sorted(times_str.split()) if times_str else []
                if not times:
                    fail += 1
                    continue

                check_in = check_out = None
                notes    = ""
                status   = "حاضر"

                if len(times) == 2:
                    t1 = datetime.strptime(times[0], "%H:%M:%S").time()
                    t2 = datetime.strptime(times[1], "%H:%M:%S").time()
                    if mins(t1, t2) <= CLOSE:
                        if mins(t1, ws_time) <= 30:
                            check_in = times[0]
                            notes = "بصمتان متقاربتان عند الدخول"
                        else:
                            check_out = times[1]
                            notes = "بصمتان متقاربتان عند الخروج"
                    else:
                        check_in  = times[0] if t1 < t2 else times[1]
                        check_out = times[1] if t1 < t2 else times[0]

                elif len(times) == 1:
                    t = datetime.strptime(times[0], "%H:%M:%S").time()
                    if mins(t, ws_time) <= mins(t, we_time):
                        check_in = times[0]
                    else:
                        check_out = times[0]
                    status = "نصف يوم"
                    notes  = "بصمة واحدة"

                else:
                    check_in  = times[0]
                    check_out = times[-1]
                    ignored = times[1:-1]
                    notes = (
                        f"تم تجاهل {len(ignored)} بصمة: "
                        + ", ".join(ignored)
                    )

                    t_i = datetime.strptime(check_in,  "%H:%M:%S").time()
                    t_o = datetime.strptime(check_out, "%H:%M:%S").time()
                    if t_i > t_o:
                        check_in, check_out = check_out, check_in

                work_h = ot_h = 0.0
                late_min = early_min = 0

                if check_in and check_out:
                    try:
                        ti = datetime.strptime(
                            f"{punch_date} {check_in}", "%Y-%m-%d %H:%M:%S")
                        to = datetime.strptime(
                            f"{punch_date} {check_out}", "%Y-%m-%d %H:%M:%S")
                        diff = (to - ti).total_seconds() / 3600
                        if include_breaks == 0 and diff > total_break_h:
                            diff -= total_break_h
                        work_h = min(diff, work_hours)
                        ot_h   = max(0, diff - work_hours) 
                    except Exception:
                        pass
                # تحديد الحالة النهائية بناءً على ساعات العمل
                if (emp_id, punch_date) in leave_days:
                    status = "إجازة"
                else:
                    if work_h <= 0:
                        status = "غائب"
                    elif work_h >= work_hours:
                        status = "حاضر"
                    else:
                        status = "نصف يوم"

                if check_in:
                    try:
                        ai = datetime.strptime(
                            f"{punch_date} {check_in[:5]}", "%Y-%m-%d %H:%M")
                        ei = datetime.strptime(
                            f"{punch_date} {work_start}", "%Y-%m-%d %H:%M")
                        ls = (ai - ei).total_seconds()
                        if ls > 0:
                            late_min = int(ls / 60)
                        # خصم الإجازة الساعية من التأخير
                        for f_t, t_t in hourly_leaves.get((emp_id, punch_date), []):
                            lf = datetime.strptime(
                                f"{punch_date} {f_t}", "%Y-%m-%d %H:%M")
                            lt = datetime.strptime(
                                f"{punch_date} {t_t}", "%Y-%m-%d %H:%M")
                            covered = max(
                                0, (lt - max(ai, lf)).total_seconds() / 60)
                            late_min = max(0, late_min - int(covered))

                    except Exception:
                        pass

                if check_out:
                    try:
                        ao = datetime.strptime(
                            f"{punch_date} {check_out[:5]}", "%Y-%m-%d %H:%M")
                        eo = datetime.strptime(
                            f"{punch_date} {work_end}", "%Y-%m-%d %H:%M")
                        es = (eo - ao).total_seconds()
                        if es > 0:
                            early_min = int(es / 60)
                        # خصم الإجازة الساعية من الخروج المبكر
                        for f_t, t_t in hourly_leaves.get((emp_id, punch_date), []):
                            lf = datetime.strptime(
                                f"{punch_date} {f_t}", "%Y-%m-%d %H:%M")
                            lt = datetime.strptime(
                                f"{punch_date} {t_t}", "%Y-%m-%d %H:%M")
                            covered = max(
                                0, (min(eo, lt) - lf).total_seconds() / 60)
                            early_min = max(0, early_min - int(covered))

                    except Exception:
                        pass

                if (emp_id, punch_date) in leave_days:
                    status = "إجازة"
                    work_h = ot_h = late_min = early_min = 0
                    check_in = check_out = None
                    emp_n = cur.execute(
                        "SELECT first_name||' '||last_name "
                        "FROM employees WHERE id=?", (emp_id,)).fetchone()
                    name = emp_n[0] if emp_n else str(emp_id)
                    leave_conflicts.setdefault(name, []).append(punch_date)
                    notes = "تحويل إلى إجازة"

                cur.execute(
                    "SELECT id FROM attendance "
                    "WHERE employee_id=? AND punch_date=?",
                    (emp_id, punch_date))
                existing = cur.fetchone()

                if existing:
                    cur.execute("""
                        UPDATE attendance
                        SET check_in=?, check_out=?, work_hours=?,
                            overtime_hours=?, late_minutes=?,
                            early_leave_minutes=?, status=?, notes=?
                        WHERE id=?
                    """, (check_in, check_out,
                          round(work_h, 2), round(ot_h, 2),
                          late_min, early_min, status, notes,
                          existing[0]))
                else:
                    cur.execute("""
                        INSERT INTO attendance
                        (employee_id, fingerprint_id, punch_date,
                         check_in, check_out, work_hours, overtime_hours,
                         late_minutes, early_leave_minutes, status, notes)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?)
                    """, (emp_id, fp_id, punch_date,
                          check_in, check_out,
                          round(work_h, 2), round(ot_h, 2),
                          late_min, early_min, status, notes))

                cur.execute(
                    "UPDATE fingerprint_raw SET processed=1 "
                    "WHERE fingerprint_id=? AND DATE(punch_datetime)=?",
                    (fp_id, punch_date))
                ok += 1

            except Exception as ex:
                logger.error("خطأ في سجل %s %s: %s", fp_id, punch_date, ex)
                fail += 1

            if idx % 200 == 0:
                conn.commit()

        conn.commit()
        conn.close()
        return ok, fail, no_emp, unmatched, leave_conflicts


# ==================== نافذة الإدخال اليدوي ====================
class ManualAttendanceDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, record=None):
        super().__init__(parent)
        self.db     = db
        self.record = record
        self.setWindowTitle(
            "إدخال حضور يدوي" if not record else "تعديل سجل حضور")
        self.setFixedSize(450, 450)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()
        if record:
            self._load_record()

    def _build(self):
        lay  = QVBoxLayout(self)
        form = QFormLayout()

        self.emp_combo = QComboBox()
        for eid, name in self.db.fetch_all(
                "SELECT id, first_name||' '||last_name "
                "FROM employees WHERE status='نشط'"):
            self.emp_combo.addItem(name, eid)
        self.emp_combo.setEditable(True)
        self.emp_combo.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.emp_combo.completer().setFilterMode(Qt.MatchContains)

        self.punch_date = QDateEdit()
        self.punch_date.setCalendarPopup(True)
        self.punch_date.setDate(QDate.currentDate())

        self.check_in  = QTimeEdit()
        self.check_in.setTime(QTime(8, 0))
        self.check_in.setDisplayFormat("HH:mm")

        self.check_out = QTimeEdit()
        self.check_out.setTime(QTime(17, 0))
        self.check_out.setDisplayFormat("HH:mm")

        self.status = QComboBox()
        self.status.addItems(["حاضر", "غائب", "إجازة", "نصف يوم"])

        self.notes = QLineEdit()
        self.notes.setPlaceholderText("اختياري")

        form.addRow("الموظف:",      self.emp_combo)
        form.addRow("التاريخ:",     self.punch_date)
        form.addRow("وقت الدخول:", self.check_in)
        form.addRow("وقت الخروج:", self.check_out)
        form.addRow("الحالة:",      self.status)
        form.addRow("ملاحظة:",      self.notes)
        lay.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _load_record(self):
        idx = self.emp_combo.findData(self.record[1])
        if idx >= 0:
            self.emp_combo.setCurrentIndex(idx)
        if self.record[3]:
            self.punch_date.setDate(
                QDate.fromString(str(self.record[3]), Qt.ISODate))
        if self.record[4]:
            self.check_in.setTime(QTime.fromString(str(self.record[4]), "HH:mm"))
        if self.record[5]:
            self.check_out.setTime(QTime.fromString(str(self.record[5]), "HH:mm"))
        self.status.setCurrentText(self.record[10] or "حاضر")
        self.notes.setText(self.record[11] or "")

    def _save(self):
        emp_id = self.emp_combo.currentData()
        pdate  = self.punch_date.date().toString(Qt.ISODate)
        cin    = (self.check_in.time().toString("HH:mm")
                  if self.check_in.time() != QTime(0, 0) else None)
        cout   = (self.check_out.time().toString("HH:mm")
                  if self.check_out.time() != QTime(0, 0) else None)
        status = self.status.currentText()
        notes  = self.notes.text().strip()

        # التحقق من تعارض الإجازة
        if status == "غائب" and not self.record:
            on_leave = self.db.fetch_one(
                """SELECT id FROM leave_requests
                   WHERE employee_id=? AND status='موافق'
                     AND start_date<=? AND end_date>=?""",
                (emp_id, pdate, pdate))
            if on_leave:
                QMessageBox.warning(
                    self, "خطأ",
                    "لا يمكن تسجيل غياب — الموظف في إجازة معتمدة.")
                return

        work_hours  = float(self.db.get_setting('working_hours', '8'))
        work_start  = self.db.get_setting('work_start_time', '08:00')
        work_end    = self.db.get_setting('work_end_time',   '17:00')

        work_h = ot_h = 0.0
        late_min = early_min = 0

        if status == "غائب":
            cin = cout = None
        elif cin and cout:
            try:
                t1    = datetime.strptime(f"{pdate} {cin}",  "%Y-%m-%d %H:%M")
                t2    = datetime.strptime(f"{pdate} {cout}", "%Y-%m-%d %H:%M")
                diff  = (t2 - t1).total_seconds() / 3600
                work_h = min(diff, work_hours)
                ot_h   = max(0, diff - work_hours)
            except Exception:
                pass

        if cin:
            try:
                ai = datetime.strptime(f"{pdate} {cin}",        "%Y-%m-%d %H:%M")
                ei = datetime.strptime(f"{pdate} {work_start}", "%Y-%m-%d %H:%M")
                ls = (ai - ei).total_seconds()
                if ls > 0:
                    late_min = int(ls / 60)
            except Exception:
                pass

        if cout:
            try:
                ao = datetime.strptime(f"{pdate} {cout}",     "%Y-%m-%d %H:%M")
                eo = datetime.strptime(f"{pdate} {work_end}", "%Y-%m-%d %H:%M")
                es = (eo - ao).total_seconds()
                if es > 0:
                    early_min = int(es / 60)
            except Exception:
                pass

        if self.record:
            self.db.execute_query(
                """UPDATE attendance
                   SET employee_id=?, punch_date=?, check_in=?, check_out=?,
                       work_hours=?, overtime_hours=?, status=?, notes=?,
                       late_minutes=?, early_leave_minutes=?
                   WHERE id=?""",
                (emp_id, pdate, cin, cout,
                 round(work_h, 2), round(ot_h, 2),
                 status, notes, late_min, early_min, self.record[0]))
        else:
            existing = self.db.fetch_one(
                "SELECT id FROM attendance "
                "WHERE employee_id=? AND punch_date=?", (emp_id, pdate))
            if existing:
                if QMessageBox.question(
                        self, "تأكيد",
                        "يوجد سجل لهذا الموظف في نفس التاريخ. "
                        "هل تريد تحديثه؟",
                        QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
                    return
                self.db.execute_query(
                    """UPDATE attendance
                       SET check_in=?, check_out=?, work_hours=?,
                           overtime_hours=?, status=?, notes=?,
                           late_minutes=?, early_leave_minutes=?
                       WHERE id=?""",
                    (cin, cout, round(work_h, 2), round(ot_h, 2),
                     status, notes, late_min, early_min, existing[0]))
            else:
                self.db.execute_query(
                    """INSERT INTO attendance
                       (employee_id, punch_date, check_in, check_out,
                        work_hours, overtime_hours, status, notes,
                        late_minutes, early_leave_minutes)
                       VALUES (?,?,?,?,?,?,?,?,?,?)""",
                    (emp_id, pdate, cin, cout,
                     round(work_h, 2), round(ot_h, 2),
                     status, notes, late_min, early_min))
        self.accept()


# ==================== نافذة التعديل الجماعي ====================
class BulkEditAttendanceDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager,
                 record_ids: list, settings: dict):
        super().__init__(parent)
        self.db         = db
        self.record_ids = record_ids
        self.settings   = settings
        self.setWindowTitle("تعديل جماعي للحضور")
        self.setFixedSize(500, 500)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)

        lbl = QLabel(f"عدد السجلات المحددة: {len(self.record_ids)}")
        lbl.setStyleSheet("font-weight:bold; padding:5px; background:#e0e0e0;")
        lay.addWidget(lbl)

        group = QGroupBox("خيارات التعديل")
        form  = QFormLayout(group)
        form.setSpacing(10)

        self.chk_in  = QCheckBox("تغيير وقت الدخول")
        self.in_edit = QTimeEdit()
        self.in_edit.setTime(QTime(8, 0))
        self.in_edit.setDisplayFormat("HH:mm")
        self.in_edit.setEnabled(False)
        self.chk_in.toggled.connect(self.in_edit.setEnabled)
        form.addRow(self.chk_in, self.in_edit)

        self.chk_out  = QCheckBox("تغيير وقت الخروج")
        self.out_edit = QTimeEdit()
        self.out_edit.setTime(QTime(17, 0))
        self.out_edit.setDisplayFormat("HH:mm")
        self.out_edit.setEnabled(False)
        self.chk_out.toggled.connect(self.out_edit.setEnabled)
        form.addRow(self.chk_out, self.out_edit)

        self.chk_status = QCheckBox("تغيير الحالة")
        self.status_cmb = QComboBox()
        self.status_cmb.addItems(["لا تغيير", "حاضر", "غائب", "إجازة", "نصف يوم"])
        self.status_cmb.setEnabled(False)
        self.chk_status.toggled.connect(self.status_cmb.setEnabled)
        form.addRow(self.chk_status, self.status_cmb)

        self.chk_notes  = QCheckBox("تغيير الملاحظات")
        self.notes_edit = QLineEdit()
        self.notes_edit.setEnabled(False)
        self.chk_notes.toggled.connect(self.notes_edit.setEnabled)
        form.addRow(self.chk_notes, self.notes_edit)

        lay.addWidget(group)

        pg   = QGroupBox("معاينة")
        pl   = QVBoxLayout(pg)
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        self.preview.setMaximumHeight(80)
        pl.addWidget(self.preview)
        lay.addWidget(pg)

        for sig in (self.chk_in.toggled, self.chk_out.toggled,
                    self.chk_status.toggled, self.chk_notes.toggled,
                    self.in_edit.timeChanged, self.out_edit.timeChanged,
                    self.status_cmb.currentTextChanged,
                    self.notes_edit.textChanged):
            sig.connect(self._update_preview)
        self._update_preview()

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _update_preview(self):
        lines = []
        if self.chk_in.isChecked():
            lines.append(f"وقت الدخول → {self.in_edit.time().toString('HH:mm')}")
        if self.chk_out.isChecked():
            lines.append(f"وقت الخروج → {self.out_edit.time().toString('HH:mm')}")
        if self.chk_status.isChecked() and self.status_cmb.currentText() != "لا تغيير":
            lines.append(f"الحالة → {self.status_cmb.currentText()}")
        if self.chk_notes.isChecked() and self.notes_edit.text():
            lines.append(f"الملاحظات → {self.notes_edit.text()}")
        self.preview.setText("\n".join(lines) if lines else "لم يتم تحديد تغييرات")

    def _save(self):
        updates = {}
        recalc  = False

        if self.chk_in.isChecked():
            updates['check_in']  = self.in_edit.time().toString("HH:mm")
            recalc = True
        if self.chk_out.isChecked():
            updates['check_out'] = self.out_edit.time().toString("HH:mm")
            recalc = True
        if self.chk_status.isChecked() and self.status_cmb.currentText() != "لا تغيير":
            updates['status'] = self.status_cmb.currentText()
        if self.chk_notes.isChecked() and self.notes_edit.text():
            updates['notes'] = self.notes_edit.text()

        if not updates:
            QMessageBox.warning(self, "تحذير", "لم يتم تحديد أي تغييرات")
            return

        wh    = self.settings['work_hours']
        ws    = self.settings['work_start']
        we    = self.settings['work_end']

        for rid in self.record_ids:
            rec = self.db.fetch_one(
                "SELECT employee_id, punch_date, check_in, check_out "
                "FROM attendance WHERE id=?", (rid,))
            if not rec:
                continue
            emp_id, pdate, old_in, old_out = rec
            new_in  = updates.get('check_in',  old_in)
            new_out = updates.get('check_out', old_out)

            row_updates = dict(updates)
            extra = []

            if recalc:
                work_h = ot_h = 0.0
                late_m = early_m = 0
                if new_in and new_out:
                    try:
                        t1    = datetime.strptime(f"{pdate} {new_in}",  "%Y-%m-%d %H:%M")
                        t2    = datetime.strptime(f"{pdate} {new_out}", "%Y-%m-%d %H:%M")
                        diff  = (t2 - t1).total_seconds() / 3600
                        work_h = min(diff, wh)
                        ot_h   = max(0, diff - wh)
                    except Exception:
                        pass
                if new_in:filter_row.addWidget(self.draft_status_filter)
                    try:
                        ai = datetime.strptime(f"{pdate} {new_in[:5]}", "%Y-%m-%d %H:%M")
                        ei = datetime.strptime(f"{pdate} {ws}",         "%Y-%m-%d %H:%M")
                        ls = (ai - ei).total_seconds()
                        if ls > 0:
                            late_m = int(ls / 60)
                    except Exception:
                        pass
                if new_out:
                    try:
                        ao = datetime.strptime(f"{pdate} {new_out[:5]}", "%Y-%m-%d %H:%M")
                        eo = datetime.strptime(f"{pdate} {we}",          "%Y-%m-%d %H:%M")
                        es = (eo - ao).total_seconds()
                        if es > 0:
                            early_m = int(es / 60)
                    except Exception:
                        pass
                if 'status' not in row_updates:
                    if new_in and new_out:
                        row_updates['status'] = "حاضر"
                    elif new_in or new_out:
                        row_updates['status'] = "نصف يوم"
                    else:
                        row_updates['status'] = "غائب"
                extra = [round(work_h,2), round(ot_h,2), late_m, early_m]
                extra_cols = ", work_hours=?, overtime_hours=?, late_minutes=?, early_leave_minutes=?"
            else:
                extra_cols = ""

            set_clause = ", ".join(f"{k}=?" for k in row_updates)
            values     = list(row_updates.values()) + extra + [rid]
            self.db.execute_query(
                f"UPDATE attendance SET {set_clause}{extra_cols} WHERE id=?",
                values)

        QMessageBox.information(
            self, "نجاح",
            f"تم تحديث {len(self.record_ids)} سجل بنجاح")
        self.accept()
