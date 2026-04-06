#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/audit_tab.py

import logging
from datetime import date

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QDateEdit,
    QPushButton, QMessageBox, QFileDialog, QGroupBox, QGridLayout,
    QDialog, QTextEdit, QDialogButtonBox, QSizePolicy
)
from PyQt5.QtCore import Qt, QDate

from database import DatabaseManager
from utils import make_table, fill_table, btn
from constants import BTN_PRIMARY, BTN_SUCCESS, BTN_GRAY

logger = logging.getLogger(__name__)

# الحد الأقصى للسجلات المعروضة في الجدول (لتجنب تجميد الواجهة)
MAX_DISPLAY_ROWS = 5_000


class AuditTab(QWidget):
    """
    تبويب سجل الأحداث — يعرض جميع العمليات المسجَّلة في audit_log.

    التحسينات في هذه النسخة:
    - استخراج بناء الاستعلام في دالة _build_query() مستقلة
      → لا تكرار بين _load() و _export_excel()
    - popup لعرض القيمة الكاملة عند النقر المزدوج على أي خلية
    - _on_data_changed لا تُعيد التحميل إلا لأنواع الأحداث المنطقية
    - حد أقصى MAX_DISPLAY_ROWS مع تحذير للمستخدم
    - تحميل الفلاتر في دالة موحدة _refresh_filter_combos()
    """

    def __init__(self, db: DatabaseManager, user: dict, comm=None):
        super().__init__()
        self.db   = db
        self.user = user
        self.comm = comm
        self._build()
        self._refresh_filter_combos()
        self._load()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    # ==================== بناء الواجهة ====================
    def _build(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        # --- مجموعة الفلاتر ---
        filter_group = QGroupBox("🔍 تصفية السجلات")
        filter_layout = QGridLayout()
        filter_layout.setSpacing(8)

        filter_layout.addWidget(QLabel("من تاريخ:"), 0, 0)
        self.date_from = QDateEdit()
        self.date_from.setCalendarPopup(True)
        self.date_from.setDate(QDate.currentDate().addDays(-30))
        filter_layout.addWidget(self.date_from, 0, 1)

        filter_layout.addWidget(QLabel("إلى تاريخ:"), 0, 2)
        self.date_to = QDateEdit()
        self.date_to.setCalendarPopup(True)
        self.date_to.setDate(QDate.currentDate())
        filter_layout.addWidget(self.date_to, 0, 3)

        filter_layout.addWidget(QLabel("المستخدم:"), 1, 0)
        self.user_filter = QComboBox()
        self.user_filter.setMinimumWidth(150)
        filter_layout.addWidget(self.user_filter, 1, 1)

        filter_layout.addWidget(QLabel("الإجراء:"), 1, 2)
        self.action_filter = QComboBox()
        filter_layout.addWidget(self.action_filter, 1, 3)

        filter_layout.addWidget(QLabel("الجدول:"), 2, 0)
        self.table_filter = QComboBox()
        filter_layout.addWidget(self.table_filter, 2, 1)

        self.search_btn = btn("🔍 بحث", BTN_PRIMARY, self._load)
        filter_layout.addWidget(self.search_btn, 2, 2)

        self.reset_btn = btn("🔄 إعادة تعيين", BTN_GRAY, self._reset_filters)
        filter_layout.addWidget(self.reset_btn, 2, 3)

        filter_group.setLayout(filter_layout)
        layout.addWidget(filter_group)

        # --- شريط الأزرار ---
        btn_row = QHBoxLayout()
        export_btn = btn("📥 تصدير Excel", BTN_SUCCESS, self._export_excel)
        btn_row.addWidget(export_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        # --- الجدول الرئيسي ---
        self.table = make_table([
            "التاريخ", "المستخدم", "الإجراء", "الجدول",
            "معرف السجل", "القيمة القديمة", "القيمة الجديدة"
        ])
        self.table.setSortingEnabled(False)
        # النقر المزدوج يفتح popup بالقيمة الكاملة
        self.table.cellDoubleClicked.connect(self._show_full_value)
        layout.addWidget(self.table)

        # --- شريط المعلومات السفلي ---
        info_row = QHBoxLayout()
        self.lbl_count = QLabel("عدد السجلات: 0")
        self.lbl_count.setStyleSheet(
            "font-weight:bold; padding:5px; background:#e3f2fd; border-radius:3px;")

        self.lbl_limit = QLabel("")
        self.lbl_limit.setStyleSheet("color:#D32F2F; font-weight:bold; padding:5px;")

        self.lbl_hint = QLabel("💡 انقر نقراً مزدوجاً على أي خلية لعرض القيمة الكاملة")
        self.lbl_hint.setStyleSheet("color:#666; font-size:11px; padding:5px;")

        info_row.addWidget(self.lbl_count)
        info_row.addWidget(self.lbl_limit)
        info_row.addStretch()
        info_row.addWidget(self.lbl_hint)
        layout.addLayout(info_row)

    # ==================== تحميل الفلاتر ====================
    def _refresh_filter_combos(self):
        """
        تحميل (أو إعادة تحميل) جميع قوائم الفلاتر من قاعدة البيانات.
        تُستدعى مرة عند البناء وعند الحاجة لتحديث الخيارات.
        """
        # فلتر المستخدمين
        current_user = self.user_filter.currentData()
        self.user_filter.clear()
        self.user_filter.addItem("جميع المستخدمين", None)
        for uid, uname in self.db.fetch_all(
                "SELECT id, username FROM users ORDER BY username"):
            self.user_filter.addItem(uname, uid)
        # استعادة الاختيار السابق إن وُجد
        idx = self.user_filter.findData(current_user)
        if idx >= 0:
            self.user_filter.setCurrentIndex(idx)

        # فلتر الإجراءات
        current_action = self.action_filter.currentData()
        self.action_filter.clear()
        self.action_filter.addItem("جميع الإجراءات", None)
        for (act,) in self.db.fetch_all(
                "SELECT DISTINCT action FROM audit_log "
                "WHERE action IS NOT NULL ORDER BY action"):
            self.action_filter.addItem(act, act)
        idx = self.action_filter.findData(current_action)
        if idx >= 0:
            self.action_filter.setCurrentIndex(idx)

        # فلتر الجداول
        current_table = self.table_filter.currentData()
        self.table_filter.clear()
        self.table_filter.addItem("جميع الجداول", None)
        for (tbl,) in self.db.fetch_all(
                "SELECT DISTINCT table_name FROM audit_log "
                "WHERE table_name IS NOT NULL ORDER BY table_name"):
            self.table_filter.addItem(tbl, tbl)
        idx = self.table_filter.findData(current_table)
        if idx >= 0:
            self.table_filter.setCurrentIndex(idx)

    def _reset_filters(self):
        """إعادة تعيين جميع الفلاتر إلى القيم الافتراضية."""
        self.date_from.setDate(QDate.currentDate().addDays(-30))
        self.date_to.setDate(QDate.currentDate())
        self.user_filter.setCurrentIndex(0)
        self.action_filter.setCurrentIndex(0)
        self.table_filter.setCurrentIndex(0)
        self._load()

    # ==================== بناء الاستعلام (مركزي) ====================
    def _build_query(self, for_export: bool = False):
        """
        بناء استعلام SQL وقائمة المعاملات بناءً على الفلاتر الحالية.

        هذه الدالة هي المصدر الوحيد للاستعلام — تُستخدَم في كل من
        _load() و _export_excel() لتفادي التكرار وضمان الاتساق.

        المعاملات:
            for_export: إذا True لا تُضيف LIMIT (لتصدير كامل البيانات)

        الإرجاع: (query_string, params_list)
        """
        d_from     = self.date_from.date().toString(Qt.ISODate)
        d_to       = self.date_to.date().toString(Qt.ISODate)
        user_id    = self.user_filter.currentData()
        action     = self.action_filter.currentData()
        table_name = self.table_filter.currentData()

        q = """
            SELECT a.created_at, u.username, a.action, a.table_name,
                   a.record_id, a.old_value, a.new_value
            FROM audit_log a
            LEFT JOIN users u ON a.user_id = u.id
            WHERE DATE(a.created_at) BETWEEN ? AND ?
        """
        params = [d_from, d_to]

        if user_id is not None:
            q += " AND a.user_id = ?"
            params.append(user_id)
        if action is not None:
            q += " AND a.action = ?"
            params.append(action)
        if table_name is not None:
            q += " AND a.table_name = ?"
            params.append(table_name)

        q += " ORDER BY a.created_at DESC"

        if not for_export:
            q += f" LIMIT {MAX_DISPLAY_ROWS}"

        return q, params

    # ==================== التحميل ====================
    def _load(self):
        """تحميل السجلات وعرضها في الجدول."""
        try:
            q, params = self._build_query(for_export=False)
            data      = self.db.fetch_all(q, params)

            display_data = []
            for row in data:
                # اختصار القيم لعرض أنظف في الجدول
                # (الضغط المزدوج يُظهر القيمة الكاملة)
                old_val = self._truncate(row[5])
                new_val = self._truncate(row[6])
                display_data.append((
                    row[0],
                    row[1] or "غير معروف",
                    row[2],
                    row[3] or "",
                    row[4] or "",
                    old_val,
                    new_val,
                ))

            fill_table(self.table, display_data)
            self.lbl_count.setText(f"عدد السجلات: {len(display_data)}")

            # تحذير إذا وصلنا للحد الأقصى
            if len(data) >= MAX_DISPLAY_ROWS:
                self.lbl_limit.setText(
                    f"⚠️ يُعرض {MAX_DISPLAY_ROWS:,} سجل كحد أقصى — "
                    f"صفِّ النتائج أو صدِّرها للاطلاع على الكامل")
                self.lbl_limit.show()
            else:
                self.lbl_limit.setText("")

        except Exception as e:
            logger.error("خطأ في تحميل سجل الأحداث: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ", f"فشل تحميل السجلات:\n{e}")

    @staticmethod
    def _truncate(value, max_len: int = 60) -> str:
        """اختصار النص الطويل لعرضه في الجدول."""
        if not value:
            return ""
        return (value[:max_len] + "…") if len(value) > max_len else value

    # ==================== Popup القيمة الكاملة ====================
    def _show_full_value(self, row: int, col: int):
        """
        عند النقر المزدوج على خلية في عمود القيمة القديمة أو الجديدة،
        تُفتح نافذة منبثقة تعرض النص كاملاً غير مُختصَر.
        """
        # عمود 5 = القيمة القديمة، عمود 6 = القيمة الجديدة
        if col not in (5, 6):
            return

        item = self.table.item(row, col)
        if not item:
            return

        # جلب النص الكامل مباشرة من قاعدة البيانات
        # (الجدول يعرض نصاً مُختصَراً فقط)
        date_item = self.table.item(row, 0)
        user_item = self.table.item(row, 1)
        act_item  = self.table.item(row, 2)

        full_data = self.db.fetch_one(
            """SELECT old_value, new_value
               FROM audit_log
               WHERE created_at = ?
                 AND action     = ?
               LIMIT 1""",
            (date_item.text() if date_item else "",
             act_item.text()  if act_item  else "")
        )

        col_label = "القيمة القديمة" if col == 5 else "القيمة الجديدة"
        full_text = ""
        if full_data:
            full_text = full_data[0] if col == 5 else full_data[1]
        full_text = full_text or "(فارغ)"

        dlg = _FullValueDialog(col_label, full_text, self)
        dlg.exec_()

    # ==================== التصدير ====================
    def _export_excel(self):
        """تصدير كامل السجلات المُصفَّاة (بدون حد) إلى ملف Excel."""
        try:
            import pandas as pd
        except ImportError:
            QMessageBox.critical(
                self, "خطأ",
                "مكتبة pandas غير مثبَّتة.\nنفِّذ: pip install pandas openpyxl")
            return

        try:
            q, params = self._build_query(for_export=True)
            data      = self.db.fetch_all(q, params)

            if not data:
                QMessageBox.warning(self, "تنبيه", "لا توجد بيانات للتصدير")
                return

            df = pd.DataFrame(data, columns=[
                "التاريخ", "المستخدم", "الإجراء", "الجدول",
                "معرف السجل", "القيمة القديمة", "القيمة الجديدة"
            ])

            filename = f"audit_log_{date.today().isoformat()}.xlsx"
            path, _  = QFileDialog.getSaveFileName(
                self, "حفظ الملف", filename, "Excel (*.xlsx)")
            if not path:
                return

            df.to_excel(path, index=False)
            QMessageBox.information(
                self, "نجاح", f"تم تصدير {len(df):,} سجل بنجاح")

        except Exception as e:
            logger.error("خطأ في تصدير سجل الأحداث: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ في التصدير", str(e))

    # ==================== التحديث التلقائي ====================
    def _on_data_changed(self, data_type: str, data):
        """
        الإصلاح: لا نُعيد تحميل الجدول عند كل حدث.
        نُعيد التحميل فقط عند الأحداث التي تُنتج سجلات جديدة فعلاً،
        أو عند تغيير المستخدمين (لتحديث فلتر المستخدمين).
        """
        # تحديث فلتر المستخدمين عند إضافة/تعديل مستخدم
        if data_type == 'user':
            self._refresh_filter_combos()

        # إعادة تحميل السجلات عند أي حدث يُنتج سجلاً في audit_log
        if data_type in (
            'employee', 'attendance', 'leave_request',
            'payroll', 'loan', 'user', 'settings',
        ):
            self._refresh_filter_combos()
            self._load()


# ==================== نافذة عرض القيمة الكاملة ====================
class _FullValueDialog(QDialog):
    """
    نافذة منبثقة بسيطة لعرض نص كامل (قيمة قديمة أو جديدة من audit_log).
    """

    def __init__(self, title: str, content: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"📋 {title}")
        self.setMinimumSize(500, 300)
        self.setLayoutDirection(Qt.RightToLeft)

        layout = QVBoxLayout(self)

        lbl = QLabel(f"<b>{title}:</b>")
        layout.addWidget(lbl)

        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setPlainText(content)
        text_edit.setStyleSheet(
            "font-family: monospace; font-size: 12px; background:#f9f9f9;")
        text_edit.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(text_edit)

        buttons = QDialogButtonBox(QDialogButtonBox.Close)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)