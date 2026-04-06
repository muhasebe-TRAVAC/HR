#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/settings_tab.py

import os
import shutil
import logging
import bcrypt
from datetime import date

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QComboBox,
    QSpinBox, QDoubleSpinBox, QTimeEdit, QPushButton, QMessageBox,
    QFileDialog, QListWidget, QListWidgetItem, QFormLayout, QGroupBox,
    QTabWidget, QInputDialog, QDialog, QDialogButtonBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QDateEdit, QCompleter
)
from PyQt5.QtCore import Qt, QTime, QDate
from PyQt5.QtGui import QPixmap

from database import DatabaseManager
from utils import (make_table, fill_table, btn,
                   can_add, can_edit, can_delete, can_manage_users)
from constants import (BTN_SUCCESS, BTN_PRIMARY, BTN_DANGER,
                       BTN_GRAY, BTN_WARNING,
                       BACKUP_FOLDER, COMPANY_LOGO_FOLDER)

logger = logging.getLogger(__name__)


class SettingsTab(QWidget):
    """
    تبويب الإعدادات.

    الإصلاحات:
    - ChangeMyPasswordDialog: يتحقق من كلمة المرور القديمة قبل السماح بالتغيير.
    - _save_language(): استخدام QApplication.exit(0) بدلاً من استدعاء quit و exit معاً.
    - إضافة logging.
    """

    def __init__(self, db: DatabaseManager, user: dict, comm=None):
        super().__init__()
        self.db        = db
        self.user      = user
        self.comm      = comm
        self.logo_path = None
        self._build()
        self._load()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    def _on_data_changed(self, data_type: str, data):
        if data_type == 'department':
            self._load_depts()
        elif data_type == 'user':
            self._load_users()

    def _build(self):
        layout = QVBoxLayout(self)
        tabs   = QTabWidget()
        tabs.addTab(self._build_company_tab(),        "🏢 بيانات الشركة")
        tabs.addTab(self._build_work_tab(),           "⏰ إعدادات العمل")
        tabs.addTab(self._build_leave_settings_tab(), "🏖️ إعدادات الإجازات")
        tabs.addTab(self._build_holidays_tab(),       "📅 العطل الرسمية")
        tabs.addTab(self._build_users_tab(),          "👤 المستخدمون")
        tabs.addTab(self._build_dept_tab(),           "🏛️ الأقسام")
        tabs.addTab(self._build_backup_tab(),         "💾 النسخ الاحتياطي")
        layout.addWidget(tabs)
        self.inner_tabs = tabs

    # ==================== تبويب بيانات الشركة ====================
    def _build_company_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        logo_group  = QGroupBox("شعار الشركة")
        logo_layout = QHBoxLayout()
        self.logo_label = QLabel()
        self.logo_label.setFixedSize(150, 150)
        self.logo_label.setStyleSheet(
            "border: 1px solid #ccc; background: #f9f9f9;")
        self.logo_label.setAlignment(Qt.AlignCenter)
        self.logo_label.setText("لا يوجد شعار")
        logo_layout.addWidget(self.logo_label)

        logo_btns = QVBoxLayout()
        self.btn_choose_logo = btn("📂 اختيار شعار", BTN_PRIMARY, self._choose_logo)
        self.btn_remove_logo = btn("🗑️ إزالة الشعار", BTN_DANGER,  self._remove_logo)
        logo_btns.addWidget(self.btn_choose_logo)
        logo_btns.addWidget(self.btn_remove_logo)
        logo_btns.addStretch()
        logo_layout.addLayout(logo_btns)
        logo_group.setLayout(logo_layout)
        lay.addWidget(logo_group)

        info_group = QGroupBox("بيانات الشركة")
        form       = QFormLayout()
        form.setSpacing(10)
        self.company_name    = QLineEdit()
        self.company_address = QLineEdit()
        self.company_phone   = QLineEdit()
        self.currency        = QComboBox()
        self.currency.setEditable(True)
        self.currency.addItems([
            "ريال سعودي (SAR)", "درهم إماراتي (AED)", "دينار كويتي (KWD)",
            "دينار بحريني (BHD)", "ريال عماني (OMR)", "ريال قطري (QAR)",
            "دولار أمريكي (USD)", "يورو (EUR)", "جنيه مصري (EGP)",
            "دينار أردني (JOD)", "جنيه إسترليني (GBP)", "ليرة تركية (TRY)",
            "دينار عراقي (IQD)", "دينار ليبي (LYD)", "درهم مغربي (MAD)",
            "أوقية موريتانية (MRU)", "جنيه سوداني (SDG)", "ليرة سورية (SYP)",
        ])
        self.currency.setCurrentText("ليرة تركية (TRY)")
        form.addRow("اسم الشركة:", self.company_name)
        form.addRow("العنوان:",    self.company_address)
        form.addRow("الهاتف:",    self.company_phone)
        form.addRow("العملة:",    self.currency)
        info_group.setLayout(form)
        lay.addWidget(info_group)

        # إعدادات اللغة
        lang_group  = QGroupBox("اللغة / Dil")
        lang_layout = QHBoxLayout()
        self.lang_combo = QComboBox()
        self.lang_combo.addItem("العربية", "ar")
        self.lang_combo.addItem("Türkçe",  "tr")
        current_lang = self.db.get_setting('language', 'ar')
        idx = self.lang_combo.findData(current_lang)
        if idx >= 0:
            self.lang_combo.setCurrentIndex(idx)
        lang_layout.addWidget(QLabel("اختر اللغة / Dil Seçin:"))
        lang_layout.addWidget(self.lang_combo)
        lang_layout.addStretch()
        lang_group.setLayout(lang_layout)
        lay.addWidget(lang_group)

        self.btn_save_lang    = btn(
            "حفظ اللغة / Dili Kaydet", BTN_SUCCESS, self._save_language)
        self.btn_save_company = btn(
            "💾 حفظ بيانات الشركة", BTN_SUCCESS, self._save_company)
        lay.addWidget(self.btn_save_lang)
        lay.addWidget(self.btn_save_company)
        lay.addStretch()

        self._apply_company_permissions()
        return w

    def _apply_company_permissions(self):
        can = self.user['role'] in ('admin', 'hr')
        self.btn_choose_logo.setVisible(can)
        self.btn_remove_logo.setVisible(can)
        self.btn_save_company.setVisible(can)

    def _save_language(self):
        """
        حفظ اللغة وإعادة تشغيل التطبيق.
        الإصلاح: استخدام QApplication.exit(0) مرة واحدة فقط.
        """
        lang = self.lang_combo.currentData()
        self.db.set_setting('language', lang)
        import constants
        constants.CURRENT_LANG = lang
        QMessageBox.information(
            self, "تغيير اللغة",
            "تم حفظ اللغة. سيتم إعادة تشغيل البرنامج لتطبيق التغيير.\n"
            "Dil kaydedildi. Değişikliği uygulamak için program yeniden başlatılacak.")
        from PyQt5.QtWidgets import QApplication
        QApplication.exit(0)

    # ==================== تبويب إعدادات العمل ====================
    def _build_work_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        main_group = QGroupBox("إعدادات الدوام والرواتب")
        form       = QFormLayout()
        form.setSpacing(10)

        self.work_start = QTimeEdit()
        self.work_start.setTime(QTime(8, 0))
        self.work_start.setDisplayFormat("HH:mm")
        self.work_end = QTimeEdit()
        self.work_end.setTime(QTime(17, 0))
        self.work_end.setDisplayFormat("HH:mm")
        self.work_start.timeChanged.connect(self._calc_work_hours)
        self.work_end.timeChanged.connect(self._calc_work_hours)

        self.lunch_break = QSpinBox()
        self.lunch_break.setRange(0, 180)
        self.lunch_break.setSuffix(" دقيقة")
        self.lunch_break.setValue(30)
        self.lunch_break.valueChanged.connect(self._calc_work_hours)

        self.num_breaks = QSpinBox()
        self.num_breaks.setRange(0, 10)
        self.num_breaks.setValue(2)
        self.num_breaks.valueChanged.connect(self._calc_work_hours)

        self.break_duration = QSpinBox()
        self.break_duration.setRange(0, 60)
        self.break_duration.setSuffix(" دقيقة")
        self.break_duration.setValue(15)
        self.break_duration.valueChanged.connect(self._calc_work_hours)

        self.include_breaks = QComboBox()
        self.include_breaks.addItems(["نعم (شامل)", "لا (غير شامل)"])
        self.include_breaks.currentIndexChanged.connect(self._calc_work_hours)

        self.work_hours = QDoubleSpinBox()
        self.work_hours.setRange(1, 24)
        self.work_hours.setReadOnly(True)
        self.work_hours.setSuffix(" ساعة")

        self.work_days_list = QListWidget()
        self.work_days_list.setSelectionMode(QAbstractItemView.MultiSelection)
        days = ["الإثنين", "الثلاثاء", "الأربعاء", "الخميس",
                "الجمعة", "السبت", "الأحد"]
        for day in days:
            item = QListWidgetItem(day)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(
                Qt.Checked if day in ["الإثنين", "الثلاثاء", "الأربعاء",
                                       "الخميس", "الجمعة"]
                else Qt.Unchecked)
            self.work_days_list.addItem(item)
        self.work_days_list.itemChanged.connect(self._calc_weekly_hours)

        self.weekly_work_hours = QDoubleSpinBox()
        self.weekly_work_hours.setRange(0, 168)
        self.weekly_work_hours.setReadOnly(True)
        self.weekly_work_hours.setSuffix(" ساعة")

        self.work_days_month = QSpinBox()
        self.work_days_month.setRange(20, 31)
        self.work_days_month.setReadOnly(True)
        self.work_days_month.setSuffix(" يوم")

        self.late_tol      = QSpinBox()
        self.late_tol.setRange(0, 120)
        self.late_tol.setSuffix(" دقيقة")
        self.late_tol_type = QComboBox()
        self.late_tol_type.addItems(["يومي", "شهري"])

        self.ot_rate = QDoubleSpinBox()
        self.ot_rate.setRange(1, 3)
        self.ot_rate.setSingleStep(0.25)
        self.ot_rate.setSuffix(" x")

        self.absence_deduction_rate = QDoubleSpinBox()
        self.absence_deduction_rate.setRange(0, 3)
        self.absence_deduction_rate.setSingleStep(0.25)
        self.absence_deduction_rate.setValue(1.0)
        self.absence_deduction_rate.setSuffix(" x")

        self.gosi_emp_pct = QDoubleSpinBox()
        self.gosi_emp_pct.setRange(0, 25)
        self.gosi_emp_pct.setSuffix("%")
        self.gosi_co_pct = QDoubleSpinBox()
        self.gosi_co_pct.setRange(0, 25)
        self.gosi_co_pct.setSuffix("%")

        for label, widget in [
            ("وقت بداية الدوام:",              self.work_start),
            ("وقت نهاية الدوام:",              self.work_end),
            ("فترة الغداء:",                   self.lunch_break),
            ("عدد الاستراحات:",                self.num_breaks),
            ("مدة كل استراحة:",               self.break_duration),
            ("الاستراحات ضمن ساعات العمل?:", self.include_breaks),
            ("ساعات العمل اليومية (محسوبة):", self.work_hours),
            ("أيام العمل الأسبوعية:",         self.work_days_list),
            ("ساعات العمل الأسبوعية (محسوبة):", self.weekly_work_hours),
            ("أيام العمل في الشهر (محسوبة):", self.work_days_month),
            ("سماحية التأخير:",               self.late_tol),
            ("طبيعة السماحية:",               self.late_tol_type),
            ("معدل الأوفرتايم:",              self.ot_rate),
            ("معدل الخصم على الغياب:",        self.absence_deduction_rate),
            ("نسبة تأمين الموظف:",            self.gosi_emp_pct),
            ("نسبة تأمين الشركة:",            self.gosi_co_pct),
        ]:
            form.addRow(label, widget)

        main_group.setLayout(form)
        lay.addWidget(main_group)
        self.btn_save_work = btn("💾 حفظ إعدادات العمل", BTN_SUCCESS, self._save_work)
        lay.addWidget(self.btn_save_work)
        lay.addStretch()
        self._apply_work_permissions()
        return w

    def _apply_work_permissions(self):
        self.btn_save_work.setVisible(self.user['role'] in ('admin', 'hr'))

    def _calc_work_hours(self):
        s = self.work_start.time()
        e = self.work_end.time()
        total = (e.hour() * 60 + e.minute()) - (s.hour() * 60 + s.minute())
        if total <= 0:
            total = 0
        breaks = self.lunch_break.value() + (
            self.num_breaks.value() * self.break_duration.value())
        if self.include_breaks.currentIndex() == 1:
            total = max(0, total - breaks)
        self.work_hours.setValue(round(total / 60.0, 2))
        self._calc_weekly_hours()

    def _calc_weekly_hours(self):
        daily      = self.work_hours.value()
        work_days  = sum(
            1 for i in range(self.work_days_list.count())
            if self.work_days_list.item(i).checkState() == Qt.Checked)
        weekly     = daily * work_days
        self.weekly_work_hours.setValue(round(weekly, 2))
        self.work_days_month.setValue(round(work_days * 4.33))

    # ==================== تبويب إعدادات الإجازات ====================
    def _build_leave_settings_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        group       = QGroupBox("قواعد استحقاق الإجازة السنوية")
        group_layout = QVBoxLayout()
        self.leave_rules_table = QTableWidget()
        self.leave_rules_table.setColumnCount(4)
        self.leave_rules_table.setHorizontalHeaderLabels(
            ["من سنة", "إلى سنة", "عدد الأيام", "ملاحظات"])
        self.leave_rules_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)

        btn_layout = QHBoxLayout()
        self.btn_add_rule    = btn("➕ إضافة قاعدة", BTN_SUCCESS, self._add_leave_rule)
        self.btn_edit_rule   = btn("✏️ تعديل قاعدة", BTN_PRIMARY, self._edit_leave_rule)
        self.btn_delete_rule = btn("🗑️ حذف قاعدة",  BTN_DANGER,  self._delete_leave_rule)
        for b in (self.btn_add_rule, self.btn_edit_rule, self.btn_delete_rule):
            btn_layout.addWidget(b)

        group_layout.addWidget(self.leave_rules_table)
        group_layout.addLayout(btn_layout)
        group.setLayout(group_layout)
        lay.addWidget(group)

        gen_group = QGroupBox("إعدادات عامة")
        gen_form  = QFormLayout()
        self.max_leave_days = QSpinBox()
        self.max_leave_days.setRange(0, 365)
        self.max_leave_days.setValue(30)
        gen_form.addRow("الحد الأقصى لترحيل الإجازات:", self.max_leave_days)
        gen_group.setLayout(gen_form)
        lay.addWidget(gen_group)

        self.btn_save_leave = btn(
            "💾 حفظ إعدادات الإجازات", BTN_SUCCESS, self._save_leave_settings)
        lay.addWidget(self.btn_save_leave)
        lay.addStretch()
        self._apply_leave_permissions()
        return w

    def _apply_leave_permissions(self):
        can = self.user['role'] in ('admin', 'hr')
        self.btn_add_rule.setVisible(can)
        self.btn_edit_rule.setVisible(can)
        self.btn_delete_rule.setVisible(can)
        self.btn_save_leave.setVisible(can)

    def _add_leave_rule(self):
        from_year, ok = QInputDialog.getInt(
            self, "إضافة قاعدة", "من سنة (مدة الخدمة):", 0, 0, 50)
        if not ok: return
        to_year, ok = QInputDialog.getInt(
            self, "إضافة قاعدة", "إلى سنة:", 5, 0, 50)
        if not ok: return
        days, ok = QInputDialog.getInt(
            self, "إضافة قاعدة", "عدد أيام الإجازة:", 21, 0, 365)
        if not ok: return
        notes, _ = QInputDialog.getText(
            self, "إضافة قاعدة", "ملاحظات (اختياري):")
        r = self.leave_rules_table.rowCount()
        self.leave_rules_table.insertRow(r)
        for c, v in enumerate([str(from_year), str(to_year), str(days), notes]):
            self.leave_rules_table.setItem(r, c, QTableWidgetItem(v))

    def _edit_leave_rule(self):
        row = self.leave_rules_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "تنبيه", "اختر قاعدة أولاً")
            return
        vals = [self.leave_rules_table.item(row, c).text() for c in range(4)]
        from_year, ok = QInputDialog.getInt(
            self, "تعديل", "من سنة:", int(vals[0]), 0, 50)
        if not ok: return
        to_year, ok = QInputDialog.getInt(
            self, "تعديل", "إلى سنة:", int(vals[1]), 0, 50)
        if not ok: return
        days, ok = QInputDialog.getInt(
            self, "تعديل", "عدد الأيام:", int(vals[2]), 0, 365)
        if not ok: return
        notes, _ = QInputDialog.getText(
            self, "تعديل", "ملاحظات:", text=vals[3])
        for c, v in enumerate([str(from_year), str(to_year), str(days), notes]):
            self.leave_rules_table.setItem(row, c, QTableWidgetItem(v))

    def _delete_leave_rule(self):
        row = self.leave_rules_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "تنبيه", "اختر قاعدة أولاً")
            return
        self.leave_rules_table.removeRow(row)

    def _save_leave_settings(self):
        self.db.execute_query("DELETE FROM leave_rules")
        for row in range(self.leave_rules_table.rowCount()):
            vals = [self.leave_rules_table.item(row, c).text()
                    for c in range(4)]
            self.db.execute_query(
                "INSERT INTO leave_rules (from_year, to_year, days, notes) "
                "VALUES (?,?,?,?)",
                (int(vals[0]), int(vals[1]), int(vals[2]), vals[3] or None))
        self.db.execute_query(
            "INSERT OR REPLACE INTO settings (setting_name, setting_value) "
            "VALUES (?,?)",
            ('max_leave_carryover', str(self.max_leave_days.value())))
        self.db.log_custom("تحديث إعدادات الإجازات", "settings")
        QMessageBox.information(self, "نجاح", "تم حفظ إعدادات الإجازات")
        if self.comm:
            self.comm.dataChanged.emit('settings', {'type': 'leave_rules'})

    # ==================== تبويب العطل الرسمية ====================
    def _build_holidays_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        # عطل ثابتة
        fg   = QGroupBox("العطل السنوية الثابتة (تتكرر كل سنة)")
        fl   = QVBoxLayout()
        self.fixed_holidays_table = QTableWidget()
        self.fixed_holidays_table.setColumnCount(3)
        self.fixed_holidays_table.setHorizontalHeaderLabels(
            ["الاسم", "التاريخ (MM-DD)", "ملاحظات"])
        self.fixed_holidays_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        fbl = QHBoxLayout()
        self.btn_add_fixed    = btn("➕ إضافة عطلة ثابتة", BTN_SUCCESS,
                                    lambda: self._add_holiday('fixed'))
        self.btn_edit_fixed   = btn("✏️ تعديل", BTN_PRIMARY,
                                    lambda: self._edit_holiday('fixed'))
        self.btn_delete_fixed = btn("🗑️ حذف", BTN_DANGER,
                                    lambda: self._delete_holiday('fixed'))
        for b in (self.btn_add_fixed, self.btn_edit_fixed, self.btn_delete_fixed):
            fbl.addWidget(b)
        fl.addWidget(self.fixed_holidays_table)
        fl.addLayout(fbl)
        fg.setLayout(fl)
        lay.addWidget(fg)

        # عطل متغيرة
        vg  = QGroupBox("العطل الدينية المتغيرة (حسب السنة)")
        vl  = QVBoxLayout()
        self.var_holidays_table = QTableWidget()
        self.var_holidays_table.setColumnCount(4)
        self.var_holidays_table.setHorizontalHeaderLabels(
            ["الاسم", "التاريخ", "السنة", "ملاحظات"])
        self.var_holidays_table.horizontalHeader().setSectionResizeMode(
            QHeaderView.Stretch)
        vbl = QHBoxLayout()
        self.btn_add_var    = btn("➕ إضافة عطلة متغيرة", BTN_SUCCESS,
                                  lambda: self._add_holiday('variable'))
        self.btn_edit_var   = btn("✏️ تعديل", BTN_PRIMARY,
                                  lambda: self._edit_holiday('variable'))
        self.btn_delete_var = btn("🗑️ حذف", BTN_DANGER,
                                  lambda: self._delete_holiday('variable'))
        for b in (self.btn_add_var, self.btn_edit_var, self.btn_delete_var):
            vbl.addWidget(b)
        vl.addWidget(self.var_holidays_table)
        vl.addLayout(vbl)
        vg.setLayout(vl)
        lay.addWidget(vg)

        self.btn_save_holidays = btn(
            "💾 حفظ العطل", BTN_SUCCESS, self._save_holidays)
        lay.addWidget(self.btn_save_holidays)
        lay.addStretch()
        self._apply_holidays_permissions()
        return w

    def _apply_holidays_permissions(self):
        can_e = self.user['role'] in ('admin', 'hr')
        can_d = can_delete(self.user['role'])
        self.btn_add_fixed.setVisible(can_e)
        self.btn_edit_fixed.setVisible(can_e)
        self.btn_delete_fixed.setVisible(can_d)
        self.btn_add_var.setVisible(can_e)
        self.btn_edit_var.setVisible(can_e)
        self.btn_delete_var.setVisible(can_d)
        self.btn_save_holidays.setVisible(can_e)

    def _add_holiday(self, holiday_type: str):
        if holiday_type == 'fixed':
            name, ok = QInputDialog.getText(self, "عطلة ثابتة", "اسم العطلة:")
            if not ok or not name: return
            date_str, ok = QInputDialog.getText(
                self, "عطلة ثابتة", "التاريخ (MM-DD):", text="01-01")
            if not ok or not date_str: return
            notes, _ = QInputDialog.getText(
                self, "عطلة ثابتة", "ملاحظات (اختياري):")
            r = self.fixed_holidays_table.rowCount()
            self.fixed_holidays_table.insertRow(r)
            for c, v in enumerate([name, date_str, notes]):
                self.fixed_holidays_table.setItem(r, c, QTableWidgetItem(v))
        else:
            name, ok = QInputDialog.getText(self, "عطلة متغيرة", "اسم العطلة:")
            if not ok or not name: return
            year, ok = QInputDialog.getInt(
                self, "عطلة متغيرة", "السنة:", date.today().year, 2000, 2100)
            if not ok: return
            month, ok = QInputDialog.getInt(
                self, "عطلة متغيرة", "الشهر:", 1, 1, 12)
            if not ok: return
            day, ok = QInputDialog.getInt(
                self, "عطلة متغيرة", "اليوم:", 1, 1, 31)
            if not ok: return
            notes, _ = QInputDialog.getText(
                self, "عطلة متغيرة", "ملاحظات (اختياري):")
            r = self.var_holidays_table.rowCount()
            self.var_holidays_table.insertRow(r)
            for c, v in enumerate([name,
                                    f"{year:04d}-{month:02d}-{day:02d}",
                                    str(year), notes]):
                self.var_holidays_table.setItem(r, c, QTableWidgetItem(v))

    def _edit_holiday(self, holiday_type: str):
        table = (self.fixed_holidays_table if holiday_type == 'fixed'
                 else self.var_holidays_table)
        row   = table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "تنبيه", "اختر عطلة أولاً")
            return

        if holiday_type == 'fixed':
            vals    = [table.item(row, c).text() for c in range(3)]
            name,ok = QInputDialog.getText(self, "تعديل", "الاسم:", text=vals[0])
            if not ok or not name: return
            dt, ok  = QInputDialog.getText(
                self, "تعديل", "التاريخ (MM-DD):", text=vals[1])
            if not ok or not dt: return
            notes,_ = QInputDialog.getText(
                self, "تعديل", "ملاحظات:", text=vals[2])
            for c, v in enumerate([name, dt, notes]):
                table.setItem(row, c, QTableWidgetItem(v))
        else:
            vals     = [table.item(row, c).text() for c in range(4)]
            name, ok = QInputDialog.getText(self, "تعديل", "الاسم:", text=vals[0])
            if not ok or not name: return
            y, ok    = QInputDialog.getInt(
                self, "تعديل", "السنة:", int(vals[2]), 2000, 2100)
            if not ok: return
            mo, ok   = QInputDialog.getInt(
                self, "تعديل", "الشهر:",
                int(vals[1][5:7]) if len(vals[1]) >= 7 else 1, 1, 12)
            if not ok: return
            d, ok    = QInputDialog.getInt(
                self, "تعديل", "اليوم:",
                int(vals[1][8:10]) if len(vals[1]) >= 10 else 1, 1, 31)
            if not ok: return
            notes, _ = QInputDialog.getText(
                self, "تعديل", "ملاحظات:", text=vals[3])
            for c, v in enumerate([name,
                                    f"{y:04d}-{mo:02d}-{d:02d}",
                                    str(y), notes]):
                table.setItem(row, c, QTableWidgetItem(v))

    def _delete_holiday(self, holiday_type: str):
        table = (self.fixed_holidays_table if holiday_type == 'fixed'
                 else self.var_holidays_table)
        row   = table.currentRow()
        if row >= 0:
            table.removeRow(row)

    def _save_holidays(self):
        self.db.execute_query("DELETE FROM holidays")
        for row in range(self.fixed_holidays_table.rowCount()):
            vals  = [self.fixed_holidays_table.item(row, c).text()
                     for c in range(3)]
            self.db.execute_query(
                "INSERT INTO holidays (name, holiday_date, type, notes) "
                "VALUES (?,?,?,?)",
                (vals[0], vals[1], 'fixed', vals[2] or None))
        for row in range(self.var_holidays_table.rowCount()):
            vals  = [self.var_holidays_table.item(row, c).text()
                     for c in range(4)]
            self.db.execute_query(
                "INSERT INTO holidays (name, holiday_date, type, year, notes) "
                "VALUES (?,?,?,?,?)",
                (vals[0], vals[1], 'variable', int(vals[2]), vals[3] or None))
        self.db.log_custom("تحديث العطل الرسمية", "settings")
        QMessageBox.information(self, "نجاح", "تم حفظ العطل")
        if self.comm:
            self.comm.dataChanged.emit('settings', {'type': 'holidays'})

    # ==================== تبويب المستخدمين ====================
    def _build_users_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        tools = QHBoxLayout()
        self.btn_new_user     = btn("➕ مستخدم جديد",            BTN_SUCCESS, self._new_user)
        self.btn_change_pw    = btn("🔑 تغيير كلمة المرور",       BTN_WARNING, self._change_pw)
        self.btn_change_my_pw = btn("🔐 تغيير كلمة مروري",        BTN_PRIMARY, self._change_my_password)
        for b in (self.btn_new_user, self.btn_change_pw, self.btn_change_my_pw):
            tools.addWidget(b)
        tools.addStretch()
        lay.addLayout(tools)

        self.users_table = make_table(
            ["#", "اسم المستخدم", "الاسم الكامل", "الصلاحية", "الحالة"])
        lay.addWidget(self.users_table)
        self._load_users()
        self._apply_users_permissions()
        return w

    def _apply_users_permissions(self):
        can = can_manage_users(self.user['role'])
        self.btn_new_user.setVisible(can)
        self.btn_change_pw.setVisible(can)

    def _load_users(self):
        fill_table(self.users_table, self.db.fetch_all(
            "SELECT id, username, full_name, role, "
            "CASE is_active WHEN 1 THEN 'نشط' ELSE 'موقوف' END FROM users"))

    def _new_user(self):
        dlg = NewUserDialog(self, self.db)
        if dlg.exec_() == QDialog.Accepted:
            self._load_users()
            if self.comm:
                self.comm.dataChanged.emit('user', {'action': 'add'})

    def _change_pw(self):
        row = self.users_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "خطأ", "اختر مستخدماً")
            return
        uid = int(self.users_table.item(row, 0).text())
        dlg = ChangePasswordDialog(self, self.db, uid)
        if dlg.exec_() == QDialog.Accepted:
            self._load_users()
            if self.comm:
                self.comm.dataChanged.emit(
                    'user', {'action': 'change_password', 'id': uid})

    def _change_my_password(self):
        """
        الإصلاح: يستخدم ChangeMyPasswordDialog الذي يتحقق من
        كلمة المرور القديمة قبل السماح بالتغيير.
        """
        dlg = ChangeMyPasswordDialog(self, self.db, self.user['id'])
        if dlg.exec_() == QDialog.Accepted:
            QMessageBox.information(self, "نجاح", "تم تغيير كلمة المرور بنجاح")
            if self.comm:
                self.comm.dataChanged.emit(
                    'user', {'action': 'change_my_password',
                             'id': self.user['id']})

    # ==================== تبويب الأقسام ====================
    def _build_dept_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        tools = QHBoxLayout()
        self.btn_new_dept    = btn("➕ قسم جديد", BTN_SUCCESS, self._new_dept)
        self.btn_delete_dept = btn("🗑️ حذف",     BTN_DANGER,  self._del_dept)
        tools.addWidget(self.btn_new_dept)
        tools.addWidget(self.btn_delete_dept)
        tools.addStretch()
        lay.addLayout(tools)

        self.dept_table = make_table(["#", "اسم القسم", "ملاحظات"])
        lay.addWidget(self.dept_table)
        self._load_depts()
        self._apply_dept_permissions()
        return w

    def _apply_dept_permissions(self):
        self.btn_new_dept.setVisible(can_add(self.user['role']))
        self.btn_delete_dept.setVisible(can_delete(self.user['role']))

    def _load_depts(self):
        fill_table(self.dept_table, self.db.fetch_all(
            "SELECT id, name, notes FROM departments ORDER BY name"))

    def _new_dept(self):
        name, ok = QInputDialog.getText(self, "قسم جديد", "اسم القسم:")
        if ok and name:
            self.db.execute_query(
                "INSERT OR IGNORE INTO departments (name) VALUES (?)", (name,))
            self.db.log_insert("departments", self.db.last_id(), {"name": name})
            self._load_depts()
            if self.comm:
                self.comm.dataChanged.emit('department', {'action': 'add', 'name': name})

    def _del_dept(self):
        row = self.dept_table.currentRow()
        if row < 0:
            return
        did       = int(self.dept_table.item(row, 0).text())
        dept_name = self.dept_table.item(row, 1).text()

        emp_count = self.db.fetch_one(
            "SELECT COUNT(*) FROM employees WHERE department_id=?", (did,))
        if emp_count and emp_count[0] > 0:
            QMessageBox.warning(
                self, "تنبيه",
                f"لا يمكن حذف هذا القسم لأنه مرتبط بـ {emp_count[0]} موظف.\n"
                "يرجى نقل الموظفين إلى قسم آخر أولاً.")
            return

        if QMessageBox.question(
                self, "تأكيد", "حذف القسم؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self.db.execute_query(
                "DELETE FROM departments WHERE id=?", (did,))
            self.db.log_delete("departments", did, {"name": dept_name})
            self._load_depts()
            if self.comm:
                self.comm.dataChanged.emit(
                    'department', {'action': 'delete', 'id': did})

    # ==================== تبويب النسخ الاحتياطي ====================
    def _build_backup_tab(self) -> QWidget:
        w   = QWidget()
        lay = QVBoxLayout(w)

        g     = QGroupBox("النسخ الاحتياطي والاستعادة")
        g_lay = QVBoxLayout()
        self.btn_backup      = btn("💾 إنشاء نسخة احتياطية الآن", BTN_PRIMARY, self._backup)
        self.btn_open_backup = btn("📂 فتح مجلد النسخ الاحتياطي", BTN_GRAY,   self._open_backup_folder)
        self.btn_restore     = btn("🔄 استعادة نسخة احتياطية",    BTN_WARNING, self._restore_backup)
        for b in (self.btn_backup, self.btn_open_backup, self.btn_restore):
            g_lay.addWidget(b)
        g.setLayout(g_lay)
        lay.addWidget(g)

        lay.addWidget(QLabel("النسخ الاحتياطية المتوفرة:"))
        self.backup_list = QListWidget()
        self._load_backups()
        lay.addWidget(self.backup_list)
        lay.addStretch()
        self._apply_backup_permissions()
        return w

    def _apply_backup_permissions(self):
        can = self.user['role'] == 'admin'
        self.btn_backup.setVisible(can)
        self.btn_restore.setVisible(can)

    def _backup(self):
        path = self.db.backup()
        if path:
            QMessageBox.information(self, "نجاح", f"تم إنشاء نسخة احتياطية:\n{path}")
            self._load_backups()
        else:
            QMessageBox.critical(self, "خطأ", "فشل في إنشاء النسخة الاحتياطية")

    def _restore_backup(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "اختر ملف النسخة الاحتياطية",
            BACKUP_FOLDER, "Database (*.db)")
        if not path:
            return
        if QMessageBox.question(
                self, "تأكيد",
                "سيتم استبدال قاعدة البيانات الحالية بالنسخة الاحتياطية.\n"
                "هل أنت متأكد؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        try:
            self.db.conn.close()
            shutil.copy2(path, self.db.db_name)
            self.db.connect()
            QMessageBox.information(
                self, "نجاح",
                "تمت استعادة النسخة الاحتياطية بنجاح.\n"
                "سيتم إعادة تشغيل البرنامج.")
        except Exception as e:
            logger.error("خطأ في استعادة النسخة: %s", e, exc_info=True)
            QMessageBox.critical(self, "خطأ", f"فشل في استعادة النسخة:\n{e}")

    def _open_backup_folder(self):
        import subprocess
        import platform
        if platform.system() == 'Windows':
            os.startfile(BACKUP_FOLDER)
        elif platform.system() == 'Darwin':
            subprocess.call(['open', BACKUP_FOLDER])
        else:
            subprocess.call(['xdg-open', BACKUP_FOLDER])

    def _load_backups(self):
        self.backup_list.clear()
        try:
            for f in sorted(os.listdir(BACKUP_FOLDER), reverse=True):
                if f.endswith('.db'):
                    size = os.path.getsize(
                        os.path.join(BACKUP_FOLDER, f)) // 1024
                    self.backup_list.addItem(f"💾 {f}  ({size} KB)")
        except Exception:
            pass

    # ==================== تحميل وحفظ الإعدادات ====================
    def _load(self):
        self.company_name.setText(self.db.get_setting('company_name', ''))
        self.company_address.setText(self.db.get_setting('company_address', ''))
        self.company_phone.setText(self.db.get_setting('company_phone', ''))

        saved_currency = self.db.get_setting('currency', 'ليرة تركية (TRY)')
        idx = self.currency.findText(saved_currency)
        self.currency.setCurrentIndex(idx if idx >= 0 else 0)
        if idx < 0:
            self.currency.setCurrentText(saved_currency)

        logo_path = self.db.get_setting('company_logo', '')
        if logo_path and os.path.exists(logo_path):
            px = QPixmap(logo_path)
            if not px.isNull():
                self.logo_label.setPixmap(
                    px.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
                self.logo_path = logo_path
            else:
                self.logo_label.setText("لا يوجد شعار")
        else:
            self.logo_label.setText("لا يوجد شعار")

        self.work_start.setTime(QTime.fromString(
            self.db.get_setting('work_start_time', '08:00'), "HH:mm"))
        self.work_end.setTime(QTime.fromString(
            self.db.get_setting('work_end_time', '17:00'), "HH:mm"))
        self.lunch_break.setValue(int(self.db.get_setting('lunch_break', '30')))
        self.num_breaks.setValue(int(self.db.get_setting('num_breaks', '2')))
        self.break_duration.setValue(int(self.db.get_setting('break_duration', '15')))
        self.include_breaks.setCurrentIndex(
            int(self.db.get_setting('include_breaks', '0')))
        self._calc_work_hours()

        days_str = self.db.get_setting('work_days', '0,1,2,3,4')
        checked  = {int(x) for x in days_str.split(',') if x.strip().isdigit()}
        for i in range(self.work_days_list.count()):
            self.work_days_list.item(i).setCheckState(
                Qt.Checked if i in checked else Qt.Unchecked)

        self.late_tol.setValue(int(self.db.get_setting('late_tolerance_minutes', '10')))
        self.late_tol_type.setCurrentIndex(
            int(self.db.get_setting('late_tolerance_type', '0')))
        self.ot_rate.setValue(float(self.db.get_setting('overtime_rate', '1.5')))
        self.absence_deduction_rate.setValue(
            float(self.db.get_setting('absence_deduction_rate', '1.0')))
        self.gosi_emp_pct.setValue(
            float(self.db.get_setting('gosi_employee_percent', '9.75')))
        self.gosi_co_pct.setValue(
            float(self.db.get_setting('gosi_company_percent', '12.0')))

        self._load_leave_rules()
        self._load_holidays()

        # تحميل الحد الأقصى لترحيل الإجازات
        self.max_leave_days.setValue(
            int(self.db.get_setting('max_leave_carryover', '30')))

    def _load_leave_rules(self):
        self.leave_rules_table.setRowCount(0)
        for from_y, to_y, days, notes in self.db.fetch_all(
                "SELECT from_year, to_year, days, notes "
                "FROM leave_rules ORDER BY from_year"):
            r = self.leave_rules_table.rowCount()
            self.leave_rules_table.insertRow(r)
            for c, v in enumerate([str(from_y), str(to_y),
                                    str(days), notes or ""]):
                self.leave_rules_table.setItem(r, c, QTableWidgetItem(v))

    def _load_holidays(self):
        self.fixed_holidays_table.setRowCount(0)
        self.var_holidays_table.setRowCount(0)
        for name, hdate, typ, year, notes in self.db.fetch_all(
                "SELECT name, holiday_date, type, year, notes FROM holidays"):
            if typ == 'fixed':
                r = self.fixed_holidays_table.rowCount()
                self.fixed_holidays_table.insertRow(r)
                for c, v in enumerate([name, hdate, notes or ""]):
                    self.fixed_holidays_table.setItem(r, c, QTableWidgetItem(v))
            else:
                r = self.var_holidays_table.rowCount()
                self.var_holidays_table.insertRow(r)
                for c, v in enumerate([name, hdate, str(year or ''), notes or ""]):
                    self.var_holidays_table.setItem(r, c, QTableWidgetItem(v))

    def _save_company(self):
        self.db.set_setting('company_name',    self.company_name.text())
        self.db.set_setting('company_address', self.company_address.text())
        self.db.set_setting('company_phone',   self.company_phone.text())
        self.db.set_setting('currency',        self.currency.currentText())
        if self.logo_path:
            self.db.set_setting('company_logo', self.logo_path)
        self.db.log_custom("تحديث بيانات الشركة", "settings")
        if self.comm:
            self.comm.dataChanged.emit('settings', {'type': 'company'})
        QMessageBox.information(self, "نجاح", "تم حفظ بيانات الشركة")

    def _save_work(self):
        settings = {
            'work_start_time':        self.work_start.time().toString("HH:mm"),
            'work_end_time':          self.work_end.time().toString("HH:mm"),
            'lunch_break':            str(self.lunch_break.value()),
            'num_breaks':             str(self.num_breaks.value()),
            'break_duration':         str(self.break_duration.value()),
            'include_breaks':         str(self.include_breaks.currentIndex()),
            'working_hours':          str(self.work_hours.value()),
            'late_tolerance_minutes': str(self.late_tol.value()),
            'late_tolerance_type':    str(self.late_tol_type.currentIndex()),
            'overtime_rate':          str(self.ot_rate.value()),
            'absence_deduction_rate': str(self.absence_deduction_rate.value()),
            'gosi_employee_percent':  str(self.gosi_emp_pct.value()),
            'gosi_company_percent':   str(self.gosi_co_pct.value()),
            'work_days_month':        str(self.work_days_month.value()),
            'work_days':              ','.join(
                str(i) for i in range(self.work_days_list.count())
                if self.work_days_list.item(i).checkState() == Qt.Checked),
        }
        for k, v in settings.items():
            self.db.set_setting(k, v)
        self.db.log_custom("تحديث إعدادات العمل", "settings")
        if self.comm:
            self.comm.dataChanged.emit('settings', {'type': 'work'})
        QMessageBox.information(self, "نجاح", "تم حفظ إعدادات العمل")

    def _choose_logo(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "اختر شعار الشركة", "",
            "Images (*.png *.jpg *.jpeg *.bmp)")
        if not path:
            return
        px = QPixmap(path)
        if px.isNull():
            QMessageBox.warning(self, "خطأ", "الملف المختار ليس صورة صالحة")
            return
        os.makedirs(COMPANY_LOGO_FOLDER, exist_ok=True)
        ext      = os.path.splitext(path)[1]
        new_path = os.path.join(COMPANY_LOGO_FOLDER, f"company_logo{ext}")
        shutil.copy2(path, new_path)
        self.logo_label.setPixmap(
            px.scaled(150, 150, Qt.KeepAspectRatio, Qt.SmoothTransformation))
        self.logo_path = new_path

    def _remove_logo(self):
        self.logo_label.clear()
        self.logo_label.setText("لا يوجد شعار")
        if self.logo_path and os.path.exists(self.logo_path):
            try:
                os.remove(self.logo_path)
            except Exception:
                pass
        self.logo_path = None
        self.db.set_setting('company_logo', '')


# ===================================================================
# كلاسات الحوارات
# ===================================================================
class NewUserDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("مستخدم جديد")
        self.setFixedSize(360, 300)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()

    def _build(self):
        lay  = QVBoxLayout(self)
        form = QFormLayout()
        form.setSpacing(10)

        self.uname = QLineEdit()
        self.fname = QLineEdit()
        self.pw    = QLineEdit()
        self.pw.setEchoMode(QLineEdit.Password)
        self.pw2   = QLineEdit()
        self.pw2.setEchoMode(QLineEdit.Password)
        self.role  = QComboBox()
        self.role.addItems(["admin", "hr", "accountant", "viewer"])

        form.addRow("اسم المستخدم:",      self.uname)
        form.addRow("الاسم الكامل:",      self.fname)
        form.addRow("كلمة المرور:",       self.pw)
        form.addRow("تأكيد كلمة المرور:", self.pw2)
        form.addRow("الصلاحية:",          self.role)
        lay.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _save(self):
        if not self.uname.text() or not self.fname.text() or not self.pw.text():
            QMessageBox.warning(self, "خطأ", "جميع الحقول مطلوبة")
            return
        if self.pw.text() != self.pw2.text():
            QMessageBox.warning(self, "خطأ", "كلمتا المرور غير متطابقتين")
            return
        pw_hash = bcrypt.hashpw(
            self.pw.text().encode(), bcrypt.gensalt()).decode()
        if self.db.execute_query(
                "INSERT INTO users (username, password_hash, full_name, role) "
                "VALUES (?,?,?,?)",
                (self.uname.text(), pw_hash,
                 self.fname.text(), self.role.currentText())):
            self.db.log_insert("users", self.db.last_id(), {
                "username":  self.uname.text(),
                "full_name": self.fname.text(),
                "role":      self.role.currentText()})
            self.accept()
        else:
            QMessageBox.critical(self, "خطأ", "اسم المستخدم موجود مسبقاً")


class ChangePasswordDialog(QDialog):
    """تغيير كلمة مرور أي مستخدم (للمدير)."""

    def __init__(self, parent, db: DatabaseManager, user_id: int):
        super().__init__(parent)
        self.db      = db
        self.user_id = user_id
        self.setWindowTitle("تغيير كلمة المرور")
        self.setFixedSize(350, 200)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()

    def _build(self):
        lay  = QVBoxLayout(self)
        form = QFormLayout()
        self.new_pw     = QLineEdit()
        self.new_pw.setEchoMode(QLineEdit.Password)
        self.confirm_pw = QLineEdit()
        self.confirm_pw.setEchoMode(QLineEdit.Password)
        form.addRow("كلمة المرور الجديدة:", self.new_pw)
        form.addRow("تأكيد:",               self.confirm_pw)
        lay.addLayout(form)
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _save(self):
        if not self.new_pw.text():
            QMessageBox.warning(self, "خطأ", "كلمة المرور مطلوبة")
            return
        if self.new_pw.text() != self.confirm_pw.text():
            QMessageBox.warning(self, "خطأ", "كلمتا المرور غير متطابقتين")
            return
        pw_hash = bcrypt.hashpw(
            self.new_pw.text().encode(), bcrypt.gensalt()).decode()
        self.db.execute_query(
            "UPDATE users SET password_hash=? WHERE id=?",
            (pw_hash, self.user_id))
        self.db.log_custom("تغيير كلمة مرور", "users", self.user_id)
        self.accept()


class ChangeMyPasswordDialog(QDialog):
    """
    تغيير كلمة مرور المستخدم الحالي — يتحقق من القديمة أولاً.

    الإصلاح: هذا النموذج يختلف عن ChangePasswordDialog لأنه يتطلب
    إدخال كلمة المرور الحالية قبل السماح بتغييرها، وهو إجراء أمني ضروري.
    """

    def __init__(self, parent, db: DatabaseManager, user_id: int):
        super().__init__(parent)
        self.db      = db
        self.user_id = user_id
        self.setWindowTitle("تغيير كلمة المرور الخاصة بي")
        self.setFixedSize(380, 240)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()

    def _build(self):
        lay  = QVBoxLayout(self)
        form = QFormLayout()
        form.setSpacing(10)

        self.old_pw  = QLineEdit()
        self.old_pw.setEchoMode(QLineEdit.Password)
        self.new_pw  = QLineEdit()
        self.new_pw.setEchoMode(QLineEdit.Password)
        self.conf_pw = QLineEdit()
        self.conf_pw.setEchoMode(QLineEdit.Password)

        form.addRow("كلمة المرور الحالية:", self.old_pw)
        form.addRow("كلمة المرور الجديدة:", self.new_pw)
        form.addRow("تأكيد كلمة المرور:",  self.conf_pw)
        lay.addLayout(form)

        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._save)
        btns.rejected.connect(self.reject)
        lay.addWidget(btns)

    def _save(self):
        if not self.old_pw.text() or not self.new_pw.text():
            QMessageBox.warning(self, "خطأ", "جميع الحقول مطلوبة")
            return
        if self.new_pw.text() != self.conf_pw.text():
            QMessageBox.warning(self, "خطأ", "كلمتا المرور الجديدتان غير متطابقتين")
            return

        # التحقق من كلمة المرور القديمة
        row = self.db.fetch_one(
            "SELECT password_hash FROM users WHERE id=?", (self.user_id,))
        if not row:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على المستخدم")
            return
        if not bcrypt.checkpw(self.old_pw.text().encode(),
                               row[0].encode()):
            QMessageBox.warning(self, "خطأ", "كلمة المرور الحالية غير صحيحة")
            return

        pw_hash = bcrypt.hashpw(
            self.new_pw.text().encode(), bcrypt.gensalt()).decode()
        self.db.execute_query(
            "UPDATE users SET password_hash=? WHERE id=?",
            (pw_hash, self.user_id))
        self.db.log_custom("تغيير كلمة مروري", "users", self.user_id)
        self.accept()
