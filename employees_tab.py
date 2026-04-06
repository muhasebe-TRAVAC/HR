#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/employees_tab.py

import os
import re
import shutil
import logging
import subprocess
import platform
from datetime import date, datetime

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QGroupBox, QFormLayout,
    QLineEdit, QComboBox, QDateEdit, QDoubleSpinBox, QCheckBox, QTextEdit,
    QScrollArea, QMessageBox, QFileDialog, QListWidget, QListWidgetItem,
    QMenu, QCompleter
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QColor

from database import DatabaseManager
from utils import make_table, fill_table, btn, can_add, can_edit, can_delete
from constants import (BTN_SUCCESS, BTN_PRIMARY, BTN_DANGER,
                       BTN_GRAY, BTN_PURPLE, BTN_TEAL, DOCUMENTS_FOLDER)

logger = logging.getLogger(__name__)

# ================================================================
# استعلام SELECT صريح بأسماء الأعمدة — يُستخدَم في _select()
# هذا يحمي الكود من الانهيار عند إضافة أعمدة جديدة للجدول
# ================================================================
_EMP_SELECT = """
    SELECT id,
           employee_code, first_name, last_name, national_id, nationality,
           birth_date, gender, position, department_id,
           hire_date, contract_type,
           basic_salary, housing_allowance, transportation_allowance,
           food_allowance, phone_allowance, other_allowances,
           bank_salary, cash_salary,
           phone, email, address, bank_name, bank_account, iban,
           fingerprint_id,
           social_security_number, social_security_registered,
           social_security_date,
           iqama_expiry, iqama_number,
           passport_expiry, passport_number,
           health_insurance_expiry, health_insurance_number,
           status, notes,
           is_exempt_from_fingerprint
    FROM employees
    WHERE id = ?
"""
# ترتيب الأعمدة في النتيجة (للمرجعية)
_C = {
    'id':                         0,
    'employee_code':              1,
    'first_name':                 2,
    'last_name':                  3,
    'national_id':                4,
    'nationality':                5,
    'birth_date':                 6,
    'gender':                     7,
    'position':                   8,
    'department_id':              9,
    'hire_date':                  10,
    'contract_type':              11,
    'basic_salary':               12,
    'housing_allowance':          13,
    'transportation_allowance':   14,
    'food_allowance':             15,
    'phone_allowance':            16,
    'other_allowances':           17,
    'bank_salary':                18,
    'cash_salary':                19,
    'phone':                      20,
    'email':                      21,
    'address':                    22,
    'bank_name':                  23,
    'bank_account':               24,
    'iban':                       25,
    'fingerprint_id':             26,
    'social_security_number':     27,
    'social_security_registered': 28,
    'social_security_date':       29,
    'iqama_expiry':               30,
    'iqama_number':               31,
    'passport_expiry':            32,
    'passport_number':            33,
    'health_insurance_expiry':    34,
    'health_insurance_number':    35,
    'status':                     36,
    'notes':                      37,
    'is_exempt_from_fingerprint': 38,
}


class EmployeesTab(QWidget):
    """
    تبويب الموظفين.

    الإصلاحات في هذه النسخة:
    - _select() يستخدم استعلاماً صريحاً بأسماء الأعمدة بدلاً من SELECT *
      مع فهرس رقمي — يحمي من الانهيار عند أي تغيير في هيكل الجدول.
    - _delete() يستخدم with self.db.transaction() بدلاً من BEGIN يدوي،
      مما يضمن الـ rollback الصحيح عند أي خطأ.
    - إضافة تحقق من صحة البريد الإلكتروني في _validate().
    - استخدام logger بدلاً من print في _import_excel().
    """

    def __init__(self, db: DatabaseManager, user: dict, comm=None):
        super().__init__()
        self.db           = db
        self.user         = user
        self.comm         = comm
        self.current_id   = None
        self.currency     = self.db.get_setting('currency', 'TRY')
        self.work_days_month  = int(self.db.get_setting('work_days_month', '26'))
        self.work_hours_daily = float(self.db.get_setting('working_hours', '8'))
        self._build()
        self._load()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    def _on_data_changed(self, data_type: str, data):
        if data_type == 'department':
            self._load_departments()
        elif data_type == 'employee':
            self._load()
        elif data_type == 'settings':
            old_currency = self.currency
            self.currency         = self.db.get_setting('currency', 'TRY')
            self.work_days_month  = int(self.db.get_setting('work_days_month', '26'))
            self.work_hours_daily = float(self.db.get_setting('working_hours', '8'))
            if old_currency != self.currency:
                for spin in (self.basic_salary, self.housing, self.transport,
                             self.food, self.phone_allow, self.other,
                             self.bank_salary, self.cash_salary):
                    spin.setSuffix(f" {self.currency}")
            self._update_rate_labels()

    # ==================== بناء الواجهة ====================
    def _build(self):
        main = QHBoxLayout(self)

        # ---- يمين: النموذج ----
        right = QScrollArea()
        right.setWidgetResizable(True)
        right.setMaximumWidth(550)
        form_widget = QWidget()
        right_lay   = QVBoxLayout(form_widget)
        right_lay.setSpacing(8)

        # --- البيانات الأساسية ---
        g1 = QGroupBox("البيانات الأساسية")
        f1 = QFormLayout()
        f1.setLabelAlignment(Qt.AlignRight)
        f1.setSpacing(8)

        self.emp_code    = QLineEdit()
        self.first_name  = QLineEdit()
        self.last_name   = QLineEdit()
        self.national_id = QLineEdit()

        self.birth_date  = QDateEdit()
        self.birth_date.setCalendarPopup(True)
        self.birth_date.setSpecialValueText("-")
        self.birth_date.setDate(QDate())

        self.gender = QComboBox()
        self.gender.addItems(["ذكر", "أنثى"])

        self.nationality = QComboBox()
        self._load_nationalities()

        self.position   = QLineEdit()
        self.department = QComboBox()

        self.hire_date = QDateEdit()
        self.hire_date.setCalendarPopup(True)
        self.hire_date.setSpecialValueText("-")
        self.hire_date.setDate(QDate())

        self.contract = QComboBox()
        self.contract.addItems(["دوام كامل", "دوام جزئي", "عقد مؤقت", "متدرب"])
        self.contract.currentIndexChanged.connect(self._update_rate_labels)

        self.status = QComboBox()
        self.status.addItems(["نشط", "غير نشط", "إجازة", "منتهي الخدمة"])

        self.fingerprint = QLineEdit()
        self.exempt_fp   = QCheckBox("معفى من البصمة")

        for label, widget in [
            ("رقم الموظف:",         self.emp_code),
            ("الاسم الأول:",        self.first_name),
            ("الاسم الأخير:",       self.last_name),
            ("رقم الهوية:",         self.national_id),
            ("تاريخ الميلاد:",      self.birth_date),
            ("الجنس:",              self.gender),
            ("الجنسية:",            self.nationality),
            ("المسمى الوظيفي:",     self.position),
            ("القسم:",              self.department),
            ("تاريخ التوظيف:",      self.hire_date),
            ("نوع العقد:",          self.contract),
            ("الحالة:",             self.status),
            ("رقم البصمة:",         self.fingerprint),
            ("",                    self.exempt_fp),
        ]:
            f1.addRow(label, widget)
        g1.setLayout(f1)
        right_lay.addWidget(g1)

        # --- الراتب والبدلات ---
        g2 = QGroupBox("الراتب والبدلات")
        f2 = QFormLayout()
        f2.setSpacing(8)

        def _spin():
            s = QDoubleSpinBox()
            s.setRange(0, 999999)
            s.setDecimals(2)
            s.setSuffix(f" {self.currency}")
            return s

        self.basic_salary = _spin()
        self.housing      = _spin()
        self.transport    = _spin()
        self.food         = _spin()
        self.phone_allow  = _spin()
        self.other        = _spin()
        self.bank_salary  = _spin()
        self.cash_salary  = _spin()

        self.basic_salary.valueChanged.connect(self._on_basic_salary_changed)
        self.bank_salary.valueChanged.connect(self._on_bank_salary_changed)
        self.cash_salary.valueChanged.connect(self._on_cash_salary_changed)

        self.lbl_hourly_rate = QLabel("0.00")
        self.lbl_daily_rate  = QLabel("0.00")

        for label, widget in [
            ("الراتب الأساسي:",   self.basic_salary),
            ("الراتب عبر البنك:", self.bank_salary),
            ("الراتب النقدي:",    self.cash_salary),
            ("بدل سكن:",          self.housing),
            ("بدل نقل:",          self.transport),
            ("بدل غذاء:",         self.food),
            ("بدل هاتف:",         self.phone_allow),
            ("بدلات أخرى:",       self.other),
            ("أجر الساعة:",       self.lbl_hourly_rate),
            ("أجر اليوم:",        self.lbl_daily_rate),
        ]:
            f2.addRow(label, widget)
        g2.setLayout(f2)
        right_lay.addWidget(g2)

        # --- التأمينات الاجتماعية ---
        g3 = QGroupBox("التأمينات الاجتماعية")
        f3 = QFormLayout()
        f3.setSpacing(8)
        self.social_sec   = QLineEdit()
        self.social_check = QCheckBox("مسجل في التأمينات")
        self.social_date  = QDateEdit()
        self.social_date.setCalendarPopup(True)
        self.social_date.setSpecialValueText("-")
        self.social_date.setDate(QDate())
        self.social_check.toggled.connect(self._update_salary_fields)
        f3.addRow("رقم التأمينات:",     self.social_sec)
        f3.addRow("مسجل في التأمينات:", self.social_check)
        f3.addRow("تاريخ التسجيل:",    self.social_date)
        g3.setLayout(f3)
        right_lay.addWidget(g3)

        # --- الوثائق والتواريخ ---
        g4 = QGroupBox("الوثائق والتواريخ")
        f4 = QFormLayout()
        f4.setSpacing(8)

        def _date_edit():
            de = QDateEdit()
            de.setCalendarPopup(True)
            de.setSpecialValueText("-")
            de.setDate(QDate())
            return de

        self.iqama_num    = QLineEdit()
        self.iqama_exp    = _date_edit()
        self.passport_num = QLineEdit()
        self.passport_exp = _date_edit()
        self.health_ins   = QLineEdit()
        self.health_exp   = _date_edit()

        for label, widget in [
            ("رقم الإقامة:",           self.iqama_num),
            ("تاريخ انتهاء الإقامة:",  self.iqama_exp),
            ("رقم الجواز:",            self.passport_num),
            ("تاريخ انتهاء الجواز:",   self.passport_exp),
            ("رقم التأمين الصحي:",    self.health_ins),
            ("تاريخ انتهاء التأمين:", self.health_exp),
        ]:
            f4.addRow(label, widget)
        g4.setLayout(f4)
        right_lay.addWidget(g4)

        # --- الاتصال والبنك ---
        g5 = QGroupBox("الاتصال والبنك")
        f5 = QFormLayout()
        f5.setSpacing(8)
        self.phone        = QLineEdit()
        self.email        = QLineEdit()
        self.email.setPlaceholderText("example@domain.com")
        self.address      = QLineEdit()
        self.bank_name    = QLineEdit()
        self.bank_account = QLineEdit()
        self.iban         = QLineEdit()

        for label, widget in [
            ("الهاتف:",      self.phone),
            ("البريد:",      self.email),
            ("العنوان:",     self.address),
            ("البنك:",       self.bank_name),
            ("رقم الحساب:", self.bank_account),
            ("IBAN:",        self.iban),
        ]:
            f5.addRow(label, widget)
        g5.setLayout(f5)
        right_lay.addWidget(g5)

        # --- ملاحظات ---
        g6 = QGroupBox("ملاحظات")
        v6 = QVBoxLayout()
        self.notes = QTextEdit()
        self.notes.setMaximumHeight(80)
        v6.addWidget(self.notes)
        g6.setLayout(v6)
        right_lay.addWidget(g6)

        # --- أزرار النموذج ---
        btns_row = QHBoxLayout()
        self.btn_add    = btn("✅ إضافة", BTN_SUCCESS, self._add)
        self.btn_edit   = btn("✏️ تعديل", BTN_PRIMARY, self._edit)
        self.btn_delete = btn("🗑️ حذف",  BTN_DANGER,  self._delete)
        self.btn_clear  = btn("🧹 مسح",  BTN_GRAY,    self._clear)
        for b in (self.btn_add, self.btn_edit, self.btn_delete, self.btn_clear):
            btns_row.addWidget(b)
        right_lay.addLayout(btns_row)

        right.setWidget(form_widget)
        main.addWidget(right)

        # ---- يسار: الجدول + وثائق ----
        left = QVBoxLayout()

        search_row = QHBoxLayout()
        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("🔍 بحث بالاسم أو الرقم أو البصمة...")
        self.search_box.textChanged.connect(self._load)
        search_row.addWidget(self.search_box)
        self.btn_import            = btn("📥 Excel",  BTN_PURPLE, self._import_excel)
        self.btn_template          = btn("📤 قالب",   BTN_TEAL,   self._export_template)
        self.btn_refresh_employees = btn("🔄 تحديث", BTN_GRAY,   self._refresh_employees)
        for b in (self.btn_import, self.btn_template, self.btn_refresh_employees):
            search_row.addWidget(b)
        left.addLayout(search_row)

        self.table = make_table([
            "#", "الرقم", "الاسم", "الوظيفة", "القسم",
            "الراتب الأساسي", "الحالة"])
        self.table.cellClicked.connect(self._select)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)
        left.addWidget(self.table)

        # وثائق الموظف
        doc_grp = QGroupBox("📎 وثائق الموظف")
        doc_lay = QVBoxLayout()
        self.docs_list = QListWidget()
        self.docs_list.setMaximumHeight(120)
        self.docs_list.itemDoubleClicked.connect(self._open_document)
        doc_btns = QHBoxLayout()
        self.btn_upload     = btn("رفع وثيقة", BTN_PRIMARY, self._upload_doc)
        self.btn_delete_doc = btn("حذف وثيقة", BTN_DANGER,  self._delete_doc)
        doc_btns.addWidget(self.btn_upload)
        doc_btns.addWidget(self.btn_delete_doc)
        doc_lay.addWidget(self.docs_list)
        doc_lay.addLayout(doc_btns)
        doc_grp.setLayout(doc_lay)
        left.addWidget(doc_grp)

        left_widget = QWidget()
        left_widget.setLayout(left)
        main.addWidget(left_widget, 1)

        self._apply_permissions()
        self._load_departments()
        self._update_salary_fields()
        self._update_rate_labels()

    # ==================== تحكم في الراتب ====================
    def _update_salary_fields(self):
        if self.social_check.isChecked():
            self.bank_salary.setEnabled(True)
            total = self.basic_salary.value()
            bank  = self.bank_salary.value()
            cash  = total - bank
            if cash < 0:
                cash = 0
                bank = total
                self.bank_salary.setValue(bank)
            self.cash_salary.setValue(cash)
        else:
            self.bank_salary.setEnabled(False)
            self.bank_salary.setValue(0)
            self.cash_salary.setValue(self.basic_salary.value())

    def _on_basic_salary_changed(self, value):
        if self.social_check.isChecked():
            self.cash_salary.setValue(value - self.bank_salary.value())
        else:
            self.cash_salary.setValue(value)
        self._update_rate_labels()

    def _on_bank_salary_changed(self, value):
        if self.social_check.isChecked():
            self.cash_salary.blockSignals(True)
            self.cash_salary.setValue(self.basic_salary.value() - value)
            self.cash_salary.blockSignals(False)

    def _on_cash_salary_changed(self, value):
        if self.social_check.isChecked():
          self.bank_salary.blockSignals(True)
          self.bank_salary.setValue(self.basic_salary.value() - value)
          self.bank_salary.blockSignals(False)


    def _update_rate_labels(self):
        basic    = self.basic_salary.value()
        contract = self.contract.currentText()
        wdm = (20 if contract == "دوام جزئي"
               else 22 if contract == "متدرب"
               else self.work_days_month)
        wdh = self.work_hours_daily
        daily  = basic / wdm        if wdm > 0 else 0
        hourly = basic / (wdm * wdh) if wdm > 0 and wdh > 0 else 0
        self.lbl_daily_rate.setText(f"{daily:,.2f} {self.currency}")
        self.lbl_hourly_rate.setText(f"{hourly:,.2f} {self.currency}")

    def _refresh_employees(self):
        self._load()
        self._load_departments()
        QMessageBox.information(self, "تحديث", "تم تحديث بيانات الموظفين")

    # ==================== تحميل البيانات ====================
    def _load_nationalities(self):
        countries = [
            "أفغانستان","ألبانيا","الجزائر","أندورا","أنغولا","الأرجنتين","أرمينيا",
            "أستراليا","النمسا","أذربيجان","البحرين","بنغلاديش","روسيا البيضاء","بلجيكا",
            "بنين","بوتان","بوليفيا","البوسنة والهرسك","بوتسوانا","البرازيل","بروناي",
            "بلغاريا","بوركينا فاسو","بوروندي","كمبوديا","الكاميرون","كندا","الرأس الأخضر",
            "جمهورية أفريقيا الوسطى","تشاد","تشيلي","الصين","كولومبيا","جزر القمر",
            "الكونغو","كوستاريكا","ساحل العاج","كرواتيا","كوبا","قبرص","جمهورية التشيك",
            "الدنمارك","جيبوتي","دومينيكا","جمهورية الدومينيكان","الإكوادور","مصر",
            "السلفادور","غينيا الاستوائية","إريتريا","إستونيا","إثيوبيا","فيجي","فنلندا",
            "فرنسا","الغابون","غامبيا","جورجيا","ألمانيا","غانا","اليونان","غرينادا",
            "غواتيمالا","غينيا","غينيا بيساو","غيانا","هايتي","هندوراس","المجر","آيسلندا",
            "الهند","إندونيسيا","إيران","العراق","أيرلندا","إسرائيل","إيطاليا","جامايكا",
            "اليابان","الأردن","كازاخستان","كينيا","كيريباتي","كوريا الشمالية",
            "كوريا الجنوبية","الكويت","قيرغيزستان","لاوس","لاتفيا","لبنان","ليسوتو",
            "ليبيريا","ليبيا","ليختنشتاين","ليتوانيا","لوكسمبورغ","مدغشقر","مالاوي",
            "ماليزيا","جزر المالديف","مالي","مالطا","جزر مارشال","موريتانيا","موريشيوس",
            "المكسيك","ميكرونيسيا","مولدوفا","موناكو","منغوليا","الجبل الأسود","المغرب",
            "موزمبيق","ميانمار","ناميبيا","ناورو","نيبال","هولندا","نيوزيلندا","نيكاراغوا",
            "النيجر","نيجيريا","مقدونيا الشمالية","النرويج","عمان","باكستان","بالاو",
            "فلسطين","بنما","بابوا غينيا الجديدة","باراغواي","بيرو","الفلبين","بولندا",
            "البرتغال","قطر","رومانيا","روسيا","رواندا","سانت كيتس ونيفيس","سانت لوسيا",
            "سانت فينسنت والغرينادين","ساموا","سان مارينو","ساو تومي وبرينسيب",
            "المملكة العربية السعودية","السنغال","صربيا","سيشل","سيراليون","سنغافورة",
            "سلوفاكيا","سلوفينيا","جزر سليمان","الصومال","جنوب أفريقيا","جنوب السودان",
            "إسبانيا","سريلانكا","السودان","سورينام","إسواتيني","السويد","سويسرا","سوريا",
            "تايوان","طاجيكستان","تنزانيا","تايلاند","تيمور الشرقية","توغو","تونغا",
            "ترينيداد وتوباغو","تونس","تركيا","تركمانستان","توفالو","أوغندا","أوكرانيا",
            "الإمارات العربية المتحدة","المملكة المتحدة","الولايات المتحدة","أوروغواي",
            "أوزبكستان","فانواتو","الفاتيكان","فنزويلا","فيتنام","اليمن","زامبيا","زيمبابوي",
        ]
        self.nationality.addItems(countries)
        self.nationality.setEditable(True)
        self.nationality.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.nationality.completer().setFilterMode(Qt.MatchContains)

    def _apply_permissions(self):
        role = self.user['role']
        self.btn_add.setVisible(can_add(role))
        self.btn_edit.setVisible(can_edit(role))
        self.btn_delete.setVisible(can_delete(role))
        self.btn_import.setVisible(can_add(role))
        self.btn_template.setVisible(True)
        self.btn_upload.setVisible(can_add(role))
        self.btn_delete_doc.setVisible(can_delete(role))

    def _load_departments(self):
        self.department.clear()
        self.department.addItem("-- اختر القسم --", None)
        for did, name in self.db.fetch_all(
                "SELECT id, name FROM departments ORDER BY name"):
            self.department.addItem(name, did)
        self.department.setEditable(True)
        self.department.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.department.completer().setFilterMode(Qt.MatchContains)

    def _load(self):
        txt = self.search_box.text().strip()
        if txt:
            data = self.db.fetch_all("""
                SELECT e.id, e.employee_code,
                       e.first_name || ' ' || e.last_name,
                       e.position, d.name, e.basic_salary, e.status
                FROM employees e
                LEFT JOIN departments d ON e.department_id = d.id
                WHERE e.employee_code LIKE ?
                   OR e.first_name    LIKE ?
                   OR e.last_name     LIKE ?
                   OR e.fingerprint_id LIKE ?
                   OR (e.first_name || ' ' || e.last_name) LIKE ?
                ORDER BY e.id DESC
            """, (f'%{txt}%',) * 5)
        else:
            data = self.db.fetch_all("""
                SELECT e.id, e.employee_code,
                       e.first_name || ' ' || e.last_name,
                       e.position, d.name, e.basic_salary, e.status
                FROM employees e
                LEFT JOIN departments d ON e.department_id = d.id
                ORDER BY e.id DESC
            """)
        fill_table(self.table, data, colors={
            6: lambda v: (
                "#388E3C" if v == "نشط"
                else "#D32F2F" if v == "منتهي الخدمة"
                else "#F57C00"
            )
        })

    # ==================== التحقق والمعاملات ====================
    def _is_fingerprint_duplicate(self, fingerprint: str,
                                   exclude_id=None) -> bool:
        if exclude_id:
            row = self.db.fetch_one(
                "SELECT id FROM employees WHERE fingerprint_id=? AND id!=?",
                (fingerprint, exclude_id))
        else:
            row = self.db.fetch_one(
                "SELECT id FROM employees WHERE fingerprint_id=?",
                (fingerprint,))
        return row is not None
    def _generate_employee_code(self) -> None:
        """
        توليد رقم موظف تلقائيًا (رقمي متسلسل)
        يُستخدم فقط عند الإدخال اليدوي إذا كان الحقل فارغًا.
        """
        if self.emp_code.text().strip():
            return

        row = self.db.fetch_one(
            "SELECT MAX(CAST(employee_code AS INTEGER)) FROM employees"
        )

        next_code = 1001
        if row and row[0]:
            try:
                next_code = int(row[0]) + 1
            except ValueError:
                pass

        self.emp_code.setText(str(next_code))


    def _validate(self) -> bool:
        # حقول نصية مطلوبة
        required = [
            (self.emp_code,   "رقم الموظف"),
            (self.first_name, "الاسم الأول"),
            (self.last_name,  "الاسم الأخير"),
            (self.position,   "المسمى الوظيفي"),
        ]

        # رقم البصمة مطلوب فقط إذا لم يكن الموظف معفى
        if not self.exempt_fp.isChecked():
            if not self.fingerprint.text().strip():
                QMessageBox.warning(self, "خطأ", "رقم البصمة مطلوب")
                return False

        for field, name in required:
            if not field.text().strip():
                QMessageBox.warning(self, "خطأ", f"{name} مطلوب")
                return False

        # القسم إلزامي فقط عند الإضافة اليدوية
        if self.current_id is None and self.department.currentData() is None:
            QMessageBox.warning(self, "خطأ", "يجب اختيار القسم")
            return False

        # تاريخ التوظيف
        if self.hire_date.date().isNull():
            QMessageBox.warning(self, "خطأ", "يجب اختيار القسم")
        return False

        # الراتب الأساسي
        if self.basic_salary.value() <= 0:
            QMessageBox.warning(self, "خطأ", "الراتب الأساسي يجب أن يكون أكبر من صفر")
            return False

        # تكرار البصمة
        fp = self.fingerprint.text().strip()
        if self._is_fingerprint_duplicate(fp, self.current_id):
            QMessageBox.warning(self, "خطأ", "رقم البصمة موجود مسبقاً لموظف آخر")
            return False

        # التأمينات → الراتب البنكي مطلوب
        if self.social_check.isChecked() and self.bank_salary.value() <= 0:
            QMessageBox.warning(
                self, "خطأ",
                "الراتب عبر البنك مطلوب (الموظف مسجل في التأمينات)")
            return False

        # مجموع البنكي + النقدي = الأساسي
        total_paid = self.bank_salary.value() + self.cash_salary.value()
        if abs(total_paid - self.basic_salary.value()) > 0.01:
            QMessageBox.warning(
                self, "خطأ",
                f"مجموع الراتب البنكي والنقدي ({total_paid:.2f}) "
                f"يجب أن يساوي الراتب الأساسي ({self.basic_salary.value():.2f})")
            return False

        # تاريخ الميلاد (إن أُدخل)
        if not self.birth_date.date().isNull():
            if self.birth_date.date().toPyDate() > date.today():
                QMessageBox.warning(self, "خطأ", "تاريخ الميلاد لا يمكن أن يكون في المستقبل")
                return False

        # رقم الهوية (إن أُدخل)
        nid = self.national_id.text().strip()
        if nid:
            if not nid.isdigit() or len(nid) < 6:
                QMessageBox.warning(self, "خطأ", "رقم الهوية غير صالح")
                return False

        # البريد الإلكتروني (إن أُدخل)
        email = self.email.text().strip()
        if email:
            pattern = r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
            if not re.match(pattern, email):
                QMessageBox.warning(
                    self, "خطأ",
                    f"صيغة البريد الإلكتروني غير صحيحة:\n{email}")
                return False

        # رقم الحساب / IBAN (إن أُدخل)
        iban = self.iban.text().strip()
        if iban and len(iban) < 8:
            QMessageBox.warning(self, "خطأ", "رقم IBAN غير صالح")
            return False

        return True

    def _params(self) -> tuple:
        """جمع قيم جميع حقول النموذج في tuple للإدراج/التحديث."""
        def _date(de: QDateEdit):
            return de.date().toString(Qt.ISODate) if not de.date().isNull() else None

        return (
            self.emp_code.text().strip(),
            self.first_name.text().strip(),
            self.last_name.text().strip(),
            self.national_id.text().strip() or None,
            self.nationality.currentText().strip(),
            _date(self.birth_date),
            self.gender.currentText(),
            self.position.text().strip() or None,
            self.department.currentData(),
            _date(self.hire_date),
            self.contract.currentText(),
            self.basic_salary.value(),
            self.housing.value(),
            self.transport.value(),
            self.food.value(),
            self.phone_allow.value(),
            self.other.value(),
            self.bank_salary.value(),
            self.cash_salary.value(),
            self.phone.text().strip() or None,
            self.email.text().strip() or None,
            self.address.text().strip() or None,
            self.bank_name.text().strip() or None,
            self.bank_account.text().strip() or None,
            self.iban.text().strip() or None,
            self.fingerprint.text().strip(),
            self.social_sec.text().strip() or None,
            1 if self.social_check.isChecked() else 0,
            _date(self.social_date),
            _date(self.iqama_exp),
            self.iqama_num.text().strip() or None,
            _date(self.passport_exp),
            self.passport_num.text().strip() or None,
            _date(self.health_exp),
            self.health_ins.text().strip() or None,
            self.status.currentText(),
            self.notes.toPlainText().strip() or None,
            1 if self.exempt_fp.isChecked() else 0,
        )

        # ==================== عمليات CRUD ====================
    def _add(self):
        self._generate_employee_code()

        if not self._validate():
            return

        q = """
            INSERT INTO employees (
                employee_code, first_name, last_name, national_id, nationality,
                birth_date, gender, position, department_id,
                hire_date, contract_type,
                basic_salary, housing_allowance, transportation_allowance,
                food_allowance, phone_allowance, other_allowances,
                bank_salary, cash_salary,
                phone, email, address, bank_name, bank_account, iban,
                fingerprint_id,
                social_security_number, social_security_registered, social_security_date,
                iqama_expiry, iqama_number,
                passport_expiry, passport_number,
                health_insurance_expiry, health_insurance_number,
                status, notes, is_exempt_from_fingerprint
            ) VALUES (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?,
                ?, ?,
                ?, ?, ?,
                ?, ?, ?,
                ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?,
                ?, ?, ?,
                ?, ?,
                ?, ?,
                ?, ?,
                ?, ?, ?
            )
        """

        if self.db.execute_query(q, self._params()):
            self.current_id = self.db.last_id()

            self.db.execute(
                "INSERT INTO employee_audit_log (employee_id, action, changed_by, changed_at) VALUES (?, ?, ?, ?)",
                (
                    self.current_id,
                    "CREATE",
                    self.user.get('role', 'system'),
                    datetime.now().isoformat(timespec='seconds')
                )
            )

            QMessageBox.information(self, "نجاح", "تم إضافة الموظف")

            self._load()
            self._clear()

            if self.comm:
                self.comm.dataChanged.emit(
                    'employee',
                    {'action': 'add', 'id': self.current_id}
                )

        else:
            QMessageBox.critical(
                self,
                "خطأ",
                "فشل في الإضافة — تأكد من عدم تكرار رقم الموظف أو البصمة"
            )

    def _edit(self):
        if not self.current_id:
            QMessageBox.warning(self, "خطأ", "اختر موظفاً أولاً")
            return
        if not self._validate():
            return

        old = self.db.fetch_one(_EMP_SELECT, (self.current_id,))
        if not old:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على الموظف")
            return
        old_values = {
            "employee_code":  old[_C['employee_code']],
            "first_name":     old[_C['first_name']],
            "last_name":      old[_C['last_name']],
            "fingerprint_id": old[_C['fingerprint_id']],
            "basic_salary":   old[_C['basic_salary']],
            "status":         old[_C['status']],
        }

        q = """UPDATE employees SET
                employee_code=?, first_name=?, last_name=?, national_id=?, nationality=?,
                birth_date=?, gender=?, position=?, department_id=?,
                hire_date=?, contract_type=?,
                basic_salary=?, housing_allowance=?, transportation_allowance=?,
                food_allowance=?, phone_allowance=?, other_allowances=?,
                bank_salary=?, cash_salary=?,
                phone=?, email=?, address=?, bank_name=?, bank_account=?, iban=?,
                fingerprint_id=?,
                social_security_number=?, social_security_registered=?, social_security_date=?,
                iqama_expiry=?, iqama_number=?,
                passport_expiry=?, passport_number=?,
                health_insurance_expiry=?, health_insurance_number=?,
                status=?, notes=?, is_exempt_from_fingerprint=?
               WHERE id=?"""

        if self.db.execute_query(q, self._params() + (self.current_id,)):
            QMessageBox.information(self, "نجاح", "تم تعديل البيانات")
            self.db.log_update("employees", self.current_id, old_values, {
                "employee_code":  self.emp_code.text(),
                "first_name":     self.first_name.text(),
                "last_name":      self.last_name.text(),
                "fingerprint_id": self.fingerprint.text(),
                "basic_salary":   self.basic_salary.value(),
                "status":         self.status.currentText(),
            })
            self._load()
            self._clear()
            if self.comm:
                self.comm.dataChanged.emit(
                    'employee', {'action': 'edit', 'id': self.current_id})
        else:
            QMessageBox.critical(self, "خطأ", "فشل في التعديل")

    def _delete(self):
        """
        حذف موظف مع جميع البيانات المرتبطة به في معاملة واحدة ذرية.

        الإصلاح: استبدال BEGIN/COMMIT اليدوي بـ with self.db.transaction()
        الذي يضمن rollback صحيح وكامل عند أي استثناء.
        ملاحظة: execute_query داخل transaction() لا تُنفِّذ commit مبكراً
        بسبب الفلاج _in_transaction الذي أضفناه في database.py.
        """
        if not self.current_id:
            QMessageBox.warning(self, "خطأ", "اختر موظفاً أولاً")
            return

        emp = self.db.fetch_one(_EMP_SELECT, (self.current_id,))
        if not emp:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على الموظف")
            return

        emp_values = {
            "employee_code":  emp[_C['employee_code']],
            "first_name":     emp[_C['first_name']],
            "last_name":      emp[_C['last_name']],
            "fingerprint_id": emp[_C['fingerprint_id']],
        }

        if QMessageBox.question(
                self, "تأكيد",
                "هل أنت متأكد من الحذف؟\n"
                "سيتم حذف جميع سجلات الحضور والإجازات والسلف والوثائق المرتبطة.",
                QMessageBox.Yes | QMessageBox.No) != QMessageBox.Yes:
            return

        try:
            # جلب الوثائق قبل المعاملة لحذف الملفات لاحقاً
            docs = self.db.fetch_all(
                "SELECT id, document_path FROM documents WHERE employee_id=?",
                (self.current_id,))
            loans = self.db.fetch_all(
                "SELECT id FROM loans WHERE employee_id=?",
                (self.current_id,))

            with self.db.transaction():
                # حذف الأقساط المرتبطة بالسلف
                for (loan_id,) in loans:
                    self.db.execute_query(
                        "DELETE FROM installments WHERE loan_id=?", (loan_id,))
                # حذف السلف
                self.db.execute_query(
                    "DELETE FROM loans WHERE employee_id=?", (self.current_id,))
                # حذف الحضور
                self.db.execute_query(
                    "DELETE FROM attendance WHERE employee_id=?", (self.current_id,))
                # حذف طلبات الإجازة
                self.db.execute_query(
                    "DELETE FROM leave_requests WHERE employee_id=?", (self.current_id,))
                # حذف أرصدة الإجازة
                self.db.execute_query(
                    "DELETE FROM leave_balance WHERE employee_id=?", (self.current_id,))
                # حذف سجل الوثائق من قاعدة البيانات
                self.db.execute_query(
                    "DELETE FROM documents WHERE employee_id=?", (self.current_id,))
                # حذف الراتب (ON DELETE CASCADE لكن نحذف صراحةً للأمان)
                self.db.execute_query(
                    "DELETE FROM payroll WHERE employee_id=?", (self.current_id,))
                # حذف الموظف نفسه
                self.db.execute_query(
                    "DELETE FROM employees WHERE id=?", (self.current_id,))

            # حذف ملفات الوثائق من القرص (بعد نجاح المعاملة)
            for _, doc_path in docs:
                try:
                    if doc_path and os.path.exists(doc_path):
                        os.remove(doc_path)
                except Exception as e:
                    logger.warning("لم يتم حذف ملف الوثيقة: %s — %s", doc_path, e)

            self.db.log_delete("employees", self.current_id, emp_values)
            self._load()
            self._clear()
            if self.comm:
                self.comm.dataChanged.emit(
                    'employee', {'action': 'delete', 'id': self.current_id})
            QMessageBox.information(
                self, "نجاح", "تم حذف الموظف وجميع البيانات المرتبطة")

        except Exception as e:
            logger.error("خطأ في حذف الموظف %d: %s", self.current_id, e, exc_info=True)
            QMessageBox.critical(self, "خطأ", f"حدث خطأ أثناء الحذف:\n{e}")

    def _clear(self):
        self.current_id = None
        self.emp_code.clear()
        self.first_name.clear()
        self.last_name.clear()
        self.national_id.clear()
        self.nationality.setCurrentIndex(0)
        self.birth_date.setDate(QDate())
        self.gender.setCurrentIndex(0)
        self.position.clear()
        self.department.setCurrentIndex(0)
        self.hire_date.setDate(QDate())
        self.contract.setCurrentIndex(0)
        self.status.setCurrentIndex(0)
        self.fingerprint.clear()
        self.exempt_fp.setChecked(False)
        for s in (self.basic_salary, self.housing, self.transport,
                  self.food, self.phone_allow, self.other,
                  self.bank_salary, self.cash_salary):
            s.setValue(0)
        self.social_sec.clear()
        self.social_check.setChecked(False)
        self.social_date.setDate(QDate())
        self.iqama_num.clear()
        self.iqama_exp.setDate(QDate())
        self.passport_num.clear()
        self.passport_exp.setDate(QDate())
        self.health_ins.clear()
        self.health_exp.setDate(QDate())
        self.phone.clear()
        self.email.clear()
        self.address.clear()
        self.bank_name.clear()
        self.bank_account.clear()
        self.iban.clear()
        self.notes.clear()
        self.docs_list.clear()
        self.lbl_hourly_rate.setText("0.00")
        self.lbl_daily_rate.setText("0.00")

    def _select(self, row: int, col: int):
        """
        ملء النموذج بيانات الموظف المحدد في الجدول.

        الإصلاح: يستخدم _EMP_SELECT الصريح بأسماء الأعمدة
        مع _C كقاموس للفهارس — محمي من تغييرات هيكل الجدول.
        """
        item = self.table.item(row, 0)
        if not item:
            return
        self.current_id = int(item.text())

        e = self.db.fetch_one(_EMP_SELECT, (self.current_id,))
        if not e:
            return

        def _set_date(de: QDateEdit, val):
            de.setDate(QDate.fromString(str(val), Qt.ISODate) if val else QDate())

        self.emp_code.setText(e[_C['employee_code']] or "")
        self.first_name.setText(e[_C['first_name']] or "")
        self.last_name.setText(e[_C['last_name']] or "")
        self.national_id.setText(e[_C['national_id']] or "")

        nat_idx = self.nationality.findText(e[_C['nationality']] or "", Qt.MatchExactly)
        if nat_idx >= 0:
            self.nationality.setCurrentIndex(nat_idx)
        else:
            self.nationality.setCurrentText(e[_C['nationality']] or "")

        _set_date(self.birth_date, e[_C['birth_date']])

        g_idx = self.gender.findText(e[_C['gender']] or "ذكر")
        if g_idx >= 0:
            self.gender.setCurrentIndex(g_idx)

        self.position.setText(e[_C['position']] or "")

        dept_idx = self.department.findData(e[_C['department_id']])
        self.department.setCurrentIndex(max(0, dept_idx))

        _set_date(self.hire_date, e[_C['hire_date']])

        c_idx = self.contract.findText(e[_C['contract_type']] or "")
        if c_idx >= 0:
            self.contract.setCurrentIndex(c_idx)

        self.basic_salary.setValue(float(e[_C['basic_salary']]  or 0))
        self.housing.setValue(float(e[_C['housing_allowance']]        or 0))
        self.transport.setValue(float(e[_C['transportation_allowance']] or 0))
        self.food.setValue(float(e[_C['food_allowance']]               or 0))
        self.phone_allow.setValue(float(e[_C['phone_allowance']]       or 0))
        self.other.setValue(float(e[_C['other_allowances']]            or 0))
        self.bank_salary.setValue(float(e[_C['bank_salary']]           or 0))
        self.cash_salary.setValue(float(e[_C['cash_salary']]           or 0))

        self.phone.setText(e[_C['phone']] or "")
        self.email.setText(e[_C['email']] or "")
        self.address.setText(e[_C['address']] or "")
        self.bank_name.setText(e[_C['bank_name']] or "")
        self.bank_account.setText(e[_C['bank_account']] or "")
        self.iban.setText(e[_C['iban']] or "")
        self.fingerprint.setText(e[_C['fingerprint_id']] or "")
        self.social_sec.setText(e[_C['social_security_number']] or "")
        self.social_check.setChecked(bool(e[_C['social_security_registered']]))
        _set_date(self.social_date, e[_C['social_security_date']])

        _set_date(self.iqama_exp,    e[_C['iqama_expiry']])
        self.iqama_num.setText(e[_C['iqama_number']] or "")
        _set_date(self.passport_exp, e[_C['passport_expiry']])
        self.passport_num.setText(e[_C['passport_number']] or "")
        _set_date(self.health_exp,   e[_C['health_insurance_expiry']])
        self.health_ins.setText(e[_C['health_insurance_number']] or "")

        s_idx = self.status.findText(e[_C['status']] or "نشط")
        if s_idx >= 0:
            self.status.setCurrentIndex(s_idx)

        self.notes.setText(e[_C['notes']] or "")
        self.exempt_fp.setChecked(bool(e[_C['is_exempt_from_fingerprint']]))

        self._load_docs()
        self._update_rate_labels()
        self._update_salary_fields()

    # ==================== القائمة المنبثقة ====================
    def _show_context_menu(self, pos):
        row = self.table.currentRow()
        if row < 0:
            return
        menu = QMenu()
        action = menu.exec_(self.table.mapToGlobal(pos))

    # ==================== الوثائق ====================
    def _load_docs(self):
        self.docs_list.clear()
        if not self.current_id:
            return
        for did, name, path in self.db.fetch_all(
                "SELECT id, document_name, document_path "
                "FROM documents WHERE employee_id=?",
                (self.current_id,)):
            item = QListWidgetItem(f"📄 {name}")
            item.setData(Qt.UserRole,     path)
            item.setData(Qt.UserRole + 1, did)
            self.docs_list.addItem(item)

    def _open_document(self, item):
        path = item.data(Qt.UserRole)
        if not path or not os.path.exists(path):
            QMessageBox.warning(self, "خطأ", "الملف غير موجود")
            return
        try:
            if platform.system() == 'Windows':
                os.startfile(path)
            elif platform.system() == 'Darwin':
                subprocess.call(['open', path])
            else:
                subprocess.call(['xdg-open', path])
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"تعذر فتح الملف:\n{e}")

    def _upload_doc(self):
        if not self.current_id:
            QMessageBox.warning(self, "خطأ", "اختر موظفاً أولاً")
            return
        path, _ = QFileDialog.getOpenFileName(self, "اختر ملف", "", "All Files (*)")
        if not path:
            return
        emp_dir = os.path.join(DOCUMENTS_FOLDER, str(self.current_id))
        os.makedirs(emp_dir, exist_ok=True)
        fname = os.path.basename(path)
        ts    = datetime.now().strftime('%Y%m%d_%H%M%S')
        dest  = os.path.join(emp_dir, f"{ts}_{fname}")
        shutil.copy2(path, dest)
        self.db.execute_query(
            "INSERT INTO documents "
            "(employee_id, document_name, document_path, upload_date) VALUES (?,?,?,?)",
            (self.current_id, fname, dest, date.today().isoformat()))
        self._load_docs()
        self.db.log_custom("رفع وثيقة", "documents", self.db.last_id(), {
            "employee_id":   self.current_id,
            "document_name": fname})
        if self.comm:
            self.comm.dataChanged.emit(
                'document', {'action': 'upload', 'employee_id': self.current_id})

    def _delete_doc(self):
        item = self.docs_list.currentItem()
        if not item:
            return
        path   = item.data(Qt.UserRole)
        doc_id = item.data(Qt.UserRole + 1)
        if QMessageBox.question(
                self, "تأكيد", "حذف الوثيقة؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            doc_name = item.text().replace("📄 ", "")
            self.db.execute_query("DELETE FROM documents WHERE id=?", (doc_id,))
            try:
                if path and os.path.exists(path):
                    os.remove(path)
            except Exception as e:
                logger.warning("لم يتم حذف ملف الوثيقة: %s", e)
            self.db.log_delete("documents", doc_id, {
                "document_name": doc_name, "employee_id": self.current_id})
            self._load_docs()
            if self.comm:
                self.comm.dataChanged.emit(
                    'document', {'action': 'delete',
                                 'employee_id': self.current_id})

    # ==================== استيراد / تصدير Excel ====================
    def _import_excel(self):
        if not can_add(self.user['role']):
            QMessageBox.warning(self, "صلاحيات", "ليس لديك صلاحية لاستيراد البيانات")
            return
        path, _ = QFileDialog.getOpenFileName(
            self, "اختر ملف Excel", "",
            "Excel Files (*.xlsx *.xls)")
        if not path:
            return
        try:
            import pandas as pd
            df = pd.read_excel(path)
            required = ['employee_code', 'first_name', 'last_name', 'fingerprint_id']
            missing  = [c for c in required if c not in df.columns]
            if missing:
                QMessageBox.warning(
                    self, "خطأ",
                    f"الأعمدة التالية ناقصة:\n{', '.join(missing)}")
                return

            ok = fail = 0
            without_department = 0

            for _, row in df.iterrows():
                try:
                    dept_name = str(row.get('department', '') or '').strip()
                    dept_id   = None
                    if dept_name:
                        dept = self.db.fetch_one(
                            "SELECT id FROM departments WHERE name=?", (dept_name,))
                        if dept:
                            dept_id = dept[0]
                        if not dept_id:
                            without_department += 1

                    p = (
                        str(row.get('employee_code', '')).strip(),
                        str(row.get('first_name', '')).strip(),
                        str(row.get('last_name', '')).strip(),
                        str(row.get('fingerprint_id', '')).strip(),
                        str(row.get('position', '') or '').strip() or None,
                        dept_id,
                        float(row.get('basic_salary', 0) or 0),
                        QDate.currentDate().toString(Qt.ISODate),
                        'نشط'
                    )
                    q = """INSERT OR IGNORE INTO employees
                           (employee_code, first_name, last_name, fingerprint_id,
                            position, department_id, basic_salary, hire_date, status)
                           VALUES (?,?,?,?,?,?,?,?,?)"""
                    if self.db.execute_query(q, p):
                        new_id = self.db.last_id()
                        self.db.log_insert("employees", new_id, {
                            "employee_code":  p[0], "first_name":    p[1],
                            "last_name":      p[2], "fingerprint_id": p[3]})
                        ok += 1
                    else:
                        fail += 1
                except Exception as ex:
                    logger.warning("خطأ في استيراد صف: %s", ex)
                    fail += 1

            msg = f"✅ نجح: {ok}\n❌ فشل: {fail}"

            if without_department > 0:
                msg += (
                    f"\n\n⚠️ تنبيه:\n"
                    f"عدد ({without_department}) موظف بدون قسم.\n"
                    f"يرجى استكمال بيانات القسم قبل اعتماد الحضور أو الرواتب."
                )

            QMessageBox.information(self, "نتيجة الاستيراد", msg)
            self._load()

        except ImportError:
            QMessageBox.critical(
                self, "خطأ", "قم بتثبيت المكتبة:\npip install pandas openpyxl")

    def _export_template(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "حفظ قالب", "employee_template.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            import pandas as pd
            df = pd.DataFrame({
                'employee_code':              [''],
                'first_name':                 ['محمد'],
                'last_name':                  ['العلي'],
                'department':                 [''],
                'status':                     ['نشط'],
                'fingerprint_id':             ['1001'],
                'is_exempt_from_fingerprint': [''],
                'basic_salary':               [5000],
                'bank_salary':                [5000],
                'cash_salary':                [0],
                'social_security_registered': [1],
                'social_security_date':       ['2024-01-01'],
                'hire_date':                  ['2024-01-01'],
            })
            df.to_excel(path, index=False)
            QMessageBox.information(self, "نجاح", "تم حفظ قالب الاستيراد المعتمد")
        except ImportError:
            QMessageBox.critical(
                self, "خطأ", "pip install pandas openpyxl")
