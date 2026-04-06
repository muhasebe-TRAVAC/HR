#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# tabs/receipts_tab.py

import os
from datetime import datetime
from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QComboBox, QPushButton,
    QMessageBox, QGroupBox, QFormLayout, QLineEdit, QTextEdit, QRadioButton,
    QButtonGroup, QDateEdit, QFileDialog, QDialog, QDialogButtonBox,
    QDoubleSpinBox, QCompleter
)
from PyQt5.QtCore import Qt, QDate
from PyQt5.QtGui import QPixmap, QFont, QPainter, QTextDocument, QTextCursor, QTextBlockFormat, QTextCharFormat
from PyQt5.QtPrintSupport import QPrinter, QPrintDialog, QPrintPreviewDialog

from database import DatabaseManager
from utils import btn, can_add, number_to_words_tr
from constants import BTN_SUCCESS, BTN_PRIMARY, BTN_GRAY

class ReceiptsTab(QWidget):
    def __init__(self, db, user, comm=None):
        super().__init__()
        self.db = db
        self.user = user
        self.comm = comm
        self._build()
        if self.comm:
            self.comm.dataChanged.connect(self._on_data_changed)

    def _on_data_changed(self, data_type, data):
        if data_type == 'employee':
            self._refresh_employee_list()
        elif data_type == 'payroll':
            self._refresh_payroll_list()

    def _refresh_employee_list(self):
        """تحديث قائمة الموظفين مع تفعيل البحث"""
        self.emp_combo.clear()
        emps = self.db.fetch_all("SELECT id, first_name||' '||last_name FROM employees WHERE status='نشط' ORDER BY first_name")
        for eid, name in emps:
            self.emp_combo.addItem(name, eid)

        # تفعيل خاصية البحث (اكتب حرفاً وستظهر القائمة المطابقة)
        self.emp_combo.setEditable(True)
        self.emp_combo.setInsertPolicy(QComboBox.NoInsert)  # لا يضيف نصاً جديداً للقائمة
        self.emp_combo.completer().setCompletionMode(QCompleter.PopupCompletion)
        self.emp_combo.completer().setFilterMode(Qt.MatchContains)
        self.emp_combo.completer().setCaseSensitivity(Qt.CaseInsensitive)

    def _refresh_payroll_list(self):
        """تحديث قائمة الرواتب (غير معتمدة)"""
        self.payroll_combo.clear()
        self.payroll_combo.addItem("-- Seçiniz --", None)
        payrolls = self.db.fetch_all("""
            SELECT p.id, e.first_name||' '||e.last_name, p.month, p.year, p.net_salary
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            WHERE p.status != 'معتمد'
            ORDER BY p.year DESC, p.month DESC
        """)
        for pid, name, month, year, net in payrolls:
            month_names = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
                           'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
            month_str = month_names[month-1] if 1 <= month <= 12 else str(month)
            self.payroll_combo.addItem(f"{name} - {month_str} {year} - {net:,.2f} TL", pid)

    def _build(self):
        layout = QVBoxLayout(self)

        # عنوان
        title = QLabel("🧾 Ödeme Makbuzu")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:18px; font-weight:bold; color:#1976D2; margin:10px;")
        layout.addWidget(title)

        # ========== Alıcı Bilgileri ==========
        receiver_group = QGroupBox("Alıcı Bilgileri")
        receiver_layout = QFormLayout()
        receiver_layout.setSpacing(10)

        self.receiver_type_group = QButtonGroup()
        rb_employee = QRadioButton("Personel")
        rb_other = QRadioButton("Diğer")
        rb_employee.setChecked(True)
        self.receiver_type_group.addButton(rb_employee, 1)
        self.receiver_type_group.addButton(rb_other, 2)

        type_layout = QHBoxLayout()
        type_layout.addWidget(rb_employee)
        type_layout.addWidget(rb_other)
        type_layout.addStretch()
        receiver_layout.addRow("Alıcı Tipi:", type_layout)

        # Personel seçimi (arama özellikli)
        self.emp_combo = QComboBox()
        self.emp_combo.setMinimumWidth(250)
        self._refresh_employee_list()
        self.emp_combo.currentIndexChanged.connect(self._on_employee_changed)
        receiver_layout.addRow("Personel Seçin:", self.emp_combo)

        # Diğer alıcı adı
        self.other_name = QLineEdit()
        self.other_name.setPlaceholderText("Alıcı adını girin")
        self.other_name.setEnabled(False)
        receiver_layout.addRow("Ad:", self.other_name)

        # ربط الإشارات
        rb_employee.toggled.connect(self._on_receiver_type_changed)
        rb_other.toggled.connect(self._on_receiver_type_changed)

        receiver_group.setLayout(receiver_layout)
        layout.addWidget(receiver_group)

        # ========== Ödeme Bilgileri ==========
        payment_group = QGroupBox("Ödeme Bilgileri")
        payment_layout = QFormLayout()
        payment_layout.setSpacing(10)

        self.payment_date = QDateEdit()
        self.payment_date.setCalendarPopup(True)
        self.payment_date.setDate(QDate.currentDate())
        payment_layout.addRow("Tarih:", self.payment_date)

        # Ödeme türü - الافتراضي "Diger"
        self.payment_type_combo = QComboBox()
        self.payment_type_combo.addItems(["Diger", "Maas", "Prim", "Harcırah", "Hediye", "Yolluk"])
        self.payment_type_combo.setCurrentIndex(0)  # "Diger" أولاً
        self.payment_type_combo.currentTextChanged.connect(self._on_payment_type_changed)
        payment_layout.addRow("Ödeme Türü:", self.payment_type_combo)

        # Maaş bordrosu listesi (sadece "Maas" seçildiğinde görünür)
        self.payroll_combo = QComboBox()
        self.payroll_combo.setVisible(False)
        self.payroll_combo.currentIndexChanged.connect(self._on_payroll_selected)
        self._refresh_payroll_list()
        payment_layout.addRow("Maaş Bordrosu:", self.payroll_combo)

        # Tutar (sayısal) - düzenlenebilir
        self.amount_spin = QDoubleSpinBox()
        self.amount_spin.setRange(0, 999999999)
        self.amount_spin.setDecimals(2)
        self.amount_spin.setSuffix(" TL")
        self.amount_spin.valueChanged.connect(self._update_amount_words)
        payment_layout.addRow("Tutar:", self.amount_spin)

        # Tutar yazıyla (otomatik)
        self.amount_words = QLineEdit()
        self.amount_words.setReadOnly(True)
        self.amount_words.setPlaceholderText("Otomatik yazıyla tutar")
        payment_layout.addRow("Tutar Yazıyla:", self.amount_words)

        # Açıklama (otomatik doldurulur, düzenlenebilir)
        self.description = QTextEdit()
        self.description.setMaximumHeight(80)
        self.description.setPlaceholderText("Açıklama (otomatik doldurulur, düzenlenebilir)")
        payment_layout.addRow("Açıklama:", self.description)

        # Notlar
        self.notes = QLineEdit()
        self.notes.setPlaceholderText("Ek not (isteğe bağlı)")
        payment_layout.addRow("Notlar:", self.notes)

        payment_group.setLayout(payment_layout)
        layout.addWidget(payment_group)

        # ========== Butonlar ==========
        btn_layout = QHBoxLayout()
        self.btn_print = btn("🖨️ Makbuz Yazdır", BTN_PRIMARY, self._print_receipt)
        self.btn_preview = btn("👁️ Önizleme", BTN_GRAY, self._preview_receipt)
        self.btn_clear = btn("🧹 Temizle", BTN_GRAY, self._clear)
        btn_layout.addWidget(self.btn_print)
        btn_layout.addWidget(self.btn_preview)
        btn_layout.addWidget(self.btn_clear)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        layout.addStretch()
        self._on_receiver_type_changed()

    # ========== Olay işleyiciler ==========

    def _on_receiver_type_changed(self):
        """Alıcı tipine göre alanları etkinleştir/devre dışı bırak"""
        if self.receiver_type_group.checkedId() == 2:  # Diğer
            self.emp_combo.setEnabled(False)
            self.other_name.setEnabled(True)
            # Ödeme türü "Diger" yap ve devre dışı bırak
            self.payment_type_combo.setCurrentIndex(0)  # "Diger"
            self.payment_type_combo.setEnabled(False)
        else:  # Personel
            self.emp_combo.setEnabled(True)
            self.other_name.setEnabled(False)
            self.other_name.clear()
            # Ödeme türünü aktif et
            self.payment_type_combo.setEnabled(True)
            # Not: Mevcut seçili değer korunur

    def _on_payment_type_changed(self, txt):
        """Ödeme tipi değişince maaş listesini göster/gizle ve açıklamayı güncelle"""
        if txt == "Maas":
            self.payroll_combo.setVisible(True)
            self._auto_fill_from_payroll()
        else:
            self.payroll_combo.setVisible(False)
            self._update_description()

    def _on_employee_changed(self):
        """Personel değişince eğer ödeme tipi maaşsa otomatik doldur"""
        if self.payment_type_combo.currentText() == "Maas":
            self._auto_fill_from_payroll()
        else:
            self._update_description()

    def _on_payroll_selected(self, index):
        """Maaş bordrosu seçilince tutar ve açıklamayı güncelle"""
        if index <= 0:
            return
        payroll_id = self.payroll_combo.currentData()
        if payroll_id:
            row = self.db.fetch_one("SELECT net_salary, month, year FROM payroll WHERE id=?", (payroll_id,))
            if row:
                self.amount_spin.setValue(row[0])
                self._update_description()

    def _auto_fill_from_payroll(self):
        """Seçili personele ait en son maaşı bul ve otomatik seç"""
        emp_id = self.emp_combo.currentData()
        if not emp_id:
            return
        # Önce en son maaş bordrosunu bul
        row = self.db.fetch_one("""
            SELECT id, net_salary FROM payroll
            WHERE employee_id=? AND status!='معتمد'
            ORDER BY year DESC, month DESC LIMIT 1
        """, (emp_id,))
        if row:
            payroll_id, net = row
            index = self.payroll_combo.findData(payroll_id)
            if index >= 0:
                self.payroll_combo.setCurrentIndex(index)
                self.amount_spin.setValue(net)
        else:
            # Eğer bordro yoksa tutarı sıfırla
            self.amount_spin.setValue(0)
        self._update_description()

    def _update_amount_words(self):
        """Sayısal tutarı Türkçe yazıya çevir"""
        amount = self.amount_spin.value()
        self.amount_words.setText(number_to_words_tr(amount, "TL"))

    def _update_description(self):
        """Açıklama alanını otomatik doldur (kullanıcı düzenleyebilir)"""
        emp_name = self._get_receiver_name()
        ptype = self.payment_type_combo.currentText()
        desc = f"{emp_name} - {ptype}"
        if ptype == "Maas" and self.payroll_combo.currentData():
            text = self.payroll_combo.currentText()
            parts = text.split('-')
            if len(parts) > 1:
                desc += f" - {parts[1].strip()}"
        self.description.setText(desc)

    def _get_receiver_name(self):
        """Alıcı adını döndür"""
        if self.receiver_type_group.checkedId() == 2:
            return self.other_name.text().strip() or "Belirtilmemiş"
        else:
            return self.emp_combo.currentText()

    def _clear(self):
        """Tüm alanları temizle"""
        self.amount_spin.setValue(0)
        self.description.clear()
        self.notes.clear()
        self.other_name.clear()
        self.payment_date.setDate(QDate.currentDate())
        self.payroll_combo.setCurrentIndex(0)
        self.payment_type_combo.setCurrentIndex(0)  # "Diger" varsayılan
        self.receiver_type_group.button(1).setChecked(True)
        self._on_receiver_type_changed()

    # ========== Makbuz Tasarımı (Geliştirilmiş) ==========

    def _generate_receipt_html(self):
        """Makbuz için HTML kodu oluştur (düzenli tasarım)"""
        company_name = self.db.get_setting('company_name', 'Şirket')
        company_address = self.db.get_setting('company_address', '')
        company_phone = self.db.get_setting('company_phone', '')
        logo_path = self.db.get_setting('company_logo', '')

        # Logo için base64 veya dosya yolu
        logo_html = ""
        if logo_path and os.path.exists(logo_path):
            logo_html = f'<img src="{logo_path}" width="150" style="float:left; margin-right:20px;">'

        # Satırları oluştur
        lines = []
        lines.append(f"<b>Tarih:</b> {self.payment_date.date().toString('dd/MM/yyyy')}")
        lines.append(f"<b>Alıcı:</b> {self._get_receiver_name()}")
        lines.append(f"<b>Tutar:</b> {self.amount_spin.value():,.2f} TL")
        lines.append(f"<b>Tutar Yazıyla:</b> {self.amount_words.text()}")
        if self.description.toPlainText():
            lines.append(f"<b>Açıklama:</b> {self.description.toPlainText()}")
        if self.notes.text():
            lines.append(f"<b>Not:</b> {self.notes.text()}")

        # HTML şablonu (محسّن)
        html = f"""
        <!DOCTYPE html>
        <html dir="ltr">
        <head>
            <meta charset="UTF-8">
            <style>
                body {{
                    font-family: Arial, sans-serif;
                    margin: 20px;
                    line-height: 1.6;
                    background-color: #ffffff;
                }}
                .header {{
                    display: flex;
                    align-items: center;
                    margin-bottom: 30px;
                    border-bottom: 2px solid #1976D2;
                    padding-bottom: 10px;
                }}
                .company-info {{
                    flex: 1;
                    text-align: right;
                }}
                .company-info h2 {{
                    color: #1976D2;
                    margin: 0;
                }}
                .title {{
                    font-size: 16pt;
                    font-weight: bold;
                    text-align: center;
                    margin: 30px 0;
                    color: #1976D2;
                    text-transform: uppercase;
                    letter-spacing: 2px;
                    border-bottom: 1px dashed #1976D2;
                    padding-bottom: 10px;
                }}
                .content {{
                    margin: 30px 0;
                    padding: 20px;
                    border-radius: 8px;
                }}
                .row {{
                    margin: 12px 0;
                    font-size: 12pt;
                    background-color: #f0f0f0;  /* تظليل رمادي */
                    padding: 8px 12px;
                    border-radius: 4px;
                }}
                .row b {{
                    color: #1976D2;
                    min-width: 120px;
                    display: inline-block;
                }}
                .signature {{
                    margin-top: 50px;
                    text-align: center;
                    width: 300px;
                    margin-left: auto;
                    margin-right: auto;
                }}
                .signature-line {{
                    border-top: 1px solid #000;
                    margin-top: 30px;
                    padding-top: 5px;
                    width: 100%;
                }}
                .footer {{
                    margin-top: 200px;
                    font-size: 9px;
                    color: #888;
                    text-align: center;
                    border-top: 1px solid #ccc;
                    padding-top: 10px;
                }}
            </style>
        </head>
        <body>
            <div class="header">
                {logo_html}
                <div class="company-info">
                    <h2>{company_name}</h2>
                    <p>{company_address}<br>Tel: {company_phone}</p>
                </div>
            </div>
            <div class="title">ÖDEME MAKBUZU</div>
            <div class="content">
                {"".join(f'<div class="row">{line}</div>' for line in lines)}
            </div>
            <div class="signature">
                <div class="signature-line">Alıcı Adı Soyadı İmzası</div>
            </div>
            <div class="footer">
                Bu belge İnsan Kaynakları Sistemi tarafından oluşturulmuştur - {datetime.now().strftime('%Y-%m-%d %H:%M')}
            </div>
        </body>
        </html>
        """
        return html

    def _print_receipt(self):
        """Makbuzu yazdır"""
        if self.amount_spin.value() <= 0:
            QMessageBox.warning(self, "Hata", "Lütfen geçerli bir tutar girin")
            return
        if self.receiver_type_group.checkedId() == 2 and not self.other_name.text():
            QMessageBox.warning(self, "Hata", "Lütfen alıcı adını girin")
            return

        html = self._generate_receipt_html()

        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        dialog = QPrintDialog(printer, self)
        if dialog.exec_() != QDialog.Accepted:
            return

        document = QTextDocument()
        document.setHtml(html)
        document.print_(printer)

        QMessageBox.information(self, "Başarılı", "Makbuz başarıyla yazdırıldı")
        self._clear()

    def _preview_receipt(self):
        """Makbuz önizlemesi göster"""
        if self.amount_spin.value() <= 0:
            QMessageBox.warning(self, "Hata", "Lütfen geçerli bir tutar girin")
            return
        if self.receiver_type_group.checkedId() == 2 and not self.other_name.text():
            QMessageBox.warning(self, "Hata", "Lütfen alıcı adını girin")
            return

        html = self._generate_receipt_html()

        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        preview = QPrintPreviewDialog(printer, self)
        preview.paintRequested.connect(lambda p: self._handle_paint_request(p, html))
        preview.exec_()

    def _handle_paint_request(self, printer, html):
        document = QTextDocument()
        document.setHtml(html)
        document.print_(printer)