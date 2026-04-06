#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# utils.py

import constants                          # استيراد المودول كاملاً — لا القيمة
from enum import Enum
from typing import Optional, Callable, Dict, List, Any

from PyQt5.QtWidgets import (
    QTableWidget, QTableWidgetItem, QPushButton,
    QAbstractItemView, QHeaderView
)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

from translations import tr as translations_tr


# ==================== دوال الترجمة ====================
def tr(key: str, section: str = "general") -> str:
    """
    دالة ترجمة مبسطة تقرأ اللغة الحالية ديناميكياً من constants.

    الإصلاح: كان الكود يستورد (CURRENT_LANG) كقيمة ثابتة عند تحميل الملف،
    مما يعني أن تغيير اللغة لاحقاً لا ينعكس على الترجمات.
    الحل: قراءة (constants.CURRENT_LANG) في كل استدعاء مباشرةً من المودول.
    """
    return translations_tr(key, constants.CURRENT_LANG, section)


# ==================== أدوار المستخدمين ====================
class Role(str, Enum):
    ADMIN      = 'admin'
    HR         = 'hr'
    ACCOUNTANT = 'accountant'
    VIEWER     = 'viewer'


# ==================== دوال الجداول ====================
def make_table(columns: List[str], parent=None) -> QTableWidget:
    """إنشاء QTableWidget موحد الشكل والسلوك."""
    t = QTableWidget(parent)
    t.setColumnCount(len(columns))
    t.setHorizontalHeaderLabels(columns)
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setAlternatingRowColors(True)
    t.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
    t.verticalHeader().setVisible(False)
    t.setStyleSheet("font-size:12px;")
    return t


def fill_table(
    table: QTableWidget,
    data: List,
    colors: Optional[Dict[int, Callable]] = None,
    row_colors: Optional[Dict[int, str]] = None
) -> None:
    """
    ملء الجدول بالبيانات مع دعم اختياري لتلوين خلايا أو صفوف كاملة.

    المعاملات:
        table      : الجدول المراد ملؤه
        data       : قائمة الصفوف (كل صف قائمة أو tuple)
        colors     : {رقم_العمود: دالة(قيمة) → لون_نص أو None}
        row_colors : {رقم_الصف: لون_خلفية} — لتلوين صفوف بأكملها
    """
    table.setRowCount(0)
    for row_idx, row in enumerate(data):
        table.insertRow(row_idx)
        bg_color = row_colors.get(row_idx) if row_colors else None
        for col_idx, val in enumerate(row):
            item = QTableWidgetItem(str(val) if val is not None else "")
            item.setTextAlignment(Qt.AlignCenter)
            # تلوين النص حسب قيمة الخلية
            if colors and col_idx in colors:
                clr = colors[col_idx](val)
                if clr:
                    item.setForeground(QColor(clr))
            # تلوين خلفية الصف كاملاً
            if bg_color:
                item.setBackground(QColor(bg_color))
            table.setItem(row_idx, col_idx, item)


def btn(text: str, style: str, callback: Callable) -> QPushButton:
    """إنشاء زر موحد الشكل والسلوك."""
    b = QPushButton(text)
    b.setStyleSheet(style)
    b.setCursor(Qt.PointingHandCursor)
    b.clicked.connect(callback)
    return b


def get_selected_id(table: QTableWidget, id_column: int = 0) -> Optional[int]:
    """
    دالة مساعدة: إرجاع قيمة عمود المعرّف للصف المحدد في الجدول.
    تُقلِّل الكود المكرر في كل tab عند الحاجة لمعرفة الصف المختار.

    المعاملات:
        table     : الجدول
        id_column : رقم العمود الذي يحتوي على المعرف (افتراضي 0)

    الإرجاع: المعرف كـ int، أو None إذا لم يُحدَّد صف.
    """
    row = table.currentRow()
    if row < 0:
        return None
    item = table.item(row, id_column)
    if not item or not item.text():
        return None
    try:
        return int(item.text())
    except ValueError:
        return None


# ==================== دوال التحقق من الصلاحيات ====================
def can_add(user_role: str) -> bool:
    """هل يملك الدور صلاحية الإضافة؟"""
    return user_role in (Role.ADMIN, Role.HR)


def can_edit(user_role: str) -> bool:
    """هل يملك الدور صلاحية التعديل؟"""
    return user_role in (Role.ADMIN, Role.HR)


def can_delete(user_role: str) -> bool:
    """هل يملك الدور صلاحية الحذف؟ (المدير فقط)"""
    return user_role == Role.ADMIN


def can_approve(user_role: str) -> bool:
    """هل يملك الدور صلاحية الاعتماد؟"""
    return user_role in (Role.ADMIN, Role.HR)


def can_process_payroll(user_role: str) -> bool:
    """هل يملك الدور صلاحية معالجة الرواتب؟"""
    return user_role in (Role.ADMIN, Role.ACCOUNTANT)


def can_manage_users(user_role: str) -> bool:
    """هل يملك الدور صلاحية إدارة المستخدمين؟ (المدير فقط)"""
    return user_role == Role.ADMIN


def can_view_reports(user_role: str) -> bool:
    """هل يملك الدور صلاحية عرض التقارير؟ (الجميع)"""
    return True


def can_export(user_role: str) -> bool:
    """هل يملك الدور صلاحية التصدير؟ (الجميع)"""
    return True


def apply_permissions_to_buttons(
    buttons: List[Optional[QPushButton]],
    user_role: str,
    allowed_roles: List[str]
) -> None:
    """إخفاء أو إظهار قائمة أزرار بناءً على دور المستخدم."""
    for b in buttons:
        if b is not None:
            b.setVisible(user_role in allowed_roles)


def filter_buttons_by_role(
    buttons_dict: Dict[Optional[QPushButton], List[str]],
    user_role: str
) -> None:
    """إخفاء أو إظهار أزرار بناءً على قاموس {زر: [الأدوار المسموحة]}."""
    for b, allowed_roles in buttons_dict.items():
        if b is not None:
            b.setVisible(user_role in allowed_roles)


# ==================== تفقيط تركي ====================
def number_to_words_tr(number: float, currency: str = "TL") -> str:
    """
    تحويل رقم إلى كتابة تركية (تفقيط).

    هذه هي النسخة الوحيدة المعتمدة في المشروع.
    لا تُضِف نسخة أخرى في أي ملف آخر — استورد هذه الدالة مباشرةً.

    الإصلاحات مقارنةً بالنسخة القديمة:
    - "bir milyon" بدلاً من "milyon" (الرقم 1,000,000 كان خاطئاً)
    - "bir milyar" بدلاً من "milyar" (نفس المشكلة)
    - مسافات صحيحة بين الأجزاء في جميع الحالات

    أمثلة:
        1_000     → "bin TL"
        1_001     → "bin bir TL"
        15_800.50 → "on beş bin sekiz yüz virgül 50 TL"
        1_000_000 → "bir milyon TL"
    """
    if number is None:
        return f"Sıfır {currency}"

    number = round(float(number), 2)

    if number == 0:
        return f"Sıfır {currency}"

    negative        = number < 0
    number          = abs(number)
    integer_part    = int(number)
    fractional_part = round((number - integer_part) * 100)

    birler   = ["", "bir", "iki", "üç", "dört", "beş",
                "altı", "yedi", "sekiz", "dokuz"]
    onlar    = ["", "on", "yirmi", "otuz", "kırk", "elli",
                "altmış", "yetmiş", "seksen", "doksan"]
    birimler = ["", "bin", "milyon", "milyar", "trilyon"]

    def _uc_basamak(n: int) -> str:
        """تحويل عدد من 1 إلى 999 إلى كلمات."""
        if n == 0:
            return ""
        s     = ""
        yuz   = n // 100
        kalan = n % 100
        on_d  = kalan // 10
        bir_d = kalan % 10
        if yuz == 1:
            s = "yüz"
        elif yuz > 1:
            s = birler[yuz] + " yüz"
        if on_d > 0:
            s += (" " if s else "") + onlar[on_d]
        if bir_d > 0:
            s += (" " if s else "") + birler[bir_d]
        return s.strip()

    # تقسيم الرقم لمجموعات ثلاثية
    gruplar: List[int] = []
    temp = integer_part
    while temp > 0:
        gruplar.append(temp % 1000)
        temp //= 1000
    if not gruplar:
        gruplar = [0]

    sonuc: List[str] = []
    for i in reversed(range(len(gruplar))):
        g = gruplar[i]
        if g == 0:
            continue
        okunus = _uc_basamak(g)
        if i == 0:
            # الجزء الأصغر (الآحاد إلى المئات)
            sonuc.append(okunus)
        elif i == 1:
            # الألوف: "bin" بدلاً من "bir bin"
            sonuc.append("bin" if g == 1 else okunus + " bin")
        else:
            # الملايين وما فوق: "bir milyon"، "iki milyon"، إلخ
            # الإصلاح: g==1 يُعطي "bir milyon" لا "milyon" فقط
            prefix = "bir" if g == 1 else okunus
            sonuc.append(prefix + " " + birimler[i])

    yazi = " ".join(sonuc).strip()

    if negative:
        yazi = "eksi " + yazi
    if fractional_part > 0:
        yazi += f" virgül {fractional_part:02d}"

    return f"{yazi} {currency}"
