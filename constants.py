# constants.py (بعد إضافة CURRENT_LANG)
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# constants.py

import os

# ==================== أسماء الملفات والمجلدات ====================
DATABASE_NAME       = 'hr_payroll.db'
DOCUMENTS_FOLDER    = 'employee_documents'
BACKUP_FOLDER       = 'backups'
PAYSLIPS_FOLDER     = 'payslips'
REPORTS_FOLDER      = 'reports'
COMPANY_LOGO_FOLDER = 'company_logo'

# إصدار البرنامج
APP_VERSION = "2.0"

# اللغة الحالية (يتم تعيينها من قاعدة البيانات عند التشغيل)
CURRENT_LANG = "ar"   # ar: عربي, tr: تركي


def setup_directories() -> None:
    """
    إنشاء المجلدات المطلوبة للتطبيق.
    تُستدعى مرة واحدة صراحةً من main.py عند بدء التشغيل،
    وليس عند استيراد الملف تلقائياً.
    """
    for folder in [DOCUMENTS_FOLDER, BACKUP_FOLDER,
                   PAYSLIPS_FOLDER, REPORTS_FOLDER, COMPANY_LOGO_FOLDER]:
        os.makedirs(folder, exist_ok=True)


# ==================== ستايل البرنامج ====================
STYLE = """
QMainWindow { background-color: #f5f5f5; }
QTabWidget::pane { border: 1px solid #ddd; background: white; }
QTabBar::tab {
    background: #e0e0e0; padding: 8px 16px; margin: 2px;
    border-radius: 4px; font-size: 12px;
}
QTabBar::tab:selected { background: #1976D2; color: white; font-weight: bold; }
QGroupBox {
    border: 1px solid #ccc; border-radius: 6px;
    margin-top: 10px; padding-top: 8px;
    font-weight: bold; color: #333;
}
QGroupBox::title { subcontrol-origin: margin; padding: 0 5px; }
QPushButton {
    padding: 6px 14px; border-radius: 4px;
    border: none; font-size: 12px; font-weight: bold;
}
QPushButton:hover { opacity: 0.85; }
QLineEdit, QComboBox, QDateEdit, QDoubleSpinBox, QSpinBox, QTextEdit {
    border: 1px solid #bbb; border-radius: 4px;
    padding: 4px 8px; background: white;
}
QLineEdit:focus, QComboBox:focus { border-color: #1976D2; }
QTableView, QTableWidget {
    gridline-color: #e0e0e0; alternate-background-color: #f9f9f9;
    selection-background-color: #1976D2; selection-color: white;
}
QHeaderView::section {
    background: #1976D2; color: white; padding: 6px;
    border: none; font-weight: bold;
}
QLabel { color: #333; }
"""

# ==================== ألوان الأزرار ====================
BTN_PRIMARY  = "background:#1976D2;color:white;padding:8px 16px;"
BTN_SUCCESS  = "background:#388E3C;color:white;padding:8px 16px;"
BTN_DANGER   = "background:#D32F2F;color:white;padding:8px 16px;"
BTN_WARNING  = "background:#F57C00;color:white;padding:8px 16px;"
BTN_PURPLE   = "background:#7B1FA2;color:white;padding:8px 16px;"
BTN_TEAL     = "background:#00796B;color:white;padding:8px 16px;"
BTN_GRAY     = "background:#616161;color:white;padding:8px 16px;"