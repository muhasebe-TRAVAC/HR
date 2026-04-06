class PayrollMaintenanceDialog(QDialog):
    CHECKS = [
        {
            "id":        "neg_net",
            "title":     "صافي الراتب سالب",
            "desc":      "صافي الراتب للدفع أقل من صفر",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, p.net_salary "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? AND p.net_salary < 0"),
            "fixable":   True,
            "fix_label": "ضبط صافي الراتب على 0",
            "fix_sql":   "UPDATE payroll SET net_salary=0, bank_salary=0, cash_salary=0 WHERE id=?",
        },
        {
            "id":        "bank_cash_mismatch",
            "title":     "مجموع بنكي + نقدي ≠ صافي الراتب",
            "desc":      "فارق > 1 بين net_salary و bank+cash",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, "
                          "p.net_salary, p.bank_salary, p.cash_salary "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND ABS(p.net_salary-(p.bank_salary+p.cash_salary)) > 1"),
            "fixable":   True,
            "fix_label": "تعديل نقدي = صافي − بنكي",
            "fix_sql":   "UPDATE payroll SET cash_salary=MAX(0,net_salary-bank_salary) WHERE id=?",
        },
        {
            "id":        "wrong_total_earn",
            "title":     "إجمالي الاستحقاق خاطئ",
            "desc":      "total_earnings ≠ أساسي+بدلات+OT+مكافأة−خصومات (فارق > 1)",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, p.total_earnings, "
                          "(p.basic_salary+p.housing_allowance+p.transportation_allowance+"
                          "p.food_allowance+p.phone_allowance+p.other_allowances+"
                          "p.overtime_amount+p.bonus-p.absence_deduction-p.late_deduction-"
                          "COALESCE(p.unpaid_leave_deduction,0)) AS exp "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND ABS(p.total_earnings-(p.basic_salary+p.housing_allowance+"
                          "p.transportation_allowance+p.food_allowance+p.phone_allowance+"
                          "p.other_allowances+p.overtime_amount+p.bonus-p.absence_deduction-"
                          "p.late_deduction-COALESCE(p.unpaid_leave_deduction,0)))>1"),
            "fixable":   True,
            "fix_label": "إعادة حساب إجمالي الاستحقاق",
            "fix_sql":   ("UPDATE payroll SET total_earnings="
                          "basic_salary+housing_allowance+transportation_allowance+"
                          "food_allowance+phone_allowance+other_allowances+"
                          "overtime_amount+bonus-absence_deduction-late_deduction-"
                          "COALESCE(unpaid_leave_deduction,0) WHERE id=?"),
        },
        {
            "id":        "wrong_total_ded",
            "title":     "إجمالي الخصومات خاطئ",
            "desc":      "total_deductions ≠ غياب+تأخير+بدون راتب (فارق > 1)",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, "
                          "p.total_deductions, "
                          "(p.absence_deduction+p.late_deduction+"
                          "COALESCE(p.unpaid_leave_deduction,0)) AS exp "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND ABS(p.total_deductions-(p.absence_deduction+p.late_deduction+"
                          "COALESCE(p.unpaid_leave_deduction,0)))>1"),
            "fixable":   True,
            "fix_label": "إعادة حساب إجمالي الخصومات",
            "fix_sql":   ("UPDATE payroll SET total_deductions="
                          "absence_deduction+late_deduction+"
                          "COALESCE(unpaid_leave_deduction,0) WHERE id=?"),
        },
        {
            "id":        "unregistered_has_bank",
            "title":     "موظف غير مسجل بالتأمينات لديه راتب بنكي",
            "desc":      "الموظف غير المسجل يجب أن يكون راتبه نقدياً كاملاً",
            "query":     ("SELECT p.id, e.first_name||' '||e.last_name, p.bank_salary "
                          "FROM payroll p JOIN employees e ON p.employee_id=e.id "
                          "WHERE p.month=? AND p.year=? "
                          "AND (e.social_security_registered=0 "
                          "OR e.social_security_registered IS NULL) "
                          "AND p.bank_salary > 0"),
            "fixable":   True,
            "fix_label": "نقل البنكي إلى النقدي",
            "fix_sql":   "UPDATE payroll SET cash_salary=net_salary, bank_salary=0 WHERE id=?",
        },
        {
            "id":        "active_no_payroll",
            "title":     "موظفون نشطون بدون راتب",
            "desc":      "موظفون بحالة نشط لم يُحسب لهم راتب هذا الشهر",
            "query":     ("SELECT e.id, e.first_name||' '||e.last_name, e.basic_salary "
                          "FROM employees e WHERE e.status='نشط' "
                          "AND e.id NOT IN "
                          "(SELECT employee_id FROM payroll WHERE month=? AND year=?)"),
            "fixable":   False,
        },
        {
            "id":        "orphan_inst",
            "title":     "أقساط مرتبطة برواتب غير موجودة",
            "desc":      "installments.payroll_id يشير لراتب محذوف",
            "query":     None,
            "fixable":   True,
            "fix_label": "فك ربط الأقساط اليتيمة",
            "fix_sql":   ("UPDATE installments SET payroll_id=NULL, paid_amount=0, "
                          "status='pending', notes=NULL WHERE id=?"),
            "is_inst":   True,
        },
    ]

    def __init__(self, parent, db: DatabaseManager, month: int, year: int):
        super().__init__(parent)
        self.db    = db
        self.month = month
        self.year  = year
        months = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                  'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
        self.setWindowTitle(f"🔧 صيانة الرواتب — {months[month-1]} {year}")
        self.setMinimumSize(940, 640)
        self.setLayoutDirection(Qt.RightToLeft)
        self._results = {}
        self._build()
        self._run_all_checks()

    def _build(self):
        layout = QVBoxLayout(self)

        hdr = QLabel(
            f"فحص رواتب شهر {self.month}/{self.year}  "
            f"— {len(self.CHECKS)} قاعدة فحص")
        hdr.setStyleSheet(
            "font-weight:bold; font-size:13px; padding:8px; "
            "background:#e3f2fd; border-radius:4px;")
        layout.addWidget(hdr)

        self.tbl = QTableWidget()
        self.tbl.setColumnCount(5)
        self.tbl.setHorizontalHeaderLabels(
            ["الحالة", "القاعدة", "بيانات", "عدد", "إجراء"])
        hh = self.tbl.horizontalHeader()
        for col, mode in [(0, QHeaderView.ResizeToContents),
                          (1, QHeaderView.Stretch),
                          (2, QHeaderView.Stretch),
                          (3, QHeaderView.ResizeToContents),
                          (4, QHeaderView.ResizeToContents)]:
            hh.setSectionResizeMode(col, mode)
        self.tbl.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tbl.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.tbl.setWordWrap(True)
        layout.addWidget(self.tbl)

        btn_row      = QHBoxLayout()
        self.btn_recheck = btn("🔄 إعادة الفحص",      BTN_PRIMARY, self._run_all_checks)
        self.btn_fix_all = btn("🔧 إصلاح كل المشاكل", BTN_SUCCESS, self._fix_all)
        self.btn_close   = btn("✅ إغلاق",              BTN_GRAY,   self.accept)
        self.lbl_sum     = QLabel("")
        self.lbl_sum.setStyleSheet("font-weight:bold; padding:4px; font-size:12px;")
        btn_row.addWidget(self.btn_recheck)
        btn_row.addWidget(self.btn_fix_all)
        btn_row.addWidget(self.lbl_sum)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_close)
        layout.addLayout(btn_row)

    def _run_all_checks(self):
        self._results.clear()
        self.tbl.setRowCount(0)
        total = fixable_cnt = 0

        for check in self.CHECKS:
            rows = self._exec_check(check)
            self._results[check['id']] = rows
            self._add_row(check, rows)
            total += len(rows)
            if rows and check.get('fixable'):
                fixable_cnt += len(rows)

        if total == 0:
            self.lbl_sum.setText("✅ لا توجد مشاكل")
            self.lbl_sum.setStyleSheet(
                "font-weight:bold; color:#388E3C; padding:4px;")
        else:
            self.lbl_sum.setText(
                f"⚠️ {total} مشكلة — {fixable_cnt} قابلة للإصلاح التلقائي")
            self.lbl_sum.setStyleSheet(
                "font-weight:bold; color:#D32F2F; padding:4px;")

    def _exec_check(self, check: dict) -> list:
        try:
            if check['id'] == 'orphan_inst':
                return self.db.fetch_all(
                    "SELECT i.id, COALESCE(e.first_name||' '||e.last_name,'?'), i.amount "
                    "FROM installments i "
                    "JOIN loans l ON i.loan_id=l.id "
                    "JOIN employees e ON l.employee_id=e.id "
                    "WHERE i.payroll_id IS NOT NULL "
                    "AND i.payroll_id NOT IN (SELECT id FROM payroll)") or []
            return self.db.fetch_all(
                check['query'], (self.month, self.year)) or []
        except Exception as e:
            return [("ERR", f"خطأ في الفحص: {e}", "—")]

    def _add_row(self, check: dict, rows: list):
        r = self.tbl.rowCount()
        self.tbl.insertRow(r)

        if not rows:
            si = QTableWidgetItem("✅ سليم")
            si.setForeground(QColor("#388E3C"))
        elif not check.get('fixable'):
            si = QTableWidgetItem(f"ℹ️ {len(rows)}")
            si.setForeground(QColor("#1565C0"))
        else:
            si = QTableWidgetItem(f"❌ {len(rows)}")
            si.setForeground(QColor("#C62828"))
        self.tbl.setItem(r, 0, si)

        self.tbl.setItem(r, 1, QTableWidgetItem(
            f"{check['title']}\n{check.get('desc','')}"))

        if rows:
            names = [str(row[1]) for row in rows[:6]]
            if len(rows) > 6:
                names.append(f"...و{len(rows)-6} آخرين")
            self.tbl.setItem(r, 2, QTableWidgetItem("\n".join(names)))
        else:
            self.tbl.setItem(r, 2, QTableWidgetItem("—"))

        self.tbl.setItem(r, 3, QTableWidgetItem(str(len(rows))))

        if rows and check.get('fixable'):
            fb = QPushButton(f"🔧 {check.get('fix_label','إصلاح')}")
            fb.setStyleSheet(BTN_WARNING)
            fb.clicked.connect(lambda _, c=check: self._fix_one(c))
            self.tbl.setCellWidget(r, 4, fb)
        else:
            self.tbl.setItem(r, 4, QTableWidgetItem("—"))

        self.tbl.setRowHeight(r, max(48, 18 * min(len(rows), 6) + 12))

    def _fix_one(self, check: dict):
        rows = self._results.get(check['id'], [])
        if not rows:
            return
        if QMessageBox.question(
                self, "تأكيد",
                f"إصلاح {len(rows)} سجل بـ «{check['fix_label']}»؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        self._apply_fix(check, rows)

    def _fix_all(self):
        fixable = [c for c in self.CHECKS
                   if c.get('fixable') and self._results.get(c['id'])]
        if not fixable:
            QMessageBox.information(self, "تنبيه", "لا توجد مشاكل قابلة للإصلاح")
            return
        total = sum(len(self._results[c['id']]) for c in fixable)
        if QMessageBox.question(
                self, "تأكيد",
                f"إصلاح {total} مشكلة دفعة واحدة؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        try:
            with self.db.transaction():
                fixed = 0
                for check in fixable:
                    for row in self._results[check['id']]:
                        self.db.conn.execute(check['fix_sql'], (row[0],))
                        fixed += 1
            self.db.log_action(
                "صيانة رواتب شاملة", "payroll", None, None,
                {"month": self.month, "year": self.year, "fixed": fixed})
            QMessageBox.information(self, "نجاح", f"تم إصلاح {fixed} مشكلة")
            self._run_all_checks()
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"فشل:\n{e}")

    def _apply_fix(self, check: dict, rows: list):
        try:
            with self.db.transaction():
                for row in rows:
                    self.db.conn.execute(check['fix_sql'], (row[0],))
            self.db.log_action(
                f"صيانة: {check['title']}", "payroll", None, None,
                {"month": self.month, "year": self.year, "fixed": len(rows)})
            QMessageBox.information(self, "نجاح", f"تم إصلاح {len(rows)} سجل")
            self._run_all_checks()
        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"فشل:\n{e}")
