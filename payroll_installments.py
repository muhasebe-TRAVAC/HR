    def _get_installments_data(self, month: int, year: int) -> dict:
        rows = self.db.fetch_all("""
            SELECT i.id, i.loan_id,
                   (i.amount - COALESCE(i.paid_amount, 0)) AS due,
                   l.employee_id, l.payment_method
            FROM installments i
            JOIN loans l ON i.loan_id = l.id
            WHERE strftime('%Y-%m', i.due_date) = ?
              AND i.status     != 'paid'
              AND i.payroll_id IS NULL
        """, (f"{year}-{month:02d}",))

        data = {}
        for inst_id, loan_id, due, emp_id, method in rows:
            if not due or due <= 0:
                continue
            data.setdefault(emp_id, []).append({
                'inst_id': inst_id,
                'loan_id': loan_id,
                'due':     due,
                'method':  method or 'cash'
            })
        return data

    # ==================== الاعتماد ====================
    def _approve_selected(self):
class EditPayrollInstallmentsDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, user: dict,
                 payroll_id: int, comm=None):
        super().__init__(parent)
        self.db         = db
        self.user       = user
        self.payroll_id = payroll_id
        self.comm       = comm
        self.setWindowTitle("تعديل خصم السلف من الراتب")
        self.setMinimumSize(860, 560)
        self.setLayoutDirection(Qt.RightToLeft)
        if self._load_data():
            self._build()

    def _load_data(self) -> bool:
        self.payroll = self.db.fetch_one("""
            SELECT id, employee_id, month, year,
                   basic_salary, total_earnings,
                   absence_deduction, late_deduction,
                   COALESCE(unpaid_leave_deduction, 0),
                   loan_deduction_bank, loan_deduction_cash, net_salary
            FROM payroll WHERE id=?
        """, (self.payroll_id,))

        if not self.payroll:
            QMessageBox.critical(self, "خطأ", "لم يتم العثور على الراتب")
            self.reject()
            return False

        self.employee_id = self.payroll[1]
        emp = self.db.fetch_one(
            "SELECT first_name || ' ' || last_name FROM employees WHERE id=?",
            (self.employee_id,))
        self.emp_name = emp[0] if emp else "غير معروف"

        self.installments = self.db.fetch_all("""
            SELECT i.id, i.loan_id, i.due_date,
                   i.amount,
                   COALESCE(i.paid_amount, 0),
                   (i.amount - COALESCE(i.paid_amount, 0)) AS remaining,
                   l.loan_type, l.payment_method,
                   CASE WHEN i.payroll_id = ? THEN 1 ELSE 0 END AS linked
            FROM installments i
            JOIN loans l ON i.loan_id = l.id
            WHERE l.employee_id = ?
              AND l.status      = 'نشط'
              AND i.status     != 'paid'
              AND (i.payroll_id IS NULL OR i.payroll_id = ?)
            ORDER BY i.due_date
        """, (self.payroll_id, self.employee_id, self.payroll_id))
        return True

    def _build(self):
        layout = QVBoxLayout(self)

        info = QLabel(
            f"الموظف: {self.emp_name}  |  "
            f"الراتب الأساسي: {self.payroll[4]:,.2f}  |  "
            f"إجمالي الاستحقاق: {self.payroll[5]:,.2f}")
        info.setStyleSheet("font-weight:bold; padding:8px; background:#e3f2fd;")
        layout.addWidget(info)

        self.table = QTableWidget()
        self.table.setColumnCount(9)
        self.table.setHorizontalHeaderLabels([
            "تحديد", "نوع السلفة", "تاريخ الاستحقاق", "المبلغ الأصلي",
            "المدفوع", "المتبقي", "طريقة الدفع",
            "المبلغ المراد خصمه", "ملاحظات"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._populate_table()
        layout.addWidget(self.table)

        total_layout = QHBoxLayout()
        total_layout.addWidget(QLabel("إجمالي الخصم المحدد:"))
        self._total_lbl = QLabel("0.00")
        self._total_lbl.setStyleSheet(
            "font-weight:bold; color:#1976D2; font-size:14px;")
        total_layout.addWidget(self._total_lbl)
        total_layout.addStretch()
        layout.addLayout(total_layout)

        btn_layout = QHBoxLayout()
        btn_layout.addWidget(btn("💾 حفظ التغييرات", BTN_SUCCESS, self._save))
        btn_layout.addWidget(btn("❌ إلغاء",          BTN_DANGER,  self.reject))
        layout.addLayout(btn_layout)

        self._update_total()
        self.table.itemChanged.connect(self._on_check_changed)

    def _populate_table(self):
        self.table.setRowCount(len(self.installments))
        for row_idx, inst in enumerate(self.installments):
            (inst_id, loan_id, due_date, amount,
             paid, remaining, loan_type, method, linked) = inst

            chk = QTableWidgetItem()
            chk.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk.setCheckState(Qt.Checked if linked else Qt.Unchecked)
            self.table.setItem(row_idx, 0, chk)

            for col, val in [(1, loan_type or ""), (2, due_date or ""),
                              (3, f"{amount:,.2f}"), (4, f"{paid:,.2f}"),
                              (5, f"{remaining:,.2f}"),
                              (6, "بنكي" if method == 'bank' else "نقدي")]:
                self.table.setItem(row_idx, col, QTableWidgetItem(val))

            spin = QDoubleSpinBox()
            spin.setRange(0.0, max(remaining, 0.0))
            spin.setDecimals(2)
            spin.setValue(remaining if linked else 0.0)
            spin.valueChanged.connect(self._update_total)
            self.table.setCellWidget(row_idx, 7, spin)

            notes_item = QTableWidgetItem("")
            notes_item.setFlags(Qt.ItemIsEditable | Qt.ItemIsEnabled)
            self.table.setItem(row_idx, 8, notes_item)

    def _on_check_changed(self, item):
        if item.column() != 0:
            return
        row     = item.row()
        checked = item.checkState() == Qt.Checked
        spin    = self.table.cellWidget(row, 7)
        if spin:
            rem_item = self.table.item(row, 5)
            if rem_item:
                try:
                    rem = float(rem_item.text().replace(',', ''))
                    spin.setValue(rem if checked else 0.0)
                except Exception:
                    pass
        self._update_total()

    def _update_total(self):
        total = sum(
            self.table.cellWidget(r, 7).value()
            for r in range(self.table.rowCount())
            if self.table.cellWidget(r, 7))
        self._total_lbl.setText(f"{total:,.2f}")

    def _save(self):
        try:
            with self.db.transaction():
                cur   = self.db.conn.cursor()
                month = self.payroll[2]
                year  = self.payroll[3]

                cur.execute("""
                    UPDATE installments
                    SET payroll_id=NULL, paid_amount=0, notes=NULL
                    WHERE payroll_id=?
                """, (self.payroll_id,))

                loan_bank = loan_cash = 0.0

                for row_idx in range(self.table.rowCount()):
                    chk = self.table.item(row_idx, 0)
                    if not chk or chk.checkState() != Qt.Checked:
                        continue
                    inst_id = self.installments[row_idx][0]
                    method  = self.installments[row_idx][7]
                    spin    = self.table.cellWidget(row_idx, 7)
                    amount  = spin.value() if spin else 0.0
                    if amount <= 0:
                        continue

                    note_item = self.table.item(row_idx, 8)
                    note_txt  = note_item.text().strip() if note_item else ""
                    note_full = (f"مدرج في راتب شهر {month}/{year}"
                                 + (f" — {note_txt}" if note_txt else ""))

                    cur.execute("""
                        UPDATE installments
                        SET payroll_id=?, paid_amount=?, notes=?
                        WHERE id=?
                    """, (self.payroll_id, amount, note_full, inst_id))

                    if method == 'bank':
                        loan_bank += amount
                    else:
                        loan_cash += amount

                total_earnings   = self.payroll[5] or 0
                absence          = self.payroll[6] or 0
                late             = self.payroll[7] or 0
                unpaid           = self.payroll[8] or 0
                total_deductions = absence + late + unpaid + loan_bank + loan_cash
                net_salary       = total_earnings - loan_bank - loan_cash

                cur.execute("""
                    UPDATE payroll
                    SET loan_deduction_bank=?, loan_deduction_cash=?,
                        total_deductions=?, net_salary=?
                    WHERE id=?
                """, (loan_bank, loan_cash, total_deductions,
                      net_salary, self.payroll_id))

            QMessageBox.information(self, "نجاح", "تم تحديث خصم السلف بنجاح")
            self.accept()

        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"حدث خطأ:\n{e}")


# ===================================================================
# نافذة إدارة المكافأة
# ===================================================================
class ManageBonusDialog(QDialog):
