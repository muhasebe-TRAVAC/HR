class ManageBonusDialog(QDialog):
    def __init__(self, parent, db: DatabaseManager, user: dict, payroll_id: int):
        super().__init__(parent)
        self.db         = db
        self.user       = user
        self.payroll_id = payroll_id
        self._load_current_bonus()
        self.setWindowTitle("إدارة المكافأة")
        self.setFixedSize(400, 240)
        self.setLayoutDirection(Qt.RightToLeft)
        self._build()

    def _load_current_bonus(self):
        row = self.db.fetch_one(
            "SELECT bonus FROM payroll WHERE id=?", (self.payroll_id,))
        self.current_bonus = float(row[0] or 0) if row else 0.0

    def _build(self):
        layout   = QVBoxLayout(self)
        currency = self.db.get_setting('currency', 'ريال')

        current_lbl = QLabel(
            f"المكافأة الحالية: {self.current_bonus:,.2f} {currency}")
        current_lbl.setStyleSheet(
            "font-weight:bold; padding:8px; background:#fff9c4; "
            "border-radius:4px; font-size:13px;")
        layout.addWidget(current_lbl)

        form = QFormLayout()
        self.bonus_spin = QDoubleSpinBox()
        self.bonus_spin.setRange(0.0, 9_999_999.99)
        self.bonus_spin.setDecimals(2)
        self.bonus_spin.setSuffix(f" {currency}")
        self.bonus_spin.setValue(self.current_bonus)
        form.addRow("القيمة الجديدة للمكافأة:", self.bonus_spin)
        layout.addLayout(form)

        btn_layout = QHBoxLayout()
        btn_save   = QPushButton("💾 حفظ")
        btn_save.setStyleSheet(BTN_SUCCESS)
        btn_save.clicked.connect(self._save)

        btn_delete = QPushButton("🗑️ حذف المكافأة")
        btn_delete.setStyleSheet(BTN_DANGER)
        btn_delete.clicked.connect(self._delete)
        btn_delete.setEnabled(self.current_bonus > 0)

        btn_cancel = QPushButton("❌ إلغاء")
        btn_cancel.setStyleSheet(BTN_GRAY)
        btn_cancel.clicked.connect(self.reject)

        btn_layout.addWidget(btn_save)
        btn_layout.addWidget(btn_delete)
        btn_layout.addWidget(btn_cancel)
        layout.addLayout(btn_layout)

    def _save(self):
        self._apply(self.bonus_spin.value(), "تعديل مكافأة")

    def _delete(self):
        if QMessageBox.question(
                self, "تأكيد", "هل تريد حذف المكافأة وضبطها على صفر؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        self._apply(0.0, "حذف مكافأة")

    def _apply(self, new_bonus: float, action: str):
        row = self.db.fetch_one("""
            SELECT total_earnings, net_salary, bonus
            FROM payroll WHERE id=?
        """, (self.payroll_id,))
        if not row:
            return
        total_earn_old, net_old, bonus_old = row
        bonus_old = float(bonus_old  or 0)
        diff      = new_bonus - bonus_old
        new_total = float(total_earn_old or 0) + diff
        new_net   = float(net_old        or 0) + diff

        self.db.execute_query("""
            UPDATE payroll
            SET bonus=?, total_earnings=?, net_salary=?
            WHERE id=?
        """, (new_bonus, new_total, new_net, self.payroll_id))

        self.db.log_action(action, "payroll", self.payroll_id,
                           {"bonus": bonus_old}, {"bonus": new_bonus})
        self.accept()


# الاسم القديم للتوافق
AddBonusDialog = ManageBonusDialog


# ===================================================================
# نافذة صيانة الرواتب
# ===================================================================
class PayrollMaintenanceDialog(QDialog):
