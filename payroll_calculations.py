            print(">>> payroll_calculations loaded")

    def _calculate(self):
        m = self.current_month.currentIndex() + 1
        y = self.current_year.value()
            print(">>> _calculate from payroll_calculations")


        # ============================================================
        # Guard: التحقق من وجود حضور معتمد للشهر المحدد
        # ============================================================
        approved_att = self.db.fetch_one(
            """SELECT COUNT(*)
               FROM attendance
               WHERE is_approved = 1
                 AND strftime('%m', punch_date) = ?
                 AND strftime('%Y', punch_date) = ?""",
            (f"{m:02d}", str(y)))
        approved_att_count = approved_att[0] if approved_att else 0

        if approved_att_count == 0:
            QMessageBox.warning(
                self, "لا يمكن الحساب",
                f"لا توجد سجلات حضور معتمدة لشهر "
                f"{self.current_month.currentText()} {y}.\n\n"
                "يجب اعتماد سجلات الحضور أولاً من تبويب الحضور "
                "← مسودة الحضور ← زر 'اعتماد الفترة'.")
            return
        # ============================================================

        # التعامل مع الرواتب الموجودة
        approved_count = (self.db.fetch_one(
            "SELECT COUNT(*) FROM payroll "
            "WHERE month=? AND year=? AND status='معتمد'",
            (m, y)) or [0])[0]

        if approved_count > 0:
            if QMessageBox.question(
                    self, "تأكيد",
                    f"يوجد {approved_count} راتب معتمد لهذا الشهر.\n"
                    "سيتم حذف المسودات فقط والإبقاء على المعتمدة.\nمتابعة؟",
                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                return
            self.db.execute_query("""
                UPDATE installments SET payroll_id=NULL, paid_amount=0, notes=NULL
                WHERE payroll_id IN (
                    SELECT id FROM payroll WHERE month=? AND year=? AND status='مسودة'
                )""", (m, y))
            self.db.execute_query(
                "DELETE FROM payroll WHERE month=? AND year=? AND status='مسودة'",
                (m, y))
        else:
            existing = (self.db.fetch_one(
                "SELECT COUNT(*) FROM payroll WHERE month=? AND year=?",
                (m, y)) or [0])[0]
            if existing > 0:
                if QMessageBox.question(
                        self, "تأكيد",
                        "يوجد رواتب لهذا الشهر. إعادة الحساب ستحذفها.\nمتابعة؟",
                        QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
                    return
                self.db.execute_query("""
                    UPDATE installments SET payroll_id=NULL, paid_amount=0, notes=NULL
                    WHERE payroll_id IN (
                        SELECT id FROM payroll WHERE month=? AND year=?
                    )""", (m, y))
                self.db.execute_query(
                    "DELETE FROM payroll WHERE month=? AND year=?", (m, y))

        employees = self.db.fetch_all("""
            SELECT id,
                   basic_salary,
                   housing_allowance, transportation_allowance,
                   food_allowance, phone_allowance, other_allowances,
                   COALESCE(bank_salary, 0),
                   COALESCE(cash_salary, 0),
                   is_exempt_from_fingerprint,
                   social_security_registered,
                   social_security_percent
            FROM employees WHERE status='نشط'
        """)

        if not employees:
            QMessageBox.warning(self, "تنبيه", "لا يوجد موظفون نشطون")
            return

        # جلب بيانات الحضور المعتمد فقط
        attendance_data   = self._get_attendance_data(m, y)
        unpaid_leave_data = self._get_unpaid_leave_data(m, y)
        installments_data = self._get_installments_data(m, y)

        progress = QProgressDialog(
            "جاري حساب الرواتب...", "إلغاء", 0, len(employees), self)
        progress.setWindowModality(Qt.WindowModal)

        ok = 0
        for idx, emp in enumerate(employees):
            progress.setValue(idx)
            if progress.wasCanceled():
                break

            (emp_id, basic, housing, transport, food, phone, other,
             bank_salary_orig, cash_salary_orig,
             is_exempt, ss_registered, ss_percent) = emp

            basic            = float(basic   or 0)
            housing          = float(housing or 0)
            transport        = float(transport or 0)
            food             = float(food    or 0)
            phone            = float(phone   or 0)
            other            = float(other   or 0)
            bank_salary_orig = float(bank_salary_orig or 0)

            wdm         = self.work_days_month  if self.work_days_month  > 0 else 26
            wdh         = self.work_hours_daily if self.work_hours_daily > 0 else 8
            daily_rate  = basic / wdm
            hourly_rate = basic / (wdm * wdh)

            att = attendance_data.get(emp_id, {
                'ot_hours': 0.0, 'absent_days': 0,
                'late_minutes': 0, 'early_minutes': 0})

            if is_exempt:
                att = {'ot_hours': 0.0, 'absent_days': 0,
                       'late_minutes': 0, 'early_minutes': 0}

            ot_hours     = float(att['ot_hours']   or 0)
            absent_hours = float(att['absent_days'] or 0) * wdh
            late_hours   = (float(att['late_minutes']  or 0) +
                            float(att['early_minutes'] or 0)) / 60.0

            net_ot, net_absence, net_late = self._calc_netting(
                ot_hours, absent_hours, late_hours)

            ot_amount         = net_ot      * hourly_rate * self.ot_rate
            absence_deduction = net_absence * hourly_rate * self.absence_rate
            late_deduction    = net_late    * hourly_rate * self.absence_rate

            unpaid = unpaid_leave_data.get(emp_id, {'days': 0.0, 'hours': 0.0})
            unpaid_deduction = (float(unpaid['days'])  * daily_rate +
                                float(unpaid['hours']) * hourly_rate)

            bonus         = 0.0
            allowances    = housing + transport + food + phone + other
            total_earnings = (basic + allowances + ot_amount + bonus
                              - absence_deduction - unpaid_deduction - late_deduction)

            inst_list = installments_data.get(emp_id, [])
            loan_bank = sum(i['due'] for i in inst_list if i['method'] == 'bank')
            loan_cash = sum(i['due'] for i in inst_list if i['method'] == 'cash')

            # الخصومات الكلية (بدون الأقساط — تُعرَض منفصلة)
            total_deductions = absence_deduction + unpaid_deduction + late_deduction

            net_payable = total_earnings - loan_bank - loan_cash

            bank_paid, cash_paid = self._calc_bank_cash(
                ss_registered, net_payable, bank_salary_orig, loan_bank)
            bank_paid = max(0.0, bank_paid)
            cash_paid = max(0.0, cash_paid)

            # تقريب الراتب النقدي لأقرب 100
            cash_rounding_note = ""
            if cash_paid > 0:
                cash_paid_orig = cash_paid
                cash_paid      = float(round(cash_paid / 100) * 100)
                if abs(cash_paid - cash_paid_orig) >= 0.5:
                    cash_rounding_note = (
                        f" | الراتب النقدي دُوِّر من {cash_paid_orig:,.0f}"
                        f" إلى {cash_paid:,.0f} (أقرب 100)")

            month_name = self.current_month.currentText()
            notes = f"تم حساب راتب شهر {month_name} {y}."
            if loan_bank + loan_cash > 0:
                notes += " تضمين أقساط شهرية."
            notes += cash_rounding_note

            self.db.execute_query("""
                INSERT INTO payroll (
                    employee_id, month, year,
                    basic_salary, housing_allowance, transportation_allowance,
                    food_allowance, phone_allowance, other_allowances,
                    overtime_hours, overtime_amount, bonus,
                    total_earnings,
                    absence_deduction, late_deduction, unpaid_leave_deduction,
                    loan_deduction_bank, loan_deduction_cash,
                    total_deductions, net_salary,
                    bank_salary, cash_salary,
                    notes, status, approved_at
                ) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                emp_id, m, y,
                basic, housing, transport, food, phone, other,
                ot_hours, ot_amount, bonus,
                total_earnings,
                absence_deduction, late_deduction, unpaid_deduction,
                loan_bank, loan_cash,
                total_deductions, net_payable,
                bank_paid, cash_paid,
                notes, 'مسودة', None
            ))

            payroll_id = self.db.last_id()

            for inst in inst_list:
                self.db.execute_query("""
                    UPDATE installments
                    SET payroll_id = ?,
                        notes = COALESCE(notes, '') ||
                                ' | مدرج في راتب شهر ' || ? || '/' || ?
                    WHERE id = ?
                """, (payroll_id, str(m), str(y), inst['inst_id']))

            ok += 1

        progress.setValue(len(employees))

        self.db.log_action("حساب رواتب", "payroll", None, None,
                           {"month": m, "year": y, "count": ok})
        if self.comm:
            self.comm.dataChanged.emit(
                'payroll', {'action': 'calculate', 'month': m, 'year': y})

        QMessageBox.information(
            self, "نجاح",
            f"تم حساب رواتب {ok} موظف بنجاح.\n"
            f"(استُخدمت {approved_att_count} سجل حضور معتمد)")
        self._load_current()

    # ==================== منطق المقاصة ====================
    @staticmethod
    def _calc_netting(ot_hours: float, absent_hours: float,
    def _calc_netting(ot_hours: float, absent_hours: float,
                      late_hours: float) -> tuple:
        """
        مقاصة بين الإضافي والغياب والتأخير.
        يُرجع (net_ot_hours, net_absence_hours, net_late_hours).
        """
        total_deficit = absent_hours + late_hours

        if ot_hours <= 0:
            return 0.0, absent_hours, late_hours
        if total_deficit <= 0:
            return ot_hours, 0.0, 0.0
        if abs(ot_hours - total_deficit) < 0.001:
            return 0.0, 0.0, 0.0
        if ot_hours > total_deficit:
            return ot_hours - total_deficit, 0.0, 0.0

        remaining = total_deficit - ot_hours
        if total_deficit > 0:
            abs_ratio  = absent_hours / total_deficit
            late_ratio = late_hours   / total_deficit
            return 0.0, remaining * abs_ratio, remaining * late_ratio
        return 0.0, 0.0, 0.0

    # ==================== توزيع الراتب ====================
    @staticmethod
    def _calc_bank_cash(ss_registered, net_payable: float,
    def _calc_bank_cash(ss_registered, net_payable: float,
                        bank_salary_orig: float, loan_bank: float) -> tuple:
        if not ss_registered:
            return 0.0, net_payable
        available_bank = bank_salary_orig - loan_bank
        if available_bank <= 0:
            return 0.0, net_payable
        if net_payable <= available_bank:
            return net_payable, 0.0
        return available_bank, net_payable - available_bank

    # ==================== جلب بيانات الحضور ====================
    def _get_attendance_data(self, month: int, year: int) -> dict:
