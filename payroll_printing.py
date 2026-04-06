    def _print_payslip(self):
        m = self.current_month.currentIndex() + 1
        y = self.current_year.value()

        data = self._fetch_payslip_data(m, y, 'معتمد')
        if not data:
            reply = QMessageBox.question(
                self, "تنبيه",
                "لا توجد رواتب معتمدة.\n"
                "هل تريد طباعة قصاصات المسودة؟",
                QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.No:
                return
            data = self._fetch_payslip_data(m, y, 'مسودة')

        if not data:
            QMessageBox.warning(self, "تنبيه", "لا توجد بيانات رواتب")
            return

        self._render_payslips(data, m, y)

    def _fetch_payslip_data(self, m: int, y: int, status: str) -> list:
    def _fetch_payslip_data(self, m: int, y: int, status: str) -> list:
        return self.db.fetch_all("""
            SELECT p.id,
                   p.basic_salary,
                   p.housing_allowance, p.transportation_allowance,
                   p.food_allowance, p.phone_allowance, p.other_allowances,
                   p.overtime_amount, p.bonus,
                   p.total_earnings,
                   p.absence_deduction, p.late_deduction,
                   COALESCE(p.unpaid_leave_deduction, 0),
                   p.loan_deduction_bank, p.loan_deduction_cash,
                   p.total_deductions, p.net_salary,
                   p.bank_salary, p.cash_salary,
                   p.notes, p.status,
                   e.first_name || ' ' || e.last_name,
                   e.employee_code,
                   COALESCE(d.name, ''),
                   COALESCE(e.iban, ''),
                   COALESCE(e.bank_name, '')
            FROM payroll p
            JOIN employees e ON p.employee_id = e.id
            LEFT JOIN departments d ON e.department_id = d.id
            WHERE p.month=? AND p.year=? AND p.status=?
            ORDER BY e.first_name
        """, (m, y, status))

    def _render_payslips(self, data: list, m: int, y: int):
    def _render_payslips(self, data: list, m: int, y: int):
        months_tr   = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
                        'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
        month_name  = months_tr[m - 1]
        company_name    = self.db.get_setting('company_name', 'Şirket')
        company_address = self.db.get_setting('company_address', '')
        company_phone   = self.db.get_setting('company_phone', '')
        logo_path       = self.db.get_setting('company_logo', '')

        logo_html = (f'<img src="{logo_path}" width="120" '
                     f'style="float:left; margin-right:15px;">'
                     if logo_path and os.path.exists(logo_path) else "")

        all_html = ""
        for row in data:
            try:
                (pay_id, basic, housing, transport, food, phone, other,
                 ot, bonus, total_earn, absence, late, unpaid,
                 loan_bank, loan_cash, total_ded, net,
                 bank_sal, cash_sal, notes, status,
                 emp_name, emp_code, dept, iban, bank_name) = row

                all_allow = ((housing or 0) + (transport or 0) + (food or 0)
                             + (phone or 0) + (other or 0))

                html = f"""
                <!DOCTYPE html><html dir="ltr">
                <head><meta charset="UTF-8">
                <style>
                body{{font-family:Arial;margin:0;padding:8px;font-size:9pt;}}
                .header{{display:flex;align-items:center;border-bottom:2px solid #1976D2;
                          padding-bottom:6px;margin-bottom:8px;}}
                .co-info{{flex:1;text-align:right;}}
                .co-info h2{{margin:0;font-size:13pt;color:#1976D2;}}
                .co-info p{{margin:2px 0;font-size:8pt;color:#555;}}
                .title{{text-align:center;font-size:12pt;font-weight:bold;color:#1976D2;
                         margin:6px 0;border-bottom:1px dashed #1976D2;padding-bottom:4px;}}
                .grid{{display:grid;grid-template-columns:1fr 1fr;gap:4px;margin:6px 0;}}
                .section{{background:#f5f9ff;border:1px solid #cce0ff;
                           border-radius:4px;padding:6px;}}
                .section h4{{margin:0 0 4px 0;color:#1976D2;font-size:9pt;
                              border-bottom:1px solid #cce0ff;padding-bottom:2px;}}
                .row{{display:flex;justify-content:space-between;padding:2px 0;font-size:8.5pt;}}
                .row .lbl{{color:#555;}}.row .val{{font-weight:bold;}}
                .total-row{{display:flex;justify-content:space-between;
                             background:#1976D2;color:white;padding:4px 6px;
                             border-radius:3px;margin-top:4px;font-weight:bold;}}
                .net-box{{background:#e8f5e9;border:1px solid #66bb6a;
                           border-radius:4px;padding:6px;margin:6px 0;text-align:center;}}
                .net-box .net-val{{font-size:14pt;font-weight:bold;color:#2e7d32;}}
                .sig{{display:flex;justify-content:space-around;margin-top:16px;font-size:8pt;}}
                .sig-box{{text-align:center;width:140px;}}
                .sig-line{{border-top:1px solid #333;margin-top:25px;padding-top:3px;}}
                .footer{{margin-top:8px;font-size:7pt;color:#999;text-align:center;
                          border-top:1px solid #eee;padding-top:4px;}}
                </style></head><body>
                <div class="header">{logo_html}
                    <div class="co-info">
                        <h2>{company_name}</h2>
                        <p>{company_address}</p><p>Tel: {company_phone}</p>
                    </div>
                </div>
                <div class="title">MAAŞ BORDROSU — {month_name} {y}</div>
                <div class="section" style="margin-bottom:6px;">
                    <div class="row"><span class="lbl">Personel:</span>
                        <span class="val">{emp_name} ({emp_code})</span></div>
                    <div class="row"><span class="lbl">Departman:</span>
                        <span class="val">{dept}</span></div>
                    <div class="row"><span class="lbl">Banka / IBAN:</span>
                        <span class="val">{bank_name} — {iban}</span></div>
                </div>
                <div class="grid">
                    <div class="section"><h4>Kazançlar</h4>
                        <div class="row"><span class="lbl">Temel Maaş</span>
                            <span class="val">{basic or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Yardımlar</span>
                            <span class="val">{all_allow:,.2f}</span></div>
                        <div class="row"><span class="lbl">Fazla Mesai</span>
                            <span class="val">{ot or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Prim</span>
                            <span class="val">{bonus or 0:,.2f}</span></div>
                        <div class="total-row"><span>Toplam Kazanç</span>
                            <span>{total_earn or 0:,.2f}</span></div>
                    </div>
                    <div class="section"><h4>Kesintiler</h4>
                        <div class="row"><span class="lbl">Devamsızlık</span>
                            <span class="val">{absence or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Geç Kalma / Erken Çıkış</span>
                            <span class="val">{late or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Ücretsiz İzin</span>
                            <span class="val">{unpaid or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Avans (Banka)</span>
                            <span class="val">{loan_bank or 0:,.2f}</span></div>
                        <div class="row"><span class="lbl">Avans (Nakit)</span>
                            <span class="val">{loan_cash or 0:,.2f}</span></div>
                        <div class="total-row"><span>Toplam Kesinti</span>
                            <span>{total_ded or 0:,.2f}</span></div>
                    </div>
                </div>
                <div class="net-box">
                    <div style="font-size:9pt;color:#555;margin-bottom:2px;">Net Ödeme</div>
                    <div class="net-val">{net or 0:,.2f}</div>
                    <div style="font-size:8pt;color:#388e3c;margin-top:2px;">
                        Banka: {bank_sal or 0:,.2f} &nbsp;|&nbsp; Nakit: {cash_sal or 0:,.2f}
                    </div>
                </div>
                <div class="sig">
                    <div class="sig-box"><div class="sig-line">Personel İmzası</div></div>
                    <div class="sig-box"><div class="sig-line">Muhasebe İmzası</div></div>
                    <div class="sig-box"><div class="sig-line">Yetkili İmzası</div></div>
                </div>
                <div class="footer">Bu belge İnsan Kaynakları Sistemi tarafından oluşturulmuştur —
                    {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
                </body></html>"""
                all_html += html + "<div style='page-break-after:always;'></div>"
            except Exception as e:
                logger.error("خطأ في إنشاء قصاصة: %s", e)

        if not all_html:
            QMessageBox.warning(self, "خطأ", "لم يتم إنشاء أي قصاصة.")
            return

        printer = QPrinter(QPrinter.HighResolution)
        printer.setPageSize(QPrinter.A4)
        printer.setOrientation(QPrinter.Portrait)
        printer.setPageMargins(5, 5, 5, 5, QPrinter.Millimeter)
        dlg = QPrintDialog(printer, self)
        if dlg.exec_() == QDialog.Accepted:
            doc = QTextDocument()
            doc.setHtml(all_html)
            doc.print_(printer)

    # ==================== طباعة كشف الرواتب ====================
    def _print_payroll_sheet(self, table, month_combo, year_spin,
    def _print_payroll_sheet(self, table, month_combo, year_spin,
                              dept_filter=None, payment_filter=None):
        try:
            m          = month_combo.currentIndex() + 1
            y          = year_spin.value()
            months_ar  = ['يناير','فبراير','مارس','أبريل','مايو','يونيو',
                           'يوليو','أغسطس','سبتمبر','أكتوبر','نوفمبر','ديسمبر']
            month_name = months_ar[m - 1]
            company_name    = self.db.get_setting('company_name', 'الشركة')
            company_address = self.db.get_setting('company_address', '')
            company_phone   = self.db.get_setting('company_phone', '')
            logo_path       = self.db.get_setting('company_logo', '')

            logo_html    = (f'<img src="{logo_path}" width="80" style="float:left;">'
                            if logo_path and os.path.exists(logo_path) else "")
            filter_info  = f"الشهر: {month_name} {y}"
            if dept_filter and dept_filter.currentData():
                filter_info += f" | القسم: {dept_filter.currentText()}"
            if payment_filter and payment_filter.currentText() != "الكل":
                filter_info += f" | طريقة الدفع: {payment_filter.currentText()}"

            headers     = [table.horizontalHeaderItem(c).text()
                           for c in range(1, table.columnCount())
                           if not table.isColumnHidden(c)]
            visible_cols = [c for c in range(1, table.columnCount())
                            if not table.isColumnHidden(c)]
            rows_data, totals = [], {}
            for r in range(table.rowCount()):
                vals = []
                for c in visible_cols:
                    item = table.item(r, c)
                    text = item.text() if item else ""
                    vals.append(text)
                    try:
                        totals[c] = totals.get(c, 0.0) + float(
                            text.replace(',', '').replace(' ', ''))
                    except Exception:
                        pass
                rows_data.append(vals)

            total_row = []
            for i, c in enumerate(visible_cols):
                if i == 0:
                    total_row.append("الإجمالي")
                elif c in totals:
                    total_row.append(f"{totals[c]:,.2f}")
                else:
                    total_row.append("")

            html = f"""
            <!DOCTYPE html><html dir="rtl">
            <head><meta charset="UTF-8"><style>
            body{{font-family:Arial;margin:2px;padding:0;}}
            .header{{display:flex;align-items:center;margin-bottom:4px;
                      border-bottom:1px solid #1976D2;}}
            .co-info{{flex:1;text-align:right;}}
            .co-info h2{{margin:0;font-size:11pt;color:#1976D2;}}
            .co-info p{{margin:1px 0;font-size:7pt;color:#555;}}
            .title{{font-size:10pt;font-weight:bold;text-align:center;
                     margin:3px 0;color:#1976D2;}}
            .fi{{background:#f0f0f0;padding:2px 4px;font-size:7pt;margin:2px 0;}}
            table{{border-collapse:collapse;width:100%;font-size:6pt;}}
            th,td{{border:1px solid #aaa;padding:1px 2px;text-align:center;}}
            th{{background:#1976D2;color:white;}}
            .tr{{background:#e0e0e0;font-weight:bold;}}
            .footer{{margin-top:3px;font-size:5pt;color:#888;text-align:center;}}
            </style></head><body>
            <div class="header">{logo_html}
                <div class="co-info"><h2>{company_name}</h2>
                    <p>{company_address} — هاتف: {company_phone}</p></div>
            </div>
            <div class="title">كشف الرواتب</div>
            <div class="fi">{filter_info}</div>
            <table><thead><tr>{"".join(f"<th>{h}</th>" for h in headers)}</tr></thead>
            <tbody>"""

            for row in rows_data:
                html += "<tr>" + "".join(f"<td>{c}</td>" for c in row) + "</tr>"
            html += ('<tr class="tr">' +
                     "".join(f"<td>{c}</td>" for c in total_row) + "</tr>")
            html += f"""</tbody></table>
            <div class="footer">تم الإنشاء: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
            </body></html>"""

            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            printer.setOrientation(QPrinter.Landscape)
            printer.setPageMargins(3, 3, 3, 3, QPrinter.Millimeter)
            dlg = QPrintDialog(printer, self)
            if dlg.exec_() == QDialog.Accepted:
                doc = QTextDocument()
                doc.setHtml(html)
                doc.print_(printer)

        except Exception as e:
            QMessageBox.critical(self, "خطأ في الطباعة", str(e))

    def _print_current_payroll(self):
    def _print_receipt_tr(self, payroll_id: int, emp_name: str,
                           amount: float, month: int, year: int):
        try:
            currency   = self.db.get_setting('currency', 'TL')
            months_tr  = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
                           'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
            month_name = months_tr[month - 1] if 1 <= month <= 12 else str(month)
            company_name    = self.db.get_setting('company_name', 'Şirket')
            company_address = self.db.get_setting('company_address', '')
            company_phone   = self.db.get_setting('company_phone', '')
            logo_path       = self.db.get_setting('company_logo', '')

            logo_html = (f'<img src="{logo_path}" width="130" '
                         f'style="float:left; margin-right:15px;">'
                         if logo_path and os.path.exists(logo_path) else "")

            cash_row = self.db.fetch_one(
                "SELECT cash_salary, notes FROM payroll WHERE id=?",
                (payroll_id,))
            cash      = float(cash_row[0] or 0) if cash_row else 0.0
            pay_notes = cash_row[1] or "" if cash_row else ""

            bank_row = self.db.fetch_one(
                "SELECT bank_salary FROM payroll WHERE id=?", (payroll_id,))
            bank = float(bank_row[0] or 0) if bank_row else 0.0

            rounding_note_html = ""
            if "دُوِّر" in pay_notes:
                import re
                m_rnd = re.search(r'الراتب النقدي دُوِّر[^|]+', pay_notes)
                if m_rnd:
                    rounding_note_html = (
                        f'<div style="font-size:8pt;color:#e65100;margin-top:4px;">'
                        f'⚠ {m_rnd.group(0).strip()}</div>')

            words_cash = number_to_words_tr(cash, currency) if cash > 0 else "—"
            words_bank = number_to_words_tr(bank, currency) if bank > 0 else "—"

            html = f"""
            <!DOCTYPE html><html dir="ltr">
            <head><meta charset="UTF-8"><style>
            body{{font-family:Arial;margin:0;padding:14px;font-size:11pt;line-height:1.8;}}
            .header{{display:flex;align-items:center;border-bottom:2px solid #1976D2;
                      padding-bottom:10px;margin-bottom:14px;}}
            .co-info{{flex:1;text-align:right;}}
            .co-info h2{{margin:0;font-size:14pt;color:#1976D2;}}
            .co-info p{{margin:3px 0;font-size:9.5pt;color:#555;}}
            .title{{font-size:15pt;font-weight:bold;text-align:center;color:#1976D2;
                     margin:12px 0;border:2px solid #1976D2;padding:8px;
                     border-radius:4px;letter-spacing:1px;}}
            .box{{background:#f5f9ff;border:1px solid #cce0ff;
                   border-radius:6px;padding:14px 18px;margin:10px 0;}}
            .row{{display:flex;justify-content:space-between;padding:7px 0;
                   border-bottom:1px dotted #ddd;font-size:11pt;}}
            .row:last-child{{border-bottom:none;}}
            .lbl{{color:#555;}}.val{{font-weight:bold;color:#1a237e;}}
            .words{{font-size:9.5pt;color:#388e3c;margin:3px 0 6px 0;
                     padding-right:4px;font-style:italic;}}
            .sig-area{{display:flex;justify-content:space-between;margin-top:40px;}}
            .sig-box{{text-align:center;width:43%;}}
            .sig-line{{border-top:1px solid #333;margin-top:36px;
                        padding-top:6px;font-size:10pt;}}
            .footer{{margin-top:20px;font-size:7.5pt;color:#aaa;text-align:center;
                      border-top:1px solid #eee;padding-top:6px;}}
            </style></head><body>
            <div class="header">{logo_html}
                <div class="co-info"><h2>{company_name}</h2>
                    <p>{company_address}</p><p>Tel: {company_phone}</p></div>
            </div>
            <div class="title">ÖDEME MAKBUZU</div>
            <div class="box">
                <div class="row"><span class="lbl">Tarih:</span>
                    <span class="val">{datetime.now().strftime('%d/%m/%Y')}</span></div>
                <div class="row"><span class="lbl">Alıcı:</span>
                    <span class="val">{emp_name}</span></div>
                <div class="row"><span class="lbl">Dönem:</span>
                    <span class="val">{month_name} {year}</span></div>
                <div class="row"><span class="lbl">Net Ödeme:</span>
                    <span class="val">{amount or 0:,.2f} {currency}</span></div>
                <div class="row"><span class="lbl">Nakit Ödeme:</span>
                    <span class="val">{cash:,.2f} {currency}</span></div>
                {rounding_note_html}
                <div class="words">{words_cash}</div>
                <div class="row"><span class="lbl">Banka Ödemesi:</span>
                    <span class="val">{bank:,.2f} {currency}</span></div>
                <div class="words">{words_bank}</div>
            </div>
            <div class="sig-area">
                <div class="sig-box">
                    <div class="sig-line">Alıcı Adı Soyadı / İmzası</div></div>
                <div class="sig-box">
                    <div class="sig-line">Yetkili Adı Soyadı / İmzası</div></div>
            </div>
            <div class="footer">Bu belge İnsan Kaynakları Sistemi tarafından oluşturulmuştur —
                {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
            </body></html>"""

            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            printer.setOrientation(QPrinter.Portrait)
            printer.setPageMargins(8, 8, 8, 8, QPrinter.Millimeter)
            dlg = QPrintDialog(printer, self)
            if dlg.exec_() == QDialog.Accepted:
                doc = QTextDocument()
                doc.setHtml(html)
                doc.print_(printer)

            self.db.execute_query("""
                UPDATE payroll
                SET notes = COALESCE(notes,'') || ' | طباعة إيصال: ' || ?
                WHERE id = ?
            """, (datetime.now().strftime('%Y-%m-%d %H:%M'), payroll_id))
            self._load_current()
            self._load_archived()

        except Exception as e:
            QMessageBox.critical(self, "Hata", f"Yazdırma hatası: {e}")

    # ==================== طباعة جميع الإيصالات ====================
    def _print_all_receipts(self):
    def _print_all_receipts(self):
        tab_idx = self.tabs.currentIndex()
        if tab_idx == 0:
            ids, month_combo, year_spin = (
                self._current_ids[:], self.current_month, self.current_year)
        else:
            ids, month_combo, year_spin = (
                self._archived_ids[:], self.archived_month, self.archived_year)

        if not ids:
            QMessageBox.warning(self, "تنبيه", "لا توجد رواتب لطباعة إيصالاتها")
            return
        if QMessageBox.question(
                self, "تأكيد",
                f"سيتم طباعة {len(ids)} إيصال دفع. متابعة؟",
                QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return

        try:
            currency     = self.db.get_setting('currency', 'TL')
            company_name = self.db.get_setting('company_name', 'Şirket')
            company_addr = self.db.get_setting('company_address', '')
            company_tel  = self.db.get_setting('company_phone', '')
            logo_path    = self.db.get_setting('company_logo', '')
            months_tr    = ['Ocak','Şubat','Mart','Nisan','Mayıs','Haziran',
                             'Temmuz','Ağustos','Eylül','Ekim','Kasım','Aralık']
            m            = month_combo.currentIndex() + 1
            y            = year_spin.value()
            month_name   = months_tr[m - 1]
            logo_html    = (f'<img src="{logo_path}" width="130" '
                            f'style="float:left; margin-right:15px;">'
                            if logo_path and os.path.exists(logo_path) else "")

            import re
            all_html = ""
            for pid in ids:
                row = self.db.fetch_one("""
                    SELECT e.first_name || ' ' || e.last_name,
                           p.net_salary, p.cash_salary,
                           p.bank_salary, p.notes
                    FROM payroll p
                    JOIN employees e ON p.employee_id = e.id
                    WHERE p.id = ?
                """, (pid,))
                if not row:
                    continue

                emp_name, net, cash, bank, pay_notes = row
                net       = float(net  or 0)
                cash      = float(cash or 0)
                bank      = float(bank or 0)
                pay_notes = pay_notes or ""

                words_cash = number_to_words_tr(cash, currency) if cash > 0 else "—"
                words_bank = number_to_words_tr(bank, currency) if bank > 0 else "—"

                rnd_note = ""
                m_rnd = re.search(r'الراتب النقدي دُوِّر[^|]+', pay_notes)
                if m_rnd:
                    rnd_note = (
                        f'<div style="font-size:8pt;color:#e65100;">'
                        f'⚠ {m_rnd.group(0).strip()}</div>')

                html = f"""
                <!DOCTYPE html><html dir="ltr">
                <head><meta charset="UTF-8"><style>
                body{{font-family:Arial;margin:0;padding:14px;font-size:11pt;}}
                .header{{display:flex;align-items:center;border-bottom:2px solid #1976D2;
                          padding-bottom:10px;margin-bottom:14px;}}
                .co-info{{flex:1;text-align:right;}}
                .co-info h2{{margin:0;font-size:14pt;color:#1976D2;}}
                .title{{font-size:15pt;font-weight:bold;text-align:center;color:#1976D2;
                         margin:12px 0;border:2px solid #1976D2;padding:8px;}}
                .box{{background:#f5f9ff;border:1px solid #cce0ff;
                       border-radius:6px;padding:14px 18px;margin:10px 0;}}
                .row{{display:flex;justify-content:space-between;padding:7px 0;
                       border-bottom:1px dotted #ddd;}}
                .lbl{{color:#555;}}.val{{font-weight:bold;color:#1a237e;}}
                .words{{font-size:9.5pt;color:#388e3c;font-style:italic;}}
                .sig-area{{display:flex;justify-content:space-between;margin-top:40px;}}
                .sig-box{{text-align:center;width:43%;}}
                .sig-line{{border-top:1px solid #333;margin-top:36px;padding-top:6px;}}
                .footer{{margin-top:20px;font-size:7.5pt;color:#aaa;text-align:center;}}
                </style></head><body>
                <div class="header">{logo_html}
                    <div class="co-info"><h2>{company_name}</h2>
                        <p>{company_addr}</p><p>Tel: {company_tel}</p></div>
                </div>
                <div class="title">ÖDEME MAKBUZU</div>
                <div class="box">
                    <div class="row"><span class="lbl">Tarih:</span>
                        <span class="val">{datetime.now().strftime('%d/%m/%Y')}</span></div>
                    <div class="row"><span class="lbl">Alıcı:</span>
                        <span class="val">{emp_name}</span></div>
                    <div class="row"><span class="lbl">Dönem:</span>
                        <span class="val">{month_name} {y}</span></div>
                    <div class="row"><span class="lbl">Net Ödeme:</span>
                        <span class="val">{net:,.2f} {currency}</span></div>
                    <div class="row"><span class="lbl">Nakit Ödeme:</span>
                        <span class="val">{cash:,.2f} {currency}</span></div>
                    {rnd_note}
                    <div class="words">{words_cash}</div>
                    <div class="row"><span class="lbl">Banka Ödemesi:</span>
                        <span class="val">{bank:,.2f} {currency}</span></div>
                    <div class="words">{words_bank}</div>
                </div>
                <div class="sig-area">
                    <div class="sig-box">
                        <div class="sig-line">Alıcı Adı Soyadı / İmzası</div></div>
                    <div class="sig-box">
                        <div class="sig-line">Yetkili Adı Soyadı / İmzası</div></div>
                </div>
                <div class="footer">Bu belge İnsan Kaynakları Sistemi tarafından oluşturulmuştur —
                    {datetime.now().strftime('%d/%m/%Y %H:%M')}</div>
                </body></html>"""
                all_html += html + "<div style='page-break-after:always;'></div>"

            if not all_html:
                QMessageBox.warning(self, "خطأ", "لم يتم إنشاء أي إيصال")
                return

            printer = QPrinter(QPrinter.HighResolution)
            printer.setPageSize(QPrinter.A4)
            printer.setOrientation(QPrinter.Portrait)
            printer.setPageMargins(8, 8, 8, 8, QPrinter.Millimeter)
            dlg = QPrintDialog(printer, self)
            if dlg.exec_() == QDialog.Accepted:
                doc = QTextDocument()
                doc.setHtml(all_html)
                doc.print_(printer)

        except Exception as e:
            QMessageBox.critical(self, "خطأ", f"خطأ في الطباعة الجماعية:\n{e}")

    # ==================== صيانة الرواتب ====================
    def _run_maintenance(self):
