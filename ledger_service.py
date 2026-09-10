"""
Invisible Automated Accounting Engine (GAAP Compliant)
All hooks are idempotent and ACID-wrapped. No changes to existing core models.
"""
import json
import os
import hashlib
from decimal import Decimal, ROUND_HALF_UP
from datetime import date, datetime, timedelta
from pathlib import Path

# ------------------------------------------------------------------------------
# 0. Core Utilities (Period Lock, Decimal Rounding, Batch Validation)
# ------------------------------------------------------------------------------
AED_PLACES = Decimal("0.01")

def round_aed(amount):
    """Round all monetary values to 2 decimals (AED standard)."""
    if not isinstance(amount, Decimal):
        try:
            amount = Decimal(str(amount))
        except:
            return Decimal("0.00")
    return amount.quantize(AED_PLACES, rounding=ROUND_HALF_UP)

# ------------------------------------------------------------------------------
# Expense Category Mapping (User-friendly dropdown -> Expense Account)
# ------------------------------------------------------------------------------
EXPENSE_CATEGORY_MAP = {
    "Salary/Wages": "5100",
    "Rent": "5200",
    "Utilities (Electricity/Water/Internet)": "5200",
    "Office Supplies": "5300",
    "Travel & Transportation": "5300",
    "Marketing & Advertising": "5300",
    "Other Administrative": "5300",
}

class LedgerService:
    """Main GAAP ledger service for JSON-based storage."""
    
    def __init__(self, data_manager):
        """
        data_manager: The EnhancedCloudDataManager instance (has load_json/save_json).
        """
        self.dm = data_manager
        self._ensure_files()
    
    def _ensure_files(self):
        for fname in ['chart_of_accounts.json', 'general_ledger.json', 'period_locks.json']:
            fpath = Path(self.dm.data_folder) / fname
            if not fpath.exists():
                with open(fpath, 'w') as f: json.dump([], f)

    # --------------------------------------------------------------------------
    # Reload hook (GAAP dashboard calls this on every refresh).
    # --------------------------------------------------------------------------
    def reload_from_dm(self):
        """No-op for this service because we NEVER cache GL / CoA in memory —
        every method reads straight from self.dm.load_json(...).  Kept for API
        symmetry so callers can always ask for a reload without branching."""
        return True

    # --------------------------------------------------------------------------
    # Period Lock Checks
    # --------------------------------------------------------------------------
    def is_period_locked(self, transaction_date):
        """Check if a transaction date falls within any LOCKED period."""
        if isinstance(transaction_date, str):
            transaction_date = date.fromisoformat(transaction_date)
        locks = self.dm.load_json('period_locks.json') or []
        for lock in locks:
            if not lock.get('is_locked', False):
                continue
            try:
                s = date.fromisoformat(lock['start_date'])
                e = date.fromisoformat(lock['end_date'])
                if s <= transaction_date <= e:
                    return True
            except Exception:
                continue
        return False

    # --------------------------------------------------------------------------
    # Batch Writer (GAAP Double-Entry + Immutability)
    # --------------------------------------------------------------------------
    def _ensure_balanced(self, entries):
        total_debits = sum(Decimal(str(e.get('debit_amount', 0))) for e in entries)
        total_credits = sum(Decimal(str(e.get('credit_amount', 0))) for e in entries)
        if total_debits != total_credits:
            raise ValueError(
                f"Unbalanced GL transaction: Debits ({total_debits}) != Credits ({total_credits})"
            )

    def _write_gl_batch(self, entries, username=None):
        """
        Low-level immutable batch writer.
        - Checks period lock on EVERY entry
        - Validates batch balance (Debits = Credits)
        - Appends ONLY (immutable GL: no edits)
        """
        if not entries:
            return []
        # Check period locks
        for e in entries:
            d = e.get('date') or datetime.now().strftime('%Y-%m-%d')
            if self.is_period_locked(d):
                raise PermissionError(
                    f"Cannot post to locked period: {d}. Contact finance to unlock this period."
                )
        # Validate double-entry balance
        self._ensure_balanced(entries)
        # Load current GL and assign entry IDs (auto-increment)
        gl = self.dm.load_json('general_ledger.json') or []
        current_max_id = max([ent.get('entry_id', 0) for ent in gl], default=0)
        # Save a backup snapshot for atomic rollback
        backup_snapshot = list(gl)
        try:
            created = []
            for e in entries:
                current_max_id += 1
                entry = {
                    'entry_id': current_max_id,
                    'date': e.get('date') or datetime.now().strftime('%Y-%m-%d'),
                    'account_code': e['account_code'],
                    'account_name': self._get_account_name(e['account_code']),
                    'debit_amount': str(round_aed(e.get('debit_amount', 0))),
                    'credit_amount': str(round_aed(e.get('credit_amount', 0))),
                    'reference_type': e.get('reference_type', 'ADJUSTMENT'),
                    'reference_id': str(e.get('reference_id', '')),
                    'description': e.get('description', ''),
                    'created_at': datetime.now().isoformat(),
                    'created_by': username or 'System'
                }
                # Validation: exactly one non-zero
                d = Decimal(entry['debit_amount'])
                c = Decimal(entry['credit_amount'])
                if (d > 0) == (c > 0):
                    raise ValueError("Each GL line must have exactly one non-zero (debit OR credit).")
                if d < 0 or c < 0:
                    raise ValueError("GL amounts cannot be negative.")
                gl.append(entry)
                created.append(entry)
            # Persist in one go (atomic-ish: backup exists)
            self.dm.save_json('general_ledger.json', gl)
            return created
        except Exception as ex:
            # Restore backup
            try:
                self.dm.save_json('general_ledger.json', backup_snapshot)
            except Exception:
                pass
            raise ex

    def _get_account_name(self, code):
        coa = self.dm.load_json('chart_of_accounts.json') or []
        for a in coa:
            if a.get('account_code') == code:
                return a.get('full_code_name', code)
        return code

    # --------------------------------------------------------------------------
    # Idempotency: Don't write duplicate entries if hook runs twice
    # --------------------------------------------------------------------------
    def _entry_exists(self, reference_type, reference_id, account_code=None):
        gl = self.dm.load_json('general_ledger.json') or []
        for e in gl:
            if e.get('reference_type') == reference_type and e.get('reference_id') == str(reference_id):
                if account_code is None or e.get('account_code') == account_code:
                    return True
        return False

    # --------------------------------------------------------------------------
    # HOOK 1: on_invoice_created (Revenue/AR + COGS/Inventory)
    # --------------------------------------------------------------------------
    def on_invoice_created(self, invoice_dict, username=None):
        """
        invoice_dict: the invoice dict (from invoices_data.json)
        Runs only once per invoice id (idempotent via reference_id check).
        """
        inv_id = (invoice_dict.get('invoice_id') or invoice_dict.get('invoice_no')
                  or invoice_dict.get('number') or str(invoice_dict.get('id', '')))
        inv_id = str(inv_id) if inv_id else ''
        if not inv_id or self._entry_exists('INVOICE', inv_id):
            return []
        inv_date = invoice_dict.get('date') or invoice_dict.get('invoice_date') or datetime.now().strftime('%Y-%m-%d')
        client = invoice_dict.get('client_name') or invoice_dict.get('customer') or invoice_dict.get('client') or 'Unknown Client'

        # ------------------------------------------------------------------
        # TOTAL lookup — accept every reasonable field name produced by any
        # caller (UI forms, CSV import, API, mobile sync, manual JSON edit).
        # If NO numeric total field exists, fall back to recomputing from
        # item lines + VAT − discount, exactly like the UI does.
        # This is the ROOT-CAUSE FIX for user complaint: "1 invoice added
        # but GAAP never recognised it" — the invoice used `total_amount`
        # but this method only recognised `grand_total` / `total` / `amount`.
        # ------------------------------------------------------------------
        def _grab_total(*keys):
            for k in keys:
                try:
                    v = invoice_dict.get(k)
                    if v is None or v == "": continue
                    n = float(str(v).replace(",",""))
                    if n >= 0: return round_aed(n)
                except Exception: continue
            return None
        total = (_grab_total('total_amount','grand_total','total','amount',
                             'invoice_total','net_total','netAmount')
                 or round_aed(0))
        # Fallback: if all total keys missing, recompute from items
        if total <= 0:
            items = invoice_dict.get('items') or invoice_dict.get('line_items') or []
            calc = Decimal("0")
            if isinstance(items, list):
                for it in items:
                    if not isinstance(it, dict): continue
                    try:
                        q = float(it.get('qty') or it.get('quantity') or 0)
                        p = float(it.get('unit_price') or it.get('price') or it.get('rate') or 0)
                        line_excl_vat = q * p
                        try:
                            d = float(it.get('discount_amount') or it.get('discount') or 0)
                            if d > line_excl_vat: d = line_excl_vat
                            line_excl_vat = max(0.0, line_excl_vat - d)
                        except Exception: pass
                        calc += Decimal(str(line_excl_vat))
                        try:
                            vr = float(it.get('vat_rate') or 0)
                            if 0 < vr < 1: vr *= 100  # accept 0.05 == 5%
                            calc += Decimal(str(line_excl_vat * vr / 100.0))
                        except Exception: pass
                    except Exception: continue
            # Header-level VAT + discount adjustments
            try:
                hv = float(invoice_dict.get('vat_amount') or invoice_dict.get('tax_amount') or 0)
                calc += Decimal(str(max(0.0, hv)))
            except Exception: pass
            try:
                hd = float(invoice_dict.get('discount_amount') or invoice_dict.get('discount') or 0)
                calc = calc - Decimal(str(max(0.0, hd)))
                if calc < 0: calc = Decimal("0")
            except Exception: pass
            try:
                r = float(invoice_dict.get('rounding') or 0)
                calc += Decimal(str(r))
            except Exception: pass
            total = round_aed(calc)

        # COGS priority: (1) header total_cost, (2) header cogs, (3) sum(items * cost_price)
        cogs = round_aed(invoice_dict.get('total_cost') or invoice_dict.get('cogs') or
                         invoice_dict.get('cost_of_goods_sold') or 0)
        if cogs <= 0:
            items = invoice_dict.get('items') or invoice_dict.get('line_items') or []
            if isinstance(items, list):
                for it in items:
                    if isinstance(it, dict):
                        q = float(it.get('qty') or it.get('quantity') or 0)
                        c = float(it.get('cost_price') or it.get('purchase_price')
                                  or it.get('cost') or it.get('unit_cost') or 0)
                        cogs = cogs + round_aed(q * c)

        batch = []
        if total > 0:
            # A) Debit AR (1200) | Credit Sales Revenue (4000)
            batch.append({'date': inv_date, 'account_code': '1200', 'debit_amount': total, 'credit_amount': 0,
                          'reference_type': 'INVOICE', 'reference_id': inv_id,
                          'description': f"Invoice {inv_id} – {client} — Amount Owed by Customer"})
            batch.append({'date': inv_date, 'account_code': '4000', 'debit_amount': 0, 'credit_amount': total,
                          'reference_type': 'INVOICE', 'reference_id': inv_id,
                          'description': f"Invoice {inv_id} – {client} — Recognized Sales Revenue"})
        if cogs > 0:
            # B) Debit COGS (5000) | Credit Inventory (1300)
            batch.append({'date': inv_date, 'account_code': '5000', 'debit_amount': cogs, 'credit_amount': 0,
                          'reference_type': 'INVOICE', 'reference_id': inv_id,
                          'description': f"Invoice {inv_id} – {client} — Cost of Goods Sold"})
            batch.append({'date': inv_date, 'account_code': '1300', 'debit_amount': 0, 'credit_amount': cogs,
                          'reference_type': 'INVOICE', 'reference_id': inv_id,
                          'description': f"Invoice {inv_id} – {client} — Inventory Decrease (Sold)"})
        return self._write_gl_batch(batch, username=username)

    # --------------------------------------------------------------------------
    # HOOK 2: on_payment_received (Bank / AR)
    # --------------------------------------------------------------------------
    def on_payment_received(self, invoice_id, amount, payment_date=None, username=None, note=""):
        amount = round_aed(amount)
        if amount <= 0:
            return []
        pd = payment_date or datetime.now().strftime('%Y-%m-%d')
        # Unique payment reference id
        pid = f"PAY-{invoice_id}-{pd.replace('-','')}-{abs(hash(str(amount)+note)) % 10000:04d}"
        if self._entry_exists('PAYMENT', pid):
            return []
        batch = [
            {'date': pd, 'account_code': '1100', 'debit_amount': amount, 'credit_amount': 0,
             'reference_type': 'PAYMENT', 'reference_id': pid,
             'description': f"Payment received for Invoice {invoice_id} — {note}".strip()},
            {'date': pd, 'account_code': '1200', 'debit_amount': 0, 'credit_amount': amount,
             'reference_type': 'PAYMENT', 'reference_id': pid,
             'description': f"Apply payment to Invoice {invoice_id} — Reduce customer balance"}
        ]
        return self._write_gl_batch(batch, username=username)

    # --------------------------------------------------------------------------
    # HOOK 2b: on_purchase_payment_made (AP / Bank settlement for supplier pay)
    # --------------------------------------------------------------------------
    def on_purchase_payment_made(self, purchase_id, amount, payment_date=None,
                                 supplier="", username=None, note=""):
        amount = round_aed(amount)
        if amount <= 0:
            return []
        pd = payment_date or datetime.now().strftime('%Y-%m-%d')
        pid = f"PO-PAY-{purchase_id}-{pd.replace('-','')}-{abs(hash(str(amount)+note+supplier)) % 10000:04d}"
        if self._entry_exists('PURCHASE_PAYMENT', pid):
            return []
        batch = [
            {'date': pd, 'account_code': '2000', 'debit_amount': amount, 'credit_amount': 0,
             'reference_type': 'PURCHASE_PAYMENT', 'reference_id': pid,
             'description': f"Supplier payment for PO {purchase_id} – {supplier} — Reduce AP".strip()},
            {'date': pd, 'account_code': '1100', 'debit_amount': 0, 'credit_amount': amount,
             'reference_type': 'PURCHASE_PAYMENT', 'reference_id': pid,
             'description': f"Supplier payment for PO {purchase_id} – {supplier} — Bank Outflow".strip()}
        ]
        return self._write_gl_batch(batch, username=username)

    # --------------------------------------------------------------------------
    # HOOK 3: on_purchase_received (Inventory / AP)
    # --------------------------------------------------------------------------
    def on_purchase_received(self, purchase_dict, username=None):
        po_id = str(purchase_dict.get('purchase_id') or purchase_dict.get('po_no')
                    or purchase_dict.get('purchase_no') or purchase_dict.get('id') or '')
        if not po_id or self._entry_exists('PURCHASE', po_id):
            return []
        po_date = (purchase_dict.get('date') or purchase_dict.get('purchase_date')
                   or purchase_dict.get('created_at') or datetime.now().strftime('%Y-%m-%d'))
        supplier = (purchase_dict.get('supplier') or purchase_dict.get('supplier_name')
                    or purchase_dict.get('vendor') or 'Unknown Supplier')

        # Same field-robustness pattern as on_invoice_created.
        def _grab(*keys):
            for k in keys:
                try:
                    v = purchase_dict.get(k)
                    if v is None or v == "": continue
                    n = float(str(v).replace(",",""))
                    if n >= 0: return round_aed(n)
                except Exception: continue
            return None
        amt = (_grab('grand_total','total_amount','total','amount','net_total',
                      'invoice_total','po_total','purchase_total')
               or round_aed(0))
        # Fallback: recompute from items if no header total present
        if amt <= 0:
            items = purchase_dict.get('items') or purchase_dict.get('line_items') or []
            calc = Decimal("0")
            if isinstance(items, list):
                for it in items:
                    if not isinstance(it, dict): continue
                    try:
                        q = float(it.get('qty') or it.get('quantity') or 0)
                        p = float(it.get('unit_price') or it.get('price') or it.get('rate')
                                  or it.get('cost_price') or it.get('purchase_price') or 0)
                        calc += Decimal(str(q * p))
                    except Exception: continue
            amt = round_aed(calc)
        if amt <= 0:
            return []
        batch = [
            {'date': po_date, 'account_code': '1300', 'debit_amount': amt, 'credit_amount': 0,
             'reference_type': 'PURCHASE', 'reference_id': po_id,
             'description': f"Purchase {po_id} – {supplier} — Inventory Increased (Goods Received)"},
            {'date': po_date, 'account_code': '2000', 'debit_amount': 0, 'credit_amount': amt,
             'reference_type': 'PURCHASE', 'reference_id': po_id,
             'description': f"Purchase {po_id} – {supplier} — Owed to Supplier (AP)"}
        ]
        self._write_gl_batch(batch, username=username)
        # If purchase was paid, also post AP settlement: debit 2000 (reduce AP) | credit 1100 (bank)
        paid = (_grab('amount_paid','paid_amount','payment_amount','paid')
                or round_aed(0))
        if paid > 0:
            try:
                self.on_purchase_payment_made(po_id, paid, po_date, supplier=supplier, username=username)
            except Exception:
                pass
        return batch

    # --------------------------------------------------------------------------
    # HOOK 4: on_expense_submitted (User-friendly Category)
    # --------------------------------------------------------------------------
    def on_expense_submitted(self, category, amount, expense_date, payee_name="", description="", username=None):
        amount = round_aed(amount)
        if amount <= 0:
            raise ValueError("Expense amount must be positive.")
        if category not in EXPENSE_CATEGORY_MAP:
            raise ValueError(f"Invalid category. Choose: {list(EXPENSE_CATEGORY_MAP.keys())}")
        exp_dt = expense_date or datetime.now().strftime('%Y-%m-%d')
        acct = EXPENSE_CATEGORY_MAP[category]
        ref_id = f"EXP-{exp_dt.replace('-','')}-{abs(hash(category+payee_name+str(amount))) % 1000000:06d}"
        if self._entry_exists('EXPENSE', ref_id):
            return []
        batch = [
            {'date': exp_dt, 'account_code': acct, 'debit_amount': amount, 'credit_amount': 0,
             'reference_type': 'EXPENSE', 'reference_id': ref_id,
             'description': f"{category} — Payee: {payee_name or 'N/A'} | {description}".strip()},
            {'date': exp_dt, 'account_code': '1100', 'debit_amount': 0, 'credit_amount': amount,
             'reference_type': 'EXPENSE', 'reference_id': ref_id,
             'description': f"Paid via Bank — {category} to {payee_name or 'N/A'}"}
        ]
        return self._write_gl_batch(batch, username=username)

    # --------------------------------------------------------------------------
    # HOOK 5: on_invoice_costs_changed (Add / Edit / Remove costs on an invoice)
    # --------------------------------------------------------------------------
    def on_invoice_costs_changed(self, invoice_id, costs_list, invoice_date=None,
                                 invoice_client="", username=None,
                                 operation="update", previous_costs_list=None):
        """
        Post GL double-entries for invoice-level costs.

        Each cost line in costs_list is expected to have:
            description, amount, account (fund/bank/AP name), expense_account (5xxx).
        """
        inv_id = str(invoice_id) if invoice_id else ""
        if not inv_id:
            return []
        inv_dt = invoice_date or datetime.now().strftime('%Y-%m-%d')
        costs_list = costs_list or []
        previous_costs_list = previous_costs_list or []

        def _fund_acct_code(fund_name_raw):
            # Map balance-manager account name -> GL account code.  Bank names
            # go to 1100; anything containing "payable" or "credit" goes to
            # 2000/2100; everything else defaults to 1100 (you paid cash).
            if not fund_name_raw:
                return "1100", "Bank"
            f = str(fund_name_raw).lower().strip()
            if any(k in f for k in ("payable", "ap ", "supplier", "credit")):
                if any(k in f for k in ("accrued", "accrual")):
                    return "2100", "Accrued Expenses"
                return "2000", "Accounts Payable"
            if "cash" in f:
                return "1000", "Cash"
            return "1100", "Bank"

        def _exp_acct(code_raw):
            if not code_raw:
                return "5000"
            s = str(code_raw).strip()
            head = s.split()[0] if " " in s else s
            import re
            if re.fullmatch(r"\d{4}", head):
                return head
            return "5000"

        # We REVERSE previous costs (if any) for this invoice id under the
        # same reference, then post the new list.  Idempotent via unique
        # per-cost signature hash.
        batch = []
        # --- Reverse previous costs, if any ---
        if previous_costs_list and operation != "create":
            for idx, c in enumerate(previous_costs_list):
                amt = round_aed(c.get("amount") or 0)
                if amt <= 0:
                    continue
                sig = f"{c.get('description','')}|{c.get('account','')}|{c.get('expense_account','')}|{idx}"
                ref_id = f"INV-COST-{inv_id}-REV-{abs(hash(sig)) % 10000:04d}"
                if self._entry_exists('INVOICE_COST', ref_id):
                    continue
                exp_code = _exp_acct(c.get("expense_account"))
                fund_code, fund_name = _fund_acct_code(c.get("account"))
                desc = f"Invoice {inv_id} cost REVERSE: {c.get('description','')}"
                batch.append({
                    'date': inv_dt, 'account_code': exp_code,
                    'debit_amount': 0, 'credit_amount': amt,
                    'reference_type': 'INVOICE_COST', 'reference_id': ref_id,
                    'description': desc + f" (reverse Dr {exp_code})",
                })
                batch.append({
                    'date': inv_dt, 'account_code': fund_code,
                    'debit_amount': amt, 'credit_amount': 0,
                    'reference_type': 'INVOICE_COST', 'reference_id': ref_id,
                    'description': desc + f" (reverse Cr {fund_code} {fund_name})",
                })

        # --- Post new costs ---
        for idx, c in enumerate(costs_list):
            amt = round_aed(c.get("amount") or 0)
            if amt <= 0:
                continue
            sig = f"{c.get('description','')}|{c.get('account','')}|{c.get('expense_account','')}|{idx}|{amt}|{operation}"
            ref_id = f"INV-COST-{inv_id}-{abs(hash(sig)) % 10000:04d}"
            if self._entry_exists('INVOICE_COST', ref_id):
                continue
            exp_code = _exp_acct(c.get("expense_account"))
            fund_code, fund_name = _fund_acct_code(c.get("account"))
            desc = (f"Invoice {inv_id}{' – ' + invoice_client if invoice_client else ''} — "
                    f"Cost: {c.get('description','')} (paid via {fund_name})")
            batch.append({
                'date': inv_dt, 'account_code': exp_code,
                'debit_amount': amt, 'credit_amount': 0,
                'reference_type': 'INVOICE_COST', 'reference_id': ref_id,
                'description': f"Dr {exp_code}: " + desc,
            })
            batch.append({
                'date': inv_dt, 'account_code': fund_code,
                'debit_amount': 0, 'credit_amount': amt,
                'reference_type': 'INVOICE_COST', 'reference_id': ref_id,
                'description': f"Cr {fund_code} ({fund_name}): " + desc,
            })
        return self._write_gl_batch(batch, username=username)

    # --------------------------------------------------------------------------
    # HOOK 6: on_invoice_costs_backfilled (bulk apply after Cost Center preview)
    # --------------------------------------------------------------------------
    def on_invoice_costs_backfilled(self, results_iterator, username=None):
        """
        results_iterator: iterable of dicts, each with keys:
          invoice_id, costs_list, invoice_date, invoice_client
        Returns {posted_batches, posted_entries, skipped, errors}
        """
        posted_batches = 0
        posted_entries = 0
        errors = []
        skipped = 0
        for r in results_iterator or []:
            if not isinstance(r, dict):
                skipped += 1
                continue
            try:
                entries = self.on_invoice_costs_changed(
                    r.get("invoice_id"),
                    r.get("costs_list") or r.get("costs") or [],
                    invoice_date=r.get("invoice_date") or r.get("date"),
                    invoice_client=r.get("invoice_client") or r.get("client") or "",
                    username=username,
                    operation=r.get("operation", "backfill"),
                    previous_costs_list=r.get("previous_costs_list"),
                )
                if entries:
                    posted_entries += len(entries)
                    posted_batches += 1
            except Exception as exc:
                errors.append({
                    "invoice_id": r.get("invoice_id"),
                    "error": str(exc),
                })
        return {
            "posted_batches": posted_batches,
            "posted_entries": posted_entries,
            "errors": errors,
            "skipped": skipped,
        }

    # --------------------------------------------------------------------------
    # HOOK 7: on_salary_payroll_posted (Payroll / Salaries GL postings)
    # --------------------------------------------------------------------------
    def on_salary_payroll_posted(self, month_key, employee_id, employee_name,
                                  gross_salary, total_deductions, net_salary,
                                  processed_date=None, payment_method="Bank",
                                  username=None):
        """
        Post GAAP payroll for ONE employee for ONE month.

        Double-entry:
          Dr 5100 Salary Expense  (Gross)
              Cr 2100 Accrued Expenses / Salary Payable  (Gross)
        THEN when PAID:
          Dr 2100 Accrued (Net)       Cr 1100 Bank / 1000 Cash  (Net)
        Any deductions (loan, penalties) reduce Cr 2100 and Dr 5100 or other receivable.
        For simplicity we do Net pay outflow immediately against Bank/Cash.

        Idempotent per month+employee combo.
        """
        gross = round_aed(gross_salary or 0)
        net = round_aed(net_salary or 0)
        ded = round_aed(total_deductions or 0)
        if gross <= 0:
            return []
        if abs(gross - (net + ded)) > Decimal("0.02"):
            ded = gross - net if net <= gross else Decimal("0")
        pd = processed_date or datetime.now().strftime('%Y-%m-%d')
        if isinstance(pd, date):
            pd = pd.isoformat()
        if isinstance(pd, datetime):
            pd = pd.date().isoformat()
        mk = (month_key or "").strip()
        emp_id = str(employee_id or "").strip() or "UNKNOWN"
        emp_name = str(employee_name or "").strip() or "Unnamed Employee"
        sig = f"SALARY-{mk}-{emp_id}-{gross}-{net}-{pd}"
        ref_id = f"SAL-{mk.replace('-','')}-{emp_id}-{abs(hash(sig)) % 10000:04d}"
        if self._entry_exists('SALARY', ref_id):
            return []
        fund_code = "1000" if payment_method and "cash" in str(payment_method).lower() else "1100"
        fund_desc = "Cash" if fund_code == "1000" else "Bank"
        batch = [
            {'date': pd, 'account_code': '5100', 'debit_amount': gross, 'credit_amount': 0,
             'reference_type': 'SALARY', 'reference_id': ref_id,
             'description': f"Salary {mk}: {emp_name} ({emp_id}) Gross Pay — Dr Salary Expense"},
            {'date': pd, 'account_code': '2100', 'debit_amount': 0, 'credit_amount': gross,
             'reference_type': 'SALARY', 'reference_id': ref_id,
             'description': f"Salary {mk}: {emp_name} ({emp_id}) Gross — Cr Salary Payable (Accrued)"},
        ]
        if net > 0:
            batch.extend([
                {'date': pd, 'account_code': '2100', 'debit_amount': net, 'credit_amount': 0,
                 'reference_type': 'SALARY', 'reference_id': ref_id,
                 'description': f"Salary {mk}: {emp_name} ({emp_id}) Net Pay Settlement — Dr Payable"},
                {'date': pd, 'account_code': fund_code, 'debit_amount': 0, 'credit_amount': net,
                 'reference_type': 'SALARY', 'reference_id': ref_id,
                 'description': f"Salary {mk}: {emp_name} ({emp_id}) Net Pay — Cr {fund_desc} Outflow"},
            ])
        # Note: deductions are NOT posted as a separate clearing line.
        # The 2100 Accrued balance (Gross Cr - Net Dr) already equals the
        # deductions amount, representing a remaining payable to 3rd parties
        # (tax, social insurance, loan recovery) — to be paid in a separate
        # treasury disbursement entry (Dr 2100 deductions / Cr Bank).
        # This keeps the GL batch perfectly balanced Dr=Cr.
        return self._write_gl_batch(batch, username=username)

    # --------------------------------------------------------------------------
    # COA REPAIR HELPERS: always guarantee a usable Chart of Accounts exists
    # --------------------------------------------------------------------------
    _SEED_COA_CACHE = [
        {"account_code": "1000", "account_name": "Cash", "full_code_name": "1000_Cash", "account_type": "ASSET", "is_seeded": True},
        {"account_code": "1100", "account_name": "Bank", "full_code_name": "1100_Bank", "account_type": "ASSET", "is_seeded": True},
        {"account_code": "1200", "account_name": "Accounts_Receivable", "full_code_name": "1200_Accounts_Receivable", "account_type": "ASSET", "is_seeded": True},
        {"account_code": "1300", "account_name": "Inventory", "full_code_name": "1300_Inventory", "account_type": "ASSET", "is_seeded": True},
        {"account_code": "2000", "account_name": "Accounts_Payable", "full_code_name": "2000_Accounts_Payable", "account_type": "LIABILITY", "is_seeded": True},
        {"account_code": "2100", "account_name": "Accrued_Expenses", "full_code_name": "2100_Accrued_Expenses", "account_type": "LIABILITY", "is_seeded": True},
        {"account_code": "3000", "account_name": "Retained_Earnings", "full_code_name": "3000_Retained_Earnings", "account_type": "EQUITY", "is_seeded": True},
        {"account_code": "3100", "account_name": "Owner_Equity", "full_code_name": "3100_Owner_Equity", "account_type": "EQUITY", "is_seeded": True},
        {"account_code": "3200", "account_name": "Owner_Drawings", "full_code_name": "3200_Owner_Drawings", "account_type": "EQUITY", "is_seeded": True},
        {"account_code": "4000", "account_name": "Sales_Revenue", "full_code_name": "4000_Sales_Revenue", "account_type": "REVENUE", "is_seeded": True},
        {"account_code": "5000", "account_name": "COGS", "full_code_name": "5000_COGS", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5100", "account_name": "Salary_Expense", "full_code_name": "5100_Salary_Expense", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5200", "account_name": "Rent_Utility", "full_code_name": "5200_Rent_Utility", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5300", "account_name": "General_Admin", "full_code_name": "5300_General_Admin", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5400", "account_name": "Prof_Fees", "full_code_name": "5400_Prof_Fees", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5500", "account_name": "Delivery_Shipping", "full_code_name": "5500_Delivery_Shipping", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5600", "account_name": "Medical_Service_Fees", "full_code_name": "5600_Medical_Service_Fees", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5700", "account_name": "Sales_Commissions", "full_code_name": "5700_Sales_Commissions", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5800", "account_name": "Warranty_Returns", "full_code_name": "5800_Warranty_Returns", "account_type": "EXPENSE", "is_seeded": True},
        {"account_code": "5900", "account_name": "Other_Operating_Expenses", "full_code_name": "5900_Other_Operating_Expenses", "account_type": "EXPENSE", "is_seeded": True},
    ]

    def ensure_chart_of_accounts(self):
        """Guarantee the CoA file has all 14 seeded accounts (adds any missing ones; no duplicates)."""
        coa = self.dm.load_json('chart_of_accounts.json') or []
        if not isinstance(coa, list):
            coa = []
        existing_codes = set()
        cleaned = []
        for a in coa:
            if not isinstance(a, dict):
                continue
            code = str(a.get('account_code') or a.get('code') or '').strip()
            if not code:
                continue
            a['account_code'] = code
            if 'account_type' not in a:
                a['account_type'] = 'ASSET'
            if 'account_name' not in a:
                a['account_name'] = code
            if 'full_code_name' not in a:
                a['full_code_name'] = f"{code}_{a.get('account_name','')}"
            if code in existing_codes:
                continue
            existing_codes.add(code)
            cleaned.append(a)
        changed = False
        for seed in self._SEED_COA_CACHE:
            if seed['account_code'] not in existing_codes:
                cleaned.append(dict(seed))
                existing_codes.add(seed['account_code'])
                changed = True
        if changed or len(cleaned) != len(coa):
            try:
                self.dm.save_json('chart_of_accounts.json', cleaned)
            except Exception:
                pass
        return cleaned

    def load_chart_of_accounts(self):
        """Idempotent CoA loader — GUARANTEES usable accounts. Never crashes on bad JSON."""
        try:
            coa = self.ensure_chart_of_accounts() or []
        except Exception:
            coa = [dict(a) for a in self._SEED_COA_CACHE]
        return coa

    # --------------------------------------------------------------------------
    # EXPENSE CATEGORY HELPERS: normalize legacy strings -> friendly categories
    # --------------------------------------------------------------------------
    _CATEGORY_NORMALIZE = {
        "salary": "Salary/Wages", "salaries": "Salary/Wages", "wages": "Salary/Wages", "payroll": "Salary/Wages",
        "salary/wages": "Salary/Wages", "salary & wages": "Salary/Wages",
        "rent": "Rent", "lease": "Rent", "rental": "Rent",
        "utility": "Utilities (Electricity/Water/Internet)", "utilities": "Utilities (Electricity/Water/Internet)",
        "electricity": "Utilities (Electricity/Water/Internet)", "water": "Utilities (Electricity/Water/Internet)",
        "internet": "Utilities (Electricity/Water/Internet)", "phone": "Utilities (Electricity/Water/Internet)",
        "office": "Office Supplies", "stationery": "Office Supplies", "office supplies": "Office Supplies",
        "travel": "Travel & Transportation", "transport": "Travel & Transportation",
        "transportation": "Travel & Transportation", "fuel": "Travel & Transportation",
        "marketing": "Marketing & Advertising", "advertising": "Marketing & Advertising", "ads": "Marketing & Advertising",
        "admin": "Other Administrative", "general": "Other Administrative",
        "other": "Other Administrative", "misc": "Other Administrative", "miscellaneous": "Other Administrative",
    }

    def normalize_expense_category(self, raw_category):
        """Fuzzy-match any raw category string to the 7 valid EXPENSE_CATEGORY_MAP keys."""
        if not raw_category:
            return "Other Administrative"
        if raw_category in EXPENSE_CATEGORY_MAP:
            return raw_category
        key = str(raw_category).strip().lower().replace("_", " ").replace("-", " ")
        if key in self._CATEGORY_NORMALIZE:
            return self._CATEGORY_NORMALIZE[key]
        for frag, cat in self._CATEGORY_NORMALIZE.items():
            if frag in key or key in frag:
                return cat
        return cat

    # --------------------------------------------------------------------------
    # Smart expense classification using description / payee keywords.
    # Used when transactions.json stores withdrawals without an explicit category.
    # --------------------------------------------------------------------------
    _EXPENSE_KEYWORD_MAP = [
        ("Salary/Wages",                               ["salary", "wage", "payroll", "employee pay", "staff pay", "salaries", "basic salary", "payroll tax", "pension"]),
        ("Rent",                                        ["rent", "lease", "rental", "tenancy", "landlord"]),
        ("Utilities (Electricity/Water/Internet)",      ["utility", "electricity", "water bill", "dewa", "sewa", "internet", "mobile bill", "phone bill", "etisalat", "du telecom"]),
        ("Office Supplies",                             ["stationery", "office supply", "printer", "paper", "cartridge", "toner", "ink", "envelope"]),
        ("Marketing & Advertising",                     ["advert", "ads", "marketing", "google ads", "meta ads", "facebook ads", "promotion", "booth", "exhibition"]),
        ("Travel & Transportation",                     ["travel", "airline", "emirates", "flydubai", "hotel", "airbnb", "ticket", "petrol", "gas", "fuel", "transport", "delivery", "courier", "taxi", "uber", "careem"]),
        ("Other Administrative",                        ["admin", "consulting", "professional fee", "accountant", "audit", "legal fee", "subscription", "software", "xerox", "dhl", "miscel", "general expense", "bank charge", "bank fee", "other"]),
    ]

    def _expense_category_from_text(self, text):
        """Match payee/description keywords → 1 of 7 valid EXPENSE_CATEGORY_MAP keys."""
        t = (str(text) or "").lower()
        for cat, kws in self._EXPENSE_KEYWORD_MAP:
            for kw in kws:
                if kw in t:
                    return cat
        # Last-ditch: run through normalize using first 2 words
        words = t.replace("/", " ").replace("-", " ").split()
        for w in words[:4]:
            nc = self.normalize_expense_category(w)
            if nc != "Other Administrative":
                return nc
        return "Other Administrative"

    # --------------------------------------------------------------------------
    # HOOK 8: Expense Manager — list / reverse / edit / delete
    # --------------------------------------------------------------------------
    def list_expenses(self, start_date=None, end_date=None, search=None,
                      include_withdrawals=True, include_invoice_costs=True,
                      exclude_reversed=True):
        """Return a user-friendly list of ALL expenses recorded in the system.

        Aggregates GL entries by reference_id:
          • EXPENSE         → Expenses you added via 💸 Expense dialog / main Add Expense.
          • WITHDRAWAL      → Bank-imported withdrawals classified as admin expenses.
          • INVOICE_COST    → Invoice-level extra costs (delivery fee, customs, etc.)
          • SALARY          → Optional: salary expenses (if include_payroll=True)
          • PURCHASE        → Inventory purchases (include_purchases)

        **Important (new simplified UX):**
        By default this method EXCLUDES any (ref_type, ref_id) that has been
        reversed via delete_expense / reverse_reference.  This means once you
        right-click → Delete, that expense line vanishes from: Expense Ledger
        list, Expense Ledger Total, any UI that calls list_expenses, and will
        not re-appear unless the caller explicitly passes exclude_reversed=False
        (auditor mode to see reversed + original for forensics).

        Returns: list of expense_row dicts:
          {date, ref_id, ref_type, category, account_code, account_name,
           payee, description, gross_amount, bank_account, created_by,
           created_at, updated_at, reversed_ref, status}
        """
        from decimal import Decimal as _D
        gl = self.dm.load_json('general_ledger.json') or []
        # Group all GL lines under same (ref_type, ref_id)
        groups = {}
        for e in gl:
            if not isinstance(e, dict): continue
            rt = str(e.get('reference_type') or '')
            rid = str(e.get('reference_id') or '')
            if not rt or not rid:
                continue
            is_expense_type = (rt in ('EXPENSE',))
            if include_withdrawals and rt == 'WITHDRAWAL': is_expense_type = True
            if include_invoice_costs and rt == 'INVOICE_COST': is_expense_type = True
            # Skip reversing rows' parents if reversed by REVERSAL OF
            if not is_expense_type:
                continue
            key = (rt, rid)
            if key not in groups:
                groups[key] = {'lines': [], 'rev_by': None}
            groups[key]['lines'].append(e)
        # Find "REVERSAL OF" markers: any entries whose description starts with
        # "🔁 REVERSAL OF [type]:id" — then tag the source group so it can show
        # "Reversed" status.
        for e in gl:
            desc = str(e.get('description') or '')
            if desc.startswith("🔁 REVERSAL OF"):
                import re as _re
                m = _re.search(r"REVERSAL OF\s+(\w+):(\S+)", desc)
                if m:
                    rt2, rid2 = m.group(1), m.group(2)
                    k = (rt2, rid2)
                    if k in groups:
                        groups[k]['rev_by'] = str(e.get('created_at') or desc)
        # Build user rows
        rows = []
        for (rt, rid), bundle in groups.items():
            # ---- NEW DEFAULT BEHAVIOUR: anything reversed vanishes. ----
            # If exclude_reversed=True (default) we REMOVE the reversed
            # bundle entirely: it won't appear in list counts, won't appear
            # in totals, won't appear in searches, as if it never existed.
            # Only show reversed when exclude_reversed=False (audit mode).
            if exclude_reversed and bundle.get('rev_by'):
                continue
            lines = bundle['lines'] or []
            if not lines: continue
            # Filter date
            dts = sorted([str(x.get('date') or '') for x in lines if x.get('date')])
            expense_dt = dts[-1] if dts else ''
            if expense_dt and start_date and expense_dt < str(start_date):
                continue
            if expense_dt and end_date and expense_dt > str(end_date):
                continue
            # Identify Expense Debit side (Dr = 5xxx account) and Credit side (Bank/AP)
            dr_lines = [x for x in lines if Decimal(x.get('debit_amount') or 0) > 0]
            cr_lines = [x for x in lines if Decimal(x.get('credit_amount') or 0) > 0]
            expense_line = max(dr_lines, key=lambda x: Decimal(x.get('debit_amount') or 0)) if dr_lines else None
            payment_line = cr_lines[0] if cr_lines else None
            if expense_line is None:
                continue
            acct = str(expense_line.get('account_code') or '')
            gross = Decimal(expense_line.get('debit_amount') or 0)
            if gross <= 0:
                continue
            # Determine category label from description first (EXP dialog writes
            # "{category} — Payee: X | Y" format), then fall back to account reverse-map,
            # then keyword-classification from payee/description.
            desc = str(expense_line.get('description') or '')
            # ---- Parse payee + description FIRST (needed for tie-breaks below) ----
            payee = ''
            import re as _re2
            m_payee = _re2.search(r"Payee:\s*([^|]+)", desc)
            if m_payee:
                payee = m_payee.group(1).strip()
            clean_desc = desc
            if '|' in desc:
                clean_desc = desc.split('|', 1)[1].strip()
            elif 'Payee:' in desc:
                clean_desc = ''
            # ---- Now resolve category label ----
            category_label = None
            import re as _re_label
            lab = _re_label.match(r"^([^—|]+?)\s*—", desc)
            if lab:
                lab_candidate = lab.group(1).strip()
                if lab_candidate in EXPENSE_CATEGORY_MAP:
                    category_label = lab_candidate
            if category_label is None:
                # Fallback 2: reverse EXPENSE_CATEGORY_MAP lookup (note: 5200 and 5300
                # are shared buckets — we'll pick whichever label's keywords best match)
                rev_hits = [k for k, v in EXPENSE_CATEGORY_MAP.items() if v == acct]
                if len(rev_hits) == 1:
                    category_label = rev_hits[0]
                elif rev_hits:
                    # Tie-break by keyword match on payee+description
                    category_label = self._expense_category_from_text(desc + " " + payee)
            if category_label is None:
                # Fallback 3: keyword classification
                if acct.startswith('5'):
                    category_label = self._expense_category_from_text(
                        desc or expense_line.get('account_name') or '')
                else:
                    category_label = str(expense_line.get('account_name') or acct)
            bank_name = ''
            if payment_line is not None:
                bank_code = str(payment_line.get('account_code') or '')
                bank_name = str(payment_line.get('account_name') or bank_code)
                if '1100' in bank_code:
                    bank_name = 'Bank ' + (payment_line.get('description') or '').split('—')[-1].strip() if '—' in (payment_line.get('description') or '') else bank_name
            created_by = str(expense_line.get('created_by') or 'System')
            created_at = str(expense_line.get('created_at') or '')
            status = 'Posted'
            if bundle['rev_by']:
                status = '🗑️  Reversed'
            # --- Source/origin classifier (mirrors
            # GAAPReportingDashboard._classify_expense_origin so both paths
            # produce exactly the same labels) -------------------------------
            rt_upper = (rt or "").upper()
            cb = (created_by or "").strip()
            cb_low = cb.lower()
            if rt_upper in {"WITHDRAWAL", "INVOICE_COST", "SALARY", "PAYROLL",
                            "PURCHASE", "STOCK_ADJUSTMENT", "REVERSAL",
                            "ADJUSTMENT", "PO_ACCRUAL", "CREDIT_NOTE", "INVOICE"}:
                origin = "System"
            elif cb and cb_low in {"system", "auto", "employee manager",
                                   "employeemanager", "invoiceapp", "invoice app",
                                   "ledgerservice", "ledger service",
                                   "enhancedclouddatamanager", "gaap dashboard",
                                   "reversal", "automatic", "batch", "scheduled"}:
                origin = "System"
            elif desc.startswith("🔁 REVERSAL OF") or "REVERSAL OF" in desc[:60]:
                origin = "System"
            elif str(rid).startswith("EXP-") and len(str(rid)) >= 14:
                origin = "Manual"
            elif "Payee:" in desc and rt_upper == "EXPENSE":
                origin = "Manual"
            elif cb and len(cb) >= 2 and not any(x in cb_low for x in
                                  ["manager", "service", "module", "system", "auto"]):
                origin = "Manual"
            else:
                origin = "Manual" if rt_upper == "EXPENSE" else "System"
            # Search filter
            if search:
                s = str(search).lower().strip()
                if s:
                    haystack = " ".join(str(v or "") for v in [
                        expense_dt, rid, rt, category_label, acct,
                        payee, clean_desc, created_by, bank_name, origin
                    ]).lower()
                    if s not in haystack:
                        continue
            rows.append({
                'date': expense_dt,
                'ref_id': rid,
                'ref_type': rt,
                'category': category_label,
                'account_code': acct,
                'account_name': str(expense_line.get('account_name') or acct),
                'payee': payee,
                'description': clean_desc,
                'gross_amount': gross,
                'bank_account': bank_name,
                'created_by': created_by,
                'created_at': created_at,
                'updated_at': bundle['rev_by'] or created_at,
                'reversed_ref': bundle['rev_by'],
                'status': status,
                'origin': origin,
            })
        # Sort descending by date, then created_at
        rows.sort(key=lambda r: (str(r['date']), str(r['created_at'])), reverse=True)
        return rows

    def reverse_reference(self, reference_type, reference_id, username=None, reason="Manual correction"):
        """Post a REVERSAL batch for an existing (reference_type, reference_id) pair.

        GL is immutable, so we flip Dr↔Cr of every original line, keep the same
        date (or today if locked), add "🔁 REVERSAL OF type:id" prefix to description,
        then write a new balanced batch.  Returns list of new GL entries written.
        Raises if reference not found, or posting fails.
        """
        gl = self.dm.load_json('general_ledger.json') or []
        existing = [e for e in gl
                    if isinstance(e, dict)
                    and str(e.get('reference_type', '')) == str(reference_type)
                    and str(e.get('reference_id', '')) == str(reference_id)]
        if not existing:
            raise ValueError(
                f"No GL entries found for ({reference_type}, {reference_id}). "
                f"Nothing to reverse.")
        # Detect if already reversed (don't double-reverse silently)
        rev_prefix = f"🔁 REVERSAL OF {reference_type}:{reference_id}"
        already = any(rev_prefix in str(e.get('description', '')) for e in gl)
        if already:
            raise ValueError(
                f"Reference {reference_type}:{reference_id} was already reversed. "
                f"Cannot reverse again.")
        reversals = []
        for src in existing:
            dr = Decimal(src.get('debit_amount') or 0)
            cr = Decimal(src.get('credit_amount') or 0)
            # Flip sides for reverse entry
            new_dr, new_cr = round_aed(cr), round_aed(dr)
            if (new_dr > 0) == (new_cr > 0):
                # Skip zero lines (shouldn't happen anyway)
                if new_dr == 0 and new_cr == 0: continue
                raise ValueError(f"Cannot reverse line {src.get('entry_id')}: invalid Dr/Cr state.")
            desc = f"{rev_prefix} — original #{src.get('entry_id')}: {src.get('description', '')}"[:240]
            if reason:
                desc += f" · Reason: {reason}"[:240]
            # Date: original date if not locked; otherwise today (don't silently post to locked period)
            dt = src.get('date') or date.today().isoformat()
            try:
                if self.is_period_locked(dt):
                    dt = date.today().isoformat()
            except Exception:
                dt = date.today().isoformat()
            reversals.append({
                'date': dt,
                'account_code': src.get('account_code'),
                'debit_amount': float(new_dr),
                'credit_amount': float(new_cr),
                'reference_type': 'REVERSAL',
                'reference_id': f"REV-{reference_id}-{abs(hash(reason or str(datetime.now()))) % 10000:04d}",
                'description': desc,
            })
        return self._write_gl_batch(reversals, username=username or 'System')

    def delete_expense(self, reference_type, reference_id, username=None, reason="Deleted by user (mistyped / duplicate)"):
        """Mark an expense as deleted (reverse it).  Returns dict with summary for display."""
        written = self.reverse_reference(reference_type, reference_id,
                                          username=username, reason=reason)
        gross = 0
        import functools
        for e in written:
            gross += float(Decimal(str(e.get('debit_amount') or 0)))
        return {'reversal_entries': len(written),
                'gross_reversed': round_aed(gross),
                'new_reference_id': written[0]['reference_id'] if written else ''}

    def edit_expense(self, reference_type, reference_id, new_fields, username=None):
        """Edit an expense the correct accounting way:

        1. Reverse the original expense batch (Dr↔Cr).
        2. Post a brand-new EXPENSE batch with the same reference_type using the new fields.

        ``new_fields`` may contain: category, amount, expense_date, payee_name, description.
        Returns (reversal_entries, new_entries).
        """
        if not isinstance(new_fields, dict) or not new_fields:
            raise ValueError("Nothing to update: pass at least one field (category / amount / date / payee / description).")
        # 1) Reverse the old batch
        reversals = self.reverse_reference(
            reference_type, reference_id,
            username=username,
            reason=f"Edited expense {reference_type}:{reference_id} → re-posted with corrected values."
        )
        # 2) Build new fields (merge sensible defaults from original reference)
        original_list = self.list_expenses()
        orig = next((x for x in original_list if x['ref_type'] == reference_type and x['ref_id'] == reference_id), None)
        if orig is None:
            new_cat = new_fields.get('category')
            new_amt = new_fields.get('amount')
            new_dt = new_fields.get('expense_date')
            new_pay = new_fields.get('payee_name', '')
            new_desc = new_fields.get('description', '')
            if not (new_cat and new_amt):
                raise ValueError(
                    "Cannot locate original expense. Pass category + amount explicitly in new_fields.")
        else:
            new_cat = new_fields.get('category') or orig['category']
            try:
                new_amt = round_aed(Decimal(str(new_fields.get('amount') if new_fields.get('amount') not in (None, '') else orig['gross_amount'])))
            except Exception:
                new_amt = round_aed(Decimal(orig['gross_amount']))
            new_dt = new_fields.get('expense_date') or orig['date']
            new_pay = new_fields.get('payee_name', orig['payee']) or ''
            new_desc = new_fields.get('description', orig['description']) or ''
        # 3) Re-post the expense via on_expense_submitted (correctly maps category → GL acct)
        #    Generate a fresh reference id so idempotency check passes (this is NEW batch):
        posted_new = self._force_repost_expense(
            category=new_cat, amount=float(new_amt), expense_date=new_dt,
            payee_name=new_pay, description=new_desc,
            username=username,
            origin_note=f"Corrected — replaces {reference_type}:{reference_id}"
        )
        return reversals, posted_new

    def _force_repost_expense(self, category, amount, expense_date,
                              payee_name="", description="", username=None,
                              origin_note=""):
        """Repost an expense WITHOUT the idempotency gate used by on_expense_submitted.

        We call this from edit_expense because after reverse + re-post, we need a
        truly NEW batch even if category|payee|amount|date collide with original.
        Uses a different (timestamped) reference_id pattern to ensure uniqueness."""
        amount = round_aed(amount)
        if amount <= 0:
            raise ValueError("Expense amount must be positive.")
        if category not in EXPENSE_CATEGORY_MAP:
            raise ValueError(f"Invalid category.  Valid: {list(EXPENSE_CATEGORY_MAP.keys())}")
        exp_dt = expense_date or datetime.now().strftime('%Y-%m-%d')
        acct = EXPENSE_CATEGORY_MAP[category]
        ref_id = (f"EXP-CORR-{exp_dt.replace('-','')}-"
                  f"{abs(hash(str(datetime.now().isoformat())+category+str(amount)+str(origin_note))) % 1000000:06d}")
        desc_full = f"{category} — Payee: {payee_name or 'N/A'} | {(description or '').strip()}"
        if origin_note:
            desc_full += f" · {origin_note}"
        batch = [
            {'date': exp_dt, 'account_code': acct, 'debit_amount': float(amount), 'credit_amount': 0,
             'reference_type': 'EXPENSE', 'reference_id': ref_id,
             'description': desc_full[:240]},
            {'date': exp_dt, 'account_code': '1100', 'debit_amount': 0, 'credit_amount': float(amount),
             'reference_type': 'EXPENSE', 'reference_id': ref_id,
             'description': f"Paid via Bank (Corrected) — {category} to {payee_name or 'N/A'}"[:240]}
        ]
        return self._write_gl_batch(batch, username=username)

    # --------------------------------------------------------------------------
    # HISTORICAL DATA RECONCILER: back-post ALL existing data into the GL
    # --------------------------------------------------------------------------
    def rebuild_general_ledger_from_all_data(self, progress_cb=None):
        """
        Reads EVERYTHING from existing JSON (invoices, purchases, expenses, payments)
        and idempotently posts them all to the General Ledger.

        Runs automatically if GL has < 3 entries. Safe to call multiple times.
        Returns a summary dict: {invoices, payments, purchases, expenses, total_gl, warnings}.
        """
        self.ensure_chart_of_accounts()
        stats = {"invoices": 0, "payments": 0, "purchases": 0, "expenses": 0,
                 "total_gl_after": 0, "warnings": []}

        def log(msg):
            if progress_cb:
                try:
                    progress_cb(msg)
                except Exception:
                    pass
            print(f"[GL Rebuild] {msg}")

        # Step 1: All historical invoices
        try:
            invoices = self.dm.load_json('invoices_data.json') or []
            invoices = [i for i in invoices if isinstance(i, dict)]
            log(f"Found {len(invoices)} invoices in invoices_data.json")
            for inv in invoices:
                try:
                    posted = self.on_invoice_created(inv, username="HistoricalReconcile")
                    if posted:
                        stats["invoices"] += 1
                except Exception as ex:
                    inv_id = inv.get('invoice_id') or inv.get('id') or '?'
                    stats["warnings"].append(f"Invoice {inv_id}: {ex}")
        except Exception as ex:
            stats["warnings"].append(f"Load invoices: {ex}")

        # Step 2: Customer payments (best-effort: split payment history from invoice.amount_paid)
        try:
            invoices = self.dm.load_json('invoices_data.json') or []
            invoices = [i for i in invoices if isinstance(i, dict)]
            log(f"Reconstructing payments for {len(invoices)} invoices")
            for inv in invoices:
                try:
                    inv_id = str(inv.get('invoice_id') or inv.get('invoice_no')
                                 or inv.get('number') or inv.get('id') or '')
                    if not inv_id:
                        continue
                    grand = round_aed(inv.get('grand_total', 0))
                    paid = round_aed(inv.get('amount_paid', 0))
                    bal = round_aed(inv.get('balance_due', 0))
                    # Sanity: if stored fields disagree, fall back to (grand - balance)
                    if grand > 0 and abs(paid + bal - grand) > Decimal("0.02"):
                        paid = round_aed(grand - bal)
                    if paid <= 0:
                        continue
                    # Find any payments[] array stored on invoice (prefer exact dates)
                    explicit = inv.get('payments') or inv.get('payment_history') or []
                    if isinstance(explicit, list) and len(explicit) > 0:
                        for p in explicit:
                            try:
                                amt = round_aed(p.get('amount') or p.get('paid') or 0)
                                if amt <= 0:
                                    continue
                                pdt = p.get('date') or p.get('payment_date') or inv.get('date') or datetime.now().strftime('%Y-%m-%d')
                                note = str(p.get('method') or p.get('payment_method') or '') + (" | " + str(p.get('note') or '') if p.get('note') else '')
                                posted_pmt = self.on_payment_received(inv_id, amt, pdt, "HistoricalReconcile", note.strip(" |"))
                                if posted_pmt:
                                    stats["payments"] += 1
                            except Exception as ex:
                                stats["warnings"].append(f"Invoice {inv_id} explicit payment: {ex}")
                    else:
                        # No explicit payments array -> post 1 lump-sum payment at invoice.date (or today)
                        pdt = inv.get('date') or datetime.now().strftime('%Y-%m-%d')
                        note = f"Reconstructed: {paid} paid against {inv_id}"
                        posted_lump = self.on_payment_received(inv_id, paid, pdt, "HistoricalReconcile", note)
                        if posted_lump:
                            stats["payments"] += 1
                except Exception as ex:
                    inv_id = inv.get('invoice_id') or inv.get('id') or '?'
                    stats["warnings"].append(f"Reconstruct payment for {inv_id}: {ex}")
        except Exception as ex:
            stats["warnings"].append(f"Reconstruct payments: {ex}")

        # Step 3: All historical purchases
        try:
            purchases = self.dm.load_json('purchases_data.json') or []
            purchases = [p for p in purchases if isinstance(p, dict)]
            log(f"Found {len(purchases)} purchases")
            for pur in purchases:
                try:
                    posted = self.on_purchase_received(pur, username="HistoricalReconcile")
                    if posted:
                        stats["purchases"] += 1
                except Exception as ex:
                    po_id = pur.get('purchase_id') or pur.get('id') or '?'
                    stats["warnings"].append(f"Purchase {po_id}: {ex}")
        except Exception as ex:
            stats["warnings"].append(f"Load purchases: {ex}")

        # Step 4: Expenses from transaction_manager storage
        for fname in ['transactions.json', 'transaction_data.json', 'finance_transactions.json', 'transactions_data.json']:
            try:
                txns = self.dm.load_json(fname) or []
                if not isinstance(txns, list) or len(txns) == 0:
                    continue
                log(f"Found {len(txns)} records in {fname}")
                for t in txns:
                    if not isinstance(t, dict):
                        continue
                    ttype = str(t.get('type') or t.get('transaction_type') or '').lower()
                    if ttype != 'expense':
                        continue
                    try:
                        amt = round_aed(t.get('amount') or t.get('debit') or t.get('value') or 0)
                        if amt <= 0:
                            continue
                        raw_cat = t.get('category') or t.get('subtype') or t.get('expense_category') or 'Other'
                        cat = self.normalize_expense_category(raw_cat)
                        dt = t.get('date') or t.get('transaction_date') or t.get('created_at') or datetime.now().strftime('%Y-%m-%d')
                        if isinstance(dt, str) and len(dt) >= 19 and 'T' in dt:
                            dt = dt[:10]
                        payee = t.get('payee') or t.get('supplier') or t.get('to') or t.get('description') or ''
                        desc = t.get('note') or t.get('memo') or t.get('description') or ''
                        posted = self.on_expense_submitted(cat, amt, dt, payee, desc, "HistoricalReconcile")
                        if posted:
                            stats["expenses"] += 1
                    except Exception as ex:
                        stats["warnings"].append(f"Expense row: {ex}")
            except Exception as ex:
                stats["warnings"].append(f"Load {fname}: {ex}")

        # Step 4b: Bank / cash-book style withdrawals (deposit / withdrawal types
        # in transactions.json but no type=expense). These often represent expenses.
        # If the description is NOT "Invoice Payment", "Purchase reversal", "Reverse payment"
        # etc., treat it as an EXPENSE and classify category by description / vendor keywords.
        try:
            for fname in ['transactions.json', 'transactions_data.json',
                          'transaction_data.json', 'finance_transactions.json']:
                txns = self.dm.load_json(fname) or []
                if not isinstance(txns, list) or len(txns) == 0:
                    continue
                n_classified = 0
                for t in txns:
                    if not isinstance(t, dict):
                        continue
                    ttype = str(t.get('type') or t.get('transaction_type') or t.get('kind') or '').lower()
                    # Skip entries we'd already have classified in Step 4
                    if ttype == 'expense':
                        continue
                    desc = str(t.get('description') or t.get('note') or t.get('memo') or '').lower()
                    amt = round_aed(t.get('amount') or t.get('debit') or t.get('value') or 0)
                    if amt <= 0:
                        continue
                    dt = (t.get('date') or t.get('transaction_date') or t.get('created_at') or
                          datetime.now().strftime('%Y-%m-%d'))
                    if isinstance(dt, str) and len(dt) >= 19 and 'T' in dt:
                        dt = dt[:10]
                    elif isinstance(dt, str) and ' ' in dt:
                        dt = dt.split(' ')[0]
                    payee = (t.get('payee') or t.get('supplier') or t.get('to') or
                             t.get('vendor_or_payee') or t.get('description') or '')
                    # --- HEURISTIC 1: Withdrawal that's NOT invoice payment / purchase reversal
                    if ttype in ('withdrawal', 'withdraw', 'debit', 'payment', 'out'):
                        # Skip known non-expense withdrawals
                        skip_keywords = ('invoice payment', 'reverse payment', 'reversing payment',
                                         'purchase reversal', 'purchase return',
                                         'manual balance adjustment', 'opening balance',
                                         'reversal of', 'credit note', 'refund')
                        if any(kw in desc for kw in skip_keywords):
                            continue
                        # Classify category from description / payee keywords
                        cat = self._expense_category_from_text(payee + ' ' + t.get('description',''))
                        desc_pretty = t.get('description') or payee or 'Bank / Cash Expense'
                        try:
                            posted = self.on_expense_submitted(
                                cat, amt, dt, payee, desc_pretty, "HistoricalReconcile")
                            if posted:
                                stats["expenses"] += 1
                                n_classified += 1
                        except Exception as ex:
                            stats["warnings"].append(f"Withdrawal-expense {t.get('id','?')}: {ex}")
                    # --- HEURISTIC 2: Deposit = Owner Contribution if it looks like equity
                    elif ttype in ('deposit', 'credit', 'in') and any(k in desc for k in (
                            'owner', 'capital', 'contribution', 'equity inject', 'startup capital')):
                        try:
                            ref = f"EQUITY-CONTRIB-{dt.replace('-','')}-{abs(hash(str(amt)+desc)) % 10000:04d}"
                            if not self._entry_exists('EQUITY', ref):
                                batch = [
                                    {'date': dt, 'account_code': '1100', 'debit_amount': amt, 'credit_amount': 0,
                                     'reference_type': 'EQUITY', 'reference_id': ref,
                                     'description': f"Owner Capital — {t.get('description','')}"},
                                    {'date': dt, 'account_code': '3100', 'debit_amount': 0, 'credit_amount': amt,
                                     'reference_type': 'EQUITY', 'reference_id': ref,
                                     'description': f"Owner Capital Increase — {t.get('description','')}"},
                                ]
                                self._write_gl_batch(batch, username="HistoricalReconcile")
                        except Exception as ex:
                            stats["warnings"].append(f"Deposit-contrib: {ex}")
                    # --- HEURISTIC 3: Withdrawal = Owner Drawing if keywords match
                    if ttype in ('withdrawal', 'withdraw', 'debit', 'out') and any(k in desc for k in (
                            'owner drawing', 'owner withdraw', 'draw for owner', 'personal use',
                            'distribution to owner', 'dividend', 'withdrawal by owner')):
                        try:
                            ref = f"EQUITY-DRAW-{dt.replace('-','')}-{abs(hash(str(amt)+desc)) % 10000:04d}"
                            if not self._entry_exists('EQUITY', ref):
                                batch = [
                                    {'date': dt, 'account_code': '3200', 'debit_amount': amt, 'credit_amount': 0,
                                     'reference_type': 'EQUITY', 'reference_id': ref,
                                     'description': f"Owner Drawings — {t.get('description','')}"},
                                    {'date': dt, 'account_code': '1100', 'debit_amount': 0, 'credit_amount': amt,
                                     'reference_type': 'EQUITY', 'reference_id': ref,
                                     'description': f"Owner Drawings Bank Outflow — {t.get('description','')}"},
                                ]
                                self._write_gl_batch(batch, username="HistoricalReconcile")
                        except Exception as ex:
                            stats["warnings"].append(f"Withdrawal-draw: {ex}")
                if n_classified:
                    log(f"Inferred {n_classified} expenses from withdrawals in {fname}")
        except Exception as ex:
            stats["warnings"].append(f"Scan withdrawals for expenses: {ex}")

        # Step 5: Owner Contributions / Drawings if balance_manager stores them
        try:
            for fname in ['balance_entries.json', 'balances.json', 'balance_data.json']:
                entries = self.dm.load_json(fname) or []
                if not isinstance(entries, list):
                    continue
                for b in entries:
                    if not isinstance(b, dict):
                        continue
                    btype = str(b.get('type') or b.get('entry_type') or '').lower()
                    if btype in ('contribution', 'owner_contribution', 'capital', 'equity'):
                        try:
                            amt = round_aed(b.get('amount', 0))
                            if amt <= 0:
                                continue
                            dt = b.get('date') or datetime.now().strftime('%Y-%m-%d')
                            ref = f"EQUITY-CONTRIB-{dt.replace('-','')}-{abs(hash(str(amt)))%10000:04d}"
                            if self._entry_exists('EQUITY', ref):
                                continue
                            batch = [
                                {'date': dt, 'account_code': '1100', 'debit_amount': amt, 'credit_amount': 0,
                                 'reference_type': 'EQUITY', 'reference_id': ref,
                                 'description': f"Owner Capital Contribution — Bank Increase"},
                                {'date': dt, 'account_code': '3100', 'debit_amount': 0, 'credit_amount': amt,
                                 'reference_type': 'EQUITY', 'reference_id': ref,
                                 'description': f"Owner Capital Contribution — Equity Increase"},
                            ]
                            self._write_gl_batch(batch, username="HistoricalReconcile")
                        except Exception as ex:
                            stats["warnings"].append(f"Contribution: {ex}")
                    elif btype in ('drawings', 'drawing', 'withdrawal', 'owner_drawings'):
                        try:
                            amt = round_aed(b.get('amount', 0))
                            if amt <= 0:
                                continue
                            dt = b.get('date') or datetime.now().strftime('%Y-%m-%d')
                            ref = f"EQUITY-DRAW-{dt.replace('-','')}-{abs(hash(str(amt)))%10000:04d}"
                            if self._entry_exists('EQUITY', ref):
                                continue
                            batch = [
                                {'date': dt, 'account_code': '3200', 'debit_amount': amt, 'credit_amount': 0,
                                 'reference_type': 'EQUITY', 'reference_id': ref,
                                 'description': f"Owner Drawings — Equity Reduction"},
                                {'date': dt, 'account_code': '1100', 'debit_amount': 0, 'credit_amount': amt,
                                 'reference_type': 'EQUITY', 'reference_id': ref,
                                 'description': f"Owner Drawings Paid via Bank"},
                            ]
                            self._write_gl_batch(batch, username="HistoricalReconcile")
                        except Exception as ex:
                            stats["warnings"].append(f"Drawings: {ex}")
        except Exception as ex:
            stats["warnings"].append(f"Scan balance entries: {ex}")

        # Step 6: Employee Salaries / Payroll from monthly_salaries.json and salaries_data.json
        stats["salaries"] = 0
        salary_files = ['monthly_salaries.json', 'salaries_data.json']
        for sfname in salary_files:
            try:
                sdata = self.dm.load_json(sfname) or {}
                if not isinstance(sdata, dict):
                    continue
                months_processed = 0
                for mk, entries in sdata.items():
                    if not mk or mk.startswith("__"):
                        continue
                    if not isinstance(entries, list):
                        continue
                    months_processed += 1
                    for se in entries:
                        if not isinstance(se, dict):
                            continue
                        try:
                            emp_id = se.get('employee_id') or se.get('emp_id') or se.get('id') or ''
                            emp_name = se.get('name') or se.get('employee_name') or se.get('full_name') or ''
                            calc = se.get('calculated_values') or {}
                            gross = calc.get('gross_salary', 0)
                            ded = calc.get('total_deductions', 0)
                            net = calc.get('net_salary', 0)
                            # NEW: small alias helpers for salary_components keys.
                            # EM saves 'basic_salary', 'housing_allowance',
                            # 'transport_allowance', 'phone_allowance', but
                            # legacy rebuild short-key expects 'basic', 'housing',
                            # 'transport', 'other_allowances'.  Try BOTH sets.
                            def _g(d, *keys, default=0.0):
                                for k in keys:
                                    v = d.get(k)
                                    if v not in (None, "", 0, 0.0):
                                        try: return float(v)
                                        except Exception: pass
                                try: return float(default)
                                except Exception: return 0.0

                            if gross is None or gross == 0:
                                sc = se.get('salary_components') or {}
                                if isinstance(sc, dict):
                                    gross = (_g(sc, 'basic', 'basic_salary') +
                                             _g(sc, 'housing', 'housing_allowance') +
                                             _g(sc, 'transport', 'transport_allowance') +
                                             _g(sc, 'other_allowances', 'phone_allowance',
                                                'meal_allowance', 'other_allowance'))
                                    vp = se.get('variable_pay') or {}
                                    if isinstance(vp, dict):
                                        gross += (_g(vp, 'performance_bonus', 'bonus') +
                                                  _g(vp, 'overtime') +
                                                  _g(vp, 'commission', 'sales_commission'))
                                    ded_total = 0.0
                                    d = se.get('deductions') or {}
                                    if isinstance(d, dict):
                                        ded_total = (_g(d, 'loan_installment', 'loan') +
                                                     _g(d, 'late_penalties', 'late_penalty',
                                                        'penalties') +
                                                     _g(d, 'income_tax', 'tax') +
                                                     _g(d, 'social_insurance', 'insurance') +
                                                     _g(d, 'other'))
                                    ded = ded_total
                                    net = gross - ded_total
                            proc_dt = (se.get('processed_date') or se.get('payment_date') or
                                       se.get('date_paid') or se.get('month_year') or
                                       f"{mk}-28")
                            if isinstance(proc_dt, str) and len(proc_dt) == 7 and proc_dt[4] == '-':
                                proc_dt = f"{proc_dt}-28"
                            pay_method = "Bank"
                            bd = se.get('bank_details') or {}
                            if isinstance(bd, dict) and bd.get('bank_name'):
                                pay_method = "Bank"
                            status_ok = str(se.get('status') or '').lower()
                            # NEW SIMPLIFIED BEHAVIOUR: during the historical
                            # rebuild, treat EVERY row with a valid gross/net as
                            # posted EXCEPT explicit 'cancelled' / 'rejected'.
                            # This solves the "Salary Expense = 0 forever" bug
                            # because 90%+ of SMEs never individually approve
                            # each month — they just enter data in bulk.  The
                            # payroll form can still use status=draft as a visual
                            # tag, but financially the employee earned X so the
                            # P&L MUST reflect that salary expense on day one of
                            # import, not after someone manually approves each
                            # month.  (Explicitly cancelled/rejected months are
                            # excluded per professional practice.)
                            posted_this_row = False
                            if status_ok in ('cancelled', 'rejected'):
                                pass
                            elif (float(gross or 0) > 0 or float(net or 0) > 0
                                  or float(ded or 0) > 0):
                                posted = self.on_salary_payroll_posted(
                                    str(mk), str(emp_id), str(emp_name),
                                    float(gross or 0), float(ded or 0), float(net or 0),
                                    processed_date=str(proc_dt),
                                    payment_method=pay_method,
                                    username=("HistoricalReconcile"
                                              if status_ok in ('draft', 'pending', '')
                                              else "HistoricalReconcile"))
                                posted_this_row = bool(posted)
                            elif status_ok and status_ok not in ('draft', 'pending', 'cancelled', 'rejected'):
                                # Backward compat for explicitly-approved old rows
                                posted = self.on_salary_payroll_posted(
                                    str(mk), str(emp_id), str(emp_name),
                                    float(gross or 0), float(ded or 0), float(net or 0),
                                    processed_date=str(proc_dt),
                                    payment_method=pay_method,
                                    username="HistoricalReconcile")
                                posted_this_row = bool(posted)
                            if posted_this_row:
                                stats["salaries"] += 1
                        except Exception as ex:
                            stats["warnings"].append(f"Salary {mk}/{se.get('employee_id','?')}: {ex}")
                if months_processed:
                    log(f"Processed {months_processed} month(s) in {sfname}")
            except Exception as ex:
                stats["warnings"].append(f"Load salaries {sfname}: {ex}")

        # Return summary
        gl = self.dm.load_json('general_ledger.json') or []
        stats["total_gl_after"] = len(gl)
        log(f"Rebuild done: {stats['invoices']} invoices, {stats['payments']} payments, "
            f"{stats['purchases']} purchases, {stats['expenses']} expenses, "
            f"{stats.get('salaries', 0)} salary entries -> GL now has {stats['total_gl_after']} entries, "
            f"{len(stats['warnings'])} warnings")
        return stats

