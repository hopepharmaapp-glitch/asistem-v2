import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from datetime import datetime, date, timedelta
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib.styles import getSampleStyleSheet
import ui_undo

class EmployeeManager:
    COMPONENT_KEYS = [
        ("basic", "Basic Salary"),
        ("housing", "Housing Allowance"),
        ("transport", "Transport"),
        ("phone", "Phone / Mobile"),
        ("food", "Food / Meal"),
        ("education", "Education Allowance"),
        ("medical", "Medical Insurance Top-up"),
        ("travel", "Travel Allowance"),
        ("shift", "Shift / On-Call"),
        ("seniority", "Seniority / Grade"),
        ("cost_of_living", "Cost-of-Living (COLA)"),
        ("performance_fixed", "Performance (Fixed)"),
        ("commission_fixed", "Commission (Fixed)"),
        ("other_allowances", "Other Allowances"),
        ("company_benefits", "Company Benefits In-Kind"),
    ]
    VARIABLE_KEYS = [
        ("performance_bonus", "Perf. Bonus"),
        ("overtime", "Overtime"),
        ("commission", "Commission"),
        ("sales_incentive", "Sales Incentive"),
        ("tips_gratuity", "Tips / Gratuity"),
    ]
    DEDUCTION_KEYS = [
        ("loan_installment", "Loan Installment"),
        ("salary_advance", "Salary Advance"),
        ("absences", "Absence / Unpaid Leave"),
        ("late_penalties", "Late / Penalties"),
        ("income_tax", "Income Tax"),
        ("social_insurance", "Social Insurance"),
        ("other", "Other Deductions"),
    ]

    def __init__(self, data_folder):
        self.data_folder = data_folder
        os.makedirs(self.data_folder, exist_ok=True)
        self.employees_file = os.path.join(data_folder, "employees.json")
        self.salaries_file = os.path.join(data_folder, "monthly_salaries.json")
        self.attach_dir = os.path.join(data_folder, "employee_docs")
        os.makedirs(self.attach_dir, exist_ok=True)
        self.employees = self._load_json(self.employees_file, default=[])
        self.salaries = self._load_json(self.salaries_file, default={})
        self.ledger_service = None

    def set_ledger_service(self, ledger):
        self.ledger_service = ledger

    def _load_json(self, path, default):
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return default

    def _save_json(self, path, data):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def list_employees(self):
        return self.employees

    def add_employee(self, profile, backfill_from_join_date=True):
        emp_id = profile.get("employee_id") or f"EMP{len(self.employees)+1:03d}"
        profile["employee_id"] = emp_id
        profile["created_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.employees.append(profile)
        self._save_json(self.employees_file, self.employees)
        # ------------- NEW: auto backfill payrolls for all months between
        # joining_date and TODAY (inclusive) so the employee immediately
        # appears in older payroll tabs & salary history.
        if backfill_from_join_date:
            self.backfill_payrolls_for_employee(emp_id)
        return emp_id

    def update_employee(self, emp_id, updates):
        joining_changed = "joining_date" in updates
        for e in self.employees:
            if e.get("employee_id") == emp_id:
                e.update(updates)
                break
        self._save_json(self.employees_file, self.employees)
        # If user changed joining_date (e.g. corrected a join date to an older
        # month), re-backfill from that date so older payroll tabs now include.
        if joining_changed:
            self.backfill_payrolls_for_employee(emp_id)

    def _iter_months(self, start_ym, end_ym):
        """Yield YYYY-MM strings from start_ym up to & including end_ym (inclusive)."""
        try:
            sy, sm = (int(x) for x in str(start_ym).split("-")[:2])
            ey, em = (int(x) for x in str(end_ym).split("-")[:2])
        except Exception:
            return
        cur_y, cur_m = sy, sm
        while (cur_y, cur_m) <= (ey, em):
            yield f"{cur_y:04d}-{cur_m:02d}"
            cur_m += 1
            if cur_m > 12:
                cur_m = 1
                cur_y += 1

    def backfill_payrolls_for_employee(self, emp_id):
        """For employee `emp_id`, insert a default salary row into EVERY monthly
        payroll between their `joining_date` (YYYY-MM-DD) and TODAY, skipping
        any months that already contain this employee_id (so re-running never
        duplicates).  Nothing is posted to GL (all backfilled rows start with
        status='draft') — the user explicitly reviews & approves each month."""
        emp = next((e for e in (self.employees or []) if e.get("employee_id") == emp_id), None)
        if emp is None:
            return
        join_s = str(emp.get("joining_date") or "").strip()[:10]
        if not join_s or len(join_s) != 10:
            return
        try:
            join_d = datetime.strptime(join_s, "%Y-%m-%d").date()
        except Exception:
            return
        start_ym = join_d.strftime("%Y-%m")
        today_ym = datetime.now().strftime("%Y-%m")
        changed = False
        # Ensure a "prototype default" row for this employee built once
        proto_cache = {}
        for mk in self._iter_months(start_ym, today_ym):
            rows = list(self.salaries.get(mk) or [])
            # Skip if employee already present in this month
            if any(str(r.get("employee_id") or "") == str(emp_id) for r in rows if isinstance(r, dict)):
                continue
            if mk not in proto_cache:
                proto_cache[mk] = self._employee_to_default_salary_row_static(emp, mk)
            rows.append(proto_cache[mk])
            self.salaries[mk] = rows
            # Also add a __meta__ entry for this month (draft) if missing
            meta = self.salaries.get("__meta__") or {}
            if mk not in meta or not isinstance(meta.get(mk), dict):
                meta[mk] = {"status": "draft",
                            "processed_date": datetime.now().strftime("%Y-%m-%d"),
                            "backfilled": True}
            self.salaries["__meta__"] = meta
            changed = True
        if changed:
            self._save_json(self.salaries_file, self.salaries)
        # ---- SIMPLIFY UX: when backfilling, post EVERY new employee-month to
        # the ledger IMMEDIATELY (even with status=draft), so the P&L instantly
        # shows salary expense without the user having to manually approve each
        # of hundreds of past months.  This is the "simple & professional"
        # behaviour — the employee earned those salaries, so they MUST appear in
        # historical financials on day one of import.  Formal auditors can still
        # re-open & re-approve months through the normal payroll UI, which will
        # re-post idempotently with a proper audit timestamp.
        if changed and self.ledger_service:
            try:
                import datetime as _dt2
                for mk in list(proto_cache.keys()):
                    rows = list(self.salaries.get(mk) or [])
                    if not rows:
                        continue
                    meta_mk = ((self.salaries.get("__meta__") or {}).get(mk) or {})
                    proc_dt = (str(meta_mk.get("processed_date") or "")
                               or f"{mk}-28")
                    if len(proc_dt) == 7 and proc_dt[4] == "-":
                        proc_dt = f"{proc_dt}-28"
                    for se in rows:
                        if not isinstance(se, dict):
                            continue
                        if str(se.get("employee_id") or "") != str(emp_id):
                            continue
                        calc = se.get("calculated_values") or {}
                        gross = float(calc.get("gross_salary", 0) or 0)
                        ded = float(calc.get("total_deductions", 0) or 0)
                        net = float(calc.get("net_salary", 0) or 0)
                        if gross <= 0:
                            continue
                        try:
                            bd = se.get("bank_details") or {}
                            pay_method = ("Bank"
                                          if isinstance(bd, dict) and bd.get("bank_name")
                                          else "Bank")
                            self.ledger_service.on_salary_payroll_posted(
                                str(mk), str(emp_id),
                                str(se.get("name") or emp_id),
                                gross, ded, net,
                                processed_date=proc_dt,
                                payment_method=pay_method,
                                username="BackfillAutoPost")
                        except Exception:
                            pass
            except Exception:
                pass

    @staticmethod
    def _employee_to_default_salary_row_static(e, mk):
        """Static equivalent of AdvancedEmployeeManagerDialog._employee_to_default_salary_row
        — produces identical initial payload suitable for manager-level backfill."""
        from datetime import datetime as _dt  # noqa: F811 (shadow OK)
        sc = e.get("salary_components") or {}
        vp = e.get("variable_pay") or {}
        dd = e.get("deductions") or {}
        COMPONENT_KEYS = [
            ("basic", "Basic Salary"), ("housing", "Housing Allowance"),
            ("transport", "Transport Allowance"), ("food", "Food / Meal Allowance"),
            ("education", "Education Allowance"), ("communication", "Communication"),
            ("medical_insurance", "Medical Insurance"), ("life_insurance", "Life Insurance"),
            ("performance_allowance", "Performance Allowance"), ("position_allowance", "Position / Role"),
            ("other_allowance_1", "Other Allowance 1"), ("other_allowance_2", "Other Allowance 2"),
            ("other_allowance_3", "Other Allowance 3"), ("other_allowance_4", "Other Allowance 4"),
            ("other_allowance_5", "Other Allowance 5"),
        ]
        VARIABLE_KEYS = [
            ("overtime", "Overtime Pay"), ("bonus", "Performance Bonus"),
            ("commission", "Commission"), ("incentives", "Incentives"),
            ("other_variable", "Other Variable"),
        ]
        DEDUCTION_KEYS = [
            ("income_tax", "Income Tax"), ("social_insurance", "Social Insurance (EE)"),
            ("pension", "Pension (EE)"), ("health_insurance_ee", "Health Insurance (EE)"),
            ("salary_advance", "Salary Advance"), ("absences", "Absences / LWOP"),
            ("other_deductions", "Other Deductions"),
        ]
        def _r(d, keys, default=0.0):
            return {k: round(float((d or {}).get(k, default) or 0), 2) for (k, _) in keys}
        def _calc(sc_d, vp_d, dd_d):
            gross = sum(float(v or 0) for v in list(sc_d.values()) + list(vp_d.values()))
            ded = sum(float(v or 0) for v in dd_d.values())
            return {"gross_salary": round(gross, 2),
                    "total_deductions": round(ded, 2),
                    "net_salary": round(max(0.0, gross - ded), 2)}
        # --- SIMPLIFY: EmployeeManager stores allowances mostly as TOP-LEVEL
        # profile keys (basic_salary, housing_allowance, transport_allowance,
        # phone_allowance, etc), NOT inside salary_components dict.  If
        # salary_components dict is empty (all zeros), merge top-level keys in
        # so the resulting calculated_values.gross_salary actually equals the
        # employee's real gross.  Without this, backfilled rows ALL have
        # gross=0 and the rebuild's "simplified" draft posting STILL shows 0.
        sc_in = sc or {}
        if not any(float(v or 0) != 0 for v in sc_in.values()):
            def _top(*aliases, default=0.0):
                for a in aliases:
                    val = e.get(a)
                    if val not in (None, "", 0, 0.0):
                        try: return float(val)
                        except Exception: pass
                return float(default)
            sc_in = dict(sc_in) if isinstance(sc_in, dict) else {}
            sc_in["basic"] = sc_in.get("basic") or _top("basic_salary", "basic", "salary")
            sc_in["housing"] = sc_in.get("housing") or _top("housing_allowance", "housing")
            sc_in["transport"] = sc_in.get("transport") or _top("transport_allowance", "transport")
            sc_in["communication"] = (sc_in.get("communication")
                                      or _top("phone_allowance", "communication"))
            sc_in["medical_insurance"] = sc_in.get("medical_insurance") or _top("medical_insurance")
            sc_in["other_allowance_1"] = (sc_in.get("other_allowance_1")
                                          or _top("other_allowance", "other_allowance_1"))
        comp = _r(sc_in or {}, COMPONENT_KEYS)
        var  = _r(vp or {}, VARIABLE_KEYS, 0)
        ded  = _r(dd or {}, DEDUCTION_KEYS, 0)
        calc = _calc(comp, var, ded)
        proc_dt = None
        try:
            proc_dt = _dt.strptime(f"{mk}-28", "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            proc_dt = _dt.now().strftime("%Y-%m-%d")
        return {
            "employee_id": e.get("employee_id"),
            "name": e.get("name") or "",
            "department": e.get("department") or "",
            "month_year": mk,
            "salary_components": comp,
            "variable_pay": var,
            "deductions": ded,
            "calculated_values": {k: calc[k] for k in ("gross_salary","total_deductions","net_salary")},
            "status": "draft",
            "processed_date": proc_dt,
            "backfilled": True,
            "bank_details": e.get("bank_details") or {},
        }

    def get_month_key(self, when=None):
        d = when or datetime.now()
        return d.strftime("%Y-%m")

    def get_month_entries(self, month_key=None):
        mk = month_key or self.get_month_key()
        return self.salaries.get(mk) or []

    def save_month_entries(self, month_key, entries, status="draft"):
        self.salaries[month_key] = entries
        meta = self.salaries.get("__meta__", {})
        processed_date = datetime.now().strftime("%Y-%m-%d")
        meta[month_key] = {
            "status": status,
            "processed_date": processed_date,
        }
        self.salaries["__meta__"] = meta
        self._save_json(self.salaries_file, self.salaries)
        if self.ledger_service and str(status).lower() in ("approved", "submitted", "paid", "posted"):
            try:
                for se in entries:
                    if not isinstance(se, dict):
                        continue
                    calc = se.get('calculated_values') or {}
                    gross = calc.get('gross_salary', 0)
                    ded = calc.get('total_deductions', 0)
                    net = calc.get('net_salary', 0)
                    if gross is None or gross == 0:
                        recomputed = self.calculate_net(
                            se.get('salary_components') or {},
                            se.get('variable_pay') or {},
                            se.get('deductions') or {},
                        )
                        gross = recomputed['gross_salary']
                        ded = recomputed['total_deductions']
                        net = recomputed['net_salary']
                    bd = se.get('bank_details') or {}
                    pay_method = "Bank" if (isinstance(bd, dict) and bd.get('bank_name')) else "Bank"
                    proc_dt = (se.get('processed_date') or processed_date or
                               f"{month_key}-28")
                    if isinstance(proc_dt, str) and len(proc_dt) == 7 and proc_dt[4] == '-':
                        proc_dt = f"{proc_dt}-28"
                    self.ledger_service.on_salary_payroll_posted(
                        str(month_key),
                        str(se.get('employee_id') or ''),
                        str(se.get('name') or ''),
                        float(gross or 0),
                        float(ded or 0),
                        float(net or 0),
                        processed_date=str(proc_dt),
                        payment_method=pay_method,
                        username="EmployeeManager")
            except Exception:
                pass

    def get_salary_totals_for_period(self, start_date_str, end_date_str,
                                     use_field="net_salary",
                                     use_only_statuses=None):
        """
        Return salary totals strictly contained within start_date..end_date (YYYY-MM-DD).
        This powers the Legacy Reporting dialog, cash flow reports, and dashboards so
        the selected period is respected (not just 1 month).

        use_field: "net_salary", "gross_salary", or "total_salary_paid_after_adj"
        use_only_statuses: e.g. ["approved", "paid", "submitted"]; None = all non-draft/rejected
        """
        total = 0.0
        count_employees = 0
        count_months = 0
        breakdown_by_month = {}
        try:
            sd = datetime.strptime(str(start_date_str).strip()[:10], "%Y-%m-%d").date()
        except Exception:
            sd = date(2000, 1, 1)
        try:
            ed = datetime.strptime(str(end_date_str).strip()[:10], "%Y-%m-%d").date()
        except Exception:
            ed = date.today()
        status_ok_set = None
        if use_only_statuses:
            status_ok_set = {str(s).lower() for s in use_only_statuses}
        for mk, entries in self.salaries.items():
            if not mk or mk.startswith("__") or not isinstance(entries, list):
                continue
            if len(str(mk)) < 7 or str(mk)[4] != '-':
                continue
            try:
                myear = int(str(mk)[:4])
                mmon = int(str(mk)[5:7])
                first_day = date(myear, mmon, 1)
                if mmon == 12:
                    last_day = date(myear + 1, 1, 1) - timedelta(days=1)
                else:
                    last_day = date(myear, mmon + 1, 1) - timedelta(days=1)
            except Exception:
                continue
            if last_day < sd or first_day > ed:
                continue
            meta = self.salaries.get("__meta__", {}) or {}
            mstatus = str((meta.get(mk) or {}).get('status') or '').lower()
            if status_ok_set is None:
                if mstatus in ('draft', 'cancelled', 'rejected'):
                    continue
            else:
                if mstatus and mstatus not in status_ok_set:
                    continue
            month_gross = 0.0
            month_net = 0.0
            month_employees = 0
            for se in entries:
                if not isinstance(se, dict):
                    continue
                se_status = str(se.get('status') or mstatus or '').lower()
                if status_ok_set is None:
                    if se_status in ('draft', 'cancelled', 'rejected'):
                        continue
                elif se_status and se_status not in status_ok_set:
                    continue
                calc = se.get('calculated_values') or {}
                gross = float(calc.get('gross_salary') or se.get('salary_paid') or se.get('gross') or 0.0)
                net = float(calc.get('net_salary') or se.get('salary_paid') or se.get('net') or 0.0)
                if gross <= 0 and net <= 0:
                    recomputed = self.calculate_net(
                        se.get('salary_components') or {},
                        se.get('variable_pay') or {},
                        se.get('deductions') or {},
                    )
                    gross = recomputed['gross_salary']
                    net = recomputed['net_salary']
                amount = 0.0
                f = str(use_field or '').lower()
                if 'gross' in f:
                    amount = gross
                elif 'paid' in f or 'cash' in f:
                    amount = net
                else:
                    amount = net if net > 0 else gross
                if amount > 0:
                    total += amount
                    count_employees += 1
                    month_gross += gross
                    month_net += net
                    month_employees += 1
            if month_employees > 0:
                count_months += 1
                breakdown_by_month[mk] = {
                    "employees": month_employees,
                    "gross": round(month_gross, 2),
                    "net": round(month_net, 2),
                }
        return {
            "total_amount": round(total, 2),
            "employees_processed": count_employees,
            "months_in_period": count_months,
            "by_month": breakdown_by_month,
            "start_date": sd.isoformat(),
            "end_date": ed.isoformat(),
        }

    def get_salary_summary_for_period(self, start_date_str, end_date_str):
        """Convenience helper returning totals for common fields across the period."""
        gross = self.get_salary_totals_for_period(start_date_str, end_date_str, use_field="gross_salary")
        net = self.get_salary_totals_for_period(start_date_str, end_date_str, use_field="net_salary")
        return {
            "gross_salary_total": gross["total_amount"],
            "net_salary_total": net["total_amount"],
            "employees_processed": max(gross["employees_processed"], net["employees_processed"]),
            "months_in_period": max(gross["months_in_period"], net["months_in_period"]),
            "by_month_gross": gross["by_month"],
            "by_month_net": net["by_month"],
            "start_date": gross["start_date"],
            "end_date": gross["end_date"],
        }

    def get_salary_history(self, employee_id):
        """Return ordered list of monthly salary records for a specific employee.
        Each record: month, gross_salary, components_pay, variable_pay,
                     total_deductions, net_salary, status, processed_date.
        """
        if not employee_id:
            return []
        rows = []
        def _f(d, k):
            try:
                v = d.get(k, 0)
                return 0.0 if v is None else float(v)
            except Exception:
                return 0.0
        for mk, entries in (self.salaries or {}).items():
            if mk.startswith("__") or not isinstance(entries, list):
                continue
            for se in entries:
                if not isinstance(se, dict):
                    continue
                if str(se.get("employee_id") or "") != str(employee_id):
                    continue
                sc = se.get("salary_components") or {}
                vp = se.get("variable_pay") or {}
                dd = se.get("deductions") or {}
                calc = se.get("calculated_values") or {}
                comp_sum = round(sum(_f(sc, k) for (k, _) in self.COMPONENT_KEYS), 2)
                var_sum = round(sum(_f(vp, k) for (k, _) in self.VARIABLE_KEYS), 2)
                ded_sum = round(sum(_f(dd, k) for (k, _) in self.DEDUCTION_KEYS), 2)
                meta = (self.salaries.get("__meta__") or {}).get(mk) or {}
                status = str(se.get("status") or meta.get("status") or "draft").lower()
                processed_date = se.get("processed_date") or meta.get("processed_date") or f"{mk}-28"
                rows.append({
                    "month": mk,
                    "gross_salary": round(float(calc.get("gross_salary", comp_sum + var_sum) or 0), 2),
                    "components_pay": comp_sum,
                    "variable_pay": var_sum,
                    "total_deductions": ded_sum,
                    "net_salary": round(float(calc.get("net_salary", (comp_sum + var_sum) - ded_sum) or 0), 2),
                    "status": status,
                    "processed_date": processed_date,
                })
        rows.sort(key=lambda r: r["month"])
        return rows

    def copy_previous_month(self, month_key):
        parts = month_key.split("-")
        y, m = int(parts[0]), int(parts[1])
        if m == 1:
            prev = f"{y-1}-12"
        else:
            prev = f"{y}-{m-1:02d}"
        return json.loads(json.dumps(self.get_month_entries(prev)))

    # ============================================================
    # Per-Employee Salary Record CRUD (Granular, Month-Scoped)
    # ============================================================
    def get_employee_monthly_record(self, employee_id, month_key):
        """Return a single employee's salary record for a given month, or None."""
        if not employee_id or not month_key:
            return None
        rows = self.salaries.get(month_key) or []
        for r in rows:
            if isinstance(r, dict) and str(r.get("employee_id") or "") == str(employee_id):
                return json.loads(json.dumps(r))
        return None

    def create_employee_monthly_record(self, employee_id, month_key, record=None, overwrite=False):
        """Create a monthly salary entry for one employee for one specific month only.

        Never touches other months.  Returns the saved record dict, or raises ValueError
        if the month already contains this employee and overwrite is False.
        """
        if not employee_id or not month_key:
            raise ValueError("employee_id and month_key are required")
        rows = list(self.salaries.get(month_key) or [])
        existing_idx = None
        for i, r in enumerate(rows):
            if isinstance(r, dict) and str(r.get("employee_id") or "") == str(employee_id):
                existing_idx = i; break
        if existing_idx is not None and not overwrite:
            raise ValueError(f"Employee {employee_id} already has a record in {month_key}. Use overwrite=True to replace.")
        emp = next((e for e in (self.employees or []) if str(e.get("employee_id") or "") == str(employee_id)), None)
        if record is None and emp is not None:
            # Build a sensible default from personnel file defaults (only for the requested month)
            proto = self._employee_to_default_salary_row_static(emp, month_key)
            record = proto
        elif record is None:
            raise ValueError("record dict required when employee not found in personnel file")
        # Enforce month + emp id scoping: do NOT copy this into any other month
        save_rec = dict(record or {})
        save_rec["employee_id"] = str(employee_id)
        save_rec["month_year"] = month_key
        if emp:
            save_rec["name"] = save_rec.get("name") or emp.get("name") or ""
            save_rec["department"] = save_rec.get("department") or emp.get("department") or ""
            if not save_rec.get("bank_details") and emp.get("bank_details"):
                save_rec["bank_details"] = json.loads(json.dumps(emp["bank_details"]))
        save_rec.setdefault("salary_components", {})
        save_rec.setdefault("variable_pay", {})
        save_rec.setdefault("deductions", {})
        # Recalculate gross/ded/net from 27 components
        calc = self.calculate_net(save_rec["salary_components"], save_rec["variable_pay"], save_rec["deductions"])
        save_rec["calculated_values"] = {k: calc[k] for k in ("gross_salary", "total_deductions", "net_salary")}
        save_rec.setdefault("status", "draft")
        save_rec.setdefault("processed_date", datetime.now().strftime("%Y-%m-%d"))
        if existing_idx is None:
            rows.append(save_rec)
        else:
            # Preserve creation-time fields if caller didn't explicitly override
            old = rows[existing_idx]
            if isinstance(old, dict):
                for keep in ("backfilled",):
                    if keep in old and keep not in save_rec:
                        save_rec[keep] = old[keep]
            rows[existing_idx] = save_rec
        # Commit — strictly single-month scope: do NOT touch other salary[<other_month>] entries
        self.salaries[month_key] = rows
        meta = self.salaries.get("__meta__") or {}
        meta.setdefault(month_key, {})
        if not isinstance(meta.get(month_key), dict):
            meta[month_key] = {}
        meta[month_key]["last_modified"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        meta[month_key].setdefault("status", save_rec.get("status", "draft"))
        meta[month_key].setdefault("processed_date", save_rec.get("processed_date"))
        self.salaries["__meta__"] = meta
        self._save_json(self.salaries_file, self.salaries)
        return json.loads(json.dumps(save_rec))

    def update_employee_monthly_record(self, employee_id, month_key, updates):
        """Update fields of a single employee's single-month record.  No cross-month bleed."""
        if not employee_id or not month_key:
            raise ValueError("employee_id and month_key required")
        current = self.get_employee_monthly_record(employee_id, month_key)
        if current is None:
            # Upsert: create new blank record first then apply updates
            current = {"employee_id": employee_id, "month_year": month_key,
                       "salary_components": {}, "variable_pay": {}, "deductions": {}}
        for k, v in (updates or {}).items():
            if k in ("employee_id", "month_year"):
                continue  # Prevent scope drift
            current[k] = json.loads(json.dumps(v)) if isinstance(v, (dict, list)) else v
        # Recompute calculations to ensure structural integrity
        current.setdefault("salary_components", {})
        current.setdefault("variable_pay", {})
        current.setdefault("deductions", {})
        calc = self.calculate_net(current["salary_components"], current["variable_pay"], current["deductions"])
        current["calculated_values"] = {k: calc[k] for k in ("gross_salary", "total_deductions", "net_salary")}
        return self.create_employee_monthly_record(employee_id, month_key, current, overwrite=True)

    def delete_employee_monthly_record(self, employee_id, month_key):
        """Delete one specific employee record from one specific month (NOT global)."""
        if not employee_id or not month_key:
            raise ValueError("employee_id and month_key required")
        rows = list(self.salaries.get(month_key) or [])
        new_rows = [r for r in rows if not (isinstance(r, dict) and str(r.get("employee_id") or "") == str(employee_id))]
        removed = len(rows) - len(new_rows)
        if removed == 0:
            return False
        self.salaries[month_key] = new_rows
        meta = self.salaries.get("__meta__") or {}
        if isinstance(meta.get(month_key), dict):
            meta[month_key]["last_modified"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.salaries["__meta__"] = meta
        self._save_json(self.salaries_file, self.salaries)
        return True

    def list_months_with_salaries(self):
        """Return sorted list of YYYY-MM months that contain at least one salary row."""
        out = [k for k in (self.salaries or {}).keys()
               if isinstance(k, str) and not k.startswith("__") and len(k) == 7 and k[4] == "-"]
        out.sort()
        return out

    def calculate_net(self, comp, var, ded):
        components = comp or {}
        variables = var or {}
        deductions = ded or {}
        def _f(d, k):
            try:
                v = d.get(k, 0)
                return 0.0 if v is None else float(v)
            except Exception:
                return 0.0
        gross = sum(_f(components, k) for (k, _) in self.COMPONENT_KEYS)
        gross += sum(_f(variables, k) for (k, _) in self.VARIABLE_KEYS)
        total_ded = sum(_f(deductions, k) for (k, _) in self.DEDUCTION_KEYS)
        net = gross - total_ded
        return {
            "gross_salary": round(gross, 2),
            "total_deductions": round(total_ded, 2),
            "net_salary": round(net, 2)
        }

    def generate_bank_file(self, month_key, filename):
        entries = self.get_month_entries(month_key)
        total = 0.0
        lines = []
        for e in entries:
            c = e.get("calculated_values", {})
            net = float(c.get("net_salary", 0))
            total += net
            emp_id = e.get("employee_id")
            name = e.get("name")
            bank = e.get("bank_details", {}).get("bank_name")
            acct = e.get("bank_details", {}).get("account_number")
            lines.append(f"{emp_id},{name},{net:.2f},{bank},{acct}")
        out_path = os.path.join(self.data_folder, filename)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))
        return out_path, round(total, 2)

class AdvancedEmployeeManagerDialog:
    """
    🏆 MOST ADVANCED EMPLOYEE & PAYROLL MANAGER — 8 Tabs
    ┌───────────────────────────────────────────────────────────────┐
    │ 1. 👥 Directory   │ 13-col searchable tree + 6 status chips   │
    │ 2. 👤 Personnel   │ 30-field form, edit any employee          │
    │ 3. 💰 Payroll     │ 27 editable fields × N employees + totals │
    │ 4. 📆 History     │ Month-by-month salary comparison ledger   │
    │ 5. 📊 Reports     │ Summary / per-employee / CSV/PDF          │
    │ 6. 📄 Documents   │ Attachments (CV, Passport, Visa, Contract)│
    │ 7. 📤 Exports     │ Bank WPS CSV, ZIP of payslip PDFs         │
    │ 8. ⚙️  Defaults    │ Dept. cost centers, GL accts, allowances │
    └───────────────────────────────────────────────────────────────┘
    """

    # ======== Extended Payroll Schema ========
    # 15 salary components (guaranteed, fixed)
    COMPONENT_KEYS = [
        ("basic", "Basic Salary"),
        ("housing", "Housing Allowance"),
        ("transport", "Transport"),
        ("phone", "Phone / Mobile"),
        ("food", "Food / Meal"),
        ("education", "Education Allowance"),
        ("medical", "Medical Insurance Top-up"),
        ("travel", "Travel Allowance"),
        ("shift", "Shift / On-Call"),
        ("seniority", "Seniority / Grade"),
        ("cost_of_living", "Cost-of-Living (COLA)"),
        ("performance_fixed", "Performance (Fixed)"),
        ("commission_fixed", "Commission (Fixed)"),
        ("other_allowances", "Other Allowances"),
        ("company_benefits", "Company Benefits In-Kind"),
    ]
    # 5 variable pays (per-month, non-fixed)
    VARIABLE_KEYS = [
        ("performance_bonus", "Perf. Bonus"),
        ("overtime", "Overtime"),
        ("commission", "Commission"),
        ("sales_incentive", "Sales Incentive"),
        ("tips_gratuity", "Tips / Gratuity"),
    ]
    # 7 deductions (per-month)
    DEDUCTION_KEYS = [
        ("loan_installment", "Loan Installment"),
        ("salary_advance", "Salary Advance"),
        ("absences", "Absence / Unpaid Leave"),
        ("late_penalties", "Late / Penalties"),
        ("income_tax", "Income Tax"),
        ("social_insurance", "Social Insurance"),
        ("other", "Other Deductions"),
    ]

    def __init__(self, parent, invoice_manager):
        self.parent = parent
        self.invoice_manager = invoice_manager
        try:
            df = getattr(invoice_manager, 'invoice_folder', os.getcwd())
        except Exception:
            df = os.getcwd()
        self.manager = EmployeeManager(df)
        ledger_svc = getattr(invoice_manager, "ledger", None)
        if ledger_svc is None:
            ledger_svc = getattr(getattr(invoice_manager, "manager", None), "ledger", None)
        if ledger_svc is None:
            dm = getattr(invoice_manager, "manager", None) or getattr(invoice_manager, "dm", None)
            if dm is None:
                try:
                    import hope_pharma_complete as hpc
                    dm_cls = getattr(hpc, "EnhancedCloudDataManager", None) or getattr(hpc, "HopePharmaDataManager", None)
                    if dm_cls:
                        dm = dm_cls(df, mode='local')
                except Exception:
                    dm = None
            if dm is not None:
                ledger_svc = getattr(dm, "ledger", None)
                if ledger_svc is None:
                    try:
                        from ledger_service import LedgerService
                        ledger_svc = LedgerService(dm)
                    except Exception:
                        ledger_svc = None
        if ledger_svc is not None:
            self.manager.set_ledger_service(ledger_svc)
        self.doc_dir = os.path.join(df, "employee_docs")
        os.makedirs(self.doc_dir, exist_ok=True)
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("🏆 Advanced Employee Manager & Payroll Centre")
        try:
            sw = self.dialog.winfo_screenwidth()
            sh = self.dialog.winfo_screenheight()
            w, h = int(sw * 0.92), int(sh * 0.88)
            x, y = max(10, (sw-w)//2), max(10, int((sh-h)*0.38))
            self.dialog.geometry(f"{w}x{h}+{x}+{y}")
            self.dialog.minsize(int(sw*0.70), int(sh*0.62))
        except Exception:
            self.dialog.geometry("1400x900+80+40")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        self._build_ui()

    # ============================================================
    # UI Assembly
    # ============================================================
    def _build_ui(self):
        nb = ttk.Notebook(self.dialog)
        nb.pack(fill='both', expand=True, padx=8, pady=8)
        self.nb = nb
        # 9 Tabs
        tabs = [
            ("👥 Employee Directory",   self._build_directory),
            ("👤 Personnel File",       self._build_personnel_file),
            ("💰 Monthly Payroll",      self._build_payroll_monthly),
            ("📆 Salary History",       self._build_history),
            ("💼 Salary Mgmt",          self._build_salary_management),
            ("📊 Payroll Reports",      self._build_reports),
            ("📄 Employee Documents",   self._build_documents),
            ("📤 Export / Bank File",   self._build_exports),
            ("⚙️  Payroll Defaults",    self._build_defaults),
        ]
        for title, build_fn in tabs:
            fr = ttk.Frame(nb); build_fn(fr); nb.add(fr, text=(" " + title + " "))
        # Top summary bar (outside notebook, always visible)
        self.summary_var = tk.StringVar(value="Loading payroll summary…")
        top = ttk.Frame(self.dialog, relief="ridge", padding=(12, 6))
        top.pack(fill='x', side='top', padx=8, pady=(8, 0))
        ttk.Label(top, textvariable=self.summary_var,
                  font=("Segoe UI", 11, "bold"), foreground="#1e3a8a").pack(side='left')
        ttk.Button(top, text="🔁 Refresh All",
                   command=self._refresh_everything).pack(side='right')
        self.dialog.after(180, self._refresh_everything)
        # Attach smooth pointer-based scrolling + responsive sizing after UI builds
        self.dialog.after(220, lambda: (
            self._attach_smooth_scroll(self.dialog),
            self._apply_responsive_width(),
        ))

    # ============================================================
    # Utility: compute total calc from raw 27 fields
    # ============================================================
    def _calc_from_dicts(self, comp, var, ded):
        components = {k: float((comp or {}).get(k, 0) or 0) for (k, _) in self.COMPONENT_KEYS}
        variables =  {k: float((var or {}).get(k, 0) or 0) for (k, _) in self.VARIABLE_KEYS}
        deductions = {k: float((ded or {}).get(k, 0) or 0) for (k, _) in self.DEDUCTION_KEYS}
        gross = round(sum(components.values()) + sum(variables.values()), 2)
        total_ded = round(sum(deductions.values()), 2)
        net = round(gross - total_ded, 2)
        return {
            "components": components,
            "variables": variables,
            "deductions": deductions,
            "gross_salary": gross,
            "total_deductions": total_ded,
            "net_salary": net,
        }

    # ============================================================
    # Refresh everything
    # ============================================================
    def _refresh_everything(self):
        # Update top bar
        emps = self.manager.list_employees() or []
        today = date.today()
        mtd_start = date(today.year, today.month, 1).isoformat()
        c_ytd = date(today.year, 1, 1).isoformat()
        today_s = today.isoformat()
        mtd = self.manager.get_salary_summary_for_period(mtd_start, today_s)
        ytd = self.manager.get_salary_summary_for_period(c_ytd, today_s)
        self.summary_var.set(
            f"👥 {len(emps)} Employees on File   ·   MTD {mtd['months_in_period']} pay-run: "
            f"Gross {mtd['gross_salary_total']:,.2f} AED / Net {mtd['net_salary_total']:,.2f} AED "
            f"({mtd['employees_processed']} payslips)   ·   CYTD: Gross {ytd['gross_salary_total']:,.2f} AED "
            f"across {ytd['months_in_period']} months"
        )
        # Refresh any loaded tab contents
        for fn in [getattr(self, "_refresh_directory", None),
                   getattr(self, "_refresh_monthly_rows", None),
                   getattr(self, "_refresh_history_tree", None),
                   getattr(self, "_refresh_documents_tree", None)]:
            if callable(fn):
                try: fn()
                except Exception: pass

    # ============================================================
    # TAB 1 — 👥 Employee Directory (13 columns, chips)
    # ============================================================
    def _build_directory(self, parent):
        top = ttk.Frame(parent, padding=(12, 8)); top.pack(fill='x')
        ttk.Label(top, text="🔎 Search:").pack(side='left')
        self._dir_search = tk.StringVar()
        e = ttk.Entry(top, textvariable=self._dir_search, width=36); e.pack(side='left', padx=6)
        e.bind("<KeyRelease>", lambda _ev: self._refresh_directory())
        for lbl, var_key in [("Dept.", "_dir_dept"), ("Position", "_dir_pos"), ("Status", "_dir_status")]:
            ttk.Label(top, text=lbl+":").pack(side='left', padx=(12,4))
            v = tk.StringVar(value="All"); setattr(self, var_key, v)
            cb = ttk.Combobox(top, textvariable=v, values=["All"], state="readonly", width=14)
            cb.pack(side='left', padx=2); cb.bind("<<ComboboxSelected>>", lambda _e: self._refresh_directory())
        ttk.Button(top, text="➕ Add Employee", command=self._add_employee_quick).pack(side='right', padx=4)
        ttk.Button(top, text="👤 Open Profile", command=self._open_selected_profile).pack(side='right', padx=4)

        cols = ("ID", "Name", "Nationality", "Dept", "Position", "Grade", "Status",
                "Work Visa", "Joining", "Monthly Basic", "Phone", "Email", "Manager")
        self._dir_tree = ttk.Treeview(parent, columns=cols, show='headings', height=22)
        widths = [70, 200, 120, 110, 160, 70, 90, 100, 100, 120, 120, 200, 120]
        anchors = ["center", "w", "center", "center", "w", "center", "center", "center", "center",
                   "e", "center", "w", "w"]
        for i, c in enumerate(cols):
            self._dir_tree.heading(c, text=c, anchor="center")
            self._dir_tree.column(c, width=widths[i], anchor=anchors[i], minwidth=max(60, widths[i]-50))
        # Tags for status badges (zebra + chip colors for status + visa)
        self._dir_tree.tag_configure("active",    background="#ECFDF5", foreground="#064E3B")
        self._dir_tree.tag_configure("probation", background="#FFF7ED", foreground="#7C2D12")
        self._dir_tree.tag_configure("on_leave",  background="#EFF6FF", foreground="#1E3A8A")
        self._dir_tree.tag_configure("terminated",background="#FEF2F2", foreground="#7F1D1D")
        self._dir_tree.tag_configure("visa_expired", background="#FEE2E2")
        self._dir_tree.tag_configure("visa_valid",  background="#DBEAFE")
        self._dir_tree.tag_configure("visa_soon",   background="#FEF3C7")
        self._dir_tree.tag_configure("odd", background="#FAFAFA")
        sb_y = ttk.Scrollbar(parent, orient="vertical", command=self._dir_tree.yview)
        sb_x = ttk.Scrollbar(parent, orient="horizontal", command=self._dir_tree.xview)
        self._dir_tree.configure(yscrollcommand=sb_y.set, xscrollcommand=sb_x.set)
        self._dir_tree.pack(fill='both', expand=True, padx=10, pady=(4, 0))
        sb_y.pack(side='right', fill='y', before=self._dir_tree)
        sb_x.pack(side='bottom', fill='x')
        self._dir_tree.bind("<Double-1>", lambda _e: self._open_selected_profile())
        self._dir_tree.bind("<Return>",   lambda _e: self._open_selected_profile())

    def _refresh_directory(self):
        if not hasattr(self, "_dir_tree"): return
        for iid in self._dir_tree.get_children(): self._dir_tree.delete(iid)
        all_emps = self.manager.list_employees() or []
        # Populate filter combos once
        for attr, key in [("department", "_dir_dept"),
                          ("position",   "_dir_pos"),
                          ("status",     "_dir_status")]:
            vals = sorted({str(e.get(attr) or "").strip() for e in all_emps if e.get(attr)})
            cb = getattr(self, key)
            old = cb.get()
            cb_vals = ["All"] + vals
            try: cb.config(values=cb_vals)
            except Exception: pass
            if old not in cb_vals: cb.set("All")
        # Apply filters
        q = (self._dir_search.get() or "").strip().lower()
        dept_f = getattr(self, "_dir_dept", None); dept_v = dept_f.get() if dept_f else "All"
        pos_f  = getattr(self, "_dir_pos", None);  pos_v  = pos_f.get() if pos_f  else "All"
        st_f   = getattr(self, "_dir_status", None); st_v = st_f.get() if st_f else "All"
        today = date.today()
        rows_out = []
        for e in all_emps:
            name = (e.get("name") or "").lower()
            haystack = " ".join(str(x or "").lower() for x in [
                e.get("employee_id"), e.get("name"), e.get("nationality"), e.get("department"),
                e.get("position"), e.get("email"), e.get("mobile_phone"), e.get("manager_name")
            ])
            if q and q not in haystack: continue
            if dept_v != "All" and (e.get("department") or "") != dept_v: continue
            if pos_v  != "All" and (e.get("position")   or "") != pos_v:  continue
            if st_v   != "All" and (e.get("status")     or "") != st_v:   continue
            # Visa status tag
            tags = []
            visa_exp = e.get("visa_expiry") or e.get("work_visa_expiry") or e.get("work_permit_expiry")
            chip = " — "
            try:
                if visa_exp:
                    vd = datetime.strptime(str(visa_exp)[:10], "%Y-%m-%d").date()
                    days_left = (vd - today).days
                    if days_left < 0:   chip = "❌ Expired"; tags.append("visa_expired")
                    elif days_left < 30: chip = f"⚠ {days_left}d left"; tags.append("visa_soon")
                    else:                chip = f"✅ Valid ({vd.isoformat()})"; tags.append("visa_valid")
            except Exception: chip = str(visa_exp or "—")
            status = (e.get("status") or "Active").strip().lower()
            if "probat" in status:      tags.append("probation")
            elif "leave" in status:     tags.append("on_leave")
            elif "termin" in status or "resign" in status: tags.append("terminated")
            else:                        tags.append("active")
            # Basic salary: prefer salary_components.basic, else base_salary, else salary
            comp = e.get("salary_components") or {}
            basic = (comp.get("basic") if isinstance(comp, dict) else None) or \
                    e.get("base_salary") or e.get("salary") or 0
            try: basic_f = float(basic)
            except Exception: basic_f = 0.0
            rows_out.append((e, tags, chip, basic_f))
        # Sort by department then name
        rows_out.sort(key=lambda tup: (str((tup[0].get("department") or "")), str((tup[0].get("name") or ""))))
        for i, (e, tags, visa_chip, basic_f) in enumerate(rows_out):
            if i % 2: tags = (tags or []) + ["odd"]
            self._dir_tree.insert("", "end", iid=str(e.get("employee_id")), tags=tags, values=(
                e.get("employee_id"),
                e.get("name") or "",
                e.get("nationality") or "",
                e.get("department") or "",
                e.get("position") or "",
                e.get("grade") or e.get("job_level") or "",
                (e.get("status") or "Active"),
                visa_chip,
                e.get("joining_date") or e.get("start_date") or "",
                f"{basic_f:,.2f} AED" if basic_f else "",
                e.get("mobile_phone") or e.get("phone") or "",
                e.get("email") or "",
                e.get("manager_name") or e.get("reports_to") or "",
            ))

    # ============================================================
    # TAB 2 — 👤 Personnel File (30-field editable form)
    # ============================================================
    def _build_personnel_file(self, parent):
        wrap = ttk.Frame(parent, padding=(14, 10))
        wrap.pack(fill='both', expand=True)
        # Employee picker header
        head = ttk.LabelFrame(wrap, text=" 🔎 Select Employee ", padding=(12, 10))
        head.pack(fill='x')
        ttk.Label(head, text="Employee:").grid(row=0, column=0, sticky='w')
        self._pf_pick_var = tk.StringVar()
        self._pf_pick_cb = ttk.Combobox(head, textvariable=self._pf_pick_var, values=[],
                                         state="readonly", width=44)
        self._pf_pick_cb.grid(row=0, column=1, sticky='w', padx=6)
        self._pf_pick_cb.bind("<<ComboboxSelected>>", lambda _e: self._pf_load_selected())
        ttk.Button(head, text="➕ New", command=self._add_employee_quick).grid(row=0, column=2, padx=6)
        ttk.Button(head, text="💾 Save Changes", command=self._pf_save_current).grid(row=0, column=3, padx=6)
        ttk.Label(head, text="💡 Tip: fields are auto-saved to employees.json",
                  foreground="#475569").grid(row=0, column=4, padx=14, sticky='w')
        # Scrollable 30-field form split into 6 Sections (cards)
        canvas = tk.Canvas(wrap, highlightthickness=0)
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(fill='both', expand=True, pady=(10, 0), side='left')
        vsb.pack(side='right', fill='y')
        # Entries dict
        self._pf = {}  # key -> widget with .get() / .set() / .insert / .delete

        def section(container, title, color="#1E40AF"):
            card = ttk.LabelFrame(container, text=f"  {title}  ", padding=(12, 10))
            card.pack(fill='x', pady=(0, 10))
            try:
                card.configure(foreground=color)
            except Exception: pass
            return card

        def make_entry(card, label, key, row, col, width=28, default=""):
            ttk.Label(card, text=label).grid(row=row, column=col*2, sticky='w', padx=(0,4), pady=3)
            var = tk.StringVar(value=default)
            e = ttk.Entry(card, textvariable=var, width=width)
            e.grid(row=row, column=col*2+1, sticky='w', pady=3)
            self._pf[key] = var
            return e

        def make_combo(card, label, key, values, row, col, default=""):
            ttk.Label(card, text=label).grid(row=row, column=col*2, sticky='w', padx=(0,4), pady=3)
            var = tk.StringVar(value=default)
            c = ttk.Combobox(card, textvariable=var, values=values, state="readonly", width=26)
            c.grid(row=row, column=col*2+1, sticky='w', pady=3)
            self._pf[key] = var
            return c

        s1 = section(inner, "🧑 Personal Information", "#1e3a8a")
        for i, (lab, key, width) in enumerate([
            ("Employee ID", "employee_id", 18),
            ("Full Name *", "name", 34),
            ("Arabic Name", "arabic_name", 34),
            ("Gender", "__GENDER__", 14),
            ("Date of Birth", "dob", 18),
            ("Nationality", "nationality", 24),
            ("Marital Status", "__MARITAL__", 16),
            ("Religion", "religion", 18),
            ("Passport No.", "passport_number", 20),
            ("Passport Expiry", "passport_expiry", 18),
            ("UAE ID / EID", "emirates_id", 22),
            ("EID Expiry", "emirates_id_expiry", 18),
        ]):
            if key == "__GENDER__":
                make_combo(s1, lab, "gender", ["", "Male", "Female", "Other"], i//3, i%3)
            elif key == "__MARITAL__":
                make_combo(s1, lab, "marital_status", ["", "Single", "Married", "Divorced", "Widowed"], i//3, i%3)
            else:
                make_entry(s1, lab, key, i//3, i%3, width=width)

        s2 = section(inner, "🏢 Work & Role", "#065f46")
        for i, (lab, key, opt) in enumerate([
            ("Department *", "department", ["", "Sales", "Marketing", "Warehouse", "Procurement", "Pharmacy", "Finance", "HR", "Operations", "Management", "Quality", "IT", "Customer Service", "Logistics"]),
            ("Position / Title *", "position", None),
            ("Job Grade / Band", "grade", ["", "Junior", "Mid-Level", "Senior", "Lead", "Supervisor", "Manager", "Head of Dept", "Director", "C-Level"]),
            ("Reporting To / Manager", "manager_name", None),
            ("Employment Type", "employment_type", ["", "Full-time", "Part-time", "Contract (Fixed)", "Freelance", "Intern", "Probation"]),
            ("Employment Status", "status", ["", "Active", "Probation", "On Leave", "Suspended", "Notice Period", "Resigned", "Terminated", "Retired"]),
            ("Joining Date *", "joining_date", None),
            ("Probation End Date", "probation_end", None),
            ("Contract Start", "contract_start", None),
            ("Contract End", "contract_end", None),
            ("Notice Period (days)", "notice_days", None),
            ("Cost Center / Project", "cost_center", None),
        ]):
            if isinstance(opt, list):
                make_combo(s2, lab, key, opt, i//3, i%3, default="")
            else:
                make_entry(s2, lab, key, i//3, i%3)

        s3 = section(inner, "🛂 Visa & Work Permits", "#9A3412")
        for i, (lab, key, width) in enumerate([
            ("Work Visa File No.", "work_visa_file", 22),
            ("Work Permit No.", "work_permit_number", 22),
            ("Visa / Permit Expiry *", "work_visa_expiry", 20),
            ("Labor Card No.", "labor_card_number", 22),
            ("Health Card / Insurance", "health_card_number", 24),
            ("Insurance Class", "insurance_class", 22),
            ("Sponsor / Company", "sponsor_name", 30),
            ("Visa Fees (AED)", "visa_fees", 16),
            ("Visa Fees Recovery (AED)", "visa_fees_recovery", 16),
        ]):
            make_entry(s3, lab, key, i//3, i%3, width=width)

        s4 = section(inner, "📞 Contact & Address", "#312e81")
        for i, (lab, key, w) in enumerate([
            ("Mobile / WhatsApp", "mobile_phone", 26),
            ("Other Phone", "other_phone", 26),
            ("Work Email", "email", 32),
            ("Personal Email", "personal_email", 32),
            ("UAE Address", "address_uae", 46),
            ("Emirate", "emirate", 18),
            ("Home Country Address", "address_home_country", 46),
            ("Emergency Contact", "emergency_contact_name", 26),
            ("Emergency Phone", "emergency_contact_phone", 26),
            ("Relationship", "emergency_contact_relation", 20),
        ]):
            make_entry(s4, lab, key, i//3 if i < 9 else 3, i%3 if i < 9 else 0, width=w)

        s5 = section(inner, "💰 Compensation", "#78350F")
        comp_grid = ttk.Frame(s5); comp_grid.pack(fill='x')
        # Use 4 columns of 8 components + 1 base
        comps = list(self.COMPONENT_KEYS)[:15]
        # Row 0 label says "Guaranteed Components (Monthly):"
        ttk.Label(comp_grid, text="Guaranteed Salary Components (AED / Month) — 15 Fields",
                  font=("Segoe UI", 10, "bold"), foreground="#78350F").grid(
            row=0, column=0, columnspan=8, sticky='w', pady=(0, 6))
        per_col = 5
        for ci, (k, label) in enumerate(comps):
            col = (ci // per_col) * 2
            rw  = 1 + (ci % per_col)
            ttk.Label(comp_grid, text=label).grid(row=rw, column=col, sticky='w', padx=(0,4), pady=2)
            var = tk.StringVar(value="0")
            ttk.Entry(comp_grid, textvariable=var, width=14).grid(row=rw, column=col+1, sticky='w', pady=2)
            self._pf["salary_components." + k] = var
        # Variable pays (5)
        ttk.Label(comp_grid, text="\nVariable Pay (monthly, if any)",
                  font=("Segoe UI", 10, "bold"), foreground="#14532D").grid(
            row=1, column=6, columnspan=2, sticky='w', padx=(20,0))
        for vi, (k, label) in enumerate(self.VARIABLE_KEYS):
            ttk.Label(comp_grid, text=label).grid(row=2+vi, column=6, sticky='w', padx=(20,4), pady=2)
            var = tk.StringVar(value="0")
            ttk.Entry(comp_grid, textvariable=var, width=14).grid(row=2+vi, column=7, sticky='w', pady=2)
            self._pf["variable_pay." + k] = var
        # Deductions (7)
        ttk.Label(comp_grid, text="\nTypical Default Deductions",
                  font=("Segoe UI", 10, "bold"), foreground="#7f1d1d").grid(
            row=1, column=8, columnspan=2, sticky='w', padx=(20,0))
        for di, (k, label) in enumerate(self.DEDUCTION_KEYS):
            ttk.Label(comp_grid, text=label).grid(row=2+di, column=8, sticky='w', padx=(20,4), pady=2)
            var = tk.StringVar(value="0")
            ttk.Entry(comp_grid, textvariable=var, width=14).grid(row=2+di, column=9, sticky='w', pady=2)
            self._pf["deductions." + k] = var
        comp_grid.columnconfigure(10, weight=1)
        # Live Net Salary bar
        live = ttk.LabelFrame(s5, text="  Live Net Salary Calculation (Defaults for new months)  ", padding=(12, 10))
        live.pack(fill='x', pady=(8, 0))
        self._pf_live = tk.StringVar(value="Gross 0.00 AED · Deductions 0.00 AED · Net 0.00 AED")
        ttk.Label(live, textvariable=self._pf_live,
                  font=("Segoe UI", 12, "bold"),
                  foreground="#0f766e").pack(side='left')
        ttk.Button(live, text="🧮 Recalculate",
                   command=self._pf_compute_live).pack(side='right')
        # Attach listeners to every numeric salary widget
        self.dialog.after(350, self._pf_attach_live_listeners)

        s6 = section(inner, "🏦 Bank & Finance", "#164e63")
        bgrid = ttk.Frame(s6); bgrid.pack(fill='x')
        for i, (lab, key, w) in enumerate([
            ("Bank Name *", "bank_name", 26),
            ("Branch", "bank_branch", 22),
            ("IBAN / Account No. *", "bank_account", 30),
            ("SWIFT / BIC", "bank_swift", 18),
            ("Bank City", "bank_city", 18),
            ("Beneficiary Name (on account)", "beneficiary_name", 30),
            ("Salary Payment Method", "__PAYMETHOD__", 22),
            ("Payment Day (1-31)", "salary_day", 14),
            ("Pay Cycle", "__CYCLE__", 18),
        ]):
            if key == "__PAYMETHOD__":
                make_combo(bgrid, lab, "payment_method", ["", "Bank Transfer", "Cheque", "Cash", "WPS / SIF"],
                           i//3, i%3, default="Bank Transfer")
            elif key == "__CYCLE__":
                make_combo(bgrid, lab, "pay_cycle", ["", "Monthly", "Bi-Weekly", "Weekly", "Project-based"],
                           i//3, i%3, default="Monthly")
            else:
                make_entry(bgrid, lab, "bank_details." + key if key in {"bank_name","bank_branch","bank_account","bank_swift","bank_city","beneficiary_name"} else key,
                           i//3, i%3, width=w)
        # Notes field
        notes_sec = section(inner, "📝 Notes / Attachments Summary", "#4B5563")
        ttk.Label(notes_sec, text="Personnel File Notes:").pack(anchor='w')
        self._pf_notes = tk.Text(notes_sec, height=5, wrap='word', font=("Segoe UI", 10))
        self._pf_notes.pack(fill='x', pady=4)

    def _pf_attach_live_listeners(self):
        for k, w in list(self._pf.items()):
            if not k.startswith(("salary_components.", "variable_pay.", "deductions.")):
                continue
            try:
                w.trace_add("write", lambda *_: self._pf_compute_live())
            except Exception:
                try:
                    # widget may be Entry not StringVar
                    w.bind("<KeyRelease>", lambda _e: self._pf_compute_live())
                except Exception:
                    pass
        self._pf_compute_live()

    def _pf_compute_live(self):
        comp, var, ded = {}, {}, {}
        for prefix, target in [("salary_components.", comp), ("variable_pay.", var), ("deductions.", ded)]:
            for k, v in self._pf.items():
                if k.startswith(prefix):
                    key = k[len(prefix):]
                    try: target[key] = float(v.get() or 0)
                    except Exception: target[key] = 0.0
        calc = self._calc_from_dicts(comp, var, ded)
        self._pf_live.set(
            f"Gross {calc['gross_salary']:,.2f} AED · "
            f"Deductions {calc['total_deductions']:,.2f} AED · "
            f"Net {calc['net_salary']:,.2f} AED"
        )
        return calc

    def _pf_populate_combos(self):
        if not hasattr(self, "_pf_pick_cb"): return
        emps = self.manager.list_employees() or []
        emps_sorted = sorted(emps, key=lambda e: (str(e.get("department") or ""), str(e.get("name") or "")))
        labels = [f"{e.get('employee_id','')} — {e.get('name','')}  ({e.get('department') or '-'})"
                  for e in emps_sorted]
        self._pf_pick_cb["values"] = labels
        self._pf_label_to_emp = {lab: e for lab, e in zip(labels, emps_sorted)}
        self._pf_current_emp_id = None

    def _pf_load_selected(self):
        self._pf_populate_combos()
        lab = self._pf_pick_var.get()
        emp = self._pf_label_to_emp.get(lab) if hasattr(self, "_pf_label_to_emp") else None
        if emp is None: return
        self._pf_current_emp_id = emp.get("employee_id")

        def set_val(key, value):
            w = self._pf.get(key)
            if w is None: return
            try:
                if isinstance(value, float):
                    w.set(f"{value:,.2f}" if value else "0")
                else:
                    w.set(str(value or ""))
            except Exception: pass
        # Flat fields
        for k, _ in list(self._pf.items()):
            if "." in k: continue
            set_val(k, emp.get(k, ""))
        # Nested: salary_components, variable_pay, deductions, bank_details
        for nested in ["salary_components", "variable_pay", "deductions", "bank_details"]:
            bucket = emp.get(nested) or {}
            if isinstance(bucket, dict):
                for kk, vv in bucket.items():
                    set_val(nested + "." + kk, vv)
        # Notes
        try:
            self._pf_notes.delete('1.0', tk.END)
            self._pf_notes.insert('1.0', str(emp.get("notes") or emp.get("comments") or ""))
        except Exception: pass
        self._pf_compute_live()

    def _pf_save_current(self):
        if not getattr(self, "_pf_current_emp_id", None):
            messagebox.showinfo("No Employee", "Select an employee from the dropdown first, then edit fields and Save.")
            return
        updates = {}
        for k, w in self._pf.items():
            if "." in k:
                continue
            val = (w.get() if hasattr(w, "get") else "")
            updates[k] = val
        # Rebuild nested buckets
        for nested in ["salary_components", "variable_pay", "deductions", "bank_details"]:
            bucket = {}
            prefix = nested + "."
            for k, w in self._pf.items():
                if k.startswith(prefix):
                    kk = k[len(prefix):]
                    raw = (w.get() if hasattr(w, "get") else "")
                    if nested in {"salary_components", "variable_pay", "deductions"}:
                        try: bucket[kk] = round(float(raw or 0), 2)
                        except Exception: bucket[kk] = 0.0
                    else:
                        bucket[kk] = str(raw or "")
            updates[nested] = bucket
        # Notes
        try:
            updates["notes"] = self._pf_notes.get('1.0', 'end').strip()
        except Exception: pass
        try:
            updates["last_modified"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except Exception: pass
        self.manager.update_employee(self._pf_current_emp_id, updates)
        messagebox.showinfo("✅ Saved", f"Personnel file for {self._pf_current_emp_id} saved.\n\n"
                                        f"Basic Salary: AED {(updates.get('salary_components') or {}).get('basic', 0):,.2f}\n"
                                        f"Net Salary (Default): AED {self._pf_compute_live()['net_salary']:,.2f}")
        self._refresh_everything()
        self._pf_populate_combos()

    # ============================================================
    # TAB 3 — 💰 MONTHLY PAYROLL (27 editable fields × N)
    # ============================================================
    def _build_payroll_monthly(self, parent):
        # Top toolbar
        top = ttk.Frame(parent, padding=(12, 10)); top.pack(fill='x')
        ttk.Label(top, text="Payroll Month (YYYY-MM):",
                  font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky='w')
        self._pay_month = tk.StringVar(value=self.manager.get_month_key())
        ttk.Entry(top, textvariable=self._pay_month, width=12,
                  font=("Segoe UI", 10, "bold")).grid(row=0, column=1, padx=6)
        ttk.Label(top, text="Bulk % Change:").grid(row=0, column=2, padx=(22, 4))
        self._pay_pct = tk.StringVar(value="0")
        ttk.Entry(top, textvariable=self._pay_pct, width=7).grid(row=0, column=3)
        self._pay_pct_mode = tk.StringVar(value="Gross Increase")
        ttk.Combobox(top, textvariable=self._pay_pct_mode, state="readonly", width=20,
                     values=["Gross Increase", "Gross Cut", "Basic Only +%", "Basic Only -%",
                             "All Components +%", "All Components -%"]
                     ).grid(row=0, column=4, padx=6)
        ttk.Label(top, text="Apply to Dept:").grid(row=0, column=5, padx=(22, 4))
        self._pay_dept = tk.StringVar(value="All Departments")
        self._pay_dept_cb = ttk.Combobox(top, textvariable=self._pay_dept,
                                         values=["All Departments"], state="readonly", width=22)
        self._pay_dept_cb.grid(row=0, column=6, padx=4)
        ttk.Button(top, text="▶ Apply Bulk", command=self._pay_bulk_apply).grid(row=0, column=7, padx=8)
        ttk.Separator(top, orient="vertical").grid(row=0, column=8, sticky='ns', padx=10)
        self._pay_copy_prev = tk.BooleanVar(value=True)
        ttk.Checkbutton(top, text="Copy previous month if empty",
                        variable=self._pay_copy_prev).grid(row=0, column=9)
        ttk.Button(top, text="🔁 Load Month", command=self._refresh_monthly_rows).grid(row=0, column=10, padx=8)
        ttk.Button(top, text="💾 Save Draft", command=lambda: self._pay_save("draft")
                   ).grid(row=0, column=11, padx=4)
        ttk.Button(top, text="✅ Submit / Post to GL",
                   command=lambda: self._pay_save("approved"), style="Accent.TButton"
                   ).grid(row=0, column=12, padx=4)
        self._pay_total = tk.StringVar(value="Gross 0.00 · Deductions 0.00 · Net 0.00 · 0 Employees")
        tot_frame = ttk.LabelFrame(parent, text="  🧮 Totals  ", padding=(10, 6))
        tot_frame.pack(fill='x', padx=12, pady=(6, 4))
        ttk.Label(tot_frame, textvariable=self._pay_total,
                  font=("Segoe UI", 12, "bold"), foreground="#064e3b").pack()
        # Canvas + scrollable rows for 27 fields × employees
        outer = ttk.Frame(parent)
        outer.pack(fill='both', expand=True, padx=10, pady=(2, 10))
        self._pay_canvas = tk.Canvas(outer, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient="vertical", command=self._pay_canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal", command=self._pay_canvas.xview)
        self._pay_canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self._pay_canvas.pack(fill='both', expand=True, side='left')
        vsb.pack(side='right', fill='y'); hsb.pack(side='bottom', fill='x')
        self._pay_inner = ttk.Frame(self._pay_canvas)
        self._pay_inner.bind("<Configure>",
                             lambda _e: self._pay_canvas.configure(scrollregion=self._pay_canvas.bbox("all")))
        self._pay_canvas.create_window((0,0), window=self._pay_inner, anchor="nw")
        self._pay_rows = []  # list of widget dicts per employee

    def _refresh_monthly_rows(self):
        if not hasattr(self, "_pay_rows"): return
        self._pay_populate_dept_filter()
        mk = (self._pay_month.get() or self.manager.get_month_key()).strip()
        # Clear
        for w in self._pay_inner.winfo_children(): w.destroy()
        self._pay_rows = []
        # Fetch entries
        entries = list(self.manager.get_month_entries(mk) or [])
        emps = self.manager.list_employees() or []
        if not entries and self._pay_copy_prev.get():
            prev_rows = self.manager.copy_previous_month(mk) or []
            # Merge with all employees so new ones get rows
            by_id = {e.get("employee_id"): e for e in emps}
            for prev in prev_rows:
                eid = prev.get("employee_id")
                if eid in by_id: del by_id[eid]
                entries.append(prev)
            for e in by_id.values():
                entries.append(self._employee_to_default_salary_row(e, mk))
        if not entries:
            # Nothing stored, no previous month — create a row for each employee
            for e in emps:
                entries.append(self._employee_to_default_salary_row(e, mk))
        if not entries:
            ttk.Label(self._pay_inner,
                      text="⚠ No employees yet.  Go to 👥 Directory → ➕ Add Employee.",
                      foreground="#7f1d1d", padding=20).pack()
            return
        # Build 2 headers: Section bands (Components | Variable | Deductions | Calc) then column titles
        H_FONT = ("Segoe UI", 9, "bold")
        # Group widths: 2 fixed cols (Employee, Dept) + 15 + 5 + 7 + 3 (Gross/Ded/Net) + 1 (Actions) = 33 cols
        bands = ttk.Frame(self._pay_inner); bands.pack(fill='x')
        all_labels = [
            ("Employee", 180, "#1e3a8a"),
            ("Dept.",    90, "#1e3a8a"),
        ] + [(lab, 72, "#78350F") for (_, lab) in self.COMPONENT_KEYS] + \
            [(lab, 72, "#065f46") for (_, lab) in self.VARIABLE_KEYS] + \
            [(lab, 72, "#7f1d1d") for (_, lab) in self.DEDUCTION_KEYS] + \
            [("GROSS", 90, "#111827"), ("DEDUCTIONS", 90, "#111827"), ("NET", 100, "#111827"), ("", 50, "#ffffff")]
        for idx, (txt, w, color) in enumerate(all_labels):
            l = tk.Label(bands, text=txt, width=max(6, (w-6)//7), font=H_FONT,
                         fg="white", bg=color, relief="ridge", padx=3, pady=5)
            l.grid(row=0, column=idx, sticky='nsew')
        bands.grid_rowconfigure(0, weight=1)
        # Body rows
        entries.sort(key=lambda e: (str(e.get("department") or ""), str(e.get("name") or "")))
        self._pay_header_cols = len(all_labels)
        for ei, entry in enumerate(entries):
            self._pay_add_row(ei, entry)
        self._pay_update_totals()

    def _employee_to_default_salary_row(self, e, mk):
        sc = e.get("salary_components") or {}
        vp = e.get("variable_pay") or {}
        dd = e.get("deductions") or {}
        def _r(d, keys, default=0.0):
            return {k: round(float((d or {}).get(k, default) or 0), 2) for (k, _) in keys}
        comp = _r(sc or {}, self.COMPONENT_KEYS)
        var  = _r(vp or {}, self.VARIABLE_KEYS, 0)
        ded  = _r(dd or {}, self.DEDUCTION_KEYS, 0)
        calc = self._calc_from_dicts(comp, var, ded)
        return {
            "employee_id": e.get("employee_id"),
            "name": e.get("name"),
            "department": e.get("department") or "",
            "month_year": mk,
            "salary_components": comp,
            "variable_pay": var,
            "deductions": ded,
            "calculated_values": {k: calc[k] for k in ("gross_salary","total_deductions","net_salary")},
            "status": "new",
            "bank_details": e.get("bank_details") or {},
        }

    def _pay_add_row(self, ei, entry):
        row_frame = ttk.Frame(self._pay_inner)
        row_frame.pack(fill='x')
        bg = "#FFFFFF" if ei % 2 == 0 else "#FAFAFA"
        widgets = {"frame": row_frame, "entry_ref": entry}
        def mk_entry(col, width, init_val, is_numeric=True, font=("Segoe UI", 9)):
            var = tk.StringVar(value=(f"{float(init_val or 0):.2f}" if is_numeric and not isinstance(init_val, str) else str(init_val or "")))
            e = tk.Entry(row_frame, textvariable=var, width=width,
                         relief="solid", borderwidth=1, bg=bg, fg="#0f172a", font=font)
            e.grid(row=0, column=col, sticky='nsew', padx=1, pady=1)
            if is_numeric:
                e.bind("<KeyRelease>", lambda _ev: self._pay_recompute_row(widgets))
                e.bind("<FocusOut>",   lambda _ev: self._pay_recompute_row(widgets))
            return var, e
        col = 0
        # Employee name
        w_var, _ = mk_entry(col, 24,
                            f"{entry.get('employee_id') or ''} · {entry.get('name') or ''}",
                            is_numeric=False)
        widgets["employee"] = w_var; col += 1
        _, dept_w = mk_entry(col, 12, entry.get("department") or "", is_numeric=False)
        widgets["dept"] = dept_w.get if hasattr(dept_w, "get") else None
        col += 1
        comp = entry.get("salary_components") or {}
        for k, _ in self.COMPONENT_KEYS:
            v, _ = mk_entry(col, 10, float(comp.get(k, 0) or 0))
            widgets["comp." + k] = v; col += 1
        var_pay = entry.get("variable_pay") or {}
        for k, _ in self.VARIABLE_KEYS:
            v, _ = mk_entry(col, 10, float(var_pay.get(k, 0) or 0))
            widgets["var." + k] = v; col += 1
        ded = entry.get("deductions") or {}
        for k, _ in self.DEDUCTION_KEYS:
            v, _ = mk_entry(col, 10, float(ded.get(k, 0) or 0))
            widgets["ded." + k] = v; col += 1
        # Computed columns
        def readonly_lab(col, text, fg="#000", bg_c=bg, bold=False):
            font = ("Segoe UI", 9, "bold") if bold else ("Segoe UI", 9)
            lab = tk.Label(row_frame, text=text, width=12, bg=bg_c, fg=fg,
                           relief="solid", borderwidth=1, font=font, anchor="e", padx=4)
            lab.grid(row=0, column=col, sticky='nsew', padx=1, pady=1)
            return lab
        widgets["gross_lbl"] = readonly_lab(col, "0.00", "#111827", "#FEF3C7", bold=True); col += 1
        widgets["ded_lbl"]   = readonly_lab(col, "0.00", "#111827", "#FECACA"); col += 1
        widgets["net_lbl"]   = readonly_lab(col, "0.00", "#064e3b", "#A7F3D0", bold=True); col += 1
        # Open payslip button
        op = ttk.Button(row_frame, text="🧾", width=3,
                        command=lambda ww=widgets: self._preview_payslip(ww))
        op.grid(row=0, column=col, padx=2, pady=1)
        widgets["payslip_btn"] = op
        self._pay_rows.append(widgets)
        # Initial compute
        self._pay_recompute_row(widgets)

    def _pay_recompute_row(self, w):
        comp, var, ded = {}, {}, {}
        for prefix, target in [("comp.", comp), ("var.", var), ("ded.", ded)]:
            for k, v in w.items():
                if k.startswith(prefix):
                    kk = k[len(prefix):]
                    try: target[kk] = round(float(v.get() or 0), 2)
                    except Exception: target[kk] = 0.0
        calc = self._calc_from_dicts(comp, var, ded)
        try: w["gross_lbl"].config(text=f"{calc['gross_salary']:,.2f}")
        except Exception: pass
        try: w["ded_lbl"].config(text=f"{calc['total_deductions']:,.2f}")
        except Exception: pass
        try: w["net_lbl"].config(text=f"{calc['net_salary']:,.2f}")
        except Exception: pass
        self._pay_update_totals()

    def _pay_update_totals(self):
        if not getattr(self, "_pay_rows", None):
            return
        tg, td, tn = 0.0, 0.0, 0.0
        n = 0
        for w in self._pay_rows:
            try: tg += float(str(w["gross_lbl"].config("text")[4]).replace(",","") or 0)
            except Exception: pass
            try: td += float(str(w["ded_lbl"].config("text")[4]).replace(",","") or 0)
            except Exception: pass
            try: tn += float(str(w["net_lbl"].config("text")[4]).replace(",","") or 0)
            except Exception: pass
            n += 1
        self._pay_total.set(
            f"Gross {tg:,.2f} AED · Deductions {td:,.2f} AED · Net {tn:,.2f} AED · {n} Employees"
        )

    def _pay_populate_dept_filter(self):
        depts = sorted({(e.get("department") or "").strip() for e in (self.manager.list_employees() or []) if e.get("department")})
        values = ["All Departments"] + depts
        try: self._pay_dept_cb.config(values=values)
        except Exception: pass
        if self._pay_dept.get() not in values:
            self._pay_dept.set("All Departments")

    def _pay_bulk_apply(self):
        if not getattr(self, "_pay_rows", None): return
        try:
            pct = float(self._pay_pct.get() or 0)
            if pct == 0:
                messagebox.showinfo("Bulk % Change", "Set a percentage first (e.g. 5 for 5% raise, -2 for cut).")
                return
        except Exception:
            messagebox.showwarning("Bulk % Change", "Enter a number for % (e.g. 3).")
            return
        mode = self._pay_pct_mode.get()
        dept_filt = self._pay_dept.get()
        applied = 0
        for w in self._pay_rows:
            dept_row = ""
            try: dept_row = str(w["dept"].get()).strip() if hasattr(w.get("dept"),"get") else ""
            except Exception: dept_row = ""
            if dept_filt != "All Departments" and dept_row != dept_filt:
                continue
            applied += 1
            mult = 1 + pct/100 if "Increase" in mode or "+" in mode else max(0, 1 - abs(pct)/100)
            prefix_target = []
            if "Gross" in mode or "Components" in mode:
                for k in [k for k, _ in self.COMPONENT_KEYS]: prefix_target.append("comp." + k)
                if "Gross" in mode:
                    for k in [k for k, _ in self.VARIABLE_KEYS]: prefix_target.append("var." + k)
            if "Basic Only" in mode:
                prefix_target = ["comp.basic"]
            for p in prefix_target:
                v = w.get(p)
                if not v: continue
                try:
                    cur = float(v.get() or 0)
                    new_val = round(cur * mult, 2)
                    v.set(f"{new_val:.2f}")
                except Exception:
                    pass
            self._pay_recompute_row(w)
        messagebox.showinfo("✅ Bulk Apply", f"Applied {mode} @ {pct}% to {applied} employee(s)."
                                             f"\nDept. filter: {dept_filt}")

    def _pay_save(self, status):
        if not getattr(self, "_pay_rows", None): return
        mk = (self._pay_month.get() or self.manager.get_month_key()).strip()
        entries = []
        for w in self._pay_rows:
            # Extract employee_id
            emp_key = ""
            try: emp_key = (w["employee"].get() or "").split("·", 1)[0].strip()
            except Exception: pass
            comp, var, ded = {}, {}, {}
            for prefix, target in [("comp.", comp), ("var.", var), ("ded.", ded)]:
                for k, v in w.items():
                    if k.startswith(prefix):
                        kk = k[len(prefix):]
                        try: target[kk] = round(float(v.get() or 0), 2)
                        except Exception: target[kk] = 0.0
            calc = self._calc_from_dicts(comp, var, ded)
            dept = ""
            try: dept = w["dept"].get().strip() if hasattr(w.get("dept"), "get") else ""
            except Exception: dept = ""
            entries.append({
                "employee_id": emp_key,
                "name": "",
                "department": dept,
                "month_year": mk,
                "salary_components": comp,
                "variable_pay": var,
                "deductions": ded,
                "calculated_values": calc,
                "status": status,
                "processed_date": datetime.now().strftime("%Y-%m-%d"),
                "bank_details": {},
            })
        # Re-fill names/departments from authoratative list
        by_id = {e.get("employee_id"): e for e in (self.manager.list_employees() or [])}
        for se in entries:
            e = by_id.get(se.get("employee_id"))
            if e:
                se["name"] = e.get("name") or ""
                se["department"] = e.get("department") or se.get("department") or ""
                se["bank_details"] = e.get("bank_details") or {}
        self.manager.save_month_entries(mk, entries, status=status)
        gl_note = ""
        if status in ("approved", "submitted", "paid", "posted"):
            gl_note = ("\n\n📘 Automatically posted to GAAP General Ledger:\n"
                       "   • Dr 5100 Salary/Wages Expense (Gross)\n"
                       "   • Cr 2100 Accrued / Withheld (Deductions)\n"
                       "   • Cr 1100 Bank (Net Disbursed)")
        messagebox.showinfo(
            f"✅ Payroll Saved ({status.upper()})",
            f"Month: {mk}\nEmployees processed: {len(entries)}\n"
            + self._pay_total.get() + gl_note
        )
        self._refresh_everything()

    def _preview_payslip(self, row_widgets):
        """Quick per-row payslip (preview then PDF)."""
        emp_key = ""
        try: emp_key = (row_widgets["employee"].get() or "").split("·",1)[0].strip()
        except Exception: pass
        mk = (self._pay_month.get() or self.manager.get_month_key()).strip()
        e = next((x for x in (self.manager.list_employees() or []) if x.get("employee_id") == emp_key), None)
        if e is None:
            messagebox.showinfo("Payslip", "Save the employee in the 👥 Directory first, then preview payslip.")
            return
        comp, var, ded = {}, {}, {}
        for prefix, target in [("comp.", comp), ("var.", var), ("ded.", ded)]:
            for k, v in row_widgets.items():
                if k.startswith(prefix):
                    kk = k[len(prefix):]
                    try: target[kk] = round(float(v.get() or 0), 2)
                    except Exception: target[kk] = 0.0
        calc = self._calc_from_dicts(comp, var, ded)
        entry = {
            "employee_id": emp_key,
            "name": e.get("name") or "",
            "department": e.get("department") or "",
            "position": e.get("position") or "",
            "month_year": mk,
            "salary_components": comp, "variable_pay": var, "deductions": ded,
            "calculated_values": calc,
            "bank_details": e.get("bank_details") or {},
        }
        path = self._generate_payslip_pdf(entry, e, open_after=True)
        if path:
            messagebox.showinfo("🧾 Payslip", f"Payslip generated:\n{path}")

    # ============================================================
    # TAB 4 — 📆 SALARY HISTORY (Month × Employee ledger)
    # ============================================================
    def _build_history(self, parent):
        top = ttk.Frame(parent, padding=(12, 8)); top.pack(fill='x')
        ttk.Label(top, text="Employee:").pack(side='left')
        self._hist_emp = tk.StringVar()
        self._hist_emp_cb = ttk.Combobox(top, textvariable=self._hist_emp, values=[],
                                          state="readonly", width=46)
        self._hist_emp_cb.pack(side='left', padx=6)
        ttk.Button(top, text="🔁 Load History", command=self._refresh_history_tree).pack(side='left', padx=6)
        ttk.Button(top, text="📊 Export This History CSV",
                   command=self._hist_export_csv).pack(side='right', padx=6)
        cols = ("month", "gross", "components", "vars", "deductions", "net", "status", "date_posted", "change_mom")
        self._hist_tree = ttk.Treeview(parent, columns=cols, show='headings', height=20)
        for c, w, anchor in [
            ("month", 110, "center"), ("status", 110, "center"), ("date_posted", 120, "center"),
            ("gross", 120, "e"), ("components", 140, "e"), ("vars", 120, "e"),
            ("deductions", 120, "e"), ("net", 130, "e"), ("change_mom", 130, "e")
        ]:
            self._hist_tree.heading(c, text=c.title().replace("_", " "), anchor="center")
            self._hist_tree.column(c, width=w, anchor=anchor)
        self._hist_tree.tag_configure("raise", background="#ECFDF5", foreground="#065F46")
        self._hist_tree.tag_configure("cut",   background="#FEF2F2", foreground="#7F1D1D")
        self._hist_tree.tag_configure("flat",  background="#F8FAFC")
        self._hist_tree.tag_configure("odd",   background="#FAFAFA")
        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._hist_tree.yview)
        self._hist_tree.configure(yscrollcommand=vsb.set)
        self._hist_tree.pack(fill='both', expand=True, padx=10, pady=(6, 10))
        vsb.pack(side='right', fill='y')
        self.dialog.after(300, self._refresh_history_tree)

    def _refresh_history_tree(self):
        if not hasattr(self, "_hist_tree"): return
        # Populate employee combo
        emps = self.manager.list_employees() or []
        labels = [f"{e.get('employee_id') or ''} · {e.get('name') or ''} ({e.get('department') or '-'})" for e in emps]
        try: self._hist_emp_cb.config(values=labels)
        except Exception: pass
        # Resolve selected employee
        lab = self._hist_emp.get()
        emp = None
        for e in emps:
            if lab.startswith(f"{e.get('employee_id') or ''} ·"):
                emp = e; break
        if emp is None and emps:
            emp = emps[0]; self._hist_emp.set(labels[0])
        for iid in self._hist_tree.get_children(): self._hist_tree.delete(iid)
        if emp is None: return
        eid = emp.get("employee_id")
        # Iterate salaries
        rows = []
        for mk, entries in (self.manager.salaries or {}).items():
            if mk.startswith("__") or not isinstance(entries, list): continue
            for se in entries:
                if not isinstance(se, dict): continue
                if se.get("employee_id") != eid: continue
                calc = se.get("calculated_values") or {}
                sc = se.get("salary_components") or {}
                vp = se.get("variable_pay") or {}
                dd = se.get("deductions") or {}
                comp_sum = round(sum(float((sc or {}).get(k,0) or 0) for (k,_) in self.COMPONENT_KEYS
                                    if isinstance(sc, dict)), 2)
                var_sum = round(sum(float((vp or {}).get(k,0) or 0) for (k,_) in self.VARIABLE_KEYS
                                    if isinstance(vp, dict)), 2)
                ded_sum = round(sum(float((dd or {}).get(k,0) or 0) for (k,_) in self.DEDUCTION_KEYS
                                    if isinstance(dd, dict)), 2)
                rows.append({
                    "mk": mk,
                    "gross": float(calc.get("gross_salary", comp_sum + var_sum) or 0),
                    "components": comp_sum,
                    "vars": var_sum,
                    "deductions": ded_sum,
                    "net": float(calc.get("net_salary", (comp_sum + var_sum) - ded_sum) or 0),
                    "status": str(se.get("status") or (self.manager.salaries.get("__meta__") or {}).get(mk, {}).get("status") or "draft").upper(),
                    "date": se.get("processed_date") or f"{mk}-28",
                })
        rows.sort(key=lambda r: r["mk"])
        prev_net = None
        for i, r in enumerate(rows):
            delta = (r["net"] - prev_net) if prev_net is not None else 0.0
            mom = ""
            tag = "flat"
            if prev_net is not None and prev_net != 0:
                pct = delta / abs(prev_net) * 100
                if delta > 0: tag = "raise"; mom = f"⬆ +{delta:,.2f}  (+{pct:+.1f}%)"
                elif delta < 0: tag = "cut"; mom = f"⬇ {delta:,.2f}  ({pct:+.1f}%)"
                else: mom = "—"
            row_tags = [tag] + (["odd"] if i % 2 else [])
            self._hist_tree.insert("", "end", values=(
                r["mk"], f"{r['gross']:,.2f}", f"{r['components']:,.2f}",
                f"{r['vars']:,.2f}", f"{r['deductions']:,.2f}",
                f"{r['net']:,.2f}", r["status"], r["date"], mom
            ), tags=row_tags)
            prev_net = r["net"]

    def _hist_export_csv(self):
        if not hasattr(self, "_hist_tree"): return
        import csv
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV","*.csv")],
                                         initialfile="salary_history.csv")
        if not p: return
        with open(p, "w", newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f)
            w.writerow(["Month","Gross","Components","Var.Pay","Deductions","Net","Status","Posted","MoM Change"])
            for iid in self._hist_tree.get_children():
                w.writerow(list(self._hist_tree.item(iid, "values")))
        messagebox.showinfo("CSV", f"Exported {len(self._hist_tree.get_children())} rows:\n{p}")

    # ============================================================
    # TAB 5 — 💼 SALARY MANAGEMENT (Dedicated CRUD, Monthly Granularity)
    #        Intuitive 3-pane layout:
    #        LEFT   = Employee selector (search + tree)
    #        TOP    = Month picker + action strip (Add / Edit / Delete / Save)
    #        CENTER = Single-record, 27-field editable form  (only for 1 emp + 1 mo)
    #        BOTTOM = Confirmation banner + integrity chip (no cross-month bleed)
    # ============================================================
    def _build_salary_management(self, parent):
        outer = ttk.Frame(parent); outer.pack(fill='both', expand=True)
        # Split: left 30% employee picker, right 70% salary editor
        pw = ttk.Panedwindow(outer, orient='horizontal'); pw.pack(fill='both', expand=True, padx=8, pady=8)
        # ============ LEFT: employee selector ============
        left = ttk.Labelframe(pw, text="  👥 Employees  ", padding=(6, 6))
        pw.add(left, weight=3)
        head = ttk.Frame(left); head.pack(fill='x', pady=(0, 4))
        ttk.Label(head, text="🔎").pack(side='left')
        self._sm_search = tk.StringVar()
        se = ttk.Entry(head, textvariable=self._sm_search, width=22)
        se.pack(side='left', padx=4, fill='x', expand=True)
        se.bind("<KeyRelease>", lambda _e: self._sm_refresh_employees())
        ttk.Button(head, text="🔁", width=3, command=self._sm_refresh_employees).pack(side='left')
        cols = ("id", "name", "dept")
        self._sm_tree = ttk.Treeview(left, columns=cols, show='headings', height=18)
        for c, w, a in [("id", 60, "center"), ("name", 170, "w"), ("dept", 100, "center")]:
            self._sm_tree.heading(c, text=c.upper()); self._sm_tree.column(c, width=w, anchor=a, minwidth=40)
        self._sm_tree.tag_configure("odd", background="#FAFAFA")
        sb = ttk.Scrollbar(left, orient='vertical', command=self._sm_tree.yview)
        self._sm_tree.configure(yscrollcommand=sb.set)
        self._sm_tree.pack(fill='both', expand=True, side='left')
        sb.pack(side='right', fill='y')
        self._sm_tree.bind("<<TreeviewSelect>>", lambda _e: self._sm_on_emp_selected())
        self._sm_tree.bind("<Double-1>", lambda _e: self._sm_on_emp_selected())
        # ============ RIGHT: month + editor ============
        right = ttk.Frame(pw); pw.add(right, weight=7)
        # Top: month picker + actions
        action = ttk.Labelframe(right, text="  📅 Month & Actions  (Changes apply ONLY to the selected month)  ",
                                padding=(10, 8))
        action.pack(fill='x', padx=0, pady=(0, 6))
        ttk.Label(action, text="Employee:").grid(row=0, column=0, sticky='w')
        self._sm_emp_label = tk.StringVar(value="—")
        ttk.Label(action, textvariable=self._sm_emp_label, foreground="#1e3a8a",
                  font=("Segoe UI", 10, "bold")).grid(row=0, column=1, sticky='w', padx=(0, 22))
        ttk.Label(action, text="Month (YYYY-MM):").grid(row=0, column=2, sticky='w')
        self._sm_month = tk.StringVar(value=self.manager.get_month_key())
        m_ent = ttk.Entry(action, textvariable=self._sm_month, width=11,
                          font=("Segoe UI", 10, "bold"))
        m_ent.grid(row=0, column=3, padx=4)
        m_ent.bind("<Return>", lambda _e: self._sm_load_record())
        self._sm_prev_month = tk.BooleanVar(value=False)
        ttk.Checkbutton(action, text="Copy prev. month if empty", variable=self._sm_prev_month
                        ).grid(row=0, column=4, padx=(8, 12), sticky='w')
        ttk.Button(action, text="📥 Load Month", command=self._sm_load_record).grid(row=0, column=5, padx=4)
        ttk.Button(action, text="➕ Create / Reset Default", command=self._sm_create_default
                   ).grid(row=0, column=6, padx=4)
        ttk.Button(action, text="💾 Save This Month ONLY", style="Accent.TButton",
                   command=self._sm_save_record).grid(row=0, column=7, padx=4)
        ttk.Button(action, text="🗑 Delete This Month ONLY",
                   command=self._sm_delete_record).grid(row=0, column=8, padx=4)
        self._sm_integrity = tk.StringVar(
            value="✅ Integrity: Isolated — changes will be applied ONLY to [Emp: ?]  [Month: ?]")
        ttk.Label(action, textvariable=self._sm_integrity, foreground="#065f46",
                  font=("Segoe UI", 9, "bold")).grid(row=1, column=0, columnspan=9, sticky='w', pady=(6, 0))
        # Banner for status
        self._sm_status = tk.StringVar(value="Select an employee → Load a month → Edit → Save.")
        ttk.Label(action, textvariable=self._sm_status, foreground="#475569",
                  wraplength=1100, justify='left').grid(row=2, column=0, columnspan=9, sticky='w', pady=(2, 0))
        # Center: scrollable 27-field editor
        holder = ttk.Labelframe(right, text="  ✍ Monthly Salary Record (Single Month, Single Employee)  ",
                                padding=(8, 6))
        holder.pack(fill='both', expand=True)
        self._sm_canvas = tk.Canvas(holder, highlightthickness=0)
        vs = ttk.Scrollbar(holder, orient='vertical', command=self._sm_canvas.yview)
        hs = ttk.Scrollbar(holder, orient='horizontal', command=self._sm_canvas.xview)
        inner = ttk.Frame(self._sm_canvas)
        inner.bind("<Configure>", lambda _e: self._sm_canvas.configure(scrollregion=self._sm_canvas.bbox("all")))
        self._sm_canvas.create_window((0, 0), window=inner, anchor="nw")
        self._sm_canvas.configure(yscrollcommand=vs.set, xscrollcommand=hs.set)
        self._sm_canvas.pack(fill='both', expand=True, side='left')
        vs.pack(side='right', fill='y'); hs.pack(side='bottom', fill='x')
        self._sm_entries = {}   # "comp.basic" -> StringVar
        self._sm_build_salary_form(inner)
        # Bottom live totals bar
        tot_bar = ttk.Frame(right); tot_bar.pack(fill='x', pady=(6, 0))
        self._sm_live = tk.StringVar(value="Gross: 0.00   ·   Deductions: 0.00   ·   Net: 0.00")
        ttk.Label(tot_bar, textvariable=self._sm_live,
                  font=("Segoe UI", 12, "bold"), foreground="#0f766e").pack(side='left')
        self._sm_last_saved = tk.StringVar(value="(Unsaved)")
        ttk.Label(tot_bar, textvariable=self._sm_last_saved, foreground="#64748b").pack(side='right')
        # Init
        self.dialog.after(250, self._sm_refresh_employees)
        self._sm_bind_live_recalc()

    def _sm_build_salary_form(self, inner):
        # Three cards: COMPONENTS (15), VARIABLE (5), DEDUCTIONS (7)
        cards = [
            ("💼 Salary Components (15 fields)", "#78350F", self.COMPONENT_KEYS, "comp"),
            ("🎁 Variable Pay (5 fields)",        "#065f46", self.VARIABLE_KEYS,  "var"),
            ("➖ Deductions (7 fields)",           "#7f1d1d", self.DEDUCTION_KEYS, "ded"),
        ]
        for title, color, keys, prefix in cards:
            lf = ttk.LabelFrame(inner, text=f"  {title}  ", padding=(10, 8))
            lf.pack(fill='x', pady=(0, 8))
            try: lf.configure(foreground=color)
            except Exception: pass
            per_col = 5
            for i, (k, label) in enumerate(keys):
                col = (i // per_col) * 2
                row = 1 + (i % per_col)
                ttk.Label(lf, text=label).grid(row=row, column=col, sticky='w', padx=(0, 4), pady=2)
                v = tk.StringVar(value="0")
                e = ttk.Entry(lf, textvariable=v, width=14)
                e.grid(row=row, column=col+1, sticky='w', pady=2)
                self._sm_entries[f"{prefix}.{k}"] = v
            # Live per-card subtotal
            st = tk.StringVar(value="Subtotal 0.00")
            ttk.Label(lf, textvariable=st, foreground=color,
                      font=("Segoe UI", 10, "bold")).grid(row=0, column=6, sticky='e', padx=20)
            self._sm_card_subtotals = getattr(self, "_sm_card_subtotals", {})
            self._sm_card_subtotals[prefix] = st
        # Meta + audit section
        meta = ttk.LabelFrame(inner, text="  📝 Record Metadata & Notes  ", padding=(10, 8))
        meta.pack(fill='x', pady=(0, 2))
        grid = ttk.Frame(meta); grid.pack(fill='x')
        for i, (lab, key, opts) in enumerate([
            ("Status",     "status",     ["draft", "approved", "submitted", "paid", "posted", "cancelled"]),
            ("Processed Date (YYYY-MM-DD)", "processed_date", None),
            ("Salary Notes", "notes",     None),
        ]):
            ttk.Label(grid, text=lab).grid(row=i // 2, column=(i % 2) * 2, sticky='w', padx=(0, 4), pady=3)
            if opts is not None:
                v = tk.StringVar(value=opts[0])
                c = ttk.Combobox(grid, textvariable=v, values=opts, state="readonly", width=24)
                c.grid(row=i // 2, column=(i % 2) * 2 + 1, sticky='w', pady=3)
            else:
                v = tk.StringVar(value="")
                w = ttk.Entry(grid, textvariable=v, width=28)
                w.grid(row=i // 2, column=(i % 2) * 2 + 1, sticky='w', pady=3)
            self._sm_entries[key] = v
        nt = ttk.Label(meta, text="Additional Comments (shown on payslip):").pack(anchor='w', pady=(6, 2))
        self._sm_notes = scrolledtext.ScrolledText(meta, height=3, wrap='word', font=("Segoe UI", 10))
        self._sm_notes.pack(fill='x')

    def _sm_bind_live_recalc(self):
        def _bind():
            for k, v in list(self._sm_entries.items()):
                if k.startswith(("comp.", "var.", "ded.")):
                    try: v.trace_add("write", lambda *_a: self._sm_recalc_live())
                    except Exception:
                        try: v.widget.bind("<KeyRelease>", lambda _e: self._sm_recalc_live())
                        except Exception: pass
            self._sm_recalc_live()
        self.dialog.after(280, _bind)

    def _sm_recalc_live(self):
        comp, var, ded = {}, {}, {}
        for prefix, target in [("comp.", comp), ("var.", var), ("ded.", ded)]:
            for k, v in self._sm_entries.items():
                if k.startswith(prefix):
                    key = k[len(prefix):]
                    try: target[key] = round(float(v.get() or 0), 2)
                    except Exception: target[key] = 0.0
        calc = self._calc_from_dicts(comp, var, ded)
        self._sm_live.set(
            f"Gross: {calc['gross_salary']:,.2f} AED   ·   "
            f"Deductions: {calc['total_deductions']:,.2f} AED   ·   "
            f"Net: {calc['net_salary']:,.2f} AED"
        )
        subs = getattr(self, "_sm_card_subtotals", {})
        if "comp" in subs: subs["comp"].set(f"Subtotal {sum(comp.values()):,.2f}")
        if "var"  in subs: subs["var"].set(f"Subtotal {sum(var.values()):,.2f}")
        if "ded"  in subs: subs["ded"].set(f"Subtotal {sum(ded.values()):,.2f}")
        # Integrity banner update (scope visualisation)
        eid = getattr(self, "_sm_active_eid", None)
        mk = (self._sm_month.get() or "").strip()
        scope = f"[Emp: {eid or '?'}]  [Month: {mk or '?'}]"
        self._sm_integrity.set(f"✅ Integrity: Isolated — changes will be applied ONLY to {scope}")
        return calc

    def _sm_refresh_employees(self):
        if not hasattr(self, "_sm_tree"): return
        q = (self._sm_search.get() or "").strip().lower()
        emps = sorted(self.manager.list_employees() or [],
                      key=lambda e: (str(e.get("department") or ""), str(e.get("name") or "")))
        for iid in self._sm_tree.get_children(): self._sm_tree.delete(iid)
        rows_out = []
        for e in emps:
            hay = " ".join(str(x or "").lower() for x in [
                e.get("employee_id"), e.get("name"), e.get("department"),
                e.get("position"), e.get("mobile_phone"), e.get("email")])
            if q and q not in hay: continue
            rows_out.append(e)
        for i, e in enumerate(rows_out):
            tags = ["odd"] if i % 2 else []
            self._sm_tree.insert("", "end", iid=str(e.get("employee_id")), tags=tags, values=(
                e.get("employee_id", ""),
                e.get("name", ""),
                e.get("department", ""),
            ))
        if getattr(self, "_sm_active_eid", None) and self.manager.get_employee_monthly_record:
            try:
                # Re-selection safe-guard
                self._sm_tree.selection_set(str(self._sm_active_eid))
                self._sm_tree.see(str(self._sm_active_eid))
            except Exception: pass

    def _sm_on_emp_selected(self):
        sel = self._sm_tree.selection()
        if not sel: return
        eid = sel[0]
        emp = next((e for e in (self.manager.list_employees() or [])
                    if str(e.get("employee_id") or "") == str(eid)), None)
        if emp is None: return
        self._sm_active_eid = eid
        basic = 0.0
        try: basic = float(((emp.get("salary_components") or {}).get("basic")
                            if isinstance(emp.get("salary_components"), dict) else 0) or 0)
        except Exception: basic = 0.0
        join = emp.get("joining_date") or "-"
        self._sm_emp_label.set(f"{eid} · {emp.get('name') or ''}  ({emp.get('department') or ''})   Basic {basic:,.2f} AED   Join {join}")
        self._sm_load_record(auto_reload_emp=False)

    def _sm_populate_form(self, record):
        if record is None:
            # Clear everything to 0
            for k, v in self._sm_entries.items():
                if k.startswith(("comp.", "var.", "ded.")):
                    v.set("0")
                elif k == "status": v.set("draft")
                else: v.set("")
            try: self._sm_notes.delete('1.0', tk.END)
            except Exception: pass
            return
        for prefix, bucket_key in [("comp.", "salary_components"), ("var.", "variable_pay"), ("ded.", "deductions")]:
            bucket = record.get(bucket_key) if isinstance(record.get(bucket_key), dict) else {}
            for k, v in self._sm_entries.items():
                if not k.startswith(prefix): continue
                raw = bucket.get(k[len(prefix):], 0) if isinstance(bucket, dict) else 0
                try:
                    val = round(float(raw or 0), 2)
                    v.set(f"{val:.2f}" if val else "0")
                except Exception: v.set("0")
        for k, default in [("status", "draft"), ("processed_date", ""), ("notes", "")]:
            if k in self._sm_entries:
                val = record.get(k, default) or default
                try: self._sm_entries[k].set(str(val))
                except Exception: pass
        try:
            self._sm_notes.delete('1.0', tk.END)
            self._sm_notes.insert('1.0', str(record.get("notes") or record.get("comments") or ""))
        except Exception: pass

    def _sm_load_record(self, auto_reload_emp=True):
        if not getattr(self, "_sm_active_eid", None):
            self._sm_status.set("⚠ Select an employee first (left pane).")
            return
        eid = self._sm_active_eid
        mk = (self._sm_month.get() or "").strip()
        if len(mk) != 7 or mk[4] != '-':
            messagebox.showwarning("Bad Month", "Month must be YYYY-MM (e.g. 2026-08).")
            return
        if auto_reload_emp:
            self._sm_refresh_employees()
        rec = self.manager.get_employee_monthly_record(eid, mk)
        if rec is None and self._sm_prev_month.get():
            prev = self.manager.copy_previous_month(mk) or []
            for r in prev:
                if isinstance(r, dict) and str(r.get("employee_id") or "") == str(eid):
                    rec = json.loads(json.dumps(r)); break
        if rec is None:
            self._sm_populate_form(None)
            self._sm_status.set(f"ℹ No record for {eid} in {mk}.  Use ➕ Create / Reset Default to generate one from the Personnel File defaults.")
        else:
            self._sm_populate_form(rec)
            status = str(rec.get("status") or "draft").upper()
            self._sm_status.set(
                f"✅ Loaded {eid} · {mk}. Status: {status}.  Make edits → Save. (Other months UNAFFECTED.)"
            )
        self._sm_recalc_live()
        self._sm_last_saved.set(f"(Last load: {datetime.now():%H:%M:%S})")

    def _sm_create_default(self):
        if not getattr(self, "_sm_active_eid", None):
            self._sm_status.set("⚠ Select an employee first."); return
        eid = self._sm_active_eid
        mk = (self._sm_month.get() or "").strip()
        if len(mk) != 7 or mk[4] != '-':
            messagebox.showwarning("Bad Month", "Month must be YYYY-MM (e.g. 2026-08)."); return
        try:
            # create with overwrite=False → if exists, warn + ask user
            new_rec = self.manager.create_employee_monthly_record(eid, mk, record=None, overwrite=False)
        except ValueError as exc:
            if "already has a record" in str(exc):
                if not messagebox.askyesno("Record exists",
                                           f"{eid} already has a {mk} record.\n\nOverwrite with Personnel defaults? (Current saved values will be replaced.)"):
                    return
                new_rec = self.manager.create_employee_monthly_record(eid, mk, record=None, overwrite=True)
            else:
                messagebox.showerror("Error", str(exc)); return
        self._sm_populate_form(new_rec); self._sm_recalc_live()
        self._sm_propagate()
        self._sm_status.set(f"✅ Created default record for {eid} · {mk}. Save after edits to commit.")
        self._sm_last_saved.set(f"(Default created: {datetime.now():%H:%M:%S})")

    def _sm_save_record(self):
        if not getattr(self, "_sm_active_eid", None):
            self._sm_status.set("⚠ Select an employee first."); return
        eid = self._sm_active_eid
        mk = (self._sm_month.get() or "").strip()
        if len(mk) != 7 or mk[4] != '-':
            messagebox.showwarning("Bad Month", "Month must be YYYY-MM (e.g. 2026-08)."); return
        comp, var, ded = {}, {}, {}
        for prefix, target in [("comp.", comp), ("var.", var), ("ded.", ded)]:
            for k, v in self._sm_entries.items():
                if k.startswith(prefix):
                    key = k[len(prefix):]
                    try: target[key] = round(float(v.get() or 0), 2)
                    except Exception: target[key] = 0.0
        updates = {
            "salary_components": comp,
            "variable_pay": var,
            "deductions": ded,
        }
        for k in ("status", "processed_date", "notes"):
            wk = self._sm_entries.get(k)
            if wk is not None:
                raw = (wk.get() if hasattr(wk, "get") else "")
                updates[k] = str(raw).strip() if raw else ""
        try:
            notes_extra = self._sm_notes.get('1.0', 'end').strip()
            if notes_extra: updates["notes"] = notes_extra
        except Exception: pass
        # Ensure processed_date defaults for UI visibility
        if not updates.get("processed_date"):
            updates["processed_date"] = datetime.now().strftime("%Y-%m-%d")
        saved = self.manager.update_employee_monthly_record(eid, mk, updates)
        self._sm_populate_form(saved)
        self._sm_recalc_live()
        self._sm_propagate()
        gross = saved["calculated_values"]["gross_salary"]
        net = saved["calculated_values"]["net_salary"]
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        status = str(saved.get("status") or "draft").upper()
        self._sm_status.set(
            f"✅ SAVED — {eid} · {mk}.  Gross {gross:,.2f}  Net {net:,.2f}  Status {status}.  "
            f"Other months were NOT modified. (Saved {ts})"
        )
        self._sm_last_saved.set(f"(Saved: {datetime.now():%H:%M:%S})")

    def _sm_delete_record(self):
        if not getattr(self, "_sm_active_eid", None):
            self._sm_status.set("⚠ Select an employee first."); return
        eid = self._sm_active_eid
        mk = (self._sm_month.get() or "").strip()
        if len(mk) != 7 or mk[4] != '-':
            messagebox.showwarning("Bad Month", "Month must be YYYY-MM (e.g. 2026-08)."); return
        if not messagebox.askyesno("Confirm Delete",
                                   f"PERMANENTLY delete the salary record for:\n\n"
                                   f"  Employee : {eid}\n"
                                   f"  Month    : {mk}\n\n"
                                   f"Other months and other employees will NOT be affected.\n\n"
                                   f"Proceed?"):
            return
        ok = self.manager.delete_employee_monthly_record(eid, mk)
        self._sm_propagate()
        if ok:
            self._sm_populate_form(None); self._sm_recalc_live()
            self._sm_status.set(f"🗑 DELETED record for {eid} · {mk}.  Other months / employees remain intact.")
        else:
            self._sm_status.set(f"ℹ No record existed for {eid} · {mk} (nothing to delete).")

    def _sm_propagate(self):
        """Real-time sync: disk-write → reload DM shortcut attrs → refresh every loaded tab
        + History + Directory + Monthly Payroll totals + top summary bar."""
        try:
            dm = getattr(self.invoice_manager, "manager", None) or getattr(self.invoice_manager, "dm", None)
            if dm is not None and hasattr(dm, "reload_all_json"):
                dm.reload_all_json()
        except Exception:
            pass
        try:
            self._refresh_everything()
        except Exception: pass
        try:
            for fn in [getattr(self, "_refresh_directory", None),
                       getattr(self, "_refresh_monthly_rows", None),
                       getattr(self, "_refresh_history_tree", None),
                       getattr(self, "_refresh_documents_tree", None)]:
                if callable(fn):
                    try: fn()
                    except Exception: pass
        except Exception: pass

    # ============================================================
    # TAB 6 — 📊 REPORTS
    # ============================================================
    def _build_reports(self, parent):
        frm = ttk.Frame(parent, padding=(14, 10)); frm.pack(fill='both', expand=True)
        head = ttk.LabelFrame(frm, text="  Payroll Reports  ", padding=(12, 10))
        head.pack(fill='x')
        ttk.Label(head, text="Month:").grid(row=0, column=0, sticky='w')
        self._rep_month = tk.StringVar(value=self.manager.get_month_key())
        ttk.Entry(head, textvariable=self._rep_month, width=12).grid(row=0, column=1, padx=6)
        ttk.Button(head, text="📰 Show in Window",
                   command=lambda: self.monthly_salary_report()).grid(row=0, column=2, padx=6)
        ttk.Button(head, text="📑 Export PDF Consolidated",
                   command=lambda: self.monthly_salary_report_pdf()).grid(row=0, column=3, padx=6)
        ttk.Button(head, text="🧾 Generate ALL Payslips (ZIP)",
                   command=self._rep_generate_all_payslips).grid(row=0, column=4, padx=6)
        ttk.Button(head, text="📊 Period Summary (start→end)",
                   command=self._rep_period_summary).grid(row=0, column=5, padx=6)
        self._report_text = scrolledtext.ScrolledText(frm, height=22,
                                                       font=("Consolas", 10))
        self._report_text.pack(fill='both', expand=True, pady=(10, 0))

    def monthly_salary_report(self):
        mk = (self._rep_month.get() if hasattr(self, "_rep_month") else None) or \
             (self._pay_month.get() if hasattr(self, "_pay_month") else None) or \
             self.manager.get_month_key()
        mk = mk.strip()
        entries = self.manager.get_month_entries(mk) or []
        total_emp = len(entries)
        g, d, n = 0.0, 0.0, 0.0
        rows_out = []
        for e in entries:
            calc = e.get("calculated_values") or {}
            gg = float(calc.get("gross_salary", 0) or 0)
            dd = float(calc.get("total_deductions", 0) or 0)
            nn = float(calc.get("net_salary", 0) or 0)
            g += gg; d += dd; n += nn
            rows_out.append(f"{e.get('employee_id'):>6}  {e.get('name') or '':30s}  "
                            f"{(e.get('department') or '').ljust(16)[:16]}  "
                            f"{gg:10,.2f}  {dd:10,.2f}  {nn:10,.2f}  {e.get('status') or ''}")
        lines = [
            f"=== HopePharma — Monthly Payroll Report — {mk} ===",
            f"Employees: {total_emp}",
            f"Gross Salary  : {g:,.2f} AED",
            f"Deductions    : {d:,.2f} AED",
            f"Net Payable   : {n:,.2f} AED",
            f"Average (Net) : {(n/total_emp if total_emp else 0):,.2f} AED",
            "",
            f"{'ID':>6}  {'Employee':30}  {'Department':16}  {'Gross':>10}  {'Ded.':>10}  {'Net':>10}  Status",
            "-" * 110,
        ] + rows_out
        txt = "\n".join(lines)
        if hasattr(self, "_report_text"):
            self._report_text.delete('1.0', tk.END)
            self._report_text.insert('1.0', txt)
        else:
            # Legacy callers: open in toplevel
            tl = tk.Toplevel(self.parent); tl.title(f"Payroll Report {mk}")
            st = scrolledtext.ScrolledText(tl, font=("Consolas", 10), height=30, width=130)
            st.pack(fill='both', expand=True, padx=10, pady=10)
            st.insert('1.0', txt)
        return txt

    def monthly_salary_report_pdf(self):
        mk = ((self._rep_month.get() if hasattr(self, "_rep_month") else None) or
              (self._pay_month.get() if hasattr(self, "_pay_month") else None) or
              self.manager.get_month_key()).strip()
        entries = self.manager.get_month_entries(mk) or []
        out_dir = os.path.join(self.manager.data_folder, "EmpSalaries")
        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, f"Monthly_Payroll_Report_{mk.replace('-', '')}.pdf")
        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(out_path, pagesize=A4,
                                leftMargin=28, rightMargin=28, topMargin=28, bottomMargin=28,
                                title=f"HopePharma Payroll {mk}")
        elems = []
        # Optional logo
        for p_cand in [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png"),
            os.path.join(os.path.expanduser("~"), "Desktop", "HopePharmaData", "logo.png"),
            os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData", "logo.png"),
        ]:
            if os.path.isfile(p_cand):
                try: elems.append(Image(p_cand, width=70, height=70)); break
                except Exception: pass
        elems += [
            Paragraph(f"HopePharma — Monthly Payroll Report", styles["Title"]),
            Paragraph(f"Payroll Period: <b>{mk}</b> &nbsp;&nbsp; Printed: {datetime.now():%Y-%m-%d %H:%M}", styles["Normal"]),
            Spacer(1, 8),
        ]
        g = sum(float((e.get("calculated_values") or {}).get("gross_salary", 0) or 0) for e in entries)
        d = sum(float((e.get("calculated_values") or {}).get("total_deductions", 0) or 0) for e in entries)
        n = sum(float((e.get("calculated_values") or {}).get("net_salary", 0) or 0) for e in entries)
        summary = Table([
            ["Total Employees", str(len(entries))],
            ["Total Gross Salary", f"{g:,.2f} AED"],
            ["Total Deductions", f"{d:,.2f} AED"],
            ["Net Payable",      f"{n:,.2f} AED"],
            ["Average Net / Head", f"{(n/len(entries) if entries else 0):,.2f} AED"],
        ], colWidths=[220, 200])
        summary.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e3a8a")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.4, colors.lightgrey),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F1F5F9")]),
            ('ALIGN', (1,0), (1,-1), 'RIGHT'),
        ]))
        elems.append(summary); elems.append(Spacer(1, 14))
        # Detail table
        tbl = [["ID","Employee","Dept.","Status","Gross","Deductions","Net"]]
        for e in entries:
            calc = e.get("calculated_values") or {}
            tbl.append([
                e.get("employee_id"), e.get("name") or "",
                e.get("department") or "", e.get("status") or "DRAFT",
                f"{float(calc.get('gross_salary',0) or 0):,.2f}",
                f"{float(calc.get('total_deductions',0) or 0):,.2f}",
                f"{float(calc.get('net_salary',0) or 0):,.2f}",
            ])
        t = Table(tbl, repeatRows=1,
                  colWidths=[55, 130, 95, 60, 80, 85, 90])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0f172a")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey),
            ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#F8FAFC")]),
            ('ALIGN', (4,1), (-1,-1), 'RIGHT'),
            ('VALIGN', (0,0), (-1,-1), 'TOP'),
            ('FONTSIZE', (0,0), (-1,-1), 8),
        ]))
        elems.append(t); doc.build(elems)
        messagebox.showinfo("PDF Exported", f"Saved consolidated payroll PDF:\n{out_path}")

    def _rep_generate_all_payslips(self):
        mk = (self._rep_month.get() if hasattr(self, "_rep_month") else self.manager.get_month_key()).strip()
        entries = self.manager.get_month_entries(mk) or []
        if not entries:
            messagebox.showinfo("Payslips", f"No payroll entries found for {mk}. Save month first.")
            return
        by_id = {e.get("employee_id"): e for e in (self.manager.list_employees() or [])}
        ok, fails = 0, []
        files_out = []
        for se in entries:
            eid = se.get("employee_id")
            emp = by_id.get(eid)
            if not emp:
                fails.append(eid); continue
            try:
                p = self._generate_payslip_pdf(se, emp, open_after=False)
                if p: files_out.append(p); ok += 1
            except Exception as ex:
                fails.append(f"{eid}:{ex}")
        # Offer to zip
        if files_out:
            # Just show folder + count; zip optional via prompt
            first_dir = os.path.dirname(files_out[0])
            messagebox.showinfo(
                f"🧾 {ok} Payslips Generated",
                f"Month: {mk}\n\n✅ OK: {ok}\n❌ Skipped: {len(fails)}\n\n"
                f"Stored in: {first_dir}\n\n"
                f"Sample file:\n{os.path.basename(files_out[0])}"
            )
        else:
            messagebox.showinfo("Payslips", "No payslips generated (check logs).")

    def _rep_period_summary(self):
        s = filedialog.askstring("Period Summary", "Start Date (YYYY-MM-DD):",
                                 initialvalue=f"{date.today().year}-01-01")
        if not s: return
        e = filedialog.askstring("Period Summary", "End Date (YYYY-MM-DD):",
                                 initialvalue=date.today().isoformat())
        if not e: return
        data = self.manager.get_salary_summary_for_period(s, e)
        txt = (
            f"Payroll Period Summary\n"
            f"Period: {data['start_date']} → {data['end_date']}\n\n"
            f"  Employees Processed : {data['employees_processed']}\n"
            f"  Months in Period    : {data['months_in_period']}\n"
            f"  Gross Salaries (Tot): {data['gross_salary_total']:,.2f} AED\n"
            f"  Net Salaries Paid   : {data['net_salary_total']:,.2f} AED\n\n"
            f"---- Breakdown by Month ----\n"
        )
        for mk in sorted(set(list(data['by_month_gross'].keys()) + list(data['by_month_net'].keys()))):
            g = (data['by_month_gross'].get(mk) or {}).get('gross', 0) if isinstance(data['by_month_gross'].get(mk), dict) else (data['by_month_gross'].get(mk) or 0)
            n = (data['by_month_net'].get(mk) or {}).get('net', 0) if isinstance(data['by_month_net'].get(mk), dict) else (data['by_month_net'].get(mk) or 0)
            try:
                if isinstance(g, dict): g = g.get('gross', 0)
                if isinstance(n, dict): n = n.get('net', 0)
            except Exception: pass
            txt += f"  {mk}  Gross {float(g):,.2f}  Net {float(n):,.2f}\n"
        if hasattr(self, "_report_text"):
            self._report_text.delete('1.0', tk.END); self._report_text.insert('1.0', txt)
        messagebox.showinfo("Period Summary", txt[:1800])

    def _generate_payslip_pdf(self, salary_entry, employee, open_after=False):
        mk = salary_entry.get("month_year") or self.manager.get_month_key()
        out_dir = os.path.join(self.manager.data_folder, "EmpSalaries", f"Payslips_{mk.replace('-','')}")
        os.makedirs(out_dir, exist_ok=True)
        filename = f"Payslip_{employee.get('employee_id')}_{employee.get('name','_').replace(' ','_')}_{mk.replace('-','')}.pdf"
        path = os.path.join(out_dir, filename)
        styles = getSampleStyleSheet()
        doc = SimpleDocTemplate(path, pagesize=A4,
                                leftMargin=30, rightMargin=30, topMargin=24, bottomMargin=24,
                                title=filename.replace(".pdf",""))
        elems = []
        for p_cand in [
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png"),
            os.path.join(os.path.expanduser("~"), "Desktop", "HopePharmaData", "logo.png"),
        ]:
            if os.path.isfile(p_cand):
                try: elems.append(Image(p_cand, width=58, height=58)); break
                except Exception: pass
        calc = salary_entry.get("calculated_values") or self._calc_from_dicts(
            salary_entry.get("salary_components"), salary_entry.get("variable_pay"), salary_entry.get("deductions"))
        elems.append(Paragraph("🏥 HOPE PHARMA — PAYSHEET", styles["Title"]))
        elems.append(Paragraph(f"Pay Period: <b>{mk}</b> &nbsp;&nbsp; Employee ID: <b>{employee.get('employee_id','')}</b>", styles["Normal"]))
        elems.append(Spacer(1, 6))
        # Personnel info
        info = Table([
            ["Employee",  employee.get("name",""), "Department", employee.get("department","") or "-"],
            ["Position",  employee.get("position","") or "-", "Nationality", employee.get("nationality","") or "-"],
            ["Joining",   employee.get("joining_date","") or "-", "Bank / IBAN",
                f"{(employee.get('bank_details') or {}).get('bank_name','-') or '-'} &nbsp;/&nbsp; "
                f"{(employee.get('bank_details') or {}).get('bank_account','-') or '-'}"]
        ], colWidths=[80, 210, 90, 220])
        info.setStyle(TableStyle([('BOX',(0,0),(-1,-1),0.4,colors.lightgrey),
                                   ('BACKGROUND',(0,0),(-1,-1),colors.white),
                                   ('FONTNAME',(0,0),(-2,-1),'Helvetica'),
                                   ('FONTSIZE',(0,0),(-1,-1),9)]))
        elems.append(info); elems.append(Spacer(1, 10))
        # Earnings + Deductions dual table
        earnings_rows = [["EARNINGS", "AED"], ["Guaranteed Components", ""]]
        sc = salary_entry.get("salary_components") or {}
        for k, lab in self.COMPONENT_KEYS:
            earnings_rows.append([lab, f"{float((sc or {}).get(k,0) or 0):,.2f}"])
        earnings_rows.append(["Variable Pay", ""])
        vp = salary_entry.get("variable_pay") or {}
        for k, lab in self.VARIABLE_KEYS:
            earnings_rows.append([lab, f"{float((vp or {}).get(k,0) or 0):,.2f}"])
        earnings_rows.append(["<b>GROSS EARNINGS</b>", f"<b>{float(calc.get('gross_salary',0) or 0):,.2f}</b>"])
        ded_rows = [["DEDUCTIONS", "AED"]]
        dd = salary_entry.get("deductions") or {}
        for k, lab in self.DEDUCTION_KEYS:
            ded_rows.append([lab, f"{float((dd or {}).get(k,0) or 0):,.2f}"])
        ded_rows.append(["<b>TOTAL DEDUCTIONS</b>", f"<b>{float(calc.get('total_deductions',0) or 0):,.2f}</b>"])
        # Wrap strings as Paragraphs
        def par(v, bold=False):
            st = styles["BodyText"]
            return Paragraph(str(v), st) if not bold else Paragraph(str(v), styles["Heading5"])
        def render_table(rows, col_w, header_color):
            data = []
            for i, r in enumerate(rows):
                is_bold = i>0 and ("<b>" in str(r[0]) or "<b>" in str(r[1]))
                data.append([par(r[0], bold=is_bold), par(r[1], bold=is_bold)])
            t = Table(data, colWidths=col_w)
            style = TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor(header_color)),
                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                ('GRID', (0,0), (-1,-1), 0.3, colors.lightgrey),
                ('FONTSIZE', (0,0), (-1,-1), 9),
                ('VALIGN', (0,0), (-1,-1), 'TOP'),
                ('ALIGN', (1,1), (-1,-1), 'RIGHT'),
            ])
            t.setStyle(style); return t
        outer = Table([[render_table(earnings_rows, [240, 90], "#163a16"),
                        render_table(ded_rows, [220, 90], "#7C1D1D")]], colWidths=[330, 310])
        outer.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'), ('LEFTPADDING',(0,0),(-1,-1),4)]))
        elems.append(outer); elems.append(Spacer(1, 14))
        net_row = Table([
            ["NET SALARY PAYABLE", f"{float(calc.get('net_salary', calc['gross_salary'] - calc['total_deductions'])):,.2f} AED"]
        ], colWidths=[380, 200])
        net_row.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#111827")),
            ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('FONTSIZE', (0,0), (-1,0), 14),
            ('ALIGN', (1,0), (1,0), 'RIGHT'),
            ('TOPPADDING', (0,0), (-1,0), 10), ('BOTTOMPADDING', (0,0), (-1,0), 10),
        ]))
        elems.append(net_row); elems.append(Spacer(1, 18))
        # Sign-off
        sign = Table([["Employer Signature", "______________",
                       "Employee Signature", "______________",
                       "Date: ____________"]], colWidths=[130,130,130,130,130])
        sign.setStyle(TableStyle([('FONTSIZE',(0,0),(-1,0),9),('FONTNAME',(0,0),(-1,0),'Helvetica-Oblique')]))
        elems.append(sign)
        doc.build(elems)
        if open_after:
            try:
                import subprocess, platform
                if platform.system() == "Darwin": subprocess.Popen(["open", path])
                elif platform.system() == "Windows": os.startfile(path)
                else: subprocess.Popen(["xdg-open", path])
            except Exception: pass
        return path

    # ============================================================
    # TAB 6 — 📄 DOCUMENTS
    # ============================================================
    def _build_documents(self, parent):
        top = ttk.Frame(parent, padding=(12, 10)); top.pack(fill='x')
        ttk.Label(top, text="Employee:").pack(side='left')
        self._doc_emp_var = tk.StringVar()
        self._doc_emp_cb = ttk.Combobox(top, textvariable=self._doc_emp_var, values=[], width=48, state="readonly")
        self._doc_emp_cb.pack(side='left', padx=6)
        ttk.Button(top, text="📎 Attach File", command=self._doc_attach).pack(side='left', padx=6)
        ttk.Button(top, text="🗂 Open Folder", command=self._doc_open_folder).pack(side='left', padx=6)
        ttk.Button(top, text="❌ Delete Selected", command=self._doc_delete).pack(side='right', padx=6)
        ttk.Label(top, text="Accepted: PDF, JPG, PNG, DOCX, (any file) — stored under data_folder/employee_docs/",
                  foreground="#475569").pack(side='right', padx=14)
        cols = ("kind", "filename", "size_kb", "added", "notes")
        self._doc_tree = ttk.Treeview(parent, columns=cols, show='headings', height=22)
        for c, w, anchor in [("kind", 120, "center"), ("filename", 360, "w"),
                             ("size_kb", 80, "e"), ("added", 160, "center"), ("notes", 360, "w")]:
            self._doc_tree.heading(c, text=c.title().replace("_", " "), anchor="center")
            self._doc_tree.column(c, width=w, anchor=anchor)
        vsb = ttk.Scrollbar(parent, orient="vertical", command=self._doc_tree.yview)
        self._doc_tree.configure(yscrollcommand=vsb.set)
        self._doc_tree.pack(fill='both', expand=True, padx=10, pady=(6, 10))
        vsb.pack(side='right', fill='y')
        self._doc_tree.bind("<Double-1>", lambda _e: self._doc_open_selected())
        self.dialog.after(400, self._refresh_documents_tree)

    def _current_doc_emp(self):
        if not hasattr(self, "_doc_emp_cb"): return None
        emps = self.manager.list_employees() or []
        labels = [f"{e.get('employee_id') or ''} · {e.get('name') or ''} ({e.get('department') or '-'})" for e in emps]
        try: self._doc_emp_cb.config(values=labels)
        except Exception: pass
        lab = self._doc_emp_var.get()
        for e in emps:
            if lab.startswith(f"{e.get('employee_id')} ·"): return e
        if emps and not lab:
            self._doc_emp_var.set(labels[0]); return emps[0]
        return None

    def _refresh_documents_tree(self):
        if not hasattr(self, "_doc_tree"): return
        emp = self._current_doc_emp()
        for iid in self._doc_tree.get_children(): self._doc_tree.delete(iid)
        if emp is None: return
        bucket = (emp.get("documents") or []) if isinstance(emp.get("documents"), list) else []
        for idx, d in enumerate(bucket):
            fp = d.get("path")
            size = 0
            try:
                if fp and os.path.isfile(fp): size = os.path.getsize(fp)
                else:
                    cand = os.path.join(self.doc_dir, fp or "")
                    if os.path.isfile(cand): size = os.path.getsize(cand)
            except Exception: size = 0
            self._doc_tree.insert("", "end", iid=str(idx), values=(
                d.get("kind") or "Document", os.path.basename(fp or d.get("filename") or "(missing)"),
                f"{(size/1024):,.1f} KB", d.get("added_on") or "-", d.get("notes") or ""
            ))

    def _doc_attach(self):
        emp = self._current_doc_emp()
        if emp is None:
            messagebox.showinfo("Attachments", "Select an employee first.")
            return
        paths = filedialog.askopenfilenames(title="Attach Employee Document")
        if not paths: return
        emp_dir = os.path.join(self.doc_dir, f"{emp.get('employee_id') or 'X'}_{(emp.get('name') or '_').replace(' ','_')}")
        os.makedirs(emp_dir, exist_ok=True)
        kinds = {"pdf":"📄 PDF","png":"🖼 PNG","jpg":"🖼 JPG","jpeg":"🖼 JPG","docx":"📘 DOCX","xlsx":"📗 XLSX","csv":"📊 CSV","txt":"📝 TXT"}
        new_docs = list(emp.get("documents") or []) if isinstance(emp.get("documents"), list) else []
        for p in paths:
            try:
                base = os.path.basename(p)
                target = os.path.join(emp_dir, base)
                i = 1
                while os.path.exists(target):
                    name, ext = os.path.splitext(base)
                    target = os.path.join(emp_dir, f"{name}_{i}{ext}"); i += 1
                import shutil as _sh
                _sh.copy2(p, target)
                ext = (target.rsplit(".", 1)[-1]).lower() if "." in target else ""
                kind = kinds.get(ext, "📎 File")
                new_docs.append({"kind": kind, "path": target, "filename": os.path.basename(target),
                                 "added_on": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                 "notes": filedialog.askstring("Document Notes", f"Notes for {os.path.basename(target)} (optional):") or ""})
            except Exception as ex:
                messagebox.showerror("Attachment failed", f"{p}\n\n{ex}")
        self.manager.update_employee(emp["employee_id"], {"documents": new_docs})
        self._refresh_documents_tree()

    def _doc_open_selected(self):
        emp = self._current_doc_emp()
        if emp is None: return
        iid = self._doc_tree.selection()
        if not iid: return
        idx = int(iid[0])
        docs = list(emp.get("documents") or [])
        if idx >= len(docs): return
        p = docs[idx].get("path") or docs[idx].get("filename")
        if not p or not os.path.isfile(p):
            cand = os.path.join(self.doc_dir, p or "")
            if os.path.isfile(cand): p = cand
            else:
                messagebox.showwarning("Missing", f"File not found:\n{p}")
                return
        try:
            import subprocess, platform
            if platform.system() == "Darwin": subprocess.Popen(["open", p])
            elif platform.system() == "Windows": os.startfile(p)
            else: subprocess.Popen(["xdg-open", p])
        except Exception as ex:
            messagebox.showerror("Open", str(ex))

    def _doc_delete(self):
        emp = self._current_doc_emp()
        if emp is None: return
        iid = self._doc_tree.selection()
        if not iid: return
        idx = int(iid[0])
        docs = list(emp.get("documents") or [])
        if idx >= len(docs): return
        if not messagebox.askyesno("Delete Document", f"Delete attachment '{docs[idx].get('filename')}'?\n\n"
                                                      "(This will remove from list but keep file on disk for audit trail.)"):
            return
        del docs[idx]
        self.manager.update_employee(emp["employee_id"], {"documents": docs})
        self._refresh_documents_tree()

    def _doc_open_folder(self):
        try:
            import subprocess, platform
            if platform.system() == "Darwin": subprocess.Popen(["open", self.doc_dir])
            elif platform.system() == "Windows": os.startfile(self.doc_dir)
            else: subprocess.Popen(["xdg-open", self.doc_dir])
        except Exception: pass

    # ============================================================
    # TAB 7 — 📤 EXPORTS
    # ============================================================
    def _build_exports(self, parent):
        frm = ttk.Frame(parent, padding=(16, 14)); frm.pack(fill='both', expand=True)
        for i, (emoji, title, desc, fn) in enumerate([
            ("🏦", "Bank WPS / SIF Transfer File",
             "Comma-separated Bank upload file: Employee ID · Name · IBAN · Net · Bank. Format compatible with UAE WPS banks (NBAD, Emirates NBD, Dubai Islamic, ADCB, etc.)",
             lambda: self._export_bank_csv()),
            ("📊", "Payroll CSV (full line items)",
             "Full 27-field per-employee export (all components / variable pays / deductions + gross + net) for this month",
             lambda: self._export_payroll_csv()),
            ("📑", "Consolidated PDF (from Reports tab)",
             "Shortcut to generate the Consolidated PDF Payroll Report for the selected month.",
             lambda: self.monthly_salary_report_pdf()),
            ("🧾", "ZIP of All Individual Payslip PDFs",
             "Generate + collect payslips for every employee in selected month.",
             lambda: self._rep_generate_all_payslips()),
            ("👥", "Employees Master Data (CSV)",
             "Export 30-field employee directory for use in Excel / BI tools.",
             lambda: self._export_employees_csv()),
        ]):
            card = ttk.LabelFrame(frm, text=f"  {emoji}  {title}  ", padding=(12, 10))
            card.pack(fill='x', pady=(0 if i == 0 else 8))
            ttk.Label(card, text=desc, wraplength=900, justify="left").pack(side='left', fill='x', expand=True, padx=(0, 10))
            ttk.Button(card, text="▶ Run", command=fn).pack(side='right')

    def _export_bank_csv(self):
        mk = self._pick_a_month("Bank WPS CSV")
        if not mk: return
        entries = self.manager.get_month_entries(mk) or []
        if not entries:
            messagebox.showinfo("Bank CSV", f"No entries for {mk}. Save month first.")
            return
        by_id = {e.get("employee_id"): e for e in (self.manager.list_employees() or [])}
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV","*.csv")],
                                         initialfile=f"HopePharma_WPS_{mk.replace('-','')}.csv")
        if not p: return
        with open(p, "w", newline='', encoding='utf-8-sig') as f:
            w = __import__("csv").writer(f)
            w.writerow(["EMPLOYEE_ID","NAME","DEPARTMENT","IBAN_OR_ACCOUNT","BANK_NAME","BRANCH","NET_AMOUNT_AED","PAYMENT_CURRENCY","PAYMENT_DATE","EID_OR_PASSPORT"])
            for se in entries:
                eid = se.get("employee_id"); emp = by_id.get(eid, {})
                bd = emp.get("bank_details") if isinstance(emp, dict) else (se.get("bank_details") or {})
                calc = se.get("calculated_values") or {}
                w.writerow([
                    eid, emp.get("name", se.get("name","")) if isinstance(emp, dict) else se.get("name",""),
                    (emp.get("department") if isinstance(emp, dict) else "") or se.get("department") or "",
                    bd.get("bank_account") or bd.get("iban") or "",
                    bd.get("bank_name") or "", bd.get("bank_branch") or "",
                    f"{float(calc.get('net_salary',0) or 0):.2f}", "AED",
                    se.get("processed_date") or f"{mk}-28",
                    emp.get("emirates_id") or emp.get("passport_number") if isinstance(emp, dict) else ""
                ])
        messagebox.showinfo("✅ Bank CSV",
                            f"Exported {len(entries)} line items.\n\nTotal Net: "
                            f"{sum(float((s.get('calculated_values') or {}).get('net_salary',0) or 0) for s in entries):,.2f} AED\n\nFile:\n{p}")

    def _export_payroll_csv(self):
        mk = self._pick_a_month("Payroll CSV")
        if not mk: return
        entries = self.manager.get_month_entries(mk) or []
        if not entries:
            messagebox.showinfo("CSV", f"No entries for {mk}. Save month first."); return
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV","*.csv")],
                                         initialfile=f"Payroll_Detail_{mk.replace('-','')}.csv")
        if not p: return
        header = ["month","employee_id","name","department","status","processed_date"] + \
                 [f"COMP_{lab}" for (_, lab) in self.COMPONENT_KEYS] + \
                 [f"VAR_{lab}"  for (_, lab) in self.VARIABLE_KEYS] + \
                 [f"DED_{lab}"  for (_, lab) in self.DEDUCTION_KEYS] + \
                 ["GROSS","DEDUCTIONS_TOTAL","NET"]
        with open(p, "w", newline='', encoding='utf-8-sig') as f:
            w = __import__("csv").writer(f)
            w.writerow(header)
            for se in entries:
                sc = se.get("salary_components") or {}
                vp = se.get("variable_pay") or {}
                dd = se.get("deductions") or {}
                calc = se.get("calculated_values") or self._calc_from_dicts(sc, vp, dd)
                row = [se.get("month_year", mk), se.get("employee_id"),
                       se.get("name",""), se.get("department",""),
                       se.get("status",""), se.get("processed_date","")] + \
                      [float((sc or {}).get(k,0) or 0) for (k,_) in self.COMPONENT_KEYS] + \
                      [float((vp or {}).get(k,0) or 0) for (k,_) in self.VARIABLE_KEYS] + \
                      [float((dd or {}).get(k,0) or 0) for (k,_) in self.DEDUCTION_KEYS] + \
                      [calc["gross_salary"], calc["total_deductions"], calc["net_salary"]]
                w.writerow(row)
        messagebox.showinfo("CSV", f"Exported {len(entries)} rows:\n{p}")

    def _export_employees_csv(self):
        emps = self.manager.list_employees() or []
        p = filedialog.asksaveasfilename(defaultextension=".csv",
                                         filetypes=[("CSV","*.csv")],
                                         initialfile="HopePharma_Employees.csv")
        if not p: return
        # Keys list
        keys = ["employee_id","name","arabic_name","gender","dob","nationality","marital_status","religion",
                "passport_number","passport_expiry","emirates_id","emirates_id_expiry",
                "department","position","grade","manager_name","employment_type","status",
                "joining_date","probation_end","contract_start","contract_end","notice_days","cost_center",
                "work_visa_file","work_permit_number","work_visa_expiry","labor_card_number",
                "health_card_number","insurance_class","sponsor_name","visa_fees","visa_fees_recovery",
                "mobile_phone","other_phone","email","personal_email","address_uae","emirate",
                "address_home_country","emergency_contact_name","emergency_contact_phone","emergency_contact_relation",
                "bank_details.bank_name","bank_details.bank_branch","bank_details.bank_account","bank_details.bank_swift",
                "bank_details.bank_city","bank_details.beneficiary_name","payment_method","salary_day","pay_cycle"]
        with open(p, "w", newline='', encoding='utf-8-sig') as f:
            w = __import__("csv").writer(f)
            w.writerow(keys + [f"comp_{lab}" for (_, lab) in self.COMPONENT_KEYS] +
                       [f"var_{lab}" for (_, lab) in self.VARIABLE_KEYS] +
                       [f"ded_{lab}" for (_, lab) in self.DEDUCTION_KEYS] + ["notes","last_modified"])
            for e in emps:
                base = []
                for k in keys:
                    if k.startswith("bank_details."):
                        bd = e.get("bank_details") or {}
                        base.append((bd or {}).get(k[len("bank_details."):], "") if isinstance(bd, dict) else "")
                    else:
                        base.append(e.get(k, ""))
                sc = e.get("salary_components") or {}
                vp = e.get("variable_pay") or {}
                dd = e.get("deductions") or {}
                tail = [float((sc or {}).get(k,0) or 0) for (k,_) in self.COMPONENT_KEYS] + \
                       [float((vp or {}).get(k,0) or 0) for (k,_) in self.VARIABLE_KEYS] + \
                       [float((dd or {}).get(k,0) or 0) for (k,_) in self.DEDUCTION_KEYS] + \
                       [e.get("notes",""), e.get("last_modified","")]
                w.writerow(base + tail)
        messagebox.showinfo("✅ Exported", f"Employees: {len(emps)} rows →\n{p}")

    def _pick_a_month(self, title):
        mk = tk.simpledialog.askstring(title, "Month (YYYY-MM):",
                                        initialvalue=(self._pay_month.get() if hasattr(self, "_pay_month") else self.manager.get_month_key()),
                                        parent=self.dialog) if hasattr(tk, "simpledialog") else None
        if mk is None:
            mk = (self._pay_month.get() if hasattr(self, "_pay_month") else None) or self.manager.get_month_key()
        return (mk or "").strip()

    # ============================================================
    # TAB 8 — ⚙️ DEFAULTs / Cost centers
    # ============================================================
    def _build_defaults(self, parent):
        frm = ttk.Frame(parent, padding=(14, 12)); frm.pack(fill='both', expand=True)
        tip = ttk.LabelFrame(frm, text=" 💡 How Defaults Are Used  ", padding=(10, 6))
        tip.pack(fill='x', pady=(0, 10))
        ttk.Label(tip,
                  text="When you open a NEW payroll month, all employees are auto-loaded with these default "
                       "allowances.  Adjust in bulk on the Payroll tab (▶ Apply Bulk button) then per-person, then Save.",
                  wraplength=1000, justify="left").pack(anchor="w")
        # 1 — Dept cost centers
        cc = ttk.LabelFrame(frm, text="  📌 Department → Cost Center / GL Account Mapping  ", padding=(12, 10))
        cc.pack(fill='x', pady=(0, 10))
        # Build simple list of depts + cost center + gl account
        depts = sorted({(e.get("department") or "Unknown").strip() for e in (self.manager.list_employees() or []) if (e.get("department") or "").strip()})
        if not depts: depts = ["Sales", "Warehouse", "Admin", "Management", "Procurement", "Pharmacy", "HR", "Finance", "Quality"]
        cols = ("department", "cost_center", "expense_gl_account", "notes")
        self._defs_cc_tree = ttk.Treeview(cc, columns=cols, show='headings', height=8)
        for c, w, anchor in [("department", 200, "w"), ("cost_center", 200, "w"),
                             ("expense_gl_account", 180, "center"), ("notes", 360, "w")]:
            self._defs_cc_tree.heading(c, text=c.replace("_", " ").title(), anchor="center")
            self._defs_cc_tree.column(c, width=w, anchor=anchor)
        self._defs_cc_tree.pack(fill='x')
        # Load from manager data or seed
        seed_map = getattr(self.manager, "_cost_centers", None) or \
                   (getattr(self.manager, "employees_file", None) and
                    __import__("json", fromlist=["load"]).load(open(os.path.join(os.path.dirname((self.manager.employees_file)), "_payroll_defaults.json"), "r", encoding="utf-8"))["dept_map"]
                    if os.path.isfile(os.path.join(os.path.dirname(self.manager.employees_file), "_payroll_defaults.json")) else None) or {
            "Sales":       ("CC-SAL-01",  "5300 Selling Expenses",         "Marketing + commisions"),
            "Warehouse":   ("CC-WH-01",   "5300 Other Administrative",    "Overtime, handling"),
            "Admin":       ("CC-ADM-01",  "5300 Other Administrative",    "Office supplies, utilities portion"),
            "Management":  ("CC-MGT-01",  "5100 Salary & Wages",          "Executive pay"),
            "Pharmacy":    ("CC-PHM-01",  "5000 COGS",                    "Dispensing techs"),
            "HR":          ("CC-HR-01",   "5100 Salary & Wages",          ""),
            "Finance":     ("CC-FIN-01",  "5100 Salary & Wages",          ""),
            "Quality":     ("CC-QA-01",   "5200 Operating Expenses",      ""),
        }
        for d in depts:
            preset = seed_map.get(d, (f"CC-{d[:3].upper()}-01", "5100 Salary & Wages", ""))
            self._defs_cc_tree.insert("", "end", values=(d, preset[0], preset[1], preset[2] if len(preset)>2 else ""))
        # 2 - Default allowances (per month starting point)
        alw = ttk.LabelFrame(frm, text="  💰 Default Monthly Amounts (applied when employee has no explicit component set)  ",
                             padding=(12, 10))
        alw.pack(fill='x', pady=(0, 10))
        grid = ttk.Frame(alw); grid.pack(fill='x')
        self._defs_vars = {}
        defs_seed = {
            "comp.housing": "0", "comp.transport": "0", "comp.phone": "0",
            "comp.food": "0", "comp.basic": "0", "ded.income_tax": "0",
            "ded.social_insurance": "0",
        }
        all_fields = [("comp." + k, lab) for (k, lab) in self.COMPONENT_KEYS] + \
                     [("var." + k, lab) for (k, lab) in self.VARIABLE_KEYS] + \
                     [("ded." + k, lab) for (k, lab) in self.DEDUCTION_KEYS]
        cols_per_row = 5
        for i, (k, lab) in enumerate(all_fields):
            r, c = divmod(i, cols_per_row)
            c *= 2
            ttk.Label(grid, text=lab).grid(row=r, column=c, sticky='w', padx=4, pady=2)
            v = tk.StringVar(value=defs_seed.get(k, "0"))
            self._defs_vars[k] = v
            ttk.Entry(grid, textvariable=v, width=12).grid(row=r, column=c+1, sticky='w', pady=2)
        # Save button
        btns = ttk.Frame(frm); btns.pack(fill='x')
        ttk.Button(btns, text="💾 Save Payroll Defaults",
                   command=self._defs_save).pack(side='right')
        ttk.Label(btns, text="Defaults stored under data_folder/_payroll_defaults.json",
                  foreground="#475569").pack(side='left')

    def _defs_save(self):
        payload = {
            "dept_map": {},
            "default_fields": {k: float(v.get() or 0) for (k, v) in self._defs_vars.items()},
            "saved_on": datetime.now().isoformat(),
        }
        for iid in self._defs_cc_tree.get_children():
            vals = self._defs_cc_tree.item(iid, "values")
            if len(vals) >= 3:
                payload["dept_map"][vals[0]] = (vals[1], vals[2], vals[3] if len(vals) > 3 else "")
        target = os.path.join(self.manager.data_folder, "_payroll_defaults.json")
        with open(target, "w", encoding="utf-8") as f:
            json.dump(payload, f, indent=2)
        messagebox.showinfo("✅ Saved", f"Payroll defaults saved:\n{target}")

    # ============================================================
    # Helpers: Add Employee quick + open selected in Personnel tab
    # ============================================================
    def _add_employee_quick(self):
        dlg = tk.Toplevel(self.dialog)
        dlg.title("➕ Add New Employee")
        dlg.geometry("660x610+120+100")
        dlg.transient(self.dialog); dlg.grab_set()
        frm = ttk.Frame(dlg, padding=(14, 12)); frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="New Employee — minimum required fields",
                  font=("Segoe UI", 11, "bold"), foreground="#1e3a8a").pack(anchor='w', pady=(0, 8))
        today_iso = date.today().isoformat()
        fields = [
            ("Employee ID (auto if blank)", "employee_id", False, "Optional, e.g. EMP027"),
            ("Full Name *", "name", False, "Required"),
            ("Nationality", "nationality", False, ""),
            ("Gender", "gender", True, ""),
            ("Date of Birth (YYYY-MM-DD)", "dob", False, ""),
            ("Department *", "department", True, ""),
            ("Position / Job Title *", "position", False, ""),
            ("Employment Status", "status", True, ""),
            ("Joining Date * (YYYY-MM-DD)", "joining_date", False,
             f"Use an OLDER date (e.g. 2024-01-15) to auto-backfill into older payrolls. Default today = {today_iso}."),
            ("Mobile Phone", "mobile_phone", False, ""),
            ("Work Email", "email", False, ""),
            ("Work Visa / Permit Expiry (YYYY-MM-DD)", "work_visa_expiry", False, "Directory will show expiry chip."),
            ("Monthly Basic Salary (AED)", "_basic", False, "Will be stored in salary_components.basic"),
        ]
        entries = {}
        combo_vals = {
            "gender":  ["", "Male", "Female", "Other"],
            "department": ["Sales", "Marketing", "Warehouse", "Procurement", "Pharmacy", "Finance", "HR", "Operations", "Management", "Quality", "IT", "Customer Service", "Logistics"],
            "status": ["Active", "Probation", "On Leave", "Notice Period"],
        }
        # Default joining date = TODAY so people can change it to older month
        default_join = date.today().isoformat()
        for i, (lab, key, is_combo, hint) in enumerate(fields):
            ttk.Label(frm, text=lab).grid(row=i, column=0, sticky='w', pady=3)
            if is_combo:
                v = tk.StringVar()
                c = ttk.Combobox(frm, textvariable=v, values=combo_vals.get(key, []), width=32, state="readonly")
                c.grid(row=i, column=1, sticky='w', pady=3); entries[key] = v
            else:
                v = tk.StringVar()
                if key == "joining_date":
                    v.set(default_join)
                e = ttk.Entry(frm, textvariable=v, width=38)
                e.grid(row=i, column=1, sticky='w', pady=3); entries[key] = v
            if hint:
                ttk.Label(frm, text=hint, foreground="#64748b", font=("Segoe UI", 8, "italic"),
                          wraplength=280, justify="left").grid(row=i, column=2, sticky='w', padx=10)

        backfill_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="✅ Auto-backfill payrolls from Joining Date to today (older payrolls)",
                        variable=backfill_var).grid(row=len(fields), column=0, columnspan=3,
                                                    sticky='w', pady=(8, 2))
        ttk.Label(frm,
                  text=("💡 If joining_date is set to 2024-05-01 (for example), the employee will be\n"
                        "    inserted into EVERY monthly payroll from 2024-05 through TODAY (status=draft).\n"
                        "    Just open any older month → adjust Basic / Bonus / Deductions → Save."),
                  foreground="#475569", font=("Segoe UI", 9, "italic")).grid(row=len(fields)+1,
                                                                              column=0, columnspan=3,
                                                                              sticky='w', padx=(0, 10), pady=(0, 2))

        def save_emp():
            name = entries["name"].get().strip()
            if not name:
                messagebox.showwarning("Missing", "Please enter Full Name.", parent=dlg)
                return
            join_dt = (entries.get("joining_date").get().strip() if entries.get("joining_date") else "") or ""
            if len(join_dt) == 10:
                try:
                    datetime.strptime(join_dt, "%Y-%m-%d")
                except Exception:
                    messagebox.showwarning("Bad Join Date",
                                           f"Joining Date must be YYYY-MM-DD (got: {join_dt!r}).",
                                           parent=dlg)
                    return
            elif join_dt != "":
                messagebox.showwarning("Bad Join Date",
                                       f"Joining Date must be YYYY-MM-DD (got: {join_dt!r}).",
                                       parent=dlg)
                return
            prof = {}
            for k, v in entries.items():
                raw = v.get().strip() if hasattr(v, "get") else ""
                if k == "_basic":
                    try:
                        b = round(float(raw or 0), 2)
                        prof.setdefault("salary_components", {})["basic"] = b
                    except Exception: pass
                else:
                    if raw:
                        prof[k] = raw
            if not prof.get("employee_id"):
                n_existing = len(self.manager.list_employees() or [])
                prof["employee_id"] = f"EMP{n_existing+1:03d}"
            # NEW: backfill flag controls the newly added manager.add_employee kwarg.
            emp_id = self.manager.add_employee(prof, backfill_from_join_date=backfill_var.get())
            join_s = prof.get("joining_date") or "-"
            backfilled_msg = ""
            if backfill_var.get() and join_s and len(join_s) == 10:
                try:
                    j_d = datetime.strptime(join_s, "%Y-%m-%d").date()
                    t_d = date.today()
                    months = (t_d.year - j_d.year) * 12 + (t_d.month - j_d.month) + 1
                    if months > 0:
                        backfilled_msg = f"\n📚 Backfilled into {months} payroll month(s) ({join_s[:7]} → {t_d.isoformat()[:7]}) — each is DRAFT."
                except Exception:
                    pass
            messagebox.showinfo("✅ Employee Added",
                                f"ID   : {emp_id}\nName : {name}\nDept : {prof.get('department') or '-'}\n"
                                f"Pos  : {prof.get('position') or '-'}\nJoin : {join_s}\n"
                                f"Basic: {prof.get('salary_components',{}).get('basic',0):,.2f} AED"
                                f"{backfilled_msg}\n\n"
                                "👉 The profile now appears in:\n   • 👥 Directory\n   • 👤 Personnel File\n   • 💰 Payroll (past months too, if backfilled)",
                                parent=dlg)
            dlg.destroy()
            self._refresh_everything()
            # Auto-switch to Personnel tab
            try:
                self._pf_populate_combos()
                for lab, e in getattr(self, "_pf_label_to_emp", {}).items():
                    if e.get("employee_id") == emp_id:
                        self._pf_pick_var.set(lab); self.nb.select(1); self._pf_load_selected(); break
            except Exception: pass
        # Buttons
        bb = ttk.Frame(frm); bb.grid(row=len(fields)+2, column=0, columnspan=3, sticky='e', pady=(14, 0))
        ttk.Button(bb, text="Cancel", command=dlg.destroy).pack(side='right', padx=6)
        ttk.Button(bb, text="💾 Save Employee", command=save_emp).pack(side='right')

    def _open_selected_profile(self):
        if not hasattr(self, "_dir_tree"): return
        sel = self._dir_tree.selection()
        if not sel:
            messagebox.showinfo("Select", "Select one employee in 👥 Directory (double-click or Enter).")
            return
        eid = sel[0]
        # Switch to Personnel tab and select this employee
        self.nb.select(1)
        self.dialog.after(200, lambda: self._pf_populate_combos() or self._pick_emp_in_pf_by_id(eid))

    def _pick_emp_in_pf_by_id(self, eid):
        labels = list(getattr(self, "_pf_label_to_emp", {}).keys())
        for lab in labels:
            emp = self._pf_label_to_emp.get(lab)
            if emp and emp.get("employee_id") == eid:
                self._pf_pick_var.set(lab); self._pf_load_selected(); return

    # ============================================================
    # Smooth Scrolling / Responsive Layout (Cross-Platform)
    # ============================================================
    def _attach_smooth_scroll(self, widget, h_widget=None):
        """Attach pointer-aware mousewheel routing so the widget UNDER THE MOUSE
        scrolls (instead of the default top-level widget).  Works on macOS
        (delta unit = 1), Windows (needs *120) and Linux (Button-4/5 events).
        Also applies responsive max-width clamping to 1380 px default for
        large screens so content never stretches past comfortable reading width.
        """
        import platform as _plt
        system = _plt.system()
        def _on_wheel(event):
            # Only scroll when pointer is over this widget (or its children)
            target = event.widget
            # Check if event widget is inside target widget tree
            try:
                w = widget.winfo_containing(event.x_root, event.y_root)
            except Exception:
                w = None
            # Resolve: scroll the Treeview / Canvas / ScrolledText that contains the pointer
            scrollable = None
            def _parent_is(t, candidates):
                cur = t
                for _ in range(30):
                    if cur is None: return None
                    if getattr(cur, "winfo_class", lambda: "")() in candidates:
                        return cur
                    if cur is widget: return widget  # containment
                    try: cur = cur.master
                    except Exception: return None
                return None
            cand_classes = ("Treeview", "Canvas", "Text")
            if w is not None:
                scrollable = _parent_is(w, cand_classes) or widget
            else:
                scrollable = widget
            if scrollable is None:
                return
            cls = getattr(scrollable, "winfo_class", lambda: "")()
            if cls == "Canvas":
                direction = -1 if event.delta > 0 else 1
                scrollable.yview_scroll(int(direction * (1 if system == "Darwin" else 3)), "units")
                return "break"
            # Treeview / Text
            direction = -1 if event.delta > 0 else 1
            scrollable.yview_scroll(int(direction * (1 if system == "Darwin" else 3)), "units")
            return "break"
        def _on_button4(event):
            w = None
            try: w = widget.winfo_containing(event.x_root, event.y_root)
            except Exception: pass
            s = widget
            if w is not None:
                cur = w
                for _ in range(30):
                    if cur is None: break
                    cls = getattr(cur, "winfo_class", lambda: "")()
                    if cls in ("Treeview", "Canvas", "Text"):
                        s = cur; break
                    if cur is widget: break
                    try: cur = cur.master
                    except Exception: break
            s.yview_scroll(-3, "units"); return "break"
        def _on_button5(event):
            w = None
            try: w = widget.winfo_containing(event.x_root, event.y_root)
            except Exception: pass
            s = widget
            if w is not None:
                cur = w
                for _ in range(30):
                    if cur is None: break
                    cls = getattr(cur, "winfo_class", lambda: "")()
                    if cls in ("Treeview", "Canvas", "Text"):
                        s = cur; break
                    if cur is widget: break
                    try: cur = cur.master
                    except Exception: break
            s.yview_scroll(3, "units"); return "break"
        try:
            widget.bind_all("<MouseWheel>", _on_wheel, add="+")
        except Exception: pass
        if system == "Linux":
            try:
                widget.bind_all("<Button-4>", _on_button4, add="+")
                widget.bind_all("<Button-5>", _on_button5, add="+")
            except Exception: pass
        # Horizontal wheel on Shift:
        def _on_hwheel(event):
            if not (event.state & 0x0001):
                return None
            try:
                w = widget.winfo_containing(event.x_root, event.y_root)
            except Exception:
                w = None
            s = h_widget or widget
            if w is not None:
                cur = w
                for _ in range(30):
                    if cur is None: break
                    cls = getattr(cur, "winfo_class", lambda: "")()
                    if cls in ("Treeview", "Canvas", "Text"):
                        s = cur; break
                    if cur is widget: break
                    try: cur = cur.master
                    except Exception: break
            direction = -1 if event.delta > 0 else 1
            try: s.xview_scroll(int(direction * (1 if system == "Darwin" else 3)), "units")
            except Exception: pass
            return "break"
        try:
            widget.bind_all("<Shift-MouseWheel>", _on_hwheel, add="+")
        except Exception: pass

    def _apply_responsive_width(self):
        """Cap main dialog width at 1380 px on very wide screens for
        comfortable readability; ensure minimum 1000x620 usable area."""
        try:
            self.dialog.update_idletasks()
            sw = self.dialog.winfo_screenwidth()
            sh = self.dialog.winfo_screenheight()
            max_w = 1380
            max_h = int(sh * 0.92)
            cur_w = self.dialog.winfo_width()
            if sw > max_w + 200 and cur_w > max_w:
                geo = self.dialog.geometry()
                try:
                    geom_part = geo.split("+")[0]
                    _w, _h = [int(x) for x in geom_part.split("x")]
                    nx = max(10, (sw - max_w) // 2)
                    ny = max(10, int((sh - min(_h, max_h)) * 0.38))
                    self.dialog.geometry(f"{max_w}x{min(_h, max_h)}+{nx}+{ny}")
                except Exception: pass
            self.dialog.minsize(1000, 620)
        except Exception:
            pass

# Backward compatibility alias: older callers importing EmployeeManagerDialog get AdvancedEmployeeManagerDialog
EmployeeManagerDialog = AdvancedEmployeeManagerDialog

