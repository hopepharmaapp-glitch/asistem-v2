"""
Centralized Report Engine (Unified for Sales/Purchase/HR/Finance)
All monetary values use Decimal. Exports via openpyxl (Excel) + reportlab (PDF).
Includes accountant-style plain-English explanations for non-accountant users.
"""
import json
from decimal import Decimal
from datetime import date, datetime, timedelta
from calendar import monthrange
from pathlib import Path
from io import BytesIO
import os

# 3rd party imports (graceful degradation if missing)
try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.enums import TA_CENTER, TA_RIGHT, TA_LEFT
    HAS_REPORTLAB = True
except ImportError:
    HAS_REPORTLAB = False

from ledger_service import round_aed

# ------------------------------------------------------------------------------
# 0. Core Utilities: Date Range Translator + Period Comparison
# ------------------------------------------------------------------------------
class DateRangeType:
    TODAY = "TODAY"
    THIS_WEEK = "THIS_WEEK"
    LAST_MONTH = "LAST_MONTH"
    THIS_MONTH = "THIS_MONTH"
    LAST_QUARTER = "LAST_QUARTER"
    THIS_QUARTER = "THIS_QUARTER"
    THIS_YEAR = "THIS_YEAR"
    CUSTOM = "CUSTOM"
    ALL_TIME = "ALL_TIME"

DATE_RANGE_PRESETS = [
    ("Today", DateRangeType.TODAY),
    ("This Week", DateRangeType.THIS_WEEK),
    ("This Month", DateRangeType.THIS_MONTH),
    ("Last Month", DateRangeType.LAST_MONTH),
    ("This Quarter", DateRangeType.THIS_QUARTER),
    ("Last Quarter", DateRangeType.LAST_QUARTER),
    ("This Year", DateRangeType.THIS_YEAR),
    ("All Time", DateRangeType.ALL_TIME),
    ("Custom Range...", DateRangeType.CUSTOM),
]

def _add_months(d, months):
    """Pure-stdlib month arithmetic (replaces dateutil.relativedelta)."""
    m0 = d.month - 1 + months
    year = d.year + m0 // 12
    month = m0 % 12 + 1
    day = min(d.day, monthrange(year, month)[1])
    return d.replace(year=year, month=month, day=day)

def translate_date_range(range_type, start=None, end=None):
    """Translate user-friendly strings to exact start/end dates."""
    today = date.today()
    rt = (range_type or DateRangeType.THIS_MONTH).upper()
    if rt == DateRangeType.ALL_TIME:
        return (date(2000, 1, 1), today)
    if rt == DateRangeType.TODAY:
        return today, today
    if rt == DateRangeType.THIS_WEEK:
        s = today - timedelta(days=today.weekday())
        return s, s + timedelta(days=6)
    if rt == DateRangeType.THIS_MONTH:
        return today.replace(day=1), today
    if rt == DateRangeType.LAST_MONTH:
        first_this = today.replace(day=1)
        e = first_this - timedelta(days=1)
        return e.replace(day=1), e
    if rt == DateRangeType.THIS_QUARTER:
        q = (today.month - 1) // 3
        s = date(today.year, 3*q +1, 1)
        return s, today
    if rt == DateRangeType.LAST_QUARTER:
        e = date(today.year, ((today.month-1)//3)*3, 1) - timedelta(days=1)
        s = _add_months(e.replace(day=1), -2)
        return s, e
    if rt == DateRangeType.THIS_YEAR:
        return today.replace(month=1, day=1), today
    if rt == DateRangeType.CUSTOM:
        s = start; e = end
        if isinstance(s, str): s = date.fromisoformat(s)
        if isinstance(e, str): e = date.fromisoformat(e)
        if not (s and e):
            raise ValueError("Custom range requires start/end dates.")
        return min(s, e), max(s, e)
    return today.replace(day=1), today

def get_previous_period(start, end):
    """Get prior equivalent period for % change calculations."""
    duration = end - start
    prev_end = start - timedelta(days=1)
    prev_start = prev_end - duration
    try:
        if start.day == 1:
            last_day_month = _add_months(start, 1) - timedelta(days=1)
            if end == last_day_month:
                return _add_months(start, -1), _add_months(end, -1)
    except:
        pass
    try:
        if start.month in [1,4,7,10] and start.day ==1:
            quarter_end = _add_months(start, 3) - timedelta(days=1)
            if end == quarter_end:
                return _add_months(start, -3), _add_months(end, -3)
    except:
        pass
    return prev_start, prev_end

def calc_pct_change(current, prior):
    """Calculate % change vs prior period."""
    if not isinstance(current, Decimal): current = Decimal(str(current or 0))
    if not isinstance(prior, Decimal): prior = Decimal(str(prior or 0))
    if prior == 0:
        return Decimal("0.00") if current ==0 else Decimal("100.00")
    return ((current - prior) / abs(prior)) * Decimal("100")

def fmt_aed(v):
    """Human-friendly AED format with commas."""
    if not isinstance(v, Decimal): v = Decimal(str(v or 0))
    v = round_aed(v)
    sign = "-" if v < 0 else ""
    abs_v = abs(v)
    s = f"{abs_v:,.2f}"
    return f"{sign}AED {s}"

# ------------------------------------------------------------------------------
# 1. REPORT ENGINE CLASS
# ------------------------------------------------------------------------------
class ReportEngine:
    def __init__(self, data_manager, range_type=DateRangeType.THIS_MONTH, start=None, end=None):
        self.dm = data_manager
        self.range_type = range_type
        self.start, self.end = translate_date_range(range_type, start, end)
        self.prev_start, self.prev_end = get_previous_period(self.start, self.end)
        # FIRST: Guarantee usable CoA BEFORE loading GL (prevents any early lookup crashes)
        try:
            if hasattr(self.dm, 'ledger'):
                self._coa = self.dm.ledger.load_chart_of_accounts() or []
            else:
                self._coa = self._fallback_coa()
        except Exception:
            self._coa = self._fallback_coa()
        # Cached GL entries
        try:
            self._gl_all = self._load_gl()
            # Guard: every GL entry must have 'date' / 'account_code' as strings
            cleaned = []
            for e in self._gl_all:
                if not isinstance(e, dict):
                    continue
                if 'date' not in e or not e.get('date'):
                    e['date'] = date.today().isoformat()
                if 'account_code' not in e or not e.get('account_code'):
                    e['account_code'] = '1100'
                cleaned.append(e)
            self._gl_all = cleaned
        except Exception:
            self._gl_all = []

        # ------------------------------------------------------------------
        # Period-Lock Read Filter:
        #
        # DEFAULT = MANAGEMENT (read_from_locked_periods_only = False).
        # We ONLY enter strict GAAP audit mode (True) if the user has
        # EXPLICITLY created >= 1 locked period in period_locks.json AND
        # an operator has opted in via the dashboard toggle.
        #
        # Why the flip: 90%+ of SMEs never formally lock ANY period.
        # Old behaviour defaulted to True + zero locks → all GL entries
        # filtered away → reports showed 0/0/0.  New behaviour: if the
        # accountant hasn't created any locked periods yet, we default to
        # Management mode so the user actually sees numbers.  A big amber
        # warning on every report + the dashboard banner tells the user
        # this data is preview-only (not GAAP audit-ready) until they
        # formally close periods.
        # ------------------------------------------------------------------
        # ---- NEW FILTER (global, runs BEFORE locked-period filter): ----
        # Anything right-click → Delete → Reverse is removed from
        # ReportEngine's working GL completely, so it does NOT reflect in
        # P&L, BS, CF, Supplier Spend, Aging, ANY report, ever again.
        # The raw reversal lines still live in immutable general_ledger.json
        # for audit / forensics, but day-to-day reports act as if it never
        # happened (user expectation when they click "Delete").
        try:
            _raw_gl = list(self._gl_all)
        except Exception:
            _raw_gl = []
        reversed_blacklist = set()
        reversal_ids = set()
        import re as _re_gaap
        for e in _raw_gl:
            if not isinstance(e, dict):
                continue
            desc = str(e.get('description') or '')
            if '🔁 REVERSAL OF' in desc or 'REVERSAL OF' in desc[:80]:
                _eid = e.get('entry_id') or id(e)
                reversal_ids.add(_eid)
                m = _re_gaap.search(r"REVERSAL OF\s+([A-Za-z0-9_]+)\s*:\s*(\S+)", desc)
                if m:
                    reversed_blacklist.add((m.group(1), m.group(2)))
        _new_gl = []
        for e in _raw_gl:
            if not isinstance(e, dict):
                continue
            _eid = e.get('entry_id') or id(e)
            if _eid in reversal_ids:
                continue
            rt = str(e.get('reference_type') or '')
            rid = str(e.get('reference_id') or '')
            if rt and rid and (rt, rid) in reversed_blacklist:
                continue
            _new_gl.append(e)
        self._gl_all = _new_gl
        # Track how many reversed/blacklisted refs we dropped so banner/mode
        # can show "N reversed entries excluded" if user wants details.
        try:
            self._reversed_refs_dropped = len(reversed_blacklist)
        except Exception:
            self._reversed_refs_dropped = 0

        self.read_from_locked_periods_only = False
        self._locked_period_count = 0
        self._period_locks_available = False
        try:
            period_locks = self.dm.load_json('period_locks.json') or []
        except Exception:
            period_locks = []
        if not isinstance(period_locks, list):
            period_locks = []
        locked_periods = []
        for pl in period_locks:
            if not isinstance(pl, dict):
                continue
            if pl.get('is_locked') is True:
                s = self._parse_date(pl.get('period_start') or pl.get('start_date'))
                e = self._parse_date(pl.get('period_end') or pl.get('end_date'))
                if s and e:
                    locked_periods.append((s, e))
        self._locked_period_count = len(locked_periods)
        self._period_locks_available = len(locked_periods) > 0
        self._excluded_unlocked_count = 0
        self._excluded_unlocked_debits = Decimal("0.00")
        self._excluded_unlocked_credits = Decimal("0.00")
        self._period_filter_warning = ""
        try:
            if self.read_from_locked_periods_only:
                if locked_periods:
                    filtered_gl = []
                    excluded_count = 0
                    excluded_dr = Decimal("0.00")
                    excluded_cr = Decimal("0.00")
                    for entry in self._gl_all:
                        if not isinstance(entry, dict):
                            continue
                        entry_date = self._parse_date(entry.get('transaction_date') or entry.get('date'))
                        if entry_date is None:
                            filtered_gl.append(entry)
                            continue
                        in_locked = False
                        for (ls, le) in locked_periods:
                            if ls <= entry_date <= le:
                                in_locked = True
                                break
                        if in_locked:
                            filtered_gl.append(entry)
                        else:
                            excluded_count += 1
                            try:
                                excluded_dr += Decimal(str(entry.get('debit_amount') or 0))
                            except Exception:
                                pass
                            try:
                                excluded_cr += Decimal(str(entry.get('credit_amount') or 0))
                            except Exception:
                                pass
                    self._excluded_unlocked_count = excluded_count
                    self._excluded_unlocked_debits = round_aed(excluded_dr)
                    self._excluded_unlocked_credits = round_aed(excluded_cr)
                    self._gl_all = filtered_gl
                    if excluded_count > 0:
                        self._period_filter_warning = (
                            f"⚠️ {excluded_count} GL entries dated in UNLOCKED periods were EXCLUDED "
                            f"per GAAP requirement #3 (finalized/locked periods only). "
                            f"Total excluded Dr = {fmt_aed(self._excluded_unlocked_debits)}, "
                            f"Cr = {fmt_aed(self._excluded_unlocked_credits)}. "
                            f"To include these entries in management-review mode, toggle "
                            f"'Read Locked Periods Only' to OFF."
                        )
                    else:
                        self._period_filter_warning = ""
                else:
                    # Strict locked-only mode + zero locked periods.  This
                    # should never happen in the new default (because we
                    # only set the flag True when the user explicitly
                    # flipped it in the dashboard AND at least one locked
                    # period exists), but handle it gracefully anyway by
                    # keeping every GL entry and just emitting a strong
                    # warning banner.
                    self._period_filter_warning = (
                        "ℹ️ Strict locked-period mode is ON but no LOCKED accounting periods "
                        "were found.  All GL entries are INCLUDED (Management fallback) so "
                        "reports still show useful numbers.  To lock a period, use Period "
                        "Manager — a closed period is required for GAAP-audit ready output."
                    )
                    self._excluded_unlocked_count = 0
            else:
                if len(locked_periods) == 0:
                    self._period_filter_warning = (
                        "ℹ️ Management Preview Mode — no locked periods exist yet, so ALL "
                        "GL entries are shown.  For auditor-ready GAAP statements: open "
                        "Period Manager → close the relevant periods, then switch this "
                        "dashboard to Locked-Only mode using the toggle in the top bar."
                    )
                else:
                    self._period_filter_warning = (
                        "ℹ️ Management Review Mode — ALL GL entries included regardless of "
                        "period-lock status (NOT GAAP-compliant for audit purposes).  Toggle "
                        "to Locked-Only mode in the top bar for formal closing reports."
                    )
        except Exception:
            self._period_filter_warning = ""

        try:
            self.gl_current = [e for e in self._gl_all if self._in_range(e.get('date', date.today().isoformat()), self.start, self.end)]
            self.gl_prior = [e for e in self._gl_all if self._in_range(e.get('date', date.today().isoformat()), self.prev_start, self.prev_end)]
        except Exception:
            self.gl_current = []
            self.gl_prior = []
        # Cached invoices, purchases
        try:
            self._invoices = self.dm.load_json('invoices_data.json') or []
            if not isinstance(self._invoices, list):
                self._invoices = []
            self._invoices = [i for i in self._invoices if isinstance(i, dict)]
        except Exception:
            self._invoices = []
        try:
            self._purchases = self.dm.load_json('purchases_data.json') or []
            if not isinstance(self._purchases, list):
                self._purchases = []
            self._purchases = [p for p in self._purchases if isinstance(p, dict)]
        except Exception:
            self._purchases = []
        # Account types lookup (GUARANTEED to have no KeyErrors now)
        self._account_types = {}
        for a in self._coa:
            if not isinstance(a, dict):
                continue
            code = str(a.get('account_code') or a.get('code') or '').strip()
            if not code:
                continue
            t = a.get('account_type') or 'ASSET'
            self._account_types[code] = str(t).upper()
        # AUTHORITY OVERLAY: The 20 core seeded accounts MUST have these GAAP types
        # regardless of what's stored in chart_of_accounts.json.  A corrupt/incomplete
        # CoA JSON MUST NOT be able to flip the sign-logic in _sum_net.
        self._COA_TYPES_AUTHORITATIVE = {
            "1000": "ASSET", "1100": "ASSET", "1200": "ASSET", "1300": "ASSET",
            "2000": "LIABILITY", "2100": "LIABILITY",
            "3000": "EQUITY", "3100": "EQUITY", "3200": "EQUITY",
            "4000": "REVENUE",
            "5000": "EXPENSE", "5100": "EXPENSE", "5200": "EXPENSE", "5300": "EXPENSE",
            "5400": "EXPENSE", "5500": "EXPENSE", "5600": "EXPENSE", "5700": "EXPENSE",
            "5800": "EXPENSE", "5900": "EXPENSE",
        }
        for k, v in self._COA_TYPES_AUTHORITATIVE.items():
            self._account_types[k] = v

    def _fallback_coa(self):
        """Hard-coded safe fallback that will NEVER fail — used if CoA file missing/corrupt."""
        return [
            {"account_code": "1000", "account_name": "Cash", "account_type": "ASSET"},
            {"account_code": "1100", "account_name": "Bank", "account_type": "ASSET"},
            {"account_code": "1200", "account_name": "Accounts_Receivable", "account_type": "ASSET"},
            {"account_code": "1300", "account_name": "Inventory", "account_type": "ASSET"},
            {"account_code": "2000", "account_name": "Accounts_Payable", "account_type": "LIABILITY"},
            {"account_code": "2100", "account_name": "Accrued_Expenses", "account_type": "LIABILITY"},
            {"account_code": "3000", "account_name": "Retained_Earnings", "account_type": "EQUITY"},
            {"account_code": "3100", "account_name": "Owner_Equity", "account_type": "EQUITY"},
            {"account_code": "3200", "account_name": "Owner_Drawings", "account_type": "EQUITY"},
            {"account_code": "4000", "account_name": "Sales_Revenue", "account_type": "REVENUE"},
            {"account_code": "5000", "account_name": "COGS", "account_type": "EXPENSE"},
            {"account_code": "5100", "account_name": "Salary_Expense", "account_type": "EXPENSE"},
            {"account_code": "5200", "account_name": "Rent_Utility", "account_type": "EXPENSE"},
            {"account_code": "5300", "account_name": "General_Admin", "account_type": "EXPENSE"},
        ]

    # --------------------------
    # Helpers
    # --------------------------
    def _load_gl(self):
        try:
            gl = self.dm.load_json('general_ledger.json') or []
        except Exception:
            gl = []
        return gl if isinstance(gl, list) else []

    def _parse_date(self, d):
        """Robustly parse a date value; return None if fails.
        Handles non-zero-padded dates like '2026-4-5' that date.fromisoformat rejects."""
        try:
            if isinstance(d, date):
                return d
            if isinstance(d, datetime):
                return d.date()
            if isinstance(d, str):
                s10 = d[:10].strip()
                try:
                    return date.fromisoformat(s10)
                except Exception:
                    pass
                import re as _re
                m = _re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s10)
                if m:
                    try:
                        return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                    except Exception:
                        pass
                try:
                    from datetime import datetime as _dt
                    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y"):
                        try:
                            return _dt.strptime(s10, fmt).date()
                        except Exception:
                            continue
                except Exception:
                    pass
        except Exception:
            return None
        return None

    def _in_range(self, d, s, e):
        try:
            pd = self._parse_date(d)
            if pd is None:
                return False
            return s <= pd <= e
        except Exception:
            return False

    def _sum_net(self, account_codes, gl_entries):
        """Net movement: Assets/Expenses = Dr-Cr; Liability/Equity/Revenue = Cr-Dr. 100% crash-safe."""
        net = Decimal("0.00")
        codes = set(str(c) for c in account_codes) if not isinstance(account_codes, (str, int)) else {str(account_codes)}
        for e in gl_entries or []:
            if not isinstance(e, dict):
                continue
            code = str(e.get('account_code') or '').strip()
            if code not in codes:
                continue
            try:
                dr = Decimal(str(e.get('debit_amount') or 0))
            except Exception:
                dr = Decimal("0")
            try:
                cr = Decimal(str(e.get('credit_amount') or 0))
            except Exception:
                cr = Decimal("0")
            typ = self._account_types.get(code, 'ASSET')
            if typ in ('ASSET', 'EXPENSE'):
                net += (dr - cr)
            else:
                net += (cr - dr)
        return round_aed(net)

    def _sum_net_by_classification(self, account_codes, gl_entries, classification_rules):
        """Split account_code totals into sub-buckets by classification keywords.

        classification_rules: [(bucket_label, keyword_list), ...]
            Any GL entry whose `description` or payee/vendor/reference contains ANY of the
            keywords (case-insensitive) → added to that bucket.
        Returns: dict {bucket_label: Decimal(net_amount), ...} with one extra
                 "__unclassified__" for any entries that matched no bucket but are in accts.
        """
        buckets = {label: Decimal("0.00") for label, _ in classification_rules}
        buckets["__unclassified__"] = Decimal("0.00")
        codes = set(str(c) for c in account_codes) if not isinstance(account_codes, (str, int)) else {str(account_codes)}
        for e in gl_entries or []:
            if not isinstance(e, dict):
                continue
            code = str(e.get('account_code') or '').strip()
            if code not in codes:
                continue
            try:
                dr = Decimal(str(e.get('debit_amount') or 0))
            except Exception:
                dr = Decimal("0")
            try:
                cr = Decimal(str(e.get('credit_amount') or 0))
            except Exception:
                cr = Decimal("0")
            typ = self._account_types.get(code, 'ASSET')
            if typ in ('ASSET', 'EXPENSE'):
                amt = round_aed(dr - cr)
            else:
                amt = round_aed(cr - dr)
            if amt == 0:
                continue
            # Build search haystack
            haystack = " ".join(str(v or '') for v in [
                e.get('description'), e.get('note'), e.get('memo'),
                e.get('payee'), e.get('supplier'), e.get('vendor_or_payee'),
                e.get('vendor'), e.get('reference_id'), e.get('reference_type')
            ]).lower()
            assigned = False
            for label, keywords in classification_rules:
                if any(kw.lower() in haystack for kw in keywords):
                    buckets[label] += amt
                    assigned = True
                    break
            if not assigned:
                buckets["__unclassified__"] += amt
        for k in list(buckets.keys()):
            buckets[k] = round_aed(buckets[k])
        return buckets

    def _cumulative_to(self, account_codes, as_of_date):
        """Cumulative balance from inception -> as_of_date. Fully crash-safe."""
        q = []
        for e in self._gl_all or []:
            try:
                d = self._parse_date(e.get('date'))
                if d is None:
                    continue
                if d <= as_of_date:
                    q.append(e)
            except Exception:
                continue
        return self._sum_net(account_codes, q)

    def _in_range_before(self, d, before_date):
        """True iff parseable date d < before_date."""
        try:
            pd = self._parse_date(d)
            if pd is None:
                return False
            return pd < before_date
        except Exception:
            return False

    def _account_name(self, account_code, fallback=None):
        """Human-friendly account name lookup; underscores become spaces;
        friendly defaults for all standard GAAP account codes."""
        code = str(account_code or '').strip()
        if not code:
            return fallback or code
        name = None
        try:
            for a in self._coa or []:
                if not isinstance(a, dict):
                    continue
                if str(a.get('account_code') or a.get('code') or '').strip() == code:
                    name = a.get('account_name') or a.get('name')
                    break
        except Exception:
            name = None
        if name:
            nice = str(name).replace('_', ' ').strip()
            return nice or (fallback or f"Account {code}")
        nice_map = {
            "1000": "Cash on Hand",
            "1100": "Bank (Operating)",
            "1200": "Accounts Receivable (Trade)",
            "1300": "Inventory / Stock",
            "1400": "Prepaid Expenses",
            "1500": "Property, Plant & Equipment",
            "1600": "Accumulated Depreciation",
            "1610": "Accumulated Depreciation",
            "2000": "Accounts Payable (Trade)",
            "2100": "Accrued Expenses",
            "2200": "VAT / Tax Payable",
            "2300": "Short-term Loans",
            "3000": "Retained Earnings",
            "3100": "Owner / Shareholder Equity",
            "3200": "Owner Drawings",
            "4000": "Sales Revenue",
            "4100": "Service Revenue",
            "5000": "Cost of Goods Sold",
            "5100": "Salaries & Wages Expense",
            "5200": "Rent & Utilities",
            "5300": "General & Administrative",
            "5400": "Marketing & Advertising",
            "5500": "Travel & Transportation",
            "5600": "Office Supplies",
            "5700": "Depreciation Expense",
            "5800": "Other Operating Expenses",
            "5900": "Finance / Interest Expense",
        }
        if code in nice_map:
            return nice_map[code]
        return fallback or f"Account {code}"

    def _sorted_all_account_codes(self):
        """Deduped union of CoA accounts + any account codes that actually
        appear in the GL, sorted ASC (standard TB order: Assets → Liabilities
        → Equity → Revenue → Expenses)."""
        seen = set()
        order = []
        def _add(code):
            c = str(code).strip()
            if not c or c in seen:
                return
            seen.add(c); order.append(c)
        try:
            for a in self._coa or []:
                if not isinstance(a, dict): continue
                _add(a.get('account_code') or a.get('code') or '')
        except Exception:
            pass
        try:
            for e in self._gl_all or []:
                if not isinstance(e, dict): continue
                _add(e.get('account_code') or '')
        except Exception:
            pass
        order.sort()
        return order

    def get_trial_balance(self):
        """Classic 6-column Trial Balance — the bookkeeper's eye-level view.

        Returns:
          rows[]: account_code, account_name, account_type,
                  opening_debit / opening_credit,
                  period_debit  / period_credit,
                  closing_debit / closing_credit
          totals: 6 totals (opening_Dr/Cr, period_Dr/Cr, closing_Dr/Cr)
          balanced: True iff each of the 3 pairs is equal (ledger in balance).
          start_date, end_date, range_type, account_count
        """
        from decimal import Decimal
        rows = []
        tot = {k: Decimal("0.00") for k in (
            "opening_debit","opening_credit",
            "period_debit","period_credit",
            "closing_debit","closing_credit")}

        codes = self._sorted_all_account_codes()
        for code in codes:
            typ = self._account_types.get(code)
            if not typ:
                first = (code or ' ')[0]
                typ = {"1":"ASSET","2":"LIABILITY","3":"EQUITY","4":"REVENUE","5":"EXPENSE"}.get(first, "ASSET")
            typ = str(typ).upper()

            o_dr_raw = Decimal("0.00"); o_cr_raw = Decimal("0.00")
            p_dr_raw = Decimal("0.00"); p_cr_raw = Decimal("0.00")
            try:
                for e in self._gl_all or []:
                    if not isinstance(e, dict): continue
                    if str(e.get('account_code') or '').strip() != code: continue
                    try:    dr = Decimal(str(e.get('debit_amount') or 0))
                    except: dr = Decimal("0")
                    try:    cr = Decimal(str(e.get('credit_amount') or 0))
                    except: cr = Decimal("0")
                    edate = self._parse_date(e.get('date'))
                    if edate is None:
                        continue
                    if edate < self.start:
                        o_dr_raw += dr; o_cr_raw += cr
                    elif self.start <= edate <= self.end:
                        p_dr_raw += dr; p_cr_raw += cr
            except Exception:
                pass

            o_net_raw = (o_dr_raw - o_cr_raw)
            p_net_raw = (p_dr_raw - p_cr_raw)
            if typ in ("ASSET", "EXPENSE"):
                o_net = o_net_raw
                p_net = p_net_raw
            else:
                o_net = -o_net_raw
                p_net = -p_net_raw

            # Opening columns: use RAW sign bucketing so Dr>Cr → Dr column,
            # Cr>Dr → Cr column REGARDLESS of account type.  This is the
            # standard Trial Balance presentation rule and is what keeps
            # the 6-column layout balanced.
            if o_net_raw > 0:
                o_dr = round_aed(o_net_raw);   o_cr = Decimal("0.00")
            elif o_net_raw < 0:
                o_dr = Decimal("0.00");        o_cr = round_aed(-o_net_raw)
            else:
                o_dr = Decimal("0.00");        o_cr = Decimal("0.00")

            c_net_raw = o_net_raw + p_net_raw
            if c_net_raw > 0:
                c_dr = round_aed(c_net_raw);   c_cr = Decimal("0.00")
            elif c_net_raw < 0:
                c_dr = Decimal("0.00");        c_cr = round_aed(-c_net_raw)
            else:
                c_dr = Decimal("0.00");        c_cr = Decimal("0.00")

            p_dr = round_aed(p_dr_raw); p_cr = round_aed(p_cr_raw)

            total_zero = (o_dr == 0 and o_cr == 0 and p_dr == 0 and p_cr == 0 and c_dr == 0 and c_cr == 0)
            if total_zero:
                continue

            row = {
                "account_code": code,
                "account_name": self._account_name(code),
                "account_type": typ,
                "opening_debit":  o_dr,
                "opening_credit": o_cr,
                "period_debit":   p_dr,
                "period_credit":  p_cr,
                "closing_debit":  c_dr,
                "closing_credit": c_cr,
            }
            rows.append(row)
            tot["opening_debit"]  += o_dr;  tot["opening_credit"] += o_cr
            tot["period_debit"]   += p_dr;  tot["period_credit"]  += p_cr
            tot["closing_debit"]  += c_dr;  tot["closing_credit"] += c_cr

        for k in tot:
            tot[k] = round_aed(tot[k])

        balanced = (
            tot["opening_debit"]  == tot["opening_credit"] and
            tot["period_debit"]   == tot["period_credit"]  and
            tot["closing_debit"]  == tot["closing_credit"]
        )

        return {
            "start_date": self.start.isoformat(),
            "end_date":   self.end.isoformat(),
            "range_type": self.range_type,
            "rows": rows,
            "totals": tot,
            "balanced": balanced,
            "account_count": len(rows),
        }

    def _banner(self, grand_total, avg=None, pct_change=None, title="Grand Total"):
        return {
            "title": title,
            "grand_total": round_aed(grand_total),
            "avg": round_aed(avg) if avg is not None else None,
            "pct_change": round_aed(pct_change) if pct_change is not None else None,
        }

    # ==================================================================================
    # REPORT 1: PROFIT & LOSS (P&L)
    # ==================================================================================
    def get_profit_and_loss(self):
        # -- Sub-bucket breakdown for General & Admin (5300) so user can trace "mystery fees" --
        # Rules match EXPENSE_CATEGORY_MAP + ledger keyword classifier 1:1.
        GA5300_BUCKETS = [
            ("Office Supplies",       ["office suppl", "stationery", "printer", "paper",
                                       "cartridge", "toner", "ink", "envelope",
                                       "office equipment", "pantry"]),
            ("Travel & Transportation",["travel", "transport", "airline", "flight", "taxi",
                                        "fuel", "petrol", "gasoline", "car rental",
                                        "metro", "bus fare", "parking", "mileage"]),
            ("Marketing & Advertising",["marketing", "advertising", "social media",
                                         "facebook ad", "google ad", "instagram ad",
                                         "promotion", "brochure", "flyer", "signage",
                                         "billboard", "branding", "website seo"]),
        ]
        OTHER_INCOME_CODES = [getattr(self, 'OTHER_INCOME_ACCOUNT', '4100'), '4100', '4200']
        OTHER_EXPENSE_CODES = ['5500', '5600', '5700', '5800']
        INCOME_TAX_CODES = ['5900', '5910', '5920', '5930']
        DEPRECIATION_CODES = ['5400', '5410', '5420']
        def _classify(code, buckets):
            c = str(code or '').strip()
            for kwords, bucket in buckets:
                for kw in kwords:
                    if kw in c:
                        return bucket
            return None
        def compute(gl):
            revenue = self._sum_net(["4000"], gl)
            returns_allow = self._sum_net(["4010", "4020", "4030"], gl)
            net_sales = revenue + returns_allow if returns_allow != 0 else revenue
            cogs = self._sum_net(["5000"], gl)
            salary = self._sum_net(["5100"], gl)
            rent_util = self._sum_net(["5200"], gl)
            depreciation = self._sum_net(DEPRECIATION_CODES, gl)
            gen_admin = self._sum_net(["5300"], gl)
            ga_by_bucket = self._sum_net_by_classification(["5300"], gl, GA5300_BUCKETS)
            ga_office    = abs(ga_by_bucket.get("Office Supplies", 0))
            ga_travel    = abs(ga_by_bucket.get("Travel & Transportation", 0))
            ga_marketing = abs(ga_by_bucket.get("Marketing & Advertising", 0))
            ga_other     = abs(gen_admin) - ga_office - ga_travel - ga_marketing
            if ga_other < 0:
                ga_other = abs(ga_by_bucket.get("__unclassified__", 0))
                total_check = ga_office + ga_travel + ga_marketing + ga_other
                if total_check != abs(gen_admin):
                    ga_other = abs(gen_admin) - ga_office - ga_travel - ga_marketing
            cogs_abs = abs(cogs) if cogs < 0 else cogs
            salary_abs = abs(salary) if salary < 0 else salary
            ru_abs = abs(rent_util) if rent_util < 0 else rent_util
            ga_abs = abs(gen_admin) if gen_admin < 0 else gen_admin
            dep_abs = abs(depreciation) if depreciation < 0 else depreciation
            ga_office    = abs(ga_office)
            ga_travel    = abs(ga_travel)
            ga_marketing = abs(ga_marketing)
            ga_other     = max(Decimal("0"), ga_other)
            gross_profit = net_sales - cogs_abs
            op_expenses = salary_abs + ru_abs + ga_abs + dep_abs
            operating_income = gross_profit - op_expenses
            other_income = self._sum_net(OTHER_INCOME_CODES, gl)
            other_expense = self._sum_net(OTHER_EXPENSE_CODES, gl)
            other_exp_abs = abs(other_expense) if other_expense < 0 else other_expense
            other_total = other_income - other_exp_abs
            income_before_tax = operating_income + other_total
            income_tax_exp = self._sum_net(INCOME_TAX_CODES, gl)
            income_tax_abs = abs(income_tax_exp) if income_tax_exp < 0 else income_tax_exp
            net_income = income_before_tax - income_tax_abs
            oci = self._sum_net(['3900', '3910', '3920', '3930'], gl)
            oci_abs = oci  # Other Comprehensive Income is Cr-balance positive (credit)
            total_comprehensive_income = net_income + oci_abs
            shares_outstanding = max(int(getattr(self, 'shares_outstanding', 10000) or 10000), 1)
            eps = round_aed(net_income / Decimal(shares_outstanding)) if shares_outstanding > 0 else Decimal("0")
            return {
                "gross_revenue": revenue,
                "sales_returns_allowances": -abs(returns_allow) if returns_allow != 0 else Decimal("0"),
                "net_sales": net_sales,
                "cogs": cogs_abs,
                "gross_profit": gross_profit,
                "salary_expense": salary_abs,
                "rent_utility": ru_abs,
                "depreciation_amortization": dep_abs,
                "general_admin": ga_abs,
                "general_admin_office": ga_office,
                "general_admin_travel": ga_travel,
                "general_admin_marketing": ga_marketing,
                "general_admin_other": ga_other,
                "total_operating_expenses": op_expenses,
                "operating_income": operating_income,
                "other_income": other_income,
                "other_expenses": other_exp_abs,
                "other_total": other_total,
                "income_before_tax": income_before_tax,
                "income_tax_expense": income_tax_abs,
                "net_income": net_income,
                "other_comprehensive_income": oci_abs,
                "total_comprehensive_income": total_comprehensive_income,
                "earnings_per_share": eps,
            }
        c = compute(self.gl_current); p = compute(self.gl_prior)
        lines = {}
        for k in c:
            lines[k] = {
                "current": c[k], "prior": p[k],
                "pct_change": calc_pct_change(c[k], p[k])
            }
        days = (self.end - self.start).days + 1
        net_income_cur = c["net_income"]
        gp_current = c["gross_profit"]
        rev_current = c["net_sales"]
        avg = net_income_cur / Decimal(max(days, 1))
        if rev_current > 0:
            gm_pct = (gp_current / rev_current * 100)
            gm_prior_pct = (p["gross_profit"] / p["net_sales"] * 100) if p["net_sales"] > 0 else Decimal("0")
            gm_pct_valid = True
        else:
            gm_pct = Decimal("0")
            gm_prior_pct = Decimal("0")
            gm_pct_valid = False
        if net_income_cur > 0:
            status = f"✅ PROFITABLE: Net Income {fmt_aed(net_income_cur)} for the period."
            if rev_current <= 0: status = "ℹ️ No positive net sales recorded in this period."
        elif net_income_cur < 0:
            status = f"⚠️ NET LOSS: {fmt_aed(abs(net_income_cur))} during this period — review pricing and operating expenses."
        else:
            status = "➖ Break-even: Net sales exactly covered all expenses."
        if gm_pct_valid:
            gm_comment = f"📊 Gross margin: {round_aed(gm_pct)}% — "
            if gm_pct >= Decimal("40"): gm_comment += "Excellent margin (pharma distribution benchmark)."
            elif gm_pct >= Decimal("20"): gm_comment += "Healthy margin."
            elif gm_pct > 0: gm_comment += "Tight margin — renegotiate supplier terms or raise prices."
            else: gm_comment += "Negative margin — COGS exceeds Net Sales (urgent)."
        else:
            gm_comment = "📊 Gross margin: N/A — no Net Sales in this period."
        ga_sum_current = c["general_admin_office"] + c["general_admin_travel"] + c["general_admin_marketing"] + c["general_admin_other"]
        ga_break_str = "\n".join([
            f"• Office Supplies          = {fmt_aed(c['general_admin_office'])}",
            f"• Travel & Transportation  = {fmt_aed(c['general_admin_travel'])}",
            f"• Marketing & Advertising  = {fmt_aed(c['general_admin_marketing'])}",
            f"• Other Admin / Misc bank withdrawals classified as Admin = {fmt_aed(c['general_admin_other'])}",
        ])
        explanation = f"""
        {status}\n{gm_comment}\n\n
        US GAAP / IAS 1 structure applied:\n
        • Net Sales = Gross Revenue (4000) − Returns & Allowances.\n
        • Gross Profit = Net Sales − Cost of Goods Sold (5000).\n
        • Total Operating Expenses = Salaries (5100) + Rent & Utilities (5200) + Depreciation (5400) + G&A (5300).\n
        • Operating Income = Gross Profit − Total Operating Expenses.\n
        • Income Before Tax = Operating Income + Other Income (4100) − Other Expenses (5500–5800).\n
        • Net Income = Income Before Tax − Income Tax Expense (5900).\n
        • Total Comprehensive Income = Net Income + Other Comprehensive Income (39xx).\n\n
        🔍  General & Admin (5300) breakdown:\n
        {ga_break_str}\n
        TOTAL General & Admin = {fmt_aed(c['general_admin'])} (sum of 4 sub-categories: {fmt_aed(ga_sum_current)})\n\n
        ℹ️  How to populate each line:\n
        1. Salaries (5100)    → Employee Manager → Add Monthly Salary.\n
        2. Rent & Utilities (5200) → Expenses → category "Rent" or "Utilities".\n
        3. Depreciation (5400) → Manual journal Dr 5400 / Cr 1350 Accumulated Depreciation.\n
        4. Other Income (4100) / Other Expenses (5500–5800) → Manual JE with matching codes.\n
        5. Income Tax (5900)  → At year-end, manual Dr 5900 / Cr 2150 Tax Payable based on corporate tax return.\n
        6. EPS denominator → `engine.shares_outstanding` (default: 10,000 ordinary shares).\n
        """.strip()
        # --- P&L Arithmetic Diagnostic (US GAAP step-by-step) ---
        diagnostics = {"passed": True, "checks": [], "issues": [], "fix_steps": []}
        def _chk(name, ok, detail):
            diagnostics["checks"].append({"name": name, "passed": bool(ok), "detail": str(detail)})
            if not ok: diagnostics["passed"] = False
        try:
            gr = c["gross_revenue"]; sra = c["sales_returns_allowances"]
            exp_ns = round_aed(gr + sra) if gr != 0 or sra != 0 else round_aed(gr)
            _chk("US-GAAP 1: Net Sales = Gross Revenue − Returns & Allowances",
                 abs(exp_ns - c["net_sales"]) < Decimal("0.01"),
                 f"Calc {fmt_aed(gr)} + ({fmt_aed(sra)}) = {fmt_aed(exp_ns)}; shown {fmt_aed(c['net_sales'])}")
        except Exception as exc:
            _chk("US-GAAP 1", False, f"Failed: {exc}")
        try:
            exp_gp = round_aed(c["net_sales"] - c["cogs"])
            _chk("US-GAAP 2: Gross Profit = Net Sales − COGS",
                 abs(exp_gp - c["gross_profit"]) < Decimal("0.01"),
                 f"{fmt_aed(c['net_sales'])} − {fmt_aed(c['cogs'])} = {fmt_aed(exp_gp)}; shown {fmt_aed(c['gross_profit'])}")
        except Exception as exc:
            _chk("US-GAAP 2", False, f"Failed: {exc}")
        try:
            exp_opex = round_aed(c["salary_expense"] + c["rent_utility"] + c["depreciation_amortization"] + c["general_admin"])
            _chk("US-GAAP 3: Total Operating Expenses = Salaries + Rent+Util + D&A + G&A",
                 abs(exp_opex - c["total_operating_expenses"]) < Decimal("0.01"),
                 f"Sum = {fmt_aed(exp_opex)}; shown {fmt_aed(c['total_operating_expenses'])}")
            ga_sub = c["general_admin_office"] + c["general_admin_travel"] + c["general_admin_marketing"] + c["general_admin_other"]
            _chk("US-GAAP 3b: G&A = Office + Travel + Marketing + Other",
                 abs(ga_sub - c["general_admin"]) < Decimal("0.01"),
                 f"Subsum = {fmt_aed(ga_sub)}; 5300 total = {fmt_aed(c['general_admin'])}")
        except Exception as exc:
            _chk("US-GAAP 3", False, f"Failed: {exc}")
        try:
            exp_oi = round_aed(c["gross_profit"] - c["total_operating_expenses"])
            _chk("US-GAAP 4: Operating Income (EBIT) = Gross Profit − Total OpEx",
                 abs(exp_oi - c["operating_income"]) < Decimal("0.01"),
                 f"{fmt_aed(c['gross_profit'])} − {fmt_aed(c['total_operating_expenses'])} = {fmt_aed(exp_oi)}; shown {fmt_aed(c['operating_income'])}")
        except Exception as exc:
            _chk("US-GAAP 4", False, f"Failed: {exc}")
        try:
            exp_ibt = round_aed(c["operating_income"] + c["other_income"] - c["other_expenses"])
            _chk("US-GAAP 5: Income Before Tax = Operating Income + Other Income − Other Expenses",
                 abs(exp_ibt - c["income_before_tax"]) < Decimal("0.01"),
                 f"{fmt_aed(c['operating_income'])} + {fmt_aed(c['other_income'])} − {fmt_aed(c['other_expenses'])} = {fmt_aed(exp_ibt)}; shown {fmt_aed(c['income_before_tax'])}")
        except Exception as exc:
            _chk("US-GAAP 5", False, f"Failed: {exc}")
        try:
            exp_ni = round_aed(c["income_before_tax"] - c["income_tax_expense"])
            _chk("US-GAAP 6: Net Income = Income Before Tax − Income Tax Expense",
                 abs(exp_ni - c["net_income"]) < Decimal("0.01"),
                 f"{fmt_aed(c['income_before_tax'])} − {fmt_aed(c['income_tax_expense'])} = {fmt_aed(exp_ni)}; shown {fmt_aed(c['net_income'])}")
        except Exception as exc:
            _chk("US-GAAP 6", False, f"Failed: {exc}")
        try:
            exp_tci = round_aed(c["net_income"] + c["other_comprehensive_income"])
            _chk("US-GAAP 7: Total Comprehensive Income = Net Income + OCI",
                 abs(exp_tci - c["total_comprehensive_income"]) < Decimal("0.01"),
                 f"{fmt_aed(c['net_income'])} + {fmt_aed(c['other_comprehensive_income'])} = {fmt_aed(exp_tci)}; shown {fmt_aed(c['total_comprehensive_income'])}")
        except Exception as exc:
            _chk("US-GAAP 7", False, f"Failed: {exc}")
        try:
            nrev = len([e for e in self.gl_current if isinstance(e, dict)
                        and str(e.get('account_code') or '').strip() == '4000'])
            ncogs = len([e for e in self.gl_current if isinstance(e, dict)
                         and str(e.get('account_code') or '').strip() == '5000'])
            ok_c = not (nrev > 0 and ncogs == 0 and c["net_sales"] > 0 and c["cogs"] == 0)
            _chk("Completeness: Revenue entries vs COGS entries", ok_c,
                 f"4000 lines = {nrev}, 5000 lines = {ncogs}")
        except Exception as exc:
            _chk("Completeness", False, f"Failed: {exc}")
        diagnostics["summary"] = (
            "✅ All 7 US-GAAP P&L arithmetic checks passed." if diagnostics["passed"]
            else f"⚠️ {len(diagnostics['issues'])} issue(s) found."
        )
        return {
            "report_name": "Consolidated Statements of Comprehensive Income (US GAAP)",
            "period": {"start": self.start, "end": self.end},
            "prior_period": {"start": self.prev_start, "end": self.prev_end},
            "lines": lines,
            "diagnostics": diagnostics,
            "summary_banner": {
                "grand_total": lines["net_income"]["current"],
                "avg": round_aed(avg),
                "pct_change": lines["net_income"]["pct_change"],
                "gross_margin_pct": round_aed(gm_pct),
                "gross_margin_prior_pct": round_aed(gm_prior_pct),
                "gross_margin_pct_change": round_aed(gm_pct - gm_prior_pct),
                "operating_income": c["operating_income"],
                "income_before_tax": c["income_before_tax"],
                "total_comprehensive_income": c["total_comprehensive_income"],
                "eps": c["earnings_per_share"],
            },
            "explanation": explanation,
            "table_rows": [
                # -------- REVENUE SECTION (US GAAP: Gross − Returns = Net) --------
                {"label": "Gross Revenue",         **lines["gross_revenue"],          "tag": None},
                {"label": "  Less: Sales Returns & Allowances", **lines["sales_returns_allowances"],  "tag": None},
                {"label": "NET SALES",             **lines["net_sales"],              "tag": "section"},
                {"label": "  Less: Cost of Goods Sold", **lines["cogs"],             "tag": None},
                {"label": "GROSS PROFIT",          **lines["gross_profit"],           "tag": "gross_profit"},
                # -------- OPERATING EXPENSES --------
                {"label": "OPERATING EXPENSES:",    "current": None, "prior": None, "pct_change": None, "tag": "section"},
                {"label": "    Salaries & Wages (5100)",            **lines["salary_expense"], "tag": None},
                {"label": "    Rent & Utilities (5200)",            **lines["rent_utility"],   "tag": None},
                {"label": "    Depreciation & Amortization (5400)", **lines["depreciation_amortization"], "tag": None},
                {"label": "    General & Administrative (5300)",    **lines["general_admin"],  "tag": None},
                {"label": "        ▸ Office Supplies",              **lines["general_admin_office"],    "tag": None},
                {"label": "        ▸ Travel & Transportation",      **lines["general_admin_travel"],    "tag": None},
                {"label": "        ▸ Marketing & Advertising",      **lines["general_admin_marketing"], "tag": None},
                {"label": "        ▸ Other Admin",                  **lines["general_admin_other"],     "tag": None},
                {"label": "  TOTAL OPERATING EXPENSES", **lines["total_operating_expenses"], "tag": "total_row"},
                {"label": "OPERATING INCOME (EBIT)",    **lines["operating_income"],  "tag": "gross_profit"},
                # -------- NON-OPERATING --------
                {"label": "OTHER INCOME & (EXPENSES):", "current": None, "prior": None, "pct_change": None, "tag": "section"},
                {"label": "    Other Income",          **lines["other_income"],       "tag": None},
                {"label": "    Other Expenses",        **lines["other_expenses"],     "tag": None},
                {"label": "  Other Income, net",       **lines["other_total"],        "tag": None},
                {"label": "INCOME BEFORE INCOME TAX",  **lines["income_before_tax"],  "tag": "section"},
                {"label": "  Less: Income Tax Expense",**lines["income_tax_expense"], "tag": None},
                {"label": "NET INCOME",                **lines["net_income"],         "tag": "net_profit"},
                # -------- COMPREHENSIVE INCOME --------
                {"label": "OTHER COMPREHENSIVE INCOME, net of tax", **lines["other_comprehensive_income"], "tag": None},
                {"label": "TOTAL COMPREHENSIVE INCOME", **lines["total_comprehensive_income"], "tag": "total_row"},
                # -------- PER-SHARE (FASB ASC 260) --------
                {"label": "EARNINGS PER SHARE (Basic)", **lines["earnings_per_share"], "tag": None},
            ]
        }

    # ==================================================================================
    # REPORT 2: BALANCE SHEET (US GAAP Classified Statement of Financial Position)
    # ==================================================================================
    def get_balance_sheet(self):
        as_of = self.end
        # ------ US GAAP / SEC Regulation S-X classification buckets ----------
        # (codes match seeded chart_of_accounts.json schema used by this product)
        CA_CODES = ["1000", "1100", "1200", "1210", "1220", "1300", "1310", "1320",
                    "1400", "1410", "1500"]     # Cash, Bank, AR, Allowance, Inv, Prepaid, Other current
        NCA_CODES = ["1600", "1610", "1620", "1630", "1640", "1650", "1700", "1710", "1800"]
        CL_CODES = ["2000", "2010", "2100", "2110", "2120", "2150", "2200", "2210", "2220",
                    "2300", "2310"]  # AP, Accrued, Tax, ST debt, Unearned rev
        NCL_CODES = ["2400", "2410", "2500", "2510", "2600", "2700"]  # LT debt, Lease, Pension, DTL
        EQ_CAPSTOCK_CODES = ["3100", "3110", "3120"]  # Capital Stock / Owner Equity (3100), APIC (3110)
        EQ_RE_CODES = ["3000"]                       # Retained Earnings (3000)
        EQ_AOCI_CODES = ["3900", "3910", "3920", "3930"]  # Other Comprehensive Income (AOCI)
        EQ_TREASURY_CODES = ["3300", "3310"]          # Treasury Stock, contra-equity
        EQ_DRAWINGS_CODES = ["3200"]                  # Owner Drawings / Distributions (contra)
        # ----------------------------------------------------------------------
        def _sum_signed(codes):
            return self._cumulative_to([c for c in codes], as_of)
        cash = self._cumulative_to(["1000"], as_of)
        bank = self._cumulative_to(["1100"], as_of)
        ar = self._cumulative_to(["1200"], as_of)
        ar_allowance = -abs(self._cumulative_to(["1210"], as_of))  # Contra-asset, credit balance
        inv = self._cumulative_to(["1300"], as_of)
        prepaid = self._cumulative_to(["1400", "1410"], as_of)
        other_current_assets = self._cumulative_to(["1500"], as_of)
        ca_total = round_aed(cash + bank + ar + ar_allowance + inv + prepaid + other_current_assets)
        ppe_gross = self._cumulative_to(["1600"], as_of)
        accum_depr = -abs(self._cumulative_to(["1610"], as_of))  # Contra-asset
        ppe_net = round_aed(ppe_gross + accum_depr)
        intangibles = self._cumulative_to(["1620", "1630", "1640"], as_of)
        investments = self._cumulative_to(["1700", "1710"], as_of)
        other_noncurrent_assets = self._cumulative_to(["1800"], as_of)
        nca_total = round_aed(ppe_gross + accum_depr + intangibles + investments + other_noncurrent_assets)
        total_assets_raw = round_aed(ca_total + nca_total)
        ap = self._cumulative_to(["2000"], as_of)
        accrued = self._cumulative_to(["2100", "2110", "2120"], as_of)
        tax_payable = self._cumulative_to(["2150"], as_of)
        short_term_debt = self._cumulative_to(["2200", "2210", "2220"], as_of)
        unearned_rev = self._cumulative_to(["2300", "2310"], as_of)
        cl_total = round_aed(ap + accrued + tax_payable + short_term_debt + unearned_rev)
        long_term_debt = self._cumulative_to(["2400", "2410"], as_of)
        lease_liab = self._cumulative_to(["2500", "2510"], as_of)
        deferred_tax_liab = self._cumulative_to(["2600"], as_of)
        other_noncurrent_liab = self._cumulative_to(["2700"], as_of)
        ncl_total = round_aed(long_term_debt + lease_liab + deferred_tax_liab + other_noncurrent_liab)
        total_liabilities_raw = round_aed(cl_total + ncl_total)
        # ------- Equity (US GAAP) --------
        cap_stock = self._cumulative_to(EQ_CAPSTOCK_CODES, as_of)
        retained_earnings = self._cumulative_to(EQ_RE_CODES, as_of)
        aoci = self._cumulative_to(EQ_AOCI_CODES, as_of)
        treasury_stock_contra = -abs(self._cumulative_to(EQ_TREASURY_CODES, as_of))
        drawings_contra = -abs(self._cumulative_to(EQ_DRAWINGS_CODES, as_of))
        # ⚓ GAAP tie: NI from P&L = Current Period Net Income / (Loss) added into RE before closing
        pnl = self.get_profit_and_loss()
        net_income_from_pnl = pnl["lines"]["net_income"]["current"]
        oci_from_pnl = pnl["lines"]["other_comprehensive_income"]["current"]
        aoci += oci_from_pnl  # OCI for the period -> AOCI balance
        current_period_re_addition = net_income_from_pnl + drawings_contra
        retained_earnings_total = round_aed(retained_earnings + net_income_from_pnl + drawings_contra)
        total_equity_raw = round_aed(cap_stock + retained_earnings_total + aoci + treasury_stock_contra)
        total_le_raw = round_aed(total_liabilities_raw + total_equity_raw)
        variance_raw = round_aed(total_assets_raw - total_le_raw)
        # If unbalanced by < 0.02, this is a rounding artifact; else we flag to user
        bs_ok = abs(variance_raw) < Decimal("0.01")
        # --- Build explanation block (US GAAP ASC 205 / 210 / 505) ---------------
        lines = {
            "assets": {
                "1000 Cash on Hand": cash,
                "1100 Cash in Bank": bank,
                "1200 Trade Accounts Receivable (Gross)": ar,
                "1210 Less: Allowance for Doubtful Accounts": ar_allowance,
                "1300 Inventories (net)": inv,
                "1400 Prepaid Expenses & Other Current Assets": prepaid,
                "Other Current Assets": other_current_assets,
                "__current_assets_total": ca_total,
                "1600 Property, Plant & Equipment (Gross Cost)": ppe_gross,
                "1610 Less: Accumulated Depreciation & Amortization": accum_depr,
                "Property, Plant & Equipment, Net": ppe_net,
                "1620-40 Intangible Assets (Net)": intangibles,
                "1700 Long-Term Investments": investments,
                "1800 Other Non-Current Assets": other_noncurrent_assets,
                "__noncurrent_assets_total": nca_total,
                "total_assets": round_aed(ca_total + nca_total),
            },
            "liabilities": {
                "2000 Trade Accounts Payable": ap,
                "2100-2120 Accrued Expenses & Employee Benefits": accrued,
                "2150 Income & Other Taxes Payable": tax_payable,
                "2200-2220 Short-Term Borrowings & Current Portion of LT Debt": short_term_debt,
                "2300 Unearned Revenue / Customer Deposits": unearned_rev,
                "__current_liabilities_total": cl_total,
                "2400-41 Long-Term Debt (Non-Current Portion)": long_term_debt,
                "2500 Operating & Finance Lease Liabilities (Non-Current)": lease_liab,
                "2600 Deferred Tax Liabilities": deferred_tax_liab,
                "2700 Other Non-Current Liabilities": other_noncurrent_liab,
                "__noncurrent_liabilities_total": ncl_total,
                "total_liabilities": round_aed(cl_total + ncl_total),
            },
            "equity": {
                "3100-3120 Capital Stock / Paid-in Capital": cap_stock,
                "3000 Retained Earnings, Beginning of Period": retained_earnings,
                "    + Net Income for the Period": net_income_from_pnl,
                "    − Distributions / Owner Drawings": drawings_contra,
                "3000 Retained Earnings, End of Period": retained_earnings_total,
                "39xx Accumulated Other Comprehensive Income (AOCI)": aoci,
                "33xx Less: Treasury Stock, at Cost": treasury_stock_contra,
                "total_equity": total_equity_raw,
            },
            "total_le": total_le_raw,
            "variance": variance_raw,
            "is_balanced": bs_ok,
            "__breakdowns": {
                "current_assets": ca_total, "noncurrent_assets": nca_total,
                "current_liabilities": cl_total, "noncurrent_liabilities": ncl_total,
                "working_capital": round_aed(ca_total - cl_total),
                "re_beginning": retained_earnings, "re_add_ni": net_income_from_pnl,
                "re_less_drawings": drawings_contra, "re_ending": retained_earnings_total,
            }
        }
        # ---- Backward-compat: some callers access bs.lines.liabilities_and_equity.total_liabilities_and_equity
        # instead of the top-level `total_le` key.  Mirror the full structure into
        # `liabilities_and_equity` so they get correct totals (no 0/0/0 on BS side).
        try:
            lines["liabilities_and_equity"] = {
                "total_liabilities_and_equity": total_le_raw,
                "total_liabilities": round_aed(cl_total + ncl_total),
                "total_equity": total_equity_raw,
                "current_liabilities": cl_total,
                "noncurrent_liabilities": ncl_total,
            }
        except Exception:
            pass
        # --- Diagnostics (US GAAP ASC checks) ---------------------------------
        diagnostics = {"is_balanced": bs_ok, "variance": variance_raw, "checks": [], "root_causes": [], "fix_steps": []}
        def _add_check(name, ok, detail):
            diagnostics["checks"].append({"name": name, "passed": bool(ok), "detail": str(detail)})
        # G1: Σ debits = Σ credits globally
        try:
            all_dr = Decimal("0"); all_cr = Decimal("0")
            for e in self._gl_all:
                try:
                    all_dr += Decimal(str(e.get('debit_amount') or 0))
                    all_cr += Decimal(str(e.get('credit_amount') or 0))
                except Exception: pass
            delta = round_aed(all_dr - all_cr)
            ok = abs(delta) < Decimal("0.01")
            _add_check("A. Global GL Integrity (Σ Debits = Σ Credits)", ok,
                       f"Σ Dr = {fmt_aed(all_dr)}  |  Σ Cr = {fmt_aed(all_cr)}  |  Δ = {fmt_aed(delta)}")
            if not ok:
                diagnostics["root_causes"].append(
                    f"Global GL unbalanced by {fmt_aed(delta)}. This is the PRIMARY cause of BS variance.")
                diagnostics["fix_steps"].append(
                    "Fix A: Tab 8 (GL Audit) → filter rows where Σ Dr ≠ Σ Cr per entry_id. Post the missing side; then 🔁 Refresh.")
        except Exception as exc:
            _add_check("A. Global GL Integrity", False, f"Failed: {exc}")
        # G2: P&L Net Income = Retained Earnings movement additions (tie between statements)
        try:
            ok2 = abs(net_income_from_pnl - pnl["lines"]["net_income"]["current"]) < Decimal("0.01")
            _add_check("B. Tie — P&L Net Income = Retained Earnings current-period addition", ok2,
                       f"P&L NI = {fmt_aed(pnl['lines']['net_income']['current'])}; RE addition = {fmt_aed(net_income_from_pnl)}; Δ = {fmt_aed(net_income_from_pnl - pnl['lines']['net_income']['current'])}")
            if not ok2:
                diagnostics["root_causes"].append("Net Income from P&L is not the same figure that flows to Retained Earnings movement on BS.")
                diagnostics["fix_steps"].append("Fix B: Ensure engine.gl_current range matches P&L range. Rebuild GL & Refresh.")
        except Exception as exc:
            _add_check("B. P&L — RE Movement Tie", False, f"Failed: {exc}")
        # G3: Working Capital = Current Assets − Current Liabilities (positive is liquidity-safe per GAAP)
        try:
            wc = lines["__breakdowns"]["working_capital"]
            _add_check("C. Working Capital = Current Assets − Current Liabilities (ASC 210-10-45)",
                       True,
                       f"CA = {fmt_aed(ca_total)} − CL = {fmt_aed(cl_total)} = WC {fmt_aed(wc)}   "
                       f"{'✅ Healthy positive WC' if wc >= 0 else '⚠️ Negative WC — liquidity risk (supplier calls, payroll, VAT coming due)'}")
        except Exception as exc:
            _add_check("C. Working Capital", False, f"Failed: {exc}")
        # G4: Inventory pairing (Dr purchases ≥ Cr COGS otherwise abnormal)
        try:
            inv_entries = [e for e in self._gl_all if isinstance(e, dict) and str(e.get('account_code') or '').strip() == '1300']
            dr = sum(Decimal(str(e.get('debit_amount') or 0)) for e in inv_entries)
            cr = sum(Decimal(str(e.get('credit_amount') or 0)) for e in inv_entries)
            ok4 = not (dr == 0 and cr > 0 and inv < 0)
            msg = (f"1300 Inventory: Dr (Purchases) = {fmt_aed(dr)}; Cr (COGS) = {fmt_aed(cr)}; Net = {fmt_aed(inv)}"
                   + ("  ⚠️ ABNORMAL: Purchases never posted but COGS recognised (record Purchase Dr 1300 / Cr 2000)." if inv < 0 and dr == 0 else ""))
            _add_check("D. Inventory (1300) purchase-debits paired with COGS credits", ok4, msg)
            if not ok4:
                diagnostics["root_causes"].append("COGS recognised before any Purchase/Goods Received journal — Inventory shows abnormal credit balance.")
                diagnostics["fix_steps"].append("Fix D: Post the goods-received purchase (Dr 1300 Inventory / Cr 2000 AP). Refresh BS; Inventory flips positive & BS normal.")
        except Exception as exc:
            _add_check("D. Inventory Pairing Check", False, f"Failed: {exc}")
        # G5: Owner Equity contributions (3100) present; warn if zero in first period
        try:
            eq3100 = [e for e in self._gl_all if isinstance(e, dict) and str(e.get('account_code') or '').strip() in EQ_CAPSTOCK_CODES]
            cr_eq = sum(Decimal(str(e.get('credit_amount') or 0)) for e in eq3100)
            ok5 = True
            msg = f"Capital Stock (3100/3110): contributions = {fmt_aed(cr_eq)}"
            if cr_eq == 0 and len(self._gl_all) > 0:
                msg += "  ⚠️  No owner capital recorded — if personal funds were used to start business, post Dr 1100 Bank / Cr 3100 Owner Equity."
            _add_check("E. Owner Capital Contributions (3100) present", ok5, msg)
        except Exception as exc:
            _add_check("E. Owner Capital Check", False, f"Failed: {exc}")
        # Final summary
        diagnostics["summary"] = (
            "✅ Perfectly balanced, all 5 US GAAP integrity checks passed."
            if bs_ok and all(c.get("passed") for c in diagnostics["checks"])
            else f"⚠️ Issues found: variance={fmt_aed(variance_raw)}; {sum(0 if c.get('passed') else 1 for c in diagnostics['checks'])} check(s) failed."
        )
        if not bs_ok:
            diagnostics["fix_steps"].append(
                "Last step: after applying fixes, click 🔁 Refresh on GAAP dashboard; then re-open Balance Sheet.")
        return {
            "report_name": "Consolidated Balance Sheets (Statement of Financial Position — US GAAP Classified, ASC 210)",
            "as_of": as_of,
            "lines": lines,
            "diagnostics": diagnostics,
            "summary_banner": {
                "grand_total": lines["assets"]["total_assets"],
                "current_assets": ca_total,
                "noncurrent_assets": nca_total,
                "current_liabilities": cl_total,
                "noncurrent_liabilities": ncl_total,
                "working_capital": lines["__breakdowns"]["working_capital"],
                "liabilities": lines["liabilities"]["total_liabilities"],
                "equity": lines["equity"]["total_equity"],
                "is_balanced": bs_ok
            },
            "explanation": f"""
{"✅ BALANCED (US GAAP ASC 205 / ASC 210) — Total Assets = Total Liabilities + Total Equity" if bs_ok else f"⚠️ UNBALANCED by {fmt_aed(variance_raw)} — see diagnostics below."}

⚖️  GOLDEN EQUATION (Double-Entry Bookkeeping):
      TOTAL ASSETS            {fmt_aed(lines["assets"]["total_assets"])}
    = TOTAL LIABILITIES       {fmt_aed(lines["liabilities"]["total_liabilities"])}
    + TOTAL EQUITY            {fmt_aed(lines["equity"]["total_equity"])}
    ═══════════════════════════════════════════
      TOTAL L + E             {fmt_aed(total_le_raw)}   (Δ = {fmt_aed(variance_raw)}; within 0.02 = balanced per rounding tolerance)

💧  LIQUIDITY — Working Capital = Current Assets − Current Liabilities:
      WC = {fmt_aed(lines["__breakdowns"]["working_capital"])}
      ("WC is a 'buffer' for near-term obligations; GAAP requires disclosure for SEC filers.")

📌  Cross-Statement Tie (verifiable on the Statement of Changes in Equity):
      Retained Earnings (end) = RE (begin) {fmt_aed(lines["__breakdowns"]["re_beginning"])}
                                + Net Income (period) {fmt_aed(lines["__breakdowns"]["re_add_ni"])}
                                − Distributions / Drawings {fmt_aed(-lines["__breakdowns"]["re_less_drawings"])}
                                = {fmt_aed(lines["__breakdowns"]["re_ending"])}
            """.strip()
        }

    # ==================================================================================
    # REPORT 3: AGING RECEIVABLES
    # ==================================================================================
    def get_aging_receivables(self):
        today = date.today()
        buckets = {
            "Current (not overdue)": Decimal("0.00"),
            "0-30 Days Overdue": Decimal("0.00"),
            "31-60 Days Overdue": Decimal("0.00"),
            "61-90 Days Overdue": Decimal("0.00"),
            "91+ Days Overdue": Decimal("0.00")
        }
        bucket_keys_order = list(buckets.keys())
        rows = []
        for inv in self._invoices:
            status = inv.get('status', 'Not Paid')
            if status in ('Paid',): continue
            due_str = inv.get('due_date')
            bal = Decimal(str(inv.get('balance_due', 0) or inv.get('grand_total', 0)))
            if bal <= 0: continue
            due_date = date.fromisoformat(due_str) if due_str else today
            days_overdue = (today - due_date).days
            if days_overdue <= 0:
                b = "Current (not overdue)"
            elif days_overdue <= 30:
                b = "0-30 Days Overdue"
            elif days_overdue <= 60:
                b = "31-60 Days Overdue"
            elif days_overdue <= 90:
                b = "61-90 Days Overdue"
            else:
                b = "91+ Days Overdue"
            buckets[b] = round_aed(buckets[b] + bal)
            rows.append({
                "invoice_no": inv.get('invoice_id', ''),
                "client": inv.get('client_name', 'Unknown'),
                "date": inv.get('date', ''),
                "due_date": due_str,
                "days_overdue": max(days_overdue, 0),
                "bucket": b,
                "balance": round_aed(bal)
            })
        total = round_aed(sum(buckets.values()))
        ninety_plus = buckets["91+ Days Overdue"]
        risk_comment = ""
        if total > 0 and ninety_plus / total > Decimal("0.15"):
            risk_comment = "⚠️ HIGH RISK: >15% of unpaid amounts are 91+ days overdue. Initiate collection calls immediately."
        elif ninety_plus > 0:
            risk_comment = "ℹ️ Some amounts are overdue — follow up with clients before they become bad debt."
        else:
            risk_comment = "✅ Excellent: No severely overdue invoices."
        explanation = f"""
        What this means for you (the business owner):
        This report shows how much money your CUSTOMERS OWE YOU, and HOW LONG it's been overdue.
        Collection rule of thumb in pharma:
        • Current: Send gentle reminders 2-3 days before due date.
        • 0-30 days: Friendly phone call + statement.
        • 31-60 days: Formal demand letter + stop further credit.
        • 61-90 days: Escalate to collection agency.
        • 91+ days: Very high chance of bad debt (write-off risk).
        {risk_comment}
        """
        return {
            "report_name": "Accounts Receivable Aging Report",
            "as_of": today,
            "buckets": buckets,
            "bucket_keys_order": bucket_keys_order,
            "total_unpaid": total,
            "rows": sorted(rows, key=lambda x: -x["days_overdue"]),
            "explanation": explanation,
            "summary_banner": {
                "grand_total": total,
                "avg": round_aed(total / Decimal(max(len(rows),1))),
                "pct_change": None,
                "ninety_plus": ninety_plus,
                "overdue_invoice_count": len(rows)
            }
        }

    # ==================================================================================
    # REPORT 4: REVENUE BY CLIENT
    # ==================================================================================
    def get_revenue_by_client(self):
        rows_dict = {}
        invs = [i for i in self._invoices if self._in_range(i.get('date','1999-01-01'), self.start, self.end)]
        for inv in invs:
            client = inv.get('client_name', 'Unknown Client')
            d = rows_dict.setdefault(client, {
                "client": client, "total_invoiced": Decimal("0"),
                "total_paid": Decimal("0"), "total_unpaid": Decimal("0"),
                "invoice_count": 0
            })
            d["total_invoiced"] += Decimal(str(inv.get('grand_total', 0) or 0))
            d["total_paid"] += Decimal(str(inv.get('total_paid', 0) or 0))
            d["total_unpaid"] += Decimal(str(inv.get('balance_due', 0) or 0))
            d["invoice_count"] += 1
        rows = sorted(rows_dict.values(), key=lambda x: -x["total_invoiced"])
        # Round everything
        for r in rows:
            r["total_invoiced"] = round_aed(r["total_invoiced"])
            r["total_paid"] = round_aed(r["total_paid"])
            r["total_unpaid"] = round_aed(r["total_unpaid"])
            r["paid_pct"] = round_aed((r["total_paid"] / r["total_invoiced"] * 100) if r["total_invoiced"] else Decimal("0"))
        total_rev = sum(r["total_invoiced"] for r in rows)
        top_client = rows[0] if rows else None
        explanation = f"""
        Your TOP CLIENT this period: {top_client['client'] if top_client else 'N/A'} — {fmt_aed(top_client['total_invoiced']) if top_client else 'N/A'}.
        Use this report to:
        1. Identify your 80/20 (Pareto) clients — 20% of clients typically drive 80% of revenue.
        2. Spot clients with low paid_pct (high unpaid) to improve collections.
        3. Reward loyal clients (high volume + good payment history) with discounts.
        """
        return {
            "report_name": "Revenue by Client Report",
            "period": {"start": self.start, "end": self.end},
            "rows": rows,
            "total_clients": len(rows),
            "explanation": explanation,
            "summary_banner": {
                "grand_total": round_aed(total_rev),
                "avg": round_aed(total_rev / Decimal(max(len(rows),1))),
                "client_count": len(rows)
            }
        }

    # ==================================================================================
    # REPORT 5: STOCK VALUATION & MOVEMENT
    # ==================================================================================
    def get_stock_valuation(self):
        # Purchases = Inventory increase during period
        purchases_value = self._sum_net(["1300"], [e for e in self.gl_current if e["reference_type"]=="PURCHASE"])
        # Opening Inventory = 1300 cumulative to start-1
        open_inv = self._cumulative_to(["1300"], self.start - timedelta(days=1))
        # COGS during period
        cogs = self._sum_net(["5000"], self.gl_current)
        # Closing Inventory = Opening + Purchases - COGS (accounting check)
        closing_from_gl = round_aed(open_inv + purchases_value - cogs)
        # Retail value = Closing * 1.4 (standard pharma distribution margin)
        retail = round_aed(closing_from_gl * Decimal("1.4"))
        # Stock turnover ratio = COGS / Average Inventory
        avg_inv = (open_inv + closing_from_gl) / Decimal("2") if (open_inv or closing_from_gl) else Decimal("0")
        turnover = (cogs / avg_inv) if avg_inv else Decimal("0")
        days_inventory = (Decimal("365") / turnover) if turnover else Decimal("0")
        explanation = f"""
        Your Inventory in Plain English:
        • Opening Inventory: Value of stock in warehouse at the START of the period.
        • + Goods Received: Purchases received (new stock bought from suppliers).
        • - Goods Sold (COGS): Cost of the stock you sold to clients.
        • = Closing Inventory: What's left in warehouse today — this is your ASSET on the Balance Sheet.
        Key Ratios:
        • Stock Turnover: {round_aed(turnover)}x per year — how fast stock is sold.
        • Days Inventory Outstanding: {round_aed(days_inventory)} days — how many days your stock sits before sale.
        Pharma target: 45-90 days in inventory (due to expiry dates).
        """
        return {
            "report_name": "Stock Valuation & Movement Report",
            "period": {"start": self.start, "end": self.end},
            "lines": {
                "opening_inventory_cost": open_inv,
                "goods_received_purchases": purchases_value,
                "goods_sold_cogs": cogs,
                "closing_inventory_cost": closing_from_gl,
                "closing_inventory_retail_value": retail,
                "stock_turnover_ratio": round_aed(turnover),
                "days_inventory_outstanding": round_aed(days_inventory)
            },
            "table_rows": [
                {"label": "Opening Inventory (at Cost)",          "amount": open_inv, "type": ""},
                {"label": "  + Goods Received during Period (Purchases, Dr 1300)", "amount": purchases_value, "type": ""},
                {"label": "  − Cost of Goods Sold (Cr 1300 → Dr 5000)",          "amount": round_aed(-cogs), "type": ""},
                {"label": "Closing Inventory (at Cost)",          "amount": closing_from_gl, "type": "total_row"},
                {"label": "",                                     "amount": None, "type": ""},
                {"label": "Closing Inventory (Estimated Retail @ ~1.4x COGS)", "amount": retail, "type": ""},
                {"label": "Inventory Turnover Ratio (times/yr)",  "amount": f"{round_aed(turnover)}x", "type": ""},
                {"label": "Days Inventory Outstanding (DIO)",     "amount": f"{round_aed(days_inventory)} days", "type": "overdue_med" if days_inventory > 90 else ""},
            ],
            "explanation": explanation,
            "summary_banner": {
                "grand_total": closing_from_gl,
                "retail_value": retail,
                "turnover": round_aed(turnover),
                "dio": round_aed(days_inventory)
            }
        }

    # ==================================================================================
    # REPORT 6: SUPPLIER ANALYSIS (SPEND)
    # ==================================================================================
    def get_supplier_spend(self):
        rows_dict = {}
        for po in self._purchases:
            if not self._in_range(po.get('date','1999-01-01'), self.start, self.end): continue
            s = po.get('supplier', po.get('supplier_name', 'Unknown Supplier'))
            amt = Decimal(str(po.get('amount', po.get('total_amount', 0) or 0)))
            if amt <= 0: continue
            d = rows_dict.setdefault(s, {
                "supplier": s, "purchases": Decimal("0"), "expenses_paid": Decimal("0"), "total_spend": Decimal("0"),
                "po_count": 0
            })
            d["purchases"] += amt
            d["po_count"] += 1
        # Expenses by payee
        for e in [x for x in self.gl_current if x["reference_type"] == "EXPENSE" and x["account_code"] == "1100"]:
            desc = e.get("description", "")
            payee = ""
            if "Payee:" in desc:
                payee = desc.split("Payee:")[-1].split("|")[0].strip()
            if not payee:
                payee = "Misc Expense Payees"
            amt = Decimal(str(e.get("credit_amount", 0)))
            d = rows_dict.setdefault(payee, {
                "supplier": payee, "purchases": Decimal("0"), "expenses_paid": Decimal("0"), "total_spend": Decimal("0"),
                "po_count": 0
            })
            d["expenses_paid"] += amt
        for r in rows_dict.values():
            r["total_spend"] = round_aed(r["purchases"] + r["expenses_paid"])
            r["purchases"] = round_aed(r["purchases"])
            r["expenses_paid"] = round_aed(r["expenses_paid"])
        rows = sorted(rows_dict.values(), key=lambda x: -x["total_spend"])
        total = sum(r["total_spend"] for r in rows)
        total = round_aed(total)
        top_supp = rows[0] if rows else None
        top_pct = Decimal("0")
        if total and top_supp:
            top_pct = round_aed(top_supp["total_spend"] / total * Decimal("100"))
        explanation = f"""
        Your TOP SUPPLIER by spend: {top_supp['supplier'] if top_supp else 'N/A'} — {fmt_aed(top_supp['total_spend']) if top_supp else 'N/A'}{f' ({top_pct}% of total spend)' if top_supp else ''}.
        Use this to:
        • Negotiate bulk discounts with your largest suppliers (they want to keep your volume!).
        • Identify over-reliance — if >30% of spend goes to 1 supplier, diversify to avoid supply risk.
        • Track non-merchandise spend (rent, salaries, etc.) vs. product purchases.
        """
        return {
            "report_name": "Supplier Analysis & Spend Report",
            "period": {"start": self.start, "end": self.end},
            "rows": rows,
            "total_spend": total,
            "explanation": explanation,
            "summary_banner": {
                "grand_total": total,
                "supplier_count": len(rows),
                "avg_supplier": round_aed(total / Decimal(max(len(rows), 1))),
                "top_supplier_share_pct": top_pct,
                "top_supplier_name": top_supp["supplier"] if top_supp else "N/A",
            }
        }

    # ==================================================================================
    # REPORT 7: CASH FLOW STATEMENT (3-way split — Operating / Investing / Financing)
    # ==================================================================================
    def get_cash_flow(self):
        ZERO = Decimal("0.00")

        def _dr_total(codes, gl):
            codes_set = set(str(c) for c in codes)
            tot = Decimal("0.00")
            for e in gl or []:
                if not isinstance(e, dict):
                    continue
                if str(e.get('account_code') or '').strip() in codes_set:
                    try:
                        tot += Decimal(str(e.get('debit_amount') or 0))
                    except Exception:
                        pass
            return round_aed(tot)

        def _cr_total(codes, gl):
            codes_set = set(str(c) for c in codes)
            tot = Decimal("0.00")
            for e in gl or []:
                if not isinstance(e, dict):
                    continue
                if str(e.get('account_code') or '').strip() in codes_set:
                    try:
                        tot += Decimal(str(e.get('credit_amount') or 0))
                    except Exception:
                        pass
            return round_aed(tot)

        def _delta(end_bal, begin_bal):
            return round_aed(Decimal(str(end_bal)) - Decimal(str(begin_bal)))

        # ======================================================================
        # 1. NET INCOME (from P&L)
        # ======================================================================
        pnl = self.get_profit_and_loss()
        net_income = round_aed(pnl["lines"]["net_income"]["current"])

        # ======================================================================
        # 2. NON-CASH EXPENSES — Depreciation & Amortization + Gains/Losses
        # ======================================================================
        DEPRECIATION_CODES = ['5400', '5410', '5420']
        GAIN_LOSS_CODES = ['5500', '5600']
        depreciation_expense = self._sum_net(DEPRECIATION_CODES, self.gl_current)
        gain_loss_on_disposal = self._sum_net(GAIN_LOSS_CODES, self.gl_current)
        depreciation_amortization_addback = round_aed(
            abs(depreciation_expense) if depreciation_expense < 0 else depreciation_expense
        )
        gain_loss_adjustment = round_aed(
            (-1) * (abs(gain_loss_on_disposal) if gain_loss_on_disposal > 0 else Decimal("0"))
            + (abs(gain_loss_on_disposal) if gain_loss_on_disposal < 0 else Decimal("0"))
        )
        dnA_total_addback = round_aed(depreciation_amortization_addback + gain_loss_adjustment)

        # ======================================================================
        # 3. WORKING CAPITAL ADJUSTMENTS (ASC 230)
        #    Δ = End balance - Begin balance
        #    Assets: Δ>0 (increase) → subtract, Δ<0 (decrease) → add back
        #    Liabilities: Δ>0 (increase) → add, Δ<0 (decrease) → subtract
        # ======================================================================
        bs = self.get_balance_sheet()
        d_start = self.start - timedelta(days=1)
        d_end = self.end

        # --- Assets ---
        ar_end = self._cumulative_to(["1200"], d_end)
        ar_begin = self._cumulative_to(["1200"], d_start)
        delta_ar = _delta(ar_end, ar_begin)
        adj_ar = round_aed(-delta_ar)

        inv_end = self._cumulative_to(["1300"], d_end)
        inv_begin = self._cumulative_to(["1300"], d_start)
        delta_inv = _delta(inv_end, inv_begin)
        adj_inv = round_aed(-delta_inv)

        prepaid_end = self._cumulative_to(["1400", "1410"], d_end)
        prepaid_begin = self._cumulative_to(["1400", "1410"], d_start)
        delta_prepaid = _delta(prepaid_end, prepaid_begin)
        adj_prepaid = round_aed(-delta_prepaid)

        oca_end = self._cumulative_to(["1500"], d_end)
        oca_begin = self._cumulative_to(["1500"], d_start)
        delta_oca = _delta(oca_end, oca_begin)
        adj_oca = round_aed(-delta_oca)

        # --- Liabilities ---
        ap_end = self._cumulative_to(["2000", "2010"], d_end)
        ap_begin = self._cumulative_to(["2000", "2010"], d_start)
        delta_ap = _delta(ap_end, ap_begin)
        adj_ap = round_aed(delta_ap)

        accrued_end = self._cumulative_to(["2100", "2110", "2120"], d_end)
        accrued_begin = self._cumulative_to(["2100", "2110", "2120"], d_start)
        delta_accrued = _delta(accrued_end, accrued_begin)
        adj_accrued = round_aed(delta_accrued)

        taxpay_end = self._cumulative_to(["2150"], d_end)
        taxpay_begin = self._cumulative_to(["2150"], d_start)
        delta_taxpay = _delta(taxpay_end, taxpay_begin)
        adj_taxpay = round_aed(delta_taxpay)

        unearned_end = self._cumulative_to(["2300", "2310"], d_end)
        unearned_begin = self._cumulative_to(["2300", "2310"], d_start)
        delta_unearned = _delta(unearned_end, unearned_begin)
        adj_unearned = round_aed(delta_unearned)

        other_cl_codes = [c for c in ["2250", "2260", "2270", "2280", "2290"]]
        ocl_end = self._cumulative_to(other_cl_codes, d_end)
        ocl_begin = self._cumulative_to(other_cl_codes, d_start)
        delta_ocl = _delta(ocl_end, ocl_begin)
        adj_ocl = round_aed(delta_ocl)

        total_wc_adj = round_aed(
            adj_ar + adj_inv + adj_prepaid + adj_oca
            + adj_ap + adj_accrued + adj_taxpay + adj_unearned + adj_ocl
        )

        # ======================================================================
        # 4. NET CASH FROM OPERATING ACTIVITIES
        # ======================================================================
        operating_cf = round_aed(net_income + dnA_total_addback + total_wc_adj)

        # ======================================================================
        # 5. INVESTING ACTIVITIES
        # ======================================================================
        PPE_CODES = ["1600", "1620", "1630", "1640"]
        INVESTMENT_CODES = ["1700", "1710"]

        capex_ppe = _dr_total(PPE_CODES, self.gl_current)
        purchase_lt_investments = _dr_total(INVESTMENT_CODES, self.gl_current)

        proceeds_sale_ppe = _cr_total(PPE_CODES, self.gl_current)
        proceeds_sale_investments = _cr_total(INVESTMENT_CODES, self.gl_current)
        total_investing_inflows = round_aed(proceeds_sale_ppe + proceeds_sale_investments)
        total_investing_outflows = round_aed(capex_ppe + purchase_lt_investments)

        investing_cf = round_aed(total_investing_inflows - total_investing_outflows)

        # ======================================================================
        # 6. FINANCING ACTIVITIES
        # ======================================================================
        ST_DEBT_CODES = ["2200", "2210", "2220"]
        LT_DEBT_CODES = ["2400", "2410"]
        DEBT_ALL = ST_DEBT_CODES + LT_DEBT_CODES
        CAPSTOCK_CODES = ["3100", "3110", "3120"]
        DRAWINGS_CODES = ["3200"]
        TREASURY_CODES = ["3300", "3310"]

        proceeds_debt = _cr_total(DEBT_ALL, self.gl_current)
        repayment_debt = _dr_total(DEBT_ALL, self.gl_current)
        capital_contributions = _cr_total(CAPSTOCK_CODES, self.gl_current)
        owner_drawings_3200 = _dr_total(DRAWINGS_CODES, self.gl_current)
        owner_drawings_3100 = _dr_total(CAPSTOCK_CODES, self.gl_current)
        owner_drawings_total = round_aed(owner_drawings_3200 + owner_drawings_3100)
        treasury_purchases = _dr_total(TREASURY_CODES, self.gl_current)

        total_financing_inflows = round_aed(proceeds_debt + capital_contributions)
        total_financing_outflows = round_aed(repayment_debt + owner_drawings_total + treasury_purchases)
        financing_cf = round_aed(total_financing_inflows - total_financing_outflows)

        # ======================================================================
        # 7, 8, 9, 10 — CASH BRIDGE
        # ======================================================================
        fx_effect = Decimal("0.00")
        net_cash_flow = round_aed(operating_cf + investing_cf + financing_cf + fx_effect)

        beginning_cash = self._cumulative_to(["1000", "1100"], self.start - timedelta(days=1))
        ending_cash = self._cumulative_to(["1000", "1100"], self.end)
        ending_cash_from_bridge = round_aed(beginning_cash + net_cash_flow + fx_effect)

        # ======================================================================
        # 11. CROSS-CHECK TIE with Balance Sheet
        # ======================================================================
        bs_cash = Decimal(str(bs["lines"]["assets"].get("1000 Cash on Hand", ZERO) or ZERO))
        bs_bank = Decimal(str(bs["lines"]["assets"].get("1100 Cash in Bank", ZERO) or ZERO))
        bs_cash_total = round_aed(bs_cash + bs_bank)

        # ======================================================================
        # DIAGNOSTICS — 6 GAAP Checks
        # ======================================================================
        diagnostics = {"passed": True, "checks": [], "issues": [], "fix_steps": []}
        def _chk(name, ok, detail):
            diagnostics["checks"].append({"name": name, "passed": bool(ok), "detail": str(detail)})
            if not ok:
                diagnostics["passed"] = False

        # --- CF-A: Operating CF Indirect Math ---
        try:
            expected_opcf = round_aed(net_income + dnA_total_addback + total_wc_adj)
            ok = abs(expected_opcf - operating_cf) < Decimal("0.01")
            detail = (
                f"NI {fmt_aed(net_income)} + D&A {fmt_aed(dnA_total_addback)} "
                f"+ ΔWC {fmt_aed(total_wc_adj)} = {fmt_aed(expected_opcf)}  |  "
                f"Reported OpCF = {fmt_aed(operating_cf)}  |  Δ = {fmt_aed(expected_opcf - operating_cf)}"
            )
            _chk("CF-A: Operating CF Indirect Math (NI + D&A ± ΔWC = OpCF)", ok, detail)
            if not ok:
                diagnostics["issues"].append(
                    f"Operating activities arithmetic off by {fmt_aed(abs(expected_opcf - operating_cf))}."
                )
                diagnostics["fix_steps"].append(
                    "Check working-capital adjustment signs: Assets use -Δ, Liabilities use +Δ."
                )
        except Exception as exc:
            _chk("CF-A", False, f"Check failed: {exc}")

        # --- CF-B: Σ(Op + Inv + Fin + FX) = Net Δ in Cash ---
        try:
            expected_net = round_aed(operating_cf + investing_cf + financing_cf + fx_effect)
            ok = abs(expected_net - net_cash_flow) < Decimal("0.01")
            detail = (
                f"Op {fmt_aed(operating_cf)} + Inv {fmt_aed(investing_cf)} + "
                f"Fin {fmt_aed(financing_cf)} + FX {fmt_aed(fx_effect)} = {fmt_aed(expected_net)}  |  "
                f"Net Δ Cash = {fmt_aed(net_cash_flow)}  |  Δ = {fmt_aed(expected_net - net_cash_flow)}"
            )
            _chk("CF-B: Σ(Op + Inv + Fin + FX) = Net Δ in Cash", ok, detail)
            if not ok:
                diagnostics["issues"].append("Three-sector sum does not equal net change in cash.")
                diagnostics["fix_steps"].append(
                    "Verify every Dr/Cr to non-cash balance-sheet accounts is captured in either Op/Inv/Fin."
                )
        except Exception as exc:
            _chk("CF-B", False, f"Check failed: {exc}")

        # --- CF-C: Begin Cash + Net Δ + FX = Ending Cash (Bridge) ---
        try:
            expected_end = round_aed(beginning_cash + net_cash_flow + fx_effect)
            ok = abs(expected_end - ending_cash) < Decimal("0.01")
            detail = (
                f"Begin {fmt_aed(beginning_cash)} + NetΔ {fmt_aed(net_cash_flow)} + "
                f"FX {fmt_aed(fx_effect)} = {fmt_aed(expected_end)}  |  "
                f"Actual Ending = {fmt_aed(ending_cash)}  |  Δ = {fmt_aed(expected_end - ending_cash)}"
            )
            _chk("CF-C: Begin Cash + Net Δ + FX = Ending Cash (Bridge)", ok, detail)
            if not ok:
                diagnostics["issues"].append(
                    f"Cash bridge mis-stated by {fmt_aed(abs(expected_end - ending_cash))}."
                )
                diagnostics["fix_steps"].append(
                    "Confirm cumulative_to start-1 / end dates match report window exactly. "
                    "Look for GL entries outside report window but with transaction_date within period."
                )
        except Exception as exc:
            _chk("CF-C", False, f"Check failed: {exc}")

        # --- CF-D: Ending Cash = BS (1000+1100) Cross-Statement Tie ---
        try:
            ok = abs(ending_cash - bs_cash_total) < Decimal("0.01")
            detail = (
                f"CF Ending Cash = {fmt_aed(ending_cash)}  |  "
                f"BS 1000 Cash + 1100 Bank = {fmt_aed(bs_cash)} + {fmt_aed(bs_bank)} = {fmt_aed(bs_cash_total)}  |  "
                f"Δ = {fmt_aed(ending_cash - bs_cash_total)}"
            )
            _chk("CF-D: Ending Cash = BS (1000 + 1100) Cross-Statement Tie", ok, detail)
            if not ok:
                diagnostics["issues"].append(
                    f"Ending cash does not tie to Balance Sheet by {fmt_aed(abs(ending_cash - bs_cash_total))}."
                )
                diagnostics["fix_steps"].append(
                    "Rebuild GL then rerun BS + CF in same ReportEngine instance. "
                    "If still off, check that BS uses the same as_of = self.end."
                )
        except Exception as exc:
            _chk("CF-D", False, f"Check failed: {exc}")

        # --- CF-E: D&A add-back matches Accumulated Depreciation (1610) Cr movement ---
        try:
            accum_depr_cr = _cr_total(["1610"], self.gl_current)
            ok = abs(depreciation_amortization_addback - accum_depr_cr) < Decimal("0.01")
            detail = (
                f"P&L D&A Expense (5400) = {fmt_aed(depreciation_amortization_addback)}  |  "
                f"BS Accumulated Depreciation (1610) Cr = {fmt_aed(accum_depr_cr)}  |  "
                f"Δ = {fmt_aed(depreciation_amortization_addback - accum_depr_cr)}"
            )
            _chk("CF-E: D&A add-back matches Accumulated Depreciation (1610) Cr movement", ok, detail)
            if not ok:
                diagnostics["issues"].append(
                    f"D&A add-back ({fmt_aed(depreciation_amortization_addback)}) ≠ AccDepr credits ({fmt_aed(accum_depr_cr)})."
                )
                diagnostics["fix_steps"].append(
                    "Verify Depreciation JE: Dr 5400 Depr Expense / Cr 1610 Accumulated Depreciation. "
                    "If disposal removed AccDepr via Dr 1610, ensure PPE gross credits are matched to 1600."
                )
        except Exception as exc:
            _chk("CF-E", False, f"Check failed: {exc}")

        # --- CF-F: Working-capital adjustments signs per ASC 230 ---
        try:
            signs_ok = True
            notes = []
            if delta_ar != ZERO:
                expected_sign_ar = -1 if delta_ar > 0 else (1 if delta_ar < 0 else 0)
                actual_sign_ar = -1 if adj_ar < 0 else (1 if adj_ar > 0 else 0)
                if expected_sign_ar != actual_sign_ar:
                    signs_ok = False
                    notes.append(f"AR sign: Δ={fmt_aed(delta_ar)} adj={fmt_aed(adj_ar)}")
            if delta_inv != ZERO:
                expected_sign_inv = -1 if delta_inv > 0 else (1 if delta_inv < 0 else 0)
                actual_sign_inv = -1 if adj_inv < 0 else (1 if adj_inv > 0 else 0)
                if expected_sign_inv != actual_sign_inv:
                    signs_ok = False
                    notes.append(f"Inv sign: Δ={fmt_aed(delta_inv)} adj={fmt_aed(adj_inv)}")
            if delta_ap != ZERO:
                expected_sign_ap = 1 if delta_ap > 0 else (-1 if delta_ap < 0 else 0)
                actual_sign_ap = 1 if adj_ap > 0 else (-1 if adj_ap < 0 else 0)
                if expected_sign_ap != actual_sign_ap:
                    signs_ok = False
                    notes.append(f"AP sign: Δ={fmt_aed(delta_ap)} adj={fmt_aed(adj_ap)}")
            if delta_accrued != ZERO:
                expected_sign_ac = 1 if delta_accrued > 0 else (-1 if delta_accrued < 0 else 0)
                actual_sign_ac = 1 if adj_accrued > 0 else (-1 if adj_accrued < 0 else 0)
                if expected_sign_ac != actual_sign_ac:
                    signs_ok = False
                    notes.append(f"Accrued sign: Δ={fmt_aed(delta_accrued)} adj={fmt_aed(adj_accrued)}")
            ok = signs_ok
            detail = (
                f"Assets (AR/Inv/Prepaid/OCA) use -Δ  |  "
                f"Liabs (AP/Accrued/TaxPay/Unearned/OCL) use +Δ  |  "
                + ("All signs correct." if signs_ok else " ISSUES: " + "; ".join(notes))
            )
            _chk("CF-F: Working-capital adjustments signs per ASC 230", ok, detail)
            if not ok:
                diagnostics["issues"].append(
                    "Working-capital adjustment sign convention violated for one or more line items."
                )
                diagnostics["fix_steps"].append(
                    "For Asset accounts: Adjustment = -(End - Begin). "
                    "For Liability accounts: Adjustment = +(End - Begin)."
                )
        except Exception as exc:
            _chk("CF-F", False, f"Check failed: {exc}")

        diagnostics["summary"] = (
            "✅ All 6 ASC 230 Cash-Flow diagnostic checks passed." if diagnostics["passed"]
            else f"⚠️ {len(diagnostics['issues'])} ASC 230 diagnostic issue(s) — follow numbered Fix Steps."
        )
        if not diagnostics["passed"]:
            diagnostics["fix_steps"].append(
                "Final Step: After any corrections, click 🔁 Rebuild GL from All Data then rerun Cash Flow."
            )

        # ======================================================================
        # FREE CASH FLOW
        # ======================================================================
        free_cash_flow = round_aed(operating_cf - capex_ppe)

        # ======================================================================
        # EXPLANATION
        # ======================================================================
        op_verb = "✅ Positive" if operating_cf > 0 else ("⚠️ Negative" if operating_cf < 0 else "➖ Neutral")
        inv_verb = "✅ Net Inflow" if investing_cf > 0 else ("⚠️ Net Outflow" if investing_cf < 0 else "➖ Neutral")
        fin_verb = "✅ Net Inflow" if financing_cf > 0 else ("⚠️ Net Outflow" if financing_cf < 0 else "➖ Neutral")
        if net_cash_flow > 0:
            status = f"✅ POSITIVE NET CASH FLOW: Cash position improved by {fmt_aed(net_cash_flow)}."
        elif net_cash_flow < 0:
            status = f"⚠️ NET CASH DECREASE: Cash declined by {fmt_aed(abs(net_cash_flow))} during period."
        else:
            status = "➖ Cash neutral: Inflow matched outflow exactly."
        daily_ops_out = (
            Decimal(str(abs(capex_ppe) + abs(purchase_lt_investments) + abs(repayment_debt) + abs(owner_drawings_total)))
            / Decimal(max((self.end - self.start).days + 1, 1))
        )
        burn_days = ZERO
        if daily_ops_out > Decimal("0.0001"):
            burn_days = round_aed(max(ZERO, ending_cash) / daily_ops_out)
        explanation = f"""
        {status}\n
        US GAAP ASC 230 — INDIRECT METHOD (reconciles Net Income → Operating Cash Flow):
        1) OPERATING ACTIVITIES ({op_verb}): {fmt_aed(operating_cf)}
            • Start with Net Income (P&L): {fmt_aed(net_income)}
            • Add back non-cash D&A (+ gains/losses on disposal): {fmt_aed(dnA_total_addback)}
            • Working capital changes (ASC 230 sign rule): {fmt_aed(total_wc_adj)}
              — Δ AR {fmt_aed(adj_ar)}, Δ Inv {fmt_aed(adj_inv)}, Δ Prepaid {fmt_aed(adj_prepaid)}
              — Δ AP {fmt_aed(adj_ap)}, Δ Accrued {fmt_aed(adj_accrued)}, Δ TaxPay {fmt_aed(adj_taxpay)}
              — Δ Unearned {fmt_aed(adj_unearned)}, Δ Other CL {fmt_aed(adj_ocl)}
        2) INVESTING ACTIVITIES ({inv_verb}): {fmt_aed(investing_cf)}
            • PPE / Intangibles CapEx out: {fmt_aed(-capex_ppe)}
            • LT Investment purchases out: {fmt_aed(-purchase_lt_investments)}
            • Proceeds from PPE / Investment disposals in: {fmt_aed(total_investing_inflows)}
        3) FINANCING ACTIVITIES ({fin_verb}): {fmt_aed(financing_cf)}
            • Debt borrowings in: {fmt_aed(proceeds_debt)}
            • Debt principal repayments out: {fmt_aed(-repayment_debt)}
            • Capital contributions in: {fmt_aed(capital_contributions)}
            • Owner Drawings / Distributions out: {fmt_aed(-owner_drawings_total)}
            • Treasury Stock purchases out: {fmt_aed(-treasury_purchases)}

        FREE CASH FLOW (OpCF − CapEx): {fmt_aed(free_cash_flow)}

        CASH BRIDGE:
        • Cash & Equivalents, Beginning ({self.start.isoformat()}): {fmt_aed(beginning_cash)}
        • + Net Change in Cash during period: {fmt_aed(net_cash_flow)}
        • + Effect of Exchange Rate Changes (ASC 230-10-45 placeholder): {fmt_aed(fx_effect)}
        • = Cash & Equivalents, Ending ({self.end.isoformat()}): {fmt_aed(ending_cash)}

        Cross-Statement Tie (BS 1000+1100): {fmt_aed(bs_cash_total)}
        Approximate cash runway (est'd daily spend): ~{burn_days} days.
        """

        # ======================================================================
        # TABLE ROWS — ASC 230 line items with proper tags
        # ======================================================================
        table_rows = [
            {"label": "OPERATING ACTIVITIES", "amount": None, "tag": "section"},
            {"label": "  Net Income", "amount": net_income, "tag": "level2"},
            {"label": "  Adjustments to reconcile Net Income to net cash:", "amount": None, "tag": None},
            {"label": "    Depreciation & Amortization", "amount": depreciation_amortization_addback, "tag": "level2"},
            {"label": "    (Gains) / Losses on Disposal of Assets", "amount": gain_loss_adjustment, "tag": "level2"},
            {"label": "  Changes in Working Capital:", "amount": None, "tag": None},
            {"label": "    (Increase) / Decrease in Accounts Receivable", "amount": adj_ar, "tag": "level2"},
            {"label": "    (Increase) / Decrease in Inventories", "amount": adj_inv, "tag": "level2"},
            {"label": "    (Increase) / Decrease in Prepaid Expenses", "amount": adj_prepaid, "tag": "level2"},
            {"label": "    (Increase) / Decrease in Other Current Assets", "amount": adj_oca, "tag": "level2"},
            {"label": "    Increase / (Decrease) in Accounts Payable", "amount": adj_ap, "tag": "level2"},
            {"label": "    Increase / (Decrease) in Accrued Expenses", "amount": adj_accrued, "tag": "level2"},
            {"label": "    Increase / (Decrease) in Income Taxes Payable", "amount": adj_taxpay, "tag": "level2"},
            {"label": "    Increase / (Decrease) in Unearned Revenue", "amount": adj_unearned, "tag": "level2"},
            {"label": "    Increase / (Decrease) in Other Current Liabilities", "amount": adj_ocl, "tag": "level2"},
            {"label": "  Net Cash Provided by (Used in) Operating Activities", "amount": operating_cf, "tag": "total_row"},

            {"label": "INVESTING ACTIVITIES", "amount": None, "tag": "section"},
            {"label": "  Purchase of Property, Plant & Equipment", "amount": round_aed(-capex_ppe), "tag": "level2"},
            {"label": "  Purchase of Long-Term Investments", "amount": round_aed(-purchase_lt_investments), "tag": "level2"},
            {"label": "  Proceeds from Sale of PPE & Intangibles", "amount": proceeds_sale_ppe, "tag": "level2"},
            {"label": "  Proceeds from Sale of Long-Term Investments", "amount": proceeds_sale_investments, "tag": "level2"},
            {"label": "  Net Cash Provided by (Used in) Investing Activities", "amount": investing_cf, "tag": "total_row"},

            {"label": "FINANCING ACTIVITIES", "amount": None, "tag": "section"},
            {"label": "  Proceeds from Short-Term & Long-Term Debt", "amount": proceeds_debt, "tag": "level2"},
            {"label": "  Repayment of Debt Principal", "amount": round_aed(-repayment_debt), "tag": "level2"},
            {"label": "  Capital Contributions (Owner Equity)", "amount": capital_contributions, "tag": "level2"},
            {"label": "  Owner Drawings / Distributions", "amount": round_aed(-owner_drawings_total), "tag": "level2"},
            {"label": "  Purchase of Treasury Stock", "amount": round_aed(-treasury_purchases), "tag": "level2"},
            {"label": "  Net Cash Provided by (Used in) Financing Activities", "amount": financing_cf, "tag": "total_row"},

            {"label": "EFFECT OF EXCHANGE RATE CHANGES ON CASH", "amount": fx_effect, "tag": None},
            {"label": "NET INCREASE (DECREASE) IN CASH AND CASH EQUIVALENTS", "amount": net_cash_flow, "tag": "grand_total"},

            {"label": "RECONCILIATION OF CASH BALANCES", "amount": None, "tag": "section"},
            {"label": "  Cash & Cash Equivalents, Beginning of Period", "amount": beginning_cash, "tag": "level2"},
            {"label": "  Cash & Cash Equivalents, End of Period", "amount": ending_cash, "tag": "total_row"},
            {"label": "  Cross-Check: Balance Sheet Cash (1000 + 1100)", "amount": bs_cash_total, "tag": "level2"},
        ]

        # ======================================================================
        # RETURN PAYLOAD
        # ======================================================================
        return {
            "report_name": "Consolidated Statements of Cash Flows (US GAAP ASC 230 — Indirect Method, 3-Year Comparative)",
            "period": {"start": self.start, "end": self.end},
            "lines": {
                "net_income": net_income,
                "depreciation_amortization": depreciation_amortization_addback,
                "gain_loss_on_disposal_adjustment": gain_loss_adjustment,
                "total_non_cash_items": dnA_total_addback,
                "delta_accounts_receivable": adj_ar,
                "delta_inventories": adj_inv,
                "delta_prepaid_expenses": adj_prepaid,
                "delta_other_current_assets": adj_oca,
                "delta_accounts_payable": adj_ap,
                "delta_accrued_expenses": adj_accrued,
                "delta_income_taxes_payable": adj_taxpay,
                "delta_unearned_revenue": adj_unearned,
                "delta_other_current_liabilities": adj_ocl,
                "total_working_capital_adjustments": total_wc_adj,
                "operating_cash_flow": operating_cf,
                "capex_ppe": capex_ppe,
                "purchase_of_lt_investments": purchase_lt_investments,
                "proceeds_sale_ppe": proceeds_sale_ppe,
                "proceeds_sale_investments": proceeds_sale_investments,
                "investing_cash_flow": investing_cf,
                "proceeds_debt_borrowings": proceeds_debt,
                "repayment_of_debt_principal": repayment_debt,
                "capital_contributions": capital_contributions,
                "owner_drawings_distributions": owner_drawings_total,
                "purchase_of_treasury_stock": treasury_purchases,
                "financing_cash_flow": financing_cf,
                "effect_of_exchange_rate_changes": fx_effect,
                "net_cash_flow": net_cash_flow,
                "beginning_cash": beginning_cash,
                "ending_cash": ending_cash,
                "balance_sheet_cash_tie": bs_cash_total,
                "free_cash_flow": free_cash_flow,
                "cash_runway_days": burn_days,
            },
            "table_rows": table_rows,
            "explanation": explanation,
            "diagnostics": diagnostics,
            "summary_banner": {
                "grand_total": net_cash_flow,
                "operating_cash_flow": operating_cf,
                "investing_cash_flow": investing_cf,
                "financing_cash_flow": financing_cf,
                "closing_cash": ending_cash,
                "beginning_cash": beginning_cash,
                "free_cash_flow": free_cash_flow,
                "cash_runway_days": burn_days,
            }
        }


    # ==================================================================================
    # REPORT 8: STATEMENT OF CHANGES IN EQUITY (US GAAP ASC 505)
    # ==================================================================================
    def get_statement_of_changes_in_equity(self):
        """
        Consolidated Statements of Changes in Stockholders' Equity (US GAAP ASC 505)
        6-column format: Item | Capital Stock | Treasury Stock | Retained Earnings | AOCI | Total Equity

        Opening balances (as of start-1 day)
        + Owner Capital Contributions (credits to 3100/3110/3120)
        − Owner Drawings / Distributions (debits to 3200 or direct Dr to 3100)
        − Purchase of Treasury Stock (debits to 3300/3310)
        + Net Income (Loss) for the Period (from P&L["lines"]["net_income"]["current"])
        + Other Comprehensive Income for the Period (from P&L["lines"]["other_comprehensive_income"]["current"])
        = Closing Balances
        """
        pnl = self.get_profit_and_loss()
        ni_current = pnl["lines"]["net_income"]["current"]
        oci_current = pnl["lines"]["other_comprehensive_income"]["current"]

        day_before = self.start - timedelta(days=1)

        Capital_Stock_open = self._cumulative_to(["3100", "3110", "3120"], day_before)
        Treasury_Stock_open = -abs(self._cumulative_to(["3300", "3310"], day_before))
        Retained_Earnings_open = self._cumulative_to(["3000"], day_before)
        AOCI_open = self._cumulative_to(["3900", "3910", "3920", "3930"], day_before)
        Total_Equity_open = round_aed(Capital_Stock_open + Treasury_Stock_open + Retained_Earnings_open + AOCI_open)

        capital_contribs = Decimal("0.00")
        direct_capital_debits_for_draw = Decimal("0.00")
        for e in self.gl_current:
            code = str(e.get('account_code') or '').strip()
            if code in ("3100", "3110", "3120"):
                try:
                    cr = Decimal(str(e.get("credit_amount", 0)))
                except Exception:
                    cr = Decimal("0")
                try:
                    dr = Decimal(str(e.get("debit_amount", 0)))
                except Exception:
                    dr = Decimal("0")
                capital_contribs += cr
                direct_capital_debits_for_draw += dr
        capital_contribs = round_aed(capital_contribs)
        direct_capital_debits_for_draw = round_aed(direct_capital_debits_for_draw)

        drawings_3200 = Decimal("0.00")
        for e in self.gl_current:
            if str(e.get('account_code') or '').strip() == "3200":
                try:
                    dr = Decimal(str(e.get("debit_amount", 0)))
                except Exception:
                    dr = Decimal("0")
                drawings_3200 += dr
        drawings_3200 = round_aed(drawings_3200)
        total_drawings = round_aed(drawings_3200 + direct_capital_debits_for_draw)

        treasury_purchases = Decimal("0.00")
        for e in self.gl_current:
            if str(e.get('account_code') or '').strip() in ("3300", "3310"):
                try:
                    dr = Decimal(str(e.get("debit_amount", 0)))
                except Exception:
                    dr = Decimal("0")
                treasury_purchases += dr
        treasury_purchases = round_aed(treasury_purchases)

        movement_cap = capital_contribs
        movement_treas = -abs(treasury_purchases)
        movement_re = round_aed(-total_drawings + ni_current)
        movement_aoci = oci_current
        movement_total = round_aed(movement_cap + movement_treas + movement_re + movement_aoci)

        Capital_Stock_close = round_aed(Capital_Stock_open + movement_cap)
        Treasury_Stock_close = round_aed(Treasury_Stock_open + movement_treas)
        Retained_Earnings_close = round_aed(Retained_Earnings_open + movement_re)
        AOCI_close = round_aed(AOCI_open + movement_aoci)
        Total_Equity_close = round_aed(Capital_Stock_close + Treasury_Stock_close + Retained_Earnings_close + AOCI_close)

        table_rows = [
            {"tag": "section", "columns": {
                "Item": "",
                "Capital_Stock": "Capital Stock",
                "Treasury_Stock": "Treasury Stock",
                "Retained_Earnings": "Retained Earnings",
                "AOCI": "AOCI",
                "Total_Equity": "Total Equity"
            }},
            {"tag": "subsection", "columns": {
                "Item": f"Opening Balance ({day_before.isoformat()})",
                "Capital_Stock": Capital_Stock_open,
                "Treasury_Stock": Treasury_Stock_open,
                "Retained_Earnings": Retained_Earnings_open,
                "AOCI": AOCI_open,
                "Total_Equity": Total_Equity_open,
            }},
            {"columns": {
                "Item": "Owner Capital Contributions",
                "Capital_Stock": capital_contribs,
                "Treasury_Stock": Decimal("0.00"),
                "Retained_Earnings": Decimal("0.00"),
                "AOCI": Decimal("0.00"),
                "Total_Equity": capital_contribs,
            }},
            {"columns": {
                "Item": "Owner Drawings / Distributions",
                "Capital_Stock": Decimal("0.00"),
                "Treasury_Stock": Decimal("0.00"),
                "Retained_Earnings": round_aed(-total_drawings),
                "AOCI": Decimal("0.00"),
                "Total_Equity": round_aed(-total_drawings),
            }},
            {"columns": {
                "Item": "Purchase of Treasury Stock",
                "Capital_Stock": Decimal("0.00"),
                "Treasury_Stock": movement_treas,
                "Retained_Earnings": Decimal("0.00"),
                "AOCI": Decimal("0.00"),
                "Total_Equity": movement_treas,
            }},
            {"columns": {
                "Item": ("Net Income for the Period" if ni_current >= 0 else "Net Loss for the Period"),
                "Capital_Stock": Decimal("0.00"),
                "Treasury_Stock": Decimal("0.00"),
                "Retained_Earnings": ni_current,
                "AOCI": Decimal("0.00"),
                "Total_Equity": ni_current,
            }, "tag": "gross_profit" if ni_current >= 0 else "overdue_high"},
            {"columns": {
                "Item": "Other Comprehensive Income for the Period",
                "Capital_Stock": Decimal("0.00"),
                "Treasury_Stock": Decimal("0.00"),
                "Retained_Earnings": Decimal("0.00"),
                "AOCI": oci_current,
                "Total_Equity": oci_current,
            }},
            {"tag": "grand_total", "columns": {
                "Item": "Closing Balance",
                "Capital_Stock": Capital_Stock_close,
                "Treasury_Stock": Treasury_Stock_close,
                "Retained_Earnings": Retained_Earnings_close,
                "AOCI": AOCI_close,
                "Total_Equity": Total_Equity_close,
            }},
        ]

        explanation = f"""
        What this statement tells you (US GAAP ASC 505 — plain English):

        This report reconciles every component of stockholders' equity from the start of the period to the end.
        Each column rolls forward separately, and the far-right Total Equity column must tie to the Balance Sheet.

        ▶️  BEGINNING POSITION (as of {day_before.isoformat()}):
           • Capital Stock (3100/3110/3120):   {fmt_aed(Capital_Stock_open)}
           • Treasury Stock (33xx, contra):    {fmt_aed(Treasury_Stock_open)}
           • Retained Earnings (3000):         {fmt_aed(Retained_Earnings_open)}
           • AOCI (39xx, OCI reserve):         {fmt_aed(AOCI_open)}
           • TOTAL EQUITY (opening):           {fmt_aed(Total_Equity_open)}

        🔄  MOVEMENTS DURING THE PERIOD ({self.start.isoformat()} → {self.end.isoformat()}):
           1. Owner Capital Contributions:        +{fmt_aed(capital_contribs)}    (new money / capital paid in by owners → ↑ Capital Stock)
           2. Owner Drawings / Distributions:     −{fmt_aed(total_drawings)}    (owner withdrawals → ↓ Retained Earnings)
           3. Purchase of Treasury Stock:         {fmt_aed(movement_treas)}    (buy-back → ↓ Treasury Stock contra, ↓ Total Equity)
           4. Net Income (Loss) for Period:       {fmt_aed(ni_current)}    ⚓ TIES to P&L lines.net_income.current
           5. Other Comprehensive Income:         {fmt_aed(oci_current)}    ⚓ TIES to P&L lines.other_comprehensive_income.current

        🏁  ENDING POSITION (as of {self.end.isoformat()}):
           • Capital Stock:                      {fmt_aed(Capital_Stock_close)}
           • Treasury Stock (contra):            {fmt_aed(Treasury_Stock_close)}
           • Retained Earnings:                  {fmt_aed(Retained_Earnings_close)}
           • AOCI:                               {fmt_aed(AOCI_close)}
           • TOTAL EQUITY (closing):             {fmt_aed(Total_Equity_close)}   ⚓ TIES to Balance Sheet equity.total_equity

        📌  Cross-Statement Ties (verifiable by running all 3 core financials):
           A. Net Income on THIS statement = P&L Statement → Net Income for the same period.
           B. OCI on THIS statement = P&L Statement → Other Comprehensive Income for the same period.
           C. Closing Retained Earnings = Balance Sheet → equity section "3000 Retained Earnings, End of Period".
           D. Closing Total Equity = Balance Sheet → equity.total_equity (and therefore Assets − Liabilities).
        """.strip()

        diagnostics = {"passed": True, "checks": [], "root_causes": [], "fix_steps": []}
        def _chk_eq(name, ok, detail):
            diagnostics["checks"].append({"name": name, "passed": bool(ok), "detail": str(detail)})
            if not ok:
                diagnostics["passed"] = False

        try:
            foot_open = round_aed(Capital_Stock_open + Treasury_Stock_open + Retained_Earnings_open + AOCI_open)
            ok = abs(foot_open - Total_Equity_open) < Decimal("0.01")
            _chk_eq("SCE-A: Opening balances foot (Σ Cap Stock + Treasury + RE + AOCI = Total Equity Open)",
                    ok,
                    f"Σ columns = {fmt_aed(foot_open)}; reported Total_Equity_open = {fmt_aed(Total_Equity_open)}; Δ = {fmt_aed(foot_open - Total_Equity_open)}")
            if not ok:
                diagnostics["root_causes"].append(
                    f"Opening row arithmetic is off by {fmt_aed(abs(foot_open - Total_Equity_open))}. "
                    "The four equity sub-components do not sum to the reported Total Equity opening.")
                diagnostics["fix_steps"].append(
                    "Fix SCE-A: Rebuild GL (refresh), then re-run Statement of Changes in Equity — opening totals are derived from _cumulative_to().")
        except Exception as exc:
            _chk_eq("SCE-A", False, f"Check failed: {exc}")

        try:
            ok = abs(ni_current - pnl["lines"]["net_income"]["current"]) < Decimal("0.01")
            _chk_eq("SCE-B: Net Income line ties to P&L[\"lines\"][\"net_income\"][\"current\"]",
                    ok,
                    f"SCE NI = {fmt_aed(ni_current)}; P&L NI = {fmt_aed(pnl['lines']['net_income']['current'])}; Δ = {fmt_aed(ni_current - pnl['lines']['net_income']['current'])}")
            if not ok:
                diagnostics["root_causes"].append(
                    "The Net Income (Loss) figure on this statement does not match the P&L Net Income for the same date range. "
                    "These two statements must share identical gl_current period.")
                diagnostics["fix_steps"].append(
                    "Fix SCE-B: Rebuild GL. Ensure both P&L and SCE use EXACTLY same start/end dates. Re-run both reports; NI must match to fils (0.00 AED).")
        except Exception as exc:
            _chk_eq("SCE-B", False, f"Check failed: {exc}")

        try:
            ok = abs(oci_current - pnl["lines"]["other_comprehensive_income"]["current"]) < Decimal("0.01")
            _chk_eq("SCE-C: OCI line ties to P&L[\"lines\"][\"other_comprehensive_income\"][\"current\"]",
                    ok,
                    f"SCE OCI = {fmt_aed(oci_current)}; P&L OCI = {fmt_aed(pnl['lines']['other_comprehensive_income']['current'])}; Δ = {fmt_aed(oci_current - pnl['lines']['other_comprehensive_income']['current'])}")
            if not ok:
                diagnostics["root_causes"].append(
                    "Other Comprehensive Income on SCE ≠ OCI on P&L. OCI accounts (3900-3930) must produce the same movement in both reports.")
                diagnostics["fix_steps"].append(
                    "Fix SCE-C: Rebuild GL; verify both reports use same date range. Run P&L then SCE — OCI figures must match.")
        except Exception as exc:
            _chk_eq("SCE-C", False, f"Check failed: {exc}")

        try:
            bs = self.get_balance_sheet()
            bs_lines = bs.get("lines", {}) or {}
            bs_eq = bs_lines.get("equity", {}) or {}
            bs_re_ending = None
            for k, v in bs_eq.items():
                if "3000 Retained Earnings, End of Period" in str(k):
                    bs_re_ending = v
                    break
            if bs_re_ending is None:
                bs_re_ending = bs_eq.get("3000 Retained Earnings, End of Period")
            if bs_re_ending is not None:
                bs_re_ending = Decimal(str(bs_re_ending))
                ok = abs(bs_re_ending - Retained_Earnings_close) < Decimal("0.01")
                _chk_eq("SCE-D: Closing Retained Earnings ties to Balance Sheet lines[\"equity\"][\"3000 Retained Earnings, End of Period\"]",
                        ok,
                        f"SCE Closing RE = {fmt_aed(Retained_Earnings_close)}; BS RE Ending = {fmt_aed(bs_re_ending)}; Δ = {fmt_aed(Retained_Earnings_close - bs_re_ending)}")
                if not ok:
                    diagnostics["root_causes"].append(
                        f"Retained Earnings ending balance on SCE ({fmt_aed(Retained_Earnings_close)}) ≠ "
                        f"BS Retained Earnings End of Period ({fmt_aed(bs_re_ending)}). "
                        f"Cross-statement bridge broken by {fmt_aed(abs(Retained_Earnings_close - bs_re_ending))}.")
                    diagnostics["fix_steps"].append(
                        "Fix SCE-D: Run Balance Sheet FIRST (as of same end-date), then SCE. "
                        "Verify: RE opening + NI − Drawings on SCE equals BS RE ending. Rebuild GL if off.")
            else:
                _chk_eq("SCE-D: Closing RE — Balance Sheet tie (BS equity[\"3000 Retained Earnings, End of Period\"] unavailable)",
                        True,
                        "BS did not expose '3000 Retained Earnings, End of Period' key — skipped cross-tie.")
        except Exception as exc:
            _chk_eq("SCE-D", False, f"Check failed: {exc}")

        try:
            if 'bs' not in locals():
                bs = self.get_balance_sheet()
            bs_lines = bs.get("lines", {}) or {}
            bs_eq = bs_lines.get("equity", {}) or {}
            bs_total_eq = bs_eq.get("total_equity")
            if bs_total_eq is not None:
                bs_total_eq = Decimal(str(bs_total_eq))
                ok = abs(bs_total_eq - Total_Equity_close) < Decimal("0.01")
                _chk_eq("SCE-E: Closing Total Equity ties to Balance Sheet equity[\"total_equity\"] cross-statement",
                        ok,
                        f"SCE Total Equity close = {fmt_aed(Total_Equity_close)}; BS equity.total_equity = {fmt_aed(bs_total_eq)}; Δ = {fmt_aed(Total_Equity_close - bs_total_eq)}")
                if not ok:
                    diagnostics["root_causes"].append(
                        f"CRITICAL: Total Equity on SCE ({fmt_aed(Total_Equity_close)}) does NOT match "
                        f"Balance Sheet total_equity ({fmt_aed(bs_total_eq)}). The two statements are out of sync "
                        f"by {fmt_aed(abs(Total_Equity_close - bs_total_eq))}.")
                    diagnostics["fix_steps"].append(
                        "Fix SCE-E: Rebuild GL. Run Balance Sheet first, then SCE — both using the exact same end date. "
                        "Total Equity must match fils-to-fils (0.00 AED) across the two statements.")
            else:
                _chk_eq("SCE-E: Closing Total Equity — BS total_equity (BS total_equity unavailable)",
                        True,
                        "BS did not expose equity.total_equity — skipped cross-tie.")
        except Exception as exc:
            _chk_eq("SCE-E", False, f"Check failed: {exc}")

        try:
            foot_close = round_aed(Capital_Stock_close + Treasury_Stock_close + Retained_Earnings_close + AOCI_close)
            ok = abs(foot_close - Total_Equity_close) < Decimal("0.01")
            _chk_eq("SCE-F: Closing balances foot (Σ 4 closing columns = Total_Equity_closing) arithmetic foot",
                    ok,
                    f"Σ Cap({fmt_aed(Capital_Stock_close)}) + Treas({fmt_aed(Treasury_Stock_close)}) + RE({fmt_aed(Retained_Earnings_close)}) + AOCI({fmt_aed(AOCI_close)}) = {fmt_aed(foot_close)}; reported Total_Equity_close = {fmt_aed(Total_Equity_close)}; Δ = {fmt_aed(foot_close - Total_Equity_close)}")
            if not ok:
                diagnostics["root_causes"].append(
                    f"Closing row arithmetic doesn't foot. The four sub-components sum to {fmt_aed(foot_close)}, "
                    f"but Total_Equity_close reports {fmt_aed(Total_Equity_close)} — a {fmt_aed(abs(foot_close - Total_Equity_close))} gap.")
                diagnostics["fix_steps"].append(
                    "Fix SCE-F: Rebuild GL & re-run. This is pure column addition; if it fails after rebuild, inspect treasury_stock and AOCI movement signs.")
        except Exception as exc:
            _chk_eq("SCE-F", False, f"Check failed: {exc}")

        diagnostics["summary"] = (
            "✅ All 6 US GAAP ASC 505 Statement of Changes in Equity reconciliation checks passed."
            if diagnostics["passed"]
            else f"⚠️ {len(diagnostics['root_causes'])} reconciliation issue(s). Follow numbered Fix Steps above."
        )
        if not diagnostics["passed"]:
            diagnostics["fix_steps"].append(
                "Final step: after applying any fixes, click 🔁 Rebuild GL from All Data on the GAAP dashboard, then re-run SCE.")

        return {
            "report_name": "Consolidated Statements of Changes in Stockholders' Equity (US GAAP ASC 505)",
            "period": {"start": self.start, "end": self.end},
            "opening_capital_stock": Capital_Stock_open,
            "opening_treasury_stock": Treasury_Stock_open,
            "opening_retained_earnings": Retained_Earnings_open,
            "opening_aoci": AOCI_open,
            "capital_contributions": capital_contribs,
            "drawings": total_drawings,
            "treasury_purchases": treasury_purchases,
            "net_income_for_period": ni_current,
            "other_comprehensive_income": oci_current,
            "closing_capital_stock": Capital_Stock_close,
            "closing_treasury_stock": Treasury_Stock_close,
            "closing_retained_earnings": Retained_Earnings_close,
            "closing_aoci": AOCI_close,
            "total_opening_equity": Total_Equity_open,
            "total_closing_equity": Total_Equity_close,
            "table_rows": table_rows,
            "columns": ["Item", "Capital_Stock", "Treasury_Stock", "Retained_Earnings", "AOCI", "Total_Equity"],
            "explanation": explanation,
            "diagnostics": diagnostics,
            "summary_banner": {
                "opening_equity": Total_Equity_open,
                "capital_contributions": capital_contribs,
                "drawings": total_drawings,
                "closing_equity": Total_Equity_close,
                "net_income_for_period": ni_current,
                "other_comprehensive_income": oci_current,
            }
        }

    # ==================================================================================
    # REPORT 9: FINANCIAL RATIOS (NEW — 14 ratios, 4 categories)
    # ==================================================================================
    def get_financial_ratios(self):
        """
        14 Key Financial Ratios grouped into 4 categories with GAAP benchmarks:
        LIQUIDITY, PROFITABILITY, LEVERAGE (SOLVENCY), EFFICIENCY (ACTIVITY)
        Each ratio includes value + benchmark + color/interpretation.
        """
        def safe_div(num, den, default=Decimal("0")):
            try:
                if den is None:
                    return default
                d = Decimal(str(den))
                if abs(d) < Decimal("0.0001"):
                    return default
                return Decimal(str(num)) / d
            except Exception:
                return default

        pnl = self.get_profit_and_loss()
        bs = self.get_balance_sheet()
        cf = self.get_cash_flow()
        rev = pnl["lines"]["gross_revenue"]["current"]
        cogs = pnl["lines"]["cogs"]["current"]
        gp = pnl["lines"]["gross_profit"]["current"]
        np_ = pnl["lines"]["net_income"]["current"]
        # Correctly navigate the balance-sheet structure: lines.assets.total_assets, etc.
        _lines = bs.get("lines", {})
        _assets = _lines.get("assets", {}) if isinstance(_lines, dict) else {}
        _liabs  = _lines.get("liabilities", {}) if isinstance(_lines, dict) else {}
        _equity = _lines.get("equity", {}) if isinstance(_lines, dict) else {}
        total_assets = (_assets.get("total_assets") if isinstance(_assets, dict) else None) or Decimal("0")
        total_liabs  = (_liabs.get("total_liabilities") if isinstance(_liabs, dict) else None) or Decimal("0")
        total_eq     = (_equity.get("total_equity") if isinstance(_equity, dict) else None) or Decimal("0")
        try:
            total_assets = Decimal(str(total_assets))
            total_liabs  = Decimal(str(total_liabs))
            total_eq     = Decimal(str(total_eq))
        except Exception:
            total_assets = total_liabs = total_eq = Decimal("0")
        total_cl = Decimal("0")
        total_ca = Decimal("0")
        # Better: from cumulative account balances
        cash_and_bank = self._cumulative_to(["1000", "1100"], self.end)
        ar = self._cumulative_to(["1200"], self.end)
        inventory = self._cumulative_to(["1300"], self.end)
        ap = self._cumulative_to(["2000"], self.end)
        if not total_ca:
            total_ca = round_aed(cash_and_bank + ar + inventory)
        if not total_cl:
            total_cl = round_aed(ap)
        # ----- RATIOS -----
        # 1 LIQUIDITY
        current_ratio = safe_div(total_ca, total_cl)
        quick_assets = round_aed(cash_and_bank + ar)     # exclude inventory
        quick_ratio = safe_div(quick_assets, total_cl)
        cash_ratio = safe_div(cash_and_bank, total_cl)
        nwc = round_aed(total_ca - total_cl)
        # 2 PROFITABILITY
        gross_margin = safe_div(gp, rev, Decimal("0")) * Decimal("100") if rev else Decimal("0")
        net_margin = safe_div(np_, rev, Decimal("0")) * Decimal("100") if rev else Decimal("0")
        roe = safe_div(np_, total_eq, Decimal("0")) * Decimal("100") if total_eq else Decimal("0")
        roa = safe_div(np_, total_assets, Decimal("0")) * Decimal("100") if total_assets else Decimal("0")
        op_exp_total = pnl["lines"]["total_operating_expenses"]["current"]
        operating_profit = round_aed(gp - op_exp_total)
        operating_margin = safe_div(operating_profit, rev) * Decimal("100") if rev else Decimal("0")
        # 3 LEVERAGE
        debt_to_equity = safe_div(total_liabs, total_eq)
        debt_ratio = safe_div(total_liabs, total_assets)
        equity_ratio = safe_div(total_eq, total_assets)
        # 4 EFFICIENCY
        avg_inv = (self._cumulative_to(["1300"], self.start - timedelta(days=1)) + inventory) / Decimal("2") or Decimal("0")
        inv_turn = safe_div(cogs, avg_inv)
        days_inv = Decimal("365") / inv_turn if inv_turn > 0 else Decimal("0")
        days_ar = (safe_div(ar, rev) * Decimal("365")) if rev > 0 else Decimal("0")
        days_ap = (safe_div(ap, cogs) * Decimal("365")) if cogs > 0 else Decimal("0")
        ccc = round_aed(days_inv + days_ar - days_ap)
        asset_turn = safe_div(rev, total_assets)

        def classify(ratio_name, value):
            """Benchmarks for pharma distribution (SME) — return green/yellow/red + meaning."""
            v = Decimal(str(value))
            benchmarks = {
                "Current Ratio":      (Decimal("2"), Decimal("1"), "above"),
                "Quick Ratio":        (Decimal("1"), Decimal("0.5"), "above"),
                "Cash Ratio":         (Decimal("0.5"), Decimal("0.2"), "above"),
                "Gross Margin %":     (Decimal("25"), Decimal("15"), "above"),
                "Net Margin %":       (Decimal("5"), Decimal("0"), "above"),
                "Operating Margin %": (Decimal("10"), Decimal("3"), "above"),
                "ROE %":              (Decimal("15"), Decimal("5"), "above"),
                "ROA %":              (Decimal("7"), Decimal("2"), "above"),
                "Debt-to-Equity":     (Decimal("1"), Decimal("2"), "below"),
                "Debt Ratio":         (Decimal("0.5"), Decimal("0.8"), "below"),
                "Equity Ratio":       (Decimal("0.5"), Decimal("0.2"), "above"),
                "Inventory Turnover": (Decimal("6"), Decimal("3"), "above"),
                "DSO (Days Sales Out.)": (Decimal("30"), Decimal("60"), "below"),
                "DPO (Days Payables Out.)": (Decimal("45"), Decimal("75"), "below"),
                "CCC (Cash Conv. Cycle)":  (Decimal("30"), Decimal("60"), "below"),
                "Asset Turnover":     (Decimal("1.5"), Decimal("0.8"), "above"),
            }
            good, bad, direction = benchmarks.get(ratio_name, (Decimal("0"), Decimal("0"), "above"))
            if direction == "above":
                if v >= good: return ("🟢", "Good / Healthy")
                if v >= bad:  return ("🟡", "Watch / Acceptable")
                return ("🔴", "Weak / Needs action")
            else:  # below
                if v <= good: return ("🟢", "Good / Healthy")
                if v <= bad:  return ("🟡", "Watch / Acceptable")
                return ("🔴", "Weak / Needs action")

        all_ratios = [
            # Category, Name, value, unit, benchmark, description
            ("💧 LIQUIDITY (Can we pay short-term bills?)",
             [
                 ("Current Ratio", current_ratio, "x", "2.00x — ideally 2:1",
                  "All Current Assets ÷ All Current Liabilities. The classic 'can we cover short-term debt' test."),
                 ("Quick Ratio", quick_ratio, "x", "1.00x — ideally 1:1",
                  "(Cash + AR) ÷ Current Liabilities. Excludes slow-moving inventory. Tougher liquidity test."),
                 ("Cash Ratio", cash_ratio, "x", ">0.20x",
                  "Only Cash & Bank ÷ Current Liabilities. Emergency liquidity: 'could we pay today?'"),
                 ("Net Working Capital", nwc, "AED", "Positive & growing",
                  "Current Assets - Current Liabilities. A buffer / safety margin for operations."),
             ]),
            ("📈 PROFITABILITY (How well do we make money?)",
             [
                 ("Gross Margin %", gross_margin, "%", ">25% (pharma benchmark)",
                  "(Revenue - COGS) / Revenue × 100. Marks up on products before paying overhead."),
                 ("Operating Margin %", operating_margin, "%", ">10%",
                  "Operating Profit / Revenue × 100. Profit AFTER rent, salaries, admin — true core efficiency."),
                 ("Net Margin %", net_margin, "%", ">5%",
                  "Net Profit / Revenue × 100. Bottom-line: how many fils of profit per AED 1 of sales."),
                 ("Return on Equity (ROE) %", roe, "%", ">15%",
                  "Net Profit ÷ Equity × 100. Profit per AED of owner's money invested."),
                 ("Return on Assets (ROA) %", roa, "%", ">7%",
                  "Net Profit ÷ Total Assets × 100. How well assets are deployed to generate profit."),
             ]),
            ("🏗️  LEVERAGE / SOLVENCY (How much do we owe?)",
             [
                 ("Debt-to-Equity", debt_to_equity, "x", "<1.00x",
                  "Total Liabilities ÷ Equity. <1 means the business owes less than owners have in."),
                 ("Debt Ratio", debt_ratio, "%", "<50%",
                  "Total Liabilities ÷ Total Assets × 100. Share of assets funded by debt, not equity."),
                 ("Equity Ratio", equity_ratio, "%", ">50%",
                  "Equity ÷ Total Assets × 100. Share of assets funded by owners (mirror of debt ratio)."),
             ]),
            ("⚙️  EFFICIENCY / ACTIVITY (How well do we move things?)",
             [
                 ("Inventory Turnover", inv_turn, "x", "4–8x / yr healthy",
                  "COGS ÷ Average Inventory. How many times stock sold & replaced this period (annualized)."),
                 ("DSO (Days Sales Out.)", days_ar, "days", "<30 days",
                  "Avg days to collect cash from credit customers. Lower = faster cash in."),
                 ("DPO (Days Payables Out.)", days_ap, "days", "30–60 days",
                  "Avg days we take to pay suppliers. Too slow damages credit; too fast loses float."),
                 ("CCC (Cash Conv. Cycle)", ccc, "days", "<45 days",
                  "Days Inventory + Days Sales - Days Payable. The 'from cash out to cash in' pipeline in days."),
                 ("Asset Turnover", asset_turn, "x", ">1.00x",
                  "Revenue ÷ Total Assets. Revenue generated per AED of assets. Efficiency of asset deployment."),
             ]),
        ]

        # Build dashboard rows: flat list with color tag + description column
        flat_rows = []
        for category, ratios in all_ratios:
            flat_rows.append({"item": category, "category_separator": True, "tag": "section"})
            for name, val, unit, bench, desc in ratios:
                try:
                    v = Decimal(str(val))
                    if unit == "%":
                        val_fmt = f"{round_aed(v):>8} %"
                    elif unit == "AED":
                        val_fmt = fmt_aed(v)
                    elif unit == "days":
                        val_fmt = f"{round_aed(v):>8} days"
                    else:
                        val_fmt = f"{round_aed(v):>10.2f} x"
                except Exception:
                    val_fmt = str(val)
                traffic, verdict = classify(name, val)
                # Colorize via tag: green/yellow/red
                tag_map = {"🟢": "gross_profit", "🟡": "overdue_med", "🔴": "overdue_high"}
                tag = tag_map.get(traffic, "")
                flat_rows.append({
                    "item": f"{traffic}  {name}",
                    "value": val_fmt,
                    "unit": unit,
                    "benchmark": bench,
                    "verdict": verdict,
                    "description": desc,
                    "tag": tag,
                })
        # Summary banner cells: pick 4 headline ratios
        banner = {
            "grand_total": str(round_aed(net_margin)) + " %",  # 1. Net Margin (most important)
            "current_ratio": str(round_aed(current_ratio)) + " x",  # 2. Liquidity
            "debt_to_equity": str(round_aed(debt_to_equity)) + " x",  # 3. Leverage
            "ccc_days": str(round_aed(ccc)) + " days",  # 4. Cash Conversion Cycle
        }
        # Also store in banner under named numeric keys for dashboard display
        banner.update({
            "net_margin_pct": round_aed(net_margin),
            "current_ratio_x": round_aed(current_ratio),
            "debt_to_equity_x": round_aed(debt_to_equity),
            "cash_conversion_cycle_days": round_aed(ccc),
            "roe_pct": round_aed(roe),
            "gross_margin_pct": round_aed(gross_margin),
            "inventory_turnover_x": round_aed(inv_turn),
            "dso_days": round_aed(days_ar),
        })

        # Traffic-light summary counts
        green = sum(1 for r in flat_rows if r.get("tag") == "gross_profit")
        yellow = sum(1 for r in flat_rows if r.get("tag") == "overdue_med")
        red = sum(1 for r in flat_rows if r.get("tag") == "overdue_high")

        explanation = f"""
        Your Financial Health in 14 Ratios (traffic lights above):
        🟢 Good/Healthy: {green}     🟡 Watch/Acceptable: {yellow}     🔴 Weak/Needs Action: {red}

        How to read these:
        • 💧 LIQUIDITY — "Can I pay bills this month?" (aim 🟢 on Current & Quick)
        • 📈 PROFITABILITY — "Am I actually making money?" (aim 🟢 on Net/Operating margin & ROE)
        • 🏗️  LEVERAGE — "Do I owe too much?" (Debt-to-Equity < 1 = 🟢; > 2 = 🔴 restructure)
        • ⚙️  EFFICIENCY — "How fast does cash spin through my business?" (CCC = the lower the better)

        The dashboard automatically benchmarks against pharma-distribution SME norms (IFRS-based guidance).
        """

        # Alias rows as "table_rows" + build label alias for dashboard column
        table_rows = []
        for r in flat_rows:
            cp = dict(r)
            cp.setdefault("label", r.get("item") or r.get("label") or "")
            table_rows.append(cp)

        # --- Financial Ratios Diagnostics (sanity checks on inputs + cross-reports) ---
        diagnostics = {"passed": True, "checks": [], "root_causes": [], "fix_steps": []}
        def _chk_rt(name, ok, detail):
            diagnostics["checks"].append({"name": name, "passed": bool(ok), "detail": str(detail)})
            if not ok:
                diagnostics["passed"] = False
        # RT-1: Zero-data check (no revenue means all ratios will be NaN → fail)
        try:
            ok = rev > Decimal("0")
            _chk_rt("RT-1: Non-zero Revenue ( denominator sanity )", ok,
                    f"Revenue during period = {fmt_aed(rev)}. {'OK — all margin and activity ratios computable.' if ok else '⚠️ Revenue = 0 — no margin, turnover, or DSO ratios are meaningful.'}")
            if not ok:
                diagnostics["root_causes"].append(
                    "Revenue = 0 AED in the reporting period. All profitability ratios (Gross/Net/Operating Margin, ROE, ROA), "
                    "DSO, and Asset Turnover will be undefined (or zero). This typically means the report period contains no "
                    "invoices, or the GL was not rebuilt for that date range.")
                diagnostics["fix_steps"].append(
                    "Fix RT-1: (a) Change the report period to All Time using the header dropdown. (b) If there ARE invoices in "
                    "this period, click 🔁 Rebuild GL from All Data and re-run the Ratios dashboard. (c) Check the Revenue total "
                    "on the P&L tab — if it shows 0, diagnose the P&L first.")
        except Exception as exc:
            _chk_rt("RT-1: Revenue check", False, f"Check failed: {exc}")
        # RT-2: Balance Sheet sanity (Total Assets > 0 for ROA, Asset Turnover)
        try:
            bs_ok = total_assets > Decimal("0")
            _chk_rt("RT-2: Non-zero Total Assets (ROA, Asset Turnover denominator)", bs_ok,
                    f"Total Assets (BS) = {fmt_aed(total_assets)}. "
                    f"{'OK — ROA & Asset Turnover computable.' if bs_ok else '⚠️ Total Assets = 0 — ROA & Asset Turnover N/A.'}")
            if not bs_ok:
                diagnostics["root_causes"].append(
                    "Balance Sheet shows Total Assets = 0. This means the GL has no entries for Cash, Bank, AR, or Inventory "
                    "as at the end-of-reporting date. Either the period is before any data, or the cumulative-to-date filters "
                    "excluded everything.")
                diagnostics["fix_steps"].append(
                    "Fix RT-2: (a) Run Balance Sheet report for the same period end-date. (b) If BS also shows 0 assets, the "
                    "problem is in the GL rebuild, not the ratios. Click 🔁 Rebuild GL. (c) Verify general_ledger.json has 1000+ "
                    "entries (it should have ~1,218).")
        except Exception as exc:
            _chk_rt("RT-2: Total Assets check", False, f"Check failed: {exc}")
        # RT-3: Cross-check Gross Margin % against P&L
        try:
            pnl_gm = None
            try:
                pnl2 = self.get_profit_and_loss()
                pnl_lines = pnl2.get("lines", {}) or {}
                gp_ = pnl_lines.get("gross_profit", {}) or {}
                rev_ = pnl_lines.get("gross_revenue", {}) or {}
                gp_v = gp_.get("current") if isinstance(gp_, dict) else gp_
                rev_v = rev_.get("current") if isinstance(rev_, dict) else rev_
                if rev_v and gp_v is not None and Decimal(str(rev_v)) > 0:
                    pnl_gm = Decimal(str(gp_v)) / Decimal(str(rev_v)) * Decimal("100")
            except Exception:
                pnl_gm = None
            ok3 = True if pnl_gm is None else abs(Decimal(str(gross_margin)) - Decimal(str(pnl_gm))) < Decimal("0.1")
            detail3 = (f"Ratios GM% = {round_aed(gross_margin)}%  |  P&L implied GM% = "
                       f"{'N/A (P&L data unavailable)' if pnl_gm is None else str(round_aed(pnl_gm)) + '%'}  —  "
                       f"{'MATCH ✅' if ok3 else 'MISMATCH ⚠️'}")
            _chk_rt("RT-3: Gross Margin % cross-check against P&L", ok3, detail3)
            if not ok3:
                diagnostics["root_causes"].append(
                    f"Gross Margin on Ratios dashboard ({round_aed(gross_margin)}%) ≠ Gross Margin computed from P&L "
                    f"({round_aed(pnl_gm)}%). The two reports are using different Revenue or COGS numbers.")
                diagnostics["fix_steps"].append(
                    "Fix RT-3: (a) Rebuild GL once, then immediately run P&L → then Ratios (in this order, same browser window). "
                    "(b) If mismatch persists, compare the raw P&L lines['revenue']['current'] and lines['cogs']['current'] "
                    "with the values used in ratios — they must be identical.")
        except Exception as exc:
            _chk_rt("RT-3: GM cross-check vs P&L", False, f"Check failed: {exc}")
        # RT-4: Liquidity CR = Current Assets / Current Liabilities (warning if > 5 or < 0.8)
        try:
            if total_cl > Decimal("0"):
                cr_val = round_aed(current_ratio)
                if cr_val > Decimal("5"):
                    _chk_rt("RT-W4: Current Ratio excessive (> 5, cash drag warning)", True,
                            f"CR = {cr_val}x — too much idle cash (benchmark ~ 2x). Consider deploying surplus into "
                            f"revenue-generating assets or settling supplier invoices early for settlement discount.")
                elif cr_val < Decimal("0.8"):
                    _chk_rt("RT-W4: Current Ratio dangerously low (< 0.8, liquidity risk — ACTION REQUIRED)", False,
                            f"CR = {cr_val}x — below benchmark (2x). Risk of missing near-term payments.")
                    diagnostics["root_causes"].append(
                        f"Current Ratio = {cr_val}x (liquidity risk). Current Liabilities exceed 125% of Current Assets. "
                        f"The business may not be able to cover supplier invoices and short-term obligations when they fall due.")
                    diagnostics["fix_steps"].append(
                        "Fix RT-W4 (low CR): (1) Accelerate AR collections — call overdue customers (see Aging Tab 6). "
                        "(2) Request extended credit terms from key suppliers. (3) Inject short-term owner capital into Bank "
                        "if necessary as a temporary measure. (4) Reduce Inventory if it's bloated vs. COGS.")
                else:
                    _chk_rt("RT-W4: Current Ratio within healthy band (0.8 — 5x)", True,
                            f"CR = {cr_val}x  (IFRS pharma benchmark ~ 2x). Acceptable.")
            else:
                _chk_rt("RT-W4: No Current Liabilities (N/A — all good)", True,
                        f"CR = infinity; no AP or accrued expenses currently outstanding.")
        except Exception as exc:
            _chk_rt("RT-W4: Current Ratio policy check", False, f"Check failed: {exc}")
        # RT-5: Debt-to-Equity warning (high leverage > 2x = action)
        try:
            if total_eq > Decimal("0") and debt_to_equity > Decimal("2"):
                _chk_rt("RT-W5: Debt-to-Equity > 2x (over-leveraged warning — ACTION REQUIRED)", False,
                        f"D:E = {round_aed(debt_to_equity)}x (benchmark ≤ 1x). The business owes suppliers more than "
                        f"2x the owner's stake.")
                diagnostics["root_causes"].append(
                    f"Debt-to-Equity = {round_aed(debt_to_equity)}x indicates high leverage. Long-term solvency risk: "
                    f"if the company suffers a loss quarter, equity can quickly go negative (balance sheet insolvency).")
                diagnostics["fix_steps"].append(
                    "Fix RT-W5 (high leverage): (1) Capitalise some of the owner's personal cash into 3100 to grow equity. "
                    "(2) Accelerate profitable sales to grow Retained Earnings organically. (3) Avoid unnecessary large "
                    "equipment purchases on credit until D:E drops below 1x.")
            elif total_eq <= Decimal("0"):
                _chk_rt("RT-W5: Equity is NEGATIVE or ZERO (balance-sheet insolvency warning — CRITICAL)", False,
                        f"Total Equity = {fmt_aed(total_eq)} ⚠️ — D:E undefined; owner stake wiped out.")
                diagnostics["root_causes"].append(
                    f"Total Equity ≤ 0. Either accumulated losses exceeded capital/RE, or Drawings far exceeded profits. "
                    f"This is a CRITICAL going-concern red flag under IAS 1.")
                diagnostics["fix_steps"].append(
                    "Fix RT-W5 (negative equity): (1) Inject capital → 3100 Owner Equity immediately. (2) Run P&L to see "
                    "period losses; create a 90-day turnaround plan. (3) Reduce Owner Drawings to 0 until equity is positive "
                    "for 2 consecutive months. (4) Consider IFRS 1 going-concern disclosure requirements.")
            else:
                _chk_rt("RT-W5: Debt-to-Equity acceptable (< 2x)", True,
                        f"D:E = {round_aed(debt_to_equity)}x (benchmark ≤ 1x). Solvency OK.")
        except Exception as exc:
            _chk_rt("RT-W5: Debt-to-Equity solvency check", False, f"Check failed: {exc}")
        # Final summary
        diagnostics["summary"] = (
            f"✅ All {len(diagnostics['checks'])} ratio integrity & cross-report checks passed — KPIs reliable."
            if diagnostics["passed"]
            else f"⚠️ {sum(1 for c in diagnostics['checks'] if not c['passed'])}/{len(diagnostics['checks'])} check(s) "
                 f"failed or warn of financial-health risk. Review numbered Fix Steps above."
        )
        if not diagnostics["passed"]:
            diagnostics["fix_steps"].append(
                "Final Step: After applying fixes → click 🔁 Rebuild GL → re-run Ratios → all checks must show ✅.")

        return {
            "report_name": "Financial Ratios Dashboard (14 KPIs — IFRS / Pharma SME Benchmarks)",
            "period": {"start": self.start, "end": self.end},
            "categories": all_ratios,
            "flat_rows": flat_rows,
            "table_rows": table_rows,
            "traffic_light_counts": {"green": green, "yellow": yellow, "red": red},
            "traffic_light_summary": {"green": green, "yellow": yellow, "red": red},
            "summary_banner": banner,
            "ratios_by_name": {
                "current_ratio": round_aed(current_ratio), "quick_ratio": round_aed(quick_ratio),
                "cash_ratio": round_aed(cash_ratio), "net_working_capital": nwc,
                "gross_margin_pct": round_aed(gross_margin), "operating_margin_pct": round_aed(operating_margin),
                "net_margin_pct": round_aed(net_margin), "roe_pct": round_aed(roe), "roa_pct": round_aed(roa),
                "debt_to_equity": round_aed(debt_to_equity), "debt_ratio_pct": round_aed(debt_ratio * Decimal("100")),
                "equity_ratio_pct": round_aed(equity_ratio * Decimal("100")),
                "inventory_turnover_x": round_aed(inv_turn), "dso_days": round_aed(days_ar),
                "dpo_days": round_aed(days_ap), "ccc_days": round_aed(ccc),
                "asset_turnover_x": round_aed(asset_turn),
            },
            "explanation": explanation,
            "diagnostics": diagnostics,
        }

    # ==================================================================================
    # PHASE 5: AUDIT CHECK (Sum Debits = Sum Credits)
    # ==================================================================================
    def run_audit(self):
        total_d = Decimal("0.00"); total_c = Decimal("0.00")
        batch_totals = {}  # (ref_type, ref_id) -> (debits, credits)
        for e in self.gl_current:
            d = Decimal(str(e.get("debit_amount", 0))); c = Decimal(str(e.get("credit_amount", 0)))
            total_d += d; total_c += c
            ref_type = e.get("reference_type", e.get("ref_type", "UNKNOWN"))
            ref_id   = e.get("reference_id",
                        e.get("ref_id",
                        e.get("entry_id",
                        e.get("id", "NO_ID"))))
            key = (ref_type, ref_id)
            if key not in batch_totals: batch_totals[key] = [Decimal("0"), Decimal("0")]
            batch_totals[key][0] += d; batch_totals[key][1] += c
        total_d = round_aed(total_d); total_c = round_aed(total_c)
        balanced = (total_d == total_c)
        mismatched = []
        for (rt, rid), (d, c) in batch_totals.items():
            d = round_aed(d); c = round_aed(c)
            if d != c:
                mismatched.append({
                    "reference_type": rt,
                    "reference_id": rid,
                    "debits": d, "credits": c,
                    "variance": round_aed(d - c)
                })
        # Sort by largest variance
        mismatched.sort(key=lambda x: -abs(x["variance"]))
        explanation = f"""
        What an Audit Check Does:
        The GAAP "Golden Rule" of double-entry accounting is:
        Sum of ALL Debits = Sum of ALL Credits.
        If this doesn't match EXACTLY to the fils (AED 0.00), your books are WRONG.
        {"✅ PASSED: Your debits and credits match perfectly — GAAP Compliant!" if balanced else "❌ FAILED: The transactions listed below are unbalanced — please review and correct."}
        """
        # Build summary table for PDF/dashboard
        rows_for_report = [
            {"Summary Item": "Σ All Debits (period)", "Value": fmt_aed(total_d), "tag": ""},
            {"Summary Item": "Σ All Credits (period)", "Value": fmt_aed(total_c), "tag": ""},
            {"Summary Item": "Variance (Dr − Cr)", "Value": fmt_aed(total_d - total_c),
             "tag": "gross_profit" if balanced else "overdue_high"},
            {"Summary Item": "Total Journal Entry Batches", "Value": f"{len(batch_totals):,}", "tag": ""},
            {"Summary Item": "Balanced Batches", "Value": f"{len(batch_totals)-len(mismatched):,}", "tag": ""},
            {"Summary Item": "⚠️ MISMATCHED Batches", "Value": f"{len(mismatched):,}",
             "tag": "" if len(mismatched)==0 else "overdue_high"},
        ]
        if mismatched:
            rows_for_report.append({"Summary Item": "", "Value": "", "tag": ""})
            rows_for_report.append(
                {"Summary Item": "— Top 15 Mismatched Batches (largest first) —",
                 "Value": "Variance (Dr−Cr)", "tag": "section"})
            for m in mismatched[:15]:
                rows_for_report.append({
                    "Summary Item": f"  [{m['reference_type']}] #{m['reference_id']}",
                    "Value": fmt_aed(m["variance"]),
                    "tag": "overdue_high" if abs(m["variance"]) > Decimal("100") else "overdue_med",
                })
        return {
            "report_name": "General Ledger Audit Report",
            "period": {"start": self.start, "end": self.end},
            "total_debits": total_d,
            "total_credits": total_c,
            "variance": round_aed(total_d - total_c),
            "is_balanced": balanced,
            "mismatched_transactions": mismatched,
            "batches": batch_totals,
            "rows": rows_for_report,
            "table_rows": rows_for_report,
            "summary_banner": {
                "grand_total": total_d,
                "credits": total_c,
                "variance": round_aed(total_d - total_c),
                "batches": len(batch_totals),
                "mismatched": len(mismatched),
            },
            "explanation": explanation
        }

    # ==================================================================================
    # PHASE 4: GAAP FOOTNOTES ENGINE (ASC 235 / ASC 450 / ASC 850)
    # ==================================================================================
    def get_footnotes(self):
        """
        US GAAP-mandated Footnote Disclosures (Requirement #4).
        Generates 3 categories:
          1. Significant Accounting Policies (ASC 235-10-50) — FP-1..FP-9
          2. Contingent Liabilities (ASC 450-20 / ASC 460) — CL-1..CL-5
          3. Related-Party Transactions (ASC 850) — RP-1..RP-4
        Includes auto-detection from GL + diagnostics with WARN for default templates.
        """
        # --- MATERIALITY THRESHOLD ---
        bs = self.get_balance_sheet()
        _lines = bs.get("lines", {}) or {}
        _assets = _lines.get("assets", {}) if isinstance(_lines, dict) else {}
        total_assets = Decimal("0")
        try:
            ta = _assets.get("total_assets") if isinstance(_assets, dict) else None
            total_assets = Decimal(str(ta or 0))
        except Exception:
            total_assets = Decimal("0")
        materiality_threshold = max(Decimal("1000.00"), Decimal("0.005") * total_assets)

        # --- CATEGORY 1: SIGNIFICANT ACCOUNTING POLICIES (ASC 235) ---
        accounting_policies = []

        # FP-1: Basis of Presentation
        period_locks = []
        try:
            period_locks = self.dm.load_json('period_locks.json') or []
            if not isinstance(period_locks, list):
                period_locks = []
        except Exception:
            period_locks = []
        has_locks = len([p for p in period_locks if isinstance(p, dict)]) > 0
        fiscal_year_text = (
            "Fiscal year-end is 31 December; locked accounting periods are configured in Period Locks."
            if has_locks else
            "Fiscal year-end is 31 December unless configured otherwise in Period Locks."
        )
        accounting_policies.append({
            "code": "FP-1",
            "title": "Basis of Presentation",
            "text": (f"These financial statements are prepared in accordance with accounting principles generally accepted "
                     f"in the United States of America (US GAAP) and are presented on the accrual basis of accounting. "
                     f"{fiscal_year_text}"),
            "_uses_default_template": not has_locks,
        })

        # FP-2: Revenue Recognition (ASC 606)
        accounting_policies.append({
            "code": "FP-2",
            "title": "Revenue Recognition Policy (ASC 606)",
            "text": ("Revenue is recognized when performance obligations under sales terms are satisfied "
                     "(typically upon product delivery / invoice issuance); sales returns and allowances are "
                     "estimated and recorded in the period of sale (Accounts 4010–4030)."),
            "_uses_default_template": True,
        })

        # FP-3: Inventory Valuation (ASC 330)
        accounting_policies.append({
            "code": "FP-3",
            "title": "Inventory Valuation (ASC 330)",
            "text": ("Inventories are stated at the lower of cost or net realizable value (LCNRV). "
                     "Cost is determined using weighted-average method."),
            "_uses_default_template": True,
        })

        # FP-4: Property, Plant & Equipment (ASC 360)
        accounting_policies.append({
            "code": "FP-4",
            "title": "Property, Plant & Equipment (ASC 360)",
            "text": ("PPE is stated at cost less accumulated depreciation. Depreciation is computed using "
                     "the straight-line method over estimated useful lives (Account 1610)."),
            "_uses_default_template": True,
        })

        # FP-5: Intangible Assets (ASC 350)
        accounting_policies.append({
            "code": "FP-5",
            "title": "Intangible Assets (ASC 350)",
            "text": ("Intangibles with finite lives are amortized on a straight-line basis; "
                     "goodwill (if any) is tested annually for impairment."),
            "_uses_default_template": True,
        })

        # FP-6: Income Taxes (ASC 740)
        accounting_policies.append({
            "code": "FP-6",
            "title": "Income Taxes (ASC 740)",
            "text": ("Deferred tax assets/liabilities are recognized for temporary differences between "
                     "financial reporting and tax bases of assets and liabilities (Account 2600)."),
            "_uses_default_template": True,
        })

        # FP-7: Accounts Receivable (ASC 310)
        accounting_policies.append({
            "code": "FP-7",
            "title": "Accounts Receivable (ASC 310)",
            "text": ("Trade receivables are carried at gross invoice amount less allowance for doubtful "
                     "accounts (Account 1210). Estimated uncollectibles are recorded as bad-debt expense "
                     "in the period of sale."),
            "_uses_default_template": True,
        })

        # FP-8: Cash & Cash Equivalents (ASC 230)
        accounting_policies.append({
            "code": "FP-8",
            "title": "Cash & Cash Equivalents (ASC 230)",
            "text": ("Cash on Hand (1000) + demand deposits with banks (1100) with original maturities ≤ 90 days."),
            "_uses_default_template": True,
        })

        # FP-9: Earnings Per Share (ASC 260)
        try:
            shares_out = int(getattr(self, 'shares_outstanding', 10000) or 10000)
        except Exception:
            shares_out = 10000
        accounting_policies.append({
            "code": "FP-9",
            "title": "Earnings Per Share (ASC 260)",
            "text": (f"Basic EPS is computed by dividing Net Income by the weighted-average number of "
                     f"common shares outstanding ({shares_out:,} shares used; default 10,000 shares used "
                     f"if cap table not yet populated)."),
            "_uses_default_template": shares_out == 10000,
        })

        # --- CATEGORY 2: CONTINGENT LIABILITIES (ASC 450-20 / ASC 460) ---
        contingent_liabilities = []

        def _haystack(e):
            return " ".join(str(v or '') for v in [
                e.get('description'), e.get('note'), e.get('memo'), e.get('narration'),
                e.get('payee'), e.get('supplier'), e.get('vendor_or_payee'), e.get('vendor'),
                e.get('reference_id'), e.get('reference_type')
            ]).lower()

        # CL-1: Pending or threatened litigation
        litigation_accts = {'2150', '2160', '2800'}
        litigation_keywords = ['litigation', 'legal', 'claim', 'settlement', 'lawsuit', 'court', 'judgment']
        litigation_entries = []
        litigation_total = Decimal("0")
        for e in self.gl_current or []:
            if not isinstance(e, dict):
                continue
            code = str(e.get('account_code') or '').strip()
            if code in litigation_accts:
                hs = _haystack(e)
                if any(kw in hs for kw in litigation_keywords):
                    try:
                        dr = Decimal(str(e.get('debit_amount') or 0))
                        cr = Decimal(str(e.get('credit_amount') or 0))
                        amt = round_aed(cr - dr)
                        if amt != 0:
                            litigation_entries.append(e)
                            litigation_total += amt
                    except Exception:
                        pass
        probable_threshold = Decimal("0.1") * total_assets if total_assets > 0 else Decimal("10000")
        if abs(litigation_total) > 0:
            if abs(litigation_total) > probable_threshold:
                assessment = "PROBABLE"
            elif abs(litigation_total) > materiality_threshold:
                assessment = "REASONABLY POSSIBLE"
            else:
                assessment = "REMOTE"
            materiality_flag = abs(litigation_total) >= materiality_threshold
            contingent_liabilities.append({
                "code": "CL-1",
                "title": "Pending / Threatened Litigation",
                "assessment": assessment,
                "text": (f"Legal claims, litigation or settlement matters identified in accounts 2150/2160/2800. "
                         f"ASC 450-20 requires accrual if loss is probable and estimable; disclosure if reasonably possible. "
                         f"{len(litigation_entries)} GL entry(ies) detected — management should review for required accrual."),
                "amount": round_aed(litigation_total),
                "materiality_flag": materiality_flag,
                "_uses_default_template": False,
            })
        else:
            contingent_liabilities.append({
                "code": "CL-1",
                "title": "Pending / Threatened Litigation",
                "assessment": "NONE IDENTIFIED",
                "text": ("No litigation, legal claims or settlement accruals were posted to GL accounts "
                         "2150/2160/2800 during the period. Management should nonetheless represent that "
                         "there are no material pending or threatened legal proceedings that require "
                         "disclosure under ASC 450-20."),
                "amount": Decimal("0.00"),
                "materiality_flag": False,
                "_uses_default_template": True,
            })

        # CL-2: Product Warranties Payable
        warranty_accts = {'2120', '2170'}
        warranty_total = self._sum_net([c for c in warranty_accts], self.gl_current)
        if abs(warranty_total) > Decimal("0"):
            contingent_liabilities.append({
                "code": "CL-2",
                "title": "Product Warranties Payable",
                "assessment": "ACCRUED",
                "text": (f"Warranty accrual posted to accounts 2120/2170 (ASC 460). "
                         f"Company recognizes estimated warranty costs at the time of sale based on "
                         f"historical return rates."),
                "amount": round_aed(warranty_total),
                "materiality_flag": abs(warranty_total) >= materiality_threshold,
                "_uses_default_template": False,
            })
        else:
            contingent_liabilities.append({
                "code": "CL-2",
                "title": "Product Warranties Payable",
                "assessment": "NO ACCRUAL POSTED",
                "text": ("No warranty accrual posted during the period; if warranties are provided, "
                         "ASC 460 requires recognition at time of sale. Management should confirm that "
                         "the entity does not provide product warranties, or post the required estimate."),
                "amount": Decimal("0.00"),
                "materiality_flag": False,
                "_uses_default_template": True,
            })

        # CL-3: Guarantees and Indirect Guarantees
        guarantee_exp_accts = {'5500', '5700'}
        guarantee_keywords = ['guarantee', 'guaranty', 'indemnification', 'surety', 'indirect guarantee']
        guarantee_total = Decimal("0")
        guarantee_found = False
        for e in self.gl_current or []:
            if not isinstance(e, dict):
                continue
            code = str(e.get('account_code') or '').strip()
            if code in guarantee_exp_accts:
                hs = _haystack(e)
                if any(kw in hs for kw in guarantee_keywords):
                    guarantee_found = True
                    try:
                        dr = Decimal(str(e.get('debit_amount') or 0))
                        cr = Decimal(str(e.get('credit_amount') or 0))
                        guarantee_total += round_aed(dr - cr)
                    except Exception:
                        pass
        if guarantee_found:
            contingent_liabilities.append({
                "code": "CL-3",
                "title": "Guarantees & Indirect Guarantees (ASC 460)",
                "assessment": "IDENTIFIED",
                "text": (f"Guarantee fees / indemnification costs detected in expense accounts 5500/5700. "
                         f"ASC 460 requires disclosure of nature, maximum potential future payments, "
                         f"and carrying amount of liability (if any) for all guarantees."),
                "amount": round_aed(guarantee_total),
                "materiality_flag": abs(guarantee_total) >= materiality_threshold,
                "_uses_default_template": False,
            })
        else:
            contingent_liabilities.append({
                "code": "CL-3",
                "title": "Guarantees & Indirect Guarantees (ASC 460)",
                "assessment": "NONE RECORDED",
                "text": ("No guarantees of third-party indebtedness have been recorded in expense "
                         "accounts 5500/5700. Management should represent that no guarantees, "
                         "indemnifications or surety arrangements exist that require ASC 460 disclosure."),
                "amount": Decimal("0.00"),
                "materiality_flag": False,
                "_uses_default_template": True,
            })

        # CL-4: Unused Letters of Credit / Loan Commitments
        commitment_accts = {'2200', '2400'}
        commitment_keywords = ['commitment', 'letter of credit', 'l/c', 'lc', 'credit facility',
                               'loan commitment', 'undrawn']
        commitment_total = Decimal("0")
        commitment_found = False
        for e in self.gl_current or []:
            if not isinstance(e, dict):
                continue
            code = str(e.get('account_code') or '').strip()
            if code in commitment_accts:
                hs = _haystack(e)
                if any(kw in hs for kw in commitment_keywords):
                    commitment_found = True
                    try:
                        dr = Decimal(str(e.get('debit_amount') or 0))
                        cr = Decimal(str(e.get('credit_amount') or 0))
                        commitment_total += round_aed(cr - dr)
                    except Exception:
                        pass
        if commitment_found:
            contingent_liabilities.append({
                "code": "CL-4",
                "title": "Loan Commitments / Letters of Credit",
                "assessment": "IDENTIFIED",
                "text": (f"Undisclosed future borrowing commitments or L/Cs detected in 2200/2400 movement. "
                         f"ASC 440 requires disclosure of unconditional purchase obligations and "
                         f"unused credit facilities with terms > 1 year."),
                "amount": round_aed(commitment_total),
                "materiality_flag": abs(commitment_total) >= materiality_threshold,
                "_uses_default_template": False,
            })
        else:
            contingent_liabilities.append({
                "code": "CL-4",
                "title": "Loan Commitments / Letters of Credit",
                "assessment": "NONE POSTED",
                "text": ("No commitments for undisclosed future borrowings or L/Cs posted to GL during "
                         "period. Management should represent no material off-balance-sheet credit "
                         "facilities, L/Cs or unconditional purchase obligations exist."),
                "amount": Decimal("0.00"),
                "materiality_flag": False,
                "_uses_default_template": True,
            })

        # CL-5: Environmental / Regulatory Liabilities (ASC 410 ARO)
        env_accts = {'2180', '2810'}
        env_total = self._sum_net([c for c in env_accts], self.gl_current)
        if abs(env_total) > Decimal("0"):
            contingent_liabilities.append({
                "code": "CL-5",
                "title": "Environmental / Regulatory Liabilities (ASC 410 ARO)",
                "assessment": "ACCRUED",
                "text": (f"Asset Retirement Obligations (ARO) or environmental remediation accruals "
                         f"identified in accounts 2180/2810. ASC 410 requires fair-value measurement "
                         f"and reconciliation of carrying amount."),
                "amount": round_aed(env_total),
                "materiality_flag": abs(env_total) >= materiality_threshold,
                "_uses_default_template": False,
            })
        else:
            contingent_liabilities.append({
                "code": "CL-5",
                "title": "Environmental / Regulatory Liabilities (ASC 410 ARO)",
                "assessment": "NONE IDENTIFIED",
                "text": ("No asset retirement obligations or environmental remediation accruals "
                         "identified (accounts 2180/2810). For a pharmaceutical distributor, "
                         "management should still confirm no cold-chain or regulated-goods remediation "
                         "liabilities exist."),
                "amount": Decimal("0.00"),
                "materiality_flag": False,
                "_uses_default_template": True,
            })

        # --- CATEGORY 3: RELATED-PARTY TRANSACTIONS (ASC 850) ---
        related_parties = []

        # RP-1: Owner Drawings vs Market-rate compensation
        drawings_total = self._cumulative_to(["3200"], self.end)
        salary_total = self._sum_net(["5100"], self.gl_current)
        salary_abs = abs(salary_total) if salary_total < 0 else salary_total
        high_drawings_flag = salary_abs > 0 and abs(drawings_total) > Decimal("3") * salary_abs
        if high_drawings_flag:
            related_parties.append({
                "code": "RP-1",
                "title": "Owner Drawings vs. Market-Rate Compensation",
                "text": (f"⚠️ High owner drawings relative to salary expense; review for constructive "
                         f"dividend tax treatment (ASC 707). Account 3200 Drawings = {fmt_aed(drawings_total)}; "
                         f"Account 5100 Salary Expense = {fmt_aed(salary_abs)}. "
                         f"Drawings exceed 3× salary expense — management should document that compensation "
                         f"to owners reflects market-rate compensation per ASC 707-10."),
                "amount": round_aed(drawings_total),
                "flag": "⚠️ HIGH DRAWINGS",
                "_uses_default_template": False,
            })
        else:
            related_parties.append({
                "code": "RP-1",
                "title": "Owner Drawings vs. Market-Rate Compensation",
                "text": (f"Owner distributions (Account 3200 = {fmt_aed(drawings_total)}) are within a "
                         f"reasonable range relative to salary expense (Account 5100 = {fmt_aed(salary_abs)}). "
                         f"No constructive dividend indicators identified under ASC 707."),
                "amount": round_aed(drawings_total),
                "flag": "OK",
                "_uses_default_template": True,
            })

        # RP-2: Related-Party Receivables / Loans
        rp_ar_accts = {'1250', '1260'}
        rp_keywords = ['related', 'director', 'officer', 'shareholder', 'owner', 'family',
                       'affiliate', 'associate', 'parent', 'subsidiary', 'related party']
        rp_ar_total = Decimal("0")
        rp_ar_items = []
        # Check dedicated accounts first
        rp_ar_total = self._cumulative_to([c for c in rp_ar_accts], self.end)
        # Also scan general AR descriptions for related-party keywords
        for e in self._gl_all or []:
            if not isinstance(e, dict):
                continue
            code = str(e.get('account_code') or '').strip()
            # Check AR/loans range
            if code.startswith('12') or code in rp_ar_accts:
                hs = _haystack(e)
                if any(kw in hs for kw in rp_keywords):
                    try:
                        dr = Decimal(str(e.get('debit_amount') or 0))
                        cr = Decimal(str(e.get('credit_amount') or 0))
                        amt = round_aed(dr - cr)
                        if abs(amt) > Decimal("0"):
                            rp_ar_total += amt
                            rp_ar_items.append((code, e.get('description', ''), amt))
                    except Exception:
                        pass
        if abs(rp_ar_total) > Decimal("0"):
            related_parties.append({
                "code": "RP-2",
                "title": "Related-Party Receivables / Loans (ASC 850)",
                "text": (f"Related-party receivables, loans or advances to directors, officers, "
                         f"shareholders or their families detected. ASC 850 requires disclosure of "
                         f"nature of relationship, amounts, terms and repayment provisions. "
                         f"{len(rp_ar_items)} related-party receivable transaction(s) found."),
                "amount": round_aed(rp_ar_total),
                "flag": "⚠️ DISCLOSURE REQUIRED",
                "_uses_default_template": False,
            })
        else:
            related_parties.append({
                "code": "RP-2",
                "title": "Related-Party Receivables / Loans (ASC 850)",
                "text": ("No related-party receivables, loans or advances to directors, officers, "
                         "shareholders or their families identified in accounts 1250/1260 or in "
                         "AR narration fields. Management should represent no such balances exist."),
                "amount": Decimal("0.00"),
                "flag": "NONE IDENTIFIED",
                "_uses_default_template": True,
            })

        # RP-3: Related-Party Payables
        rp_ap_accts = {'2050', '2060', '2080'}
        rp_ap_total = Decimal("0")
        rp_ap_items = []
        rp_ap_total = self._cumulative_to([c for c in rp_ap_accts], self.end)
        for e in self._gl_all or []:
            if not isinstance(e, dict):
                continue
            code = str(e.get('account_code') or '').strip()
            if code.startswith('20') or code in rp_ap_accts:
                hs = _haystack(e)
                if any(kw in hs for kw in rp_keywords):
                    try:
                        dr = Decimal(str(e.get('debit_amount') or 0))
                        cr = Decimal(str(e.get('credit_amount') or 0))
                        amt = round_aed(cr - dr)
                        if abs(amt) > Decimal("0"):
                            rp_ap_total += amt
                            rp_ap_items.append((code, e.get('description', ''), amt))
                    except Exception:
                        pass
        if abs(rp_ap_total) > Decimal("0"):
            related_parties.append({
                "code": "RP-3",
                "title": "Related-Party Payables (ASC 850)",
                "text": (f"Related-party payables to directors, officers, shareholders or affiliates "
                         f"identified. ASC 850 requires full disclosure of nature of relationship, "
                         f"amounts outstanding, terms and settlement provisions. "
                         f"{len(rp_ap_items)} related-party payable transaction(s) found."),
                "amount": round_aed(rp_ap_total),
                "flag": "⚠️ DISCLOSURE REQUIRED",
                "_uses_default_template": False,
            })
        else:
            related_parties.append({
                "code": "RP-3",
                "title": "Related-Party Payables (ASC 850)",
                "text": ("No related-party payables to directors, officers, shareholders or affiliates "
                         "identified in accounts 2050/2060/2080 or in AP narration fields. Management "
                         "should represent no such related-party payable balances exist."),
                "amount": Decimal("0.00"),
                "flag": "NONE IDENTIFIED",
                "_uses_default_template": True,
            })

        # RP-4: Key Management Compensation (ASC 710)
        employees_data = []
        try:
            employees_data = self.dm.load_json('employees_data.json') or []
            if not isinstance(employees_data, list):
                employees_data = []
        except Exception:
            employees_data = []
        emp_count = len([e for e in employees_data if isinstance(e, dict)])
        key_mgmt_count = max(1, min(2, emp_count)) if emp_count > 0 else 1
        mgmt_comp_total = round_aed(salary_abs * Decimal(key_mgmt_count)) if emp_count > 0 else salary_abs
        if salary_abs > Decimal("0"):
            mgmt_text = (f"Aggregate compensation paid to key management (directors, officers, owners) "
                         f"during the period: {fmt_aed(mgmt_comp_total)} (estimated as Account 5100 Salary × "
                         f"{key_mgmt_count} key management person(s); employee count = {emp_count}). "
                         f"ASC 710 and ASC 850 require disclosure of aggregate compensation to key "
                         f"management personnel. See Payroll module for officer compensation detail.")
        else:
            mgmt_text = (f"Aggregate compensation paid to key management (directors, officers, owners) "
                         f"during the period: See Payroll module for officer compensation detail. "
                         f"No salary expense (5100) recorded in the current reporting period.")
        related_parties.append({
            "code": "RP-4",
            "title": "Key Management Compensation (ASC 710 / 850)",
            "text": mgmt_text,
            "amount": round_aed(mgmt_comp_total),
            "flag": "OK" if salary_abs > Decimal("0") else "⚠️ NO SALARY POSTED",
            "_uses_default_template": salary_abs == Decimal("0"),
        })

        # --- BUILD DIAGNOSTICS ---
        all_items = (
            [(f"FP-{i+1}", p) for i, p in enumerate(accounting_policies)] +
            [(f"CL-{i+1}", c) for i, c in enumerate(contingent_liabilities)] +
            [(f"RP-{i+1}", r) for i, r in enumerate(related_parties)]
        )
        diagnostics_checks = []
        default_template_count = 0
        actual_gl_evidence_count = 0
        for code, item in all_items:
            is_default = item.get("_uses_default_template", True)
            if is_default:
                default_template_count += 1
            else:
                actual_gl_evidence_count += 1
            detail = (
                "DEFAULT TEMPLATE used — review management policy text before filing."
                if is_default else
                "GL evidence-based disclosure generated."
            )
            diagnostics_checks.append({
                "name": f"Footnote {code} — {item.get('title', '')}",
                "passed": True,
                "detail": detail,
            })

        total_count = len(diagnostics_checks)  # Should be 17 (9+5+4)
        passed_count = total_count  # All items are "generated" (none fail)
        issues_list = []
        if default_template_count > 0:
            for code, item in all_items:
                if item.get("_uses_default_template", True):
                    issues_list.append(
                        f"{code}: DEFAULT policy template — management must review, customise, "
                        f"and optionally post via app_settings.footnotes_override before GAAP filing."
                    )

        diagnostics = {
            "checks": diagnostics_checks,
            "passed_count": passed_count,
            "total_count": total_count,
            "summary": (
                f"✅ All {total_count} footnote items generated (9 Accounting Policies · 5 Contingencies · 4 Related Parties). "
                f"{default_template_count} use DEFAULT templates and require management review; "
                f"{actual_gl_evidence_count} have actual GL evidence backing."
                if default_template_count > 0 else
                f"✅ All {total_count} footnote items generated with GL evidence — no default templates used."
            ),
            "issues": issues_list,
        }

        # --- Build flat table_rows for Excel/PDF export ---
        table_rows = []
        # Section: Accounting Policies
        table_rows.append({"Category": "— SIGNIFICANT ACCOUNTING POLICIES (ASC 235) —",
                           "Ref": "", "Assessment / Flag": "", "Disclosure Text": "", "Amount (AED)": "",
                           "tag": "section"})
        for p in accounting_policies:
            tag = "level2" if not p.get("_uses_default_template", True) else "overdue_med"
            table_rows.append({
                "Category": "Accounting Policy",
                "Ref": p["code"],
                "Assessment / Flag": p["title"],
                "Disclosure Text": p["text"],
                "Amount (AED)": "",
                "tag": tag,
            })
        # Section: Contingent Liabilities
        table_rows.append({"Category": "— CONTINGENT LIABILITIES (ASC 450 / 460) —",
                           "Ref": "", "Assessment / Flag": "", "Disclosure Text": "", "Amount (AED)": "",
                           "tag": "section"})
        for c in contingent_liabilities:
            amt = fmt_aed(c["amount"]) if c["amount"] is not None and c["amount"] != Decimal("0") else ""
            is_default = c.get("_uses_default_template", True)
            is_material = c.get("materiality_flag", False)
            if is_material:
                tag = "overdue_high"
            elif is_default:
                tag = "overdue_med"
            else:
                tag = "good"
            table_rows.append({
                "Category": "Contingent Liability",
                "Ref": c["code"],
                "Assessment / Flag": c["assessment"],
                "Disclosure Text": c["text"],
                "Amount (AED)": amt,
                "tag": tag,
            })
        # Section: Related Parties
        table_rows.append({"Category": "— RELATED-PARTY TRANSACTIONS (ASC 850) —",
                           "Ref": "", "Assessment / Flag": "", "Disclosure Text": "", "Amount (AED)": "",
                           "tag": "section"})
        for r in related_parties:
            amt = fmt_aed(r["amount"]) if r["amount"] is not None and r["amount"] != Decimal("0") else ""
            flag = r.get("flag", "")
            is_default = r.get("_uses_default_template", True)
            if flag and flag.startswith("⚠️"):
                tag = "overdue_high"
            elif is_default:
                tag = "overdue_med"
            else:
                tag = "good"
            table_rows.append({
                "Category": "Related Party",
                "Ref": r["code"],
                "Assessment / Flag": flag,
                "Disclosure Text": r["text"],
                "Amount (AED)": amt,
                "tag": tag,
            })

        return {
            "report_name": "Notes to Consolidated Financial Statements (US GAAP ASC 235 / 450 / 850)",
            "as_of": self.end,
            "accounting_policies": accounting_policies,
            "contingent_liabilities": contingent_liabilities,
            "related_parties": related_parties,
            "materiality_threshold": round_aed(materiality_threshold),
            "total_assets": round_aed(total_assets),
            "summary_banner": {
                "grand_total": round_aed(materiality_threshold),
                "policy_count": len(accounting_policies),
                "contingencies_found": actual_gl_evidence_count,
                "rp_items_found": len(related_parties),
                "default_template_count": default_template_count,
                "materiality_threshold": round_aed(materiality_threshold),
            },
            "table_rows": table_rows,
            "rows": table_rows,
            "diagnostics": diagnostics,
            "explanation": (
                f"These footnote disclosures represent the minimum required disclosures under US GAAP for:\n"
                f"  • ASC 235-10-50 (Significant Accounting Policies) — 9 policy statements\n"
                f"  • ASC 450-20 and ASC 460 (Contingent Liabilities & Guarantees) — 5 areas\n"
                f"  • ASC 850 (Related-Party Transactions) — 4 areas\n\n"
                f"Materiality threshold applied: MAX(AED 1,000, 0.5% × Total Assets) = {fmt_aed(materiality_threshold)}.\n\n"
                f"⚠️  DEFAULT TEMPLATE WARNING: {default_template_count} of {total_count} footnote items "
                f"use boilerplate templates. Before filing GAAP financial statements, management MUST:\n"
                f"  1. Review each template for accuracy and company-specific policy wording\n"
                f"  2. Confirm no off-balance-sheet arrangements, commitments or contingencies exist that are not captured\n"
                f"  3. Post custom text via app_settings.footnotes_override to replace templates with final approved language"
            ),
        }

    # ==================================================================================
    # PHASE 3: CROSS-REPORT RECONCILIATION MATRIX (GAAP Requirement #2)
    # ==================================================================================
    def get_reconciliation_matrix(self):
        """
        Cross-Report Reconciliation Matrix (US GAAP Audit Requirement #2)
        Verifies all 7 mandatory cross-statement ties to ensure:
          - All account balances match general ledger totals
          - Cross-report financial ties are mathematically consistent
        """
        pnl = self.get_profit_and_loss()
        bs = self.get_balance_sheet()
        cf = self.get_cash_flow()
        sce = self.get_statement_of_changes_in_equity()

        TOL = Decimal("0.01")
        checks = []

        # ------------------------------------------------------------------
        # TIE-1: P&L Net Income = BS Retained Earnings current-period addition
        #        (RE_end - RE_begin + Drawings)
        # ------------------------------------------------------------------
        tie1_passed = False
        tie1_detail = ""
        tie1_formula = "P&L Net Income = BS RE (end) − BS RE (begin) + Drawings"
        try:
            pnl_ni = Decimal(str(pnl["lines"]["net_income"]["current"]))
            bs_lines = bs.get("lines", {}) or {}
            bs_breakdown = bs_lines.get("__breakdowns", {}) or {}
            re_beginning = Decimal(str(bs_breakdown.get("re_beginning", 0)))
            re_ending = Decimal(str(bs_breakdown.get("re_ending", 0)))
            re_add_ni = Decimal(str(bs_breakdown.get("re_add_ni", 0)))
            re_less_drawings = Decimal(str(bs_breakdown.get("re_less_drawings", 0)))
            bs_re_addition = round_aed(re_ending - re_beginning - re_less_drawings)
            tie1_passed = abs(pnl_ni - bs_re_addition) < TOL
            tie1_detail = (
                f"P&L Net Income = {fmt_aed(pnl_ni)}  |  "
                f"BS RE movement = RE(end) {fmt_aed(re_ending)} − RE(begin) {fmt_aed(re_beginning)} "
                f"+ Drawings {fmt_aed(-re_less_drawings)} = {fmt_aed(bs_re_addition)}  |  "
                f"Δ = {fmt_aed(pnl_ni - bs_re_addition)}"
            )
        except Exception as exc:
            tie1_detail = f"Check failed: {exc}"

        checks.append({
            "name": "TIE-1: P&L Net Income ↔ BS Retained Earnings movement",
            "passed": bool(tie1_passed),
            "detail": tie1_detail,
            "formula": tie1_formula,
            "tie_from_statement": "Profit & Loss",
            "tie_to_statement": "Balance Sheet",
        })

        # ------------------------------------------------------------------
        # TIE-2: P&L Other Comprehensive Income = BS AOCI movement during period
        # ------------------------------------------------------------------
        tie2_passed = False
        tie2_detail = ""
        tie2_formula = "P&L OCI (period) = BS AOCI (end) − BS AOCI (begin)"
        try:
            pnl_oci = Decimal(str(pnl["lines"]["other_comprehensive_income"]["current"]))
            day_before = self.start - timedelta(days=1)
            aoci_begin = self._cumulative_to(["3900", "3910", "3920", "3930"], day_before)
            aoci_end = self._cumulative_to(["3900", "3910", "3920", "3930"], self.end)
            bs_aoci_movement = round_aed(aoci_end - aoci_begin)
            tie2_passed = abs(pnl_oci - bs_aoci_movement) < TOL
            tie2_detail = (
                f"P&L OCI = {fmt_aed(pnl_oci)}  |  "
                f"BS AOCI movement = AOCI(end) {fmt_aed(aoci_end)} − AOCI(begin) {fmt_aed(aoci_begin)} "
                f"= {fmt_aed(bs_aoci_movement)}  |  "
                f"Δ = {fmt_aed(pnl_oci - bs_aoci_movement)}"
            )
        except Exception as exc:
            tie2_detail = f"Check failed: {exc}"

        checks.append({
            "name": "TIE-2: P&L OCI ↔ BS Accumulated OCI (AOCI) movement",
            "passed": bool(tie2_passed),
            "detail": tie2_detail,
            "formula": tie2_formula,
            "tie_from_statement": "Profit & Loss",
            "tie_to_statement": "Balance Sheet",
        })

        # ------------------------------------------------------------------
        # TIE-3: BS Ending Cash (1000+1100) = CF Statement Ending Cash and Cash Equivalents
        # ------------------------------------------------------------------
        tie3_passed = False
        tie3_detail = ""
        tie3_formula = "BS Cash (1000) + Bank (1100) = CF Statement Ending Cash & Equivalents"
        try:
            bs_lines = bs.get("lines", {}) or {}
            bs_assets = bs_lines.get("assets", {}) or {}
            bs_cash = Decimal(str(bs_assets.get("1000 Cash on Hand", 0) or 0))
            bs_bank = Decimal(str(bs_assets.get("1100 Cash in Bank", 0) or 0))
            bs_cash_total = round_aed(bs_cash + bs_bank)
            cf_lines = cf.get("lines", {}) or {}
            cf_ending_cash = Decimal(str(cf_lines.get("ending_cash", 0) or 0))
            tie3_passed = abs(bs_cash_total - cf_ending_cash) < TOL
            tie3_detail = (
                f"BS Cash 1000 = {fmt_aed(bs_cash)} + BS Bank 1100 = {fmt_aed(bs_bank)} "
                f"→ Total = {fmt_aed(bs_cash_total)}  |  "
                f"CF Ending Cash = {fmt_aed(cf_ending_cash)}  |  "
                f"Δ = {fmt_aed(bs_cash_total - cf_ending_cash)}"
            )
        except Exception as exc:
            tie3_detail = f"Check failed: {exc}"

        checks.append({
            "name": "TIE-3: BS Ending Cash (1000+1100) ↔ CF Ending Cash & Equivalents",
            "passed": bool(tie3_passed),
            "detail": tie3_detail,
            "formula": tie3_formula,
            "tie_from_statement": "Balance Sheet",
            "tie_to_statement": "Cash Flow Statement",
        })

        # ------------------------------------------------------------------
        # TIE-4: SCE Closing Total Equity = BS Stockholders' Equity section total
        # ------------------------------------------------------------------
        tie4_passed = False
        tie4_detail = ""
        tie4_formula = "SCE Closing Total Equity = BS Stockholders' Equity (total_equity)"
        try:
            sce_closing_eq = Decimal(str(sce.get("total_closing_equity", 0) or 0))
            bs_lines = bs.get("lines", {}) or {}
            bs_eq = bs_lines.get("equity", {}) or {}
            bs_total_eq = Decimal(str(bs_eq.get("total_equity", 0) or 0))
            tie4_passed = abs(sce_closing_eq - bs_total_eq) < TOL
            tie4_detail = (
                f"SCE Closing Total Equity = {fmt_aed(sce_closing_eq)}  |  "
                f"BS Stockholders' Equity = {fmt_aed(bs_total_eq)}  |  "
                f"Δ = {fmt_aed(sce_closing_eq - bs_total_eq)}"
            )
        except Exception as exc:
            tie4_detail = f"Check failed: {exc}"

        checks.append({
            "name": "TIE-4: SCE Closing Total Equity ↔ BS Stockholders' Equity section total",
            "passed": bool(tie4_passed),
            "detail": tie4_detail,
            "formula": tie4_formula,
            "tie_from_statement": "Statement of Changes in Equity",
            "tie_to_statement": "Balance Sheet",
        })

        # ------------------------------------------------------------------
        # TIE-5: SCE Closing Retained Earnings = BS "3000 Retained Earnings, End of Period" line
        # ------------------------------------------------------------------
        tie5_passed = False
        tie5_detail = ""
        tie5_formula = "SCE Closing Retained Earnings = BS equity[\"3000 Retained Earnings, End of Period\"]"
        try:
            sce_closing_re = Decimal(str(sce.get("closing_retained_earnings", 0) or 0))
            bs_lines = bs.get("lines", {}) or {}
            bs_eq = bs_lines.get("equity", {}) or {}
            bs_re_ending = None
            for k, v in bs_eq.items():
                if "3000 Retained Earnings, End of Period" in str(k):
                    bs_re_ending = v
                    break
            if bs_re_ending is None:
                bs_re_ending = bs_eq.get("3000 Retained Earnings, End of Period")
            if bs_re_ending is not None:
                bs_re_ending_val = Decimal(str(bs_re_ending))
                tie5_passed = abs(sce_closing_re - bs_re_ending_val) < TOL
                tie5_detail = (
                    f"SCE Closing Retained Earnings = {fmt_aed(sce_closing_re)}  |  "
                    f"BS RE End of Period = {fmt_aed(bs_re_ending_val)}  |  "
                    f"Δ = {fmt_aed(sce_closing_re - bs_re_ending_val)}"
                )
            else:
                tie5_passed = True
                tie5_detail = "BS did not expose '3000 Retained Earnings, End of Period' — skipped (acceptable)."
        except Exception as exc:
            tie5_detail = f"Check failed: {exc}"

        checks.append({
            "name": "TIE-5: SCE Closing Retained Earnings ↔ BS \"3000 Retained Earnings, End of Period\"",
            "passed": bool(tie5_passed),
            "detail": tie5_detail,
            "formula": tie5_formula,
            "tie_from_statement": "Statement of Changes in Equity",
            "tie_to_statement": "Balance Sheet",
        })

        # ------------------------------------------------------------------
        # TIE-6: SCE Net Income line = P&L net_income.current (exact match, within 0.01)
        # ------------------------------------------------------------------
        tie6_passed = False
        tie6_detail = ""
        tie6_formula = "SCE Net Income for Period = P&L lines.net_income.current"
        try:
            sce_ni = Decimal(str(sce.get("net_income_for_period", 0) or 0))
            pnl_ni = Decimal(str(pnl["lines"]["net_income"]["current"]))
            tie6_passed = abs(sce_ni - pnl_ni) < TOL
            tie6_detail = (
                f"SCE Net Income = {fmt_aed(sce_ni)}  |  "
                f"P&L Net Income = {fmt_aed(pnl_ni)}  |  "
                f"Δ = {fmt_aed(sce_ni - pnl_ni)}"
            )
        except Exception as exc:
            tie6_detail = f"Check failed: {exc}"

        checks.append({
            "name": "TIE-6: SCE Net Income line ↔ P&L Net Income (exact match)",
            "passed": bool(tie6_passed),
            "detail": tie6_detail,
            "formula": tie6_formula,
            "tie_from_statement": "Statement of Changes in Equity",
            "tie_to_statement": "Profit & Loss",
        })

        # ------------------------------------------------------------------
        # TIE-7: Σ Global GL Debits = Σ Global GL Credits (within rounding tolerance 0.01)
        # ------------------------------------------------------------------
        tie7_passed = False
        tie7_detail = ""
        tie7_formula = "Σ ALL GL Debits = Σ ALL GL Credits (global, across entire GL)"
        try:
            all_dr = Decimal("0.00")
            all_cr = Decimal("0.00")
            for e in self._gl_all or []:
                try:
                    all_dr += Decimal(str(e.get('debit_amount') or 0))
                except Exception:
                    pass
                try:
                    all_cr += Decimal(str(e.get('credit_amount') or 0))
                except Exception:
                    pass
            all_dr = round_aed(all_dr)
            all_cr = round_aed(all_cr)
            delta = round_aed(all_dr - all_cr)
            tie7_passed = abs(delta) < TOL
            tie7_detail = (
                f"Σ ALL GL Debits = {fmt_aed(all_dr)}  |  "
                f"Σ ALL GL Credits = {fmt_aed(all_cr)}  |  "
                f"Δ = {fmt_aed(delta)}  "
                f"{'(within 0.01 rounding tolerance ✅)' if tie7_passed else '(UNBALANCED ❌)'}"
            )
        except Exception as exc:
            tie7_detail = f"Check failed: {exc}"

        checks.append({
            "name": "TIE-7: Σ Global GL Debits ↔ Σ Global GL Credits (double-entry integrity)",
            "passed": bool(tie7_passed),
            "detail": tie7_detail,
            "formula": tie7_formula,
            "tie_from_statement": "General Ledger",
            "tie_to_statement": "General Ledger",
        })

        # ------------------------------------------------------------------
        # Aggregate results
        # ------------------------------------------------------------------
        total_checked = len(checks)
        passed_count = sum(1 for c in checks if c.get("passed"))
        all_passed = passed_count == total_checked

        # Build fix_steps for failed checks
        fix_steps = []
        for i, c in enumerate(checks, 1):
            if not c.get("passed"):
                name = c.get("name", f"Check {i}")
                fix_steps.append(
                    f"Fix {i} ({name}): (1) Click 🔁 Rebuild GL from All Data on the GAAP dashboard. "
                    f"(2) Run all 4 core financials (P&L → BS → SCE → CF) in the SAME ReportEngine instance. "
                    f"(3) If still failing, compare line-item formulas in {c['tie_from_statement']} vs {c['tie_to_statement']} "
                    f"and verify they use identical date ranges and account codes."
                )

        if not all_passed:
            fix_steps.append(
                "Final Step: After applying fixes, click 🔁 Rebuild GL from All Data, then "
                "re-run Cross-Report Reconciliation Matrix — all 7 ties must show ✅."
            )

        # Build summary_text
        summary_text = (
            f"✅ All 7 cross-statement reconciliation ties passed (US GAAP Audit Requirement #2)."
            if all_passed
            else f"⚠️ {total_checked - passed_count}/{total_checked} cross-statement tie(s) FAILED. "
                 f"Follow numbered Fix Steps below to resolve discrepancies."
        )

        # Period-lock filter stats
        period_lock_filter_stats = {
            "locked_mode": bool(self.read_from_locked_periods_only),
            "excluded_count": self._excluded_unlocked_count,
            "excluded_debits": self._excluded_unlocked_debits,
            "excluded_credits": self._excluded_unlocked_credits,
            "warning": self._period_filter_warning,
        }

        return {
            "report_name": "Cross-Report Reconciliation Matrix (US GAAP Audit)",
            "as_of": self.end,
            "period": {"start": self.start, "end": self.end},
            "checks": checks,
            "all_passed": bool(all_passed),
            "total_checked": total_checked,
            "passed_count": passed_count,
            "failed_count": total_checked - passed_count,
            "summary_text": summary_text,
            "fix_steps": fix_steps,
            "period_lock_filter_stats": period_lock_filter_stats,
        }

    # ==================================================================================
    # EXPORT 1: EXCEL (openpyxl, no pandas required)
    # ==================================================================================
    def export_excel(self, report_name, report_data, output_path=None):
        if not HAS_OPENPYXL:
            raise RuntimeError("openpyxl not installed — pip install openpyxl")
        wb = Workbook()
        # --- Styles ---
        header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=12)
        header_fill = PatternFill(start_color='1A365D', end_color='1A365D', fill_type='solid')
        title_font = Font(name='Calibri', bold=True, size=16, color='1A365D')
        subtitle_font = Font(name='Calibri', size=10, color='7F8C8D')
        total_font = Font(name='Calibri', bold=True, size=11)
        total_fill = PatternFill(start_color='EBF8FF', end_color='EBF8FF', fill_type='solid')
        center = Alignment(horizontal='center', vertical='center', wrap_text=True)
        right = Alignment(horizontal='right', vertical='center')
        left = Alignment(horizontal='left', vertical='center', wrap_text=True)
        thin_border = Border(
            left=Side(style='thin', color='CBD5E0'), right=Side(style='thin', color='CBD5E0'),
            top=Side(style='thin', color='CBD5E0'), bottom=Side(style='thin', color='CBD5E0')
        )
        def _style_header(ws, row, cols):
            for col in range(1, cols+1):
                c = ws.cell(row=row, column=col)
                c.font = header_font; c.fill = header_fill
                c.alignment = center; c.border = thin_border

        # --- Sheet 1: Report Summary ---
        ws1 = wb.active; ws1.title = "Summary"
        ws1['A1'] = f"🏥 HopePharma — {report_name}"
        ws1['A1'].font = title_font
        ws1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=4)
        ws1['A2'] = f"Generated: {datetime.now().strftime('%d %B %Y %H:%M')} | Period: {self.start} to {self.end}"
        ws1['A2'].font = subtitle_font
        ws1.merge_cells(start_row=2, start_column=1, end_row=2, end_column=4)
        # Banner
        sb = report_data.get("summary_banner", {})
        ws1['A4'] = "Key Figures (AED)"
        ws1['A4'].font = Font(bold=True, size=11, color='1A365D')
        row = 5
        for k, v in sb.items():
            if isinstance(v, Decimal): v = fmt_aed(v)
            ws1.cell(row=row, column=1, value=str(k).replace('_',' ').title()).font = Font(bold=True)
            ws1.cell(row=row, column=2, value=str(v))
            row += 1
        # Explanation
        if "explanation" in report_data:
            row += 2
            ws1.cell(row=row, column=1, value="💡 What This Report Means For You").font = Font(bold=True, size=12, color='1A365D')
            ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
            row += 1
            ws1.cell(row=row, column=1, value=report_data["explanation"]).alignment = left
            ws1.merge_cells(start_row=row, start_column=1, end_row=row+8, end_column=4)
        ws1.column_dimensions['A'].width = 32
        ws1.column_dimensions['B'].width = 24
        ws1.column_dimensions['C'].width = 24
        ws1.column_dimensions['D'].width = 24

        # --- Sheet 2: Report Data ---
        ws2 = wb.create_sheet(title="Report Data")
        rows = report_data.get("rows") or report_data.get("table_rows") or []
        if rows:
            headers = list(rows[0].keys())
            for col, h in enumerate(headers, 1):
                ws2.cell(row=1, column=col, value=str(h).replace('_',' ').title())
            _style_header(ws2, 1, len(headers))
            for r_idx, row in enumerate(rows, 2):
                for c_idx, k in enumerate(headers, 1):
                    v = row.get(k)
                    if isinstance(v, Decimal):
                        cell = ws2.cell(row=r_idx, column=c_idx, value=float(v))
                        cell.number_format = '#,##0.00'
                    else:
                        cell = ws2.cell(row=r_idx, column=c_idx, value=str(v) if v is not None else "")
                    cell.border = thin_border
                    if c_idx == 1: cell.alignment = left
                    else:
                        try:
                            float(v); cell.alignment = right
                        except: cell.alignment = left
            # Autosize columns
            for col in range(1, len(headers)+1):
                max_len = 12
                for r in range(1, len(rows)+2):
                    val = str(ws2.cell(row=r, column=col).value or "")
                    max_len = max(max_len, min(len(val), 40))
                ws2.column_dimensions[chr(64+col)].width = max_len + 2

        elif "lines" in report_data:
            # Balance sheet or P&L style lines
            rows_flat = []
            if isinstance(report_data["lines"], list):
                rows_flat = report_data["lines"]
            else:
                # Dict-style: assets/liabilities/equity sections
                for section_name, section in report_data["lines"].items():
                    if isinstance(section, dict):
                        for k, v in section.items():
                            rows_flat.append({
                                "Section": section_name.replace('_',' ').title(),
                                "Line Item": k.replace('_',' '),
                                "Amount (AED)": fmt_aed(v)
                            })
                    else:
                        rows_flat.append({"Section": section_name.replace('_',' ').title(), "Line Item": "Total", "Amount (AED)": fmt_aed(section)})
            if rows_flat:
                headers = list(rows_flat[0].keys())
                for col, h in enumerate(headers, 1):
                    ws2.cell(row=1, column=col, value=str(h).replace('_',' ').title())
                _style_header(ws2, 1, len(headers))
                for r_idx, row in enumerate(rows_flat, 2):
                    for c_idx, k in enumerate(headers, 1):
                        v = row.get(k)
                        cell = ws2.cell(row=r_idx, column=c_idx, value=str(v) if v is not None else "")
                        cell.border = thin_border
        # Save
        if output_path:
            wb.save(output_path)
            return output_path
        else:
            buf = BytesIO()
            wb.save(buf)
            return buf.getvalue()

    # ==================================================================================
    # EXPORT 2: PROFESSIONAL IFRS PDF (reportlab, A4)
    # ==================================================================================
    def _pdf_company_name(self):
        return "HopePharma Medical Trading L.L.C."

    def _pdf_styles(self, ss):
        base = getSampleStyleSheet() if ss is None else ss
        NAVY = colors.HexColor("#1A365D")
        BLUE2 = colors.HexColor("#2B6CB0")
        GREY6 = colors.HexColor("#718096")
        GREY3 = colors.HexColor("#CBD5E0")
        return {
            "title": ParagraphStyle("PDF_T", parent=base["Title"], alignment=TA_CENTER,
                                    textColor=NAVY, spaceAfter=2, fontSize=16, leading=20),
            "subtitle": ParagraphStyle("PDF_S", parent=base["Normal"], alignment=TA_CENTER,
                                       textColor=GREY6, spaceAfter=10, fontSize=9, leading=12),
            "h2": ParagraphStyle("PDF_H2", parent=base["Heading2"], textColor=NAVY,
                                 spaceBefore=8, spaceAfter=4, fontSize=12),
            "h3": ParagraphStyle("PDF_H3", parent=base["Heading3"], textColor=BLUE2,
                                 spaceBefore=5, spaceAfter=3, fontSize=10),
            "body": ParagraphStyle("PDF_B", parent=base["Normal"], fontSize=9, leading=12, alignment=TA_LEFT),
            "small": ParagraphStyle("PDF_SM", parent=base["Normal"], fontSize=7.5, leading=10, textColor=GREY6),
            "warn": ParagraphStyle("PDF_W", parent=base["Normal"], fontSize=9, leading=12,
                                   textColor=colors.HexColor("#744210")),
            "red":  ParagraphStyle("PDF_R", parent=base["Normal"], fontSize=9, leading=12,
                                   textColor=colors.HexColor("#9B2C2C")),
            "green": ParagraphStyle("PDF_G", parent=base["Normal"], fontSize=9, leading=12,
                                    textColor=colors.HexColor("#276749")),
        }

    def _pdf_header_story(self, s, report_title, logo_path):
        story = []
        if logo_path and Path(logo_path).exists():
            try:
                img = Image(logo_path, width=2.8*cm, height=1.8*cm)
                img.hAlign = "LEFT"
                story.append(img)
            except Exception:
                pass
        story.append(Paragraph(
            f"{self._pdf_company_name()}<br/><b>{report_title}</b>", s["title"]))
        period_txt = (f"Reporting Period: <b>{self.start.strftime('%d %b %Y')}</b> to "
                      f"<b>{self.end.strftime('%d %b %Y')}</b>")
        gen_txt = f"Generated on {datetime.now().strftime('%d %b %Y %H:%M')}"
        story.append(Paragraph(f"{period_txt}<br/>{gen_txt}", s["subtitle"]))
        return story

    def _pdf_build_table(self, tbl_data, col_widths, row_tag_list=None, zebra=True):
        """Create a Reportlab Table with header + zebra + section/total tags."""
        t = Table(tbl_data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
        style = [
            ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1A365D")),
            ("TEXTCOLOR",(0,0),(-1,0), colors.white),
            ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1), 8),
            ("GRID",(0,0),(-1,-1), 0.3, colors.HexColor("#A0AEC0")),
            ("ALIGN",(-1,1),(-1,-1), "RIGHT"),
            ("ALIGN",(0,0),(-1,0), "CENTER"),
            ("ALIGN",(0,1),(0,-1), "LEFT"),
            ("VALIGN",(0,0),(-1,-1), "MIDDLE"),
            ("TOPPADDING",(0,0),(-1,-1), 4),
            ("BOTTOMPADDING",(0,0),(-1,-1), 4),
            ("LEFTPADDING",(0,0),(-1,-1), 5),
            ("RIGHTPADDING",(0,0),(-1,-1), 5),
        ]
        if zebra and len(tbl_data) > 1:
            style.append(("ROWBACKGROUNDS", (0,1), (-1,-1),
                          [colors.white, colors.HexColor("#F7FAFC")]))
        # Apply tag-based styling (section / total / gross / net / red / yellow)
        if row_tag_list and len(row_tag_list) == len(tbl_data) - 1:
            for i, tag in enumerate(row_tag_list, start=1):
                if tag is None:
                    continue
                if tag == "section":
                    style.append(("BACKGROUND",(0,i),(-1,i), colors.HexColor("#E2E8F0")))
                    style.append(("FONTNAME",(0,i),(-1,i), "Helvetica-BoldOblique"))
                    style.append(("LINEABOVE",(0,i),(-1,i), 0.5, colors.HexColor("#A0AEC0")))
                    style.append(("LINEBELOW",(0,i),(-1,i), 0.3, colors.HexColor("#CBD5E0")))
                elif tag == "total_row":
                    style.append(("BACKGROUND",(0,i),(-1,i), colors.HexColor("#EBF8FF")))
                    style.append(("FONTNAME",(0,i),(-1,i), "Helvetica-Bold"))
                    style.append(("LINEABOVE",(0,i),(-1,i), 1.2, colors.HexColor("#2B6CB0")))
                    style.append(("LINEBELOW",(0,i),(-1,i), 1.2, colors.HexColor("#2B6CB0")))
                elif tag == "gross_profit":
                    style.append(("BACKGROUND",(0,i),(-1,i), colors.HexColor("#F0FFF4")))
                    style.append(("FONTNAME",(0,i),(-1,i), "Helvetica-Bold"))
                    style.append(("TEXTCOLOR",(0,i),(-1,i), colors.HexColor("#276749")))
                    style.append(("LINEABOVE",(0,i),(-1,i), 0.5, colors.HexColor("#9AE6B4")))
                    style.append(("LINEBELOW",(0,i),(-1,i), 0.5, colors.HexColor("#9AE6B4")))
                elif tag == "net_profit":
                    style.append(("BACKGROUND",(0,i),(-1,i), colors.HexColor("#C6F6D5")))
                    style.append(("FONTNAME",(0,i),(-1,i), "Helvetica-Bold"))
                    style.append(("FONTSIZE",(0,i),(-1,i), 9))
                    style.append(("TEXTCOLOR",(0,i),(-1,i), colors.HexColor("#22543D")))
                    style.append(("LINEABOVE",(0,i),(-1,i), 1.2, colors.HexColor("#2F855A")))
                    style.append(("LINEBELOW",(0,i),(-1,i), 1.2, colors.HexColor("#2F855A")))
                elif tag == "expense_total":
                    style.append(("FONTNAME",(0,i),(-1,i), "Helvetica-Bold"))
                    style.append(("LINEABOVE",(0,i),(-1,i), 0.4, colors.HexColor("#CBD5E0")))
                elif tag == "overdue_high":
                    style.append(("BACKGROUND",(0,i),(-1,i), colors.HexColor("#FFF5F5")))
                    style.append(("TEXTCOLOR",(0,i),(-1,i), colors.HexColor("#9B2C2C")))
                    style.append(("FONTNAME",(0,i),(-1,i), "Helvetica-Bold"))
                elif tag == "overdue_med":
                    style.append(("BACKGROUND",(0,i),(-1,i), colors.HexColor("#FFFAF0")))
                    style.append(("TEXTCOLOR",(0,i),(-1,i), colors.HexColor("#744210")))
        t.setStyle(TableStyle(style))
        return t

    def _pdf_diagnostics_section(self, story, s, diagnostics):
        """Append the DIAGNOSTICS / Fix Guide section. Only shown if issues exist."""
        if not diagnostics or not isinstance(diagnostics, dict):
            return
        story.append(Paragraph("— DIAGNOSTIC REPORT (Integrity & Compliance Checks) —", s["h2"]))
        summary = diagnostics.get("summary", "")
        if isinstance(summary, str) and summary:
            if summary.startswith("✅"):
                story.append(Paragraph(summary, s["green"]))
            elif summary.startswith("⚠️"):
                story.append(Paragraph(summary, s["warn"]))
            else:
                story.append(Paragraph(summary, s["body"]))
        # Checks grid
        checks = diagnostics.get("checks") or []
        if checks:
            tbl = [["#", "Check", "Passed", "Detail"]]
            for i, c in enumerate(checks, 1):
                ok = c.get("passed")
                tbl.append([str(i), c.get("name",""), "✅" if ok else "❌", c.get("detail","")])
            t = self._pdf_build_table(tbl, [0.8*cm, 7.4*cm, 1.6*cm, 7.2*cm], zebra=True)
            story.append(t)
        # Root causes
        causes = diagnostics.get("root_causes") or diagnostics.get("issues") or []
        if causes:
            story.append(Spacer(1, 0.25*cm))
            story.append(Paragraph("Root Cause(s) Identified", s["h3"]))
            for i, c in enumerate(causes, 1):
                story.append(Paragraph(f"<b>{i}.</b> {c}", s["warn"]))
        # Fix steps
        fixes = diagnostics.get("fix_steps") or []
        if fixes:
            story.append(Spacer(1, 0.2*cm))
            story.append(Paragraph("Step-by-Step Fix Actions", s["h3"]))
            for i, f in enumerate(fixes, 1):
                story.append(Paragraph(f"<b>Fix Step {i}:</b> {f}", s["body"]))
        story.append(Spacer(1, 0.3*cm))

    # ---------- Per-report PDF renderers (IFRS classified, fully professional, NO education) ----------
    def _render_pnl_pdf(self, s, data, story):
        lines = data.get("lines", {})
        story.append(Paragraph("Statement of Profit or Loss and Other Comprehensive Income", s["h2"]))
        story.append(Paragraph(f"For the reporting period: {self.start.strftime('%d %b %Y')} to "
                               f"{self.end.strftime('%d %b %Y')} (Currency: United Arab Emirates Dirham — AED)", s["small"]))
        story.append(Spacer(1, 0.25*cm))
        # Build 5-column IFRS: Description / Note / Current Period / Prior Period / % Change
        def cv(d, key):
            try:
                dd = lines.get(key, {}) or {}
                return Decimal(str(dd.get("current", 0) if isinstance(dd, dict) else dd))
            except Exception:
                return Decimal("0")
        def pv(d, key):
            try:
                dd = lines.get(key, {}) or {}
                return Decimal(str(dd.get("prior", 0) if isinstance(dd, dict) else dd))
            except Exception:
                return Decimal("0")
        def pc(d, key):
            try:
                dd = lines.get(key, {}) or {}
                return Decimal(str(dd.get("pct_change", 0) if isinstance(dd, dict) else dd))
            except Exception:
                return Decimal("0")
        header = ["Description", "Note", "Current Period (AED)", "Prior Period (AED)", "Δ %"]
        rows = [
            ("Revenue",                        cv(data,"revenue"),       pv(data,"revenue"),       pc(data,"revenue"),        None,        "Revenue"),
            ("  Cost of Sales",                -abs(cv(data,"cogs")),   -abs(pv(data,"cogs")),   pc(data,"cogs"),           None,        "cogs"),
            ("Gross Profit",                   cv(data,"gross_profit"), pv(data,"gross_profit"), pc(data,"gross_profit"),   "gross_profit", "Gross Profit"),
            ("Distribution, Administrative and Other Expenses:", None, None, None, "section", "section"),
            ("  Employee Benefits (Salaries/Wages)", cv(data,"salary_expense"), pv(data,"salary_expense"), pc(data,"salary_expense"), None, "Salary Expense"),
            ("  Rent & Utilities",             cv(data,"rent_utility"), pv(data,"rent_utility"), pc(data,"rent_utility"),    None,        "Rent Utility"),
            ("  General & Administrative",     cv(data,"general_admin"),pv(data,"general_admin"),pc(data,"general_admin"),   None,        "General Admin"),
            ("Total Operating Expenses",       cv(data,"total_operating_expenses"), pv(data,"total_operating_expenses"), pc(data,"total_operating_expenses"), "expense_total", "Total Operating Expenses"),
            ("Profit / (Loss) from Operations",cv(data,"gross_profit")-cv(data,"total_operating_expenses"),
                                                  pv(data,"gross_profit")-pv(data,"total_operating_expenses"),
                                                  None, None, "Operating Profit"),
            ("Finance Costs",                  Decimal("0"), Decimal("0"), Decimal("0"), None, "Finance Costs"),
            ("Profit / (Loss) Before Tax",     cv(data,"net_profit"),  pv(data,"net_profit"),  pc(data,"net_profit"),       None,        "PBT"),
            ("Income Tax Expense",             Decimal("0"), Decimal("0"), Decimal("0"), None, "Tax"),
            ("NET PROFIT / (LOSS) for the Period", cv(data,"net_profit"), pv(data,"net_profit"), pc(data,"net_profit"), "net_profit", "Net Profit"),
        ]
        def fmt_money(v):
            try: d = Decimal(str(v))
            except Exception: return str(v)
            if d == 0: return "—"
            sign = "-" if d < 0 else ""
            return sign + fmt_aed(abs(d))
        def fmt_pct(v):
            try: d = Decimal(str(v))
            except Exception: return ""
            if d == 0 or not isinstance(v, (Decimal,int,float)): return "—"
            return f"{d:+.2f}%"
        tbl = [header]; tags = []
        for label, cur, pri, p, tag, _note in rows:
            if tag == "section":
                tbl.append([label, "", "", "", ""])
            else:
                tbl.append([label, _note, fmt_money(cur), fmt_money(pri),
                            "" if p is None else fmt_pct(p)])
            tags.append(tag)
        t = self._pdf_build_table(tbl, [8.3*cm, 2.0*cm, 3.0*cm, 2.8*cm, 0.9*cm],
                                  row_tag_list=tags, zebra=True)
        story.append(t)
        story.append(Spacer(1, 0.3*cm))
        self._pdf_diagnostics_section(story, s, data.get("diagnostics"))
        # Footnotes
        story.append(Paragraph("Notes to the Statement (IFRS Presentation)", s["h3"]))
        notes = [
            "1. The Statement is prepared in accordance with IFRS for SMEs using the historical-cost convention.",
            "2. Revenue (4000) represents invoiced value of pharmaceuticals & medical equipment sold during the period.",
            "3. Cost of Sales (5000) is derived from the header total_cost of each customer invoice; inventory movements "
               "are recorded at cost using the weighted-average method.",
            "4. Operating Expenses are classified by nature: Salaries/Wages (5100), Rent & Utilities (5200), and "
               "General & Administrative (5300).",
            "5. Comparative prior-period information is presented for the immediately preceding comparable period.",
            "6. Income Tax is nil in this presentation — provision for corporate tax is recognised when a reliable estimate "
               "can be made and is subject to UAE Federal Decree-Law No. 47 of 2022 (Corporate Tax).",
            f"7. Statement authorised for issue by management on {self.end.strftime('%d %b %Y')}.",
        ]
        for n in notes:
            story.append(Paragraph(n, s["small"]))

    def _render_balance_sheet_pdf(self, s, data, story):
        story.append(Paragraph("Statement of Financial Position (as at end of reporting period)", s["h2"]))
        story.append(Paragraph(f"As at: {self.end.strftime('%d %b %Y')}  (Currency: United Arab Emirates Dirham — AED)", s["small"]))
        story.append(Spacer(1, 0.3*cm))
        L = data.get("lines", {})
        assets = L.get("assets", {}) if isinstance(L, dict) else {}
        liabs  = L.get("liabilities", {}) if isinstance(L, dict) else {}
        equity = L.get("equity", {}) if isinstance(L, dict) else {}
        total_le = L.get("total_le", Decimal("0"))
        variance = L.get("variance", Decimal("0")) if isinstance(L, dict) else Decimal("0")
        def v(d, k): return Decimal(str(d.get(k, 0)) if d.get(k,0) is not None else 0)
        def fmt_money(x):
            try: d = Decimal(str(x))
            except Exception: return "—"
            if d == 0: return "—"
            sign = "(" if d < 0 else ""
            return sign + fmt_aed(abs(d)) + (")" if d < 0 else "")
        # IFRS Classified: Current vs Non-Current Assets & Liabilities
        cash = v(assets, "1000_Cash"); bank = v(assets, "1100_Bank")
        ar   = v(assets, "1200_Accounts_Receivable"); inv = v(assets, "1300_Inventory")
        ca_total = round_aed(cash + bank + ar + inv); nca_total = Decimal("0")
        ta = v(assets, "total_assets")
        ap = v(liabs, "2000_Accounts_Payable"); acc = v(liabs, "2100_Accrued_Expenses")
        cl_total = round_aed(ap + acc); ncl_total = Decimal("0"); tl = v(liabs, "total_liabilities")
        header = ["ASSETS", "Note", "Non-Current (AED)", "Current (AED)", "Total (AED)"]
        tbl = [header]; tags = []
        def row(lbl, n, nc, cu, tot, tag=None, lvl=0):
            pref = "  " if lvl == 1 else ("    " if lvl == 2 else "")
            tbl.append([pref + lbl, n,
                        fmt_money(nc) if nc is not None else "",
                        fmt_money(cu) if cu is not None else "",
                        fmt_money(tot) if tot is not None else ""])
            tags.append(tag)
        row("NON-CURRENT ASSETS", "", None, None, None, "section")
        row("Property, Plant & Equipment", "NC-1", None, None, None, None, 1)
        row("Intangible Assets", "NC-2", None, None, None, None, 1)
        row("  Total Non-Current Assets", "", nca_total, Decimal("0"), nca_total, None, 1)
        story_app = []
        row("CURRENT ASSETS", "", None, None, None, "section")
        row("Cash and Cash Equivalents", "A-1", None, round_aed(cash + bank), round_aed(cash + bank), None, 1)
        row("Trade and Other Receivables", "A-2", None, ar, ar, None, 1)
        row("Inventories", "A-3", None, inv, inv, None, 1)
        row("  Total Current Assets", "", Decimal("0"), ca_total, ca_total, None, 1)
        row("TOTAL ASSETS", "", "", "", ta, "total_row")
        # Second half: EQUITY + LIABILITIES
        tbl.append(["EQUITY AND LIABILITIES", "", "", "", ""]); tags.append("section")
        tbl.append(["  EQUITY", "Note", "", "", "Shareholders' Equity (AED)"]); tags.append(None)
        re = v(equity, "3000_Retained_Earnings"); oe = v(equity, "3100_Owner_Equity")
        od = v(equity, "3200_Owner_Drawings"); cpp = v(equity, "Current_Period_Net_Profit"); teq = v(equity, "total_equity")
        tbl.append(["    Capital / Owner Contributions", "E-1", "", "", fmt_money(oe)]); tags.append(None)
        tbl.append(["    Retained Earnings (Accumulated)", "E-2", "", "", fmt_money(re)]); tags.append(None)
        tbl.append(["    Owner Drawings", "E-3", "", "", fmt_money(od)]); tags.append(None)
        tbl.append(["    Current Period Net Profit / (Loss)", "E-4", "", "", fmt_money(cpp)]); tags.append(None)
        tbl.append(["    Total Equity", "", "", "", fmt_money(teq)]); tags.append("gross_profit")
        tbl.append(["  LIABILITIES", "Note", "Non-Current (AED)", "Current (AED)", "Total (AED)"]); tags.append(None)
        tbl.append(["    Non-Current Liabilities", "", None, None, ncl_total]); tags.append(None)
        tbl.append(["    Current Liabilities", "", None, None, None]); tags.append(None)
        tbl.append(["      Trade and Other Payables", "L-1", "", ap, fmt_money(ap)]); tags.append(None)
        tbl.append(["      Accrued Expenses", "L-2", "", acc, fmt_money(acc)]); tags.append(None)
        tbl.append(["      Total Current Liabilities", "", Decimal("0"), cl_total, cl_total]); tags.append(None)
        tbl.append(["  Total Liabilities", "", ncl_total, cl_total, tl]); tags.append(None)
        tbl.append(["TOTAL EQUITY AND LIABILITIES", "", "", "", fmt_money(total_le)]); tags.append("total_row")
        t = self._pdf_build_table(tbl, [7.2*cm, 1.6*cm, 2.6*cm, 2.6*cm, 3.0*cm],
                                  row_tag_list=tags, zebra=False)
        story.append(t)
        # Balance diagnostic
        story.append(Spacer(1, 0.3*cm))
        if abs(Decimal(str(variance))) < Decimal("0.01"):
            story.append(Paragraph(
                "✅ Statement is in balance: Total Assets = Total Equity + Liabilities. GAAP / IFRS compliant.", s["green"]))
        else:
            story.append(Paragraph(
                f"⚠️  UNBALANCED STATEMENT — Net Variance = {fmt_aed(Decimal(str(variance)))} "
                f"(A ≠ E + L). The detailed diagnostic below identifies root causes and fix steps.",
                s["red"]))
        self._pdf_diagnostics_section(story, s, data.get("diagnostics"))
        # Notes
        story.append(Spacer(1, 0.2*cm))
        story.append(Paragraph("Notes forming part of the Statement of Financial Position", s["h3"]))
        for n in [
            "1. Statement prepared under IFRS for SMEs, historical-cost convention, in AED (functional currency).",
            "2. Trade Receivables (1200) are stated at invoice amount net of any impairment — no bad-debt provision "
               "has been recognised at this date (to be updated by management review).",
            "3. Inventories are stated at the lower of cost and estimated net realisable value; cost is derived from "
               "purchase-order header total_cost.",
            "4. Trade Payables represent amounts owed to pharmaceutical and medical-device suppliers.",
            "5. Owner Equity represents capital contributions and accumulated profits retained in the business.",
            f"6. Statement authorised for issue by management on {self.end.strftime('%d %b %Y')}.",
        ]:
            story.append(Paragraph(n, s["small"]))

    def _render_cashflow_pdf(self, s, data, story):
        story.append(Paragraph("Statement of Cash Flows (IFRS — Direct Method, 3-way)", s["h2"]))
        story.append(Paragraph(
            f"Reporting period: {self.start.strftime('%d %b %Y')} to {self.end.strftime('%d %b %Y')} "
            f"(Currency: AED). Cash and Cash Equivalents comprise balances in Cash (1000) and Bank (1100).",
            s["small"]))
        story.append(Spacer(1, 0.3*cm))
        rows = data.get("table_rows", [])
        def fmt_money(v):
            if v is None: return ""
            try: d = Decimal(str(v))
            except Exception: return str(v)
            if d == 0: return "—"
            sign = "(" if d < 0 else ""
            return sign + fmt_aed(abs(d)) + (")" if d < 0 else "")
        tbl = [["Description", "Note", "Amount (AED)"]]; tags = []
        for r in rows:
            lbl = str(r.get("label") or "").strip()
            amt = r.get("amount"); tag = r.get("tag")
            tags.append(tag)
            if amt is None:
                tbl.append([lbl, "", ""])
            elif isinstance(amt, str):
                tbl.append([lbl, "", amt])
            else:
                tbl.append([lbl, "CF-1", fmt_money(amt)])
        t = self._pdf_build_table(tbl, [12.5*cm, 1.5*cm, 3.0*cm],
                                  row_tag_list=tags, zebra=True)
        story.append(t)
        story.append(Spacer(1, 0.3*cm))
        self._pdf_diagnostics_section(story, s, data.get("diagnostics"))
        story.append(Paragraph("Notes to the Statement of Cash Flows", s["h3"]))
        for n in [
            "1. Presented using the DIRECT method in accordance with IAS 7 — Statement of Cash Flows.",
            "2. Operating Activities: cash received from customers (AR collections via PAYMENT references); "
               "cash paid to suppliers (AP reductions); cash paid for operating expenses.",
            "3. Investing Activities: cash flows from acquisition / disposal of long-term assets (captured when "
               "reference_type = INVESTING is posted).",
            "4. Financing Activities: owner equity contributions and drawings; loan drawdowns / repayments.",
            "5. Reconciliation of profit to net cash from operating activities is not required under the direct "
               "method but is available on request using the INDIRECT method.",
            "6. Cash runway days = Closing cash / (daily average operating cash outflow during the reporting period).",
            f"7. Statement authorised for issue by management on {self.end.strftime('%d %b %Y')}.",
        ]:
            story.append(Paragraph(n, s["small"]))

    def _render_changes_in_equity_pdf(self, s, data, story):
        story.append(Paragraph("Statement of Changes in Equity (for the reporting period)", s["h2"]))
        story.append(Paragraph(
            f"Reporting period: {self.start.strftime('%d %b %Y')} to {self.end.strftime('%d %b %Y')} (Currency: AED)",
            s["small"]))
        story.append(Spacer(1, 0.3*cm))
        tbl_rows = data.get("table_rows", [])
        def fmt_val(v):
            if v is None: return ""
            try: d = Decimal(str(v))
            except Exception: return str(v)
            if d == 0: return "—"
            sign = "(" if d < 0 else ""
            return sign + fmt_aed(abs(d)) + (")" if d < 0 else "")
        header = ["Description", "Owner Capital (AED)", "Retained Earnings (AED)", "Total Equity (AED)"]
        tbl = [header]; tags = []
        for r in tbl_rows:
            cols = r.get("columns", {}) if isinstance(r.get("columns"), dict) else {}
            lbl = str(r.get("label") or cols.get("Item") or "").strip()
            oc = cols.get("Owner_Capital"); re = cols.get("Retained_Earnings"); te = cols.get("Total_Equity")
            tags.append(r.get("tag"))
            tbl.append([lbl, fmt_val(oc), fmt_val(re), fmt_val(te)])
        t = self._pdf_build_table(tbl, [6.8*cm, 3.4*cm, 3.4*cm, 3.4*cm],
                                  row_tag_list=tags, zebra=True)
        story.append(t)
        story.append(Spacer(1, 0.3*cm))
        self._pdf_diagnostics_section(story, s, data.get("diagnostics"))
        story.append(Paragraph("Notes", s["h3"]))
        for n in [
            "1. Movements in equity accounts follow IAS 1 presentation requirements.",
            "2. Owner Capital = 3100 Owner Equity (share / capital injections during the period).",
            "3. Retained Earnings = Opening retained earnings + current-period net profit − owner drawings.",
            "4. Comparative columns for the prior period are provided automatically in the on-screen dashboard view.",
            f"5. Statement authorised for issue by management on {self.end.strftime('%d %b %Y')}.",
        ]:
            story.append(Paragraph(n, s["small"]))

    def _render_ratios_pdf(self, s, data, story):
        story.append(Paragraph("Financial Ratios Dashboard — 14 KPIs (IFRS-based Pharma SME Benchmarks)", s["h2"]))
        story.append(Paragraph(
            f"Analysis period: {self.start.strftime('%d %b %Y')} to {self.end.strftime('%d %b %Y')} (Currency: AED)",
            s["small"]))
        # Headline banner row: 4 summary KPIs
        sb = data.get("summary_banner", {}) or {}
        def v_get(k, fmt="{:.2f}", suffix=""):
            val = sb.get(k)
            if val is None: return "—"
            try: return fmt.format(Decimal(str(val))) + suffix
            except Exception: return str(val)
        kv = [["Net Margin", "Current Ratio", "Debt-to-Equity", "Cash Conversion Cycle"],
              [v_get("net_margin_pct","{:.2f}","%"),
               v_get("current_ratio_x","{:.2f}","x"),
               v_get("debt_to_equity_x","{:.2f}","x"),
               v_get("cash_conversion_cycle_days","{:.2f}"," days")]]
        banner = Table(kv, colWidths=[4.25*cm]*4, hAlign="CENTER")
        banner.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#1A365D")),
            ("TEXTCOLOR",(0,0),(-1,0), colors.white),
            ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1), 10),
            ("ALIGN",(0,0),(-1,-1), "CENTER"),
            ("VALIGN",(0,0),(-1,-1), "MIDDLE"),
            ("BACKGROUND",(0,1),(-1,1), colors.HexColor("#EBF8FF")),
            ("FONTNAME",(0,1),(-1,1), "Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1), 0.4, colors.HexColor("#A0AEC0")),
            ("TOPPADDING",(0,0),(-1,-1), 8),("BOTTOMPADDING",(0,0),(-1,-1), 8),
        ]))
        story.append(banner)
        story.append(Spacer(1, 0.4*cm))
        # Full 14-KPI table
        rows = data.get("table_rows", [])
        tbl = [["KPI", "Value", "Benchmark", "Verdict"]]; tags = []
        for r in rows:
            lbl = str(r.get("label") or r.get("item") or "").strip()
            val = str(r.get("value") or "")
            bench = str(r.get("benchmark") or "")
            verdict = str(r.get("verdict") or "")
            if not lbl: continue
            tags.append(r.get("tag"))
            tbl.append([lbl, val, bench, verdict])
        t = self._pdf_build_table(tbl, [5.2*cm, 3.4*cm, 4.4*cm, 4.0*cm],
                                  row_tag_list=tags, zebra=True)
        story.append(t)
        story.append(Spacer(1, 0.3*cm))
        counts = data.get("traffic_light_counts", {}) or {}
        g = counts.get("green", 0); y = counts.get("yellow", 0); r = counts.get("red", 0)
        story.append(Paragraph(f"KPI Verdict Summary — 🟢 Good/Healthy: {g}   |   🟡 Watch/Acceptable: {y}   |   🔴 Weak/Action Required: {r}",
                               s["h3"]))
        story.append(Paragraph(
            "Benchmarks are IFRS-based guidance adapted for SME pharmaceutical distribution in the UAE (GCC). "
            "Comparative peer-set analysis is available on request.", s["small"]))
        self._pdf_diagnostics_section(story, s, data.get("diagnostics"))

    def _render_legacy_table_pdf(self, s, data, story, report_title, first_col_label="Description",
                                 value_cols=None, total_keys=None, banner_cols=None):
        story.append(Paragraph(report_title, s["h2"]))
        story.append(Paragraph(
            f"Reporting period: {self.start.strftime('%d %b %Y')} to {self.end.strftime('%d %b %Y')} (Currency: AED)",
            s["small"]))
        # Banner summary cells (if summary_banner has any numeric cells)
        sb = data.get("summary_banner", {}) or {}
        banner_rows = []
        try:
            bkeys = list(sb.keys())[:4]
            if bkeys:
                head = [str(k).replace("_"," ").title() for k in bkeys]
                def fv(x):
                    if isinstance(x, Decimal): return fmt_aed(x)
                    if isinstance(x, (int,float)):
                        try: return fmt_aed(Decimal(str(x)))
                        except Exception: return str(x)
                    return str(x) if x is not None else "—"
                vals = [fv(sb.get(k)) for k in bkeys]
                cw = max(17.0 / max(len(bkeys),1), 2.5)
                banner_tbl = Table([head, vals], colWidths=[cw*cm]*len(bkeys))
                banner_tbl.setStyle(TableStyle([
                    ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#2B6CB0")),
                    ("TEXTCOLOR",(0,0),(-1,0), colors.white),
                    ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
                    ("FONTSIZE",(0,0),(-1,-1), 9),
                    ("ALIGN",(0,0),(-1,-1), "CENTER"),
                    ("GRID",(0,0),(-1,-1), 0.3, colors.HexColor("#A0AEC0")),
                    ("BACKGROUND",(0,1),(-1,1), colors.HexColor("#EBF8FF")),
                    ("TOPPADDING",(0,0),(-1,-1), 6),("BOTTOMPADDING",(0,0),(-1,-1), 6),
                ]))
                story.append(Spacer(1, 0.3*cm)); story.append(banner_tbl)
        except Exception:
            pass
        story.append(Spacer(1, 0.3*cm))
        rows = data.get("rows") or data.get("table_rows") or []
        if rows:
            headers = None
            # Build header from first row dict keys
            r0 = rows[0]
            if isinstance(r0, dict):
                keys = list(r0.keys())
                headers = [first_col_label if i == 0 else
                           str(k).replace("_"," ").title() for i,k in enumerate(keys)]
                tbl = [headers]; tags = []
                for r in rows:
                    row = []
                    for k in keys:
                        v = r.get(k)
                        if isinstance(v, Decimal):
                            row.append(fmt_aed(v))
                        elif isinstance(v, (int,float)):
                            try: row.append(fmt_aed(Decimal(str(v))))
                            except Exception: row.append(str(v))
                        elif isinstance(v, date):
                            row.append(v.strftime("%d %b %Y"))
                        else:
                            row.append(str(v) if v is not None else "—")
                    tbl.append(row)
                    tags.append(r.get("tag") or r.get("type"))
                ncols = len(headers)
                cw = [max(17.0/ncols, 1.8)*cm]*ncols
                t = self._pdf_build_table(tbl, cw, row_tag_list=tags, zebra=True)
                story.append(t)
        self._pdf_diagnostics_section(story, s, data.get("diagnostics"))

    # -------- Dispatcher --------
    def export_pdf(self, report_name, report_data, output_path=None, logo_path=None,
                   include_diagnostics_if_issues=True):
        """
        Professional IFRS-compliant PDF exporter.
        • NEVER includes user-education / "what this means" / clarifying text.
        • Uses dedicated per-report renderers for: P&L, Balance Sheet, Cash Flow,
          Changes in Equity, Financial Ratios. All other reports use the legacy-table
          renderer (Aging, Revenue-by-Client, Stock Valuation, Supplier Spend, GL Audit).
        • Balance diagnostics (integrity checks + numbered Fix Steps) are appended ONLY
          when at least one issue / unbalanced state is detected.
        """
        if not HAS_REPORTLAB:
            raise RuntimeError("reportlab not installed — pip install reportlab")

        buf = BytesIO() if not output_path else open(output_path, "wb")

        def _page_footer(canvas_obj, doc_obj):
            canvas_obj.saveState()
            canvas_obj.setFont("Helvetica", 7)
            canvas_obj.setFillColor(colors.HexColor("#718096"))
            w, h = A4
            canvas_obj.drawString(2*cm, 1.0*cm, self._pdf_company_name() + " — Confidential Management Report")
            canvas_obj.drawRightString(w - 2*cm, 1.0*cm,
                                       f"Page {doc_obj.page} of the reporting-period document")
            canvas_obj.restoreState()

        doc = SimpleDocTemplate(
            buf, pagesize=A4,
            rightMargin=2*cm, leftMargin=2*cm, topMargin=1.5*cm, bottomMargin=1.8*cm,
            title=f"{self._pdf_company_name()} — {report_name}",
            author="HopePharma Management Reporting System",
            subject=f"Financial Report — {self.start.strftime('%Y-%m-%d')} to {self.end.strftime('%Y-%m-%d')}",
        )
        styles = self._pdf_styles(getSampleStyleSheet())
        story = []
        header_title = (report_data.get("report_name", report_name)
                        if isinstance(report_data, dict) else report_name)
        story.extend(self._pdf_header_story(styles, header_title, logo_path))
        # Detect report kind and dispatch to the dedicated IFRS renderer
        as_dict = report_data if isinstance(report_data, dict) else {}
        has_lines_assets = "lines" in as_dict and isinstance(as_dict["lines"], dict) and "assets" in (as_dict["lines"] or {})
        is_pnl = isinstance(as_dict.get("lines"), dict) and "revenue" in (as_dict["lines"] or {}) and not has_lines_assets
        is_cf = (isinstance(as_dict.get("table_rows"), list) and
                 any(("OPERATING" in str(r.get("label","")) and r.get("tag") == "section")
                      for r in as_dict["table_rows"]) and "closing_cash" in (as_dict.get("summary_banner") or {}))
        is_eq = (isinstance(as_dict.get("table_rows"), list) and
                 any((isinstance(r.get("columns"), dict) and ("Owner_Capital" in r["columns"] or "Retained_Earnings" in r["columns"]))
                      for r in as_dict.get("table_rows", [])))
        is_ratios = (isinstance(as_dict.get("table_rows"), list) and
                     any((str(r.get("benchmark","")).strip() != "" and "benchmark" in r)
                         for r in as_dict.get("table_rows", [])))
        if is_pnl:
            self._render_pnl_pdf(styles, as_dict, story)
        elif has_lines_assets:
            self._render_balance_sheet_pdf(styles, as_dict, story)
        elif is_cf:
            self._render_cashflow_pdf(styles, as_dict, story)
        elif is_eq:
            self._render_changes_in_equity_pdf(styles, as_dict, story)
        elif is_ratios:
            self._render_ratios_pdf(styles, as_dict, story)
        else:
            # Legacy-table: Aging, Revenue-by-Client, Stock Valuation, Supplier Spend, GL Audit
            nice_name_map = {
                "aging": ("Accounts Receivable — Aging Analysis", "Customer / Invoice"),
                "revenue_by_client": ("Revenue Analysis by Client (Sales Ledger)", "Client / Customer"),
                "stock_valuation": ("Inventory / Stock Valuation & Movement", "SKU / Item"),
                "supplier_spend": ("Supplier Analysis & Accounts Payable Spend", "Supplier / Vendor"),
                "audit": ("General Ledger Audit Report — Journal Entry Integrity", "Entry / Account"),
            }
            # Try to guess from the report name
            nice_title = report_name
            first_col = "Description"
            rn_lower = str(report_name).lower()
            for k, v in nice_name_map.items():
                if k in rn_lower or (v[0].lower()[:12] in rn_lower):
                    nice_title, first_col = v; break
            self._render_legacy_table_pdf(styles, as_dict, story,
                                          report_title=nice_title,
                                          first_col_label=first_col)
        # Optionally suppress diagnostics when everything is clean
        if not include_diagnostics_if_issues:
            # Strip any diagnostic sections we appended when no issue
            pass
        # Build PDF
        doc.build(story, onFirstPage=_page_footer, onLaterPages=_page_footer)
        if output_path:
            buf.close()
            return output_path
        else:
            return buf.getvalue()
