"""
cost_center_ui.py
-----------------
Two standalone Tkinter dialogs that live in the InvoiceManager main dashboard
(user asked Cost Centers to NOT live inside Create/Edit Invoice anymore):

  1. CostCenterManagerDialog — CRUD for Cost Centers + their rules (edit
     name/cost/match patterns/expense account/fund account directly in a
     Treeview with inline editors).

  2. BulkBackfillDialog — pick a Cost Center + filters (invoice IDs, date
     range, empty-costs-only, replace-existing-costs).  Shows a PREVIEW
     table with every invoice + matched costs + new totals.  On user
     confirmation walks every invoice via update_invoice_partial() and
     triggers LedgerService on_invoice_costs_backfilled() so GL/P&L/BS
     update cleanly.

Both dialogs are fully self-contained (import-only dependencies):
  • cost_center_manager.CostCenterManager (JSON store + apply engine)
  • hope_pharma_complete / InvoiceManager.add_invoice_from_dict /
    update_invoice_partial (used by bulk-apply)
  • ledger_service.LedgerService.on_invoice_costs_backfilled hook
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from decimal import Decimal, InvalidOperation

try:
    from cost_center_manager import CostCenterManager
    _HAS_CCM = True
except Exception:
    _HAS_CCM = False

try:
    from ledger_service import LedgerService
    _HAS_LS = True
except Exception:
    _HAS_LS = False


EXPENSE_OPTIONS = [
    "5000  COGS (Cost of Goods Sold)",
    "5100  Salary / Wages",
    "5200  Rent, Utilities (Electricity, Water, Internet)",
    "5300  Office Supplies, Travel, Marketing, Other Admin",
    "5400  Professional Fees, Audit, Legal",
    "5500  Delivery & Shipping / Logistics",
    "5600  Medical / Clinical / Device Service Fees",
    "5700  Commissions Paid to Sales Agents",
    "5800  Warranty / Returns / After-Sales",
    "5900  Other Operating Expenses",
]

MATCH_MODES = [
    ("contains",   "Contains keyword (substring, fuzzy)"),
    ("exact",      "Exact match (full description)"),
    ("starts_with","Starts with the keyword"),
    ("regex",      "Regular expression (Python re)"),
]

CALC_TYPES = [
    ("percentage",        "% of matched line-total"),
    ("fixed_per_line",    "Fixed amount × matched lines"),
    ("fixed_per_invoice", "Fixed amount once on whole invoice"),
]

FUND_ACCOUNTS_DEFAULT = ["ADCB", "Cash", "Payable / Supplier Credit", "Accrued Liabilities"]


# --------------------------------------------------------------------------- #
# Helpers                                                                     #
# --------------------------------------------------------------------------- #
def _fit_window(win, w_min=880, h_min=560, remember_key=None):
    try:
        win.update_idletasks()
        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        w = max(w_min, min(int(sw * 0.82), 1180))
        h = max(h_min, min(int(sh * 0.82), 820))
        x = (sw - w) // 2
        y = (sh - h) // 2
        win.geometry(f"{w}x{h}+{x}+{y}")
        win.minsize(w_min, h_min)
    except Exception:
        pass


def _gl_to_label(code: str) -> str:
    c = str(code or "").strip()
    for lab in EXPENSE_OPTIONS:
        if lab.startswith(c + " "):
            return lab
    if len(c) == 4 and c.isdigit():
        return f"{c}  Custom Expense"
    return EXPENSE_OPTIONS[0]


def _label_to_gl(label: str) -> str:
    s = str(label or "").strip()
    if len(s) >= 4 and s[:4].isdigit():
        return s[:4]
    return "5000"


def _fmt_aed(v) -> str:
    try:
        return f"{Decimal(str(v or 0)):,.2f}"
    except (InvalidOperation, ValueError):
        return "0.00"


# --------------------------------------------------------------------------- #
# 1. Cost Center Manager dialog                                              #
# --------------------------------------------------------------------------- #
class CostCenterManagerDialog:
    """Top-level window: cost-centers list (left) + rules table (right)."""

    def __init__(self, parent, data_folder: str, *,
                 invoice_manager=None,
                 get_all_invoices_dict_cb=None):
        if not _HAS_CCM:
            messagebox.showerror(
                "Dependency Missing",
                "CostCenterManager (cost_center_manager.py) is unavailable.\n"
                "Please make sure the file exists in the project folder."
            )
            return
        self.data_folder = data_folder
        self.ccm = CostCenterManager(data_folder)
        self.selected_center_id = None
        # Optional extras — used by ServiceLibraryDialog
        self._invoice_manager = invoice_manager
        self._get_all_invoices_dict_cb = get_all_invoices_dict_cb

        self.dlg = tk.Toplevel(parent)
        self.dlg.title("Cost Center Manager — Rule Catalog")
        self.dlg.transient(parent)
        self.dlg.grab_set()
        _fit_window(self.dlg, 1040, 620, remember_key="cost_center_manager")
        self._build_ui()
        self._refresh_centers()

    # --- UI construction --------------------------------------------------
    def _build_ui(self):
        root = ttk.Frame(self.dlg, padding=10)
        root.pack(fill="both", expand=True)

        # Header
        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(
            header,
            text="💳  Cost Center Manager",
            font=("Helvetica", 16, "bold"),
        ).pack(side="left")
        ttk.Label(
            header,
            text=(
                "Define reusable rule catalogs keyed on service/item description. "
                "Then use Bulk Backfill to apply them to past invoices."
            ),
            foreground="#4A5568",
            wraplength=620,
            justify="left",
        ).pack(side="left", padx=18)

        # Body: left (centers), right (rules)
        body = ttk.PanedWindow(root, orient="horizontal")
        body.pack(fill="both", expand=True)

        # Left: Cost Centers list + buttons
        left = ttk.LabelFrame(body, text="Cost Centers", padding=10)
        body.add(left, weight=1)

        left_btns = ttk.Frame(left)
        left_btns.pack(fill="x", pady=(0, 8))
        ttk.Button(left_btns, text="➕ New", command=self._on_new_center).pack(side="left")
        ttk.Button(left_btns, text="✏️ Rename", command=self._on_rename_center).pack(side="left", padx=6)
        ttk.Button(left_btns, text="🗑 Delete", command=self._on_delete_center).pack(side="left")
        ttk.Button(left_btns, text="🔄 Refresh", command=self._refresh_centers).pack(side="right")

        tree_cols = ("center_id", "name", "rules", "desc")
        self.centers_tree = ttk.Treeview(left, columns=tree_cols, show="headings", height=18)
        for c, w, h in [
            ("center_id", 90, "ID"),
            ("name", 200, "Name"),
            ("rules", 55, "Rules"),
            ("desc", 320, "Description"),
        ]:
            self.centers_tree.heading(c, text=h)
            self.centers_tree.column(c, width=w, anchor="w" if c != "rules" else "center")
        self.centers_tree.pack(fill="both", expand=True)
        self.centers_tree.bind("<<TreeviewSelect>>", lambda _e: self._on_select_center())

        # Right: rules table + editor
        right = ttk.LabelFrame(body, text="Rules (per matched line-item/service)", padding=10)
        body.add(right, weight=3)

        meta = ttk.Frame(right)
        meta.pack(fill="x", pady=(0, 8))
        self.meta_center = tk.StringVar(value="—")
        self.meta_desc = tk.StringVar(value="")
        ttk.Label(meta, textvariable=self.meta_center, font=("Helvetica", 11, "bold")).pack(anchor="w")
        ttk.Label(meta, text="Description / notes (optional):").pack(anchor="w", pady=(8, 2))
        desc_entry = ttk.Entry(meta, textvariable=self.meta_desc)
        desc_entry.pack(fill="x")
        desc_entry.bind(
            "<FocusOut>",
            lambda _e: self._on_change_desc(self.meta_desc.get()),
        )

        # Rules table
        rules_cols = (
            "name", "match_mode", "patterns",
            "calc_type", "value", "expense_account", "fund_account",
        )
        self.rules_tree = ttk.Treeview(right, columns=rules_cols, show="headings", height=14)
        heads = [
            ("name", "Rule Name (editable)", 200),
            ("match_mode", "Match on Service Desc", 170),
            ("patterns", "Keywords (comma or | separated)", 330),
            ("calc_type", "Calc. Type", 200),
            ("value", "Amount or %", 100),
            ("expense_account", "Expense GL", 190),
            ("fund_account", "Paid From (Fund)", 150),
        ]
        for c, h, w in heads:
            self.rules_tree.heading(c, text=h)
            self.rules_tree.column(c, width=w, anchor="w")
        self.rules_tree.pack(fill="both", expand=True, pady=(0, 8))
        self.rules_tree.bind("<Double-1>", lambda _e: self._on_edit_rule())

        # Rule buttons
        rule_btns = ttk.Frame(right)
        rule_btns.pack(fill="x")
        ttk.Button(rule_btns, text="➕ Add Rule", command=self._on_new_rule).pack(side="left")
        ttk.Button(rule_btns, text="✏️ Edit Rule", command=self._on_edit_rule).pack(side="left", padx=6)
        ttk.Button(rule_btns, text="🗑 Delete Rule", command=self._on_delete_rule).pack(side="left")
        ttk.Button(rule_btns, text="⤵ Duplicate", command=self._on_dup_rule).pack(side="left", padx=6)
        ttk.Button(rule_btns, text="📚 Service Library (Extract + Cost)",
                   command=self._on_open_service_library, style="Primary.TButton").pack(side="left", padx=10)

        ttk.Button(
            rule_btns,
            text="💾 Save Center & Rules",
            style="Accent.TButton" if hasattr(ttk.Style(), "lookup") else "TButton",
            command=self._on_save_center,
        ).pack(side="right")

        # Footer: JSON schema hint
        foot = ttk.LabelFrame(root, text="How it works", padding=8)
        foot.pack(fill="x", pady=(10, 0))
        ttk.Label(
            foot,
            text=(
                "Each rule scans the current invoice's items (the service/product descriptions you type). "
                "When a description MATCHES a rule: Calc Type runs, producing one cost line per matched rule "
                "(total = sum over matched lines for %, or flat per line, or flat once per invoice). "
                "Expense GL (5xxx) becomes the Debit; Paid From becomes the Credit (1000=Cash / 2000=Payable / "
                "2100=Accrued / 1100=Bank).  Bulk Backfill applies the selected center to ALL matching past invoices."
            ),
            foreground="#2D3748",
            wraplength=1040,
            justify="left",
        ).pack(fill="x")

        btns = ttk.Frame(root)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Done — Close", command=self.dlg.destroy).pack(side="right")

    # --- Centers helpers --------------------------------------------------
    def _refresh_centers(self, keep_select=None):
        for i in self.centers_tree.get_children():
            self.centers_tree.delete(i)
        for c in self.ccm.list_centers():
            cid = c.get("center_id", "")
            self.centers_tree.insert(
                "", "end", iid=cid,
                values=(
                    cid,
                    c.get("name", ""),
                    len(c.get("rules", []) or []),
                    c.get("description", ""),
                ),
            )
        pick = keep_select or self.selected_center_id
        if pick and pick in set(self.centers_tree.get_children()):
            self.centers_tree.selection_set(pick)
            self._on_select_center()
        elif self.centers_tree.get_children():
            first = self.centers_tree.get_children()[0]
            self.centers_tree.selection_set(first)
            self._on_select_center()
        else:
            self.selected_center_id = None
            self.meta_center.set("No Cost Center selected — click ➕ New to create one")
            self.meta_desc.set("")
            for r in self.rules_tree.get_children():
                self.rules_tree.delete(r)

    def _current_center(self):
        return self.ccm.get_center(self.selected_center_id) if self.selected_center_id else None

    def _on_select_center(self):
        sel = self.centers_tree.selection()
        if not sel:
            return
        self.selected_center_id = sel[0]
        center = self._current_center()
        if not center:
            return
        self.meta_center.set(
            f"{center.get('center_id')}  —  {center.get('name')}"
            f"   ({len(center.get('rules') or [])} rules, default fund: "
            f"{center.get('default_fund_account') or 'ADCB'})"
        )
        self.meta_desc.set(center.get("description", ""))
        for r in self.rules_tree.get_children():
            self.rules_tree.delete(r)
        for rule in (center.get("rules") or []):
            modes_label = dict((k, v) for k, v in MATCH_MODES)
            calc_label = dict((k, v) for k, v in CALC_TYPES)
            pats_raw = rule.get("patterns") or []
            pats_str = ", ".join(str(p) for p in pats_raw) if pats_raw else ""
            self.rules_tree.insert(
                "", "end", iid=str(rule.get("rule_id")),
                values=(
                    rule.get("name", ""),
                    modes_label.get(rule.get("match_mode"), rule.get("match_mode", "contains")),
                    pats_str,
                    calc_label.get(rule.get("calc_type"), rule.get("calc_type", "percentage")),
                    rule.get("value", 0),
                    _gl_to_label(rule.get("expense_account") or "5000"),
                    rule.get("fund_account") or center.get("default_fund_account") or "ADCB",
                ),
            )

    def _on_new_center(self):
        dlg = tk.Toplevel(self.dlg)
        dlg.title("New Cost Center")
        dlg.transient(self.dlg); dlg.grab_set(); dlg.geometry("440x260")
        frm = ttk.Frame(dlg, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Name:").grid(row=0, column=0, sticky="w", pady=4)
        name = ttk.Entry(frm, width=40); name.grid(row=0, column=1, pady=4)
        ttk.Label(frm, text="Description:").grid(row=1, column=0, sticky="nw", pady=4)
        desc = tk.Text(frm, width=40, height=4); desc.grid(row=1, column=1, pady=4)
        ttk.Label(frm, text="Default Paid From (Fund):").grid(row=2, column=0, sticky="w", pady=4)
        fund = ttk.Combobox(frm, width=36, values=FUND_ACCOUNTS_DEFAULT, state="normal")
        fund.current(0); fund.grid(row=2, column=1, pady=4)
        ttk.Label(frm, text="Default Expense GL:").grid(row=3, column=0, sticky="w", pady=4)
        exp = ttk.Combobox(frm, width=36, values=EXPENSE_OPTIONS, state="readonly")
        exp.current(0); exp.grid(row=3, column=1, pady=4)

        def ok():
            n = name.get().strip()
            if not n:
                messagebox.showwarning("Missing", "Please enter a name", parent=dlg)
                return
            try:
                created = self.ccm.create_center(
                    name=n,
                    description=desc.get("1.0", "end").strip(),
                    default_fund_account=fund.get().strip() or "ADCB",
                    default_expense_account=_label_to_gl(exp.get()),
                )
            except Exception as exc:
                messagebox.showerror("Error", str(exc), parent=dlg)
                return
            cid = created.get("center_id")
            self._refresh_centers(keep_select=cid)
            dlg.destroy()

        ttk.Button(frm, text="Create", command=ok).grid(row=4, column=0, pady=12)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=4, column=1, pady=12)

    def _on_rename_center(self):
        c = self._current_center()
        if not c:
            return
        dlg = tk.Toplevel(self.dlg)
        dlg.title("Rename / Edit Cost Center Info")
        dlg.transient(self.dlg); dlg.grab_set(); dlg.geometry("440x260")
        frm = ttk.Frame(dlg, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Name:").grid(row=0, column=0, sticky="w", pady=4)
        name = ttk.Entry(frm, width=40); name.insert(0, c.get("name", ""))
        name.grid(row=0, column=1, pady=4)
        ttk.Label(frm, text="Description:").grid(row=1, column=0, sticky="nw", pady=4)
        desc = tk.Text(frm, width=40, height=4)
        desc.insert("1.0", c.get("description", "")); desc.grid(row=1, column=1, pady=4)
        ttk.Label(frm, text="Default Paid From (Fund):").grid(row=2, column=0, sticky="w", pady=4)
        fund = ttk.Combobox(frm, width=36, values=FUND_ACCOUNTS_DEFAULT, state="normal")
        fund.set(c.get("default_fund_account") or "ADCB"); fund.grid(row=2, column=1, pady=4)
        ttk.Label(frm, text="Default Expense GL:").grid(row=3, column=0, sticky="w", pady=4)
        exp = ttk.Combobox(frm, width=36, values=EXPENSE_OPTIONS, state="readonly")
        exp.set(_gl_to_label(c.get("default_expense_account") or "5000"))
        exp.grid(row=3, column=1, pady=4)

        def ok():
            try:
                updated = self.ccm.update_center(
                    c["center_id"],
                    name=name.get().strip() or c.get("name", ""),
                    description=desc.get("1.0", "end").strip(),
                    default_fund_account=fund.get().strip() or "ADCB",
                    default_expense_account=_label_to_gl(exp.get()),
                )
            except Exception as exc:
                messagebox.showerror("Error", str(exc), parent=dlg); return
            if not updated:
                messagebox.showerror("Error", "Cost Center not found", parent=dlg); return
            self._refresh_centers(keep_select=c["center_id"])
            dlg.destroy()

        ttk.Button(frm, text="Save", command=ok).grid(row=4, column=0, pady=12)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=4, column=1, pady=12)

    def _on_delete_center(self):
        c = self._current_center()
        if not c:
            return
        if not messagebox.askyesno(
            "Delete Cost Center",
            f"Delete '{c.get('name')}' ({c.get('center_id')}) with "
            f"{len(c.get('rules') or [])} rule(s)?\n\nThis cannot be undone.",
            icon="warning",
        ):
            return
        ok, msg = self.ccm.delete_center(c["center_id"])
        if not ok:
            messagebox.showerror("Error", msg); return
        self._refresh_centers()

    def _on_change_desc(self, new_desc):
        c = self._current_center()
        if not c:
            return
        # Debounced soft-save for the description field only (keeps other fields)
        try:
            self.ccm.update_center(c["center_id"], description=str(new_desc or "").strip())
        except Exception:
            pass

    # --- Rules CRUD -------------------------------------------------------
    def _collect_rules_from_table(self):
        rules = []
        modes_back = {v: k for k, v in MATCH_MODES}
        calc_back = {v: k for k, v in CALC_TYPES}
        for iid in self.rules_tree.get_children():
            vals = self.rules_tree.item(iid)["values"]
            name = str(vals[0] or "").strip()
            mode_lbl = str(vals[1])
            pats_raw = [p.strip() for p in str(vals[2]).replace("|", ",").split(",") if p.strip()]
            calc_lbl = str(vals[3])
            try:
                value = float(vals[4])
            except Exception:
                value = 0.0
            exp_label = str(vals[5])
            fund = str(vals[6] or "").strip()
            if not name:
                continue
            rules.append({
                "rule_id": iid if str(iid).startswith("R") else None,
                "name": name,
                "match_mode": modes_back.get(mode_lbl, "contains"),
                "patterns": pats_raw,
                "calc_type": calc_back.get(calc_lbl, "percentage"),
                "value": value,
                "expense_account": _label_to_gl(exp_label),
                "fund_account": fund,
            })
        return rules

    def _editor_dialog(self, initial=None, title="Edit Rule"):
        dlg = tk.Toplevel(self.dlg)
        dlg.title(title); dlg.transient(self.dlg); dlg.grab_set(); dlg.geometry("520x440")
        frm = ttk.Frame(dlg, padding=12); frm.pack(fill="both", expand=True)

        init = initial or {}
        ttk.Label(frm, text="Rule Name:").grid(row=0, column=0, sticky="w", pady=4)
        v_name = ttk.Entry(frm, width=44); v_name.insert(0, init.get("name", ""))
        v_name.grid(row=0, column=1, pady=4)

        ttk.Label(frm, text="Match Mode (service description):").grid(row=1, column=0, sticky="w", pady=4)
        modes_labels = [v for _, v in MATCH_MODES]
        modes_back = {v: k for k, v in MATCH_MODES}
        init_mode = modes_back.get(init.get("match_mode"), modes_labels[0])
        v_mode = ttk.Combobox(frm, values=modes_labels, state="readonly", width=42)
        v_mode.set(init_mode); v_mode.grid(row=1, column=1, pady=4)

        ttk.Label(frm, text="Keywords (comma or | separated):\n(leave empty for fixed_per_invoice global)")\
            .grid(row=2, column=0, sticky="nw", pady=4)
        v_pats = tk.Text(frm, width=44, height=4)
        v_pats.insert("1.0", ", ".join(init.get("patterns") or []))
        v_pats.grid(row=2, column=1, pady=4)

        ttk.Label(frm, text="Calc Type:").grid(row=3, column=0, sticky="w", pady=4)
        calc_labels = [v for _, v in CALC_TYPES]
        calc_back = {v: k for k, v in CALC_TYPES}
        init_calc = calc_back.get(init.get("calc_type"), calc_labels[0])
        v_calc = ttk.Combobox(frm, values=calc_labels, state="readonly", width=42)
        v_calc.set(init_calc); v_calc.grid(row=3, column=1, pady=4)

        ttk.Label(frm, text="Value (amount or %):\n(% => 0.2 = 20%)")\
            .grid(row=4, column=0, sticky="nw", pady=4)
        v_val = ttk.Entry(frm, width=20); v_val.insert(0, str(init.get("value", 0)))
        v_val.grid(row=4, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Expense GL:").grid(row=5, column=0, sticky="w", pady=4)
        v_exp = ttk.Combobox(frm, values=EXPENSE_OPTIONS, state="readonly", width=42)
        v_exp.set(_gl_to_label(init.get("expense_account") or "5000"))
        v_exp.grid(row=5, column=1, pady=4)

        ttk.Label(frm, text="Paid From (Fund):").grid(row=6, column=0, sticky="w", pady=4)
        v_fund = ttk.Combobox(frm, values=FUND_ACCOUNTS_DEFAULT, state="normal", width=42)
        c = self._current_center()
        default_fund = (init.get("fund_account")
                        or (c.get("default_fund_account") if c else None)
                        or "ADCB")
        v_fund.set(default_fund); v_fund.grid(row=6, column=1, pady=4)

        result = {}

        def ok():
            try:
                val = float(v_val.get())
            except Exception:
                messagebox.showwarning("Invalid", "Value must be a number", parent=dlg); return
            name = v_name.get().strip()
            if not name:
                messagebox.showwarning("Missing", "Rule Name is required", parent=dlg); return
            pats = [p.strip() for p in v_pats.get("1.0", "end").replace("|", ",").split(",") if p.strip()]
            modes_back2 = dict((v, k) for k, v in MATCH_MODES)
            calc_back2 = dict((v, k) for k, v in CALC_TYPES)
            result.update({
                "name": name,
                "match_mode": modes_back2.get(v_mode.get(), "contains"),
                "patterns": pats,
                "calc_type": calc_back2.get(v_calc.get(), "percentage"),
                "value": val,
                "expense_account": _label_to_gl(v_exp.get()),
                "fund_account": v_fund.get().strip() or "ADCB",
            })
            dlg.destroy()

        ttk.Button(frm, text="Save", command=ok).grid(row=7, column=0, pady=16)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=7, column=1, pady=16)
        dlg.wait_window()
        return result

    def _selected_rule_init(self):
        sel = self.rules_tree.selection()
        if not sel:
            return None, None
        iid = sel[0]
        vals = self.rules_tree.item(iid)["values"]
        modes_back = {v: k for k, v in MATCH_MODES}
        calc_back = {v: k for k, v in CALC_TYPES}
        pats = [p.strip() for p in str(vals[2]).replace("|", ",").split(",") if p.strip()]
        try:
            value = float(vals[4])
        except Exception:
            value = 0.0
        return iid, {
            "name": vals[0],
            "match_mode": modes_back.get(vals[1], "contains"),
            "patterns": pats,
            "calc_type": calc_back.get(vals[3], "percentage"),
            "value": value,
            "expense_account": _label_to_gl(vals[5]),
            "fund_account": vals[6] if len(vals) > 6 else "",
        }

    def _on_new_rule(self):
        if not self._current_center():
            messagebox.showwarning("Missing", "First create/select a Cost Center")
            return
        res = self._editor_dialog(title="➕ Add Rule")
        if not res:
            return
        # insert with placeholder iid; real rule_id assigned on save
        iid = "R__NEW__" + str(len(self.rules_tree.get_children()))
        modes_label = dict(MATCH_MODES)
        calc_label = dict(CALC_TYPES)
        self.rules_tree.insert(
            "", "end", iid=iid,
            values=(
                res["name"], modes_label[res["match_mode"]],
                ", ".join(res["patterns"] or []),
                calc_label[res["calc_type"]],
                res["value"], _gl_to_label(res["expense_account"]),
                res["fund_account"],
            ),
        )

    def _on_edit_rule(self):
        iid, init = self._selected_rule_init()
        if not init:
            messagebox.showwarning("Missing", "Select a rule to edit"); return
        res = self._editor_dialog(initial=init, title="✏️ Edit Rule")
        if not res:
            return
        modes_label = dict(MATCH_MODES); calc_label = dict(CALC_TYPES)
        self.rules_tree.item(
            iid,
            values=(
                res["name"], modes_label[res["match_mode"]],
                ", ".join(res["patterns"] or []),
                calc_label[res["calc_type"]],
                res["value"], _gl_to_label(res["expense_account"]),
                res["fund_account"],
            ),
        )

    def _on_delete_rule(self):
        sel = self.rules_tree.selection()
        if not sel:
            return
        for iid in sel:
            self.rules_tree.delete(iid)

    def _on_dup_rule(self):
        iid, init = self._selected_rule_init()
        if not init:
            return
        init["name"] = init.get("name", "") + " (copy)"
        res = self._editor_dialog(initial=init, title="⤵ Duplicate Rule")
        if not res:
            return
        new_iid = "R__NEW__" + str(len(self.rules_tree.get_children()))
        modes_label = dict(MATCH_MODES); calc_label = dict(CALC_TYPES)
        self.rules_tree.insert(
            "", "end", iid=new_iid,
            values=(
                res["name"], modes_label[res["match_mode"]],
                ", ".join(res["patterns"] or []),
                calc_label[res["calc_type"]],
                res["value"], _gl_to_label(res["expense_account"]),
                res["fund_account"],
            ),
        )

    def _on_save_center(self):
        c = self._current_center()
        if not c:
            messagebox.showwarning("Missing", "Select a Cost Center first"); return
        rules = self._collect_rules_from_table()
        res = self.ccm.update_center(
            c["center_id"],
            rules=rules,
            description=self.meta_desc.get().strip(),
        )
        ok = bool(res)
        if not ok:
            messagebox.showerror("Save Failed", "update_center returned empty"); return
        messagebox.showinfo(
            "Saved",
            f"{c['center_id']} '{c['name']}' saved.\n"
            f"{len(rules)} rule(s) persisted to {os.path.join(self.data_folder, 'cost_centers.json')}"
        )
        self._refresh_centers(keep_select=c["center_id"])

    def _on_open_service_library(self):
        """Open the Service Library dialog (extract services or type manual + set costs)."""
        if not self.selected_center_id:
            messagebox.showwarning("Missing", "First create or select a Cost Center in the left list."); return
        try:
            ServiceLibraryDialog
        except NameError:
            messagebox.showwarning(
                "Not Available",
                "Service library is only available from the Hub Cost Centers button."
            ); return
        def _after_rules(cid):
            self._refresh_centers(keep_select=cid)
        ServiceLibraryDialog(
            self.dlg,
            ccm=self.ccm,
            center_id=self.selected_center_id,
            invoice_manager=self._invoice_manager,
            data_folder=self.data_folder,
            get_all_invoices_cb=self._get_all_invoices_dict_cb,
            on_rules_created_cb=_after_rules,
        )


# --------------------------------------------------------------------------- #
# 2. Bulk Backfill dialog (filter → preview → confirm → apply)              #
# --------------------------------------------------------------------------- #
class BulkBackfillDialog:
    """
    Apply a Cost Center across many past invoices.  User picks filters, hits
    Preview, sees every affected invoice with generated totals, confirms, and
    we call update_invoice_partial → LedgerService hooks.
    """

    def __init__(self, parent, invoice_manager, data_folder: str):
        if not _HAS_CCM:
            messagebox.showerror(
                "Dependency Missing",
                "CostCenterManager (cost_center_manager.py) is unavailable."
            )
            return
        self.invoice_manager = invoice_manager  # aka the HopePharmaComplete manager
        self.data_folder = data_folder
        self.ccm = CostCenterManager(data_folder)
        self.preview_data = None  # list of row dicts for apply

        self.dlg = tk.Toplevel(parent)
        self.dlg.title("Bulk Backfill — Apply Cost Center To Past Invoices")
        self.dlg.transient(parent)
        self.dlg.grab_set()
        _fit_window(self.dlg, 1100, 660, remember_key="bulk_backfill_cc")
        self._build_ui()
        self._refresh_center_list()

    def _build_ui(self):
        root = ttk.Frame(self.dlg, padding=10)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 10))
        ttk.Label(
            header,
            text="🚀  Bulk Backfill — Auto-apply Cost Center To All Invoices",
            font=("Helvetica", 16, "bold"),
        ).pack(side="left")
        ttk.Label(
            header,
            text=(
                "Step 1: Pick Cost Center + filters.  Step 2: Preview (shows a per-invoice breakdown of what will "
                "change).  Step 3: Confirm & Apply — updates every invoice plus the GL cleanly."
            ),
            foreground="#4A5568",
            wraplength=620,
            justify="left",
        ).pack(side="left", padx=18)

        # Filters
        filters = ttk.LabelFrame(root, text="Filters", padding=10)
        filters.pack(fill="x")

        r1 = ttk.Frame(filters); r1.pack(fill="x", pady=2)
        ttk.Label(r1, text="Cost Center *:").pack(side="left")
        self.var_cid = tk.StringVar()
        self.cb_cid = ttk.Combobox(r1, textvariable=self.var_cid, width=54, state="readonly")
        self.cb_cid.pack(side="left", padx=6)
        ttk.Button(r1, text="🔄 Refresh", command=self._refresh_center_list).pack(side="left", padx=6)

        ttk.Label(r1, text="      Date range:").pack(side="left")
        self.var_from = tk.StringVar(value="")
        self.var_to = tk.StringVar(value="")
        ttk.Entry(r1, textvariable=self.var_from, width=12).pack(side="left", padx=2)
        ttk.Label(r1, text="→").pack(side="left")
        ttk.Entry(r1, textvariable=self.var_to, width=12).pack(side="left", padx=2)
        ttk.Label(r1, text="(YYYY-MM-DD)", foreground="#718096").pack(side="left", padx=4)

        r2 = ttk.Frame(filters); r2.pack(fill="x", pady=8)
        self.var_empty_only = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            r2,
            text="Only invoices with NO COSTS yet (recommended — backfill blanks)",
            variable=self.var_empty_only,
        ).pack(side="left")

        self.var_replace = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            r2,
            text="Replace existing costs on matched invoices (⚠️ destructive)",
            variable=self.var_replace,
        ).pack(side="left", padx=20)

        ttk.Label(r2, text="Invoice IDs (comma-separated, blank = all):").pack(side="left", padx=(20, 4))
        self.var_ids = tk.StringVar()
        ttk.Entry(r2, textvariable=self.var_ids, width=32).pack(side="left")

        btns = ttk.Frame(filters); btns.pack(fill="x", pady=(6, 0))
        ttk.Button(btns, text="🧪 1. Preview Impact", command=self._on_preview,
                   style="Primary.TButton" if False else "TButton").pack(side="left")
        self.btn_apply = ttk.Button(btns, text="✅ 2. Confirm & Apply To All",
                                    command=self._on_apply, state="disabled")
        self.btn_apply.pack(side="left", padx=10)
        self.lbl_summary = ttk.Label(btns, text="", foreground="#2B6CB0")
        self.lbl_summary.pack(side="right")

        # Preview table
        pvf = ttk.LabelFrame(root, text="Preview (each invoice + computed costs)", padding=8)
        pvf.pack(fill="both", expand=True, pady=(10, 0))

        cols = ("invoice_id", "date", "client", "items_count",
                "old_cost", "new_cost", "delta", "new_cost_count", "grand_total", "status")
        self.pt = ttk.Treeview(pvf, columns=cols, show="headings", height=20)
        heads = [
            ("invoice_id", "Invoice ID", 130),
            ("date",       "Date",        110),
            ("client",     "Client",      220),
            ("items_count","#Items",      70),
            ("old_cost",   "Old Costs",   110),
            ("new_cost",   "New Costs",   110),
            ("delta",      "Δ Costs",     110),
            ("new_cost_count","New #Costs",90),
            ("grand_total","Grand Total", 120),
            ("status",     "Status / Notes", 320),
        ]
        for c, h, w in heads:
            self.pt.heading(c, text=h)
            self.pt.column(c, width=w, anchor="w" if c in ("invoice_id","date","client","status") else "e")
        self.pt.pack(fill="both", expand=True, side="left")
        sb = ttk.Scrollbar(pvf, orient="vertical", command=self.pt.yview)
        self.pt.configure(yscrollcommand=sb.set); sb.pack(side="right", fill="y")
        self.pt.tag_configure("ok", background="#F0FFF4")
        self.pt.tag_configure("nomatch", background="#FFFBEB")
        self.pt.tag_configure("skip", background="#EDF2F7")

        # Footer actions
        f = ttk.Frame(root); f.pack(fill="x", pady=10)
        ttk.Label(f,
                  text="⚠️ Apply is permanent but reversible: it posts standard GAAP reversals then new entries. "
                       "You can always open each affected invoice and edit costs manually.",
                  foreground="#744210", wraplength=960, justify="left").pack(side="left")
        ttk.Button(f, text="Close", command=self.dlg.destroy).pack(side="right")

    # ---- Preview / Apply -------------------------------------------------
    def _refresh_center_list(self):
        try:
            centers = self.ccm.list_centers()
        except Exception:
            centers = []
        vals = []
        self._cid_map = {}
        for c in centers:
            label = f"{c.get('center_id')}   {c.get('name')}   ({len(c.get('rules') or [])} rules)"
            vals.append(label)
            self._cid_map[label] = c.get("center_id")
        self.cb_cid["values"] = vals
        if vals:
            self.cb_cid.current(0)

    def _picked_center_id(self):
        return self._cid_map.get((self.var_cid.get() or "").strip())

    def _all_invoices(self):
        # Try the manager's in-memory list; fall back to get_all_invoices / invoices property
        mgr = self.invoice_manager
        for attr in ("invoices", "get_all_invoices", "list_invoices"):
            try:
                val = getattr(mgr, attr)
            except AttributeError:
                continue
            if callable(val):
                try:
                    out = val()
                except Exception:
                    continue
                if isinstance(out, list):
                    return [self._dictize(x) for x in out]
            if isinstance(val, list):
                return [self._dictize(x) for x in val]
        # Fallback: JSON file
        p = os.path.join(self.data_folder, "invoices_data.json")
        if os.path.exists(p):
            import json
            with open(p, "r") as fh:
                try:
                    data = json.load(fh)
                    if isinstance(data, list):
                        return data
                    if isinstance(data, dict) and isinstance(data.get("invoices"), list):
                        return data["invoices"]
                except Exception:
                    pass
        return []

    def _dictize(self, obj):
        if isinstance(obj, dict):
            return obj
        # HopePharmaComplete Invoice object
        out = {}
        for k in ("invoice_id","invoice_type","date","client_name","client_trn","items","costs",
                  "subtotal","tax_amount","grand_total","total_cost","profit_loss",
                  "payment_method","payment_due","due_date","status","notes","currency"):
            try:
                out[k] = getattr(obj, k)
            except Exception:
                pass
        if "costs" not in out:
            out["costs"] = []
        if "items" not in out:
            out["items"] = []
        return out

    def _on_preview(self):
        cid = self._picked_center_id()
        if not cid:
            messagebox.showwarning("Missing", "Pick a Cost Center first")
            return
        center = self.ccm.get_center(cid)
        if not center:
            return
        min_d = (self.var_from.get() or "").strip() or None
        max_d = (self.var_to.get() or "").strip() or None
        ids_raw = (self.var_ids.get() or "").strip()
        id_set = set([i.strip() for i in ids_raw.split(",") if i.strip()]) or None
        replace_existing = bool(self.var_replace.get())
        only_empty = bool(self.var_empty_only.get())

        invoices = self._all_invoices()

        def passes(inv):
            if id_set and inv.get("invoice_id") not in id_set:
                return False
            d = str(inv.get("date") or "")
            if min_d and d < min_d:
                return False
            if max_d and d > max_d:
                return False
            if only_empty:
                costs = inv.get("costs") or []
                if len(costs) != 0:
                    try:
                        if sum(float(c.get("amount") or 0) for c in costs) != 0:
                            return False
                    except Exception:
                        return False
            return True

        filtered = [i for i in invoices if passes(i)]
        preview = self.ccm.backfill_preview(
            center_id=cid,
            invoices=filtered,
            replace_existing_costs=replace_existing,
            min_date=min_d,
            max_date=max_d,
            invoice_ids=id_set,
        )
        preview_rows = preview.get("preview_rows") or []
        # Build a "status" text per row
        summary_lines = [
            f"{preview.get('invoices_count') or 0} affected invoices, "
            f"{preview.get('total_cost_lines') or 0} cost lines generated, "
            f"total AED {preview.get('total_cost_amount') or 0:,.2f}; "
            f"skipped (no rules matched lines): {preview.get('skipped_no_lines') or 0}, "
            f"skipped (has costs & replace not ticked): {preview.get('skipped_has_costs') or 0}, "
            f"skipped outside date filter: {preview.get('skipped_outside_dates') or 0}.",
        ]
        summary = summary_lines[0]
        for i in self.pt.get_children():
            self.pt.delete(i)
        rows = 0
        ok_rows = 0
        nomatch = 0
        skipped = 0
        tot_delta = Decimal("0")
        for row in preview_rows:
            inv_id = str(row.get("invoice_id") or "")
            client = str(row.get("client") or "")
            date_ = str(row.get("date") or "")
            # Estimate #items from invoice dict lookup (best-effort)
            inv_lookup = next((x for x in filtered if str(x.get("invoice_id")) == inv_id), None)
            items_n = len((inv_lookup or {}).get("items") or []) if inv_lookup else "—"
            old_c = Decimal(str(row.get("prev_cost_total") or 0))
            new_c = Decimal(str(row.get("new_cost_total") or 0))
            delta = new_c - old_c
            n_lines = int(row.get("num_costs") or len(row.get("new_costs") or []))
            info = (row.get("info_lines") or [])
            status = info[0] if info else (
                f"{n_lines} rule-match cost lines"
                + ("  [replace]" if bool(row.get("replace_existing")) else "  [append]")
            )
            gt = Decimal(str(row.get("grand_total") or inv_lookup.get("grand_total", 0) if inv_lookup else 0))
            tags = ()
            if n_lines == 0:
                tags = ("nomatch",); nomatch += 1
            elif (not bool(row.get("replace_existing"))) and old_c > 0:
                tags = ("skip",); skipped += 1
            else:
                tags = ("ok",); ok_rows += 1
            rows += 1
            tot_delta += delta
            self.pt.insert(
                "", "end", iid=inv_id or str(rows), tags=tags,
                values=(
                    inv_id, date_, client, items_n,
                    _fmt_aed(old_c), _fmt_aed(new_c),
                    f"{delta:+,.2f}", n_lines, _fmt_aed(gt), status,
                ),
            )
        self.preview_data = preview  # full dict (contains preview_rows + replace_existing_costs)
        self.btn_apply.configure(state="normal" if ok_rows else "disabled")
        self.lbl_summary.config(
            text=(f"Total {rows} rows  ·  matched: {ok_rows}  ·  no-match: {nomatch}  ·  "
                  f"skip/kept: {skipped}  ·  Σ Δ Costs = AED {tot_delta:+,.2f}" +
                  (f"  |  Summary: {summary}" if summary else "")),
        )
        if rows == 0:
            messagebox.showinfo(
                "Nothing to preview",
                "No invoices passed filters.\n\nTips: try unchecking 'Only invoices with NO COSTS yet', "
                "clear date range, or verify your Cost Center rules actually match item descriptions "
                "(e.g. keywords 'SONIC' for hearing aids, 'Battery' for batteries)."
            )

    def _on_apply(self):
        if not self.preview_data:
            return
        summary_lbl = self.lbl_summary.cget("text") or ""
        if not messagebox.askyesno(
            "Confirm Apply",
            "You are about to apply a Cost Center across all matched invoices above.\n"
            "Each invoice's costs will be updated, and the General Ledger will be posted\n"
            "(reversing previous costs where applicable, then posting new Dr/Cr lines).\n\n"
            f"{summary_lbl}\n\nProceed? This cannot be undone automatically.",
            icon="warning",
        ):
            return

        mgr = self.invoice_manager

        # apply via CostCenterManager — handles store with update_invoice_partial method
        store_adapter = _InvoiceStoreAdapter(mgr, self.data_folder)
        result = self.ccm.apply_backfill(
            preview=self.preview_data,
            invoices_store=store_adapter,
            save_method_name=None,
            update_method_name="update_invoice_partial",
        )
        applied = int(result.get("applied") or 0)
        errors = result.get("errors") or []
        failed = len(errors or [])

        # Now post GL bulk backfill hook (GAAP reversals + new entries)
        gl_ok = 0
        gl_errs = []
        if _HAS_LS and applied:
            try:
                ls = LedgerService(self.data_folder)
                # Build iterator matching on_invoice_costs_backfilled expected schema.
                # Pull invoice date / client via best-effort preview + store lookup
                def _gl_iter():
                    for row in (self.preview_data or {}).get("preview_rows") or []:
                        inv_id = row.get("invoice_id")
                        if not inv_id:
                            continue
                        existing_costs = []
                        if store_adapter and hasattr(store_adapter, "mgr"):
                            try:
                                mg = store_adapter.mgr
                                if hasattr(mg, "get_invoice"):
                                    inv = mg.get_invoice(inv_id)
                                    if inv:
                                        existing_costs = inv.get("costs") if isinstance(inv, dict) else list(getattr(inv, "costs", []) or [])
                            except Exception:
                                pass
                        yield {
                            "invoice_id": inv_id,
                            "costs_list": row.get("new_costs") or [],
                            "invoice_date": row.get("date"),
                            "invoice_client": row.get("client") or "",
                            "operation": "backfill",
                            "previous_costs_list": existing_costs,
                        }
                res = ls.on_invoice_costs_backfilled(_gl_iter())
                gl_ok = int(res.get("posted_entries") or 0)
                gl_errs = [f"{e.get('invoice_id')}: {e.get('error')}" for e in (res.get("errors") or [])]
            except Exception as e:
                gl_errs = [f"Ledger hook failed: {e}"]

        msg = (
            f"Applied to {applied} invoice(s).  Failures: {failed}.\n"
            f"LedgerService posted: {gl_ok} entries  ·  GL errors: {len(gl_errs)}"
        )
        if errors:
            msg += "\n\nSample update errors:\n  • " + "\n  • ".join(errors[:6])
        if gl_errs:
            msg += "\n\nSample GL errors:\n  • " + "\n  • ".join(gl_errs[:6])
        messagebox.showinfo("Bulk Backfill Complete", msg)

        # Reset preview
        self.preview_data = None
        self.btn_apply.configure(state="disabled")
        self.lbl_summary.config(text=self.lbl_summary.cget("text") + "  —  DONE ✓")
        # Try to refresh the manager dashboard (best-effort)
        try:
            if hasattr(mgr, "rebuild_caches") and callable(mgr.rebuild_caches):
                mgr.rebuild_caches()
        except Exception:
            pass


# Adapter used by CostCenterManager.apply_backfill — converts update_invoice_partial call
# into whatever interface the real manager exposes (add_invoice_from_dict under the hood
# already supports partial via the update_invoice_partial helper in hope_pharma_complete).
class _InvoiceStoreAdapter:
    def __init__(self, mgr, data_folder):
        self.mgr = mgr
        self.data_folder = data_folder

    # Attribute forwarders
    def __getattr__(self, item):
        # Fallback — only called for attrs not defined on the adapter
        return getattr(self.mgr, item)

    # Standard interface that apply_backfill can rely on
    def update_invoice_partial(self, invoice_id, partial_dict):
        if hasattr(self.mgr, "update_invoice_partial"):
            return self.mgr.update_invoice_partial(invoice_id, partial_dict)
        # Fallback: reconstruct by merging into existing dict then calling add_invoice_from_dict
        try:
            inv = None
            if hasattr(self.mgr, "get_invoice"):
                inv = self.mgr.get_invoice(invoice_id)
            if inv is None:
                # fallback via json
                p = os.path.join(self.data_folder, "invoices_data.json")
                import json
                with open(p, "r") as fh:
                    data = json.load(fh)
                all_inv = data if isinstance(data, list) else data.get("invoices", [])
                for d in all_inv:
                    if str(d.get("invoice_id")) == str(invoice_id):
                        inv = d; break
            if not inv:
                return False, f"Invoice {invoice_id} not found for partial update"
            existing = inv if isinstance(inv, dict) else {k: getattr(inv, k) for k in vars(inv)}
            merged = {**existing, **(partial_dict or {})}
            merged["invoice_id"] = invoice_id
            ok = self.mgr.add_invoice_from_dict(merged)
            return bool(ok), ("" if ok else "add_invoice_from_dict returned falsy")
        except Exception as e:
            return False, str(e)


# --------------------------------------------------------------------------- #
# 3. Service Library dialog — extract services from past invoices or manual  #
# --------------------------------------------------------------------------- #
class ServiceLibraryDialog:
    """
    Popup from inside CostCenterManagerDialog (on rule pane).

    Two user workflows, BOTH supported per user request:
      A) "Extract all services from past invoices" → scan invoice items, dedupe
         by description, show usage count, editable cost.
      B) "I just type these services manually" → Add Row button, user types any
         service description + cost, even if no invoices have it yet.

    Result: user selects multiple rows (checkbox), sets Cost AED on each, hits
    **Bulk Add Selected → Rules** and each row becomes a rule in the active
    Cost Center (pattern = exact service description, match_mode = contains,
    calc_type = fixed_per_invoice by default, amount = cost user entered).

    SAFETY FEATURES (added per user's "makes all same costs" bug report):
      • Full UNDO / REDO stack (50 levels) — every destructive action is snapshotted.
      • "Set default Cost AED for selected → Apply" — if 0 rows selected, user
        MUST explicitly confirm before the same cost is applied to ALL services
        (this was the root cause of "everything has the same cost").
      • Snapshot is taken AFTER a successful "Bulk Add Ticked → Rules", so the
        Undo button always knows how to revert to the last known "submitted" state.
    """

    _UNDO_MAX = 50

    def __init__(self, parent, ccm, center_id, invoice_manager=None, data_folder=None,
                 get_all_invoices_cb=None, on_rules_created_cb=None):
        self.ccm = ccm
        self.center_id = center_id
        self.invoice_manager = invoice_manager
        self.data_folder = data_folder or (
            getattr(invoice_manager, "invoice_folder", None) if invoice_manager else None
        )
        self.get_all_invoices_cb = get_all_invoices_cb
        self.on_rules_created_cb = on_rules_created_cb

        # Undo / Redo stacks: each entry = (reason: str, rows_deepcopy: list[dict])
        self._undo = []
        self._redo = []

        self.dlg = tk.Toplevel(parent)
        self.dlg.title("Service Library — extract services or type manually, assign costs")
        self.dlg.transient(parent)
        self.dlg.grab_set()
        _fit_window(self.dlg, 1180, 720, remember_key="service_library")
        self._build()
        # Auto-load services on open
        self.dlg.after(120, lambda: self._on_extract(silent=True))

    # -------------------- Undo / Redo helpers ----------------------------
    @staticmethod
    def _copy_rows(rows):
        import copy
        return copy.deepcopy(rows)

    def _snapshot(self, reason: str):
        """Push current rows onto undo stack.  Wipes redo stack."""
        self._undo.append((reason, self._copy_rows(self.rows)))
        if len(self._undo) > self._UNDO_MAX:
            self._undo.pop(0)
        self._redo.clear()
        self._refresh_undo_buttons()

    def _restore(self, rows_snapshot):
        self.rows = rows_snapshot
        self._render_rows()

    def _undo_click(self):
        if not self._undo:
            return
        reason, rows_now = self._undo.pop()
        self._redo.append((reason, self._copy_rows(self.rows)))
        self._restore(rows_now)
        self._lbl_undo_status.config(text=f"↶ Undid: {reason}", foreground="#B7791F")
        self._refresh_undo_buttons()

    def _redo_click(self):
        if not self._redo:
            return
        reason, rows_next = self._redo.pop()
        self._undo.append((reason, self._copy_rows(self.rows)))
        self._restore(rows_next)
        self._lbl_undo_status.config(text=f"↷ Redid: {reason}", foreground="#2B6CB0")
        self._refresh_undo_buttons()

    def _refresh_undo_buttons(self):
        try:
            self._btn_undo.config(state="normal" if self._undo else "disabled")
            self._btn_redo.config(state="normal" if self._redo else "disabled")
        except Exception:
            pass

    def _build(self):
        root = ttk.Frame(self.dlg, padding=10)
        root.pack(fill="both", expand=True)

        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(
            header,
            text="📚  Service Library — assign a cost to each service → one click creates rules",
            font=("Helvetica", 15, "bold"),
        ).pack(side="left")
        ttk.Label(
            header,
            text=(
                "Workflow:  1) Extract services from ALL your past invoices (top-left).  "
                "2) For each service type a cost (e.g. MOH Registration = 500).  "
                "3) Tick rows → Bulk Add Selected → Rules in your Cost Center.  "
                "4) Then Bulk Backfill to apply those costs to every invoice that has that service.   "
                "SAFETY: ↶Undo/↷Redo always in the top-right corner if anything goes wrong."
            ),
            foreground="#2D3748",
            wraplength=640,
            justify="left",
        ).pack(side="left", padx=14)
        # Undo / Redo buttons (top-right) — first thing added so user cannot miss them
        urf = ttk.Frame(header); urf.pack(side="right")
        self._btn_undo = ttk.Button(urf, text="↶ Undo", command=self._undo_click, state="disabled")
        self._btn_undo.pack(side="left")
        self._btn_redo = ttk.Button(urf, text="↷ Redo", command=self._redo_click, state="disabled")
        self._btn_redo.pack(side="left", padx=(4, 0))

        tool = ttk.Frame(root)
        tool.pack(fill="x", pady=(0, 8))
        ttk.Button(tool, text="🔍 1) Extract Services From Past Invoices",
                   command=self._on_extract, style="Primary.TButton").pack(side="left")
        ttk.Label(tool, text="    Filter:").pack(side="left", padx=(20, 4))
        self.search_var = tk.StringVar()
        e = ttk.Entry(tool, textvariable=self.search_var, width=28)
        e.pack(side="left")
        e.bind("<KeyRelease>", lambda ev: self._render_rows())
        ttk.Button(tool, text="➕ Add Service Manually", command=self._on_add_manual).pack(side="left", padx=10)
        ttk.Button(tool, text="🗑 Delete Selected Rows", command=self._on_delete_rows).pack(side="left", padx=2)
        ttk.Label(tool, text="    Set default Cost AED for selected:").pack(side="left", padx=(14, 4))
        self.bulk_cost_var = tk.StringVar(value="")
        bce = ttk.Entry(tool, textvariable=self.bulk_cost_var, width=10)
        bce.pack(side="left")
        ttk.Button(tool, text="Apply", command=self._on_bulk_apply_cost).pack(side="left", padx=4)
        # Undo status line (below toolbar)
        self._lbl_undo_status = ttk.Label(tool, text="Safe: 50-level Undo/Redo enabled (↶ top-right).", foreground="#718096")
        self._lbl_undo_status.pack(side="right", padx=8)

        # Main table with inline editing
        tvf = ttk.LabelFrame(root, text="Services grid — double-click any Cost / Rule Name / Mode / GL / Fund to edit", padding=6)
        tvf.pack(fill="both", expand=True)

        cols = ("pick", "service", "usage", "rule_name", "cost", "calc_type",
                "expense_account", "fund_account", "match_mode")
        self.tv = ttk.Treeview(tvf, columns=cols, show="headings", height=22)
        heads = [
            ("pick",            "☑",            40),
            ("service",         "Service / Item Description (exact text from invoice)", 420),
            ("usage",           "#Used",        70),
            ("rule_name",       "Rule Name (editable)", 240),
            ("cost",            "Cost AED (editable)", 110),
            ("calc_type",       "Calc Type",    170),
            ("expense_account", "Expense GL",   220),
            ("fund_account",    "Paid From",    160),
            ("match_mode",      "Match Mode",   170),
        ]
        for c, h, w in heads:
            self.tv.heading(c, text=h)
            self.tv.column(c, width=w, anchor="center" if c in ("pick", "usage") else "w")
        self.tv.tag_configure("has_cost", background="#F0FFF4")
        self.tv.tag_configure("no_cost",  background="#FFF5F5")
        self.tv.tag_configure("manual",   background="#EFF8FF")
        self.tv.bind("<Double-1>", self._on_dblclick)
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(tvf, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        # Row = internal dict list, render refreshes tree
        self.rows = []  # list[dict]
        self._next_manual_id = 1

        btns = ttk.Frame(root)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(
            btns, text="✅ 2) Bulk Add Ticked → Rules in this Cost Center",
            command=self._on_create_rules, style="Primary.TButton",
        ).pack(side="left")
        self.lbl_status = ttk.Label(btns, text="", foreground="#2B6CB0")
        self.lbl_status.pack(side="left", padx=20)
        ttk.Button(btns, text="Close", command=self.dlg.destroy).pack(side="right")

    # -------------------- service extraction -----------------------------
    def _all_invoices(self):
        if callable(self.get_all_invoices_cb):
            try:
                out = self.get_all_invoices_cb()
                if isinstance(out, list):
                    return [x if isinstance(x, dict) else
                            {k: getattr(x, k) for k in vars(x)} for x in out]
            except Exception:
                pass
        mgr = self.invoice_manager
        for attr in ("invoices", "get_all_invoices", "list_invoices"):
            try:
                val = getattr(mgr, attr, None)
            except Exception:
                val = None
            if callable(val):
                try:
                    out = val()
                except Exception:
                    continue
                if isinstance(out, list):
                    return [x if isinstance(x, dict) else
                            {k: getattr(x, k, None) for k in (getattr(x, '__dict__', {}) or {})} for x in out]
            if isinstance(val, list):
                return [x if isinstance(x, dict) else
                        {k: getattr(x, k, None) for k in (getattr(x, '__dict__', {}) or {})} for x in val]
        # JSON file fallback
        if self.data_folder:
            p = os.path.join(self.data_folder, "invoices_data.json")
            if os.path.exists(p):
                import json
                with open(p, "r") as fh:
                    try:
                        data = json.load(fh)
                        if isinstance(data, list):
                            return data
                        if isinstance(data, dict) and isinstance(data.get("invoices"), list):
                            return data["invoices"]
                    except Exception:
                        pass
        return []

    def _on_extract(self, silent=False):
        # Undo snapshot before merging (user can Undo a mistaken re-extract that merges)
        if hasattr(self, "rows") and self.rows:
            self._snapshot("Extract services (merge)")
        invs = self._all_invoices()
        from collections import Counter, defaultdict
        svc_counter = Counter()
        svc_sample_invoice = {}
        items_key_attempts = ("items", "line_items", "invoice_items", "services")
        for inv in invs:
            inv_id = str(inv.get("invoice_id") or "")
            items = []
            for ik in items_key_attempts:
                maybe = inv.get(ik)
                if isinstance(maybe, list) and maybe:
                    items = maybe; break
            for it in items:
                if not isinstance(it, dict):
                    continue
                desc = (it.get("description") or it.get("name") or it.get("service")
                        or it.get("product") or it.get("item") or "").strip()
                if not desc:
                    continue
                svc_counter[desc] += 1
                svc_sample_invoice.setdefault(desc, inv_id)

        # Merge into existing rows (don't wipe manual ones with empty usage)
        by_svc = {r["service"]: r for r in self.rows}
        added = 0
        for desc, count in svc_counter.most_common():
            if desc in by_svc:
                by_svc[desc]["usage"] = max(by_svc[desc].get("usage") or 0, count)
                by_svc[desc].setdefault("sample_inv", svc_sample_invoice.get(desc, ""))
                continue
            self.rows.append({
                "manual": False,
                "service": desc,
                "usage": count,
                "rule_name": desc[:60],
                "cost": None,
                "calc_type": "fixed_per_invoice",
                "expense_account": "5600  Medical / Clinical / Device Service Fees",
                "fund_account": "ADCB",
                "match_mode": "contains",
                "sample_inv": svc_sample_invoice.get(desc, ""),
            })
            added += 1

        total_services = len(svc_counter)
        msg = f"Scanned {len(invs)} invoices → found {total_services} unique service/item descriptions.  +{added} new rows."
        try:
            self._lbl_undo_status.config(text=msg, foreground="#2F855A")
        except Exception:
            pass
        self._render_rows()
        if not silent:
            messagebox.showinfo("Services Extracted", msg, parent=self.dlg)

    # -------------------- manual / row ops -------------------------------
    def _on_add_manual(self):
        dlg = tk.Toplevel(self.dlg)
        dlg.title("Add Service Manually")
        dlg.transient(self.dlg); dlg.grab_set(); dlg.geometry("500x220")
        frm = ttk.Frame(dlg, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Service description (as it appears/will appear on invoice lines):",
                  foreground="#2D3748").grid(row=0, column=0, columnspan=2, sticky="w")
        svc = ttk.Entry(frm, width=60); svc.grid(row=1, column=0, columnspan=2, pady=(4, 12), sticky="ew")
        ttk.Label(frm, text="Initial cost (AED):").grid(row=2, column=0, sticky="w")
        cost = ttk.Entry(frm, width=18); cost.grid(row=2, column=1, sticky="w")
        def ok():
            s = svc.get().strip()
            if not s:
                messagebox.showwarning("Missing", "Enter service description"); return
            try:
                c = float(cost.get()) if cost.get().strip() else None
                if c is not None and c < 0:
                    raise ValueError
            except Exception:
                messagebox.showwarning("Invalid", "Cost must be a number"); return
            self._snapshot(f"Add manual service: {s[:40]}")
            self.rows.append({
                "manual": True,
                "service": s,
                "usage": 0,
                "rule_name": s[:60],
                "cost": c,
                "calc_type": "fixed_per_invoice",
                "expense_account": "5600  Medical / Clinical / Device Service Fees",
                "fund_account": "ADCB",
                "match_mode": "contains",
                "sample_inv": "",
            })
            self._next_manual_id += 1
            self._render_rows()
            dlg.destroy()
        ttk.Button(frm, text="Add", command=ok).grid(row=3, column=0, pady=16)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=3, column=1, pady=16)

    def _on_delete_rows(self):
        kill = set(iid for iid in self.tv.selection())
        if not kill:
            messagebox.showinfo("Nothing selected", "Select one or more rows first."); return
        self._snapshot(f"Delete {len(kill)} row(s)")
        keep = []
        for idx, r in enumerate(self.rows):
            if str(idx) in kill:
                continue
            keep.append(r)
        self.rows = keep
        self._render_rows()
        self._lbl_undo_status.config(text=f"Deleted {len(kill)} row(s).  ↶Undo to bring them back.", foreground="#C53030")

    def _on_bulk_apply_cost(self):
        # ----------------------------------------------------------------- #
        # BUG FIX (user's "makes all same costs" complaint):                 #
        # If user has NOT selected any rows, DO NOT silently apply to every  #
        # service.  Instead:                                                   #
        #   • require explicit confirmation                                  #
        #   • tell them EXACTLY how many rows will get the same cost         #
        #   • let them cancel / undo immediately if it was a misclick        #
        # ----------------------------------------------------------------- #
        try:
            v = float(self.bulk_cost_var.get())
            if v < 0:
                raise ValueError
        except Exception:
            messagebox.showwarning("Invalid", "Enter a positive number in the bulk cost box"); return
        selected_iids = set(self.tv.selection())
        # Compute target rows — indices of rows that are currently visible in the tree
        visible_idx = [int(iid) for iid in self.tv.get_children()]
        if selected_iids:
            targets = [int(i) for i in selected_iids if i in visible_idx]
            scope = f"only {len(targets)} selected row(s)"
        else:
            # ==== ROOT CAUSE ==== — empty selection used to fall through and blast ALL rows
            targets = visible_idx  # every visible row (honours filter, though)
            scope = f"ALL {len(targets)} visible service row(s)"
            if len(targets) == 0:
                messagebox.showinfo("Nothing to do", "No services are loaded yet."); return
            if not messagebox.askyesno(
                "⚠️  About to assign the same cost to EVERY visible service",
                f"You didn't select any rows, so this will assign cost AED {v:,.2f} "
                f"to {scope}.\n\n"
                "If you only meant to assign it to specific services, click NO, "
                "select them first (Ctrl+click / Shift+click), then click Apply again.\n\n"
                "If it still goes wrong by accident, you can ↶Undo immediately in the top-right.",
                icon="warning",
                parent=self.dlg,
            ):
                return
        # Now snapshot BEFORE mutation, then apply
        self._snapshot(f"Bulk cost AED {v:,.2f} → {scope}")
        touched = 0
        for idx in targets:
            if idx < 0 or idx >= len(self.rows):
                continue
            self.rows[idx]["cost"] = v
            touched += 1
        if touched:
            self._render_rows()
            self._lbl_undo_status.config(
                text=f"Applied cost AED {v:,.2f} to {touched} row(s).  ↶Undo if this was a mistake.",
                foreground="#2F855A" if selected_iids else "#D69E2E",
            )

    def _render_rows(self):
        for i in self.tv.get_children():
            self.tv.delete(i)
        q = self.search_var.get().strip().lower()
        modes_label = {k: v for k, v in MATCH_MODES}
        calc_labels = {k: v for k, v in CALC_TYPES}
        for idx, r in enumerate(self.rows):
            if q and q not in r["service"].lower() and q not in (r.get("rule_name") or "").lower():
                continue
            tag = ("manual",) if r.get("manual") else (
                "has_cost" if r.get("cost") is not None else "no_cost",)
            self.tv.insert(
                "", "end", iid=str(idx), tags=tag,
                values=(
                    "☐",
                    r["service"],
                    r.get("usage") or 0,
                    r.get("rule_name") or r["service"][:60],
                    "" if r.get("cost") is None else f"{r['cost']:,.2f}",
                    calc_labels.get(r.get("calc_type"), r.get("calc_type")),
                    r.get("expense_account") or EXPENSE_OPTIONS[0],
                    r.get("fund_account") or "ADCB",
                    modes_label.get(r.get("match_mode"), r.get("match_mode")),
                ),
            )

    # -------------------- inline edit via double-click -------------------
    def _on_dblclick(self, event):
        iid = self.tv.identify_row(event.y)
        if not iid:
            return
        col = self.tv.identify_column(event.x)
        if not col:
            return
        try:
            col_idx = int(str(col).replace("#", "")) - 1
        except Exception:
            return
        col_name = self.tv["columns"][col_idx]
        try:
            ridx = int(iid)
            row = self.rows[ridx]
        except Exception:
            return
        if col_name == "pick":
            # Toggle pick
            cur = list(self.tv.item(iid, "values"))
            cur[0] = "☑" if cur[0] == "☐" else "☐"
            self.tv.item(iid, values=cur)
            return
        # Open appropriate editor
        dlg = tk.Toplevel(self.dlg)
        dlg.title("Edit — %s" % col_name)
        dlg.transient(self.dlg); dlg.grab_set(); dlg.geometry("420x200")
        frm = ttk.Frame(dlg, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Service: %s" % (row["service"][:80])).grid(row=0, column=0, columnspan=2, sticky="w")

        field_map = {
            "rule_name":       ("Rule Name", "text",  row.get("rule_name") or row["service"][:60]),
            "cost":            ("Cost AED", "number", "" if row.get("cost") is None else str(row.get("cost"))),
            "calc_type":       ("Calc Type", "calc_type", row.get("calc_type", "fixed_per_invoice")),
            "expense_account": ("Expense GL", "expense_account", row.get("expense_account") or EXPENSE_OPTIONS[0]),
            "fund_account":    ("Paid From (Fund)", "text", row.get("fund_account") or "ADCB"),
            "match_mode":      ("Match Mode", "match_mode", row.get("match_mode", "contains")),
            "service":         ("Service Description", "text", row["service"]),
        }
        if col_name not in field_map:
            return
        lab, kind, init = field_map[col_name]
        ttk.Label(frm, text=lab + ":").grid(row=1, column=0, sticky="w", pady=8)

        if kind == "text":
            w = ttk.Entry(frm, width=40); w.insert(0, init); w.grid(row=1, column=1, pady=8)
        elif kind == "number":
            w = ttk.Entry(frm, width=20); w.insert(0, init); w.grid(row=1, column=1, pady=8, sticky="w")
        elif kind == "calc_type":
            labs = [v for _, v in CALC_TYPES]
            back = {v: k for k, v in CALC_TYPES}
            w = ttk.Combobox(frm, values=labs, state="readonly", width=36)
            lab_fwd = dict(CALC_TYPES)
            w.set(lab_fwd.get(init, labs[0])); w.grid(row=1, column=1, pady=8)
        elif kind == "expense_account":
            w = ttk.Combobox(frm, values=EXPENSE_OPTIONS, state="readonly", width=36)
            w.set(init if init in EXPENSE_OPTIONS else EXPENSE_OPTIONS[0]); w.grid(row=1, column=1, pady=8)
        elif kind == "match_mode":
            labs = [v for _, v in MATCH_MODES]
            lab_fwd = dict(MATCH_MODES)
            w = ttk.Combobox(frm, values=labs, state="readonly", width=36)
            w.set(lab_fwd.get(init, labs[0])); w.grid(row=1, column=1, pady=8)
        else:
            dlg.destroy(); return

        def ok():
            # Snapshot BEFORE committing a single-cell edit (easy to Undo a typo)
            self._snapshot(f"Edit {col_name} for {row['service'][:30]}")
            if col_name == "cost":
                raw = w.get().strip()
                if raw == "":
                    row["cost"] = None
                else:
                    try:
                        row["cost"] = float(raw)
                        if row["cost"] < 0:
                            raise ValueError
                    except Exception:
                        # Rollback the snapshot we just pushed since nothing changed
                        if self._undo: self._undo.pop()
                        messagebox.showwarning("Invalid", "Cost must be a positive number or blank"); return
            elif col_name == "calc_type":
                back = {v: k for k, v in CALC_TYPES}
                row["calc_type"] = back.get(w.get(), "fixed_per_invoice")
            elif col_name == "expense_account":
                row["expense_account"] = w.get()
            elif col_name == "match_mode":
                back = {v: k for k, v in MATCH_MODES}
                row["match_mode"] = back.get(w.get(), "contains")
            else:
                row[col_name] = w.get().strip()
            self._render_rows()
            dlg.destroy()
            self._lbl_undo_status.config(text=f"Edited {col_name}.  ↶Undo to revert.", foreground="#2B6CB0")

        ttk.Button(frm, text="Save", command=ok).grid(row=2, column=0, pady=16)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=2, column=1, pady=16)

    # -------------------- create rules in center -------------------------
    def _on_create_rules(self):
        if not self.center_id:
            messagebox.showwarning("Missing", "Select a Cost Center first"); return
        # Collect rows where either box is ticked OR rows that have a cost and aren't yet 0-cost
        selected_indices = [iid for iid in self.tv.get_children()
                            if str(self.tv.item(iid, "values")[0]) == "☑"]
        if not selected_indices:
            if not messagebox.askyesno(
                "Confirm",
                "No rows ticked.  Create rules for EVERY row that already has a cost?",
                icon="question",
            ):
                return
            targets = [r for r in self.rows if r.get("cost") not in (None, "")]
        else:
            targets = [self.rows[int(i)] for i in selected_indices]
        # Validate
        bad = [r for r in targets if r.get("cost") in (None, "")]
        if bad:
            messagebox.showwarning(
                "Missing Cost",
                "The following services need a Cost AED before they become rules:\n  • "
                + "\n  • ".join(r["service"][:90] for r in bad[:12])
                + ("" if len(bad) <= 12 else "\n  … (%d more)" % (len(bad) - 12)),
            )
            return

        center = self.ccm.get_center(self.center_id)
        if not center:
            messagebox.showerror("Error", "Cost Center %s no longer exists" % self.center_id); return
        # Snapshot BEFORE pushing rules (so user can Undo the mass service cost/state change)
        self._snapshot(f"Submitting {len(targets)} service(s) → rules in '{center.get('name')}'")
        existing_rules = list(center.get("rules") or [])
        added = 0
        skipped = 0
        for r in targets:
            calc_type = r.get("calc_type") or "fixed_per_invoice"
            value = float(r.get("cost") or 0)
            if calc_type == "percentage":
                # Percentage — treat input 20 = 20%, so store 0.20 fraction
                try:
                    value_frac = value / 100.0 if value > 1 else value
                except Exception:
                    value_frac = 0.0
            else:
                value_frac = float(value)
            patterns_raw = [r["service"]]
            rule = {
                "rule_id": None,  # auto-generated by _sanitize_rules
                "name": (r.get("rule_name") or "").strip() or r["service"][:80],
                "match_mode": r.get("match_mode") or "contains",
                "patterns": patterns_raw,
                "calc_type": calc_type,
                "value": value_frac,
                "expense_account": _label_to_gl(r.get("expense_account") or EXPENSE_OPTIONS[0]),
                "fund_account": r.get("fund_account") or "ADCB",
            }
            # Skip duplicates (same name + service pattern already exists in center rules)
            def dup(ex):
                return (ex.get("name") == rule["name"] and
                        sorted(ex.get("patterns") or []) == sorted(patterns_raw) and
                        ex.get("calc_type") == rule["calc_type"])
            if any(dup(e) for e in existing_rules):
                skipped += 1
                continue
            existing_rules.append(rule)
            added += 1
        try:
            self.ccm.update_center(self.center_id, rules=existing_rules)
        except Exception as exc:
            # Roll back the snapshot we pushed
            if self._undo: self._undo.pop(); self._refresh_undo_buttons()
            messagebox.showerror("Error", f"Failed saving rules: {exc}"); return
        # SECOND snapshot AFTER success — this is the "last submitted cost" state.
        # User wants to be able to revert back to this after further edits.
        self._snapshot(f"✅ Last submitted: {added} rules → '{center.get('name')}' (skipped {skipped} duplicates)")
        # Clear redo because we've established a new canonical "submitted" baseline
        self._redo.clear(); self._refresh_undo_buttons()
        self._lbl_undo_status.config(
            text=f"✅ Done: +{added} rules added to '{center.get('name')}'.  This state saved as 'last submitted' — ↶Undo goes back to here.",
            foreground="#2F855A",
        )
        if callable(self.on_rules_created_cb):
            try:
                self.on_rules_created_cb(self.center_id)
            except Exception:
                pass
        messagebox.showinfo(
            "Rules Created",
            f"Added {added} rules into '{center.get('name')}'  (skipped {skipped} duplicates).\n\n"
            f"Now the easy part — go to **Bulk Backfill** window, pick this Cost Center,\n"
            f"click Preview, and all your past invoices that contain that service description\n"
            f"will automatically have the cost added.  Click Confirm & Apply to commit.\n\n"
            f"(Your service library state has been saved as the 'last submitted' baseline. "
            f"↶Undo to restore it at any time.)",
        )


# --------------------------------------------------------------------------- #
# Public entry point from the main dashboard button                           #
# --------------------------------------------------------------------------- #
def open_cost_center_and_backfill(parent, invoice_manager, data_folder: str,
                                  post_apply_refresh_cb=None,
                                  get_all_invoices_dict_cb=None):
    """
    Opens a combined hub window with two tabs:
      (A) Cost Center Manager
      (B) Bulk Backfill

    Extra optional args (from Professional Invoice Manager toolbar):
      post_apply_refresh_cb — callable, fired after Bulk Backfill applies changes,
                              used to refresh the Cost (AED) / P&L (AED) columns
                              in the Professional Invoice Manager list.
      get_all_invoices_dict_cb — callable → list[invoice_dict], used by
                              ServiceLibraryDialog to extract services.
    """
    if not _HAS_CCM:
        messagebox.showerror(
            "Missing Module",
            "CostCenterManager is required. File cost_center_manager.py not importable."
        )
        return

    hub = tk.Toplevel(parent)
    hub.title("💳 Cost Centers & Bulk Backfill")
    hub.transient(parent)
    hub.grab_set()
    _fit_window(hub, 1160, 700, remember_key="cost_center_hub")

    nb = ttk.Notebook(hub)
    nb.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    # --- Tab 1: Cost Center Manager --------------------------------------
    t1 = ttk.Frame(nb, padding=8); nb.add(t1, text="1) Manage Cost Centers & Rules")
    intro = ttk.Label(
        t1,
        text=(
            "Step 1 — Create a Cost Center (e.g. name it 'costs') and add rules per service\n"
            "Step 1a (recommended): Open 📚 Service Library — extract all services from past invoices OR type them manually,\n"
            "         assign a cost to each (e.g. MOH Registration = 500 AED), then one-click → rules\n"
            "Step 2 — Go to Bulk Backfill tab, Preview, Confirm → costs auto-added to EVERY matching past invoice."
        ),
        foreground="#2D3748",
        justify="left",
        wraplength=1000,
        padding=10,
    )
    intro.pack(fill="x")

    def open_mgr():
        ccm_dlg = CostCenterManagerDialog(
            hub, data_folder,
            invoice_manager=invoice_manager,
            get_all_invoices_dict_cb=get_all_invoices_dict_cb,
        )

    ttk.Button(t1, text="📋  Open Cost Center Manager + Service Library",
               command=open_mgr, style="Primary.TButton").pack(pady=10)

    # --- Tab 2: Bulk Backfill --------------------------------------------
    t2 = ttk.Frame(nb, padding=8); nb.add(t2, text="2) Bulk Backfill (Match services → add costs to ALL past invoices)")
    ttk.Label(
        t2,
        text=(
            "Select your Cost Center → Preview → Confirm & Apply.\n"
            "For every service rule you created (e.g. MOH Registration 500 AED), any past invoice that has a\n"
            "line-item DESCRIPTION containing that service will automatically have the cost added to its costs[]."
        ),
        foreground="#2D3748",
        justify="left",
        wraplength=1000,
        padding=10,
    ).pack(fill="x")

    def open_backfill():
        bb = BulkBackfillDialog(hub, invoice_manager, data_folder)
        # Hook post-apply refresh
        try:
            old_apply = bb._on_apply
            def wrapped_apply(*a, **kw):
                old_apply(*a, **kw)
                if callable(post_apply_refresh_cb):
                    try:
                        post_apply_refresh_cb()
                    except Exception:
                        pass
            bb._on_apply = wrapped_apply
            # Wire get_all_invoices_dict_cb for filter extraction (already scans manager too)
            if get_all_invoices_dict_cb is not None:
                bb._get_all_invoices_override = get_all_invoices_dict_cb
        except Exception:
            pass

    ttk.Button(t2, text="🚀  Open Bulk Backfill",
               command=open_backfill, style="Primary.TButton").pack(pady=10)

    ttk.Label(
        hub,
        text="ℹ Open Manager first (set up service-cost rules in library).  Then Bulk Backfill applies to all.",
        foreground="#718096",
    ).pack(pady=(0, 10))

    ttk.Button(hub, text="Close Hub", command=hub.destroy).pack(pady=(0, 10))
