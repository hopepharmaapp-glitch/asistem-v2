
# balance_manager.py
"""
Single source of truth for account balances and financial transactions.
- Stores accounts in balances.json
- Stores transactions in transactions.json
- Exposes high-level APIs used by the app & invoice manager:
    * process_invoice_payment(invoice_data, payment_amount, account_name)
    * reverse_invoice_payment(invoice_data, payment_amount, account_name)
    * record_purchase(purchase_dict)   # deducts from account and records VAT if present
    * update_balance(account_key, amount, tx_type, description, meta=None)
- Provides an optional lightweight UI via show_balance_manager(...), kept for compatibility.
"""
from __future__ import annotations

import json
import os
from datetime import datetime
from typing import Dict, Any, List, Tuple, Optional
import ui_undo

# ---------------------------
# Core storage & data schema
# ---------------------------

class BalanceManager:
    def __init__(self, data_folder: str):
        self.data_folder = data_folder
        os.makedirs(self.data_folder, exist_ok=True)
        self.balance_file = os.path.join(data_folder, "balances.json")
        self.transaction_file = os.path.join(data_folder, "transactions.json")

        self.accounts: Dict[str, Dict[str, Any]] = self._load_json(self.balance_file, default={})
        self.transactions: List[Dict[str, Any]] = self._load_json(self.transaction_file, default=[])

        # Ensure a useful default chart of accounts (keys can be codes or names)
        if not self.accounts:
            self.accounts = {
                "100": {"name": "Cash", "type": "asset", "currency": "AED", "balance": 0.0},
                "101": {"name": "Bank", "type": "asset", "currency": "AED", "balance": 0.0},
                "110": {"name": "Accounts Receivable", "type": "asset", "currency": "AED", "balance": 0.0},
                "200": {"name": "Accounts Payable", "type": "liability", "currency": "AED", "balance": 0.0},
                "300": {"name": "Sales Revenue", "type": "income", "currency": "AED", "balance": 0.0},
                "400": {"name": "Purchases/COGS", "type": "expense", "currency": "AED", "balance": 0.0},
                "202": {"name": "VAT Payable", "type": "liability", "currency": "AED", "balance": 0.0},
            }
            self._save_json(self.balance_file, self.accounts)

    # ---------- JSON helpers ----------
    def _load_json(self, path: str, default):
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return default
        return default

    def _save_json(self, path: str, data):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def _now(self) -> str:
        return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def add_transaction(self, account_name: str, amount: float, transaction_type: str, description: str, reference_id: Optional[str] = None, reference_type: Optional[str] = None, date: Optional[str] = None) -> Tuple[bool, str]:
        if account_name not in self.accounts:
            self.add_account(account_name, account_name, acc_type="asset", opening_balance=0.0)
        amt = round(float(amount), 2)
        if amt <= 0:
            return False, "Amount must be > 0"
        if transaction_type not in ("deposit","withdrawal"):
            return False, "Invalid transaction type"
        sign = 1 if transaction_type == "deposit" else -1
        new_bal = round(self.accounts[account_name]["balance"] + sign * amt, 2)
        if self.accounts[account_name].get("type") == "asset" and new_bal < 0:
            return False, "Insufficient funds"
        self.accounts[account_name]["balance"] = new_bal
        self.accounts[account_name]["last_updated"] = self._now()
        txn = {
            "transaction_id": f"TXN-{datetime.now().strftime('%Y%m%d')}-{len(self.transactions)+1:03d}",
            "timestamp": self._now(),
            "date": date or datetime.now().strftime("%Y-%m-%d"),
            "account_name": account_name,
            "amount": amt,
            "transaction_type": transaction_type,
            "description": description,
            "reference_type": reference_type,
            "reference_id": reference_id,
            "status": "completed",
        }
        self.transactions.append(txn)
        self._save_json(self.balance_file, self.accounts)
        self._save_json(self.transaction_file, self.transactions)
        return True, "Transaction recorded"

    # ---------- Account operations ----------
    def get_account_names(self) -> List[str]:
        """Return list of keys (codes or names) present in accounts."""
        return list(self.accounts.keys())

    def get_account_balance(self, account_name: str) -> float:
        acc = self.accounts.get(account_name)
        return float(acc.get("balance", 0.0)) if acc else 0.0

    def get_all_accounts(self) -> List[Dict[str, Any]]:
        out = []
        for name, acc in self.accounts.items():
            out.append({
                "account_name": name,
                "current_balance": acc.get("balance", 0.0),
                "currency": acc.get("currency", "AED"),
                "account_type": acc.get("type", "Bank"),
                "created_date": acc.get("created_date") or datetime.now().strftime("%Y-%m-%d"),
                "last_updated": acc.get("last_updated") or self._now(),
            })
        return out

    def create_account(self, account_name: str, initial_balance: float = 0.0, currency: str = "AED", account_type: str = "Bank") -> Tuple[bool, str]:
        raw = (account_type or "bank").lower()
        if raw in ("bank", "cash", "credit", "credit card"):
            acc_type_norm = "asset"
        elif raw in ("ap", "accounts payable", "liability"):
            acc_type_norm = "liability"
        elif raw in ("income", "revenue"):
            acc_type_norm = "income"
        elif raw in ("expense", "cogs"):
            acc_type_norm = "expense"
        else:
            acc_type_norm = raw
        return self.add_account(account_name, account_name, acc_type=acc_type_norm, currency=currency, opening_balance=initial_balance)

    def add_account(self, key: str, name: str, acc_type: str = "asset", currency: str = "AED", opening_balance: float = 0.0):
        if key in self.accounts:
            return False, f"Account '{key}' already exists"
        self.accounts[key] = {"name": name, "type": acc_type, "currency": currency, "balance": float(opening_balance)}
        self._save_json(self.balance_file, self.accounts)
        if opening_balance:
            txn = {
                "transaction_id": f"TXN-{datetime.now().strftime('%Y%m%d')}-{len(self.transactions)+1:03d}",
                "timestamp": self._now(),
                "date": datetime.now().strftime("%Y-%m-%d"),
                "account_name": key,
                "amount": round(abs(float(opening_balance)), 2),
                "transaction_type": "deposit" if opening_balance > 0 else "withdrawal",
                "description": f"Opening balance for {key} - {name}",
                "reference_type": "opening_balance",
                "reference_id": None,
                "status": "completed",
            }
            self.transactions.append(txn)
            self._save_json(self.transaction_file, self.transactions)
        return True, "Account created"

    def delete_account(self, key: str, force: bool = False, transfer_to: Optional[str] = None) -> Tuple[bool, str]:
        if key not in self.accounts:
            return False, "Account not found"
        bal = float(self.accounts[key].get("balance", 0.0))
        has_history = any(t.get("account_name") == key and t.get("status") == "completed" for t in self.transactions)
        if (abs(bal) > 1e-9 or has_history) and not force:
            return False, "Account has balance or history; use force delete"
        if abs(bal) > 1e-9:
            target = transfer_to or "Suspense"
            if target not in self.accounts:
                self.add_account(target, target, acc_type="asset", opening_balance=0.0)
            if bal > 0:
                self.add_transaction(key, bal, "withdrawal", f"Force delete transfer to {target}", reference_type="account_delete")
                self.add_transaction(target, bal, "deposit", f"Force delete transfer from {key}", reference_type="account_delete")
            else:
                amt = abs(bal)
                self.add_transaction(key, amt, "deposit", f"Force delete transfer from {target}", reference_type="account_delete")
                self.add_transaction(target, amt, "withdrawal", f"Force delete transfer to {key}", reference_type="account_delete")
        del self.accounts[key]
        self._save_json(self.balance_file, self.accounts)
        return True, "Account deleted"

    def adjust_balance(self, account_name: str, new_balance: float, reason: str = "Manual adjustment") -> Tuple[bool, str]:
        if account_name not in self.accounts:
            return False, "Account not found"
        current = float(self.accounts[account_name].get("balance", 0.0))
        delta = round(float(new_balance) - current, 2)
        if abs(delta) < 1e-9:
            return True, "No change"
        tx_type = "deposit" if delta > 0 else "withdrawal"
        return self.add_transaction(account_name, abs(delta), tx_type, reason, reference_id=None, reference_type="adjustment")

    def update_balance(self, key: str, amount: float, tx_type: str, description: str, meta: Optional[Dict[str, Any]] = None) -> Tuple[bool, str]:
        """
        Update account balance and record a simple cash movement transaction.
        tx_type: 'deposit' (increase asset) or 'withdrawal' (decrease asset)
        """
        if key not in self.accounts:
            return False, f"Account '{key}' not found"

        amount = float(amount)
        if tx_type not in ("deposit", "withdrawal"):
            return False, "tx_type must be 'deposit' or 'withdrawal'"

        sign = 1 if tx_type == "deposit" else -1
        new_bal = round(self.accounts[key]["balance"] + sign * amount, 2)
        if self.accounts[key].get("type") == "asset" and new_bal < 0:
            return False, "Insufficient funds"
        self.accounts[key]["balance"] = new_bal
        self.accounts[key]["last_updated"] = self._now()
        # Append non-double-entry transaction record
        txn = {
            "transaction_id": f"TXN-{datetime.now().strftime('%Y%m%d')}-{len(self.transactions)+1:03d}",
            "timestamp": self._now(),
            "date": datetime.now().strftime("%Y-%m-%d"),
            "account_name": key,
            "amount": round(amount, 2),
            "transaction_type": tx_type,
            "description": description,
            "reference_type": (meta or {}).get("kind"),
            "reference_id": None,
            "status": "completed",
        }
        self.transactions.append(txn)
        self._save_json(self.balance_file, self.accounts)
        self._save_json(self.transaction_file, self.transactions)
        return True, "Balance updated"

    # ---------- Transaction operations ----------
    def record_transaction(
        self,
        date: str,
        description: str,
        amount: float,
        debit_account: Optional[str],
        credit_account: Optional[str],
        reference: Optional[str] = None,
        meta: Optional[Dict[str, Any]] = None,
    ):
        """Append a double-entry style transaction and update balances for asset/liability types simply."""
        tx = {
            "transaction_id": f"TX{len(self.transactions)+1:06d}",
            "date": date,
            "description": description,
            "amount": round(float(amount), 2),
            "debit_account": debit_account,
            "credit_account": credit_account,
            "reference": reference,
            "meta": meta or {},
            "created_at": self._now(),
        }
        self.transactions.append(tx)

        # Update balances crudely for common accounts (assets/liabilities)
        def apply_delta(acc_key: str, delta: float):
            if acc_key and acc_key in self.accounts:
                self.accounts[acc_key]["balance"] = round(self.accounts[acc_key]["balance"] + delta, 2)

        # For assets (e.g., cash/bank/AR): debit increases, credit decreases
        # For liabilities (e.g., AP/VAT): debit decreases, credit increases
        def is_liability(acc_key: str) -> bool:
            return self.accounts.get(acc_key, {}).get("type") == "liability"

        if debit_account:
            apply_delta(debit_account, -float(amount) if is_liability(debit_account) else float(amount))
        if credit_account:
            apply_delta(credit_account, float(amount) if is_liability(credit_account) else -float(amount))

        self._save_json(self.transaction_file, self.transactions)
        self._save_json(self.balance_file, self.accounts)

    # ---------- High-level domain ops ----------
    def process_invoice_payment(self, invoice_data: Dict[str, Any], payment_amount: float, account_key: str) -> Tuple[bool, str]:
        if account_key not in self.accounts:
            self.add_account(account_key, account_key, acc_type="asset", opening_balance=0.0)
        amount = round(float(payment_amount), 2)
        inv_id = invoice_data.get("invoice_id") or invoice_data.get("id") or "UNKNOWN"
        desc = f"Invoice payment {inv_id}"
        ok, msg = self.update_balance(account_key, amount, "deposit", desc, meta={"kind": "invoice"})
        if not ok:
            return False, msg
        return True, "Payment recorded"

    def reverse_invoice_payment(self, invoice_data: Dict[str, Any], payment_amount: float, account_key: str) -> Tuple[bool, str]:
        if account_key not in self.accounts:
            self.add_account(account_key, account_key, acc_type="asset", opening_balance=0.0)
        amount = round(float(payment_amount), 2)
        inv_id = invoice_data.get("invoice_id") or invoice_data.get("id") or "UNKNOWN"
        desc = f"Reverse payment {inv_id}"
        ok, msg = self.update_balance(account_key, amount, "withdrawal", desc, meta={"kind": "invoice"})
        if not ok:
            return False, msg
        return True, "Payment reversed"

    def record_purchase(self, purchase: Dict[str, Any], pay_from_account: str) -> Tuple[bool, str]:
        """
        Record a purchase (inventory/expense) and deduct from the chosen account.
        purchase fields expected: {purchase_id, description, amount, date, category, vat_rate(optional)}
        """
        if pay_from_account not in self.accounts:
            self.add_account(pay_from_account, pay_from_account, acc_type="asset", opening_balance=0.0)

        amount = round(float(purchase.get("amount", 0)), 2)
        if amount <= 0:
            return False, "Purchase amount must be > 0"

        date = purchase.get("date") or datetime.now().strftime("%Y-%m-%d")
        pid = purchase.get("purchase_id") or f"PUR-{int(datetime.now().timestamp())}"
        desc = purchase.get("description") or "Purchase"

        expense_key = "400" if "400" in self.accounts else self._find_account_by_name_fuzzy("purchases") or pay_from_account
        vat_rate = float(purchase.get("vat_rate", 0.0))
        vat_amount = round(amount * vat_rate, 2) if vat_rate > 0 else 0.0

        # Split: net + VAT (liability)
        net_amount = amount - vat_amount if vat_amount else amount

        if net_amount:
            ok_net, msg_net = self.update_balance(pay_from_account, net_amount, "withdrawal", f"Purchase {pid} - {desc}", meta={"kind": "purchase"})
            if not ok_net:
                return False, msg_net
        # VAT (debit VAT Paid asset? or reduce VAT payable). For simplicity, reduce VAT Payable
        if vat_amount:
            vat_key = "202" if "202" in self.accounts else self._find_account_by_name_fuzzy("vat")
            if vat_key:
                ok_v, msg_v = self.update_balance(pay_from_account, vat_amount, "withdrawal", f"VAT on Purchase {pid}", meta={"kind": "purchase_vat"})
                if not ok_v:
                    return False, msg_v
            else:
                # If no VAT account configured, include full amount as expense
                pass

        return True, "Purchase recorded"

    def process_purchase(self, purchase_data: Dict[str, Any]) -> Tuple[bool, str]:
        acc = purchase_data.get("account")
        amt = float(purchase_data.get("amount", 0) or 0)
        pid = purchase_data.get("purchase_id") or f"PUR-{int(datetime.now().timestamp())}"
        desc = purchase_data.get("description") or "Purchase"
        date = purchase_data.get("date") or datetime.now().strftime("%Y-%m-%d")
        for t in self.transactions:
            if t.get("reference_type") == "purchase" and t.get("reference_id") == pid and t.get("status") == "completed":
                return False, "Duplicate purchase"
        if acc not in self.accounts:
            self.add_account(acc, acc, acc_type="asset", opening_balance=0.0)
        ok, msg = self.update_balance(acc, amt, "withdrawal", desc, meta={"kind": "purchase"})
        return ok, msg

    def reverse_purchase(self, purchase_id: str, purchase_data: Dict[str, Any]) -> Tuple[bool, str]:
        acc = purchase_data.get("account")
        amt = float(purchase_data.get("amount", 0) or 0)
        date = purchase_data.get("date") or datetime.now().strftime("%Y-%m-%d")
        if acc not in self.accounts:
            self.add_account(acc, acc, acc_type="asset", opening_balance=0.0)
        ok, msg = self.update_balance(acc, amt, "deposit", f"Reverse {purchase_id}", meta={"kind": "purchase"})
        if not ok:
            return False, msg
        for t in self.transactions:
            if t.get("reference_type") == "purchase" and t.get("reference_id") == purchase_id and t.get("status") == "completed":
                t["status"] = "reversed"
                break
        self._save_json(self.transaction_file, self.transactions)
        return True, "Purchase reversed"

    # ---------- helpers ----------
    def _find_account_by_name_fuzzy(self, needle: str) -> Optional[str]:
        needle = needle.lower()
        for key, acc in self.accounts.items():
            name = str(acc.get("name", "")).lower()
            if needle in name:
                return key
        return None

# ---------------------------
# Optional minimal UI hooks
# ---------------------------
# We keep the UI entrypoint for compatibility with the rest of the app.
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, simpledialog

    class BalanceManagerDialog:
        def __init__(self, parent, invoice_manager, data_folder=None):
            self.parent = parent
            self.invoice_manager = invoice_manager
            # Accept either the app instance or the data manager
            self.data_folder = data_folder or getattr(invoice_manager, "invoice_folder", None)
            if not self.data_folder and hasattr(invoice_manager, 'manager'):
                self.data_folder = getattr(invoice_manager.manager, 'invoice_folder', os.getcwd())
            if not self.data_folder:
                self.data_folder = os.getcwd()
            self.manager = BalanceManager(self.data_folder)
            self._admin_ok = bool(getattr(invoice_manager, 'user_role', '') == 'Admin')
            self._build()

        def _build(self):
            self.dialog = tk.Toplevel(self.parent)
            self.dialog.title("Balance Manager")
            self.dialog.geometry("700x520")
            self.dialog.grab_set()

            # Accounts table
            frm = ttk.Frame(self.dialog, padding=10)
            frm.pack(fill="both", expand=True)

            ttk.Label(frm, text="Accounts", font=("Arial", 11, "bold")).pack(anchor="w", pady=(0,6))
            cols = ("key","name","type","currency","balance")
            self.tree = ttk.Treeview(frm, columns=cols, show="headings", height=8)
            for c in cols:
                self.tree.heading(c, text=c.title())
                self.tree.column(c, width=120, anchor="center")
            self.tree.pack(fill="x")

            self._refresh_accounts()

            # Buttons
            btns = ttk.Frame(frm)
            btns.pack(fill="x", pady=8)
            ttk.Button(btns, text="Add Account", command=self._add_account).pack(side="left", padx=4)
            ttk.Button(btns, text="Delete Account", command=self._delete_account).pack(side="left", padx=4)
            ttk.Button(btns, text="Edit Balance", command=self._edit_balance).pack(side="left", padx=4)
            ttk.Button(btns, text="Refresh", command=self._refresh_all).pack(side="left", padx=4)
            ttk.Button(btns, text="Close", command=self.dialog.destroy).pack(side="right", padx=4)

            # Transactions list (brief)
            ttk.Label(frm, text="Recent Transactions", font=("Arial", 11, "bold")).pack(anchor="w", pady=(10,6))
            self.tx_list = tk.Listbox(frm, height=10)
            self.tx_list.pack(fill="both", expand=True)
            self._refresh_transactions()

        def _refresh_accounts(self):
            for i in self.tree.get_children():
                self.tree.delete(i)
            for key, acc in self.manager.accounts.items():
                self.tree.insert("", "end", values=(key, acc.get("name"), acc.get("type"), acc.get("currency"), f"{acc.get('balance',0):,.2f}"))
            # Auto-select first row for convenience
            try:
                first = self.tree.get_children()[0]
                self.tree.selection_set(first)
            except Exception:
                pass

        def _selected_account_key(self):
            sel = self.tree.selection()
            if not sel:
                # Fallback: select first row
                try:
                    first = self.tree.get_children()[0]
                    self.tree.selection_set(first)
                    sel = self.tree.selection()
                except Exception:
                    return None
            vals = self.tree.item(sel[0]).get('values', [])
            return vals[0] if vals else None

        def _refresh_transactions(self):
            self.tx_list.delete(0, "end")
            for tx in self.manager.transactions[-50:]:
                try:
                    date = tx.get('date', '')
                    tid = tx.get('transaction_id') or tx.get('id') or ''
                    desc = tx.get('description', '')
                    amt = float(tx.get('amount', 0) or 0)
                    self.tx_list.insert("end", f"{date} {tid} | {desc} | AED {amt:,.2f}")
                except Exception:
                    continue

        def _refresh_all(self):
            self._refresh_accounts()
            self._refresh_transactions()

        def _add_account(self):
            if not self._require_admin():
                return
            dlg = simpledialog.askstring("New Account", "Enter account key (e.g., 105):", parent=self.dialog)
            if not dlg:
                return
            name = simpledialog.askstring("New Account", "Enter account name:", parent=self.dialog) or "Account"
            ok, msg = self.manager.add_account(dlg.strip(), name.strip())
            messagebox.showinfo("Result", msg, parent=self.dialog)
            self._refresh_accounts()

        def _delete_account(self):
            if not self._require_admin():
                return
            key = self._selected_account_key()
            if not key:
                messagebox.showwarning("Warning", "Please select an account in the table", parent=self.dialog)
                return
            ok, msg = self.manager.delete_account(str(key).strip())
            if ok:
                messagebox.showinfo("Result", msg, parent=self.dialog)
                self._refresh_accounts()
                self._refresh_transactions()
            else:
                # Automatically force delete if blocked by balance/history
                ok2, msg2 = self.manager.delete_account(str(key).strip(), force=True, transfer_to="Suspense")
                if ok2:
                    messagebox.showinfo("Result", f"{msg2} (forced)", parent=self.dialog)
                    self._refresh_accounts()
                    self._refresh_transactions()
                else:
                    messagebox.showerror("Error", msg, parent=self.dialog)

        def _edit_balance(self):
            if not self._require_admin():
                return
            key = self._selected_account_key()
            if not key:
                messagebox.showwarning("Warning", "Please select an account in the table", parent=self.dialog)
                return
            new_bal_str = simpledialog.askstring("Edit Balance", "Enter new balance (AED):", parent=self.dialog)
            if not new_bal_str:
                return
            try:
                new_bal = float(new_bal_str)
            except Exception:
                messagebox.showerror("Error", "Invalid amount", parent=self.dialog)
                return
            ok, msg = self.manager.adjust_balance(str(key).strip(), new_bal, reason="Admin adjustment")
            if ok:
                messagebox.showinfo("Result", msg, parent=self.dialog)
                self._refresh_accounts()
                self._refresh_transactions()
            else:
                messagebox.showerror("Error", msg, parent=self.dialog)

        def _require_admin(self) -> bool:
            # Honor app role if already admin
            if getattr(self.invoice_manager, 'user_role', None) == 'Admin' or self._admin_ok:
                return True
            # Use app's admin login if available
            if hasattr(self.invoice_manager, 'ensure_admin_then'):
                def _mark_admin():
                    self._admin_ok = True
                try:
                    self.invoice_manager.ensure_admin_then(_mark_admin)
                except Exception:
                    self._admin_ok = False
                return bool(self._admin_ok)
            # Staff cannot proceed
            try:
                messagebox.showerror("Error", "Can only be accessed by admin")
            except Exception:
                pass
            return False

    class BalancePasswordDialog:
        def __init__(self, parent):
            self.dialog = tk.Toplevel(parent)
            self.dialog.title("Enter Password")
            self.dialog.geometry("300x120")
            self.dialog.grab_set()
            ttk.Label(self.dialog, text="Password").pack(pady=8)
            self.entry = ttk.Entry(self.dialog, show="*")
            self.entry.pack(padx=10, pady=4, fill="x")
            ttk.Button(self.dialog, text="OK", command=self._ok).pack(pady=8)
            self.result = False

        def _ok(self):
            try:
                pwd = self.entry.get().strip()
                self.result = (pwd == 'admin')
            except Exception:
                self.result = False
            finally:
                self.dialog.destroy()

    def show_balance_manager(parent, invoice_manager):
        try:
            if getattr(invoice_manager, 'user_role', 'Staff') == 'Admin':
                BalanceManagerDialog(parent, invoice_manager)
            else:
                messagebox.showerror("Error", "Can only be accessed by admin")
        except Exception:
            messagebox.showerror("Error", "Can only be accessed by admin")

except Exception:
    # If Tk isn't available (e.g., headless), ignore UI portion gracefully.
    pass
    def get_transactions(self, account_name: Optional[str] = None, start_date: Optional[str] = None, end_date: Optional[str] = None, type: Optional[str] = None, reference_id: Optional[str] = None, reference_type: Optional[str] = None) -> List[Dict[str, Any]]:
        out = []
        for t in self.transactions:
            if account_name and t.get("account_name") != account_name:
                continue
            if type and t.get("transaction_type") != type:
                continue
            if reference_id and t.get("reference_id") != reference_id:
                continue
            if reference_type and t.get("reference_type") != reference_type:
                continue
            if start_date and t.get("date") < start_date:
                continue
            if end_date and t.get("date") > end_date:
                continue
            out.append(t)
        return out

    def get_account_statement(self, account_name: str, start_date: str, end_date: str) -> Dict[str, Any]:
        bal = float(self.accounts.get(account_name, {}).get("balance", 0.0))
        txs = self.get_transactions(account_name=account_name, start_date=start_date, end_date=end_date)
        deposits = sum(t.get("amount", 0.0) for t in txs if t.get("transaction_type") == "deposit")
        withdrawals = sum(t.get("amount", 0.0) for t in txs if t.get("transaction_type") == "withdrawal")
        # Approximate opening by reversing movements within the range
        opening = bal
        for t in txs:
            if t.get("transaction_type") == "deposit":
                opening -= t.get("amount", 0.0)
            else:
                opening += t.get("amount", 0.0)
        return {
            "account_name": account_name,
            "start_date": start_date,
            "end_date": end_date,
            "opening_balance": round(opening, 2),
            "closing_balance": round(bal, 2),
            "deposits": round(deposits, 2),
            "withdrawals": round(withdrawals, 2),
            "transactions": txs,
        }

    def save_data(self):
        self._save_json(self.balance_file, self.accounts)
        self._save_json(self.transaction_file, self.transactions)

    def load_data(self):
        self.accounts = self._load_json(self.balance_file, default={})
        self.transactions = self._load_json(self.transaction_file, default=[])
