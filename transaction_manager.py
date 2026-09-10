# transaction_manager.py
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import json
import os
import shutil
from datetime import datetime, timedelta
from balance_manager import BalanceManager
import ui_undo

class TransactionManager:
    def __init__(self, data_folder, invoice_manager=None):
        self.data_folder = data_folder
        self.invoice_manager = invoice_manager  # Reference to main app's invoice manager
        self.transactions_file = os.path.join(data_folder, "financial_transactions.json")
        self.accounts_file = os.path.join(data_folder, "financial_accounts.json")
        self.transactions = self.load_transactions()
        self.accounts = self.load_accounts()
        self.setup_default_accounts()
        self.balance_manager = BalanceManager(getattr(invoice_manager, 'invoice_folder', data_folder) if invoice_manager else data_folder)
        self.sync_with_invoice_data()  # Sync with existing data
    
    def sync_with_invoice_data(self):
        """Sync transaction data with existing invoices and purchases"""
        if not self.invoice_manager:
            return
            
        try:
            # Sync invoices as revenue transactions
            invoices = self.invoice_manager.get_all_invoices_dict()
            for invoice in invoices:
                self.sync_invoice_transaction(invoice)
            
            # Sync purchases as expense transactions
            purchases = self.invoice_manager.get_all_purchases_dict()
            for purchase in purchases:
                self.sync_purchase_transaction(purchase)
                
            self.save_transactions()
            self.save_accounts()
            
        except Exception as e:
            print(f"Warning: Could not sync with invoice data: {e}")
    
    def sync_invoice_transaction(self, invoice):
        invoice_id = invoice.get('invoice_id')
        existing_trx = [t for t in self.transactions if t.get('reference') == invoice_id]
        if existing_trx:
            return
        revenue_account = "401" if invoice.get('invoice_type') == 'sales' else "402"
        amount = float(invoice.get('grand_total', 0))
        date_str = invoice.get('date', datetime.now().strftime("%Y-%m-%d"))
        self.balance_manager.record_transaction(
            date=date_str,
            description=f"Invoice {invoice_id} - {invoice.get('client_name', '')}",
            amount=amount,
            debit_account="104",
            credit_account=revenue_account,
            reference=invoice_id,
            meta={"kind": "invoice"}
        )
        self.transactions.append({
            'transaction_id': f"INV{len(self.transactions) + 1:06d}",
            'date': date_str,
            'description': f"Invoice {invoice_id} - {invoice.get('client_name', '')}",
            'debit_account': "104",
            'credit_account': revenue_account,
            'amount': amount,
            'vat_amount': float(invoice.get('tax_amount', 0)),
            'reference': invoice_id,
            'type': 'invoice',
            'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        if invoice.get('tax_amount', 0) > 0:
            self.handle_vat_transaction(date_str, f"VAT - Invoice {invoice_id}", float(invoice.get('tax_amount', 0)), "104")
        total_paid = invoice.get('total_paid', 0)
        if total_paid > 0:
            self.sync_invoice_payment(invoice_id, invoice.get('client_name', ''), total_paid)
    
    def sync_invoice_payment(self, invoice_id, client_name, amount):
        date_str = datetime.now().strftime("%Y-%m-%d")
        self.balance_manager.record_transaction(
            date=date_str,
            description=f"Payment received - {client_name}",
            amount=float(amount),
            debit_account="102",
            credit_account="104",
            reference=f"PAY-{invoice_id}",
            meta={"kind": "payment"}
        )
        self.transactions.append({
            'transaction_id': f"PMT{len(self.transactions) + 1:06d}",
            'date': date_str,
            'description': f"Payment received - {client_name}",
            'debit_account': "102",
            'credit_account': "104",
            'amount': float(amount),
            'vat_amount': 0,
            'reference': f"PAY-{invoice_id}",
            'type': 'payment',
            'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
    
    def sync_purchase_transaction(self, purchase):
        purchase_id = purchase.get('purchase_id')
        existing_trx = [t for t in self.transactions if t.get('reference') == purchase_id]
        if existing_trx:
            return
        date_str = purchase.get('date', datetime.now().strftime("%Y-%m-%d"))
        category = purchase.get('category', 'Other')
        expense_account = self.map_category_to_account(category)
        amount = float(purchase.get('amount', 0))
        self.balance_manager.record_transaction(
            date=date_str,
            description=f"Purchase - {purchase.get('description', '')}",
            amount=amount,
            debit_account=expense_account,
            credit_account=purchase.get('account') or "101",
            reference=purchase_id,
            meta={"kind": "purchase", "category": category}
        )
        vat_amount = amount * 0.05
        self.transactions.append({
            'transaction_id': f"PUR{len(self.transactions) + 1:06d}",
            'date': date_str,
            'description': f"Purchase - {purchase.get('description', '')}",
            'debit_account': expense_account,
            'credit_account': purchase.get('account') or "101",
            'amount': amount,
            'vat_amount': vat_amount,
            'reference': purchase_id,
            'type': 'purchase',
            'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        if vat_amount > 0:
            self.handle_vat_transaction(date_str, f"VAT - Purchase {purchase_id}", vat_amount, expense_account)
    
    def map_category_to_account(self, category):
        """Map purchase category to appropriate expense account"""
        category_mapping = {
            "Office Supplies": "505",
            "Equipment": "505",
            "Software": "505",
            "Services": "508",
            "Travel": "507",
            "Marketing": "506",
            "Utilities": "504",
            "Other": "505"
        }
        return category_mapping.get(category, "505")  # Default to Office Supplies
    
    def load_transactions(self):
        """Load transactions from JSON file"""
        if os.path.exists(self.transactions_file):
            try:
                with open(self.transactions_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading transactions: {e}")
                return []
        return []
    
    def load_accounts(self):
        """Load accounts from JSON file"""
        if os.path.exists(self.accounts_file):
            try:
                with open(self.accounts_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading accounts: {e}")
                return {}
        return {}
    
    def setup_default_accounts(self):
        """Setup default chart of accounts"""
        default_accounts = {
            "assets": {
                "101": {"name": "Cash on Hand", "type": "current_asset", "balance": 0},
                "102": {"name": "Emirates NBD", "type": "bank", "balance": 0},
                "103": {"name": "ADCB", "type": "bank", "balance": 0},
                "104": {"name": "Accounts Receivable", "type": "current_asset", "balance": 0},
                "105": {"name": "VAT Receivable", "type": "current_asset", "balance": 0}
            },
            "liabilities": {
                "201": {"name": "Accounts Payable", "type": "current_liability", "balance": 0},
                "202": {"name": "VAT Payable", "type": "current_liability", "balance": 0}
            },
            "equity": {
                "301": {"name": "Share Capital", "type": "equity", "balance": 300000},
                "302": {"name": "Retained Earnings", "type": "equity", "balance": 0}
            },
            "revenue": {
                "401": {"name": "Medicine Sales", "type": "revenue", "balance": 0},
                "402": {"name": "Service Revenue", "type": "revenue", "balance": 0}
            },
            "expenses": {
                "501": {"name": "Cost of Goods Sold", "type": "expense", "balance": 0},
                "502": {"name": "Salaries & Wages", "type": "expense", "balance": 0},
                "503": {"name": "Rent Expense", "type": "expense", "balance": 0},
                "504": {"name": "Utilities Expense", "type": "expense", "balance": 0},
                "505": {"name": "Office Supplies", "type": "expense", "balance": 0},
                "506": {"name": "Marketing Expense", "type": "expense", "balance": 0},
                "507": {"name": "Travel Expense", "type": "expense", "balance": 0},
                "508": {"name": "Professional Fees", "type": "expense", "balance": 0}
            }
        }
        
        # Only update if accounts file doesn't exist or is empty
        if not self.accounts:
            self.accounts = default_accounts
            self.save_accounts()
        else:
            # Merge with existing accounts, preserving balances
            for category, accounts in default_accounts.items():
                if category not in self.accounts:
                    self.accounts[category] = accounts
                else:
                    for acc_id, acc_data in accounts.items():
                        if acc_id not in self.accounts[category]:
                            self.accounts[category][acc_id] = acc_data
    
    def save_transactions(self):
        """Save transactions to JSON file"""
        try:
            with open(self.transactions_file, 'w', encoding='utf-8') as f:
                json.dump(self.transactions, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving transactions: {e}")
            return False
    
    def save_accounts(self):
        """Save accounts to JSON file"""
        try:
            with open(self.accounts_file, 'w', encoding='utf-8') as f:
                json.dump(self.accounts, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error saving accounts: {e}")
            return False
    
    def add_transaction(self, date, description, debit_account, credit_account, amount, vat_amount=0, reference="", transaction_type="manual", currency="AED", exchange_rate=1.0, category="", attachment=None, party="", payment_method="", status="completed"):
        try:
            self.balance_manager.record_transaction(
                date=date,
                description=description,
                amount=float(amount),
                debit_account=debit_account,
                credit_account=credit_account,
                reference=reference,
                meta={"kind": transaction_type, "currency": currency, "fx": float(exchange_rate), "category": category, "party": party, "method": payment_method}
            )
            if float(vat_amount) > 0:
                self.handle_vat_transaction(date, f"VAT - {description}", float(vat_amount), debit_account)
            transaction_id = f"TRX{len(self.transactions) + 1:06d}"
            self.transactions.append({
                'transaction_id': transaction_id,
                'date': date,
                'description': description,
                'debit_account': debit_account,
                'credit_account': credit_account,
                'amount': float(amount),
                'vat_amount': float(vat_amount),
                'reference': reference,
                'type': transaction_type,
                'currency': currency,
                'exchange_rate': float(exchange_rate),
                'category': category,
                'attachment': attachment,
                'party': party,
                'payment_method': payment_method,
                'status': status,
                'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            # GAAP Ledger: Auto-record Expense-type transactions through friendly interface
            if str(transaction_type).lower() == "expense" and float(amount) > 0:
                try:
                    if hasattr(self.invoice_manager, 'ledger') and self.invoice_manager.ledger:
                        from ledger_service import EXPENSE_CATEGORY_MAP
                        # Map legacy category -> ledger category, fallback
                        legacy_to_friendly = {
                            "Payroll": "Salary/Wages",
                            "Operating Expenses": "Other Administrative",
                            "Administrative": "Other Administrative",
                            "Marketing": "Marketing & Advertising",
                            "COGS": "Other Administrative",
                            "Revenue": "Other Administrative",
                            "Other": "Other Administrative",
                        }
                        friendly_cat = legacy_to_friendly.get(str(category or "").strip(), "Other Administrative")
                        # Find closest match if not exact
                        if friendly_cat not in EXPENSE_CATEGORY_MAP:
                            friendly_cat = "Other Administrative"
                        try:
                            self.invoice_manager.ledger.on_expense_submitted(
                                category=friendly_cat,
                                amount=float(amount),
                                expense_date=date,
                                payee_name=(party or "").strip(),
                                description=f"{description} | Ref: {reference or 'N/A'}".strip(),
                                username="System"
                            )
                        except Exception as le:
                            print(f"Ledger hook warning (manual expense TRX): {le}")
                except Exception as outer:
                    pass
            if self.save_transactions():
                return True, f"Transaction {transaction_id} added successfully"
            return False, "Failed to save transaction data"
        except Exception as e:
            return False, f"Error adding transaction: {str(e)}"
    
    def account_exists(self, account_id):
        """Check if account exists in chart of accounts"""
        for category in self.accounts.values():
            if account_id in category:
                return True
        return False
    
    def update_account_balance(self, account_id, amount, entry_type):
        """Update account balance based on debit/credit"""
        try:
            for category in self.accounts.values():
                if account_id in category:
                    account = category[account_id]
                    
                    # Determine if account normally has debit or credit balance
                    account_type = account['type']
                    
                    if entry_type == 'debit':
                        # Assets and expenses increase with debit
                        if account_type in ['current_asset', 'bank', 'expense']:
                            account['balance'] += amount
                        else:
                            account['balance'] -= amount
                    else:  # credit
                        # Liabilities, equity, revenue increase with credit
                        if account_type in ['current_liability', 'equity', 'revenue']:
                            account['balance'] += amount
                        else:
                            account['balance'] -= amount
                    
                    return True
            return False
        except Exception as e:
            print(f"Error updating account balance: {e}")
            return False
    
    def handle_vat_transaction(self, date, description, vat_amount, related_account):
        """Handle VAT portion of transaction"""
        try:
            # VAT collected (credit VAT Payable)
            vat_transaction_id = f"VAT{len(self.transactions) + 1:06d}"
            
            vat_transaction = {
                'transaction_id': vat_transaction_id,
                'date': date,
                'description': description,
                'debit_account': related_account,  # Same as main transaction
                'credit_account': "202",  # VAT Payable
                'amount': float(vat_amount),
                'vat_amount': 0,
                'reference': "VAT",
                'type': 'vat',
                'created_at': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            # Update VAT Payable account
            success = self.update_account_balance("202", float(vat_amount), 'credit')
            if success:
                self.transactions.append(vat_transaction)
                return True
            return False
            
        except Exception as e:
            print(f"Error handling VAT transaction: {e}")
            return False
    
    def get_account_balance(self, account_id):
        """Get current balance of an account"""
        for category in self.accounts.values():
            if account_id in category:
                return category[account_id]['balance']
        return 0
    
    def get_transactions_by_date_range(self, start_date, end_date):
        """Get transactions within date range"""
        return [t for t in self.transactions if start_date <= t['date'] <= end_date]
    
    def get_trial_balance(self, as_of_date=None):
        """Generate trial balance"""
        trial_balance = {
            'assets': [],
            'liabilities': [],
            'equity': [],
            'revenue': [],
            'expenses': []
        }
        
        total_debits = 0
        total_credits = 0
        
        for category_name, accounts in self.accounts.items():
            for acc_id, acc_data in accounts.items():
                balance = acc_data['balance']
                
                if category_name in ['assets', 'expenses']:
                    # Normally have debit balances
                    if balance >= 0:
                        trial_balance[category_name].append({
                            'account_id': acc_id,
                            'name': acc_data['name'],
                            'debit': balance,
                            'credit': 0
                        })
                        total_debits += balance
                    else:
                        trial_balance[category_name].append({
                            'account_id': acc_id,
                            'name': acc_data['name'],
                            'debit': 0,
                            'credit': abs(balance)
                        })
                        total_credits += abs(balance)
                else:
                    # Liabilities, equity, revenue normally have credit balances
                    if balance >= 0:
                        trial_balance[category_name].append({
                            'account_id': acc_id,
                            'name': acc_data['name'],
                            'debit': 0,
                            'credit': balance
                        })
                        total_credits += balance
                    else:
                        trial_balance[category_name].append({
                            'account_id': acc_id,
                            'name': acc_data['name'],
                            'debit': abs(balance),
                            'credit': 0
                        })
                        total_debits += abs(balance)
        
        return {
            'trial_balance': trial_balance,
            'total_debits': total_debits,
            'total_credits': total_credits,
            'is_balanced': abs(total_debits - total_credits) < 0.01
        }
    
    def get_vat_report(self, start_date, end_date):
        """Generate VAT report for period"""
        vat_transactions = [
            t for t in self.transactions 
            if start_date <= t['date'] <= end_date and t.get('type') == 'vat'
        ]
        
        vat_collected = sum(t['amount'] for t in vat_transactions if t['credit_account'] == '202')
        vat_paid = sum(t['amount'] for t in vat_transactions if t['debit_account'] == '202')
        
        return {
            'period': f"{start_date} to {end_date}",
            'vat_collected': vat_collected,
            'vat_paid': vat_paid,
            'vat_payable': vat_collected - vat_paid,
            'transactions': vat_transactions
        }
    
    def get_income_statement(self, start_date, end_date):
        """Generate income statement for period"""
        # Get transactions in date range
        period_transactions = self.get_transactions_by_date_range(start_date, end_date)
        
        revenue = 0
        expenses = 0
        
        for transaction in period_transactions:
            # Revenue transactions (credit to revenue accounts)
            if transaction['credit_account'] in ['401', '402']:  # Revenue accounts
                revenue += transaction['amount']
            
            # Expense transactions (debit to expense accounts)
            if transaction['debit_account'].startswith('5'):  # Expense accounts (501-508)
                expenses += transaction['amount']
        
        net_income = revenue - expenses
        
        return {
            'period': f"{start_date} to {end_date}",
            'revenue': revenue,
            'expenses': expenses,
            'net_income': net_income,
            'revenue_breakdown': {
                'medicine_sales': sum(t['amount'] for t in period_transactions if t['credit_account'] == '401'),
                'service_revenue': sum(t['amount'] for t in period_transactions if t['credit_account'] == '402')
            },
            'expense_breakdown': {
                acc_id: self.accounts['expenses'][acc_id]['balance'] 
                for acc_id in self.accounts['expenses'] 
                if self.accounts['expenses'][acc_id]['balance'] > 0
            }
        }
    
    def get_balance_sheet(self, as_of_date):
        """Generate balance sheet as of date"""
        # Calculate account balances up to the as_of_date
        assets = sum(
            self.accounts['assets'][acc_id]['balance'] 
            for acc_id in self.accounts['assets']
        )
        
        liabilities = sum(
            self.accounts['liabilities'][acc_id]['balance'] 
            for acc_id in self.accounts['liabilities']
        )
        
        equity = sum(
            self.accounts['equity'][acc_id]['balance'] 
            for acc_id in self.accounts['equity']
        )
        
        # Add net income to retained earnings
        income_stmt = self.get_income_statement("2000-01-01", as_of_date)
        equity += income_stmt['net_income']
        
        return {
            'as_of_date': as_of_date,
            'assets': assets,
            'liabilities': liabilities,
            'equity': equity,
            'is_balanced': abs(assets - (liabilities + equity)) < 0.01,
            'asset_details': self.accounts['assets'],
            'liability_details': self.accounts['liabilities'],
            'equity_details': self.accounts['equity']
        }

    def delete_transaction(self, transaction_id):
        transaction_to_delete = None
        for transaction in self.transactions:
            if transaction.get('transaction_id') == transaction_id:
                transaction_to_delete = transaction
                break
        if not transaction_to_delete:
            return False, "Transaction not found"
        try:
            debit_account = transaction_to_delete['debit_account']
            credit_account = transaction_to_delete['credit_account']
            amount = transaction_to_delete['amount']
            date_str = datetime.now().strftime("%Y-%m-%d")
            self.balance_manager.record_transaction(
                date=date_str,
                description=f"Reversal of {transaction_id}",
                amount=float(amount),
                debit_account=credit_account,
                credit_account=debit_account,
                reference=transaction_id,
                meta={"kind": "reversal"}
            )
            vat_amount = transaction_to_delete.get('vat_amount', 0)
            if vat_amount > 0:
                self.balance_manager.record_transaction(
                    date=date_str,
                    description=f"VAT reversal {transaction_id}",
                    amount=float(vat_amount),
                    debit_account="202",
                    credit_account=debit_account,
                    reference=transaction_id,
                    meta={"kind": "vat_reversal"}
                )
            self.transactions = [t for t in self.transactions if t.get('transaction_id') != transaction_id]
            if self.save_transactions():
                return True, f"Transaction {transaction_id} deleted successfully"
            return False, "Failed to save changes after deletion"
        except Exception as e:
            return False, f"Error deleting transaction: {str(e)}"

class ScrollableFrame:
    """A scrollable frame that can be used in any dialog"""
    def __init__(self, parent):
        self.canvas = tk.Canvas(parent)
        self.scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Mouse wheel binding
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        
    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")
        
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def pack(self, **kwargs):
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

class TransactionDialog:
    def __init__(self, parent, manager):
        self.manager = manager
        try:
            self.balance_manager = manager.balance_manager
        except Exception:
            self.balance_manager = BalanceManager(getattr(manager, 'data_folder', os.getcwd()))
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("💼 Financial Transaction Manager")
        self.dialog.geometry("1400x800")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        self.refresh_data()
    
    def setup_ui(self):
        # Create main notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.dialog)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Tab 1: Record Transactions
        self.setup_transaction_tab()
        
        # Tab 2: Chart of Accounts
        self.setup_accounts_tab()
        
        # Tab 3: Financial Reports
        self.setup_reports_tab()
        
        # Tab 4: VAT Management
        self.setup_vat_tab()
    
    def setup_transaction_tab(self):
        """Setup the transaction recording tab"""
        transaction_frame = ttk.Frame(self.notebook)
        self.notebook.add(transaction_frame, text="📝 Record Transactions")
        
        # Create scrollable frame
        main_container = tk.Frame(transaction_frame)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_content = scroll_frame.scrollable_frame
        
        ttk.Label(main_content, text="Double-Entry Transaction Recording", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        
        # Transaction Form
        form_frame = ttk.LabelFrame(main_content, text="New Transaction", padding="15")
        form_frame.pack(fill='x', pady=(0, 20))
        
        # Date
        date_frame = ttk.Frame(form_frame)
        date_frame.pack(fill='x', pady=5)
        ttk.Label(date_frame, text="Date *:", font=('Helvetica', 10, 'bold')).pack(side='left')
        self.trans_date = ttk.Entry(date_frame, width=12, font=('Helvetica', 10))
        self.trans_date.pack(side='left', padx=10)
        self.trans_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # Transaction Type
        type_frame = ttk.Frame(form_frame)
        type_frame.pack(fill='x', pady=5)
        ttk.Label(type_frame, text="Type *:", font=('Helvetica', 10, 'bold')).pack(side='left')
        self.trans_type = ttk.Combobox(type_frame, state="readonly", width=24, font=('Helvetica', 10),
                                       values=["Sales","Purchase","Expense","Income","Payment","Receipt","Transfer","Adjustment","Payroll"])
        self.trans_type.pack(side='left', padx=10)
        self.trans_type.set("Expense")

        # Description
        desc_frame = ttk.Frame(form_frame)
        desc_frame.pack(fill='x', pady=5)
        ttk.Label(desc_frame, text="Description *:", font=('Helvetica', 10, 'bold')).pack(side='left')
        self.trans_desc = ttk.Entry(desc_frame, width=60, font=('Helvetica', 10))
        self.trans_desc.pack(side='left', padx=10, fill='x', expand=True)
        
        # Amount, VAT, Currency, FX
        amount_frame = ttk.Frame(form_frame)
        amount_frame.pack(fill='x', pady=5)
        ttk.Label(amount_frame, text="Amount (AED) *:", font=('Helvetica', 10, 'bold')).pack(side='left')
        self.trans_amount = ttk.Entry(amount_frame, width=15, font=('Helvetica', 10))
        self.trans_amount.pack(side='left', padx=10)
        
        ttk.Label(amount_frame, text="VAT Amount:", font=('Helvetica', 10, 'bold')).pack(side='left', padx=(20,0))
        self.trans_vat = ttk.Entry(amount_frame, width=12, font=('Helvetica', 10))
        self.trans_vat.pack(side='left', padx=10)
        self.trans_vat.insert(0, "0.00")

        ttk.Label(amount_frame, text="Currency:", font=('Helvetica', 10, 'bold')).pack(side='left', padx=(20,0))
        self.trans_currency = ttk.Combobox(amount_frame, state="readonly", width=8, font=('Helvetica', 10),
                                           values=["AED","USD","EUR","GBP","SAR","QAR","OMR","KWD","BHD","EGP"])
        self.trans_currency.pack(side='left', padx=6)
        self.trans_currency.set("AED")

        ttk.Label(amount_frame, text="FX Rate:", font=('Helvetica', 10, 'bold')).pack(side='left', padx=(20,0))
        self.trans_fx = ttk.Entry(amount_frame, width=8, font=('Helvetica', 10))
        self.trans_fx.pack(side='left', padx=6)
        self.trans_fx.insert(0, "1.00")
        
        # Account Selection
        accounts_frame = ttk.Frame(form_frame)
        accounts_frame.pack(fill='x', pady=10)
        
        # Debit Account
        debit_frame = ttk.LabelFrame(accounts_frame, text="Debit Account *", padding="10")
        debit_frame.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        ttk.Label(debit_frame, text="Account:", font=('Helvetica', 9, 'bold')).pack(anchor='w')
        self.debit_account = ttk.Combobox(debit_frame, values=self.get_accounts_list(), 
                                         state="readonly", width=25, font=('Helvetica', 9))
        self.debit_account.pack(fill='x', pady=2)
        
        ttk.Label(debit_frame, text="Current Balance:", font=('Helvetica', 9)).pack(anchor='w')
        self.debit_balance = ttk.Label(debit_frame, text="AED 0.00", font=('Helvetica', 9, 'bold'))
        self.debit_balance.pack(anchor='w')
        
        # Credit Account
        credit_frame = ttk.LabelFrame(accounts_frame, text="Credit Account *", padding="10")
        credit_frame.pack(side='left', fill='x', expand=True, padx=(10, 0))
        
        ttk.Label(credit_frame, text="Account:", font=('Helvetica', 9, 'bold')).pack(anchor='w')
        self.credit_account = ttk.Combobox(credit_frame, values=self.get_accounts_list(), 
                                          state="readonly", width=25, font=('Helvetica', 9))
        self.credit_account.pack(fill='x', pady=2)
        
        ttk.Label(credit_frame, text="Current Balance:", font=('Helvetica', 9)).pack(anchor='w')
        self.credit_balance = ttk.Label(credit_frame, text="AED 0.00", font=('Helvetica', 9, 'bold'))
        self.credit_balance.pack(anchor='w')
        
        # Category, Party, Payment Method
        meta_frame = ttk.Frame(form_frame)
        meta_frame.pack(fill='x', pady=5)
        ttk.Label(meta_frame, text="Category:", font=('Helvetica', 10, 'bold')).pack(side='left')
        self.trans_category = ttk.Combobox(meta_frame, state="readonly", width=24, font=('Helvetica', 10),
                                           values=["Operating Expenses","COGS","Administrative","Marketing","Payroll","Revenue","Other"])
        self.trans_category.pack(side='left', padx=10)
        self.trans_category.set("Operating Expenses")
        ttk.Label(meta_frame, text="Party:", font=('Helvetica', 10, 'bold')).pack(side='left', padx=(20,0))
        self.trans_party = ttk.Entry(meta_frame, width=24, font=('Helvetica', 10))
        self.trans_party.pack(side='left', padx=10)
        ttk.Label(meta_frame, text="Method:", font=('Helvetica', 10, 'bold')).pack(side='left', padx=(20,0))
        self.trans_method = ttk.Combobox(meta_frame, state="readonly", width=18, font=('Helvetica', 10),
                                         values=["Cash","Bank Transfer","Card","Cheque","Online"])
        self.trans_method.pack(side='left', padx=10)
        self.trans_method.set("Cash")

        # Reference
        ref_frame = ttk.Frame(form_frame)
        ref_frame.pack(fill='x', pady=5)
        ttk.Label(ref_frame, text="Reference:", font=('Helvetica', 10, 'bold')).pack(side='left')
        self.trans_ref = ttk.Entry(ref_frame, width=30, font=('Helvetica', 10))
        self.trans_ref.pack(side='left', padx=10)

        # Attachment
        attach_frame = ttk.Frame(form_frame)
        attach_frame.pack(fill='x', pady=5)
        ttk.Label(attach_frame, text="Attachment:", font=('Helvetica', 10, 'bold')).pack(side='left')
        self.attach_path = ttk.Entry(attach_frame, width=50, font=('Helvetica', 10))
        self.attach_path.pack(side='left', padx=10)
        def _browse_attach():
            try:
                from tkinter.filedialog import askopenfilename
                p = askopenfilename(title="Select attachment")
                if p:
                    self.attach_path.delete(0, tk.END)
                    self.attach_path.insert(0, p)
            except Exception:
                pass
        ttk.Button(attach_frame, text="Browse", command=_browse_attach).pack(side='left', padx=6)
        
        # Buttons
        button_frame = ttk.Frame(form_frame)
        button_frame.pack(fill='x', pady=10)
        
        ttk.Button(button_frame, text="💾 Record Transaction", 
                  command=self.record_transaction, style='Accent.TButton').pack(side='left', padx=5)
        ttk.Button(button_frame, text="🔄 Clear Form", 
                  command=self.clear_form).pack(side='left', padx=5)
        ttk.Button(button_frame, text="🔄 Sync with Invoices", 
                  command=self.sync_with_invoices).pack(side='left', padx=5)
        
        # Bind account selection events
        self.debit_account.bind('<<ComboboxSelected>>', self.update_account_balances)
        self.credit_account.bind('<<ComboboxSelected>>', self.update_account_balances)
        
        # Recent Transactions
        recent_frame = ttk.LabelFrame(main_content, text="Recent Transactions", padding="10")
        recent_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Create transactions treeview
        columns = ('ID', 'Date', 'Description', 'Debit', 'Credit', 'Amount', 'VAT', 'Reference', 'Type')
        self.trans_tree = ttk.Treeview(recent_frame, columns=columns, show='headings', height=12)
        
        # Define headings
        self.trans_tree.heading('ID', text='ID')
        self.trans_tree.heading('Date', text='Date')
        self.trans_tree.heading('Description', text='Description')
        self.trans_tree.heading('Debit', text='Debit')
        self.trans_tree.heading('Credit', text='Credit')
        self.trans_tree.heading('Amount', text='Amount')
        self.trans_tree.heading('VAT', text='VAT')
        self.trans_tree.heading('Reference', text='Reference')
        self.trans_tree.heading('Type', text='Type')
        
        # Define columns
        self.trans_tree.column('ID', width=80)
        self.trans_tree.column('Date', width=90)
        self.trans_tree.column('Description', width=200)
        self.trans_tree.column('Debit', width=80)
        self.trans_tree.column('Credit', width=80)
        self.trans_tree.column('Amount', width=90)
        self.trans_tree.column('VAT', width=80)
        self.trans_tree.column('Reference', width=100)
        self.trans_tree.column('Type', width=80)
        
        # Scrollbar
        trans_scrollbar = ttk.Scrollbar(recent_frame, orient=tk.VERTICAL, command=self.trans_tree.yview)
        self.trans_tree.configure(yscrollcommand=trans_scrollbar.set)
        
        self.trans_tree.pack(side='left', fill='both', expand=True)
        trans_scrollbar.pack(side='right', fill='y')
        
        # Add delete transaction button
        def delete_selected_transaction():
            selection = self.trans_tree.selection()
            if not selection:
                messagebox.showwarning("Warning", "Please select a transaction to delete")
                return
                
            transaction_id = self.trans_tree.item(selection[0])['values'][0]
            if messagebox.askyesno("Confirm Delete", 
                                  f"Are you sure you want to delete this transaction?\nID: {transaction_id}"):
                success, message = self.manager.delete_transaction(transaction_id)
                if success:
                    messagebox.showinfo("Success", message)
                    self.refresh_data()
                else:
                    messagebox.showerror("Error", message)
        
        delete_button_frame = ttk.Frame(recent_frame)
        delete_button_frame.pack(fill='x', pady=5)
        
        ttk.Button(delete_button_frame, text="🗑️ Delete Selected Transaction", 
                  command=delete_selected_transaction).pack(side='left', padx=5)
    
    def setup_accounts_tab(self):
        """Setup chart of accounts tab"""
        accounts_frame = ttk.Frame(self.notebook)
        self.notebook.add(accounts_frame, text="📊 Chart of Accounts")
        
        # Create scrollable frame
        main_container = tk.Frame(accounts_frame)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_content = scroll_frame.scrollable_frame
        
        ttk.Label(main_content, text="Chart of Accounts", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        
        # Trial Balance Summary
        trial_frame = ttk.LabelFrame(main_content, text="Trial Balance Summary", padding="10")
        trial_frame.pack(fill='x', pady=(0, 20))
        
        self.trial_balance_text = scrolledtext.ScrolledText(trial_frame, height=8, font=('Consolas', 9))
        self.trial_balance_text.pack(fill='x')
        
        # Accounts by Category
        categories = [
            ("💳 Assets", "assets"),
            ("📋 Liabilities", "liabilities"), 
            ("🏛️ Equity", "equity"),
            ("💰 Revenue", "revenue"),
            ("💸 Expenses", "expenses")
        ]
        
        for title, category in categories:
            cat_frame = ttk.LabelFrame(main_content, text=title, padding="10")
            cat_frame.pack(fill='x', pady=5)
            
            # Create treeview for this category
            columns = ('Account ID', 'Account Name', 'Type', 'Balance')
            tree = ttk.Treeview(cat_frame, columns=columns, show='headings', height=4)
            
            for col in columns:
                tree.heading(col, text=col)
                if col == 'Account Name':
                    tree.column(col, width=200)
                else:
                    tree.column(col, width=120)
            
            # Add accounts
            for acc_id, acc_data in self.manager.accounts.get(category, {}).items():
                tree.insert('', tk.END, values=(
                    acc_id,
                    acc_data['name'],
                    acc_data['type'],
                    f"AED {acc_data['balance']:,.2f}"
                ))
            
            tree.pack(fill='x')
    
    def setup_reports_tab(self):
        """Setup financial reports tab"""
        reports_frame = ttk.Frame(self.notebook)
        self.notebook.add(reports_frame, text="📈 Financial Reports")
        
        # Create scrollable frame
        main_container = tk.Frame(reports_frame)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_content = scroll_frame.scrollable_frame
        
        ttk.Label(main_content, text="Financial Statements", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        
        # Date Range
        date_frame = ttk.LabelFrame(main_content, text="Report Period", padding="10")
        date_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(date_frame, text="From:").grid(row=0, column=0, padx=5, pady=5)
        self.report_start = ttk.Entry(date_frame, width=12)
        self.report_start.grid(row=0, column=1, padx=5, pady=5)
        self.report_start.insert(0, (datetime.now().replace(day=1) - timedelta(days=30)).strftime("%Y-%m-%d"))
        
        ttk.Label(date_frame, text="To:").grid(row=0, column=2, padx=5, pady=5)
        self.report_end = ttk.Entry(date_frame, width=12)
        self.report_end.grid(row=0, column=3, padx=5, pady=5)
        self.report_end.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # Report Buttons
        button_frame = ttk.Frame(date_frame)
        button_frame.grid(row=1, column=0, columnspan=4, pady=10)
        
        ttk.Button(button_frame, text="💰 Income Statement", 
                  command=self.generate_income_statement).pack(side='left', padx=5)
        ttk.Button(button_frame, text="🏛️ Balance Sheet", 
                  command=self.generate_balance_sheet).pack(side='left', padx=5)
        ttk.Button(button_frame, text="⚖️ Trial Balance", 
                  command=self.generate_trial_balance).pack(side='left', padx=5)
        
        # Report Display
        report_display_frame = ttk.LabelFrame(main_content, text="Financial Report", padding="10")
        report_display_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        self.report_text = scrolledtext.ScrolledText(report_display_frame, height=20, font=('Consolas', 9))
        self.report_text.pack(fill='both', expand=True)
    
    def setup_vat_tab(self):
        """Setup VAT management tab"""
        vat_frame = ttk.Frame(self.notebook)
        self.notebook.add(vat_frame, text="🧾 VAT Management")
        
        # Create scrollable frame
        main_container = tk.Frame(vat_frame)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_content = scroll_frame.scrollable_frame
        
        ttk.Label(main_content, text="Value Added Tax (VAT) Management", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        
        # VAT Period
        vat_period_frame = ttk.LabelFrame(main_content, text="VAT Reporting Period", padding="10")
        vat_period_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(vat_period_frame, text="From:").grid(row=0, column=0, padx=5, pady=5)
        self.vat_start = ttk.Entry(vat_period_frame, width=12)
        self.vat_start.grid(row=0, column=1, padx=5, pady=5)
        self.vat_start.insert(0, (datetime.now().replace(day=1) - timedelta(days=30)).strftime("%Y-%m-%d"))
        
        ttk.Label(vat_period_frame, text="To:").grid(row=0, column=2, padx=5, pady=5)
        self.vat_end = ttk.Entry(vat_period_frame, width=12)
        self.vat_end.grid(row=0, column=3, padx=5, pady=5)
        self.vat_end.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # VAT Buttons
        vat_button_frame = ttk.Frame(vat_period_frame)
        vat_button_frame.grid(row=1, column=0, columnspan=4, pady=10)
        
        ttk.Button(vat_button_frame, text="📊 Generate VAT Report", 
                  command=self.generate_vat_report).pack(side='left', padx=5)
        ttk.Button(vat_button_frame, text="💾 Export VAT Return", 
                  command=self.export_vat_return).pack(side='left', padx=5)
        
        # VAT Summary
        vat_summary_frame = ttk.LabelFrame(main_content, text="VAT Summary", padding="10")
        vat_summary_frame.pack(fill='x', pady=(0, 10))
        
        self.vat_summary_text = scrolledtext.ScrolledText(vat_summary_frame, height=6, font=('Consolas', 9))
        self.vat_summary_text.pack(fill='x')
        
        # VAT Transactions
        vat_trans_frame = ttk.LabelFrame(main_content, text="VAT Transactions", padding="10")
        vat_trans_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        columns = ('Date', 'Description', 'VAT Amount', 'Type', 'Related Account')
        self.vat_tree = ttk.Treeview(vat_trans_frame, columns=columns, show='headings', height=8)
        
        for col in columns:
            self.vat_tree.heading(col, text=col)
            if col == 'Description':
                self.vat_tree.column(col, width=200)
            else:
                self.vat_tree.column(col, width=120)
        
        vat_scrollbar = ttk.Scrollbar(vat_trans_frame, orient=tk.VERTICAL, command=self.vat_tree.yview)
        self.vat_tree.configure(yscrollcommand=vat_scrollbar.set)
        
        self.vat_tree.pack(side='left', fill='both', expand=True)
        vat_scrollbar.pack(side='right', fill='y')
    
    def get_accounts_list(self):
        accounts = []
        for acc_id, acc_data in self.balance_manager.accounts.items():
            accounts.append(f"{acc_id} - {acc_data.get('name')}")
        return sorted(accounts)
    
    def update_account_balances(self, event=None):
        try:
            if self.debit_account.get():
                acc_id = self.debit_account.get().split(' - ')[0]
                balance = self.balance_manager.accounts.get(acc_id, {}).get('balance', 0)
                self.debit_balance.config(text=f"AED {balance:,.2f}")
            if self.credit_account.get():
                acc_id = self.credit_account.get().split(' - ')[0]
                balance = self.balance_manager.accounts.get(acc_id, {}).get('balance', 0)
                self.credit_balance.config(text=f"AED {balance:,.2f}")
        except Exception:
            pass
    
    def sync_with_invoices(self):
        """Sync with invoice and purchase data"""
        try:
            self.manager.sync_with_invoice_data()
            messagebox.showinfo("Success", "Transaction data synchronized with invoices and purchases!")
            self.refresh_data()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to sync data: {e}")
    
    def record_transaction(self):
        """Record a new transaction"""
        try:
            # Validate form
            if not all([self.trans_date.get(), self.trans_desc.get(), self.trans_amount.get(),
                       self.debit_account.get(), self.credit_account.get()]):
                messagebox.showwarning("Warning", "Please fill all required fields (*)")
                return
            
            # Get account IDs
            debit_acc = self.debit_account.get().split(' - ')[0]
            credit_acc = self.credit_account.get().split(' - ')[0]
            
            # Record transaction
            # Copy attachment if provided
            attachment_saved = None
            try:
                ap = (self.attach_path.get() or '').strip()
                if ap and os.path.isfile(ap):
                    attach_dir = os.path.join(self.manager.data_folder, 'attachments')
                    os.makedirs(attach_dir, exist_ok=True)
                    fname = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.path.basename(ap)}"
                    dst = os.path.join(attach_dir, fname)
                    shutil.copy2(ap, dst)
                    attachment_saved = dst
            except Exception:
                attachment_saved = None

            success, message = self.manager.add_transaction(
                date=self.trans_date.get(),
                description=self.trans_desc.get(),
                debit_account=debit_acc,
                credit_account=credit_acc,
                amount=float(self.trans_amount.get()),
                vat_amount=float(self.trans_vat.get() or 0),
                reference=self.trans_ref.get(),
                transaction_type=self.trans_type.get().lower(),
                currency=self.trans_currency.get(),
                exchange_rate=float(self.trans_fx.get() or 1.0),
                category=self.trans_category.get(),
                attachment=attachment_saved,
                party=(self.trans_party.get() or '').strip(),
                payment_method=self.trans_method.get()
            )
            
            if success:
                messagebox.showinfo("Success", message)
                self.clear_form()
                self.refresh_data()
            else:
                messagebox.showerror("Error", message)
                
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values for amount and VAT")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to record transaction: {str(e)}")
    
    def clear_form(self):
        """Clear the transaction form"""
        self.trans_desc.delete(0, tk.END)
        self.trans_amount.delete(0, tk.END)
        self.trans_vat.delete(0, tk.END)
        self.trans_vat.insert(0, "0.00")
        self.trans_ref.delete(0, tk.END)
        self.attach_path.delete(0, tk.END)
        self.debit_account.set('')
        self.credit_account.set('')
        self.debit_balance.config(text="AED 0.00")
        self.credit_balance.config(text="AED 0.00")
    
    def refresh_data(self):
        """Refresh all data displays"""
        self.refresh_transactions()
        self.refresh_trial_balance()
        self.refresh_vat_data()
    
    def refresh_transactions(self):
        """Refresh transactions list"""
        for item in self.trans_tree.get_children():
            self.trans_tree.delete(item)
        
        # Optional filters (simple phase 1): show last 50
        recent_trans = self.manager.transactions[-50:]
        for trans in recent_trans:
            self.trans_tree.insert('', tk.END, values=(
                trans['transaction_id'],
                trans['date'],
                trans['description'],
                trans['debit_account'],
                trans['credit_account'],
                f"AED {trans['amount']:,.2f}",
                f"AED {trans['vat_amount']:,.2f}",
                trans.get('reference', ''),
                trans.get('type', 'manual')
            ))
    
    def refresh_trial_balance(self):
        """Refresh trial balance display"""
        trial_balance = self.manager.get_trial_balance()
        
        self.trial_balance_text.delete('1.0', tk.END)
        
        self.trial_balance_text.insert(tk.END, "TRIAL BALANCE\n")
        self.trial_balance_text.insert(tk.END, "="*50 + "\n\n")
        
        total_debits = 0
        total_credits = 0
        
        for category, accounts in trial_balance['trial_balance'].items():
            if accounts:
                self.trial_balance_text.insert(tk.END, f"\n{category.upper()}:\n")
                for acc in accounts:
                    self.trial_balance_text.insert(tk.END, 
                        f"  {acc['account_id']} - {acc['name']:<30} {acc['debit']:>10,.2f} {acc['credit']:>10,.2f}\n")
                    total_debits += acc['debit']
                    total_credits += acc['credit']
        
        self.trial_balance_text.insert(tk.END, "\n" + "="*50 + "\n")
        self.trial_balance_text.insert(tk.END, f"Total Debits:  AED {total_debits:>15,.2f}\n")
        self.trial_balance_text.insert(tk.END, f"Total Credits: AED {total_credits:>15,.2f}\n")
        
        if trial_balance['is_balanced']:
            self.trial_balance_text.insert(tk.END, "✅ ACCOUNTS ARE BALANCED\n")
        else:
            self.trial_balance_text.insert(tk.END, f"❌ OUT OF BALANCE: AED {abs(total_debits - total_credits):,.2f}\n")
    
    def refresh_vat_data(self):
        """Refresh VAT data"""
        # This will be populated when generating VAT reports
        pass
    
    def generate_income_statement(self):
        """Generate income statement report"""
        try:
            report = self.manager.get_income_statement(
                self.report_start.get(),
                self.report_end.get()
            )
            
            self.report_text.delete('1.0', tk.END)
            
            self.report_text.insert(tk.END, "ASSISTEM\n")
            self.report_text.insert(tk.END, "STATEMENT OF COMPREHENSIVE INCOME\n")
            self.report_text.insert(tk.END, f"FOR THE PERIOD {report['period']}\n")
            self.report_text.insert(tk.END, "="*60 + "\n\n")
            
            self.report_text.insert(tk.END, "REVENUE:\n")
            self.report_text.insert(tk.END, 
                f"  Medicine Sales{'':<26} AED {report['revenue_breakdown']['medicine_sales']:>12,.2f}\n")
            self.report_text.insert(tk.END, 
                f"  Service Revenue{'':<25} AED {report['revenue_breakdown']['service_revenue']:>12,.2f}\n")
            
            self.report_text.insert(tk.END, f"{'Total Revenue':<40} AED {report['revenue']:>12,.2f}\n\n")
            
            self.report_text.insert(tk.END, "EXPENSES:\n")
            for acc_id, amount in report['expense_breakdown'].items():
                account_name = self.manager.accounts['expenses'][acc_id]['name']
                self.report_text.insert(tk.END, 
                    f"  {account_name:<36} AED {amount:>12,.2f}\n")
            
            self.report_text.insert(tk.END, f"{'Total Expenses':<40} AED {report['expenses']:>12,.2f}\n\n")
            self.report_text.insert(tk.END, "="*60 + "\n")
            self.report_text.insert(tk.END, f"{'NET INCOME':<40} AED {report['net_income']:>12,.2f}\n")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate income statement: {str(e)}")
    
    def generate_balance_sheet(self):
        """Generate balance sheet report"""
        try:
            report = self.manager.get_balance_sheet(self.report_end.get())
            
            self.report_text.delete('1.0', tk.END)
            
            self.report_text.insert(tk.END, "ASSISTEM\n")
            self.report_text.insert(tk.END, "STATEMENT OF FINANCIAL POSITION\n")
            self.report_text.insert(tk.END, f"AS AT {report['as_of_date']}\n")
            self.report_text.insert(tk.END, "="*60 + "\n\n")
            
            self.report_text.insert(tk.END, "ASSETS:\n")
            for acc_id, acc_data in report['asset_details'].items():
                if acc_data['balance'] != 0:
                    self.report_text.insert(tk.END, 
                        f"  {acc_data['name']:<40} AED {acc_data['balance']:>12,.2f}\n")
            self.report_text.insert(tk.END, f"{'Total Assets':<40} AED {report['assets']:>12,.2f}\n\n")
            
            self.report_text.insert(tk.END, "LIABILITIES:\n")
            for acc_id, acc_data in report['liability_details'].items():
                if acc_data['balance'] != 0:
                    self.report_text.insert(tk.END, 
                        f"  {acc_data['name']:<40} AED {acc_data['balance']:>12,.2f}\n")
            self.report_text.insert(tk.END, f"{'Total Liabilities':<40} AED {report['liabilities']:>12,.2f}\n\n")
            
            self.report_text.insert(tk.END, "EQUITY:\n")
            for acc_id, acc_data in report['equity_details'].items():
                if acc_data['balance'] != 0:
                    self.report_text.insert(tk.END, 
                        f"  {acc_data['name']:<40} AED {acc_data['balance']:>12,.2f}\n")
            self.report_text.insert(tk.END, f"{'Total Equity':<40} AED {report['equity']:>12,.2f}\n\n")
            
            self.report_text.insert(tk.END, "="*60 + "\n")
            if report['is_balanced']:
                self.report_text.insert(tk.END, f"Assets = Liabilities + Equity: AED {report['assets']:,.2f} = AED {report['liabilities'] + report['equity']:,.2f}\n")
                self.report_text.insert(tk.END, "✅ BALANCE SHEET IS BALANCED\n")
            else:
                self.report_text.insert(tk.END, "❌ BALANCE SHEET IS NOT BALANCED\n")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate balance sheet: {str(e)}")
    
    def generate_trial_balance(self):
        """Generate trial balance report"""
        try:
            report = self.manager.get_trial_balance()
            
            self.report_text.delete('1.0', tk.END)
            
            self.report_text.insert(tk.END, "ASSISTEM\n")
            self.report_text.insert(tk.END, "TRIAL BALANCE\n")
            self.report_text.insert(tk.END, f"AS AT {datetime.now().strftime('%Y-%m-%d')}\n")
            self.report_text.insert(tk.END, "="*70 + "\n\n")
            
            self.report_text.insert(tk.END, f"{'Account':<40} {'Debit (AED)':>15} {'Credit (AED)':>15}\n")
            self.report_text.insert(tk.END, "-"*70 + "\n")
            
            total_debits = 0
            total_credits = 0
            
            for category, accounts in report['trial_balance'].items():
                if accounts:
                    self.report_text.insert(tk.END, f"\n{category.upper()}:\n")
                    for acc in accounts:
                        if acc['debit'] > 0 or acc['credit'] > 0:
                            self.report_text.insert(tk.END, 
                                f"  {acc['name']:<38} {acc['debit']:>15,.2f} {acc['credit']:>15,.2f}\n")
                            total_debits += acc['debit']
                            total_credits += acc['credit']
            
            self.report_text.insert(tk.END, "-"*70 + "\n")
            self.report_text.insert(tk.END, f"{'TOTALS':<38} {total_debits:>15,.2f} {total_credits:>15,.2f}\n")
            
            if report['is_balanced']:
                self.report_text.insert(tk.END, "✅ TRIAL BALANCE IS BALANCED\n")
            else:
                self.report_text.insert(tk.END, f"❌ OUT OF BALANCE BY: AED {abs(total_debits - total_credits):,.2f}\n")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate trial balance: {str(e)}")
    
    def generate_vat_report(self):
        """Generate VAT report"""
        try:
            report = self.manager.get_vat_report(
                self.vat_start.get(),
                self.vat_end.get()
            )
            
            self.vat_summary_text.delete('1.0', tk.END)
            
            self.vat_summary_text.insert(tk.END, "VAT RETURN SUMMARY\n")
            self.vat_summary_text.insert(tk.END, "="*50 + "\n\n")
            self.vat_summary_text.insert(tk.END, f"Period: {report['period']}\n\n")
            self.vat_summary_text.insert(tk.END, f"VAT Collected:  AED {report['vat_collected']:>12,.2f}\n")
            self.vat_summary_text.insert(tk.END, f"VAT Paid:       AED {report['vat_paid']:>12,.2f}\n")
            self.vat_summary_text.insert(tk.END, "="*50 + "\n")
            self.vat_summary_text.insert(tk.END, f"Net VAT Payable: AED {report['vat_payable']:>12,.2f}\n")
            
            # Refresh VAT transactions
            for item in self.vat_tree.get_children():
                self.vat_tree.delete(item)
            
            for trans in report['transactions']:
                vat_type = "Collected" if trans['credit_account'] == '202' else "Paid"
                self.vat_tree.insert('', tk.END, values=(
                    trans['date'],
                    trans['description'],
                    f"AED {trans['amount']:,.2f}",
                    vat_type,
                    trans['debit_account']
                ))
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate VAT report: {str(e)}")
    
    def export_vat_return(self):
        """Export VAT return as text file"""
        try:
            report = self.manager.get_vat_report(
                self.vat_start.get(),
                self.vat_end.get()
            )
            
            # Create VAT returns folder
            vat_folder = os.path.join(self.manager.data_folder, "VATReturns")
            if not os.path.exists(vat_folder):
                os.makedirs(vat_folder)
            
            filename = f"VAT_Return_{self.vat_start.get()}_to_{self.vat_end.get()}.txt"
            filepath = os.path.join(vat_folder, filename)
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("ASSISTEM - VAT RETURN\n")
                f.write("="*50 + "\n\n")
                f.write(f"TRN: 100466797600003\n")
                f.write(f"Period: {report['period']}\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n\n")
                f.write(f"VAT Collected on Sales:    AED {report['vat_collected']:>12,.2f}\n")
                f.write(f"VAT Paid on Purchases:     AED {report['vat_paid']:>12,.2f}\n")
                f.write("="*50 + "\n")
                f.write(f"Net VAT Payable to FTA:    AED {report['vat_payable']:>12,.2f}\n\n")
                f.write("VAT Transactions:\n")
                f.write("-"*50 + "\n")
                for trans in report['transactions']:
                    vat_type = "Collected" if trans['credit_account'] == '202' else "Paid"
                    f.write(f"{trans['date']} - {trans['description']} - AED {trans['amount']:,.2f} ({vat_type})\n")
            
            messagebox.showinfo("Success", f"VAT return exported to:\n{filepath}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export VAT return: {str(e)}")
