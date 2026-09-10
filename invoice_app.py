# invoice_app.py
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from PIL import Image, ImageTk
import json
import os
import shutil
import webbrowser
import subprocess
import sys
from datetime import datetime, timedelta, date
import platform
import uuid
import urllib.parse
import math
from decimal import Decimal, ROUND_HALF_UP
import time
import csv
import hashlib
from pathlib import Path
import ui_undo
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except Exception:
    OPENPYXL_AVAILABLE = False
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.units import inch
    REPORTLAB_PDF_AVAILABLE = True
except Exception:
    REPORTLAB_PDF_AVAILABLE = False

try:
    from temperature_log import TemperatureLogScreen
    TEMPERATURE_LOG_AVAILABLE = True
except Exception:
    TEMPERATURE_LOG_AVAILABLE = False

try:
    from inventory_report import InventoryReportScreen
    INVENTORY_REPORT_AVAILABLE = True
except Exception:
    INVENTORY_REPORT_AVAILABLE = False

try:
    from cost_center_manager import CostCenterManager
    COST_CENTER_AVAILABLE = True
except Exception:
    COST_CENTER_AVAILABLE = False

# Import the InvoiceManager from invoice_manager module
try:
    from invoice_manager import InvoiceManager
except ImportError:
    # Create a simple InvoiceManager class if the module is not available
        class InvoiceManager:
            def __init__(self):
                self.invoice_folder = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")
                # Try to detect if it exists elsewhere
                possible_roots = [
                    os.path.join(os.path.expanduser("~"), "Google Drive"),
                    os.path.join(os.path.expanduser("~"), "GoogleDrive"),
                    os.path.join(os.path.expanduser("~"), "Documents"),
                    os.path.join(os.path.expanduser("~"), "Desktop")
                ]
                for root in possible_roots:
                    path = os.path.join(root, "HopePharmaData")
                    if os.path.exists(path):
                        self.invoice_folder = path
                        break
                
                os.makedirs(self.invoice_folder, exist_ok=True)
        
        def generate_invoice_id(self):
            return f"INV-{datetime.now().strftime('%Y%m%d')}-001"
        
        def add_invoice_from_dict(self, invoice_dict):
            return True
        
        def get_invoice(self, invoice_id):
            return None
        
        def get_all_invoices_dict(self):
            return []
        
        def search_invoices(self, query):
            return []
        
        def delete_invoice(self, invoice_id):
            return True
        
        def update_invoice_id(self, old_id, new_id):
            return True, "Success"
        
        def save_invoices(self):
            return True
        
        def update_invoice(self, invoice_obj):
            return True
        
        def get_clients(self):
            return []
        
        def get_sales_report(self, start_date, end_date):
            return {}
        
        def get_client_report(self, client_name, start_date, end_date):
            return {}
        
        def get_company_report(self, start_date, end_date):
            return {}
        
        def create_purchase(self, **kwargs):
            return type('Purchase', (), {'purchase_id': 'PUR-001'})()
        
        def get_purchase(self, purchase_id):
            return None
        
        def get_all_purchases_dict(self):
            return []
        
        def search_purchases(self, query):
            return []
        
        def delete_purchase(self, purchase_id):
            return True
        
        def update_purchase(self, purchase_id, **kwargs):
            return True
        
        def upload_purchase_receipt(self, purchase_id, file_path):
            return True
        
        def open_purchase_receipt(self, purchase_id):
            return True
        
        def upload_receipt(self, invoice_id, cost_index, file_path):
            return "receipt.pdf"
        
        def open_receipt(self, invoice_id, cost_index):
            return True

        # NEW METHODS FOR BALANCE ADJUSTMENT
        def delete_invoice_with_balance_adjustment(self, invoice_id):
            """Delete an invoice and adjust balances for any payments made"""
            try:
                # Get the invoice first
                invoice = self.get_invoice(invoice_id)
                if not invoice:
                    return False, "Invoice not found"
                
                invoice_dict = invoice.to_dict() if hasattr(invoice, 'to_dict') else invoice
                
                # Check if there are payments that need to be reversed
                total_paid = invoice_dict.get('total_paid', 0)
                if total_paid > 0:
                    # We need to reverse the payments by withdrawing from accounts
                    try:
                        from balance_manager import BalanceManager
                        balance_manager = BalanceManager(self.invoice_folder)
                        
                        # Get payment history (you might need to store this in your invoice data)
                        # For now, we'll assume payments were made to the first account
                        accounts = balance_manager.get_account_names()
                        if accounts:
                            # Withdraw the total paid amount from the first account
                            success, message = balance_manager.update_balance(
                                accounts[0],  # Use first available account
                                total_paid,
                                "withdrawal",
                                f"Payment reversal for deleted invoice: {invoice_id}"
                            )
                            if not success:
                                return False, f"Failed to adjust balance: {message}"
                    except ImportError:
                        print("Balance manager not available for balance adjustment")
                    except Exception as e:
                        print(f"Error adjusting balance: {e}")
                
                # Reverse Inventory Sales
                try:
                    from inventory_system import InventoryManager
                    # Assuming inventory.json is in the same folder or standard location
                    # The InventoryManager usually takes the data folder
                    inventory_manager = InventoryManager(self.invoice_folder)
                    inv_success, inv_msg = inventory_manager.reverse_invoice_sales(invoice_id)
                    if not inv_success:
                        print(f"Warning: Inventory reversal failed: {inv_msg}")
                except ImportError:
                    pass # Inventory system might not be present
                except Exception as e:
                    print(f"Error reversing inventory: {e}")

                # Now delete the invoice
                success = self.delete_invoice(invoice_id)
                if success:
                    return True, f"Invoice {invoice_id} deleted successfully. Balance adjusted."
                else:
                    return False, "Failed to delete invoice"
                    
            except Exception as e:
                return False, f"Error deleting invoice: {str(e)}"

        def delete_purchase_with_balance_adjustment(self, purchase_id):
            """Delete a purchase and adjust balances"""
            try:
                # Get the purchase first
                purchase = self.get_purchase(purchase_id)
                if not purchase:
                    return False, "Purchase not found"
                
                purchase_dict = purchase.to_dict() if hasattr(purchase, 'to_dict') else purchase
                
                # Get purchase amount and account
                amount = purchase_dict.get('amount', 0)
                account = purchase_dict.get('account', 'Cash')
                
                if amount > 0:
                    try:
                        from balance_manager import BalanceManager
                        balance_manager = BalanceManager(self.invoice_folder)
                        expense_key = "400" if "400" in balance_manager.accounts else (balance_manager._find_account_by_name_fuzzy("purchases") or "400")
                        balance_manager.record_transaction(
                            date=purchase_dict.get('date', datetime.now().strftime("%Y-%m-%d")),
                            description=f"Purchase reversal for deleted purchase: {purchase_id}",
                            amount=float(amount),
                            debit_account=account,
                            credit_account=expense_key,
                            reference=str(purchase_id),
                            meta={"kind": "purchase_reversal"},
                        )
                    except ImportError:
                        print("Balance manager not available for balance adjustment")
                    except Exception as e:
                        print(f"Error adjusting balance: {e}")
                
                # Now delete the purchase
                success = self.delete_purchase(purchase_id)
                if success:
                    return True, f"Purchase {purchase_id} deleted successfully. Balance adjusted."
                else:
                    return False, "Failed to delete purchase"
                    
            except Exception as e:
                return False, f"Error deleting purchase: {str(e)}"

        # ADD THESE METHODS FOR BALANCE INTEGRATION:
    
        def process_purchase_with_balance(self, purchase_data):
            """Process purchase and update balance manager"""
            try:
                from balance_manager import BalanceManager
                balance_manager = BalanceManager(self.invoice_folder)
                success, message = balance_manager.process_purchase(purchase_data)
                return success, message
                
            except ImportError:
                return False, "Balance manager not available"
            except Exception as e:
                return False, f"Error processing purchase in balance manager: {str(e)}"
    
        def process_invoice_payment_with_balance(self, invoice_data, payment_amount, account_name):
            """Process invoice payment and update balance manager"""
            try:
                from balance_manager import BalanceManager
                balance_manager = BalanceManager(self.invoice_folder)
                
                success, message = balance_manager.process_invoice_payment(
                    invoice_data, payment_amount, account_name
                )
                return success, message
                
            except ImportError:
                return False, "Balance manager not available"
            except Exception as e:
                return False, f"Error processing invoice payment in balance manager: {str(e)}"
    
        def reverse_invoice_payment_with_balance(self, invoice_data, payment_amount, account_name):
            """Reverse invoice payment when invoice is deleted"""
            try:
                from balance_manager import BalanceManager
                balance_manager = BalanceManager(self.invoice_folder)
                
                success, message = balance_manager.reverse_invoice_payment(
                    invoice_data, payment_amount, account_name
                )
                return success, message
                
            except ImportError:
                return False, "Balance manager not available"
            except Exception as e:
                return False, f"Error reversing invoice payment in balance manager: {str(e)}"

# ... rest of your invoice_app.py file continues ...

# PDF generator wiring: use EnhancedPDFGenerator from hope_pharma_complete
try:
    from hope_pharma_complete import EnhancedPDFGenerator as PDFGenerator
    PDF_AVAILABLE = True
    print("✅ PDF Generator wired to EnhancedPDFGenerator")
except Exception as e:
    PDF_AVAILABLE = False
    print(f"❌ PDF generation not available: {e}")

# Import Inventory System
# Add to imports section
try:
    from balance_manager import show_balance_manager, BalanceManager
    BALANCE_MANAGER_AVAILABLE = True
    print("✅ Balance Manager is available")
except ImportError as e:
    BALANCE_MANAGER_AVAILABLE = False
    print(f"❌ Balance manager not available: {e}")

try:
    from inventory_system import InventorySystem, InventoryManager
    INVENTORY_AVAILABLE = True
    print("✅ Inventory System is available")
except ImportError as e:
    INVENTORY_AVAILABLE = False
    print(f"❌ Inventory system not available: {e}")

try:
    from warehouse_extension import (
        CompanyConfigurationFrame,
        CompanySettingsDialog,
        WarehouseManagementDialog,
        WarehouseManager,
        MODULE_LABELS,
        default_company_configuration,
        ensure_company_extension_storage,
        load_company_configuration,
        normalize_company_record,
        save_company_configuration,
    )
    WAREHOUSE_EXTENSION_AVAILABLE = True
except Exception as e:
    WAREHOUSE_EXTENSION_AVAILABLE = False
    print(f"❌ Warehouse extension not available: {e}")


def _get_inventory_manager_for_invoice(invoice_manager):
    if not INVENTORY_AVAILABLE:
        return None
    try:
        data_folder = getattr(invoice_manager, 'invoice_folder', None)
        if not data_folder:
            data_folder = os.getcwd()
        return InventoryManager(data_folder)
    except Exception:
        return None

try:
    from marketplace_manager import MarketplaceManager
    MARKETPLACE_MANAGER_AVAILABLE = True
    print("✅ Marketplace Manager backend is available")
except ImportError as e:
    MARKETPLACE_MANAGER_AVAILABLE = False
    print(f"❌ Marketplace manager backend not available: {e}")
    # Create a dummy InventorySystem class
    class InventorySystem:
        def __init__(self, parent, data_folder, invoice_manager=None):
            self.parent = parent
            ttk.Label(parent, text="Inventory system not available", 
                     font=('Helvetica', 16, 'bold')).pack(pady=50)

# Import Transaction Manager
try:
    from transaction_manager import TransactionManager, TransactionDialog
    TRANSACTION_MANAGER_AVAILABLE = True
    print("✅ Transaction Manager is available")
except ImportError as e:
    TRANSACTION_MANAGER_AVAILABLE = False
    print(f"❌ Transaction manager not available: {e}")

# Import GAAP Financial Reporting Dashboard
try:
    from gaap_reporting_dashboard import GAAPReportingDashboard
    GAAP_DASHBOARD_AVAILABLE = True
    print("✅ GAAP Financial Reporting Dashboard is available")
except ImportError as e:
    GAAP_DASHBOARD_AVAILABLE = False
    print(f"❌ GAAP Reporting Dashboard not available: {e}")

# Import Quotation Manager
try:
    from quotation_manager import QuotationManagerDialog
    QUOTATION_MANAGER_AVAILABLE = True
    print("✅ Quotation Manager is available")
except ImportError as e:
    QUOTATION_MANAGER_AVAILABLE = False
    print(f"❌ Quotation Manager not available: {e}")

class ScrollableFrame:
    """A scrollable frame that can be used in any dialog"""
    def __init__(self, parent):
        try:
            style = ttk.Style()
            canvas_bg = style.lookup("App.TFrame", "background") or style.lookup("TFrame", "background") or "#f8fafc"
        except Exception:
            canvas_bg = "#f8fafc"
        self.canvas = tk.Canvas(parent, highlightthickness=0, bd=0, background=canvas_bg)
        self.scrollbar = ttk.Scrollbar(parent, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self._window_item = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self._window_item, width=e.width))
        
        # Mouse wheel binding (scroll anywhere inside)
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        self.scrollable_frame.bind("<Enter>", self._bind_mousewheel)
        self.scrollable_frame.bind("<Leave>", self._unbind_mousewheel)
        try:
            top = parent.winfo_toplevel()
            top.bind("<MouseWheel>", self._on_mousewheel)
            top.bind("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
            top.bind("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))
        except Exception:
            pass
        
    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        # Linux support
        self.canvas.bind_all("<Button-4>", lambda e: self.canvas.yview_scroll(-1, "units"))
        self.canvas.bind_all("<Button-5>", lambda e: self.canvas.yview_scroll(1, "units"))
        
    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")
        
    def _on_mousewheel(self, event):
        try:
            d = int(event.delta)
        except Exception:
            d = 0
        steps = 0
        if abs(d) >= 120:
            steps = int(-d/120)
        else:
            if d < 0:
                steps = 1
            elif d > 0:
                steps = -1
        if steps != 0:
            self.canvas.yview_scroll(steps, "units")
        
    def pack(self, **kwargs):
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
    
class LogoManager:
    def __init__(self, data_folder=None):
        self.data_folder = data_folder
        # Look for logo in multiple locations
        possible_logo_paths = []
        
        if data_folder:
            possible_logo_paths.extend([
                os.path.join(data_folder, "logo.png"),
                os.path.join(data_folder, "logo.jpg"),
                os.path.join(data_folder, "logo.jpeg"),
                os.path.join(data_folder, "HopePharmaInvoice", "logo.png"),
                os.path.join(data_folder, "HopePharmaInvoices", "logo.png")
            ])
        
        desktop_path = os.path.join(os.path.expanduser("~"), "Documents")
        hope_pharma_folder = os.path.join(desktop_path, "HopePharmaInvoices")
        
        # Only create default folder if no data_folder provided or if explicitly requested
        if not data_folder and not os.path.exists(hope_pharma_folder):
            try:
                os.makedirs(hope_pharma_folder)
                print(f"Created HopePharmaInvoices folder: {hope_pharma_folder}")
            except Exception:
                pass
        
        # Check multiple possible locations for logo
        possible_logo_paths.extend([
            os.path.join(hope_pharma_folder, "logo.png"),
            os.path.join(hope_pharma_folder, "logo.jpg"),
            os.path.join(hope_pharma_folder, "logo.jpeg"),
            os.path.join(desktop_path, "logo.png"),
            os.path.join(desktop_path, "logo.jpg"),
            os.path.join(desktop_path, "logo.jpeg"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "HopePharma.app", "Contents", "Resources", "logo.png"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "logo.png"),
        ])
        
        self.logo_path = None
        for path in possible_logo_paths:
            if os.path.exists(path):
                self.logo_path = path
                print(f"Found logo at: {path}")
                break
        
        if not self.logo_path:
            # Default location
            self.logo_path = os.path.join(data_folder if data_folder else hope_pharma_folder, "logo.png")
            print(f"No logo file found. Defaulting to: {self.logo_path}")
            
        self.logo_image = None
        self.logo_photo = None
        
    def load_logo(self, size=(100, 100)):
        """Load and resize logo"""
        try:
            if os.path.exists(self.logo_path):
                image = Image.open(self.logo_path)
                image = image.resize(size, Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(image)
                self.logo_image = image
                print(f"✅ Logo loaded successfully from: {self.logo_path}")
                return True
            else:
                print(f"❌ Logo file not found at: {self.logo_path}")
                # Create a simple colored rectangle as placeholder
                image = Image.new('RGB', size, color='lightblue')
                self.logo_photo = ImageTk.PhotoImage(image)
                self.logo_image = image
                return False
        except Exception as e:
            print(f"❌ Error loading logo: {e}")
            # Create a simple colored rectangle as placeholder
            image = Image.new('RGB', size, color='lightblue')
            self.logo_photo = ImageTk.PhotoImage(image)
            self.logo_image = image
            return False
    
    def get_logo_for_html(self, size=(150, 150)):
        """Get logo as base64 for HTML embedding"""
        try:
            if os.path.exists(self.logo_path):
                from io import BytesIO
                import base64
                
                image = Image.open(self.logo_path)
                image = image.resize(size, Image.Resampling.LANCZOS)
                
                buffered = BytesIO()
                image.save(buffered, format="PNG")
                img_str = base64.b64encode(buffered.getvalue()).decode()
                return f"data:image/png;base64,{img_str}"
            return None
        except Exception as e:
            print(f"Error preparing logo for HTML: {e}")
            return None
    
    def check_signatures(self):
        """Check if signature files exist and return status"""
        hope_pharma_folder = None
        if hasattr(self, 'data_folder') and self.data_folder:
            # Check data folder first
            hope_pharma_folder = os.path.join(self.data_folder, "HopePharmaInvoices")
            if not os.path.exists(hope_pharma_folder):
                hope_pharma_folder = os.path.join(self.data_folder, "HopePharmaInvoice")
        
        if not hope_pharma_folder or not os.path.exists(hope_pharma_folder):
            # Fallback to Documents
            desktop_path = os.path.join(os.path.expanduser("~"), "Documents")
            hope_pharma_folder = os.path.join(desktop_path, "HopePharmaInvoices")

        sign1_path = os.path.join(hope_pharma_folder, "sign1.png")
        sign2_path = os.path.join(hope_pharma_folder, "sign2.png")
        
        has_sign1 = os.path.exists(sign1_path)
        has_sign2 = os.path.exists(sign2_path)
        
        print(f"Signature 1 (sign1.png): {'✅ Found' if has_sign1 else '❌ Missing'}")
        print(f"Signature 2 (sign2.png): {'✅ Found' if has_sign2 else '❌ Missing'}")
        
        return has_sign1, has_sign2

class DataMemoryManager:
    """Manages memory for clients, products, and other frequently used data"""
    
    def __init__(self, data_folder):
        self.data_folder = data_folder
        self.clients_file = os.path.join(data_folder, "clients_memory.json")
        self.products_file = os.path.join(data_folder, "products_memory.json")
        self.clients = self.load_clients()
        self.products = self.load_products()
    
    def load_clients(self):
        """Load clients from memory file"""
        if os.path.exists(self.clients_file):
            try:
                with open(self.clients_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading clients memory: {e}")
                return []
        return []
    
    def load_products(self):
        """Load products from memory file"""
        if os.path.exists(self.products_file):
            try:
                with open(self.products_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading products memory: {e}")
                return []
        return []
    
    def save_clients(self):
        """Save clients to memory file"""
        try:
            with open(self.clients_file, 'w') as f:
                json.dump(self.clients, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving clients memory: {e}")
            return False
    
    def save_products(self):
        """Save products to memory file"""
        try:
            with open(self.products_file, 'w') as f:
                json.dump(self.products, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving products memory: {e}")
            return False
    
    def add_client(self, client_name, client_trn=""):
        """Add a new client to memory"""
        client_data = {
            'name': client_name,
            'trn': client_trn,
            'first_added': datetime.now().strftime("%Y-%m-%d"),
            'last_used': datetime.now().strftime("%Y-%m-%d")
        }
        
        # Check if client already exists
        for client in self.clients:
            if client['name'].lower() == client_name.lower():
                client['last_used'] = datetime.now().strftime("%Y-%m-%d")
                if client_trn:
                    client['trn'] = client_trn
                return self.save_clients()
        
        self.clients.append(client_data)
        return self.save_clients()
    
    def add_product(self, product_name, default_price=0.0):
        """Add a new product to memory"""
        product_data = {
            'name': product_name,
            'default_price': default_price,
            'first_added': datetime.now().strftime("%Y-%m-%d"),
            'last_used': datetime.now().strftime("%Y-%m-%d")
        }
        
        # Check if product already exists
        for product in self.products:
            if product['name'].lower() == product_name.lower():
                product['last_used'] = datetime.now().strftime("%Y-%m-%d")
                if default_price > 0:
                    product['default_price'] = default_price
                return self.save_products()
        
        self.products.append(product_data)
        return self.save_products()
    
    def get_clients(self):
        """Get all clients sorted by last used"""
        return sorted(self.clients, key=lambda x: x['last_used'], reverse=True)
    
    def get_products(self):
        """Get all products sorted by last used"""
        return sorted(self.products, key=lambda x: x['last_used'], reverse=True)
    
    def search_clients(self, query):
        """Search clients by name"""
        query = query.lower()
        return [client for client in self.clients if query in client['name'].lower()]
    
    def search_products(self, query):
        """Search products by name"""
        query = query.lower()
        return [product for product in self.products if query in product['name'].lower()]

class EmployeeManager:
    def __init__(self, data_folder):
        self.data_folder = data_folder
        from employee_manager import EmployeeManager as _ProEmpMgr
        self._pro = _ProEmpMgr(data_folder)
        self.employees_file = os.path.join(data_folder, "employees.json")
        self.salary_reports_file = os.path.join(data_folder, "salary_reports.json")
        # --- Authoritative list is self._pro.employees; keep wrapper in sync ---
        self.employees = self._pro.employees  # SHARED REFERENCE (fixes "Add -> invisible")
        self.salary_reports = self.load_salary_reports()

    def _pro_salary_reports_view(self):
        compat = []
        try:
            for mk, entries in list(self._pro.salaries.items()):
                if not mk or mk.startswith("__") or not isinstance(entries, list):
                    continue
                total = 0.0
                salary_data = []
                for se in entries:
                    if not isinstance(se, dict):
                        continue
                    calc = se.get('calculated_values', {}) or {}
                    gross = float(calc.get('gross_salary') or se.get('salary_paid') or se.get('gross') or 0.0)
                    net = float(calc.get('net_salary') or se.get('salary_paid') or se.get('net') or 0.0)
                    paid = net if net > 0 else gross
                    total += paid
                    salary_data.append({
                        'employee_id': se.get('employee_id') or se.get('id') or '',
                        'name': se.get('name') or se.get('employee_name') or '',
                        'base_salary': se.get('salary') or se.get('base_salary') or gross,
                        'salary_paid': paid,
                    })
                compat.append({
                    'report_id': f"SAL-{mk.replace('-', '')}",
                    'month_year': mk,
                    'salary_data': salary_data,
                    'total_salary': round(total, 2),
                    'account': ((self._pro.salaries.get('__meta__', {}) or {}).get(mk, {}).get('account', 'Bank') or 'Bank'),
                    'created_date': ((self._pro.salaries.get('__meta__', {}) or {}).get(mk, {}).get('processed_date') or f"{mk}-28")
                })
        except Exception:
            pass
        try:
            if os.path.exists(self.salary_reports_file):
                with open(self.salary_reports_file, 'r') as f:
                    extra = json.load(f)
                if isinstance(extra, list):
                    seen = {r.get('month_year') for r in compat if isinstance(r, dict)}
                    for r in extra:
                        if isinstance(r, dict) and r.get('month_year') and r.get('month_year') not in seen:
                            compat.append(r)
        except Exception:
            pass
        return compat

    def set_ledger_service(self, ledger):
        try:
            return self._pro.set_ledger_service(ledger)
        except Exception:
            return None
    
    def load_employees(self):
        """Load employees (AUTHORITATIVE: from pro EmployeeManager.employees).
        Wrapper.employees is a shared ref so this is a no-op on happy path; we also
        re-read employees.json to import any legacy records not yet in pro list."""
        try:
            pro_list = self._pro.employees or []
            if os.path.exists(self.employees_file):
                with open(self.employees_file, 'r') as f:
                    legacy = json.load(f) or []
                pro_ids = {e.get("employee_id") for e in pro_list if isinstance(e, dict)}
                added_any = False
                for leg in legacy:
                    if not isinstance(leg, dict):
                        continue
                    eid = leg.get("employee_id")
                    if not eid or eid in pro_ids:
                        continue
                    # Import legacy record into pro list
                    self._pro.employees.append(leg)
                    pro_ids.add(eid)
                    added_any = True
                if added_any:
                    try:
                        self._pro._save_json(self._pro.employees_file, self._pro.employees)
                    except Exception:
                        pass
            self.employees = self._pro.employees
            return self._pro.employees
        except Exception as e:
            print(f"Error loading employees: {e}")
            return self._pro.employees or []
    
    def load_salary_reports(self):
        """Return compatibility view: all months from professional monthly_salaries.json,
        merged with any historical PDFs in salary_reports.json.  This gives ALL months
        when iterating self.salary_reports (previously it returned only 1 month's PDF)."""
        try:
            return self._pro_salary_reports_view()
        except Exception:
            # Fallback: old salary_reports.json
            if os.path.exists(self.salary_reports_file):
                try:
                    with open(self.salary_reports_file, 'r') as f:
                        return json.load(f)
                except Exception:
                    return []
            return []
    
    def save_employees(self):
        """Save employees (AUTHORITATIVE: via pro EmployeeManager._save_json).
        Both wrapper.employees and pro.employees share the same list object."""
        try:
            self._pro._save_json(self._pro.employees_file, self._pro.employees)
            # Also mirror to legacy employees.json for compatibility
            try:
                with open(self.employees_file, 'w') as f:
                    json.dump(self._pro.employees, f, indent=2, ensure_ascii=False)
            except Exception:
                pass
            return True
        except Exception as e:
            print(f"Error saving employees: {e}")
            return False
    
    def save_salary_reports(self):
        """Save salary reports to JSON file"""
        try:
            with open(self.salary_reports_file, 'w') as f:
                json.dump(self.salary_reports, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving salary reports: {e}")
            return False
    
    def add_employee(self, name, email, phone, salary, start_date, end_date=None,
                     residency_expiry=None, department=None, designation=None,
                     employee_type="Full-time", id_number=None, nationality=None,
                     joining_date=None, bank_name=None, bank_iban=None,
                     emergency_contact=None, emergency_phone=None, notes=None):
        """Add a new employee (AUTHORITATIVE: delegates to pro employee manager,
        so records are always visible in EmployeeManagerDialog.refresh_employees).

        Supports expanded professional HR fields in addition to basic ones."""
        # Build a professional profile dictionary
        profile = {
            "name": (name or "").strip(),
            "email": (email or "").strip(),
            "phone": (phone or "").strip(),
            "salary": float(salary or 0),
            "start_date": start_date,
            "end_date": end_date or None,
            "residency_expiry": residency_expiry or None,
            "department": (department or "").strip() if department else "",
            "designation": (designation or "").strip() if designation else "",
            "employee_type": (employee_type or "Full-time").strip(),
            "id_number": (id_number or "").strip() if id_number else "",
            "nationality": (nationality or "").strip() if nationality else "",
            "joining_date": joining_date or start_date,
            "bank_name": (bank_name or "").strip() if bank_name else "",
            "bank_iban": (bank_iban or "").strip() if bank_iban else "",
            "emergency_contact": (emergency_contact or "").strip() if emergency_contact else "",
            "emergency_phone": (emergency_phone or "").strip() if emergency_phone else "",
            "notes": (notes or "").strip() if notes else "",
            "status": "Active" if not end_date else "Inactive",
        }
        # Pass through to pro manager (creates ID, sets created_at, saves JSON)
        emp_id = self._pro.add_employee(profile)
        # wrapper.employees is same list object as pro.employees -> auto-updated
        # Double-save legacy employees.json for old readers
        try:
            self.save_employees()
        except Exception:
            pass
        return emp_id
    
    def update_employee(self, employee_id, **kwargs):
        """Update employee information.  DELEGATES to pro manager so updates
        are reflected everywhere including GAAP reports."""
        if 'end_date' in kwargs:
            kwargs['status'] = 'Active' if not kwargs.get('end_date') else 'Inactive'
        self._pro.update_employee(employee_id, kwargs)
        # Also save legacy file for old code paths
        return self.save_employees()
    
    def delete_employee(self, employee_id):
        """Delete employee.  Removes from pro list (mirrored to wrapper list via shared ref)."""
        self._pro.employees = [emp for emp in (self._pro.employees or [])
                               if isinstance(emp, dict) and emp.get('employee_id') != employee_id]
        self.employees = self._pro.employees
        return self.save_employees()
    
    def get_all_employees(self):
        """Get all employees — always pulls fresh from pro manager to avoid stale cache."""
        # Re-import any legacy records silently (safe idempotent)
        try:
            self.load_employees()
        except Exception:
            pass
        return list(self._pro.employees or [])
    
    def search_employees(self, query):
        """Search employees by name / ID / phone / email / department / designation / id_number."""
        query = str(query or "").lower().strip()
        results = []
        for employee in self.get_all_employees():
            if not isinstance(employee, dict):
                continue
            haystack = " ".join(str(v) for k, v in employee.items()
                                if k in ("name", "employee_id", "phone", "email",
                                         "department", "designation", "id_number",
                                         "nationality", "emergency_contact"))
            if query in haystack.lower():
                results.append(employee)
        return results
    
    def add_monthly_salary_report(self, month_year, salary_data, account=None):
        """
        Add monthly payroll.  NOW WRITES TO monthly_salaries.json (authoritative)
        AND calls save_month_entries (status='approved') so the real-time Ledger
        posting hook runs (Dr 5100 Salary / Cr 2100 Accrued + Cr 1100 Bank Net Pay).
        Also keeps old salary_reports.json cache for backward PDF compatibility.
        """
        report_id = f"SAL-{month_year.replace('-', '')}"
        total_salary = sum(float(item.get('salary_paid') or 0) for item in salary_data)
        # Convert legacy salary_data to professional month entries
        if isinstance(salary_data, list):
            entries = []
            for sd in salary_data:
                if not isinstance(sd, dict):
                    continue
                paid = float(sd.get('salary_paid') or 0)
                if paid <= 0:
                    continue
                entries.append({
                    'employee_id': sd.get('employee_id') or '',
                    'name': sd.get('name') or '',
                    'salary': sd.get('base_salary') or paid,
                    'status': 'approved',
                    'processed_date': datetime.now().strftime("%Y-%m-%d"),
                    'bank_details': ({'bank_name': str(account or 'Bank')}
                                     if account and 'ash' not in str(account).lower() else {}),
                    'salary_components': {'basic': paid},
                    'calculated_values': {
                        'gross_salary': paid,
                        'net_salary': paid,
                        'total_deductions': 0.0
                    },
                    'deductions': {},
                    'variable_pay': {},
                })
            try:
                self._pro.save_month_entries(str(month_year), entries, status='approved')
            except Exception:
                try:
                    self._pro.salaries[str(month_year)] = entries
                    meta = self._pro.salaries.get("__meta__", {}) or {}
                    meta[str(month_year)] = {
                        "status": "approved",
                        "processed_date": datetime.now().strftime("%Y-%m-%d"),
                        "account": account}
                    self._pro.salaries["__meta__"] = meta
                    self._pro._save_json(self._pro.salaries_file, self._pro.salaries)
                except Exception:
                    pass
        report = {
            'report_id': report_id,
            'month_year': month_year,
            'salary_data': list(salary_data) if isinstance(salary_data, list) else [],
            'total_salary': total_salary,
            'account': account or 'Cash',
            'created_date': datetime.now().strftime("%Y-%m-%d")
        }
        # Refresh compatibility view AND write to old salary_reports.json for PDF archiving
        try:
            self.salary_reports = self._pro_salary_reports_view()
        except Exception:
            self.salary_reports = []
        try:
            existing = []
            if os.path.exists(self.salary_reports_file):
                with open(self.salary_reports_file, 'r') as f:
                    existing = json.load(f)
                if not isinstance(existing, list):
                    existing = []
        except Exception:
            existing = []
        existing = [r for r in existing if isinstance(r, dict) and r.get('month_year') != str(month_year)]
        existing.append(report)
        try:
            with open(self.salary_reports_file, 'w') as f:
                json.dump(existing, f, indent=2)
        except Exception:
            pass
        return report_id
    
    def get_monthly_report(self, month_year):
        """Get monthly salary report"""
        for report in self.salary_reports:
            if report['month_year'] == month_year:
                return report
        return None
    
    def get_yearly_report(self, year):
        """Get yearly salary report"""
        yearly_reports = [r for r in self.salary_reports if r['month_year'].startswith(year)]
        total_salary = sum(r['total_salary'] for r in yearly_reports)
        
        return {
            'year': year,
            'reports': yearly_reports,
            'total_salary': total_salary,
            'report_count': len(yearly_reports)
        }

class EmployeePasswordDialog:
    def __init__(self, parent):
        self.parent = parent
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Employee Manager - Authentication")
        self.dialog.geometry("420x240")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.result = False
        self.setup_ui()
        _fit_window(self.dialog, 420, 240, mode="compact", remember_key="employee_password")
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(main_frame, text="Enter Password", 
                 font=('Helvetica', 12, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Password:").pack(anchor='w', pady=(5, 0))
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(main_frame, textvariable=self.password_var, show="*", width=20)
        self.password_entry.pack(fill='x', pady=5)
        self.password_entry.bind('<Return>', lambda e: self.check_password())
        try:
            self.dialog.bind('<Return>', lambda e: self.check_password())
        except Exception:
            pass
        ttk.Label(main_frame, text="Press Enter to login", style="TopbarMeta.TLabel").pack(anchor='w', pady=(0, 8))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        ttk.Button(button_frame, text="Login", 
                  command=self.check_password).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side='left', padx=5)
    
    def check_password(self):
        """Check if password is correct"""
        if self.password_var.get() == "5410954":
            self.result = True
            try:
                messagebox.showinfo("Success", "Successfully logged in")
            except Exception:
                pass
            self.dialog.destroy()
        else:
            try:
                messagebox.showerror("Error", "Incorrect password!")
            except Exception:
                pass
            self.password_var.set("")
            self.password_entry.focus()

class LegacyEmployeeManagerDialog:
    def __init__(self, parent, manager):
        self.manager = manager
        # Get the data folder from invoice_folder attribute
        self.employee_manager = EmployeeManager(manager.invoice_folder)
        # Wire real-time payroll postings to the Ledger Service (GAAP double-entry)
        ledger_svc = getattr(self.manager, "ledger", None)
        if ledger_svc is None:
            dm = getattr(self.manager, "dm", None)
            if dm is not None:
                ledger_svc = getattr(dm, "ledger", None)
        if ledger_svc is None:
            try:
                import hope_pharma_complete as hpc
                import os
                dm_cls = getattr(hpc, "EnhancedCloudDataManager", None) or getattr(hpc, "HopePharmaDataManager", None)
                if dm_cls and callable(dm_cls):
                    inst = dm_cls(manager.invoice_folder, mode='local')
                    ledger_svc = getattr(inst, "ledger", None)
                    if ledger_svc is None:
                        from ledger_service import LedgerService
                        ledger_svc = LedgerService(inst)
            except Exception:
                ledger_svc = None
        if ledger_svc is not None:
            self.employee_manager.set_ledger_service(ledger_svc)
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Employee Management")
        self.dialog.geometry("1200x700")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        self.refresh_employees()
        _fit_window(self.dialog, 1100, 700, mode="workspace", remember_key="employee_manager")
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="👥 Employee Management",
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 4))
        self.emp_summary_var = tk.StringVar(value="Computing payroll summary…")
        ttk.Label(main_frame, textvariable=self.emp_summary_var,
                  wraplength=1050, justify="left",
                  font=('Helvetica', 10, 'italic')).pack(pady=(0, 14))

        # Quick Period Filter (determines what totals appear in banner)
        period_bar = ttk.Frame(main_frame)
        period_bar.pack(fill='x', pady=(0, 8))
        ttk.Label(period_bar, text="Period:", font=('Helvetica', 10, 'bold')).pack(side='left', padx=(0, 6))
        self.emp_period_var = tk.StringVar(value="CYTD")
        for label, value in [("This MTD", "MTD"), ("Prev Month", "PM"), ("CY YTD", "CYTD"), ("All Time", "ALL")]:
            rb = ttk.Radiobutton(period_bar, text=label, value=value, variable=self.emp_period_var,
                                 command=self._refresh_emp_summary)
            rb.pack(side='left', padx=3)
        ttk.Label(period_bar, text="   |   ").pack(side='left')
        ttk.Button(period_bar, text="📅 Payroll Register (Export CSV)",
                   command=self._export_payroll_csv).pack(side='left', padx=4)
        ttk.Button(period_bar, text="🔁 Sync to Reports / Rebuild GL",
                   command=self._emp_rebuild_gl).pack(side='left', padx=4)

        # KPI Payroll Summary Cards (4 cards: Total Employees | Gross Salaries Paid | Net Cash Out | Months Processed)
        cards_frame = tk.Frame(main_frame, bg="#FFFFFF")
        cards_frame.pack(fill='x', pady=(0, 12))
        cards_frame.columnconfigure(0, weight=1)
        cards_frame.columnconfigure(1, weight=1)
        cards_frame.columnconfigure(2, weight=1)
        cards_frame.columnconfigure(3, weight=1)
        self._emp_card_labels = {}
        today = date.today()
        card_defs = [
            ("EMPLOYEES", "emp_count", "#1E3A5F", "On Payroll"),
            ("GROSS SALARIES", "emp_gross", "#0B6623", "Period Total"),
            ("NET SALARIES PAID", "emp_net", "#7C2D12", "Cash / Bank Out"),
            ("MONTHS PROCESSED", "emp_months", "#4C1D95", "In Period"),
        ]
        for ci, (title, key, color, sub) in enumerate(card_defs):
            card = tk.Frame(cards_frame, bg="white", highlightbackground="#E5E7EB",
                            highlightthickness=1, bd=0)
            card.grid(row=0, column=ci, sticky="nsew", padx=6, pady=2)
            tk.Label(card, text=title, fg=color, bg="white",
                     font=('Helvetica', 9, 'bold')).pack(anchor="w", padx=12, pady=(10, 2))
            val_label = tk.Label(card, text="-", fg="#111827", bg="white",
                                 font=('Helvetica', 18, 'bold'))
            val_label.pack(anchor="w", padx=12, pady=0)
            tk.Label(card, text=sub, fg="#6B7280", bg="white",
                     font=('Helvetica', 9)).pack(anchor="w", padx=12, pady=(2, 10))
            self._emp_card_labels[key] = val_label

        # Search and controls
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(controls_frame, text="Search:").pack(side='left', padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(controls_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side='left', padx=(0, 10))
        self.search_entry.bind('<KeyRelease>', self.on_search)
        
        # Sort options
        ttk.Label(controls_frame, text="Sort by:").pack(side='left', padx=(20, 5))
        self.sort_var = tk.StringVar(value="name_asc")
        sort_combo = ttk.Combobox(controls_frame, textvariable=self.sort_var, 
                                 values=["Name (A-Z)", "Name (Z-A)", "ID (Asc)", "ID (Desc)", "Date (New-Old)", "Date (Old-New)"],
                                 state="readonly", width=15)
        sort_combo.pack(side='left', padx=(0, 10))
        sort_combo.bind('<<ComboboxSelected>>', self.on_sort_change)
        
        # Buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.pack(side='right')
        
        ttk.Button(button_frame, text="➕ Add Employee", 
                  command=self.add_employee).pack(side='left', padx=5)
        ttk.Button(button_frame, text="✏️ Edit Employee", 
                  command=self.edit_employee).pack(side='left', padx=5)
        ttk.Button(button_frame, text="🗑️ Delete Employee", 
                  command=self.delete_employee).pack(side='left', padx=5)
        ttk.Button(button_frame, text="💰 Monthly Salary", 
                  command=self.monthly_salary_report).pack(side='left', padx=5)
        ttk.Button(button_frame, text="📊 Yearly Report", 
                  command=self.yearly_report).pack(side='left', padx=5)
        ttk.Button(button_frame, text="🔄 Refresh", 
                  command=self.refresh_employees).pack(side='left', padx=5)
        
        # Employees list
        list_frame = ttk.LabelFrame(main_frame, text="Employees — Personnel Registry", padding="10")
        list_frame.pack(fill='both', expand=True)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill='both', expand=True)
        
        columns = ('ID', 'Name', 'Department', 'Designation', 'Type', 'Nationality',
                   'ID Number', 'Salary', 'Joining', 'Start', 'End', 'Residency',
                   'Status')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Define headings
        self.tree.heading('ID', text='Emp ID')
        self.tree.heading('Name', text='Full Name')
        self.tree.heading('Department', text='Department')
        self.tree.heading('Designation', text='Job Title / Designation')
        self.tree.heading('Type', text='Emp Type')
        self.tree.heading('Nationality', text='Nationality')
        self.tree.heading('ID Number', text='Iqama / Passport')
        self.tree.heading('Salary', text='Gross Salary (AED)')
        self.tree.heading('Joining', text='Joining Date')
        self.tree.heading('Start', text='Start Date')
        self.tree.heading('End', text='End Date')
        self.tree.heading('Residency', text='Visa / Residency Expiry')
        self.tree.heading('Status', text='Status')
        
        # Define columns (right-align currency/date, width each field)
        self.tree.column('ID',           width=90,   anchor='center')
        self.tree.column('Name',         width=200,  anchor='w')
        self.tree.column('Department',   width=120,  anchor='w')
        self.tree.column('Designation',  width=180,  anchor='w')
        self.tree.column('Type',         width=95,   anchor='center')
        self.tree.column('Nationality',  width=110,  anchor='w')
        self.tree.column('ID Number',    width=140,  anchor='w')
        self.tree.column('Salary',       width=140,  anchor='e')
        self.tree.column('Joining',      width=105,  anchor='center')
        self.tree.column('Start',        width=105,  anchor='center')
        self.tree.column('End',          width=105,  anchor='center')
        self.tree.column('Residency',    width=150,  anchor='center')
        self.tree.column('Status',       width=190,  anchor='w')
        
        # Scrollbars for treeview (horizontal + vertical since 13+ columns)
        tree_scroll_y = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        tree_scroll_x = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        self.tree.pack(side='top', fill='both', expand=True)
        tree_scroll_y.pack(side='right', fill='y')
        tree_scroll_x.pack(side='bottom', fill='x')
        
        # Status
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var,
                              relief=tk.SUNKEN, font=('Helvetica', 10))
        status_bar.pack(fill='x', side='bottom', pady=(10, 0))

    def _get_period_range(self):
        today = date.today()
        mode = str(self.emp_period_var.get() or 'CYTD').upper()
        y, m = today.year, today.month
        if mode == 'MTD':
            sd = date(y, m, 1).isoformat()
            ed = today.isoformat()
        elif mode == 'PM':
            if m == 1:
                ly, lm = y - 1, 12
            else:
                ly, lm = y, m - 1
            last = (date(y, m, 1) - timedelta(days=1))
            sd = date(ly, lm, 1).isoformat()
            ed = last.isoformat()
        elif mode == 'ALL':
            sd = date(2000, 1, 1).isoformat()
            ed = date(2099, 12, 31).isoformat()
        else:
            sd = date(y, 1, 1).isoformat()
            ed = today.isoformat()
        return sd, ed

    def _refresh_emp_summary(self):
        try:
            em = self.employee_manager
            employees = list(em.get_all_employees() or [])
            active_count = sum(1 for e in employees if str(e.get('status', '')).lower() in ('', 'active'))
            try:
                sd, ed = self._get_period_range()
                s = em.get_salary_summary_for_period(sd, ed)
                gross = float(s["gross_salary_total"])
                net = float(s["net_salary_total"])
                months = int(s["months_in_period"])
            except Exception:
                s = {"gross_salary_total": 0, "net_salary_total": 0, "months_in_period": 0}
                gross = net = 0; months = 0
            self._emp_card_labels["emp_count"].configure(text=str(len(employees)))
            self._emp_card_labels["emp_gross"].configure(text=f"AED {gross:,.2f}")
            self._emp_card_labels["emp_net"].configure(text=f"AED {net:,.2f}")
            self._emp_card_labels["emp_months"].configure(text=f"{months}")
            sd, ed = self._get_period_range()
            self.emp_summary_var.set(
                f"Period {sd} → {ed}  ·  Active Employees: {active_count}  ·  "
                f"Gross AED {gross:,.2f}  ·  Net Cash Out AED {net:,.2f}.  "
                f"Payroll auto-posts to GL (5100 Salary Expense Dr / 2100+1100 Cr).  "
                f"Click 'Sync to Reports' to push into all 10 GAAP reports.")
            self.status_var.set(
                f"Ready · {len(employees)} total employees · Payroll processed: {months} months in period")
        except Exception as ex:
            self.status_var.set(f"Summary compute: {ex}")

    def _export_payroll_csv(self):
        try:
            em = self.employee_manager
            sd, ed = self._get_period_range()
            s = em.get_salary_totals_for_period(sd, ed, use_field="net_salary")
            by_month = s.get("by_month") or {}
            default_path = os.path.join(str(self.employee_manager.data_folder),
                                        f"Payroll_Register_{sd}_{ed}.csv")
            import csv
            with open(default_path, 'w', newline='', encoding='utf-8') as f:
                w = csv.writer(f)
                w.writerow(['Month', 'Employees', 'Gross (AED)', 'Net (AED)'])
                for mk in sorted(by_month.keys()):
                    bm = by_month[mk]
                    w.writerow([mk, bm.get('employees', 0),
                                f"{float(bm.get('gross',0)):.2f}",
                                f"{float(bm.get('net',0)):.2f}"])
                gross_sum = sum(float(by_month[m].get('gross', 0)) for m in by_month)
                net_sum = sum(float(by_month[m].get('net', 0)) for m in by_month)
                w.writerow([])
                w.writerow(['TOTAL', s.get('employees_processed', 0), f"{gross_sum:.2f}", f"{net_sum:.2f}"])
            try:
                messagebox.showinfo("Exported", f"Payroll Register CSV exported to:\n{default_path}")
            except Exception:
                pass
        except Exception as ex:
            try: messagebox.showerror("Export Error", str(ex))
            except Exception: pass

    def _emp_rebuild_gl(self):
        ledger_svc = getattr(self.manager, "ledger", None)
        if ledger_svc is None:
            dm = getattr(self.manager, "dm", None)
            if dm is None:
                try:
                    import hope_pharma_complete as hpc
                    dm_cls = getattr(hpc, "EnhancedCloudDataManager", None) or getattr(hpc, "HopePharmaDataManager", None)
                    if dm_cls and callable(dm_cls):
                        dm = dm_cls(self.manager.invoice_folder, mode='local')
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
        if ledger_svc is None:
            messagebox.showinfo("Rebuild GL",
                                "Ledger not available.  Open the GAAP Reports dashboard and click Rebuild GL.")
            return
        try:
            ans = messagebox.askyesno(
                "Confirm: Rebuild GL from All Data",
                "This will re-post every invoice, payment, purchase, expense, AND all payroll\n"
                "months into the General Ledger.  All existing GL entries are replaced.\n\n"
                "This operation is safe (idempotent), but takes ~1-2s per 100 invoices.\n"
                "Continue?")
            if not ans: return
            stats = ledger_svc.rebuild_general_ledger_from_all_data()
            self._refresh_emp_summary()
            messagebox.showinfo(
                "Rebuild Complete",
                f"Posted to GL:\n"
                f"  • {stats.get('invoices',0)} invoices\n"
                f"  • {stats.get('payments',0)} payments\n"
                f"  • {stats.get('purchases',0)} purchases\n"
                f"  • {stats.get('expenses',0)} expenses\n"
                f"  • {stats.get('salaries',0)} employee payroll months\n"
                f"GL now has {stats.get('total_gl_after',0)} double-entry rows.\n\n"
                f"All 10 GAAP financial reports now reflect every payroll run.")
        except Exception as ex:
            messagebox.showerror("Rebuild Failed", str(ex))

    def refresh_employees(self):
        """Refresh the employees list (13 professional columns + status tagging)"""
        for item in self.tree.get_children():
            self.tree.delete(item)

        # zebra stripe + semantic status tag config (with colored backgrounds for chips)
        try:
            self.tree.tag_configure('odd',  background='#FAFBFC')
            self.tree.tag_configure('even', background='#FFFFFF')
            self.tree.tag_configure('status_active',  foreground='#059669', font=('Helvetica', 9, 'bold'))
            self.tree.tag_configure('status_inactive',foreground='#6B7280')
            self.tree.tag_configure('status_leave',   foreground='#B45309', font=('Helvetica', 9, 'bold'))
            self.tree.tag_configure('status_expired', foreground='#991B1B', background='#FEF2F2',
                                    font=('Helvetica', 9, 'bold'))
            self.tree.tag_configure('status_visa_soon', foreground='#B45309', background='#FFFBEB',
                                    font=('Helvetica', 9, 'bold'))
            self.tree.tag_configure('salary_zero', foreground='#9CA3AF', font=('Helvetica', 9, 'italic'))
        except Exception:
            pass

        employees = self.employee_manager.get_all_employees()
        employees = self.sort_employees(employees, self.sort_var.get())

        today = date.today()
        # Date helpers for visa-expiry warning chip
        def days_until_exp(d_str):
            try:
                d = datetime.strptime(str(d_str).strip()[:10], "%Y-%m-%d").date()
                return (d - today).days
            except Exception:
                return None

        for idx, employee in enumerate(employees):
            if not isinstance(employee, dict):
                continue
            st_raw = str(employee.get('status', '') or '').lower()
            status_txt = str(employee.get('status', '') or '')
            if not status_txt:
                status_txt = 'Active'
            status_tag = 'status_active'
            if 'inactive' in st_raw or 'terminated' in st_raw or 'resigned' in st_raw:
                status_tag = 'status_inactive'
                if not status_txt: status_txt = 'Inactive'
            elif 'leave' in st_raw:
                status_tag = 'status_leave'
            elif 'expir' in st_raw:
                status_tag = 'status_expired'
            # Auto-flag residency expiry / visa chip
            rex = str(employee.get('residency_expiry', '') or '').strip()[:10]
            days_re = days_until_exp(rex) if rex and rex != 'N/A' else None
            if days_re is not None:
                if days_re < 0:
                    if status_tag not in ('status_inactive',):
                        status_tag = 'status_expired'
                    status_txt = f"● {status_txt} · Visa Expired ({abs(days_re)}d ago)"
                elif days_re <= 30:
                    if status_tag == 'status_active':
                        status_tag = 'status_visa_soon'
                    status_txt = f"● {status_txt} · Visa expires in {days_re}d"
                else:
                    status_txt = f"● {status_txt} · Visa valid {days_re}d"
            else:
                status_txt = f"● {status_txt}"
            # Add end-date chip
            end_dt = str(employee.get('end_date', '') or '').strip()[:10]
            if end_dt and end_dt != 'N/A':
                de = days_until_exp(end_dt)
                if de is not None:
                    if de < 0:
                        status_txt += " · Contract Ended"
                    elif de <= 30:
                        status_txt += f" · Contract ends in {de}d"
            # Salary chip styling
            extra_tags = ()
            try:
                sal = float(employee.get('salary') or 0)
                if sal <= 0:
                    extra_tags = ('salary_zero',)
            except Exception:
                extra_tags = ('salary_zero',)
            stripetag = ('odd' if idx % 2 else 'even')
            self.tree.insert('', tk.END,
                             tags=(stripetag, status_tag) + extra_tags,
                             values=(
                employee.get('employee_id', ''),
                employee.get('name', ''),
                employee.get('department', '') or '-',
                employee.get('designation', '') or '-',
                employee.get('employee_type', '') or 'Full-time',
                employee.get('nationality', '') or '-',
                employee.get('id_number', '') or '-',
                f"AED {float(employee.get('salary') or 0):,.2f}",
                employee.get('joining_date', '') or employee.get('start_date', '') or '-',
                employee.get('start_date', '') or '-',
                employee.get('end_date', '') or '-',
                rex if rex else '-',
                status_txt,
            ))

        self.status_var.set(f"Ready · Total Employees: {len(employees)}")
        try:
            self._refresh_emp_summary()
        except Exception:
            pass
    
    def sort_employees(self, employees, sort_option):
        """Sort employees based on selected option"""
        if sort_option == "Name (A-Z)":
            return sorted(employees, key=lambda x: x.get('name', '').lower())
        elif sort_option == "Name (Z-A)":
            return sorted(employees, key=lambda x: x.get('name', '').lower(), reverse=True)
        elif sort_option == "ID (Asc)":
            return sorted(employees, key=lambda x: x.get('employee_id', ''))
        elif sort_option == "ID (Desc)":
            return sorted(employees, key=lambda x: x.get('employee_id', ''), reverse=True)
        elif sort_option == "Date (New-Old)":
            return sorted(employees, key=lambda x: x.get('start_date', ''), reverse=True)
        elif sort_option == "Date (Old-New)":
            return sorted(employees, key=lambda x: x.get('start_date', ''))
        else:
            return employees
    
    def on_sort_change(self, event=None):
        """Handle sort change"""
        self.refresh_employees()
    
    def on_search(self, event=None):
        """Handle search functionality (13 columns matching refresh_employees layout)."""
        query = self.search_var.get().strip()
        if not query:
            self.refresh_employees()
            return
            
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        results = self.employee_manager.search_employees(query)
        results = self.sort_employees(results, self.sort_var.get())
        
        for employee in results:
            if not isinstance(employee, dict):
                continue
            rex = str(employee.get('residency_expiry', '') or '').strip()[:10]
            status_txt = str(employee.get('status', '') or 'Active')
            st_raw = status_txt.lower()
            status_tag = 'status_active'
            if 'inactive' in st_raw or 'terminated' in st_raw or 'resigned' in st_raw:
                status_tag = 'status_inactive'
            elif 'leave' in st_raw: status_tag = 'status_leave'
            elif 'expir' in st_raw: status_tag = 'status_expired'
            extra_tags = ()
            try:
                if float(employee.get('salary') or 0) <= 0:
                    extra_tags = ('salary_zero',)
            except Exception:
                extra_tags = ('salary_zero',)
            self.tree.insert('', tk.END,
                             tags=('odd', status_tag) + extra_tags,
                             values=(
                employee.get('employee_id', ''),
                employee.get('name', ''),
                employee.get('department', '') or '-',
                employee.get('designation', '') or '-',
                employee.get('employee_type', '') or 'Full-time',
                employee.get('nationality', '') or '-',
                employee.get('id_number', '') or '-',
                f"AED {float(employee.get('salary') or 0):,.2f}",
                employee.get('joining_date', '') or employee.get('start_date', '') or '-',
                employee.get('start_date', '') or '-',
                employee.get('end_date', '') or '-',
                rex if rex else '-',
                f"● {status_txt}",
            ))
        
        self.status_var.set(f"Found {len(results)} employees matching '{query}'")
    
    def _employee_form_sections(self):
        """Return ordered list of (section_title, field_defs[]) where each
        field_def = (label, key, kind, placeholder_or_values).  Used by
        both Add and Edit Employee dialogs to keep them in sync."""
        today = date.today().isoformat()
        default_joining = today
        sections = [
            ("👤  Personal Information", [
                ("Full Name *",        "name",               "entry",     "e.g. John Smith"),
                ("Nationality",        "nationality",        "entry",     "e.g. Indian / UAE"),
                ("Iqama / Passport #", "id_number",          "entry",     "Iqama or passport number"),
                ("Date of Birth",      "date_of_birth",      "entry",     "YYYY-MM-DD"),
                ("Gender",             "gender",             "combo",     ["-", "Male", "Female", "Prefer not to say"]),
                ("Marital Status",     "marital_status",     "combo",     ["-", "Single", "Married", "Divorced", "Widowed"]),
            ]),
            ("💼  Employment", [
                ("Department",        "department",     "entry",     "e.g. Sales, Pharmacy, Admin, Finance"),
                ("Designation / Job Title *", "designation", "entry", "e.g. Pharmacist, Sales Executive, Accountant"),
                ("Employee Type",     "employee_type",  "combo",
                    ["Full-time", "Part-time", "Contract", "Temporary", "Intern", "Consultant"]),
                ("Joining Date *",    "joining_date",   "entry",     default_joining),
                ("Start Date *",      "start_date",     "entry",     default_joining),
                ("Probation End",     "probation_end",  "entry",     "YYYY-MM-DD"),
                ("Contract End Date", "end_date",       "entry",     "YYYY-MM-DD (leave blank for permanent)"),
                ("Reporting Manager", "reporting_to",   "entry",     "Manager name"),
                ("Work Location",     "work_location",  "entry",     "e.g. Dubai HQ, Sharjah Branch"),
            ]),
            ("💳  Payroll & Banking", [
                ("Gross Salary (AED) *", "salary",      "entry",     "Monthly gross salary (e.g. 8000)"),
                ("Housing Allowance",    "allowance_housing",   "entry", "AED (0 if included in base)"),
                ("Transport Allowance",  "allowance_transport", "entry", "AED"),
                ("Other Allowances",     "allowance_other",     "entry", "AED"),
                ("Bank Name",            "bank_name",           "entry", "e.g. Emirates NBD, ADCB"),
                ("Bank IBAN / Account",  "bank_iban",           "entry", "IBAN or account number"),
                ("Payment Mode",         "payment_mode",        "combo",
                    ["Bank Transfer", "Cheque", "Cash", "WPS"]),
                ("Pay Cycle",            "pay_cycle",           "combo",
                    ["Monthly", "Bi-weekly", "Weekly", "Project-based"]),
            ]),
            ("📞  Contact & Emergency", [
                ("Email",                "email",              "entry", "Work or personal email"),
                ("Mobile Phone *",       "phone",              "entry", "e.g. +971 50 XXX XXXX"),
                ("Home Address",         "address",            "entry", "Full residential address"),
                ("Emergency Contact",    "emergency_contact",  "entry", "Contact person name"),
                ("Emergency Phone",      "emergency_phone",    "entry", "Phone number"),
                ("Relationship to Emp",  "emergency_relation", "entry", "e.g. Spouse / Sibling / Father"),
            ]),
            ("🛂  Visa / Residency", [
                ("Residency / Visa Expiry",  "residency_expiry",  "entry",   "YYYY-MM-DD"),
                ("Visa Type",                "visa_type",         "combo",   ["-", "Employment Visa", "Family Visa", "Golden Visa", "Visit Visa", "Freelance"]),
                ("Work Permit #",            "work_permit",       "entry",   "MOHRE work permit number"),
                ("Labour Card #",            "labour_card",       "entry",   "MOHRE labour card"),
                ("Health Insurance #",       "insurance_id",      "entry",   "Insurance member ID"),
                ("Insurance Provider",       "insurance_provider","entry",   "e.g. Daman, NEXtcare"),
            ]),
            ("📝  Notes", [
                ("Additional Notes", "notes", "text", "Any HR notes / disciplinary / appraisal (optional)"),
            ]),
        ]
        return sections

    def _render_employee_form(self, form_parent, prefill=None):
        """Render the professional employee form.

        Returns a dict: widgets -> {key: widget}, plus a list of warnings_labels for inline errors.
        """
        prefill = prefill or {}
        widgets = {}
        sections = self._employee_form_sections()
        form_parent.columnconfigure(0, weight=1)
        form_parent.columnconfigure(1, weight=1)
        total_row = 0
        for si, (section_title, fields) in enumerate(sections):
            # ---- Section Card (subtle background) ----
            card = tk.Frame(form_parent, bg="#FFFFFF",
                            highlightbackground="#E5E7EB",
                            highlightthickness=1, bd=0)
            # 2-column card layout
            card.grid(row=total_row, column=0, columnspan=2, sticky="ew",
                      padx=4, pady=(14 if si else 0, 4))
            card.columnconfigure(0, weight=1)
            card.columnconfigure(1, weight=1)
            # Section heading
            hd = tk.Label(card, text=section_title,
                          font=("Helvetica", 11, "bold"),
                          fg="#1E3A5F", bg="WHITE", anchor="w")
            hd.grid(row=0, column=0, columnspan=2, sticky="ew",
                    padx=12, pady=(10, 6))
            tk.Frame(card, height=2, bg="#DBEAFE", bd=0).grid(
                row=1, column=0, columnspan=2, sticky="ew", padx=12, pady=(0, 8))
            # Fields grid within card: 2 columns
            field_row = 2
            left = True
            for idx, (label, key, kind, extra) in enumerate(fields):
                col = 0 if left else 1
                left = not left
                cell = tk.Frame(card, bg="WHITE")
                cell.grid(row=field_row, column=col, sticky="ew",
                          padx=10, pady=4)
                cell.columnconfigure(0, weight=1)
                required_marker = " *" if (key in ("name","designation","joining_date","start_date","salary","phone")) and key != "-" else ""
                required_marker = required_marker if (label.endswith("*") or required_marker.strip()) else ""
                lbl = tk.Label(cell, text=label.rstrip("*") + required_marker,
                               font=("Helvetica", 9, "bold"), fg="#374151", bg="WHITE", anchor="w")
                lbl.pack(fill="x")
                # hint / placeholder row
                hint_lbl = tk.Label(cell,
                                    text=(("Example: " + extra) if kind != "combo" and isinstance(extra, str) else ""),
                                    font=("Helvetica", 8, "italic"), fg="#9CA3AF", bg="WHITE", anchor="w")
                hint_lbl.pack(fill="x")
                if kind == "entry":
                    e = ttk.Entry(cell)
                    pre_val = prefill.get(key, '')
                    if pre_val is None: pre_val = ''
                    if key in ("joining_date","start_date") and not str(pre_val).strip():
                        pre_val = date.today().isoformat()
                    try:
                        if key == "salary":
                            e.insert(0, (f"{float(pre_val):,.2f}" if float(pre_val or 0) > 0 else ""))
                        else:
                            e.insert(0, str(pre_val))
                    except Exception:
                        e.insert(0, str(pre_val) if pre_val is not None else "")
                    e.pack(fill="x", pady=(2, 0))
                    widgets[key] = e
                elif kind == "combo":
                    values = list(extra)
                    val = str(prefill.get(key, '') or '').strip()
                    if val and val not in values:
                        values.insert(0, val)
                    if not values:
                        values = ["-"]
                    cb = ttk.Combobox(cell, values=values, state="readonly")
                    if val and val in values:
                        cb.set(val)
                    elif not val and values:
                        cb.set(values[0])
                    cb.pack(fill="x", pady=(2, 0))
                    widgets[key] = cb
                elif kind == "text":
                    t = tk.Text(cell, height=3, bd=1, relief="solid",
                                bg="#FCFCFC", fg="#111827", font=("Helvetica", 10))
                    pre_val = prefill.get(key, '') or ''
                    t.insert("1.0", str(pre_val))
                    t.pack(fill="x", pady=(2, 0))
                    widgets[key] = t
                else:
                    e = ttk.Entry(cell)
                    e.pack(fill="x", pady=(2, 0))
                    widgets[key] = e
                if not left:
                    field_row += 1
            # If odd number of fields, increment
            if len(fields) % 2 == 1:
                field_row += 1
            total_row += 1
        return widgets

    def _collect_employee_form(self, widgets):
        """Extract values from the form widgets dict and validate.
        Returns (data_dict, errors_list)."""
        errors = []
        data = {}
        for key, w in widgets.items():
            try:
                if isinstance(w, tk.Text):
                    v = (w.get("1.0", "end") or "").strip()
                elif isinstance(w, ttk.Combobox):
                    v = (w.get() or "").strip()
                    if v == "-":
                        v = ""
                else:
                    v = (w.get() or "").strip()
            except Exception:
                v = ""
            # Normalize salary (clean ",")
            if key == "salary":
                try:
                    cleaned = str(v).replace(",", "").strip()
                    v = float(cleaned) if cleaned else 0.0
                except ValueError:
                    errors.append(f"Invalid Salary amount: '{v}'. Use digits and optional decimals (e.g. 8000.00).")
                    v = 0.0
                if v <= 0:
                    errors.append("Gross Salary must be a positive number.")
            # Date validations
            for dk in ("joining_date", "start_date", "end_date",
                       "probation_end", "residency_expiry", "date_of_birth"):
                if key == dk and v:
                    try:
                        datetime.strptime(str(v).strip()[:10], "%Y-%m-%d")
                    except Exception:
                        errors.append(f"{dk.replace('_',' ').title()} '{v}' should be YYYY-MM-DD.")
            if key == "phone":
                if not v:
                    errors.append("Mobile Phone is required.")
            if key == "name":
                if not v:
                    errors.append("Full Name is required.")
            if key == "designation":
                if not v:
                    errors.append("Designation / Job Title is required.")
            if key == "start_date":
                if not v:
                    errors.append("Start Date is required.")
            if key == "joining_date":
                if not v:
                    v = data.get("start_date") or date.today().isoformat()
            if not isinstance(v, (int, float)):
                v = str(v)
            data[key] = v
        # Cross-field checks
        endd = data.get("end_date")
        startd = data.get("start_date")
        if endd and startd:
            try:
                if datetime.strptime(str(endd)[:10],"%Y-%m-%d") < datetime.strptime(str(startd)[:10],"%Y-%m-%d"):
                    errors.append("Contract End Date cannot be earlier than Start Date.")
            except Exception:
                pass
        return data, errors

    def add_employee(self):
        """Open professional Add Employee dialog with 6 section cards + 30 fields + inline validation."""
        dialog = tk.Toplevel(self.dialog)
        dialog.title("➕  Add Employee — Professional Personnel File")
        dialog.geometry("980x780")
        dialog.transient(self.dialog)
        dialog.resizable(True, True)

        main_container = tk.Frame(dialog)
        main_container.pack(fill='both', expand=True)

        # Title row
        header = tk.Frame(main_container, bg="WHITE",
                          highlightbackground="#E5E7EB",
                          highlightthickness=1, bd=0)
        header.pack(fill="x", padx=10, pady=(10, 4))
        tk.Label(header, text="➕  Create New Employee Personnel File",
                 font=("Helvetica", 15, "bold"), fg="#1E3A5F",
                 bg="WHITE").pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(header,
                 text="Fill in at least required fields (*) and click Save. "
                      "All records post automatically to Employee Manager & GAAP Payroll Reports.",
                 wraplength=900, justify="left",
                 font=("Helvetica", 9, "italic"),
                 fg="#6B7280", bg="WHITE").pack(anchor="w", padx=16, pady=(0, 12))

        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True, padx=10, pady=4)
        form_frame = scroll_frame.scrollable_frame
        form_frame.configure(bg="#F8FAFC")

        widgets = self._render_employee_form(form_frame)

        # Error banner + Buttons row
        footer = tk.Frame(main_container, bg="WHITE",
                          highlightbackground="#E5E7EB",
                          highlightthickness=1, bd=0)
        footer.pack(fill="x", padx=10, pady=(4, 10))
        errors_var = tk.StringVar(value="")
        error_banner = tk.Label(footer, textvariable=errors_var,
                                fg="#991B1B", bg="#FEE2E2", anchor="w",
                                justify="left", wraplength=720,
                                font=("Helvetica", 9, "bold"),
                                padx=10, pady=6)
        error_banner.pack(fill="x", padx=10, pady=(10, 4))
        error_banner.pack_forget()

        btn_box = tk.Frame(footer, bg="WHITE")
        btn_box.pack(fill="x", padx=10, pady=10)

        def save_employee():
            data, errors = self._collect_employee_form(widgets)
            if errors:
                error_banner.pack(fill="x", padx=10, pady=(10, 4))
                errors_var.set("⚠️  " + "   |   ".join(errors))
                return
            error_banner.pack_forget()
            errors_var.set("")
            # Pass all expanded fields to wrapper.add_employee (it supports the extras)
            try:
                employee_id = self.employee_manager.add_employee(
                    name=data['name'],
                    email=data.get('email', ''),
                    phone=data['phone'],
                    salary=data['salary'],
                    start_date=data['start_date'],
                    end_date=data.get('end_date') or None,
                    residency_expiry=data.get('residency_expiry') or None,
                    department=data.get('department'),
                    designation=data.get('designation'),
                    employee_type=data.get('employee_type') or "Full-time",
                    id_number=data.get('id_number'),
                    nationality=data.get('nationality'),
                    joining_date=data.get('joining_date') or data.get('start_date'),
                    bank_name=data.get('bank_name'),
                    bank_iban=data.get('bank_iban'),
                    emergency_contact=data.get('emergency_contact'),
                    emergency_phone=data.get('emergency_phone'),
                    notes=data.get('notes'),
                )
                # Also persist additional HR fields the wrapper currently ignores
                extras = {k: v for k, v in data.items() if k not in (
                    "name","email","phone","salary","start_date","end_date",
                    "residency_expiry","department","designation","employee_type",
                    "id_number","nationality","joining_date","bank_name","bank_iban",
                    "emergency_contact","emergency_phone","notes") and v}
                if extras and employee_id:
                    self.employee_manager.update_employee(employee_id, **extras)
                if employee_id:
                    messagebox.showinfo(
                        "✅  Employee Added Successfully",
                        f"Personnel file created.\n\n  Employee ID : {employee_id}\n  Name        : {data['name']}\n  Designation : {data.get('designation') or '-'}\n  Salary      : AED {data['salary']:,.2f}\n  Joining     : {data.get('joining_date') or data['start_date']}\n\nRecord is now visible in Employee Manager and will feed into the GAAP Payroll ledger when you run monthly payroll.")
                    self.refresh_employees()
                    dialog.destroy()
                else:
                    errors_var.set("Save to JSON failed (write error). Try again.")
                    error_banner.pack(fill="x", padx=10, pady=(10, 4))
            except Exception as ex:
                errors_var.set(f"Save failed: {ex}")
                error_banner.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Button(btn_box, text="💾  Save Employee", style="Accent.TButton",
                   command=save_employee).pack(side='right', padx=5)
        ttk.Button(btn_box, text="Cancel", style="Ghost.TButton",
                   command=dialog.destroy).pack(side='right', padx=5)
        _fit_window(dialog, 1020, 820, mode="dialog", remember_key="employee_add")

    def edit_employee(self):
        """Edit selected employee — professional form (same 6-section 30-field layout)."""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select one employee row to edit.")
            return
        employee_id = str(self.tree.item(selection[0])['values'][0])
        employee = None
        for emp in self.employee_manager.employees:
            if str(emp.get('employee_id')) == employee_id:
                employee = emp
                break
        if not employee:
            messagebox.showerror("Not Found", f"Employee ID {employee_id} not found in personnel list.")
            return

        dialog = tk.Toplevel(self.dialog)
        dialog.title(f"✏️  Edit Employee — {employee.get('name','')}  ({employee_id})")
        dialog.geometry("980x780")
        dialog.transient(self.dialog)
        dialog.resizable(True, True)

        main_container = tk.Frame(dialog)
        main_container.pack(fill='both', expand=True)

        header = tk.Frame(main_container, bg="WHITE",
                          highlightbackground="#E5E7EB",
                          highlightthickness=1, bd=0)
        header.pack(fill="x", padx=10, pady=(10, 4))
        tk.Label(header,
                 text=f"✏️  Edit Personnel File · {employee.get('name','')}  ·  ID: {employee_id}",
                 font=("Helvetica", 15, "bold"), fg="#1E3A5F",
                 bg="WHITE").pack(anchor="w", padx=16, pady=(12, 2))
        tk.Label(header,
                 text="Update any of the 6 sections below and click Save Changes. "
                      "Update flows to Employee Manager immediately and into GAAP reports after next payroll run / GL rebuild.",
                 wraplength=900, justify="left",
                 font=("Helvetica", 9, "italic"),
                 fg="#6B7280", bg="WHITE").pack(anchor="w", padx=16, pady=(0, 12))

        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True, padx=10, pady=4)
        form_frame = scroll_frame.scrollable_frame
        form_frame.configure(bg="#F8FAFC")

        prefill = dict(employee)
        widgets = self._render_employee_form(form_frame, prefill=prefill)

        footer = tk.Frame(main_container, bg="WHITE",
                          highlightbackground="#E5E7EB",
                          highlightthickness=1, bd=0)
        footer.pack(fill="x", padx=10, pady=(4, 10))
        errors_var = tk.StringVar(value="")
        error_banner = tk.Label(footer, textvariable=errors_var,
                                fg="#991B1B", bg="#FEE2E2", anchor="w",
                                justify="left", wraplength=720,
                                font=("Helvetica", 9, "bold"),
                                padx=10, pady=6)
        error_banner.pack(fill="x", padx=10, pady=(10, 4))
        error_banner.pack_forget()

        btn_box = tk.Frame(footer, bg="WHITE")
        btn_box.pack(fill="x", padx=10, pady=10)

        def save_changes():
            data, errors = self._collect_employee_form(widgets)
            if errors:
                error_banner.pack(fill="x", padx=10, pady=(10, 4))
                errors_var.set("⚠️  " + "   |   ".join(errors))
                return
            error_banner.pack_forget()
            errors_var.set("")
            try:
                ok = self.employee_manager.update_employee(employee_id, **data)
                if ok:
                    messagebox.showinfo(
                        "✅  Changes Saved",
                        f"Employee record updated.\n\n  Employee ID : {employee_id}\n  Name        : {data['name']}\n  Salary      : AED {data['salary']:,.2f}\n\nRun '🔁 Sync to Reports / Rebuild GL' to reflect new salary data in GAAP reports.")
                    self.refresh_employees()
                    dialog.destroy()
                else:
                    errors_var.set("Failed to save updates (JSON write returned False).")
                    error_banner.pack(fill="x", padx=10, pady=(10, 4))
            except Exception as ex:
                errors_var.set(f"Save failed: {ex}")
                error_banner.pack(fill="x", padx=10, pady=(10, 4))

        ttk.Button(btn_box, text="💾  Save Changes", style="Accent.TButton",
                   command=save_changes).pack(side='right', padx=5)
        ttk.Button(btn_box, text="Cancel", style="Ghost.TButton",
                   command=dialog.destroy).pack(side='right', padx=5)
        _fit_window(dialog, 1020, 820, mode="dialog", remember_key="employee_edit")
    
    def delete_employee(self):
        """Delete selected employee"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an employee to delete")
            return
        
        employee_id = self.tree.item(selection[0])['values'][0]
        employee_name = self.tree.item(selection[0])['values'][1]
        
        if messagebox.askyesno("Confirm Delete", 
                              f"Are you sure you want to delete employee:\n{employee_name} ({employee_id})?"):
            if self.employee_manager.delete_employee(employee_id):
                messagebox.showinfo("Success", "Employee deleted successfully!")
                self.refresh_employees()
            else:
                messagebox.showerror("Error", "Failed to delete employee")
    
    def monthly_salary_report(self):
        """Generate monthly salary report"""
        dialog = tk.Toplevel(self.dialog)
        dialog.title("Monthly Salary Report")
        dialog.geometry("700x600")
        dialog.transient(self.dialog)
        dialog.resizable(True, True)
        
        # Create scrollable frame
        main_container = tk.Frame(dialog)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="💰 Monthly Salary Report", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Month selection
        month_frame = ttk.Frame(main_frame)
        month_frame.pack(fill='x', pady=10)
        
        ttk.Label(month_frame, text="Select Month:").pack(side='left')
        self.month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        month_entry = ttk.Entry(month_frame, textvariable=self.month_var, width=10)
        month_entry.pack(side='left', padx=10)
        ttk.Label(month_frame, text="(YYYY-MM)").pack(side='left')
        
        # Employees list for salary input
        list_frame = ttk.LabelFrame(main_frame, text="Employee Salaries", padding="10")
        list_frame.pack(fill='both', expand=True, pady=10)
        
        # Create scrollable frame for employees
        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        self.salary_entries = {}
        employees = self.employee_manager.get_all_employees()
        
        # Header
        ttk.Label(scrollable_frame, text="Employee", font=('Helvetica', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Label(scrollable_frame, text="Base Salary", font=('Helvetica', 10, 'bold')).grid(row=0, column=1, padx=5, pady=5)
        ttk.Label(scrollable_frame, text="Salary Paid (AED)", font=('Helvetica', 10, 'bold')).grid(row=0, column=2, padx=5, pady=5)
        
        for i, employee in enumerate(employees):
            ttk.Label(scrollable_frame, text=employee['name']).grid(row=i+1, column=0, padx=5, pady=2, sticky='w')
            ttk.Label(scrollable_frame, text=f"AED {employee['salary']:.2f}").grid(row=i+1, column=1, padx=5, pady=2)
            
            salary_var = tk.StringVar(value=str(employee['salary']))
            salary_entry = ttk.Entry(scrollable_frame, textvariable=salary_var, width=15)
            salary_entry.grid(row=i+1, column=2, padx=5, pady=2)
            self.salary_entries[employee['employee_id']] = salary_var
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        _fit_window(dialog, 700, 600, mode="large", remember_key="salary_monthly")
        
        # Account selection
        account_frame = ttk.LabelFrame(main_frame, text="Payment Account", padding="10")
        account_frame.pack(fill='x', pady=10)
        ttk.Label(account_frame, text="Pay salaries from:").pack(side='left')
        self.salary_account = ttk.Combobox(account_frame, width=25)
        try:
            from balance_manager import BalanceManager
            bm = BalanceManager(self.employee_manager.data_folder)
            self.salary_account['values'] = bm.get_account_names()
        except Exception:
            self.salary_account['values'] = ["Cash"]
        self.salary_account.set("Cash")

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=10)
        
        def generate_report():
            month_year = self.month_var.get().strip()
            if not month_year:
                messagebox.showerror("Error", "Please enter month (YYYY-MM)")
                return
            
            salary_data = []
            total_salary = 0
            for emp_id, salary_var in self.salary_entries.items():
                try:
                    salary_paid = float(salary_var.get()) if salary_var.get() else 0.0
                    if salary_paid > 0:
                        # Find employee details
                        employee = None
                        for emp in employees:
                            if emp['employee_id'] == emp_id:
                                employee = emp
                                break
                        
                        if employee:
                            salary_data.append({
                                'employee_id': emp_id,
                                'employee_name': employee['name'],
                                'base_salary': employee['salary'],
                                'salary_paid': salary_paid
                            })
                            total_salary += salary_paid
                except ValueError:
                    messagebox.showerror("Error", f"Invalid salary amount for employee {emp_id}")
                    return
            
            if not salary_data:
                messagebox.showwarning("Warning", "No salary data to save")
                return
            
            account = self.salary_account.get().strip()
            report_id = self.employee_manager.add_monthly_salary_report(month_year, salary_data, account)
            if report_id:
                # Generate PDF report
                self.generate_salary_pdf(month_year, salary_data, total_salary)
                try:
                    from balance_manager import BalanceManager
                    bm = BalanceManager(self.employee_manager.data_folder)
                    ok, msg = bm.update_balance(account or 'Cash', total_salary, 'withdrawal', f"Salaries {month_year}")
                except Exception:
                    pass
                messagebox.showinfo("Success", f"Monthly salary report generated!\nTotal: AED {total_salary:.2f}\nPDF report saved to HopePharmaInvoice folder")
                dialog.destroy()
            else:
                messagebox.showerror("Error", "Failed to generate salary report")
        
        ttk.Button(button_frame, text="Generate Report & PDF", 
                  command=generate_report).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=dialog.destroy).pack(side='left', padx=5)
    
    def generate_salary_pdf(self, month_year, salary_data, total_salary):
        """Generate PDF salary report"""
        try:
            # Create HTML content for the report
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Salary Report - {month_year}</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    .header {{ text-align: center; border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 30px; }}
                    .company-name {{ font-size: 24px; font-weight: bold; color: #2c5aa0; }}
                    .report-title {{ font-size: 20px; margin: 10px 0; }}
                    .report-date {{ font-size: 16px; color: #666; }}
                    table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                    th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                    th {{ background-color: #f2f2f2; font-weight: bold; }}
                    .total-row {{ background-color: #e8f4f8; font-weight: bold; }}
                    .footer {{ margin-top: 50px; text-align: center; color: #666; font-size: 12px; }}
                </style>
            </head>
            <body>
                <div class="header">
                    <div class="company-name">Hope Pharma Medicine Trading</div>
                    <div class="report-title">Monthly Salary Report</div>
                    <div class="report-date">Period: {month_year}</div>
                    <div class="report-date">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
                </div>
                
                <table>
                    <thead>
                        <tr>
                            <th>Employee ID</th>
                            <th>Employee Name</th>
                            <th>Base Salary (AED)</th>
                            <th>Salary Paid (AED)</th>
                        </tr>
                    </thead>
                    <tbody>
            """
            
            for item in salary_data:
                html_content += f"""
                        <tr>
                            <td>{item['employee_id']}</td>
                            <td>{item['employee_name']}</td>
                            <td>{item['base_salary']:,.2f}</td>
                            <td>{item['salary_paid']:,.2f}</td>
                        </tr>
                """
            
            html_content += f"""
                        <tr class="total-row">
                            <td colspan="3" style="text-align: right;"><strong>Total Salary Paid:</strong></td>
                            <td><strong>AED {total_salary:,.2f}</strong></td>
                        </tr>
                    </tbody>
                </table>
                
                <div class="footer">
                    <p>This is an automatically generated salary report.</p>
                    <p>Hope Pharma Medicine Trading | Dubai, UAE</p>
                </div>
            </body>
            </html>
            """
            
            # Save HTML file
            pdf_filename = f"Salary_Report_{month_year}.html"
            pdf_path = os.path.join(self.employee_manager.data_folder, pdf_filename)
            
            with open(pdf_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            # Open in browser for printing/PDF conversion
            webbrowser.open('file://' + os.path.abspath(pdf_path))
            
        except Exception as e:
            print(f"Error generating PDF: {e}")
            messagebox.showerror("Error", f"Failed to generate PDF: {e}")
    
    def yearly_report(self):
        """Show yearly salary report"""
        dialog = tk.Toplevel(self.dialog)
        dialog.title("Yearly Salary Report")
        dialog.geometry("800x600")
        dialog.transient(self.dialog)
        dialog.resizable(True, True)
        
        # Create scrollable frame
        main_container = tk.Frame(dialog)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="📊 Yearly Salary Report", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Year selection
        year_frame = ttk.Frame(main_frame)
        year_frame.pack(fill='x', pady=10)
        
        ttk.Label(year_frame, text="Select Year:").pack(side='left')
        self.year_var = tk.StringVar(value=datetime.now().strftime("%Y"))
        year_entry = ttk.Entry(year_frame, textvariable=self.year_var, width=10)
        year_entry.pack(side='left', padx=10)
        
        ttk.Button(year_frame, text="Generate Report", 
                  command=self.show_yearly_report).pack(side='left', padx=10)
        ttk.Button(year_frame, text="Export to PDF", 
                  command=self.export_yearly_pdf).pack(side='left', padx=10)
        
        # Report display
        report_frame = ttk.LabelFrame(main_frame, text="Report Results", padding="10")
        report_frame.pack(fill='both', expand=True, pady=10)
        
        self.report_text = scrolledtext.ScrolledText(report_frame, height=20, width=80)
        self.report_text.pack(fill='both', expand=True)
        
        # Show current year report by default
        self.show_yearly_report()
        _fit_window(dialog, 800, 600, mode="large", remember_key="salary_yearly")
    
    def show_yearly_report(self):
        """Display yearly report"""
        year = self.year_var.get().strip()
        if not year:
            messagebox.showerror("Error", "Please enter a year")
            return
        
        report = self.employee_manager.get_yearly_report(year)
        
        self.report_text.delete('1.0', tk.END)
        
        if not report['reports']:
            self.report_text.insert(tk.END, f"No salary reports found for year {year}")
            return
        
        self.report_text.insert(tk.END, f"YEARLY SALARY REPORT - {year}\n")
        self.report_text.insert(tk.END, "=" * 60 + "\n\n")
        
        total_salary = 0
        monthly_totals = []
        
        for monthly_report in report['reports']:
            self.report_text.insert(tk.END, f"Month: {monthly_report['month_year']}\n")
            self.report_text.insert(tk.END, f"Total Salary Paid: AED {monthly_report['total_salary']:.2f}\n")
            self.report_text.insert(tk.END, "Employees Paid:\n")
            
            monthly_total = 0
            for salary_item in monthly_report['salary_data']:
                employee_name = salary_item.get('employee_name', 'Unknown')
                self.report_text.insert(tk.END, f"  - {employee_name}: AED {salary_item['salary_paid']:.2f}\n")
                monthly_total += salary_item['salary_paid']
            
            monthly_totals.append((monthly_report['month_year'], monthly_total))
            self.report_text.insert(tk.END, "\n")
            total_salary += monthly_report['total_salary']
        
        # Monthly breakdown
        self.report_text.insert(tk.END, "MONTHLY BREAKDOWN:\n")
        self.report_text.insert(tk.END, "-" * 40 + "\n")
        for month_year, amount in monthly_totals:
            self.report_text.insert(tk.END, f"{month_year}: AED {amount:,.2f}\n")
        
        self.report_text.insert(tk.END, "\n" + "=" * 60 + "\n")
        self.report_text.insert(tk.END, f"YEARLY TOTAL: AED {total_salary:,.2f}\n")
        self.report_text.insert(tk.END, f"Total Reports: {len(report['reports'])} months\n")
        
        # Store current report data for PDF export
        self.current_yearly_report = report
        self.current_year = year
    
    def export_yearly_pdf(self):
        """Export yearly report to PDF"""
        if not hasattr(self, 'current_yearly_report'):
            messagebox.showwarning("Warning", "No yearly report to export")
            return
        
        try:
            year = self.current_year
            report = self.current_yearly_report
            
            # Create HTML content
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Yearly Salary Report - {year}</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    .header {{ text-align: center; border-bottom: 2px solid #333; padding-bottom: 20px; margin-bottom: 30px; }}
                    .company-name {{ font-size: 24px; font-weight: bold; color: #2c5aa0; }}
                    .report-title {{ font-size: 20px; margin: 10px 0; }}
                    .report-date {{ font-size: 16px; color: #666; }}
                    table {{ width: 100%; border-collapse: collapse; margin: 20px 0; }}
                    th, td {{ border: 1px solid #ddd; padding: 12px; text-align: left; }}
                    th {{ background-color: #f2f2f2; font-weight: bold; }}
                    .month-section {{ margin: 30px 0; }}
                    .month-header {{ background-color: #e8f4f8; padding: 10px; font-weight: bold; }}
                    .total-row {{ background-color: #2c5aa0; color: white; font-weight: bold; }}
                    .footer {{ margin-top: 50px; text-align: center; color: #666; font-size: 12px; }}
                </style>
            </head>
            <body>
                <div class="header">
                    <div class="company-name">Hope Pharma Medicine Trading</div>
                    <div class="report-title">Yearly Salary Report</div>
                    <div class="report-date">Year: {year}</div>
                    <div class="report-date">Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}</div>
                </div>
            """
            
            total_yearly = 0
            for monthly_report in report['reports']:
                html_content += f"""
                <div class="month-section">
                    <div class="month-header">Month: {monthly_report['month_year']} - Total: AED {monthly_report['total_salary']:,.2f}</div>
                    <table>
                        <thead>
                            <tr>
                                <th>Employee ID</th>
                                <th>Employee Name</th>
                                <th>Base Salary (AED)</th>
                                <th>Salary Paid (AED)</th>
                            </tr>
                        </thead>
                        <tbody>
                """
                
                for salary_item in monthly_report['salary_data']:
                    html_content += f"""
                            <tr>
                                <td>{salary_item['employee_id']}</td>
                                <td>{salary_item.get('employee_name', 'Unknown')}</td>
                                <td>{salary_item.get('base_salary', 0):,.2f}</td>
                                <td>{salary_item['salary_paid']:,.2f}</td>
                            </tr>
                    """
                
                html_content += """
                        </tbody>
                    </table>
                </div>
                """
                total_yearly += monthly_report['total_salary']
            
            html_content += f"""
                <table>
                    <tr class="total-row">
                        <td colspan="3" style="text-align: right;"><strong>YEARLY TOTAL SALARY:</strong></td>
                        <td><strong>AED {total_yearly:,.2f}</strong></td>
                    </tr>
                </table>
                
                <div class="footer">
                    <p>This is an automatically generated yearly salary report.</p>
                    <p>Hope Pharma Medicine Trading | Dubai, UAE</p>
                </div>
            </body>
            </html>
            """
            
            # Save HTML file
            pdf_filename = f"Yearly_Salary_Report_{year}.html"
            pdf_path = os.path.join(self.employee_manager.data_folder, pdf_filename)
            
            with open(pdf_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            # Open in browser for printing/PDF conversion
            webbrowser.open('file://' + os.path.abspath(pdf_path))
            messagebox.showinfo("Success", f"Yearly report exported to:\n{pdf_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")

# =============================================================================
# COST CENTER HELPERS (shared by CreateInvoiceDialog + EditInvoiceDialog)
# =============================================================================

def _get_ccm(data_folder):
    if not COST_CENTER_AVAILABLE:
        return None
    try:
        return CostCenterManager(data_folder)
    except Exception:
        return None


def _confirm_cost_change(parent, *, action_title, summary_text, old_total, new_total,
                         grand_total, expense_account, fund_account, line_items_preview=None):
    """
    Confirmation dialog shown BEFORE any cost add / edit / delete / apply.
    """
    try:
        from decimal import Decimal as _D
        old_t = _D(str(old_total or 0))
        new_t = _D(str(new_total or 0))
        gt = _D(str(grand_total or 0))
    except Exception:
        from decimal import Decimal as _D
        old_t = _D("0"); new_t = _D("0"); gt = _D("0")
    delta = new_t - old_t
    new_pl = gt - new_t
    old_pl = gt - old_t
    pl_delta = new_pl - old_pl

    dlg = tk.Toplevel(parent)
    dlg.title(f"Confirm — {action_title}")
    dlg.transient(parent)
    dlg.grab_set()
    dlg.geometry("640x620")
    frm = ttk.Frame(dlg, padding=14)
    frm.pack(fill='both', expand=True)

    ttk.Label(frm, text=action_title, font=("Helvetica", 13, "bold")).pack(anchor='w')
    ttk.Label(frm, text=summary_text, foreground="#2D3748", wraplength=580,
              justify='left').pack(anchor='w', pady=(4, 10))

    rows = [
        ("Item", "Before", "After", "Δ"),
        ("Total Costs (AED)", f"{old_t:,.2f}", f"{new_t:,.2f}", f"{delta:+,.2f}"),
        ("Gross / Grand Total (AED)", f"{gt:,.2f}", f"{gt:,.2f}", "—"),
        ("Net Profit (AED)", f"{old_pl:,.2f}", f"{new_pl:,.2f}", f"{pl_delta:+,.2f}"),
    ]
    cols = ["#1", "#2", "#3", "#4"]
    tv = ttk.Treeview(frm, columns=cols, show='headings', height=4)
    widths = [240, 120, 120, 120]
    for c, w, h in zip(cols, widths, rows[0]):
        tv.heading(c, text=h)
        tv.column(c, width=w, anchor='center' if c != '#1' else 'w')
    for r in rows[1:]:
        tags = ()
        if r[0] == "Net Profit (AED)":
            tags = ("net",) if pl_delta >= 0 else ("net_bad",)
        tv.insert('', tk.END, values=r, tags=tags)
    try:
        tv.tag_configure("net", background="#F0FFF4", foreground="#276749")
        tv.tag_configure("net_bad", background="#FFF5F5", foreground="#9B2C2C")
    except Exception:
        pass
    tv.pack(fill='x', pady=6)

    gl_frame = ttk.LabelFrame(frm, text="GL Impact Preview (Double-Entry)", padding=8)
    gl_frame.pack(fill='x', pady=(8, 6))
    ea = str(expense_account or "5000")
    fa = str(fund_account or "1100")
    amt = abs(delta)
    if abs(delta) > 0:
        direction = "INCREASE in costs" if delta > 0 else "REDUCTION in costs"
        ttk.Label(gl_frame, text=f"{direction} of AED {amt:,.2f} will post:",
                  font=("Helvetica", 9, "bold")).pack(anchor='w')
        if delta > 0:
            ttk.Label(gl_frame, text=f"  Dr {ea}  Expense / COGS     AED {amt:,.2f}").pack(anchor='w')
            ttk.Label(gl_frame, text=f"  Cr {fa}  Bank/Cash/AP           AED {amt:,.2f}").pack(anchor='w')
        else:
            ttk.Label(gl_frame, text=f"  Dr {fa}  Bank/Cash/AP           AED {amt:,.2f}").pack(anchor='w')
            ttk.Label(gl_frame, text=f"  Cr {ea}  Expense / COGS     AED {amt:,.2f}").pack(anchor='w')
    else:
        ttk.Label(gl_frame,
                  text="No change in cost totals → no incremental GL entry expected on Save.").pack(anchor='w')

    if line_items_preview:
        preview_frame = ttk.LabelFrame(frm, text=f"Generated Cost Lines ({len(line_items_preview)})", padding=8)
        preview_frame.pack(fill='both', expand=True, pady=(6, 6))
        pcols = ["desc", "amt", "exp", "fund"]
        ptv = ttk.Treeview(preview_frame, columns=pcols, show='headings', height=7)
        headers = [("Description", 240), ("Amount (AED)", 110), ("Expense Acc.", 120), ("Fund Account", 130)]
        for c, (h, w) in zip(pcols, headers):
            ptv.heading(c, text=h); ptv.column(c, width=w, anchor='center' if c != 'desc' else 'w')
        for li in line_items_preview:
            try:
                from decimal import Decimal as _D
                if isinstance(li, dict):
                    a = _D(str(li.get("amount") or 0))
                    d = str(li.get("description") or "")
                    e = str(li.get("expense_account") or ea)
                    f = str(li.get("account") or li.get("fund_account") or fa)
                else:
                    t = tuple(li) + ("", "", "", "")
                    a = _D(str(t[1] or 0))
                    d = str(t[0] or "")
                    e = str(t[2] or ea)
                    f = str(t[3] or fa)
            except Exception:
                from decimal import Decimal as _D
                a = _D("0")
                d = ""; e = ea; f = fa
            ptv.insert('', tk.END, values=(d[:70], f"{a:,.2f}", e, f))
        ptv.pack(fill='both', expand=True)

    result = {"ok": False}

    def _ok():
        result["ok"] = True
        dlg.destroy()

    btns = ttk.Frame(frm); btns.pack(fill='x', pady=(10, 0))
    ttk.Button(btns, text="✅ Confirm & Apply", command=_ok).pack(side='right', padx=6)
    ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side='right')
    dlg.wait_window()
    return result["ok"]


def _inject_cost_center_ui(self, parent_frame, mode="create"):
    cc_frame = ttk.LabelFrame(parent_frame,
                              text="🎯 Cost Center — Auto-Apply Costs By Service/Item Rules", padding=10)
    cc_frame.pack(fill='x', pady=(0, 10))
    ttk.Label(cc_frame, text="Cost Center:").grid(row=0, column=0, padx=(0, 5), pady=4, sticky='w')
    self.cost_center_combo = ttk.Combobox(cc_frame, width=34, state="readonly")
    self.cost_center_combo.grid(row=0, column=1, padx=(0, 10), pady=4, sticky='w')
    ttk.Button(cc_frame, text="🔄 Refresh",
               command=lambda: _refresh_cost_centers_list(self)).grid(row=0, column=2, padx=4)
    ttk.Button(cc_frame, text="⚙️  Manage Cost Centers",
               command=lambda: _open_cost_center_manager(self)).grid(row=0, column=3, padx=4)
    ttk.Button(cc_frame, text="🧪 Preview Impact",
               command=lambda: _preview_cost_center(self)).grid(row=0, column=4, padx=4)
    ttk.Button(cc_frame, text="✅ Apply Cost Center To This Invoice",
               command=lambda: _apply_cost_center_to_invoice(self)).grid(row=0, column=5, padx=(10, 0))
    ttk.Label(cc_frame, text="Replace Existing Costs?", foreground="#4A5568").grid(row=0, column=6, padx=(20, 4))
    self.cc_replace_var = tk.BooleanVar(value=(mode != "create"))
    ttk.Checkbutton(cc_frame, variable=self.cc_replace_var, text="Replace").grid(row=0, column=7)
    self.cost_center_info = tk.Label(
        cc_frame,
        text="(Create a Cost Center first via Manage → add rules keyed on your service item names. Select it, Preview, then Apply.)",
        fg="#4A5568", anchor='w', justify='left',
    )
    self.cost_center_info.grid(row=1, column=0, columnspan=8, sticky='we', padx=4, pady=(4, 0))
    try:
        cc_frame.columnconfigure(1, weight=1)
    except Exception:
        pass


def _refresh_cost_centers_list(self):
    try:
        folder = (getattr(self.manager, "invoice_folder", None)
                  or getattr(self.manager, "data_folder", None) or os.getcwd())
    except Exception:
        folder = os.getcwd()
    vals = ["<None / Manual Costs Only>"]
    ids = [None]
    ccm = _get_ccm(folder)
    if ccm:
        for c in ccm.list_centers():
            rules_n = len(c.get("rules") or [])
            vals.append(f"{c.get('name')}  ({c.get('center_id')}) — {rules_n} rule(s)")
            ids.append(c.get("center_id"))
    self._cc_ids = ids
    try:
        self.cost_center_combo['values'] = vals
        if not self.cost_center_combo.get():
            self.cost_center_combo.current(0)
    except Exception:
        pass


def _selected_cost_center(self):
    try:
        idx = self.cost_center_combo.current()
        if idx <= 0:
            return None
        return self._cc_ids[idx]
    except Exception:
        return None


def _collect_items_for_cc(self):
    items = []
    for iid in self.items_tree.get_children():
        v = self.items_tree.item(iid)['values']
        if not v or len(v) < 5:
            continue
        try:
            qty = float(v[1] or 0)
            price = float(v[2] or 0)
        except Exception:
            qty = 0; price = 0
        items.append({
            "description": v[0] or "",
            "quantity": qty,
            "unit_price": price,
            "taxable": str(v[3] or "").strip().lower() in {"yes", "true", "1", "vat", "taxable"},
            "total": float(v[4] or 0),
        })
    return items


def _invoice_dict_for_cc(self):
    try:
        inv_date = (self.invoice_date.get() or "").strip()
    except Exception:
        inv_date = None
    try:
        gt = 0.0
        for iid in self.items_tree.get_children():
            v = self.items_tree.item(iid)['values']
            if v and len(v) >= 5:
                try: gt += float(v[4] or 0)
                except Exception: pass
    except Exception:
        gt = 0
    return {"date": inv_date, "items": _collect_items_for_cc(self), "grand_total": gt}


def _set_cost_center_info(self, msg, ok=True):
    try:
        self.cost_center_info.config(text=msg, fg="#276749" if ok else "#9B2C2C")
    except Exception:
        pass


def _q_cc(value):
    try:
        from decimal import Decimal as _D
        return _D(str(value or 0))
    except Exception:
        from decimal import Decimal as _D
        return _D("0")


def _aed_cc(value):
    from decimal import Decimal as _D
    try:
        return _q_cc(value).quantize(_D("0.01"), rounding=ROUND_HALF_UP)
    except Exception:
        from decimal import Decimal as _D
        return _D("0.00")


def _preview_cost_center(self):
    try:
        folder = (getattr(self.manager, "invoice_folder", None)
                  or getattr(self.manager, "data_folder", None) or os.getcwd())
    except Exception:
        folder = os.getcwd()
    cid = _selected_cost_center(self)
    if not cid:
        messagebox.showinfo("Cost Center", "Please select a Cost Center first.")
        return
    ccm = _get_ccm(folder)
    if not ccm:
        messagebox.showerror("Error", "Cost Center Manager unavailable.")
        return
    center = ccm.get_center(cid)
    if not center:
        messagebox.showerror("Error", "Cost Center not found.")
        return
    inv_dict = _invoice_dict_for_cc(self)
    new_costs, total, info = ccm.compute_costs_for_invoice(center, inv_dict)
    existing_total = _q_cc(0)
    for iid in self.costs_tree.get_children():
        v = self.costs_tree.item(iid)['values']
        existing_total = existing_total + _q_cc(v[1] if len(v) > 1 else 0)
    if not new_costs:
        _set_cost_center_info(self,
            "⚠️ No cost lines matched. Edit the Cost Center's rules → add patterns matching the item/service descriptions.",
            ok=False)
        messagebox.showwarning("No Matches",
            "No cost rules matched any invoice line.\n\n"
            "Open Manage Cost Centers → Add/Edit Rules → add patterns (e.g. 'Hearing Aid', 'SONIC', 'Battery', 'Pack') "
            "that appear in your item descriptions.")
        return
    if self.cc_replace_var.get():
        after_total = _aed_cc(total)
    else:
        after_total = _aed_cc(existing_total + total)
    _set_cost_center_info(self,
        f"✅ {len(new_costs)} cost lines, AED {_aed_cc(total):,.2f} "
        f"{'replace' if self.cc_replace_var.get() else 'added to'} existing (final AED {after_total:,.2f}). Click Apply to confirm.")
    lines = "\n".join(info[:20]) + ("" if len(info) <= 20 else f"\n... and {len(info) - 20} more")
    messagebox.showinfo("Cost Center Preview",
        f"Cost Center: {center.get('name')} ({center.get('center_id')})\n\n"
        f"New cost lines generated: {len(new_costs)}\n"
        f"New cost subtotal:        AED {_aed_cc(total):,.2f}\n"
        f"Existing cost subtotal:   AED {_aed_cc(existing_total):,.2f}\n"
        f"Mode:                     {'REPLACE existing costs' if self.cc_replace_var.get() else 'ADD to existing costs'}\n"
        f"Final cost subtotal:      AED {after_total:,.2f}\n\n"
        f"Breakdown:\n{lines}")


def _open_cost_center_manager(self):
    try:
        folder = (getattr(self.manager, "invoice_folder", None)
                  or getattr(self.manager, "data_folder", None) or os.getcwd())
    except Exception:
        folder = os.getcwd()
    ccm = _get_ccm(folder)
    if not ccm:
        messagebox.showerror("Error", "Cost Center Manager unavailable (cost_center_manager.py missing/broken).")
        return
    try:
        from cost_center_manager_ui import CostCenterManagerDialog
        dlg = CostCenterManagerDialog(self.dialog, ccm)
        self.dialog.wait_window(dlg.dialog)
    except Exception as e:
        messagebox.showinfo("Info",
            f"Cost Center Manager UI module not present ({e}).\n\n"
            "You can still manage cost centers directly in cost_centers.json (data folder).\n"
            "Schema per center: { center_id, name, description, default_expense_account, rules: [ "
            "{ rule_id, name, match_mode (contains|exact|starts_with|regex), patterns[], "
            "calc_type ('percentage'/'fixed_per_line'/'fixed_per_invoice'), value, expense_account, "
            "fund_account, apply_to_taxable_only, apply_to_non_taxable_only, notes } ] }")
    finally:
        try:
            _refresh_cost_centers_list(self)
        except Exception:
            pass


def _apply_cost_center_to_invoice(self):
    try:
        folder = (getattr(self.manager, "invoice_folder", None)
                  or getattr(self.manager, "data_folder", None) or os.getcwd())
    except Exception:
        folder = os.getcwd()
    cid = _selected_cost_center(self)
    if not cid:
        messagebox.showinfo("Cost Center", "Please select a Cost Center first.")
        return
    ccm = _get_ccm(folder)
    if not ccm:
        return
    center = ccm.get_center(cid)
    if not center:
        return
    inv_dict = _invoice_dict_for_cc(self)
    new_costs, total, info = ccm.compute_costs_for_invoice(center, inv_dict)
    if not new_costs:
        _set_cost_center_info(self, "⚠️ No cost lines matched — nothing to apply.", ok=False)
        return

    existing_total = _q_cc(0)
    fund_set = set()
    exp_set = set([str(center.get("default_expense_account") or "5000")])
    for iid in self.costs_tree.get_children():
        v = self.costs_tree.item(iid)['values']
        existing_total = existing_total + _q_cc(v[1] if len(v) > 1 else 0)
        if len(v) > 2 and v[2]:
            fund_set.add(str(v[2]))
    for c in new_costs:
        exp_set.add(str(c.get("expense_account") or "5000"))
        fund_set.add(str(c.get("account") or str(c.get("fund_account") or "ADCB")))

    replace = bool(self.cc_replace_var.get())
    after_total = _aed_cc(total) if replace else _aed_cc(existing_total + total)

    ok = _confirm_cost_change(
        self.dialog,
        action_title=f"Apply Cost Center «{center.get('name')}»",
        summary_text=(
            f"Applying cost center '{center.get('name')}' ({center.get('center_id')}) will generate "
            f"{len(new_costs)} cost line(s) totalling AED {_aed_cc(total):,.2f}.\n\n"
            + ("All EXISTING cost lines on this invoice will be REPLACED." if replace
               else "New cost lines will be ADDED alongside existing costs.")
        ),
        old_total=existing_total,
        new_total=after_total,
        grand_total=inv_dict.get("grand_total", 0),
        expense_account=" / ".join(sorted(exp_set)[:3]) + (" …" if len(exp_set) > 3 else ""),
        fund_account=" / ".join(sorted(fund_set)[:3]) + (" …" if len(fund_set) > 3 else ""),
        line_items_preview=new_costs,
    )
    if not ok:
        return

    if replace:
        for iid in list(self.costs_tree.get_children()):
            self.costs_tree.delete(iid)
    ncols = len(self.costs_tree["columns"])
    for c in new_costs:
        try: a = float(c.get("amount") or 0)
        except Exception: a = 0.0
        desc = str(c.get("description") or "")
        fund = str(c.get("account") or c.get("fund_account") or "ADCB")
        exp = str(c.get("expense_account") or "5000")
        if len(exp) > 4 and exp[:4].isdigit():
            exp = exp[:4]
        if not (len(exp) == 4 and exp.isdigit()):
            exp = "5000"
        receipt_mark = "📎" if c.get("receipt_attached") else ""
        if ncols == 5:
            self.costs_tree.insert('', tk.END, values=(desc, f"{a:.2f}", fund, exp, receipt_mark))
        else:
            self.costs_tree.insert('', tk.END, values=(desc, f"{a:.2f}", fund, receipt_mark))
    try:
        self._adjust_tree_heights()
    except Exception:
        pass
    try:
        self.update_totals_preview()
    except Exception:
        pass
    _set_cost_center_info(self,
        f"✅ Applied — {len(new_costs)} cost line(s), AED {_aed_cc(total):,.2f}. Edit below if needed, then Save Invoice.", ok=True)


class CreateInvoiceDialog:
    def __init__(self, parent, manager, data_memory):
        self.manager = manager
        self.data_memory = data_memory
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Create New Invoice")
        self.dialog.geometry("1100x720")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        _fit_window(self.dialog, 800, 520, mode="workspace", remember_key="create_invoice")
    
    def setup_ui(self):
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="Create New Invoice", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Invoice Type
        type_frame = ttk.LabelFrame(main_frame, text="Invoice Type", padding="10")
        type_frame.pack(fill='x', pady=5)
        
        self.invoice_type = tk.StringVar(value="sales")
        ttk.Radiobutton(type_frame, text="Sales Invoice", variable=self.invoice_type, value="sales").pack(side='left', padx=10)
        ttk.Radiobutton(type_frame, text="Service Invoice", variable=self.invoice_type, value="service").pack(side='left', padx=10)
        
        # Client Information
        client_frame = ttk.LabelFrame(main_frame, text="Client Information", padding="10")
        client_frame.pack(fill='x', pady=5)
        
        ttk.Label(client_frame, text="Client Name *:").grid(row=0, column=0, sticky='w', pady=5)
        self.client_name = ttk.Combobox(client_frame, width=30)
        self.client_name.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        # Load existing clients
        clients = self.data_memory.get_clients()
        self.client_name['values'] = [client['name'] for client in clients]
        self.client_name.bind('<KeyRelease>', self.on_client_search)
        
        ttk.Label(client_frame, text="Client TRN:").grid(row=1, column=0, sticky='w', pady=5)
        self.client_trn = ttk.Entry(client_frame, width=30)
        self.client_trn.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        
        # ADDED: Emirate Selection
        ttk.Label(client_frame, text="Location (City, Country):").grid(row=2, column=0, sticky='w', pady=5)
        self.client_emirate = ttk.Combobox(client_frame, 
                                         values=[
                                             "Dubai, United Arab Emirates", "Abu Dhabi, United Arab Emirates",
                                             "Sharjah, United Arab Emirates", "Ajman, United Arab Emirates",
                                             "Umm Al Quwain, United Arab Emirates", "Ras Al Khaimah, United Arab Emirates",
                                             "Fujairah, United Arab Emirates",
                                             "Cairo, Egypt", "Alexandria, Egypt", "Riyadh, Saudi Arabia",
                                             "Jeddah, Saudi Arabia", "Doha, Qatar", "Kuwait City, Kuwait",
                                             "Muscat, Oman"
                                         ],
                                         state="normal", width=27)
        self.client_emirate.set("")
        self.client_emirate.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        
        # Invoice Details
        details_frame = ttk.LabelFrame(main_frame, text="Invoice Details", padding="10")
        details_frame.pack(fill='x', pady=5)
        
        # ADDED: Date field
        ttk.Label(details_frame, text="Invoice Date:").grid(row=0, column=0, sticky='w', pady=5)
        self.invoice_date = ttk.Entry(details_frame, width=15)
        self.invoice_date.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        self.invoice_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        ttk.Label(details_frame, text="Payment Method:").grid(row=1, column=0, sticky='w', pady=5)
        self.payment_method = ttk.Combobox(details_frame,
                                           values=["Cash", "Transfer"],
                                           state="readonly", width=27)
        self.payment_method.set("Cash")
        self.payment_method.grid(row=1, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(details_frame, text="Payment Due:").grid(row=2, column=0, sticky='w', pady=5)
        self.payment_due = ttk.Combobox(details_frame,
                                        values=["On receipt", "30 Days", "60 Days", "90 Days", "120 Days"],
                                        state="readonly", width=27)
        self.payment_due.set("On receipt")
        self.payment_due.grid(row=2, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(details_frame, text="Currency:").grid(row=3, column=0, sticky='w', pady=5)
        self.currency_var = tk.StringVar(value="AED")
        self.currency_combo = ttk.Combobox(
            details_frame,
            textvariable=self.currency_var,
            values=["AED", "USD", "EUR", "GBP", "SAR", "QAR", "OMR", "KWD", "BHD", "EGP"],
            state="normal",
            width=10
        )
        self.currency_combo.grid(row=3, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(details_frame, text="VAT Rate %:").grid(row=4, column=0, sticky='w', pady=5)
        vat_frame = ttk.Frame(details_frame)
        vat_frame.grid(row=4, column=1, sticky='w', pady=5, padx=10)
        
        self.vat_rate_var = tk.StringVar(value="0")
        self.vat_rate = ttk.Combobox(vat_frame, textvariable=self.vat_rate_var, values=["0", "5", "10", "15"], state="readonly", width=5)
        self.vat_rate.pack(side='left')
        
        ttk.Label(vat_frame, text="%").pack(side='left', padx=(5, 0))
        
        # ... rest of the existing code remains the same ...
        
        # Items Section
        items_frame = ttk.LabelFrame(main_frame, text="Items / Services", padding="10")
        items_frame.pack(fill='both', expand=True, pady=5)
        
        # Items list with scrollbar
        tree_frame = ttk.Frame(items_frame)
        tree_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        columns = ('Description', 'Quantity', 'Unit Price', 'VAT', 'Total')
        self.items_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=6)
        
        for col in columns:
            self.items_tree.heading(col, text=col)
            self.items_tree.column(col, width=100, stretch=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=scrollbar.set)
        
        self.items_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        try:
            self.items_tree.bind('<Configure>', lambda e: self._resize_items_tree_columns())
            self.dialog.after(0, self._resize_items_tree_columns)
        except Exception:
            pass
        
        # Add item controls
        add_frame = ttk.Frame(items_frame)
        add_frame.pack(fill='x')
        
        ttk.Label(add_frame, text="Description:").grid(row=0, column=0, padx=(0, 5))
        self.item_desc = ttk.Combobox(add_frame, width=15)
        self.item_desc.grid(row=0, column=1, padx=(0, 10))
        # Load existing products
        products = self.data_memory.get_products()
        self.item_desc['values'] = [product['name'] for product in products]
        self.item_desc.bind('<KeyRelease>', self.on_product_search)
        
        ttk.Label(add_frame, text="Qty:").grid(row=0, column=2, padx=(0, 5))
        self.item_qty = ttk.Entry(add_frame, width=6)
        self.item_qty.grid(row=0, column=3, padx=(0, 10))
        self.item_qty.insert(0, "1")
        
        self.unit_price_label = ttk.Label(add_frame, text="Unit Price (AED):")
        self.unit_price_label.grid(row=0, column=4, padx=(0, 5))
        self.item_price = ttk.Entry(add_frame, width=10)
        self.item_price.grid(row=0, column=5, padx=(0, 10))
        
        ttk.Label(add_frame, text="VAT:").grid(row=0, column=6, padx=(0, 5))
        self.item_vat = ttk.Combobox(add_frame, values=["Yes", "No"], state="readonly", width=5)
        self.item_vat.set("No")
        self.item_vat.grid(row=0, column=7, padx=(0, 10))
        
        ttk.Button(add_frame, text="Add Item", command=self.add_item).grid(row=0, column=8, padx=(10, 0))
        ttk.Button(add_frame, text="Edit Item", command=self.edit_item_dialog).grid(row=0, column=9, padx=5)
        ttk.Button(add_frame, text="Remove", command=self.remove_item).grid(row=0, column=10, padx=5)
        ttk.Button(add_frame, text="Import Excel", command=self.import_items_from_excel).grid(row=0, column=11, padx=5)
        
        # Invoice Costs Section
        costs_frame = ttk.LabelFrame(main_frame, text="Invoice Costs (Internal Only)", padding="10")
        costs_frame.pack(fill='x', pady=5)

        # --- Cost Center applicator band was here; moved to InvoiceManager main menu (Cost Centers & Bulk Backfill) ---

        # Costs list with scrollbar
        costs_tree_frame = ttk.Frame(costs_frame)
        costs_tree_frame.pack(fill='both', expand=True, pady=(0, 10))

        cost_columns = ('Description', 'Amount', 'Account', 'Expense GL', 'Receipt')
        self.costs_tree = ttk.Treeview(costs_tree_frame, columns=cost_columns, show='headings', height=2)

        for col in cost_columns:
            self.costs_tree.heading(col, text=col)
            if col == 'Description':
                self.costs_tree.column(col, width=190)
            elif col == 'Amount':
                self.costs_tree.column(col, width=95)
            elif col == 'Account':
                self.costs_tree.column(col, width=120)
            elif col == 'Expense GL':
                self.costs_tree.column(col, width=110)
            else:
                self.costs_tree.column(col, width=70)

        costs_scrollbar = ttk.Scrollbar(costs_tree_frame, orient=tk.VERTICAL, command=self.costs_tree.yview)
        self.costs_tree.configure(yscrollcommand=costs_scrollbar.set)

        self.costs_tree.pack(side='left', fill='both', expand=True)
        costs_scrollbar.pack(side='right', fill='y')

        # Add cost controls
        add_cost_frame = ttk.Frame(costs_frame)
        add_cost_frame.pack(fill='x')

        ttk.Label(add_cost_frame, text="Cost Description:").grid(row=0, column=0, padx=(0, 5))
        self.cost_desc = ttk.Entry(add_cost_frame, width=20)
        self.cost_desc.grid(row=0, column=1, padx=(0, 10))

        self.cost_amount_label = ttk.Label(add_cost_frame, text="Amount (AED):")
        self.cost_amount_label.grid(row=0, column=2, padx=(0, 5))
        self.cost_amount = ttk.Entry(add_cost_frame, width=10)
        self.cost_amount.grid(row=0, column=3, padx=(0, 10))

        ttk.Label(add_cost_frame, text="Paid From:").grid(row=0, column=4, padx=(0, 5))
        self.cost_account = ttk.Combobox(add_cost_frame, width=15, state="readonly")
        try:
            from balance_manager import BalanceManager
            bm = BalanceManager(self.manager.invoice_folder)
            accounts = bm.get_account_names() or ["Cash"]
            self.cost_account['values'] = ["Select Account"] + accounts
        except Exception:
            self.cost_account['values'] = ["Select Account", "Cash", "ADCB"]
        try:
            last = getattr(self.manager, "last_selected_account", None)
            if last and last in self.cost_account['values']:
                self.cost_account.set(last)
            else:
                self.cost_account.set("Select Account")
        except Exception:
            self.cost_account.set("Select Account")
        self.cost_account.grid(row=0, column=5, padx=(0, 10))

        ttk.Label(add_cost_frame, text="Expense GL:").grid(row=0, column=6, padx=(0, 5))
        self.cost_expense_account = ttk.Combobox(add_cost_frame, width=22, state="readonly")
        _ea_labels = ["5000  COGS (Cost of Goods Sold)",
                      "5100  Salary / Wages",
                      "5200  Rent, Utilities (Electricity, Water, Internet)",
                      "5300  Office Supplies, Travel, Marketing, Other Admin",
                      "5400  Professional Fees, Audit, Legal",
                      "5500  Delivery & Shipping / Logistics",
                      "5600  Medical / Clinical / Device Service Fees",
                      "5700  Commissions Paid to Sales Agents",
                      "5800  Warranty / Returns / After-Sales",
                      "5900  Other Operating Expenses"]
        self.cost_expense_account['values'] = _ea_labels
        self.cost_expense_account.current(0)
        self.cost_expense_account.grid(row=0, column=7, padx=(0, 10))

        ttk.Button(add_cost_frame, text="Add Cost", command=self.add_cost).grid(row=0, column=8, padx=(10, 5))
        ttk.Button(add_cost_frame, text="Edit Cost", command=self.edit_cost_dialog).grid(row=0, column=9, padx=5)
        ttk.Button(add_cost_frame, text="Upload Receipt", command=self.upload_receipt).grid(row=0, column=10, padx=5)
        ttk.Button(add_cost_frame, text="Remove Cost", command=self.remove_cost).grid(row=0, column=11, padx=5)

        # Totals preview
        totals_frame = ttk.Frame(main_frame)
        totals_frame.pack(fill='x', pady=5)

        self.totals_label = ttk.Label(totals_frame,
                                     text="Subtotal: AED 0.00 | VAT: AED 0.00 | Total: AED 0.00 | Total Cost: AED 0.00 | P/L: AED 0.00",
                                     font=('Helvetica', 9, 'bold'))
        self.totals_label.pack()

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Create Invoice", command=self.create_invoice).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy).pack(side='left', padx=5)
        
        notes_frame = ttk.LabelFrame(main_frame, text="Notes", padding="10")
        notes_frame.pack(fill='x', pady=5)
        
        self.notes_text = scrolledtext.ScrolledText(notes_frame, height=3)
        self.notes_text.pack(fill='x')

        try:
            self.currency_var.trace('w', lambda *args: self.update_currency_ui())
        except Exception:
            pass
        try:
            self.vat_rate_var.trace('w', lambda *args: self.update_totals_preview())
        except Exception:
            pass
        self.update_currency_ui()
    
    def on_client_search(self, event=None):
        """Handle client search in combobox"""
        value = self.client_name.get().lower()
        if value == '':
            clients = self.data_memory.get_clients()
            self.client_name['values'] = [client['name'] for client in clients]
        else:
            data = []
            clients = self.data_memory.search_clients(value)
            for client in clients:
                data.append(client['name'])
            self.client_name['values'] = data
    
    def on_product_search(self, event=None):
        """Handle product search in combobox"""
        value = self.item_desc.get().lower()
        if value == '':
            products = self.data_memory.get_products()
            self.item_desc['values'] = [product['name'] for product in products]
        else:
            data = []
            products = self.data_memory.search_products(value)
            for product in products:
                data.append(product['name'])
            self.item_desc['values'] = data

    def _parse_invoice_items_excel(self, file_path):
        warnings = []
        errors = []
        if not OPENPYXL_AVAILABLE:
            errors.append("Excel library (openpyxl) is not available in this build.")
            return [], warnings, errors

        def norm(value):
            """Enhanced normalization: lowercase, strip, remove special chars, collapse whitespace"""
            if value is None:
                return ""
            import re
            s = str(value).strip().lower()
            s = re.sub(r'[^a-z0-9\s]', '', s)
            s = re.sub(r'\s+', ' ', s).strip()
            return s

        def fuzzy_match(a, b, threshold=0.8):
            """Fuzzy string match using difflib SequenceMatcher"""
            import difflib
            if not a or not b:
                return False, 0.0
            ratio = difflib.SequenceMatcher(None, a, b).ratio()
            return ratio >= threshold, ratio

        def to_float(value, default=0.0):
            try:
                if value is None:
                    return default
                if isinstance(value, str):
                    cleaned = value.strip().replace(",", "")
                    filtered = "".join(ch for ch in cleaned if ch.isdigit() or ch in ".-")
                    if filtered and any(ch.isdigit() for ch in filtered):
                        cleaned = filtered
                    if not cleaned:
                        return default
                    return float(cleaned)
                return float(value)
            except Exception:
                return default

        def vat_to_yes_no(value):
            if value is None:
                return "No"
            if isinstance(value, (int, float)):
                return "Yes" if float(value) > 0 else "No"
            text = norm(value)
            if not text:
                return "No"
            if text in {"yes", "y", "true", "1", "vat", "tax", "taxable", "included"}:
                return "Yes"
            if text in {"no", "n", "false", "0", "non-taxable", "without vat", "excluded"}:
                return "No"
            return "Yes" if to_float(text, 0.0) > 0 else "No"

        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as exc:
            errors.append(f"Error reading Excel file: {exc}")
            return [], warnings, errors

        header_keywords = {
            'description': ['description', 'item', 'item description', 'product', 'service', 'details', 'name'],
            'quantity': ['qty', 'quantity', 'qnty', 'count'],
            'actual_rate': ['actual rate', 'actual unit price', 'actual price'],
            'rate': ['unit price', 'rate', 'price', 'unit cost', 'amount per unit'],
            'total': ['total', 'amount', 'line total', 'net'],
            'vat': ['vat', 'tax', 'vat %', 'tax %', 'vat amount', 'tax amount', 'taxable'],
        }

        worksheet = workbook.active
        header_row_idx = None
        header_map = {}
        best_score = 0
        max_scan_cols = 0
        potential_mismatches = []
        for candidate_sheet in workbook.worksheets:
            candidate_max_scan_rows = min(50, candidate_sheet.max_row or 0)
            candidate_max_scan_cols = min(60, candidate_sheet.max_column or 0)
            for row_index in range(1, candidate_max_scan_rows + 1):
                row_text = [norm(candidate_sheet.cell(row=row_index, column=col_index).value) for col_index in range(1, candidate_max_scan_cols + 1)]
                score = 0
                current_map = {}
                used_cols = set()
                # First try exact matches
                for field, keywords in header_keywords.items():
                    found = False
                    for col_index, cell_text in enumerate(row_text, start=1):
                        if col_index in used_cols or not cell_text:
                            continue
                        for keyword in keywords:
                            norm_keyword = norm(keyword)
                            if cell_text == norm_keyword:
                                current_map[field] = col_index
                                used_cols.add(col_index)
                                score += 10  # Exact match gets high score
                                found = True
                                break
                        if found:
                            break
                    if found:
                        continue
                    # Then try fuzzy matches
                    best_ratio = 0
                    best_col = None
                    for col_index, cell_text in enumerate(row_text, start=1):
                        if col_index in used_cols or not cell_text:
                            continue
                        for keyword in keywords:
                            norm_keyword = norm(keyword)
                            is_match, ratio = fuzzy_match(cell_text, norm_keyword, threshold=0.7)
                            if is_match and ratio > best_ratio:
                                best_ratio = ratio
                                best_col = col_index
                    if best_col is not None:
                        current_map[field] = best_col
                        used_cols.add(best_col)
                        score += best_ratio
                        if best_ratio < 0.9:
                            potential_mismatches.append(f"Potential header mismatch in sheet '{candidate_sheet.title}' row {row_index}: column {best_col} mapped to '{field}' (similarity: {best_ratio:.2f})")
                if score > best_score:
                    best_score = score
                    worksheet = candidate_sheet
                    header_row_idx = row_index
                    header_map = current_map
                    max_scan_cols = candidate_max_scan_cols
        
        if potential_mismatches:
            warnings.extend(potential_mismatches)

        if not header_row_idx or 'description' not in header_map:
            errors.append("Could not detect the invoice items table in the Excel file. Make sure it has a Description or Item column.")
            return [], warnings, errors

        imported_items = []
        blank_streak = 0
        stop_labels = (
            "subtotal", "sub total", "total", "grand total", "invoice total",
            "amount due", "net total", "balance due", "tax", "vat",
            "tax amount", "vat amount"
        )
        for row_index in range(header_row_idx + 1, (worksheet.max_row or header_row_idx) + 1):
            desc_value = worksheet.cell(row=row_index, column=header_map['description']).value
            description = str(desc_value).strip() if desc_value is not None else ""
            if not description:
                empty_row = True
                for check_col in range(1, max(1, max_scan_cols) + 1):
                    if worksheet.cell(row=row_index, column=check_col).value not in (None, ""):
                        empty_row = False
                        break
                if empty_row:
                    blank_streak += 1
                    if imported_items and blank_streak >= 5:
                        break
                else:
                    blank_streak = 0
                continue
            blank_streak = 0

            normalized_description = norm(description)
            if any(
                normalized_description == label
                or normalized_description.startswith(label + " ")
                or normalized_description.endswith(" " + label)
                for label in stop_labels
            ):
                continue

            quantity = 1.0
            if 'quantity' in header_map:
                quantity = to_float(worksheet.cell(row=row_index, column=header_map['quantity']).value, 1.0)
                if quantity <= 0:
                    quantity = 1.0

            unit_price = 0.0
            if 'actual_rate' in header_map:
                unit_price = to_float(worksheet.cell(row=row_index, column=header_map['actual_rate']).value, 0.0)
            elif 'rate' in header_map:
                unit_price = to_float(worksheet.cell(row=row_index, column=header_map['rate']).value, 0.0)

            line_total = None
            if 'total' in header_map:
                line_total = to_float(worksheet.cell(row=row_index, column=header_map['total']).value, None)
            if line_total is None:
                line_total = quantity * unit_price
            elif unit_price == 0 and quantity:
                unit_price = line_total / quantity

            vat_text = "No"
            if 'vat' in header_map:
                vat_text = vat_to_yes_no(worksheet.cell(row=row_index, column=header_map['vat']).value)

            imported_items.append({
                'description': description,
                'quantity': float(quantity),
                'unit_price': float(unit_price),
                'total': float(line_total),
                'vat': vat_text,
            })

        if not imported_items:
            warnings.append("No invoice items were found under the detected Excel header row.")
        return imported_items, warnings, errors

    def import_items_from_excel(self):
        file_path = filedialog.askopenfilename(
            title="Import Invoice Items From Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return

        imported_items, warnings, errors = self._parse_invoice_items_excel(file_path)
        if errors and not imported_items:
            messagebox.showerror("Import Excel", "\n".join(errors))
            return

        if self.items_tree.get_children() and imported_items:
            replace_existing = messagebox.askyesno(
                "Import Excel",
                "The invoice already has items.\n\nClick Yes to replace them, or No to append the imported items."
            )
            if replace_existing:
                for item_id in self.items_tree.get_children():
                    self.items_tree.delete(item_id)

        imported_count = 0
        for item in imported_items:
            description = str(item.get('description') or "").strip()
            if not description:
                continue
            quantity = float(item.get('quantity', 0) or 0)
            unit_price = float(item.get('unit_price', 0) or 0)
            line_total = float(item.get('total', quantity * unit_price) or 0)
            vat_text = "Yes" if str(item.get('vat', 'No')).strip().lower() == "yes" else "No"
            self.items_tree.insert('', tk.END, values=(
                description,
                quantity,
                f"{unit_price:.2f}",
                vat_text,
                f"{line_total:.2f}"
            ))
            try:
                self.data_memory.add_product(description, unit_price)
            except Exception:
                pass
            imported_count += 1

        if imported_count:
            self.update_totals_preview()
            self._adjust_tree_heights()

        message_lines = [f"Imported {imported_count} item(s) from Excel."]
        if warnings:
            message_lines.append("")
            message_lines.append("Warnings:")
            message_lines.extend(warnings)
        if errors:
            message_lines.append("")
            message_lines.append("Notes:")
            message_lines.extend(errors)
        messagebox.showinfo("Import Excel", "\n".join(message_lines))
    
    def add_item(self):
        """Add item to the items list"""
        desc = self.item_desc.get().strip()
        qty = self.item_qty.get().strip()
        price = self.item_price.get().strip()
        vat_applicable = self.item_vat.get() == "Yes"
        
        if not desc:
            messagebox.showwarning("Warning", "Please enter item description")
            return
            
        try:
            qty_val = float(qty)
            price_val = float(price)
            if qty_val <= 0 or price_val < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Warning", "Please enter valid quantity and price")
            return
        
        # Add to product memory
        self.data_memory.add_product(desc, price_val)
            
        total = qty_val * price_val
        vat_text = "Yes" if vat_applicable else "No"
        self.items_tree.insert('', tk.END, values=(
            desc, 
            qty_val, 
            f"{price_val:.2f}", 
            vat_text, 
            f"{total:.2f}"
        ))
        self._adjust_tree_heights()
        
        # Clear fields
        self.item_desc.delete(0, tk.END)
        self.item_qty.delete(0, tk.END)
        self.item_qty.insert(0, "1")
        self.item_price.delete(0, tk.END)
        
        # Update totals preview
        self.update_totals_preview()
    
    def remove_item(self):
        """Remove selected item"""
        selection = self.items_tree.selection()
        if selection:
            self.items_tree.delete(selection[0])
            self.update_totals_preview()
            self._adjust_tree_heights()
    
    def add_cost(self):
        """Add cost to the costs list — with financial-impact confirmation BEFORE applying."""
        desc = self.cost_desc.get().strip()
        amount = self.cost_amount.get().strip()
        account_val = self.cost_account.get().strip()
        expense_gl = getattr(self, "cost_expense_account", None)
        expense_gl_val = ""
        if expense_gl is not None:
            raw = expense_gl.get().strip()
            if raw:
                expense_gl_val = (raw.split(None, 1)[0] if raw[:4].isdigit() else "5000")

        if not desc:
            messagebox.showwarning("Warning", "Please enter cost description")
            return

        try:
            amount_val = float(amount)
            if amount_val < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Warning", "Please enter valid cost amount")
            return

        if not account_val or account_val == "Select Account":
            messagebox.showwarning("Warning", "Please select an account for this cost")
            return

        try:
            setattr(self.manager, "last_selected_account", account_val)
            self.cost_account.set(account_val)
        except Exception:
            pass

        # Build preview single line
        single_line = [(desc, amount_val, expense_gl_val or "5000", account_val)]

        # Confirm with financial impact
        old_total = sum(float(c.get('amount', 0)) for c in self._collect_costs_for_save())
        new_total = old_total + amount_val
        # Compute grand_total via update_totals_preview side-effect free
        try:
            gt = self._preview_grand_total()
        except Exception:
            gt = 0.0

        confirmed = _confirm_cost_change(
            self.dialog,
            action_title="Add Manual Cost",
            summary_text=f"Adding cost '{desc}' for {amount_val:,.2f} AED from '{account_val}'.",
            old_total=old_total,
            new_total=new_total,
            grand_total=gt,
            expense_account=expense_gl_val or "5000",
            fund_account=account_val,
            line_items_preview=single_line,
        )
        if not confirmed:
            return

        # Do NOT post BalanceManager here — defer to LedgerService on Invoice save.

        cols_count = len(self.costs_tree["columns"])
        if cols_count == 5:
            self.costs_tree.insert('', tk.END, values=(
                desc,
                f"{amount_val:.2f}",
                account_val,
                expense_gl_val,
                ""
            ))
        else:
            self.costs_tree.insert('', tk.END, values=(
                desc,
                f"{amount_val:.2f}",
                account_val,
                ""
            ))
        self._adjust_tree_heights()

        # Clear fields
        self.cost_desc.delete(0, tk.END)
        self.cost_amount.delete(0, tk.END)

        # Update totals preview
        self.update_totals_preview()

    def _preview_grand_total(self):
        """Helper to compute grand_total without side effects."""
        subtotal = 0.0
        for item in self.items_tree.get_children():
            vals = self.items_tree.item(item)['values']
            try:
                qty = float(vals[1])
                price = float(vals[2])
                vat = str(vals[3]).strip().lower() == 'yes'
                line = qty * price
                if vat:
                    line *= 1.05
                subtotal += line
            except Exception:
                continue
        return subtotal

    def _collect_costs_for_save(self):
        """Helper: extract current costs list from tree as dict list."""
        costs = []
        for item in self.costs_tree.get_children():
            vals = self.costs_tree.item(item)['values']
            try:
                if len(vals) >= 5:
                    desc, amt, acc, exp_gl, receipt = (vals[0], vals[1], vals[2], vals[3], vals[4])
                elif len(vals) == 4:
                    desc, amt, acc, receipt = vals
                    exp_gl = ""
                else:
                    continue
                costs.append({
                    'description': str(desc),
                    'amount': float(amt),
                    'account': str(acc),
                    'expense_account': str(exp_gl) if str(exp_gl) not in ("", "None") else "5000",
                    'receipt': str(receipt) if str(receipt) not in ("", "None") else "",
                })
            except Exception:
                continue
        return costs

    def remove_cost(self):
        """Remove selected cost — confirm with financial impact first."""
        selection = self.costs_tree.selection()
        if not selection:
            return
        values = self.costs_tree.item(selection[0])['values']
        try:
            desc = str(values[0])
            amt = float(values[1])
            if len(values) >= 5:
                acc = values[2]
                exp_gl = values[3] if values[3] else "5000"
            else:
                acc = values[2]
                exp_gl = "5000"
        except Exception:
            return

        old_total = sum(float(c.get('amount', 0)) for c in self._collect_costs_for_save())
        new_total = old_total - amt
        try:
            gt = self._preview_grand_total()
        except Exception:
            gt = 0.0

        confirmed = _confirm_cost_change(
            self.dialog,
            action_title="Remove Cost",
            summary_text=f"Removing cost '{desc}' ({amt:,.2f} AED) funded from '{acc}'.",
            old_total=old_total,
            new_total=new_total,
            grand_total=gt,
            expense_account=str(exp_gl),
            fund_account=str(acc),
            line_items_preview=None,
        )
        if not confirmed:
            return

        # Don't touch BalanceManager here; LedgerService auto-reverses on save.
        self.costs_tree.delete(selection[0])
        self.update_totals_preview()
        self._adjust_tree_heights()
    
    def _adjust_tree_heights(self):
        try:
            items_count = len(self.items_tree.get_children())
            costs_count = len(self.costs_tree.get_children())
            self.items_tree.configure(height=max(6, min(items_count + 1, 18)))
            self.costs_tree.configure(height=max(4, min(costs_count + 1, 10)))
            self._resize_items_tree_columns()
        except Exception:
            pass

    def _resize_items_tree_columns(self):
        try:
            total_width = self.items_tree.winfo_width()
            if total_width <= 60:
                return
            available = max(0, total_width - 24)
            weights = {
                'Description': 0.48,
                'Quantity': 0.10,
                'Unit Price': 0.15,
                'VAT': 0.07,
                'Total': 0.20,
            }
            mins = {
                'Description': 240,
                'Quantity': 80,
                'Unit Price': 120,
                'VAT': 70,
                'Total': 140,
            }
            widths = {}
            for col, w in weights.items():
                widths[col] = max(mins.get(col, 80), int(available * w))
            leftover = available - sum(widths.values())
            if leftover != 0:
                widths['Description'] = max(mins['Description'], widths['Description'] + leftover)
            for col, w in widths.items():
                self.items_tree.column(col, width=w, minwidth=mins.get(col, 60), stretch=True)
        except Exception:
            pass
            pass

    def _resize_items_tree_columns(self):
        try:
            total_width = self.items_tree.winfo_width()
            if total_width <= 60:
                return
            available = max(0, total_width - 24)
            weights = {
                'Description': 0.48,
                'Quantity': 0.10,
                'Unit Price': 0.15,
                'VAT': 0.07,
                'Total': 0.20,
            }
            mins = {
                'Description': 240,
                'Quantity': 80,
                'Unit Price': 120,
                'VAT': 70,
                'Total': 140,
            }
            widths = {}
            for col, w in weights.items():
                widths[col] = max(mins.get(col, 80), int(available * w))
            leftover = available - sum(widths.values())
            if leftover != 0:
                widths['Description'] = max(mins['Description'], widths['Description'] + leftover)
            for col, w in widths.items():
                self.items_tree.column(col, width=w, minwidth=mins.get(col, 60), stretch=True)
        except Exception:
            pass
    
    def upload_receipt(self):
        """Upload receipt for selected cost"""
        selection = self.costs_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a cost first")
            return
        
        # For new invoices, we can't upload receipts until invoice is created
        messagebox.showinfo("Info", "Please create the invoice first, then you can upload receipts when editing the invoice.")
    
    def update_totals_preview(self):
        """Update the totals preview"""
        subtotal = 0
        vat_amount = 0
        total_cost = 0
        
        # Calculate items subtotal and VAT
        for item in self.items_tree.get_children():
            values = self.items_tree.item(item)['values']
            item_total = float(values[4])
            subtotal += item_total
            
            # Calculate VAT for taxable items
            if values[3] == "Yes":  # If VAT applicable
                vat_rate = float(self.vat_rate_var.get()) / 100
                vat_amount += item_total * vat_rate
        
        # Calculate total costs
        for cost in self.costs_tree.get_children():
            values = self.costs_tree.item(cost)['values']
            total_cost += float(values[1])
        
        total_amount = subtotal + vat_amount
        profit_loss = total_amount - total_cost

        currency = (getattr(self, "currency_var", None).get() if getattr(self, "currency_var", None) else "AED") or "AED"
        currency = str(currency).strip() or "AED"
        self.totals_label.config(
            text=f"Subtotal: {currency} {subtotal:.2f} | VAT: {currency} {vat_amount:.2f} | Total: {currency} {total_amount:.2f} | Total Cost: {currency} {total_cost:.2f} | P/L: {currency} {profit_loss:.2f}"
        )

    def update_currency_ui(self):
        currency = str(self.currency_var.get() or "").strip() or "AED"
        self.items_tree.heading('Unit Price', text=f"Unit Price ({currency})")
        self.items_tree.heading('Total', text=f"Total ({currency})")
        self.costs_tree.heading('Amount', text=f"Amount ({currency})")
        try:
            self.unit_price_label.config(text=f"Unit Price ({currency}):")
        except Exception:
            pass
        try:
            self.cost_amount_label.config(text=f"Amount ({currency}):")
        except Exception:
            pass
        self.update_totals_preview()

    def create_invoice(self):
        """Create the invoice"""
        client_name = self.client_name.get().strip()
        
        if not client_name:
            messagebox.showwarning("Warning", "Please enter client name")
            return
            
        if len(self.items_tree.get_children()) == 0:
            messagebox.showwarning("Warning", "Please add at least one item")
            return
        
        try:
            vat_rate = float(self.vat_rate_var.get())
        except ValueError:
            messagebox.showwarning("Warning", "Please select a valid VAT rate")
            return
        
        # Validate date
        invoice_date = self.invoice_date.get().strip()
        if not invoice_date:
            invoice_date = datetime.now().strftime("%Y-%m-%d")
        else:
            try:
                datetime.strptime(invoice_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Warning", "Please enter a valid date (YYYY-MM-DD)")
                return
        
        # Add client to memory
        self.data_memory.add_client(client_name, self.client_trn.get().strip())
        
        # Calculate totals
        subtotal = 0
        taxable_amount = 0
        non_taxable_amount = 0
        total_cost = 0
        
        items = []
        for item in self.items_tree.get_children():
            values = self.items_tree.item(item)['values']
            description = values[0]
            quantity = float(values[1])
            unit_price = float(values[2])
            taxable = values[3] == "Yes"
            total = quantity * unit_price
            
            subtotal += total
            
            if taxable:
                taxable_amount += total
            else:
                non_taxable_amount += total
            
            items.append({
                'description': description,
                'quantity': quantity,
                'unit_price': unit_price,
                'total': total,
                'taxable': taxable
            })
        
        # Calculate costs
        costs = []
        for cost in self.costs_tree.get_children():
            values = self.costs_tree.item(cost)['values']
            vlen = len(values)
            desc = values[0]
            amount = float(values[1])
            fund = values[2] if vlen > 2 else ""
            exp_gl = (values[3] if vlen > 3 else "") or "5000"
            receipt = (values[4] if vlen > 4 else "") if vlen > 3 else (values[2] if vlen > 2 else "")
            receipt_attached = str(receipt) in ('📎', 'yes', 'Yes', 'Y', 'y', 'true')
            # normalize exp_gl to 4-digit code
            exp_code = str(exp_gl).strip()
            if len(exp_code) > 4 and exp_code[:4].isdigit():
                exp_code = exp_code[:4]
            if not (len(exp_code) == 4 and exp_code.isdigit()):
                exp_code = "5000"
            costs.append({
                'description': str(desc),
                'amount': amount,
                'account': str(fund) if str(fund) not in ("", "None") else "Cash",
                'expense_account': exp_code,
                'receipt_attached': receipt_attached,
            })
            total_cost += amount

        # Calculate VAT
        vat_amount = taxable_amount * (vat_rate / 100)
        grand_total = subtotal + vat_amount
        profit_loss = grand_total - total_cost

        # Create invoice as dictionary first
        invoice_id = self.manager.generate_invoice_id()
        loc_value = (self.client_emirate.get() or "").strip()
        try:
            low = loc_value.lower()
            if ',' in loc_value and 'united arab emirates' in low:
                emirate_guess = loc_value.split(',', 1)[0].strip()
            else:
                emirate_guess = loc_value
        except Exception:
            emirate_guess = loc_value
        invoice_dict = {
            'invoice_id': invoice_id,
            'invoice_type': self.invoice_type.get(),
            'client_name': client_name,
            'client_trn': self.client_trn.get().strip(),
            'client_emirate': emirate_guess,
            'client_location': loc_value,
            'date': invoice_date,  # CHANGED: Use the entered date
            'payment_method': self.payment_method.get(),
            'payment_due': self.payment_due.get(),
            'due_date': self._compute_due_date(invoice_date, self.payment_due.get()),
            'tax_rate': vat_rate,
            'items': items,
            'costs': costs,
            'subtotal': subtotal,
            'taxable_amount': taxable_amount,
            'non_taxable_amount': non_taxable_amount,
            'tax_amount': vat_amount,
            'grand_total': grand_total,
            'total_cost': total_cost,
            'profit_loss': profit_loss,
            'total_paid': 0.0,
            'status': 'Not Paid',
            'notes': self.notes_text.get('1.0', tk.END).strip(),
            'currency': (str(self.currency_var.get() or "").strip() or "AED"),
            'inventory_links': []
        }

        inventory_sync_message = ""
        inventory_manager = _get_inventory_manager_for_invoice(self.manager)
        if inventory_manager:
            sync_ok, sync_msg, inventory_links = inventory_manager.sync_invoice_to_inventory(invoice_dict)
            if not sync_ok:
                override = messagebox.askyesno(
                    "Inventory Check",
                    f"{sync_msg}\n\n"
                    f"Invoice save would normally stop here because inventory is insufficient.\n"
                    f"Proceed anyway?  The invoice will still save, but NO inventory will be deducted.\n\n"
                    f"Create invoice without inventory deduction?",
                    icon="warning"
                )
                if not override:
                    return
                inventory_links = []
                inventory_sync_message = "Saved with override (no stock deduction): " + sync_msg
            else:
                invoice_dict['inventory_links'] = inventory_links
                inventory_sync_message = sync_msg

        # Use the manager to create the invoice properly using the new method
        success = self.manager.add_invoice_from_dict(invoice_dict)
        if not success and inventory_manager and invoice_dict.get('inventory_links'):
            try:
                inventory_manager.reverse_invoice_sales(invoice_id)
            except Exception:
                pass
        
        if success:
            currency = invoice_dict.get('currency', 'AED')
            success_message = (
                f"Invoice {invoice_id} created successfully!\n\n"
                f"Client: {client_name}\n"
                f"Emirate: {self.client_emirate.get()}\n"
                f"Date: {invoice_date}\n"
                f"Type: {'Sales' if self.invoice_type.get() == 'sales' else 'Service'}\n"
                f"Total: {currency} {grand_total:.2f}\n"
                f"Cost: {currency} {total_cost:.2f}\n"
                f"P/L: {currency} {profit_loss:.2f}\n"
                f"VAT: {currency} {vat_amount:.2f}"
            )
            if inventory_sync_message:
                success_message += f"\n\nInventory: {inventory_sync_message}"
            messagebox.showinfo("Success", success_message)
            self.dialog.destroy()
        else:
            messagebox.showerror("Error", "Failed to create invoice. Please try again.")

    def _compute_due_date(self, base_date_str, payment_due_str):
        try:
            base = datetime.strptime(base_date_str, "%Y-%m-%d")
        except Exception:
            base = datetime.now()
        s = (payment_due_str or "On receipt").lower()
        if s.startswith("on receipt"):
            return base.strftime("%Y-%m-%d")
        days = 0
        if s.startswith("30"):
            days = 30
        elif s.startswith("60"):
            days = 60
        elif s.startswith("90"):
            days = 90
        elif s.startswith("120"):
            days = 120
        return (base + timedelta(days=days)).strftime("%Y-%m-%d")

    def edit_item_dialog(self):
        sel = self.items_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select an item to edit")
            return
        item_id = sel[0]
        vals = self.items_tree.item(item_id)['values']
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Edit Item")
        dlg.geometry("420x260")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Description:").grid(row=0, column=0, sticky='w')
        desc = ttk.Entry(frm, width=24)
        desc.grid(row=0, column=1, padx=8, pady=4)
        desc.insert(0, vals[0])
        ttk.Label(frm, text="Qty:").grid(row=1, column=0, sticky='w')
        qty = ttk.Entry(frm, width=8)
        qty.grid(row=1, column=1, padx=8, pady=4)
        qty.insert(0, str(vals[1]))
        ttk.Label(frm, text="Unit Price:").grid(row=2, column=0, sticky='w')
        price = ttk.Entry(frm, width=10)
        price.grid(row=2, column=1, padx=8, pady=4)
        price.insert(0, str(vals[2]))
        ttk.Label(frm, text="VAT:").grid(row=3, column=0, sticky='w')
        vat = ttk.Combobox(frm, values=["Yes","No"], state='readonly', width=10)
        vat.grid(row=3, column=1, padx=8, pady=4)
        vat.set(vals[3])
        def save():
            try:
                q = float(qty.get())
                p = float(price.get())
                if q <= 0 or p < 0:
                    raise ValueError
                total = q * p
                self.items_tree.item(item_id, values=(desc.get().strip(), q, f"{p:.2f}", vat.get(), f"{total:.2f}"))
                self.update_totals_preview()
                dlg.destroy()
            except Exception:
                messagebox.showwarning("Warning", "Please enter valid quantity and price")
        ttk.Button(frm, text="Save", command=save).grid(row=4, column=0, pady=8)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=4, column=1, pady=8)
        _fit_window(dlg, 420, 260, mode="compact", remember_key="invoice_edit_item")

    def edit_cost_dialog(self):
        sel = self.costs_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a cost to edit")
            return
        item_id = sel[0]
        vals = self.costs_tree.item(item_id)['values']
        has_exp = len(vals) >= 5
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Edit Cost")
        dlg.geometry("420x320")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Description:").grid(row=0, column=0, sticky='w')
        desc = ttk.Entry(frm, width=24)
        desc.grid(row=0, column=1, padx=8, pady=4)
        desc.insert(0, vals[0])
        ttk.Label(frm, text="Amount:").grid(row=1, column=0, sticky='w')
        amount = ttk.Entry(frm, width=12)
        amount.grid(row=1, column=1, padx=8, pady=4)
        try:
            amt_init = str(float(vals[1]))
        except Exception:
            amt_init = str(vals[1])
        amount.insert(0, amt_init)
        ttk.Label(frm, text="Account (Paid From):").grid(row=2, column=0, sticky='w')
        acc = ttk.Combobox(frm, width=18)
        try:
            from balance_manager import BalanceManager
            bm = BalanceManager(self.manager.invoice_folder)
            acc['values'] = bm.get_account_names()
        except Exception:
            acc['values'] = ["Cash"]
        acc.set(vals[2] if len(vals) > 2 and vals[2] else "Cash")
        ttk.Label(frm, text="Expense GL:").grid(row=3, column=0, sticky='w')
        exp_gl = ttk.Combobox(frm, width=26, state="readonly")
        _ea = ["5000  COGS (Cost of Goods Sold)",
               "5100  Salary / Wages",
               "5200  Rent, Utilities",
               "5300  Office Supplies, Travel, Marketing, Other Admin",
               "5400  Professional Fees, Audit, Legal",
               "5500  Delivery & Shipping",
               "5600  Medical / Clinical / Device Service Fees",
               "5700  Commissions Paid to Sales Agents",
               "5800  Warranty / Returns / After-Sales",
               "5900  Other Operating Expenses"]
        exp_gl['values'] = _ea
        init_exp = vals[3] if has_exp else ""
        exp_gl_str = str(init_exp or "").strip()
        found_idx = 0
        for i, lab in enumerate(_ea):
            code = lab.split(None, 1)[0]
            if exp_gl_str and (exp_gl_str == code or exp_gl_str.startswith(code)):
                found_idx = i
                break
        exp_gl.current(found_idx)
        exp_gl.grid(row=3, column=1, padx=8, pady=4)
        def save():
            try:
                a = float(amount.get())
                if a < 0:
                    raise ValueError
            except Exception:
                messagebox.showwarning("Warning", "Please enter valid cost amount")
                return
            new_desc = desc.get().strip()
            new_acc = acc.get().strip() or "Cash"
            exp_code = exp_gl.get().strip().split(None, 1)[0] if exp_gl.get().strip() else "5000"
            old_amt = float(vals[1])
            old_acc = vals[2]
            old_exp = (vals[3] if has_exp else "5000") or "5000"

            # Financial-impact confirmation before committing edit
            old_total = sum(float(c.get('amount', 0)) for c in self._collect_costs_for_save())
            new_total = old_total - old_amt + a
            try:
                gt = self._preview_grand_total()
            except Exception:
                gt = 0.0
            confirmed = _confirm_cost_change(
                self.dialog,
                action_title="Edit Cost",
                summary_text=(
                    f"Editing cost: '{vals[0]}' → '{new_desc}'\n"
                    f"Amount: {old_amt:,.2f} → {a:,.2f}  |  Fund: {old_acc} → {new_acc}"
                ),
                old_total=old_total,
                new_total=new_total,
                grand_total=gt,
                expense_account=exp_code,
                fund_account=new_acc,
                line_items_preview=None,
            )
            if not confirmed:
                return
            if has_exp:
                self.costs_tree.item(item_id, values=(new_desc, f"{a:.2f}", new_acc, exp_code, vals[4] if len(vals) > 4 else ""))
            else:
                self.costs_tree.item(item_id, values=(new_desc, f"{a:.2f}", new_acc, vals[3] if len(vals) > 3 else ""))
            self.update_totals_preview()
            dlg.destroy()
        ttk.Button(frm, text="Save", command=save).grid(row=4, column=0, pady=8)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=4, column=1, pady=8)
        _fit_window(dlg, 420, 320, mode="compact", remember_key="invoice_edit_cost")

class EditInvoiceDialog:
    def __init__(self, parent, manager, invoice_id, invoice, data_memory):
        self.manager = manager
        self.invoice_id = invoice_id
        self.invoice = invoice
        self.data_memory = data_memory
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Edit Invoice - {invoice_id}")
        self.dialog.geometry("1100x720")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        self.load_invoice_data()
        _fit_window(self.dialog, 800, 520, mode="workspace", remember_key="edit_invoice")
    
    def setup_ui(self):
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text=f"Edit Invoice: {self.invoice_id}", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Invoice Type
        type_frame = ttk.LabelFrame(main_frame, text="Invoice Type", padding="10")
        type_frame.pack(fill='x', pady=5)
        
        self.invoice_type = tk.StringVar(value=self.invoice.get('invoice_type', 'sales'))
        ttk.Radiobutton(type_frame, text="Sales Invoice", variable=self.invoice_type, value="sales").pack(side='left', padx=10)
        ttk.Radiobutton(type_frame, text="Service Invoice", variable=self.invoice_type, value="service").pack(side='left', padx=10)
        
        # Client Information
        client_frame = ttk.LabelFrame(main_frame, text="Client Information", padding="10")
        client_frame.pack(fill='x', pady=5)
        
        ttk.Label(client_frame, text="Client Name *:").grid(row=0, column=0, sticky='w', pady=5)
        self.client_name = ttk.Combobox(client_frame, width=30)
        self.client_name.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        # Load existing clients
        clients = self.data_memory.get_clients()
        self.client_name['values'] = [client['name'] for client in clients]
        self.client_name.bind('<KeyRelease>', self.on_client_search)
        
        ttk.Label(client_frame, text="Client TRN:").grid(row=1, column=0, sticky='w', pady=5)
        self.client_trn = ttk.Entry(client_frame, width=30)
        self.client_trn.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        
        # ADDED: Emirate Selection
        ttk.Label(client_frame, text="Location (City, Country):").grid(row=2, column=0, sticky='w', pady=5)
        self.client_emirate = ttk.Combobox(client_frame, 
                                         values=[
                                             "Dubai, United Arab Emirates", "Abu Dhabi, United Arab Emirates",
                                             "Sharjah, United Arab Emirates", "Ajman, United Arab Emirates",
                                             "Umm Al Quwain, United Arab Emirates", "Ras Al Khaimah, United Arab Emirates",
                                             "Fujairah, United Arab Emirates",
                                             "Cairo, Egypt", "Alexandria, Egypt", "Riyadh, Saudi Arabia",
                                             "Jeddah, Saudi Arabia", "Doha, Qatar", "Kuwait City, Kuwait",
                                             "Muscat, Oman"
                                         ],
                                         state="normal", width=27)
        self.client_emirate.set("")
        self.client_emirate.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        
        # Invoice Details
        details_frame = ttk.LabelFrame(main_frame, text="Invoice Details", padding="10")
        details_frame.pack(fill='x', pady=5)
        
        # ADDED: Date field
        ttk.Label(details_frame, text="Invoice Date:").grid(row=0, column=0, sticky='w', pady=5)
        self.invoice_date = ttk.Entry(details_frame, width=15)
        self.invoice_date.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Payment Method:").grid(row=1, column=0, sticky='w', pady=5)
        self.payment_method = ttk.Combobox(details_frame,
                                           values=["Cash", "Transfer"],
                                           state="readonly", width=27)
        self.payment_method.set("Cash")
        self.payment_method.grid(row=1, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(details_frame, text="Payment Due:").grid(row=2, column=0, sticky='w', pady=5)
        self.payment_due = ttk.Combobox(details_frame,
                                        values=["On receipt", "30 Days", "60 Days", "90 Days", "120 Days"],
                                        state="readonly", width=27)
        self.payment_due.set("On receipt")
        self.payment_due.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Invoice Status:").grid(row=3, column=0, sticky='w', pady=5)
        self.workflow_status = ttk.Combobox(details_frame,
                                            values=["Under Process", "Closed"],
                                            state="readonly", width=27)
        self.workflow_status.set("Under Process")
        self.workflow_status.grid(row=3, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(details_frame, text="Currency:").grid(row=4, column=0, sticky='w', pady=5)
        self.currency_var = tk.StringVar(value=str(self.invoice.get('currency', 'AED') or 'AED'))
        self.currency_combo = ttk.Combobox(
            details_frame,
            textvariable=self.currency_var,
            values=["AED", "USD", "EUR", "GBP", "SAR", "QAR", "OMR", "KWD", "BHD", "EGP"],
            state="normal",
            width=10
        )
        self.currency_combo.grid(row=4, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(details_frame, text="VAT Rate %:").grid(row=5, column=0, sticky='w', pady=5)
        vat_frame = ttk.Frame(details_frame)
        vat_frame.grid(row=5, column=1, sticky='w', pady=5, padx=10)
        
        self.vat_rate_var = tk.StringVar(value="0")
        self.vat_rate = ttk.Combobox(vat_frame, textvariable=self.vat_rate_var, values=["0", "5", "10", "15"], state="readonly", width=5)
        self.vat_rate.pack(side='left')
        
        ttk.Label(vat_frame, text="%").pack(side='left', padx=(5, 0))
        
        # Items Section
        items_frame = ttk.LabelFrame(main_frame, text="Items / Services", padding="10")
        items_frame.pack(fill='both', expand=True, pady=5)
        
        # Items list with scrollbar
        tree_frame = ttk.Frame(items_frame)
        tree_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        columns = ('Description', 'Quantity', 'Unit Price', 'VAT', 'Total')
        self.items_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=6)
        
        for col in columns:
            self.items_tree.heading(col, text=col)
            self.items_tree.column(col, width=100, stretch=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=scrollbar.set)
        
        self.items_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        try:
            self.items_tree.bind('<Configure>', lambda e: self._resize_items_tree_columns())
            self.dialog.after(0, self._resize_items_tree_columns)
        except Exception:
            pass
        
        # Add item controls
        add_frame = ttk.Frame(items_frame)
        add_frame.pack(fill='x')
        
        ttk.Label(add_frame, text="Description:").grid(row=0, column=0, padx=(0, 5))
        self.item_desc = ttk.Combobox(add_frame, width=15)
        self.item_desc.grid(row=0, column=1, padx=(0, 10))
        # Load existing products
        products = self.data_memory.get_products()
        self.item_desc['values'] = [product['name'] for product in products]
        self.item_desc.bind('<KeyRelease>', self.on_product_search)
        
        ttk.Label(add_frame, text="Qty:").grid(row=0, column=2, padx=(0, 5))
        self.item_qty = ttk.Entry(add_frame, width=6)
        self.item_qty.grid(row=0, column=3, padx=(0, 10))
        self.item_qty.insert(0, "1")
        
        self.unit_price_label = ttk.Label(add_frame, text="Unit Price (AED):")
        self.unit_price_label.grid(row=0, column=4, padx=(0, 5))
        self.item_price = ttk.Entry(add_frame, width=10)
        self.item_price.grid(row=0, column=5, padx=(0, 10))
        
        ttk.Label(add_frame, text="VAT:").grid(row=0, column=6, padx=(0, 5))
        self.item_vat = ttk.Combobox(add_frame, values=["Yes", "No"], state="readonly", width=5)
        self.item_vat.set("No")
        self.item_vat.grid(row=0, column=7, padx=(0, 10))
        
        ttk.Button(add_frame, text="Add Item", command=self.add_item).grid(row=0, column=8, padx=(10, 0))
        ttk.Button(add_frame, text="Edit Item", command=self.edit_item_dialog).grid(row=0, column=9, padx=5)
        ttk.Button(add_frame, text="Remove", command=self.remove_item).grid(row=0, column=10, padx=5)
        ttk.Button(add_frame, text="Import Excel", command=self.import_items_from_excel).grid(row=0, column=11, padx=5)

        # Invoice Costs Section
        costs_frame = ttk.LabelFrame(main_frame, text="Invoice Costs (Internal Only)", padding="10")
        costs_frame.pack(fill='x', pady=5)

        # --- Cost Center applicator band was here; moved to InvoiceManager main menu (Cost Centers & Bulk Backfill) ---

        # Costs list with scrollbar
        costs_tree_frame = ttk.Frame(costs_frame)
        costs_tree_frame.pack(fill='both', expand=True, pady=(0, 10))

        cost_columns = ('Description', 'Amount', 'Account', 'Expense GL', 'Receipt')
        self.costs_tree = ttk.Treeview(costs_tree_frame, columns=cost_columns, show='headings', height=2)

        for col in cost_columns:
            self.costs_tree.heading(col, text=col)
            if col == 'Description':
                self.costs_tree.column(col, width=190)
            elif col == 'Amount':
                self.costs_tree.column(col, width=95)
            elif col == 'Account':
                self.costs_tree.column(col, width=120)
            elif col == 'Expense GL':
                self.costs_tree.column(col, width=110)
            else:
                self.costs_tree.column(col, width=70)

        costs_scrollbar = ttk.Scrollbar(costs_tree_frame, orient=tk.VERTICAL, command=self.costs_tree.yview)
        self.costs_tree.configure(yscrollcommand=costs_scrollbar.set)

        self.costs_tree.pack(side='left', fill='both', expand=True)
        costs_scrollbar.pack(side='right', fill='y')

        # Add cost controls
        add_cost_frame = ttk.Frame(costs_frame)
        add_cost_frame.pack(fill='x')

        ttk.Label(add_cost_frame, text="Cost Description:").grid(row=0, column=0, padx=(0, 5))
        self.cost_desc = ttk.Entry(add_cost_frame, width=20)
        self.cost_desc.grid(row=0, column=1, padx=(0, 10))

        self.cost_amount_label = ttk.Label(add_cost_frame, text="Amount (AED):")
        self.cost_amount_label.grid(row=0, column=2, padx=(0, 5))
        self.cost_amount = ttk.Entry(add_cost_frame, width=10)
        self.cost_amount.grid(row=0, column=3, padx=(0, 10))

        ttk.Label(add_cost_frame, text="Paid From:").grid(row=0, column=4, padx=(0, 5))
        self.cost_account = ttk.Combobox(add_cost_frame, width=15, state="readonly")
        try:
            from balance_manager import BalanceManager
            bm = BalanceManager(self.manager.invoice_folder)
            accounts = bm.get_account_names() or ["Cash"]
            self.cost_account['values'] = ["Select Account"] + accounts
        except Exception:
            self.cost_account['values'] = ["Select Account", "Cash", "ADCB"]
        try:
            last = getattr(self.manager, "last_selected_account", None)
            if last and last in self.cost_account['values']:
                self.cost_account.set(last)
            else:
                self.cost_account.set("Select Account")
        except Exception:
            self.cost_account.set("Select Account")
        self.cost_account.grid(row=0, column=5, padx=(0, 10))

        ttk.Label(add_cost_frame, text="Expense GL:").grid(row=0, column=6, padx=(0, 5))
        self.cost_expense_account = ttk.Combobox(add_cost_frame, width=22, state="readonly")
        _ea_labels_edit = ["5000  COGS (Cost of Goods Sold)",
                           "5100  Salary / Wages",
                           "5200  Rent, Utilities (Electricity, Water, Internet)",
                           "5300  Office Supplies, Travel, Marketing, Other Admin",
                           "5400  Professional Fees, Audit, Legal",
                           "5500  Delivery & Shipping / Logistics",
                           "5600  Medical / Clinical / Device Service Fees",
                           "5700  Commissions Paid to Sales Agents",
                           "5800  Warranty / Returns / After-Sales",
                           "5900  Other Operating Expenses"]
        self.cost_expense_account['values'] = _ea_labels_edit
        self.cost_expense_account.current(0)
        self.cost_expense_account.grid(row=0, column=7, padx=(0, 10))

        ttk.Button(add_cost_frame, text="Add Cost", command=self.add_cost).grid(row=0, column=8, padx=(10, 5))
        ttk.Button(add_cost_frame, text="Edit Cost", command=self.edit_cost_dialog).grid(row=0, column=9, padx=5)
        ttk.Button(add_cost_frame, text="Upload Receipt", command=self.upload_receipt).grid(row=0, column=10, padx=5)
        ttk.Button(add_cost_frame, text="Remove Cost", command=self.remove_cost).grid(row=0, column=11, padx=5)

        # Totals preview
        totals_frame = ttk.Frame(main_frame)
        totals_frame.pack(fill='x', pady=5)

        self.totals_label = ttk.Label(totals_frame,
                                     text="Subtotal: AED 0.00 | VAT: AED 0.00 | Total: AED 0.00 | Total Cost: AED 0.00 | P/L: AED 0.00",
                                     font=('Helvetica', 9, 'bold'))
        self.totals_label.pack()

        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        ttk.Button(button_frame, text="Save Changes", command=self.save_changes).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy).pack(side='left', padx=5)
        
        notes_frame = ttk.LabelFrame(main_frame, text="Notes", padding="10")
        notes_frame.pack(fill='x', pady=5)
        
        self.notes_text = scrolledtext.ScrolledText(notes_frame, height=3)
        self.notes_text.pack(fill='x')

        try:
            self.currency_var.trace('w', lambda *args: self.update_currency_ui())
        except Exception:
            pass
        try:
            self.vat_rate_var.trace('w', lambda *args: self.update_totals_preview())
        except Exception:
            pass
        self.update_currency_ui()
    
    def load_invoice_data(self):
        """Load existing invoice data into form"""
        # Client information
        self.client_name.insert(0, self.invoice.get('client_name', ''))
        self.client_trn.insert(0, self.invoice.get('client_trn', ''))
        try:
            loc = self.invoice.get('client_location')
            if not loc:
                loc = str(self.invoice.get('client_emirate', '') or '').strip()
            if not loc:
                loc = ""
            uae_emirates = {
                'Dubai',
                'Abu Dhabi',
                'Sharjah',
                'Ajman',
                'Umm Al Quwain',
                'Ras Al Khaimah',
                'Fujairah',
            }
            if ',' not in loc and loc in uae_emirates:
                loc = f"{loc}, United Arab Emirates"
            self.client_emirate.set(loc)
        except Exception:
            self.client_emirate.set(str(self.invoice.get('client_emirate', '') or '').strip())
        
        # Invoice details
        self.invoice_date.insert(0, self.invoice.get('date', datetime.now().strftime("%Y-%m-%d")))
        pm = self.invoice.get('payment_method', self.invoice.get('payment_terms', 'Cash'))
        pd = self.invoice.get('payment_due', 'On receipt')
        self.payment_method.set(pm)
        self.payment_due.set(pd)
        self.vat_rate_var.set(str(self.invoice.get('tax_rate', 0)))
        try:
            self.currency_var.set(str(self.invoice.get('currency', 'AED') or 'AED'))
        except Exception:
            pass
        self.update_currency_ui()
        
        # Load items
        for item in self.invoice.get('items', []):
            vat_text = "Yes" if item.get('taxable', True) else "No"
            self.items_tree.insert('', tk.END, values=(
                item.get('description', ''),
                item.get('quantity', 0),
                f"{item.get('unit_price', 0):.2f}",
                vat_text,
                f"{item.get('total', 0):.2f}"
            ))
        try:
            self._adjust_tree_heights()
        except Exception:
            pass
        
        # Load costs
        for cost in self.invoice.get('costs', []):
            receipt_indicator = "📎" if cost.get('receipt_attached') else ""
            exp_code_raw = str(cost.get('expense_account') or "").strip()
            exp_code = exp_code_raw
            if len(exp_code) > 4 and exp_code[:4].isdigit():
                exp_code = exp_code[:4]
            if not (len(exp_code) == 4 and exp_code.isdigit()):
                exp_code = "5000"
            cols = len(self.costs_tree["columns"])
            if cols == 5:
                self.costs_tree.insert('', tk.END, values=(
                    cost.get('description', ''),
                    f"{cost.get('amount', 0):.2f}",
                    cost.get('account', 'Cash'),
                    exp_code,
                    receipt_indicator,
                ))
            else:
                self.costs_tree.insert('', tk.END, values=(
                    cost.get('description', ''),
                    f"{cost.get('amount', 0):.2f}",
                    cost.get('account', 'Cash'),
                    receipt_indicator
                ))
        try:
            self._adjust_tree_heights()
        except Exception:
            pass
        
        # Load notes
        if self.invoice.get('notes'):
            self.notes_text.insert('1.0', self.invoice.get('notes', ''))
        
        # Update totals preview
        self.update_totals_preview()
    
    def on_client_search(self, event=None):
        """Handle client search in combobox"""
        value = self.client_name.get().lower()
        if value == '':
            clients = self.data_memory.get_clients()
            self.client_name['values'] = [client['name'] for client in clients]
        else:
            data = []
            clients = self.data_memory.search_clients(value)
            for client in clients:
                data.append(client['name'])
            self.client_name['values'] = data
    
    def on_product_search(self, event=None):
        """Handle product search in combobox"""
        value = self.item_desc.get().lower()
        if value == '':
            products = self.data_memory.get_products()
            self.item_desc['values'] = [product['name'] for product in products]
        else:
            data = []
            products = self.data_memory.search_products(value)
            for product in products:
                data.append(product['name'])
            self.item_desc['values'] = data

    def _parse_invoice_items_excel(self, file_path):
        warnings = []
        errors = []
        if not OPENPYXL_AVAILABLE:
            errors.append("Excel library (openpyxl) is not available in this build.")
            return [], warnings, errors

        def norm(value):
            """Enhanced normalization: lowercase, strip, remove special chars, collapse whitespace"""
            if value is None:
                return ""
            import re
            s = str(value).strip().lower()
            s = re.sub(r'[^a-z0-9\s]', '', s)
            s = re.sub(r'\s+', ' ', s).strip()
            return s

        def fuzzy_match(a, b, threshold=0.8):
            """Fuzzy string match using difflib SequenceMatcher"""
            import difflib
            if not a or not b:
                return False, 0.0
            ratio = difflib.SequenceMatcher(None, a, b).ratio()
            return ratio >= threshold, ratio

        def to_float(value, default=0.0):
            try:
                if value is None:
                    return default
                if isinstance(value, str):
                    cleaned = value.strip().replace(",", "")
                    filtered = "".join(ch for ch in cleaned if ch.isdigit() or ch in ".-")
                    if filtered and any(ch.isdigit() for ch in filtered):
                        cleaned = filtered
                    if not cleaned:
                        return default
                    return float(cleaned)
                return float(value)
            except Exception:
                return default

        def vat_to_yes_no(value):
            if value is None:
                return "No"
            if isinstance(value, (int, float)):
                return "Yes" if float(value) > 0 else "No"
            text = norm(value)
            if not text:
                return "No"
            if text in {"yes", "y", "true", "1", "vat", "tax", "taxable", "included"}:
                return "Yes"
            if text in {"no", "n", "false", "0", "non-taxable", "without vat", "excluded"}:
                return "No"
            return "Yes" if to_float(text, 0.0) > 0 else "No"

        try:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as exc:
            errors.append(f"Error reading Excel file: {exc}")
            return [], warnings, errors

        header_keywords = {
            'description': ['description', 'item', 'item description', 'product', 'service', 'details', 'name'],
            'quantity': ['qty', 'quantity', 'qnty', 'count'],
            'actual_rate': ['actual rate', 'actual unit price', 'actual price'],
            'rate': ['unit price', 'rate', 'price', 'unit cost', 'amount per unit'],
            'total': ['total', 'amount', 'line total', 'net'],
            'vat': ['vat', 'tax', 'vat %', 'tax %', 'vat amount', 'tax amount', 'taxable'],
        }

        worksheet = workbook.active
        header_row_idx = None
        header_map = {}
        best_score = 0
        max_scan_cols = 0
        potential_mismatches = []
        for candidate_sheet in workbook.worksheets:
            candidate_max_scan_rows = min(50, candidate_sheet.max_row or 0)
            candidate_max_scan_cols = min(60, candidate_sheet.max_column or 0)
            for row_index in range(1, candidate_max_scan_rows + 1):
                row_text = [norm(candidate_sheet.cell(row=row_index, column=col_index).value) for col_index in range(1, candidate_max_scan_cols + 1)]
                score = 0
                current_map = {}
                used_cols = set()
                # First try exact matches
                for field, keywords in header_keywords.items():
                    found = False
                    for col_index, cell_text in enumerate(row_text, start=1):
                        if col_index in used_cols or not cell_text:
                            continue
                        for keyword in keywords:
                            norm_keyword = norm(keyword)
                            if cell_text == norm_keyword:
                                current_map[field] = col_index
                                used_cols.add(col_index)
                                score += 10  # Exact match gets high score
                                found = True
                                break
                        if found:
                            break
                    if found:
                        continue
                    # Then try fuzzy matches
                    best_ratio = 0
                    best_col = None
                    for col_index, cell_text in enumerate(row_text, start=1):
                        if col_index in used_cols or not cell_text:
                            continue
                        for keyword in keywords:
                            norm_keyword = norm(keyword)
                            is_match, ratio = fuzzy_match(cell_text, norm_keyword, threshold=0.7)
                            if is_match and ratio > best_ratio:
                                best_ratio = ratio
                                best_col = col_index
                    if best_col is not None:
                        current_map[field] = best_col
                        used_cols.add(best_col)
                        score += best_ratio
                        if best_ratio < 0.9:
                            potential_mismatches.append(f"Potential header mismatch in sheet '{candidate_sheet.title}' row {row_index}: column {best_col} mapped to '{field}' (similarity: {best_ratio:.2f})")
                if score > best_score:
                    best_score = score
                    worksheet = candidate_sheet
                    header_row_idx = row_index
                    header_map = current_map
                    max_scan_cols = candidate_max_scan_cols
        
        if potential_mismatches:
            warnings.extend(potential_mismatches)

        if not header_row_idx or 'description' not in header_map:
            errors.append("Could not detect the invoice items table in the Excel file. Make sure it has a Description or Item column.")
            return [], warnings, errors

        imported_items = []
        blank_streak = 0
        stop_labels = (
            "subtotal", "sub total", "total", "grand total", "invoice total",
            "amount due", "net total", "balance due", "tax", "vat",
            "tax amount", "vat amount"
        )
        for row_index in range(header_row_idx + 1, (worksheet.max_row or header_row_idx) + 1):
            desc_value = worksheet.cell(row=row_index, column=header_map['description']).value
            description = str(desc_value).strip() if desc_value is not None else ""
            if not description:
                empty_row = True
                for check_col in range(1, max(1, max_scan_cols) + 1):
                    if worksheet.cell(row=row_index, column=check_col).value not in (None, ""):
                        empty_row = False
                        break
                if empty_row:
                    blank_streak += 1
                    if imported_items and blank_streak >= 5:
                        break
                else:
                    blank_streak = 0
                continue
            blank_streak = 0

            normalized_description = norm(description)
            if any(
                normalized_description == label
                or normalized_description.startswith(label + " ")
                or normalized_description.endswith(" " + label)
                for label in stop_labels
            ):
                continue

            quantity = 1.0
            if 'quantity' in header_map:
                quantity = to_float(worksheet.cell(row=row_index, column=header_map['quantity']).value, 1.0)
                if quantity <= 0:
                    quantity = 1.0

            unit_price = 0.0
            if 'actual_rate' in header_map:
                unit_price = to_float(worksheet.cell(row=row_index, column=header_map['actual_rate']).value, 0.0)
            elif 'rate' in header_map:
                unit_price = to_float(worksheet.cell(row=row_index, column=header_map['rate']).value, 0.0)

            line_total = None
            if 'total' in header_map:
                line_total = to_float(worksheet.cell(row=row_index, column=header_map['total']).value, None)
            if line_total is None:
                line_total = quantity * unit_price
            elif unit_price == 0 and quantity:
                unit_price = line_total / quantity

            vat_text = "No"
            if 'vat' in header_map:
                vat_text = vat_to_yes_no(worksheet.cell(row=row_index, column=header_map['vat']).value)

            imported_items.append({
                'description': description,
                'quantity': float(quantity),
                'unit_price': float(unit_price),
                'total': float(line_total),
                'vat': vat_text,
            })

        if not imported_items:
            warnings.append("No invoice items were found under the detected Excel header row.")
        return imported_items, warnings, errors

    def import_items_from_excel(self):
        file_path = filedialog.askopenfilename(
            title="Import Invoice Items From Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not file_path:
            return

        imported_items, warnings, errors = self._parse_invoice_items_excel(file_path)
        if errors and not imported_items:
            messagebox.showerror("Import Excel", "\n".join(errors))
            return

        if self.items_tree.get_children() and imported_items:
            replace_existing = messagebox.askyesno(
                "Import Excel",
                "The invoice already has items.\n\nClick Yes to replace them, or No to append the imported items."
            )
            if replace_existing:
                for item_id in self.items_tree.get_children():
                    self.items_tree.delete(item_id)

        imported_count = 0
        for item in imported_items:
            description = str(item.get('description') or "").strip()
            if not description:
                continue
            quantity = float(item.get('quantity', 0) or 0)
            unit_price = float(item.get('unit_price', 0) or 0)
            line_total = float(item.get('total', quantity * unit_price) or 0)
            vat_text = "Yes" if str(item.get('vat', 'No')).strip().lower() == "yes" else "No"
            self.items_tree.insert('', tk.END, values=(
                description,
                quantity,
                f"{unit_price:.2f}",
                vat_text,
                f"{line_total:.2f}"
            ))
            try:
                self.data_memory.add_product(description, unit_price)
            except Exception:
                pass
            imported_count += 1

        if imported_count:
            self.update_totals_preview()
            self._adjust_tree_heights()

        message_lines = [f"Imported {imported_count} item(s) from Excel."]
        if warnings:
            message_lines.append("")
            message_lines.append("Warnings:")
            message_lines.extend(warnings)
        if errors:
            message_lines.append("")
            message_lines.append("Notes:")
            message_lines.extend(errors)
        messagebox.showinfo("Import Excel", "\n".join(message_lines))
    
    def add_item(self):
        """Add item to the items list"""
        desc = self.item_desc.get().strip()
        qty = self.item_qty.get().strip()
        price = self.item_price.get().strip()
        vat_applicable = self.item_vat.get() == "Yes"
        
        if not desc:
            messagebox.showwarning("Warning", "Please enter item description")
            return
            
        try:
            qty_val = float(qty)
            price_val = float(price)
            if qty_val <= 0 or price_val < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Warning", "Please enter valid quantity and price")
            return
        
        # Add to product memory
        self.data_memory.add_product(desc, price_val)
            
        total = qty_val * price_val
        vat_text = "Yes" if vat_applicable else "No"
        self.items_tree.insert('', tk.END, values=(
            desc, 
            qty_val, 
            f"{price_val:.2f}", 
            vat_text, 
            f"{total:.2f}"
        ))
        
        # Clear fields
        self.item_desc.delete(0, tk.END)
        self.item_qty.delete(0, tk.END)
        self.item_qty.insert(0, "1")
        self.item_price.delete(0, tk.END)
        
        self.update_totals_preview()
        try:
            self._adjust_tree_heights()
        except Exception:
            pass
    
    def remove_item(self):
        """Remove selected item"""
        selection = self.items_tree.selection()
        if selection:
            self.items_tree.delete(selection[0])
            self.update_totals_preview()
            try:
                self._adjust_tree_heights()
            except Exception:
                pass
    
    def add_cost(self):
        """Add cost to the costs list — confirm with financial impact before applying."""
        desc = self.cost_desc.get().strip()
        amount = self.cost_amount.get().strip()
        account_val = self.cost_account.get().strip()
        expense_gl = getattr(self, "cost_expense_account", None)
        expense_gl_val = ""
        if expense_gl is not None:
            raw = expense_gl.get().strip()
            if raw:
                expense_gl_val = (raw.split(None, 1)[0] if raw[:4].isdigit() else "5000")

        if not desc:
            messagebox.showwarning("Warning", "Please enter cost description")
            return

        try:
            amount_val = float(amount)
            if amount_val < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Warning", "Please enter valid cost amount")
            return

        if not account_val or account_val == "Select Account":
            messagebox.showwarning("Warning", "Please select an account for this cost")
            return

        try:
            setattr(self.manager, "last_selected_account", account_val)
            self.cost_account.set(account_val)
        except Exception:
            pass

        old_total = sum(float(c.get('amount', 0)) for c in self._collect_costs_for_save())
        new_total = old_total + amount_val
        try:
            gt = self._preview_grand_total()
        except Exception:
            gt = 0.0

        single_line = [(desc, amount_val, expense_gl_val or "5000", account_val)]
        confirmed = _confirm_cost_change(
            self.dialog,
            action_title="Add Manual Cost",
            summary_text=f"Adding cost '{desc}' for {amount_val:,.2f} AED from '{account_val}'.",
            old_total=old_total,
            new_total=new_total,
            grand_total=gt,
            expense_account=expense_gl_val or "5000",
            fund_account=account_val,
            line_items_preview=single_line,
        )
        if not confirmed:
            return

        cols_count = len(self.costs_tree["columns"])
        if cols_count == 5:
            self.costs_tree.insert('', tk.END, values=(desc, f"{amount_val:.2f}", account_val, expense_gl_val, ""))
        else:
            self.costs_tree.insert('', tk.END, values=(desc, f"{amount_val:.2f}", account_val, ""))

        # Clear fields
        self.cost_desc.delete(0, tk.END)
        self.cost_amount.delete(0, tk.END)

        self.update_totals_preview()
        try:
            self._adjust_tree_heights()
        except Exception:
            pass

    def _preview_grand_total(self):
        """Helper to compute grand_total without side effects."""
        subtotal = 0.0
        for item in self.items_tree.get_children():
            vals = self.items_tree.item(item)['values']
            try:
                qty = float(vals[1])
                price = float(vals[2])
                vat = str(vals[3]).strip().lower() == 'yes'
                line = qty * price
                if vat:
                    line *= 1.05
                subtotal += line
            except Exception:
                continue
        return subtotal

    def _collect_costs_for_save(self):
        """Helper: extract current costs list from tree as dict list."""
        costs = []
        for item in self.costs_tree.get_children():
            vals = self.costs_tree.item(item)['values']
            try:
                if len(vals) >= 5:
                    desc, amt, acc, exp_gl, receipt = (vals[0], vals[1], vals[2], vals[3], vals[4])
                elif len(vals) == 4:
                    desc, amt, acc, receipt = vals
                    exp_gl = ""
                else:
                    continue
                costs.append({
                    'description': str(desc),
                    'amount': float(amt),
                    'account': str(acc),
                    'expense_account': str(exp_gl) if str(exp_gl) not in ("", "None") else "5000",
                    'receipt': str(receipt) if str(receipt) not in ("", "None") else "",
                })
            except Exception:
                continue
        return costs

    def remove_cost(self):
        """Remove selected cost — confirm with financial impact first."""
        selection = self.costs_tree.selection()
        if not selection:
            return
        values = self.costs_tree.item(selection[0])['values']
        try:
            desc = str(values[0])
            amt = float(values[1])
            if len(values) >= 5:
                acc = values[2]
                exp_gl = values[3] if values[3] else "5000"
            else:
                acc = values[2]
                exp_gl = "5000"
        except Exception:
            return

        old_total = sum(float(c.get('amount', 0)) for c in self._collect_costs_for_save())
        new_total = old_total - amt
        try:
            gt = self._preview_grand_total()
        except Exception:
            gt = 0.0

        confirmed = _confirm_cost_change(
            self.dialog,
            action_title="Remove Cost",
            summary_text=f"Removing cost '{desc}' ({amt:,.2f} AED) funded from '{acc}'.",
            old_total=old_total,
            new_total=new_total,
            grand_total=gt,
            expense_account=str(exp_gl),
            fund_account=str(acc),
            line_items_preview=None,
        )
        if not confirmed:
            return

        self.costs_tree.delete(selection[0])
        self.update_totals_preview()
        try:
            self._adjust_tree_heights()
        except Exception:
            pass
    
    def _adjust_tree_heights(self):
        items_count = len(self.items_tree.get_children())
        costs_count = len(self.costs_tree.get_children())
        self.items_tree.configure(height=max(1, min(items_count, 8)))
        self.costs_tree.configure(height=max(1, min(costs_count, 8)))

    def edit_item_dialog(self):
        sel = self.items_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select an item to edit")
            return
        item_id = sel[0]
        vals = self.items_tree.item(item_id)['values']
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Edit Item")
        dlg.geometry("360x220")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Description:").grid(row=0, column=0, sticky='w')
        desc = ttk.Entry(frm, width=24)
        desc.grid(row=0, column=1, padx=8, pady=4)
        desc.insert(0, vals[0])
        ttk.Label(frm, text="Qty:").grid(row=1, column=0, sticky='w')
        qty = ttk.Entry(frm, width=8)
        qty.grid(row=1, column=1, padx=8, pady=4)
        qty.insert(0, str(vals[1]))
        ttk.Label(frm, text="Unit Price:").grid(row=2, column=0, sticky='w')
        price = ttk.Entry(frm, width=10)
        price.grid(row=2, column=1, padx=8, pady=4)
        price.insert(0, str(vals[2]))
        ttk.Label(frm, text="VAT:").grid(row=3, column=0, sticky='w')
        vat = ttk.Combobox(frm, values=["Yes","No"], state='readonly', width=10)
        vat.grid(row=3, column=1, padx=8, pady=4)
        vat.set(vals[3])
        def save():
            try:
                q = float(qty.get())
                p = float(price.get())
                if q <= 0 or p < 0:
                    raise ValueError
                total = q * p
                self.items_tree.item(item_id, values=(desc.get().strip(), q, f"{p:.2f}", vat.get(), f"{total:.2f}"))
                self.update_totals_preview()
                dlg.destroy()
            except Exception:
                messagebox.showwarning("Warning", "Please enter valid quantity and price")
        ttk.Button(frm, text="Save", command=save).grid(row=4, column=0, pady=8)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=4, column=1, pady=8)
        _fit_window(dlg, 360, 220, mode="compact", remember_key="edit_invoice_item")

    def edit_cost_dialog(self):
        sel = self.costs_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a cost to edit")
            return
        item_id = sel[0]
        vals = self.costs_tree.item(item_id)['values']
        has_exp = len(vals) >= 5
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Edit Cost")
        dlg.geometry("420x320")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Description:").grid(row=0, column=0, sticky='w')
        desc = ttk.Entry(frm, width=24)
        desc.grid(row=0, column=1, padx=8, pady=4)
        desc.insert(0, vals[0])
        ttk.Label(frm, text="Amount:").grid(row=1, column=0, sticky='w')
        amount = ttk.Entry(frm, width=12)
        amount.grid(row=1, column=1, padx=8, pady=4)
        try:
            amt_init = str(float(vals[1]))
        except Exception:
            amt_init = str(vals[1])
        amount.insert(0, amt_init)
        ttk.Label(frm, text="Account (Paid From):").grid(row=2, column=0, sticky='w')
        acc = ttk.Combobox(frm, width=18)
        try:
            from balance_manager import BalanceManager
            bm = BalanceManager(self.manager.invoice_folder)
            acc['values'] = bm.get_account_names()
        except Exception:
            acc['values'] = ["Cash"]
        acc.set(vals[2] if len(vals) > 2 and vals[2] else "Cash")
        ttk.Label(frm, text="Expense GL:").grid(row=3, column=0, sticky='w')
        exp_gl = ttk.Combobox(frm, width=26, state="readonly")
        _ea = ["5000  COGS (Cost of Goods Sold)",
               "5100  Salary / Wages",
               "5200  Rent, Utilities",
               "5300  Office Supplies, Travel, Marketing, Other Admin",
               "5400  Professional Fees, Audit, Legal",
               "5500  Delivery & Shipping",
               "5600  Medical / Clinical / Device Service Fees",
               "5700  Commissions Paid to Sales Agents",
               "5800  Warranty / Returns / After-Sales",
               "5900  Other Operating Expenses"]
        exp_gl['values'] = _ea
        init_exp = vals[3] if has_exp else ""
        exp_gl_str = str(init_exp or "").strip()
        found_idx = 0
        for i, lab in enumerate(_ea):
            code = lab.split(None, 1)[0]
            if exp_gl_str and (exp_gl_str == code or exp_gl_str.startswith(code)):
                found_idx = i
                break
        exp_gl.current(found_idx)
        exp_gl.grid(row=3, column=1, padx=8, pady=4)
        def save():
            try:
                a = float(amount.get())
                if a < 0:
                    raise ValueError
            except Exception:
                messagebox.showwarning("Warning", "Please enter valid cost amount")
                return
            new_desc = desc.get().strip()
            new_acc = acc.get().strip() or "Cash"
            exp_code = exp_gl.get().strip().split(None, 1)[0] if exp_gl.get().strip() else "5000"
            old_amt = float(vals[1])

            old_total = sum(float(c.get('amount', 0)) for c in self._collect_costs_for_save())
            new_total = old_total - old_amt + a
            try:
                gt = self._preview_grand_total()
            except Exception:
                gt = 0.0
            confirmed = _confirm_cost_change(
                self.dialog,
                action_title="Edit Cost",
                summary_text=(
                    f"Editing cost: '{vals[0]}' → '{new_desc}'\n"
                    f"Amount: {old_amt:,.2f} → {a:,.2f}  |  Fund: {vals[2]} → {new_acc}"
                ),
                old_total=old_total,
                new_total=new_total,
                grand_total=gt,
                expense_account=exp_code,
                fund_account=new_acc,
                line_items_preview=None,
            )
            if not confirmed:
                return
            if has_exp:
                self.costs_tree.item(item_id, values=(new_desc, f"{a:.2f}", new_acc, exp_code, vals[4] if len(vals) > 4 else ""))
            else:
                self.costs_tree.item(item_id, values=(new_desc, f"{a:.2f}", new_acc, vals[3] if len(vals) > 3 else ""))
            self.update_totals_preview()
            dlg.destroy()
        ttk.Button(frm, text="Save", command=save).grid(row=4, column=0, pady=8)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=4, column=1, pady=8)
        _fit_window(dlg, 420, 320, mode="compact", remember_key="edit_invoice_cost")
    
    def upload_receipt(self):
        """Upload receipt for selected cost"""
        selection = self.costs_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a cost first")
            return

        cost_index = self.costs_tree.index(selection[0])
        file_path = filedialog.askopenfilename(
            title="Select Receipt File",
            filetypes=[
                ("All supported files", "*.pdf *.jpg *.jpeg *.png *.gif *.bmp"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                ("All files", "*.*")
            ]
        )

        if file_path:
            # Upload receipt using manager
            receipt_filename = self.manager.upload_receipt(self.invoice_id, cost_index, file_path)
            if receipt_filename:
                # Update the treeview to show receipt indicator in the last column
                values = list(self.costs_tree.item(selection[0])['values'])
                last_idx = len(values) - 1 if values else 3
                values[last_idx] = "📎"
                self.costs_tree.item(selection[0], values=values)
                messagebox.showinfo("Success", f"Receipt uploaded successfully: {receipt_filename}")
            else:
                messagebox.showerror("Error", "Failed to upload receipt")
    
    def update_totals_preview(self):
        """Update the totals preview"""
        subtotal = 0
        vat_amount = 0
        total_cost = 0
        
        # Calculate items subtotal and VAT
        for item in self.items_tree.get_children():
            values = self.items_tree.item(item)['values']
            item_total = float(values[4])
            subtotal += item_total
            
            # Calculate VAT for taxable items
            if values[3] == "Yes":  # If VAT applicable
                vat_rate = float(self.vat_rate_var.get()) / 100
                vat_amount += item_total * vat_rate
        
        # Calculate total costs
        for cost in self.costs_tree.get_children():
            values = self.costs_tree.item(cost)['values']
            total_cost += float(values[1])
        
        total_amount = subtotal + vat_amount
        profit_loss = total_amount - total_cost

        currency = (getattr(self, "currency_var", None).get() if getattr(self, "currency_var", None) else "AED") or "AED"
        currency = str(currency).strip() or "AED"
        self.totals_label.config(
            text=f"Subtotal: {currency} {subtotal:.2f} | VAT: {currency} {vat_amount:.2f} | Total: {currency} {total_amount:.2f} | Total Cost: {currency} {total_cost:.2f} | P/L: {currency} {profit_loss:.2f}"
        )

    def update_currency_ui(self):
        currency = str(self.currency_var.get() or "").strip() or "AED"
        self.items_tree.heading('Unit Price', text=f"Unit Price ({currency})")
        self.items_tree.heading('Total', text=f"Total ({currency})")
        self.costs_tree.heading('Amount', text=f"Amount ({currency})")
        try:
            self.unit_price_label.config(text=f"Unit Price ({currency}):")
        except Exception:
            pass
        try:
            self.cost_amount_label.config(text=f"Amount ({currency}):")
        except Exception:
            pass
        self.update_totals_preview()

    def _adjust_tree_heights(self):
        try:
            items_count = len(self.items_tree.get_children())
            costs_count = len(self.costs_tree.get_children())
            self.items_tree.configure(height=max(6, min(items_count + 1, 18)))
            self.costs_tree.configure(height=max(4, min(costs_count + 1, 10)))
            self._resize_items_tree_columns()
        except Exception:
            pass

    def _resize_items_tree_columns(self):
        try:
            total_width = self.items_tree.winfo_width()
            if total_width <= 60:
                return
            available = max(0, total_width - 24)
            weights = {
                'Description': 0.48,
                'Quantity': 0.10,
                'Unit Price': 0.15,
                'VAT': 0.07,
                'Total': 0.20,
            }
            mins = {
                'Description': 240,
                'Quantity': 80,
                'Unit Price': 120,
                'VAT': 70,
                'Total': 140,
            }
            widths = {}
            for col, w in weights.items():
                widths[col] = max(mins.get(col, 80), int(available * w))
            leftover = available - sum(widths.values())
            if leftover != 0:
                widths['Description'] = max(mins['Description'], widths['Description'] + leftover)
            for col, w in widths.items():
                self.items_tree.column(col, width=w, minwidth=mins.get(col, 60), stretch=True)
        except Exception:
            pass

    def _compute_due_date(self, base_date_str, payment_due_str):
        try:
            base = datetime.strptime(base_date_str, "%Y-%m-%d")
        except Exception:
            base = datetime.now()
        s = (payment_due_str or "On receipt").lower()
        if s.startswith("on receipt"):
            return base.strftime("%Y-%m-%d")
        days = 0
        if s.startswith("30"):
            days = 30
        elif s.startswith("60"):
            days = 60
        elif s.startswith("90"):
            days = 90
        elif s.startswith("120"):
            days = 120
        return (base + timedelta(days=days)).strftime("%Y-%m-%d")

    def save_changes(self):
        """Save changes to the invoice"""
        client_name = self.client_name.get().strip()
        
        if not client_name:
            messagebox.showwarning("Warning", "Please enter client name")
            return
            
        if len(self.items_tree.get_children()) == 0:
            messagebox.showwarning("Warning", "Please add at least one item")
            return
        
        try:
            vat_rate = float(self.vat_rate_var.get())
        except ValueError:
            messagebox.showwarning("Warning", "Please select a valid VAT rate")
            return
        
        # Validate date
        invoice_date = self.invoice_date.get().strip()
        if not invoice_date:
            invoice_date = datetime.now().strftime("%Y-%m-%d")
        else:
            try:
                datetime.strptime(invoice_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Warning", "Please enter a valid date (YYYY-MM-DD)")
                return
        
        # Add client to memory
        self.data_memory.add_client(client_name, self.client_trn.get().strip())
        
        # Calculate totals
        subtotal = 0
        taxable_amount = 0
        non_taxable_amount = 0
        total_cost = 0
        
        items = []
        for item in self.items_tree.get_children():
            values = self.items_tree.item(item)['values']
            description = values[0]
            quantity = float(values[1])
            unit_price = float(values[2])
            taxable = values[3] == "Yes"
            total = quantity * unit_price
            
            subtotal += total
            
            if taxable:
                taxable_amount += total
            else:
                non_taxable_amount += total
            
            items.append({
                'description': description,
                'quantity': quantity,
                'unit_price': unit_price,
                'total': total,
                'taxable': taxable
            })
        
        # Calculate costs
        costs = []
        original_costs = self.invoice.get('costs', [])
        for idx, cost in enumerate(self.costs_tree.get_children()):
            values = self.costs_tree.item(cost)['values']
            vlen = len(values)
            desc = values[0]
            amount = float(values[1])
            fund = values[2] if vlen > 2 else ""
            exp_gl = (values[3] if vlen > 3 else "") or "5000"
            receipt_cell = (values[4] if vlen > 4 else "") if vlen > 3 else (values[3] if vlen > 3 else "")
            receipt_attached = str(receipt_cell) in ('📎', 'yes', 'Yes', 'Y', 'y', 'true')
            # normalize exp_gl to 4-digit code
            exp_code = str(exp_gl).strip()
            if len(exp_code) > 4 and exp_code[:4].isdigit():
                exp_code = exp_code[:4]
            if not (len(exp_code) == 4 and exp_code.isdigit()):
                exp_code = "5000"
            # Preserve receipt_filename if available from original costs
            receipt_filename = None
            if idx < len(original_costs):
                receipt_filename = original_costs[idx].get('receipt_filename', None)
            cost_dict = {
                'description': str(desc),
                'amount': amount,
                'account': str(fund) if str(fund) not in ("", "None") else "Cash",
                'expense_account': exp_code,
                'receipt_attached': receipt_attached
            }
            if receipt_filename:
                cost_dict['receipt_filename'] = receipt_filename
            costs.append(cost_dict)
            total_cost += amount

        # Calculate VAT
        vat_amount = taxable_amount * (vat_rate / 100)
        grand_total = subtotal + vat_amount
        profit_loss = grand_total - total_cost

        # Update invoice dictionary
        # Derive location/emirate for save
        loc_value = (self.client_emirate.get() or "").strip()
        try:
            low = loc_value.lower()
            if ',' in loc_value and 'united arab emirates' in low:
                emirate_guess = loc_value.split(',', 1)[0].strip()
            else:
                emirate_guess = loc_value
        except Exception:
            emirate_guess = loc_value
        previous_invoice = dict(self.invoice)
        old_inventory_links = list(previous_invoice.get('inventory_links') or [])
        currency = str(self.currency_var.get() or "").strip() or str(self.invoice.get('currency', 'AED') or 'AED')
        self.invoice.update({
            'invoice_type': self.invoice_type.get(),
            'client_name': client_name,
            'client_trn': self.client_trn.get().strip(),
            'client_emirate': emirate_guess,
            'client_location': loc_value,
            'date': invoice_date,  # UPDATED: Use the edited date
            'payment_method': self.payment_method.get(),
            'payment_due': self.payment_due.get(),
            'due_date': self._compute_due_date(invoice_date, self.payment_due.get()),
            'tax_rate': vat_rate,
            'workflow_status': self.workflow_status.get(),
            'items': items,
            'costs': costs,
            'subtotal': subtotal,
            'taxable_amount': taxable_amount,
            'non_taxable_amount': non_taxable_amount,
            'tax_amount': vat_amount,
            'grand_total': grand_total,
            'total_cost': total_cost,
            'profit_loss': profit_loss,
            'notes': self.notes_text.get('1.0', tk.END).strip(),
            'currency': currency,
            'inventory_links': old_inventory_links
        })
        inventory_sync_message = ""
        inventory_manager = _get_inventory_manager_for_invoice(self.manager)
        if inventory_manager:
            try:
                inventory_manager.reverse_invoice_sales(self.invoice_id)
            except Exception:
                pass
            sync_ok, sync_msg, inventory_links = inventory_manager.sync_invoice_to_inventory(self.invoice)
            if not sync_ok:
                if old_inventory_links:
                    try:
                        inventory_manager.sync_invoice_to_inventory(previous_invoice)
                    except Exception:
                        pass
                override = messagebox.askyesno(
                    "Inventory Check",
                    f"{sync_msg}\n\n"
                    f"Save changes would normally stop here because inventory is insufficient.\n"
                    f"Proceed anyway?  The invoice will still save, but NO inventory will be deducted.\n\n"
                    f"Save without inventory deduction?",
                    icon="warning"
                )
                if not override:
                    return
                inventory_links = []
                inventory_sync_message = "Saved with override (no stock deduction): " + sync_msg
            else:
                self.invoice['inventory_links'] = inventory_links
                inventory_sync_message = sync_msg
        # Persist updates by writing the updated dictionary directly to storage
        success = False
        try:
            success = self.manager.add_invoice_from_dict(self.invoice)
        except Exception:
            try:
                # Fallback: update via object if available
                invoice_obj = self.manager.get_invoice(self.invoice_id)
                if invoice_obj:
                    for key, value in self.invoice.items():
                        try:
                            setattr(invoice_obj, key, value)
                        except Exception:
                            pass
                    if hasattr(self.manager, 'update_invoice'):
                        success = self.manager.update_invoice(invoice_obj)
            except Exception:
                success = False
        if not success and inventory_manager:
            try:
                inventory_manager.reverse_invoice_sales(self.invoice_id)
            except Exception:
                pass
            if old_inventory_links:
                try:
                    inventory_manager.sync_invoice_to_inventory(previous_invoice)
                except Exception:
                    pass
            self.invoice['inventory_links'] = old_inventory_links
        if success:
            success_message = f"Invoice {self.invoice_id} updated successfully!"
            if inventory_sync_message:
                success_message += f"\n\nInventory: {inventory_sync_message}"
            messagebox.showinfo("Success", success_message)
            self.dialog.destroy()
        else:
            messagebox.showerror("Error", "Failed to update invoice")

class AddPaymentDialog:
    def __init__(self, parent, manager, invoice_id, invoice):
        self.manager = manager
        self.invoice_id = invoice_id
        self.invoice = invoice
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Add Payment - {invoice_id}")
        self.dialog.geometry("700x560")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        _fit_window(self.dialog, 700, 560)
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill='both', expand=True)
        curr = str(self.invoice.get('currency', 'AED') or 'AED').strip() or 'AED'
        
        ttk.Label(main_frame, text=f"Add Payment to {self.invoice_id}", 
                 font=('Helvetica', 12, 'bold')).pack(pady=(0, 10))
        
        # Invoice info
        info_frame = ttk.LabelFrame(main_frame, text="Invoice Summary", padding="10")
        info_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(info_frame, text=f"Client: {self.invoice.get('client_name', '')}").pack(anchor='w')
        ttk.Label(info_frame, text=f"Total: {curr} {self.invoice.get('grand_total', 0):.2f}").pack(anchor='w')
        ttk.Label(info_frame, text=f"Cost: {curr} {self.invoice.get('total_cost', 0):.2f}").pack(anchor='w')
        ttk.Label(info_frame, text=f"P/L: {curr} {self.invoice.get('profit_loss', 0):.2f}").pack(anchor='w')
        ttk.Label(info_frame, text=f"Paid: {curr} {self.invoice.get('total_paid', 0):.2f}").pack(anchor='w')
        
        balance = self.invoice.get('grand_total', 0) - self.invoice.get('total_paid', 0)
        ttk.Label(info_frame, text=f"Balance Due: {curr} {balance:.2f}").pack(anchor='w')
        
        # Payment details
        payment_frame = ttk.LabelFrame(main_frame, text="Payment Details", padding="10")
        payment_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(payment_frame, text=f"Amount ({curr}):").grid(row=0, column=0, sticky='w', pady=5)
        self.payment_amount = ttk.Entry(payment_frame, width=15)
        self.payment_amount.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        self.payment_amount.insert(0, str(balance))
        
        ttk.Label(payment_frame, text="Date:").grid(row=1, column=0, sticky='w', pady=5)
        self.payment_date = ttk.Entry(payment_frame, width=15)
        self.payment_date.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        self.payment_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        ttk.Label(payment_frame, text="Payment Method:").grid(row=2, column=0, sticky='w', pady=5)
        self.payment_method = ttk.Combobox(payment_frame, 
                                         values=["Cash", "Transfer"],
                                         state="readonly", width=15)
        self.payment_method.set("Cash")
        self.payment_method.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        
        # Account selection - NEW
        ttk.Label(payment_frame, text="Account *:").grid(row=3, column=0, sticky='w', pady=5)
        self.payment_account = ttk.Combobox(payment_frame, width=15)
        self.payment_account.grid(row=3, column=1, sticky='w', pady=5, padx=10)
        self.load_accounts()
        
        ttk.Label(payment_frame, text="Notes:").grid(row=4, column=0, sticky='w', pady=5)
        self.payment_notes = ttk.Entry(payment_frame, width=15)
        self.payment_notes.grid(row=4, column=1, sticky='w', pady=5, padx=10)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        ttk.Button(button_frame, text="Add Payment", 
                  command=self.add_payment).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side='left', padx=5)

        payments_frame = ttk.LabelFrame(main_frame, text="Payments Done", padding="10")
        payments_frame.pack(fill='both', expand=True, pady=(10, 0))
        cols = ('Date','Amount','Method','Account','Notes')
        self.payments_tree = ttk.Treeview(payments_frame, columns=cols, show='headings', height=5)
        for c in cols:
            self.payments_tree.heading(c, text=c)
            self.payments_tree.column(c, width=120 if c!='Notes' else 220)
        self.payments_tree.pack(fill='both', expand=True)
        self.load_payments_into_tree()
        actions = ttk.Frame(payments_frame)
        actions.pack(pady=8)
        ttk.Button(actions, text="Edit Selected", command=self.edit_selected_payment).pack(side='left', padx=6)
        ttk.Button(actions, text="Delete Selected", command=self.delete_selected_payment).pack(side='left', padx=6)
    
    def load_accounts(self):
        """Load available accounts from balance manager"""
        try:
            balance_manager = BalanceManager(self.manager.invoice_folder)
            accounts = balance_manager.get_account_names()
            self.payment_account['values'] = accounts
            if accounts:
                self.payment_account.set(accounts[0])
        except Exception as e:
            print(f"Error loading accounts: {e}")
            self.payment_account['values'] = ["Cash", "Emirates NBD", "ADCB", "DrAhmed ADCB"]
            self.payment_account.set("Cash")

    def load_payments_into_tree(self):
        for i in self.payments_tree.get_children():
            self.payments_tree.delete(i)
        payments = self.invoice.get('payment_history', [])
        for p in payments:
            self.payments_tree.insert('', tk.END, values=(
                p.get('date',''),
                f"{float(p.get('amount',0)):.2f}",
                p.get('method',''),
                p.get('account', p.get('method','')),
                p.get('notes','')
            ))

    def _recalc_summary(self):
        try:
            inv_obj = self.manager.get_invoice(self.invoice_id)
            if inv_obj:
                d = inv_obj.to_dict()
                balance = d.get('grand_total',0) - d.get('total_paid',0)
                for child in self.dialog.winfo_children():
                    pass
        except Exception:
            pass

    def edit_selected_payment(self):
        sel = self.payments_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a payment to edit")
            return
        idx = self.payments_tree.index(sel[0])
        vals = self.payments_tree.item(sel[0])['values']
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Edit Payment")
        dlg.geometry("520x360")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Amount (AED):").grid(row=0, column=0, sticky='w', pady=5)
        amt = ttk.Entry(frm, width=15)
        amt.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        amt.insert(0, str(vals[1]))
        ttk.Label(frm, text="Date:").grid(row=1, column=0, sticky='w', pady=5)
        dt = ttk.Entry(frm, width=15)
        dt.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        dt.insert(0, vals[0])
        ttk.Label(frm, text="Method:").grid(row=2, column=0, sticky='w', pady=5)
        meth = ttk.Combobox(frm, values=["Cash","Transfer"], state='readonly', width=15)
        meth.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        meth.set(vals[2] or "Cash")
        ttk.Label(frm, text="Account:").grid(row=3, column=0, sticky='w', pady=5)
        acc = ttk.Combobox(frm, width=15)
        acc.grid(row=3, column=1, sticky='w', pady=5, padx=10)
        try:
            balance_manager = BalanceManager(self.manager.invoice_folder)
            acc['values'] = balance_manager.get_account_names()
        except Exception:
            acc['values'] = ["Cash"]
        acc.set(vals[3] or meth.get())
        ttk.Label(frm, text="Notes:").grid(row=4, column=0, sticky='w', pady=5)
        nts = ttk.Entry(frm, width=20)
        nts.grid(row=4, column=1, sticky='w', pady=5, padx=10)
        nts.insert(0, vals[4] or "")
        def save():
            try:
                new_amount = float(amt.get())
                if new_amount <= 0:
                    raise ValueError
                try:
                    datetime.strptime(dt.get(), "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter a valid date (YYYY-MM-DD)")
                    return
                old_amount = float(vals[1])
                old_account = vals[3] or vals[2] or 'Cash'
                if hasattr(self.manager, 'reverse_invoice_payment_with_balance'):
                    ok, msg = self.manager.reverse_invoice_payment_with_balance(self.invoice, old_amount, old_account)
                    if not ok:
                        messagebox.showerror("Error", msg)
                        return
                if hasattr(self.manager, 'process_invoice_payment_with_balance'):
                    ok, msg = self.manager.process_invoice_payment_with_balance(self.invoice, new_amount, acc.get())
                    if not ok:
                        messagebox.showerror("Error", msg)
                        return
                inv_obj = self.manager.get_invoice(self.invoice_id)
                if inv_obj and hasattr(inv_obj, 'edit_payment'):
                    ok, msg = inv_obj.edit_payment(idx, new_amount, dt.get(), meth.get(), nts.get(), acc.get())
                    if not ok:
                        messagebox.showerror("Error", msg)
                        return
                    if hasattr(self.manager, 'update_invoice'):
                        save_ok = self.manager.update_invoice(inv_obj)
                    else:
                        save_ok = self.manager.add_invoice_from_dict(inv_obj.to_dict())
                    if not save_ok:
                        messagebox.showerror("Error", "Failed to save changes")
                        return
                self.load_payments_into_tree()
                dlg.destroy()
                messagebox.showinfo("Success", "Payment updated successfully")
            except Exception:
                messagebox.showwarning("Warning", "Please enter a valid amount")
        ttk.Button(frm, text="Save", command=save).grid(row=5, column=0, pady=10)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=5, column=1, pady=10)
        _fit_window(dlg, 520, 360, mode="dialog", remember_key="edit_payment")

    def delete_selected_payment(self):
        sel = self.payments_tree.selection()
        if not sel:
            messagebox.showwarning("Warning", "Please select a payment to delete")
            return
        idx = self.payments_tree.index(sel[0])
        vals = self.payments_tree.item(sel[0])['values']
        try:
            old_amount = float(vals[1])
            old_account = vals[3] or vals[2] or 'Cash'
            if hasattr(self.manager, 'reverse_invoice_payment_with_balance'):
                ok, msg = self.manager.reverse_invoice_payment_with_balance(self.invoice, old_amount, old_account)
                if not ok:
                    messagebox.showerror("Error", msg)
                    return
            inv_obj = self.manager.get_invoice(self.invoice_id)
            if inv_obj and hasattr(inv_obj, 'delete_payment'):
                ok, msg = inv_obj.delete_payment(idx)
                if not ok:
                    messagebox.showerror("Error", msg)
                    return
                if hasattr(self.manager, 'update_invoice'):
                    save_ok = self.manager.update_invoice(inv_obj)
                else:
                    save_ok = self.manager.add_invoice_from_dict(inv_obj.to_dict())
                if not save_ok:
                    messagebox.showerror("Error", "Failed to save changes")
                    return
            self.load_payments_into_tree()
            messagebox.showinfo("Success", "Payment deleted successfully")
        except Exception:
            messagebox.showwarning("Warning", "Unexpected error while deleting payment")
    
    def add_payment(self):
        """Add payment to invoice with account update"""
        try:
            amount = float(self.payment_amount.get())
            payment_date = self.payment_date.get()
            payment_method = self.payment_method.get()
            payment_account = self.payment_account.get()
            payment_notes = self.payment_notes.get()
            
            if amount <= 0:
                messagebox.showwarning("Warning", "Please enter a valid payment amount")
                return
                
            if not payment_account:
                messagebox.showwarning("Warning", "Please select an account")
                return
                
            current_paid = self.invoice.get('total_paid', 0)
            total = self.invoice.get('grand_total', 0)
            
            if amount > (total - current_paid):
                messagebox.showwarning("Warning", "Payment amount cannot exceed balance due")
                return
            
            # Validate date
            try:
                datetime.strptime(payment_date, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Warning", "Please enter a valid date (YYYY-MM-DD)")
                return
            
            # Process payment in balance manager FIRST
            if hasattr(self.manager, 'process_invoice_payment_with_balance'):
                success, message = self.manager.process_invoice_payment_with_balance(
                    self.invoice, amount, payment_account
                )
                
                if not success:
                    messagebox.showerror("Error", f"Failed to update account balance: {message}")
                    return
            
            # Get the actual invoice object from manager
            invoice_obj = self.manager.get_invoice(self.invoice_id)
            if invoice_obj:
                print(f"Before payment - Total Paid: {invoice_obj.total_paid}, Status: {invoice_obj.status}")
                # Add the payment using the fixed method
                success = invoice_obj.add_payment(amount, payment_date, payment_method, payment_notes, payment_account)
                if success:
                    print(f"After payment - Total Paid: {invoice_obj.total_paid}, Status: {invoice_obj.status}")
                    
                    # Save the updated invoice
                    if hasattr(self.manager, 'update_invoice'):
                        # Use the new update_invoice method if available
                        save_success = self.manager.update_invoice(invoice_obj)
                    else:
                        # Fallback to using add_invoice_from_dict
                        save_success = self.manager.add_invoice_from_dict(invoice_obj.to_dict())
                    
                    if save_success:
                        # Also save invoices to ensure data persistence
                        self.manager.save_invoices()
                        # GAAP Ledger: Record payment received
                        if hasattr(self.manager, 'ledger') and self.manager.ledger:
                            try:
                                self.manager.ledger.on_payment_received(
                                    invoice_id=self.invoice_id,
                                    amount=amount,
                                    payment_date=payment_date,
                                    username="System",
                                    note=f"{payment_method} | {payment_notes}".strip()
                                )
                            except Exception as le:
                                print(f"Ledger hook warning (payment): {le}")
                        try:
                            folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                            path = os.path.join(folder, 'audit_log.jsonl')
                            rec = {
                                'timestamp': datetime.now().isoformat(),
                                'action': 'payment_added',
                                'details': {
                                    'invoice_id': self.invoice_id,
                                    'amount': amount,
                                    'account': payment_account,
                                    'method': payment_method
                                }
                            }
                            with open(path, 'a') as f:
                                f.write(json.dumps(rec) + "\n")
                        except Exception:
                            pass
                        
                        curr = str(getattr(invoice_obj, 'currency', None) or self.invoice.get('currency', 'AED') or 'AED').strip() or 'AED'
                        messagebox.showinfo(
                            "Success",
                            f"Payment of {curr} {amount:.2f} added successfully!\n"
                            f"Account: {payment_account}\n"
                            f"New Total Paid: {curr} {invoice_obj.total_paid:.2f}\n"
                            f"New Status: {invoice_obj.status}\n"
                            f"Balance Due: {curr} {invoice_obj.balance_due:.2f}"
                        )
                        self.dialog.destroy()
                    else:
                        messagebox.showerror("Error", "Payment added but failed to save changes to file")
                else:
                    messagebox.showerror("Error", "Failed to add payment to invoice")
            else:
                messagebox.showerror("Error", "Invoice not found in database")
            
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a valid payment amount")
        except Exception as e:
            messagebox.showerror("Error", f"Unexpected error: {str(e)}")
            print(f"Error in add_payment: {e}")

class BulkPaymentsDialog:
    """
    Add payments to MANY invoices in a single click.

    Features:
      • Loads every invoice with balance_due > 0 (user can filter / sort).
      • Multi-select checkboxes per row (Ctrl+A select all / Ctrl+click).
      • Each row has an editable PAYMENT AMOUNT — defaults to the full balance.
        User can overwrite with a partial payment.
      • Single shared Payment Date, Payment Method, Account, Notes — applied
        to every row in one go.
      • Live running totals: #Selected, Total Payments, Total Outstanding.
      • "Apply same fixed amount" + "Apply as % of balance" quick helpers.
      • Apply loop calls the EXACT same internal flow as the single-invoice
        Add Payment dialog (balance manager → invoice.add_payment → save →
        ledger.on_payment_received → audit_log).
      • Result dialog lists per-row successes + any errors (doesn't stop on
        a single failure).
    """

    _CURRENCY = "AED"

    def __init__(self, parent, manager):
        self.manager = manager
        self.dlg = tk.Toplevel(parent)
        self.dlg.title("💰 Bulk Add Payments — record one payment for many invoices")
        self.dlg.transient(parent)
        self.dlg.grab_set()
        self.dlg.resizable(True, True)
        _fit_window(self.dlg, 1200, 760, remember_key="bulk_payments")

        self.rows = []       # list[dict]: each = {invoice_id, inv_dict, selected (bool), amount, currency}
        self._build()
        self._reload_invoices()

    # ---------------------------------------------------------------
    # UI
    # ---------------------------------------------------------------
    def _build(self):
        root = ttk.Frame(self.dlg, padding=10)
        root.pack(fill="both", expand=True)

        # Header
        header = ttk.Frame(root)
        header.pack(fill="x", pady=(0, 8))
        ttk.Label(header, text="💰 Bulk Add Payments",
                  font=("Helvetica", 16, "bold")).pack(side="left")
        ttk.Label(header,
                  text=(
                      "Tick the invoices you want to pay, edit the Payment Amount per row "
                      "(defaults to full balance due), set a Date / Method / Account / Notes at the "
                      "bottom, then click ✅ Record All Payments.  Uses exactly the same ledger + "
                      "balance flow as the normal single-invoice Add Payment."
                  ),
                  foreground="#2D3748", wraplength=800, justify="left",
        ).pack(side="left", padx=20)

        # Filter toolbar
        ft = ttk.Frame(root); ft.pack(fill="x", pady=(0, 8))
        ttk.Label(ft, text="Search:").pack(side="left")
        self.sv = tk.StringVar()
        e = ttk.Entry(ft, textvariable=self.sv, width=32); e.pack(side="left", padx=4)
        e.bind("<KeyRelease>", lambda ev: self._render())
        ttk.Label(ft, text="  Filter:").pack(side="left", padx=(10, 4))
        self.filter_var = tk.StringVar(value="With Balance Due only")
        cb = ttk.Combobox(ft, values=["All invoices", "With Balance Due only", "Not Paid (Partial/None)"],
                          state="readonly", width=26, textvariable=self.filter_var)
        cb.pack(side="left")
        cb.bind("<<ComboboxSelected>>", lambda ev: self._render())
        ttk.Button(ft, text="🔄 Reload invoices",
                   command=self._reload_invoices).pack(side="left", padx=10)

        # Quick-select
        qs = ttk.Frame(ft); qs.pack(side="left", padx=14)
        ttk.Button(qs, text="☑ All Visible", command=self._sel_all_visible).pack(side="left")
        ttk.Button(qs, text="☐ None", command=self._sel_none).pack(side="left", padx=4)
        ttk.Button(qs, text="☑ Invert", command=self._sel_invert).pack(side="left")

        # Quick-apply payment amounts
        qa = ttk.LabelFrame(root, text="Quick-apply payment amounts for TICKED rows", padding=8)
        qa.pack(fill="x", pady=(0, 8))
        ttk.Button(qa, text="Reset ticked → full balance due",
                   command=self._qa_reset_ticked_full).pack(side="left")
        ttk.Label(qa, text="  Fixed amount per ticked invoice:").pack(side="left", padx=(14, 4))
        self.qa_fixed = ttk.Entry(qa, width=14); self.qa_fixed.pack(side="left")
        ttk.Button(qa, text="Apply Fixed", command=self._qa_apply_fixed).pack(side="left", padx=4)
        ttk.Label(qa, text="  % of balance per ticked invoice:").pack(side="left", padx=(14, 4))
        self.qa_pct = ttk.Entry(qa, width=10); self.qa_pct.pack(side="left")
        ttk.Label(qa, text="%").pack(side="left")
        ttk.Button(qa, text="Apply %", command=self._qa_apply_pct).pack(side="left", padx=4)

        # Running totals bar
        self.totals_var = tk.StringVar(value="")
        ttk.Label(root, textvariable=self.totals_var,
                  font=("Helvetica", 12, "bold"), foreground="#2B6CB0",
                  anchor="w").pack(fill="x", pady=(0, 8))

        # Main Tree
        cols = ("pick", "invoice_id", "date", "client", "grand_total",
                "total_paid", "balance_due", "payment_amount", "status")
        self.tv = ttk.Treeview(root, columns=cols, show="headings", height=22)
        head_w = [40, 150, 100, 300, 120, 120, 140, 160, 150]
        head_t = ["☑", "Invoice ID", "Date", "Client", f"Total ({self._CURRENCY})",
                  f"Paid ({self._CURRENCY})", f"Balance Due ({self._CURRENCY})",
                  f"Payment Amount (edit) ✎", "Status"]
        for c, h, w in zip(cols, head_t, head_w):
            self.tv.heading(c, text=h)
            anchor = "center" if c in ("pick", "status") else "e" if "amount" in c or "total" in c or "paid" in c or "balance" in c else "w"
            self.tv.column(c, width=w, anchor=anchor)
        self.tv.tag_configure("selected_bg", background="#EBF8FF")
        self.tv.tag_configure("fullypaid",    foreground="#A0AEC0")
        self.tv.tag_configure("partial",      foreground="#2C5282")
        self.tv.tag_configure("unpaid",       foreground="#C53030")
        self.tv.bind("<Double-1>", self._on_dblclick)
        self.tv.bind("<Button-1>", self._on_single_click_toggle)
        self.tv.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(root, orient="vertical", command=self.tv.yview)
        self.tv.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        # Shared payment fields
        pf = ttk.LabelFrame(root, text="Shared payment details — applied to EVERY ticked invoice", padding=10)
        pf.pack(fill="x", pady=(10, 0))
        ttk.Label(pf, text="Date:").grid(row=0, column=0, sticky="w", pady=4)
        self.p_date = ttk.Entry(pf, width=18); self.p_date.grid(row=0, column=1, sticky="w", pady=4, padx=(4, 20))
        self.p_date.insert(0, datetime.now().strftime("%Y-%m-%d"))

        ttk.Label(pf, text="Method:").grid(row=0, column=2, sticky="w", pady=4)
        self.p_method = ttk.Combobox(pf, values=["Cash", "Transfer", "Card", "Check", "Other"],
                                     state="readonly", width=14)
        self.p_method.set("Cash"); self.p_method.grid(row=0, column=3, sticky="w", pady=4, padx=(4, 20))

        ttk.Label(pf, text="Account *:").grid(row=0, column=4, sticky="w", pady=4)
        self.p_account = ttk.Combobox(pf, width=22); self.p_account.grid(row=0, column=5, sticky="w", pady=4, padx=(4, 20))
        self._load_accounts()

        ttk.Label(pf, text="Notes:").grid(row=0, column=6, sticky="w", pady=4)
        self.p_notes = ttk.Entry(pf, width=32); self.p_notes.grid(row=0, column=7, sticky="w", pady=4, padx=(4, 0))

        ttk.Label(pf, text="⚠️  Each individual payment is still checked against that invoice's balance — no overpay.",
                  foreground="#B7791F").grid(row=1, column=0, columnspan=8, sticky="w", pady=(6, 0))

        # Bottom buttons
        btns = ttk.Frame(root); btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="✅  Record All Payments for Ticked Invoices",
                   command=self._on_apply, style="Primary.TButton").pack(side="left")
        ttk.Label(btns, text="   (after record, the invoice manager list is auto-refreshed)").pack(side="left", padx=10)
        ttk.Button(btns, text="Close", command=self.dlg.destroy).pack(side="right")

    # ---------------------------------------------------------------
    # Data
    # ---------------------------------------------------------------
    def _load_accounts(self):
        try:
            bal = BalanceManager(getattr(self.manager, "invoice_folder", None))
            names = bal.get_account_names()
        except Exception:
            names = ["ADCB", "Cash"]
        if not names:
            names = ["ADCB", "Cash"]
        self.p_account["values"] = names
        try:
            if "ADCB" in names:
                self.p_account.set("ADCB")
            else:
                self.p_account.set(names[0])
        except Exception:
            pass

    def _all_invoice_dicts(self):
        """Return list[invoice_dict] using manager or JSON fallback."""
        # Prefer manager invoices_data load to stay consistent with source of truth
        try:
            if hasattr(self.manager, "load_json"):
                raw = self.manager.load_json("invoices_data.json") or []
                if isinstance(raw, list):
                    return [x for x in raw if isinstance(x, dict) and "invoice_id" in x]
        except Exception:
            pass
        try:
            if hasattr(self.manager, "get_invoice_ids"):
                out = []
                for iid in self.manager.get_invoice_ids():
                    obj = self.manager.get_invoice(iid)
                    if obj is None: continue
                    d = obj.to_dict() if hasattr(obj, "to_dict") else dict(obj)
                    out.append(d)
                return out
        except Exception:
            pass
        return []

    def _reload_invoices(self):
        # Build internal rows indexed by invoice_id for stability across filters
        invs = self._all_invoice_dicts()
        by_id = {r["invoice_id"]: r for r in self.rows}
        new_rows = []
        for d in invs:
            iid = d["invoice_id"]
            balance = float(d.get("balance_due") or 0)
            if abs(balance - (float(d.get("grand_total") or 0) - float(d.get("total_paid") or 0))) > 0.005:
                balance = float(d.get("grand_total") or 0) - float(d.get("total_paid") or 0)
            cur = str(d.get("currency") or "AED").strip() or "AED"
            if iid in by_id:
                existing = by_id[iid]
                existing["inv_dict"] = d
                existing["balance"] = balance
                # Clamp payment amount to new balance if stale
                try:
                    amt = float(existing.get("amount") or 0)
                except Exception:
                    amt = 0.0
                if amt > balance or abs(amt - 0) < 0.0001:
                    amt = balance
                existing["amount"] = amt
                existing["currency"] = cur
                new_rows.append(existing)
            else:
                new_rows.append(dict(
                    invoice_id=iid,
                    inv_dict=d,
                    selected=False,
                    amount=balance,
                    balance=balance,
                    currency=cur,
                ))
        self.rows = new_rows
        self._render()

    def _filtered_rows(self):
        q = self.sv.get().strip().lower()
        mode = self.filter_var.get()
        out = []
        for idx, r in enumerate(self.rows):
            bal = r.get("balance") or 0.0
            if mode == "With Balance Due only" and bal <= 0.001:
                continue
            if mode == "Not Paid (Partial/None)" and bal <= 0.001:
                continue
            if q:
                d = r["inv_dict"]
                hay = " ".join(str(x or "").lower() for x in [
                    r["invoice_id"], d.get("client_name"), d.get("status"),
                    d.get("date"), d.get("due_date"), d.get("grand_total"),
                ])
                if q not in hay:
                    continue
            out.append((idx, r))
        return out

    def _render(self):
        for i in self.tv.get_children():
            self.tv.delete(i)
        sel_count = 0
        sel_total = 0.0
        bal_total = 0.0
        for idx, r in self._filtered_rows():
            d = r["inv_dict"]
            bal = float(r.get("balance") or 0)
            paid = float(d.get("total_paid") or 0)
            total = float(d.get("grand_total") or 0)
            status = d.get("status") or ""
            amt = float(r.get("amount") or 0)
            tags = []
            if r.get("selected"):
                tags.append("selected_bg")
                sel_count += 1
                sel_total += amt
            bal_total += max(bal, 0.0)
            if bal <= 0.001:
                tags.append("fullypaid")
            elif paid <= 0.001:
                tags.append("unpaid")
            else:
                tags.append("partial")
            self.tv.insert(
                "", "end", iid=str(idx), tags=tuple(tags),
                values=(
                    "☑" if r.get("selected") else "☐",
                    r["invoice_id"],
                    d.get("date", ""),
                    d.get("client_name", ""),
                    f"{total:,.2f}",
                    f"{paid:,.2f}",
                    f"{bal:,.2f}",
                    "" if amt <= 0 else f"{amt:,.2f}",
                    status,
                ),
            )
        self.totals_var.set(
            f"☑ {sel_count} invoice(s) selected  ·  "
            f"Payments total: {self._CURRENCY} {sel_total:,.2f}   ·   "
            f"Total balance due (visible): {self._CURRENCY} {bal_total:,.2f}"
        )

    # ---------------------------------------------------------------
    # Pick + row interactions
    # ---------------------------------------------------------------
    def _toggle_by_idx(self, idx):
        if idx < 0 or idx >= len(self.rows): return
        self.rows[idx]["selected"] = not bool(self.rows[idx].get("selected"))
        self._render()

    def _on_single_click_toggle(self, event):
        col = self.tv.identify_column(event.x)
        iid = self.tv.identify_row(event.y)
        if not iid: return
        if col and str(col).endswith("1"):
            # first column = pick
            self._toggle_by_idx(int(iid))
            return "break"

    def _on_dblclick(self, event):
        iid = self.tv.identify_row(event.y)
        if not iid: return
        col = self.tv.identify_column(event.x)
        if not col: return
        try:
            col_idx = int(str(col).replace("#", "")) - 1
        except Exception:
            return
        if col_idx == 0:
            self._toggle_by_idx(int(iid)); return
        if col_idx != 7:
            return  # only Payment Amount is editable via dbl-click (or via quick apply)
        try:
            idx = int(iid)
            r = self.rows[idx]
        except Exception:
            return
        # Inline editor for payment amount
        dlg = tk.Toplevel(self.dlg); dlg.title("Edit Payment Amount"); dlg.transient(self.dlg); dlg.grab_set(); dlg.geometry("420x200")
        frm = ttk.Frame(dlg, padding=12); frm.pack(fill="both", expand=True)
        ttk.Label(frm,
                  text=f"Invoice: {r['invoice_id']}\nClient: {r['inv_dict'].get('client_name','')}\n"
                       f"Balance due: {self._CURRENCY} {float(r.get('balance') or 0):,.2f}").grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(frm, text="Payment amount:").grid(row=1, column=0, sticky="w", pady=10)
        v = tk.StringVar(value="" if (r.get("amount") or 0) <= 0 else f"{float(r.get('amount') or 0):.2f}")
        e = ttk.Entry(frm, textvariable=v, width=20); e.grid(row=1, column=1, pady=10, sticky="w")
        def ok():
            raw = e.get().strip()
            try:
                val = float(raw) if raw else 0.0
                if val < 0: raise ValueError
            except Exception:
                messagebox.showwarning("Invalid", "Enter a positive number (or blank for 0)"); return
            bal = float(r.get("balance") or 0)
            if val > bal + 0.001:
                if not messagebox.askyesno("Exceeds Balance?",
                                           f"Amount {val:,.2f} > balance {bal:,.2f}.  Clamp to balance?\n"
                                           f"(No individual invoice can ever be overpaid.)",
                                           parent=dlg):
                    return
                val = bal
            # Take mini-snapshot internally by cloning rows shallowly isn't critical here.
            r["amount"] = float(val)
            # Auto-tick this row when user edits it non-zero
            if val > 0:
                r["selected"] = True
            self._render(); dlg.destroy()
        ttk.Button(frm, text="Save", command=ok).grid(row=2, column=0, pady=12)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=2, column=1, pady=12)

    # ---------------------------------------------------------------
    # Quick selects + quick apply
    # ---------------------------------------------------------------
    def _sel_all_visible(self):
        for idx, r in self._filtered_rows():
            if (r.get("balance") or 0) > 0.001:
                r["selected"] = True
        self._render()

    def _sel_none(self):
        for _, r in self._filtered_rows(): r["selected"] = False
        self._render()

    def _sel_invert(self):
        for _, r in self._filtered_rows(): r["selected"] = not r.get("selected")
        self._render()

    def _qa_reset_ticked_full(self):
        n = 0
        for r in self.rows:
            if r.get("selected"):
                r["amount"] = float(r.get("balance") or 0)
                n += 1
        self._render()
        if n: messagebox.showinfo("Reset", f"Reset {n} ticked invoice(s) → full balance due.", parent=self.dlg)

    def _qa_apply_fixed(self):
        raw = self.qa_fixed.get().strip()
        try:
            v = float(raw)
            if v < 0: raise ValueError
        except Exception:
            messagebox.showwarning("Invalid", "Enter a positive fixed amount"); return
        n = 0
        for r in self.rows:
            if r.get("selected"):
                bal = float(r.get("balance") or 0)
                r["amount"] = max(0.0, min(v, bal))
                n += 1
        self._render()
        messagebox.showinfo("Applied", f"Fixed amount {v:,.2f} clamped to balance → {n} ticked rows.", parent=self.dlg)

    def _qa_apply_pct(self):
        raw = self.qa_pct.get().strip()
        try:
            pct = float(raw)
            if pct < 0 or pct > 100: raise ValueError
        except Exception:
            messagebox.showwarning("Invalid", "Enter 0-100 percent"); return
        n = 0
        for r in self.rows:
            if r.get("selected"):
                bal = float(r.get("balance") or 0)
                r["amount"] = round(bal * (pct / 100.0), 2)
                n += 1
        self._render()
        messagebox.showinfo("Applied", f"{pct}% of balance → {n} ticked rows.", parent=self.dlg)

    # ---------------------------------------------------------------
    # Apply payments (single-shot)
    # ---------------------------------------------------------------
    def _apply_one(self, iid, inv_dict, amount, payment_date, payment_method, payment_account, payment_notes):
        """Replicate the exact single-payment flow from AddPaymentDialog.add_payment().

        Returns (ok: bool, message: str).
        """
        try:
            amount = float(amount)
        except Exception:
            return False, "Invalid amount"
        if amount <= 0.0001:
            return False, "Amount is zero — skipped"
        # 1) balance manager FIRST
        if hasattr(self.manager, 'process_invoice_payment_with_balance'):
            ok_bal, msg_bal = self.manager.process_invoice_payment_with_balance(
                inv_dict, amount, payment_account
            )
            if not ok_bal:
                return False, f"BalanceManager: {msg_bal}"
        # 2) load invoice object
        invoice_obj = self.manager.get_invoice(iid)
        if invoice_obj is None:
            return False, "Invoice object not found in manager"
        # 3) add_payment (auto-clamps to balance via invoice method)
        ok_inv, msg_inv = invoice_obj.add_payment(
            amount, payment_date, payment_method, payment_notes, payment_account
        )
        if not ok_inv:
            # Roll back balance-manager change if possible
            if hasattr(self.manager, 'reverse_invoice_payment_with_balance'):
                try:
                    self.manager.reverse_invoice_payment_with_balance(
                        inv_dict, amount, payment_account
                    )
                except Exception:
                    pass
            return False, f"invoice.add_payment: {msg_inv}"
        # 4) Save invoice
        try:
            if hasattr(self.manager, 'update_invoice'):
                saved = self.manager.update_invoice(invoice_obj)
            else:
                saved = self.manager.add_invoice_from_dict(invoice_obj.to_dict())
        except Exception as se:
            return False, f"save invoice: {se}"
        if not saved:
            if hasattr(self.manager, 'reverse_invoice_payment_with_balance'):
                try:
                    self.manager.reverse_invoice_payment_with_balance(
                        inv_dict, amount, payment_account
                    )
                except Exception:
                    pass
            return False, "Failed to persist invoice after payment"
        # 5) save_invoices + ledger
        try:
            if hasattr(self.manager, 'save_invoices'):
                self.manager.save_invoices()
        except Exception:
            pass
        if hasattr(self.manager, 'ledger') and self.manager.ledger:
            try:
                self.manager.ledger.on_payment_received(
                    invoice_id=iid,
                    amount=amount,
                    payment_date=payment_date,
                    username="System",
                    note=f"{payment_method} | {payment_notes}".strip(),
                )
            except Exception:
                pass
        # 6) audit log
        try:
            folder = (getattr(self.manager, 'invoice_folder', None)
                      or os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
            path = os.path.join(folder, 'audit_log.jsonl')
            rec = {
                'timestamp': datetime.now().isoformat(),
                'action': 'payment_added_bulk',
                'details': {
                    'invoice_id': iid, 'amount': amount,
                    'account': payment_account, 'method': payment_method,
                    'date': payment_date,
                },
            }
            with open(path, 'a') as f:
                f.write(json.dumps(rec) + "\n")
        except Exception:
            pass
        return True, f"{self._CURRENCY} {amount:,.2f} applied; new status={invoice_obj.status}; balance={float(invoice_obj.balance_due):,.2f}"

    def _on_apply(self):
        # Gather ticked rows
        targets = [(idx, r) for idx, r in enumerate(self.rows)
                   if r.get("selected") and float(r.get("amount") or 0) > 0.0001]
        if not targets:
            messagebox.showwarning("Nothing to do", "Tick one or more invoices and ensure they have a payment amount."); return
        # Validate shared fields
        payment_date = self.p_date.get().strip()
        try:
            datetime.strptime(payment_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid Date", "Date must be YYYY-MM-DD"); return
        payment_method = (self.p_method.get() or "").strip() or "Cash"
        payment_account = (self.p_account.get() or "").strip()
        if not payment_account:
            messagebox.showwarning("Missing Account", "Please choose an Account (ADCB / Cash / ...)"); return
        payment_notes = (self.p_notes.get() or "").strip()
        # Summary confirmation BEFORE posting:
        total_amt = sum(float(r.get("amount") or 0) for _, r in targets)
        if not messagebox.askyesno(
            "Confirm Bulk Payments",
            f"About to record {len(targets)} payment(s), TOTAL = {self._CURRENCY} {total_amt:,.2f}\n\n"
            f"Date: {payment_date}\n"
            f"Method: {payment_method}\n"
            f"Account: {payment_account}\n"
            f"Notes: {payment_notes or '(none)'}\n\n"
            f"Each payment is still validated against its invoice balance — no overpayments.  Continue?",
            icon="question",
            parent=self.dlg,
        ):
            return
        # Apply sequentially, collect results
        results = []
        applied = 0
        errors = []
        for idx, r in targets:
            iid = r["invoice_id"]
            amt = float(r.get("amount") or 0)
            ok, msg = self._apply_one(
                iid, r["inv_dict"], amt,
                payment_date, payment_method, payment_account, payment_notes,
            )
            results.append((iid, ok, msg, amt))
            if ok:
                applied += 1
                # Mark row with new balance / updated dict for immediate visual
                try:
                    new_obj = self.manager.get_invoice(iid)
                    if new_obj is not None:
                        r["inv_dict"] = new_obj.to_dict()
                        r["balance"] = float(getattr(new_obj, "balance_due",
                                                     float(r["inv_dict"].get("balance_due") or 0)))
                        r["amount"] = 0.0
                        r["selected"] = False
                except Exception:
                    pass
            else:
                errors.append((iid, msg, amt))
        # Reload to reflect new state
        self._reload_invoices()
        # Result dialog
        res = tk.Toplevel(self.dlg); res.title("Bulk Payments — Result"); res.transient(self.dlg); res.grab_set()
        _fit_window(res, 760, 520)
        fr = ttk.Frame(res, padding=12); fr.pack(fill="both", expand=True)
        ttk.Label(fr,
                  text=f"Applied {applied} / {len(targets)} payment(s)  ·  "
                       f"Total: {self._CURRENCY} {total_amt:,.2f}  ·  Errors: {len(errors)}",
                  font=("Helvetica", 14, "bold"), foreground=("#2F855A" if not errors else "#C53030"),
                  ).pack(anchor="w", pady=(0, 8))
        cols = ("invoice_id", "amount", "ok", "message")
        tv = ttk.Treeview(fr, columns=cols, show="headings", height=18)
        for c, h, w in zip(cols, ("Invoice ID", "Amount", "OK?", "Detail"), (140, 130, 60, 420)):
            tv.heading(c, text=h); tv.column(c, width=w, anchor="w" if c in ("invoice_id", "message") else "e" if c == "amount" else "center")
        tv.tag_configure("ok", foreground="#2F855A")
        tv.tag_configure("bad", foreground="#C53030")
        for iid, ok, msg, amt in results:
            tv.insert("", "end", tags=("ok",) if ok else ("bad",),
                      values=(iid, f"{self._CURRENCY} {amt:,.2f}", "✅" if ok else "❌", msg))
        tv.pack(fill="both", expand=True, pady=(8, 8))
        ttk.Button(fr, text="Close", command=res.destroy).pack()


class ViewInvoiceDialog:
    def __init__(self, parent, manager, invoice_id, invoice):
        self.manager = manager  # Add manager attribute
        self.invoice_id = invoice_id
        self.invoice = invoice
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Invoice {invoice_id}")
        self.dialog.geometry("900x750")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        # Header with View Receipts button
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(header_frame, text=f"INVOICE: {self.invoice_id}", 
                 font=('Helvetica', 16, 'bold')).pack(side='left')
        right_info = f"Date: {self.invoice.get('date','')}  |  ID: {self.invoice_id}"
        ttk.Label(header_frame, text=right_info, font=('Helvetica', 11)).pack(side='right')
        
        # View Receipts button
        ttk.Button(header_frame, text="📁 View Receipts", 
                  command=self.view_receipts).pack(side='right')
        
        status_color = 'green' if self.invoice.get('status') == 'Paid' else 'orange' if self.invoice.get('status') == 'Partial Paid' else 'red'
        status_label = ttk.Label(main_frame, text=f"Status: {self.invoice.get('status', 'Unknown')}", 
                                foreground=status_color, font=('Helvetica', 12, 'bold'))
        status_label.pack(anchor='w', pady=(0, 20))
        
        # Client and Invoice info
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill='x', pady=(0, 20))
        
        # Client info
        client_frame = ttk.LabelFrame(info_frame, text="Client Information", padding="10")
        client_frame.pack(side='left', fill='x', expand=True, padx=(0, 10))
        
        ttk.Label(client_frame, text=f"Name: {self.invoice.get('client_name', '')}").pack(anchor='w', pady=2)
        ttk.Label(client_frame, text=f"TRN: {self.invoice.get('client_trn', 'N/A')}").pack(anchor='w', pady=2)
        loc = (self.invoice.get('client_location') or '').strip()
        if not loc:
            loc = str(self.invoice.get('client_emirate', '') or '').strip()
        if not loc:
            loc = 'Dubai, United Arab Emirates'
        uae_emirates = {
            'Dubai',
            'Abu Dhabi',
            'Sharjah',
            'Ajman',
            'Umm Al Quwain',
            'Ras Al Khaimah',
            'Fujairah',
        }
        if ',' not in loc and loc in uae_emirates:
            loc = f"{loc}, United Arab Emirates"
        ttk.Label(client_frame, text=f"Location: {loc}").pack(anchor='w', pady=2)
        
        # Invoice info
        invoice_frame = ttk.LabelFrame(info_frame, text="Invoice Details", padding="10")
        invoice_frame.pack(side='left', fill='x', expand=True)
        ttk.Label(invoice_frame, text=f"Due Date: {self.invoice.get('due_date', '')}").pack(anchor='w', pady=2)
        ttk.Label(invoice_frame, text=f"Payment Method: {self.invoice.get('payment_method', 'N/A')}\nPayment Due: {self.invoice.get('payment_due', 'N/A')}").pack(anchor='w', pady=2)
        ttk.Label(invoice_frame, text=f"Type: {self.invoice.get('invoice_type', 'Sales').title()}").pack(anchor='w', pady=2)
        
        # Items
        items_frame = ttk.LabelFrame(main_frame, text="Items", padding="10")
        items_frame.pack(fill='both', expand=True, pady=(0, 20))
        
        currency = str(self.invoice.get('currency', 'AED') or 'AED').strip() or 'AED'

        # Create treeview
        columns = ('Description', 'Quantity', 'Unit Price', 'VAT', 'Total')
        tree = ttk.Treeview(items_frame, columns=columns, show='headings', height=6)
        
        for col in columns:
            if col == 'Unit Price':
                tree.heading(col, text=f"Unit Price ({currency})")
            elif col == 'Total':
                tree.heading(col, text=f"Total ({currency})")
            else:
                tree.heading(col, text=col)
            tree.column(col, width=120)
        
        # Add items
        for item in self.invoice.get('items', []):
            vat_text = "Yes" if item.get('taxable', True) else "No"
            tree.insert('', tk.END, values=(
                item.get('description', ''),
                item.get('quantity', 0),
                f"{item.get('unit_price', 0):.2f}",
                vat_text,
                f"{item.get('total', 0):.2f}"
            ))
        
        scrollbar = ttk.Scrollbar(items_frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Costs section
        if self.invoice.get('costs'):
            costs_frame = ttk.LabelFrame(main_frame, text="Invoice Costs (Internal)", padding="10")
            costs_frame.pack(fill='x', pady=(0, 20))
            
            cost_columns = ('Description', 'Amount', 'Receipt')
            self.costs_tree = ttk.Treeview(costs_frame, columns=cost_columns, show='headings', height=3)
            
            for col in cost_columns:
                if col == 'Amount':
                    self.costs_tree.heading(col, text=f"Amount ({currency})")
                else:
                    self.costs_tree.heading(col, text=col)
                self.costs_tree.column(col, width=150)
            
            # Add costs
            for i, cost in enumerate(self.invoice.get('costs', [])):
                receipt_indicator = "📎" if cost.get('receipt_attached') else ""
                self.costs_tree.insert('', tk.END, values=(
                    cost.get('description', ''),
                    f"{currency} {cost.get('amount', 0):.2f}",
                    receipt_indicator
                ))
            
            self.costs_tree.pack(fill='x')
            
            # Bind double-click to open receipt
            self.costs_tree.bind('<Double-1>', self.on_cost_double_click)
        
        # Totals
        totals_frame = ttk.Frame(main_frame)
        totals_frame.pack(fill='x', pady=(0, 20))
        
        ttk.Label(totals_frame, text=f"Subtotal: {currency} {self.invoice.get('subtotal', 0):.2f}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='e', pady=2)
        ttk.Label(totals_frame, text=f"Taxable Amount: {currency} {self.invoice.get('taxable_amount', 0):.2f}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='e', pady=2)
        ttk.Label(totals_frame, text=f"Non-Taxable Amount: {currency} {self.invoice.get('non_taxable_amount', 0):.2f}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='e', pady=2)
        ttk.Label(totals_frame, text=f"VAT ({self.invoice.get('tax_rate', 0)}%): {currency} {self.invoice.get('tax_amount', 0):.2f}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='e', pady=2)
        ttk.Label(totals_frame, text=f"GRAND TOTAL: {currency} {self.invoice.get('grand_total', 0):.2f}", 
                 font=('Helvetica', 12, 'bold')).pack(anchor='e', pady=2)
        ttk.Label(totals_frame, text=f"Total Paid: {currency} {self.invoice.get('total_paid', 0):.2f}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='e', pady=2)
        
        # Internal costs and P/L (not shown to clients)
        ttk.Label(totals_frame, text=f"Total Cost: {currency} {self.invoice.get('total_cost', 0):.2f}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='e', pady=2)
        
        profit_loss = self.invoice.get('profit_loss', 0)
        pl_color = 'green' if profit_loss >= 0 else 'red'
        ttk.Label(totals_frame, text=f"PROFIT/LOSS: {currency} {profit_loss:.2f}", 
                 font=('Helvetica', 12, 'bold'), foreground=pl_color).pack(anchor='e', pady=2)
        
        ttk.Label(totals_frame, text=f"BALANCE DUE: {currency} {self.invoice.get('grand_total', 0) - self.invoice.get('total_paid', 0):.2f}", 
                 font=('Helvetica', 12, 'bold')).pack(anchor='e', pady=2)

        # Notes section
        inv_notes = (self.invoice.get('notes') or '').strip()
        if inv_notes:
            notes_frame = ttk.LabelFrame(main_frame, text="Notes", padding="10")
            notes_frame.pack(fill='x', pady=(0, 10))
            notes_text = scrolledtext.ScrolledText(notes_frame, height=4)
            notes_text.pack(fill='x')
            notes_text.insert('1.0', inv_notes)
            notes_text.config(state='disabled')
        
        # Close button
        ttk.Button(main_frame, text="Close", 
                  command=self.dialog.destroy).pack(pady=10)
    
    def on_cost_double_click(self, event):
        """Handle double-click on cost item to open receipt"""
        selection = self.costs_tree.selection()
        if not selection:
            return
        
        cost_index = self.costs_tree.index(selection[0])
        cost = self.invoice.get('costs', [])[cost_index]
        
        if cost.get('receipt_attached'):
            # Open the receipt file using the manager
            success = self.manager.open_receipt(self.invoice_id, cost_index)
            if not success:
                messagebox.showerror("Error", "Could not open receipt file. The file may have been moved or deleted.")
        else:
            messagebox.showinfo("Info", "No receipt attached to this cost")
    
    def view_receipts(self):
        """Show all receipts for this invoice"""
        if not self.invoice.get('costs'):
            messagebox.showinfo("Info", "No costs found in this invoice")
            return
        
        receipts_found = False
        for i, cost in enumerate(self.invoice.get('costs', [])):
            if cost.get('receipt_attached'):
                receipts_found = True
                # Open the receipt file using the manager
                success = self.manager.open_receipt(self.invoice_id, i)
                if not success:
                    messagebox.showerror("Error", f"Could not open receipt for cost: {cost.get('description', 'Unknown')}")
        
        if not receipts_found:
            messagebox.showinfo("Info", "No receipts found for this invoice")

class ChangeInvoiceIdDialog:
    def __init__(self, parent, manager, invoice_id, invoice):
        self.manager = manager
        self.old_invoice_id = invoice_id
        self.invoice = invoice
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Change Invoice ID")
        self.dialog.geometry("400x200")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        _fit_window(self.dialog, 400, 200, mode="compact", remember_key="change_invoice_id")
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(main_frame, text="Change Invoice ID", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        ttk.Label(main_frame, text=f"Current Invoice ID: {self.old_invoice_id}").pack(anchor='w', pady=5)
        ttk.Label(main_frame, text=f"Client: {self.invoice.get('client_name', '')}").pack(anchor='w', pady=5)
        
        ttk.Label(main_frame, text="New Invoice ID:").pack(anchor='w', pady=(10, 5))
        self.new_invoice_id = ttk.Entry(main_frame, width=30)
        self.new_invoice_id.pack(fill='x', pady=5)
        self.new_invoice_id.insert(0, self.old_invoice_id)
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Update ID", 
                  command=self.update_invoice_id).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side='left', padx=5)
    
    def update_invoice_id(self):
        new_id = self.new_invoice_id.get().strip()
        
        if new_id == self.old_invoice_id:
            messagebox.showinfo("Info", "Invoice ID unchanged")
            self.dialog.destroy()
            return
        
        success, message = self.manager.update_invoice_id(self.old_invoice_id, new_id)
        
        if success:
            messagebox.showinfo("Success", message)
            self.dialog.destroy()
        else:
            messagebox.showerror("Error", message)

class CompanyReportDialog:
    def __init__(self, parent, manager):
        self.manager = manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Company Financial Report")
        self.dialog.geometry("1000x750")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        try:
            lm = LogoManager()
            if lm.load_logo(size=(64, 64)):
                self.dialog.iconphoto(True, lm.logo_photo)
        except Exception:
            pass
        
        self.setup_ui()
        _fit_window(self.dialog, 1000, 750, mode="workspace", remember_key="company_report")
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="🏢 Company Financial Statements", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        
        # Date range
        date_frame = ttk.LabelFrame(main_frame, text="Report Period", padding="10")
        date_frame.pack(fill='x', pady=5)
        
        ttk.Label(date_frame, text="Start Date (YYYY-MM-DD):").grid(row=0, column=0, padx=5, pady=5)
        self.start_date = ttk.Entry(date_frame, width=12)
        self.start_date.grid(row=0, column=1, padx=5, pady=5)
        self.start_date.insert(0, (datetime.now().replace(day=1) - timedelta(days=365)).strftime("%Y-%m-%d"))
        
        ttk.Label(date_frame, text="End Date (YYYY-MM-DD):").grid(row=0, column=2, padx=5, pady=5)
        self.end_date = ttk.Entry(date_frame, width=12)
        self.end_date.grid(row=0, column=3, padx=5)
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # Buttons
        button_frame = ttk.Frame(date_frame)
        button_frame.grid(row=1, column=0, columnspan=4, padx=10, pady=(10, 4))
        pdf_button_frame = ttk.Frame(date_frame)
        pdf_button_frame.grid(row=2, column=0, columnspan=4, padx=10, pady=(0, 10))

        ttk.Button(button_frame, text="Generate Full Financial Statements",
                  command=self.generate_comprehensive_report, width=25).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Export to HTML",
                  command=self.export_report, width=15).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Balance Sheet",
                  command=self.generate_balance_sheet, width=15).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Income Statement",
                  command=self.generate_income_statement, width=15).pack(side='left', padx=5)
        ttk.Button(button_frame, text="AR Aging",
                  command=self.generate_ar_aging, width=15).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cash Flow",
                  command=self.generate_cash_flow_statement, width=15).pack(side='left', padx=5)

        # Professional PDF export row
        ttk.Label(pdf_button_frame, text="📄 Export Professional PDF:",
                 font=('Helvetica', 9, 'bold')).pack(side='left', padx=(0, 8))
        ttk.Button(pdf_button_frame, text="Balance Sheet PDF",
                  command=self.export_balance_sheet_pdf, width=18).pack(side='left', padx=3)
        ttk.Button(pdf_button_frame, text="Income Statement PDF",
                  command=self.export_income_statement_pdf, width=20).pack(side='left', padx=3)
        ttk.Button(pdf_button_frame, text="AR Aging PDF",
                  command=self.export_ar_aging_pdf, width=14).pack(side='left', padx=3)
        ttk.Button(pdf_button_frame, text="Cash Flow PDF",
                  command=self.export_cash_flow_pdf, width=15).pack(side='left', padx=3)
        ttk.Button(pdf_button_frame, text="Full Report PDF",
                  command=self.export_comprehensive_report_pdf, width=16).pack(side='left', padx=3)
        ttk.Button(pdf_button_frame, text="🗂 Batch Export All (5 PDFs)",
                  command=self.export_all_reports_batch_pdf, width=24,
                  style='Accent.TButton').pack(side='left', padx=5)
        
        # Report content
        report_frame = ttk.LabelFrame(main_frame, text="Financial Statements", padding="10")
        report_frame.pack(fill='both', expand=True, pady=10)
        
        self.report_text = scrolledtext.ScrolledText(report_frame, height=35, font=('Consolas', 9))
        self.report_text.pack(fill='both', expand=True)
        
        # Generate initial report
        self.generate_comprehensive_report()
    
    def get_actual_financial_data(self, start_date, end_date):
        """Extract ACTUAL financial data from all application sources"""
        try:
            # Get all invoices within period
            all_invoices = self.manager.get_all_invoices_dict()
            period_invoices = []
            for inv in all_invoices:
                inv_date = inv.get('date', '')
                if inv_date and start_date <= inv_date <= end_date:
                    period_invoices.append(inv)
            
            # Get all purchases within period
            all_purchases = self.manager.get_all_purchases_dict()
            period_purchases = []
            for purchase in all_purchases:
                purchase_date = purchase.get('date', '')
                if purchase_date and start_date <= purchase_date <= end_date:
                    period_purchases.append(purchase)
            
            # Calculate ACTUAL revenue and costs from invoices
            total_revenue = sum(inv.get('grand_total', 0) for inv in period_invoices)
            total_cost_of_sales = sum(inv.get('total_cost', 0) for inv in period_invoices)
            gross_profit = total_revenue - total_cost_of_sales
            
            # Calculate ACTUAL operating expenses from purchases
            operating_expenses = sum(purchase.get('amount', 0) for purchase in period_purchases)
            
            # Calculate ACTUAL salary expenses from employee records
            # IMPORTANT: Use professional employee_manager.py EmployeeManager
            # (reads monthly_salaries.json) and new get_salary_totals_for_period helper
            # that respects the exact chosen date range and counts ALL months in range
            # (not just 1 month PDF cache from the old local salary_reports.json copy)
            from employee_manager import EmployeeManager as ProEmployeeManager
            pro_emp_mgr = ProEmployeeManager(self.manager.invoice_folder)
            salary_summary = pro_emp_mgr.get_salary_summary_for_period(start_date, end_date)
            salary_expenses = float(salary_summary["gross_salary_total"])
            salary_net_paid = float(salary_summary["net_salary_total"])
            salary_months_used = salary_summary["months_in_period"]
            salary_employees_used = salary_summary["employees_processed"]
            
            total_operating_expenses = operating_expenses + salary_expenses
            net_profit = gross_profit - total_operating_expenses
            
            # Calculate ACTUAL accounts receivable (unpaid invoices)
            accounts_receivable = sum(
                inv.get('grand_total', 0) - inv.get('total_paid', 0) 
                for inv in period_invoices 
                if inv.get('grand_total', 0) > inv.get('total_paid', 0)
            )
            
            # Calculate ACTUAL accounts payable (unpaid purchases)
            accounts_payable = sum(
                purchase.get('amount', 0) 
                for purchase in period_purchases 
                if not purchase.get('paid', False)  # Assuming you track payment status
            )
            
            # Calculate ACTUAL cash position from all transactions
            cash_balance = self.calculate_actual_cash_balance(start_date, end_date)
            
            # Calculate ACTUAL VAT collected and paid
            vat_collected = sum(inv.get('tax_amount', 0) for inv in period_invoices)
            vat_paid = sum(purchase.get('amount', 0) * 0.05 for purchase in period_purchases)  # Assuming 5% VAT on purchases
            
            return {
                'period': {'start_date': start_date, 'end_date': end_date},
                'income_statement': {
                    'revenue': total_revenue,
                    'cost_of_sales': total_cost_of_sales,
                    'gross_profit': gross_profit,
                    'gross_profit_margin': (gross_profit / total_revenue * 100) if total_revenue > 0 else 0,
                    'operating_expenses': total_operating_expenses,
                    'salary_expenses': salary_expenses,
                    'purchase_expenses': operating_expenses,
                    'net_profit': net_profit,
                    'net_profit_margin': (net_profit / total_revenue * 100) if total_revenue > 0 else 0,
                    'vat_collected': vat_collected,
                    'vat_paid': vat_paid
                },
                'balance_sheet': {
                    'accounts_receivable': accounts_receivable,
                    'accounts_payable': accounts_payable,
                    'cash_balance': cash_balance,
                    'vat_payable': vat_collected - vat_paid,  # VAT owed to government
                    'retained_earnings': net_profit  # Simplified - should accumulate over time
                },
                'cash_flow': {
                    'cash_from_sales': sum(inv.get('total_paid', 0) for inv in period_invoices),
                    'cash_to_suppliers': sum(purchase.get('amount', 0) for purchase in period_purchases if purchase.get('paid', True)),
                    'cash_to_employees': salary_expenses,
                    'net_cash_flow': self.calculate_net_cash_flow(period_invoices, period_purchases, salary_expenses)
                },
                'metrics': {
                    'total_invoices': len(period_invoices),
                    'total_purchases': len(period_purchases),
                    'paid_invoices': len([inv for inv in period_invoices if inv.get('total_paid', 0) >= inv.get('grand_total', 0)]),
                    'unpaid_invoices': len([inv for inv in period_invoices if inv.get('total_paid', 0) < inv.get('grand_total', 0)]),
                    'paid_purchases': len([p for p in period_purchases if p.get('paid', False)]),
                    'unpaid_purchases': len([p for p in period_purchases if not p.get('paid', False)])
                }
            }
            
        except Exception as e:
            print(f"Error getting actual financial data: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def calculate_actual_cash_balance(self, start_date, end_date):
        """Calculate ACTUAL cash balance from all cash movements"""
        try:
            # Get all invoices and calculate cash received
            all_invoices = self.manager.get_all_invoices_dict()
            cash_from_customers = sum(
                inv.get('total_paid', 0) 
                for inv in all_invoices 
                if inv.get('date', '') and start_date <= inv.get('date', '') <= end_date
            )
            
            # Get all purchases and calculate cash paid to suppliers
            all_purchases = self.manager.get_all_purchases_dict()
            cash_to_suppliers = sum(
                purchase.get('amount', 0) 
                for purchase in all_purchases 
                if purchase.get('date', '') and start_date <= purchase.get('date', '') <= end_date and purchase.get('paid', False)
            )
            
            # Get salary payments — use professional EmployeeManager with correct period filter
            from employee_manager import EmployeeManager as ProEmployeeManager2
            pro_emp_mgr2 = ProEmployeeManager2(self.manager.invoice_folder)
            salary_period = pro_emp_mgr2.get_salary_totals_for_period(
                start_date, end_date, use_field="net_salary")
            cash_to_employees = float(salary_period["total_amount"])
            
            # Simple cash calculation (starting balance + inflows - outflows)
            # You might want to track opening cash balance separately
            opening_cash_balance = 0  # This should be tracked in your system
            net_cash_flow = cash_from_customers - cash_to_suppliers - cash_to_employees
            
            return opening_cash_balance + net_cash_flow
            
        except Exception as e:
            print(f"Error calculating cash balance: {e}")
            return 0
    
    def calculate_net_cash_flow(self, invoices, purchases, salary_expenses):
        """Calculate net cash flow from operations"""
        cash_inflows = sum(inv.get('total_paid', 0) for inv in invoices)
        cash_outflows = sum(purchase.get('amount', 0) for purchase in purchases if purchase.get('paid', False)) + salary_expenses
        return cash_inflows - cash_outflows
    
    def generate_comprehensive_report(self):
        """Generate full financial statements using ACTUAL data"""
        try:
            start_date = self.start_date.get()
            end_date = self.end_date.get()
            
            financial_data = self.get_actual_financial_data(start_date, end_date)
            if not financial_data:
                self.report_text.delete('1.0', tk.END)
                self.report_text.insert('1.0', "No financial data available for the selected period.")
                return
            
            report = self._format_comprehensive_report(financial_data)
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', report)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate comprehensive report: {e}")
    
    def generate_balance_sheet(self):
        """Generate balance sheet only — pulls from GAAP ReportEngine for fully balanced, inventory-inclusive figures"""
        try:
            from report_engine import ReportEngine
            from datetime import date as _date
            sd = self.start_date.get().strip()
            ed = self.end_date.get().strip()
            # Try to use the PROPER ReportEngine first (always balanced, uses GL single-source)
            try:
                sd_d = _date.fromisoformat(sd)
                ed_d = _date.fromisoformat(ed)
                engine = ReportEngine(self.manager.invoice_folder, sd_d, ed_d)
                bs = engine.get_balance_sheet()
                L = bs.get("lines", {}) or {}
                assets = L.get("assets", {}) or {}
                liabs = L.get("liabilities", {}) or {}
                equity = L.get("equity", {}) or {}
                is_ok = L.get("is_balanced", True)
                variance = Decimal(str(L.get("variance", 0)))
                def _g(d, k, fallback=Decimal("0")):
                    try:
                        v = d.get(k, fallback)
                        if v is None: return Decimal("0")
                        return Decimal(str(v))
                    except Exception:
                        return Decimal("0")
                cash = _g(assets, "1000_Cash") + _g(assets, "1100_Bank")
                ar = _g(assets, "1200_Accounts_Receivable")
                inv = _g(assets, "1300_Inventory")
                ta = _g(assets, "total_assets")
                ap = _g(liabs, "2000_Accounts_Payable")
                accrued = _g(liabs, "2100_Accrued_Expenses")
                tl = _g(liabs, "total_liabilities")
                oe = _g(equity, "3100_Owner_Equity")
                re = _g(equity, "3000_Retained_Earnings")
                od = abs(_g(equity, "3200_Owner_Drawings"))
                cpp = _g(equity, "Current_Period_Net_Profit")
                teq = _g(equity, "total_equity")
                tle = _g(L, "total_le")
                def fmt(v):
                    try: d = Decimal(str(v))
                    except Exception: return "           0.00"
                    sign = "(" if d < 0 else ""
                    close = ")" if d < 0 else ""
                    return f"{sign}{abs(d):>12,.2f}{close}"
                bal_status = "✅ BALANCED (GAAP Compliant)" if is_ok else f"⚠️ UNBALANCED by {fmt(variance)} — run GL Rebuild"
                report = f"""
HOPE PHARMA MEDICINE TRADING
STATEMENT OF FINANCIAL POSITION (BALANCE SHEET)
AS AT {ed}
{'='*80}
[Source: GAAP ReportEngine — General Ledger single source of truth]
STATUS: {bal_status}

ASSETS                                                AED
-------------------------------------------------------
CURRENT ASSETS:
  Cash and Cash Equivalents (1000+1100)             {fmt(cash)}
  Trade Accounts Receivable (1200)                   {fmt(ar)}
  Inventories — Pharma / Medical Stock (1300)        {fmt(inv)}
  Total Current Assets                               {fmt(cash + ar + inv)}

NON-CURRENT ASSETS:
  Property, Plant & Equipment (net)                  {fmt(ta - cash - ar - inv) if ta > (cash+ar+inv) else fmt(0)}
  Total Non-Current Assets                           {fmt(ta - cash - ar - inv) if ta > (cash+ar+inv) else fmt(0)}

TOTAL ASSETS                                         {fmt(ta)}

LIABILITIES AND EQUITY
-------------------------------------------------------
CURRENT LIABILITIES:
  Trade Accounts Payable — Suppliers (2000)          {fmt(ap)}
  Accrued Expenses / Provisions (2100)               {fmt(accrued)}
  Total Current Liabilities                          {fmt(tl)}

NON-CURRENT LIABILITIES:
  Long-term Loans / Borrowings                       {fmt(0)}
  Total Non-Current Liabilities                      {fmt(0)}

TOTAL LIABILITIES                                    {fmt(tl)}

EQUITY:
  Owner / Share Capital (3100)                       {fmt(oe)}
  Retained Earnings — Accumulated (3000)             {fmt(re)}
  Current Period Net Profit / (Loss)                 {fmt(cpp)}
  Less: Owner Drawings during period                 {fmt(od) if od != 0 else fmt(0)}
  TOTAL EQUITY (Net Assets attributable to Owner)    {fmt(teq)}

TOTAL LIABILITIES AND EQUITY                         {fmt(tle)}

{'='*80}
Accounting Equation Check:
  Total Assets:      {fmt(ta)}
  Liabilities + Eq:  {fmt(tle)}
  Variance (A − L−E):{fmt(variance)}   →   {"✅ MATCH (0.00 AED)" if abs(variance) < Decimal("0.01") else "❌ MISMATCH — Rebuild GL Now"}
{'='*80}
Note: This legacy report now pulls directly from the GAAP General Ledger. For the fully
professional IFRS presentation, use the GAAP Financial Reporting Centre → Balance Sheet
tab and click [Export PDF].
{'='*80}
"""
                self.report_text.delete('1.0', tk.END)
                self.report_text.insert('1.0', report)
                return
            except Exception as exc_inner:
                print(f"Legacy BS: GAAP engine fallback unavailable ({exc_inner}); using ad-hoc calc mode")
        except Exception as exc_import:
            pass
        # --- FALLBACK: Original ad-hoc calc but with bugs fixed ---
        try:
            start_date = self.start_date.get()
            end_date = self.end_date.get()
            financial_data = self.get_actual_financial_data(start_date, end_date)
            if not financial_data:
                self.report_text.delete('1.0', tk.END)
                self.report_text.insert('1.0', "No financial data available for the selected period.")
                return
            balance = financial_data['balance_sheet']
            income = financial_data['income_statement']
            # Pull inventory from InventoryManager
            inventory_val = Decimal("0.00")
            try:
                from inventory_system import InventoryManager
                im = InventoryManager(self.manager.invoice_folder)
                for it in (getattr(im, 'items', None) or []):
                    try:
                        q = float(getattr(it, 'quantity', 0) or 0)
                        cp = float(getattr(it, 'cost_price', 0) or getattr(it, 'unit_cost', 0) or 0)
                        inventory_val += Decimal(str(q * cp))
                    except Exception:
                        continue
                inventory_val = inventory_val.quantize(Decimal("0.01"))
            except Exception:
                pass
            # Fixed: vat_paid lives in income_statement, NOT balance dict!
            vat_paid = Decimal(str(income.get('vat_paid', 0) or 0))
            vat_collected = Decimal(str(income.get('vat_collected', 0) or 0))
            vat_receivable = max(Decimal("0"), vat_paid - vat_collected)
            vat_payable = max(Decimal("0"), vat_collected - vat_paid)
            # Fixed: Accumulated Retained Earnings (not just current period)
            # Hardcoded 300k initial capital replaced with: owner_capital is set from 3100 Owner Equity if possible, else 300k
            share_capital = Decimal("300000.00")
            try:
                from report_engine import ReportEngine
                from datetime import date as _date2
                _e = ReportEngine(self.manager.invoice_folder,
                                  _date2.fromisoformat(start_date), _date2.fromisoformat(end_date))
                _bs2 = _e.get_balance_sheet()
                _L2 = _bs2.get("lines", {}) or {}
                _E2 = _L2.get("equity", {}) or {}
                v = _E2.get("3100_Owner_Equity")
                if v is not None:
                    share_capital = Decimal(str(v)).quantize(Decimal("0.01"))
            except Exception:
                pass
            cash_bal = Decimal(str(balance.get('cash_balance', 0) or 0))
            ar_bal = Decimal(str(balance.get('accounts_receivable', 0) or 0))
            ap_bal = Decimal(str(balance.get('accounts_payable', 0) or 0))
            net_profit = Decimal(str(income.get('net_profit', 0) or 0))
            tca = cash_bal + ar_bal + inventory_val + vat_receivable
            ta_total = tca  # no NCA in fallback
            tcl_total = ap_bal + vat_payable
            eq_total = share_capital + net_profit
            tle_total = tcl_total + eq_total
            def fmt2(v):
                try: d = Decimal(str(v))
                except Exception: return "           0.00"
                return f"{d:>12,.2f}"
            report = f"""
HOPE PHARMA MEDICINE TRADING
STATEMENT OF FINANCIAL POSITION (BALANCE SHEET)
AS AT {end_date}
{'='*80}
[Fallback Mode: ad-hoc data aggregation — for professional balanced BS, use GAAP Centre]

ASSETS                                                AED
-------------------------------------------------------
CURRENT ASSETS:
  Cash and Cash Equivalents                          {fmt2(cash_bal)}
  Accounts Receivable (Trade Debtors)                {fmt2(ar_bal)}
  Inventories (Pharma / Medical Stock)               {fmt2(inventory_val)}
  VAT Receivable                                     {fmt2(vat_receivable)}
  Total Current Assets                               {fmt2(tca)}

NON-CURRENT ASSETS:
  Property, Plant and Equipment                      {fmt2(0)}
  Total Non-Current Assets                           {fmt2(0)}

TOTAL ASSETS                                         {fmt2(ta_total)}

LIABILITIES AND EQUITY
-------------------------------------------------------
CURRENT LIABILITIES:
  Accounts Payable (Trade Creditors)                 {fmt2(ap_bal)}
  VAT Payable                                        {fmt2(vat_payable)}
  Total Current Liabilities                          {fmt2(tcl_total)}

EQUITY:
  Share Capital                                      {fmt2(share_capital)}
  Retained Earnings (incl. current period profit)    {fmt2(net_profit)}
  Total Equity                                       {fmt2(eq_total)}

TOTAL LIABILITIES AND EQUITY                         {fmt2(tle_total)}

{'='*80}
Variance (A − (L + E)):                              {fmt2(ta_total - tle_total)}
For a perfectly balanced GAAP statement, go to:
  Main Menu → Finance → GAAP Financial Reports → Balance Sheet
{'='*80}
"""
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', report)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to generate balance sheet: {e}")
    
    def generate_income_statement(self):
        """Generate income statement only"""
        try:
            start_date = self.start_date.get()
            end_date = self.end_date.get()
            
            financial_data = self.get_actual_financial_data(start_date, end_date)
            if not financial_data:
                self.report_text.delete('1.0', tk.END)
                self.report_text.insert('1.0', "No financial data available for the selected period.")
                return
            
            income = financial_data['income_statement']
            
            report = f"""
HOPE PHARMA MEDICINE TRADING
STATEMENT OF COMPREHENSIVE INCOME (PROFIT & LOSS)
FOR THE PERIOD {start_date} TO {end_date}
{'='*80}

REVENUE:
  Sales Revenue                                      {income['revenue']:>12,.2f}

COST OF SALES:
  Cost of Goods/Services Sold                       {-income['cost_of_sales']:>12,.2f}
-------------------------------------------------------
GROSS PROFIT                                         {income['gross_profit']:>12,.2f}

OPERATING EXPENSES:
  Employee Salaries and Benefits                    {-income['salary_expenses']:>12,.2f}
  Purchase Expenses (Supplies, Services)           {-income['purchase_expenses']:>12,.2f}
  Total Operating Expenses                          {-income['operating_expenses']:>12,.2f}
-------------------------------------------------------
OPERATING PROFIT                                     {income['gross_profit'] - income['operating_expenses']:>12,.2f}

OTHER ITEMS:
  VAT Collected                                      {income['vat_collected']:>12,.2f}
  VAT Paid                                          {-income['vat_paid']:>12,.2f}
-------------------------------------------------------
NET PROFIT BEFORE TAX                                {income['net_profit']:>12,.2f}

Income Tax Expense                                  {0:>12,.2f}  /* Consult tax advisor */
-------------------------------------------------------
NET PROFIT FOR THE PERIOD                            {income['net_profit']:>12,.2f}

PROFITABILITY RATIOS:
  Gross Profit Margin: {income['gross_profit_margin']:.2f}%
  Net Profit Margin: {income['net_profit_margin']:.2f}%

{'='*80}
"""
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', report)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate income statement: {e}")
    
    def generate_ar_aging(self):
        try:
            end_date = self.end_date.get()
            invoices = self.manager.get_all_invoices_dict()
            buckets = {
                'Current': 0.0,
                '1-30': 0.0,
                '31-60': 0.0,
                '61-90': 0.0,
                '>90': 0.0
            }
            details = {k: [] for k in buckets.keys()}
            today = datetime.strptime(end_date, "%Y-%m-%d")
            for inv in invoices:
                total = float(inv.get('grand_total', 0) or 0)
                paid = float(inv.get('total_paid', 0) or 0)
                balance = max(total - paid, 0.0)
                if balance <= 0:
                    continue
                due = inv.get('due_date') or inv.get('date')
                try:
                    due_dt = datetime.strptime(due, "%Y-%m-%d")
                except Exception:
                    due_dt = today
                days = (today - due_dt).days
                if days <= 0:
                    bucket = 'Current'
                elif days <= 30:
                    bucket = '1-30'
                elif days <= 60:
                    bucket = '31-60'
                elif days <= 90:
                    bucket = '61-90'
                else:
                    bucket = '>90'
                buckets[bucket] += balance
                details[bucket].append((inv.get('invoice_id',''), inv.get('client_name',''), balance, days))
            total_ar = sum(buckets.values())
            out = []
            out.append("HOPE PHARMA MEDICINE TRADING")
            out.append("ACCOUNTS RECEIVABLE AGING")
            out.append(f"AS AT {end_date}")
            out.append("="*70)
            out.append("")
            out.append(f"Total AR: AED {total_ar:,.2f}")
            out.append("")
            for k in ['Current','1-30','31-60','61-90','>90']:
                out.append(f"{k:<10} AED {buckets[k]:>12,.2f}")
            out.append("")
            for k in ['>90','61-90','31-60','1-30','Current']:
                if details[k]:
                    out.append("")
                    out.append(f"{k} Details:")
                    for inv_id, client, bal, d in details[k]:
                        out.append(f"  {inv_id} - {client:<30} AED {bal:>12,.2f} ({d} days)")
            text = "\n".join(out)
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate AR aging: {e}")

    def generate_cash_flow_statement(self):
        try:
            start_date = self.start_date.get()
            end_date = self.end_date.get()
            data = self.get_actual_financial_data(start_date, end_date)
            if not data:
                self.report_text.delete('1.0', tk.END)
                self.report_text.insert('1.0', "No financial data available for the selected period.")
                return
            cf = data['cash_flow']
            balance = data['balance_sheet']
            out = []
            out.append("HOPE PHARMA MEDICINE TRADING")
            out.append("STATEMENT OF CASH FLOWS")
            out.append(f"FOR THE PERIOD {start_date} TO {end_date}")
            out.append("="*80)
            out.append("")
            out.append("Cash flows from operating activities")
            out.append(f"  Cash received from customers      {cf['cash_from_sales']:>12,.2f}")
            out.append(f"  Cash paid to suppliers            {-sum(p.get('amount',0) for p in self.manager.get_all_purchases_dict() if p.get('paid', False) and start_date <= (p.get('date','') or '') <= end_date):>12,.2f}")
            out.append(f"  Cash paid to employees            {-cf['cash_to_employees']:>12,.2f}")
            out.append(f"Net cash from operating activities  {cf['net_cash_flow']:>12,.2f}")
            out.append("")
            out.append(f"Cash and cash equivalents, ending   {balance['cash_balance']:>12,.2f}")
            text = "\n".join(out)
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', text)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate cash flow: {e}")
    
    def _format_comprehensive_report(self, data):
        """Format the comprehensive financial report using ACTUAL data"""
        income = data['income_statement']
        balance = data['balance_sheet']
        cash_flow = data['cash_flow']
        metrics = data['metrics']
        
        report = f"""
HOPE PHARMA MEDICINE TRADING
DUBAI - UNITED ARAB EMIRATES
COMPREHENSIVE FINANCIAL STATEMENTS
{data['period']['end_date']}

{'='*80}

MANAGING DIRECTOR'S REPORT

The Managing Director presents the financial statements for the period 
{data['period']['start_date']} to {data['period']['end_date']}.

FINANCIAL HIGHLIGHTS:
• Total Revenue: AED {income['revenue']:,.2f}
• Gross Profit: AED {income['gross_profit']:,.2f} ({income['gross_profit_margin']:.2f}%)
• Net Profit: AED {income['net_profit']:,.2f} ({income['net_profit_margin']:.2f}%)
• Cash Balance: AED {balance['cash_balance']:,.2f}
• Accounts Receivable: AED {balance['accounts_receivable']:,.2f}

OPERATIONAL SUMMARY:
• Invoices Issued: {metrics['total_invoices']}
• Purchase Orders: {metrics['total_purchases']}
• Collection Rate: {(metrics['paid_invoices']/metrics['total_invoices']*100) if metrics['total_invoices'] > 0 else 0:.1f}%

{'='*80}

STATEMENT OF FINANCIAL POSITION
AS AT {data['period']['end_date']}
{'='*80}

ASSETS                                                AED
-------------------------------------------------------
Current Assets:
  Cash and Bank Balances                            {balance['cash_balance']:>12,.2f}
  Accounts Receivable                               {balance['accounts_receivable']:>12,.2f}
  Total Current Assets                              {balance['cash_balance'] + balance['accounts_receivable']:>12,.2f}

TOTAL ASSETS                                         {balance['cash_balance'] + balance['accounts_receivable']:>12,.2f}

LIABILITIES AND EQUITY
-------------------------------------------------------
Current Liabilities:
  Accounts Payable                                  {balance['accounts_payable']:>12,.2f}
  VAT Payable                                       {max(0, income['vat_collected'] - income['vat_paid']):>12,.2f}
  Total Current Liabilities                         {balance['accounts_payable'] + max(0, income['vat_collected'] - income['vat_paid']):>12,.2f}

Equity:
  Share Capital                                     {300000:>12,.2f}
  Retained Earnings                                 {income['net_profit']:>12,.2f}
  Total Equity                                      {300000 + income['net_profit']:>12,.2f}

TOTAL LIABILITIES AND EQUITY                         {balance['accounts_payable'] + max(0, income['vat_collected'] - income['vat_paid']) + 300000 + income['net_profit']:>12,.2f}

{'='*80}

STATEMENT OF COMPREHENSIVE INCOME
FOR THE PERIOD ENDED {data['period']['end_date']}
{'='*80}

                                                  AED
Revenue                                      {income['revenue']:>12,.2f}
Cost of Sales                               {-income['cost_of_sales']:>12,.2f}
-------------------------------------------------------
Gross Profit                                 {income['gross_profit']:>12,.2f}

Operating Expenses:
  Salaries and Wages                        {-income['salary_expenses']:>12,.2f}
  Other Operating Expenses                  {-income['purchase_expenses']:>12,.2f}
-------------------------------------------------------
Total Operating Expenses                    {-income['operating_expenses']:>12,.2f}
-------------------------------------------------------
Operating Profit                             {income['gross_profit'] - income['operating_expenses']:>12,.2f}

VAT (Net)                                    {income['vat_collected'] - income['vat_paid']:>12,.2f}
-------------------------------------------------------
NET PROFIT                                   {income['net_profit']:>12,.2f}

{'='*80}

STATEMENT OF CASH FLOWS
FOR THE PERIOD ENDED {data['period']['end_date']}
{'='*80}

Cash flows from operating activities:              AED
Cash received from customers                      {cash_flow['cash_from_sales']:>12,.2f}
Cash paid to suppliers and employees             {-cash_flow['cash_to_suppliers'] - cash_flow['cash_to_employees']:>12,.2f}
-------------------------------------------------------
Net cash from operating activities                {cash_flow['net_cash_flow']:>12,.2f}

Net increase in cash                              {cash_flow['net_cash_flow']:>12,.2f}
Cash at beginning of period                       {0:>12,.2f}  /* Track opening balance */
Cash at end of period                             {balance['cash_balance']:>12,.2f}

{'='*80}

FINANCIAL RATIOS
{'='*80}

Profitability:
  Gross Profit Margin: {income['gross_profit_margin']:.2f}%
  Net Profit Margin: {income['net_profit_margin']:.2f}%

Liquidity:
  Current Ratio: {(balance['cash_balance'] + balance['accounts_receivable']) / (balance['accounts_payable'] + max(0.01, income['vat_collected'] - income['vat_paid'])):.2f}

Efficiency:
  Receivable Collection: {(cash_flow['cash_from_sales'] / max(1, income['revenue'])):.2%}

{'='*80}

NOTES:
1. All figures are generated from actual transaction data in the system
2. Revenue represents invoiced amounts, not necessarily collected
3. Expenses include all recorded purchases and salary payments
4. VAT calculations based on 5% standard rate
5. Fixed assets tracking to be implemented separately

{'='*80}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
{'='*80}
"""
        return report
    
    def export_report(self):
        """Export the current report to HTML"""
        try:
            current_content = self.report_text.get('1.0', tk.END)
            if not current_content.strip():
                messagebox.showwarning("Warning", "No report content to export")
                return
            
            # Create financial reports folder
            reports_folder = os.path.join(self.manager.invoice_folder, "FinancialReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)
            
            # Export as HTML
            filename = f"Financial_Statements_{self.start_date.get()}_to_{self.end_date.get()}.html"
            filepath = os.path.join(reports_folder, filename)
            
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Financial Statements - Hope Pharma</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    .header {{ text-align: center; border-bottom: 2px solid #2c5aa0; padding-bottom: 20px; }}
                    .company-name {{ font-size: 24px; font-weight: bold; color: #2c5aa0; }}
                    .content {{ font-family: 'Courier New', monospace; font-size: 12px; white-space: pre-wrap; line-height: 1.2; }}
                    .footer {{ margin-top: 30px; text-align: center; color: #666; font-size: 11px; }}
                </style>
            </head>
            <body>
                <div class="header">
                    <div class="company-name">HOPE PHARMA MEDICINE TRADING</div>
                    <div>Comprehensive Financial Statements</div>
                    <div>Period: {self.start_date.get()} to {self.end_date.get()}</div>
                </div>
                <div class="content">{current_content}</div>
                <div class="footer">
                    Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} | From Actual System Data
                </div>
            </body>
            </html>
            """
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            # Open in browser
            webbrowser.open('file://' + os.path.abspath(filepath))
            messagebox.showinfo("Success", f"Financial statements exported to:\n{filepath}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report: {e}")

    # ------------------------------------------------------------------
    # PDF Export methods (Professional brand PDFs for the 4 GAAP statements + batch
    # comprehensive + one-click batch export (5 reports + index).
    # ------------------------------------------------------------------
    def _get_report_engine(self):
        from report_engine import ReportEngine, DateRangeType
        start = self.start_date.get().strip()
        end = self.end_date.get().strip()
        try:
            return ReportEngine(self.manager, DateRangeType.CUSTOM, start, end)
        except Exception as e:
            messagebox.showerror("Error", f"Could not build report engine for the chosen period:\n{e}")
            return None

    def export_balance_sheet_pdf(self):
        try:
            engine = self._get_report_engine()
            if engine is None: return
            bs = engine.get_balance_sheet()
            path = PDFGenerator().generate_balance_sheet_pdf(bs)
            messagebox.showinfo("Success", f"Balance Sheet PDF exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Balance Sheet PDF: {e}")

    def export_income_statement_pdf(self):
        try:
            engine = self._get_report_engine()
            if engine is None: return
            pnl = engine.get_profit_and_loss()
            path = PDFGenerator().generate_income_statement_pdf(pnl)
            messagebox.showinfo("Success", f"Income Statement PDF exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Income Statement PDF: {e}")

    def export_ar_aging_pdf(self):
        try:
            engine = self._get_report_engine()
            if engine is None: return
            ar = engine.get_aging_receivables()
            path = PDFGenerator().generate_ar_aging_pdf(ar)
            messagebox.showinfo("Success", f"AR Aging PDF exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export AR Aging PDF: {e}")

    def export_cash_flow_pdf(self):
        try:
            engine = self._get_report_engine()
            if engine is None: return
            cf = engine.get_cash_flow()
            path = PDFGenerator().generate_cash_flow_pdf(cf)
            messagebox.showinfo("Success", f"Cash Flow PDF exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export Cash Flow PDF: {e}")

    def export_comprehensive_report_pdf(self):
        """Build a one-shot consolidated PDF that stacks all 4 statements.
        Uses the batch generator and uses the folder it produces, then produces a combined
        single PDF via a cover sheet with individual PDFs merged via PyPDF2 if available.
        Fallback: notifies user that batch folder was created (merges only on systems with PyPDF2 installed).
        """
        try:
            result = PDFGenerator().generate_all_reports_pdf_batch(
                data_manager=self.manager,
                start_date=self.start_date.get().strip(),
                end_date=self.end_date.get().strip()
            )
            files = result.get("files", {})
            # Attempt to merge all PDFs produced (cover PDF; fall back to opening batch folder
            merged = None
            try:
                order = ["__index","balance_sheet","income_statement","cash_flow","ar_aging","trial_balance"]
                ordered_paths = [files[k] for k in order if files.get(k)]
                # Try PyPDF2 / pypdf mergers if present
                merger = None
                try:
                    from PyPDF2 import PdfMerger as _Merger; merger = _Merger()
                except Exception:
                    try:
                        from pypdf import PdfMerger as _Merger; merger = _Merger()
                    except Exception:
                        merger = None
                if merger is not None and ordered_paths:
                    for p in ordered_paths:
                        try: merger.append(p)
                        except Exception: pass
                    folder = result["folder"]
                    merged_pdf = os.path.join(folder,
                        f"Full_Financial_Statements_{self.start_date.get().strip()}_to_{self.end_date.get().strip()}.pdf")
                    merger.write(merged_pdf)
                    merger.close()
                    merged = merged_pdf
                    # Open merged file
                    try:
                        import platform, subprocess
                        if platform.system() == 'Darwin': subprocess.run(['open', merged], check=False)
                        elif platform.system() == 'Windows': os.startfile(merged) # noqa attr-defined
                        else: subprocess.run(['xdg-open', merged], check=False)
                    except Exception: pass
            except Exception as exc:
                print(f"[INFO] PDF merge skipped ({exc})")

            msg = f"Full financial statements exported:\nBatch folder: {result['folder']}"
            if merged: msg += f"\n\nCombined single PDF:\n{merged}"
            else: msg += "\n\n(Tip: pip install PyPDF2 to produce a single combined PDF instead of 5 separate files.)"
            messagebox.showinfo("Success", msg)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export full report PDF: {e}")

    def export_all_reports_batch_pdf(self):
        try:
            result = PDFGenerator().generate_all_reports_pdf_batch(
                data_manager=self.manager,
                start_date=self.start_date.get().strip(),
                end_date=self.end_date.get().strip()
            )
            files = result.get("files", {})
            names = []
            for k, v in files.items():
                if k == "__index": names.append(f"  INDEX: {os.path.basename(v)}")
                else: names.append(f"  {k}: {os.path.basename(v)}")
            names.sort()
            listing = "\n".join(names)
            messagebox.showinfo("Success",
                f"Batch export completed successfully!\n\n"
                f"Folder: {result['folder']}\n\n"
                f"Generated {len(files)} report PDFs:\n{listing}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to batch export reports PDF: {e}")


class TrialBalanceDialog:
    """Trial Balance dialog: period picker + 6-column TB table + PDF export."""
    def __init__(self, parent, manager):
        self.manager = manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Trial Balance — General Ledger")
        self.dialog.geometry("1280x780")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        try:
            lm = LogoManager()
            if lm.load_logo(size=(64, 64)):
                self.dialog.iconphoto(True, lm.logo_photo)
        except Exception:
            pass
        self.last_trial_balance = None
        self.setup_ui()
        _fit_window(self.dialog, 1280, 780, mode="workspace", remember_key="trial_balance")

    def setup_ui(self):
        main = ttk.Frame(self.dialog, padding=10)
        main.pack(fill='both', expand=True)
        # Title bar
        title_row = ttk.Frame(main)
        title_row.pack(fill='x')
        ttk.Label(title_row, text="🧾  Trial Balance",
                  font=('Helvetica', 18, 'bold')).pack(side='left')
        self.balance_lbl = ttk.Label(title_row, text="—",
                                     font=('Helvetica', 11, 'bold'))
        self.balance_lbl.pack(side='right')

        # Period bar
        period_bar = ttk.LabelFrame(main, text="Report Period", padding=8)
        period_bar.pack(fill='x', pady=(10, 8))
        ttk.Label(period_bar, text="Start Date:").grid(row=0, column=0, padx=4, pady=2, sticky='e')
        self.start_date = ttk.Entry(period_bar, width=12)
        self.start_date.grid(row=0, column=1, padx=4, pady=2)
        self.start_date.insert(0, (datetime.now().replace(month=1, day=1)).strftime("%Y-%m-%d"))
        ttk.Label(period_bar, text="End Date:").grid(row=0, column=2, padx=(12, 4), pady=2, sticky='e')
        self.end_date = ttk.Entry(period_bar, width=12)
        self.end_date.grid(row=0, column=3, padx=4, pady=2)
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))

        btns = ttk.Frame(period_bar)
        btns.grid(row=0, column=4, padx=(20, 0))
        ttk.Button(btns, text="🔄  Generate Trial Balance",
                   style="Primary.TButton",
                   command=self.generate).pack(side='left', padx=6)
        ttk.Button(btns, text="📤  Export to PDF",
                   command=self.export_pdf).pack(side='left', padx=6)
        ttk.Button(btns, text="↻  This Year",
                   command=self.preset_this_year).pack(side='left', padx=6)
        ttk.Button(btns, text="↻  All Time",
                   command=self.preset_all_time).pack(side='left', padx=6)

        # Summary bar
        self.summary_lbl = ttk.Label(main, text="",
                                     font=('Helvetica', 10), foreground="#2D3748")
        self.summary_lbl.pack(fill='x', pady=(0, 8))

        # Table
        table_frame = ttk.LabelFrame(main, text="Trial Balance (Opening / Period / Closing)", padding=4)
        table_frame.pack(fill='both', expand=True)
        cols = ("account_code","account_name","account_type",
                "opening_debit","opening_credit","period_debit","period_credit",
                "closing_debit","closing_credit")
        headers = [
            ("account_code", "Code", 70),
            ("account_name", "Account Name", 240),
            ("account_type", "Type", 90),
            ("opening_debit", "Opening Dr", 120),
            ("opening_credit", "Opening Cr", 120),
            ("period_debit",  "Period Dr", 120),
            ("period_credit", "Period Cr", 120),
            ("closing_debit", "Closing Dr", 120),
            ("closing_credit", "Closing Cr", 120),
        ]
        ysb = ttk.Scrollbar(table_frame, orient='vertical')
        xsb = ttk.Scrollbar(table_frame, orient='horizontal')
        self.tree = ttk.Treeview(table_frame, columns=cols, show='headings',
                                 yscrollcommand=ysb.set, xscrollcommand=xsb.set)
        ysb.config(command=self.tree.yview); xsb.config(command=self.tree.xview)
        ysb.pack(side='right', fill='y'); xsb.pack(side='bottom', fill='x')
        self.tree.pack(fill='both', expand=True)
        for c, t, w in headers:
            self.tree.heading(c, text=t)
            anchor = 'e' if c in ("opening_debit","opening_credit","period_debit",
                                   "period_credit","closing_debit","closing_credit") else 'w'
            self.tree.column(c, width=w, anchor=anchor, stretch=True)
        # Zebra styling
        self.tree.tag_configure('odd', background='#FFFFFF')
        self.tree.tag_configure('even', background='#F7FAFC')
        self.tree.tag_configure('totals',
                                background='#EDF2F7',
                                font=('Helvetica', 10, 'bold'))

        # Generate automatically on open
        try: self.generate()
        except Exception: pass

    def preset_this_year(self):
        self.start_date.delete(0, 'end')
        self.start_date.insert(0, datetime.now().replace(month=1, day=1).strftime("%Y-%m-%d"))
        self.end_date.delete(0, 'end')
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.generate()

    def preset_all_time(self):
        self.start_date.delete(0, 'end')
        self.start_date.insert(0, "2000-01-01")
        self.end_date.delete(0, 'end')
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.generate()

    @staticmethod
    def _fmt(v):
        try:    fv = float(v or 0)
        except: return ""
        if fv == 0: return ""
        return f"{fv:,.2f}"

    def generate(self):
        try:
            s = self.start_date.get().strip()
            e = self.end_date.get().strip()
            import datetime as _dt
            s_d = _dt.date.fromisoformat(s); e_d = _dt.date.fromisoformat(e)
            if s_d > e_d: s_d, e_d = e_d, s_d
        except Exception as e:
            messagebox.showerror("Error", f"Invalid dates: {e}")
            return
        try:
            from report_engine import ReportEngine, DateRangeType
            eng = ReportEngine(self.manager, range_type=DateRangeType.CUSTOM,
                               start=s_d, end=e_d)
            tb = eng.get_trial_balance()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to build Trial Balance: {e}")
            import traceback; traceback.print_exc()
            return
        self.last_trial_balance = tb

        # Populate tree
        for i in self.tree.get_children():
            self.tree.delete(i)
        tot = tb.get('totals') or {}
        rows = tb.get('rows') or []
        for idx, r in enumerate(rows):
            tag = 'even' if idx % 2 == 0 else 'odd'
            self.tree.insert('', 'end', iid=str(idx), tags=(tag,), values=(
                r.get('account_code',''),
                r.get('account_name',''),
                r.get('account_type',''),
                self._fmt(r.get('opening_debit')), self._fmt(r.get('opening_credit')),
                self._fmt(r.get('period_debit')),  self._fmt(r.get('period_credit')),
                self._fmt(r.get('closing_debit')), self._fmt(r.get('closing_credit')),
            ))
        # Totals row
        def _ft(x):
            try: fv = float(x or 0); return f"{fv:,.2f}"
            except: return "0.00"
        self.tree.insert('', 'end', iid='__totals__', tags=('totals',), values=(
            "", "TOTALS", "",
            _ft(tot.get('opening_debit')), _ft(tot.get('opening_credit')),
            _ft(tot.get('period_debit')),  _ft(tot.get('period_credit')),
            _ft(tot.get('closing_debit')), _ft(tot.get('closing_credit')),
        ))

        # Summary + balance banner
        ok = bool(tb.get('balanced', False))
        if ok:
            self.balance_lbl.config(text="✔  In Balance", foreground="#276749")
        else:
            self.balance_lbl.config(text="⚠  OUT OF BALANCE", foreground="#9B2C2C")
        bal_amt = (self._safer(tot,'closing_debit') - self._safer(tot,'closing_credit'))
        self.summary_lbl.config(
            text=(f"Period: {tb.get('start_date')} → {tb.get('end_date')}   ·   "
                  f"Active Accounts: {tb.get('account_count', 0)}   ·   "
                  f"Closing Dr Total: AED {self._safer(tot,'closing_debit'):,.2f}   ·   "
                  f"Closing Cr Total: AED {self._safer(tot,'closing_credit'):,.2f}"
                  + (f"   ·   Net Difference: AED {bal_amt:,.2f}" if not ok else ""))
        )

    @staticmethod
    def _safer(d, k):
        try: return float(d.get(k, 0) or 0)
        except Exception: return 0.0

    def export_pdf(self):
        if not self.last_trial_balance:
            messagebox.showerror("Error", "Generate the Trial Balance first, then export.")
            return
        try:
            gen = getattr(self.manager, 'pdf_generator', None)
            if gen is None or not callable(getattr(gen, 'generate_trial_balance_pdf', None)):
                from hope_pharma_complete import EnhancedPDFGenerator
                gen = EnhancedPDFGenerator()
            path = gen.generate_trial_balance_pdf(self.last_trial_balance)
            messagebox.showinfo("PDF Exported",
                                f"Trial Balance PDF saved successfully:\n\n{path}")
            try:
                if sys.platform == 'darwin':
                    subprocess.run(['open', path])
                elif os.name == 'nt':
                    os.startfile(path)
                else:
                    subprocess.run(['xdg-open', path])
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Export Failed", f"Error generating PDF:\n{e}")
            import traceback; traceback.print_exc()


class ReportsDialog:
    def __init__(self, parent, manager):
        self.manager = manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Sales Reports")
        self.dialog.geometry("800x600")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        try:
            lm = LogoManager()
            if lm.load_logo(size=(64, 64)):
                self.dialog.iconphoto(True, lm.logo_photo)
        except Exception:
            pass
            
        self.last_report_data = None
        self.last_report_type = None
        
        self.setup_ui()
        _fit_window(self.dialog, 900, 640, mode="large", remember_key="sales_reports")
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="📊 Sales Reports", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        
        # Report controls
        controls_frame = ttk.LabelFrame(main_frame, text="Report Filters", padding="10")
        controls_frame.pack(fill='x', pady=(0, 20))
        
        # Date range
        date_frame = ttk.Frame(controls_frame)
        date_frame.pack(fill='x', pady=5)
        
        ttk.Label(date_frame, text="Start Date:").grid(row=0, column=0, padx=(0, 5))
        self.start_date = ttk.Entry(date_frame, width=12)
        self.start_date.grid(row=0, column=1, padx=(0, 15))
        self.start_date.insert(0, (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))
        
        ttk.Label(date_frame, text="End Date:").grid(row=0, column=2, padx=(0, 5))
        self.end_date = ttk.Entry(date_frame, width=12)
        self.end_date.grid(row=0, column=3, padx=(0, 15))
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        # Client filter
        client_frame = ttk.Frame(controls_frame)
        client_frame.pack(fill='x', pady=5)
        
        ttk.Label(client_frame, text="Client:").grid(row=0, column=0, padx=(0, 5))
        self.client_filter = ttk.Combobox(client_frame, width=30)
        self.client_filter.grid(row=0, column=1, padx=(0, 15))
        self.update_client_list()
        self.client_filter.set("All Clients")
        
        # Report buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.pack(fill='x', pady=10)
        
        ttk.Button(button_frame, text="Generate Sales Report", 
                  command=self.generate_sales_report).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Client Report", 
                  command=self.generate_client_report).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Client Summary", 
                  command=self.generate_client_summary_report).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Export Text", 
                  command=self.export_report).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Export CSV", 
                  command=self.export_csv, style='Accent.TButton').pack(side='left', padx=5)
        ttk.Button(button_frame, text="Export PDF", 
                  command=self.export_pdf).pack(side='left', padx=5)
        
        # Report results
        report_frame = ttk.LabelFrame(main_frame, text="Report Results", padding="10")
        report_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        self.report_text = scrolledtext.ScrolledText(report_frame, height=20)
        self.report_text.pack(fill='both', expand=True)
        
        # Close button
        ttk.Button(main_frame, text="Close", 
                  command=self.dialog.destroy).pack(pady=10)
    
    def update_client_list(self):
        """Update the client list in reports"""
        clients = self.manager.get_clients()
        values = ['All Clients']
        for client in clients:
            name = str(client or '').strip()
            if name and name not in values:
                values.append(name)
        self.client_filter['values'] = values

    def _collect_filtered_invoices(self):
        start_date = self.start_date.get().strip()
        end_date = self.end_date.get().strip()
        selected_client = self.client_filter.get().strip()
        if selected_client.lower() == 'all clients':
            selected_client = ''

        invoices = []
        try:
            invoices = self.manager.get_all_invoices_dict() or []
        except Exception:
            invoices = []

        filtered = []
        for invoice in invoices:
            if not isinstance(invoice, dict):
                continue
            invoice_date = str(invoice.get('date') or '').strip()
            if start_date and invoice_date and invoice_date < start_date:
                continue
            if end_date and invoice_date and invoice_date > end_date:
                continue
            client_name = str(invoice.get('client_name') or '').strip() or 'Unknown Client'
            if selected_client and client_name.lower() != selected_client.lower():
                continue
            filtered.append(invoice)
        return filtered

    def _get_reports_logo_path(self):
        try:
            data_folder = getattr(self.manager, 'data_folder', None) or getattr(self.manager, 'invoice_folder', None)
            logo_manager = LogoManager(data_folder)
            if logo_manager.logo_path and os.path.exists(logo_manager.logo_path):
                return logo_manager.logo_path
        except Exception:
            pass
        return None

    def generate_client_summary_report(self):
        invoices = self._collect_filtered_invoices()
        if not invoices:
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', "No invoices found for the selected filters.")
            self.last_report_data = None
            return

        summary_map = {}
        for invoice in invoices:
            client_name = str(invoice.get('client_name') or '').strip() or 'Unknown Client'
            summary = summary_map.setdefault(client_name, {
                'client_name': client_name,
                'invoice_count': 0,
                'total_sales': 0.0,
                'total_tax': 0.0,
                'total_paid': 0.0,
                'total_cost': 0.0,
                'total_profit': 0.0,
                'outstanding_balance': 0.0,
                'invoices': [],
            })

            grand_total = float(invoice.get('grand_total', 0) or 0)
            total_paid = float(invoice.get('total_paid', 0) or 0)
            tax_amount = float(invoice.get('tax_amount', 0) or 0)
            total_cost = float(invoice.get('total_cost', 0) or 0)
            profit_loss = float(invoice.get('profit_loss', grand_total - total_cost) or 0)
            balance = max(grand_total - total_paid, 0.0)

            invoice_row = {
                'invoice_id': invoice.get('invoice_id', ''),
                'date': invoice.get('date', ''),
                'grand_total': grand_total,
                'total_paid': total_paid,
                'tax_amount': tax_amount,
                'total_cost': total_cost,
                'profit_loss': profit_loss,
                'outstanding_balance': balance,
                'status': invoice.get('status', ''),
            }
            summary['invoices'].append(invoice_row)
            summary['invoice_count'] += 1
            summary['total_sales'] += grand_total
            summary['total_tax'] += tax_amount
            summary['total_paid'] += total_paid
            summary['total_cost'] += total_cost
            summary['total_profit'] += profit_loss
            summary['outstanding_balance'] += balance

        client_summaries = []
        for summary in summary_map.values():
            summary['invoices'].sort(key=lambda row: (row.get('date', ''), row.get('invoice_id', '')))
            for key in ('total_sales', 'total_tax', 'total_paid', 'total_cost', 'total_profit', 'outstanding_balance'):
                summary[key] = round(summary[key], 2)
            client_summaries.append(summary)
        client_summaries.sort(key=lambda row: row.get('client_name', '').lower())

        report = {
            'start_date': self.start_date.get().strip(),
            'end_date': self.end_date.get().strip(),
            'client_filter': self.client_filter.get().strip() or 'All Clients',
            'client_summaries': client_summaries,
            'total_clients': len(client_summaries),
            'total_invoices': sum(item.get('invoice_count', 0) for item in client_summaries),
            'grand_total_sales': round(sum(item.get('total_sales', 0.0) for item in client_summaries), 2),
            'grand_total_paid': round(sum(item.get('total_paid', 0.0) for item in client_summaries), 2),
            'grand_outstanding': round(sum(item.get('outstanding_balance', 0.0) for item in client_summaries), 2),
        }
        self.last_report_data = report
        self.last_report_type = 'client_summary'

        lines = []
        lines.append("CLIENT SUMMARY REPORT")
        lines.append(f"Period: {report['start_date']} to {report['end_date']}")
        lines.append(f"Client Filter: {report['client_filter']}")
        lines.append("=" * 72)
        lines.append("")
        for summary in client_summaries:
            lines.append(f"Client: {summary['client_name']}")
            lines.append(f"Number of invoices: {summary['invoice_count']}")
            for index, invoice_row in enumerate(summary.get('invoices', []), start=1):
                lines.append(
                    f"Invoice {index}: {invoice_row.get('invoice_id', '')} | "
                    f"Date: {invoice_row.get('date', '')} | "
                    f"Cost: AED {invoice_row.get('grand_total', 0.0):.2f}"
                )
            lines.append(f"Total for this client: AED {summary['total_sales']:.2f}")
            lines.append(f"Outstanding for this client: AED {summary['outstanding_balance']:.2f}")
            lines.append("-" * 72)
        lines.append(f"Clients: {report['total_clients']}")
        lines.append(f"Invoices: {report['total_invoices']}")
        lines.append(f"Grand Total Sales: AED {report['grand_total_sales']:.2f}")
        lines.append(f"Grand Total Paid: AED {report['grand_total_paid']:.2f}")
        lines.append(f"Grand Outstanding: AED {report['grand_outstanding']:.2f}")
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert('1.0', "\n".join(lines))
    
    def generate_sales_report(self):
        """Generate sales report for date range"""
        try:
            report = self.manager.get_sales_report(
                self.start_date.get(),
                self.end_date.get()
            )
            
            if not report or report['total_invoices'] == 0:
                self.report_text.delete('1.0', tk.END)
                self.report_text.insert('1.0', "No data found for the selected date range.")
                self.last_report_data = None
                return
            
            self.last_report_data = report
            self.last_report_type = 'sales'
            
            report_text = f"""
SALES REPORT
Period: {report['start_date']} to {report['end_date']}
{'='*60}

SUMMARY:
Total Invoices: {report['total_invoices']}
Total Sales: AED {report['total_sales']:.2f}
Total VAT: AED {report['total_tax']:.2f}
Total Paid: AED {report['total_paid']:.2f}
Total Cost: AED {report['total_cost']:.2f}
Total Profit: AED {report['total_profit']:.2f}
Outstanding Balance: AED {report['outstanding_balance']:.2f}

DETAILED INVOICES:
{'='*60}
"""
            for invoice in report['invoices']:
                balance = invoice.get('grand_total', 0) - invoice.get('total_paid', 0)
                report_text += f"""
Invoice: {invoice.get('invoice_id', '')}
Client: {invoice.get('client_name', '')}
Date: {invoice.get('date', '')} | Due: {invoice.get('due_date', '')}
Total: AED {invoice.get('grand_total', 0):.2f} | Paid: AED {invoice.get('total_paid', 0):.2f} | Balance: AED {balance:.2f}
Cost: AED {invoice.get('total_cost', 0):.2f} | P/L: AED {invoice.get('profit_loss', 0):.2f}
Status: {invoice.get('status', 'Unknown')}
{'-'*50}
"""
            
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', report_text)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate report: {e}")
    
    def generate_client_report(self):
        """Generate client-specific report"""
        client = self.client_filter.get()
        if not client or client == "All Clients":
            messagebox.showwarning("Warning", "Please select a client")
            return
            
        try:
            report = self.manager.get_client_report(
                client,
                self.start_date.get(),
                self.end_date.get()
            )
            
            if not report or report['total_invoices'] == 0:
                self.report_text.delete('1.0', tk.END)
                self.report_text.insert('1.0', f"No invoices found for {client} in the selected date range.")
                self.last_report_data = None
                return
            
            self.last_report_data = report
            self.last_report_type = 'client'
            
            report_text = f"""
CLIENT REPORT: {report['client_name']}
Period: {report['start_date']} to {report['end_date']}
{'='*60}

SUMMARY:
Total Invoices: {report['total_invoices']}
Total Sales: AED {report['total_sales']:.2f}
Total VAT: AED {report['total_tax']:.2f}
Total Paid: AED {report['total_paid']:.2f}
Total Cost: AED {report['total_cost']:.2f}
Total Profit: AED {report['total_profit']:.2f}
Outstanding Balance: AED {report['outstanding_balance']:.2f}

INVOICE DETAILS:
{'='*60}
"""
            for invoice in report['invoices']:
                balance = invoice.get('grand_total', 0) - invoice.get('total_paid', 0)
                report_text += f"""
Invoice: {invoice.get('invoice_id', '')}
Date: {invoice.get('date', '')} | Due: {invoice.get('due_date', '')}
Total: AED {invoice.get('grand_total', 0):.2f} | Paid: AED {invoice.get('total_paid', 0):.2f} | Balance: AED {balance:.2f}
Cost: AED {invoice.get('total_cost', 0):.2f} | P/L: AED {invoice.get('profit_loss', 0):.2f}
Status: {invoice.get('status', 'Unknown')}
{'-'*50}
"""
            
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', report_text)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate client report: {e}")
            
    def export_csv(self):
        """Export current report data to CSV"""
        if not self.last_report_data:
            messagebox.showwarning("Warning", "Please generate a report first.")
            return
            
        try:
            # Generate default filename
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            if self.last_report_type == 'sales':
                prefix = "Sales_Report"
            elif self.last_report_type == 'client_summary':
                prefix = "Client_Summary_Report"
            else:
                prefix = f"Client_Report_{self.last_report_data.get('client_name', 'Client')}"
            default_filename = f"{prefix}_{timestamp}.csv"
            
            filename = filedialog.asksaveasfilename(
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
                initialfile=default_filename,
                title="Export Data to CSV"
            )
            
            if not filename:
                return
                
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                if self.last_report_type == 'client_summary':
                    writer.writerow(['CLIENT SUMMARY REPORT'])
                    writer.writerow(['Period Start', self.last_report_data.get('start_date', '')])
                    writer.writerow(['Period End', self.last_report_data.get('end_date', '')])
                    writer.writerow(['Client Filter', self.last_report_data.get('client_filter', 'All Clients')])
                    writer.writerow(['Total Clients', self.last_report_data.get('total_clients', 0)])
                    writer.writerow(['Total Invoices', self.last_report_data.get('total_invoices', 0)])
                    writer.writerow(['Grand Total Sales', self.last_report_data.get('grand_total_sales', 0)])
                    writer.writerow(['Grand Total Paid', self.last_report_data.get('grand_total_paid', 0)])
                    writer.writerow(['Grand Outstanding', self.last_report_data.get('grand_outstanding', 0)])
                    writer.writerow([])
                    writer.writerow(['CLIENT TOTALS'])
                    writer.writerow(['Client', 'Invoices', 'Total Sales', 'Total VAT', 'Total Paid', 'Total Cost', 'Total Profit', 'Outstanding'])
                    for summary in self.last_report_data.get('client_summaries', []):
                        writer.writerow([
                            summary.get('client_name', ''),
                            summary.get('invoice_count', 0),
                            f"{summary.get('total_sales', 0):.2f}",
                            f"{summary.get('total_tax', 0):.2f}",
                            f"{summary.get('total_paid', 0):.2f}",
                            f"{summary.get('total_cost', 0):.2f}",
                            f"{summary.get('total_profit', 0):.2f}",
                            f"{summary.get('outstanding_balance', 0):.2f}",
                        ])
                    writer.writerow([])
                    writer.writerow(['INVOICE BREAKDOWN'])
                    writer.writerow(['Client', 'Invoice ID', 'Date', 'Grand Total', 'Paid', 'Outstanding', 'Cost', 'Profit/Loss', 'Status'])
                    for summary in self.last_report_data.get('client_summaries', []):
                        for inv in summary.get('invoices', []):
                            writer.writerow([
                                summary.get('client_name', ''),
                                inv.get('invoice_id', ''),
                                inv.get('date', ''),
                                f"{inv.get('grand_total', 0):.2f}",
                                f"{inv.get('total_paid', 0):.2f}",
                                f"{inv.get('outstanding_balance', 0):.2f}",
                                f"{inv.get('total_cost', 0):.2f}",
                                f"{inv.get('profit_loss', 0):.2f}",
                                inv.get('status', ''),
                            ])
                else:
                    # Write Summary Header
                    writer.writerow(['REPORT SUMMARY'])
                    if self.last_report_type == 'client':
                        writer.writerow(['Client Name', self.last_report_data.get('client_name', '')])
                    writer.writerow(['Period Start', self.last_report_data.get('start_date', '')])
                    writer.writerow(['Period End', self.last_report_data.get('end_date', '')])
                    writer.writerow(['Total Invoices', self.last_report_data.get('total_invoices', 0)])
                    writer.writerow(['Total Sales', self.last_report_data.get('total_sales', 0)])
                    writer.writerow(['Total VAT', self.last_report_data.get('total_tax', 0)])
                    writer.writerow(['Total Paid', self.last_report_data.get('total_paid', 0)])
                    writer.writerow(['Total Cost', self.last_report_data.get('total_cost', 0)])
                    writer.writerow(['Total Profit', self.last_report_data.get('total_profit', 0)])
                    writer.writerow(['Outstanding', self.last_report_data.get('outstanding_balance', 0)])
                    writer.writerow([])
                    
                    # Write Invoices Header
                    writer.writerow(['INVOICE DETAILS'])
                    headers = ['Invoice ID', 'Date', 'Due Date', 'Client', 'Grand Total', 'Total Paid', 'Balance', 'Total Cost', 'Profit/Loss', 'Status']
                    writer.writerow(headers)
                    
                    # Write Invoices Data
                    for inv in self.last_report_data.get('invoices', []):
                        balance = inv.get('grand_total', 0) - inv.get('total_paid', 0)
                        row = [
                            inv.get('invoice_id', ''),
                            inv.get('date', ''),
                            inv.get('due_date', ''),
                            inv.get('client_name', ''),
                            f"{inv.get('grand_total', 0):.2f}",
                            f"{inv.get('total_paid', 0):.2f}",
                            f"{balance:.2f}",
                            f"{inv.get('total_cost', 0):.2f}",
                            f"{inv.get('profit_loss', 0):.2f}",
                            inv.get('status', '')
                        ]
                        writer.writerow(row)
            
            messagebox.showinfo("Success", f"Data exported successfully to:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export CSV: {e}")
    
    def export_report(self):
        """Export report to file"""
        try:
            content = self.report_text.get('1.0', tk.END)
            if not content.strip():
                messagebox.showwarning("Warning", "No report content to export")
                return
                
            # Save report to HopePharmaInvoice folder
            filename = os.path.join(self.manager.invoice_folder, f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")
            with open(filename, 'w') as f:
                f.write(content)
                
            messagebox.showinfo("Success", f"Report exported to {filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report: {e}")

    def export_pdf(self):
        if not self.last_report_data:
            messagebox.showwarning("Warning", "Please generate a report first.")
            return
        if not REPORTLAB_PDF_AVAILABLE:
            messagebox.showwarning("Warning", "PDF export is not available in this build.")
            return

        try:
            reports_folder = os.path.join(self.manager.invoice_folder, "FinancialReports")
            os.makedirs(reports_folder, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            if self.last_report_type == 'sales':
                pdf_path = PDFGenerator(output_folder=reports_folder).generate_sales_report_pdf(self.last_report_data)
                if pdf_path:
                    webbrowser.open('file://' + os.path.abspath(pdf_path))
                    messagebox.showinfo("Success", f"PDF exported to:\n{pdf_path}")
                return

            if self.last_report_type == 'client_summary':
                filename = os.path.join(reports_folder, f"Client_Summary_Report_{timestamp}.pdf")
                title_text = "Client Summary Report"
            else:
                filename = os.path.join(reports_folder, f"Client_Report_{self.last_report_data.get('client_name', 'Client')}_{timestamp}.pdf")
                title_text = f"Client Report: {self.last_report_data.get('client_name', 'Client')}"

            doc = SimpleDocTemplate(filename, pagesize=landscape(A4), leftMargin=0.35 * inch, rightMargin=0.35 * inch, topMargin=0.4 * inch, bottomMargin=0.4 * inch)
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle('ClientReportTitle', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor("#2c3e50"))
            subtitle_style = ParagraphStyle('ClientReportSub', parent=styles['Normal'], fontSize=9, alignment=1, textColor=colors.grey)
            body_style = ParagraphStyle('ClientReportBody', parent=styles['BodyText'], fontSize=8, leading=10)
            elements = []

            logo_path = self._get_reports_logo_path()
            if logo_path:
                try:
                    elements.append(RLImage(logo_path, width=1.0 * inch, height=1.0 * inch))
                    elements.append(Spacer(1, 0.08 * inch))
                except Exception:
                    pass

            elements.append(Paragraph(title_text, title_style))
            elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_style))
            elements.append(Paragraph(f"Period: {self.last_report_data.get('start_date', '')} to {self.last_report_data.get('end_date', '')}", subtitle_style))
            elements.append(Spacer(1, 10))

            if self.last_report_type == 'client_summary':
                elements.append(Paragraph(f"Client Filter: {self.last_report_data.get('client_filter', 'All Clients')}", body_style))
                elements.append(Spacer(1, 8))
                summary_data = [["Client", "Invoices", "Total Sales", "Total Paid", "Outstanding"]]
                for summary in self.last_report_data.get('client_summaries', []):
                    summary_data.append([
                        summary.get('client_name', ''),
                        str(summary.get('invoice_count', 0)),
                        f"AED {summary.get('total_sales', 0.0):.2f}",
                        f"AED {summary.get('total_paid', 0.0):.2f}",
                        f"AED {summary.get('outstanding_balance', 0.0):.2f}",
                    ])
                table = Table(summary_data, colWidths=[3.2 * inch, 0.9 * inch, 1.5 * inch, 1.5 * inch, 1.5 * inch])
                table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b5394')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ]))
                elements.append(table)
                elements.append(Spacer(1, 10))
                elements.append(Paragraph(f"Clients: {self.last_report_data.get('total_clients', 0)}", body_style))
                elements.append(Paragraph(f"Invoices: {self.last_report_data.get('total_invoices', 0)}", body_style))
                elements.append(Paragraph(f"Grand Total Sales: AED {self.last_report_data.get('grand_total_sales', 0.0):.2f}", body_style))
                elements.append(Paragraph(f"Grand Total Paid: AED {self.last_report_data.get('grand_total_paid', 0.0):.2f}", body_style))
                elements.append(Paragraph(f"Grand Outstanding: AED {self.last_report_data.get('grand_outstanding', 0.0):.2f}", body_style))
            else:
                summary_data = [
                    ['Client', self.last_report_data.get('client_name', '')],
                    ['Total Invoices', str(self.last_report_data.get('total_invoices', 0))],
                    ['Total Sales', f"AED {self.last_report_data.get('total_sales', 0.0):.2f}"],
                    ['Total Paid', f"AED {self.last_report_data.get('total_paid', 0.0):.2f}"],
                    ['Outstanding', f"AED {self.last_report_data.get('outstanding_balance', 0.0):.2f}"],
                ]
                summary_table = Table(summary_data, colWidths=[2.0 * inch, 2.5 * inch])
                summary_table.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                    ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#eaf2f8')),
                    ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                ]))
                elements.append(summary_table)
                elements.append(Spacer(1, 10))

                invoice_data = [["Invoice ID", "Date", "Due Date", "Total", "Paid", "Balance", "Status"]]
                for inv in self.last_report_data.get('invoices', []):
                    balance = float(inv.get('grand_total', 0) or 0) - float(inv.get('total_paid', 0) or 0)
                    invoice_data.append([
                        inv.get('invoice_id', ''),
                        inv.get('date', ''),
                        inv.get('due_date', ''),
                        f"AED {float(inv.get('grand_total', 0) or 0):.2f}",
                        f"AED {float(inv.get('total_paid', 0) or 0):.2f}",
                        f"AED {balance:.2f}",
                        inv.get('status', ''),
                    ])
                details = Table(invoice_data, colWidths=[1.4 * inch, 1.0 * inch, 1.0 * inch, 1.2 * inch, 1.2 * inch, 1.2 * inch, 1.2 * inch])
                details.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b5394')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 9),
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ]))
                elements.append(details)

            doc.build(elements)
            webbrowser.open('file://' + os.path.abspath(filename))
            messagebox.showinfo("Success", f"PDF exported to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")

class SalesServicesReportsDialog:
    def __init__(self, parent, manager):
        self.manager = manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Sales and Services Reports")
        self.dialog.geometry("980x680")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        try:
            lm = LogoManager()
            if lm.load_logo(size=(64, 64)):
                self.dialog.iconphoto(True, lm.logo_photo)
        except Exception:
            pass

        self.current_mode = 'all'
        self.invoice_item_map = {}
        self.summary_var = tk.StringVar(value="")
        self.current_grouped_report = []

        self.setup_ui()
        self.load_report('all')
        _fit_window(self.dialog, 980, 680, mode="workspace", remember_key="sales_services_reports")

    def setup_ui(self):
        main_frame = ttk.Frame(self.dialog, padding="12")
        main_frame.pack(fill='both', expand=True)

        ttk.Label(
            main_frame,
            text="Sales and Services Reports",
            font=('Helvetica', 16, 'bold')
        ).pack(anchor='w', pady=(0, 10))

        ttk.Label(
            main_frame,
            text="Clients are grouped first. Double-click any invoice row to open its details.",
            font=('Helvetica', 10)
        ).pack(anchor='w', pady=(0, 10))

        filters_frame = ttk.LabelFrame(main_frame, text="Filters", padding="10")
        filters_frame.pack(fill='x', pady=(0, 10))

        ttk.Label(filters_frame, text="Start Date:").grid(row=0, column=0, padx=(0, 5), pady=2, sticky='w')
        self.start_date = ttk.Entry(filters_frame, width=12)
        self.start_date.grid(row=0, column=1, padx=(0, 15), pady=2, sticky='w')
        self.start_date.insert(0, (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))

        ttk.Label(filters_frame, text="End Date:").grid(row=0, column=2, padx=(0, 5), pady=2, sticky='w')
        self.end_date = ttk.Entry(filters_frame, width=12)
        self.end_date.grid(row=0, column=3, padx=(0, 15), pady=2, sticky='w')
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))

        mode_frame = ttk.Frame(filters_frame)
        mode_frame.grid(row=1, column=0, columnspan=4, sticky='w', pady=(10, 0))
        ttk.Button(mode_frame, text="All Invoices", command=lambda: self.load_report('all')).pack(side='left', padx=(0, 6))
        ttk.Button(mode_frame, text="Sales Only", command=lambda: self.load_report('sales')).pack(side='left', padx=6)
        ttk.Button(mode_frame, text="Services Only", command=lambda: self.load_report('service')).pack(side='left', padx=6)
        ttk.Button(mode_frame, text="Export Outbound PDF", command=self.export_outbound_pdf).pack(side='left', padx=(18, 6))
        ttk.Button(mode_frame, text="Expand All", command=self.expand_all).pack(side='left', padx=6)
        ttk.Button(mode_frame, text="Collapse All", command=self.collapse_all).pack(side='left', padx=6)

        ttk.Label(main_frame, textvariable=self.summary_var, font=('Helvetica', 10, 'bold')).pack(anchor='w', pady=(0, 8))

        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill='both', expand=True)

        columns = ('Category', 'Location', 'Invoices', 'Date', 'Total', 'Paid', 'Balance', 'Status')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='tree headings', height=22)
        self.tree.heading('#0', text='Client / Invoice')
        self.tree.heading('Category', text='Category')
        self.tree.heading('Location', text='Location')
        self.tree.heading('Invoices', text='Invoices')
        self.tree.heading('Date', text='Date')
        self.tree.heading('Total', text='Total')
        self.tree.heading('Paid', text='Paid')
        self.tree.heading('Balance', text='Balance')
        self.tree.heading('Status', text='Status')

        self.tree.column('#0', width=260, stretch=True)
        self.tree.column('Category', width=90, anchor='center')
        self.tree.column('Location', width=140, anchor='w')
        self.tree.column('Invoices', width=80, anchor='center')
        self.tree.column('Date', width=100, anchor='center')
        self.tree.column('Total', width=110, anchor='e')
        self.tree.column('Paid', width=110, anchor='e')
        self.tree.column('Balance', width=110, anchor='e')
        self.tree.column('Status', width=120, anchor='center')

        y_scroll = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        x_scroll = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.tree.pack(side='left', fill='both', expand=True)
        y_scroll.pack(side='right', fill='y')
        x_scroll.pack(side='bottom', fill='x')
        self.tree.bind('<Double-1>', self.on_tree_double_click)

        try:
            self.tree.tag_configure('client', font=('Helvetica', 10, 'bold'))
        except Exception:
            pass

        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x', pady=(10, 0))
        ttk.Button(buttons_frame, text="Refresh", command=lambda: self.load_report(self.current_mode)).pack(side='left')
        ttk.Button(buttons_frame, text="Close", command=self.dialog.destroy).pack(side='right')

    def _validate_date_filters(self):
        start_date = self.start_date.get().strip()
        end_date = self.end_date.get().strip()
        try:
            if start_date:
                datetime.strptime(start_date, "%Y-%m-%d")
            if end_date:
                datetime.strptime(end_date, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid Date", "Please use the date format YYYY-MM-DD.")
            return None, None
        if start_date and end_date and start_date > end_date:
            messagebox.showwarning("Invalid Date", "Start date cannot be after end date.")
            return None, None
        return start_date, end_date

    def _normalize_invoice_type(self, invoice):
        invoice_type = str(invoice.get('invoice_type', 'sales') or 'sales').strip().lower()
        if invoice_type not in ('sales', 'service'):
            return 'sales'
        return invoice_type

    def _mode_label(self, mode=None):
        mode = mode or self.current_mode
        return {
            'all': 'All Invoices',
            'sales': 'Sales',
            'service': 'Service',
        }.get(mode, 'All Invoices')

    def _get_logo_path(self):
        try:
            data_folder = getattr(self.manager, 'data_folder', None) or getattr(self.manager, 'invoice_folder', None)
            logo_manager = LogoManager(data_folder)
            if logo_manager.logo_path and os.path.exists(logo_manager.logo_path):
                return logo_manager.logo_path
        except Exception:
            pass
        return None

    def _status_for_invoice(self, invoice):
        total = float(invoice.get('grand_total', 0) or 0)
        paid = float(invoice.get('total_paid', 0) or 0)
        balance = max(total - paid, 0.0)
        if paid >= total:
            status = "Paid"
        elif paid > 0:
            status = "Partial Paid"
        else:
            status = "Not Paid"
        return status, balance

    def _collect_invoices(self, mode):
        start_date, end_date = self._validate_date_filters()
        if start_date is None and end_date is None:
            return None

        try:
            invoices = self.manager.get_all_invoices_dict() or []
        except Exception:
            invoices = []

        filtered = []
        for invoice in invoices:
            if not isinstance(invoice, dict):
                continue
            invoice_type = self._normalize_invoice_type(invoice)
            if mode in ('sales', 'service') and invoice_type != mode:
                continue
            invoice_date = str(invoice.get('date') or '').strip()
            if start_date and invoice_date and invoice_date < start_date:
                continue
            if end_date and invoice_date and invoice_date > end_date:
                continue
            filtered.append(invoice)
        return filtered

    def load_report(self, mode):
        invoices = self._collect_invoices(mode)
        if invoices is None:
            return

        self.current_mode = mode
        self.invoice_item_map = {}
        self.current_grouped_report = []
        for item_id in self.tree.get_children():
            self.tree.delete(item_id)

        report_title = {
            'all': 'All Invoices',
            'sales': 'Sales Only',
            'service': 'Services Only',
        }.get(mode, 'All Invoices')

        grouped = {}
        total_sales = 0.0
        total_paid = 0.0
        for invoice in invoices:
            client_name = str(invoice.get('client_name') or '').strip() or 'Unknown Client'
            grouped.setdefault(client_name, []).append(invoice)
            total_sales += float(invoice.get('grand_total', 0) or 0)
            total_paid += float(invoice.get('total_paid', 0) or 0)

        if not grouped:
            self.summary_var.set(
                f"{report_title}: no invoices found from {self.start_date.get().strip()} to {self.end_date.get().strip()}."
            )
            return

        total_balance = max(total_sales - total_paid, 0.0)
        self.summary_var.set(
            f"{report_title} | Clients: {len(grouped)} | Invoices: {len(invoices)} | "
            f"Total: AED {total_sales:.2f} | Outstanding: AED {total_balance:.2f}"
        )

        for client_name in sorted(grouped, key=lambda value: value.lower()):
            client_invoices = sorted(
                grouped[client_name],
                key=lambda row: (str(row.get('date') or ''), str(row.get('invoice_id') or ''))
            )
            client_total = 0.0
            client_paid = 0.0
            invoice_rows = []
            client_location = ''
            for invoice in client_invoices:
                if not client_location:
                    loc_value = str(
                        invoice.get('client_location')
                        or invoice.get('client_emirate')
                        or ''
                    ).strip()
                    if loc_value:
                        client_location = loc_value
                invoice_total = float(invoice.get('grand_total', 0) or 0)
                invoice_paid = float(invoice.get('total_paid', 0) or 0)
                client_total += invoice_total
                client_paid += invoice_paid
                invoice_rows.append({
                    'invoice_id': str(invoice.get('invoice_id') or ''),
                    'date': str(invoice.get('date') or ''),
                    'total': round(invoice_total, 2),
                    'paid': round(invoice_paid, 2),
                })
            client_balance = max(client_total - client_paid, 0.0)
            self.current_grouped_report.append({
                'client_name': client_name,
                'client_location': client_location,
                'invoice_count': len(client_invoices),
                'total_amount': round(client_total, 2),
                'total_paid': round(client_paid, 2),
                'total_balance': round(client_balance, 2),
                'invoice_ids': [row['invoice_id'] for row in invoice_rows if row['invoice_id']],
                'invoices': invoice_rows,
            })
            parent_id = self.tree.insert(
                '',
                'end',
                text=client_name,
                values=(
                    'Client',
                    client_location,
                    len(client_invoices),
                    '',
                    f"AED {client_total:.2f}",
                    f"AED {client_paid:.2f}",
                    f"AED {client_balance:.2f}",
                    ''
                ),
                open=True,
                tags=('client',)
            )

            for invoice in client_invoices:
                invoice_type = self._normalize_invoice_type(invoice)
                type_label = 'Sales' if invoice_type == 'sales' else 'Service'
                status, balance = self._status_for_invoice(invoice)
                child_id = self.tree.insert(
                    parent_id,
                    'end',
                    text=str(invoice.get('invoice_id') or ''),
                    values=(
                        type_label,
                        '',
                        '',
                        str(invoice.get('date') or ''),
                        f"AED {float(invoice.get('grand_total', 0) or 0):.2f}",
                        f"AED {float(invoice.get('total_paid', 0) or 0):.2f}",
                        f"AED {balance:.2f}",
                        status
                    )
                )
                self.invoice_item_map[child_id] = str(invoice.get('invoice_id') or '')

    def export_outbound_pdf(self):
        if not REPORTLAB_PDF_AVAILABLE:
            messagebox.showwarning("Warning", "PDF export is not available in this build.")
            return

        if not self.current_grouped_report:
            messagebox.showwarning("Warning", "No outbound report data to export.")
            return

        try:
            reports_folder = os.path.join(self.manager.invoice_folder, "FinancialReports")
            os.makedirs(reports_folder, exist_ok=True)
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            mode_slug = self.current_mode or 'all'
            filename = os.path.join(reports_folder, f"Outbound_Report_{mode_slug}_{timestamp}.pdf")

            doc = SimpleDocTemplate(
                filename,
                pagesize=landscape(A4),
                leftMargin=0.35 * inch,
                rightMargin=0.35 * inch,
                topMargin=0.4 * inch,
                bottomMargin=0.4 * inch
            )
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'OutboundTitle',
                parent=styles['Heading1'],
                fontSize=16,
                alignment=1,
                textColor=colors.HexColor("#2c3e50")
            )
            subtitle_style = ParagraphStyle(
                'OutboundSub',
                parent=styles['Normal'],
                fontSize=9,
                alignment=1,
                textColor=colors.grey
            )
            section_style = ParagraphStyle(
                'OutboundSection',
                parent=styles['Heading3'],
                fontSize=11,
                textColor=colors.HexColor("#0b5394"),
                spaceAfter=4
            )
            body_style = ParagraphStyle(
                'OutboundBody',
                parent=styles['BodyText'],
                fontSize=8.5,
                leading=10
            )

            elements = []
            logo_path = self._get_logo_path()
            if logo_path:
                try:
                    elements.append(RLImage(logo_path, width=1.0 * inch, height=1.0 * inch))
                    elements.append(Spacer(1, 0.08 * inch))
                except Exception:
                    pass

            mode_label = self._mode_label()
            elements.append(Paragraph(f"Outbound Report - {mode_label}", title_style))
            elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_style))
            elements.append(Paragraph(
                f"Period: {self.start_date.get().strip()} to {self.end_date.get().strip()}",
                subtitle_style
            ))
            elements.append(Spacer(1, 10))

            total_clients = len(self.current_grouped_report)
            total_invoices = sum(item.get('invoice_count', 0) for item in self.current_grouped_report)
            grand_total = sum(float(item.get('total_amount', 0) or 0) for item in self.current_grouped_report)
            elements.append(Paragraph(
                f"Clients: {total_clients} | Invoices: {total_invoices} | Total Amount: AED {grand_total:.2f}",
                body_style
            ))
            elements.append(Spacer(1, 8))

            for client in self.current_grouped_report:
                client_name = client.get('client_name', 'Unknown Client')
                client_location = str(client.get('client_location') or '').strip() or '-'
                invoice_ids = client.get('invoice_ids', [])
                invoice_ids_text = ", ".join(invoice_ids) if invoice_ids else "No invoice IDs"
                elements.append(Paragraph(client_name, section_style))

                summary_table = Table([
                    ["Location", client_location],
                    ["Number of Invoices", str(client.get('invoice_count', 0))],
                    ["Total Amount", f"AED {float(client.get('total_amount', 0) or 0):.2f}"],
                    ["Total Paid", f"AED {float(client.get('total_paid', 0) or 0):.2f}"],
                    ["Outstanding", f"AED {float(client.get('total_balance', 0) or 0):.2f}"],
                    ["Invoice IDs", invoice_ids_text],
                ], colWidths=[1.8 * inch, 7.8 * inch])
                summary_table.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                    ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#eaf2f8')),
                    ('FONTNAME', (0, 0), (0, -1), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 8),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                    ('WORDWRAP', (0, 0), (-1, -1), 'CJK'),
                ]))
                elements.append(summary_table)
                elements.append(Spacer(1, 6))

                invoice_rows = [["Invoice ID", "Date", "Amount", "Paid"]]
                for invoice in client.get('invoices', []):
                    invoice_rows.append([
                        invoice.get('invoice_id', ''),
                        invoice.get('date', ''),
                        f"AED {float(invoice.get('total', 0) or 0):.2f}",
                        f"AED {float(invoice.get('paid', 0) or 0):.2f}",
                    ])
                invoice_table = Table(invoice_rows, colWidths=[2.0 * inch, 1.5 * inch, 1.5 * inch, 1.5 * inch])
                invoice_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b5394')),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 8.5),
                    ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                    ('FONTSIZE', (0, 1), (-1, -1), 8),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ]))
                elements.append(invoice_table)
                elements.append(Spacer(1, 12))

            doc.build(elements)
            webbrowser.open('file://' + os.path.abspath(filename))
            messagebox.showinfo("Success", f"Outbound report exported to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export outbound report PDF: {e}")

    def expand_all(self):
        for item_id in self.tree.get_children():
            self.tree.item(item_id, open=True)

    def collapse_all(self):
        for item_id in self.tree.get_children():
            self.tree.item(item_id, open=False)

    def on_tree_double_click(self, event=None):
        selection = self.tree.selection()
        if not selection:
            return
        item_id = selection[0]
        invoice_id = self.invoice_item_map.get(item_id)
        if not invoice_id:
            is_open = bool(self.tree.item(item_id, 'open'))
            self.tree.item(item_id, open=not is_open)
            return

        invoice_obj = self.manager.get_invoice(invoice_id)
        if not invoice_obj:
            messagebox.showwarning("Invoice Not Found", f"Could not find invoice {invoice_id}.")
            return
        invoice_data = invoice_obj.to_dict() if hasattr(invoice_obj, 'to_dict') else invoice_obj
        ViewInvoiceDialog(self.dialog, self.manager, invoice_id, invoice_data)

class CreatePurchaseDialog:
    def __init__(self, parent, manager):
        self.manager = manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Add New Purchase")
        self.dialog.geometry("620x680")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="Add New Purchase", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Purchase Details
        details_frame = ttk.LabelFrame(main_frame, text="Purchase Details", padding="10")
        details_frame.pack(fill='x', pady=5)
        
        # Use grid for the form fields
        ttk.Label(details_frame, text="Description *:").grid(row=0, column=0, sticky='w', pady=5)
        self.description = ttk.Entry(details_frame, width=40)
        self.description.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Amount (AED) *:").grid(row=1, column=0, sticky='w', pady=5)
        self.amount = ttk.Entry(details_frame, width=20)
        self.amount.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Date:").grid(row=2, column=0, sticky='w', pady=5)
        self.date = ttk.Entry(details_frame, width=20)
        self.date.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        self.date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        
        ttk.Label(details_frame, text="Category:").grid(row=3, column=0, sticky='w', pady=5)
        self.category = ttk.Combobox(details_frame, 
                                   values=["Office Supplies", "Equipment", "Software", "Services", "Travel", "Marketing", "Utilities", "Other"],
                                   state="readonly", width=20)
        self.category.set("Office Supplies")
        self.category.grid(row=3, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Supplier:").grid(row=4, column=0, sticky='w', pady=5)
        self.supplier = ttk.Entry(details_frame, width=30)
        self.supplier.grid(row=4, column=1, sticky='w', pady=5, padx=10)
        
        # Account Selection - FIXED: Use grid instead of pack
        ttk.Label(details_frame, text="Account *:").grid(row=5, column=0, sticky='w', pady=5)
        self.account = ttk.Combobox(details_frame, state="readonly", width=20)
        try:
            bm = BalanceManager(self.manager.invoice_folder)
            self.account['values'] = bm.get_account_names()
            if self.account['values']:
                self.account.set(self.account['values'][0])
            else:
                self.account['values'] = ["Cash"]
                self.account.set("Cash")
        except Exception:
            self.account['values'] = ["Cash"]
            self.account.set("Cash")
        self.account.grid(row=5, column=1, sticky='w', pady=5, padx=10)

        # Payment Status section
        pay_frame = ttk.LabelFrame(details_frame, text="Payment (at purchase time)", padding="8")
        pay_frame.grid(row=6, column=0, columnspan=2, sticky='we', pady=8)

        ttk.Label(pay_frame, text="Paid Status:").grid(row=0, column=0, sticky='w', padx=(0,6), pady=4)
        self.paid_status = ttk.Combobox(pay_frame, values=["Pending","Partial","Paid"], state="readonly", width=10)
        self.paid_status.set("Pending")
        self.paid_status.grid(row=0, column=1, sticky='w', pady=4)
        self.paid_status.bind('<<ComboboxSelected>>', self._on_paid_change)

        ttk.Label(pay_frame, text="Amount Paid (AED):").grid(row=1, column=0, sticky='w', padx=(0,6), pady=4)
        self.amount_paid = ttk.Entry(pay_frame, width=18)
        self.amount_paid.insert(0, "0.00")
        self.amount_paid.grid(row=1, column=1, sticky='w', pady=4)
        self.amount_paid.config(state='disabled')

        ttk.Label(pay_frame, text="Payment Date:").grid(row=2, column=0, sticky='w', padx=(0,6), pady=4)
        self.payment_date = ttk.Entry(pay_frame, width=18)
        self.payment_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.payment_date.grid(row=2, column=1, sticky='w', pady=4)
        self.payment_date.config(state='disabled')

        ttk.Label(pay_frame, text="Paid From (Bank):").grid(row=3, column=0, sticky='w', padx=(0,6), pady=4)
        self.bank_account = ttk.Combobox(pay_frame, state="readonly", width=20)
        try:
            bm = BalanceManager(self.manager.invoice_folder)
            vals = bm.get_account_names()
            if vals:
                self.bank_account['values'] = vals
                self.bank_account.set(vals[0] if vals else "Emirates NBD")
            else:
                self.bank_account['values'] = ["Emirates NBD","ADCB","DrAhmed ADCB","Cash"]
                self.bank_account.set("Emirates NBD")
        except Exception:
            self.bank_account['values'] = ["Emirates NBD","ADCB","DrAhmed ADCB","Cash"]
            self.bank_account.set("Emirates NBD")
        self.bank_account.grid(row=3, column=1, sticky='w', pady=4)
        self.bank_account.config(state='disabled')
        
        # Receipt Section
        receipt_frame = ttk.LabelFrame(main_frame, text="Receipt", padding="10")
        receipt_frame.pack(fill='x', pady=5)
        
        self.receipt_attached = False
        self.receipt_filename = ""
        
        ttk.Button(receipt_frame, text="📎 Attach Receipt", 
                  command=self.attach_receipt).pack(pady=5)
        
        self.receipt_label = ttk.Label(receipt_frame, text="No receipt attached", 
                                      foreground="gray")
        self.receipt_label.pack(pady=5)
        
        # Notes
        notes_frame = ttk.LabelFrame(main_frame, text="Notes", padding="10")
        notes_frame.pack(fill='x', pady=5)
        
        self.notes_text = scrolledtext.ScrolledText(notes_frame, height=4)
        self.notes_text.pack(fill='x')
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Add Purchase", 
                  command=self.add_purchase).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side='left', padx=5)

    def _on_paid_change(self, event=None):
        status = self.paid_status.get()
        if status == "Pending":
            self.amount_paid.config(state='normal')
            self.amount_paid.delete(0, tk.END)
            self.amount_paid.insert(0, "0.00")
            self.amount_paid.config(state='disabled')
            self.payment_date.config(state='disabled')
            self.bank_account.config(state='disabled')
        elif status == "Paid":
            try:
                amt = float(self.amount.get() or 0)
            except Exception:
                amt = 0.0
            self.amount_paid.config(state='normal')
            self.amount_paid.delete(0, tk.END)
            self.amount_paid.insert(0, f"{amt:.2f}")
            self.payment_date.config(state='normal')
            self.bank_account.config(state='normal')
        elif status == "Partial":
            self.amount_paid.config(state='normal')
            self.payment_date.config(state='normal')
            self.bank_account.config(state='normal')
    
    def attach_receipt(self):
        """Attach receipt file to purchase"""
        file_path = filedialog.askopenfilename(
            title="Select Receipt File",
            filetypes=[
                ("All supported files", "*.pdf *.jpg *.jpeg *.png *.gif *.bmp"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.receipt_filename = file_path
            self.receipt_attached = True
            self.receipt_label.config(text=f"📎 {os.path.basename(file_path)}", foreground="green")
    
    def add_purchase(self):
        """Add the purchase to the system with balance integration"""
        description = self.description.get().strip()
        amount_str = self.amount.get().strip()
        date = self.date.get().strip()
        account = self.account.get().strip()
        
        if not description:
            messagebox.showwarning("Warning", "Please enter purchase description")
            return
            
        try:
            amount = float(amount_str)
            if amount <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a valid amount")
            return
        
        if not account:
            messagebox.showwarning("Warning", "Please select an account")
            return
        
        # Validate date
        if not date:
            date = datetime.now().strftime("%Y-%m-%d")
        else:
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Warning", "Please enter a valid date (YYYY-MM-DD)")
                return

        status = self.paid_status.get()
        try:
            amount_paid = float(self.amount_paid.get() or 0)
        except Exception:
            amount_paid = 0.0
        if amount_paid < 0:
            amount_paid = 0.0
        if status == "Pending":
            amount_paid = 0.0
        if status == "Paid":
            amount_paid = amount
        if amount_paid > amount:
            amount_paid = amount
        payment_date = ''
        bank_account = ''
        if amount_paid > 0:
            payment_date = self.payment_date.get().strip() or datetime.now().strftime("%Y-%m-%d")
            try:
                datetime.strptime(payment_date, "%Y-%m-%d")
            except Exception:
                payment_date = datetime.now().strftime("%Y-%m-%d")
            bank_account = self.bank_account.get().strip() or "Emirates NBD"
        
        # Create purchase record
        purchase = self.manager.create_purchase(
            description=description,
            amount=amount,
            date=date,
            category=self.category.get(),
            supplier=self.supplier.get().strip(),
            account=account,
            amount_paid=amount_paid,
            payment_date=payment_date,
            bank_account=bank_account,
            paid_status=status if amount_paid <= 0 else (
                "Paid" if amount_paid >= amount - 0.005 else "Partial"
            )
        )

        if not purchase:
            messagebox.showerror("Error", "Failed to create purchase record.")
            return
        
        # Also notify balance manager of purchase (legacy system)
        try:
            from balance_manager import BalanceManager
            bm = BalanceManager(self.manager.invoice_folder)
            bm.record_purchase({
                'purchase_id': purchase.purchase_id,
                'description': description,
                'amount': amount,
                'date': date,
                'account': account,
                'category': self.category.get(),
                'supplier': self.supplier.get().strip()
            }, account)
        except Exception:
            pass
        
        # Add notes
        notes = self.notes_text.get('1.0', tk.END).strip()
        if notes:
            self.manager.update_purchase(purchase.purchase_id, notes=notes)
        
        # Upload receipt if attached
        if self.receipt_attached and self.receipt_filename:
            self.manager.upload_purchase_receipt(purchase.purchase_id, self.receipt_filename)

        paid_label = f" (Paid AED {amount_paid:.2f} / {amount:.2f})" if amount_paid > 0 else " (Pending — not paid)"
        messagebox.showinfo("Success", f"Purchase added successfully!\n\n"
                                     f"ID: {purchase.purchase_id}\n"
                                     f"Description: {description}\n"
                                     f"Amount: AED {amount:.2f}{paid_label}\n"
                                     f"Category: {self.category.get()}\n"
                                     f"Account: {account}")
        self.dialog.destroy()

class EditPurchaseDialog:
    def __init__(self, parent, manager, purchase_id, purchase):
        self.manager = manager
        self.purchase_id = purchase_id
        self.purchase = purchase
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Edit Purchase {purchase_id}")
        self.dialog.geometry("640x700")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        self.load_purchase_data()
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text=f"Edit Purchase: {self.purchase_id}", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Purchase Details
        details_frame = ttk.LabelFrame(main_frame, text="Purchase Details", padding="10")
        details_frame.pack(fill='x', pady=5)
        
        # Use grid for form fields
        ttk.Label(details_frame, text="Description *:").grid(row=0, column=0, sticky='w', pady=5)
        self.description = ttk.Entry(details_frame, width=40)
        self.description.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Amount (AED) *:").grid(row=1, column=0, sticky='w', pady=5)
        self.amount = ttk.Entry(details_frame, width=20)
        self.amount.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Date:").grid(row=2, column=0, sticky='w', pady=5)
        self.date = ttk.Entry(details_frame, width=20)
        self.date.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Category:").grid(row=3, column=0, sticky='w', pady=5)
        self.category = ttk.Combobox(details_frame, 
                                   values=["Office Supplies", "Equipment", "Software", "Services", "Travel", "Marketing", "Utilities", "Other"],
                                   state="readonly", width=20)
        self.category.set("Office Supplies")
        self.category.grid(row=3, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(details_frame, text="Supplier:").grid(row=4, column=0, sticky='w', pady=5)
        self.supplier = ttk.Entry(details_frame, width=30)
        self.supplier.grid(row=4, column=1, sticky='w', pady=5, padx=10)
        
        # Account Selection - FIXED: Use grid instead of pack
        ttk.Label(details_frame, text="Account *:").grid(row=5, column=0, sticky='w', pady=5)
        self.account = ttk.Combobox(details_frame, 
                                   values=["Emirates NBD", "ADCB", "DrAhmed ADCB", "Cash"],
                                   state="readonly", width=20)
        self.account.set("Emirates NBD")
        self.account.grid(row=5, column=1, sticky='w', pady=5, padx=10)

        # Payment Info section (for read-only status display)
        pay_frame = ttk.LabelFrame(details_frame, text="Payment Info", padding="8")
        pay_frame.grid(row=6, column=0, columnspan=2, sticky='we', pady=8)
        total_amt = float(self.purchase.get('amount', 0) or 0)
        prev_paid = float(self.purchase.get('amount_paid', 0) or 0)
        status_txt = self.purchase.get('paid_status', 'Pending')
        self.summary_lbl = ttk.Label(pay_frame,
            text=f"Status: {status_txt}   |   Paid: AED {prev_paid:.2f} / AED {total_amt:.2f}   |   Outstanding: AED {max(total_amt-prev_paid,0):.2f}")
        self.summary_lbl.grid(row=0, column=0, columnspan=3, sticky='w', pady=4)

        ttk.Label(pay_frame, text="Last Payment Date:").grid(row=1, column=0, sticky='w', padx=(0,6), pady=4)
        self.lst_pay_date = ttk.Label(pay_frame, text=str(self.purchase.get('payment_date') or 'N/A'))
        self.lst_pay_date.grid(row=1, column=1, sticky='w', pady=4)

        ttk.Label(pay_frame, text="Bank Account Used:").grid(row=1, column=2, sticky='w', padx=(30,6), pady=4)
        self.lst_bank = ttk.Label(pay_frame, text=str(self.purchase.get('bank_account') or 'N/A'))
        self.lst_bank.grid(row=1, column=3, sticky='w', pady=4)
        
        # Receipt Section
        receipt_frame = ttk.LabelFrame(main_frame, text="Receipt", padding="10")
        receipt_frame.pack(fill='x', pady=5)
        
        self.receipt_attached = False
        self.receipt_filename = ""
        
        ttk.Button(receipt_frame, text="📎 Attach Receipt", 
                  command=self.attach_receipt).pack(side='left', padx=5)
        ttk.Button(receipt_frame, text="📁 View Receipt", 
                  command=self.view_receipt).pack(side='left', padx=5)
        ttk.Button(receipt_frame, text="🗑️ Remove Receipt", 
                  command=self.remove_receipt).pack(side='left', padx=5)
        
        self.receipt_label = ttk.Label(receipt_frame, text="No receipt attached", 
                                      foreground="gray")
        self.receipt_label.pack(pady=5)
        
        # Notes
        notes_frame = ttk.LabelFrame(main_frame, text="Notes", padding="10")
        notes_frame.pack(fill='x', pady=5)
        
        self.notes_text = scrolledtext.ScrolledText(notes_frame, height=4)
        self.notes_text.pack(fill='x')
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)

        outstanding = max(total_amt - prev_paid, 0.0)
        if outstanding > 0.005:
            ttk.Button(button_frame, text=f"💰 Record Payment (AED {outstanding:.2f})",
                       command=self.record_payment).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Save Changes", 
                  command=self.save_changes).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side='left', padx=5)

    def record_payment(self):
        dlg = PayPendingPurchaseDialog(self.dialog, self.manager, self.purchase_id, self.purchase)
        self.dialog.wait_window(dlg.dialog)
        # Refresh summary after payment
        updated = self.manager.get_purchase(self.purchase_id)
        if updated:
            self.purchase = updated.to_dict() if hasattr(updated, 'to_dict') else updated
            total_amt = float(self.purchase.get('amount', 0) or 0)
            prev_paid = float(self.purchase.get('amount_paid', 0) or 0)
            status_txt = self.purchase.get('paid_status', 'Pending')
            self.summary_lbl.config(
                text=f"Status: {status_txt}   |   Paid: AED {prev_paid:.2f} / AED {total_amt:.2f}   |   Outstanding: AED {max(total_amt-prev_paid,0):.2f}")
            self.lst_pay_date.config(text=str(self.purchase.get('payment_date') or 'N/A'))
            self.lst_bank.config(text=str(self.purchase.get('bank_account') or 'N/A'))
    
    def load_purchase_data(self):
        """Load existing purchase data into form"""
        self.description.delete(0, tk.END)
        self.description.insert(0, self.purchase.get('description', ''))
        
        self.amount.delete(0, tk.END)
        self.amount.insert(0, str(self.purchase.get('amount', 0)))
        
        self.date.delete(0, tk.END)
        self.date.insert(0, self.purchase.get('date', datetime.now().strftime("%Y-%m-%d")))
        
        self.category.set(self.purchase.get('category', 'Office Supplies'))
        
        self.supplier.delete(0, tk.END)
        self.supplier.insert(0, self.purchase.get('supplier', ''))
        
        self.account.set(self.purchase.get('account', 'Emirates NBD'))
        
        # Load receipt status
        self.receipt_attached = self.purchase.get('receipt_attached', False)
        if self.receipt_attached:
            self.receipt_label.config(text="📎 Receipt attached", foreground="green")
        
        # Load notes
        if self.purchase.get('notes'):
            self.notes_text.delete('1.0', tk.END)
            self.notes_text.insert('1.0', self.purchase.get('notes', ''))
    
    def attach_receipt(self):
        """Attach receipt file to purchase"""
        file_path = filedialog.askopenfilename(
            title="Select Receipt File",
            filetypes=[
                ("All supported files", "*.pdf *.jpg *.jpeg *.png *.gif *.bmp"),
                ("PDF files", "*.pdf"),
                ("Image files", "*.jpg *.jpeg *.png *.gif *.bmp"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.receipt_filename = file_path
            self.receipt_attached = True
            self.receipt_label.config(text=f"📎 {os.path.basename(file_path)}", foreground="green")
    
    def view_receipt(self):
        """View attached receipt"""
        if self.receipt_attached:
            success = self.manager.open_purchase_receipt(self.purchase_id)
            if not success:
                messagebox.showerror("Error", "Could not open receipt file. The file may have been moved or deleted.")
        else:
            messagebox.showinfo("Info", "No receipt attached to this purchase")
    
    def remove_receipt(self):
        """Remove attached receipt"""
        if self.receipt_attached:
            if messagebox.askyesno("Confirm", "Are you sure you want to remove the receipt?"):
                self.receipt_attached = False
                self.receipt_filename = ""
                self.receipt_label.config(text="No receipt attached", foreground="gray")
                # Also remove from purchase data
                self.manager.update_purchase(self.purchase_id, receipt_attached=False)
    
    def save_changes(self):
        """Save changes to purchase"""
        description = self.description.get().strip()
        amount_str = self.amount.get().strip()
        date = self.date.get().strip()
        account = self.account.get().strip()
        
        if not description:
            messagebox.showwarning("Warning", "Please enter purchase description")
            return
            
        try:
            amount = float(amount_str)
            if amount <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Warning", "Please enter a valid amount")
            return
        
        if not account:
            messagebox.showwarning("Warning", "Please select an account")
            return
        
        # Validate date
        if not date:
            date = datetime.now().strftime("%Y-%m-%d")
        else:
            try:
                datetime.strptime(date, "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("Warning", "Please enter a valid date (YYYY-MM-DD)")
                return
        
        # Keep existing payment info if amount changed - recompute status
        prev_paid = float(self.purchase.get('amount_paid', 0) or 0)
        if prev_paid <= 0:
            new_status = 'Pending'
        elif prev_paid >= amount - 0.005:
            new_status = 'Paid'
        else:
            new_status = 'Partial'

        # Update purchase
        success = self.manager.update_purchase(
            self.purchase_id,
            description=description,
            amount=amount,
            date=date,
            category=self.category.get(),
            supplier=self.supplier.get().strip(),
            account=account,
            notes=self.notes_text.get('1.0', tk.END).strip(),
            receipt_attached=self.receipt_attached,
            paid_status=new_status
        )
        
        # Upload new receipt if attached
        if self.receipt_attached and self.receipt_filename:
            self.manager.upload_purchase_receipt(self.purchase_id, self.receipt_filename)
        
        if success:
            messagebox.showinfo("Success", "Purchase updated successfully!")
            self.dialog.destroy()
        else:
            messagebox.showerror("Error", "Failed to update purchase")


class PayPendingPurchaseDialog:
    """Dialog to record a payment for an outstanding (pending/partial) purchase."""
    def __init__(self, parent, manager, purchase_id, purchase):
        self.manager = manager
        self.purchase_id = purchase_id
        self.purchase = purchase
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Record Payment — {purchase_id}")
        self.dialog.geometry("520x360")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(False, False)
        self.setup_ui()

    def setup_ui(self):
        frm = ttk.Frame(self.dialog, padding=16)
        frm.pack(fill='both', expand=True)

        total = float(self.purchase.get('amount', 0) or 0)
        paid = float(self.purchase.get('amount_paid', 0) or 0)
        outstanding = max(total - paid, 0.0)
        supplier = self.purchase.get('supplier') or 'N/A'
        desc = self.purchase.get('description') or ''

        ttk.Label(frm, text=f"Record Payment: {self.purchase_id}",
                  font=('Helvetica', 14, 'bold')).grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 12))

        info = ttk.LabelFrame(frm, text="Purchase Summary", padding=10)
        info.grid(row=1, column=0, columnspan=2, sticky='we', pady=(0, 12))
        ttk.Label(info, text=f"Supplier: {supplier}").grid(row=0, column=0, sticky='w')
        ttk.Label(info, text=f"Description: {desc}").grid(row=1, column=0, sticky='w', pady=2)
        ttk.Label(info, text=f"Total: AED {total:.2f}   |   Already Paid: AED {paid:.2f}   |   Outstanding: AED {outstanding:.2f}").grid(row=2, column=0, sticky='w', pady=2)

        form = ttk.LabelFrame(frm, text="Payment Details", padding=10)
        form.grid(row=2, column=0, columnspan=2, sticky='we')
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="Payment Date *:").grid(row=0, column=0, sticky='w', pady=4, padx=(0,10))
        self.payment_date = ttk.Entry(form)
        self.payment_date.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.payment_date.grid(row=0, column=1, sticky='we', pady=4)

        ttk.Label(form, text="Amount (AED) *:").grid(row=1, column=0, sticky='w', pady=4, padx=(0,10))
        self.payment_amount = ttk.Entry(form)
        self.payment_amount.insert(0, f"{outstanding:.2f}")
        self.payment_amount.grid(row=1, column=1, sticky='we', pady=4)

        ttk.Label(form, text="Paid From (Bank) *:").grid(row=2, column=0, sticky='w', pady=4, padx=(0,10))
        self.bank_account = ttk.Combobox(form, state="readonly")
        try:
            from balance_manager import BalanceManager
            bm = BalanceManager(self.manager.invoice_folder)
            vals = bm.get_account_names()
            if vals:
                self.bank_account['values'] = vals
                self.bank_account.set(self.purchase.get('bank_account') or vals[0])
            else:
                self.bank_account['values'] = ["Emirates NBD","ADCB","DrAhmed ADCB","Cash"]
                self.bank_account.set(self.purchase.get('bank_account') or "Emirates NBD")
        except Exception:
            self.bank_account['values'] = ["Emirates NBD","ADCB","DrAhmed ADCB","Cash"]
            self.bank_account.set(self.purchase.get('bank_account') or "Emirates NBD")
        self.bank_account.grid(row=2, column=1, sticky='we', pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=3, column=0, columnspan=2, sticky='e', pady=(16, 0))
        ttk.Button(btns, text="Save Payment", command=self.save_payment).pack(side='right')
        ttk.Button(btns, text="Cancel", command=self.dialog.destroy).pack(side='right', padx=6)

    def save_payment(self):
        pay_date = self.payment_date.get().strip()
        amt_str = self.payment_amount.get().strip()
        bank = (self.bank_account.get() or "").strip()
        try:
            datetime.strptime(pay_date, "%Y-%m-%d")
        except Exception:
            messagebox.showwarning("Warning", "Please enter a valid date (YYYY-MM-DD)")
            return
        try:
            amt = float(amt_str)
            if amt <= 0:
                raise ValueError
        except Exception:
            messagebox.showwarning("Warning", "Please enter a valid payment amount greater than 0")
            return
        if not bank:
            messagebox.showwarning("Warning", "Please select a bank account")
            return
        if hasattr(self.manager, 'record_purchase_payment'):
            ok, msg = self.manager.record_purchase_payment(
                self.purchase_id, amt, payment_date=pay_date, bank_account=bank, username="Operator")
        else:
            # Fallback if manager is stub
            prev_paid = float(self.purchase.get('amount_paid', 0) or 0)
            total = float(self.purchase.get('amount', 0) or 0)
            new_paid = min(prev_paid + amt, total)
            new_status = 'Paid' if new_paid >= total - 0.005 else ('Partial' if new_paid > 0 else 'Pending')
            ok = self.manager.update_purchase(
                self.purchase_id, amount_paid=round(new_paid, 2),
                payment_date=pay_date, bank_account=bank, paid_status=new_status)
            msg = "Payment recorded (fallback path — no GL posting)." if ok else "Failed to update purchase."
        if ok:
            messagebox.showinfo("Payment Saved", msg)
            self.dialog.destroy()
        else:
            messagebox.showerror("Error", msg)

class ViewPurchaseDialog:
    def __init__(self, parent, manager, purchase_id, purchase):
        self.manager = manager
        self.purchase_id = purchase_id
        self.purchase = purchase
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Purchase {purchase_id}")
        self.dialog.geometry("540x560")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text=f"PURCHASE: {self.purchase_id}", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 10))
        
        # Purchase Details
        details_frame = ttk.LabelFrame(main_frame, text="Purchase Details", padding="10")
        details_frame.pack(fill='x', pady=5)
        
        ttk.Label(details_frame, text=f"Description: {self.purchase.get('description', '')}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='w', pady=2)
        ttk.Label(details_frame, text=f"Amount: AED {float(self.purchase.get('amount', 0) or 0):.2f}", 
                 font=('Helvetica', 10, 'bold')).pack(anchor='w', pady=2)
        ttk.Label(details_frame, text=f"Date: {self.purchase.get('date', '')}").pack(anchor='w', pady=2)
        ttk.Label(details_frame, text=f"Category: {self.purchase.get('category', '')}").pack(anchor='w', pady=2)
        ttk.Label(details_frame, text=f"Supplier: {self.purchase.get('supplier', 'N/A')}").pack(anchor='w', pady=2)
        ttk.Label(details_frame, text=f"Account: {self.purchase.get('account', 'N/A')}").pack(anchor='w', pady=2)

        # Payment info block
        pay_frame = ttk.LabelFrame(details_frame, text="Payment Status", padding="8")
        pay_frame.pack(fill='x', pady=6)
        total = float(self.purchase.get('amount', 0) or 0)
        paid = float(self.purchase.get('amount_paid', 0) or 0)
        outstanding = max(total - paid, 0.0)
        status = self.purchase.get('paid_status', 'Pending')
        if status == 'Paid':
            fg = '#1b7a3d'
            tag = '✔ PAID'
        elif status == 'Partial':
            fg = '#b45309'
            tag = '◐ PARTIAL'
        else:
            fg = '#b91c1c'
            tag = '◯ PENDING'
        ttk.Label(pay_frame, text=f"{tag} — {status}", foreground=fg,
                  font=('Helvetica', 11, 'bold')).pack(anchor='w', pady=1)
        ttk.Label(pay_frame,
                  text=f"Total Amount:  AED {total:.2f}").pack(anchor='w', pady=1)
        ttk.Label(pay_frame,
                  text=f"Amount Paid:   AED {paid:.2f}").pack(anchor='w', pady=1)
        ttk.Label(pay_frame,
                  text=f"Outstanding:   AED {outstanding:.2f}").pack(anchor='w', pady=1)
        if self.purchase.get('payment_date'):
            ttk.Label(pay_frame,
                      text=f"Last Payment Date: {self.purchase.get('payment_date')}").pack(anchor='w', pady=1)
        if self.purchase.get('bank_account'):
            ttk.Label(pay_frame,
                      text=f"Paid From (Bank): {self.purchase.get('bank_account')}").pack(anchor='w', pady=1)
        
        # Receipt Section
        receipt_frame = ttk.LabelFrame(main_frame, text="Receipt", padding="10")
        receipt_frame.pack(fill='x', pady=5)
        
        if self.purchase.get('receipt_attached'):
            ttk.Button(receipt_frame, text="📁 View Receipt", 
                      command=self.view_receipt).pack(pady=5)
            ttk.Label(receipt_frame, text="Receipt attached", 
                     foreground="green").pack(pady=2)
        else:
            ttk.Label(receipt_frame, text="No receipt attached", 
                     foreground="gray").pack(pady=5)
        
        # Notes
        if self.purchase.get('notes'):
            notes_frame = ttk.LabelFrame(main_frame, text="Notes", padding="10")
            notes_frame.pack(fill='x', pady=5)
            
            notes_text = scrolledtext.ScrolledText(notes_frame, height=4)
            notes_text.pack(fill='x')
            notes_text.insert('1.0', self.purchase.get('notes', ''))
            notes_text.config(state='disabled')

        # Actions
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=12)
        if outstanding > 0.005:
            ttk.Button(btn_frame, text=f"💰 Record Payment (AED {outstanding:.2f})",
                       command=self.record_payment).pack(side='left', padx=6)
        ttk.Button(btn_frame, text="Close", 
                  command=self.dialog.destroy).pack(side='left', padx=6)

    def record_payment(self):
        dlg = PayPendingPurchaseDialog(self.dialog, self.manager, self.purchase_id, self.purchase)
        self.dialog.wait_window(dlg.dialog)
        # Refresh summary
        updated = self.manager.get_purchase(self.purchase_id)
        if updated:
            self.purchase = updated.to_dict() if hasattr(updated, 'to_dict') else updated
            # Redraw view
            for w in self.dialog.winfo_children():
                w.destroy()
            self.setup_ui()
    
    def view_receipt(self):
        """Open the purchase receipt"""
        success = self.manager.open_purchase_receipt(self.purchase_id)
        if not success:
            messagebox.showerror("Error", "Could not open receipt file. The file may have been moved or deleted.")

class PurchasesDialog:
    def __init__(self, parent, manager):
        self.manager = manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Purchases Management")
        self.dialog.geometry("1320x700")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        self.refresh_purchases()
        _fit_window(self.dialog, 1320, 700)
    
    def setup_ui(self):
        # Create main container with scrollbar
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        # Create scrollable frame
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="🛒 Purchases Management", 
                 font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))
        
        # Search and controls
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(controls_frame, text="Search:").pack(side='left', padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(controls_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side='left', padx=(0, 10))
        self.search_entry.bind('<KeyRelease>', self.on_search)

        ttk.Label(controls_frame, text="Payment Filter:").pack(side='left', padx=(20, 5))
        self.payment_filter = ttk.Combobox(controls_frame, values=["All","Pending","Partial","Paid"], state="readonly", width=10)
        self.payment_filter.set("All")
        self.payment_filter.pack(side='left', padx=(0, 10))
        self.payment_filter.bind('<<ComboboxSelected>>', lambda e: self.refresh_purchases())
        
        # Sort options
        ttk.Label(controls_frame, text="Sort by:").pack(side='left', padx=(20, 5))
        self.sort_var = tk.StringVar(value="date_desc")
        sort_combo = ttk.Combobox(controls_frame, textvariable=self.sort_var, 
                                 values=["Date (New-Old)", "Date (Old-New)", "Amount (High-Low)", "Amount (Low-High)", "Description (A-Z)", "Description (Z-A)", "Outstanding (High-Low)"],
                                 state="readonly", width=18)
        sort_combo.pack(side='left', padx=(0, 10))
        sort_combo.bind('<<ComboboxSelected>>', self.on_sort_change)
        
        # Buttons
        button_frame = ttk.Frame(controls_frame)
        button_frame.pack(side='right')
        
        ttk.Button(button_frame, text="➕ New Purchase", 
                  command=self.create_purchase).pack(side='left', padx=5)
        ttk.Button(button_frame, text="💰 Record Payment",
                  command=self.record_payment_for_selected).pack(side='left', padx=5)
        ttk.Button(button_frame, text="👀 View Selected", 
                  command=self.view_purchase).pack(side='left', padx=5)
        if WAREHOUSE_EXTENSION_AVAILABLE:
            ttk.Button(button_frame, text="📥 Receive Stock",
                      command=self.receive_stock).pack(side='left', padx=5)
        ttk.Button(button_frame, text="✏️ Edit Purchase", 
                  command=self.edit_purchase).pack(side='left', padx=5)
        ttk.Button(button_frame, text="🗑️ Delete Purchase", 
                  command=self.delete_purchase).pack(side='left', padx=5)
        ttk.Button(button_frame, text="🔄 Refresh", 
                  command=self.refresh_purchases).pack(side='left', padx=5)
        
        # Purchases list
        list_frame = ttk.LabelFrame(main_frame, text="Purchases", padding="10")
        list_frame.pack(fill='both', expand=True)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill='both', expand=True)
        
        columns = ('ID', 'Date', 'Supplier', 'Description', 'Amount', 'Paid', 'Balance', 'Status', 'Category', 'Account', 'Receipt')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Define headings
        for c, label, w, align in [
            ('ID','Purchase ID', 150, 'w'),
            ('Date','Date', 95, 'center'),
            ('Supplier','Supplier', 170, 'w'),
            ('Description','Description', 180, 'w'),
            ('Amount','Amount (AED)', 105, 'e'),
            ('Paid','Paid (AED)', 105, 'e'),
            ('Balance','Owed (AED)', 105, 'e'),
            ('Status','Status', 90, 'center'),
            ('Category','Category', 110, 'w'),
            ('Account','Account', 120, 'w'),
            ('Receipt','Rec', 45, 'center'),
        ]:
            self.tree.heading(c, text=label)
            self.tree.column(c, width=w, anchor=align)

        # Color tagging for statuses
        self.tree.tag_configure('pending', background='#FEF2F2')
        self.tree.tag_configure('partial', background='#FFFBEB')
        self.tree.tag_configure('paid', background='#ECFDF5')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Double click to view
        self.tree.bind('<Double-1>', lambda e: self.view_purchase())

        # Summary totals bar
        self.summary_var = tk.StringVar(value="—")
        ttk.Label(main_frame, textvariable=self.summary_var,
                  font=('Helvetica', 10, 'bold'),
                  anchor='w', relief='ridge', padding=8).pack(fill='x', pady=(8, 0))
        
        # Status
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, font=('Helvetica', 10))
        status_bar.pack(fill='x', side='bottom', pady=(10, 0))

    def _format_status(self, status):
        if status == 'Paid':   return "✔ Paid"
        if status == 'Partial':return "◐ Partial"
        return "◯ Pending"

    def _tag_for_status(self, status):
        if status == 'Paid':   return 'paid'
        if status == 'Partial':return 'partial'
        return 'pending'
    
    def refresh_purchases(self):
        """Refresh the purchases list"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        purchases = self.manager.get_all_purchases_dict()
        flt = self.payment_filter.get()
        if flt and flt != "All":
            purchases = [p for p in purchases if p.get('paid_status','Pending') == flt]
        purchases = self.sort_purchases(purchases, self.sort_var.get())

        total = sum(float(p.get('amount', 0) or 0) for p in purchases)
        total_paid = sum(float(p.get('amount_paid', 0) or 0) for p in purchases)
        total_owed = sum(max(float(p.get('amount', 0) or 0) - float(p.get('amount_paid', 0) or 0), 0.0) for p in purchases)
        pending_count = sum(1 for p in purchases if p.get('paid_status','Pending') == 'Pending')
        partial_count = sum(1 for p in purchases if p.get('paid_status','Pending') == 'Partial')
        paid_count = sum(1 for p in purchases if p.get('paid_status','Pending') == 'Paid')
        self.summary_var.set(
            f"Showing {len(purchases)}: {pending_count} Pending | {partial_count} Partial | {paid_count} Paid   "
            f"|   Totals — Amount: AED {total:.2f}   Paid: AED {total_paid:.2f}   Owed: AED {total_owed:.2f}"
        )
        
        for p in purchases:
            amount = float(p.get('amount', 0) or 0)
            paid = float(p.get('amount_paid', 0) or 0)
            owed = max(amount - paid, 0.0)
            status = p.get('paid_status', 'Pending')
            receipt_indicator = "📎" if p.get('receipt_attached') else ""
            self.tree.insert('', tk.END, values=(
                p.get('purchase_id', ''),
                p.get('date', ''),
                p.get('supplier', 'N/A'),
                p.get('description', ''),
                f"AED {amount:.2f}",
                f"AED {paid:.2f}",
                f"AED {owed:.2f}",
                self._format_status(status),
                p.get('category', ''),
                p.get('account', 'N/A'),
                receipt_indicator
            ), tags=(self._tag_for_status(status),))
        
        self.status_var.set(f"Loaded {len(purchases)} purchases (filter: {flt})")
    
    def sort_purchases(self, purchases, sort_option):
        """Sort purchases based on selected option"""
        if sort_option == "Date (New-Old)":
            return sorted(purchases, key=lambda x: x.get('date', ''), reverse=True)
        elif sort_option == "Date (Old-New)":
            return sorted(purchases, key=lambda x: x.get('date', ''))
        elif sort_option == "Amount (High-Low)":
            return sorted(purchases, key=lambda x: x.get('amount', 0), reverse=True)
        elif sort_option == "Amount (Low-High)":
            return sorted(purchases, key=lambda x: x.get('amount', 0))
        elif sort_option == "Description (A-Z)":
            return sorted(purchases, key=lambda x: x.get('description', '').lower())
        elif sort_option == "Description (Z-A)":
            return sorted(purchases, key=lambda x: x.get('description', '').lower(), reverse=True)
        elif sort_option == "Outstanding (High-Low)":
            return sorted(purchases, key=lambda x: (float(x.get('amount',0) or 0) - float(x.get('amount_paid',0) or 0)), reverse=True)
        else:
            return purchases
    
    def on_sort_change(self, event=None):
        """Handle sort change"""
        self.refresh_purchases()
    
    def on_search(self, event=None):
        """Handle search functionality"""
        query = self.search_var.get().strip()
        if not query:
            self.refresh_purchases()
            return
            
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        results = self.manager.search_purchases(query)
        purchases = [p.to_dict() for p in results]
        flt = self.payment_filter.get()
        if flt and flt != "All":
            purchases = [p for p in purchases if p.get('paid_status','Pending') == flt]
        results = self.sort_purchases(purchases, self.sort_var.get())

        for p in purchases:
            amount = float(p.get('amount', 0) or 0)
            paid = float(p.get('amount_paid', 0) or 0)
            owed = max(amount - paid, 0.0)
            status = p.get('paid_status', 'Pending')
            receipt_indicator = "📎" if p.get('receipt_attached') else ""
            self.tree.insert('', tk.END, values=(
                p.get('purchase_id', ''),
                p.get('date', ''),
                p.get('supplier', 'N/A'),
                p.get('description', ''),
                f"AED {amount:.2f}",
                f"AED {paid:.2f}",
                f"AED {owed:.2f}",
                self._format_status(status),
                p.get('category', ''),
                p.get('account', 'N/A'),
                receipt_indicator
            ), tags=(self._tag_for_status(status),))
        
        self.status_var.set(f"Found {len(purchases)} purchases matching '{query}' (filter: {flt})")
    
    def get_selected_purchase(self):
        """Get currently selected purchase"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a purchase first")
            return None, None
            
        item = self.tree.item(selection[0])
        purchase_id = item['values'][0]
        
        # Get the purchase as dictionary for GUI compatibility
        purchase_obj = self.manager.get_purchase(purchase_id)
        if purchase_obj:
            return purchase_id, purchase_obj.to_dict()
        return None, None
    
    def create_purchase(self):
        """Open create purchase dialog"""
        dialog = CreatePurchaseDialog(self.dialog, self.manager)
        self.dialog.wait_window(dialog.dialog)
        self.refresh_purchases()
    
    def view_purchase(self):
        """View selected purchase"""
        purchase_id, purchase = self.get_selected_purchase()
        if purchase:
            ViewPurchaseDialog(self.dialog, self.manager, purchase_id, purchase)
    
    def edit_purchase(self):
        """Edit selected purchase"""
        purchase_id, purchase = self.get_selected_purchase()
        if purchase:
            EditPurchaseDialog(self.dialog, self.manager, purchase_id, purchase)
            self.refresh_purchases()

    def record_payment_for_selected(self):
        """Record payment for selected pending/partial purchase"""
        purchase_id, purchase = self.get_selected_purchase()
        if not purchase:
            return
        total = float(purchase.get('amount', 0) or 0)
        paid = float(purchase.get('amount_paid', 0) or 0)
        if paid >= total - 0.005:
            messagebox.showinfo("Already Paid", f"Purchase {purchase_id} is already fully paid (AED {total:.2f}).")
            return
        dlg = PayPendingPurchaseDialog(self.dialog, self.manager, purchase_id, purchase)
        self.dialog.wait_window(dlg.dialog)
        self.refresh_purchases()

    def receive_stock(self):
        if not WAREHOUSE_EXTENSION_AVAILABLE:
            messagebox.showerror("Error", "Warehouse extension is not available")
            return
        purchase_id, purchase = self.get_selected_purchase()
        if not purchase:
            return
        warehouse_manager = WarehouseManager(getattr(self.manager, "invoice_folder", os.getcwd()))
        warehouse_codes = [row.get("warehouse_code") for row in warehouse_manager.list_warehouses()]
        dlg = tk.Toplevel(self.dialog)
        dlg.title(f"Receive Stock from {purchase_id}")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        vars_ = {
            "client_name": tk.StringVar(),
            "supplier": tk.StringVar(value=str(purchase.get("supplier") or "")),
            "purchase_id": tk.StringVar(value=str(purchase_id or "")),
            "product_name": tk.StringVar(value=str(purchase.get("description") or "")),
            "brand": tk.StringVar(),
            "batch_number": tk.StringVar(),
            "lot_number": tk.StringVar(),
            "manufacture_date": tk.StringVar(),
            "expiry_date": tk.StringVar(),
            "quantity_received": tk.StringVar(value="1"),
            "unit_cost": tk.StringVar(value=str(purchase.get("amount") or 0)),
            "warehouse_code": tk.StringVar(value=warehouse_codes[0] if warehouse_codes else ""),
            "bin_location": tk.StringVar(),
            "notes": tk.StringVar(value=str(purchase.get("notes") or "")),
        }
        fields = [
            ("Company / Client", "client_name", "entry"),
            ("Supplier", "supplier", "entry"),
            ("Purchase ID", "purchase_id", "entry"),
            ("Product Name", "product_name", "entry"),
            ("Brand", "brand", "entry"),
            ("Batch Number", "batch_number", "entry"),
            ("Lot Number", "lot_number", "entry"),
            ("Manufacture Date", "manufacture_date", "entry"),
            ("Expiry Date", "expiry_date", "entry"),
            ("Quantity Received", "quantity_received", "entry"),
            ("Unit Cost", "unit_cost", "entry"),
            ("Warehouse", "warehouse_code", "combo"),
            ("Shelf / Bin / Location Code", "bin_location", "entry"),
            ("Notes", "notes", "entry"),
        ]
        for idx, (label, key, kind) in enumerate(fields):
            ttk.Label(frm, text=label).grid(row=idx, column=0, sticky="w", pady=4, padx=(0, 8))
            if kind == "combo":
                ttk.Combobox(frm, textvariable=vars_[key], values=warehouse_codes, state="readonly").grid(row=idx, column=1, sticky="ew", pady=4)
            else:
                ttk.Entry(frm, textvariable=vars_[key]).grid(row=idx, column=1, sticky="ew", pady=4)
        frm.columnconfigure(1, weight=1)

        def save_receipt():
            payload = {key: var.get() for key, var in vars_.items()}
            try:
                warehouse_manager.receive_goods(payload, user_name="Purchase Receipt")
                dlg.destroy()
                messagebox.showinfo("Success", f"Stock received successfully for {purchase_id}")
            except Exception as exc:
                messagebox.showerror("Error", f"Could not receive stock:\n{exc}")

        btns = ttk.Frame(frm)
        btns.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Receive", command=save_receipt, style="Primary.TButton").pack(side="left")
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        _fit_window(dlg, 720, 640, mode="large", remember_key="purchase_receive_stock")
    
    def delete_purchase(self):
        """Delete selected purchase with balance adjustment"""
        purchase_id, purchase = self.get_selected_purchase()
        if purchase:
            if messagebox.askyesno("Confirm Delete", 
                                 f"Are you sure you want to delete purchase {purchase_id}?\n"
                                 f"Description: {purchase.get('description', '')}\n"
                                 f"Amount: AED {purchase.get('amount', 0):.2f}\n\n"
                                 f"This will reverse the purchase amount in your balance."):
                
                # Use the new method that adjusts balances
                if hasattr(self.manager, 'delete_purchase_with_balance_adjustment'):
                    success, message = self.manager.delete_purchase_with_balance_adjustment(purchase_id)
                else:
                    # Fallback to old method
                    success = self.manager.delete_purchase(purchase_id)
                    message = "Purchase deleted (balance not adjusted)"
                
                if success:
                    messagebox.showinfo("Success", message)
                    try:
                        folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                        path = os.path.join(folder, 'audit_log.jsonl')
                        rec = {
                            'timestamp': datetime.now().isoformat(),
                            'action': 'purchase_deleted',
                            'details': {
                                'purchase_id': purchase_id,
                                'amount': purchase.get('amount', 0),
                                'account': purchase.get('account', 'Cash')
                            }
                        }
                        with open(path, 'a') as f:
                            f.write(json.dumps(rec) + "\n")
                    except Exception:
                        pass
                    self.refresh_purchases()
                else:
                    messagebox.showerror("Error", message)

class InvoiceApp:
    def __init__(self, root, manager=None):
        self.root = root
        self.root.title("ASISTEM — Multi-Company Business Suite")
        self.root.geometry("1200x800")
        self.root.resizable(True, True)
        try:
            self.root.columnconfigure(0, weight=1)
            self.root.rowconfigure(0, weight=1)
            self.root.bind("<Configure>", self._on_window_resize)
        except Exception:
            pass
        
        if manager:
            self.manager = manager
        else:
            self.manager = InvoiceManager()
            try:
                from hope_pharma_complete import CloudDataManager
                
                # Detect data folder location
                data_folder = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")
                possible_roots = [
                    os.path.join(os.path.expanduser("~"), "Google Drive"),
                    os.path.join(os.path.expanduser("~"), "GoogleDrive"),
                    os.path.join(os.path.expanduser("~"), "Documents"),
                    os.path.join(os.path.expanduser("~"), "Desktop")
                ]
                
                for root in possible_roots:
                    path = os.path.join(root, "HopePharmaData")
                    if os.path.exists(path):
                        data_folder = path
                        break
                        
                self.manager = CloudDataManager(data_folder, "single")
            except Exception:
                pass

        # Track file modification times for auto-refresh
        self.last_sync_times = {}
        self.files_to_watch = []
        df = None
        if hasattr(self.manager, 'data_folder'):
            df = str(self.manager.data_folder)
        elif hasattr(self.manager, 'invoice_folder'):
            df = str(self.manager.invoice_folder)
            
        if df:
            self.files_to_watch = [
                os.path.join(df, "invoices_data.json"),
                os.path.join(df, "purchases_data.json"),
                os.path.join(df, "inventory.json"),
                os.path.join(df, "balance_data.json")
            ]
        
        self.logo_manager = LogoManager(df)
        self.data_memory = DataMemoryManager(df if df else self.manager.invoice_folder)
        try:
            if self.logo_manager.load_logo(size=(64, 64)):
                self.root.iconphoto(True, self.logo_manager.logo_photo)
        except Exception:
            pass
        
        try:
            self._ensure_app_menu()
        except Exception:
            pass

        try:
            self._load_ui_prefs()
        except Exception:
            pass

        try:
            self.root._invoice_app = self
            self.root._ui_window_bounds = getattr(self, "ui_window_bounds", {})
            self.root.protocol("WM_DELETE_WINDOW", self._on_app_close)
        except Exception:
            pass

        try:
            self._apply_saas_theme()
        except Exception:
            pass

        try:
            self._restore_root_window_geometry()
        except Exception:
            pass

        self._show_start_animation()
        try:
            self.root.after(2000, self.update_notifications_badge)
            self.root.after(3000, self.start_auto_refresh)
        except Exception:
            pass
    
    def start_auto_refresh(self):
        """Start the auto-refresh cycle for the main app"""
        try:
            if not self.root.winfo_exists():
                return
            self.check_for_updates()
            # Check every 7 seconds for the main app (slightly slower than inventory)
            self.root.after(7000, self.start_auto_refresh)
        except Exception:
            pass

    def check_for_updates(self):
        """Check if any data files have been modified by another user"""
        try:
            modified = False
            for file_path in self.files_to_watch:
                if os.path.exists(file_path):
                    mtime = os.path.getmtime(file_path)
                    if file_path not in self.last_sync_times:
                        self.last_sync_times[file_path] = mtime
                    elif mtime > self.last_sync_times[file_path]:
                        self.last_sync_times[file_path] = mtime
                        modified = True
            
            if modified:
                print(f"DEBUG: Changes detected in {len(self.files_to_watch)} data files, refreshing UI...")
                # Refresh invoices list if it exists
                if hasattr(self, 'tree'):
                    self.refresh_invoices()
                # Refresh purchases if it exists
                if hasattr(self, 'purchase_tree'):
                    self.refresh_purchases()
                # Refresh dashboard if active
                if hasattr(self, 'refresh_dashboard_metrics'):
                    self.refresh_dashboard_metrics()
        except Exception as e:
            print(f"Main app auto-refresh error: {e}")
    
    def _on_window_resize(self, event):
        try:
            if getattr(event, "widget", None) is not self.root:
                return
            self._resize_treeviews()
            self._schedule_remember_window("main_window", self.root)
        except Exception:
            pass
    
    def _iter_widgets(self, parent):
        try:
            for w in parent.winfo_children():
                yield w
                for c in self._iter_widgets(w):
                    yield c
        except Exception:
            return
    
    def _resize_treeviews(self):
        try:
            for w in self._iter_widgets(self.root):
                if isinstance(w, ttk.Treeview):
                    self._attach_treeview_resize_hooks(w)
                    if getattr(w, "_manual_column_resize", False):
                        continue
                    width = w.winfo_width()
                    cols = w["columns"]
                    if not cols:
                        continue
                    if width <= 0:
                        continue
                    pad = 24
                    usable_width = max(120, width - pad)
                    sample_items = list(w.get_children(""))[:40]
                    base_sizes = []
                    for c in cols:
                        try:
                            heading_text = str(w.heading(c, option="text") or c)
                        except Exception:
                            heading_text = str(c)
                        max_chars = max(8, len(heading_text))
                        for iid in sample_items:
                            try:
                                row_vals = w.item(iid, "values") or []
                                col_index = list(cols).index(c)
                                cell_text = str(row_vals[col_index] if col_index < len(row_vals) else "")
                                max_chars = max(max_chars, min(len(cell_text), 32))
                            except Exception:
                                continue
                        base_sizes.append(max(80, min(320, int(max_chars * 9.5) + 28)))
                    total_base = sum(base_sizes) or 1
                    scale = usable_width / total_base
                    for c, base in zip(cols, base_sizes):
                        try:
                            target = max(80, int(base * scale))
                            w.column(c, width=target, stretch=True)
                        except Exception:
                            pass
        except Exception:
            pass

    def _attach_treeview_resize_hooks(self, tree):
        try:
            if getattr(tree, "_resize_hooks_attached", False):
                return
            tree._resize_hooks_attached = True
            tree._manual_column_resize = False
            tree.bind("<ButtonRelease-1>", lambda e, tv=tree: self._mark_treeview_manual_resize(tv, e), add="+")
            tree.bind("<Double-Button-1>", lambda _e, tv=tree: self._reset_treeview_auto_resize(tv), add="+")
        except Exception:
            pass

    def _mark_treeview_manual_resize(self, tree, event=None):
        try:
            if event is not None and tree.identify_region(event.x, event.y) == "separator":
                tree._manual_column_resize = True
        except Exception:
            pass

    def _reset_treeview_auto_resize(self, tree):
        try:
            tree._manual_column_resize = False
            self._resize_treeviews()
        except Exception:
            pass

    def _theme_palette(self):
        mode = str(getattr(self, "ui_theme_mode", "light") or "light").strip().lower()
        if mode == "dark":
            return {
                "bg": "#0b1220",
                "card_bg": "#111827",
                "text": "#e5e7eb",
                "muted": "#9ca3af",
                "border": "#1f2937",
                "blue": "#60a5fa",
                "green": "#34d399",
                "orange": "#fbbf24",
                "danger": "#f87171",
                "nav_selected": "#1f2937",
                "tree_selected": "#243247",
            }
        return {
            "bg": "#f5f6f8",
            "card_bg": "#ffffff",
            "text": "#0f172a",
            "muted": "#64748b",
            "border": "#e5e7eb",
            "blue": "#2563eb",
            "green": "#16a34a",
            "orange": "#f59e0b",
            "danger": "#c1121f",
            "nav_selected": "#fee2e2",
            "tree_selected": "#e0e7ff",
        }

    def show_balance_manager(self):
        """Show balance management system"""
        if not self._guard_module_access("accounting", "Balance Manager"):
            return
        if not BALANCE_MANAGER_AVAILABLE:
            messagebox.showerror("Error", "Balance Manager is not available")
            return
        try:
            if getattr(self, 'user_role', 'Staff') == 'Admin':
                from balance_manager import BalanceManagerDialog
                BalanceManagerDialog(self.root, self)
            else:
                show_balance_manager(self.root, self.manager)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open balance manager: {e}")
    
    def show_business_dashboard(self):
        """Show the Business Dashboard with Date Filter"""
        if not self._guard_module_access("dashboard", "Business Dashboard"):
            return
        try:
            dashboard = tk.Toplevel(self.root)
            dashboard.title("Business Dashboard")
            dashboard.geometry("1100x800")
            palette = self._theme_palette()
            bg = palette["bg"]
            card_bg = palette["card_bg"]
            text = palette["text"]
            muted = palette["muted"]
            blue = palette["blue"]
            danger = palette["danger"]
            try:
                dashboard.configure(bg=bg)
            except Exception:
                pass

            style = ttk.Style()
            try:
                style.configure("Dash.TFrame", background=bg)
                style.configure("DashCard.TLabelframe", background=card_bg, padding=15)
                style.configure("DashCard.TLabelframe.Label", background=card_bg, foreground=text, font=("Helvetica", 11, "bold"))
                style.configure("DashLabel.TLabel", background=bg, foreground=muted, font=("Helvetica", 10))
                style.configure("DashValue.TLabel", background=card_bg, foreground=blue, font=("Helvetica", 18, "bold"))
                style.configure("DashAlert.TLabel", background=card_bg, foreground=danger, font=("Helvetica", 18, "bold"))
                style.configure("DashHeader.TLabel", background=bg, foreground=text, font=("Helvetica", 20, "bold"))
            except Exception:
                pass

            # --- Filter Section ---
            filter_frame = ttk.Frame(dashboard, style='Dash.TFrame', padding=15)
            filter_frame.pack(fill='x')
            
            today = datetime.now().date()
            start_of_month = today.replace(day=1)
            
            ttk.Label(filter_frame, text="Period Start:", style='DashLabel.TLabel').pack(side='left', padx=5)
            start_var = tk.StringVar(value=start_of_month.strftime("%Y-%m-%d"))
            start_entry = ttk.Entry(filter_frame, textvariable=start_var, width=12)
            start_entry.pack(side='left', padx=5)
            
            ttk.Label(filter_frame, text="End:", style='DashLabel.TLabel').pack(side='left', padx=5)
            end_var = tk.StringVar(value=today.strftime("%Y-%m-%d"))
            end_entry = ttk.Entry(filter_frame, textvariable=end_var, width=12)
            end_entry.pack(side='left', padx=5)

            # --- Main Content ---
            main_frame = ttk.Frame(dashboard, style='Dash.TFrame', padding=20)
            main_frame.pack(fill='both', expand=True)
            
            # Header
            header_lbl = ttk.Label(main_frame, text="Dashboard", style="DashHeader.TLabel")
            header_lbl.pack(anchor='w', pady=(0, 20))
            
            # KPI Cards
            kpi_frame = ttk.Frame(main_frame, style='Dash.TFrame')
            kpi_frame.pack(fill='x', pady=(0, 20))
            
            # Placeholders for KPI labels
            sales_val = tk.StringVar()
            cost_val = tk.StringVar()
            profit_val = tk.StringVar()
            outstanding_val = tk.StringVar()
            count_val = tk.StringVar()
            stock_val = tk.StringVar()
            
            def create_card(parent, title, var):
                card = ttk.LabelFrame(parent, text=title, style='DashCard.TLabelframe')
                lbl = ttk.Label(card, textvariable=var, style='DashValue.TLabel')
                lbl.pack(anchor='w')
                return card, lbl

            # Row 1: Financials
            row1 = ttk.Frame(kpi_frame, style='Dash.TFrame')
            row1.pack(fill='x', pady=(0, 10))
            
            c1, l1 = create_card(row1, "Sales (Period)", sales_val)
            c1.pack(side='left', fill='both', expand=True, padx=(0, 10))
            
            c_cost, l_cost = create_card(row1, "Costs (Period)", cost_val)
            c_cost.pack(side='left', fill='both', expand=True, padx=(0, 10))
            
            c_profit, l_profit = create_card(row1, "Profit (Period)", profit_val)
            c_profit.pack(side='left', fill='both', expand=True)

            # Row 2: Operational
            row2 = ttk.Frame(kpi_frame, style='Dash.TFrame')
            row2.pack(fill='x')
            
            c2, l2 = create_card(row2, "Total Outstanding", outstanding_val)
            c2.pack(side='left', fill='both', expand=True, padx=(0, 10))
            
            c3, l3 = create_card(row2, "Invoices (Period)", count_val)
            c3.pack(side='left', fill='both', expand=True, padx=(0, 10))
            
            c4, l4 = create_card(row2, "Low Stock Alerts", stock_val)
            c4.pack(side='left', fill='both', expand=True)

            # Lists
            lists_frame = ttk.Frame(main_frame, style='Dash.TFrame')
            lists_frame.pack(fill='both', expand=True)
            
            # Overdue
            overdue_frame = ttk.LabelFrame(lists_frame, text="Top Overdue Invoices", style='DashCard.TLabelframe')
            overdue_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
            ov_tree = ttk.Treeview(overdue_frame, columns=('Invoice', 'Client', 'Amount', 'Days'), show='headings', height=10)
            for c in ('Invoice', 'Client', 'Amount', 'Days'):
                ov_tree.heading(c, text=c)
                ov_tree.column(c, width=90)
            ov_tree.pack(fill='both', expand=True)
            
            # Recent
            recent_frame = ttk.LabelFrame(lists_frame, text="Invoices in Period", style='DashCard.TLabelframe')
            recent_frame.pack(side='left', fill='both', expand=True)
            rc_tree = ttk.Treeview(recent_frame, columns=('Invoice', 'Date', 'Client', 'Total'), show='headings', height=10)
            for c in ('Invoice', 'Date', 'Client', 'Total'):
                rc_tree.heading(c, text=c)
                rc_tree.column(c, width=90)
            rc_tree.pack(fill='both', expand=True)

            # --- Logic ---
            def refresh():
                try:
                    s_str = start_var.get()
                    e_str = end_var.get()
                    s_date = datetime.strptime(s_str, "%Y-%m-%d").date()
                    e_date = datetime.strptime(e_str, "%Y-%m-%d").date()
                except:
                    messagebox.showerror("Error", "Invalid date format (YYYY-MM-DD)")
                    return

                # Fetch Data
                try:
                    all_invoices = self.manager.get_all_invoices_dict()
                except:
                    all_invoices = []
                
                period_sales = 0.0
                period_cost = 0.0
                period_profit = 0.0
                period_count = 0
                total_outstanding = 0.0
                
                filtered_invoices = []
                overdue_list = []
                
                for inv in all_invoices:
                    # Parse basics
                    try:
                        gt = float(inv.get('grand_total') or inv.get('total') or 0)
                    except: gt = 0.0
                    
                    paid = 0.0
                    # Try to get total_paid directly first
                    try:
                        paid = float(inv.get('total_paid', 0) or 0)
                    except: paid = 0.0
                    
                    # If 0, try to sum payments (fallback)
                    if paid == 0:
                        for p in inv.get('payments', []) or []:
                            try: paid += float(p.get('amount') or 0)
                            except: pass
                    
                    outstanding = max(gt - paid, 0)
                    
                    # Calculate Total Outstanding (Exclude Paid/Draft/Cancelled)
                    _status = str(inv.get('status', '')).lower()
                    _wf_status = str(inv.get('workflow_status', '')).lower()
                    _excluded = ('paid', 'closed', 'fully paid', 'cancelled', 'draft', 'void')
                    
                    if outstanding > 0 and _status not in _excluded and _wf_status not in _excluded:
                        total_outstanding += outstanding
                    
                    # Date check
                    inv_date_str = inv.get('date', '')
                    inv_date = None
                    try:
                        inv_date = datetime.strptime(inv_date_str, "%Y-%m-%d").date()
                    except: pass
                    
                    # Add to period stats
                    if inv_date and s_date <= inv_date <= e_date:
                        period_sales += gt
                        
                        # Cost and Profit
                        try:
                            c = float(inv.get('total_cost') or 0)
                        except: c = 0.0
                        
                        try:
                            p = float(inv.get('profit_loss') or 0)
                        except: p = 0.0
                        
                        # If profit is 0 but we have sales and cost, calc it
                        if p == 0 and gt > 0 and c > 0:
                            p = gt - c
                            
                        period_cost += c
                        period_profit += p
                        
                        period_count += 1
                        filtered_invoices.append(inv)
                        
                    # Overdue check (independent of period)
                    status = str(inv.get('status', '')).lower()
                    wf_status = str(inv.get('workflow_status', '')).lower()
                    if outstanding > 0 and status not in ('paid', 'closed', 'fully paid') and wf_status not in ('paid', 'closed', 'fully paid'):
                        due_str = inv.get('due_date', '')
                        if due_str:
                            try:
                                due = datetime.strptime(due_str, "%Y-%m-%d").date()
                                if due < datetime.now().date():
                                    overdue_list.append({
                                        'id': inv.get('invoice_id'),
                                        'client': inv.get('client_name'),
                                        'amount': outstanding,
                                        'days': (datetime.now().date() - due).days
                                    })
                            except: pass

                # Inventory
                low_stock = 0
                if INVENTORY_AVAILABLE:
                    try:
                        from inventory_system import InventoryManager
                        im = InventoryManager(self.manager.invoice_folder)
                        low_stock = len(im.get_low_stock_items())
                    except: pass

                # Update UI
                sales_val.set(f"AED {period_sales:,.2f}")
                cost_val.set(f"AED {period_cost:,.2f}")
                profit_val.set(f"AED {period_profit:,.2f}")
                outstanding_val.set(f"AED {total_outstanding:,.2f}")
                count_val.set(f"{period_count}")
                stock_val.set(f"{low_stock} Items")
                
                # Update Colors
                l2.configure(style='DashAlert.TLabel' if total_outstanding > 0 else 'DashValue.TLabel')
                l4.configure(style='DashAlert.TLabel' if low_stock > 0 else 'DashValue.TLabel')
                l_profit.configure(style='DashAlert.TLabel' if period_profit < 0 else 'DashValue.TLabel')
                
                # Populate Trees
                for item in ov_tree.get_children(): ov_tree.delete(item)
                overdue_list.sort(key=lambda x: x['days'], reverse=True)
                for item in overdue_list[:15]:
                    ov_tree.insert('', 'end', values=(item['id'], item['client'], f"{item['amount']:,.2f}", f"{item['days']} days"))
                    
                for item in rc_tree.get_children(): rc_tree.delete(item)
                filtered_invoices.sort(key=lambda x: x.get('date', ''), reverse=True)
                for inv in filtered_invoices:
                    try:
                        rc_tree.insert('', 'end', values=(
                            inv.get('invoice_id'),
                            inv.get('date'),
                            inv.get('client_name'),
                            f"{float(inv.get('grand_total') or 0):,.2f}"
                        ))
                    except: pass
                    
            ttk.Button(filter_frame, text="Apply Filter", command=refresh, style='Accent.TButton').pack(side='left', padx=10)
            
            # Initial call
            refresh()
            _fit_window(dashboard, 1100, 800, mode="workspace", remember_key="business_dashboard")
            
        except Exception as e:
            messagebox.showerror("Dashboard Error", f"Could not load dashboard: {e}")

    def create_main_menu(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        try:
            self._ensure_app_menu()
        except Exception:
            pass

        try:
            self._apply_saas_theme()
        except Exception:
            pass

        if not hasattr(self, "sidebar_collapsed"):
            self.sidebar_collapsed = False

        root_wrap = ttk.Frame(self.root, style="App.TFrame")
        root_wrap.pack(fill="both", expand=True)
        root_wrap.columnconfigure(1, weight=1)
        root_wrap.rowconfigure(0, weight=1)

        sidebar_width = 78 if self.sidebar_collapsed else 270
        sidebar = ttk.Frame(root_wrap, style="Sidebar.TFrame", width=sidebar_width)
        sidebar.grid(row=0, column=0, sticky="nsw")
        sidebar.grid_propagate(False)

        content = ttk.Frame(root_wrap, style="App.TFrame")
        content.grid(row=0, column=1, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        side_top = ttk.Frame(sidebar, style="Sidebar.TFrame")
        side_top.pack(fill="x", padx=10, pady=(10, 8))

        toggle_text = "☰" if self.sidebar_collapsed else "☰ Menu"
        ttk.Button(side_top, text=toggle_text, style="SidebarToggle.TButton", command=self._toggle_sidebar).pack(fill="x")

        if self.logo_manager.load_logo(size=(48, 48)):
            ttk.Label(side_top, image=self.logo_manager.logo_photo, style="SidebarLogo.TLabel").pack(anchor="w", pady=(10, 0))

        if not self.sidebar_collapsed:
            ttk.Label(side_top, text="Hope Pharma", style="SidebarTitle.TLabel").pack(anchor="w", pady=(6, 0))

        nav_frame = ttk.Frame(sidebar, style="Sidebar.TFrame")
        nav_frame.pack(fill="both", expand=True, padx=8, pady=(0, 8))
        nav_frame.columnconfigure(0, weight=1)
        nav_frame.rowconfigure(0, weight=1)

        self._nav_callbacks = {}
        self._nav_tree = ttk.Treeview(nav_frame, show="tree", selectmode="browse", style="Nav.Treeview", height=20)
        self._nav_tree.grid(row=0, column=0, sticky="nsew")
        nav_scroll = ttk.Scrollbar(nav_frame, orient="vertical", command=self._nav_tree.yview)
        nav_scroll.grid(row=0, column=1, sticky="ns")
        self._nav_tree.configure(yscrollcommand=nav_scroll.set)

        self._nav_item_keys = {}
        for group_label, children in self._sidebar_menu_structure():
            default_open = str(group_label).startswith("📊 Dashboard")
            group_id = self._nav_tree.insert("", "end", text=group_label, open=default_open, tags=("group",))
            for item in children:
                item_id = self._nav_tree.insert(group_id, "end", text=item["label"], tags=("item",))
                self._nav_callbacks[item_id] = item["command"]
                self._nav_item_keys[item_id] = str(item.get("key") or "")

        self._nav_tree.bind("<<TreeviewSelect>>", self._on_nav_select)
        self._nav_tree.bind("<Double-1>", self._on_nav_select)
        self._nav_tree.bind("<Button-3>", self._on_nav_right_click)
        self._nav_tree.bind("<Button-2>", self._on_nav_right_click)

        topbar = ttk.Frame(content, style="Topbar.TFrame")
        topbar.grid(row=0, column=0, sticky="ew")
        topbar.columnconfigure(1, weight=1)

        left_top = ttk.Frame(topbar, style="Topbar.TFrame")
        left_top.grid(row=0, column=0, sticky="w", padx=14, pady=10)
        ttk.Label(left_top, text="Dashboard", style="TopbarTitle.TLabel").pack(side="left")

        center_top = ttk.Frame(topbar, style="Topbar.TFrame")
        center_top.grid(row=0, column=1, sticky="ew", padx=14, pady=10)
        center_top.columnconfigure(0, weight=1)

        self.global_search_var = tk.StringVar()
        self.global_search_entry = ttk.Entry(center_top, textvariable=self.global_search_var, style="GlobalSearchPlaceholder.TEntry")
        self.global_search_entry.grid(row=0, column=0, sticky="ew")
        self.global_search_entry.bind("<KeyRelease>", self._on_global_search_change)
        self.global_search_entry.bind("<Return>", self._open_first_global_search_result)
        self.global_search_entry.bind("<Escape>", self._hide_global_search_results)
        self.global_search_entry.bind("<FocusIn>", self._on_global_search_focus_in)
        self.global_search_entry.bind("<FocusOut>", self._on_global_search_focus_out)
        self._set_global_search_placeholder()

        ttk.Button(center_top, text="+ New", command=self._open_quick_create_menu, style="Primary.TButton").grid(row=0, column=1, sticky="e", padx=(10, 0))

        right_top = ttk.Frame(topbar, style="Topbar.TFrame")
        right_top.grid(row=0, column=2, sticky="e", padx=14, pady=10)

        if not hasattr(self, "user_role"):
            self.user_role = "Staff"
        ttk.Label(right_top, text=f"Role: {self.user_role}", style="TopbarMeta.TLabel").pack(side="left", padx=(0, 10))
        try:
            company_label_text = f"Company: {getattr(self, 'current_company', 'HopePharma')}"
        except Exception:
            company_label_text = "Company: HopePharma"
        ttk.Label(right_top, text=company_label_text, style="TopbarMeta.TLabel").pack(side="left", padx=(0, 14))

        mode = str(getattr(self, "ui_theme_mode", "light") or "light").strip().lower()
        theme_btn_text = "🌙" if mode != "dark" else "☀️"
        ttk.Button(right_top, text=theme_btn_text, command=self.toggle_theme_mode, style="TopbarIcon.TButton").pack(side="left", padx=(0, 6))

        self.notif_btn = ttk.Button(right_top, text="Notifications (0)", command=self.show_notifications_dialog, style="TopbarPill.TButton")
        self.notif_btn.pack(side="left", padx=(0, 12))

        ttk.Button(right_top, text="Login as Admin", command=self.login_admin, style="Topbar.TButton").pack(side="left", padx=4)
        ttk.Button(right_top, text="Change Password", command=self._show_change_password_dialog, style="Topbar.TButton").pack(side="left", padx=4)
        ttk.Button(right_top, text="Switch Company", command=self.logout_company, style="Topbar.TButton").pack(side="left", padx=4)
        ttk.Button(right_top, text="Admin Logout", command=self.logout_admin, style="TopbarDanger.TButton").pack(side="left", padx=(4, 0))

        body = ttk.Frame(content, style="App.TFrame")
        body.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 14))
        body.columnconfigure(0, weight=1)
        body.rowconfigure(0, weight=1)

        self._build_home_overview(body)

        try:
            self.update_notifications_badge()
        except Exception:
            pass

        _fit_window(self.root, 1200, 800, mode="workspace", remember_key="main_window")

    def _toggle_sidebar(self):
        try:
            self.sidebar_collapsed = not bool(getattr(self, "sidebar_collapsed", False))
        except Exception:
            self.sidebar_collapsed = False
        try:
            self._save_ui_prefs()
        except Exception:
            pass
        self.create_main_menu()

    def _on_nav_select(self, _event=None):
        try:
            sel = self._nav_tree.selection()
            if not sel:
                return
            item_id = sel[0]
            cmd = self._nav_callbacks.get(item_id)
            if cmd:
                cmd()
        except Exception:
            return

    def _on_nav_right_click(self, event):
        try:
            row_id = self._nav_tree.identify_row(event.y)
            if not row_id:
                return
            if self._nav_tree.parent(row_id) == "":
                return
            self._nav_tree.selection_set(row_id)
            key = str((getattr(self, "_nav_item_keys", {}) or {}).get(row_id) or "").strip()
            if not key:
                return
            favs = list(getattr(self, "ui_favorites", []) or [])
            is_fav = key in favs
            menu = tk.Menu(self.root, tearoff=0)
            if is_fav:
                menu.add_command(label="Remove from Favorites", command=lambda: self._remove_favorite(key))
            else:
                menu.add_command(label="Add to Favorites", command=lambda: self._add_favorite(key))
            menu.tk_popup(event.x_root, event.y_root)
        except Exception:
            return

    def _add_favorite(self, key):
        favs = list(getattr(self, "ui_favorites", []) or [])
        k = str(key or "").strip()
        if not k:
            return
        if k not in favs:
            favs.append(k)
        self.ui_favorites = favs
        try:
            self._save_ui_prefs()
        except Exception:
            pass
        self.create_main_menu()

    def _remove_favorite(self, key):
        favs = [x for x in (getattr(self, "ui_favorites", []) or []) if str(x) != str(key)]
        self.ui_favorites = favs
        try:
            self._save_ui_prefs()
        except Exception:
            pass
        self.create_main_menu()

    def _open_quick_create_menu(self, event=None):
        try:
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Invoice", command=self._qc_new_invoice)
            menu.add_command(label="Purchase", command=self._qc_new_purchase)
            menu.add_command(label="Client", command=self._qc_new_client)
            menu.add_command(label="Product", command=self._qc_new_product)
            menu.add_command(label="Employee", command=self._qc_new_employee)
            menu.add_command(label="Supplier", command=self._qc_new_supplier)
            x = self.root.winfo_pointerx()
            y = self.root.winfo_pointery()
            menu.tk_popup(x, y)
        except Exception:
            return

    def _qc_new_invoice(self):
        try:
            CreateInvoiceDialog(self.root, self.manager, self.data_memory)
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open Create Invoice:\n{exc}")

    def open_cost_center_and_backfill(self):
        """Hub dialog: Manage Cost Centers + Bulk Backfill past invoices.
        User explicitly requested Cost Centers live HERE (InvoiceManager dashboard),
        NOT inside individual Create/Edit Invoice dialogs.
        """
        try:
            from cost_center_ui import open_cost_center_and_backfill as _runner
            folder = (
                getattr(self.manager, "invoice_folder", None)
                or getattr(self.manager, "data_folder", None)
                or os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")
            )
            _runner(self.root, self.manager, folder)
            # Refresh dashboard KPI/tables after user closes (costs changed)
            try:
                if hasattr(self, "render_dashboard") and callable(self.render_dashboard):
                    self.render_dashboard()
                elif hasattr(self, "refresh_invoice_table") and callable(self.refresh_invoice_table):
                    self.refresh_invoice_table()
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror(
                "Cost Centers Error",
                f"Could not open Cost Centers & Bulk Backfill:\n{exc}"
            )

    def _qc_new_purchase(self):
        try:
            CreatePurchaseDialog(self.root, self.manager)
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open Create Purchase:\n{exc}")

    def _qc_new_client(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("New Client")
        dlg.geometry("420x220")
        dlg.transient(self.root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Client Name *", style="TopbarMeta.TLabel").pack(anchor="w")
        name_var = tk.StringVar()
        e = ttk.Entry(frm, textvariable=name_var)
        e.pack(fill="x", pady=(6, 12))
        e.focus_set()
        def save():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Warning", "Client Name is required")
                return
            try:
                self.data_memory.add_client(name)
            except Exception:
                pass
            dlg.destroy()
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Save", command=save, style="Primary.TButton").pack(side="left")
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        dlg.bind("<Return>", lambda _e: save())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        _fit_window(dlg, 420, 220, mode="compact", remember_key="quick_client")

    def _qc_new_product(self):
        try:
            from inventory_system import InventoryManager, AddItemDialog
            folder = getattr(self.manager, "invoice_folder", None) or getattr(self.manager, "data_folder", None) or os.getcwd()
            AddItemDialog(self.root, InventoryManager(folder))
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open Add Product:\n{exc}")

    def _qc_new_employee(self):
        try:
            dlg = LegacyEmployeeManagerDialog(self.root, self.manager)
            try:
                dlg.add_employee()
            except Exception:
                pass
        except Exception as exc:
            messagebox.showerror("Error", f"Could not open Add Employee:\n{exc}")

    def _qc_new_supplier(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("New Supplier")
        dlg.geometry("480x260")
        dlg.transient(self.root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Supplier Name *", style="TopbarMeta.TLabel").pack(anchor="w")
        name_var = tk.StringVar()
        e = ttk.Entry(frm, textvariable=name_var)
        e.pack(fill="x", pady=(6, 12))
        e.focus_set()
        def save():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Warning", "Supplier Name is required")
                return
            try:
                from inventory_system import InventoryManager
                folder = getattr(self.manager, "invoice_folder", None) or getattr(self.manager, "data_folder", None) or os.getcwd()
                inv = InventoryManager(folder)
                ok, msg, _rec = inv.add_supplier_record(name)
                if not ok:
                    messagebox.showwarning("Warning", msg)
                    return
            except Exception:
                pass
            dlg.destroy()
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Save", command=save, style="Primary.TButton").pack(side="left")
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        dlg.bind("<Return>", lambda _e: save())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        _fit_window(dlg, 480, 260, mode="compact", remember_key="quick_supplier")

    def _global_search_placeholder_text(self):
        return "🔍 Search..."

    def _is_global_search_placeholder(self):
        try:
            return str(self.global_search_var.get() or "") == self._global_search_placeholder_text()
        except Exception:
            return False

    def _set_global_search_placeholder(self):
        try:
            if not str(self.global_search_var.get() or "").strip():
                self.global_search_var.set(self._global_search_placeholder_text())
                try:
                    self.global_search_entry.configure(style="GlobalSearchPlaceholder.TEntry")
                except Exception:
                    pass
        except Exception:
            return

    def _clear_global_search_placeholder(self):
        try:
            if self._is_global_search_placeholder():
                self.global_search_var.set("")
            try:
                self.global_search_entry.configure(style="GlobalSearch.TEntry")
            except Exception:
                pass
        except Exception:
            return

    def _on_global_search_focus_in(self, event=None):
        try:
            if self._is_global_search_placeholder():
                self._clear_global_search_placeholder()
        except Exception:
            return

    def _on_global_search_focus_out(self, event=None):
        try:
            if not str(self.global_search_var.get() or "").strip():
                self._set_global_search_placeholder()
        except Exception:
            return

    def _on_global_search_change(self, event=None):
        q = ""
        try:
            q = str(self.global_search_var.get() or "").strip()
        except Exception:
            q = ""
        if self._is_global_search_placeholder():
            self._hide_global_search_results()
            return
        if not q:
            self._hide_global_search_results()
            return
        self._show_global_search_results(q)

    def _hide_global_search_results(self, event=None):
        try:
            if hasattr(self, "_search_popup") and self._search_popup and self._search_popup.winfo_exists():
                self._search_popup.destroy()
        except Exception:
            pass
        self._search_popup = None
        self._search_results = []

    def _open_first_global_search_result(self, event=None):
        try:
            if self._is_global_search_placeholder():
                return
            results = list(getattr(self, "_search_results", []) or [])
            if not results:
                return
            self._open_global_search_result(results[0])
        finally:
            self._hide_global_search_results()

    def _show_global_search_results(self, query):
        q = str(query or "").strip().lower()
        results = []
        for rec in self._build_global_search_index():
            hay = " ".join([str(rec.get("title", "")), str(rec.get("subtitle", "")), str(rec.get("extra", ""))]).lower()
            if q in hay:
                results.append(rec)
            if len(results) >= 20:
                break
        self._search_results = results
        if not results:
            self._hide_global_search_results()
            return

        try:
            if hasattr(self, "_search_popup") and self._search_popup and self._search_popup.winfo_exists():
                self._search_popup.destroy()
        except Exception:
            pass

        popup = tk.Toplevel(self.root)
        popup.wm_overrideredirect(True)
        popup.attributes("-topmost", True)
        self._search_popup = popup

        try:
            x = self.global_search_entry.winfo_rootx()
            y = self.global_search_entry.winfo_rooty() + self.global_search_entry.winfo_height()
            w = max(520, self.global_search_entry.winfo_width() + 120)
        except Exception:
            x, y, w = 200, 120, 520
        popup.geometry(f"{w}x320+{x}+{y}")

        frame = ttk.Frame(popup, padding=6, style="CardBody.TFrame")
        frame.pack(fill="both", expand=True)

        cols = ("Type", "Result", "Info")
        tree = ttk.Treeview(frame, columns=cols, show="headings", height=10)
        for c in cols:
            tree.heading(c, text=c)
        tree.column("Type", width=120, anchor="w")
        tree.column("Result", width=260, anchor="w")
        tree.column("Info", width=220, anchor="w")
        tree.pack(fill="both", expand=True)
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        yscroll.place(relx=1.0, rely=0.0, relheight=1.0, anchor="ne")
        tree.configure(yscrollcommand=yscroll.set)

        for idx, rec in enumerate(results):
            tree.insert("", "end", iid=str(idx), values=(rec.get("type", ""), rec.get("title", ""), rec.get("subtitle", "")))

        def open_selected(_e=None):
            try:
                sel = tree.selection()
                if not sel:
                    return
                rec = results[int(sel[0])]
                self._open_global_search_result(rec)
            except Exception:
                pass
            self._hide_global_search_results()

        tree.bind("<Double-1>", open_selected)
        tree.bind("<Return>", open_selected)
        popup.bind("<Escape>", self._hide_global_search_results)
        try:
            tree.selection_set("0")
            tree.focus("0")
        except Exception:
            pass

    def _build_global_search_index(self):
        now_ts = time.time() if "time" in globals() else None
        try:
            cache = getattr(self, "_search_cache", None) or {}
            ts = cache.get("ts")
            if ts and now_ts and (now_ts - ts) < 20 and cache.get("rows"):
                return cache["rows"]
        except Exception:
            pass

        rows = []
        try:
            invoices = self.manager.get_all_invoices_dict()
        except Exception:
            invoices = []
        for inv in invoices or []:
            invoice_id = str(inv.get("invoice_id") or inv.get("id") or "").strip()
            if not invoice_id:
                continue
            client = str(inv.get("client_name") or "").strip()
            total = inv.get("grand_total", "")
            rows.append({
                "type": "Invoice",
                "title": invoice_id,
                "subtitle": client,
                "extra": str(total),
                "open": ("invoice", invoice_id),
            })

        try:
            purchases = self.manager.get_purchases() if hasattr(self.manager, "get_purchases") else self.manager.load_purchases()
        except Exception:
            purchases = []
        for pur in purchases or []:
            pid = str(pur.get("purchase_id") or pur.get("id") or "").strip()
            if not pid:
                continue
            desc = str(pur.get("description") or pur.get("supplier") or "").strip()
            amt = pur.get("amount", "")
            rows.append({
                "type": "Purchase Order",
                "title": pid,
                "subtitle": desc,
                "extra": str(amt),
                "open": ("purchase", pid),
            })

        try:
            folder = getattr(self.manager, "invoice_folder", None) or getattr(self.manager, "data_folder", None) or os.getcwd()
            from inventory_system import InventoryManager
            invm = InventoryManager(folder)
            for item in invm.items:
                code = str(getattr(item, "item_id", "") or getattr(item, "code", "") or "").strip()
                name = str(getattr(item, "name", "") or "").strip()
                batch = str(getattr(item, "batch_number", "") or "").strip()
                if not (code or name):
                    continue
                rows.append({
                    "type": "Product",
                    "title": code or name,
                    "subtitle": name if code else batch,
                    "extra": batch,
                    "open": ("product", code),
                })
                if batch:
                    rows.append({
                        "type": "Batch Number",
                        "title": batch,
                        "subtitle": f"{name} ({code})" if name and code else (name or code),
                        "extra": code,
                        "open": ("product", code),
                    })
        except Exception:
            pass

        try:
            clients = []
            try:
                clients = self.data_memory.load_clients()
            except Exception:
                clients = []
            for c in clients or []:
                name = str(c or "").strip()
                if name:
                    rows.append({"type": "Client", "title": name, "subtitle": "", "extra": "", "open": ("client", name)})
        except Exception:
            pass

        try:
            em = EmployeeManager(getattr(self.manager, "invoice_folder", os.getcwd()))
            for emp in em.load_employees():
                emp_id = str(emp.get("employee_id") or "").strip()
                emp_name = str(emp.get("name") or "").strip()
                if not emp_name and not emp_id:
                    continue
                rows.append({"type": "Employee", "title": emp_name or emp_id, "subtitle": emp_id, "extra": emp.get("phone", ""), "open": ("employee", emp_id)})
        except Exception:
            pass

        try:
            self._search_cache = {"ts": now_ts or 0, "rows": rows}
        except Exception:
            pass
        return rows

    def _open_global_search_result(self, rec):
        kind, ident = rec.get("open", (None, None))
        if kind == "invoice":
            try:
                inv = self.manager.get_invoice(ident)
                inv_data = inv.to_dict() if hasattr(inv, "to_dict") else inv
                ViewInvoiceDialog(self.root, self.manager, ident, inv_data)
            except Exception as exc:
                messagebox.showerror("Error", f"Could not open invoice:\n{exc}")
            return
        if kind == "purchase":
            try:
                pur = self.manager.get_purchase(ident) if hasattr(self.manager, "get_purchase") else None
                if not pur:
                    purchases = self.manager.load_purchases()
                    pur = next((p for p in purchases if str(p.get("purchase_id")) == ident), None)
                ViewPurchaseDialog(self.root, self.manager, ident, pur or {})
            except Exception as exc:
                messagebox.showerror("Error", f"Could not open purchase:\n{exc}")
            return
        if kind == "client":
            try:
                self._open_client_detail(ident)
            except Exception:
                self.show_clients()
            return
        if kind == "product":
            try:
                self._open_product_detail(ident)
            except Exception:
                self.show_inventory_system()
            return
        if kind == "employee":
            try:
                self._open_employee_detail(ident)
            except Exception:
                self.show_employee_manager()
            return

    def _open_client_detail(self, client_name):
        name = str(client_name or "").strip()
        if not name:
            return
        dlg = tk.Toplevel(self.root)
        dlg.title(f"Client: {name}")
        dlg.geometry("900x650")
        dlg.transient(self.root)
        dlg.resizable(True, True)
        main = ttk.Frame(dlg, padding=12)
        main.pack(fill="both", expand=True)
        header = ttk.Frame(main)
        header.pack(fill="x")
        ttk.Label(header, text=name, font=("Helvetica", 16, "bold")).pack(side="left")
        ttk.Button(header, text="Open Clients Module", command=self.show_clients).pack(side="right")

        invoices = []
        try:
            invoices = self.manager.get_all_invoices_dict()
        except Exception:
            invoices = []
        rows = [inv for inv in invoices if str(inv.get("client_name") or "").strip().lower() == name.lower()]

        total_sales = 0.0
        total_paid = 0.0
        total_outstanding = 0.0
        for inv in rows:
            try:
                total_sales += float(inv.get("grand_total") or 0)
            except Exception:
                pass
            try:
                total_paid += float(inv.get("total_paid") or 0)
            except Exception:
                pass
            try:
                total_outstanding += float(inv.get("balance_due") or 0)
            except Exception:
                pass

        stats = ttk.Frame(main)
        stats.pack(fill="x", pady=(10, 10))
        for i in range(3):
            stats.columnconfigure(i, weight=1, uniform="c")
        def _kpi(col, label, value):
            card = ttk.LabelFrame(stats, text=label, style="CardInfo.TLabelframe")
            card.grid(row=0, column=col, sticky="nsew", padx=(0 if col == 0 else 10, 0))
            ttk.Label(card, text=value, style="KPIValue.TLabel").pack(anchor="w", padx=12, pady=(10, 12))
        _kpi(0, "Total Sales", f"AED {total_sales:,.2f}")
        _kpi(1, "Total Paid", f"AED {total_paid:,.2f}")
        _kpi(2, "Outstanding", f"AED {total_outstanding:,.2f}")

        table_frame = ttk.LabelFrame(main, text="Invoices", style="Card.TLabelframe")
        table_frame.pack(fill="both", expand=True)
        cols = ("Invoice ID", "Date", "Status", "Total", "Paid", "Balance")
        tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        for c in cols:
            tree.heading(c, text=c)
        tree.column("Invoice ID", width=140)
        tree.column("Date", width=120)
        tree.column("Status", width=140)
        tree.column("Total", width=120, anchor="e")
        tree.column("Paid", width=120, anchor="e")
        tree.column("Balance", width=120, anchor="e")
        tree.pack(fill="both", expand=True)
        y = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        y.place(relx=1.0, rely=0.0, relheight=1.0, anchor="ne")
        tree.configure(yscrollcommand=y.set)

        def _safe(v):
            try:
                return f"{float(v or 0):,.2f}"
            except Exception:
                return str(v or "")

        for inv in sorted(rows, key=lambda r: str(r.get("date", "")), reverse=True):
            invoice_id = str(inv.get("invoice_id") or inv.get("id") or "").strip()
            tree.insert("", "end", iid=invoice_id, values=(
                invoice_id,
                inv.get("date", ""),
                inv.get("status", ""),
                _safe(inv.get("grand_total")),
                _safe(inv.get("total_paid")),
                _safe(inv.get("balance_due")),
            ))

        def open_invoice(_e=None):
            sel = tree.selection()
            if not sel:
                return
            invoice_id = sel[0]
            try:
                inv_obj = self.manager.get_invoice(invoice_id)
                inv_data = inv_obj.to_dict() if hasattr(inv_obj, "to_dict") else inv_obj
                ViewInvoiceDialog(dlg, self.manager, invoice_id, inv_data)
            except Exception:
                return

        tree.bind("<Double-1>", open_invoice)
        tree.bind("<Return>", open_invoice)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

    def _open_product_detail(self, item_code_or_id):
        ident = str(item_code_or_id or "").strip()
        if not ident:
            return
        try:
            from inventory_system import InventoryManager
            folder = getattr(self.manager, "invoice_folder", None) or getattr(self.manager, "data_folder", None) or os.getcwd()
            invm = InventoryManager(folder)
        except Exception:
            return
        item = None
        for it in invm.items:
            code = str(getattr(it, "item_id", "") or getattr(it, "code", "") or "").strip()
            if code and code == ident:
                item = it
                break
        if not item:
            return
        dlg = tk.Toplevel(self.root)
        dlg.title("Product")
        dlg.geometry("640x520")
        dlg.transient(self.root)
        dlg.resizable(True, True)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=str(getattr(item, "name", "") or ""), font=("Helvetica", 16, "bold")).pack(anchor="w")
        ttk.Label(frm, text=str(getattr(item, "item_id", "") or ""), style="TopbarMeta.TLabel").pack(anchor="w", pady=(2, 12))
        grid = ttk.Frame(frm)
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)
        fields = [
            ("Code", getattr(item, "item_id", "")),
            ("Batch Number", getattr(item, "batch_number", "")),
            ("Manufacture Date", getattr(item, "manufacture_date", "")),
            ("Expiry Date", getattr(item, "expiry_date", "")),
            ("Quantity", getattr(item, "quantity", "")),
            ("Min Stock", getattr(item, "min_stock", "")),
            ("Package Type", getattr(item, "package_type", "")),
            ("Brand", getattr(item, "brand", "")),
            ("Client Name", getattr(item, "client_name", "")),
            ("Supplier", getattr(item, "supplier", "")),
        ]
        for r, (label, val) in enumerate(fields):
            ttk.Label(grid, text=label).grid(row=r, column=0, sticky="w", pady=4, padx=(0, 10))
            ttk.Label(grid, text=str(val or "")).grid(row=r, column=1, sticky="w", pady=4)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="Open Inventory System", command=self.show_inventory_system, style="Primary.TButton").pack(side="left")
        ttk.Button(btns, text="Close", command=dlg.destroy).pack(side="left", padx=8)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

    def _open_employee_detail(self, employee_id):
        ident = str(employee_id or "").strip()
        if not ident:
            return
        em = EmployeeManager(getattr(self.manager, "invoice_folder", os.getcwd()))
        emp = None
        for e in em.load_employees():
            if str(e.get("employee_id") or "").strip() == ident:
                emp = e
                break
        if not emp:
            return
        dlg = tk.Toplevel(self.root)
        dlg.title("Employee")
        dlg.geometry("640x520")
        dlg.transient(self.root)
        dlg.resizable(True, True)
        frm = ttk.Frame(dlg, padding=14)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text=str(emp.get("name") or ""), font=("Helvetica", 16, "bold")).pack(anchor="w")
        ttk.Label(frm, text=str(emp.get("employee_id") or ""), style="TopbarMeta.TLabel").pack(anchor="w", pady=(2, 12))
        grid = ttk.Frame(frm)
        grid.pack(fill="x")
        grid.columnconfigure(1, weight=1)
        fields = [
            ("Email", emp.get("email", "")),
            ("Phone", emp.get("phone", "")),
            ("Salary", f"AED {float(emp.get('salary') or 0):,.2f}" if str(emp.get("salary") or "").strip() else ""),
            ("Start Date", emp.get("start_date", "")),
            ("End Date", emp.get("end_date", "")),
            ("Residency Expiry", emp.get("residency_expiry", "")),
            ("Status", emp.get("status", "")),
        ]
        for r, (label, val) in enumerate(fields):
            ttk.Label(grid, text=label).grid(row=r, column=0, sticky="w", pady=4, padx=(0, 10))
            ttk.Label(grid, text=str(val or "")).grid(row=r, column=1, sticky="w", pady=4)
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(18, 0))
        ttk.Button(btns, text="Open Employee Manager", command=self.show_employee_manager, style="Primary.TButton").pack(side="left")
        ttk.Button(btns, text="Close", command=dlg.destroy).pack(side="left", padx=8)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

    def _sidebar_menu_structure(self):
        items = {
            "business_dashboard": {"key": "business_dashboard", "module_key": "dashboard", "label": "📊 Business Dashboard", "command": self.show_business_dashboard},
            "balance_manager": {"key": "balance_manager", "module_key": "accounting", "label": "💰 Balance Manager", "command": self.show_balance_manager},
            "transaction_manager": {"key": "transaction_manager", "module_key": "accounting", "label": "💼 Transaction Manager", "command": self.show_transaction_manager},
            "gaap_dashboard": {"key": "gaap_dashboard", "module_key": "reports", "label": "📈 GAAP Financial Reports", "command": self.show_gaap_dashboard},
            "inventory_system": {"key": "inventory_system", "module_key": "inventory", "label": "📦 Inventory", "command": self.show_inventory_system},
            "warehouse_management": {"key": "warehouse_management", "module_key": "warehouse_management", "label": "🏬 Warehouse Management", "command": self.show_warehouse_management},
            "inventory_report": {"key": "inventory_report", "module_key": "reports", "label": "📑 Inventory Reports", "command": self.show_inventory_report},
            "temperature_log": {"key": "temperature_log", "module_key": "temperature_log", "label": "🌡️ Temperature Log", "command": self.show_temperature_log},
            "quotations": {"key": "quotations", "module_key": "invoice_management", "label": "📋 Quotations", "command": self.show_quotations},
            "invoice_management": {"key": "invoice_management", "module_key": "invoice_management", "label": "🧾 Invoice Management", "command": self.show_invoice_management},
            "clients": {"key": "clients", "module_key": "clients", "label": "👤 Clients", "command": self.show_clients},
            "purchases": {"key": "purchases", "module_key": "purchases", "label": "🛒 Purchases", "command": self.show_purchases},
            "reports": {"key": "reports", "module_key": "reports", "label": "📊 Legacy Reports", "command": self.show_reports},
            "ecommerce": {"key": "ecommerce", "module_key": "ecommerce_marketplace", "label": "🌐 E-Commerce Marketplaces", "command": self.show_marketplace_manager},
            "employee_manager": {"key": "employee_manager", "module_key": "employee_manager", "label": "👥 Employee Manager", "command": self.show_employee_manager},
            "company_settings": {"key": "company_settings", "label": "🏢 Company Settings", "command": self._open_company_settings},
            "subscription": {"key": "subscription", "label": "🪪 Subscription", "command": self.show_subscription_details},
            "exit": {"key": "exit", "label": "🚪 Exit", "command": self.root.quit},
        }

        groups = [
            ("📊 Dashboard", [items["business_dashboard"]]),
            ("💰 Finance", [items["balance_manager"], items["transaction_manager"], items["gaap_dashboard"]]),
            ("📦 Warehouse", [items["inventory_system"], items["warehouse_management"], items["inventory_report"]]),
            ("🏥 Quality Assurance", [items["temperature_log"]]),
            ("💼 Sales & CRM", [items["quotations"], items["invoice_management"], items["clients"]]),
            ("🚚 Procurement", [items["purchases"]]),
            ("📈 Analytics & Reports", [items["gaap_dashboard"], items["reports"]]),
            ("🛍️ E-Commerce", [items["ecommerce"]]),
            ("👥 Human Resources", [items["employee_manager"]]),
            ("⚙️ System Administration", [items["company_settings"], items["subscription"], items["exit"]]),
        ]
        filtered_groups = []
        for group_label, children in groups:
            visible_children = []
            for child in children:
                module_key = str(child.get("module_key") or "").strip()
                if not module_key or self._is_module_enabled(module_key, True):
                    visible_children.append(child)
            if visible_children:
                filtered_groups.append((group_label, visible_children))
        groups = filtered_groups

        fav_keys = list(getattr(self, "ui_favorites", []) or [])
        fav_set = {str(x or "").strip() for x in fav_keys if str(x or "").strip()}
        if fav_set:
            for k in fav_set:
                try:
                    if k in items:
                        lbl = str(items[k].get("label") or "")
                        if "★" not in lbl:
                            items[k]["label"] = f"{lbl} ★"
                except Exception:
                    pass
        if fav_keys:
            fav_children = []
            for k in fav_keys:
                k = str(k or "").strip()
                if not (k and k in items):
                    continue
                module_key = str(items[k].get("module_key") or "").strip()
                if module_key and not self._is_module_enabled(module_key, True):
                    continue
                if items[k] not in fav_children:
                    fav_children.append(items[k])
            if fav_children:
                groups = [("⭐ Favorites", fav_children)] + groups

        return groups

    def _build_home_overview(self, parent):
        for w in parent.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass

        sf = ScrollableFrame(parent)
        sf.pack(fill="both", expand=True)
        container = sf.scrollable_frame

        now = datetime.now()
        hour = now.hour
        greeting = "Good Morning" if hour < 12 else ("Good Afternoon" if hour < 17 else "Good Evening")
        role = str(getattr(self, "user_role", "User") or "User").strip()
        ui_mode = str(getattr(self, "ui_theme_mode", "light") or "light").strip().lower()
        if ui_mode == "dark":
            lb_bg = "#111827"
            lb_fg = "#e5e7eb"
            lb_sel_bg = "#243247"
            lb_sel_fg = "#e5e7eb"
            lb_border = "#1f2937"
        else:
            lb_bg = "#ffffff"
            lb_fg = "#0f172a"
            lb_sel_bg = "#e0e7ff"
            lb_sel_fg = "#0f172a"
            lb_border = "#e5e7eb"

        header = ttk.Frame(container, style="App.TFrame")
        header.pack(fill="x", pady=(0, 12))
        ttk.Label(header, text=f"{greeting}, {role}", style="PageTitle.TLabel").pack(side="left")

        def _to_float(v):
            try:
                return float(v or 0)
            except Exception:
                return 0.0

        today = now.strftime("%Y-%m-%d")
        yesterday = (now - timedelta(days=1)).strftime("%Y-%m-%d")

        try:
            invoices = self.manager.get_all_invoices_dict() or []
        except Exception:
            invoices = []
        try:
            purchases = self.manager.get_all_purchases_dict() if hasattr(self.manager, "get_all_purchases_dict") else (self.manager.load_purchases() if hasattr(self.manager, "load_purchases") else [])
        except Exception:
            purchases = []

        today_sales_total = sum(_to_float(i.get("grand_total")) for i in invoices if str(i.get("date") or "") == today)
        yesterday_sales_total = sum(_to_float(i.get("grand_total")) for i in invoices if str(i.get("date") or "") == yesterday)
        today_sales_count = sum(1 for i in invoices if str(i.get("date") or "") == today)
        yesterday_sales_count = sum(1 for i in invoices if str(i.get("date") or "") == yesterday)

        purchases_today_total = sum(_to_float(p.get("amount")) for p in purchases if str(p.get("date") or "") == today)
        outstanding_receivables = sum(
            max(0.0, _to_float(i.get("grand_total")) - _to_float(i.get("total_paid")))
            for i in invoices
        )
        outstanding_payables = sum(_to_float(p.get("amount")) for p in purchases if not bool(p.get("paid", False)))

        inv_folder = getattr(self.manager, "invoice_folder", None) or getattr(self.manager, "data_folder", None) or os.getcwd()
        inv_items = []
        try:
            invm = InventoryManager(inv_folder)
            inv_items = list(invm.items or [])
        except Exception:
            inv_items = []

        low_stock = []
        for it in inv_items:
            try:
                qty = float(getattr(it, "quantity", 0) or 0)
            except Exception:
                qty = 0.0
            try:
                mn = float(getattr(it, "min_stock", 0) or 0)
            except Exception:
                mn = 0.0
            if mn > 0 and qty <= mn:
                low_stock.append(it)

        def _parse_date(s):
            txt = str(s or "").strip()
            if not txt:
                return None
            for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
                try:
                    return datetime.strptime(txt, fmt)
                except Exception:
                    continue
            try:
                return datetime.fromisoformat(txt)
            except Exception:
                return None

        expiring = []
        horizon = now + timedelta(days=30)
        for it in inv_items:
            dt = _parse_date(getattr(it, "expiry_date", ""))
            if dt and now.date() <= dt.date() <= horizon.date():
                expiring.append((dt, it))
        expiring.sort(key=lambda x: x[0])

        temp_alerts = 0
        missing_temp_details = []
        try:
            from temperature_log import TemperatureLogManager
            tman = TemperatureLogManager(inv_folder)
            logs = tman.get_grouped_logs(start_date=today, end_date=today) or []
            rooms = sorted({str(r.get("room_name") or "").strip() for r in logs if str(r.get("room_name") or "").strip()})
            by_room_slot = {}
            for r in logs:
                room = str(r.get("room_name") or "").strip()
                slot = str(r.get("slot_name") or "").strip()
                if not room or not slot:
                    continue
                by_room_slot.setdefault(room, set()).add(slot.lower())
            for room in rooms:
                present = by_room_slot.get(room, set())
                for slot in ("morning", "evening"):
                    if slot not in present:
                        temp_alerts += 1
                        missing_temp_details.append(f"{room}: missing {slot.title()}")
        except Exception:
            pass

        notifications = []
        try:
            notifications = self.collect_notifications() or []
        except Exception:
            notifications = []

        warehouse_metrics = {}
        recent_goods_received = []
        recent_stock_transfers = []
        pending_approvals = 0
        total_inventory_value = sum(_to_float(getattr(it, "total_cost", 0)) for it in inv_items)
        if WAREHOUSE_EXTENSION_AVAILABLE:
            try:
                warehouse_manager = WarehouseManager(inv_folder, str(getattr(self, "current_company", "") or ""))
                warehouse_metrics = warehouse_manager.get_dashboard_metrics() or {}
                total_inventory_value = _to_float(warehouse_metrics.get("total_inventory_value", total_inventory_value))
                recent_goods_received = list(warehouse_metrics.get("recent_goods_received") or [])
                recent_stock_transfers = list(warehouse_metrics.get("recent_stock_transfers") or [])
                pending_approvals = int(warehouse_metrics.get("pending_approvals") or 0)
            except Exception:
                warehouse_metrics = {}

        company_cfg = self._get_company_configuration() if WAREHOUSE_EXTENSION_AVAILABLE else {}
        dashboard_widgets = ((company_cfg.get("settings") or {}).get("dashboard_widgets") or {}) if company_cfg else {}

        def widget_enabled(key, default=True):
            return bool(dashboard_widgets.get(key, default))

        def trend_delta(cur, prev, suffix="today"):
            if prev is None:
                return ("", "zero")
            try:
                cur = float(cur)
                prev = float(prev)
            except Exception:
                return ("", "zero")
            delta = cur - prev
            if delta == 0:
                return (f"— 0 {suffix}".strip(), "zero")
            sign = "▲" if delta > 0 else "▼"
            pm = "+" if delta > 0 else "-"
            return (f"{sign} {pm}{abs(delta):,.0f} {suffix}".strip(), "pos" if delta > 0 else "neg")

        def trend_percent(cur, prev, suffix="vs yesterday"):
            if prev is None:
                return ("", "zero")
            try:
                cur = float(cur)
                prev = float(prev)
            except Exception:
                return ("", "zero")
            if prev == 0:
                return ("", "zero")
            pct = ((cur - prev) / prev) * 100.0
            if pct == 0:
                return (f"— 0% {suffix}".strip(), "zero")
            sign = "▲" if pct > 0 else "▼"
            return (f"{sign} {abs(pct):.0f}% {suffix}".strip(), "pos" if pct > 0 else "neg")

        purchases_yesterday_total = sum(_to_float(p.get("amount")) for p in purchases if str(p.get("date") or "") == yesterday)
        sales_trend_txt, sales_trend_kind = trend_percent(today_sales_total, yesterday_sales_total)
        purch_trend_txt, purch_trend_kind = trend_percent(purchases_today_total, purchases_yesterday_total)
        inv_trend_txt, inv_trend_kind = trend_delta(today_sales_count, yesterday_sales_count)

        def _sub_style(kind):
            if kind == "pos":
                return "KPISubPos.TLabel"
            if kind == "neg":
                return "KPISubNeg.TLabel"
            return "KPISub.TLabel"

        def render_kpi_row(parent, entries, uniform_name):
            visible_entries = [entry for entry in entries if widget_enabled(entry["widget_key"], True)]
            if not visible_entries:
                return
            row = ttk.Frame(parent, style="App.TFrame")
            row.pack(fill="x", pady=(0, 12))
            for i in range(len(visible_entries)):
                row.columnconfigure(i, weight=1, uniform=uniform_name)
            for idx, entry in enumerate(visible_entries):
                card = ttk.LabelFrame(row, text=entry["title"], style=entry["style_name"])
                card.grid(row=0, column=idx, sticky="nsew", padx=(0 if idx == 0 else 10, 0))
                ttk.Label(card, text=entry["value"], style="KPIValue.TLabel").pack(anchor="w", padx=14, pady=(12, 2))
                ttk.Label(card, text=entry["sub"], style=entry.get("sub_style", "KPISub.TLabel")).pack(anchor="w", padx=14, pady=(0, 12))

        kpis = [
            {"widget_key": "todays_sales", "title": "Today's Sales", "value": f"AED {today_sales_total:,.2f}", "sub": sales_trend_txt or "Today", "style_name": "CardInfo.TLabelframe", "sub_style": _sub_style(sales_trend_kind)},
            {"widget_key": "purchases_today", "title": "Purchases Today", "value": f"AED {purchases_today_total:,.2f}", "sub": purch_trend_txt or "Today", "style_name": "CardInfo.TLabelframe", "sub_style": _sub_style(purch_trend_kind)},
            {"widget_key": "outstanding_receivables", "title": "Outstanding Receivables", "value": f"AED {outstanding_receivables:,.2f}", "sub": "Unpaid invoices", "style_name": "CardWarn.TLabelframe"},
            {"widget_key": "temperature_alerts_card", "title": "Temperature Alerts", "value": str(temp_alerts), "sub": "Missing readings", "style_name": "CardWarn.TLabelframe"},
        ]
        render_kpi_row(container, kpis, "kpi")

        kpis2 = [
            {"widget_key": "invoices", "title": "Invoices", "value": str(len(invoices)), "sub": inv_trend_txt or "Today", "style_name": "CardInfo.TLabelframe", "sub_style": _sub_style(inv_trend_kind)},
            {"widget_key": "outstanding_payables", "title": "Outstanding Payables", "value": f"AED {outstanding_payables:,.2f}", "sub": "Unpaid purchases", "style_name": "CardWarn.TLabelframe"},
            {"widget_key": "low_stock_products", "title": "Low Stock Products", "value": str(len(low_stock)), "sub": "Needs reorder", "style_name": "CardWarn.TLabelframe"},
            {"widget_key": "expiring_products", "title": "Expiring Products (30d)", "value": str(len(expiring)), "sub": "Check batches", "style_name": "CardWarn.TLabelframe"},
        ]
        render_kpi_row(container, kpis2, "kpi2")

        if WAREHOUSE_EXTENSION_AVAILABLE and self._is_module_enabled("warehouse_management", True):
            kpis3 = [
                {"widget_key": "inventory_value", "title": "Inventory Value", "value": f"AED {total_inventory_value:,.2f}", "sub": "At cost", "style_name": "CardInfo.TLabelframe"},
                {"widget_key": "total_stock_items", "title": "Total Stock Items", "value": str(int(warehouse_metrics.get("total_stock_items", len(inv_items)) or 0)), "sub": "Tracked SKUs", "style_name": "CardInfo.TLabelframe"},
                {"widget_key": "expired_items", "title": "Expired Items", "value": str(int(warehouse_metrics.get("expired_items", len([x for x in inv_items if getattr(x, 'is_expired', lambda: False)()])) or 0)), "sub": "Requires action", "style_name": "CardWarn.TLabelframe"},
                {"widget_key": "pending_approvals", "title": "Pending Approvals", "value": str(pending_approvals), "sub": "Warehouse workflow", "style_name": "CardWarn.TLabelframe"},
            ]
            render_kpi_row(container, kpis3, "kpi3")

        if widget_enabled("quick_actions", True):
            actions = ttk.LabelFrame(container, text="Quick Actions", style="Card.TLabelframe")
            actions.pack(fill="x", pady=(0, 12))
            btn_row = ttk.Frame(actions, style="CardBody.TFrame")
            btn_row.pack(fill="x", padx=14, pady=14)
            ttk.Button(btn_row, text="New Invoice", command=self._qc_new_invoice, style="Primary.TButton").pack(side="left", padx=(0, 10))
            ttk.Button(btn_row, text="Receive Stock", command=self.show_inventory_system, style="Secondary.TButton").pack(side="left", padx=(0, 10))
            if WAREHOUSE_EXTENSION_AVAILABLE and self._is_module_enabled("warehouse_management", True):
                ttk.Button(btn_row, text="Warehouse", command=self.show_warehouse_management, style="Secondary.TButton").pack(side="left", padx=(0, 10))
            ttk.Button(btn_row, text="Reports", command=self.show_reports, style="Secondary.TButton").pack(side="left", padx=(0, 10))
            ttk.Button(btn_row, text="💳 Cost Centers", command=self.open_cost_center_and_backfill, style="Secondary.TButton").pack(side="left", padx=(0, 10))
            ttk.Button(btn_row, text="Clients", command=self.show_clients, style="Secondary.TButton").pack(side="left")

        left_panel_enabled = any(widget_enabled(key, True) for key in ("low_stock_table", "expiring_products_table"))
        right_panel_enabled = any(widget_enabled(key, True) for key in ("notifications_panel", "temperature_alerts_panel", "recent_activities"))
        if left_panel_enabled or right_panel_enabled:
            grid = ttk.Frame(container, style="App.TFrame")
            grid.pack(fill="both", expand=True)
            if left_panel_enabled:
                grid.columnconfigure(0, weight=1, uniform="cols")
            if right_panel_enabled:
                grid.columnconfigure(1 if left_panel_enabled else 0, weight=1, uniform="cols")

            left = None
            right = None
            if left_panel_enabled:
                left = ttk.Frame(grid, style="App.TFrame")
                left.grid(row=0, column=0, sticky="nsew", padx=(0, 8 if right_panel_enabled else 0))
            if right_panel_enabled:
                right_col = 1 if left_panel_enabled else 0
                right = ttk.Frame(grid, style="App.TFrame")
                right.grid(row=0, column=right_col, sticky="nsew", padx=((8 if left_panel_enabled else 0), 0))

            if left is not None and widget_enabled("low_stock_table", True):
                low_frame = ttk.LabelFrame(left, text="Low Stock Products", style="Card.TLabelframe")
                low_frame.pack(fill="both", expand=True, pady=(0, 12))
                low_tree = ttk.Treeview(low_frame, columns=("Code", "Name", "Qty", "Min"), show="headings", height=7)
                for c in ("Code", "Name", "Qty", "Min"):
                    low_tree.heading(c, text=c)
                low_tree.column("Code", width=120)
                low_tree.column("Name", width=260)
                low_tree.column("Qty", width=80, anchor="e")
                low_tree.column("Min", width=80, anchor="e")
                low_tree.pack(fill="both", expand=True)
                for it in sorted(low_stock, key=lambda x: float(getattr(x, "quantity", 0) or 0)):
                    low_tree.insert("", "end", values=(
                        str(getattr(it, "item_id", "") or ""),
                        str(getattr(it, "name", "") or ""),
                        str(getattr(it, "quantity", "") or ""),
                        str(getattr(it, "min_stock", "") or ""),
                    ))

            if left is not None and widget_enabled("expiring_products_table", True):
                exp_frame = ttk.LabelFrame(left, text="Expiring Products (Next 30 Days)", style="Card.TLabelframe")
                exp_frame.pack(fill="both", expand=True)
                exp_tree = ttk.Treeview(exp_frame, columns=("Code", "Name", "Batch", "Expiry"), show="headings", height=7)
                for c in ("Code", "Name", "Batch", "Expiry"):
                    exp_tree.heading(c, text=c)
                exp_tree.column("Code", width=120)
                exp_tree.column("Name", width=240)
                exp_tree.column("Batch", width=140)
                exp_tree.column("Expiry", width=120)
                exp_tree.pack(fill="both", expand=True)
                for dt, it in expiring[:20]:
                    exp_tree.insert("", "end", values=(
                        str(getattr(it, "item_id", "") or ""),
                        str(getattr(it, "name", "") or ""),
                        str(getattr(it, "batch_number", "") or ""),
                        str(getattr(it, "expiry_date", "") or dt.strftime("%Y-%m-%d")),
                    ))

            if right is not None and widget_enabled("notifications_panel", True):
                notif_frame = ttk.LabelFrame(right, text="Notifications", style="Card.TLabelframe")
                notif_frame.pack(fill="both", expand=True, pady=(0, 12))
                notif_list = tk.Listbox(notif_frame, height=8, bg=lb_bg, fg=lb_fg, selectbackground=lb_sel_bg, selectforeground=lb_sel_fg, highlightthickness=1, highlightbackground=lb_border, bd=0)
                notif_list.pack(fill="both", expand=True)
                for n in notifications[:50]:
                    notif_list.insert("end", str(n))

            if right is not None and widget_enabled("temperature_alerts_panel", True):
                temp_frame = ttk.LabelFrame(right, text="Temperature Alerts", style="Card.TLabelframe")
                temp_frame.pack(fill="both", expand=True, pady=(0, 12))
                temp_list = tk.Listbox(temp_frame, height=6, bg=lb_bg, fg=lb_fg, selectbackground=lb_sel_bg, selectforeground=lb_sel_fg, highlightthickness=1, highlightbackground=lb_border, bd=0)
                temp_list.pack(fill="both", expand=True)
                if missing_temp_details:
                    for line in missing_temp_details[:50]:
                        temp_list.insert("end", line)
                else:
                    temp_list.insert("end", "No missing readings detected for today.")

            if right is not None and widget_enabled("recent_activities", True):
                act_frame = ttk.LabelFrame(right, text="Recent Activities", style="Card.TLabelframe")
                act_frame.pack(fill="both", expand=True)
                audit_path = os.path.join(str(inv_folder), "audit_log.jsonl")
                act_toolbar = ttk.Frame(act_frame)
                act_toolbar.pack(fill="x", padx=4, pady=(2, 6))
                act_status = tk.StringVar(value="Activities: 0")
                ttk.Label(act_toolbar, textvariable=act_status).pack(side="left")

                act_tree_wrap = ttk.Frame(act_frame)
                act_tree_wrap.pack(fill="both", expand=True)
                act_tree = ttk.Treeview(act_tree_wrap, columns=("Time", "Action", "Module", "Reference"), show="headings", height=10)
                for col, width in (("Time", 150), ("Action", 220), ("Module", 130), ("Reference", 150)):
                    act_tree.heading(col, text=col)
                    act_tree.column(col, width=width, stretch=True)
                act_scroll = ttk.Scrollbar(act_tree_wrap, orient="vertical", command=act_tree.yview)
                act_tree.configure(yscrollcommand=act_scroll.set)
                act_tree.pack(side="left", fill="both", expand=True)
                act_scroll.pack(side="right", fill="y")

                def load_activity_rows():
                    for child in act_tree.get_children():
                        act_tree.delete(child)
                    count = 0
                    try:
                        if os.path.exists(audit_path):
                            with open(audit_path, "r", encoding="utf-8") as f:
                                lines = f.readlines()
                            for line in reversed(lines):
                                try:
                                    rec = json.loads(line.strip() or "{}")
                                except Exception:
                                    continue
                                ts = str(rec.get("timestamp") or rec.get("time") or "")[:19]
                                action = str(rec.get("action") or rec.get("event_type") or rec.get("type") or "Activity")
                                module_name = str(rec.get("module") or rec.get("source") or "")
                                reference = str(rec.get("reference_id") or rec.get("details") or "")
                                act_tree.insert("", "end", values=(ts, action, module_name, reference))
                                count += 1
                    except Exception:
                        count = 0
                    act_status.set(f"Activities: {count}")

                def clear_activity_rows():
                    if not messagebox.askyesno("Clear Activities", "Clear all logged activities from Recent Activities?"):
                        return
                    if self.clear_activity_log():
                        load_activity_rows()
                    else:
                        messagebox.showerror("Error", "Failed to clear activities log")

                ttk.Button(act_toolbar, text="Clear Activities", command=clear_activity_rows).pack(side="right")
                load_activity_rows()

        warehouse_strip_enabled = (
            WAREHOUSE_EXTENSION_AVAILABLE
            and self._is_module_enabled("warehouse_management", True)
            and any(widget_enabled(key, True) for key in ("recent_goods_received", "recent_stock_transfers"))
        )
        if warehouse_strip_enabled:
            warehouse_strip = ttk.Frame(container, style="App.TFrame")
            warehouse_strip.pack(fill="both", expand=True, pady=(12, 0))
            col = 0

            if widget_enabled("recent_goods_received", True):
                warehouse_strip.columnconfigure(col, weight=1, uniform="warehouse-strip")
                grn_frame = ttk.LabelFrame(warehouse_strip, text="Recent Goods Received", style="Card.TLabelframe")
                grn_frame.grid(row=0, column=col, sticky="nsew", padx=(0, 8 if widget_enabled("recent_stock_transfers", True) else 0))
                grn_tree = ttk.Treeview(grn_frame, columns=("GRN", "Product", "Qty", "Warehouse", "Date"), show="headings", height=6)
                for c in ("GRN", "Product", "Qty", "Warehouse", "Date"):
                    grn_tree.heading(c, text=c)
                grn_tree.pack(fill="both", expand=True)
                for row in recent_goods_received:
                    grn_tree.insert("", "end", values=(
                        row.get("grn_id"),
                        row.get("product_name"),
                        row.get("quantity_received"),
                        row.get("warehouse_code"),
                        row.get("received_date"),
                    ))
                col += 1

            if widget_enabled("recent_stock_transfers", True):
                warehouse_strip.columnconfigure(col, weight=1, uniform="warehouse-strip")
                trf_frame = ttk.LabelFrame(warehouse_strip, text="Recent Stock Transfers", style="Card.TLabelframe")
                trf_frame.grid(row=0, column=col, sticky="nsew")
                trf_tree = ttk.Treeview(trf_frame, columns=("Transfer", "Product", "Qty", "Route", "Status"), show="headings", height=6)
                for c in ("Transfer", "Product", "Qty", "Route", "Status"):
                    trf_tree.heading(c, text=c)
                trf_tree.pack(fill="both", expand=True)
                for row in recent_stock_transfers:
                    trf_tree.insert("", "end", values=(
                        row.get("transfer_id"),
                        row.get("product_name") or row.get("item_id"),
                        row.get("quantity"),
                        f"{row.get('warehouse_from')} -> {row.get('warehouse_to')}",
                        row.get("status"),
                    ))

        try:
            floating = ttk.Button(parent, text="+ New", command=self._open_quick_create_menu, style="Primary.TButton")
            floating.place(relx=1.0, rely=1.0, x=-18, y=-18, anchor="se")
            self._floating_new_btn = floating
        except Exception:
            pass

    def _apply_saas_theme(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass

        palette = self._theme_palette()
        bg = palette["bg"]
        card_bg = palette["card_bg"]
        text = palette["text"]
        muted = palette["muted"]
        border = palette["border"]
        blue = palette["blue"]
        green = palette["green"]
        orange = palette["orange"]
        danger = palette["danger"]
        nav_selected = palette["nav_selected"]
        tree_selected = palette["tree_selected"]
        mode = str(getattr(self, "ui_theme_mode", "light") or "light").strip().lower()
        brand_red = "#c1121f"

        try:
            import tkinter.font as tkfont
            families = set(tkfont.families(self.root))
            font_base = "Inter" if "Inter" in families else ("Segoe UI" if "Segoe UI" in families else ("SF Pro Display" if "SF Pro Display" in families else "Helvetica"))
        except Exception:
            font_base = "Helvetica"

        style.configure(".", font=(font_base, 11), background=bg, foreground=text)
        style.configure("App.TFrame", background=bg)
        style.configure("CardBody.TFrame", background=card_bg)
        style.configure("TFrame", background=bg)
        style.configure("TLabel", background=bg, foreground=text)
        style.configure("PageTitle.TLabel", font=(font_base, 18, "bold"), background=bg, foreground=text)
        style.configure("TopbarTitle.TLabel", font=(font_base, 16, "bold"), background=card_bg, foreground=text)
        style.configure("TopbarMeta.TLabel", font=(font_base, 10), background=card_bg, foreground=muted)
        style.configure("SidebarTitle.TLabel", font=(font_base, 12, "bold"), background=card_bg, foreground=text)
        style.configure("SidebarLogo.TLabel", background=card_bg)
        style.configure("Sidebar.TFrame", background=card_bg)
        style.configure("Topbar.TFrame", background=card_bg)

        style.configure("TSeparator", background=border)

        style.configure("TEntry", padding=8, fieldbackground=card_bg, foreground=text, insertcolor=text, bordercolor=border)
        style.configure("TCombobox", padding=6, fieldbackground=card_bg, foreground=text, background=card_bg, arrowcolor=text, bordercolor=border)

        style.configure("GlobalSearch.TEntry", padding=10, fieldbackground=card_bg, foreground=text)
        style.configure("GlobalSearchPlaceholder.TEntry", padding=10, fieldbackground=card_bg, foreground=muted)

        style.configure("Topbar.TButton", padding=(10, 8))
        style.configure("TopbarIcon.TButton", padding=(10, 8))
        style.configure("TopbarPill.TButton", padding=(12, 8), background=blue, foreground="white")
        try:
            style.map("TopbarPill.TButton", background=[("active", "#1d4ed8")], foreground=[("active", "white")])
        except Exception:
            pass

        style.configure("TopbarDanger.TButton", padding=(10, 8), background=brand_red, foreground="white")
        try:
            style.map("TopbarDanger.TButton", background=[("active", "#a10f1a")], foreground=[("active", "white")])
        except Exception:
            pass

        style.configure("SidebarToggle.TButton", padding=(10, 10), background=brand_red, foreground="white")
        try:
            style.map("SidebarToggle.TButton", background=[("active", "#a10f1a")], foreground=[("active", "white")])
        except Exception:
            pass

        style.configure("Primary.TButton", padding=(14, 10), background=brand_red, foreground="white")
        style.configure("Secondary.TButton", padding=(14, 10), background=blue, foreground="white")
        try:
            style.map("Primary.TButton", background=[("active", "#a10f1a")], foreground=[("active", "white")])
            style.map("Secondary.TButton", background=[("active", "#1d4ed8")], foreground=[("active", "white")])
        except Exception:
            pass

        style.configure("Card.TLabelframe", background=card_bg, bordercolor=border, relief="solid")
        style.configure("Card.TLabelframe.Label", font=(font_base, 11, "bold"), background=card_bg, foreground=text)

        style.configure("CardInfo.TLabelframe", background=card_bg, bordercolor=border, relief="solid")
        style.configure("CardInfo.TLabelframe.Label", font=(font_base, 11, "bold"), background=card_bg, foreground=blue)
        style.configure("CardWarn.TLabelframe", background=card_bg, bordercolor=border, relief="solid")
        style.configure("CardWarn.TLabelframe.Label", font=(font_base, 11, "bold"), background=card_bg, foreground=orange)

        style.configure("KPIValue.TLabel", font=(font_base, 22, "bold"), background=card_bg, foreground=text)
        style.configure("KPISub.TLabel", font=(font_base, 10), background=card_bg, foreground=muted)
        style.configure("KPISubPos.TLabel", font=(font_base, 10), background=card_bg, foreground=green)
        style.configure("KPISubNeg.TLabel", font=(font_base, 10), background=card_bg, foreground=danger)

        style.configure("Nav.Treeview", background=card_bg, fieldbackground=card_bg, foreground=text, borderwidth=0, relief="flat", rowheight=32)
        style.configure("Nav.Treeview.Heading", font=(font_base, 10, "bold"))
        try:
            style.map("Nav.Treeview", background=[("selected", nav_selected)], foreground=[("selected", text)])
        except Exception:
            pass

        style.configure("Treeview", font=(font_base, 11), rowheight=28, background=card_bg, fieldbackground=card_bg, foreground=text, bordercolor=border)
        style.configure("Treeview.Heading", font=(font_base, 11, "bold"), background=card_bg, foreground=text)
        try:
            style.map("Treeview", background=[("selected", tree_selected)], foreground=[("selected", text)])
            style.map("TCombobox", fieldbackground=[("readonly", card_bg)], foreground=[("readonly", text)], selectforeground=[("readonly", text)], selectbackground=[("readonly", card_bg)])
        except Exception:
            pass

        style.configure("TopbarDanger.TButton", padding=(10, 8), background=danger, foreground="white")
        try:
            if mode == "dark":
                style.map("TopbarDanger.TButton", background=[("active", "#ef4444")], foreground=[("active", "white")])
            else:
                style.map("TopbarDanger.TButton", background=[("active", "#a10f1a")], foreground=[("active", "white")])
        except Exception:
            pass

    def _ui_settings_path(self):
        folder = None
        try:
            folder = getattr(self.manager, "invoice_folder", None) or getattr(self.manager, "data_folder", None)
        except Exception:
            folder = None
        if not folder:
            folder = os.getcwd()
        return os.path.join(str(folder), "app_settings.json")

    def _load_ui_prefs(self):
        self.ui_theme_mode = "light"
        self.ui_favorites = []
        self.ui_window_bounds = {}
        if not hasattr(self, "sidebar_collapsed"):
            self.sidebar_collapsed = False
        path = self._ui_settings_path()
        try:
            if not os.path.exists(path):
                return
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f) or {}
        except Exception:
            return
        try:
            self.ui_theme_mode = str(data.get("ui_theme_mode", self.ui_theme_mode) or self.ui_theme_mode).strip().lower()
        except Exception:
            pass
        try:
            fav = data.get("ui_favorites", []) or []
            if isinstance(fav, list):
                self.ui_favorites = [str(x) for x in fav if str(x).strip()]
        except Exception:
            pass
        try:
            self.sidebar_collapsed = bool(data.get("ui_sidebar_collapsed", self.sidebar_collapsed))
        except Exception:
            pass
        try:
            window_bounds = data.get("ui_window_bounds", {}) or {}
            if isinstance(window_bounds, dict):
                self.ui_window_bounds = {
                    str(k): v for k, v in window_bounds.items()
                    if str(k).strip() and isinstance(v, dict)
                }
        except Exception:
            pass

    def _save_ui_prefs(self):
        path = self._ui_settings_path()
        data = {}
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f) or {}
        except Exception:
            data = {}
        data["ui_theme_mode"] = str(getattr(self, "ui_theme_mode", "light") or "light").strip().lower()
        data["ui_favorites"] = list(getattr(self, "ui_favorites", []) or [])
        data["ui_sidebar_collapsed"] = bool(getattr(self, "sidebar_collapsed", False))
        data["ui_window_bounds"] = dict(getattr(self, "ui_window_bounds", {}) or {})
        try:
            folder = os.path.dirname(path)
            if folder:
                os.makedirs(folder, exist_ok=True)
        except Exception:
            pass
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _schedule_ui_prefs_save(self):
        try:
            job = getattr(self, "_prefs_save_job", None)
            if job:
                self.root.after_cancel(job)
        except Exception:
            pass
        try:
            self._prefs_save_job = self.root.after(350, self._save_ui_prefs)
        except Exception:
            pass

    def _get_saved_window_geometry(self, key):
        key = str(key or "").strip()
        if not key:
            return None
        try:
            rec = (getattr(self, "ui_window_bounds", {}) or {}).get(key, {}) or {}
            geom = str(rec.get("geometry") or "").strip()
            return geom or None
        except Exception:
            return None

    def _remember_window_geometry(self, win, key, force=False):
        key = str(key or "").strip()
        if not key:
            return
        try:
            if not win.winfo_exists():
                return
            state = str(win.state() or "normal").lower()
            if state == "iconic":
                return
            geom = str(win.geometry() or "").strip()
            if not geom:
                return
            if not hasattr(self, "ui_window_bounds") or not isinstance(self.ui_window_bounds, dict):
                self.ui_window_bounds = {}
            self.ui_window_bounds[key] = {"geometry": geom, "state": state}
            self.root._ui_window_bounds = self.ui_window_bounds
            if force:
                self._save_ui_prefs()
            else:
                self._schedule_ui_prefs_save()
        except Exception:
            pass

    def _schedule_remember_window(self, key, win=None):
        target = win or self.root
        try:
            job = getattr(target, "_remember_window_job", None)
            if job:
                target.after_cancel(job)
        except Exception:
            pass
        try:
            target._remember_window_job = target.after(300, lambda: self._remember_window_geometry(target, key))
        except Exception:
            pass

    def _restore_root_window_geometry(self):
        saved = self._get_saved_window_geometry("main_window")
        if saved:
            _fit_window(self.root, 1200, 800, mode="workspace", remember_key="main_window", restore=True)
            return
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            _fit_window(
                self.root,
                1200,
                800,
                pref_w=int(sw * 0.94),
                pref_h=int(sh * 0.90),
                mode="workspace",
                remember_key="main_window",
                restore=False,
            )
        except Exception:
            _fit_window(self.root, 1200, 800, mode="workspace", remember_key="main_window", restore=False)

    def _on_app_close(self):
        try:
            self._remember_window_geometry(self.root, "main_window", force=True)
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    def toggle_theme_mode(self):
        current = str(getattr(self, "ui_theme_mode", "light") or "light").strip().lower()
        self.ui_theme_mode = "dark" if current != "dark" else "light"
        try:
            self._save_ui_prefs()
        except Exception:
            pass
        try:
            self._apply_saas_theme()
        except Exception:
            pass
        self.create_main_menu()
        
    def update_notifications_badge(self):
        count = 0
        try:
            count = len(self.collect_notifications())
        except Exception:
            count = 0
        try:
            if hasattr(self, 'notif_btn'):
                self.notif_btn.config(text=f"🔔 Notifications ({count})")
        except Exception:
            pass
        self.root.after(20000, self.update_notifications_badge)

    def collect_notifications(self):
        notes = []
        try:
            invoices = []
            try:
                invoices = self.manager.get_all_invoices_dict()
            except Exception:
                invoices = []
            today = datetime.now().date()
            for inv in invoices:
                status = str(inv.get('status', '')).lower()
                wf_status = str(inv.get('workflow_status', '')).lower()
                if status in ('closed','paid','fully paid') or wf_status in ('closed','paid','fully paid'):
                    continue
                due_date_str = inv.get('due_date', '')
                due_date = None
                try:
                    if due_date_str:
                        due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
                except Exception:
                    due_date = None
                total = 0.0
                try:
                    gt = inv.get('grand_total')
                    if gt is None:
                        st = inv.get('subtotal') or 0
                        ta = inv.get('tax_amount') or 0
                        gt = (st or 0) + (ta or 0)
                    total = float(gt or 0)
                except Exception:
                    total = 0.0
                paid = 0.0
                for p in inv.get('payments', []) or []:
                    try:
                        paid += float(p.get('amount') or 0)
                    except Exception:
                        pass
                outstanding = max(total - paid, 0.0)
                if outstanding > 0:
                    overdue = (due_date and due_date < today)
                    soon = (due_date and 0 <= (due_date - today).days <= 7)
                    if overdue or soon:
                        days = None
                        try:
                            days = (due_date - today).days if due_date else None
                        except Exception:
                            days = None
                        notes.append({
                            'type': 'payment',
                            'invoice_id': inv.get('invoice_id'),
                            'client': inv.get('client_name'),
                            'outstanding': round(outstanding, 2),
                            'due_in_days': days
                        })
        except Exception:
            pass
        try:
            if INVENTORY_AVAILABLE:
                from inventory_system import InventoryManager
                im = InventoryManager(self.manager.invoice_folder)
                seen_expiry_keys = set()
                for item in im.get_low_stock_items():
                    notes.append({
                        'type': 'stock',
                        'item_id': item.item_id,
                        'name': item.name,
                        'qty': item.quantity,
                        'min': item.min_stock
                    })
                for item in im.get_expired_items():
                    days = None
                    try:
                        if item.expiry_date:
                            d = datetime.strptime(item.expiry_date, "%Y-%m-%d").date()
                            days = (d - datetime.now().date()).days
                    except Exception:
                        days = None
                    notes.append({
                        'type': 'expiry',
                        'item_id': item.item_id,
                        'name': item.name,
                        'expiry_date': item.expiry_date,
                        'days_left': days
                    })
                    seen_expiry_keys.add((str(item.item_id), str(item.expiry_date)))
                for item in im.get_expiring_soon_items(14):
                    days = None
                    try:
                        if item.expiry_date:
                            d = datetime.strptime(item.expiry_date, "%Y-%m-%d").date()
                            days = (d - datetime.now().date()).days
                    except Exception:
                        days = None
                    expiry_key = (str(item.item_id), str(item.expiry_date))
                    if expiry_key in seen_expiry_keys:
                        continue
                    notes.append({
                        'type': 'expiry',
                        'item_id': item.item_id,
                        'name': item.name,
                        'expiry_date': item.expiry_date,
                        'days_left': days
                    })
                    seen_expiry_keys.add(expiry_key)
        except Exception:
            pass
        try:
            license_path = os.path.join(getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaInvoice")), "license.json")
            if os.path.exists(license_path):
                with open(license_path, 'r') as f:
                    lic = json.load(f)
                exp = str(lic.get('expires', '') or '')
                if exp:
                    try:
                        d = datetime.strptime(exp, "%Y-%m-%d").date()
                        days = (d - datetime.now().date()).days
                        if days <= 14:
                            notes.append({'type': 'license', 'expires': exp, 'days_left': days})
                    except Exception:
                        pass
        except Exception:
            pass
        return notes

    def _show_start_animation(self):
        palette = self._theme_palette()
        overlay = tk.Frame(self.root, background=palette["bg"])
        overlay.pack(fill='both', expand=True)
        inner = ttk.Frame(overlay)
        inner.pack(fill='both', expand=True)
        try:
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            size = (int(sw*0.35), int(sh*0.35))
            candidates = [
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "AssistemLogo.png"),
                os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData", "AssistemLogo.png"),
                os.path.join(os.getcwd(), "AssistemLogo.png")
            ]
            loaded = False
            for p in candidates:
                if os.path.exists(p):
                    try:
                        img = Image.open(p)
                        img = img.resize(size, Image.Resampling.LANCZOS)
                        photo = ImageTk.PhotoImage(img)
                        lbl = ttk.Label(inner, image=photo, anchor='center')
                        lbl.image = photo
                        lbl.pack(fill='both', expand=True)
                        loaded = True
                        break
                    except Exception:
                        pass
            if not loaded:
                ttk.Label(inner, text="HopePharma", font=('Helvetica', 56, 'bold'), anchor='center').pack(fill='both', expand=True)
        except Exception:
            ttk.Label(inner, text="HopePharma", font=('Helvetica', 56, 'bold')).pack(fill='both', expand=True)
        def proceed():
            overlay.destroy()
            self._show_company_login()
        self.root.after(1200, proceed)

    def _base_data_root(self):
        try:
            if hasattr(self.manager, "data_folder"):
                p = Path(self.manager.data_folder)
                for parent in [p] + list(p.parents):
                    if parent.name == "companies":
                        return parent.parent
                return p
        except Exception:
            pass
        try:
            p = Path(getattr(self.manager, "invoice_folder", Path.home() / "Documents" / "HopePharmaData"))
            for parent in [p] + list(p.parents):
                if parent.name == "companies":
                    return parent.parent
            return p
        except Exception:
            return Path.home() / "Documents" / "HopePharmaData"

    def _companies_file_path(self):
        return self._base_data_root() / "companies.json"

    def _load_companies(self):
        p = self._companies_file_path()
        if p.exists():
            try:
                with open(p, "r", encoding="utf-8") as f:
                    data = json.load(f)
                changed = False
                companies = []
                for row in (data.get("companies", []) or []):
                    normalized = normalize_company_record(row) if WAREHOUSE_EXTENSION_AVAILABLE else row
                    companies.append(normalized)
                    if normalized != row:
                        changed = True
                data["companies"] = companies
                if changed:
                    self._save_companies(data)
                return data
            except Exception:
                return {"companies": [], "active": None}
        return {"companies": [], "active": None}

    def _save_companies(self, data):
        p = self._companies_file_path()
        try:
            p.parent.mkdir(parents=True, exist_ok=True)
            with open(p, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            return True
        except Exception:
            return False

    def _hash_password(self, password, salt):
        try:
            return hashlib.sha256((salt + password).encode()).hexdigest()
        except Exception:
            return hashlib.sha256(password.encode()).hexdigest()

    def _verify_password(self, stored_hash, salt, password):
        try:
            return stored_hash == self._hash_password(password, salt)
        except Exception:
            return False

    def _ensure_default_company(self):
        data = self._load_companies()
        if not data.get("companies"):
            salt = uuid.uuid4().hex
            default = {
                "name": "HopePharma",
                "username": "HopePharma",
                "salt": salt,
                "password_hash": self._hash_password("Hope2025", salt),
                "info": {
                    "trn": "",
                    "address": "",
                    "phone": "",
                    "email": ""
                }
            }
            if WAREHOUSE_EXTENSION_AVAILABLE:
                default.update(normalize_company_record({}))
            data["companies"] = [default]
            data["active"] = "HopePharma"
            self._save_companies(data)
        return data

    def _company_folder(self, company_name):
        root = self._base_data_root()
        try:
            if str(company_name).strip() == "HopePharma":
                return root
        except Exception:
            pass
        return root / "companies" / company_name

    def _switch_company_context(self, company_name):
        folder = self._company_folder(company_name)
        try:
            folder.mkdir(parents=True, exist_ok=True)
        except Exception:
            pass
        try:
            if hasattr(self.manager, "data_folder"):
                self.manager.data_folder = folder
            if hasattr(self.manager, "invoice_folder"):
                self.manager.invoice_folder = str(folder)
            if hasattr(self.manager, "pdf_generator") and hasattr(self.manager.pdf_generator, "output_folder"):
                try:
                    self.manager.pdf_generator.output_folder = folder / "HopePharmaInvoices"
                except Exception:
                    pass
            if hasattr(self.manager, "initialize_data_files"):
                try:
                    self.manager.initialize_data_files()
                except Exception:
                    pass
            if WAREHOUSE_EXTENSION_AVAILABLE:
                try:
                    ensure_company_extension_storage(str(folder), company_name)
                except Exception:
                    pass
        except Exception:
            pass
        try:
            df = str(folder)
            self.files_to_watch = [
                os.path.join(df, "invoices_data.json"),
                os.path.join(df, "purchases_data.json"),
                os.path.join(df, "inventory.json"),
                os.path.join(df, "balance_data.json"),
            ]
            self.last_sync_times = {}
        except Exception:
            pass
        try:
            self.data_memory = DataMemoryManager(str(folder))
        except Exception:
            pass
        try:
            logo_path = folder / "logo.png"
            if logo_path.exists():
                self.logo_manager.logo_path = str(logo_path)
                self.logo_manager.data_folder = str(folder)
                self.logo_manager.load_logo(size=(64, 64))
        except Exception:
            pass
        self.current_company = company_name

    def _get_current_company_record(self):
        current_name = str(getattr(self, "current_company", "") or "").strip()
        data = self._load_companies()
        for row in data.get("companies", []) or []:
            if str(row.get("name") or "").strip() == current_name:
                return row
        return None

    def _get_company_configuration(self):
        base = default_company_configuration("Full ERP") if WAREHOUSE_EXTENSION_AVAILABLE else {}
        if not WAREHOUSE_EXTENSION_AVAILABLE:
            return base
        current_name = str(getattr(self, "current_company", "") or "").strip()
        company_rec = self._get_current_company_record() or {}
        folder = self._company_folder(current_name or company_rec.get("name") or "HopePharma")
        try:
            cfg = load_company_configuration(str(folder))
        except Exception:
            cfg = default_company_configuration(company_rec.get("company_type") or "Full ERP")
        if company_rec:
            cfg["company_type"] = str(company_rec.get("company_type") or cfg.get("company_type") or "Full ERP")
            cfg["enabled_modules"].update(company_rec.get("enabled_modules") or {})
        return cfg

    def _save_company_configuration(self, company_name, payload):
        if not WAREHOUSE_EXTENSION_AVAILABLE:
            return
        payload = payload or {}
        data = self._load_companies()
        target = None
        for row in data.get("companies", []) or []:
            if str(row.get("name") or "").strip() == str(company_name or "").strip():
                target = row
                break
        if target is None:
            return
        target["company_type"] = str(payload.get("company_type") or "Full ERP").strip() or "Full ERP"
        target["enabled_modules"] = dict(payload.get("enabled_modules") or {})
        self._save_companies(data)
        folder = self._company_folder(company_name)
        save_company_configuration(
            str(folder),
            target["company_type"],
            target["enabled_modules"],
            payload.get("settings") or {},
        )

    def _is_module_enabled(self, module_key, default=True):
        if not WAREHOUSE_EXTENSION_AVAILABLE:
            return default
        key = str(module_key or "").strip()
        if not key:
            return default
        cfg = self._get_company_configuration()
        return bool((cfg.get("enabled_modules") or {}).get(key, default))

    def _guard_module_access(self, module_key, module_label=None, admin_only=False):
        label = module_label or MODULE_LABELS.get(str(module_key or "").strip(), str(module_key or "Module"))
        if admin_only and getattr(self, "user_role", "Staff") != "Admin":
            messagebox.showerror("Error", f"{label} can only be accessed by admin")
            return False
        if not self._is_module_enabled(module_key, True):
            messagebox.showerror("Error", f"{label} is disabled for this company")
            return False
        return True

    def _open_company_settings(self):
        if not WAREHOUSE_EXTENSION_AVAILABLE:
            messagebox.showerror("Error", "Company settings extension is not available")
            return
        if not getattr(self, "current_company", None):
            messagebox.showerror("Error", "Please log into a company first")
            return
        current_name = str(getattr(self, "current_company", "") or "").strip() or "HopePharma"
        cfg = self._get_company_configuration()
        def _save(payload):
            self._save_company_configuration(current_name, payload)
            if str(current_name) == str(getattr(self, "current_company", "")):
                self.create_main_menu()
        CompanySettingsDialog(self.root, str(self._company_folder(current_name)), current_name, cfg, _save)

    def _ensure_app_menu(self):
        m = tk.Menu(self.root)
        account = tk.Menu(m, tearoff=0)
        account.add_command(label="Switch Company...", command=self.logout_company)
        if getattr(self, "current_company", None):
            account.add_command(label="Company Settings...", command=self._open_company_settings)
        if getattr(self, "user_role", "Staff") == "Admin":
            account.add_command(label="Logout Admin", command=self.logout_admin)
        else:
            account.add_command(label="Login as Admin...", command=self.login_admin)
        account.add_command(label="Change Password...", command=self._show_change_password_dialog)
        account.add_separator()
        account.add_command(label="Exit", command=self.root.quit)
        m.add_cascade(label="Account", menu=account)
        self.root.config(menu=m)
        self._menubar = m

    def logout_company(self):
        try:
            self.user_role = "Staff"
        except Exception:
            pass
        try:
            self.current_company = None
        except Exception:
            pass
        try:
            self.last_sync_times = {}
            self.files_to_watch = []
        except Exception:
            pass
        try:
            self._show_company_login()
        except Exception:
            pass

    def _show_company_login(self):
        for w in self.root.winfo_children():
            try:
                w.destroy()
            except Exception:
                pass
        try:
            self._ensure_app_menu()
        except Exception:
            pass
        data = self._ensure_default_company()
        frame = ttk.Frame(self.root, padding="24")
        frame.pack(fill="both", expand=True)
        header = ttk.Frame(frame)
        header.pack(pady=10)
        ttk.Label(header, text="Company Login", font=("Helvetica", 20, "bold")).pack()
        body = ttk.Frame(frame)
        body.pack(pady=10, fill="x")
        companies = [c.get("name") for c in data.get("companies", [])]
        selected_company = tk.StringVar(value=(data.get("active") or (companies[0] if companies else "")))
        ttk.Label(body, text="Company").pack(anchor="w")
        company_combo = ttk.Combobox(body, textvariable=selected_company, values=companies, state="readonly")
        company_combo.pack(fill="x", pady=4)
        ttk.Label(body, text="Username").pack(anchor="w")
        # Prefer last used username if available, otherwise the company's stored username
        def _default_username_for_company(name):
            try:
                for c in data.get("companies", []):
                    if c.get("name") == name:
                        return c.get("username", "")
            except Exception:
                pass
            return ""
        username_var = tk.StringVar(value=(data.get("last_username") or _default_username_for_company(selected_company.get())))
        username_entry = ttk.Entry(body, textvariable=username_var)
        username_entry.pack(fill="x", pady=4)
        ttk.Label(body, text="Password").pack(anchor="w")
        password_var = tk.StringVar()
        password_entry = ttk.Entry(body, textvariable=password_var, show="*")
        password_entry.pack(fill="x", pady=4)
        btns = ttk.Frame(frame)
        btns.pack(pady=12, fill="x")
        # Update username when company selection changes
        def on_company_change(event=None):
            try:
                uname = _default_username_for_company(selected_company.get())
                if uname:
                    username_var.set(uname)
            except Exception:
                pass
        try:
            company_combo.bind("<<ComboboxSelected>>", on_company_change)
        except Exception:
            pass
        def do_login():
            uname = username_var.get().strip()
            comp = selected_company.get().strip()
            pwd = password_var.get().strip()
            if not comp or not uname or not pwd:
                return
            rec = None
            for c in data.get("companies", []):
                if c.get("name") == comp and c.get("username") == uname:
                    rec = c
                    break
            if not rec:
                messagebox.showerror("Error", "Invalid credentials")
                return
            salt = rec.get("salt", "")
            ph = rec.get("password_hash", "")
            if not self._verify_password(ph, salt, pwd):
                messagebox.showerror("Error", "Invalid credentials")
                return
            # Persist last used company and username
            try:
                data["active"] = comp
                data["last_username"] = uname
                self._save_companies(data)
            except Exception:
                pass
            self._switch_company_context(comp)
            try:
                self.root.unbind_all("<Return>")
            except Exception:
                pass
            self.create_main_menu()
            try:
                if not rec.get("recovery_hash"):
                    if messagebox.askyesno("Recovery PIN", "Do you want to set a Recovery PIN now?\n\nThis lets you reset your password if you forget it."):
                        self._set_recovery_pin_dialog(data, comp, uname)
            except Exception:
                pass
        try:
            password_entry.bind("<Return>", lambda e: do_login())
            frame.bind("<Return>", lambda e: do_login())
        except Exception:
            pass
        ttk.Label(body, text="Press Enter to login", style="TopbarMeta.TLabel").pack(anchor="w", pady=(0, 8))
        def create_company():
            dlg = tk.Toplevel(self.root)
            dlg.title("Create Company")
            dlg.geometry("920x760")
            dlg.transient(self.root)
            dlg.grab_set()
            scroll_wrap = ScrollableFrame(dlg)
            scroll_wrap.pack(fill="both", expand=True)
            cf = ttk.Frame(scroll_wrap.scrollable_frame, padding="12")
            cf.pack(fill="both", expand=True)
            name_var = tk.StringVar()
            user_var = tk.StringVar()
            pwd_var = tk.StringVar()
            pin_var = tk.StringVar()
            trn_var = tk.StringVar()
            addr_var = tk.StringVar()
            phone_var = tk.StringVar()
            email_var = tk.StringVar()
            logo_path_var = tk.StringVar()
            ttk.Label(cf, text="Company Name").pack(anchor="w"); ttk.Entry(cf, textvariable=name_var).pack(fill="x", pady=4)
            ttk.Label(cf, text="Username").pack(anchor="w"); ttk.Entry(cf, textvariable=user_var).pack(fill="x", pady=4)
            ttk.Label(cf, text="Password").pack(anchor="w"); ttk.Entry(cf, textvariable=pwd_var, show="*").pack(fill="x", pady=4)
            ttk.Label(cf, text="Recovery PIN (optional)").pack(anchor="w"); ttk.Entry(cf, textvariable=pin_var, show="*").pack(fill="x", pady=4)
            ttk.Label(cf, text="TRN").pack(anchor="w"); ttk.Entry(cf, textvariable=trn_var).pack(fill="x", pady=4)
            ttk.Label(cf, text="Address").pack(anchor="w"); ttk.Entry(cf, textvariable=addr_var).pack(fill="x", pady=4)
            ttk.Label(cf, text="Phone").pack(anchor="w"); ttk.Entry(cf, textvariable=phone_var).pack(fill="x", pady=4)
            ttk.Label(cf, text="Email").pack(anchor="w"); ttk.Entry(cf, textvariable=email_var).pack(fill="x", pady=4)
            lf = ttk.Frame(cf); lf.pack(fill="x", pady=6)
            ttk.Label(lf, text="Logo").pack(side="left")
            ttk.Entry(lf, textvariable=logo_path_var).pack(side="left", fill="x", expand=True, padx=6)
            def pick_logo():
                p = filedialog.askopenfilename(title="Select Logo", filetypes=[("Image files","*.png;*.jpg;*.jpeg")])
                if p:
                    logo_path_var.set(p)
            ttk.Button(lf, text="Browse", command=pick_logo).pack(side="left")
            company_config_widget = None
            if WAREHOUSE_EXTENSION_AVAILABLE:
                initial_cfg = default_company_configuration("Full ERP")
                company_config_widget = CompanyConfigurationFrame(cf, initial_cfg)
                company_config_widget.pack(fill="both", expand=True, pady=(8, 0))
            bf = ttk.Frame(cf); bf.pack(pady=10, fill="x")
            def save_company():
                name = name_var.get().strip()
                user = user_var.get().strip()
                pwd = pwd_var.get().strip()
                if not name or not user or not pwd:
                    return
                existing = [c.get("name") for c in data.get("companies", [])]
                if name in existing:
                    messagebox.showerror("Error", "Company already exists")
                    return
                salt = uuid.uuid4().hex
                pin = pin_var.get().strip()
                rec_salt = uuid.uuid4().hex if pin else ""
                rec = {
                    "name": name,
                    "username": user,
                    "salt": salt,
                    "password_hash": self._hash_password(pwd, salt),
                    "recovery_salt": rec_salt,
                    "recovery_hash": (self._hash_password(pin, rec_salt) if pin else ""),
                    "info": {
                        "trn": trn_var.get().strip(),
                        "address": addr_var.get().strip(),
                        "phone": phone_var.get().strip(),
                        "email": email_var.get().strip()
                    }
                }
                if WAREHOUSE_EXTENSION_AVAILABLE and company_config_widget:
                    rec.update(normalize_company_record(company_config_widget.get_configuration()))
                data["companies"].append(rec)
                data["active"] = name
                self._save_companies(data)
                dest_folder = self._company_folder(name)
                try:
                    dest_folder.mkdir(parents=True, exist_ok=True)
                except Exception:
                    pass
                lp = logo_path_var.get().strip()
                if lp and os.path.exists(lp):
                    try:
                        shutil.copy(lp, str(dest_folder / "logo.png"))
                    except Exception:
                        pass
                if WAREHOUSE_EXTENSION_AVAILABLE and company_config_widget:
                    try:
                        ensure_company_extension_storage(str(dest_folder), name)
                        self._save_company_configuration(name, company_config_widget.get_configuration())
                    except Exception:
                        pass
                dlg.destroy()
                self._switch_company_context(name)
                self.create_main_menu()
            ttk.Button(bf, text="Create", command=save_company).pack(side="left", padx=4)
            ttk.Button(bf, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
            _fit_window(dlg, 920, 760, mode="workspace", remember_key="create_company")
            self.root.wait_window(dlg)
        def forgot_password():
            self._forgot_password_dialog(data, selected_company.get())
        ttk.Button(btns, text="Login", command=do_login).pack(side="left", padx=6)
        ttk.Button(btns, text="Forgot Password", command=forgot_password).pack(side="left", padx=6)
        ttk.Button(btns, text="Create Company", command=create_company).pack(side="left", padx=6)

    def _set_recovery_pin_dialog(self, data, company_name, username):
        dlg = tk.Toplevel(self.root)
        dlg.title("Set Recovery PIN")
        dlg.geometry("420x240")
        dlg.transient(self.root)
        dlg.grab_set()
        f = ttk.Frame(dlg, padding="12")
        f.pack(fill="both", expand=True)
        pin_var = tk.StringVar()
        pin2_var = tk.StringVar()
        ttk.Label(f, text="Recovery PIN").pack(anchor="w")
        ttk.Entry(f, textvariable=pin_var, show="*").pack(fill="x", pady=6)
        ttk.Label(f, text="Confirm Recovery PIN").pack(anchor="w")
        ttk.Entry(f, textvariable=pin2_var, show="*").pack(fill="x", pady=6)
        def save():
            pin = pin_var.get().strip()
            pin2 = pin2_var.get().strip()
            if not pin or pin != pin2:
                messagebox.showerror("Error", "PINs do not match")
                return
            for c in data.get("companies", []):
                if c.get("name") == company_name and c.get("username") == username:
                    rs = uuid.uuid4().hex
                    c["recovery_salt"] = rs
                    c["recovery_hash"] = self._hash_password(pin, rs)
                    break
            self._save_companies(data)
            dlg.destroy()
        b = ttk.Frame(f)
        b.pack(pady=10)
        ttk.Button(b, text="Save", command=save).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        _fit_window(dlg, 420, 240, mode="compact", remember_key="set_recovery_pin")
        self.root.wait_window(dlg)

    def _forgot_password_dialog(self, data, default_company=""):
        dlg = tk.Toplevel(self.root)
        dlg.title("Reset Password")
        dlg.geometry("520x320")
        dlg.transient(self.root)
        dlg.grab_set()
        f = ttk.Frame(dlg, padding="12")
        f.pack(fill="both", expand=True)
        companies = [c.get("name") for c in data.get("companies", [])]
        company_var = tk.StringVar(value=(default_company or (data.get("active") or (companies[0] if companies else ""))))
        user_var = tk.StringVar(value=(data.get("last_username") or ""))
        pin_var = tk.StringVar()
        newp_var = tk.StringVar()
        newp2_var = tk.StringVar()
        ttk.Label(f, text="Company").pack(anchor="w")
        ttk.Combobox(f, textvariable=company_var, values=companies, state="readonly").pack(fill="x", pady=4)
        ttk.Label(f, text="Username").pack(anchor="w")
        ttk.Entry(f, textvariable=user_var).pack(fill="x", pady=4)
        ttk.Label(f, text="Recovery PIN").pack(anchor="w")
        ttk.Entry(f, textvariable=pin_var, show="*").pack(fill="x", pady=4)
        ttk.Label(f, text="New Password").pack(anchor="w")
        ttk.Entry(f, textvariable=newp_var, show="*").pack(fill="x", pady=4)
        ttk.Label(f, text="Confirm New Password").pack(anchor="w")
        ttk.Entry(f, textvariable=newp2_var, show="*").pack(fill="x", pady=4)
        def reset():
            comp = company_var.get().strip()
            uname = user_var.get().strip()
            pin = pin_var.get().strip()
            newp = newp_var.get().strip()
            newp2 = newp2_var.get().strip()
            if not comp or not uname or not pin or not newp or newp != newp2:
                messagebox.showerror("Error", "Please fill all fields and ensure passwords match")
                return
            rec = None
            for c in data.get("companies", []):
                if c.get("name") == comp and c.get("username") == uname:
                    rec = c
                    break
            if not rec:
                messagebox.showerror("Error", "Account not found")
                return
            rh = rec.get("recovery_hash", "")
            rs = rec.get("recovery_salt", "")
            if not rh or not rs:
                messagebox.showerror("Error", "No Recovery PIN is set for this account.\nLogin normally and set one, or contact the company owner.")
                return
            if not self._verify_password(rh, rs, pin):
                messagebox.showerror("Error", "Recovery PIN is incorrect")
                return
            salt = uuid.uuid4().hex
            rec["salt"] = salt
            rec["password_hash"] = self._hash_password(newp, salt)
            self._save_companies(data)
            dlg.destroy()
            try:
                messagebox.showinfo("Success", "Password reset successfully. You can login now.")
            except Exception:
                pass
        b = ttk.Frame(f)
        b.pack(pady=10, fill="x")
        ttk.Button(b, text="Reset Password", command=reset).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        _fit_window(dlg, 520, 320, mode="dialog", remember_key="forgot_password")
        self.root.wait_window(dlg)

    def _show_change_password_dialog(self):
        data = self._load_companies()
        comp = getattr(self, "current_company", None) or data.get("active")
        if not comp:
            messagebox.showerror("Error", "No company selected")
            return
        rec = None
        for c in data.get("companies", []):
            if c.get("name") == comp:
                rec = c
                break
        if not rec:
            messagebox.showerror("Error", "Company not found")
            return
        dlg = tk.Toplevel(self.root)
        dlg.title("Change Password")
        dlg.geometry("420x240")
        dlg.transient(self.root)
        dlg.grab_set()
        f = ttk.Frame(dlg, padding="12"); f.pack(fill="both", expand=True)
        cur_var = tk.StringVar(); new_var = tk.StringVar()
        ttk.Label(f, text="Current Password").pack(anchor="w"); ttk.Entry(f, textvariable=cur_var, show="*").pack(fill="x", pady=6)
        ttk.Label(f, text="New Password").pack(anchor="w"); ttk.Entry(f, textvariable=new_var, show="*").pack(fill="x", pady=6)
        def save_pw():
            cur = cur_var.get().strip(); newp = new_var.get().strip()
            if not cur or not newp:
                return
            if not self._verify_password(rec.get("password_hash",""), rec.get("salt",""), cur):
                messagebox.showerror("Error", "Current password invalid")
                return
            salt = uuid.uuid4().hex
            rec["salt"] = salt
            rec["password_hash"] = self._hash_password(newp, salt)
            for i, c in enumerate(data["companies"]):
                if c.get("name") == comp:
                    data["companies"][i] = rec
                    break
            self._save_companies(data)
            dlg.destroy()
        b = ttk.Frame(f); b.pack(pady=10)
        ttk.Button(b, text="Save", command=save_pw).pack(side="left", padx=4)
        ttk.Button(b, text="Cancel", command=dlg.destroy).pack(side="left", padx=4)
        _fit_window(dlg, 420, 240, mode="compact", remember_key="change_password")

    def show_notifications_dialog(self):
        notes = self.collect_notifications()
        dlg = tk.Toplevel(self.root)
        dlg.title("Notifications and Alerts")
        dlg.geometry("700x520")
        dlg.transient(self.root)
        dlg.resizable(True, True)
        palette = self._theme_palette()
        try:
            dlg.configure(bg=palette["bg"])
        except Exception:
            pass
        sf = ScrollableFrame(dlg)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        header = ttk.Frame(container)
        header.pack(fill='x', pady=(8, 4))
        ttk.Label(header, text="Notifications", font=('Helvetica', 18, 'bold')).pack(side='left')
        ttk.Button(header, text="Refresh", command=lambda: [dlg.destroy(), self.show_notifications_dialog()]).pack(side='right')
        sections = {
            'Payments Due': [n for n in notes if n.get('type') == 'payment'],
            'Low Stock': [n for n in notes if n.get('type') == 'stock'],
            'Expiring Items': [n for n in notes if n.get('type') == 'expiry'],
            'License': [n for n in notes if n.get('type') == 'license']
        }
        for title, items in sections.items():
            lf = ttk.LabelFrame(container, text=title, padding="8")
            lf.pack(fill='x', pady=6)
            if not items:
                ttk.Label(lf, text="No alerts").pack(anchor='w')
                continue
            for it in items:
                if title == 'Payments Due':
                    inv = it.get('invoice_id') or ''
                    cl = it.get('client') or ''
                    amt = it.get('outstanding')
                    dd = it.get('due_in_days')
                    badge = "OVERDUE" if isinstance(dd, int) and dd < 0 else ("DUE SOON" if isinstance(dd, int) and dd <= 7 else "DUE")
                    color = palette["danger"] if badge == 'OVERDUE' else (palette["orange"] if badge == 'DUE SOON' else palette["text"])
                    row = ttk.Frame(lf)
                    row.pack(fill='x', pady=3)
                    ttk.Label(row, text=f"{badge}", foreground=color).pack(side='left', padx=(0,8))
                    lbl = ttk.Label(row, text=f"Invoice {inv} • {cl} • Outstanding AED {amt:,.2f}")
                    lbl.pack(side='left')
                    lbl.bind('<Button-1>', lambda e: self.show_invoice_management())
                elif title == 'Low Stock':
                    row = ttk.Frame(lf)
                    row.pack(fill='x', pady=3)
                    ttk.Label(row, text="LOW", foreground=palette["danger"]).pack(side='left', padx=(0,8))
                    lbl = ttk.Label(row, text=f"{it.get('name')} (ID {it.get('item_id')}) • Qty {it.get('qty')} / Min {it.get('min')}")
                    lbl.pack(side='left')
                    try:
                        lbl.bind('<Button-1>', lambda e, item_id=it.get('item_id'): self._open_product_detail(item_id))
                    except Exception:
                        pass
                elif title == 'Expiring Items':
                    row = ttk.Frame(lf)
                    row.pack(fill='x', pady=3)
                    days = it.get('days_left')
                    expired = isinstance(days, int) and days < 0
                    badge = 'EXPIRED' if expired else 'EXPIRY SOON'
                    color = palette["danger"] if expired else palette["orange"]
                    ttk.Label(row, text=badge, foreground=color).pack(side='left', padx=(0,8))
                    txt = f"{it.get('name')} (ID {it.get('item_id')}) • Expiry {it.get('expiry_date')}"
                    if isinstance(days, int):
                        if expired:
                            txt += f" • {abs(days)} days overdue"
                        else:
                            txt += f" • {days} days left"
                    lbl = ttk.Label(row, text=txt)
                    lbl.pack(side='left')
                    try:
                        lbl.bind('<Button-1>', lambda e, item_id=it.get('item_id'): self._open_product_detail(item_id))
                    except Exception:
                        pass
                elif title == 'License':
                    days = it.get('days_left')
                    badge = "EXPIRING" if isinstance(days, int) and days <= 14 else "INFO"
                    color = palette["orange"] if badge == 'EXPIRING' else palette["text"]
                    row = ttk.Frame(lf)
                    row.pack(fill='x', pady=3)
                    ttk.Label(row, text=badge, foreground=color).pack(side='left', padx=(0,8))
                    if isinstance(days, int):
                        ttk.Label(row, text=f"License expires on {it.get('expires')} • {days} days left").pack(side='left')
                    else:
                        ttk.Label(row, text=f"License expires on {it.get('expires')}").pack(side='left')
        btns = ttk.Frame(container)
        btns.pack(fill='x', pady=10)
        ttk.Button(btns, text="Close", command=dlg.destroy).pack(side='right')
        _fit_window(dlg, 700, 520, mode="large", remember_key="notifications_dialog")
    
    def show_transaction_manager(self):
        """Show the transaction management system"""
        try:
            if not self._guard_module_access("accounting", "Transaction Manager", admin_only=True):
                return
            if TRANSACTION_MANAGER_AVAILABLE:
                transaction_manager = TransactionManager(self.manager.invoice_folder)
                TransactionDialog(self.root, transaction_manager)
            else:
                messagebox.showerror("Error", "Transaction Manager is not available. Please check if transaction_manager.py is in the same directory.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open transaction manager: {e}")

    def show_gaap_dashboard(self):
        """Open the professional GAAP-compliant Financial Reporting Centre"""
        try:
            if not self._guard_module_access("reports", "GAAP Financial Reports", admin_only=True):
                return
            if not GAAP_DASHBOARD_AVAILABLE:
                messagebox.showerror("Error", "GAAP Reporting Dashboard is not available. Please check if gaap_reporting_dashboard.py is in the same directory.")
                return
            GAAPReportingDashboard(self.root, self.manager)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to open GAAP Reporting Dashboard: {e}")

    def show_quotations(self):
        """Open the Quotation Management Centre"""
        try:
            if not self._guard_module_access("invoice_management", "Quotations"):
                return
            if not QUOTATION_MANAGER_AVAILABLE:
                messagebox.showerror("Error", "Quotation Manager is not available. Please check if quotation_manager.py is in the same directory.")
                return
            QuotationManagerDialog(self.root, self.manager)
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Error", f"Failed to open Quotation Manager: {e}")

    def show_inventory_system(self):
        """Show the inventory system"""
        if not self._guard_module_access("inventory", "Inventory"):
            return
        if not INVENTORY_AVAILABLE:
            messagebox.showerror("Error", "Inventory system is not available")
            return
            
        # Clear existing widgets and show inventory system
        for widget in self.root.winfo_children():
            widget.destroy()
        
        sf = ScrollableFrame(self.root)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        
        inventory_system = InventorySystem(container, self.manager.invoice_folder, self.manager)
        
        # Connect the back button to main menu
        inventory_system.show_main_menu_callback = self.create_main_menu
        _fit_window(self.root, 1000, 700, mode="workspace", remember_key="main_window")

    def show_temperature_log(self):
        if not self._guard_module_access("temperature_log", "Temperature Log"):
            return
        if not TEMPERATURE_LOG_AVAILABLE:
            messagebox.showerror("Error", "Temperature Log module is not available")
            return
        for widget in self.root.winfo_children():
            widget.destroy()
        sf = ScrollableFrame(self.root)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        TemperatureLogScreen(container, self.manager, back_callback=self.create_main_menu)
        _fit_window(self.root, 1000, 700, mode="workspace", remember_key="main_window")

    def show_inventory_report(self):
        if not self._guard_module_access("reports", "Inventory Reports"):
            return
        if not INVENTORY_REPORT_AVAILABLE:
            messagebox.showerror("Error", "Inventory Report module is not available")
            return
        for widget in self.root.winfo_children():
            widget.destroy()
        sf = ScrollableFrame(self.root)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        InventoryReportScreen(container, self.manager, back_callback=self.create_main_menu)
        _fit_window(self.root, 1100, 750, mode="workspace", remember_key="main_window")
    
    def show_employee_manager(self):
        """Show employee manager (admin only)"""
        try:
            if not self._guard_module_access("employee_manager", "Employee Manager", admin_only=True):
                return
            from employee_manager import EmployeeManagerDialog
            EmployeeManagerDialog(self.root, self.manager)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open employee manager: {e}")
    
    def show_company_report(self):
        """Show company financial report dialog"""
        if not self._guard_module_access("reports", "Company Financial Report"):
            return
        CompanyReportDialog(self.root, self.manager)
    
    def show_invoice_management(self):
        """Show the invoice management screen (existing functionality)"""
        if not self._guard_module_access("invoice_management", "Invoice Management"):
            return
        # This will show the original invoice management interface
        self.setup_ui()
    
    def show_purchases(self):
        """Show purchases management dialog"""
        if not self._guard_module_access("purchases", "Purchases"):
            return
        PurchasesDialog(self.root, self.manager)
    
    def show_reports(self):
        if not self._guard_module_access("reports", "Reports"):
            return
        for widget in self.root.winfo_children():
            widget.destroy()
        sf = ScrollableFrame(self.root)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        back_frame = ttk.Frame(container)
        back_frame.pack(fill='x', padx=0, pady=0)
        ttk.Button(back_frame, text="← Back to Main Menu", command=self.create_main_menu).pack(anchor='w')
        main = ttk.Frame(container, padding="0")
        main.pack(fill='both', expand=True)
        ttk.Label(main, text="Reports", font=('Helvetica', 20, 'bold')).pack()

        # ===== Recommended: Professional GAAP Financial Centre =====
        hero = ttk.LabelFrame(main, text="✨ Recommended — Professional Financial Reporting Centre (GAAP Compliant)", padding="10")
        hero.pack(fill='x', pady=(10, 6))
        ttk.Label(hero,
                  text="The all-in-one accountant-style dashboard: P&L, Balance Sheet, Aging Receivables, "
                       "Revenue by Client, Stock Valuation, Supplier Spend, Cash Flow, and GL Audit — with "
                       "plain-English explanations, prior-period comparisons, and branded Excel/PDF exports.",
                  wraplength=820, foreground="#1e3a5f").pack(anchor='w', pady=(0, 8))
        hero_btns = ttk.Frame(hero)
        hero_btns.pack(fill='x')
        ttk.Button(hero_btns, text="📈 Open GAAP Financial Reports", style="Primary.TButton",
                   command=self.show_gaap_dashboard).pack(side='left', padx=(0, 10))
        ttk.Button(hero_btns, text="📋 Open Quotation Manager",
                   command=self.show_quotations).pack(side='left')

        # Categories
        fin = ttk.LabelFrame(main, text="Financial (Legacy)", padding="0")
        fin.pack(fill='x')
        ttk.Button(fin, text="🧾 Trial Balance (NEW)",
                   command=lambda: self.ensure_admin_then(lambda: TrialBalanceDialog(self.root, self.manager))
                   ).pack(side='left', padx=6)
        ttk.Button(fin, text="Company Financial Statements", command=lambda: CompanyReportDialog(self.root, self.manager)).pack(side='left', padx=6)
        ttk.Button(fin, text="Sales & Client Reports", command=lambda: ReportsDialog(self.root, self.manager)).pack(side='left', padx=6)
        ttk.Button(fin, text="Sales & Services Outbound Reports", command=lambda: SalesServicesReportsDialog(self.root, self.manager)).pack(side='left', padx=6)
        ttk.Button(fin, text="AR Aging", command=self.show_ar_aging).pack(side='left', padx=6)
        ttk.Button(fin, text="Cash Flow", command=self.show_cash_flow).pack(side='left', padx=6)
        ops = ttk.LabelFrame(main, text="Operations", padding="0")
        ops.pack(fill='x')
        ttk.Button(ops, text="Inventory Reports", command=self.show_inventory_system).pack(side='left', padx=6)
        sal = ttk.LabelFrame(main, text="Salaries", padding="0")
        sal.pack(fill='x')
        ttk.Button(sal, text="Monthly Salary Report", command=lambda: self.ensure_admin_then(lambda: EmployeeManagerDialog(self.root, self.manager).monthly_salary_report())).pack(side='left', padx=6)
        _fit_window(self.root, 980, 700, mode="workspace", remember_key="main_window")

    def show_marketplace_manager(self):
        if not self._guard_module_access("ecommerce_marketplace", "E-Commerce Marketplace"):
            return
        mm = getattr(self.manager, 'marketplace_manager', None)
        if not mm and MARKETPLACE_MANAGER_AVAILABLE:
            try:
                invoice_folder = getattr(self.manager, 'invoice_folder', None)
                if invoice_folder:
                    self.manager.marketplace_manager = MarketplaceManager(invoice_folder)
                    mm = self.manager.marketplace_manager
            except Exception:
                mm = None
        if not mm:
            messagebox.showerror("Error", "E-commerce manager is not available in this mode")
            return
        for widget in self.root.winfo_children():
            widget.destroy()
        sf = ScrollableFrame(self.root)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        back_frame = ttk.Frame(container)
        back_frame.pack(fill='x', padx=0, pady=0)
        ttk.Button(back_frame, text="← Back to Main Menu", command=self.create_main_menu).pack(anchor='w')
        header = ttk.Frame(container, padding="10")
        header.pack(fill='x')
        ttk.Label(header, text="E-Commerce Marketplaces Manager", font=('Helvetica', 20, 'bold')).pack(anchor='w')
        ttk.Label(header, text="Manage platforms, consignments, orders, and settlements", font=('Helvetica', 10)).pack(anchor='w')
        main = ttk.Frame(container, padding="0")
        main.pack(fill='both', expand=True)
        notebook = ttk.Notebook(main)
        notebook.pack(fill='both', expand=True)
        platforms_frame = ttk.Frame(notebook)
        consignments_frame = ttk.Frame(notebook)
        orders_frame = ttk.Frame(notebook)
        settlements_frame = ttk.Frame(notebook)
        analytics_frame = ttk.Frame(notebook)
        notebook.add(platforms_frame, text="Platforms")
        notebook.add(consignments_frame, text="Consignments")
        notebook.add(orders_frame, text="Orders")
        notebook.add(settlements_frame, text="Settlements")
        notebook.add(analytics_frame, text="Analytics")

        pf_top = ttk.Frame(platforms_frame)
        pf_top.pack(fill='x', pady=(8, 4), padx=8)
        ttk.Button(pf_top, text="Add Platform", command=lambda: self._show_add_platform_dialog(mm, platforms_tree)).pack(side='left', padx=4)
        ttk.Button(pf_top, text="Edit Platform", command=lambda: self._edit_platform(mm, platforms_tree)).pack(side='left', padx=4)
        ttk.Button(pf_top, text="Delete Platform", command=lambda: self._delete_platform(mm, platforms_tree)).pack(side='left', padx=4)
        ttk.Button(pf_top, text="Refresh", command=lambda: self._refresh_platforms_tree(mm, platforms_tree)).pack(side='left', padx=4)
        columns = ("id", "name", "seller_id", "commission", "terms", "returns")
        platforms_tree = ttk.Treeview(platforms_frame, columns=columns, show="headings")
        for col, text, width in [
            ("id", "ID", 80),
            ("name", "Name", 160),
            ("seller_id", "Seller ID", 140),
            ("commission", "Commission %", 110),
            ("terms", "Payment Terms", 120),
            ("returns", "Return Policy", 120),
        ]:
            platforms_tree.heading(col, text=text)
            platforms_tree.column(col, width=width, anchor='center', stretch=True)
        platforms_tree.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self._refresh_platforms_tree(mm, platforms_tree)

        cf_top = ttk.Frame(consignments_frame)
        cf_top.pack(fill='x', pady=(8, 4), padx=8)
        ttk.Button(cf_top, text="Allocate Inventory", command=lambda: self._show_allocate_inventory_dialog(mm, consignments_tree)).pack(side='left', padx=4)
        ttk.Button(cf_top, text="Record Settlement", command=lambda: self._record_consignment_settlement(mm, consignments_tree)).pack(side='left', padx=4)
        ttk.Button(cf_top, text="Delete Consignment", command=lambda: self._delete_consignment(mm, consignments_tree)).pack(side='left', padx=4)
        ttk.Button(cf_top, text="Refresh", command=lambda: self._refresh_consignments_tree(mm, consignments_tree)).pack(side='left', padx=4)
        cons_columns = ("consignment_id", "platform", "date", "total_qty", "sold_qty", "current_stock", "status")
        consignments_tree = ttk.Treeview(consignments_frame, columns=cons_columns, show="headings")
        for col, text, width in [
            ("consignment_id", "Consignment", 140),
            ("platform", "Platform", 120),
            ("date", "Date", 100),
            ("total_qty", "Total Qty", 90),
            ("sold_qty", "Sold", 80),
            ("current_stock", "Current", 90),
            ("status", "Status", 90),
        ]:
            consignments_tree.heading(col, text=text)
            consignments_tree.column(col, width=width, anchor='center', stretch=True)
        consignments_tree.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self._refresh_consignments_tree(mm, consignments_tree)

        of_top = ttk.Frame(orders_frame)
        of_top.pack(fill='x', pady=(8, 4), padx=8)
        ttk.Button(of_top, text="Import Orders CSV", command=lambda: self._import_orders_csv(mm, orders_tree)).pack(side='left', padx=4)
        ttk.Button(of_top, text="Delete Order", command=lambda: self._delete_order(mm, orders_tree)).pack(side='left', padx=4)
        ttk.Button(of_top, text="Refresh", command=lambda: self._refresh_orders_tree(mm, orders_tree)).pack(side='left', padx=4)
        orders_columns = ("order_id", "platform", "date", "customer", "item_count", "total", "status")
        orders_tree = ttk.Treeview(orders_frame, columns=orders_columns, show="headings")
        for col, text, width in [
            ("order_id", "Order ID", 140),
            ("platform", "Platform", 120),
            ("date", "Date", 100),
            ("customer", "Customer", 160),
            ("item_count", "Items", 70),
            ("total", "Total (AED)", 110),
            ("status", "Status", 100),
        ]:
            orders_tree.heading(col, text=text)
            orders_tree.column(col, width=width, anchor='center', stretch=True)
        orders_tree.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self._refresh_orders_tree(mm, orders_tree)

        sf_top = ttk.Frame(settlements_frame)
        sf_top.pack(fill='x', pady=(8, 4), padx=8)
        ttk.Button(sf_top, text="Refresh", command=lambda: self._refresh_settlements_tree(mm, settlements_tree)).pack(side='left', padx=4)
        ttk.Button(sf_top, text="Import Settlement CSV", command=lambda: self._import_settlements_csv(mm, settlements_tree)).pack(side='left', padx=4)
        ttk.Button(sf_top, text="Set Payment Today", command=lambda: self._set_selected_settlement_today(mm, settlements_tree)).pack(side='left', padx=4)
        ttk.Button(sf_top, text="Generate Invoice", command=lambda: self._generate_settlement_invoice(mm, settlements_tree)).pack(side='left', padx=4)
        ttk.Button(sf_top, text="Delete Settlement", command=lambda: self._delete_settlement(mm, settlements_tree)).pack(side='left', padx=4)
        ttk.Button(sf_top, text="Mark As Paid", command=lambda: self._mark_selected_settlement_paid(mm, settlements_tree)).pack(side='left', padx=4)
        qty_frame = ttk.Frame(sf_top)
        qty_frame.pack(side='right')
        ttk.Label(qty_frame, text="Qty").pack(side='left')
        self.partial_qty_var = tk.StringVar(value="0")
        qty_entry = ttk.Entry(qty_frame, textvariable=self.partial_qty_var, width=6)
        qty_entry.pack(side='left', padx=(2, 8))
        ttk.Label(qty_frame, text="Unit").pack(side='left')
        self.partial_unit_var = tk.StringVar(value="0")
        unit_entry = ttk.Entry(qty_frame, textvariable=self.partial_unit_var, width=8)
        unit_entry.pack(side='left', padx=(2, 8))
        ttk.Label(qty_frame, text="Amount").pack(side='left')
        self.partial_amount_var = tk.StringVar(value="0.00")
        amt_lbl = ttk.Label(qty_frame, textvariable=self.partial_amount_var, width=10)
        amt_lbl.pack(side='left', padx=(2, 8))
        self.partial_method_var = tk.StringVar(value="Transfer")
        self.partial_account_var = tk.StringVar(value="Bank")
        method_box = ttk.Combobox(qty_frame, textvariable=self.partial_method_var, values=["Transfer", "Cash", "Card"], state="readonly", width=10)
        method_box.pack(side='left', padx=(2, 8))
        account_entry = ttk.Entry(qty_frame, textvariable=self.partial_account_var, width=10)
        account_entry.pack(side='left', padx=(2, 8))
        self.partial_notes_var = tk.StringVar(value="")
        notes_entry = ttk.Entry(qty_frame, textvariable=self.partial_notes_var, width=18)
        notes_entry.pack(side='left', padx=(2, 8))
        ttk.Button(qty_frame, text="Apply Partial Payment", command=lambda: self._apply_partial_settlement_payment(mm, settlements_tree)).pack(side='left', padx=4)
        ttk.Button(qty_frame, text="View Statement", command=lambda: self._show_settlement_statement(mm, settlements_tree)).pack(side='left', padx=4)
        def _recalc_amount(*args):
            try:
                q = int(self.partial_qty_var.get() or "0")
                u = float(self.partial_unit_var.get() or "0")
                self.partial_amount_var.set(f"{q * u:.2f}")
            except Exception:
                self.partial_amount_var.set("0.00")
        self.partial_qty_var.trace_add("write", _recalc_amount)
        self.partial_unit_var.trace_add("write", _recalc_amount)
        columns_set = ("platform", "month", "total_sales", "fees", "refunds", "promotions", "net", "status", "invoice")
        settlements_tree = ttk.Treeview(settlements_frame, columns=columns_set, show="headings")
        for col, text, width in [
            ("platform", "Platform", 120),
            ("month", "Month", 80),
            ("total_sales", "Total Sales", 110),
            ("fees", "Fees", 90),
            ("refunds", "Refunds", 90),
            ("promotions", "Promotions", 110),
            ("net", "Net", 110),
            ("status", "Payment Status", 120),
            ("invoice", "Invoice", 130),
        ]:
            settlements_tree.heading(col, text=text)
            settlements_tree.column(col, width=width, anchor='center', stretch=True)
        settlements_tree.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        self._refresh_settlements_tree(mm, settlements_tree)

        analytics_top = ttk.Frame(analytics_frame)
        analytics_top.pack(fill='x', pady=(8, 4), padx=8)
        ttk.Label(analytics_top, text="Analytics Month (YYYY-MM):").pack(side='left', padx=(0, 4))
        self.ecom_analytics_month_var = tk.StringVar(value=datetime.now().strftime("%Y-%m"))
        ttk.Entry(analytics_top, textvariable=self.ecom_analytics_month_var, width=10).pack(side='left')
        ttk.Button(analytics_top, text="Refresh Analytics", command=lambda: self._refresh_ecommerce_analytics(mm, platform_perf_tree, product_perf_tree)).pack(side='left', padx=6)
        analytics_split = ttk.Frame(analytics_frame)
        analytics_split.pack(fill='both', expand=True, padx=8, pady=(0, 8))
        platform_perf_frame = ttk.LabelFrame(analytics_split, text="Platform Performance", padding="6")
        product_perf_frame = ttk.LabelFrame(analytics_split, text="Product Performance", padding="6")
        platform_perf_frame.pack(side='left', fill='both', expand=True, padx=(0, 4))
        product_perf_frame.pack(side='left', fill='both', expand=True, padx=(4, 0))
        platform_perf_tree = ttk.Treeview(platform_perf_frame, columns=("platform", "total_sales", "net", "settlements"), show="headings", height=8)
        for col, text, width in [
            ("platform", "Platform", 140),
            ("total_sales", "Total Sales", 110),
            ("net", "Net Settlement", 120),
            ("settlements", "Settlements", 100),
        ]:
            platform_perf_tree.heading(col, text=text)
            platform_perf_tree.column(col, width=width, anchor='center', stretch=True)
        platform_perf_tree.pack(fill='both', expand=True)
        product_perf_tree = ttk.Treeview(product_perf_frame, columns=("product_id", "quantity", "sales"), show="headings", height=8)
        for col, text, width in [
            ("product_id", "Product ID", 140),
            ("quantity", "Qty Sold", 100),
            ("sales", "Sales (AED)", 120),
        ]:
            product_perf_tree.heading(col, text=text)
            product_perf_tree.column(col, width=width, anchor='center', stretch=True)
        product_perf_tree.pack(fill='both', expand=True)
        self._refresh_ecommerce_analytics(mm, platform_perf_tree, product_perf_tree)

        alerts_frame = ttk.LabelFrame(container, text="Alerts", padding="6")
        alerts_frame.pack(fill='x', padx=8, pady=(4, 8))
        self._refresh_ecommerce_alerts(mm, alerts_frame)
        _fit_window(self.root, 1100, 800)

    def show_warehouse_management(self):
        if not WAREHOUSE_EXTENSION_AVAILABLE:
            messagebox.showerror("Error", "Warehouse extension is not available")
            return
        if not self._guard_module_access("warehouse_management", "Warehouse Management"):
            return
        current_company = str(getattr(self, "current_company", "") or "").strip()
        WarehouseManagementDialog(
            self.root,
            WarehouseManager(self.manager.invoice_folder, current_company),
            company_name=current_company,
            user_role=getattr(self, "user_role", "Staff"),
        )

    def show_subscription_details(self):
        try:
            license_path = os.path.join(getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaInvoice")), "license.json")
        except Exception:
            license_path = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaInvoice", "license.json")
        lic = {}
        if os.path.exists(license_path):
            try:
                with open(license_path, 'r') as f:
                    lic = json.load(f)
            except Exception:
                lic = {}
        name = str(lic.get('customer', '') or '')
        expires = str(lic.get('expires', '') or '')
        key = str(lic.get('key', '') or '')
        mid = f"{platform.node()}-{uuid.getnode()}"
        days_left = None
        try:
            if expires:
                d = datetime.strptime(expires, "%Y-%m-%d")
                days_left = (d - datetime.now()).days
        except Exception:
            days_left = None
        dlg = tk.Toplevel(self.root)
        dlg.title("Subscription Details")
        dlg.geometry("520x320")
        dlg.transient(self.root)
        dlg.resizable(True, True)
        frame = ttk.Frame(dlg, padding="16")
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text="Subscription Details", font=('Helvetica', 16, 'bold')).pack(pady=(0, 12))
        ttk.Label(frame, text=f"Name: {name if name else 'Not activated'}").pack(anchor='w', pady=4)
        ttk.Label(frame, text=f"Expiry: {expires if expires else 'N/A'}").pack(anchor='w', pady=4)
        ttl = f"Days Remaining: {days_left}" if isinstance(days_left, int) else "Days Remaining: N/A"
        ttk.Label(frame, text=ttl).pack(anchor='w', pady=4)
        ttk.Label(frame, text=f"Machine ID: {mid}").pack(anchor='w', pady=4)
        btns = ttk.Frame(frame)
        btns.pack(fill='x', pady=12)
        ttk.Button(btns, text="Request Extension", command=lambda: self._request_extension_email(name, expires, mid)).pack(side='left', padx=6)
        ttk.Button(btns, text="Close", command=dlg.destroy).pack(side='right', padx=6)
        _fit_window(dlg, 520, 320, mode="dialog", remember_key="subscription_details")

    def _request_extension_email(self, name, expires, mid):
        try:
            subject = urllib.parse.quote("HopePharma Subscription Extension Request")
            body_lines = [
                "Dear Admin,",
                "",
                f"Please extend my subscription.",
                f"Name: {name or 'N/A'}",
                f"Machine ID: {mid}",
                f"Current Expiry: {expires or 'N/A'}",
                "",
                "Thank you."
            ]
            body = urllib.parse.quote("\n".join(body_lines))
            url = f"mailto:atarekk99@gmail.com?subject={subject}&body={body}"
            webbrowser.open(url)
        except Exception:
            pass

    def _refresh_platforms_tree(self, marketplace_manager, tree):
        try:
            for row in tree.get_children():
                tree.delete(row)
            platforms = marketplace_manager.list_platforms()
            for p in platforms:
                terms = p.get("consignment_terms", {})
                tree.insert(
                    "",
                    "end",
                    values=(
                        p.get("id"),
                        p.get("name"),
                        p.get("seller_id"),
                        terms.get("commission_rate", 0),
                        terms.get("payment_terms_days", 0),
                        terms.get("return_policy_days", 0),
                    ),
                )
        except Exception:
            pass

    def _edit_platform(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a platform first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if not vals:
                return
            platform_id = vals[0]
            platform = marketplace_manager.get_platform(platform_id)
            if not platform:
                messagebox.showerror("Error", "Platform not found")
                return
            terms = platform.get("consignment_terms", {})
            dlg = tk.Toplevel(self.root)
            dlg.title("Edit Platform")
            dlg.geometry("420x260")
            dlg.transient(self.root)
            dlg.resizable(True, True)
            frame = ttk.Frame(dlg, padding="16")
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="Platform Name").grid(row=0, column=0, sticky="w", pady=4)
            name_var = tk.StringVar(value=platform.get("name", ""))
            ttk.Entry(frame, textvariable=name_var).grid(row=0, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Seller ID").grid(row=1, column=0, sticky="w", pady=4)
            seller_var = tk.StringVar(value=platform.get("seller_id", ""))
            ttk.Entry(frame, textvariable=seller_var).grid(row=1, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Commission %").grid(row=2, column=0, sticky="w", pady=4)
            comm_var = tk.StringVar(value=str(terms.get("commission_rate", 0)))
            ttk.Entry(frame, textvariable=comm_var).grid(row=2, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Payment Terms (days)").grid(row=3, column=0, sticky="w", pady=4)
            pay_var = tk.StringVar(value=str(terms.get("payment_terms_days", 30)))
            ttk.Entry(frame, textvariable=pay_var).grid(row=3, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Return Policy (days)").grid(row=4, column=0, sticky="w", pady=4)
            ret_var = tk.StringVar(value=str(terms.get("return_policy_days", 14)))
            ttk.Entry(frame, textvariable=ret_var).grid(row=4, column=1, sticky="ew", pady=4)
            frame.columnconfigure(1, weight=1)

            def save():
                try:
                    name = name_var.get().strip()
                    seller = seller_var.get().strip()
                    if not name or not seller:
                        messagebox.showerror("Error", "Name and Seller ID are required")
                        return
                    commission = float(comm_var.get() or 0)
                    payment_days = int(pay_var.get() or 30)
                    return_days = int(ret_var.get() or 14)
                    marketplace_manager.update_platform(
                        platform_id=platform_id,
                        name=name,
                        seller_id=seller,
                        commission_rate=commission,
                        payment_terms_days=payment_days,
                        return_policy_days=return_days,
                    )
                    self._refresh_platforms_tree(marketplace_manager, tree)
                    dlg.destroy()
                except Exception as e:
                    messagebox.showerror("Error", str(e))

            btns = ttk.Frame(frame)
            btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=10)
            ttk.Button(btns, text="Save", command=save).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=4)
            _fit_window(dlg, 420, 260, mode="compact", remember_key="edit_platform")
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _delete_platform(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a platform first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if not vals:
                return
            platform_id = vals[0]
            name = vals[1] if len(vals) > 1 else platform_id
            if not messagebox.askyesno("Confirm", f"Delete platform {name}?"):
                return
            marketplace_manager.delete_platform(platform_id)
            self._refresh_platforms_tree(marketplace_manager, tree)
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass


    def _show_add_platform_dialog(self, marketplace_manager, tree):
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Platform")
        dlg.geometry("420x260")
        dlg.transient(self.root)
        dlg.resizable(True, True)
        frame = ttk.Frame(dlg, padding="16")
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Platform Name").grid(row=0, column=0, sticky="w", pady=4)
        name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=name_var).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Seller ID").grid(row=1, column=0, sticky="w", pady=4)
        seller_var = tk.StringVar()
        ttk.Entry(frame, textvariable=seller_var).grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Commission %").grid(row=2, column=0, sticky="w", pady=4)
        comm_var = tk.StringVar(value="15")
        ttk.Entry(frame, textvariable=comm_var).grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Payment Terms (days)").grid(row=3, column=0, sticky="w", pady=4)
        pay_var = tk.StringVar(value="30")
        ttk.Entry(frame, textvariable=pay_var).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Return Policy (days)").grid(row=4, column=0, sticky="w", pady=4)
        ret_var = tk.StringVar(value="14")
        ttk.Entry(frame, textvariable=ret_var).grid(row=4, column=1, sticky="ew", pady=4)
        frame.columnconfigure(1, weight=1)

        def save():
            try:
                name = name_var.get().strip()
                seller = seller_var.get().strip()
                if not name or not seller:
                    messagebox.showerror("Error", "Name and Seller ID are required")
                    return
                commission = float(comm_var.get() or 0)
                payment_days = int(pay_var.get() or 30)
                return_days = int(ret_var.get() or 14)
                marketplace_manager.add_platform(
                    name=name,
                    seller_id=seller,
                    commission_rate=commission,
                    payment_terms_days=payment_days,
                    return_policy_days=return_days,
                )
                self._refresh_platforms_tree(marketplace_manager, tree)
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        btns = ttk.Frame(frame)
        btns.grid(row=5, column=0, columnspan=2, sticky="e", pady=10)
        ttk.Button(btns, text="Save", command=save).pack(side="right", padx=4)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=4)
        _fit_window(dlg, 420, 260, mode="compact", remember_key="add_platform")

    def _refresh_consignments_tree(self, marketplace_manager, tree):
        try:
            for row in tree.get_children():
                tree.delete(row)
            consignments = marketplace_manager.list_consignments()
            platforms = {p.get("id"): p for p in marketplace_manager.list_platforms()}
            for c in consignments:
                s = c.get("summary", {})
                pid = c.get("platform_id")
                p = platforms.get(pid, {})
                tree.insert(
                    "",
                    "end",
                    values=(
                        c.get("consignment_id"),
                        p.get("name", pid),
                        c.get("date"),
                        s.get("total_quantity", 0),
                        s.get("sold_quantity", 0),
                        s.get("current_stock", 0),
                        c.get("status", "active"),
                    ),
                )
        except Exception:
            pass

    def _refresh_orders_tree(self, marketplace_manager, tree):
        try:
            for row in tree.get_children():
                tree.delete(row)
            orders = marketplace_manager.list_orders()
            platforms = {p.get("id"): p for p in marketplace_manager.list_platforms()}
            for o in orders:
                pid = o.get("platform_id")
                p = platforms.get(pid, {})
                platform_name = p.get("name", pid)
                items = o.get("items", [])
                item_count = len(items)
                tree.insert(
                    "",
                    "end",
                    values=(
                        o.get("order_id"),
                        platform_name,
                        o.get("order_date"),
                        (o.get("customer") or {}).get("name", ""),
                        item_count,
                        round(float(o.get("total_amount", 0) or 0), 2),
                        o.get("status", ""),
                    ),
                )
        except Exception:
            pass

    def _record_consignment_settlement(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a consignment first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if not vals:
                return
            consignment_id = vals[0]
            all_cons = marketplace_manager.list_consignments()
            consignment = None
            for c in all_cons:
                if c.get("consignment_id") == consignment_id:
                    consignment = c
                    break
            if not consignment:
                messagebox.showerror("Error", "Consignment not found")
                return
            platforms = {p.get("id"): p for p in marketplace_manager.list_platforms()}
            pid = consignment.get("platform_id")
            platform = platforms.get(pid, {})
            summary = consignment.get("summary", {})
            total_qty = int(summary.get("total_quantity", 0))
            current_sold = int(summary.get("sold_quantity", 0))
            default_sold = total_qty if total_qty > 0 and current_sold == 0 else current_sold
            dlg = tk.Toplevel(self.root)
            dlg.title("Record Consignment Settlement")
            dlg.geometry("520x260")
            dlg.transient(self.root)
            dlg.resizable(True, True)
            frame = ttk.Frame(dlg, padding="16")
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="Platform").grid(row=0, column=0, sticky="w", pady=4)
            ttk.Label(frame, text=platform.get("name", pid)).grid(row=0, column=1, sticky="w", pady=4)
            ttk.Label(frame, text="Consignment").grid(row=1, column=0, sticky="w", pady=4)
            ttk.Label(frame, text=consignment_id).grid(row=1, column=1, sticky="w", pady=4)
            ttk.Label(frame, text="Quantity sent").grid(row=2, column=0, sticky="w", pady=4)
            ttk.Label(frame, text=str(total_qty)).grid(row=2, column=1, sticky="w", pady=4)
            ttk.Label(frame, text="Quantity sold").grid(row=3, column=0, sticky="w", pady=4)
            sold_var = tk.StringVar(value=str(default_sold if default_sold > 0 else total_qty))
            ttk.Entry(frame, textvariable=sold_var).grid(row=3, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Gross sales (AED)").grid(row=4, column=0, sticky="w", pady=4)
            gross_var = tk.StringVar()
            ttk.Entry(frame, textvariable=gross_var, state="readonly").grid(row=4, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Platform fees (AED)").grid(row=5, column=0, sticky="w", pady=4)
            fees_var = tk.StringVar()
            ttk.Entry(frame, textvariable=fees_var, state="readonly").grid(row=5, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Refunds (AED)").grid(row=6, column=0, sticky="w", pady=4)
            refunds_var = tk.StringVar()
            ttk.Entry(frame, textvariable=refunds_var).grid(row=6, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Promotions (AED)").grid(row=7, column=0, sticky="w", pady=4)
            promos_var = tk.StringVar()
            ttk.Entry(frame, textvariable=promos_var).grid(row=7, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Net amount received (AED)").grid(row=8, column=0, sticky="w", pady=4)
            net_var = tk.StringVar()
            ttk.Entry(frame, textvariable=net_var, state="readonly").grid(row=8, column=1, sticky="ew", pady=4)
            ttk.Label(frame, text="Payment date").grid(row=9, column=0, sticky="w", pady=4)
            pay_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
            ttk.Entry(frame, textvariable=pay_date_var).grid(row=9, column=1, sticky="ew", pady=4)
            frame.columnconfigure(1, weight=1)

            def _recalc(*_args):
                try:
                    sold_q = int(sold_var.get() or "0")
                except Exception:
                    sold_q = 0
                gross_amount = 0.0
                try:
                    items = consignment.get("items", [])
                    if items:
                        alloc_total = sum(float(i.get("platform_price", i.get("selling_price", 0)) or 0) * float(i.get("quantity", 0) or 0) for i in items)
                        if alloc_total > 0 and total_qty > 0:
                            gross_amount = (sold_q / total_qty) * alloc_total
                        else:
                            price = float(items[0].get("platform_price", items[0].get("selling_price", 0)) or 0)
                            gross_amount = sold_q * price
                except Exception:
                    pass
                try:
                    gross_var.set(f"{gross_amount:.2f}")
                except Exception:
                    pass
                try:
                    commission_rate = marketplace_manager.get_platform_commission_rate(pid)
                except Exception:
                    commission_rate = 0.0
                fees_amount = round(gross_amount * commission_rate, 2)
                try:
                    fees_var.set(f"{fees_amount:.2f}")
                except Exception:
                    pass
                try:
                    refunds_amount = float(refunds_var.get() or "0")
                except Exception:
                    refunds_amount = 0.0
                try:
                    promos_amount = float(promos_var.get() or "0")
                except Exception:
                    promos_amount = 0.0
                net_amount = gross_amount - fees_amount - refunds_amount - promos_amount
                try:
                    net_var.set(f"{net_amount:.2f}")
                except Exception:
                    pass

            try:
                sold_var.trace_add("write", _recalc)
                refunds_var.trace_add("write", _recalc)
                promos_var.trace_add("write", _recalc)
            except Exception:
                pass
            try:
                _recalc()
            except Exception:
                pass

            def save():
                try:
                    sold_q = int(sold_var.get() or "0")
                    if sold_q <= 0:
                        messagebox.showerror("Error", "Quantity sold must be greater than zero")
                        return
                    if total_qty > 0 and sold_q > total_qty:
                        messagebox.showerror("Error", f"Sold quantity cannot exceed sent quantity ({total_qty})")
                        return
                    # Auto-calculate gross from consignment items if not provided
                    gross_amount_input = gross_var.get().strip()
                    gross_amount = float(gross_amount_input or "0")
                    if gross_amount <= 0:
                        try:
                            items = consignment.get("items", [])
                            if items:
                                # Compute proportional gross based on allocation prices
                                alloc_total = sum(float(i.get("platform_price", i.get("selling_price", 0)) or 0) * float(i.get("quantity", 0) or 0) for i in items)
                                if alloc_total > 0 and total_qty > 0:
                                    gross_amount = (sold_q / total_qty) * alloc_total
                                else:
                                    # Fallback to single item price if available
                                    price = float(items[0].get("platform_price", items[0].get("selling_price", 0)) or 0)
                                    gross_amount = sold_q * price
                            else:
                                gross_amount = 0.0
                            gross_var.set(f"{gross_amount:.2f}")
                        except Exception:
                            pass
                    # Auto-calculate platform fees from commission rate if not provided
                    fees_amount_input = fees_var.get().strip()
                    fees_amount = float(fees_amount_input or "0")
                    if fees_amount <= 0:
                        try:
                            commission_rate = marketplace_manager.get_platform_commission_rate(pid)
                            fees_amount = round(gross_amount * commission_rate, 2)
                            fees_var.set(f"{fees_amount:.2f}")
                        except Exception:
                            pass
                    refunds_amount = float(refunds_var.get() or "0")
                    promos_amount = float(promos_var.get() or "0")
                    # Auto-calculate net if not provided
                    net_amount_input = net_var.get().strip()
                    net_amount = float(net_amount_input or "0")
                    if net_amount <= 0:
                        net_amount = gross_amount - fees_amount - refunds_amount - promos_amount
                        net_var.set(f"{net_amount:.2f}")
                    if net_amount <= 0:
                        messagebox.showerror("Error", "Net amount must be greater than zero")
                        return
                    pay_date = pay_date_var.get().strip() or datetime.now().strftime("%Y-%m-%d")
                    month_str = pay_date[:7]
                    delta_sold = sold_q - current_sold
                    if delta_sold != 0:
                        marketplace_manager.update_consignment_stock(consignment_id, sold_delta=delta_sold)
                    # Record monthly settlement with linkage to allocation invoice
                    allocation_inv_no = (consignment.get("invoice_number") or "")
                    commission_rate = marketplace_manager.get_platform_commission_rate(pid)
                    marketplace_manager.record_monthly_settlement(
                        platform_id=pid,
                        month=month_str,
                        total_sales=gross_amount or net_amount,
                        platform_fees=fees_amount,
                        refunds=refunds_amount,
                        promotions=promos_amount,
                        payment_date=pay_date,
                        payment_status="pending",
                        allocation_invoice_number=allocation_inv_no or None,
                        payment_gross=gross_amount,
                        commission_rate=commission_rate,
                    )
                    # Add platform fees as cost to allocation invoice
                    try:
                        if allocation_inv_no and fees_amount > 0:
                            inv_obj_cost = self.manager.get_invoice(allocation_inv_no)
                            if inv_obj_cost:
                                try:
                                    current_costs_len = 0
                                    if hasattr(inv_obj_cost, "costs"):
                                        try:
                                            current_costs_len = len(inv_obj_cost.costs or [])
                                        except Exception:
                                            current_costs_len = 0
                                    cost_desc = f"Platform fees ({month_str})"
                                    ok_cost = inv_obj_cost.edit_cost(
                                        current_costs_len,
                                        cost_desc,
                                        float(fees_amount),
                                        account="Platform Fees",
                                        notes=f"Consignment {consignment_id}"
                                    )
                                    if ok_cost and hasattr(self.manager, "update_invoice"):
                                        _ = self.manager.update_invoice(inv_obj_cost)
                                except Exception:
                                    pass
                    except Exception:
                        pass
                    # Apply payment against the original allocation invoice (gross units price)
                    ok, msg = True, ""
                    applied_to = None
                    try:
                        if allocation_inv_no:
                            inv_obj = self.manager.get_invoice(allocation_inv_no)
                            if inv_obj:
                                inv_data = inv_obj.to_dict() if hasattr(inv_obj, "to_dict") else inv_obj
                                grand_total = float(inv_data.get("grand_total", 0) or 0)
                                total_paid = float(inv_data.get("total_paid", 0) or 0)
                                remaining = max(0.0, grand_total - total_paid)
                                payment_amount = min(remaining, gross_amount)
                                if payment_amount <= 0:
                                    ok, msg = False, "No outstanding balance to apply payment"
                                else:
                                    method = "Transfer"
                                    account = "Bank"
                                    notes = f"Consignment {consignment_id} settlement {month_str}"
                                    if hasattr(inv_obj, "add_payment"):
                                        ok_add, msg_add = inv_obj.add_payment(payment_amount, pay_date, method, notes, account)
                                        if not ok_add:
                                            ok, msg = False, msg_add
                                        else:
                                            applied_to = allocation_inv_no
                                            if hasattr(self.manager, "update_invoice"):
                                                _ = self.manager.update_invoice(inv_obj)
                                            else:
                                                _ = self.manager.add_invoice_from_dict(inv_obj.to_dict())
                                            try:
                                                self.manager.process_invoice_payment_with_balance(inv_data, payment_amount, account)
                                            except Exception:
                                                pass
                                    else:
                                        ok, msg = False, "Invoice object does not support payments"
                            else:
                                ok, msg = False, "Allocation invoice not found"
                        else:
                            # Fallback: apply payment to monthly settlement invoice
                            ok, msg = marketplace_manager.apply_settlement_payment(
                                self.manager,
                                pid,
                                month_str,
                                net_amount,
                                "Transfer",
                                "Bank",
                                f"Consignment {consignment_id} payment"
                            )
                    except Exception as e:
                        ok, msg = False, str(e)
                    self._refresh_consignments_tree(marketplace_manager, tree)
                    if ok:
                        try:
                            if applied_to:
                                inv_obj2 = self.manager.get_invoice(applied_to)
                                if inv_obj2:
                                    data2 = inv_obj2.to_dict() if hasattr(inv_obj2, "to_dict") else inv_obj2
                                    due2 = float(data2.get("grand_total", 0) or 0) - float(data2.get("total_paid", 0) or 0)
                                    messagebox.showinfo("Success", f"Settlement recorded. Allocation invoice {applied_to} updated. Outstanding AED {due2:.2f}")
                                else:
                                    messagebox.showinfo("Success", "Settlement recorded and payment applied")
                            else:
                                inv = marketplace_manager.ensure_settlement_invoice(self.manager, pid, month_str)
                                inv_id = getattr(inv, "invoice_id", None)
                                due = getattr(inv, "balance_due", None)
                                if inv_id is not None and due is not None:
                                    messagebox.showinfo("Success", f"Settlement recorded. Invoice {inv_id} updated. Outstanding AED {due:.2f}")
                                else:
                                    messagebox.showinfo("Success", "Settlement recorded")
                        except Exception:
                            pass
                    else:
                        try:
                            messagebox.showerror("Error", msg)
                        except Exception:
                            pass
                    # Audit log
                    try:
                        folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                        path = os.path.join(folder, 'audit_log.jsonl')
                        rec = {
                            'timestamp': datetime.now().isoformat(),
                            'action': 'marketplace_settlement_applied',
                            'details': {
                                'platform_id': pid,
                                'consignment_id': consignment_id,
                                'month': month_str,
                                'gross': gross_amount,
                                'fees': fees_amount,
                                'refunds': refunds_amount,
                                'promotions': promos_amount,
                                'payment_date': pay_date,
                                'applied_to_invoice': applied_to or inv_id,
                                'allocation_invoice_cost_added': bool(allocation_inv_no and fees_amount > 0)
                            }
                        }
                        with open(path, 'a') as f:
                            f.write(json.dumps(rec) + "\n")
                    except Exception:
                        pass
                    dlg.destroy()
                except Exception as e:
                    messagebox.showerror("Error", str(e))

            btns = ttk.Frame(frame)
            btns.grid(row=6, column=0, columnspan=2, sticky="e", pady=10)
            ttk.Button(btns, text="Save", command=save).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=4)
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _delete_consignment(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a consignment first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if not vals:
                return
            consignment_id = vals[0]
            if not messagebox.askyesno("Confirm", f"Delete consignment {consignment_id}?"):
                return
            cons = marketplace_manager.get_consignment(consignment_id)
            try:
                from inventory_system import InventoryManager
                folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                inv = InventoryManager(folder)
                items = (cons or {}).get("items") or []
                for it in items:
                    pid = it.get("product_id") or it.get("item_id")
                    qty = int(it.get("quantity", 0) or 0)
                    total_cost = float(it.get("total_cost", 0) or 0)
                    try:
                        obj = inv.get_item(pid)
                        if not obj:
                            candidates = inv.search_items(pid)
                            obj = candidates[0] if candidates else None
                        if obj and qty > 0:
                            obj.quantity = int(obj.quantity) + qty
                            obj.total_cost = float(obj.total_cost or 0) + total_cost
                            obj.last_updated = datetime.now().strftime("%Y-%m-%d")
                    except Exception:
                        pass
                try:
                    inv.save_inventory()
                except Exception:
                    pass
            except Exception:
                pass
            try:
                alloc_invoice_id = (cons or {}).get("invoice_number")
                if alloc_invoice_id:
                    if hasattr(self.manager, "delete_invoice_with_balance_adjustment"):
                        ok_del, msg_del = self.manager.delete_invoice_with_balance_adjustment(alloc_invoice_id)
                    else:
                        ok_del = self.manager.delete_invoice(alloc_invoice_id)
                        msg_del = "Invoice deleted"
                    try:
                        if ok_del:
                            messagebox.showinfo("Success", f"{msg_del}: {alloc_invoice_id}")
                        else:
                            messagebox.showerror("Error", f"Failed to delete invoice {alloc_invoice_id}: {msg_del}")
                    except Exception:
                        pass
                    try:
                        folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                        path = os.path.join(folder, 'audit_log.jsonl')
                        rec = {
                            'timestamp': datetime.now().isoformat(),
                            'action': 'consignment_deleted_with_invoice',
                            'details': {
                                'consignment_id': consignment_id,
                                'allocation_invoice_id': alloc_invoice_id
                            }
                        }
                        with open(path, 'a') as f:
                            f.write(json.dumps(rec) + "\n")
                    except Exception:
                        pass
            except Exception:
                pass
            marketplace_manager.delete_consignment(consignment_id)
            self._refresh_consignments_tree(marketplace_manager, tree)
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _delete_order(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select an order first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if not vals:
                return
            order_id = vals[0]
            if not messagebox.askyesno("Confirm", f"Delete order {order_id}?"):
                return
            marketplace_manager.delete_order(order_id)
            self._refresh_orders_tree(marketplace_manager, tree)
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _delete_settlement(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a settlement first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if len(vals) < 2:
                return
            platform_name = vals[0]
            month = vals[1]
            platform_id = None
            platform = None
            for p in marketplace_manager.list_platforms():
                if p.get("name") == platform_name or p.get("id") == platform_name:
                    platform = p
                    break
            if platform:
                platform_id = platform.get("id")
            else:
                platform_id = platform_name
            if not messagebox.askyesno("Confirm", f"Delete settlement for {platform_name} {month}?"):
                return
            # Attempt to reverse payment applied to allocation invoice if present
            try:
                settlements = marketplace_manager.list_settlements(platform_id=platform_id, month=month)
                if settlements:
                    s = settlements[0]
                    inv_no = s.get("allocation_invoice_number") or s.get("invoice_number")
                    amt = float(s.get("payment_gross", 0) or 0)
                    if inv_no and amt > 0:
                        inv_obj = self.manager.get_invoice(inv_no)
                        if inv_obj:
                            data = inv_obj.to_dict() if hasattr(inv_obj, "to_dict") else inv_obj
                            account = "Bank"
                            try:
                                self.manager.reverse_invoice_payment_with_balance(data, amt, account)
                            except Exception:
                                pass
            except Exception:
                pass
            marketplace_manager.delete_settlement(platform_id, month)
            self._refresh_settlements_tree(marketplace_manager, tree)
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _show_allocate_inventory_dialog(self, marketplace_manager, tree):
        if not INVENTORY_AVAILABLE:
            messagebox.showerror("Error", "Inventory system is not available")
            return
        try:
            invoice_folder = getattr(self.manager, "invoice_folder", None)
        except Exception:
            invoice_folder = None
        if not invoice_folder:
            messagebox.showerror("Error", "Invoice folder is not configured")
            return
        platforms = marketplace_manager.list_platforms()
        if not platforms:
            messagebox.showerror("Error", "Please add a platform first")
            return
        inv_manager = InventoryManager(invoice_folder)
        items = [i for i in inv_manager.items if getattr(i, "quantity", 0) > 0]
        if not items:
            messagebox.showerror("Error", "No inventory items with stock are available")
            return
        display_to_item = {}
        item_display_values = []
        for it in items:
            label = f"{it.item_id} - {it.name} (qty {it.quantity})"
            display_to_item[label] = it
            item_display_values.append(label)
        dlg = tk.Toplevel(self.root)
        dlg.title("Allocate Inventory to Platform")
        dlg.geometry("700x380")
        dlg.transient(self.root)
        dlg.resizable(True, True)
        frame = ttk.Frame(dlg, padding="16")
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Platform").grid(row=0, column=0, sticky="w", pady=4)
        platform_var = tk.StringVar()
        platform_names = [p.get("name", p.get("id")) for p in platforms]
        platform_combo = ttk.Combobox(frame, textvariable=platform_var, values=platform_names, state="readonly")
        platform_combo.grid(row=0, column=1, sticky="ew", pady=4)
        if not platform_names:
            platform_combo.current(0)
        rows_frame = ttk.Frame(frame)
        rows_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(8, 4))
        frame.columnconfigure(1, weight=1)
        rows_frame.columnconfigure(0, weight=3)
        rows_frame.columnconfigure(1, weight=1)
        rows_frame.columnconfigure(2, weight=2)
        rows_frame.columnconfigure(3, weight=1)
        ttk.Label(rows_frame, text="Item").grid(row=0, column=0, sticky="w", padx=2)
        ttk.Label(rows_frame, text="Qty").grid(row=0, column=1, sticky="w", padx=2)
        ttk.Label(rows_frame, text="Platform SKU").grid(row=0, column=2, sticky="w", padx=2)
        ttk.Label(rows_frame, text="Price (AED)").grid(row=0, column=3, sticky="w", padx=2)
        rows = []

        def add_row():
            row_index = len(rows) + 1
            item_display_var = tk.StringVar()
            item_combo = ttk.Combobox(rows_frame, textvariable=item_display_var, values=item_display_values, state="readonly")
            item_combo.grid(row=row_index, column=0, sticky="ew", padx=2, pady=2)
            qty_var = tk.StringVar(value="1")
            qty_entry = ttk.Entry(rows_frame, textvariable=qty_var)
            qty_entry.grid(row=row_index, column=1, sticky="ew", padx=2, pady=2)
            sku_var = tk.StringVar()
            sku_entry = ttk.Entry(rows_frame, textvariable=sku_var)
            sku_entry.grid(row=row_index, column=2, sticky="ew", padx=2, pady=2)
            price_var = tk.StringVar()
            price_entry = ttk.Entry(rows_frame, textvariable=price_var)
            price_entry.grid(row=row_index, column=3, sticky="ew", padx=2, pady=2)
            item_id_var = tk.StringVar()

            def on_item_selected(event=None):
                try:
                    display = item_display_var.get()
                    it = display_to_item.get(display)
                    if not it:
                        return
                    item_id_var.set(it.item_id)
                    if it.selling_price:
                        try:
                            price_var.set(str(it.selling_price))
                        except Exception:
                            pass
                    name = platform_var.get().strip()
                    if name and not sku_var.get().strip():
                        try:
                            sku_var.set(f"{name.upper()}-{it.item_id}")
                        except Exception:
                            pass
                except Exception:
                    pass

            if item_display_values:
                item_combo.current(0)
                item_display_var.set(item_display_values[0])
                on_item_selected()
            item_combo.bind("<<ComboboxSelected>>", on_item_selected)
            rows.append(
                {
                    "item_display_var": item_display_var,
                    "item_id_var": item_id_var,
                    "qty_var": qty_var,
                    "sku_var": sku_var,
                    "price_var": price_var,
                }
            )

        add_row()

        def save():
            try:
                name = platform_var.get().strip()
                if not name:
                    messagebox.showerror("Error", "Platform is required")
                    return
                platform = None
                for p in platforms:
                    if p.get("name") == name or p.get("id") == name:
                        platform = p
                        break
                if not platform:
                    messagebox.showerror("Error", "Selected platform not found")
                    return
                cons_items = []
                total_cost_all = 0.0
                subtotal = 0.0
                inv_manager_local = InventoryManager(invoice_folder)
                for row in rows:
                    display = row["item_display_var"].get().strip()
                    if not display:
                        continue
                    item_id = row["item_id_var"].get().strip()
                    qty_str = row["qty_var"].get()
                    sku = row["sku_var"].get().strip()
                    price_str = row["price_var"].get()
                    if not item_id or not qty_str or not sku or not price_str:
                        continue
                    try:
                        qty = int(qty_str)
                        price = float(price_str)
                    except Exception:
                        messagebox.showerror("Error", "Quantity and price must be valid numbers")
                        return
                    if qty <= 0 or price <= 0:
                        messagebox.showerror("Error", "Quantity and price must be greater than zero")
                        return
                    item = inv_manager_local.get_item(item_id)
                    if not item:
                        candidates = inv_manager_local.search_items(item_id)
                        if len(candidates) == 1:
                            item = candidates[0]
                        else:
                            messagebox.showerror("Error", f"Inventory item not found: {item_id}")
                            return
                    if item.quantity < qty:
                        messagebox.showerror("Error", f"Insufficient stock for {item.name}. Available: {item.quantity}")
                        return
                    cost_per = item.get_cost_per_item()
                    total_cost = cost_per * qty
                    remaining = item.quantity - qty
                    if remaining > 0:
                        item.total_cost = cost_per * remaining
                    else:
                        item.total_cost = 0
                    item.quantity = remaining
                    item.last_updated = datetime.now().strftime("%Y-%m-%d")
                    cons_items.append(
                        {
                            "product_id": item.item_id,
                            "quantity": qty,
                            "platform_sku": sku,
                            "selling_price": float(item.selling_price or price),
                            "platform_price": price,
                            "cost_per_item": cost_per,
                            "total_cost": total_cost,
                            "item_name": item.name,
                        }
                    )
                    total_cost_all += total_cost
                    subtotal += price * qty
                if not cons_items:
                    messagebox.showerror("Error", "Add at least one valid item")
                    return
                inv_manager_local.save_inventory()
                cons = marketplace_manager.create_consignment(platform.get("id"), cons_items)
                # Auto-generate allocation invoice for the full consignment
                try:
                    client_name = platform.get("name", platform.get("id"))
                    today = datetime.now().strftime("%Y-%m-%d")
                    # Determine due date from platform terms (default 30 days)
                    try:
                        terms_days = int((platform.get("consignment_terms") or {}).get("payment_terms_days", 30))
                    except Exception:
                        terms_days = 30
                    due_date = (datetime.now() + timedelta(days=terms_days)).strftime("%Y-%m-%d")
                    # Build invoice items from consignment items (non-taxable)
                    inv_items = []
                    subtotal_alloc = 0.0
                    total_cost_alloc = 0.0
                    for it in cons_items:
                        qty = float(it.get("quantity", 0) or 0)
                        price = float(it.get("platform_price", it.get("selling_price", 0)) or 0)
                        total = qty * price
                        inv_items.append({
                            "description": f"{it.get('item_name', it.get('product_id'))} ({it.get('platform_sku', '')})",
                            "quantity": qty,
                            "unit_price": price,
                            "total": total,
                            "taxable": False
                        })
                        subtotal_alloc += total
                        total_cost_alloc += float(it.get("total_cost", 0) or 0)
                    grand_total_alloc = subtotal_alloc  # VAT 0
                    profit_loss_alloc = grand_total_alloc - total_cost_alloc
                    invoice_id = self.manager.generate_invoice_id()
                    allocation_invoice = {
                        "invoice_id": invoice_id,
                        "invoice_type": "sales",
                        "client_name": client_name,
                        "client_trn": (platform.get("seller_id") or ""),
                        "client_emirate": "Dubai",
                        "date": today,
                        "payment_method": "Transfer",
                        "payment_due": f"{terms_days} Days" if terms_days in (30,60,90,120) else "30 Days",
                        "due_date": due_date,
                        "tax_rate": 0.0,
                        "items": inv_items,
                        "costs": [],
                        "subtotal": subtotal_alloc,
                        "taxable_amount": 0.0,
                        "non_taxable_amount": subtotal_alloc,
                        "tax_amount": 0.0,
                        "grand_total": grand_total_alloc,
                        "total_cost": total_cost_alloc,
                        "profit_loss": profit_loss_alloc,
                        "total_paid": 0.0,
                        "status": "Not Paid",
                        "notes": f"Marketplace allocation invoice for consignment {cons.get('consignment_id')}",
                        "currency": "AED"
                    }
                    ok = self.manager.add_invoice_from_dict(allocation_invoice)
                    if ok:
                        try:
                            marketplace_manager.set_consignment_invoice(cons.get("consignment_id"), invoice_id)
                        except Exception:
                            pass
                        messagebox.showinfo("Success", f"Consignment {cons.get('consignment_id')} created and invoice {invoice_id} generated for {client_name}\nTotal: AED {grand_total_alloc:.2f}")
                        # Audit log
                        try:
                            folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                            path = os.path.join(folder, 'audit_log.jsonl')
                            rec = {
                                'timestamp': datetime.now().isoformat(),
                                'action': 'allocation_invoice_created',
                                'details': {
                                    'consignment_id': cons.get('consignment_id'),
                                    'invoice_id': invoice_id,
                                    'platform': client_name,
                                    'grand_total': grand_total_alloc
                                }
                            }
                            with open(path, 'a') as f:
                                f.write(json.dumps(rec) + "\n")
                        except Exception:
                            pass
                    else:
                        messagebox.showerror("Error", "Consignment created but failed to create allocation invoice")
                except Exception:
                    pass
                self._refresh_consignments_tree(marketplace_manager, tree)
                dlg.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e))

        btns = ttk.Frame(frame)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=10)
        ttk.Button(btns, text="Add Item", command=add_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Allocate", command=save).pack(side="right", padx=4)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=4)

    def _refresh_settlements_tree(self, marketplace_manager, tree):
        try:
            for row in tree.get_children():
                tree.delete(row)
            settlements = marketplace_manager.list_settlements()
            platforms = {p.get("id"): p for p in marketplace_manager.list_platforms()}
            for s in settlements:
                pid = s.get("platform_id")
                p = platforms.get(pid, {})
                tree.insert(
                    "",
                    "end",
                    values=(
                        p.get("name", pid),
                        s.get("month"),
                        round(float(s.get("total_sales", 0) or 0), 2),
                        round(float(s.get("platform_fees", 0) or 0), 2),
                        round(float(s.get("refunds", 0) or 0), 2),
                        round(float(s.get("promotions", 0) or 0), 2),
                        round(float(s.get("net_settlement", 0) or 0), 2),
                        s.get("payment_status", ""),
                        s.get("invoice_number", ""),
                    ),
                )
        except Exception:
            pass

    def _mark_selected_settlement_paid(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a settlement first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if len(vals) < 2:
                return
            platform_name = vals[0]
            month = vals[1]
            platform = None
            for p in marketplace_manager.list_platforms():
                if p.get("name") == platform_name:
                    platform = p
                    break
            if not platform:
                messagebox.showerror("Error", "Platform not found")
                return
            marketplace_manager.mark_settlement_paid(platform.get("id"), month)
            try:
                try:
                    invoice_folder = getattr(self.manager, "invoice_folder", None)
                except Exception:
                    invoice_folder = None
                if invoice_folder and BALANCE_MANAGER_AVAILABLE:
                    settlements = marketplace_manager.list_settlements(platform_id=platform.get("id"), month=month)
                    if settlements:
                        s = settlements[0]
                        net_amount = float(s.get("net_settlement", 0) or 0)
                        if net_amount > 0:
                            from balance_manager import BalanceManager as _BM

                            bm = _BM(invoice_folder)
                            desc = f"Marketplace settlement {platform_name} {month}"
                            inv_no = s.get("invoice_number") or ""
                            bm.add_transaction("Bank", net_amount, "deposit", desc, reference_id=inv_no, reference_type="marketplace_settlement", date=s.get("payment_date"))
            except Exception:
                pass
            self._refresh_settlements_tree(marketplace_manager, tree)
        except Exception:
            pass

    def _set_selected_settlement_today(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a settlement first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if len(vals) < 2:
                return
            platform_name = vals[0]
            month = vals[1]
            platform = None
            for p in marketplace_manager.list_platforms():
                if p.get("name") == platform_name:
                    platform = p
                    break
            if not platform:
                messagebox.showerror("Error", "Platform not found")
                return
            marketplace_manager.set_settlement_payment_date(
                platform.get("id"), month, datetime.now().strftime("%Y-%m-%d"), status="due"
            )
            self._refresh_settlements_tree(marketplace_manager, tree)
            try:
                messagebox.showinfo(
                    "Updated",
                    f"Payment date set to today for {platform_name} {month}",
                )
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass
    
    def _apply_partial_settlement_payment(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a settlement first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if len(vals) < 2:
                return
            platform_name = vals[0]
            month = vals[1]
            platform = None
            for p in marketplace_manager.list_platforms():
                if p.get("name") == platform_name or p.get("id") == platform_name:
                    platform = p
                    break
            if not platform:
                messagebox.showerror("Error", "Platform not found")
                return
            try:
                amount = float(self.partial_amount_var.get() or "0")
            except Exception:
                amount = 0.0
            if amount <= 0:
                messagebox.showerror("Error", "Enter a valid quantity and unit price")
                return
            method = self.partial_method_var.get().strip() or "Transfer"
            account = self.partial_account_var.get().strip() or method
            notes = self.partial_notes_var.get().strip()
            ok, msg = marketplace_manager.apply_settlement_payment(
                self.manager,
                platform.get("id"),
                month,
                amount,
                method,
                account,
                notes
            )
            self._refresh_settlements_tree(marketplace_manager, tree)
            try:
                if ok:
                    inv = marketplace_manager.ensure_settlement_invoice(self.manager, platform.get("id"), month)
                    if inv:
                        message = f"Payment applied to invoice {inv.invoice_id}. Outstanding AED {float(inv.balance_due):.2f}"
                    else:
                        message = "Payment applied"
                    messagebox.showinfo("Success", message)
                else:
                    messagebox.showerror("Error", msg)
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass
    
    def _show_settlement_statement(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a settlement first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if len(vals) < 2:
                return
            platform_name = vals[0]
            month = vals[1]
            platform = None
            for p in marketplace_manager.list_platforms():
                if p.get("name") == platform_name or p.get("id") == platform_name:
                    platform = p
                    break
            if not platform:
                messagebox.showerror("Error", "Platform not found")
                return
            stmt = marketplace_manager.settlement_statement(self.manager, platform.get("id"), month)
            if not stmt:
                messagebox.showerror("Error", "Statement not available")
                return
            dlg = tk.Toplevel(self.root)
            dlg.title("Settlement Statement")
            dlg.geometry("640x420")
            dlg.transient(self.root)
            dlg.resizable(True, True)
            frame = ttk.Frame(dlg, padding="12")
            frame.pack(fill='both', expand=True)
            top = ttk.Frame(frame)
            top.pack(fill='x', pady=(0,8))
            status_text = "Paid" if float(stmt.get("outstanding_balance", 0) or 0) == 0 else "Not Paid"
            ttk.Label(top, text=f"Platform: {platform_name}").pack(side='left', padx=6)
            ttk.Label(top, text=f"Month: {month}").pack(side='left', padx=6)
            ttk.Label(top, text=f"Invoice: {stmt.get('invoice_id','')}").pack(side='left', padx=6)
            ttk.Label(top, text=f"Status: {status_text}").pack(side='left', padx=6)
            info = ttk.Frame(frame)
            info.pack(fill='x', pady=(0,8))
            ttk.Label(info, text=f"Original Amount: {float(stmt.get('original_amount',0) or 0):,.2f}").pack(side='left', padx=6)
            ttk.Label(info, text=f"Total Paid: {float(stmt.get('cumulative_payments',0) or 0):,.2f}").pack(side='left', padx=6)
            ttk.Label(info, text=f"Outstanding: {float(stmt.get('outstanding_balance',0) or 0):,.2f}").pack(side='left', padx=6)
            ttk.Label(info, text=f"Aging Days: {int(stmt.get('aging_days',0) or 0)}").pack(side='left', padx=6)
            tree_frame = ttk.LabelFrame(frame, text="Payment History", padding="6")
            tree_frame.pack(fill='both', expand=True)
            hist_cols = ("date","amount","method","account","notes")
            hist_tree = ttk.Treeview(tree_frame, columns=hist_cols, show="headings")
            for col, text, width in [
                ("date","Date",120),
                ("amount","Amount",100),
                ("method","Method",100),
                ("account","Account",120),
                ("notes","Notes",180),
            ]:
                hist_tree.heading(col, text=text)
                hist_tree.column(col, width=width, anchor='center')
            for pmt in stmt.get("payment_history", []):
                hist_tree.insert("", "end", values=(
                    pmt.get("date",""),
                    round(float(pmt.get("amount",0) or 0),2),
                    pmt.get("method",""),
                    pmt.get("account",""),
                    pmt.get("note","") or pmt.get("notes","")
                ))
            hist_tree.pack(fill='both', expand=True)
            btns = ttk.Frame(frame)
            btns.pack(fill='x', pady=8)
            ttk.Button(btns, text="Close", command=dlg.destroy).pack(side='right')
            _fit_window(dlg, 640, 420, mode="dialog", remember_key="settlement_statement")
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _refresh_ecommerce_analytics(self, marketplace_manager, platform_tree, product_tree):
        try:
            month = ""
            try:
                month = self.ecom_analytics_month_var.get().strip()
            except Exception:
                month = ""
            for row in platform_tree.get_children():
                platform_tree.delete(row)
            for row in product_tree.get_children():
                product_tree.delete(row)
            platform_rows = marketplace_manager.platform_performance_report(month=month or None)
            for r in platform_rows:
                platform_tree.insert(
                    "",
                    "end",
                    values=(
                        r.get("platform_name"),
                        round(float(r.get("total_sales", 0) or 0), 2),
                        round(float(r.get("net_settlement", 0) or 0), 2),
                        r.get("settlement_count", 0),
                    ),
                )
            product_rows = marketplace_manager.product_performance_report(platform_id=None, month=month or None)
            for r in product_rows:
                product_tree.insert(
                    "",
                    "end",
                    values=(
                        r.get("product_id"),
                        r.get("quantity_sold", 0),
                        round(float(r.get("sales", 0) or 0), 2),
                    ),
                )
        except Exception:
            pass

    def _refresh_ecommerce_alerts(self, marketplace_manager, alerts_frame):
        try:
            for w in alerts_frame.winfo_children():
                w.destroy()
            low_stock_alerts = []
            consignments = marketplace_manager.list_consignments()
            platforms = {p.get("id"): p for p in marketplace_manager.list_platforms()}
            for c in consignments:
                s = c.get("summary", {})
                current = int(s.get("current_stock", 0))
                if current <= 0:
                    continue
                if current <= 5:
                    pid = c.get("platform_id")
                    pname = platforms.get(pid, {}).get("name", pid)
                    low_stock_alerts.append(f"Low marketplace stock: {pname} consignment {c.get('consignment_id')} current {current}")
            overdue_settlements = []
            today_str = datetime.now().strftime("%Y-%m-%d")
            settlements = marketplace_manager.list_settlements()
            for s in settlements:
                status = s.get("payment_status", "")
                if status == "paid":
                    continue
                pay_date = s.get("payment_date") or ""
                if pay_date and pay_date < today_str:
                    overdue_settlements.append(f"Overdue settlement: {s.get('month')} for platform {s.get('platform_id')} status {status}")
            if not low_stock_alerts and not overdue_settlements:
                ttk.Label(alerts_frame, text="No critical alerts", foreground="green").pack(anchor="w")
                return
            if low_stock_alerts:
                ttk.Label(alerts_frame, text="Low Stock:", font=('Helvetica', 10, 'bold')).pack(anchor="w")
                for msg in low_stock_alerts:
                    ttk.Label(alerts_frame, text=f"• {msg}").pack(anchor="w")
            if overdue_settlements:
                ttk.Label(alerts_frame, text="Overdue Settlements:", font=('Helvetica', 10, 'bold')).pack(anchor="w", pady=(4, 0))
                for msg in overdue_settlements:
                    ttk.Label(alerts_frame, text=f"• {msg}").pack(anchor="w")
        except Exception:
            try:
                ttk.Label(alerts_frame, text="Unable to load alerts").pack(anchor="w")
            except Exception:
                pass

    def _import_settlements_csv(self, marketplace_manager, tree):
        try:
            platforms = marketplace_manager.list_platforms()
            if not platforms:
                messagebox.showerror("Error", "Please add a platform first")
                return
            file_path = filedialog.askopenfilename(
                title="Select platform settlement CSV",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            )
            if not file_path:
                return
            dlg = tk.Toplevel(self.root)
            dlg.title("Import Settlement CSV")
            dlg.geometry("420x200")
            dlg.transient(self.root)
            dlg.resizable(True, True)
            frame = ttk.Frame(dlg, padding="16")
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="Platform").grid(row=0, column=0, sticky="w", pady=4)
            platform_var = tk.StringVar()
            platform_names = [p.get("name", p.get("id")) for p in platforms]
            platform_combo = ttk.Combobox(
                frame, textvariable=platform_var, values=platform_names, state="readonly"
            )
            platform_combo.grid(row=0, column=1, sticky="ew", pady=4)
            if platform_names:
                platform_combo.current(0)
            ttk.Label(frame, text="CSV File").grid(row=1, column=0, sticky="w", pady=4)
            path_var = tk.StringVar(value=file_path)
            path_entry = ttk.Entry(frame, textvariable=path_var, state="readonly")
            path_entry.grid(row=1, column=1, sticky="ew", pady=4)
            frame.columnconfigure(1, weight=1)

            def do_import():
                try:
                    name = platform_var.get().strip()
                    if not name:
                        messagebox.showerror("Error", "Select a platform")
                        return
                    platform = None
                    for p in platforms:
                        if p.get("name") == name or p.get("id") == name:
                            platform = p
                            break
                    if not platform:
                        messagebox.showerror("Error", "Platform not found")
                        return
                    imported = 0
                    with open(path_var.get(), "r", encoding="utf-8-sig") as f:
                        reader = csv.DictReader(f)
                        for row in reader:
                            month = (
                                row.get("month")
                                or row.get("Month")
                                or row.get("settlement_month")
                                or row.get("SettlementMonth")
                            )
                            if not month:
                                continue
                            def num(name):
                                v = (
                                    row.get(name)
                                    or row.get(name.capitalize())
                                    or row.get(name.replace("_", " ").title())
                                )
                                if not v:
                                    return 0.0
                                try:
                                    return float(str(v).replace(",", ""))
                                except Exception:
                                    return 0.0
                            total_sales = num("total_sales")
                            fees = num("platform_fees")
                            refunds = num("refunds")
                            promotions = num("promotions")
                            payment_date = (
                                row.get("payment_date")
                                or row.get("PaymentDate")
                                or row.get("payment_date_str")
                            )
                            if not total_sales and not fees and not refunds and not promotions:
                                continue
                            marketplace_manager.record_monthly_settlement(
                                platform_id=platform.get("id"),
                                month=month,
                                total_sales=total_sales,
                                platform_fees=fees,
                                refunds=refunds,
                                promotions=promotions,
                                payment_date=payment_date,
                                payment_status="pending",
                            )
                            imported += 1
                    self._refresh_settlements_tree(marketplace_manager, tree)
                    dlg.destroy()
                    try:
                        messagebox.showinfo(
                            "Import complete",
                            f"Imported {imported} settlement rows for {name}",
                        )
                    except Exception:
                        pass
                except Exception as e:
                    messagebox.showerror("Error", str(e))

            btns = ttk.Frame(frame)
            btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=10)
            ttk.Button(btns, text="Import", command=do_import).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=4)
            _fit_window(dlg, 420, 200, mode="compact", remember_key="import_settlement_csv")
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _import_orders_csv(self, marketplace_manager, tree):
        try:
            platforms = marketplace_manager.list_platforms()
            if not platforms:
                messagebox.showerror("Error", "Please add a platform first")
                return
            file_path = filedialog.askopenfilename(
                title="Select platform orders CSV",
                filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            )
            if not file_path:
                return
            dlg = tk.Toplevel(self.root)
            dlg.title("Import Orders CSV")
            dlg.geometry("420x200")
            dlg.transient(self.root)
            dlg.resizable(True, True)
            frame = ttk.Frame(dlg, padding="16")
            frame.pack(fill="both", expand=True)
            ttk.Label(frame, text="Platform").grid(row=0, column=0, sticky="w", pady=4)
            platform_var = tk.StringVar()
            platform_names = [p.get("name", p.get("id")) for p in platforms]
            platform_combo = ttk.Combobox(
                frame, textvariable=platform_var, values=platform_names, state="readonly"
            )
            platform_combo.grid(row=0, column=1, sticky="ew", pady=4)
            if platform_names:
                platform_combo.current(0)
            ttk.Label(frame, text="CSV File").grid(row=1, column=0, sticky="w", pady=4)
            path_var = tk.StringVar(value=file_path)
            path_entry = ttk.Entry(frame, textvariable=path_var, state="readonly")
            path_entry.grid(row=1, column=1, sticky="ew", pady=4)
            frame.columnconfigure(1, weight=1)

            def do_import():
                try:
                    name = platform_var.get().strip()
                    if not name:
                        messagebox.showerror("Error", "Select a platform")
                        return
                    platform = None
                    for p in platforms:
                        if p.get("name") == name or p.get("id") == name:
                            platform = p
                            break
                    if not platform:
                        messagebox.showerror("Error", "Platform not found")
                        return
                    imported = 0

                    def num(name):
                        v = row.get(name) or row.get(name.capitalize()) or row.get(name.replace("_", " ").title())
                        if not v:
                            return 0.0
                        try:
                            return float(str(v).replace(",", ""))
                        except Exception:
                            return 0.0

                    with open(path_var.get(), "r", encoding="utf-8-sig") as f:
                        reader = csv.DictReader(f)
                        for row in reader:
                            order_id = (
                                row.get("order_id")
                                or row.get("OrderID")
                                or row.get("Order Id")
                                or row.get("AmazonOrderId")
                            )
                            if not order_id:
                                continue
                            order_date = (
                                row.get("order_date")
                                or row.get("OrderDate")
                                or row.get("purchase_date")
                                or row.get("PurchaseDate")
                            )
                            status = row.get("status") or row.get("Status") or "Delivered"
                            customer_name = (
                                row.get("customer")
                                or row.get("CustomerName")
                                or row.get("buyer_name")
                                or ""
                            )
                            total_amount = num("total_amount") or num("total") or num("grand_total")
                            fees = num("fees") or num("platform_fees")
                            refunds = num("refunds")
                            customer = {"name": customer_name}
                            items = []
                            marketplace_manager.record_order(
                                platform_id=platform.get("id"),
                                order_id=order_id,
                                order_date=order_date or datetime.now().strftime("%Y-%m-%d"),
                                status=status,
                                items=items,
                                customer=customer,
                                total_amount=total_amount,
                                fees=fees,
                                refunds=refunds,
                            )
                            imported += 1
                    self._refresh_orders_tree(marketplace_manager, tree)
                    dlg.destroy()
                    try:
                        messagebox.showinfo(
                            "Import complete",
                            f"Imported {imported} orders for {name}",
                        )
                    except Exception:
                        pass
                except Exception as e:
                    messagebox.showerror("Error", str(e))

            btns = ttk.Frame(frame)
            btns.grid(row=2, column=0, columnspan=2, sticky="e", pady=10)
            ttk.Button(btns, text="Import", command=do_import).pack(side="right", padx=4)
            ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=4)
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def _generate_settlement_invoice(self, marketplace_manager, tree):
        try:
            sel = tree.selection()
            if not sel:
                messagebox.showerror("Error", "Select a settlement first")
                return
            item = tree.item(sel[0])
            vals = item.get("values") or []
            if len(vals) < 2:
                return
            platform_name = vals[0]
            month = vals[1]
            platform = None
            for p in marketplace_manager.list_platforms():
                if p.get("name") == platform_name:
                    platform = p
                    break
            if not platform:
                messagebox.showerror("Error", "Platform not found")
                return
            settlements = marketplace_manager.list_settlements(
                platform_id=platform.get("id"), month=month
            )
            if not settlements:
                messagebox.showerror("Error", "Settlement details not found")
                return
            settlement = settlements[0]
            if settlement.get("invoice_generated") and settlement.get("invoice_number"):
                inv = marketplace_manager.ensure_settlement_invoice(self.manager, platform.get("id"), month)
                self._refresh_settlements_tree(marketplace_manager, tree)
                try:
                    if inv:
                        messagebox.showinfo("Info", f"Invoice already generated: {inv.invoice_id}")
                    else:
                        messagebox.showinfo("Info", f"Invoice already generated: {settlement.get('invoice_number')}")
                except Exception:
                    pass
                return
            try:
                invoice_folder = getattr(self.manager, "invoice_folder", None)
            except Exception:
                invoice_folder = None
            if not invoice_folder:
                messagebox.showerror("Error", "Invoice folder is not configured")
                return
            cons = marketplace_manager.list_consignments(platform_id=platform.get("id"))
            month_prefix = f"{month}-"
            product_rows = {}
            for c in cons:
                date_str = c.get("date", "")
                if not date_str.startswith(month_prefix):
                    continue
                for it in c.get("items", []):
                    pid = it.get("product_id")
                    name = it.get("item_name", pid)
                    qty = int(it.get("quantity", 0))
                    price = float(it.get("platform_price", 0) or 0)
                    key = pid or name
                    row = product_rows.setdefault(
                        key,
                        {
                            "product_id": pid,
                            "name": name,
                            "quantity": 0,
                            "unit_price": price,
                            "total": 0.0,
                        },
                    )
                    row["quantity"] += qty
                    row["total"] += qty * price
            total_sales = float(settlement.get("total_sales", 0) or 0)
            net_settlement = float(settlement.get("net_settlement", 0) or 0)
            fees = float(settlement.get("platform_fees", 0) or 0)
            refunds = float(settlement.get("refunds", 0) or 0)
            promotions = float(settlement.get("promotions", 0) or 0)
            inv = marketplace_manager.ensure_settlement_invoice(self.manager, platform.get("id"), month)
            self._refresh_settlements_tree(marketplace_manager, tree)
            try:
                if inv:
                    messagebox.showinfo("Success", f"Settlement invoice ready: {inv.invoice_id}")
                else:
                    messagebox.showerror("Error", "Failed to prepare settlement invoice")
            except Exception:
                pass
        except Exception as e:
            try:
                messagebox.showerror("Error", str(e))
            except Exception:
                pass

    def login_admin(self):
        try:
            dlg = EmployeePasswordDialog(self.root)
            self.root.wait_window(dlg.dialog)
            if dlg.result:
                self.user_role = 'Admin'
                try:
                    messagebox.showinfo("Success", "Successfully logged in as Admin")
                except Exception:
                    pass
                self.create_main_menu()
        except Exception:
            pass

    def logout_admin(self):
        self.user_role = 'Staff'
        self.create_main_menu()

    def show_ar_aging(self):
        try:
            dlg = CompanyReportDialog(self.root, self.manager)
            dlg.generate_ar_aging()
        except Exception:
            CompanyReportDialog(self.root, self.manager)

    def show_cash_flow(self):
        try:
            dlg = CompanyReportDialog(self.root, self.manager)
            dlg.generate_cash_flow_statement()
        except Exception:
            CompanyReportDialog(self.root, self.manager)

    def ensure_admin_then(self, action):
        try:
            if getattr(self, 'user_role', 'Staff') == 'Admin':
                action()
                return
            dlg = EmployeePasswordDialog(self.root)
            self.root.wait_window(dlg.dialog)
            if dlg.result:
                self.user_role = 'Admin'
                action()
        except Exception:
            pass

    def show_dashboard(self):
        for widget in self.root.winfo_children():
            widget.destroy()
        sf = ScrollableFrame(self.root)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        back_frame = ttk.Frame(container)
        back_frame.pack(fill='x', padx=10, pady=5)
        ttk.Button(back_frame, text="← Back to Main Menu", command=self.create_main_menu).pack(anchor='w')
        style = ttk.Style()
        try:
            if 'aqua' in style.theme_names():
                style.theme_use('aqua')
            else:
                style.theme_use('clam')
        except Exception:
            style.theme_use('clam')
        style.configure('.', background='#f8fafc')
        style.configure('TFrame', background='#f8fafc')
        style.configure('Card.TLabelframe', background='white', borderwidth=0, padding=16)
        style.configure('Card.TLabelframe.Label', font=('Helvetica', 12, 'bold'), foreground='#0f172a', background='white')
        style.configure('Treeview.Heading', font=('Helvetica', 11, 'bold'))
        style.configure('Treeview', font=('Helvetica', 10), rowheight=26, background='white', fieldbackground='white')
        style.map('Treeview', background=[('selected', '#e0e7ff')])
        
        main_frame = ttk.Frame(container, padding="16")
        main_frame.pack(fill='both', expand=True)
        header = ttk.Label(main_frame, text="Company Dashboard", font=('Helvetica', 20, 'bold'))
        header.pack()
        try:
            from datetime import datetime, timedelta, date
            end_date = datetime.now().strftime("%Y-%m-%d")
            start_date = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")
            report = self.manager.get_company_report(start_date, end_date)
        except Exception:
            report = {}
        stats = ttk.Frame(main_frame)
        stats.pack(fill='x')
        def fmt(v):
            try:
                return f"{float(v):,.2f}"
            except Exception:
                return str(v)
        kpis = [
            ("Total Revenue", fmt(report.get('total_revenue', 0))),
            ("Total Expenses", fmt(report.get('total_expenses', 0))),
            ("Net Profit", fmt(report.get('net_profit', 0))),
            ("Outstanding", fmt(report.get('outstanding_balance', 0))),
            ("Invoices", str(report.get('total_invoices', 0))),
            ("Purchases", str(report.get('total_purchases', 0)))
        ]
        for i, (label, value) in enumerate(kpis):
            card = ttk.LabelFrame(stats, text=label, padding="0")
            card.grid(row=0, column=i, padx=0, pady=0, sticky='nsew')
            ttk.Label(card, text=value, font=('Helvetica', 16, 'bold')).pack()
        detail_frame = ttk.Frame(main_frame)
        detail_frame.pack(fill='both', expand=True)
        left = ttk.LabelFrame(detail_frame, text="Top Clients", padding="0")
        left.pack(side='left', fill='both', expand=True)
        right = ttk.LabelFrame(detail_frame, text="Payment Status", padding="0")
        right.pack(side='right', fill='both', expand=True)
        top_clients = report.get('top_clients', [])
        tree = ttk.Treeview(left, columns=("client","revenue"), show='headings')
        tree.heading("client", text="Client")
        tree.heading("revenue", text="Revenue")
        tree.pack(fill='both', expand=True)
        for c, rev in top_clients:
            tree.insert('', 'end', values=(c, fmt(rev)))
        ps = report.get('payment_status_breakdown', {})
        for k in ["paid","partial","unpaid"]:
            ttk.Label(right, text=f"{k.capitalize()}: {ps.get(k,0)}", font=('Helvetica', 12)).pack(anchor='w')

    def show_clients(self):
        if not self._guard_module_access("clients", "Clients"):
            return
        for widget in self.root.winfo_children():
            widget.destroy()
        sf = ScrollableFrame(self.root)
        sf.pack(fill='both', expand=True)
        container = sf.scrollable_frame
        back_frame = ttk.Frame(container)
        back_frame.pack(fill='x', padx=10, pady=5)
        ttk.Button(back_frame, text="← Back to Main Menu", command=self.create_main_menu).pack(anchor='w')
        main_frame = ttk.Frame(container, padding="10")
        main_frame.pack(fill='both', expand=True)
        ttk.Label(main_frame, text="Clients", font=('Helvetica', 20, 'bold')).pack()
        controls = ttk.Frame(main_frame)
        controls.pack(fill='x')
        tk.Label(controls, text="Search:").pack(side='left')
        search_var = tk.StringVar()
        search_entry = ttk.Entry(controls, textvariable=search_var, width=30)
        search_entry.pack(side='left', padx=8)
        tree = ttk.Treeview(main_frame, columns=("client","invoices","sales","outstanding"), show='headings')
        for h in ("client","invoices","sales","outstanding"):
            tree.heading(h, text=h.capitalize())
        tree.pack(fill='both', expand=True)
        def load():
            tree.delete(*tree.get_children())
            try:
                invoices = self.manager.get_all_invoices_dict()
            except Exception:
                invoices = []
            data = {}
            for inv in invoices:
                c = inv.get('client_name','')
                if not c:
                    continue
                d = data.setdefault(c, {"invoices":0,"sales":0.0,"paid":0.0})
                d["invoices"] += 1
                d["sales"] += float(inv.get('grand_total',0) or 0)
                d["paid"] += float(inv.get('total_paid',0) or 0)
            q = search_var.get().lower()
            for c, d in sorted(data.items(), key=lambda x: x[1]["sales"], reverse=True):
                if q and q not in c.lower():
                    continue
                outstanding = max(d["sales"] - d["paid"], 0.0)
                tree.insert('', 'end', values=(c, d["invoices"], f"{d['sales']:,.2f}", f"{outstanding:,.2f}"))
        search_entry.bind('<KeyRelease>', lambda e: load())
        load()
        actions = ttk.Frame(main_frame)
        actions.pack(fill='x')
        def open_client_report():
            item = tree.selection()
            if not item:
                return
            client = tree.item(item[0], 'values')[0]
            try:
                from datetime import datetime, timedelta, date
                end_date = datetime.now().strftime("%Y-%m-%d")
                start_date = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")
                _ = self.manager.get_client_report(client, start_date, end_date)
                ReportsDialog(self.root, self.manager)
            except Exception:
                ReportsDialog(self.root, self.manager)
        ttk.Button(actions, text="Open Reports", command=open_client_report).pack(side='left')

    def log_event(self, action, details):
        try:
            folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
            path = os.path.join(folder, 'audit_log.jsonl')
            rec = {
                'timestamp': datetime.now().isoformat(),
                'action': action,
                'details': details
            }
            with open(path, 'a') as f:
                f.write(json.dumps(rec) + "\n")
        except Exception:
            pass

    def clear_activity_log(self):
        try:
            folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
            path = os.path.join(folder, 'audit_log.jsonl')
            with open(path, 'w', encoding='utf-8') as f:
                f.write("")
            return True
        except Exception:
            return False
    
    def setup_ui(self):
        """Original invoice management UI"""
        # Clear existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
        
        container = ttk.Frame(self.root)
        container.pack(fill='both', expand=True)
        
        # Back to main menu button
        back_frame = ttk.Frame(container)
        back_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(back_frame, text="← Back to Main Menu", 
                  command=self.create_main_menu).pack(anchor='w')
        
        main_frame = ttk.Frame(container, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Header with Logo and Title
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill='x')
        
        # Load and display logo
        if self.logo_manager.load_logo(size=(80, 80)):
            logo_label = ttk.Label(header_frame, image=self.logo_manager.logo_photo)
            logo_label.pack(side='left', padx=(0, 20))
        
        title_col = ttk.Frame(header_frame)
        title_col.pack(side='left', fill='y')
        ttk.Label(title_col, text="ASISTEM · Invoicing", 
                  font=('Helvetica', 20, 'bold'), foreground='#1A365D').pack(anchor='w')
        try:
            acp = self.manager.get_active_company_profile() if hasattr(self.manager, "get_active_company_profile") else None
            if acp:
                name = (acp.get("display_name") or acp.get("name") or "HopePharma").strip()
                prefix = acp.get("prefix") or "HPMT"
                ttk.Label(title_col, text=f"Working as: {name} · Invoice prefix: {prefix}####",
                          font=('Helvetica', 10), foreground='#4A5568').pack(anchor='w', pady=(2, 0))
        except Exception:
            ttk.Label(title_col, text="Working as: HopePharma Medical Trading · Invoice prefix: HPMT####",
                      font=('Helvetica', 10), foreground='#4A5568').pack(anchor='w', pady=(2, 0))
        
        # Search frame
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill='x')
        
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side='left', padx=(0, 10))
        self.search_entry.bind('<KeyRelease>', self.on_search)
        
        # Sort options
        ttk.Label(search_frame, text="Sort by:").pack(side='left', padx=(20, 5))
        self.sort_var = tk.StringVar(value="Date (Old-New)")
        sort_combo = ttk.Combobox(search_frame, textvariable=self.sort_var, 
                                 values=["Date (New-Old)", "Date (Old-New)", "Client (A-Z)", "Client (Z-A)", "Amount (High-Low)", "Amount (Low-High)", "ID (Asc)", "ID (Desc)"],
                                 state="readonly", width=15)
        sort_combo.pack(side='left', padx=(0, 10))
        sort_combo.bind('<<ComboboxSelected>>', self.on_sort_change)
        
        # Buttons frame
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill='x')
        
        ttk.Button(buttons_frame, text="➕ New Invoice", 
              command=self.create_invoice).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="👀 View Selected", 
              command=self.view_invoice).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="✏️ Edit Invoice", 
              command=self.edit_invoice).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="💰 Add Payment", 
              command=self.add_payment).pack(side='left', padx=5)  # ADDED PAYMENT BUTTON
        ttk.Button(buttons_frame, text="🆔 Change ID", 
              command=self.change_invoice_id).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="🗑️ Delete Invoice", 
              command=self.delete_invoice).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="🖨️ Print Invoice", 
              command=self.print_invoice).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="Sales and Services Reports",
              command=self.open_sales_services_reports).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="💳 Cost Centers",
              command=self._open_cost_centers_from_invoice_manager).pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="💰 Bulk Add Payments",
              command=self._open_bulk_add_payments, style="Primary.TButton").pack(side='left', padx=5)
        ttk.Button(buttons_frame, text="🔄 Refresh", 
              command=self.refresh_invoices).pack(side='left', padx=5)
        
        # Invoices list (two tabs: Active / Deleted Archive)
        list_frame = ttk.LabelFrame(main_frame, text="Invoices", padding="0", style='Card.TLabelframe')
        list_frame.pack(fill='both', expand=True)

        self.notebook = ttk.Notebook(list_frame)
        self.notebook.pack(fill='both', expand=True)

        # --- Tab 1: Active Invoices (scrollable tree, original layout) ---
        active_tab = ttk.Frame(self.notebook); self.notebook.add(active_tab, text="📋 Active Invoices")
        tree_frame = ttk.Frame(active_tab); tree_frame.pack(fill='both', expand=True)

        columns = ('ID', 'Client', 'Date', 'Due Date', 'Total', 'Paid', 'Balance', 'Status', 'Cost', 'P/L', 'Type', 'Process')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)
        
        # Define headings
        self.tree.heading('ID', text='Invoice ID')
        self.tree.heading('Client', text='Client Name')
        self.tree.heading('Date', text='Date')
        self.tree.heading('Due Date', text='Due Date')
        self.tree.heading('Total', text='Total Amount')
        self.tree.heading('Paid', text='Paid')
        self.tree.heading('Balance', text='Balance Due')
        self.tree.heading('Status', text='Status')
        self.tree.heading('Cost', text='Cost (AED)')
        self.tree.heading('P/L', text='P/L (AED)')
        self.tree.heading('Type', text='Type')
        self.tree.heading('Process', text='Process')
        
        # Define columns
        self.tree.column('ID', width=150)
        self.tree.column('Client', width=150)
        self.tree.column('Date', width=100)
        self.tree.column('Due Date', width=100)
        self.tree.column('Total', width=100)
        self.tree.column('Paid', width=100)
        self.tree.column('Balance', width=100)
        self.tree.column('Status', width=100)
        self.tree.column('Cost', width=100)
        self.tree.column('P/L', width=100)
        self.tree.column('Type', width=80)
        self.tree.column('Process', width=140)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Mouse wheel scroll for the tree only
        def _on_tree_mousewheel(event):
            try:
                self.tree.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
        self.tree.bind("<Enter>", lambda e: self.tree.bind_all("<MouseWheel>", _on_tree_mousewheel))
        self.tree.bind("<Leave>", lambda e: self.tree.unbind_all("<MouseWheel>"))
        
        # Double click to view
        self.tree.bind('<Double-1>', lambda e: self.view_invoice())
        # Click to toggle Process checkbox (admin only)
        self.tree.bind('<Button-1>', self.on_process_click)

        # --- Tab 2: Deleted Archive ---
        deleted_tab = ttk.Frame(self.notebook); self.notebook.add(deleted_tab, text="🗑️ Deleted Archive")

        # Button strip (top of tab)
        del_btn_frame = ttk.Frame(deleted_tab); del_btn_frame.pack(fill='x', padx=10, pady=8)
        ttk.Button(del_btn_frame, text="♻️ Restore Selected (to Active)",
                   command=self.restore_selected_invoice).pack(side='left', padx=4)
        ttk.Button(del_btn_frame, text="⚠️ Permanently Erase Selected (Unrecoverable)",
                   command=self.permanent_erase_selected_invoice).pack(side='left', padx=4)
        ttk.Label(del_btn_frame, text="(Use permanent erase only if legally required — GDPR/data retention. "
                                      "Soft-deleted invoices are hidden everywhere but kept for audit.)",
                  foreground='#6b7280', font=('Helvetica', 9)).pack(side='left', padx=10)

        del_tree_frame = ttk.Frame(deleted_tab); del_tree_frame.pack(fill='both', expand=True, padx=8, pady=(0, 10))
        del_cols = ('ID', 'Client', 'Date', 'Total', 'Paid', 'Balance', 'Deleted At', 'Deleted By')
        self.deleted_tree = ttk.Treeview(del_tree_frame, columns=del_cols, show='headings', height=14)
        for c, label, w in [
            ('ID', 'Invoice ID', 140), ('Client', 'Client Name', 220), ('Date', 'Invoice Date', 110),
            ('Total', 'Total Amount', 120), ('Paid', 'Paid', 120), ('Balance', 'Balance', 120),
            ('Deleted At', 'Deleted At', 170), ('Deleted By', 'Deleted By', 160),
        ]:
            self.deleted_tree.heading(c, text=label)
            self.deleted_tree.column(c, width=w)
        sb2 = ttk.Scrollbar(del_tree_frame, orient=tk.VERTICAL, command=self.deleted_tree.yview)
        self.deleted_tree.configure(yscrollcommand=sb2.set)
        self.deleted_tree.pack(side='left', fill='both', expand=True)
        sb2.pack(side='right', fill='y')
        self.deleted_tree.tag_configure('archived', background='#F7FAFC', foreground='#4A5568')

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var,
                              relief=tk.SUNKEN, font=('Helvetica', 10))
        status_bar.pack(fill='x', side='bottom')

        # Load initial data (both tabs)
        self.refresh_invoices()
        self.refresh_deleted_archive()

    def refresh_deleted_archive(self):
        for item in self.deleted_tree.get_children():
            self.deleted_tree.delete(item)
        deleted = self.manager.get_deleted_invoices_dict() if hasattr(self.manager, 'get_deleted_invoices_dict') else []
        for inv in deleted:
            total = inv.get('grand_total', 0); paid = inv.get('total_paid', 0)
            bal = total - paid
            currency = inv.get('currency', 'AED')
            self.deleted_tree.insert('', tk.END, iid=str(inv.get('invoice_id', '')),
                                     values=(
                                         inv.get('invoice_id', ''),
                                         inv.get('client_name', ''),
                                         inv.get('date', ''),
                                         f"{currency} {total:.2f}",
                                         f"{currency} {paid:.2f}",
                                         f"{currency} {bal:.2f}",
                                         inv.get('deleted_at', ''),
                                         inv.get('deleted_by', ''),
                                     ), tags=('archived',))
        if self.notebook and self.notebook.tabs():
            try:
                n = self.notebook.index(self.notebook.select())
            except Exception:
                n = 0
            if n == 1:
                self.status_var.set(f"Deleted Archive: {len(deleted)} invoices")
            else:
                self.status_var.set(f"Loaded {len(self.tree.get_children())} active invoices · "
                                    f"{len(deleted)} in archive")

    def get_selected_deleted_id(self):
        sel = self.deleted_tree.selection()
        if not sel:
            return None
        return str(sel[0])

    def restore_selected_invoice(self):
        iid = self.get_selected_deleted_id()
        if not iid:
            messagebox.showinfo("Nothing selected", "Select an invoice from the Deleted Archive first.")
            return
        if not messagebox.askyesno("Restore?",
                                   f"Restore invoice {iid} to the Active Invoices list?\n\n"
                                   f"Note: If its HPMT number has already been re-assigned since deletion, "
                                   f"you may want to use Change ID afterwards to avoid duplicate display."):
            return
        if hasattr(self.manager, 'restore_invoice'):
            ok, msg = self.manager.restore_invoice(iid)
        else:
            ok, msg = False, "Manager has no restore_invoice method"
        if ok:
            messagebox.showinfo("Restored", msg)
        else:
            messagebox.showerror("Restore failed", msg)
        self.refresh_invoices()
        self.refresh_deleted_archive()

    def permanent_erase_selected_invoice(self):
        iid = self.get_selected_deleted_id()
        if not iid:
            messagebox.showinfo("Nothing selected", "Select an invoice from the Deleted Archive first.")
            return
        if not messagebox.askyesno("⚠️ PERMANENTLY ERASE?",
                                   f"Invoice {iid} will be physically DELETED from disk — this is UNRECOVERABLE.\n\n"
                                   f"Action cannot be undone. Continue?"):
            return
        if hasattr(self.manager, 'permanent_erase_invoice'):
            ok = self.manager.permanent_erase_invoice(iid)
        else:
            ok = self.manager.delete_invoice(iid)
        if ok:
            messagebox.showinfo("Erased", f"Invoice {iid} permanently erased.")
        else:
            messagebox.showerror("Erase failed", f"Could not erase invoice {iid}.")
        self.refresh_deleted_archive()
    
    def refresh_invoices(self):
        """Refresh the invoices list"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Use the get_all_invoices_dict method for compatibility
        invoices = self.manager.get_all_invoices_dict()
        invoices = self.sort_invoices(invoices, self.sort_var.get())
        
        for invoice in invoices:
            total = invoice.get('grand_total', 0)
            paid = invoice.get('total_paid', 0)
            balance = total - paid
            cost = invoice.get('total_cost', 0)
            pl = invoice.get('profit_loss', 0)
            invoice_type = invoice.get('invoice_type', 'sales')
            currency = invoice.get('currency', 'AED')
            
            # Determine status with emojis
            if paid >= total:
                status = "✅ Paid"
            elif paid > 0:
                status = "🟡 Partial Paid"
            else:
                status = "🔴 Not Paid"
                
            # Type indicator
            type_display = "🛒" if invoice_type == 'sales' else "🔧"
                
            process = invoice.get('workflow_status', 'Under Process')
            process_display = '⬜ Undergoing' if str(process).lower().startswith('under') else '🟩 Closed'
            self.tree.insert('', tk.END, values=(
                invoice.get('invoice_id', ''),
                invoice.get('client_name', ''),
                invoice.get('date', ''),
                invoice.get('due_date', ''),
                f"{currency} {total:.2f}",
                f"{currency} {paid:.2f}",
                f"{currency} {balance:.2f}",
                status,
                f"{currency} {cost:.2f}",
                f"{currency} {pl:.2f}",
                type_display,
                process_display
            ))
        
        self.status_var.set(f"Loaded {len(invoices)} invoices")
    
    def sort_invoices(self, invoices, sort_option):
        """Sort invoices based on selected option"""
        if sort_option == "Date (New-Old)":
            return sorted(invoices, key=lambda x: x.get('date', ''), reverse=True)
        elif sort_option == "Date (Old-New)":
            return sorted(invoices, key=lambda x: x.get('date', ''))
        elif sort_option == "Client (A-Z)":
            return sorted(invoices, key=lambda x: x.get('client_name', '').lower())
        elif sort_option == "Client (Z-A)":
            return sorted(invoices, key=lambda x: x.get('client_name', '').lower(), reverse=True)
        elif sort_option == "Amount (High-Low)":
            return sorted(invoices, key=lambda x: x.get('grand_total', 0), reverse=True)
        elif sort_option == "Amount (Low-High)":
            return sorted(invoices, key=lambda x: x.get('grand_total', 0))
        elif sort_option == "ID (Asc)":
            return sorted(invoices, key=lambda x: x.get('invoice_id', ''))
        elif sort_option == "ID (Desc)":
            return sorted(invoices, key=lambda x: x.get('invoice_id', ''), reverse=True)
        else:
            return invoices
    
    def on_sort_change(self, event=None):
        """Handle sort change"""
        self.refresh_invoices()
    
    def on_search(self, event=None):
        """Handle search functionality"""
        query = self.search_var.get().strip()
        if not query:
            self.refresh_invoices()
            return
            
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        results = self.manager.search_invoices(query)
        results = self.sort_invoices([inv.to_dict() for inv in results], self.sort_var.get())
        
        for invoice in results:
            total = invoice.get('grand_total', 0)
            paid = invoice.get('total_paid', 0)
            balance = total - paid
            cost = invoice.get('total_cost', 0)
            pl = invoice.get('profit_loss', 0)
            invoice_type = invoice.get('invoice_type', 'sales')
            currency = invoice.get('currency', 'AED')
            
            if paid >= total:
                status = "✅ Paid"
            elif paid > 0:
                status = "🟡 Partial Paid"
            else:
                status = "🔴 Not Paid"
                
            type_display = "🛒" if invoice_type == 'sales' else "🔧"
                
            process = invoice.get('workflow_status', 'Under Process')
            process_display = '⬜ Undergoing' if str(process).lower().startswith('under') else '🟩 Closed'
            self.tree.insert('', tk.END, values=(
                invoice.get('invoice_id', ''),
                invoice.get('client_name', ''),
                invoice.get('date', ''),
                invoice.get('due_date', ''),
                f"{currency} {total:.2f}",
                f"{currency} {paid:.2f}",
                f"{currency} {balance:.2f}",
                status,
                f"{currency} {cost:.2f}",
                f"{currency} {pl:.2f}",
                type_display,
                process_display
            ))
        
        self.status_var.set(f"Found {len(results)} invoices matching '{query}'")
    
    def get_selected_invoice(self):
        """Get currently selected invoice"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an invoice first")
            return None, None
            
        item = self.tree.item(selection[0])
        invoice_id = item['values'][0]
        
        # Get the invoice as dictionary for GUI compatibility
        invoice_obj = self.manager.get_invoice(invoice_id)
        if invoice_obj:
            return invoice_id, invoice_obj.to_dict()
        return None, None
    
    def on_process_click(self, event):
        try:
            # Identify column and row
            col_id = self.tree.identify_column(event.x)
            row_id = self.tree.identify_row(event.y)
            if not row_id:
                return
            columns = list(self.tree['columns'])
            if 'Process' not in columns:
                return
            target_col = f"#{columns.index('Process') + 1}"
            if col_id != target_col:
                return
            # Enforce admin role
            if getattr(self, 'user_role', 'Staff') != 'Admin':
                messagebox.showwarning("Permission", "Only admin can change process status")
                return
            # Toggle checkbox visually and persist
            values = list(self.tree.item(row_id)['values'])
            invoice_id = values[0]
            current = str(values[columns.index('Process')])
            checked = current.startswith('🟩')
            new_display = '⬜ Undergoing' if checked else '🟩 Closed'
            values[columns.index('Process')] = new_display
            self.tree.item(row_id, values=values)
            inv_obj = self.manager.get_invoice(invoice_id)
            if inv_obj:
                inv_obj.workflow_status = 'Under Process' if new_display.startswith('⬜') else 'Closed'
                if hasattr(self.manager, 'update_invoice'):
                    ok = self.manager.update_invoice(inv_obj)
                else:
                    ok = self.manager.add_invoice_from_dict(inv_obj.to_dict())
                if not ok:
                    messagebox.showerror("Error", "Failed to save process change")
        except Exception:
            pass
    
    def create_invoice(self):
        """Open create invoice dialog"""
        dialog = CreateInvoiceDialog(self.root, self.manager, self.data_memory)
        self.root.wait_window(dialog.dialog)
        self.refresh_invoices()
    
    def view_invoice(self):
        """View selected invoice"""
        invoice_id, invoice = self.get_selected_invoice()
        if invoice:
            ViewInvoiceDialog(self.root, self.manager, invoice_id, invoice)

    def open_sales_services_reports(self):
        """Open the grouped sales and services reports dialog"""
        SalesServicesReportsDialog(self.root, self.manager)

    def _open_cost_centers_from_invoice_manager(self):
        """Open Cost Centers (Manager + Bulk Backfill) from the Professional Invoice Manager toolbar.

        User explicitly requested Cost Centers live INSIDE this Professional Invoice Manager window
        (the window they have open in the screenshot), NOT inside Create/Edit Invoice.
        From here user can:
          1) Create a Cost Center (e.g. called 'costs')
          2) EXTRACT all service/item descriptions from every past invoice automatically
             (e.g. 'DM Registration and Classification', 'Classification & DM Fees',
                  'MOH Registration Service', 'Hearing Test Consultation', etc.)
          3) For each extracted service user types a cost (e.g. MOH Registration = 500 AED)
             OR just types new future services manually + cost
          4) One-click → all selected services become rules in the Cost Center
          5) Bulk Backfill → scan every past invoice in the list above.  If an invoice's line-items
             contain the service description → the corresponding cost is AUTOMATICALLY added
             to that invoice (all past invoices, all current ones, no errors).
        """
        try:
            from cost_center_ui import open_cost_center_and_backfill as _runner
            folder = (
                getattr(self.manager, "invoice_folder", None)
                or getattr(self.manager, "data_folder", None)
                or os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")
            )
            # Pass the invoice_manager instance (for extract-services + backfill)
            # AND the professional-invoice-manager reference so it can auto-refresh
            # the invoices list after bulk apply.
            try:
                from cost_center_ui import CostCenterManagerDialog, BulkBackfillDialog
            except Exception:
                pass
            try:
                _runner(self.root, self.manager, folder,
                        post_apply_refresh_cb=self.refresh_invoices,
                        get_all_invoices_dict_cb=getattr(self.manager, "get_all_invoices_dict", None))
            except TypeError:
                # Older signature without callbacks
                _runner(self.root, self.manager, folder)
            # Refresh the list anyway
            try:
                self.refresh_invoices()
            except Exception:
                pass
        except Exception as exc:
            import traceback
            traceback.print_exc()
            messagebox.showerror(
                "Cost Centers Error",
                f"Could not open Cost Centers from Invoice Manager:\n{exc}"
            )
    
    def edit_invoice(self):
        """Edit selected invoice"""
        invoice_id, invoice = self.get_selected_invoice()
        if invoice:
            dialog = EditInvoiceDialog(self.root, self.manager, invoice_id, invoice, self.data_memory)
            self.root.wait_window(dialog.dialog)
            self.refresh_invoices()
    
    def add_payment(self):
        """Add payment to selected invoice"""
        invoice_id, invoice = self.get_selected_invoice()
        if invoice:
            dialog = AddPaymentDialog(self.root, self.manager, invoice_id, invoice)
            self.root.wait_window(dialog.dialog)
            self.refresh_invoices()

    def _open_bulk_add_payments(self):
        """Open Bulk Add Payments dialog — record ONE payment for MANY invoices at once."""
        try:
            dialog = BulkPaymentsDialog(self.root, self.manager)
            self.root.wait_window(dialog.dlg)
        except Exception as exc:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Bulk Add Payments Error", str(exc))
        finally:
            try:
                self.refresh_invoices()
            except Exception:
                pass

    def change_invoice_id(self):
        """Change invoice ID"""
        invoice_id, invoice = self.get_selected_invoice()
        if invoice:
            dialog = ChangeInvoiceIdDialog(self.root, self.manager, invoice_id, invoice)
            self.root.wait_window(dialog.dialog)
            self.refresh_invoices()
    
    def delete_invoice(self):
        """Soft-delete selected invoice (move to archive) + reverse payments so number is free for re-use."""
        invoice_id, invoice = self.get_selected_invoice()
        if invoice:
            if not messagebox.askyesno("Confirm Archive (Soft Delete)",
                                 f"Archive invoice {invoice_id}?\n"
                                 f"Client: {invoice.get('client_name', '')}\n"
                                 f"Amount: AED {invoice.get('grand_total', 0):.2f}\n\n"
                                 f"✅ Payments will be reversed (balances adjusted).\n"
                                 f"✅ Invoice moves to 🗑️ Deleted Archive (never permanently erased unless you manually Erase).\n"
                                 f"✅ The HPMT number '{invoice_id}' will immediately become available for the NEXT new invoice."):
                return

            # Prefer soft-delete with balance adjustment. Fallback to old hard-delete if method missing.
            if hasattr(self.manager, 'soft_delete_invoice_with_balance_adjustment'):
                success, message = self.manager.soft_delete_invoice_with_balance_adjustment(
                    invoice_id, username=getattr(self, 'username', None) or "AppUser"
                )
            elif hasattr(self.manager, 'delete_invoice_with_balance_adjustment'):
                success, message = self.manager.delete_invoice_with_balance_adjustment(invoice_id)
            else:
                success = self.manager.delete_invoice(invoice_id)
                message = "Invoice deleted (legacy hard-delete; balance not adjusted)"

            if success:
                messagebox.showinfo("Success", message)
                try:
                    folder = getattr(self.manager, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                    path = os.path.join(folder, 'audit_log.jsonl')
                    rec = {
                        'timestamp': datetime.now().isoformat(),
                        'action': 'invoice_soft_deleted',
                        'details': {
                            'invoice_id': invoice_id,
                            'grand_total': invoice.get('grand_total', 0),
                            'total_paid': invoice.get('total_paid', 0)
                        }
                    }
                    with open(path, 'a') as f:
                        f.write(json.dumps(rec) + "\n")
                except Exception:
                    pass
                self.refresh_invoices()
                try:
                    self.refresh_deleted_archive()
                except Exception:
                    pass
            else:
                messagebox.showerror("Error", message)
    
    def print_invoice(self):
        """Print selected invoice"""
        if not PDF_AVAILABLE:
            messagebox.showwarning("Warning", "PDF generation is not available. Exporting as HTML instead.")
            
        invoice_id, invoice = self.get_selected_invoice()
        if invoice:
            self.export_invoice_pdf(invoice_id, invoice)
    
    def export_invoice_pdf(self, invoice_id, invoice):
        """Export invoice as PDF for printing"""
        try:
            # Use PDF generator if available
            if PDF_AVAILABLE:
                # Pass the invoices folder explicitly to ensure correct location
                invoices_folder = None
                if hasattr(self.manager, 'data_folder'):
                    invoices_folder = os.path.join(self.manager.data_folder, "HopePharmaInvoices")
                elif hasattr(self.manager, 'invoice_folder'):
                    invoices_folder = os.path.join(self.manager.invoice_folder, "HopePharmaInvoices")
                
                if invoices_folder:
                    pdf_gen = PDFGenerator(self.logo_manager, output_folder=invoices_folder)
                else:
                    pdf_gen = PDFGenerator(self.logo_manager)
                    
                pdf_path = pdf_gen.generate_invoice_pdf(invoice)
                if pdf_path:
                    messagebox.showinfo("Success", f"Invoice exported to:\n{pdf_path}")
                    return
            # Fallback to HTML
            self.export_invoice_html(invoice_id, invoice)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export invoice: {e}")
            # Fallback to HTML
            self.export_invoice_html(invoice_id, invoice)
    
    def export_invoice_html(self, invoice_id, invoice):
        """Export invoice as HTML for printing - UPDATED SIGNATURE LAYOUT"""
        try:
            # Determine invoices folder
            invoices_folder = None
            if hasattr(self.manager, 'data_folder'):
                invoices_folder = os.path.join(self.manager.data_folder, "HopePharmaInvoices")
            elif hasattr(self.manager, 'invoice_folder'):
                invoices_folder = os.path.join(self.manager.invoice_folder, "HopePharmaInvoices")
            
            # Fallback if folder doesn't exist
            if not invoices_folder or not os.path.exists(invoices_folder):
                if hasattr(self.manager, 'invoice_folder'):
                    invoices_folder = self.manager.invoice_folder
                else:
                    invoices_folder = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")

            # Save HTML invoice to HopePharmaInvoices folder
            filename = os.path.join(invoices_folder, f"{invoice_id}.html")
            
            # Get signature images - check HopePharmaInvoices first, then main folder
            sign1_path = os.path.join(invoices_folder, "sign1.png")
            sign2_path = os.path.join(invoices_folder, "sign2.png")
            
            # Fallback to main data folder for signatures
            if not os.path.exists(sign1_path) and hasattr(self.manager, 'invoice_folder'):
                sign1_fallback = os.path.join(self.manager.invoice_folder, "sign1.png")
                if os.path.exists(sign1_fallback):
                    sign1_path = sign1_fallback
                    
            if not os.path.exists(sign2_path) and hasattr(self.manager, 'invoice_folder'):
                sign2_fallback = os.path.join(self.manager.invoice_folder, "sign2.png")
                if os.path.exists(sign2_fallback):
                    sign2_path = sign2_fallback
            
            signatures_html = ""
            has_sign1 = os.path.exists(sign1_path)
            has_sign2 = os.path.exists(sign2_path)
            
            if has_sign1 or has_sign2:
                signatures_html = '''
                <div style="display: flex; align-items: flex-start; margin: 30px 0 20px 0; width: 100%;">
                '''
                
                # Left signatures - sticking together with minimal space
                if has_sign1 or has_sign2:
                    signatures_html += '<div style="display: flex; align-items: center; gap: 5px;">'
                    
                    # Signature 1 - larger circular stamp (same size as WhatsApp image)
                    if has_sign1:
                        try:
                            from io import BytesIO
                            import base64
                            from PIL import Image as PILImage
                            
                            sign1_img = PILImage.open(sign1_path)
                            # Create larger circular signature (120px to match WhatsApp image)
                            size = 120  # Fixed size to match WhatsApp image
                            # Create a circular mask
                            mask = PILImage.new('L', (size, size), 0)
                            # Create a white circle
                            for i in range(size):
                                for j in range(size):
                                    if (i - size/2)**2 + (j - size/2)**2 <= (size/2)**2:
                                        mask.putpixel((i, j), 255)
                            
                            # Resize image to square for circle
                            sign1_img = sign1_img.resize((size, size), PILImage.Resampling.LANCZOS)
                            # Apply circular mask
                            sign1_img.putalpha(mask)
                            
                            buffered = BytesIO()
                            sign1_img.save(buffered, format="PNG")
                            sign1_str = base64.b64encode(buffered.getvalue()).decode()
                            signatures_html += f'''
                            <div style="text-align: center;">
                                <img src="data:image/png;base64,{sign1_str}" style="height: 120px; width: 120px; border-radius: 50%; object-fit: cover;">
                            </div>
                            '''
                        except Exception as e:
                            print(f"Error processing sign1: {e}")
                            signatures_html += '<div style="width: 120px;"></div>'
                    
                    # Minimal space between signatures
                    signatures_html += '<div style="width: 10px;"></div>'
                    
                    # Signature 2 - rectangular with "Signature" text below
                    if has_sign2:
                        try:
                            from io import BytesIO
                            import base64
                            from PIL import Image as PILImage
                            
                            sign2_img = PILImage.open(sign2_path)
                            # Resize signature maintaining aspect ratio (same height as sign1)
                            original_width, original_height = sign2_img.size
                            new_height = 120  # Same height as sign1
                            new_width = int((original_width / original_height) * new_height)
                            sign2_img = sign2_img.resize((new_width, new_height), PILImage.Resampling.LANCZOS)
                            
                            buffered = BytesIO()
                            sign2_img.save(buffered, format="PNG")
                            sign2_str = base64.b64encode(buffered.getvalue()).decode()
                            signatures_html += f'''
                            <div style="text-align: center;">
                                <img src="data:image/png;base64,{sign2_str}" style="height: 120px; width: auto; max-width: 150px;">
                                <p style="margin-top: 2px; font-size: 10px;">Signature</p>
                            </div>
                            '''
                        except Exception as e:
                            print(f"Error processing sign2: {e}")
                            signatures_html += '<div style="width: 150px;"></div>'
                    
                    signatures_html += '</div>'
                
                signatures_html += '</div>'
            
            # Get logo for HTML
            logo_html = ""
            logo_base64 = self.logo_manager.get_logo_for_html(size=(100, 100))
            if logo_base64:
                logo_html = f'<div style="text-align: center; margin-bottom: 15px;"><img src="{logo_base64}" alt="Hope Pharma Logo" style="height: 70px;"></div>'
            
            # Get client-specific data
            client_name = invoice.get('client_name', '')
            client_trn = invoice.get('client_trn', 'N/A')
            client_location = (invoice.get('client_location') or '').strip()
            if not client_location:
                client_location = str(invoice.get('client_emirate', '') or '').strip()
            if not client_location:
                client_location = 'Dubai, United Arab Emirates'
            uae_emirates = {
                'Dubai',
                'Abu Dhabi',
                'Sharjah',
                'Ajman',
                'Umm Al Quwain',
                'Ras Al Khaimah',
                'Fujairah',
            }
            if ',' not in client_location and client_location in uae_emirates:
                client_location = f"{client_location}, United Arab Emirates"
            
            # Calculate totals for the right side
            subtotal = invoice.get('subtotal', 0)
            taxable_amount = invoice.get('taxable_amount', 0)
            non_taxable_amount = invoice.get('non_taxable_amount', 0)
            tax_amount = invoice.get('tax_amount', 0)
            tax_rate = invoice.get('tax_rate', 5)
            grand_total = invoice.get('grand_total', 0)
            
            # Create HTML content with letter style
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>TAX INVOICE - {invoice_id}</title>
                <style>
                    @page {{
                        size: A4;
                        margin: 0.6in;
                    }}
                    body {{ 
                        font-family: 'Arial', sans-serif; 
                        margin: 0;
                        padding: 0;
                        font-size: 12px;
                        line-height: 1.3;
                        background-color: #ffffff;
                    }}
                    .container {{
                        width: 100%;
                        max-width: 21cm;
                        margin: 0 auto;
                        padding: 20px;
                    }}
                    .letter-header {{
                        border-bottom: 2px solid #2c5aa0;
                        padding-bottom: 15px;
                        margin-bottom: 20px;
                    }}
                    .company-info {{
                        text-align: center;
                        margin-bottom: 5px;
                        line-height: 1.4;
                    }}
                    .company-name {{
                        font-size: 16px;
                        font-weight: bold;
                        color: #2c5aa0;
                        margin-bottom: 5px;
                    }}
                    .invoice-title {{
                        font-size: 18px;
                        font-weight: bold;
                        text-align: center;
                        margin: 15px 0;
                        color: #333;
                    }}
                    .content-section {{
                        margin-bottom: 20px;
                    }}
                    .section-title {{
                        font-weight: bold;
                        margin-bottom: 8px;
                        color: #2c5aa0;
                        font-size: 13px;
                    }}
                    table {{
                        width: 100%;
                        border-collapse: collapse;
                        margin: 10px 0;
                        font-size: 11px;
                    }}
                    th, td {{
                        border: 1px solid #ddd;
                        padding: 8px;
                        text-align: left;
                    }}
                    th {{
                        background-color: #f8f9fa;
                        font-weight: bold;
                        color: #333;
                    }}
                    .totals-section {{
                        display: flex;
                        justify-content: space-between;
                        margin-top: 20px;
                    }}
                    .signatures {{
                        flex: 1;
                    }}
                    .amounts {{
                        flex: 1;
                        text-align: right;
                    }}
                    .amount-line {{
                        margin-bottom: 5px;
                        display: flex;
                        justify-content: space-between;
                    }}
                    .amount-label {{
                        font-weight: bold;
                        margin-right: 10px;
                    }}
                    .amount-value {{
                        text-align: right;
                        min-width: 100px;
                    }}
                    .grand-total {{
                        font-weight: bold;
                        font-size: 13px;
                        border-top: 2px solid #333;
                        padding-top: 5px;
                        margin-top: 5px;
                    }}
                    .terms-section {{
                        margin-top: 25px;
                        padding: 15px;
                        background-color: #f8f9fa;
                        border-radius: 5px;
                        font-size: 11px;
                    }}
                    .terms-title {{
                        font-weight: bold;
                        margin-bottom: 8px;
                        color: #2c5aa0;
                    }}
                    .footer {{
                        margin-top: 30px;
                        text-align: center;
                        color: #666;
                        font-size: 10px;
                        border-top: 1px solid #ddd;
                        padding-top: 10px;
                    }}
                    .separator {{
                        border-top: 1px solid #ddd;
                        margin: 15px 0;
                    }}
                </style>
            </head>
            <body>
                <div class="container">
                    <!-- Letter Header -->
                    <div class="letter-header">
                        {logo_html}
                        <div class="company-info">
                            <div class="company-name">HOPE PHARMA MEDICINE TRADING</div>
                            <div>Dubai, International City, Morocco Cluster, Bldg. J-16, Store 08</div>
                            <div>TRN: 100466797600003</div>
                        </div>
                    </div>
                    
                    <!-- Invoice Title -->
                    <div class="invoice-title">TAX INVOICE</div>
                    
                    <!-- Client Information -->
                    <div class="content-section">
                        <div class="section-title">BILL TO</div>
                        <div>
                            <strong>{client_name}</strong><br>
                            {client_location}<br>
                            TRN: {client_trn}<br>
                            Invoice Date: {invoice.get('date', 'N/A')}<br>
                            Invoice ID: {invoice_id}
                        </div>
                    </div>
                    
                    <div class="separator"></div>
                    
                    <!-- Items Table -->
                    <div class="content-section">
                        <div class="section-title">INVOICE ITEMS</div>
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 8%">QTY</th>
                                    <th style="width: 52%">DESCRIPTION</th>
                                    <th style="width: 10%">VAT</th>
                                    <th style="width: 15%">UNIT PRICE (AED)</th>
                                    <th style="width: 15%">AMOUNT (AED)</th>
                                </tr>
                            </thead>
                            <tbody>
            """
            
            # Add items with VAT column
            for item in invoice.get('items', []):
                vat_status = "Yes" if item.get('taxable', True) else "No"
                html_content += f"""
                                <tr>
                                    <td>{int(item.get('quantity', 1))}</td>
                                    <td>{item.get('description', '')}</td>
                                    <td style="text-align: center;">{vat_status}</td>
                                    <td style="text-align: right;">{item.get('unit_price', 0):,.2f}</td>
                                    <td style="text-align: right;">{item.get('total', 0):,.2f}</td>
                                </tr>
                """
            
            html_content += f"""
                            </tbody>
                        </table>
                    </div>
                    
                    <!-- Totals and Signatures Section -->
                    <div class="totals-section">
                        <div class="signatures">
                            {signatures_html}
                        </div>
                        
                        <div class="amounts">
                            <div class="amount-line">
                                <span class="amount-label">Subtotal:</span>
                                <span class="amount-value">AED {subtotal:,.2f}</span>
                            </div>
                            <div class="amount-line">
                                <span class="amount-label">Taxable Amount:</span>
                                <span class="amount-value">AED {taxable_amount:,.2f}</span>
                            </div>
                            <div class="amount-line">
                                <span class="amount-label">Non-Taxable Amount:</span>
                                <span class="amount-value">AED {non_taxable_amount:,.2f}</span>
                            </div>
                            <div class="amount-line">
                                <span class="amount-label">VAT ({tax_rate}%):</span>
                                <span class="amount-value">AED {tax_amount:,.2f}</span>
                            </div>
                            <div class="amount-line grand-total">
                                <span class="amount-label">GRAND TOTAL:</span>
                                <span class="amount-value">AED {grand_total:,.2f}</span>
                            </div>
                        </div>
                    </div>
                    
                    <!-- Terms and Conditions -->
                    <div class="terms-section">
                        <div class="terms-title">TERMS & CONDITIONS</div>
                        <div>
                            All payments should be in favour of Hope Pharma Medicine Trading<br>
                            Payment Terms: Cash<br>
                            Bank: ADCB | Branch: 259 Dubai Mall, Dubai<br>
                            Account No: 14198059920001<br>
                            IBAN: AE980030014198059920001 | Swift Code: ADCBAEAAXXX
                        </div>
                    </div>
                    
                    <!-- Footer -->
                    <div class="footer">
                        <p>Thank you for your business!</p>
                        <p>Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}</p>
                    </div>
                </div>
            </body>
            </html>
            """
            
            with open(filename, 'w') as f:
                f.write(html_content)
                
            messagebox.showinfo("Success", f"Invoice exported to:\n{filename}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export invoice: {e}")

def main():
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()

def _find_invoice_app(win):
    try:
        root = win.winfo_toplevel()
        return getattr(root, "_invoice_app", None)
    except Exception:
        return None


def _parse_geometry(geometry_text):
    try:
        text = str(geometry_text or "").strip()
        if not text or "x" not in text:
            return None
        size_part, _, pos_part = text.partition("+")
        width_text, _, height_text = size_part.partition("x")
        width = int(float(width_text))
        height = int(float(height_text))
        x = 0
        y = 0
        if pos_part:
            pos_parts = text.split("+")
            if len(pos_parts) >= 3:
                x = int(float(pos_parts[1]))
                y = int(float(pos_parts[2]))
        return width, height, x, y
    except Exception:
        return None


def _screen_metrics(win):
    try:
        scaling = float(win.tk.call("tk", "scaling"))
    except Exception:
        scaling = 1.0
    scaling = max(1.0, min(scaling, 3.0))
    return max(800, int(win.winfo_screenwidth())), max(600, int(win.winfo_screenheight())), scaling


def _clamp_geometry(parsed, screen_w, screen_h, min_w, min_h, max_w, max_h):
    if not parsed:
        return None
    width, height, x, y = parsed
    width = max(min_w, min(width, max_w))
    height = max(min_h, min(height, max_h))
    x = min(max(0, x), max(0, screen_w - width))
    y = min(max(0, y), max(0, screen_h - height))
    return width, height, x, y


def _bind_window_persistence(win, remember_key):
    key = str(remember_key or "").strip()
    if not key:
        return
    try:
        if getattr(win, "_remember_bound_key", None) == key:
            return
    except Exception:
        pass

    def _schedule_store(_event=None):
        app = _find_invoice_app(win)
        if app:
            app._schedule_remember_window(key, win)

    try:
        win._remember_bound_key = key
        win.bind("<Configure>", _schedule_store, add="+")
    except Exception:
        pass


def _fit_window(
    win,
    min_w=640,
    min_h=480,
    pref_w=None,
    pref_h=None,
    mode="dialog",
    remember_key=None,
    restore=True,
):
    try:
        win.update_idletasks()
        screen_w, screen_h, scaling = _screen_metrics(win)
        pad = int(24 * scaling)
        req_w = max(min_w, win.winfo_reqwidth() + pad)
        req_h = max(min_h, win.winfo_reqheight() + pad)

        max_w_ratio = 0.98 if mode == "workspace" else 0.94
        max_h_ratio = 0.95 if mode == "workspace" else 0.92
        max_w = max(min_w, int(screen_w * max_w_ratio))
        max_h = max(min_h, int(screen_h * max_h_ratio))

        target_w = max(min_w, req_w, int(pref_w or 0))
        target_h = max(min_h, req_h, int(pref_h or 0))

        if mode == "workspace":
            target_w = max(target_w, int(screen_w * 0.90))
            target_h = max(target_h, int(screen_h * 0.86))
        elif mode == "large":
            target_w = max(target_w, int(screen_w * 0.76))
            target_h = max(target_h, int(screen_h * 0.72))
        elif mode == "compact":
            target_w = min(max(target_w, min_w), int(screen_w * 0.55))
            target_h = min(max(target_h, min_h), int(screen_h * 0.55))

        target_w = max(min_w, min(target_w, max_w))
        target_h = max(min_h, min(target_h, max_h))

        app = _find_invoice_app(win)
        saved = None
        if restore and app and remember_key:
            saved = app._get_saved_window_geometry(remember_key)
        parsed_saved = _clamp_geometry(_parse_geometry(saved), screen_w, screen_h, min_w, min_h, max_w, max_h) if saved else None
        if parsed_saved:
            width, height, x, y = parsed_saved
        else:
            width = target_w
            height = target_h
            x = max(0, int((screen_w - width) / 2))
            y = max(0, int((screen_h - height) / 2))

        try:
            win.minsize(min_w, min_h)
        except Exception:
            pass
        try:
            win.maxsize(max_w, max_h)
        except Exception:
            pass
        win.geometry(f"{int(width)}x{int(height)}+{int(x)}+{int(y)}")

        if remember_key:
            _bind_window_persistence(win, remember_key)
            if app:
                app._schedule_remember_window(remember_key, win)
    except Exception:
        pass
