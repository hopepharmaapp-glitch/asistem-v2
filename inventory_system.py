# inventory_system.py
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
import json
import os
import re
import shutil
import subprocess
import textwrap
import hashlib
import getpass
from datetime import datetime
import webbrowser
try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
try:
    from pypdf import PdfReader
    PDF_READER_AVAILABLE = True
except ImportError:
    try:
        from PyPDF2 import PdfReader
        PDF_READER_AVAILABLE = True
    except ImportError:
        PdfReader = None
        PDF_READER_AVAILABLE = False
from PIL import Image, ImageTk
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, KeepInFrame
from reportlab.lib.units import inch
from pathlib import Path
import ui_undo

def reports_debug_log(data_folder, msg):
    try:
        path = os.path.join(data_folder, "reports_debug_log.txt")
        with open(path, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")
    except Exception:
        pass


_INVENTORY_MATCH_TOKEN_SYNONYMS = {
    "tab": "tablet",
    "tabs": "tablet",
    "tablets": "tablet",
    "cap": "capsule",
    "caps": "capsule",
    "capsules": "capsule",
    "amp": "ampoule",
    "amps": "ampoule",
    "ampoules": "ampoule",
    "inj": "injection",
    "syr": "syrup",
    "sol": "solution",
}

_INVENTORY_MATCH_IGNORABLE_TOKENS = {
    "item",
    "items",
    "product",
    "products",
    "ref",
    "reference",
    "code",
    "no",
    "number",
    "unit",
    "units",
    "pc",
    "pcs",
    "piece",
    "pieces",
}

def resolve_logo_path(data_folder):
    try:
        documents = os.path.join(os.path.expanduser("~"), "Documents")
        hope_folder = os.path.join(documents, "HopePharmaInvoices")
        candidates = [
            os.path.join(data_folder, "logo.png"),
            os.path.join(data_folder, "logo.jpg"),
            os.path.join(data_folder, "logo.jpeg"),
            os.path.join(data_folder, "AssistemLogo.png"),
            os.path.join(hope_folder, "logo.png"),
            os.path.join(hope_folder, "logo.jpg"),
            os.path.join(hope_folder, "logo.jpeg"),
            os.path.join(documents, "logo.png"),
            os.path.join(documents, "logo.jpg"),
            os.path.join(documents, "logo.jpeg"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.png"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpg"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "logo.jpeg"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "AssistemLogo.png"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "HopePharma.app", "Contents", "Resources", "logo.png"),
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist", "logo.png"),
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
    except Exception:
        return None
    return None

def build_pdf_logo_flowables(data_folder, width=1.1 * inch, height=1.1 * inch, spacer_height=0.1 * inch):
    flowables = []
    try:
        logo_path = resolve_logo_path(data_folder)
        if logo_path and os.path.exists(logo_path):
            flowables.append(RLImage(logo_path, width=width, height=height))
            if spacer_height:
                flowables.append(Spacer(1, spacer_height))
    except Exception:
        return []
    return flowables

def attach_clock_label(parent, anchor_side="right"):
    try:
        var = tk.StringVar()
        lbl = ttk.Label(parent, textvariable=var)
        def _tick():
            try:
                var.set(datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                parent.after(1000, _tick)
            except Exception:
                pass
        _tick()
        if isinstance(parent, (tk.Tk, tk.Toplevel)):
            frame = ttk.Frame(parent)
            frame.pack(fill="x", side="top", anchor="e")
            lbl.pack(in_=frame, side=anchor_side, padx=8, pady=2)
        else:
            lbl.pack(side=anchor_side, padx=8, pady=2)
        return lbl
    except Exception:
        return None

class InventoryItem:
    def __init__(self, item_id, name, category, description, total_cost, selling_price,
                 quantity, min_stock, supplier, batch_number=None, expiry_date=None,
                 created_date=None, last_updated=None, costs=None, addition_source="Manual",
                 invoice_date=None, is_storage_item=False, source_document_id=None,
                 source_invoice_number=None, source_file_name=None, max_sell_qty=999999, price_multiplier=1.0,
                 manufacture_date=None, package_type="", brand="", client_name=None,
                 sku="", subcategory="", unit_of_measure="PCS", pack_size="", vat_category="Standard",
                 barcode="", lot_number="", reserved_quantity=0, warehouse_code="", bin_location="",
                 max_stock=0, reorder_point=0, reorder_quantity=0, status="Available",
                 temperature_controlled=False, quarantine=False, supplier_name="", product_status="Stored"):
        self.item_id = item_id
        self.name = name
        self.category = category
        self.description = description
        self.total_cost = total_cost  # Total cost for all items
        self.selling_price = selling_price  # Selling price per item (optional)
        self.quantity = quantity
        self.min_stock = min_stock
        self.supplier = supplier
        self.batch_number = batch_number
        self.expiry_date = expiry_date
        self.created_date = created_date or datetime.now().strftime("%Y-%m-%d")
        self.last_updated = last_updated or datetime.now().strftime("%Y-%m-%d")
        self.costs = costs or []  # Track multiple costs
        self.addition_source = addition_source
        self.invoice_date = invoice_date or None
        self.is_storage_item = is_storage_item
        self.source_document_id = source_document_id or None
        self.source_invoice_number = source_invoice_number or None
        self.source_file_name = source_file_name or None
        self.max_sell_qty = max_sell_qty
        self.price_multiplier = price_multiplier
        self.manufacture_date = manufacture_date or None
        self.package_type = str(package_type or "").strip()
        self.brand = str(brand or "").strip()
        self.client_name = str(client_name or supplier or "").strip()
        self.sku = str(sku or item_id or "").strip()
        self.subcategory = str(subcategory or "").strip()
        self.unit_of_measure = str(unit_of_measure or "PCS").strip()
        self.pack_size = str(pack_size or "").strip()
        self.vat_category = str(vat_category or "Standard").strip()
        self.barcode = str(barcode or "").strip()
        self.lot_number = str(lot_number or "").strip()
        try:
            self.reserved_quantity = max(0.0, float(reserved_quantity or 0))
        except Exception:
            self.reserved_quantity = 0.0
        self.warehouse_code = str(warehouse_code or "").strip()
        self.bin_location = str(bin_location or "").strip()
        try:
            self.max_stock = max(0.0, float(max_stock or 0))
        except Exception:
            self.max_stock = 0.0
        try:
            self.reorder_point = max(0.0, float(reorder_point or 0))
        except Exception:
            self.reorder_point = 0.0
        try:
            self.reorder_quantity = max(0.0, float(reorder_quantity or 0))
        except Exception:
            self.reorder_quantity = 0.0
        self.status = str(status or "Available").strip() or "Available"
        self.temperature_controlled = bool(temperature_controlled)
        self.quarantine = bool(quarantine)
        self.supplier_name = str(supplier_name or supplier or "").strip()
        self.product_status = str(product_status or "Stored").strip() or "Stored"
    
    def to_dict(self):
        return {
            'item_id': self.item_id,
            'name': self.name,
            'category': self.category,
            'description': self.description,
            'total_cost': self.total_cost,
            'selling_price': self.selling_price,
            'quantity': self.quantity,
            'min_stock': self.min_stock,
            'supplier': self.supplier,
            'batch_number': self.batch_number,
            'expiry_date': self.expiry_date,
            'created_date': self.created_date,
            'last_updated': self.last_updated,
            'costs': self.costs,
            'addition_source': self.addition_source,
            'invoice_date': self.invoice_date,
            'is_storage_item': self.is_storage_item,
            'source_document_id': self.source_document_id,
            'source_invoice_number': self.source_invoice_number,
            'source_file_name': self.source_file_name,
            'max_sell_qty': self.max_sell_qty,
            'price_multiplier': self.price_multiplier,
            'manufacture_date': self.manufacture_date,
            'package_type': self.package_type,
            'brand': self.brand,
            'client_name': self.client_name,
            'sku': self.sku,
            'subcategory': self.subcategory,
            'unit_of_measure': self.unit_of_measure,
            'pack_size': self.pack_size,
            'vat_category': self.vat_category,
            'barcode': self.barcode,
            'lot_number': self.lot_number,
            'reserved_quantity': self.reserved_quantity,
            'available_quantity': self.get_available_quantity(),
            'warehouse_code': self.warehouse_code,
            'bin_location': self.bin_location,
            'max_stock': self.max_stock,
            'reorder_point': self.reorder_point,
            'reorder_quantity': self.reorder_quantity,
            'status': self.status,
            'temperature_controlled': self.temperature_controlled,
            'quarantine': self.quarantine,
            'supplier_name': self.supplier_name,
            'product_status': self.product_status,
        }
    
    @classmethod
    def from_dict(cls, data):
        """Create InventoryItem from dictionary, handling missing fields"""
        data = dict(data or {})

        # Handle old format compatibility
        if 'total_cost' not in data:
            # Calculate total cost from cost_price if it exists in old data
            cost_price = data.get('cost_price', 0)
            quantity = data.get('quantity', 1)
            data['total_cost'] = cost_price * quantity
        
        # Remove cost_price if it exists to avoid conflicts
        if 'cost_price' in data:
            del data['cost_price']
        # available_quantity is a derived value stored for reporting; do not pass it
        # back into the constructor during reload.
        if 'available_quantity' in data:
            del data['available_quantity']
            
        # Ensure all required fields are present with defaults
        defaults = {
            'total_cost': 0,
            'selling_price': 0,
            'quantity': 0,
            'min_stock': 0,
            'batch_number': None,
            'expiry_date': None,
            'created_date': datetime.now().strftime("%Y-%m-%d"),
            'last_updated': datetime.now().strftime("%Y-%m-%d"),
            'costs': [],
            'addition_source': "Manual",
            'invoice_date': None,
            'is_storage_item': False,
            'source_document_id': None,
            'source_invoice_number': None,
            'source_file_name': None,
            'max_sell_qty': 999999,
            'price_multiplier': None,
            'manufacture_date': None,
            'package_type': "",
            'brand': "",
            'client_name': "",
            'sku': "",
            'subcategory': "",
            'unit_of_measure': "PCS",
            'pack_size': "",
            'vat_category': "Standard",
            'barcode': "",
            'lot_number': "",
            'reserved_quantity': 0,
            'warehouse_code': "",
            'bin_location': "",
            'max_stock': 0,
            'reorder_point': 0,
            'reorder_quantity': 0,
            'status': "Available",
            'temperature_controlled': False,
            'quarantine': False,
            'supplier_name': "",
            'product_status': "Stored",
        }
        
        for field, default in defaults.items():
            if field not in data:
                data[field] = default

        if data.get('price_multiplier') is None:
            try:
                selling_price = float(data.get('selling_price') or 0.0)
                quantity = float(data.get('quantity') or 0.0)
                total_cost = float(data.get('total_cost') or 0.0)
                if selling_price > 0 and quantity > 0:
                    cost_per_item = total_cost / quantity
                    if cost_per_item > 0:
                        data['price_multiplier'] = selling_price / cost_per_item
                    else:
                        data['price_multiplier'] = 1.0
                else:
                    data['price_multiplier'] = 1.0
            except Exception:
                data['price_multiplier'] = 1.0

        try:
            data['max_sell_qty'] = int(data.get('max_sell_qty') or 999999)
        except Exception:
            data['max_sell_qty'] = 999999
        if data['max_sell_qty'] <= 0:
            data['max_sell_qty'] = 999999

        try:
            data['price_multiplier'] = float(data.get('price_multiplier') or 1.0)
        except Exception:
            data['price_multiplier'] = 1.0
        if data['price_multiplier'] <= 0:
            data['price_multiplier'] = 1.0
        try:
            data['reserved_quantity'] = max(0.0, float(data.get('reserved_quantity') or 0))
        except Exception:
            data['reserved_quantity'] = 0.0
                
        return cls(**data)
    
    def update_quantity(self, new_quantity):
        self.quantity = new_quantity
        self.last_updated = datetime.now().strftime("%Y-%m-%d")
    
    def is_low_stock(self):
        return self.quantity <= self.min_stock

    def get_available_quantity(self):
        try:
            return max(0.0, float(self.quantity or 0) - float(self.reserved_quantity or 0))
        except Exception:
            return max(0.0, float(self.quantity or 0))
    
    def is_expired(self):
        if not self.expiry_date:
            return False
        try:
            d = datetime.strptime(str(self.expiry_date).strip(), "%Y-%m-%d").date()
            return datetime.now().date() > d
        except Exception:
            return False

    def is_expiring_soon(self, days: int = 14):
        if not self.expiry_date:
            return False
        try:
            d = datetime.strptime(str(self.expiry_date).strip(), "%Y-%m-%d").date()
            diff = (d - datetime.now().date()).days
            return 0 <= diff <= days
        except Exception:
            return False
    
    def add_cost(self, cost_type, amount, description, date=None):
        """Add additional cost to item"""
        cost_id = f"COST{len(self.costs) + 1:03d}"
        cost_data = {
            'cost_id': cost_id,
            'cost_type': cost_type,
            'amount': amount,
            'description': description,
            'date': date or datetime.now().strftime("%Y-%m-%d")
        }
        self.costs.append(cost_data)
        self.total_cost += amount  # Update total cost
        self.last_updated = datetime.now().strftime("%Y-%m-%d")
        return cost_id
    
    def remove_cost(self, cost_id):
        """Remove cost from item"""
        for cost in self.costs:
            if cost['cost_id'] == cost_id:
                self.total_cost -= cost['amount']
                self.costs.remove(cost)
                break
        self.last_updated = datetime.now().strftime("%Y-%m-%d")
    
    def get_cost_per_item(self):
        """Calculate cost per individual item"""
        if self.quantity == 0:
            return 0
        return self.total_cost / self.quantity
    
    def get_total_value(self):
        """Get total inventory value at cost"""
        return self.total_cost
    
    def get_sales_value(self):
        """Get potential sales value"""
        return self.quantity * self.selling_price if self.selling_price else 0
    
    def get_profit_per_item(self):
        """Calculate profit per item"""
        if not self.selling_price:
            return 0
        cost_per_item = self.get_cost_per_item()
        return self.selling_price - cost_per_item
    
    def get_total_profit_potential(self):
        """Calculate total potential profit"""
        return self.get_profit_per_item() * self.quantity
    
    def get_cost_breakdown(self):
        """Get detailed cost breakdown"""
        cost_per_item = self.get_cost_per_item()
        breakdown = {
            'total_cost': self.total_cost,
            'cost_per_item': cost_per_item,
            'selling_price_per_item': self.selling_price,
            'profit_per_item': self.get_profit_per_item(),
            'total_profit_potential': self.get_total_profit_potential(),
            'additional_costs': sum(cost['amount'] for cost in self.costs),
            'cost_details': self.costs
        }
        return breakdown

class SaleRecord:
    def __init__(self, sale_id, item_id, item_name, quantity, selling_price, cost_per_item=None, 
                 total_amount=None, total_cost=None, total_profit=None, customer_name=None, 
                 sale_date=None, invoice_id=None, promotion_quantity: int = 0,
                 vat_applied: bool = False, vat_rate: float = 0.05, vat_amount: float = 0.0):
        self.sale_id = sale_id
        self.item_id = item_id
        self.item_name = item_name
        self.quantity = quantity
        self.selling_price = selling_price  # Selling price per item
        self.promotion_quantity = promotion_quantity
        self.vat_applied = vat_applied
        self.vat_rate = vat_rate
        self.vat_amount = vat_amount
        
        # Calculate missing values if not provided (for backward compatibility)
        if cost_per_item is None:
            cost_per_item = 0
        if total_amount is None:
            total_amount = quantity * selling_price
        if total_cost is None:
            total_cost = cost_per_item * (quantity + promotion_quantity)
        if total_profit is None:
            total_profit = total_amount - total_cost
            
        self.cost_per_item = cost_per_item  # Cost per item at time of sale
        self.total_amount = total_amount  # Total sales amount
        self.total_cost = total_cost  # Total cost for sold items
        self.total_profit = total_profit  # Total profit
        self.customer_name = customer_name or "Unknown"
        self.sale_date = sale_date or datetime.now().strftime("%Y-%m-%d")
        self.invoice_id = invoice_id
    
    def to_dict(self):
        return {
            'sale_id': self.sale_id,
            'item_id': self.item_id,
            'item_name': self.item_name,
            'quantity': self.quantity,
            'promotion_quantity': self.promotion_quantity,
            'selling_price': self.selling_price,
            'cost_per_item': self.cost_per_item,
            'total_amount': self.total_amount,
            'total_cost': self.total_cost,
            'total_profit': self.total_profit,
            'vat_applied': self.vat_applied,
            'vat_rate': self.vat_rate,
            'vat_amount': self.vat_amount,
            'customer_name': self.customer_name,
            'sale_date': self.sale_date,
            'invoice_id': self.invoice_id
        }
    
    @classmethod
    def from_dict(cls, data):
        defaults = {
            'cost_per_item': 0,
            'total_amount': data.get('quantity', 0) * data.get('selling_price', 0),
            'total_cost': 0,
            'total_profit': 0,
            'promotion_quantity': data.get('promotion_quantity', 0),
            'vat_applied': data.get('vat_applied', False),
            'vat_rate': data.get('vat_rate', 0.05),
            'vat_amount': data.get('vat_amount', 0.0),
            'customer_name': 'Unknown',
            'sale_date': datetime.now().strftime("%Y-%m-%d"),
            'invoice_id': None
        }
        
        for field, default in defaults.items():
            if field not in data:
                data[field] = default
                
        if data['total_amount'] == 0 and 'quantity' in data and 'selling_price' in data:
            data['total_amount'] = data['quantity'] * data['selling_price']
        if data['total_cost'] == 0 and 'cost_per_item' in data and 'quantity' in data:
            data['total_cost'] = data['cost_per_item'] * (data.get('quantity', 0) + data.get('promotion_quantity', 0))
        if data['total_profit'] == 0 and data['total_amount'] > 0:
            data['total_profit'] = data['total_amount'] - data['total_cost']
            
        return cls(**data)

class ReturnRecord:
    def __init__(
        self,
        return_id,
        sale_id,
        item_id,
        item_name,
        quantity,
        unit_price,
        customer_name,
        return_date=None,
        reason="",
        processed_by="",
        invoice_id=None,
        vat_applied=False,
        vat_rate=0.0,
        vat_amount=0.0,
        subtotal=0.0,
        total_credit=0.0,
        restocking_fee=0.0,
        credit_note_id=None,
        notes="",
    ):
        self.return_id = return_id
        self.sale_id = sale_id
        self.item_id = item_id
        self.item_name = item_name
        self.quantity = quantity
        self.unit_price = unit_price
        self.customer_name = customer_name or "Unknown"
        self.return_date = return_date or datetime.now().strftime("%Y-%m-%d")
        self.reason = reason or ""
        self.processed_by = processed_by or ""
        self.invoice_id = invoice_id
        self.vat_applied = bool(vat_applied)
        self.vat_rate = float(vat_rate or 0.0)
        self.vat_amount = float(vat_amount or 0.0)
        self.subtotal = float(subtotal or 0.0)
        self.total_credit = float(total_credit or 0.0)
        self.restocking_fee = float(restocking_fee or 0.0)
        self.credit_note_id = credit_note_id
        self.notes = notes or ""

    def to_dict(self):
        return {
            "return_id": self.return_id,
            "sale_id": self.sale_id,
            "item_id": self.item_id,
            "item_name": self.item_name,
            "quantity": self.quantity,
            "unit_price": self.unit_price,
            "customer_name": self.customer_name,
            "return_date": self.return_date,
            "reason": self.reason,
            "processed_by": self.processed_by,
            "invoice_id": self.invoice_id,
            "vat_applied": self.vat_applied,
            "vat_rate": self.vat_rate,
            "vat_amount": self.vat_amount,
            "subtotal": self.subtotal,
            "total_credit": self.total_credit,
            "restocking_fee": self.restocking_fee,
            "credit_note_id": self.credit_note_id,
            "notes": self.notes,
        }

    @classmethod
    def from_dict(cls, data):
        defaults = {
            "return_date": datetime.now().strftime("%Y-%m-%d"),
            "reason": "",
            "processed_by": "",
            "invoice_id": None,
            "vat_applied": False,
            "vat_rate": 0.0,
            "vat_amount": 0.0,
            "subtotal": 0.0,
            "total_credit": 0.0,
            "restocking_fee": 0.0,
            "credit_note_id": None,
            "notes": "",
        }
        for field, default in defaults.items():
            if field not in data:
                data[field] = default
        return cls(**data)

class InventoryManager:
    def __init__(self, data_folder):
        self.data_folder = data_folder
        self.inventory_file = os.path.join(data_folder, "inventory.json")
        self.sales_file = os.path.join(data_folder, "sales_records.json")
        self.invoices_file = os.path.join(data_folder, "invoices_data.json")
        self.settings_file = os.path.join(data_folder, "app_settings.json")
        self.categories_file = os.path.join(data_folder, "inventory_categories.json")
        self.cost_types_file = os.path.join(data_folder, "cost_types.json")
        self.delivery_notes_file = os.path.join(data_folder, "delivery_notes.json")
        self.customers_file = os.path.join(data_folder, "customers.json")
        self.suppliers_file = os.path.join(data_folder, "suppliers.json")
        self.supplier_records_file = os.path.join(data_folder, "supplier_records.json")
        self.supplier_invoice_imports_file = os.path.join(data_folder, "supplier_invoice_imports.json")
        self.supplier_documents_folder = os.path.join(data_folder, "SupplierDocuments")
        self.returns_file = os.path.join(data_folder, "returns.json")
        self.credit_notes_file = os.path.join(data_folder, "credit_notes.json")
        self.audit_log_file = os.path.join(data_folder, "audit_log.jsonl")
        
        try:
            if not os.path.exists(self.inventory_file):
                # Create default empty inventory if missing
                with open(self.inventory_file, 'w') as f:
                    json.dump([], f)
                print(f"Created missing inventory file: {self.inventory_file}")

            if not os.path.exists(self.sales_file):
                # Create default empty sales records if missing
                with open(self.sales_file, 'w') as f:
                    json.dump([], f)
                print(f"Created missing sales records file: {self.sales_file}")

            if not os.path.exists(self.delivery_notes_file):
                with open(self.delivery_notes_file, 'w') as f:
                    json.dump([], f)
                print(f"Created missing delivery notes file: {self.delivery_notes_file}")

            if not os.path.exists(self.returns_file):
                with open(self.returns_file, 'w') as f:
                    json.dump([], f)
                print(f"Created missing returns file: {self.returns_file}")

            if not os.path.exists(self.credit_notes_file):
                with open(self.credit_notes_file, 'w') as f:
                    json.dump([], f)
                print(f"Created missing credit notes file: {self.credit_notes_file}")

            self.items = self.load_inventory()
            self.sales = self.load_sales()
            self.returns = self.load_returns()
            self.credit_notes = self.load_credit_notes()
            self.categories = self.load_categories()
            self.cost_types = self.load_cost_types()
        except Exception as e:
            print(f"Error initializing data files: {e}")
    
    def load_customers(self):
        """Load list of unique customer names from customers.json"""
        try:
            if os.path.exists(self.customers_file):
                with open(self.customers_file, 'r') as f:
                    return sorted(json.load(f))
            return []
        except Exception:
            return []

    def save_customer(self, name):
        """Add new customer name to customers.json if it doesn't exist"""
        if not name or not isinstance(name, str):
            return False
        name = name.strip()
        if not name:
            return False
            
        try:
            customers = self.load_customers()
            if name not in customers:
                customers.append(name)
                with open(self.customers_file, 'w') as f:
                    json.dump(customers, f, indent=2)
                return True
        except Exception:
            pass
        return False

    def load_suppliers(self):
        """Load list of unique supplier names from suppliers.json"""
        try:
            if os.path.exists(self.suppliers_file):
                with open(self.suppliers_file, 'r') as f:
                    data = json.load(f) or []
                if isinstance(data, list):
                    return sorted({str(x).strip() for x in data if str(x).strip()})
            return []
        except Exception:
            return []

    def save_supplier(self, name):
        """Add new supplier name to suppliers.json if it doesn't exist"""
        if not name or not isinstance(name, str):
            return False
        name = name.strip()
        if not name:
            return False
        try:
            suppliers = self.load_suppliers()
            if name not in suppliers:
                suppliers.append(name)
                with open(self.suppliers_file, 'w') as f:
                    json.dump(sorted(suppliers), f, indent=2)
                return True
        except Exception:
            pass
        return False

    def _load_json_file(self, file_path, default):
        try:
            if os.path.exists(file_path):
                with open(file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return data if data is not None else default
        except Exception:
            pass
        return default

    def _save_json_file(self, file_path, data):
        try:
            folder = os.path.dirname(file_path)
            if folder:
                os.makedirs(folder, exist_ok=True)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2)
            return True
        except Exception:
            return False

    def _normalize_supplier_name(self, name):
        return str(name or "").strip()

    def _safe_filename(self, value):
        cleaned = re.sub(r'[^A-Za-z0-9._-]+', '_', str(value or "").strip())
        return cleaned.strip("._") or "document"

    def _normalize_date_value(self, value):
        text = str(value or "").strip()
        if not text:
            return ""
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y", "%d.%m.%Y"):
            try:
                return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
            except Exception:
                continue
        try:
            return datetime.fromisoformat(text).strftime("%Y-%m-%d")
        except Exception:
            return text

    def _to_float(self, value, default=0.0):
        try:
            if value is None:
                return default
            if isinstance(value, str):
                cleaned = value.replace(",", "").strip()
                filtered = "".join(ch for ch in cleaned if ch.isdigit() or ch in ".-")
                if filtered and any(ch.isdigit() for ch in filtered):
                    cleaned = filtered
                if not cleaned:
                    return default
                return float(cleaned)
            return float(value)
        except Exception:
            return default

    def generate_supplier_id(self, records=None):
        records = records if records is not None else self.load_supplier_records()
        existing = set()
        for record in records:
            supplier_id = str(record.get('supplier_id', ''))
            if supplier_id.startswith('SUP') and supplier_id[3:].isdigit():
                existing.add(int(supplier_id[3:]))
        next_id = 1
        while next_id in existing:
            next_id += 1
        return f"SUP{next_id:03d}"

    def load_supplier_records(self):
        raw_records = self._load_json_file(self.supplier_records_file, [])
        records = []
        used_names = set()
        used_numbers = set()
        changed = False

        def _alloc_supplier_id():
            next_num = 1
            while next_num in used_numbers:
                next_num += 1
            used_numbers.add(next_num)
            return f"SUP{next_num:03d}"

        if isinstance(raw_records, list):
            for entry in raw_records:
                if isinstance(entry, str):
                    entry = {'name': entry}
                if not isinstance(entry, dict):
                    continue
                name = self._normalize_supplier_name(entry.get('name'))
                if not name:
                    continue
                key = name.lower()
                if key in used_names:
                    continue
                used_names.add(key)
                supplier_id = str(entry.get('supplier_id') or "").strip()
                if supplier_id.startswith('SUP') and supplier_id[3:].isdigit():
                    try:
                        num = int(supplier_id[3:])
                        if num in used_numbers:
                            supplier_id = _alloc_supplier_id()
                            changed = True
                        else:
                            used_numbers.add(num)
                    except Exception:
                        supplier_id = _alloc_supplier_id()
                        changed = True
                else:
                    supplier_id = _alloc_supplier_id()
                    changed = True
                records.append({
                    'supplier_id': supplier_id,
                    'name': name,
                    'contact_person': str(entry.get('contact_person') or "").strip(),
                    'phone': str(entry.get('phone') or "").strip(),
                    'email': str(entry.get('email') or "").strip(),
                    'address': str(entry.get('address') or "").strip(),
                    'notes': str(entry.get('notes') or "").strip(),
                    'created_date': self._normalize_date_value(entry.get('created_date')) or datetime.now().strftime("%Y-%m-%d"),
                    'last_updated': self._normalize_date_value(entry.get('last_updated')) or datetime.now().strftime("%Y-%m-%d"),
                })

        for supplier_name in self.load_suppliers():
            key = supplier_name.lower()
            if key in used_names:
                continue
            records.append({
                'supplier_id': _alloc_supplier_id(),
                'name': supplier_name,
                'contact_person': "",
                'phone': "",
                'email': "",
                'address': "",
                'notes': "",
                'created_date': datetime.now().strftime("%Y-%m-%d"),
                'last_updated': datetime.now().strftime("%Y-%m-%d"),
            })
            used_names.add(key)
            changed = True

        records.sort(key=lambda r: r.get('name', '').lower())
        if changed:
            self._save_json_file(self.supplier_records_file, records)
        return records

    def save_supplier_records(self, records):
        normalized = []
        seen = set()
        for record in records:
            if not isinstance(record, dict):
                continue
            name = self._normalize_supplier_name(record.get('name'))
            if not name:
                continue
            key = name.lower()
            if key in seen:
                continue
            seen.add(key)
            normalized_record = {
                'supplier_id': str(record.get('supplier_id') or self.generate_supplier_id(normalized)),
                'name': name,
                'contact_person': str(record.get('contact_person') or "").strip(),
                'phone': str(record.get('phone') or "").strip(),
                'email': str(record.get('email') or "").strip(),
                'address': str(record.get('address') or "").strip(),
                'notes': str(record.get('notes') or "").strip(),
                'created_date': self._normalize_date_value(record.get('created_date')) or datetime.now().strftime("%Y-%m-%d"),
                'last_updated': self._normalize_date_value(record.get('last_updated')) or datetime.now().strftime("%Y-%m-%d"),
            }
            normalized.append(normalized_record)
            try:
                self.save_supplier(name)
            except Exception:
                pass
        normalized.sort(key=lambda r: r.get('name', '').lower())
        return self._save_json_file(self.supplier_records_file, normalized)

    def get_supplier_record(self, supplier_id):
        supplier_id = str(supplier_id or "").strip()
        for record in self.load_supplier_records():
            if record.get('supplier_id') == supplier_id:
                return record
        return None

    def add_supplier_record(self, name, contact_person="", phone="", email="", address="", notes=""):
        name = self._normalize_supplier_name(name)
        if not name:
            return False, "Supplier name is required.", None

        records = self.load_supplier_records()
        if any(self._normalize_supplier_name(r.get('name')).lower() == name.lower() for r in records):
            return False, "A supplier with this name already exists.", None

        record = {
            'supplier_id': self.generate_supplier_id(records),
            'name': name,
            'contact_person': str(contact_person or "").strip(),
            'phone': str(phone or "").strip(),
            'email': str(email or "").strip(),
            'address': str(address or "").strip(),
            'notes': str(notes or "").strip(),
            'created_date': datetime.now().strftime("%Y-%m-%d"),
            'last_updated': datetime.now().strftime("%Y-%m-%d"),
        }
        records.append(record)
        if self.save_supplier_records(records):
            return True, "Supplier added successfully.", record
        return False, "Failed to save supplier.", None

    def update_supplier_record(self, supplier_id, name, contact_person="", phone="", email="", address="", notes=""):
        supplier_id = str(supplier_id or "").strip()
        name = self._normalize_supplier_name(name)
        if not supplier_id:
            return False, "Supplier ID is missing."
        if not name:
            return False, "Supplier name is required."

        records = self.load_supplier_records()
        for record in records:
            other_name = self._normalize_supplier_name(record.get('name'))
            if record.get('supplier_id') != supplier_id and other_name.lower() == name.lower():
                return False, "Another supplier already uses this name."

        old_name = None
        updated = False
        for record in records:
            if record.get('supplier_id') != supplier_id:
                continue
            old_name = self._normalize_supplier_name(record.get('name'))
            record['name'] = name
            record['contact_person'] = str(contact_person or "").strip()
            record['phone'] = str(phone or "").strip()
            record['email'] = str(email or "").strip()
            record['address'] = str(address or "").strip()
            record['notes'] = str(notes or "").strip()
            record['last_updated'] = datetime.now().strftime("%Y-%m-%d")
            updated = True
            break

        if not updated:
            return False, "Supplier not found."

        if old_name and old_name != name:
            for item in self.items:
                if self._normalize_supplier_name(getattr(item, 'supplier', '')) == old_name:
                    item.supplier = name
            invoice_imports = self.load_supplier_invoice_imports()
            changed = False
            for invoice_import in invoice_imports:
                if invoice_import.get('supplier_id') == supplier_id:
                    invoice_import['supplier_name'] = name
                    changed = True
            if changed:
                self.save_supplier_invoice_imports(invoice_imports)
            self.save_inventory()

        if self.save_supplier_records(records):
            return True, "Supplier updated successfully."
        return False, "Failed to update supplier."

    def delete_supplier_record(self, supplier_id):
        supplier = self.get_supplier_record(supplier_id)
        if not supplier:
            return False, "Supplier not found."

        supplier_name = self._normalize_supplier_name(supplier.get('name'))
        linked_items = [item for item in self.items if self._normalize_supplier_name(getattr(item, 'supplier', '')) == supplier_name]
        if linked_items:
            return False, "This supplier is still linked to inventory items. Update or remove those items first."

        linked_docs = [doc for doc in self.load_supplier_invoice_imports() if doc.get('supplier_id') == supplier_id]
        if linked_docs:
            return False, "This supplier still has uploaded invoices. Remove those invoices first."

        records = [record for record in self.load_supplier_records() if record.get('supplier_id') != supplier_id]
        if self.save_supplier_records(records):
            try:
                suppliers = self.load_suppliers()
                remaining = [s for s in suppliers if str(s).strip().lower() != supplier_name.lower()]
                self._save_json_file(self.suppliers_file, sorted({str(x).strip() for x in remaining if str(x).strip()}))
            except Exception:
                pass
            return True, "Supplier deleted successfully."
        return False, "Failed to delete supplier."

    def generate_supplier_document_id(self, records=None):
        imports = records if records is not None else self.load_supplier_invoice_imports()
        existing = []
        for record in imports:
            document_id = str(record.get('document_id', ''))
            if document_id.startswith('SUPDOC') and document_id[6:].isdigit():
                existing.append(int(document_id[6:]))
        next_id = (max(existing) + 1) if existing else 1
        return f"SUPDOC{next_id:04d}"

    def load_supplier_invoice_imports(self):
        raw_records = self._load_json_file(self.supplier_invoice_imports_file, [])
        normalized = []
        next_id = 1

        def _next_document_id():
            nonlocal next_id
            document_id = f"SUPDOC{next_id:04d}"
            next_id += 1
            return document_id

        if not isinstance(raw_records, list):
            return normalized

        for entry in raw_records:
            if not isinstance(entry, dict):
                continue
            document_id = str(entry.get('document_id') or "").strip()
            if document_id.startswith('SUPDOC') and document_id[6:].isdigit():
                try:
                    next_id = max(next_id, int(document_id[6:]) + 1)
                except Exception:
                    pass
            else:
                document_id = _next_document_id()
            normalized.append({
                'document_id': document_id,
                'supplier_id': str(entry.get('supplier_id') or "").strip(),
                'supplier_name': self._normalize_supplier_name(entry.get('supplier_name')) or "Unknown Supplier",
                'file_name': str(entry.get('file_name') or "").strip(),
                'stored_path': str(entry.get('stored_path') or "").strip(),
                'file_type': str(entry.get('file_type') or "").strip(),
                'invoice_number': str(entry.get('invoice_number') or "").strip(),
                'document_date': self._normalize_date_value(entry.get('document_date')),
                'upload_date': str(entry.get('upload_date') or datetime.now().strftime("%Y-%m-%d %H:%M:%S")).strip(),
                'due_date': self._normalize_date_value(entry.get('due_date')),
                'subtotal': self._to_float(entry.get('subtotal'), 0.0),
                'tax_amount': self._to_float(entry.get('tax_amount'), 0.0),
                'grand_total': self._to_float(entry.get('grand_total'), 0.0),
                'currency': str(entry.get('currency') or "AED").strip(),
                'status': str(entry.get('status') or "parsed").strip(),
                'notes': str(entry.get('notes') or "").strip(),
                'warnings': [str(x) for x in (entry.get('warnings') or [])],
                'errors': [str(x) for x in (entry.get('errors') or [])],
                'inventory_item_ids': [
                    str(item_id).strip()
                    for item_id in (entry.get('inventory_item_ids') or [])
                    if str(item_id).strip()
                ],
                'items': entry.get('items') or [],
            })
        normalized.sort(key=lambda r: (r.get('supplier_name', '').lower(), r.get('document_date', ''), r.get('upload_date', '')))
        return normalized

    def save_supplier_invoice_imports(self, records):
        normalized = []
        for record in records:
            if not isinstance(record, dict):
                continue
            normalized.append({
                'document_id': str(record.get('document_id') or self.generate_supplier_document_id(normalized)),
                'supplier_id': str(record.get('supplier_id') or "").strip(),
                'supplier_name': self._normalize_supplier_name(record.get('supplier_name')) or "Unknown Supplier",
                'file_name': str(record.get('file_name') or "").strip(),
                'stored_path': str(record.get('stored_path') or "").strip(),
                'file_type': str(record.get('file_type') or "").strip(),
                'invoice_number': str(record.get('invoice_number') or "").strip(),
                'document_date': self._normalize_date_value(record.get('document_date')),
                'upload_date': str(record.get('upload_date') or datetime.now().strftime("%Y-%m-%d %H:%M:%S")).strip(),
                'due_date': self._normalize_date_value(record.get('due_date')),
                'subtotal': round(self._to_float(record.get('subtotal'), 0.0), 2),
                'tax_amount': round(self._to_float(record.get('tax_amount'), 0.0), 2),
                'grand_total': round(self._to_float(record.get('grand_total'), 0.0), 2),
                'currency': str(record.get('currency') or "AED").strip(),
                'status': str(record.get('status') or "parsed").strip(),
                'notes': str(record.get('notes') or "").strip(),
                'warnings': [str(x) for x in (record.get('warnings') or [])],
                'errors': [str(x) for x in (record.get('errors') or [])],
                'inventory_item_ids': [
                    str(item_id).strip()
                    for item_id in (record.get('inventory_item_ids') or [])
                    if str(item_id).strip()
                ],
                'items': [
                    {
                        'description': str(item.get('description') or "").strip(),
                        'quantity': self._to_float(item.get('quantity'), 0.0),
                        'unit_price': self._to_float(item.get('unit_price'), 0.0),
                        'total': self._to_float(item.get('total'), 0.0),
                        'taxable': bool(item.get('taxable', False)),
                    }
                    for item in (record.get('items') or [])
                    if isinstance(item, dict)
                ],
            })
        normalized.sort(key=lambda r: (r.get('supplier_name', '').lower(), r.get('document_date', ''), r.get('upload_date', '')))
        return self._save_json_file(self.supplier_invoice_imports_file, normalized)

    def get_supplier_invoice_import(self, document_id):
        document_id = str(document_id or "").strip()
        for record in self.load_supplier_invoice_imports():
            if record.get('document_id') == document_id:
                return record
        return None

    def update_supplier_invoice_import(self, document_id, **updates):
        document_id = str(document_id or "").strip()
        if not document_id:
            return False, "Invoice upload not found.", None

        records = self.load_supplier_invoice_imports()
        updated_record = None
        for record in records:
            if record.get('document_id') != document_id:
                continue

            if 'invoice_number' in updates:
                record['invoice_number'] = str(updates.get('invoice_number') or "").strip()
            if 'document_date' in updates:
                record['document_date'] = self._normalize_date_value(updates.get('document_date'))
            if 'due_date' in updates:
                record['due_date'] = self._normalize_date_value(updates.get('due_date'))
            if 'subtotal' in updates:
                record['subtotal'] = round(self._to_float(updates.get('subtotal'), 0.0), 2)
            if 'tax_amount' in updates:
                record['tax_amount'] = round(self._to_float(updates.get('tax_amount'), 0.0), 2)
            if 'grand_total' in updates:
                record['grand_total'] = round(self._to_float(updates.get('grand_total'), 0.0), 2)
            if 'notes' in updates:
                record['notes'] = str(updates.get('notes') or "").strip()
            if 'items' in updates:
                record['items'] = [
                    {
                        'description': str(item.get('description') or "").strip(),
                        'quantity': self._to_float(item.get('quantity'), 0.0),
                        'unit_price': self._to_float(item.get('unit_price'), 0.0),
                        'total': self._to_float(item.get('total'), 0.0),
                        'taxable': bool(item.get('taxable', False)),
                    }
                    for item in (updates.get('items') or [])
                    if isinstance(item, dict) and str(item.get('description') or "").strip()
                ]
            updated_record = record
            break

        if not updated_record:
            return False, "Invoice upload not found.", None

        synced, sync_message, synced_ids = self.sync_supplier_invoice_to_inventory(updated_record)
        updated_record['inventory_item_ids'] = synced_ids

        if synced and self.save_supplier_invoice_imports(records):
            refreshed = self.get_supplier_invoice_import(document_id)
            message = "Supplier invoice updated successfully."
            if sync_message:
                message += f"\n{sync_message}"
            return True, message, refreshed
        if not synced and self.save_supplier_invoice_imports(records):
            refreshed = self.get_supplier_invoice_import(document_id)
            message = f"Supplier invoice updated, but inventory sync had a warning.\n{sync_message}"
            return True, message, refreshed
        return False, "Failed to update supplier invoice.", None

    def get_supplier_invoice_summary(self, start_date=None, end_date=None, supplier_id=None):
        start_date = str(start_date or "").strip()
        end_date = str(end_date or "").strip()
        supplier_id = str(supplier_id or "").strip()

        supplier_records = {record.get('supplier_id', ''): record for record in self.load_supplier_records()}
        summary_map = {}
        invoice_rows = []

        for record in self.load_supplier_invoice_imports():
            record_supplier_id = str(record.get('supplier_id') or "").strip()
            if supplier_id and record_supplier_id != supplier_id:
                continue

            report_date = str(record.get('document_date') or "").strip()
            if not report_date:
                report_date = str(record.get('upload_date') or "").strip()[:10]

            if start_date and report_date and report_date < start_date:
                continue
            if end_date and report_date and report_date > end_date:
                continue

            supplier_name = self._normalize_supplier_name(record.get('supplier_name')) or "Unknown Supplier"
            key = record_supplier_id or supplier_name.lower()
            summary = summary_map.setdefault(key, {
                'supplier_id': record_supplier_id,
                'supplier_name': supplier_name,
                'invoice_count': 0,
                'total_amount': 0.0,
                'invoice_rows': [],
                'contact_person': '',
                'phone': '',
                'email': '',
            })

            supplier_info = supplier_records.get(record_supplier_id, {})
            if supplier_info:
                summary['contact_person'] = supplier_info.get('contact_person', '')
                summary['phone'] = supplier_info.get('phone', '')
                summary['email'] = supplier_info.get('email', '')

            invoice_row = {
                'supplier_id': record_supplier_id,
                'supplier_name': supplier_name,
                'document_id': record.get('document_id', ''),
                'invoice_number': record.get('invoice_number', '') or record.get('document_id', ''),
                'document_date': report_date,
                'grand_total': round(self._to_float(record.get('grand_total'), 0.0), 2),
                'subtotal': round(self._to_float(record.get('subtotal'), 0.0), 2),
                'tax_amount': round(self._to_float(record.get('tax_amount'), 0.0), 2),
                'file_name': record.get('file_name', ''),
                'item_count': len(record.get('items') or []),
            }
            invoice_rows.append(invoice_row)
            summary['invoice_rows'].append(invoice_row)
            summary['invoice_count'] += 1
            summary['total_amount'] += invoice_row['grand_total']

        supplier_summaries = []
        for summary in summary_map.values():
            summary['total_amount'] = round(summary['total_amount'], 2)
            summary['invoice_rows'].sort(key=lambda row: (row.get('document_date', ''), row.get('invoice_number', '')))
            supplier_summaries.append(summary)

        supplier_summaries.sort(key=lambda row: row.get('supplier_name', '').lower())
        invoice_rows.sort(key=lambda row: (row.get('supplier_name', '').lower(), row.get('document_date', ''), row.get('invoice_number', '')))
        return {
            'supplier_summaries': supplier_summaries,
            'invoice_rows': invoice_rows,
            'supplier_count': len(supplier_summaries),
            'invoice_count': len(invoice_rows),
            'grand_total': round(sum(row.get('grand_total', 0.0) for row in invoice_rows), 2),
        }

    def _extract_labeled_text(self, lines, labels):
        label_set = [label.lower() for label in labels]
        for line in lines:
            lower = line.lower()
            for label in label_set:
                if label not in lower:
                    continue
                if ':' in line:
                    value = line.split(':', 1)[1].strip()
                    if value:
                        return value
                match = re.search(re.escape(label) + r'\s+([A-Za-z0-9./_-]+.*)$', lower)
                if match:
                    original_value = line[len(line) - len(match.group(1)):].strip()
                    if original_value:
                        return original_value
        return ""

    def _extract_labeled_amount(self, lines, labels):
        label_set = [label.lower() for label in labels]
        for line in reversed(lines):
            lower = line.lower()
            if not any(label in lower for label in label_set):
                continue
            matches = re.findall(r'-?\d[\d,]*\.?\d*', line)
            if matches:
                return self._to_float(matches[-1], 0.0)
        return 0.0

    def _extract_labeled_date(self, lines, labels):
        text = self._extract_labeled_text(lines, labels)
        if not text:
            return ""
        match = re.search(r'(\d{4}-\d{2}-\d{2}|\d{2}[/-]\d{2}[/-]\d{4}|\d{2}\.\d{2}\.\d{4})', text)
        if match:
            return self._normalize_date_value(match.group(1))
        return self._normalize_date_value(text)

    def _extract_pdf_invoice_data(self, file_path, supplier_name=""):
        if not PDF_READER_AVAILABLE:
            return None, ["PDF reader library is not available."], ["PDF extraction is unavailable in this build."]

        warnings = []
        errors = []
        try:
            reader = PdfReader(file_path)
            text_parts = []
            for page in reader.pages:
                try:
                    page_text = page.extract_text() or ""
                except Exception:
                    page_text = ""
                if page_text.strip():
                    text_parts.append(page_text)
            text = "\n".join(text_parts).strip()
        except Exception as exc:
            return None, warnings, [f"Error reading PDF file: {exc}"]

        if not text:
            return None, warnings, ["No readable text was found in this PDF."]

        lines = [re.sub(r'\s+', ' ', line).strip() for line in text.splitlines() if str(line).strip()]
        invoice_number = self._extract_labeled_text(lines, ['invoice no', 'invoice number', 'invoice #', 'invoice'])
        document_date = self._extract_labeled_date(lines, ['invoice date', 'date'])
        due_date = self._extract_labeled_date(lines, ['due date', 'payment due'])
        subtotal = self._extract_labeled_amount(lines, ['subtotal', 'sub total', 'net total'])
        tax_amount = self._extract_labeled_amount(lines, ['vat amount', 'tax amount', 'vat', 'tax'])
        grand_total = self._extract_labeled_amount(lines, ['grand total', 'amount due', 'total due', 'invoice total', 'total'])

        item_rows = []
        skip_keywords = (
            'invoice', 'date', 'subtotal', 'tax', 'vat', 'total', 'amount due',
            'grand total', 'phone', 'email', 'address', 'terms', 'supplier', 'customer'
        )
        for line in lines:
            lower = line.lower()
            if any(keyword in lower for keyword in skip_keywords):
                continue
            parts = re.split(r'\s{2,}', line)
            if len(parts) < 4:
                match = re.match(r'^(?P<desc>.+?)\s+(?P<qty>\d+(?:\.\d+)?)\s+(?P<unit>\d[\d,]*\.?\d*)\s+(?P<total>\d[\d,]*\.?\d*)$', line)
                if not match:
                    continue
                parts = [match.group('desc'), match.group('qty'), match.group('unit'), match.group('total')]
            desc = str(parts[0]).strip()
            qty = self._to_float(parts[-3], 0.0)
            unit_price = self._to_float(parts[-2], 0.0)
            line_total = self._to_float(parts[-1], 0.0)
            if not desc or qty <= 0 or line_total <= 0:
                continue
            item_rows.append({
                'description': desc,
                'quantity': qty,
                'unit_price': unit_price,
                'total': line_total,
                'taxable': tax_amount > 0,
            })

        status = 'parsed'
        if not item_rows and grand_total > 0:
            warnings.append("PDF line items were not detected clearly, so one summary line was created from the invoice total.")
            item_rows.append({
                'description': invoice_number or f"Imported PDF invoice from {supplier_name or 'supplier'}",
                'quantity': 1.0,
                'unit_price': grand_total,
                'total': grand_total,
                'taxable': tax_amount > 0,
            })
            status = 'partial'
        elif not item_rows:
            status = 'failed'
            errors.append("Could not detect invoice items in the PDF.")
        elif not invoice_number or not document_date:
            status = 'partial'
            warnings.append("Some PDF invoice header fields could not be detected automatically.")

        if subtotal <= 0 and item_rows:
            subtotal = sum(self._to_float(item.get('total'), 0.0) for item in item_rows)
        if grand_total <= 0 and (subtotal > 0 or tax_amount > 0):
            grand_total = subtotal + tax_amount

        return {
            'invoice_number': invoice_number or os.path.splitext(os.path.basename(file_path))[0],
            'document_date': document_date or datetime.now().strftime("%Y-%m-%d"),
            'due_date': due_date,
            'subtotal': subtotal,
            'tax_amount': tax_amount,
            'grand_total': grand_total,
            'currency': 'AED',
            'status': status,
            'notes': 'Imported from PDF',
            'items': item_rows,
        }, warnings, errors

    def _extract_excel_invoice_data(self, file_path):
        invoice_dict, warnings, errors = self.parse_invoice_excel(file_path, prefer_actual_rate=False)
        if not invoice_dict:
            return None, warnings, errors
        status = 'parsed' if not warnings and not errors else 'partial'
        items = invoice_dict.get('items') or []
        return {
            'invoice_number': os.path.splitext(os.path.basename(file_path))[0],
            'document_date': self._normalize_date_value(invoice_dict.get('date')) or datetime.now().strftime("%Y-%m-%d"),
            'due_date': self._normalize_date_value(invoice_dict.get('due_date')),
            'subtotal': self._to_float(invoice_dict.get('subtotal'), 0.0),
            'tax_amount': self._to_float(invoice_dict.get('tax_amount'), 0.0),
            'grand_total': self._to_float(invoice_dict.get('grand_total'), 0.0),
            'currency': str(invoice_dict.get('currency') or 'AED'),
            'status': status,
            'notes': str(invoice_dict.get('notes') or 'Imported from Excel'),
            'items': [
                {
                    'description': str(item.get('description') or "").strip(),
                    'quantity': self._to_float(item.get('quantity'), 0.0),
                    'unit_price': self._to_float(item.get('unit_price'), 0.0),
                    'total': self._to_float(item.get('total'), 0.0),
                    'taxable': bool(item.get('taxable', False)),
                }
                for item in items
                if isinstance(item, dict) and str(item.get('description') or "").strip()
            ],
        }, warnings, errors

    def extract_supplier_invoice_data(self, file_path, supplier_name=""):
        ext = os.path.splitext(file_path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            return self._extract_excel_invoice_data(file_path)
        if ext == '.pdf':
            return self._extract_pdf_invoice_data(file_path, supplier_name=supplier_name)
        return None, ["Unsupported file type."], ["Only PDF and Excel invoice files are supported."]

    def _build_supplier_inventory_item_payload(self, record, invoice_item):
        item_name = str(invoice_item.get('description') or "").strip()
        quantity = max(self._to_float(invoice_item.get('quantity'), 0.0), 0.0)
        unit_price = max(self._to_float(invoice_item.get('unit_price'), 0.0), 0.0)
        line_total = self._to_float(invoice_item.get('total'), 0.0)
        if line_total <= 0 and quantity > 0 and unit_price > 0:
            line_total = quantity * unit_price
        if line_total < 0:
            line_total = 0.0

        invoice_number = str(record.get('invoice_number') or record.get('document_id') or "").strip()
        file_name = str(record.get('file_name') or "").strip()
        supplier_name = self._normalize_supplier_name(record.get('supplier_name')) or "Unknown Supplier"
        document_date = self._normalize_date_value(record.get('document_date')) or datetime.now().strftime("%Y-%m-%d")
        description_bits = []
        if invoice_number:
            description_bits.append(f"Invoice {invoice_number}")
        if file_name:
            description_bits.append(f"File {file_name}")
        payload = {
            'name': item_name,
            'category': 'Supplier Import',
            'description': " | ".join(description_bits) if description_bits else "Imported from supplier invoice",
            'total_cost': round(line_total, 2),
            'selling_price': 0.0,
            'quantity': quantity,
            'min_stock': 0,
            'supplier': supplier_name,
            'batch_number': None,
            'expiry_date': None,
            'created_date': document_date,
            'invoice_date': document_date,
            'addition_source': f"Supplier Invoice {invoice_number}".strip() if invoice_number else "Supplier Invoice Import",
            'is_storage_item': False,
            'source_document_id': str(record.get('document_id') or "").strip() or None,
            'source_invoice_number': invoice_number or None,
            'source_file_name': file_name or None,
            'manufacture_date': None,
            'package_type': "",
            'brand': "",
            'client_name': supplier_name,
        }
        return payload

    def sync_supplier_invoice_to_inventory(self, record):
        if not isinstance(record, dict):
            return False, "Supplier invoice record is invalid.", []

        items = [item for item in (record.get('items') or []) if isinstance(item, dict) and str(item.get('description') or "").strip()]
        if not items:
            record['inventory_item_ids'] = []
            return True, "No invoice items were found to add to inventory.", []

        linked_ids = [str(item_id).strip() for item_id in (record.get('inventory_item_ids') or []) if str(item_id).strip()]
        synced_ids = []
        existing_map = {item.item_id: item for item in self.items}

        for index, invoice_item in enumerate(items):
            payload = self._build_supplier_inventory_item_payload(record, invoice_item)
            if not payload.get('name') or payload.get('quantity', 0) <= 0:
                continue

            existing_id = linked_ids[index] if index < len(linked_ids) else ""
            existing_item = existing_map.get(existing_id) if existing_id else None

            if existing_item:
                existing_item.name = payload['name']
                existing_item.category = payload['category']
                existing_item.description = payload['description']
                existing_item.total_cost = payload['total_cost']
                existing_item.quantity = payload['quantity']
                existing_item.min_stock = payload['min_stock']
                existing_item.supplier = payload['supplier']
                existing_item.batch_number = payload['batch_number']
                existing_item.expiry_date = payload['expiry_date']
                existing_item.created_date = payload['created_date']
                existing_item.invoice_date = payload['invoice_date']
                existing_item.addition_source = payload['addition_source']
                existing_item.is_storage_item = payload['is_storage_item']
                existing_item.source_document_id = payload['source_document_id']
                existing_item.source_invoice_number = payload['source_invoice_number']
                existing_item.source_file_name = payload['source_file_name']
                existing_item.manufacture_date = payload.get('manufacture_date')
                existing_item.package_type = payload.get('package_type', '')
                existing_item.brand = payload.get('brand', '')
                existing_item.client_name = payload.get('client_name', payload.get('supplier', ''))
                existing_item.last_updated = datetime.now().strftime("%Y-%m-%d")
                synced_ids.append(existing_item.item_id)
                continue

            new_item = InventoryItem(
                self.generate_item_id(),
                payload['name'],
                payload['category'],
                payload['description'],
                payload['total_cost'],
                payload['selling_price'],
                payload['quantity'],
                payload['min_stock'],
                payload['supplier'],
                payload['batch_number'],
                payload['expiry_date'],
                created_date=payload['created_date'],
                addition_source=payload['addition_source'],
                invoice_date=payload['invoice_date'],
                is_storage_item=payload['is_storage_item'],
                source_document_id=payload['source_document_id'],
                source_invoice_number=payload['source_invoice_number'],
                source_file_name=payload['source_file_name'],
                manufacture_date=payload.get('manufacture_date'),
                package_type=payload.get('package_type', ''),
                brand=payload.get('brand', ''),
                client_name=payload.get('client_name', payload.get('supplier', '')),
            )
            self.items.append(new_item)
            existing_map[new_item.item_id] = new_item
            synced_ids.append(new_item.item_id)

        extra_ids = linked_ids[len(synced_ids):]
        if extra_ids:
            self.items = [item for item in self.items if item.item_id not in extra_ids]

        record['inventory_item_ids'] = synced_ids
        if self.save_inventory():
            return True, f"{len(synced_ids)} invoice item(s) synced to inventory.", synced_ids
        return False, "Invoice items could not be synced to inventory.", synced_ids

    def save_supplier_invoice_import(self, supplier_id, file_path):
        supplier = self.get_supplier_record(supplier_id)
        if not supplier:
            return False, "Supplier not found.", None
        if not file_path or not os.path.exists(file_path):
            return False, "Selected file was not found.", None

        document_id = self.generate_supplier_document_id()
        extension = os.path.splitext(file_path)[1].lower()
        safe_name = self._safe_filename(os.path.basename(file_path))
        supplier_folder = os.path.join(self.supplier_documents_folder, supplier.get('supplier_id'))
        os.makedirs(supplier_folder, exist_ok=True)
        stored_path = os.path.join(supplier_folder, f"{document_id}_{safe_name}")

        try:
            shutil.copy2(file_path, stored_path)
        except Exception as exc:
            return False, f"Failed to copy file: {exc}", None

        extracted, warnings, errors = self.extract_supplier_invoice_data(stored_path, supplier_name=supplier.get('name', ''))
        if extracted is None:
            extracted = {
                'invoice_number': os.path.splitext(os.path.basename(file_path))[0],
                'document_date': datetime.now().strftime("%Y-%m-%d"),
                'due_date': "",
                'subtotal': 0.0,
                'tax_amount': 0.0,
                'grand_total': 0.0,
                'currency': 'AED',
                'status': 'failed',
                'notes': 'Uploaded, but extraction failed.',
                'items': [],
            }

        record = {
            'document_id': document_id,
            'supplier_id': supplier.get('supplier_id'),
            'supplier_name': supplier.get('name'),
            'file_name': os.path.basename(file_path),
            'stored_path': stored_path,
            'file_type': extension.lstrip('.'),
            'invoice_number': extracted.get('invoice_number', ''),
            'document_date': extracted.get('document_date', datetime.now().strftime("%Y-%m-%d")),
            'upload_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'due_date': extracted.get('due_date', ''),
            'subtotal': self._to_float(extracted.get('subtotal'), 0.0),
            'tax_amount': self._to_float(extracted.get('tax_amount'), 0.0),
            'grand_total': self._to_float(extracted.get('grand_total'), 0.0),
            'currency': extracted.get('currency', 'AED'),
            'status': extracted.get('status', 'parsed'),
            'notes': extracted.get('notes', ''),
            'warnings': warnings,
            'errors': errors,
            'inventory_item_ids': [],
            'items': extracted.get('items', []),
        }

        imports = self.load_supplier_invoice_imports()
        imports.append(record)
        if not self.save_supplier_invoice_imports(imports):
            return False, "Invoice was extracted but could not be saved.", None

        synced, sync_message, synced_ids = self.sync_supplier_invoice_to_inventory(record)
        if synced:
            record['inventory_item_ids'] = synced_ids
            imports = self.load_supplier_invoice_imports()
            for entry in imports:
                if entry.get('document_id') == record.get('document_id'):
                    entry['inventory_item_ids'] = synced_ids
                    break
            self.save_supplier_invoice_imports(imports)
        else:
            warnings.append(sync_message)
            record['warnings'] = warnings

        message = f"Invoice uploaded for {supplier.get('name')}."
        if record.get('status') == 'partial':
            message += "\nSome fields were partially detected."
        elif record.get('status') == 'failed':
            message += "\nThe file was saved, but extraction failed."
        if synced:
            message += f"\n{sync_message}"
        else:
            message += f"\nInventory sync warning: {sync_message}"
        return True, message, record

    def delete_supplier_invoice_import(self, document_id):
        imports = self.load_supplier_invoice_imports()
        target = None
        kept = []
        for entry in imports:
            if entry.get('document_id') == document_id:
                target = entry
            else:
                kept.append(entry)

        if not target:
            return False, "Invoice upload not found."

        stored_path = target.get('stored_path')
        if stored_path and os.path.exists(stored_path):
            try:
                os.remove(stored_path)
            except Exception:
                pass

        if self.save_supplier_invoice_imports(kept):
            return True, "Uploaded invoice removed successfully."
        return False, "Failed to remove uploaded invoice."

    def get_supplier_inbound_rows(self, start_date=None, end_date=None, supplier_id=None):
        rows = []
        for record in self.load_supplier_invoice_imports():
            if supplier_id and record.get('supplier_id') != supplier_id:
                continue
            document_date = record.get('document_date') or ""
            if start_date and document_date and document_date < start_date:
                continue
            if end_date and document_date and document_date > end_date:
                continue

            items = record.get('items') or []
            if not items:
                rows.append({
                    'supplier_id': record.get('supplier_id', ''),
                    'supplier_name': record.get('supplier_name', ''),
                    'document_id': record.get('document_id', ''),
                    'document_date': document_date,
                    'invoice_number': record.get('invoice_number', ''),
                    'file_name': record.get('file_name', ''),
                    'item_description': '',
                    'quantity': 0.0,
                    'unit_price': 0.0,
                    'line_total': 0.0,
                    'grand_total': self._to_float(record.get('grand_total'), 0.0),
                    'status': record.get('status', ''),
                })
                continue

            for item in items:
                rows.append({
                    'supplier_id': record.get('supplier_id', ''),
                    'supplier_name': record.get('supplier_name', ''),
                    'document_id': record.get('document_id', ''),
                    'document_date': document_date,
                    'invoice_number': record.get('invoice_number', ''),
                    'file_name': record.get('file_name', ''),
                    'item_description': str(item.get('description') or "").strip(),
                    'quantity': self._to_float(item.get('quantity'), 0.0),
                    'unit_price': self._to_float(item.get('unit_price'), 0.0),
                    'line_total': self._to_float(item.get('total'), 0.0),
                    'grand_total': self._to_float(record.get('grand_total'), 0.0),
                    'status': record.get('status', ''),
                })

        rows.sort(key=lambda row: (row.get('supplier_name', '').lower(), row.get('document_date', ''), row.get('invoice_number', ''), row.get('item_description', '').lower()))
        return rows

    def import_from_excel(self, file_path):
        """Import inventory items from an Excel file"""
        if not OPENPYXL_AVAILABLE:
            return False, "Excel library (openpyxl) not available. Please install it."

        def norm(s):
            """Enhanced normalization: lowercase, strip, remove special chars, collapse whitespace"""
            if s is None:
                return ""
            import re
            s = str(s).strip().lower()
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

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            
            # Find the header row (look in first 10 rows for any keyword)
            header_row_idx = 1
            headers = {}
            raw_headers = []
            keywords = ['item', 'name', 'qty', 'quantity', 'price', 'cost', 'description']
            
            for r in range(1, 11):
                row_vals = [norm(cell.value) for cell in ws[r]]
                if any(any(kw in val for kw in keywords) for val in row_vals if val):
                    header_row_idx = r
                    for cell in ws[r]:
                        if cell.value:
                            val = str(cell.value).strip()
                            raw_headers.append(val)
                            headers[norm(val)] = cell.column - 1  # Use enhanced norm for header keys
                    break
            
            if not headers:
                return False, "Could not find any header row in the first 10 rows of the Excel sheet."
            
            # Column mapping (normalized to lowercase)
            # Priority: 'name' gets 'description' if 'item name' isn't found
            mapping_configs = [
                ('name', ['item name', 'name', 'product name', 'item', 'product', 'description']),
                ('category', ['category', 'group', 'type']),
                ('description', ['details', 'notes', 'info']), # removed 'description' from here to avoid conflict
                ('quantity', ['quantity', 'qty', 'stock', 'balance', 'count', 'amount']),
                ('cost_per_item', ['cost', 'cost per item', 'unit cost', 'purchase price', 'buying price']),
                ('selling_price', ['price', 'selling price', 'unit price', 'mrp', 'sale price', 'actual price', 'actual unit price', 'actual rate', 'rate']),
                ('min_stock', ['min stock', 'minimum stock', 'reorder level', 'alert level']),
                ('supplier', ['supplier', 'vendor', 'manufacturer', 'source']),
                ('client_name', ['client', 'client name', 'customer', 'customer name']),
                ('batch_number', ['batch', 'batch number', 'lot', 'lot number', 'batch no']),
                ('manufacture_date', ['manufacture date', 'manufacturing date', 'mfg date', 'mfd date', 'manf date']),
                ('expiry_date', ['expiry', 'expiry date', 'exp date', 'valid until']),
                ('package_type', ['package type', 'packaging', 'pack type', 'package', 'storage type']),
                ('brand', ['brand', 'brand name', 'product brand']),
                ('is_storage_item', ['storage', 'is storage', 'stored goods', 'stored', 'consignment'])
            ]
            
            found_mapping = {}
            used_cols = set()
            potential_mismatches = []
            
            for key, aliases in mapping_configs:
                # First try exact matches
                matched = False
                for alias in aliases:
                    norm_alias = norm(alias)
                    if norm_alias in headers and headers[norm_alias] not in used_cols:
                        found_mapping[key] = headers[norm_alias]
                        used_cols.add(headers[norm_alias])
                        matched = True
                        break
                if matched:
                    continue
                # Then try fuzzy matches
                best_ratio = 0
                best_idx = None
                best_header = None
                for header_text, header_idx in headers.items():
                    if header_idx in used_cols:
                        continue
                    for alias in aliases:
                        norm_alias = norm(alias)
                        is_match, ratio = fuzzy_match(header_text, norm_alias, threshold=0.7)
                        if is_match and ratio > best_ratio:
                            best_ratio = ratio
                            best_idx = header_idx
                            best_header = header_text
                if best_idx is not None:
                    found_mapping[key] = best_idx
                    used_cols.add(best_idx)
                    if best_ratio < 0.9:
                        potential_mismatches.append(f"Potential mismatch: mapped header '{best_header}' to '{key}' (similarity: {best_ratio:.2f})")
            
            # Fallback for description if it wasn't used for name and is available
            if 'description' not in found_mapping:
                for h in headers:
                    if 'description' in h and headers[h] not in used_cols:
                        found_mapping['description'] = headers[h]
                        used_cols.add(headers[h])
                        break

            if 'name' not in found_mapping:
                headers_str = ", ".join(raw_headers)
                return False, f"Could not find a column for 'Item Name'.\n\nFound columns: {headers_str}\n\nPlease ensure your Excel has a column named 'Item Name', 'Name', or 'Description'."

            items_added = 0
            errors = []
            for row_idx, row in enumerate(ws.iter_rows(min_row=header_row_idx + 1, values_only=True), start=header_row_idx + 1):
                try:
                    name = row[found_mapping['name']]
                    if not name:
                        continue # Skip empty rows
                    
                    # Extract values with defaults
                    category = row[found_mapping['category']] if 'category' in found_mapping else "Other"
                    description = row[found_mapping['description']] if 'description' in found_mapping else ""
                    
                    quantity = self._to_float(
                        row[found_mapping['quantity']] if 'quantity' in found_mapping else None,
                        0
                    )
                    cost_per_item = self._to_float(
                        row[found_mapping['cost_per_item']] if 'cost_per_item' in found_mapping else None,
                        0
                    )
                    selling_price = self._to_float(
                        row[found_mapping['selling_price']] if 'selling_price' in found_mapping else None,
                        0
                    )
                    min_stock = self._to_float(
                        row[found_mapping['min_stock']] if 'min_stock' in found_mapping else None,
                        0
                    )
                    
                    supplier = row[found_mapping['supplier']] if 'supplier' in found_mapping and row[found_mapping['supplier']] is not None else "Unknown"
                    client_name = row[found_mapping['client_name']] if 'client_name' in found_mapping and row[found_mapping['client_name']] is not None else supplier
                    batch_number = str(row[found_mapping['batch_number']]) if 'batch_number' in found_mapping and row[found_mapping['batch_number']] is not None else ""
                    manufacture_date = None
                    if 'manufacture_date' in found_mapping and row[found_mapping['manufacture_date']]:
                        val = row[found_mapping['manufacture_date']]
                        if isinstance(val, datetime):
                            manufacture_date = val.strftime("%Y-%m-%d")
                        else:
                            manufacture_date = str(val)
                    
                    expiry_date = None
                    if 'expiry_date' in found_mapping and row[found_mapping['expiry_date']]:
                        val = row[found_mapping['expiry_date']]
                        if isinstance(val, datetime):
                            expiry_date = val.strftime("%Y-%m-%d")
                        else:
                            expiry_date = str(val)
                    package_type = str(row[found_mapping['package_type']]).strip() if 'package_type' in found_mapping and row[found_mapping['package_type']] is not None else ""
                    brand = str(row[found_mapping['brand']]).strip() if 'brand' in found_mapping and row[found_mapping['brand']] is not None else ""
                    
                    is_storage = False
                    if 'is_storage_item' in found_mapping and row[found_mapping['is_storage_item']]:
                        val = str(row[found_mapping['is_storage_item']]).lower()
                        is_storage = val in ['yes', 'y', '1', 'true', 'stored', 'consignment']

                    # Total cost calculation
                    total_cost = cost_per_item * quantity
                    
                    # Add to inventory
                    self.add_item(
                        name=str(name),
                        category=str(category),
                        description=str(description),
                        total_cost=total_cost,
                        selling_price=selling_price,
                        quantity=int(quantity),
                        min_stock=int(min_stock),
                        supplier=str(supplier),
                        batch_number=batch_number,
                        expiry_date=expiry_date,
                        manufacture_date=manufacture_date,
                        package_type=package_type,
                        brand=brand,
                        client_name=str(client_name),
                        addition_source="Excel Import",
                        is_storage_item=is_storage
                    )
                    items_added += 1
                except Exception as e:
                    errors.append(f"Row {row_idx}: {str(e)}")
            
            report_msg = f"Successfully imported {items_added} items."
            all_notes = []
            if potential_mismatches:
                all_notes.extend(potential_mismatches)
            if errors:
                all_notes.extend(errors)
            
            if all_notes:
                report_file = os.path.join(os.path.dirname(file_path), "import_problems.txt")
                with open(report_file, 'w') as f:
                    f.write("IMPORT PROBLEM REPORT\n")
                    f.write("=====================\n\n")
                    f.write("\n".join(all_notes))
                report_msg += f"\n\nNote: {len(all_notes)} warnings/errors were found.\nA report has been saved to:\n{report_file}"
                
            return True, report_msg
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            return False, f"Error reading Excel file: {str(e)}"

    def parse_invoice_excel(self, file_path, prefer_actual_rate=True):
        """Parse an invoice-like Excel sheet into an invoice dict"""
        if not OPENPYXL_AVAILABLE:
            return None, ["Excel library (openpyxl) not available"], []

        warnings = []
        errors = []

        def norm(s):
            """Enhanced normalization: lowercase, strip, remove special chars, collapse whitespace"""
            if s is None:
                return ""
            import re
            s = str(s).strip().lower()
            # Remove special characters except alphanumeric and spaces
            s = re.sub(r'[^a-z0-9\s]', '', s)
            # Collapse multiple spaces
            s = re.sub(r'\s+', ' ', s).strip()
            return s

        def fuzzy_match(a, b, threshold=0.8):
            """Fuzzy string match using difflib SequenceMatcher"""
            import difflib
            if not a or not b:
                return False, 0.0
            ratio = difflib.SequenceMatcher(None, a, b).ratio()
            return ratio >= threshold, ratio

        def to_float(v, default=0.0):
            try:
                if v is None:
                    return default
                if isinstance(v, str):
                    vv = v.strip().replace(",", "")
                    if not vv:
                        return default
                    return float(vv)
                return float(v)
            except Exception:
                return default

        def to_date_str(v):
            if v is None:
                return None
            try:
                if isinstance(v, datetime):
                    return v.strftime("%Y-%m-%d")
                s = str(v).strip()
                if not s:
                    return None
                return s
            except Exception:
                return None

        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
        except Exception as e:
            return None, [f"Error reading Excel file: {e}"], []

        max_scan_rows = min(50, ws.max_row or 0)
        max_scan_cols = min(60, ws.max_column or 0)

        meta = {
            'client_name': None,
            'client_trn': None,
            'date': None,
            'due_date': None,
            'notes': None,
        }

        meta_keys = {
            'client_name': ['customer', 'customer name', 'client', 'client name', 'bill to', 'billed to', 'sold to'],
            'client_trn': ['trn', 'vat no', 'vat number', 'tax no', 'tax number'],
            'date': ['date', 'invoice date'],
            'due_date': ['due date', 'payment due'],
            'notes': ['notes', 'remark', 'remarks', 'comment'],
        }

        for r in range(1, max_scan_rows + 1):
            for c in range(1, max_scan_cols + 1):
                v = ws.cell(row=r, column=c).value
                if not v:
                    continue
                key_text = norm(v)
                for field, aliases in meta_keys.items():
                    if meta[field] is not None:
                        continue
                    if key_text in aliases:
                        val = ws.cell(row=r, column=c + 1).value
                        if field in ['date', 'due_date']:
                            meta[field] = to_date_str(val)
                        else:
                            meta[field] = str(val).strip() if val is not None else None

        item_header_keywords = {
            'description': ['description', 'item', 'item description', 'product', 'service', 'details'],
            'quantity': ['qty', 'quantity', 'qnty', 'count'],
            'actual_rate': ['actual rate', 'actual unit price', 'actual price'],
            'rate': ['unit price', 'rate', 'price', 'unit cost'],
            'total': ['total', 'amount', 'line total', 'net'],
            'vat': ['vat', 'tax', 'vat %', 'tax %', 'vat amount', 'tax amount'],
        }

        header_row_idx = None
        header_map = {}
        best_score = 0
        used_cols_in_best = set()

        for r in range(1, max_scan_rows + 1):
            row_text = [norm(ws.cell(row=r, column=c).value) for c in range(1, max_scan_cols + 1)]
            score = 0
            col_map = {}
            used_cols = set()
            # First, try exact matches
            for field, keys in item_header_keywords.items():
                for idx, cell_txt in enumerate(row_text, start=1):
                    if idx in used_cols:
                        continue
                    if cell_txt in keys:
                        col_map[field] = idx
                        used_cols.add(idx)
                        score += 10  # Exact match gets high score
                        break
                if field in col_map:
                    continue
                # Then try fuzzy matches
                best_ratio = 0
                best_idx = None
                for idx, cell_txt in enumerate(row_text, start=1):
                    if idx in used_cols:
                        continue
                    for key in keys:
                        is_match, ratio = fuzzy_match(cell_txt, key, threshold=0.7)
                        if is_match and ratio > best_ratio:
                            best_ratio = ratio
                            best_idx = idx
                if best_idx is not None:
                    col_map[field] = best_idx
                    used_cols.add(best_idx)
                    score += best_ratio  # Use ratio as partial score
            # Now see if this row is better than previous best
            if score > best_score:
                best_score = score
                header_row_idx = r
                header_map = col_map
                used_cols_in_best = used_cols

        if not header_row_idx or 'description' not in header_map:
            if meta.get('client_name'):
                warnings.append("Could not detect line items table; only invoice header fields were found")
            else:
                warnings.append("Could not detect invoice table headers (need a 'Description' column)")
            return None, warnings, errors

        items = []
        first_item_row = header_row_idx + 1
        last_row = ws.max_row or first_item_row
        for r in range(first_item_row, last_row + 1):
            desc = ws.cell(row=r, column=header_map['description']).value
            desc_txt = str(desc).strip() if desc is not None else ""
            if not desc_txt:
                empty_row = True
                for check_c in range(1, min(8, max_scan_cols) + 1):
                    if ws.cell(row=r, column=check_c).value:
                        empty_row = False
                        break
                if empty_row:
                    break
                continue

            qty = 1.0
            if 'quantity' in header_map:
                qty = to_float(ws.cell(row=r, column=header_map['quantity']).value, default=1.0)
                if qty == 0:
                    qty = 1.0

            unit_price = 0.0
            preferred_field = 'actual_rate' if prefer_actual_rate else 'rate'
            fallback_field = 'rate' if prefer_actual_rate else 'actual_rate'
            if preferred_field in header_map:
                unit_price = to_float(ws.cell(row=r, column=header_map[preferred_field]).value, default=0.0)
            elif fallback_field in header_map:
                unit_price = to_float(ws.cell(row=r, column=header_map[fallback_field]).value, default=0.0)

            line_total = None
            if 'total' in header_map:
                line_total = to_float(ws.cell(row=r, column=header_map['total']).value, default=None)

            if line_total is None:
                line_total = unit_price * qty
            else:
                if unit_price == 0 and qty != 0:
                    unit_price = line_total / qty

            taxable = False
            if 'vat' in header_map:
                vat_val = ws.cell(row=r, column=header_map['vat']).value
                v = norm(vat_val)
                if v and v not in ['0', '0.0', 'no', 'n', 'false']:
                    taxable = True

            items.append({
                'description': desc_txt,
                'quantity': float(qty),
                'unit_price': float(unit_price),
                'total': float(line_total),
                'taxable': taxable,
            })

        if not items:
            warnings.append("No invoice items found under the header row")
            return None, warnings, errors

        subtotal = sum(float(i.get('total', 0.0)) for i in items)
        taxable_amount = sum(float(i.get('total', 0.0)) for i in items if i.get('taxable'))
        non_taxable_amount = subtotal - taxable_amount
        tax_rate = 0
        tax_amount = taxable_amount * (tax_rate / 100.0)
        grand_total = subtotal + tax_amount

        invoice_dict = {
            'source': 'excel_import',
            'invoice_type': 'sales',
            'client_name': meta.get('client_name') or 'N/A',
            'client_trn': meta.get('client_trn') or '',
            'client_emirate': self._get_default_emirate(),
            'client_location': self._get_default_location(),
            'date': meta.get('date') or datetime.now().strftime("%Y-%m-%d"),
            'due_date': meta.get('due_date') or datetime.now().strftime("%Y-%m-%d"),
            'payment_terms': 'Cash',
            'tax_rate': tax_rate,
            'items': items,
            'costs': [],
            'subtotal': subtotal,
            'taxable_amount': taxable_amount,
            'non_taxable_amount': non_taxable_amount,
            'tax_amount': tax_amount,
            'grand_total': grand_total,
            'total_cost': 0.0,
            'profit_loss': grand_total,
            'total_paid': 0.0,
            'status': 'Not Paid',
            'notes': meta.get('notes') or 'Imported from Excel',
            'currency': 'AED',
        }

        if meta.get('client_name') is None:
            warnings.append("Client name not found; invoice will use 'N/A'")

        return invoice_dict, warnings, errors

    def get_supplier_activity_rows(self, start_date=None, end_date=None, supplier_filter=None):
        """Build supplier -> item -> sales activity rows"""
        def _norm_supplier(s):
            return (s or "").strip() or "Unknown Supplier"

        def _in_range(d):
            if not d:
                return False
            if start_date and d < start_date:
                return False
            if end_date and d > end_date:
                return False
            return True

        supplier_filter_norm = None
        if supplier_filter and supplier_filter != "All Suppliers":
            supplier_filter_norm = _norm_supplier(supplier_filter)

        items = list(self.items or [])
        sales = list(self.sales or [])

        sales_by_item = {}
        for s in sales:
            try:
                sales_by_item.setdefault(s.item_id, []).append(s)
            except Exception:
                continue

        for k, lst in sales_by_item.items():
            try:
                lst.sort(key=lambda x: getattr(x, 'sale_date', '') or '')
            except Exception:
                pass

        rows = []
        for item in items:
            supplier = _norm_supplier(getattr(item, 'supplier', None))
            if supplier_filter_norm and supplier != supplier_filter_norm:
                continue

            entry_date = (getattr(item, 'created_date', '') or '').strip()
            item_id = getattr(item, 'item_id', '')
            item_name = getattr(item, 'name', '')
            batch = getattr(item, 'batch_number', '') or ''
            current_qty = getattr(item, 'quantity', 0)
            storage_flag = bool(getattr(item, 'is_storage_item', False))

            item_sales = sales_by_item.get(item_id, [])
            any_sale_in_range = False
            for s in item_sales:
                sd = (getattr(s, 'sale_date', '') or '').strip()
                if start_date or end_date:
                    if not _in_range(sd):
                        continue
                any_sale_in_range = True
                rows.append({
                    'supplier': supplier,
                    'item_id': item_id,
                    'item_name': item_name,
                    'batch_number': batch,
                    'is_storage_item': storage_flag,
                    'entry_date': entry_date,
                    'sold_date': sd,
                    'sold_to': getattr(s, 'customer_name', 'Unknown') or 'Unknown',
                    'sold_qty': getattr(s, 'quantity', 0),
                    'invoice_id': getattr(s, 'invoice_id', '') or '',
                    'current_qty': current_qty,
                })

            include_unsold = True
            if start_date or end_date:
                include_unsold = _in_range(entry_date) or any_sale_in_range

            if include_unsold and not any_sale_in_range:
                rows.append({
                    'supplier': supplier,
                    'item_id': item_id,
                    'item_name': item_name,
                    'batch_number': batch,
                    'is_storage_item': storage_flag,
                    'entry_date': entry_date,
                    'sold_date': '',
                    'sold_to': '',
                    'sold_qty': 0,
                    'invoice_id': '',
                    'current_qty': current_qty,
                })

        try:
            rows.sort(key=lambda r: (r.get('supplier', ''), r.get('item_name', ''), r.get('sold_date', '')))
        except Exception:
            pass

        return rows

    def _get_default_emirate(self):
        """Get default emirate from settings, fallback to empty"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f) or {}
                em = settings.get('default_emirate')
                if isinstance(em, str) and em.strip():
                    return em.strip()
        except Exception:
            pass
        return ""
    
    def _get_default_location(self):
        """Get default location (city, country), fallback to derived from default emirate or empty"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f) or {}
                loc = settings.get('default_location')
                if isinstance(loc, str) and loc.strip():
                    return loc.strip()
        except Exception:
            pass
        em = (self._get_default_emirate() or "").strip()
        if not em:
            return ""
        uae_emirates = {
            'Dubai',
            'Abu Dhabi',
            'Sharjah',
            'Ajman',
            'Umm Al Quwain',
            'Ras Al Khaimah',
            'Fujairah',
        }
        if ',' not in em and em in uae_emirates:
            return f"{em}, United Arab Emirates"
        return em
    
    def generate_invoice_id(self):
        """Generate inventory-only invoice ID using app_settings.json"""
        prefix = "HPMI"
        number = int(datetime.now().strftime("%y%m%d%H%M%S"))
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f) or {}
                prefix = settings.get('inventory_invoice_prefix', prefix)
                last = int(settings.get('last_inventory_invoice_number', number))
                number = last + 1
                settings['last_inventory_invoice_number'] = number
                if 'inventory_invoice_prefix' not in settings:
                    settings['inventory_invoice_prefix'] = prefix
                with open(self.settings_file, 'w') as f:
                    json.dump(settings, f, indent=2)
        except Exception:
            pass
        return f"{prefix}{number}"
    
    def _append_invoice(self, invoice_data):
        """Append an invoice to invoices_data.json directly (decoupled from InvoiceManager)"""
        try:
            data = []
            if os.path.exists(self.invoices_file):
                with open(self.invoices_file, 'r') as f:
                    data = json.load(f) or []
            data.append(invoice_data)
            with open(self.invoices_file, 'w') as f:
                json.dump(data, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving invoice: {e}")
            return False
    
    def create_sales_invoice(self, sale_record, is_paid=True):
        """Create and store an invoice for a sale inside InventoryManager only"""
        try:
            invoice_id = self.generate_invoice_id()
            grand_total = sale_record.total_amount + getattr(sale_record, 'vat_amount', 0.0)
            
            invoice_data = {
                'invoice_id': invoice_id,
                'source': 'inventory',
                'invoice_type': 'sales',
                'client_name': sale_record.customer_name,
                'client_trn': '',
                'client_emirate': self._get_default_emirate(),
                'client_location': self._get_default_location(),
                'date': datetime.now().strftime("%Y-%m-%d"),
                'due_date': datetime.now().strftime("%Y-%m-%d"),
                'payment_terms': 'Cash',
                'tax_rate': 5 if getattr(sale_record, 'vat_applied', False) else 0,
                'items': [
                    {
                        'description': sale_record.item_name,
                        'quantity': sale_record.quantity,
                        'unit_price': sale_record.selling_price,
                        'total': sale_record.total_amount,
                        'taxable': getattr(sale_record, 'vat_applied', False)
                    },
                    {
                        'description': f"Promotion - {sale_record.item_name}",
                        'quantity': getattr(sale_record, 'promotion_quantity', 0),
                        'unit_price': 0.0,
                        'total': 0.0,
                        'taxable': False
                    }
                ],
                'inventory_links': [
                    {
                        'item_id': sale_record.item_id,
                        'item_name': sale_record.item_name,
                        'quantity': sale_record.quantity,
                        'promotion_quantity': getattr(sale_record, 'promotion_quantity', 0),
                        'cost_per_item': sale_record.cost_per_item
                    }
                ],
                'costs': [
                    {
                        'description': 'Inventory cost (COGS)',
                        'amount': round(float(sale_record.total_cost), 2),
                        'account': '400'
                    }
                ],
                'subtotal': sale_record.total_amount,
                'taxable_amount': sale_record.total_amount if getattr(sale_record, 'vat_applied', False) else 0,
                'non_taxable_amount': 0,
                'tax_amount': getattr(sale_record, 'vat_amount', 0.0),
                'grand_total': grand_total,
                'total_cost': sale_record.total_cost,
                'profit_loss': sale_record.total_profit,
                'total_paid': grand_total if is_paid else 0.0,
                'status': 'Paid' if is_paid else 'Not Paid',
                'notes': f"Inventory sale: {sale_record.item_name}\nPromotion quantity: {getattr(sale_record, 'promotion_quantity', 0)}\nCost per item: AED {sale_record.cost_per_item:.2f}",
                'currency': 'AED'
            }
            ok = self._append_invoice(invoice_data)
            if ok:
                sale_record.invoice_id = invoice_id
                self.save_sales()
                if is_paid:
                    try:
                        # Record payment into BalanceManager to keep finances connected
                        from balance_manager import BalanceManager
                        bm = BalanceManager(self.data_folder)
                        _ = bm.process_invoice_payment(invoice_data, invoice_data['grand_total'], 'Cash')
                    except Exception:
                        pass
                return invoice_id
        except Exception as e:
            print(f"Error creating inventory sales invoice: {e}")
        return None

    def create_sales_invoice_multi(self, sale_records, customer_name, is_paid=True, delivery_fee=0.0, is_delivery_only=False):
        """Create and store a single invoice for multiple sales"""
        if not sale_records:
            return None
        try:
            invoice_id = self.generate_invoice_id()
            
            # Aggregate totals
            total_amount = sum(s.total_amount for s in sale_records)
            total_cost = sum(s.total_cost for s in sale_records)
            total_profit = sum(s.total_profit for s in sale_records)
            vat_amount = sum(getattr(s, 'vat_amount', 0.0) for s in sale_records)
            
            # Adjust totals for delivery only
            if is_delivery_only:
                total_amount = delivery_fee
                vat_amount = 0.0 # Delivery fees usually don't have VAT in this simple logic unless specified
                # Grand total = sum of all item totals (should be exactly delivery_fee)
                grand_total = 0.0
                for s in sale_records:
                    pass  # items will be 0, then add delivery fee line item
                # (the actual grand total will be calculated after building items list, below)
                total_profit = total_amount - total_cost
            else:
                grand_total = total_amount + vat_amount + delivery_fee
                total_profit += delivery_fee

            # Aggregate items
            items = []
            for s in sale_records:
                items.append({
                    'description': s.item_name,
                    'quantity': s.quantity,
                    'unit_price': 0.0 if is_delivery_only else s.selling_price,
                    'total': 0.0 if is_delivery_only else s.total_amount,
                    'taxable': False if is_delivery_only else getattr(s, 'vat_applied', False)
                })
                if getattr(s, 'promotion_quantity', 0) > 0:
                    items.append({
                        'description': f"Promotion - {s.item_name}",
                        'quantity': s.promotion_quantity,
                        'unit_price': 0.0,
                        'total': 0.0,
                        'taxable': False
                    })
            
            # Add delivery fee as a line item if present
            if delivery_fee > 0:
                items.append({
                    'description': 'Delivery Fee',
                    'quantity': 1,
                    'unit_price': delivery_fee,
                    'total': delivery_fee,
                    'taxable': False
                })
            
            # Now calculate final grand_total based on items for delivery-only
            if is_delivery_only:
                grand_total = sum(float(i.get('total', 0)) for i in items)

            # Aggregate inventory links
            inventory_links = []
            for s in sale_records:
                inventory_links.append({
                    'item_id': s.item_id,
                    'item_name': s.item_name,
                    'quantity': s.quantity,
                    'promotion_quantity': getattr(s, 'promotion_quantity', 0),
                    'cost_per_item': s.cost_per_item
                })
            
            invoice_data = {
                'invoice_id': invoice_id,
                'source': 'inventory',
                'invoice_type': 'delivery' if is_delivery_only else 'sales',
                'client_name': customer_name,
                'client_trn': '',
                'client_emirate': self._get_default_emirate(),
                'client_location': self._get_default_location(),
                'date': datetime.now().strftime("%Y-%m-%d"),
                'due_date': datetime.now().strftime("%Y-%m-%d"),
                'payment_terms': 'Cash',
                'tax_rate': 0 if is_delivery_only else (5 if any(getattr(s, 'vat_applied', False) for s in sale_records) else 0),
                'items': items,
                'inventory_links': inventory_links,
                'costs': [
                    {
                        'description': 'Inventory cost (COGS)',
                        'amount': round(float(total_cost), 2),
                        'account': '400'
                    }
                ],
                'subtotal': total_amount,
                'taxable_amount': 0 if is_delivery_only else sum(s.total_amount for s in sale_records if getattr(s, 'vat_applied', False)),
                'non_taxable_amount': total_amount if is_delivery_only else sum(s.total_amount for s in sale_records if not getattr(s, 'vat_applied', False)),
                'tax_amount': vat_amount,
                'grand_total': grand_total,
                'total_cost': total_cost,
                'profit_loss': total_profit,
                'total_paid': grand_total if is_paid else 0.0,
                'status': 'Paid' if is_paid else 'Not Paid',
                'notes': f"{'Delivery only' if is_delivery_only else 'Inventory sale'}: {len(sale_records)} items",
                'currency': 'AED'
            }
            
            ok = self._append_invoice(invoice_data)
            if ok:
                for s in sale_records:
                    s.invoice_id = invoice_id
                self.save_sales()
                if is_paid:
                    try:
                        # Record payment into BalanceManager
                        from balance_manager import BalanceManager
                        bm = BalanceManager(self.data_folder)
                        _ = bm.process_invoice_payment(invoice_data, invoice_data['grand_total'], 'Cash')
                    except Exception:
                        pass
                return invoice_id
        except Exception as e:
            print(f"Error creating multi-item invoice: {e}")
        return None
    
    def load_delivery_notes(self):
        """Load delivery notes from JSON file"""
        if os.path.exists(self.delivery_notes_file):
            try:
                with open(self.delivery_notes_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading delivery notes: {e}")
        return []

    def save_delivery_note(self, note_data):
        """Save a new delivery note"""
        try:
            notes = self.load_delivery_notes()
            notes.append(note_data)
            with open(self.delivery_notes_file, 'w') as f:
                json.dump(notes, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving delivery note: {e}")
            return False

    def load_inventory(self):
        """Load inventory from JSON file"""
        if os.path.exists(self.inventory_file):
            try:
                with open(self.inventory_file, 'r') as f:
                    data = json.load(f)
                    items = []
                    for item_data in data:
                        try:
                            # Use from_dict method for compatibility
                            item = InventoryItem.from_dict(item_data)
                            items.append(item)
                        except Exception as e:
                            print(f"Error loading item {item_data.get('item_id', 'unknown')}: {e}")
                            # Try to create a basic item with available data
                            try:
                                item = InventoryItem(
                                    item_id=item_data.get('item_id', 'UNKNOWN'),
                                    name=item_data.get('name', 'Unknown Item'),
                                    category=item_data.get('category', 'Other'),
                                    description=item_data.get('description', ''),
                                    total_cost=item_data.get('total_cost', 0),
                                    selling_price=item_data.get('selling_price', 0),
                                    quantity=item_data.get('quantity', 0),
                                    min_stock=item_data.get('min_stock', 0),
                                    supplier=item_data.get('supplier', 'Unknown')
                                )
                                items.append(item)
                            except Exception as e2:
                                print(f"Failed to create basic item: {e2}")
                    return items
            except Exception as e:
                print(f"Error loading inventory file: {e}")
                import traceback
                traceback.print_exc()
        return []
    
    def load_sales(self):
        """Load sales records from JSON file"""
        if os.path.exists(self.sales_file):
            try:
                with open(self.sales_file, 'r') as f:
                    data = json.load(f)
                    sales = []
                    for sale_data in data:
                        try:
                            # Use from_dict method for compatibility
                            sale = SaleRecord.from_dict(sale_data)
                            sales.append(sale)
                        except Exception as e:
                            print(f"Error loading sale {sale_data.get('sale_id', 'unknown')}: {e}")
                            # Try to create a basic sale record
                            try:
                                sale = SaleRecord(
                                    sale_id=sale_data.get('sale_id', 'UNKNOWN'),
                                    item_id=sale_data.get('item_id', 'UNKNOWN'),
                                    item_name=sale_data.get('item_name', 'Unknown Item'),
                                    quantity=sale_data.get('quantity', 0),
                                    selling_price=sale_data.get('selling_price', 0),
                                    cost_per_item=sale_data.get('cost_per_item', 0),
                                    total_amount=sale_data.get('total_amount', 0),
                                    total_cost=sale_data.get('total_cost', 0),
                                    total_profit=sale_data.get('total_profit', 0),
                                    customer_name=sale_data.get('customer_name', 'Unknown'),
                                    sale_date=sale_data.get('sale_date', datetime.now().strftime("%Y-%m-%d")),
                                    invoice_id=sale_data.get('invoice_id')
                                )
                                sales.append(sale)
                            except Exception as e2:
                                print(f"Failed to create basic sale: {e2}")
                    return sales
            except Exception as e:
                print(f"Error loading sales file: {e}")
        return []

    def load_returns(self):
        if os.path.exists(self.returns_file):
            try:
                with open(self.returns_file, 'r') as f:
                    data = json.load(f) or []
                rows = []
                for row in data:
                    try:
                        rows.append(ReturnRecord.from_dict(row))
                    except Exception:
                        pass
                return rows
            except Exception:
                return []
        return []

    def save_returns(self):
        try:
            with open(self.returns_file, 'w') as f:
                json.dump([r.to_dict() for r in (self.returns or [])], f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving returns: {e}")
            return False

    def load_credit_notes(self):
        if os.path.exists(self.credit_notes_file):
            try:
                with open(self.credit_notes_file, 'r') as f:
                    data = json.load(f) or []
                return [row for row in data if isinstance(row, dict)]
            except Exception:
                return []
        return []

    def save_credit_notes(self):
        try:
            with open(self.credit_notes_file, 'w') as f:
                json.dump(list(self.credit_notes or []), f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving credit notes: {e}")
            return False

    def append_audit_log(self, event_type, details):
        try:
            entry = {
                "ts": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "event": str(event_type or "").strip(),
                "user": getpass.getuser(),
                "details": details if isinstance(details, dict) else {"value": str(details)},
            }
            with open(self.audit_log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(entry, ensure_ascii=False) + "\n")
            return True
        except Exception:
            return False
    
    def load_categories(self):
        """Load categories from JSON file"""
        if os.path.exists(self.categories_file):
            try:
                with open(self.categories_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading categories: {e}")
        return ["Medicines", "Medical Equipment", "Supplements", "Personal Care", "First Aid", "Other"]
    
    def load_cost_types(self):
        """Load cost types from JSON file"""
        if os.path.exists(self.cost_types_file):
            try:
                with open(self.cost_types_file, 'r') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading cost types: {e}")
        return ["Shipping", "Storage", "Handling", "Import Duty", "Insurance", "Testing", "Packaging", "Other"]
    
    def save_inventory(self):
        """Save inventory to JSON file"""
        try:
            with open(self.inventory_file, 'w') as f:
                json.dump([item.to_dict() for item in self.items], f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving inventory: {e}")
            return False
    
    def save_sales(self):
        """Save sales records to JSON file"""
        try:
            with open(self.sales_file, 'w') as f:
                json.dump([sale.to_dict() for sale in self.sales], f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving sales records: {e}")
            return False
    
    def save_categories(self):
        """Save categories to JSON file"""
        try:
            with open(self.categories_file, 'w') as f:
                json.dump(self.categories, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving categories: {e}")
            return False
    
    def save_cost_types(self):
        """Save cost types to JSON file"""
        try:
            with open(self.cost_types_file, 'w') as f:
                json.dump(self.cost_types, f, indent=2)
            return True
        except Exception as e:
            print(f"Error saving cost types: {e}")
            return False
    
    def generate_item_id(self):
        """Generate unique item ID without skipping deleted numbers - ITEM001, ITEM002, etc."""
        # Get all existing item IDs
        existing_ids = []
        for item in self.items:
            if item.item_id.startswith('ITEM') and item.item_id[4:].isdigit():
                existing_ids.append(int(item.item_id[4:]))
        
        if not existing_ids:
            return "ITEM001"
        
        # Find the next available number
        max_id = max(existing_ids)
        new_id = max_id + 1
        
        return f"ITEM{new_id:03d}"
    
    def generate_sale_id(self):
        """Generate unique sale ID"""
        existing_ids = [sale.sale_id for sale in self.sales]
        new_id = 1
        while f"SALE{new_id:03d}" in existing_ids:
            new_id += 1
        return f"SALE{new_id:03d}"

    def generate_return_id(self):
        existing = []
        for r in (self.returns or []):
            rid = str(getattr(r, "return_id", "") or "")
            if rid.startswith("RET") and rid[3:].isdigit():
                existing.append(int(rid[3:]))
        nxt = (max(existing) + 1) if existing else 1
        return f"RET{nxt:05d}"

    def generate_credit_note_id(self):
        prefix = "CN"
        number = int(datetime.now().strftime("%y%m%d%H%M%S"))
        try:
            settings = {}
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    settings = json.load(f) or {}
            prefix = settings.get('credit_note_prefix', prefix)
            last = int(settings.get('last_credit_note_number', number))
            number = last + 1
            settings['last_credit_note_number'] = number
            if 'credit_note_prefix' not in settings:
                settings['credit_note_prefix'] = prefix
            with open(self.settings_file, 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception:
            pass
        return f"{prefix}{number}"

    def _load_settings_dict(self):
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r') as f:
                    return json.load(f) or {}
        except Exception:
            pass
        return {}

    def _save_settings_dict(self, settings):
        try:
            with open(self.settings_file, 'w') as f:
                json.dump(settings or {}, f, indent=2)
            return True
        except Exception:
            return False

    def _hash_pin(self, pin_value):
        return hashlib.sha256(str(pin_value or "").encode("utf-8")).hexdigest()

    def require_returns_permission(self, parent):
        settings = self._load_settings_dict()
        required = settings.get("returns_permission_required", True)
        if not required:
            return True
        pin_hash = (settings.get("returns_pin_hash") or "").strip()
        if not pin_hash:
            dlg = tk.Toplevel(parent)
            dlg.title("Set Returns PIN")
            dlg.geometry("420x240")
            dlg.transient(parent)
            dlg.grab_set()
            frm = ttk.Frame(dlg, padding=12)
            frm.pack(fill="both", expand=True)
            ttk.Label(frm, text="Create a PIN to authorize returns & credit notes.").pack(anchor="w", pady=(0, 8))
            ttk.Label(frm, text="New PIN").pack(anchor="w")
            pin1 = tk.StringVar()
            ttk.Entry(frm, textvariable=pin1, show="*").pack(fill="x", pady=(0, 8))
            ttk.Label(frm, text="Confirm PIN").pack(anchor="w")
            pin2 = tk.StringVar()
            ttk.Entry(frm, textvariable=pin2, show="*").pack(fill="x", pady=(0, 12))
            result = {"ok": False}
            def save():
                p1 = pin1.get().strip()
                p2 = pin2.get().strip()
                if not p1 or not p2:
                    messagebox.showwarning("Warning", "PIN cannot be empty")
                    return
                if p1 != p2:
                    messagebox.showwarning("Warning", "PINs do not match")
                    return
                settings["returns_pin_hash"] = self._hash_pin(p1)
                settings["returns_permission_required"] = True
                if not self._save_settings_dict(settings):
                    messagebox.showerror("Error", "Failed to save settings")
                    return
                result["ok"] = True
                dlg.destroy()
            ttk.Button(frm, text="Save PIN", command=save).pack(side="left")
            ttk.Button(frm, text="Cancel", command=dlg.destroy).pack(side="left", padx=6)
            parent.wait_window(dlg)
            if not result["ok"]:
                return False
            pin_hash = (settings.get("returns_pin_hash") or "").strip()

        dlg = tk.Toplevel(parent)
        dlg.title("Authorization Required")
        dlg.geometry("420x200")
        dlg.transient(parent)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Enter PIN to process returns / create credit notes.").pack(anchor="w", pady=(0, 8))
        pin_var = tk.StringVar()
        ttk.Entry(frm, textvariable=pin_var, show="*").pack(fill="x", pady=(0, 12))
        result = {"ok": False}
        def verify():
            if self._hash_pin(pin_var.get().strip()) != pin_hash:
                messagebox.showerror("Error", "Invalid PIN")
                return
            result["ok"] = True
            dlg.destroy()
        ttk.Button(frm, text="Authorize", command=verify).pack(side="left")
        ttk.Button(frm, text="Cancel", command=dlg.destroy).pack(side="left", padx=6)
        parent.wait_window(dlg)
        return bool(result["ok"])

    def create_manual_credit_note(self, customer_name, items, notes="", restocking_fee=0.0, adjustment=0.0, reference=""):
        customer_name = str(customer_name or "").strip()
        if not customer_name:
            return False, "Customer name is required", None
        ok_items, msg_items, cleaned_items, subtotal, vat_total = self._prepare_manual_credit_note_items(items)
        if not ok_items:
            return False, msg_items, None
        try:
            restocking_fee = float(restocking_fee or 0.0)
        except Exception:
            restocking_fee = 0.0
        if restocking_fee < 0:
            return False, "Restocking fee cannot be negative", None
        try:
            adjustment = float(adjustment or 0.0)
        except Exception:
            adjustment = 0.0
        total_credit = subtotal + vat_total - restocking_fee + adjustment
        credit_note_id = self.generate_credit_note_id()
        credit_note = {
            "credit_note_id": credit_note_id,
            "date": datetime.now().strftime("%Y-%m-%d"),
            "source": "manual",
            "reference": str(reference or "").strip(),
            "customer_name": customer_name,
            "items": cleaned_items,
            "subtotal": subtotal,
            "vat_amount": vat_total,
            "restocking_fee": restocking_fee,
            "adjustment": adjustment,
            "total_credit": total_credit,
            "notes": str(notes or "").strip(),
            "processed_by": getpass.getuser(),
        }
        self.credit_notes = list(self.credit_notes or []) + [credit_note]
        if not self.save_credit_notes():
            self.credit_notes = [c for c in (self.credit_notes or []) if c.get("credit_note_id") != credit_note_id]
            return False, "Failed to save credit note", None
        self.append_audit_log(
            "credit_note_created",
            {
                "credit_note_id": credit_note_id,
                "source": "manual",
                "customer": customer_name,
                "total_credit": total_credit,
            },
        )
        return True, "Credit note created", credit_note

    def _prepare_manual_credit_note_items(self, items):
        if not isinstance(items, list) or not items:
            return False, "At least one item is required", None, 0.0, 0.0
        cleaned_items = []
        subtotal = 0.0
        vat_total = 0.0
        for row in items:
            if not isinstance(row, dict):
                continue
            desc = str(row.get("description") or "").strip()
            if not desc:
                continue
            try:
                qty = float(row.get("quantity") or 0)
                unit_price = float(row.get("unit_price") or 0)
            except Exception:
                return False, "Invalid quantity or unit price", None, 0.0, 0.0
            if qty <= 0 or unit_price < 0:
                return False, "Quantity must be > 0 and unit price cannot be negative", None, 0.0, 0.0
            taxable = bool(row.get("taxable", False))
            try:
                vat_rate = float(row.get("vat_rate") or 0.0)
            except Exception:
                vat_rate = 0.0
            if vat_rate > 1:
                vat_rate = vat_rate / 100.0
            line_subtotal = qty * unit_price
            line_vat = (line_subtotal * vat_rate) if taxable else 0.0
            subtotal += line_subtotal
            vat_total += line_vat
            cleaned_items.append({
                "item_code": str(row.get("item_code") or row.get("item_id") or "").strip(),
                "description": desc,
                "quantity": qty,
                "unit_price": unit_price,
                "taxable": taxable,
                "vat_rate": vat_rate,
                "line_total": line_subtotal,
            })
        if not cleaned_items:
            return False, "At least one valid item row is required", None, 0.0, 0.0
        return True, "OK", cleaned_items, subtotal, vat_total

    def update_manual_credit_note(self, credit_note_id, customer_name, items, notes="", restocking_fee=0.0, adjustment=0.0, reference=""):
        credit_note_id = str(credit_note_id or "").strip()
        if not credit_note_id:
            return False, "Credit note ID is required", None
        customer_name = str(customer_name or "").strip()
        if not customer_name:
            return False, "Customer name is required", None
        target = None
        for cn in (self.credit_notes or []):
            if cn.get("credit_note_id") == credit_note_id:
                target = cn
                break
        if not target:
            return False, "Credit note not found", None
        if str(target.get("source") or "") != "manual":
            return False, "Only manually created credit notes can be edited", None

        ok_items, msg_items, cleaned_items, subtotal, vat_total = self._prepare_manual_credit_note_items(items)
        if not ok_items:
            return False, msg_items, None
        try:
            restocking_fee = float(restocking_fee or 0.0)
        except Exception:
            restocking_fee = 0.0
        if restocking_fee < 0:
            return False, "Restocking fee cannot be negative", None
        try:
            adjustment = float(adjustment or 0.0)
        except Exception:
            adjustment = 0.0
        total_credit = subtotal + vat_total - restocking_fee + adjustment

        target["customer_name"] = customer_name
        target["reference"] = str(reference or "").strip()
        target["items"] = cleaned_items
        target["subtotal"] = subtotal
        target["vat_amount"] = vat_total
        target["restocking_fee"] = restocking_fee
        target["adjustment"] = adjustment
        target["total_credit"] = total_credit
        target["notes"] = str(notes or "").strip()
        target["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        target["processed_by"] = getpass.getuser()

        if not self.save_credit_notes():
            return False, "Failed to save credit note changes", None
        self.append_audit_log(
            "credit_note_updated",
            {
                "credit_note_id": credit_note_id,
                "source": "manual",
                "customer": customer_name,
                "total_credit": total_credit,
            },
        )
        return True, "Credit note updated", target

    def delete_manual_credit_note(self, credit_note_id):
        credit_note_id = str(credit_note_id or "").strip()
        if not credit_note_id:
            return False, "Credit note ID is required"
        target = None
        for cn in (self.credit_notes or []):
            if cn.get("credit_note_id") == credit_note_id:
                target = cn
                break
        if not target:
            return False, "Credit note not found"
        if str(target.get("source") or "") != "manual":
            return False, "Only manually created credit notes can be deleted"
        before = len(self.credit_notes or [])
        self.credit_notes = [cn for cn in (self.credit_notes or []) if cn.get("credit_note_id") != credit_note_id]
        if len(self.credit_notes or []) == before:
            return False, "Credit note not found"
        if not self.save_credit_notes():
            return False, "Failed to save credit note deletion"
        self.append_audit_log(
            "credit_note_deleted",
            {
                "credit_note_id": credit_note_id,
                "source": "manual",
                "customer": str(target.get("customer_name") or ""),
                "total_credit": float(target.get("total_credit") or 0.0),
            },
        )
        return True, "Credit note deleted"

    def generate_credit_note_pdf(self, credit_note_dict):
        try:
            from hope_pharma_complete import EnhancedPDFGenerator
            out_folder = os.path.join(self.data_folder, "HopePharmaCreditNotes")
            gen = EnhancedPDFGenerator(output_folder=out_folder)
            return gen.generate_credit_note_pdf(credit_note_dict)
        except Exception as e:
            raise RuntimeError(str(e))

    def _require_positive_int(self, value, field_name):
        try:
            parsed = int(value)
        except Exception:
            raise ValueError(f"{field_name} must be a positive integer greater than zero")
        if parsed <= 0:
            raise ValueError(f"{field_name} must be a positive integer greater than zero")
        return parsed

    def _require_positive_float(self, value, field_name):
        try:
            parsed = float(value)
        except Exception:
            raise ValueError(f"{field_name} must be a positive number greater than zero")
        if parsed <= 0:
            raise ValueError(f"{field_name} must be a positive number greater than zero")
        return parsed

    def _calculate_selling_price_from_multiplier(self, total_cost, quantity, price_multiplier):
        qty = float(quantity or 0.0)
        if qty <= 0:
            return 0.0
        cost_per_item = float(total_cost or 0.0) / qty
        return cost_per_item * float(price_multiplier or 1.0)
    
    def add_item(self, name, category, description, total_cost, selling_price, 
                 quantity, min_stock, supplier, batch_number=None, expiry_date=None,
                 created_date=None, invoice_date=None, addition_source="Manual", is_storage_item=False,
                 max_sell_qty=999999, price_multiplier=None, manufacture_date=None,
                 package_type="", brand="", client_name=None, sku="", subcategory="",
                 unit_of_measure="PCS", pack_size="", vat_category="Standard",
                 barcode="", lot_number="", reserved_quantity=0, warehouse_code="",
                 bin_location="", max_stock=0, reorder_point=0, reorder_quantity=0,
                 status="Available", temperature_controlled=False, quarantine=False,
                 supplier_name="", product_status="Stored"):
        """Add new item to inventory"""
        item_id = self.generate_item_id()
        try:
            self.save_supplier(supplier)
        except Exception:
            pass
        max_sell_qty = self._require_positive_int(max_sell_qty, "Max quantity per sale")
        if price_multiplier is None:
            try:
                explicit_selling_price = float(selling_price or 0.0)
            except Exception:
                explicit_selling_price = 0.0
            if explicit_selling_price > 0 and float(quantity or 0.0) > 0:
                cost_per_item = float(total_cost or 0.0) / float(quantity or 1.0)
                if cost_per_item > 0:
                    price_multiplier = explicit_selling_price / cost_per_item
                else:
                    price_multiplier = 1.0
            else:
                price_multiplier = 1.0
                selling_price = explicit_selling_price
        else:
            price_multiplier = self._require_positive_float(price_multiplier, "Price multiplier")
            selling_price = self._calculate_selling_price_from_multiplier(total_cost, quantity, price_multiplier)
        new_item = InventoryItem(
            item_id, name, category, description, total_cost, selling_price,
            quantity, min_stock, supplier, batch_number, expiry_date,
            created_date=created_date, addition_source=addition_source,
            invoice_date=invoice_date, is_storage_item=is_storage_item,
            max_sell_qty=max_sell_qty, price_multiplier=price_multiplier,
            manufacture_date=manufacture_date, package_type=package_type,
            brand=brand, client_name=client_name, sku=sku,
            subcategory=subcategory, unit_of_measure=unit_of_measure,
            pack_size=pack_size, vat_category=vat_category, barcode=barcode,
            lot_number=lot_number, reserved_quantity=reserved_quantity,
            warehouse_code=warehouse_code, bin_location=bin_location,
            max_stock=max_stock, reorder_point=reorder_point,
            reorder_quantity=reorder_quantity, status=status,
            temperature_controlled=temperature_controlled, quarantine=quarantine,
            supplier_name=supplier_name,
            product_status=product_status
        )
        self.items.append(new_item)
        
        if self.save_inventory():
            return new_item
        return None
    
    def update_item(self, item_id, **kwargs):
        """Update item information"""
        for item in self.items:
            if item.item_id == item_id:
                if 'supplier' in kwargs:
                    try:
                        self.save_supplier(kwargs.get('supplier'))
                    except Exception:
                        pass
                if 'max_sell_qty' in kwargs:
                    kwargs['max_sell_qty'] = self._require_positive_int(kwargs.get('max_sell_qty'), "Max quantity per sale")
                if 'price_multiplier' in kwargs:
                    kwargs['price_multiplier'] = self._require_positive_float(kwargs.get('price_multiplier'), "Price multiplier")

                if 'selling_price' in kwargs and 'price_multiplier' not in kwargs:
                    try:
                        new_sp = float(kwargs.get('selling_price') or 0.0)
                        if new_sp > 0:
                            total_cost = float(kwargs.get('total_cost', getattr(item, 'total_cost', 0.0)) or 0.0)
                            quantity = float(kwargs.get('quantity', getattr(item, 'quantity', 0.0)) or 0.0)
                            if quantity > 0:
                                cost_per_item = total_cost / quantity
                                if cost_per_item > 0:
                                    kwargs['price_multiplier'] = new_sp / cost_per_item
                    except Exception:
                        pass

                should_recalc_price = False
                if 'price_multiplier' in kwargs:
                    should_recalc_price = True
                elif ('total_cost' in kwargs or 'quantity' in kwargs) and 'selling_price' not in kwargs:
                    should_recalc_price = True

                if should_recalc_price:
                    try:
                        multiplier = float(kwargs.get('price_multiplier', getattr(item, 'price_multiplier', 1.0)) or 1.0)
                        if multiplier <= 0:
                            raise ValueError
                        total_cost = float(kwargs.get('total_cost', getattr(item, 'total_cost', 0.0)) or 0.0)
                        quantity = float(kwargs.get('quantity', getattr(item, 'quantity', 0.0)) or 0.0)
                        kwargs['selling_price'] = self._calculate_selling_price_from_multiplier(total_cost, quantity, multiplier)
                        kwargs['price_multiplier'] = multiplier
                    except Exception:
                        pass

                for key, value in kwargs.items():
                    if hasattr(item, key):
                        setattr(item, key, value)
                item.last_updated = datetime.now().strftime("%Y-%m-%d")
                return self.save_inventory()
        return False
    
    def delete_item(self, item_id):
        """Delete item from inventory"""
        self.items = [item for item in self.items if item.item_id != item_id]
        return self.save_inventory()
    
    def get_item(self, item_id):
        """Get item by ID"""
        for item in self.items:
            if item.item_id == item_id:
                return item
        return None

    def _tokenize_inventory_lookup_text(self, value):
        text = str(value or "").strip().lower().replace("&", " and ")
        raw_tokens = re.findall(r"[a-z]+|\d+(?:\.\d+)?", text)
        normalized = []
        for token in raw_tokens:
            normalized.append(_INVENTORY_MATCH_TOKEN_SYNONYMS.get(token, token))
        return normalized

    def _normalize_inventory_lookup_text(self, value):
        return " ".join(self._tokenize_inventory_lookup_text(value))

    def _inventory_lookup_signature(self, value):
        tokens = self._tokenize_inventory_lookup_text(value)
        filtered_tokens = [token for token in tokens if token not in _INVENTORY_MATCH_IGNORABLE_TOKENS]
        alpha_tokens = tuple(token for token in filtered_tokens if re.search(r"[a-z]", token))
        numeric_tokens = tuple(token for token in filtered_tokens if re.fullmatch(r"\d+(?:\.\d+)?", token))
        return {
            "normalized": " ".join(tokens),
            "compact": "".join(tokens),
            "tokens": tuple(tokens),
            "filtered_tokens": tuple(filtered_tokens),
            "alpha_tokens": alpha_tokens,
            "numeric_tokens": numeric_tokens,
        }

    def _iter_inventory_match_candidates(self, item):
        yield "name", str(getattr(item, "name", "") or "")
        description = str(getattr(item, "description", "") or "").strip()
        if description and self._normalize_inventory_lookup_text(description) != self._normalize_inventory_lookup_text(getattr(item, "name", "")):
            yield "description", description

    def _is_strong_inventory_signature_match(self, target_sig, candidate_sig):
        if not target_sig.get("normalized") or not candidate_sig.get("normalized"):
            return False
        if target_sig["compact"] == candidate_sig["compact"]:
            return True

        target_numbers = set(target_sig["numeric_tokens"])
        candidate_numbers = set(candidate_sig["numeric_tokens"])
        if target_numbers or candidate_numbers:
            if target_numbers != candidate_numbers:
                return False

        target_alpha = set(target_sig["alpha_tokens"])
        candidate_alpha = set(candidate_sig["alpha_tokens"])
        if not target_alpha or not candidate_alpha:
            return False

        overlap = target_alpha & candidate_alpha
        if not overlap:
            return False

        extra_target = target_alpha - candidate_alpha
        extra_candidate = candidate_alpha - target_alpha
        if extra_target and not extra_target.issubset(_INVENTORY_MATCH_IGNORABLE_TOKENS):
            return False
        if extra_candidate and not extra_candidate.issubset(_INVENTORY_MATCH_IGNORABLE_TOKENS):
            return False
        return True

    def find_item_for_invoice_line(self, description):
        """Match an invoice line description to an inventory item using guarded real-world normalization."""
        target_sig = self._inventory_lookup_signature(description)
        if not target_sig["normalized"]:
            return None

        exact_matches = []
        for item in self.items:
            for _source_name, source_value in self._iter_inventory_match_candidates(item):
                candidate_sig = self._inventory_lookup_signature(source_value)
                if candidate_sig["normalized"] == target_sig["normalized"] or candidate_sig["compact"] == target_sig["compact"]:
                    exact_matches.append(item)
                    break
        unique_exact = []
        seen_item_ids = set()
        for item in exact_matches:
            if item.item_id not in seen_item_ids:
                unique_exact.append(item)
                seen_item_ids.add(item.item_id)
        if len(unique_exact) == 1:
            return unique_exact[0]

        strong_matches = []
        seen_item_ids = set()
        for item in self.items:
            for _source_name, source_value in self._iter_inventory_match_candidates(item):
                candidate_sig = self._inventory_lookup_signature(source_value)
                if self._is_strong_inventory_signature_match(target_sig, candidate_sig):
                    if item.item_id not in seen_item_ids:
                        strong_matches.append(item)
                        seen_item_ids.add(item.item_id)
                    break
        if len(strong_matches) == 1:
            return strong_matches[0]
        return None

    def _expiry_sort_key(self, item):
        expiry = str(getattr(item, "expiry_date", "") or "").strip()
        if not expiry:
            return (1, datetime.max.date())
        for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
            try:
                return (0, datetime.strptime(expiry, fmt).date())
            except Exception:
                continue
        return (1, datetime.max.date())

    def find_items_for_invoice_line_fefo(self, description):
        target_sig = self._inventory_lookup_signature(description)
        if not target_sig["normalized"]:
            return []
        matches = []
        seen_item_ids = set()
        for item in self.items:
            for _source_name, source_value in self._iter_inventory_match_candidates(item):
                candidate_sig = self._inventory_lookup_signature(source_value)
                if (
                    candidate_sig["normalized"] == target_sig["normalized"]
                    or candidate_sig["compact"] == target_sig["compact"]
                    or self._is_strong_inventory_signature_match(target_sig, candidate_sig)
                ):
                    if item.item_id not in seen_item_ids:
                        matches.append(item)
                        seen_item_ids.add(item.item_id)
                    break
        matches.sort(
            key=lambda item: (
                self._expiry_sort_key(item),
                str(getattr(item, "batch_number", "") or "").strip(),
                str(getattr(item, "item_id", "") or "").strip(),
            )
        )
        return matches

    def sync_invoice_to_inventory(self, invoice_dict):
        """Deduct inventory for invoice lines using FEFO across matching batches/items."""
        invoice_id = str((invoice_dict or {}).get('invoice_id') or "").strip()
        client_name = str((invoice_dict or {}).get('client_name') or "").strip() or "Invoice Client"
        items = (invoice_dict or {}).get('items') or []
        tax_rate = self._to_float((invoice_dict or {}).get('tax_rate'), 0.0)
        if tax_rate > 1:
            tax_rate = tax_rate / 100.0

        matched_rows = []
        unmatched_rows = []

        for row in items:
            if not isinstance(row, dict):
                continue
            description = str(row.get('description') or "").strip()
            quantity = max(self._to_float(row.get('quantity'), 0.0), 0.0)
            if not description or quantity <= 0:
                continue

            normalized_desc = self._normalize_inventory_lookup_text(description)
            if normalized_desc in {"delivery fee"} or normalized_desc.startswith("promotion "):
                continue

            candidate_items = self.find_items_for_invoice_line_fefo(description)
            if not candidate_items:
                unmatched_rows.append(description)
                continue
            remaining = quantity
            allocated = []
            total_available = 0.0
            for inventory_item in candidate_items:
                available_qty = float(getattr(inventory_item, 'quantity', 0) or 0)
                try:
                    available_qty = float(inventory_item.get_available_quantity())
                except Exception:
                    pass
                if available_qty <= 0:
                    continue
                total_available += available_qty
                take_qty = min(remaining, available_qty)
                if take_qty <= 0:
                    continue
                allocated.append((inventory_item, take_qty))
                remaining -= take_qty
                if remaining <= 0:
                    break
            if remaining > 0:
                unmatched_rows.append(f"{description} (needed {quantity:g}, available {total_available:g})")
                continue
            for inventory_item, take_qty in allocated:
                matched_rows.append({
                    'inventory_item': inventory_item,
                    'description': description,
                    'quantity': take_qty,
                    'unit_price': self._to_float(row.get('unit_price'), 0.0),
                    'taxable': bool(row.get('taxable', False)),
                })

        inventory_links = []
        applied_sales = []
        try:
            for row in matched_rows:
                success, result = self.record_sale(
                    row['inventory_item'].item_id,
                    row['quantity'],
                    row['unit_price'],
                    client_name,
                    invoice_id=invoice_id or None,
                    promotion_quantity=0,
                    vat_applied=row['taxable'],
                    vat_rate=tax_rate
                )
                if not success:
                    raise RuntimeError(str(result))
                applied_sales.append(result)
                inventory_links.append({
                    'item_id': row['inventory_item'].item_id,
                    'item_name': row['inventory_item'].name,
                    'quantity': row['quantity'],
                    'promotion_quantity': 0,
                    'cost_per_item': getattr(result, 'cost_per_item', row['inventory_item'].get_cost_per_item()),
                    'batch_number': getattr(row['inventory_item'], 'batch_number', ''),
                    'expiry_date': getattr(row['inventory_item'], 'expiry_date', ''),
                    'warehouse_code': getattr(row['inventory_item'], 'warehouse_code', ''),
                    'bin_location': getattr(row['inventory_item'], 'bin_location', ''),
                })
        except Exception as exc:
            if invoice_id:
                try:
                    self.reverse_invoice_sales(invoice_id)
                except Exception:
                    pass
            return False, f"Failed to sync invoice with inventory: {exc}", []

        if matched_rows:
            message = f"Matched {len(matched_rows)} FEFO inventory allocation(s) and deducted stock."
        else:
            message = "No invoice lines matched existing inventory items."
        if unmatched_rows:
            unique_unmatched = []
            seen = set()
            for line in unmatched_rows:
                key = line.lower()
                if key not in seen:
                    unique_unmatched.append(line)
                    seen.add(key)
            preview = ", ".join(unique_unmatched[:5])
            if len(unique_unmatched) > 5:
                preview += f" and {len(unique_unmatched) - 5} more"
            message += f" Unmatched lines: {preview}."
        return True, message, inventory_links
    
    def search_items(self, query):
        """Search items by name or description"""
        query = query.lower()
        results = []
        for item in self.items:
            if (query in item.name.lower() or 
                query in item.description.lower() or
                query in item.category.lower()):
                results.append(item)
        return results
    
    def get_low_stock_items(self):
        """Get items with low stock"""
        return [item for item in self.items if item.is_low_stock()]
    
    def get_expired_items(self):
        """Get expired items"""
        return [item for item in self.items if item.is_expired()]

    def get_expiring_soon_items(self, days: int = 14):
        """Get items expiring within the next 'days' days"""
        return [item for item in self.items if item.is_expiring_soon(days)]
    
    def record_sale(self, item_id, quantity, selling_price, customer_name, invoice_id=None, promotion_quantity: int = 0, vat_applied: bool = False, vat_rate: float = 0.05):
        item = self.get_item(item_id)
        if not item:
            return False, "Item not found"
        
        total_units = quantity + max(0, promotion_quantity)
        try:
            max_sell_qty = int(getattr(item, 'max_sell_qty', 999999) or 999999)
        except Exception:
            max_sell_qty = 999999
        if max_sell_qty <= 0:
            max_sell_qty = 999999
        if total_units > max_sell_qty:
            return False, f"Max quantity per sale for this item is {max_sell_qty}"
        if item.quantity < total_units:
            return False, f"Insufficient stock. Available: {item.quantity}"
        
        cost_per_item = item.get_cost_per_item()
        total_cost = cost_per_item * total_units
        total_amount = quantity * selling_price
        total_profit = total_amount - total_cost
        vat_amount = (total_amount * vat_rate) if vat_applied else 0.0
        
        remaining_quantity = item.quantity - total_units
        if remaining_quantity > 0:
            item.total_cost = cost_per_item * remaining_quantity
        else:
            item.total_cost = 0
            
        item.quantity = remaining_quantity
        item.last_updated = datetime.now().strftime("%Y-%m-%d")
        
        sale_id = self.generate_sale_id()
        new_sale = SaleRecord(
            sale_id, item_id, item.name, quantity, selling_price, cost_per_item,
            total_amount, total_cost, total_profit, customer_name, invoice_id=invoice_id,
            promotion_quantity=promotion_quantity, vat_applied=vat_applied, vat_rate=vat_rate, vat_amount=vat_amount
        )
        self.sales.append(new_sale)
        
        if self.save_inventory() and self.save_sales():
            return True, new_sale
        return False, "Failed to save sale record"
    
    def reverse_invoice_sales(self, invoice_id):
        """Reverse all sales associated with an invoice ID"""
        sales_to_reverse = [s for s in self.sales if getattr(s, 'invoice_id', None) == invoice_id]
        
        if not sales_to_reverse:
            return True, "No sales found for this invoice"
            
        try:
            for sale in sales_to_reverse:
                item = self.get_item(sale.item_id)
                if item:
                    # Calculate total quantity to restore (sold + promo)
                    promo_qty = getattr(sale, 'promotion_quantity', 0)
                    qty_to_restore = sale.quantity + promo_qty
                    
                    # Restore quantity
                    item.quantity += qty_to_restore
                    
                    # Restore inventory value (cost)
                    # Use sale.total_cost which represents the COGS for this sale
                    item.total_cost += sale.total_cost
                    
                    item.last_updated = datetime.now().strftime("%Y-%m-%d")
            
            # Remove sales records
            self.sales = [s for s in self.sales if getattr(s, 'invoice_id', None) != invoice_id]
            
            if self.save_inventory() and self.save_sales():
                return True, f"Reversed {len(sales_to_reverse)} sales items"
            else:
                return False, "Failed to save inventory/sales changes"
                
        except Exception as e:
            return False, f"Error reversing sales: {str(e)}"

    def get_sale_by_id(self, sale_id):
        for sale in (self.sales or []):
            if str(getattr(sale, "sale_id", "") or "") == str(sale_id or ""):
                return sale
        return None

    def get_returned_quantity_for_sale(self, sale_id):
        total = 0
        for r in (self.returns or []):
            if str(getattr(r, "sale_id", "") or "") == str(sale_id or ""):
                try:
                    total += int(getattr(r, "quantity", 0) or 0)
                except Exception:
                    pass
        return total

    def _build_credit_note_from_return(self, return_record, sale_record, custom_notes="", manual_total_override=None):
        qty = int(getattr(return_record, "quantity", 0) or 0)
        unit_price = float(getattr(return_record, "unit_price", 0.0) or 0.0)
        subtotal = qty * unit_price
        vat_rate = float(getattr(return_record, "vat_rate", 0.0) or 0.0)
        vat_amount = (subtotal * vat_rate) if getattr(return_record, "vat_applied", False) else 0.0
        restocking_fee = float(getattr(return_record, "restocking_fee", 0.0) or 0.0)
        total_credit = subtotal + vat_amount - restocking_fee
        if manual_total_override is not None:
            total_credit = float(manual_total_override or 0.0)
            vat_amount = float(getattr(return_record, "vat_amount", vat_amount) or vat_amount)

        credit_note_id = self.generate_credit_note_id()
        credit_note = {
            "credit_note_id": credit_note_id,
            "date": datetime.now().strftime("%Y-%m-%d"),
            "source": "return",
            "return_id": getattr(return_record, "return_id", None),
            "sale_id": getattr(return_record, "sale_id", None),
            "invoice_id": getattr(sale_record, "invoice_id", None),
            "customer_name": getattr(return_record, "customer_name", "") or getattr(sale_record, "customer_name", ""),
            "items": [
                {
                    "item_id": getattr(return_record, "item_id", ""),
                    "description": getattr(return_record, "item_name", ""),
                    "quantity": qty,
                    "unit_price": unit_price,
                    "taxable": bool(getattr(return_record, "vat_applied", False)),
                    "vat_rate": vat_rate,
                    "line_total": subtotal,
                }
            ],
            "subtotal": subtotal,
            "vat_amount": vat_amount,
            "restocking_fee": restocking_fee,
            "total_credit": total_credit,
            "notes": (custom_notes or "").strip(),
            "processed_by": getattr(return_record, "processed_by", ""),
            "reason": getattr(return_record, "reason", ""),
        }
        return credit_note

    def process_return(
        self,
        sale_id,
        return_quantity,
        reason,
        processed_by=None,
        create_credit_note=True,
        credit_note_notes="",
        restocking_fee=0.0,
        manual_total_override=None,
    ):
        sale = self.get_sale_by_id(sale_id)
        if not sale:
            return False, "Sale record not found", None, None
        try:
            return_quantity = int(return_quantity)
        except Exception:
            return False, "Return quantity must be a whole number", None, None
        if return_quantity <= 0:
            return False, "Return quantity must be > 0", None, None
        already_returned = self.get_returned_quantity_for_sale(sale_id)
        sold_qty = int(getattr(sale, "quantity", 0) or 0)
        remaining = sold_qty - already_returned
        if return_quantity > remaining:
            return False, f"Cannot return more than remaining sold quantity. Remaining: {remaining}", None, None

        item = self.get_item(getattr(sale, "item_id", None))
        if not item:
            return False, "Inventory item for this sale no longer exists", None, None

        try:
            restocking_fee = float(restocking_fee or 0.0)
        except Exception:
            restocking_fee = 0.0
        if restocking_fee < 0:
            return False, "Restocking fee cannot be negative", None, None

        processed_by = (processed_by or "").strip() or getpass.getuser()
        reason = (reason or "").strip()
        if not reason:
            return False, "Return reason is required", None, None

        prev_qty = float(getattr(item, "quantity", 0) or 0)
        prev_total_cost = float(getattr(item, "total_cost", 0) or 0)

        credit_note = None
        return_record = None
        try:
            item.quantity = prev_qty + return_quantity
            try:
                unit_cost = float(getattr(sale, "cost_per_item", 0.0) or 0.0)
            except Exception:
                unit_cost = 0.0
            item.total_cost = prev_total_cost + (unit_cost * return_quantity)
            item.last_updated = datetime.now().strftime("%Y-%m-%d")

            vat_applied = bool(getattr(sale, "vat_applied", False))
            vat_rate = float(getattr(sale, "vat_rate", 0.0) or 0.0)
            unit_price = float(getattr(sale, "selling_price", 0.0) or 0.0)
            subtotal = float(return_quantity) * unit_price
            vat_amount = (subtotal * vat_rate) if vat_applied else 0.0
            total_credit = subtotal + vat_amount - restocking_fee

            return_id = self.generate_return_id()
            return_record = ReturnRecord(
                return_id=return_id,
                sale_id=str(getattr(sale, "sale_id", "")),
                item_id=str(getattr(sale, "item_id", "")),
                item_name=str(getattr(sale, "item_name", "")),
                quantity=int(return_quantity),
                unit_price=unit_price,
                customer_name=str(getattr(sale, "customer_name", "")),
                return_date=datetime.now().strftime("%Y-%m-%d"),
                reason=reason,
                processed_by=processed_by,
                invoice_id=getattr(sale, "invoice_id", None),
                vat_applied=vat_applied,
                vat_rate=vat_rate,
                vat_amount=vat_amount,
                subtotal=subtotal,
                total_credit=total_credit,
                restocking_fee=restocking_fee,
                credit_note_id=None,
                notes=(credit_note_notes or "").strip(),
            )

            if create_credit_note:
                credit_note = self._build_credit_note_from_return(
                    return_record,
                    sale,
                    custom_notes=credit_note_notes,
                    manual_total_override=manual_total_override,
                )
                return_record.credit_note_id = credit_note.get("credit_note_id")

            self.returns = list(self.returns or []) + [return_record]
            if credit_note:
                self.credit_notes = list(self.credit_notes or []) + [credit_note]

            if not self.save_inventory():
                raise RuntimeError("Failed to save inventory changes")
            if not self.save_returns():
                raise RuntimeError("Failed to save return record")
            if credit_note and not self.save_credit_notes():
                raise RuntimeError("Failed to save credit note")

            self.append_audit_log(
                "inventory_return",
                {
                    "return_id": return_record.return_id,
                    "sale_id": return_record.sale_id,
                    "item_id": return_record.item_id,
                    "quantity": return_record.quantity,
                    "customer": return_record.customer_name,
                    "credit_note_id": return_record.credit_note_id,
                    "reason": return_record.reason,
                    "processed_by": return_record.processed_by,
                },
            )
            if credit_note:
                self.append_audit_log(
                    "credit_note_created",
                    {
                        "credit_note_id": credit_note.get("credit_note_id"),
                        "source": "return",
                        "return_id": return_record.return_id,
                        "sale_id": return_record.sale_id,
                        "total_credit": credit_note.get("total_credit"),
                        "customer": credit_note.get("customer_name"),
                    },
                )

            return True, "Return processed successfully", return_record, credit_note
        except Exception as exc:
            try:
                item.quantity = prev_qty
                item.total_cost = prev_total_cost
                item.last_updated = datetime.now().strftime("%Y-%m-%d")
                self.save_inventory()
            except Exception:
                pass
            if return_record is not None:
                try:
                    self.returns = [r for r in (self.returns or []) if getattr(r, "return_id", None) != getattr(return_record, "return_id", None)]
                    self.save_returns()
                except Exception:
                    pass
            if credit_note is not None:
                try:
                    self.credit_notes = [c for c in (self.credit_notes or []) if c.get("credit_note_id") != credit_note.get("credit_note_id")]
                    self.save_credit_notes()
                except Exception:
                    pass
            return False, f"Failed to process return: {exc}", None, None
    
    def get_sales_report(self, start_date, end_date):
        """Get sales report for date range with profit analysis"""
        report_sales = []
        total_sales = 0
        total_cost = 0
        total_profit = 0
        total_quantity = 0
        total_promotion_quantity = 0
        valid_item_ids = {item.item_id for item in self.items}
        
        for sale in self.sales:
            if sale.item_id not in valid_item_ids:
                continue
            if (sale.total_amount == 0 and sale.total_cost == 0 and sale.total_profit == 0):
                continue
            if start_date <= sale.sale_date <= end_date:
                report_sales.append(sale)
                total_sales += sale.total_amount
                total_cost += sale.total_cost
                total_profit += sale.total_profit
                total_quantity += sale.quantity + getattr(sale, 'promotion_quantity', 0)
                total_promotion_quantity += getattr(sale, 'promotion_quantity', 0)
        
        return {
            'sales': report_sales,
            'total_sales': total_sales,
            'total_cost': total_cost,
            'total_profit': total_profit,
            'total_quantity': total_quantity,
            'total_promotion_quantity': total_promotion_quantity,
            'profit_margin': (total_profit / total_sales * 100) if total_sales > 0 else 0,
            'start_date': start_date,
            'end_date': end_date
        }
    
    def add_category(self, category_name):
        """Add new category"""
        if category_name not in self.categories:
            self.categories.append(category_name)
            return self.save_categories()
        return True
    
    def add_cost_type(self, cost_type_name):
        """Add new cost type"""
        if cost_type_name not in self.cost_types:
            self.cost_types.append(cost_type_name)
            return self.save_cost_types()
        return True
    
    def get_inventory_value(self):
        """Calculate total inventory value at cost"""
        return sum(item.get_total_value() for item in self.items)
    
    def get_inventory_summary(self):
        """Get inventory summary"""
        total_items = len(self.items)
        low_stock_count = len(self.get_low_stock_items())
        expired_count = len(self.get_expired_items())
        total_value = self.get_inventory_value()
        total_potential_sales = sum(item.get_sales_value() for item in self.items)
        total_potential_profit = sum(item.get_total_profit_potential() for item in self.items)
        
        # Calculate total additional costs
        total_additional_costs = sum(
            sum(cost['amount'] for cost in item.costs) 
            for item in self.items
        )
        
        return {
            'total_items': total_items,
            'low_stock_count': low_stock_count,
            'expired_count': expired_count,
            'total_value': total_value,
            'total_additional_costs': total_additional_costs,
            'total_potential_sales': total_potential_sales,
            'total_potential_profit': total_potential_profit
        }

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
        
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # Ensure the inner frame takes up the full width of the canvas
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        
        # Mouse wheel binding
        self.canvas.bind("<Enter>", self._bind_mousewheel)
        self.canvas.bind("<Leave>", self._unbind_mousewheel)
        
    def _on_canvas_configure(self, event):
        """Update the width of the scrollable frame to match the canvas width"""
        self.canvas.itemconfig(self.canvas_window, width=event.width)
        
    def _bind_mousewheel(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        
    def _unbind_mousewheel(self, event):
        self.canvas.unbind_all("<MouseWheel>")
        
    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
    def pack(self, **kwargs):
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")

class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)
    def show(self, event=None):
        if self.tip:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + 20
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        frm = ttk.Frame(self.tip, padding="6", relief="solid", borderwidth=1)
        frm.pack()
        ttk.Label(frm, text=self.text, wraplength=280).pack()
    def hide(self, event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None

class AddItemDialog:
    def __init__(self, parent, inventory_manager):
        self.inventory_manager = inventory_manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Add New Inventory Item")
        self.dialog.geometry("600x700")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
    
    def setup_ui(self):
        # Create main container with scrollbar
        attach_clock_label(self.dialog)
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="Add New Inventory Item", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Basic Information
        basic_frame = ttk.LabelFrame(main_frame, text="Basic Information", padding="10")
        basic_frame.pack(fill='x', pady=5)
        
        ttk.Label(basic_frame, text="Item Name *:").grid(row=0, column=0, sticky='w', pady=5)
        self.name_var = tk.StringVar()
        ttk.Entry(basic_frame, textvariable=self.name_var, width=30).grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(basic_frame, text="Category *:").grid(row=1, column=0, sticky='w', pady=5)
        self.category_var = tk.StringVar()
        category_combo = ttk.Combobox(basic_frame, textvariable=self.category_var, 
                                     values=self.inventory_manager.categories, width=27)
        category_combo.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(basic_frame, text="Description:").grid(row=2, column=0, sticky='w', pady=5)
        self.description_text = scrolledtext.ScrolledText(basic_frame, height=3, width=40)
        self.description_text.grid(row=2, column=1, sticky='w', pady=5, padx=10)
        
        # Cost Information - UPDATED FOR TOTAL COST
        cost_frame = ttk.LabelFrame(main_frame, text="Cost Information", padding="10")
        cost_frame.pack(fill='x', pady=5)
        
        ttk.Label(cost_frame, text="Total Cost (AED) *:").grid(row=0, column=0, sticky='w', pady=5)
        self.total_cost_var = tk.StringVar()
        total_cost_entry = ttk.Entry(cost_frame, textvariable=self.total_cost_var, width=15)
        total_cost_entry.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        ttk.Label(cost_frame, text="Total cost for all items").grid(row=0, column=2, sticky='w', pady=5, padx=5)
        
        ttk.Label(cost_frame, text="Selling Price (AED):").grid(row=1, column=0, sticky='w', pady=5)
        self.selling_price_var = tk.StringVar()
        selling_price_entry = ttk.Entry(cost_frame, textvariable=self.selling_price_var, width=15)
        selling_price_entry.grid(row=1, column=1, sticky='w', pady=5, padx=10)
        ttk.Label(cost_frame, text="Per item (optional)").grid(row=1, column=2, sticky='w', pady=5, padx=5)
        
        # Auto-calculated fields
        self.cost_per_item_var = tk.StringVar(value="Cost per item: AED 0.00")
        ttk.Label(cost_frame, textvariable=self.cost_per_item_var).grid(row=2, column=0, columnspan=3, sticky='w', pady=5)
        
        self.profit_per_item_var = tk.StringVar(value="Profit per item: AED 0.00")
        ttk.Label(cost_frame, textvariable=self.profit_per_item_var).grid(row=3, column=0, columnspan=3, sticky='w', pady=5)
        
        # Bind events for auto-calculation
        self.total_cost_var.trace('w', self.update_calculations)
        self.selling_price_var.trace('w', self.update_calculations)
        
        # Stock Information
        stock_frame = ttk.LabelFrame(main_frame, text="Stock Information", padding="10")
        stock_frame.pack(fill='x', pady=5)
        
        ttk.Label(stock_frame, text="Quantity *:").grid(row=0, column=0, sticky='w', pady=5)
        self.quantity_var = tk.StringVar(value="1")
        quantity_entry = ttk.Entry(stock_frame, textvariable=self.quantity_var, width=15)
        quantity_entry.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        quantity_entry.bind('<KeyRelease>', self.update_calculations)
        
        ttk.Label(stock_frame, text="Minimum Stock Level *:").grid(row=1, column=0, sticky='w', pady=5)
        self.min_stock_var = tk.StringVar(value="5")
        ttk.Entry(stock_frame, textvariable=self.min_stock_var, width=15).grid(row=1, column=1, sticky='w', pady=5, padx=10)

        # Supplier Information
        supplier_frame = ttk.LabelFrame(main_frame, text="Supplier Information", padding="10")
        supplier_frame.pack(fill='x', pady=5)
        
        ttk.Label(supplier_frame, text="Supplier Name *:").grid(row=0, column=0, sticky='w', pady=5)
        self.supplier_var = tk.StringVar()
        self.supplier_combo = ttk.Combobox(supplier_frame, textvariable=self.supplier_var, width=27)
        try:
            self.supplier_combo['values'] = self.inventory_manager.load_suppliers()
        except Exception:
            self.supplier_combo['values'] = []
        self.supplier_combo.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        # Dates Information
        dates_frame = ttk.LabelFrame(main_frame, text="Dates", padding="10")
        dates_frame.pack(fill='x', pady=5)
        
        ttk.Label(dates_frame, text="Date of Entry (YYYY-MM-DD):").grid(row=0, column=0, sticky='w', pady=5)
        self.date_of_entry_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(dates_frame, textvariable=self.date_of_entry_var, width=15).grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(dates_frame, text="Invoice Date (YYYY-MM-DD):").grid(row=1, column=0, sticky='w', pady=5)
        self.invoice_date_var = tk.StringVar()
        ttk.Entry(dates_frame, textvariable=self.invoice_date_var, width=15).grid(row=1, column=1, sticky='w', pady=5, padx=10)
        
        # Batch Information (for medicines)
        batch_frame = ttk.LabelFrame(main_frame, text="Batch Information (Optional)", padding="10")
        batch_frame.pack(fill='x', pady=5)
        
        ttk.Label(batch_frame, text="Batch Number:").grid(row=0, column=0, sticky='w', pady=5)
        self.batch_var = tk.StringVar()
        ttk.Entry(batch_frame, textvariable=self.batch_var, width=20).grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(batch_frame, text="Manufacturing Date (YYYY-MM-DD):").grid(row=1, column=0, sticky='w', pady=5)
        self.manufacture_var = tk.StringVar()
        ttk.Entry(batch_frame, textvariable=self.manufacture_var, width=15).grid(row=1, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(batch_frame, text="Package Type:").grid(row=2, column=0, sticky='w', pady=5)
        self.package_type_var = tk.StringVar(value="Other")
        ttk.Combobox(
            batch_frame,
            textvariable=self.package_type_var,
            values=["Box", "Carton", "Other"],
            width=17,
            state="normal"
        ).grid(row=2, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(batch_frame, text="Brand Name:").grid(row=3, column=0, sticky='w', pady=5)
        self.brand_var = tk.StringVar()
        ttk.Entry(batch_frame, textvariable=self.brand_var, width=20).grid(row=3, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(batch_frame, text="Expiry Date (YYYY-MM-DD):").grid(row=4, column=0, sticky='w', pady=5)
        self.expiry_var = tk.StringVar()
        ttk.Entry(batch_frame, textvariable=self.expiry_var, width=15).grid(row=4, column=1, sticky='w', pady=5, padx=10)
        
        # Addition Source
        source_frame = ttk.LabelFrame(main_frame, text="System Information", padding="10")
        source_frame.pack(fill='x', pady=5)
        
        ttk.Label(source_frame, text="Addition Method *:").grid(row=0, column=0, sticky='w', pady=5)
        self.source_var = tk.StringVar(value="Manual Entry")
        source_combo = ttk.Combobox(source_frame, textvariable=self.source_var, 
                                   values=["Manual Entry", "Bulk Import", "Supplier Delivery", "Inventory Adjustment"], width=27)
        source_combo.grid(row=0, column=1, sticky='w', pady=5, padx=10)

        ttk.Label(source_frame, text="Product Status:").grid(row=1, column=0, sticky='w', pady=5)
        self.product_status_var = tk.StringVar(value="Stored")
        ttk.Combobox(
            source_frame,
            textvariable=self.product_status_var,
            values=["Marketed", "Registered", "Stored"],
            state="readonly",
            width=27
        ).grid(row=1, column=1, sticky='w', pady=5, padx=10)
        
        # Storage Item Flag
        self.is_storage_item_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(main_frame, text="Storage Item (Consignment/Stored Goods - No Direct Sale)", 
                        variable=self.is_storage_item_var).pack(fill='x', pady=5)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="Add Item", 
                  command=self.add_item).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side='left', padx=5)
    
    def update_calculations(self, *args):
        """Update auto-calculated fields"""
        try:
            total_cost = float(self.total_cost_var.get()) if self.total_cost_var.get() else 0
            quantity = int(self.quantity_var.get()) if self.quantity_var.get() else 1
            selling_price = float(self.selling_price_var.get()) if self.selling_price_var.get() else 0
            
            if quantity > 0:
                cost_per_item = total_cost / quantity
                self.cost_per_item_var.set(f"Cost per item: AED {cost_per_item:.2f}")
                
                if selling_price > 0:
                    profit_per_item = selling_price - cost_per_item
                    self.profit_per_item_var.set(f"Profit per item: AED {profit_per_item:.2f}")
                else:
                    self.profit_per_item_var.set("Profit per item: Set selling price to calculate")
            else:
                self.cost_per_item_var.set("Cost per item: Enter quantity > 0")
                self.profit_per_item_var.set("Profit per item: Enter quantity > 0")
                
        except (ValueError, ZeroDivisionError):
            self.cost_per_item_var.set("Cost per item: AED 0.00")
            self.profit_per_item_var.set("Profit per item: AED 0.00")
    
    def add_item(self):
        """Add the new item to inventory"""
        try:
            # Validate required fields
            if not self.name_var.get().strip():
                messagebox.showwarning("Warning", "Please enter item name")
                return
            
            if not self.category_var.get().strip():
                messagebox.showwarning("Warning", "Please select category")
                return
            
            # Validate costs and quantities
            try:
                total_cost = float(self.total_cost_var.get())
                quantity = int(self.quantity_var.get())
                min_stock = int(self.min_stock_var.get())
                
                if total_cost < 0 or quantity <= 0 or min_stock < 0:
                    raise ValueError
                    
                selling_price = 0
                if self.selling_price_var.get():
                    selling_price = float(self.selling_price_var.get())
                    if selling_price < 0:
                        raise ValueError
                        
            except ValueError:
                messagebox.showwarning("Warning", "Please enter valid numeric values for costs and quantities")
                return
            
            if not self.supplier_var.get().strip():
                messagebox.showwarning("Warning", "Please enter supplier name")
                return
            
            # Validate manufacturing/expiry dates if provided
            manufacture_date = self.manufacture_var.get().strip()
            if manufacture_date:
                try:
                    datetime.strptime(manufacture_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid Manufacturing Date (YYYY-MM-DD)")
                    return

            expiry_date = self.expiry_var.get().strip()
            if expiry_date:
                try:
                    datetime.strptime(expiry_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid expiry date (YYYY-MM-DD)")
                    return
            
            # Validate dates of entry and invoice date
            created_date = self.date_of_entry_var.get().strip()
            if created_date:
                try:
                    datetime.strptime(created_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid Date of Entry (YYYY-MM-DD)")
                    return
            else:
                created_date = datetime.now().strftime("%Y-%m-%d")
            
            invoice_date = self.invoice_date_var.get().strip()
            if invoice_date:
                try:
                    datetime.strptime(invoice_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid Invoice Date (YYYY-MM-DD)")
                    return
            
            # Add new category if it doesn't exist
            self.inventory_manager.add_category(self.category_var.get().strip())
            
            # Add item
            new_item = self.inventory_manager.add_item(
                name=self.name_var.get().strip(),
                category=self.category_var.get().strip(),
                description=self.description_text.get('1.0', tk.END).strip(),
                total_cost=total_cost,
                selling_price=selling_price,
                quantity=quantity,
                min_stock=min_stock,
                supplier=self.supplier_var.get().strip(),
                batch_number=self.batch_var.get().strip() or None,
                manufacture_date=manufacture_date or None,
                package_type=self.package_type_var.get().strip() or "Other",
                brand=self.brand_var.get().strip(),
                expiry_date=expiry_date or None,
                created_date=created_date,
                invoice_date=invoice_date or None,
                addition_source=self.source_var.get().strip(),
                is_storage_item=self.is_storage_item_var.get(),
                product_status=self.product_status_var.get().strip() or "Stored"
            )
            
            if new_item:
                cost_per_item = new_item.get_cost_per_item()
                messagebox.showinfo("Success", 
                                  f"Item added successfully!\n\n"
                                  f"Item ID: {new_item.item_id}\n"
                                  f"Name: {new_item.name}\n"
                                  f"Total Cost: AED {total_cost:.2f}\n"
                                  f"Cost per Item: AED {cost_per_item:.2f}\n"
                                  f"Quantity: {new_item.quantity}")
                self.dialog.destroy()
            else:
                messagebox.showerror("Error", "Failed to add item")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add item: {e}")

class EditItemDialog(AddItemDialog):
    def __init__(self, parent, inventory_manager, item_id):
        self.item_id = item_id
        self.item = inventory_manager.get_item(item_id)
        super().__init__(parent, inventory_manager)
        self.dialog.title(f"Edit Item: {item_id}")
        self.load_item_data()
    
    def load_item_data(self):
        """Load existing item data into form"""
        if not self.item:
            messagebox.showerror("Error", "Item not found")
            self.dialog.destroy()
            return
        
        self.name_var.set(self.item.name)
        self.category_var.set(self.item.category)
        self.description_text.delete('1.0', tk.END)
        self.description_text.insert('1.0', self.item.description)
        self.total_cost_var.set(str(self.item.total_cost))
        self.selling_price_var.set(str(self.item.selling_price))
        self.quantity_var.set(str(self.item.quantity))
        self.min_stock_var.set(str(self.item.min_stock))
        self.supplier_var.set(self.item.supplier)
        self.batch_var.set(self.item.batch_number or "")
        self.manufacture_var.set(getattr(self.item, 'manufacture_date', "") or "")
        self.package_type_var.set(getattr(self.item, 'package_type', "") or "Other")
        self.brand_var.set(getattr(self.item, 'brand', "") or "")
        self.expiry_var.set(self.item.expiry_date or "")
        try:
            self.date_of_entry_var.set(self.item.created_date or "")
        except Exception:
            pass
        try:
            self.invoice_date_var.set(self.item.invoice_date or "")
        except Exception:
            pass
        
        self.is_storage_item_var.set(getattr(self.item, 'is_storage_item', False))
        try:
            self.product_status_var.set(getattr(self.item, "product_status", "") or ("Stored" if getattr(self.item, "is_storage_item", False) else "Marketed"))
        except Exception:
            self.product_status_var.set("Stored")

        # Update calculations
        self.update_calculations()
    
    def add_item(self):
        """Update the existing item"""
        try:
            # Validate required fields
            if not self.name_var.get().strip():
                messagebox.showwarning("Warning", "Please enter item name")
                return
            
            if not self.category_var.get().strip():
                messagebox.showwarning("Warning", "Please select category")
                return
            
            # Validate costs and quantities
            try:
                total_cost = float(self.total_cost_var.get())
                quantity = int(self.quantity_var.get())
                min_stock = int(self.min_stock_var.get())
                
                if total_cost < 0 or quantity <= 0 or min_stock < 0:
                    raise ValueError
                    
                selling_price = 0
                if self.selling_price_var.get():
                    selling_price = float(self.selling_price_var.get())
                    if selling_price < 0:
                        raise ValueError
                        
            except ValueError:
                messagebox.showwarning("Warning", "Please enter valid numeric values for costs and quantities")
                return
            
            if not self.supplier_var.get().strip():
                messagebox.showwarning("Warning", "Please enter supplier name")
                return
            
            # Validate manufacturing/expiry dates if provided
            manufacture_date = self.manufacture_var.get().strip()
            if manufacture_date:
                try:
                    datetime.strptime(manufacture_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid Manufacturing Date (YYYY-MM-DD)")
                    return

            expiry_date = self.expiry_var.get().strip()
            if expiry_date:
                try:
                    datetime.strptime(expiry_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid expiry date (YYYY-MM-DD)")
                    return

            # Validate dates of entry and invoice date
            created_date = self.date_of_entry_var.get().strip()
            if created_date:
                try:
                    datetime.strptime(created_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid Date of Entry (YYYY-MM-DD)")
                    return
            else:
                created_date = datetime.now().strftime("%Y-%m-%d")

            invoice_date = self.invoice_date_var.get().strip()
            if invoice_date:
                try:
                    datetime.strptime(invoice_date, "%Y-%m-%d")
                except ValueError:
                    messagebox.showwarning("Warning", "Please enter valid Invoice Date (YYYY-MM-DD)")
                    return
            
            # Add new category if it doesn't exist
            self.inventory_manager.add_category(self.category_var.get().strip())
            
            # Update item
            success = self.inventory_manager.update_item(
                self.item_id,
                name=self.name_var.get().strip(),
                category=self.category_var.get().strip(),
                description=self.description_text.get('1.0', tk.END).strip(),
                total_cost=total_cost,
                selling_price=selling_price,
                quantity=quantity,
                min_stock=min_stock,
                supplier=self.supplier_var.get().strip(),
                batch_number=self.batch_var.get().strip() or None,
                manufacture_date=manufacture_date or None,
                package_type=self.package_type_var.get().strip() or "Other",
                brand=self.brand_var.get().strip(),
                expiry_date=expiry_date or None,
                created_date=created_date,
                invoice_date=invoice_date or None,
                is_storage_item=self.is_storage_item_var.get(),
                product_status=self.product_status_var.get().strip() or "Stored"
            )
            
            if success:
                messagebox.showinfo("Success", "Item updated successfully!")
                self.dialog.destroy()
            else:
                messagebox.showerror("Error", "Failed to update item")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update item: {e}")

class ManageCostsDialog:
    def __init__(self, parent, inventory_manager, item_id):
        self.inventory_manager = inventory_manager
        self.item_id = item_id
        self.item = inventory_manager.get_item(item_id)
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Manage Costs - {item_id}")
        self.dialog.geometry("600x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        self.setup_ui()
        self.refresh_costs()
    
    def setup_ui(self):
        attach_clock_label(self.dialog)
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        ttk.Label(main_frame, text=f"Manage Additional Costs\n{self.item.name}", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # Cost breakdown
        breakdown = self.item.get_cost_breakdown()
        breakdown_frame = ttk.LabelFrame(main_frame, text="Cost Breakdown", padding="10")
        breakdown_frame.pack(fill='x', pady=5)
        
        ttk.Label(breakdown_frame, text=f"Total Cost: AED {breakdown['total_cost']:.2f}").pack(anchor='w')
        ttk.Label(breakdown_frame, text=f"Cost per Item: AED {breakdown['cost_per_item']:.2f}").pack(anchor='w')
        ttk.Label(breakdown_frame, text=f"Additional Costs: AED {breakdown['additional_costs']:.2f}").pack(anchor='w')
        if breakdown['selling_price_per_item']:
            ttk.Label(breakdown_frame, text=f"Profit per Item: AED {breakdown['profit_per_item']:.2f}", 
                     font=('Helvetica', 10, 'bold')).pack(anchor='w')
        
        # Add cost form
        add_cost_frame = ttk.LabelFrame(main_frame, text="Add New Cost", padding="10")
        add_cost_frame.pack(fill='x', pady=5)
        
        ttk.Label(add_cost_frame, text="Cost Type:").grid(row=0, column=0, sticky='w', pady=5)
        self.cost_type_var = tk.StringVar()
        cost_type_combo = ttk.Combobox(add_cost_frame, textvariable=self.cost_type_var,
                                      values=self.inventory_manager.cost_types, width=15)
        cost_type_combo.grid(row=0, column=1, sticky='w', pady=5, padx=10)
        
        ttk.Label(add_cost_frame, text="Amount (AED):").grid(row=0, column=2, sticky='w', pady=5)
        self.cost_amount_var = tk.StringVar()
        ttk.Entry(add_cost_frame, textvariable=self.cost_amount_var, width=10).grid(row=0, column=3, sticky='w', pady=5, padx=10)
        
        ttk.Label(add_cost_frame, text="Description:").grid(row=1, column=0, sticky='w', pady=5)
        self.cost_desc_var = tk.StringVar()
        ttk.Entry(add_cost_frame, textvariable=self.cost_desc_var, width=30).grid(row=1, column=1, columnspan=3, sticky='w', pady=5, padx=10)
        
        ttk.Button(add_cost_frame, text="Add Cost", 
                  command=self.add_cost).grid(row=2, column=0, pady=10, padx=(0,10), sticky='w')
        ttk.Button(add_cost_frame, text="Edit Selected", 
                  command=self.edit_cost).grid(row=2, column=1, pady=10, padx=(0,10), sticky='w')
        ttk.Button(add_cost_frame, text="Remove Selected", 
                  command=self.remove_cost).grid(row=2, column=2, pady=10, sticky='w')
        
        # Costs list
        costs_frame = ttk.LabelFrame(main_frame, text="Additional Costs", padding="10")
        costs_frame.pack(fill='both', expand=True, pady=5)
        
        # Create treeview
        tree_frame = ttk.Frame(costs_frame)
        tree_frame.pack(fill='both', expand=True)
        
        columns = ('Type', 'Amount', 'Description', 'Date')
        self.costs_tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=8)
        
        for col in columns:
            self.costs_tree.heading(col, text=col)
            self.costs_tree.column(col, width=120)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.costs_tree.yview)
        self.costs_tree.configure(yscrollcommand=scrollbar.set)
        
        self.costs_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        ttk.Button(costs_frame, text="Edit Selected Cost", 
                  command=self.edit_cost).pack(pady=5)
        # Remove cost button
        ttk.Button(costs_frame, text="Remove Selected Cost", 
                  command=self.remove_cost).pack(pady=5)
        
        # Close button
        ttk.Button(main_frame, text="Close", 
                  command=self.dialog.destroy).pack(pady=10)
    
    def refresh_costs(self):
        """Refresh costs list"""
        for item in self.costs_tree.get_children():
            self.costs_tree.delete(item)
        
        for cost in self.item.costs:
            self.costs_tree.insert('', tk.END, values=(
                cost['cost_type'],
                f"AED {cost['amount']:.2f}",
                cost['description'],
                cost['date']
            ))
    
    def add_cost(self):
        """Add new cost to item"""
        try:
            cost_type = self.cost_type_var.get().strip()
            amount_str = self.cost_amount_var.get().strip()
            description = self.cost_desc_var.get().strip()
            
            if not cost_type:
                messagebox.showwarning("Warning", "Please select cost type")
                return
            
            try:
                amount = float(amount_str)
                if amount <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Warning", "Please enter valid cost amount")
                return
            
            if not description:
                messagebox.showwarning("Warning", "Please enter cost description")
                return
            
            # Add new cost type if it doesn't exist
            self.inventory_manager.add_cost_type(cost_type)
            
            # Add cost to item
            self.item.add_cost(cost_type, amount, description)
            
            # Save inventory
            if self.inventory_manager.save_inventory():
                messagebox.showinfo("Success", "Cost added successfully!")
                self.refresh_costs()
                # Clear form
                self.cost_amount_var.set("")
                self.cost_desc_var.set("")
            else:
                messagebox.showerror("Error", "Failed to save cost")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add cost: {e}")
    
    def edit_cost(self):
        selection = self.costs_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a cost to edit")
            return
        vals = self.costs_tree.item(selection[0])['values']
        old_type = vals[0]
        old_amount_str = vals[1]
        old_desc = vals[2]
        old_date = vals[3]
        try:
            old_amount = float(str(old_amount_str).split()[-1])
        except Exception:
            try:
                old_amount = float(old_amount_str)
            except Exception:
                messagebox.showwarning("Warning", "Invalid amount format")
                return
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Edit Cost")
        dlg.geometry("420x240")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Type").grid(row=0, column=0, sticky='w')
        type_var = tk.StringVar(value=old_type)
        type_combo = ttk.Combobox(frm, textvariable=type_var, values=self.inventory_manager.cost_types, width=18)
        type_combo.grid(row=0, column=1, padx=8, pady=4)
        ttk.Label(frm, text="Amount").grid(row=1, column=0, sticky='w')
        amount_var = tk.StringVar(value=str(old_amount))
        amount_entry = ttk.Entry(frm, textvariable=amount_var, width=12)
        amount_entry.grid(row=1, column=1, padx=8, pady=4)
        ttk.Label(frm, text="Description").grid(row=2, column=0, sticky='w')
        desc_var = tk.StringVar(value=old_desc)
        desc_entry = ttk.Entry(frm, textvariable=desc_var, width=28)
        desc_entry.grid(row=2, column=1, padx=8, pady=4)
        ttk.Label(frm, text="Date (YYYY-MM-DD)").grid(row=3, column=0, sticky='w')
        date_var = tk.StringVar(value=old_date)
        date_entry = ttk.Entry(frm, textvariable=date_var, width=16)
        date_entry.grid(row=3, column=1, padx=8, pady=4)
        def save():
            try:
                new_type = type_var.get().strip()
                if not new_type:
                    raise ValueError
                new_amount = float(amount_var.get().strip())
                if new_amount < 0:
                    raise ValueError
                new_desc = desc_var.get().strip()
                new_date = date_var.get().strip() or old_date
                self.inventory_manager.add_cost_type(new_type)
                target = None
                for cost in self.item.costs:
                    if (cost.get('cost_type') == old_type and 
                        abs(cost.get('amount', 0) - old_amount) < 1e-9 and 
                        cost.get('description') == old_desc and 
                        str(cost.get('date')) == str(old_date)):
                        target = cost
                        break
                if target is None:
                    for cost in self.item.costs:
                        if cost.get('cost_type') == old_type and abs(cost.get('amount', 0) - old_amount) < 1e-9:
                            target = cost
                            break
                if target is None:
                    messagebox.showerror("Error", "Original cost not found")
                    return
                delta = new_amount - target.get('amount', 0)
                target['cost_type'] = new_type
                target['amount'] = new_amount
                target['description'] = new_desc
                target['date'] = new_date
                self.item.total_cost += delta
                self.item.last_updated = datetime.now().strftime("%Y-%m-%d")
                if self.inventory_manager.save_inventory():
                    self.refresh_costs()
                    dlg.destroy()
                    messagebox.showinfo("Success", "Cost updated successfully!")
                else:
                    messagebox.showerror("Error", "Failed to save changes")
            except Exception:
                messagebox.showwarning("Warning", "Please enter valid cost data")
        ttk.Button(frm, text="Save", command=save).grid(row=4, column=0, pady=8)
        ttk.Button(frm, text="Cancel", command=dlg.destroy).grid(row=4, column=1, pady=8)
    
    def remove_cost(self):
        """Remove selected cost"""
        selection = self.costs_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select a cost to remove")
            return
        
        # Get cost ID from selection
        selected_item = self.costs_tree.item(selection[0])
        cost_type = selected_item['values'][0]
        amount = selected_item['values'][1]
        
        if messagebox.askyesno("Confirm", f"Remove cost: {cost_type} - {amount}?"):
            # Find and remove cost
            for cost in self.item.costs:
                if (cost['cost_type'] == cost_type and 
                    f"AED {cost['amount']:.2f}" == amount):
                    self.item.remove_cost(cost['cost_id'])
                    break
            
            # Save inventory
            if self.inventory_manager.save_inventory():
                messagebox.showinfo("Success", "Cost removed successfully!")
                self.refresh_costs()
            else:
                messagebox.showerror("Error", "Failed to remove cost")

class SellItemDialog:
    def __init__(self, parent, inventory_manager, invoice_manager=None, preselected_item_ids=None):
        self.inventory_manager = inventory_manager
        self.invoice_manager = invoice_manager
        self.preselected_item_ids = [str(item_id) for item_id in (preselected_item_ids or []) if item_id]
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Sell Items - Point of Sale")
        self.dialog.geometry("900x750")  # Increased size for cart
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        try:
            self.dialog.minsize(800, 600)
        except Exception:
            pass
        
        self.selected_item = None
        self.cart_items = []  # List to store cart items
        self.setup_ui()
    
    def setup_ui(self):
        attach_clock_label(self.dialog)
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)
        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)
        main_frame = scroll_frame.scrollable_frame
        
        ttk.Label(main_frame, text="Point of Sale", 
                 font=('Helvetica', 14, 'bold')).pack(pady=(0, 20))
        
        # --- Top Section: Item Selection and Details ---
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill='x', expand=True)
        
        # Left: Item Selection
        left_frame = ttk.LabelFrame(top_frame, text="1. Select Item", padding="10")
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        columns = ('ID', 'Name', 'Stock', 'Price')
        self.items_tree = ttk.Treeview(left_frame, columns=columns, show='headings', height=10, selectmode='extended')
        
        for col in columns:
            self.items_tree.heading(col, text=col)
            self.items_tree.column(col, width=80 if col != 'Name' else 150)
        
        scrollbar = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.items_tree.yview)
        self.items_tree.configure(yscrollcommand=scrollbar.set)
        
        self.items_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Load items
        self.refresh_items()
        
        # Right: Item Details & Add to Cart
        right_frame = ttk.LabelFrame(top_frame, text="2. Item Details", padding="10")
        right_frame.pack(side='right', fill='both', padx=(5, 0))
        
        # Selected Item Info
        self.selected_item_label = ttk.Label(right_frame, text="No item selected", font=('Helvetica', 10, 'bold'))
        self.selected_item_label.grid(row=0, column=0, columnspan=2, pady=(0, 10))

        self.item_cost_label_var = tk.StringVar(value="Base Cost Per Item: AED 0.00")
        ttk.Label(right_frame, textvariable=self.item_cost_label_var).grid(row=1, column=0, columnspan=2, sticky='w', pady=5)

        ttk.Label(right_frame, text="Max Qty Per Sale:").grid(row=2, column=0, sticky='w', pady=5)
        self.item_max_sell_qty_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.item_max_sell_qty_var, width=10).grid(row=2, column=1, sticky='w', pady=5)

        ttk.Label(right_frame, text="Price Multiplier:").grid(row=3, column=0, sticky='w', pady=5)
        self.item_price_multiplier_var = tk.StringVar()
        ttk.Entry(right_frame, textvariable=self.item_price_multiplier_var, width=10).grid(row=3, column=1, sticky='w', pady=5)

        ttk.Label(right_frame, text="Pricing Mode:").grid(row=4, column=0, sticky='w', pady=5)
        self.pricing_mode_var = tk.StringVar(value="multiplier")
        pricing_mode_frame = ttk.Frame(right_frame)
        pricing_mode_frame.grid(row=4, column=1, sticky='w', pady=5)
        ttk.Radiobutton(
            pricing_mode_frame,
            text="Use Multiplier",
            value="multiplier",
            variable=self.pricing_mode_var,
            command=self.on_pricing_mode_change
        ).pack(side='left')
        ttk.Radiobutton(
            pricing_mode_frame,
            text="Custom Price",
            value="manual",
            variable=self.pricing_mode_var,
            command=self.on_pricing_mode_change
        ).pack(side='left', padx=(8, 0))

        ttk.Label(right_frame, text="Quantity:").grid(row=5, column=0, sticky='w', pady=5)
        self.quantity_var = tk.StringVar(value="1")
        qty_entry = ttk.Entry(right_frame, textvariable=self.quantity_var, width=10)
        qty_entry.grid(row=5, column=1, sticky='w', pady=5)

        self.use_max_quantity_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            right_frame,
            text="Use Max Quantity Automatically",
            variable=self.use_max_quantity_var,
            command=self.on_toggle_use_max_quantity
        ).grid(row=6, column=0, columnspan=2, sticky='w', pady=(0, 5))
        
        ttk.Label(right_frame, text="Promo Qty:").grid(row=7, column=0, sticky='w', pady=5)
        self.promo_quantity_var = tk.StringVar(value="0")
        promo_entry = ttk.Entry(right_frame, textvariable=self.promo_quantity_var, width=10)
        promo_entry.grid(row=7, column=1, sticky='w', pady=5)
        
        ttk.Label(right_frame, text="Selling Price (AED):").grid(row=8, column=0, sticky='w', pady=5)
        self.selling_price_var = tk.StringVar()
        self.price_entry = ttk.Entry(right_frame, textvariable=self.selling_price_var, width=10, state='readonly')
        self.price_entry.grid(row=8, column=1, sticky='w', pady=5)
        
        self.vat_var = tk.BooleanVar(value=True)
        vat_check = ttk.Checkbutton(right_frame, text="Apply VAT (5%)", variable=self.vat_var)
        vat_check.grid(row=9, column=0, columnspan=2, sticky='w', pady=5)
        
        # Profit Preview
        self.profit_preview_var = tk.StringVar(value="")
        ttk.Label(right_frame, textvariable=self.profit_preview_var, foreground="gray").grid(row=10, column=0, columnspan=2, sticky='w', pady=5)

        self.sale_settings_status_var = tk.StringVar(value="")
        ttk.Label(right_frame, textvariable=self.sale_settings_status_var, foreground="gray").grid(row=11, column=0, columnspan=2, sticky='w', pady=(0, 5))

        ttk.Button(right_frame, text="Save Item Sale Settings", command=self.save_selected_item_sale_settings).grid(row=12, column=0, columnspan=2, pady=(0, 10), sticky='ew')
        ttk.Button(right_frame, text="Add to Cart ⬇", command=self.add_to_cart).grid(row=13, column=0, columnspan=2, pady=10, sticky='ew')
        ttk.Button(right_frame, text="Add Selected Items ⬇", command=self.add_selected_items_to_cart).grid(row=14, column=0, columnspan=2, pady=(0, 10), sticky='ew')
        
        # --- Middle Section: Shopping Cart ---
        cart_frame = ttk.LabelFrame(main_frame, text="3. Shopping Cart", padding="10")
        cart_frame.pack(fill='both', expand=True, pady=10)
        
        cart_columns = ('Item', 'Qty', 'Promo', 'Price', 'VAT', 'Total', 'Profit')
        self.cart_tree = ttk.Treeview(cart_frame, columns=cart_columns, show='headings', height=6)
        
        for col in cart_columns:
            self.cart_tree.heading(col, text=col)
            self.cart_tree.column(col, width=80 if col != 'Item' else 150)
            
        cart_scroll = ttk.Scrollbar(cart_frame, orient=tk.VERTICAL, command=self.cart_tree.yview)
        self.cart_tree.configure(yscrollcommand=cart_scroll.set)
        
        self.cart_tree.pack(side='left', fill='both', expand=True)
        cart_scroll.pack(side='right', fill='y')
        
        ttk.Button(cart_frame, text="Remove Selected", command=self.remove_from_cart).pack(anchor='e', pady=5)
        
        # --- Bottom Section: Order Details & Checkout ---
        bottom_frame = ttk.LabelFrame(main_frame, text="4. Finalize Order", padding="10")
        bottom_frame.pack(fill='x', pady=5)
        
        ttk.Label(bottom_frame, text="Customer Name:").pack(side='left')
        self.customer_var = tk.StringVar()
        self.customer_combo = ttk.Combobox(bottom_frame, textvariable=self.customer_var, width=30)
        self.customer_combo['values'] = self.inventory_manager.load_customers()
        self.customer_combo.pack(side='left', padx=10)
        
        if self.invoice_manager:
            self.generate_invoice_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(bottom_frame, text="Generate Invoice", variable=self.generate_invoice_var).pack(side='left', padx=10)
            
            # New checkbox for payment status
            self.is_paid_var = tk.BooleanVar(value=True)
            ttk.Checkbutton(bottom_frame, text="Paid (Cash)", variable=self.is_paid_var).pack(side='left', padx=10)

        # Delivery Mode & Fee
        self.delivery_mode_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(bottom_frame, text="🚚 Delivery Only", variable=self.delivery_mode_var, 
                        command=self.toggle_delivery_mode).pack(side='left', padx=10)
        
        ttk.Label(bottom_frame, text="Delivery Fee:").pack(side='left', padx=(10, 0))
        self.delivery_fee_var = tk.StringVar(value="0.00")
        self.delivery_fee_entry = ttk.Entry(bottom_frame, textvariable=self.delivery_fee_var, width=10, state='disabled')
        self.delivery_fee_entry.pack(side='left', padx=5)
        self.delivery_fee_var.trace('w', self.update_cart_display)
        
        # Totals Display
        self.cart_total_var = tk.StringVar(value="Total: AED 0.00 | Items: 0")
        ttk.Label(bottom_frame, textvariable=self.cart_total_var, font=('Helvetica', 12, 'bold')).pack(side='right', padx=20)

        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        ttk.Button(button_frame, text="CHECKOUT & PRINT", command=self.checkout, width=20).pack(side='left', padx=10)
        ttk.Button(button_frame, text="Cancel", command=self.dialog.destroy).pack(side='left', padx=10)
        
        # Bind events
        self.items_tree.bind('<<TreeviewSelect>>', self.on_item_select)
        self.item_max_sell_qty_var.trace('w', self.update_sale_settings_preview)
        self.item_price_multiplier_var.trace('w', self.update_sale_settings_preview)
        self.pricing_mode_var.trace('w', self.update_sale_settings_preview)
        self.quantity_var.trace('w', self.update_profit_calculation)
        self.promo_quantity_var.trace('w', self.update_profit_calculation)
        self.selling_price_var.trace('w', self.update_profit_calculation)
        self.vat_var.trace('w', self.update_profit_calculation)
        self.dialog.after(0, self.apply_preselected_items)

    def refresh_items(self):
        """Refresh items list"""
        for item in self.items_tree.get_children():
            self.items_tree.delete(item)
        
        for item in self.inventory_manager.items:
            if item.quantity > 0:  # Only show items with stock
                name_display = item.name
                if getattr(item, 'is_storage_item', False):
                    name_display = f"📦 {item.name} (Stored)"
                
                self.items_tree.insert('', tk.END, values=(
                    item.item_id,
                    name_display,
                    item.quantity,
                    f"AED {item.selling_price:.2f}" if item.selling_price else "Not set"
                ))

    def apply_preselected_items(self):
        """Preselect inventory rows passed from the main inventory window."""
        if not self.preselected_item_ids:
            return
        selected_rows = []
        wanted = set(self.preselected_item_ids)
        for row_id in self.items_tree.get_children():
            values = self.items_tree.item(row_id).get('values') or []
            item_id = str(values[0]) if values else ""
            if item_id in wanted:
                selected_rows.append(row_id)
        if selected_rows:
            self.items_tree.selection_set(selected_rows)
            self.items_tree.focus(selected_rows[0])
            self.items_tree.see(selected_rows[0])
            self.on_item_select(None)
    
    def toggle_delivery_mode(self):
        """Handle delivery mode toggle"""
        is_delivery = self.delivery_mode_var.get()
        if is_delivery:
            self.delivery_fee_entry.config(state='normal')
            self.selling_price_var.set("0.00")
            # If cart already has items, we might want to update them to 0 price?
            # User requirement says "invoice issued for delivery fees only", 
            # so items should probably be 0 in that invoice.
            for item in self.cart_items:
                item['price'] = 0.0
                item['total'] = 0.0
                item['vat_amount'] = 0.0
                item['grand_total'] = 0.0
        else:
            self.delivery_fee_entry.config(state='disabled')
            self.delivery_fee_var.set("0.00")
            if self.selected_item:
                self.update_sale_settings_preview()
        self._sync_selling_price_entry_state()
        
        self.update_cart_display()
        self.update_profit_calculation()

    def update_sale_settings_preview(self, *args):
        if not self.selected_item:
            self.item_cost_label_var.set("Base Cost Per Item: AED 0.00")
            self.sale_settings_status_var.set("")
            return

        cost_per_item = float(self.selected_item.get_cost_per_item() or 0.0)
        self.item_cost_label_var.set(f"Base Cost Per Item: AED {cost_per_item:.2f}")
        pricing_mode = self.pricing_mode_var.get()
        self._sync_selling_price_entry_state()
        
        # Check if item has selling_price set
        item_selling_price = getattr(self.selected_item, 'selling_price', None)
        
        try:
            if not self.delivery_mode_var.get() and pricing_mode == "multiplier":
                if item_selling_price and float(item_selling_price) > 0:
                    # Use item's set selling price
                    self.selling_price_var.set(f"{float(item_selling_price):.2f}")
                else:
                    # Fallback to multiplier logic
                    multiplier = float(self.item_price_multiplier_var.get() or "0")
                    if multiplier <= 0:
                        raise ValueError
                    selling_price = cost_per_item * multiplier
                    self.selling_price_var.set(f"{selling_price:.2f}")
            
            try:
                max_sell_qty = int(self.item_max_sell_qty_var.get() or "0")
                if max_sell_qty <= 0:
                    raise ValueError
                price_mode_label = "Multiplier price" if pricing_mode == "multiplier" else "Custom sale price"
                self.sale_settings_status_var.set(f"{price_mode_label} | Max per sale: {max_sell_qty} unit(s)")
            except ValueError:
                self.sale_settings_status_var.set("Max per sale must be a whole number > 0")
        except ValueError:
            if not self.delivery_mode_var.get() and pricing_mode == "multiplier":
                if item_selling_price and float(item_selling_price) > 0:
                    self.selling_price_var.set(f"{float(item_selling_price):.2f}")
                else:
                    self.selling_price_var.set("")
            if pricing_mode == "multiplier":
                self.sale_settings_status_var.set("Multiplier must be a number > 0")
            else:
                self.sale_settings_status_var.set("Custom sale price mode enabled")

    def _sync_selling_price_entry_state(self):
        if not hasattr(self, 'price_entry'):
            return
        if self.delivery_mode_var.get():
            self.price_entry.config(state='readonly')
            return
        if self.pricing_mode_var.get() == "manual":
            self.price_entry.config(state='normal')
        else:
            self.price_entry.config(state='readonly')

    def on_pricing_mode_change(self):
        self._sync_selling_price_entry_state()
        if not self.selected_item:
            return
        if self.pricing_mode_var.get() == "manual":
            try:
                current_price = float(getattr(self.selected_item, 'selling_price', 0.0) or 0.0)
                self.selling_price_var.set(f"{current_price:.2f}")
            except Exception:
                self.selling_price_var.set("")
        else:
            self.update_sale_settings_preview()
        self.update_profit_calculation()

    def _get_in_cart_qty_for_item(self, item_id):
        return sum(
            (existing.get('quantity', 0) or 0) + (existing.get('promo_qty', 0) or 0)
            for existing in self.cart_items
            if existing.get('item_id') == item_id
        )

    def _get_auto_quantity_for_item(self, item):
        if not item:
            return 0
        try:
            max_sell_qty = int(getattr(item, 'max_sell_qty', 999999) or 999999)
        except Exception:
            max_sell_qty = 999999
        if max_sell_qty <= 0:
            max_sell_qty = 999999
        current_stock = int(float(getattr(item, 'quantity', 0) or 0))
        in_cart_qty = self._get_in_cart_qty_for_item(getattr(item, 'item_id', ''))
        allowed_total = min(current_stock, max_sell_qty)
        return max(0, allowed_total - in_cart_qty)

    def on_toggle_use_max_quantity(self):
        if self.use_max_quantity_var.get():
            self.promo_quantity_var.set("0")
            if self.selected_item:
                auto_qty = self._get_auto_quantity_for_item(self.selected_item)
                self.quantity_var.set(str(auto_qty))

    def apply_auto_quantity_for_selected_item(self):
        if not self.selected_item:
            return 0
        auto_qty = self._get_auto_quantity_for_item(self.selected_item)
        self.quantity_var.set(str(auto_qty))
        self.promo_quantity_var.set("0")
        return auto_qty

    def save_selected_item_sale_settings(self, show_message=True):
        if not self.selected_item:
            messagebox.showwarning("Warning", "Please select exactly one item first")
            return False
        try:
            max_sell_qty = self.inventory_manager._require_positive_int(self.item_max_sell_qty_var.get(), "Max quantity per sale")
            price_multiplier = self.inventory_manager._require_positive_float(self.item_price_multiplier_var.get(), "Price multiplier")
            success = self.inventory_manager.update_item(
                self.selected_item.item_id,
                max_sell_qty=max_sell_qty,
                price_multiplier=price_multiplier
            )
            if not success:
                raise ValueError("Failed to save item sale settings")
            self.selected_item = self.inventory_manager.get_item(self.selected_item.item_id)
            self.refresh_items()
            self.update_sale_settings_preview()
            self.update_profit_calculation()
            if show_message:
                messagebox.showinfo("Success", "Item sale settings saved successfully.")
            return True
        except Exception as exc:
            if show_message:
                messagebox.showwarning("Warning", str(exc))
            return False

    def on_item_select(self, event):
        """Handle item selection"""
        selection = self.items_tree.selection()
        if not selection:
            self.selected_item = None
            self.selected_item_label.config(text="No item selected")
            self.selling_price_var.set("")
            self.item_max_sell_qty_var.set("")
            self.item_price_multiplier_var.set("")
            self.item_cost_label_var.set("Base Cost Per Item: AED 0.00")
            self.sale_settings_status_var.set("")
            self.profit_preview_var.set("")
            self._sync_selling_price_entry_state()
            return

        if len(selection) > 1:
            self.selected_item = None
            self.selected_item_label.config(text=f"Selected: {len(selection)} items")
            self.selling_price_var.set("")
            self.item_max_sell_qty_var.set("")
            self.item_price_multiplier_var.set("")
            self.item_cost_label_var.set("Base Cost Per Item: Select one item to edit")
            self.sale_settings_status_var.set("Bulk add uses each item's saved sale settings.")
            self.profit_preview_var.set("Bulk add uses each item's saved price.")
            self._sync_selling_price_entry_state()
            return

        item_id = self.items_tree.item(selection[0])['values'][0]
        self.selected_item = self.inventory_manager.get_item(item_id)
        
        if self.selected_item:
            self.selected_item_label.config(text=f"Selected: {self.selected_item.name}")
            self.item_max_sell_qty_var.set(str(getattr(self.selected_item, 'max_sell_qty', 999999)))
            self.item_price_multiplier_var.set(str(getattr(self.selected_item, 'price_multiplier', 1.0)))
            
            # Set selling price from item's attribute
            try:
                item_sp = float(getattr(self.selected_item, 'selling_price', 0.0) or 0.0)
                if item_sp > 0:
                    self.selling_price_var.set(f"{item_sp:.2f}")
            except Exception:
                pass
            
            if self.pricing_mode_var.get() == "manual":
                try:
                    self.selling_price_var.set(f"{float(getattr(self.selected_item, 'selling_price', 0.0) or 0.0):.2f}")
                except Exception:
                    self.selling_price_var.set("")
            self.update_sale_settings_preview()
            if self.use_max_quantity_var.get():
                self.apply_auto_quantity_for_selected_item()
            self.update_profit_calculation()

    def _add_item_to_cart(self, item, quantity, promo_qty, selling_price):
        """Add one inventory item to the cart after stock validation."""
        current_stock = item.quantity
        in_cart_qty = sum(
            existing['quantity'] + existing['promo_qty']
            for existing in self.cart_items
            if existing['item_id'] == item.item_id
        )

        try:
            max_sell_qty = int(getattr(item, 'max_sell_qty', 999999) or 999999)
        except Exception:
            max_sell_qty = 999999
        if max_sell_qty <= 0:
            max_sell_qty = 999999
        if (quantity + promo_qty + in_cart_qty) > max_sell_qty:
            return False, f"{item.name}: max per sale is {max_sell_qty}"

        if (quantity + promo_qty + in_cart_qty) > current_stock:
            return False, f"{item.name}: available {current_stock}, already in cart {in_cart_qty}"

        total_sales = selling_price * quantity
        vat_amount = (total_sales * 0.05) if self.vat_var.get() else 0.0
        self.cart_items.append({
            'item_id': item.item_id,
            'name': item.name,
            'quantity': quantity,
            'promo_qty': promo_qty,
            'price': selling_price,
            'vat_applied': self.vat_var.get(),
            'vat_amount': vat_amount,
            'total': total_sales,
            'grand_total': total_sales + vat_amount
        })
        return True, None
    
    def update_profit_calculation(self, *args):
        """Update profit calculation preview"""
        if not self.selected_item:
            self.profit_preview_var.set("")
            return
        
        try:
            quantity = int(self.quantity_var.get()) if self.quantity_var.get() else 0
            promo_qty = int(self.promo_quantity_var.get()) if self.promo_quantity_var.get() else 0
            selling_price = float(self.selling_price_var.get()) if self.selling_price_var.get() else 0
            
            if quantity <= 0:
                self.profit_preview_var.set("")
                return
            
            cost_per_item = self.selected_item.get_cost_per_item()
            total_cost = cost_per_item * (quantity + promo_qty)
            total_sales = selling_price * quantity
            profit = total_sales - total_cost
            
            self.profit_preview_var.set(f"Est. Profit: AED {profit:.2f}")
            
        except ValueError:
            self.profit_preview_var.set("")

    def add_to_cart(self):
        """Add selected item to cart"""
        if not self.selected_item:
            messagebox.showwarning("Warning", "Please select an item")
            return
        
        try:
            if self.use_max_quantity_var.get():
                auto_qty = self.apply_auto_quantity_for_selected_item()
                if auto_qty <= 0:
                    messagebox.showwarning("Warning", "No remaining quantity is available for this item")
                    return
            quantity = int(self.quantity_var.get())
            promo_qty = int(self.promo_quantity_var.get() or "0")
            selling_price = float(self.selling_price_var.get())
            
            if quantity <= 0:
                messagebox.showwarning("Warning", "Quantity must be > 0")
                return
            if promo_qty < 0:
                messagebox.showwarning("Warning", "Promo quantity cannot be negative")
                return
            if selling_price < 0:
                messagebox.showwarning("Warning", "Price cannot be negative")
                return
            if not self.delivery_mode_var.get() and self.pricing_mode_var.get() == "manual" and selling_price == 0:
                messagebox.showwarning("Warning", "Enter a custom selling price greater than 0")
                return

        except ValueError:
            messagebox.showwarning("Warning", "Please enter valid numeric values")
            return

        success, message = self._add_item_to_cart(self.selected_item, quantity, promo_qty, selling_price)
        if not success:
            messagebox.showwarning("Warning", f"Insufficient stock! {message}")
            return

        self.update_cart_display()
        
        # Reset inputs
        self.quantity_var.set("1")
        self.promo_quantity_var.set("0")
        # Keep price as it might be same for next item

    def add_selected_items_to_cart(self):
        """Bulk-add all currently selected items using each item's saved price."""
        selection = self.items_tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select one or more items")
            return

        try:
            quantity = int(self.quantity_var.get() or "1")
            promo_qty = int(self.promo_quantity_var.get() or "0")
        except ValueError:
            messagebox.showwarning("Warning", "Please enter valid numeric values")
            return

        if quantity <= 0:
            messagebox.showwarning("Warning", "Quantity must be > 0")
            return
        if promo_qty < 0:
            messagebox.showwarning("Warning", "Promo quantity cannot be negative")
            return

        added_count = 0
        errors = []
        for row_id in selection:
            values = self.items_tree.item(row_id).get('values') or []
            if not values:
                continue
            item = self.inventory_manager.get_item(values[0])
            if not item:
                continue
            if self.use_max_quantity_var.get():
                quantity = self._get_auto_quantity_for_item(item)
                promo_qty = 0
                if quantity <= 0:
                    errors.append(f"{item.name}: no remaining quantity available")
                    continue
            selling_price = float(item.selling_price or 0)
            success, message = self._add_item_to_cart(item, quantity, promo_qty, selling_price)
            if success:
                added_count += 1
            elif message:
                errors.append(message)

        if added_count:
            self.update_cart_display()
            self.quantity_var.set("1")
            self.promo_quantity_var.set("0")

        if errors:
            messagebox.showwarning("Warning", "Some selected items could not be added:\n" + "\n".join(errors[:10]))
        elif added_count:
            messagebox.showinfo("Success", f"Added {added_count} item(s) to cart.")
        
    def remove_from_cart(self):
        """Remove selected item from cart"""
        selection = self.cart_tree.selection()
        if not selection:
            return
        
        idx = self.cart_tree.index(selection[0])
        del self.cart_items[idx]
        self.update_cart_display()
    
    def update_cart_display(self):
        """Update cart treeview and totals"""
        for item in self.cart_tree.get_children():
            self.cart_tree.delete(item)
            
        grand_total = 0
        total_items = 0
        
        for item in self.cart_items:
            self.cart_tree.insert('', 'end', values=(
                item['name'],
                item['quantity'],
                item['promo_qty'],
                f"{item['price']:.2f}",
                f"{item['vat_amount']:.2f}",
                f"{item['grand_total']:.2f}",
                "" # Profit calculated at checkout
            ))
            grand_total += item['grand_total']
            total_items += item['quantity']
            
        # Add delivery fee if in delivery mode
        delivery_fee = 0.0
        if self.delivery_mode_var.get():
            try:
                delivery_fee = float(self.delivery_fee_var.get())
                if delivery_fee > 0:
                    self.cart_tree.insert('', 'end', values=(
                        "🚚 Delivery Fee",
                        "1",
                        "0",
                        f"{delivery_fee:.2f}",
                        "0.00",
                        f"{delivery_fee:.2f}",
                        ""
                    ))
                    grand_total += delivery_fee
            except ValueError:
                pass
            
        self.cart_total_var.set(f"Total: AED {grand_total:.2f} | Items: {total_items}")

    def checkout(self):
        """Process all items in cart"""
        if not self.cart_items:
            messagebox.showwarning("Warning", "Cart is empty!")
            return
        
        customer_name = self.customer_var.get().strip()
        if not customer_name:
            messagebox.showwarning("Warning", "Please enter customer name")
            return
        
        if not messagebox.askyesno("Confirm Checkout", f"Process sale for {len(self.cart_items)} items?"):
            return
            
        sale_records = []
        errors = []
        
        # Process each item
        for item in self.cart_items:
            success, result = self.inventory_manager.record_sale(
                item['item_id'], 
                item['quantity'], 
                item['price'], 
                customer_name,
                promotion_quantity=item['promo_qty'],
                vat_applied=item['vat_applied'],
                vat_rate=0.05
            )
            
            if success:
                sale_records.append(result)
            else:
                errors.append(f"{item['name']}: {result}")
        
        if errors:
            messagebox.showerror("Partial Error", "Some items failed:\n" + "\n".join(errors))
            
        if sale_records:
            # Save customer name for future use
            self.inventory_manager.save_customer(customer_name)
            
            # Generate single invoice for all successful sales
            if self.invoice_manager and self.generate_invoice_var.get():
                is_paid = self.is_paid_var.get()
                is_delivery_only = self.delivery_mode_var.get()
                delivery_fee = 0.0
                try:
                    delivery_fee = float(self.delivery_fee_var.get())
                except ValueError:
                    pass
                
                invoice_id = self.inventory_manager.create_sales_invoice_multi(
                    sale_records, 
                    customer_name, 
                    is_paid=is_paid,
                    delivery_fee=delivery_fee,
                    is_delivery_only=is_delivery_only
                )
                if invoice_id:
                    status_msg = "Paid" if is_paid else "Not Paid"
                    messagebox.showinfo("Success", f"Order processed successfully!\nInvoice generated: {invoice_id}\nStatus: {status_msg}")
                    self.dialog.destroy()
                else:
                    messagebox.showinfo("Success", "Order processed, but invoice generation failed.")
                    self.dialog.destroy()
            else:
                messagebox.showinfo("Success", "Order processed successfully (No Invoice).")
                self.dialog.destroy()

class InventoryReportsDialog:
    def __init__(self, parent, inventory_manager, preselected_item_id=None):
        self.inventory_manager = inventory_manager
        self.preselected_item_id = preselected_item_id
        try:
            reports_debug_log(self.inventory_manager.data_folder, "InventoryReportsDialog __init__ start")
        except Exception:
            pass
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("ASSISTEM - Inventory Intelligence & Reports")
        self.dialog.geometry("1100x850")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (height // 2)
        self.dialog.geometry(f'+{x}+{y}')
        try:
            reports_debug_log(self.inventory_manager.data_folder, "InventoryReportsDialog setup_ui start")
            self.setup_ui()
            reports_debug_log(self.inventory_manager.data_folder, "InventoryReportsDialog setup_ui ok")
        except Exception as e:
            try:
                reports_debug_log(self.inventory_manager.data_folder, f"InventoryReportsDialog setup_ui error: {e}")
            except Exception:
                pass
            fallback = tk.Frame(self.dialog, padx=20, pady=20)
            fallback.pack(fill="both", expand=True)
            ttk.Label(fallback, text="Inventory Reports could not be loaded.", font=('Helvetica', 14, 'bold')).pack(pady=(0, 10))
            ttk.Label(fallback, text=str(e), foreground="red", wraplength=800, justify="left").pack()


class SupplierActivityReportDialog:
    def __init__(self, parent, inventory_manager):
        self.inventory_manager = inventory_manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Supplier Activity Report")
        self.dialog.geometry("1200x850")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)

        self._rows_cache = []

        self.setup_ui()
        try:
            self.dialog.after(0, self.generate_report)
        except Exception:
            pass

    def setup_ui(self):
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)

        header = ttk.Frame(main_container)
        header.pack(fill='x', padx=12, pady=(12, 6))
        ttk.Label(header, text="Supplier Activity Report", font=('Helvetica', 16, 'bold')).pack(side='left')

        controls = ttk.LabelFrame(main_container, text="Filters", padding=10)
        controls.pack(fill='x', padx=12, pady=6)

        suppliers = sorted({(getattr(i, 'supplier', '') or '').strip() or 'Unknown Supplier' for i in (self.inventory_manager.items or [])})
        suppliers = ['All Suppliers'] + suppliers

        ttk.Label(controls, text="Supplier:").grid(row=0, column=0, sticky='w')
        self.supplier_var = tk.StringVar(value='All Suppliers')
        self.supplier_combo = ttk.Combobox(controls, textvariable=self.supplier_var, values=suppliers, width=38, state='readonly')
        self.supplier_combo.grid(row=0, column=1, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="Start Date:").grid(row=0, column=2, sticky='w')
        self.start_date_var = tk.StringVar(value="")
        self.start_entry = ttk.Entry(controls, textvariable=self.start_date_var, width=12)
        self.start_entry.grid(row=0, column=3, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="End Date:").grid(row=0, column=4, sticky='w')
        self.end_date_var = tk.StringVar(value="")
        self.end_entry = ttk.Entry(controls, textvariable=self.end_date_var, width=12)
        self.end_entry.grid(row=0, column=5, sticky='w', padx=(6, 18))

        btns = ttk.Frame(controls)
        btns.grid(row=0, column=6, sticky='e')
        ttk.Button(btns, text="Generate", command=self.generate_report).pack(side='left', padx=4)
        ttk.Button(btns, text="Export PDF", command=self.export_pdf).pack(side='left', padx=4)
        ttk.Button(btns, text="Close", command=self.dialog.destroy).pack(side='left', padx=4)

        controls.columnconfigure(6, weight=1)

        results = ttk.LabelFrame(main_container, text="Results", padding=10)
        results.pack(fill='both', expand=True, padx=12, pady=(6, 12))

        columns = (
            'Supplier', 'Item ID', 'Item Name', 'Batch', 'Stored', 'Entry Date',
            'Sold Date', 'Sold To', 'Sold Qty', 'Current Qty', 'Invoice ID'
        )
        self.tree = ttk.Treeview(results, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, stretch=True)
        self.tree.column('Supplier', width=170)
        self.tree.column('Item Name', width=240)
        self.tree.column('Sold To', width=170)
        self.tree.column('Invoice ID', width=140)

        yscroll = ttk.Scrollbar(results, orient='vertical', command=self.tree.yview)
        xscroll = ttk.Scrollbar(results, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.pack(side='top', fill='both', expand=True)
        yscroll.pack(side='right', fill='y')
        xscroll.pack(side='bottom', fill='x')

        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(main_container, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w')
        status.pack(fill='x', side='bottom')

    def _parse_date(self, s):
        s = (s or '').strip()
        if not s:
            return None
        try:
            return datetime.strptime(s, "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            return s

    def generate_report(self):
        start = self._parse_date(self.start_date_var.get())
        end = self._parse_date(self.end_date_var.get())
        supplier = self.supplier_var.get()

        rows = self.inventory_manager.get_supplier_activity_rows(start_date=start, end_date=end, supplier_filter=supplier)
        self._rows_cache = rows

        for r in self.tree.get_children():
            self.tree.delete(r)

        for row in rows:
            stored_txt = "Yes" if row.get('is_storage_item') else "No"
            self.tree.insert('', 'end', values=(
                row.get('supplier', ''),
                row.get('item_id', ''),
                row.get('item_name', ''),
                row.get('batch_number', ''),
                stored_txt,
                row.get('entry_date', ''),
                row.get('sold_date', ''),
                row.get('sold_to', ''),
                row.get('sold_qty', 0),
                row.get('current_qty', 0),
                row.get('invoice_id', ''),
            ))

        self.status_var.set(f"Rows: {len(rows)}")

    def export_pdf(self):
        rows = list(self._rows_cache or [])
        if not rows:
            messagebox.showwarning("Export PDF", "No rows to export.")
            return

        reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
        try:
            os.makedirs(reports_folder, exist_ok=True)
        except Exception:
            pass

        filename = f"Supplier_Activity_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        filepath = os.path.join(reports_folder, filename)

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor("#2c3e50"))
        subtitle_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=9, alignment=1, textColor=colors.grey)

        doc = SimpleDocTemplate(filepath, pagesize=landscape(A4), leftMargin=0.35 * inch, rightMargin=0.35 * inch, topMargin=0.4 * inch, bottomMargin=0.4 * inch)
        elements = []
        elements.extend(build_pdf_logo_flowables(self.inventory_manager.data_folder, width=1.0 * inch, height=1.0 * inch, spacer_height=0.08 * inch))
        elements.append(Paragraph("Supplier Activity Report", title_style))
        elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_style))
        elements.append(Spacer(1, 10))

        grouped = {}
        for r in rows:
            grouped.setdefault(r.get('supplier', 'Unknown Supplier'), []).append(r)

        for supplier, sup_rows in grouped.items():
            elements.append(Paragraph(f"Supplier: {supplier}", ParagraphStyle('H2', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#0b5394'))))
            data = [["Item", "Entry", "Sold", "Sold To", "Sold Qty", "Current Qty", "Invoice"]]
            for r in sup_rows:
                item_label = r.get('item_name', '')
                if r.get('is_storage_item'):
                    item_label = f"[Stored] {item_label}"
                if r.get('batch_number'):
                    item_label = f"{item_label} ({r.get('batch_number')})"
                data.append([
                    item_label,
                    r.get('entry_date', ''),
                    r.get('sold_date', ''),
                    r.get('sold_to', ''),
                    str(r.get('sold_qty', 0)),
                    str(r.get('current_qty', 0)),
                    r.get('invoice_id', ''),
                ])

            table = Table(data, colWidths=[260, 85, 85, 160, 65, 70, 120])
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
            elements.append(Spacer(1, 12))

        try:
            doc.build(elements)
            try:
                webbrowser.open("file://" + os.path.abspath(filepath))
            except Exception:
                pass
            messagebox.showinfo("Export PDF", f"Report exported to:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export PDF", f"Failed to export PDF:\n{e}")
    
    def setup_ui(self):
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)

        scroll_frame = ScrollableFrame(main_container)
        scroll_frame.pack(fill='both', expand=True)

        main_frame = scroll_frame.scrollable_frame

        ttk.Label(main_frame, text="Inventory Reports", font=('Helvetica', 16, 'bold')).pack(pady=(0, 20))

        controls_frame = ttk.LabelFrame(main_frame, text="Report Options", padding="10")
        controls_frame.pack(fill='x', pady=5)

        date_frame = ttk.Frame(controls_frame)
        date_frame.pack(fill='x', pady=5)

        ttk.Label(date_frame, text="Start Date:").grid(row=0, column=0, padx=(0, 5))
        self.start_date = ttk.Entry(date_frame, width=12)
        self.start_date.grid(row=0, column=1, padx=(0, 15))
        self.start_date.insert(0, (datetime.now().replace(day=1)).strftime("%Y-%m-%d"))

        ttk.Label(date_frame, text="End Date:").grid(row=0, column=2, padx=(0, 5))
        self.end_date = ttk.Entry(date_frame, width=12)
        self.end_date.grid(row=0, column=3, padx=(0, 15))
        self.end_date.insert(0, datetime.now().strftime("%Y-%m-%d"))

        report_frame = ttk.Frame(controls_frame)
        report_frame.pack(fill='x', pady=5)

        ttk.Label(report_frame, text="Report Type:").grid(row=0, column=0, padx=(0, 5))
        self.report_type = tk.StringVar(value="full_summary")
        ttk.Radiobutton(report_frame, text="Full Summary", variable=self.report_type, value="full_summary").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(report_frame, text="System Health", variable=self.report_type, value="system_health").grid(row=0, column=2, padx=5)
        ttk.Radiobutton(report_frame, text="Movement", variable=self.report_type, value="movement").grid(row=0, column=3, padx=5)
        ttk.Radiobutton(report_frame, text="Sales", variable=self.report_type, value="sales").grid(row=0, column=4, padx=5)

        report_frame2 = ttk.Frame(controls_frame)
        report_frame2.pack(fill='x', pady=5)
        ttk.Radiobutton(report_frame2, text="Stock", variable=self.report_type, value="stock").grid(row=0, column=0, padx=5)
        ttk.Radiobutton(report_frame2, text="Low Stock", variable=self.report_type, value="low_stock").grid(row=0, column=1, padx=5)
        ttk.Radiobutton(report_frame2, text="Costs", variable=self.report_type, value="cost_analysis").grid(row=0, column=2, padx=5)
        ttk.Radiobutton(report_frame2, text="New Items", variable=self.report_type, value="new_items").grid(row=0, column=3, padx=5)
        ttk.Radiobutton(report_frame2, text="Per Item", variable=self.report_type, value="per_item").grid(row=0, column=4, padx=5)
        ttk.Radiobutton(report_frame2, text="Custom Combo", variable=self.report_type, value="custom_multi").grid(row=0, column=5, padx=5)

        self.custom_frame = ttk.LabelFrame(controls_frame, text="Custom Combination", padding="5")
        self.custom_sales = tk.BooleanVar(value=True)
        self.custom_costs = tk.BooleanVar(value=False)
        self.custom_movement = tk.BooleanVar(value=False)
        self.custom_low_stock = tk.BooleanVar(value=False)
        self.custom_new_items = tk.BooleanVar(value=False)
        self.custom_system_health = tk.BooleanVar(value=False)
        ttk.Checkbutton(self.custom_frame, text="Sales", variable=self.custom_sales).grid(row=0, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(self.custom_frame, text="Costs", variable=self.custom_costs).grid(row=0, column=1, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(self.custom_frame, text="Movement", variable=self.custom_movement).grid(row=0, column=2, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(self.custom_frame, text="Low Stock", variable=self.custom_low_stock).grid(row=1, column=0, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(self.custom_frame, text="New Items", variable=self.custom_new_items).grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttk.Checkbutton(self.custom_frame, text="System Health", variable=self.custom_system_health).grid(row=1, column=2, padx=5, pady=2, sticky="w")

        item_frame = ttk.Frame(controls_frame)
        item_frame.pack(fill='x', pady=5)
        ttk.Label(item_frame, text="Specific Item(s):").grid(row=0, column=0, padx=(0, 5), sticky="nw")
        self.item_selector = tk.Listbox(item_frame, selectmode='extended', height=6, exportselection=False)
        self.item_selector.grid(row=0, column=1, padx=(0, 0), sticky="nsew")
        item_frame.columnconfigure(1, weight=1)
        item_scroll = ttk.Scrollbar(item_frame, orient="vertical", command=self.item_selector.yview)
        item_scroll.grid(row=0, column=2, sticky="ns", padx=(5, 15))
        self.item_selector.configure(yscrollcommand=item_scroll.set)

        button_frame = ttk.Frame(controls_frame)
        button_frame.pack(fill='x', pady=10)

        ttk.Button(button_frame, text="Generate Report", command=self.refresh_and_generate).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Quick Summary", command=self.generate_quick_summary).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Inventory Management Report", command=self.generate_inventory_management_report).pack(side='left', padx=5)

        export_html_cmd = getattr(self, "export_report", None)
        if export_html_cmd is None:
            def _no_export_html():
                messagebox.showerror("Error", "Export to HTML is not available.")
            export_html_cmd = _no_export_html
        ttk.Button(button_frame, text="Export to HTML", command=export_html_cmd).pack(side='left', padx=5)

        export_pdf_cmd = getattr(self, "export_to_pdf", None)
        if export_pdf_cmd is None:
            def _no_export_pdf():
                messagebox.showerror("Error", "Export to PDF is not available.")
            export_pdf_cmd = _no_export_pdf
        ttk.Button(button_frame, text="Export to PDF", command=export_pdf_cmd).pack(side='left', padx=5)

        ttk.Button(button_frame, text="Close", command=self.dialog.destroy).pack(side='left', padx=5)

        report_display_frame = ttk.LabelFrame(main_frame, text="Report Results", padding="10")
        report_display_frame.pack(fill='both', expand=True, pady=10)

        self.report_text = scrolledtext.ScrolledText(report_display_frame, height=20, font=('Consolas', 9))
        self.report_text.pack(fill='both', expand=True)

        self.status_var = tk.StringVar(value="Ready.")
        status_bar = ttk.Label(self.dialog, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill='x', side='bottom')

        self.update_item_selector()
        try:
            self.report_type.trace_add("write", self.on_report_type_change)
        except Exception:
            pass

    def generate_quick_summary(self):
        selections = []
        try:
            indices = self.item_selector.curselection()
            selections = [self.item_selector.get(i) for i in indices]
        except Exception:
            selections = []
        if not selections:
            messagebox.showwarning("Quick Summary", "Please select one or more items for the summary.")
            return
        items = []
        for raw in selections:
            try:
                item_id = raw.split('(')[-1].strip(')')
                item = next((i for i in self.inventory_manager.items if i.item_id == item_id), None)
                if item:
                    items.append(item)
            except Exception:
                continue
        if not items:
            messagebox.showwarning("Quick Summary", "Selected items could not be resolved from inventory.")
            return
        lines = []
        lines.append("Inventory Summary Report")
        lines.append("HopePharma Inventory")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        lines.append("=" * 100)
        lines.append("")
        for item in items:
            name = item.name or ""
            batch = item.batch_number or ""
            qty = str(item.quantity)
            date_entered = item.created_date or ""
            warehouse = "International City, Morroco Cluster, I-16 Shop 8"
            lines.append(f"Product Name     : {name}")
            lines.append(f"Batch No         : {batch}")
            lines.append(f"Quantity         : {qty}")
            lines.append(f"Date of Entry    : {date_entered}")
            lines.append(f"Warehouse Address: {warehouse}")
            lines.append("-" * 60)
            lines.append("")
        content = "\n".join(lines)
        try:
            self.report_text.delete("1.0", tk.END)
            self.report_text.insert("1.0", content)
            try:
                self.report_type.set("quick_summary")
            except Exception:
                pass
        except Exception:
            pass

    def generate_inventory_management_report(self):
        try:
            start_date = self.start_date.get()
            end_date = self.end_date.get()
        except Exception:
            start_date = ""
            end_date = ""
        items = list(self.inventory_manager.items or [])
        suppliers = {}
        for item in items:
            try:
                created = getattr(item, "created_date", "") or ""
                if start_date and end_date:
                    if not (start_date <= created <= end_date):
                        continue
                supplier = (item.supplier or "").strip() or "Unknown Supplier"
                suppliers.setdefault(supplier, []).append(item)
            except Exception:
                continue
        lines = []
        lines.append("Inventory Management Report")
        lines.append("HopePharma Inventory")
        if start_date and end_date:
            lines.append(f"Period: {start_date} to {end_date}")
        lines.append("=" * 100)
        lines.append("")
        for supplier, supp_items in sorted(suppliers.items(), key=lambda x: x[0].lower()):
            total_cost = 0.0
            invoice_date = None
            for item in supp_items:
                try:
                    total_cost += float(getattr(item, "total_cost", 0) or 0)
                    inv_date = getattr(item, "invoice_date", None) or getattr(item, "created_date", "") or ""
                    if inv_date:
                        if invoice_date is None or inv_date < invoice_date:
                            invoice_date = inv_date
                except Exception:
                    continue
            lines.append(f"Supplier Name : {supplier}")
            lines.append(f"Invoice Value : AED {total_cost:,.2f}")
            lines.append(f"Invoice Date  : {invoice_date or '-'}")
            lines.append("")
            header = f"{'Product Name':30} | {'Batch No':12} | {'Quantity':8} | {'Total Stock':11} | {'Date Added':12}"
            lines.append(header)
            lines.append("-" * len(header))
            for item in supp_items:
                try:
                    name = (item.name or "")[:30]
                    batch = (item.batch_number or "")[:12]
                    qty = int(getattr(item, "quantity", 0) or 0)
                    date_added = getattr(item, "created_date", "") or ""
                    total_stock = qty
                    try:
                        data = self.inventory_manager.get_sales_report("2000-01-01", end_date or "2099-12-31")
                        sold = 0
                        for s in data["sales"]:
                            if getattr(s, "item_id", None) == item.item_id:
                                sold += s.quantity + getattr(s, "promotion_quantity", 0)
                        total_stock = qty + sold
                    except Exception:
                        pass
                    line = f"{name:30} | {batch:12} | {qty:8d} | {total_stock:11d} | {date_added:12}"
                    lines.append(line)
                except Exception:
                    continue
            lines.append("-" * len(header))
            lines.append("")
        if not suppliers:
            lines.append("No inventory items found for the selected period.")
        content = "\n".join(lines)
        try:
            self.report_text.delete("1.0", tk.END)
            self.report_text.insert("1.0", content)
            try:
                self.report_type.set("inventory_management")
            except Exception:
                pass
        except Exception:
            pass

    def refresh_and_generate(self):
        """Refresh inventory data from disk and then generate report"""
        self.status_var.set("Refreshing data from disk...")
        self.dialog.update_idletasks()
        self.inventory_manager.load_inventory()
        self.inventory_manager.load_sales()
        self.update_item_selector()
        self.generate_report()
        self.status_var.set("Ready. Report generated.")

    def on_report_type_change(self, *args):
        """Show/hide custom combination options based on report type"""
        try:
            if self.report_type.get() == "custom_multi":
                self.custom_frame.pack(fill='x', pady=5)
            else:
                self.custom_frame.forget()
        except Exception:
            pass

    def update_item_selector(self):
        """Update the item combobox with current inventory items"""
        items = [f"{item.name} ({item.item_id})" for item in self.inventory_manager.items]
        items = sorted(items)
        try:
            self.item_selector.delete(0, tk.END)
            for entry in items:
                self.item_selector.insert(tk.END, entry)
            if items:
                self.item_selector.selection_clear(0, tk.END)
                self.item_selector.selection_set(0)
        except Exception:
            pass

    def generate_report(self):
        """Proxy to shared report generation logic"""
        return InventorySalesDialog.generate_report(self)

    def generate_sales_report(self, start_date, end_date):
        return InventorySalesDialog.generate_sales_report(self, start_date, end_date)

    def generate_movement_report(self, start_date, end_date):
        return InventorySalesDialog.generate_movement_report(self, start_date, end_date)

    def generate_full_summary_report(self, start_date, end_date):
        return InventorySalesDialog.generate_full_summary_report(self, start_date, end_date)

    def generate_system_health_report(self, start_date, end_date):
        return InventorySalesDialog.generate_system_health_report(self, start_date, end_date)

    def generate_new_items_report(self, start_date, end_date):
        return InventorySalesDialog.generate_new_items_report(self, start_date, end_date)

    def generate_stock_report(self):
        return InventorySalesDialog.generate_stock_report(self)

    def generate_per_item_report(self):
        return InventorySalesDialog.generate_per_item_report(self)

    def generate_low_stock_report(self):
        return InventorySalesDialog.generate_low_stock_report(self)

    def generate_cost_analysis_report(self):
        return InventorySalesDialog.generate_cost_analysis_report(self)

    def export_report(self):
        return InventorySalesDialog.export_report(self)

    def export_to_pdf(self):
        return InventorySalesDialog.export_to_pdf(self)

    def setup_ui(self):
        main_container = tk.Frame(self.dialog)
        main_container.pack(fill='both', expand=True)

        header = ttk.Frame(main_container)
        header.pack(fill='x', padx=12, pady=(12, 6))
        ttk.Label(header, text="Supplier Activity Report", font=('Helvetica', 16, 'bold')).pack(side='left')

        controls = ttk.LabelFrame(main_container, text="Filters", padding=10)
        controls.pack(fill='x', padx=12, pady=6)

        suppliers = sorted({(getattr(i, 'supplier', '') or '').strip() or 'Unknown Supplier' for i in (self.inventory_manager.items or [])})
        suppliers = ['All Suppliers'] + suppliers

        ttk.Label(controls, text="Supplier:").grid(row=0, column=0, sticky='w')
        self.supplier_var = tk.StringVar(value='All Suppliers')
        self.supplier_combo = ttk.Combobox(controls, textvariable=self.supplier_var, values=suppliers, width=38, state='readonly')
        self.supplier_combo.grid(row=0, column=1, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="Start Date:").grid(row=0, column=2, sticky='w')
        self.start_date_var = tk.StringVar(value="")
        self.start_entry = ttk.Entry(controls, textvariable=self.start_date_var, width=12)
        self.start_entry.grid(row=0, column=3, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="End Date:").grid(row=0, column=4, sticky='w')
        self.end_date_var = tk.StringVar(value="")
        self.end_entry = ttk.Entry(controls, textvariable=self.end_date_var, width=12)
        self.end_entry.grid(row=0, column=5, sticky='w', padx=(6, 18))

        btns = ttk.Frame(controls)
        btns.grid(row=0, column=6, sticky='e')
        ttk.Button(btns, text="Generate", command=self.generate_report).pack(side='left', padx=4)
        ttk.Button(btns, text="Export PDF", command=self.export_pdf).pack(side='left', padx=4)
        ttk.Button(btns, text="Close", command=self.dialog.destroy).pack(side='left', padx=4)

        controls.columnconfigure(6, weight=1)

        results = ttk.LabelFrame(main_container, text="Results", padding=10)
        results.pack(fill='both', expand=True, padx=12, pady=(6, 12))

        columns = ('Entry Date', 'Sold Date', 'Sold To', 'Sold Qty', 'Invoice ID', 'Current Qty')
        self.tree = ttk.Treeview(results, columns=columns, show='tree headings')
        self.tree.heading('#0', text='Supplier / Item')
        self.tree.column('#0', width=420, stretch=True)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=130, stretch=True)
        self.tree.column('Sold To', width=220)
        self.tree.column('Invoice ID', width=150)

        yscroll = ttk.Scrollbar(results, orient='vertical', command=self.tree.yview)
        xscroll = ttk.Scrollbar(results, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        self.tree.pack(side='top', fill='both', expand=True)
        yscroll.pack(side='right', fill='y')
        xscroll.pack(side='bottom', fill='x')

        self.status_var = tk.StringVar(value="Ready")
        status = ttk.Label(main_container, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w')
        status.pack(fill='x', side='bottom')

    def _parse_date(self, s):
        s = (s or '').strip()
        if not s:
            return None
        try:
            return datetime.strptime(s, "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            return s

    def generate_report(self):
        start = self._parse_date(self.start_date_var.get())
        end = self._parse_date(self.end_date_var.get())
        supplier = self.supplier_var.get()

        rows = self.inventory_manager.get_supplier_activity_rows(start_date=start, end_date=end, supplier_filter=supplier)
        self._rows_cache = rows

        for r in self.tree.get_children():
            self.tree.delete(r)

        supplier_nodes = {}
        item_nodes = {}
        for row in rows:
            sup = row.get('supplier', 'Unknown Supplier')
            sup_id = supplier_nodes.get(sup)
            if not sup_id:
                sup_id = self.tree.insert('', 'end', text=f"Supplier: {sup}", open=True, values=('', '', '', '', '', ''))
                supplier_nodes[sup] = sup_id

            item_key = (sup, row.get('item_id', ''))
            it_id = item_nodes.get(item_key)
            if not it_id:
                item_label = row.get('item_name', '')
                if row.get('is_storage_item'):
                    item_label = f"📦 {item_label}"
                if row.get('batch_number'):
                    item_label = f"{item_label} ({row.get('batch_number')})"
                item_label = f"{item_label} [{row.get('item_id', '')}]"
                it_id = self.tree.insert(sup_id, 'end', text=item_label, open=True, values=(
                    row.get('entry_date', ''),
                    '',
                    '',
                    '',
                    '',
                    row.get('current_qty', 0),
                ))
                item_nodes[item_key] = it_id

            if row.get('sold_date'):
                sale_text = f"Sold to: {row.get('sold_to', '')}"
                self.tree.insert(it_id, 'end', text=sale_text, values=(
                    row.get('entry_date', ''),
                    row.get('sold_date', ''),
                    row.get('sold_to', ''),
                    row.get('sold_qty', 0),
                    row.get('invoice_id', ''),
                    row.get('current_qty', 0),
                ))
            else:
                self.tree.insert(it_id, 'end', text="Not sold", values=(
                    row.get('entry_date', ''),
                    '',
                    '',
                    0,
                    '',
                    row.get('current_qty', 0),
                ))

        self.status_var.set(f"Suppliers: {len(supplier_nodes)} | Rows: {len(rows)}")

class InventoryInvoicesDialog:
    def __init__(self, parent, inventory_manager, invoice_manager=None, on_data_changed=None):
        self.inventory_manager = inventory_manager
        self.invoice_manager = invoice_manager
        self.on_data_changed = on_data_changed
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Inventory Invoices")
        self.dialog.geometry("900x600")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        self.setup_ui()
        self.refresh_invoices()
    
    def setup_ui(self):
        attach_clock_label(self.dialog)
        main = ttk.Frame(self.dialog, padding="10")
        main.pack(fill='both', expand=True)
        toolbar = ttk.Frame(main)
        toolbar.pack(fill='x', pady=(0,10))
        ttk.Button(toolbar, text="Refresh", command=self.refresh_invoices).pack(side='left', padx=5)
        ttk.Button(toolbar, text="View / Print PDF", command=self.print_selected).pack(side='left', padx=5)
        ttk.Button(toolbar, text="Delete", command=self.delete_selected).pack(side='left', padx=5)
        ttk.Button(toolbar, text="↩ Return Items", command=self.return_selected).pack(side='left', padx=5)
        columns = ('ID','Date','Client','Total','Status')
        self.tree = ttk.Treeview(main, columns=columns, show='headings', height=16)
        for c in columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=140 if c in ('Client','ID') else 100)
        yscroll = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side='left', fill='both', expand=True)
        yscroll.pack(side='right', fill='y')
    
    def load_inventory_invoices(self):
        data = []
        try:
            if os.path.exists(self.inventory_manager.invoices_file):
                with open(self.inventory_manager.invoices_file, 'r') as f:
                    data = json.load(f) or []
        except Exception:
            data = []
        inv_invoices = []
        for inv in data:
            try:
                if isinstance(inv, dict) and (inv.get('source') == 'inventory' or str(inv.get('invoice_id','')).startswith('HPMI')):
                    inv_invoices.append(inv)
            except Exception:
                pass
        return inv_invoices
    
    def refresh_invoices(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        invoices = self.load_inventory_invoices()
        for inv in sorted(invoices, key=lambda x: x.get('date',''), reverse=True):
            try:
                self.tree.insert('', 'end', values=(
                    inv.get('invoice_id'),
                    inv.get('date'),
                    inv.get('client_name'),
                    f"{float(inv.get('grand_total') or 0):,.2f}",
                    inv.get('status','')
                ))
            except Exception:
                pass
    
    def _get_selected_invoice(self):
        sel = self.tree.selection()
        if not sel:
            return None
        inv_id = self.tree.item(sel[0])['values'][0]
        for inv in self.load_inventory_invoices():
            if inv.get('invoice_id') == inv_id:
                return inv
        return None
    
    def print_selected(self):
        inv = self._get_selected_invoice()
        if not inv:
            messagebox.showwarning("Warning","Select an invoice to print")
            return
        try:
            from hope_pharma_complete import EnhancedPDFGenerator
            # Pass the invoices folder explicitly to ensure correct location
            invoices_folder = os.path.join(self.inventory_manager.data_folder, "HopePharmaInvoices")
            gen = EnhancedPDFGenerator(output_folder=invoices_folder)
            pdf_path = gen.generate_invoice_pdf(inv)
            messagebox.showinfo("Success", f"PDF generated:\n{pdf_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {e}")
    
    def delete_selected(self):
        inv = self._get_selected_invoice()
        if not inv:
            messagebox.showwarning("Warning","Select an invoice to delete")
            return
        inv_id = inv.get('invoice_id')
        if not messagebox.askyesno("Confirm Delete", f"Delete inventory invoice {inv_id}?"):
            return
        try:
            # Prefer full app balance reversal if available
            if self.invoice_manager and hasattr(self.invoice_manager, 'delete_invoice_with_balance_adjustment'):
                ok, msg = self.invoice_manager.delete_invoice_with_balance_adjustment(inv_id)
                if not ok:
                    messagebox.showerror("Error", msg)
                    return
            else:
                # Reverse balance using BalanceManager
                try:
                    from balance_manager import BalanceManager
                    bm = BalanceManager(self.inventory_manager.data_folder)
                    amount = float(inv.get('total_paid', inv.get('grand_total', 0)) or 0)
                    account = inv.get('payment_method') or 'Cash'
                    if amount > 0:
                        ok_bm, msg_bm = bm.reverse_invoice_payment(inv, amount, account)
                        if not ok_bm:
                            raise Exception(msg_bm)
                except Exception as e:
                    # If balance reversal fails, continue but warn
                    print(f"Balance reversal warning: {e}")
                ok_sales, _ = self.inventory_manager.reverse_invoice_sales(inv_id)
                # Remove from invoices file
                data = []
                if os.path.exists(self.inventory_manager.invoices_file):
                    with open(self.inventory_manager.invoices_file, 'r') as f:
                        data = json.load(f) or []
                    data = [i for i in data if not (isinstance(i, dict) and i.get('invoice_id') == inv_id)]
                    with open(self.inventory_manager.invoices_file, 'w') as f:
                        json.dump(data, f, indent=2)
            self.refresh_invoices()
            messagebox.showinfo("Success", f"Deleted {inv_id}")
            try:
                if callable(self.on_data_changed):
                    self.on_data_changed()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete: {e}")
    
    def return_selected(self):
        inv = self._get_selected_invoice()
        if not inv:
            messagebox.showwarning("Warning","Select an invoice to return")
            return
        inv_id = inv.get('invoice_id')
        if not messagebox.askyesno("Confirm Return", f"Return items for invoice {inv_id}?"):
            return
        try:
            try:
                from balance_manager import BalanceManager
                bm = BalanceManager(self.inventory_manager.data_folder)
                amount = float(inv.get('total_paid', inv.get('grand_total', 0)) or 0)
                account = inv.get('payment_method') or 'Cash'
                if amount > 0:
                    ok_bm, msg_bm = bm.reverse_invoice_payment(inv, amount, account)
                    if not ok_bm:
                        raise Exception(msg_bm)
            except Exception as e:
                print(f"Balance reversal warning: {e}")
            ok_sales, msg_sales = self.inventory_manager.reverse_invoice_sales(inv_id)
            data = []
            if os.path.exists(self.inventory_manager.invoices_file):
                with open(self.inventory_manager.invoices_file, 'r') as f:
                    data = json.load(f) or []
                changed = False
                for i in data:
                    if isinstance(i, dict) and i.get('invoice_id') == inv_id:
                        i['status'] = 'Returned'
                        i['return_processed'] = True
                        i['return_date'] = datetime.now().strftime("%Y-%m-%d")
                        changed = True
                        break
                if changed:
                    with open(self.inventory_manager.invoices_file, 'w') as f:
                        json.dump(data, f, indent=2)
            self.refresh_invoices()
            try:
                if callable(self.on_data_changed):
                    self.on_data_changed()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to return items: {e}")


class InventorySalesDialog:
    def __init__(self, parent, inventory_manager, on_data_changed=None):
        self.inventory_manager = inventory_manager
        self.on_data_changed = on_data_changed
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Saved Sales Records")
        self.dialog.geometry("950x600")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        self.setup_ui()
        self.refresh_sales()

    def setup_ui(self):
        attach_clock_label(self.dialog)
        main = ttk.Frame(self.dialog, padding="10")
        main.pack(fill='both', expand=True)

        toolbar = ttk.Frame(main)
        toolbar.pack(fill='x', pady=(0, 10))

        ttk.Button(toolbar, text="Refresh", command=self.refresh_sales).pack(side='left', padx=5)
        ttk.Button(toolbar, text="Delete Selected", command=self.delete_selected).pack(side='left', padx=5)
        ttk.Button(toolbar, text="Return Selected", command=self.return_selected).pack(side='left', padx=5)
        ttk.Button(toolbar, text="Returns & Credit Notes", command=self.open_credit_notes).pack(side='left', padx=5)

        columns = ("ID", "Date", "Item", "Quantity", "Customer", "Amount", "Invoice")
        self.tree = ttk.Treeview(main, columns=columns, show="headings", height=18)
        for c in columns:
            self.tree.heading(c, text=c)
        self.tree.column("ID", width=120)
        self.tree.column("Date", width=100)
        self.tree.column("Item", width=220)
        self.tree.column("Quantity", width=80)
        self.tree.column("Customer", width=160)
        self.tree.column("Amount", width=100)
        self.tree.column("Invoice", width=120)

        yscroll = ttk.Scrollbar(main, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

    def refresh_sales(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        try:
            sales = list(self.inventory_manager.sales or [])
        except Exception:
            sales = []
        try:
            sales = sorted(sales, key=lambda s: getattr(s, "sale_date", ""), reverse=True)
        except Exception:
            pass
        for sale in sales:
            try:
                sale_id = getattr(sale, "sale_id", "")
                date = getattr(sale, "sale_date", "")
                item_name = getattr(sale, "item_name", "")
                qty = getattr(sale, "quantity", 0) + getattr(sale, "promotion_quantity", 0)
                customer = getattr(sale, "customer_name", getattr(sale, "client_name", ""))
                amount = getattr(sale, "total_amount", 0)
                invoice_id = getattr(sale, "invoice_id", "")
                self.tree.insert(
                    "",
                    "end",
                    values=(
                        sale_id,
                        date,
                        item_name,
                        qty,
                        customer,
                        f"{amount:,.2f}",
                        invoice_id,
                    ),
                )
            except Exception:
                pass

    def _get_selected_ids(self):
        ids = []
        for sel in self.tree.selection():
            try:
                vals = self.tree.item(sel)["values"]
                if vals:
                    ids.append(vals[0])
            except Exception:
                pass
        return ids

    def delete_selected(self):
        ids = self._get_selected_ids()
        if not ids:
            messagebox.showwarning("Warning", "Select one or more sales to delete")
            return
        confirm_text = (
            "Delete the selected sales permanently?\n\n"
            "This only removes the records from saved sales and reports.\n"
            "Inventory quantities will not be changed."
        )
        if not messagebox.askyesno("Confirm Delete", confirm_text):
            return
        try:
            before_count = len(self.inventory_manager.sales or [])
            remaining = []
            for sale in list(self.inventory_manager.sales or []):
                if getattr(sale, "sale_id", None) in ids:
                    continue
                remaining.append(sale)
            self.inventory_manager.sales = remaining
            ok = self.inventory_manager.save_sales()
            if not ok:
                messagebox.showerror("Error", "Failed to save updated sales records")
                return
            self.refresh_sales()
            removed = before_count - len(remaining)
            messagebox.showinfo(
                "Success",
                f"Deleted {removed} sale record(s). They will no longer appear in reports.",
            )
            try:
                if callable(self.on_data_changed):
                    self.on_data_changed()
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete sales: {e}")

    def return_selected(self):
        ids = self._get_selected_ids()
        if len(ids) != 1:
            messagebox.showwarning("Warning", "Select exactly one sale to return")
            return
        sale = self.inventory_manager.get_sale_by_id(ids[0])
        if not sale:
            messagebox.showerror("Error", "Sale record not found")
            return
        if not self.inventory_manager.require_returns_permission(self.dialog):
            return
        dlg = ReturnSaleDialog(self.dialog, self.inventory_manager, sale)
        self.dialog.wait_window(dlg.dialog)
        self.refresh_sales()
        try:
            if callable(self.on_data_changed):
                self.on_data_changed()
        except Exception:
            pass

    def open_credit_notes(self):
        if not self.inventory_manager.require_returns_permission(self.dialog):
            return
        dlg = ReturnsAndCreditNotesDialog(self.dialog, self.inventory_manager)
        self.dialog.wait_window(dlg.dialog)

    def generate_report(self):
        return

class ReturnSaleDialog:
    def __init__(self, parent, inventory_manager, sale_record):
        self.inventory_manager = inventory_manager
        self.sale = sale_record
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Return Items")
        self.dialog.geometry("640x520")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.setup_ui()

    def setup_ui(self):
        attach_clock_label(self.dialog)
        main = ttk.Frame(self.dialog, padding=12)
        main.pack(fill="both", expand=True)

        sold_qty = int(getattr(self.sale, "quantity", 0) or 0)
        returned_qty = int(self.inventory_manager.get_returned_quantity_for_sale(getattr(self.sale, "sale_id", "")) or 0)
        remaining = max(0, sold_qty - returned_qty)

        header = ttk.Label(main, text="Process Return", font=('Helvetica', 14, 'bold'))
        header.pack(anchor="w", pady=(0, 10))

        info = ttk.LabelFrame(main, text="Original Sale", padding=10)
        info.pack(fill="x", pady=(0, 10))
        ttk.Label(info, text=f"Sale ID: {getattr(self.sale, 'sale_id', '')}").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=2)
        ttk.Label(info, text=f"Invoice: {getattr(self.sale, 'invoice_id', '') or 'N/A'}").grid(row=0, column=1, sticky="w", pady=2)
        ttk.Label(info, text=f"Customer: {getattr(self.sale, 'customer_name', '')}").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=2)
        ttk.Label(info, text=f"Item: {getattr(self.sale, 'item_name', '')}").grid(row=1, column=1, sticky="w", pady=2)
        ttk.Label(info, text=f"Sold Qty: {sold_qty}").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=2)
        ttk.Label(info, text=f"Already Returned: {returned_qty}").grid(row=2, column=1, sticky="w", pady=2)
        ttk.Label(info, text=f"Remaining Returnable: {remaining}").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=2)
        ttk.Label(info, text=f"Unit Price: AED {float(getattr(self.sale, 'selling_price', 0.0) or 0.0):.2f}").grid(row=3, column=1, sticky="w", pady=2)

        form = ttk.LabelFrame(main, text="Return Details", padding=10)
        form.pack(fill="x", pady=(0, 10))

        ttk.Label(form, text="Return Quantity *").grid(row=0, column=0, sticky="w", pady=4)
        self.qty_var = tk.StringVar(value=str(remaining if remaining > 0 else 0))
        ttk.Entry(form, textvariable=self.qty_var, width=12).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(form, text="Reason *").grid(row=1, column=0, sticky="w", pady=4)
        self.reason_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.reason_var, width=46).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(form, text="Restocking Fee (AED)").grid(row=2, column=0, sticky="w", pady=4)
        self.restock_fee_var = tk.StringVar(value="0.00")
        ttk.Entry(form, textvariable=self.restock_fee_var, width=12).grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(form, text="Notes").grid(row=3, column=0, sticky="nw", pady=4)
        self.notes_text = scrolledtext.ScrolledText(form, height=4, width=46)
        self.notes_text.grid(row=3, column=1, sticky="w", pady=4)

        self.generate_credit_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(form, text="Generate Credit Note", variable=self.generate_credit_var).grid(row=4, column=1, sticky="w", pady=(6, 0))

        btns = ttk.Frame(main)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Process Return", command=self.process).pack(side="left")
        ttk.Button(btns, text="Cancel", command=self.dialog.destroy).pack(side="left", padx=6)

    def process(self):
        try:
            qty = int(self.qty_var.get().strip() or "0")
        except Exception:
            messagebox.showwarning("Warning", "Return quantity must be a whole number")
            return
        reason = self.reason_var.get().strip()
        if not reason:
            messagebox.showwarning("Warning", "Return reason is required")
            return
        try:
            restock_fee = float(self.restock_fee_var.get().strip() or "0")
        except Exception:
            messagebox.showwarning("Warning", "Restocking fee must be a number")
            return
        notes = self.notes_text.get("1.0", tk.END).strip()

        ok, msg, ret, credit = self.inventory_manager.process_return(
            sale_id=getattr(self.sale, "sale_id", ""),
            return_quantity=qty,
            reason=reason,
            processed_by=getpass.getuser(),
            create_credit_note=bool(self.generate_credit_var.get()),
            credit_note_notes=notes,
            restocking_fee=restock_fee,
        )
        if not ok:
            messagebox.showerror("Error", msg)
            return

        pdf_path = None
        if credit:
            try:
                pdf_path = self.inventory_manager.generate_credit_note_pdf(credit)
            except Exception as e:
                messagebox.showwarning("Warning", f"Return processed, but PDF failed: {e}")

        if pdf_path:
            try:
                webbrowser.open('file://' + os.path.abspath(pdf_path))
            except Exception:
                pass
            messagebox.showinfo("Success", f"{msg}\nCredit Note: {credit.get('credit_note_id')}\nPDF:\n{pdf_path}")
        else:
            messagebox.showinfo("Success", msg)
        self.dialog.destroy()


class ReturnsAndCreditNotesDialog:
    def __init__(self, parent, inventory_manager):
        self.inventory_manager = inventory_manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Returns & Credit Notes")
        self.dialog.geometry("1050x650")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.setup_ui()
        self.refresh_all()

    def setup_ui(self):
        attach_clock_label(self.dialog)
        main = ttk.Frame(self.dialog, padding=10)
        main.pack(fill="both", expand=True)

        top = ttk.Frame(main)
        top.pack(fill="x", pady=(0, 8))
        ttk.Label(top, text="Search").pack(side="left")
        self.search_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.search_var, width=40).pack(side="left", padx=6)
        ttk.Button(top, text="Refresh", command=self.refresh_all).pack(side="left", padx=6)
        ttk.Button(top, text="Create Manual Credit Note", command=self.create_manual).pack(side="left", padx=6)
        ttk.Button(top, text="Edit Selected (Manual)", command=self.edit_selected_credit_note).pack(side="left", padx=6)
        ttk.Button(top, text="Delete Selected (Manual)", command=self.delete_selected_credit_note).pack(side="left", padx=6)
        ttk.Button(top, text="Print Selected Credit Note", command=self.print_selected_credit_note).pack(side="left", padx=6)
        ttk.Button(top, text="Close", command=self.dialog.destroy).pack(side="right")

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill="both", expand=True)

        self.credit_frame = ttk.Frame(self.notebook)
        self.returns_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.credit_frame, text="Credit Notes")
        self.notebook.add(self.returns_frame, text="Returns")

        credit_cols = ("ID", "Date", "Customer", "Total Credit", "Source", "Invoice/Ref")
        self.credit_tree = ttk.Treeview(self.credit_frame, columns=credit_cols, show="headings", height=18)
        for c in credit_cols:
            self.credit_tree.heading(c, text=c)
        self.credit_tree.column("ID", width=130)
        self.credit_tree.column("Date", width=100)
        self.credit_tree.column("Customer", width=220)
        self.credit_tree.column("Total Credit", width=120)
        self.credit_tree.column("Source", width=100)
        self.credit_tree.column("Invoice/Ref", width=220)
        y1 = ttk.Scrollbar(self.credit_frame, orient=tk.VERTICAL, command=self.credit_tree.yview)
        self.credit_tree.configure(yscrollcommand=y1.set)
        self.credit_tree.pack(side="left", fill="both", expand=True)
        y1.pack(side="right", fill="y")

        return_cols = ("Return ID", "Date", "Customer", "Item", "Qty", "Credit Note", "Reason")
        self.return_tree = ttk.Treeview(self.returns_frame, columns=return_cols, show="headings", height=18)
        for c in return_cols:
            self.return_tree.heading(c, text=c)
        self.return_tree.column("Return ID", width=130)
        self.return_tree.column("Date", width=100)
        self.return_tree.column("Customer", width=180)
        self.return_tree.column("Item", width=250)
        self.return_tree.column("Qty", width=70)
        self.return_tree.column("Credit Note", width=130)
        self.return_tree.column("Reason", width=220)
        y2 = ttk.Scrollbar(self.returns_frame, orient=tk.VERTICAL, command=self.return_tree.yview)
        self.return_tree.configure(yscrollcommand=y2.set)
        self.return_tree.pack(side="left", fill="both", expand=True)
        y2.pack(side="right", fill="y")

        self.search_var.trace("w", lambda *args: self.refresh_all())

    def refresh_all(self):
        q = (self.search_var.get() or "").strip().lower()

        for item in self.credit_tree.get_children():
            self.credit_tree.delete(item)
        notes = list(self.inventory_manager.credit_notes or [])
        try:
            notes.sort(key=lambda x: str(x.get("date", "")), reverse=True)
        except Exception:
            pass
        for cn in notes:
            try:
                cid = str(cn.get("credit_note_id") or "")
                date = str(cn.get("date") or "")
                cust = str(cn.get("customer_name") or "")
                total = float(cn.get("total_credit") or 0.0)
                source = str(cn.get("source") or "")
                ref = str(cn.get("invoice_id") or cn.get("reference") or "")
                hay = " ".join([cid, date, cust, source, ref]).lower()
                if q and q not in hay:
                    continue
                self.credit_tree.insert("", "end", values=(cid, date, cust, f"{total:,.2f}", source, ref))
            except Exception:
                pass

        for item in self.return_tree.get_children():
            self.return_tree.delete(item)
        rets = list(self.inventory_manager.returns or [])
        try:
            rets.sort(key=lambda x: str(getattr(x, "return_date", "")), reverse=True)
        except Exception:
            pass
        for r in rets:
            try:
                rid = str(getattr(r, "return_id", "") or "")
                date = str(getattr(r, "return_date", "") or "")
                cust = str(getattr(r, "customer_name", "") or "")
                item_name = str(getattr(r, "item_name", "") or "")
                qty = int(getattr(r, "quantity", 0) or 0)
                cnid = str(getattr(r, "credit_note_id", "") or "")
                reason = str(getattr(r, "reason", "") or "")
                hay = " ".join([rid, date, cust, item_name, cnid, reason]).lower()
                if q and q not in hay:
                    continue
                self.return_tree.insert("", "end", values=(rid, date, cust, item_name, qty, cnid, reason))
            except Exception:
                pass

    def _get_selected_credit_note(self):
        sel = self.credit_tree.selection()
        if not sel:
            return None
        vals = self.credit_tree.item(sel[0]).get("values") or []
        if not vals:
            return None
        cid = vals[0]
        for cn in (self.inventory_manager.credit_notes or []):
            if cn.get("credit_note_id") == cid:
                return cn
        return None

    def print_selected_credit_note(self):
        cn = self._get_selected_credit_note()
        if not cn:
            messagebox.showwarning("Warning", "Select a credit note first")
            return
        try:
            path = self.inventory_manager.generate_credit_note_pdf(cn)
            try:
                webbrowser.open('file://' + os.path.abspath(path))
            except Exception:
                pass
            messagebox.showinfo("Success", f"Credit note PDF generated:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {e}")

    def create_manual(self):
        dlg = ManualCreditNoteDialog(self.dialog, self.inventory_manager)
        self.dialog.wait_window(dlg.dialog)
        self.refresh_all()

    def edit_selected_credit_note(self):
        cn = self._get_selected_credit_note()
        if not cn:
            messagebox.showwarning("Warning", "Select a credit note first")
            return
        if str(cn.get("source") or "") != "manual":
            messagebox.showwarning("Warning", "Only manual credit notes can be edited")
            return
        dlg = ManualCreditNoteDialog(self.dialog, self.inventory_manager, existing_credit_note=cn)
        self.dialog.wait_window(dlg.dialog)
        self.refresh_all()

    def delete_selected_credit_note(self):
        cn = self._get_selected_credit_note()
        if not cn:
            messagebox.showwarning("Warning", "Select a credit note first")
            return
        if str(cn.get("source") or "") != "manual":
            messagebox.showwarning("Warning", "Only manual credit notes can be deleted")
            return
        cid = cn.get("credit_note_id")
        if not messagebox.askyesno("Confirm Delete", f"Delete manual credit note {cid}?"):
            return
        ok, msg = self.inventory_manager.delete_manual_credit_note(cid)
        if not ok:
            messagebox.showerror("Error", msg)
            return
        messagebox.showinfo("Success", msg)
        self.refresh_all()


class ManualCreditNoteDialog:
    def __init__(self, parent, inventory_manager, existing_credit_note=None):
        self.inventory_manager = inventory_manager
        self.items = []
        self.existing_credit_note = existing_credit_note if isinstance(existing_credit_note, dict) else None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Edit Manual Credit Note" if self.existing_credit_note else "Create Manual Credit Note")
        self.dialog.geometry("900x650")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.setup_ui()
        if self.existing_credit_note:
            self.load_existing()

    def setup_ui(self):
        attach_clock_label(self.dialog)
        main = ttk.Frame(self.dialog, padding=10)
        main.pack(fill="both", expand=True)

        top = ttk.LabelFrame(main, text="Credit Note Details", padding=10)
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Customer Name *").grid(row=0, column=0, sticky="w", pady=4)
        self.customer_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.customer_var, width=40).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(top, text="Reference (optional)").grid(row=0, column=2, sticky="w", padx=(12, 0), pady=4)
        self.reference_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.reference_var, width=26).grid(row=0, column=3, sticky="w", pady=4)

        ttk.Label(top, text="Restocking Fee (AED)").grid(row=1, column=0, sticky="w", pady=4)
        self.restock_fee_var = tk.StringVar(value="0.00")
        ttk.Entry(top, textvariable=self.restock_fee_var, width=12).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(top, text="Adjustment (+/- AED)").grid(row=1, column=2, sticky="w", padx=(12, 0), pady=4)
        self.adjustment_var = tk.StringVar(value="0.00")
        ttk.Entry(top, textvariable=self.adjustment_var, width=12).grid(row=1, column=3, sticky="w", pady=4)

        ttk.Label(top, text="Notes").grid(row=2, column=0, sticky="nw", pady=4)
        self.notes_text = scrolledtext.ScrolledText(top, height=3, width=72)
        self.notes_text.grid(row=2, column=1, columnspan=3, sticky="w", pady=4)

        items_frame = ttk.LabelFrame(main, text="Items", padding=10)
        items_frame.pack(fill="both", expand=True, pady=(0, 10))

        btns = ttk.Frame(items_frame)
        btns.pack(fill="x", pady=(0, 8))
        ttk.Button(btns, text="Add Item", command=self.add_item).pack(side="left")
        ttk.Button(btns, text="Edit Selected", command=self.edit_item).pack(side="left", padx=6)
        ttk.Button(btns, text="Remove Selected", command=self.remove_item).pack(side="left", padx=6)

        cols = ("Code", "Description", "Qty", "Unit Price", "Taxable", "VAT Rate")
        self.tree = ttk.Treeview(items_frame, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
        self.tree.column("Code", width=120)
        self.tree.column("Description", width=320)
        self.tree.column("Qty", width=80)
        self.tree.column("Unit Price", width=120)
        self.tree.column("Taxable", width=80)
        self.tree.column("VAT Rate", width=80)
        y = ttk.Scrollbar(items_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=y.set)
        self.tree.pack(side="left", fill="both", expand=True)
        y.pack(side="right", fill="y")

        bottom = ttk.Frame(main)
        bottom.pack(fill="x")
        self.generate_pdf_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="Generate PDF after saving", variable=self.generate_pdf_var).pack(side="left")
        ttk.Button(bottom, text="Save Credit Note", command=self.save).pack(side="right")
        ttk.Button(bottom, text="Cancel", command=self.dialog.destroy).pack(side="right", padx=6)

    def _refresh_items(self):
        for x in self.tree.get_children():
            self.tree.delete(x)
        for row in self.items:
            self.tree.insert(
                "",
                "end",
                values=(
                    row.get("item_code", ""),
                    row.get("description", ""),
                    row.get("quantity", ""),
                    row.get("unit_price", ""),
                    "Yes" if row.get("taxable") else "No",
                    row.get("vat_rate", ""),
                ),
            )

    def _get_selected_index(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.tree.index(sel[0])

    def add_item(self):
        dlg = CreditNoteItemDialog(self.dialog)
        self.dialog.wait_window(dlg.dialog)
        if dlg.result:
            self.items.append(dlg.result)
            self._refresh_items()

    def edit_item(self):
        idx = self._get_selected_index()
        if idx is None:
            messagebox.showwarning("Warning", "Select an item first")
            return
        dlg = CreditNoteItemDialog(self.dialog, initial=self.items[idx])
        self.dialog.wait_window(dlg.dialog)
        if dlg.result:
            self.items[idx] = dlg.result
            self._refresh_items()

    def remove_item(self):
        idx = self._get_selected_index()
        if idx is None:
            return
        del self.items[idx]
        self._refresh_items()

    def load_existing(self):
        cn = self.existing_credit_note or {}
        self.customer_var.set(str(cn.get("customer_name") or ""))
        self.reference_var.set(str(cn.get("reference") or ""))
        self.restock_fee_var.set(str(cn.get("restocking_fee") if cn.get("restocking_fee") is not None else "0.00"))
        self.adjustment_var.set(str(cn.get("adjustment") if cn.get("adjustment") is not None else "0.00"))
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", str(cn.get("notes") or ""))
        self.items = []
        for row in (cn.get("items") or []):
            if not isinstance(row, dict):
                continue
            self.items.append({
                "item_code": str(row.get("item_code") or row.get("item_id") or ""),
                "description": str(row.get("description") or ""),
                "quantity": row.get("quantity", 1),
                "unit_price": row.get("unit_price", 0.0),
                "taxable": bool(row.get("taxable", False)),
                "vat_rate": row.get("vat_rate", 0.0),
            })
        self._refresh_items()

    def save(self):
        customer = self.customer_var.get().strip()
        if not customer:
            messagebox.showwarning("Warning", "Customer name is required")
            return
        try:
            restock_fee = float(self.restock_fee_var.get().strip() or "0")
            adjustment = float(self.adjustment_var.get().strip() or "0")
        except Exception:
            messagebox.showwarning("Warning", "Restocking fee / adjustment must be numbers")
            return
        notes = self.notes_text.get("1.0", tk.END).strip()
        if self.existing_credit_note:
            ok, msg, credit = self.inventory_manager.update_manual_credit_note(
                credit_note_id=self.existing_credit_note.get("credit_note_id"),
                customer_name=customer,
                items=self.items,
                notes=notes,
                restocking_fee=restock_fee,
                adjustment=adjustment,
                reference=self.reference_var.get().strip(),
            )
        else:
            ok, msg, credit = self.inventory_manager.create_manual_credit_note(
                customer_name=customer,
                items=self.items,
                notes=notes,
                restocking_fee=restock_fee,
                adjustment=adjustment,
                reference=self.reference_var.get().strip(),
            )
        if not ok:
            messagebox.showerror("Error", msg)
            return
        pdf_path = None
        if self.generate_pdf_var.get():
            try:
                pdf_path = self.inventory_manager.generate_credit_note_pdf(credit)
            except Exception as e:
                messagebox.showwarning("Warning", f"Credit note saved, but PDF failed: {e}")
        if pdf_path:
            try:
                webbrowser.open('file://' + os.path.abspath(pdf_path))
            except Exception:
                pass
            messagebox.showinfo("Success", f"{msg}\nCredit Note: {credit.get('credit_note_id')}\nPDF:\n{pdf_path}")
        else:
            messagebox.showinfo("Success", f"{msg}\nCredit Note: {credit.get('credit_note_id')}")
        self.dialog.destroy()


class CreditNoteItemDialog:
    def __init__(self, parent, initial=None):
        self.result = None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Credit Note Item")
        self.dialog.geometry("520x360")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.setup_ui(initial or {})

    def setup_ui(self, initial):
        frm = ttk.Frame(self.dialog, padding=12)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Item Code").grid(row=0, column=0, sticky="w", pady=4)
        self.code_var = tk.StringVar(value=str(initial.get("item_code", "") or ""))
        ttk.Entry(frm, textvariable=self.code_var, width=30).grid(row=0, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Description *").grid(row=1, column=0, sticky="w", pady=4)
        self.desc_var = tk.StringVar(value=str(initial.get("description", "") or ""))
        ttk.Entry(frm, textvariable=self.desc_var, width=40).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Quantity *").grid(row=2, column=0, sticky="w", pady=4)
        self.qty_var = tk.StringVar(value=str(initial.get("quantity", "1") or "1"))
        ttk.Entry(frm, textvariable=self.qty_var, width=12).grid(row=2, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="Unit Price *").grid(row=3, column=0, sticky="w", pady=4)
        self.price_var = tk.StringVar(value=str(initial.get("unit_price", "0.00") or "0.00"))
        ttk.Entry(frm, textvariable=self.price_var, width=12).grid(row=3, column=1, sticky="w", pady=4)

        self.taxable_var = tk.BooleanVar(value=bool(initial.get("taxable", False)))
        ttk.Checkbutton(frm, text="Taxable", variable=self.taxable_var).grid(row=4, column=1, sticky="w", pady=4)

        ttk.Label(frm, text="VAT Rate").grid(row=5, column=0, sticky="w", pady=4)
        self.vat_rate_var = tk.StringVar(value=str(initial.get("vat_rate", "0.05") or "0.05"))
        ttk.Entry(frm, textvariable=self.vat_rate_var, width=12).grid(row=5, column=1, sticky="w", pady=4)

        btns = ttk.Frame(frm)
        btns.grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 0))
        ttk.Button(btns, text="Save", command=self.save).pack(side="left")
        ttk.Button(btns, text="Cancel", command=self.dialog.destroy).pack(side="left", padx=6)

    def save(self):
        desc = self.desc_var.get().strip()
        if not desc:
            messagebox.showwarning("Warning", "Description is required")
            return
        try:
            qty = float(self.qty_var.get().strip() or "0")
            price = float(self.price_var.get().strip() or "0")
            vat_rate = float(self.vat_rate_var.get().strip() or "0")
        except Exception:
            messagebox.showwarning("Warning", "Quantity, price and VAT rate must be numbers")
            return
        if qty <= 0 or price < 0:
            messagebox.showwarning("Warning", "Quantity must be > 0 and price cannot be negative")
            return
        if vat_rate > 1:
            vat_rate = vat_rate / 100.0
        self.result = {
            "item_code": self.code_var.get().strip(),
            "description": desc,
            "quantity": qty,
            "unit_price": price,
            "taxable": bool(self.taxable_var.get()),
            "vat_rate": vat_rate,
        }
        self.dialog.destroy()
    def generate_report(self):
        """Generate the selected report"""
        report_type = self.report_type.get()
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        
        # Clear existing content
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert(tk.END, f"Generating {report_type} report...\n")
        self.report_text.update()
        
        try:
            if report_type == "custom_multi":
                parts = []
                if getattr(self, "custom_sales", None) is not None and self.custom_sales.get():
                    parts.append(self.generate_sales_report(start_date, end_date))
                if getattr(self, "custom_costs", None) is not None and self.custom_costs.get():
                    parts.append(self.generate_cost_analysis_report())
                if getattr(self, "custom_movement", None) is not None and self.custom_movement.get():
                    parts.append(self.generate_movement_report(start_date, end_date))
                if getattr(self, "custom_low_stock", None) is not None and self.custom_low_stock.get():
                    parts.append(self.generate_low_stock_report())
                if getattr(self, "custom_new_items", None) is not None and self.custom_new_items.get():
                    parts.append(self.generate_new_items_report(start_date, end_date))
                if getattr(self, "custom_system_health", None) is not None and self.custom_system_health.get():
                    parts.append(self.generate_system_health_report(start_date, end_date))
                parts = [p for p in parts if p and str(p).strip()]
                if not parts:
                    report = "No sections selected for custom combined report."
                else:
                    report = ("\n" + "="*100 + "\n\n").join(parts)
            elif report_type == "per_item":
                # The per_item report handles its own insertion into the text widget
                # to allow for professional styling with tags.
                # However, if it returns a non-empty string, it's an error/info message.
                report = self.generate_per_item_report()
                if report:
                    self.report_text.delete('1.0', tk.END)
                    self.report_text.insert('1.0', f"\n{report}\n", "error")
            elif report_type == "full_summary":
                report = self.generate_full_summary_report(start_date, end_date)
            elif report_type == "system_health":
                report = self.generate_system_health_report(start_date, end_date)
            elif report_type == "movement":
                report = self.generate_movement_report(start_date, end_date)
            elif report_type == "sales":
                report = self.generate_sales_report(start_date, end_date)
                if "No sales recorded" in report or "0.00" in report:
                    report += "\n\nTip: Make sure you have recorded sales in the inventory system for this date range."
            elif report_type == "stock":
                report = self.generate_stock_report()
            elif report_type == "low_stock":
                report = self.generate_low_stock_report()
            elif report_type == "cost_analysis":
                report = self.generate_cost_analysis_report()
            elif report_type == "new_items":
                report = self.generate_new_items_report(start_date, end_date)
            else:
                report = "Invalid report type selected."
            
            # For per_item, we don't want to overwrite what it just inserted
            if report_type != "per_item":
                self.report_text.delete('1.0', tk.END)
                if not report or not str(report).strip():
                    report = f"The {report_type.replace('_', ' ')} report is currently empty.\n\n"
                    report += "Possible reasons:\n"
                    report += f"1. No data matches the selected date range ({start_date} to {end_date}).\n"
                    report += "2. No records have been created in the system yet.\n"
                    report += "3. Try clicking 'REFRESH DATA' to reload from the database."
                
                self.report_text.insert('1.0', report)
            
        except Exception as e:
            error_msg = f"ERROR: Failed to generate report.\n\nDetails: {str(e)}"
            self.report_text.delete('1.0', tk.END)
            self.report_text.insert('1.0', error_msg)
            messagebox.showerror("Report Error", f"An error occurred while generating the report:\n{e}")
            import traceback
            traceback.print_exc()
    
    def generate_sales_report(self, start_date, end_date):
        """Generate sales report with profit analysis"""
        report_data = self.inventory_manager.get_sales_report(start_date, end_date)
        
        report = f"""
ASSISTEM - SALES REPORT WITH PROFIT ANALYSIS
Period: {start_date} to {end_date}
{'='*90}

SUMMARY:
Total Sales: AED {report_data['total_sales']:,.2f}
Total Cost: AED {report_data['total_cost']:,.2f}
Total Profit: AED {report_data['total_profit']:,.2f}
Profit Margin: {report_data['profit_margin']:.2f}%
Total Quantity Sold: {report_data['total_quantity']}
Number of Sales: {len(report_data['sales'])}

DETAILED SALES WITH PROFIT:
{'='*90}
"""
        for sale in report_data['sales']:
            profit_margin = (sale.total_profit / sale.total_amount * 100) if sale.total_amount > 0 else 0
            report += f"""
Sale ID: {sale.sale_id}
Date: {sale.sale_date}
Customer: {sale.customer_name}
Item: {sale.item_name}
Quantity: {sale.quantity}
Cost per Item: AED {sale.cost_per_item:.2f}
Selling Price: AED {sale.selling_price:.2f}
Total Sales: AED {sale.total_amount:.2f}
Total Cost: AED {sale.total_cost:.2f}
PROFIT: AED {sale.total_profit:.2f} ({profit_margin:.2f}%)
Invoice: {sale.invoice_id or 'N/A'}
{'-'*70}
"""
        
        return report
    
    def generate_movement_report(self, start_date, end_date):
        """Generate full inventory movement report grouped by item"""
        report_data = self.inventory_manager.get_sales_report(start_date, end_date)
        sales = report_data['sales']
        
        # Index items by ID for quick lookup
        items_by_id = {item.item_id: item for item in self.inventory_manager.items}
        
        # Aggregate movement per item
        movement = {}
        for sale in sales:
            entry = movement.setdefault(sale.item_id, {
                'item_id': sale.item_id,
                'item_name': sale.item_name,
                'total_qty_sold': 0,
                'total_promo_qty': 0,
                'total_revenue': 0.0,
                'total_cost': 0.0,
                'total_profit': 0.0,
            })
            entry['total_qty_sold'] += sale.quantity
            entry['total_promo_qty'] += getattr(sale, 'promotion_quantity', 0)
            entry['total_revenue'] += sale.total_amount
            entry['total_cost'] += sale.total_cost
            entry['total_profit'] += sale.total_profit
        
        report_lines = []
        report_lines.append("ASSISTEM - FULL INVENTORY MOVEMENT REPORT")
        report_lines.append(f"Period: {start_date} to {end_date}")
        report_lines.append("="*90)
        report_lines.append("")
        report_lines.append(f"Total Items in Inventory: {len(self.inventory_manager.items)}")
        report_lines.append(f"Total Sales (AED): {report_data['total_sales']:,.2f}")
        report_lines.append(f"Total Profit (AED): {report_data['total_profit']:,.2f}")
        report_lines.append(f"Total Quantity Sold: {report_data['total_quantity']}")
        report_lines.append("")
        report_lines.append("MOVEMENT BY ITEM:")
        report_lines.append("="*90)
        
        if not self.inventory_manager.items:
            report_lines.append("No items found in inventory.")
        else:
            for item in self.inventory_manager.items:
                mov = movement.get(item.item_id, None)
                sold_qty = mov['total_qty_sold'] if mov else 0
                promo_qty = mov['total_promo_qty'] if mov else 0
                revenue = mov['total_revenue'] if mov else 0.0
                profit = mov['total_profit'] if mov else 0.0
                
                report_lines.append(f"Item ID: {item.item_id}")
                report_lines.append(f"Name: {item.name}")
                report_lines.append(f"Category: {item.category}")
                report_lines.append(f"Supplier: {item.supplier}")
                report_lines.append(f"Date Added: {item.created_date}")
                report_lines.append(f"Cost per Item: AED {item.get_cost_per_item():.2f}")
                report_lines.append(f"Selling Price: AED {item.selling_price:.2f}")
                report_lines.append(f"Current Stock: {item.quantity} (Min: {item.min_stock})")
                report_lines.append(f"Sold in Period: {sold_qty} (Promo: {promo_qty})")
                report_lines.append(f"Sales Revenue in Period: AED {revenue:.2f}")
                report_lines.append(f"Profit in Period: AED {profit:.2f}")
                report_lines.append("-"*80)
        
        return "\n".join(report_lines)
    
    def generate_full_summary_report(self, start_date, end_date):
        inv_summary = self.inventory_manager.get_inventory_summary()
        sales_summary = self.inventory_manager.get_sales_report(start_date, end_date)
        
        total_items = inv_summary['total_items']
        low_stock = inv_summary['low_stock_count']
        expired = inv_summary['expired_count']
        total_value = inv_summary['total_value']
        total_additional = inv_summary['total_additional_costs']
        potential_sales = inv_summary['total_potential_sales']
        potential_profit = inv_summary['total_potential_profit']
        
        total_sales = sales_summary['total_sales']
        total_cost = sales_summary['total_cost']
        total_profit = sales_summary['total_profit']
        margin = sales_summary['profit_margin']
        total_qty = sales_summary['total_quantity']
        num_sales = len(sales_summary['sales'])
        
        lines = []
        lines.append("ASSISTEM - FULL INVENTORY & SALES SUMMARY")
        lines.append(f"Period: {start_date} to {end_date}")
        lines.append("="*90)
        lines.append("")
        lines.append("INVENTORY SNAPSHOT")
        lines.append("-"*90)
        lines.append(f"Total Items: {total_items}")
        lines.append(f"Low Stock Items: {low_stock}")
        lines.append(f"Expired Items: {expired}")
        lines.append(f"Total Inventory Value (Cost): AED {total_value:,.2f}")
        lines.append(f"Total Additional Costs: AED {total_additional:,.2f}")
        lines.append(f"Potential Sales Value: AED {potential_sales:,.2f}")
        lines.append(f"Potential Profit: AED {potential_profit:,.2f}")
        lines.append("")
        lines.append("SALES PERFORMANCE (INVENTORY LINKED)")
        lines.append("-"*90)
        lines.append(f"Total Sales: AED {total_sales:,.2f}")
        lines.append(f"Total Cost of Goods Sold: AED {total_cost:,.2f}")
        lines.append(f"Total Profit: AED {total_profit:,.2f}")
        lines.append(f"Profit Margin: {margin:.2f}%")
        lines.append(f"Total Quantity Sold: {total_qty}")
        lines.append(f"Number of Sales: {num_sales}")
        lines.append("")
        lines.append("KEY INDICATORS")
        lines.append("-"*90)
        if low_stock > 0:
            lines.append(f"- {low_stock} item(s) are below minimum stock.")
        else:
            lines.append("- No items are currently below minimum stock.")
        if expired > 0:
            lines.append(f"- {expired} item(s) are expired and should be reviewed.")
        else:
            lines.append("- No expired items detected.")
        if potential_profit > 0:
            lines.append(f"- Inventory holds potential profit of AED {potential_profit:,.2f} if fully sold.")
        else:
            lines.append("- Potential profit could not be calculated (check item prices and costs).")
        
        base_summary = "\n".join(lines)
        parts = [base_summary]
        try:
            parts.append("")
            parts.append(self.generate_system_health_report(start_date, end_date))
        except Exception:
            pass
        try:
            parts.append("")
            parts.append(self.generate_movement_report(start_date, end_date))
        except Exception:
            pass
        try:
            parts.append("")
            parts.append(self.generate_low_stock_report())
        except Exception:
            pass
        try:
            parts.append("")
            parts.append(self.generate_cost_analysis_report())
        except Exception:
            pass
        try:
            parts.append("")
            parts.append(self.generate_new_items_report(start_date, end_date))
        except Exception:
            pass
        
        return "\n".join(parts)
    
    def generate_system_health_report(self, start_date, end_date):
        """Generate system health report highlighting potential data issues"""
        items = self.inventory_manager.items
        low_stock_items = [i for i in items if i.is_low_stock()]
        expired_items = [i for i in items if i.is_expired()]
        no_price_items = [i for i in items if not i.selling_price or i.selling_price <= 0]
        zero_cost_items = [i for i in items if i.get_cost_per_item() <= 0]
        negative_qty_items = [i for i in items if i.quantity < 0]
        
        # Sales without matching inventory item
        orphan_sales = []
        known_ids = {i.item_id for i in items}
        for s in self.inventory_manager.sales:
            if s.item_id not in known_ids:
                orphan_sales.append(s)
        
        lines = []
        lines.append("ASSISTEM - INVENTORY SYSTEM HEALTH CHECK")
        lines.append(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        lines.append("="*90)
        lines.append("")
        
        if not items:
            lines.append("No inventory items found. Please verify inventory.json.")
            return "\n".join(lines)
        
        def section(title):
            lines.append("")
            lines.append(title)
            lines.append("-"*90)
        
        section("LOW STOCK ITEMS")
        if low_stock_items:
            for i in low_stock_items:
                lines.append(f"{i.item_id} - {i.name} | Qty: {i.quantity} / Min: {i.min_stock} | Supplier: {i.supplier}")
        else:
            lines.append("All items are above minimum stock levels.")
        
        section("EXPIRED ITEMS")
        if expired_items:
            for i in expired_items:
                lines.append(f"{i.item_id} - {i.name} | Expiry: {i.expiry_date or 'N/A'} | Qty: {i.quantity}")
        else:
            lines.append("No expired items detected.")
        
        section("ITEMS WITH MISSING OR ZERO SELLING PRICE")
        if no_price_items:
            for i in no_price_items:
                lines.append(f"{i.item_id} - {i.name} | Price: {i.selling_price}")
        else:
            lines.append("All items have a selling price set.")
        
        section("ITEMS WITH ZERO OR MISSING COST")
        if zero_cost_items:
            for i in zero_cost_items:
                lines.append(f"{i.item_id} - {i.name} | Cost/Item: {i.get_cost_per_item():.2f}")
        else:
            lines.append("All items have a non-zero cost per item.")
        
        section("ITEMS WITH NEGATIVE QUANTITY")
        if negative_qty_items:
            for i in negative_qty_items:
                lines.append(f"{i.item_id} - {i.name} | Quantity: {i.quantity}")
        else:
            lines.append("No items with negative quantity detected.")
        
        section("SALES WITHOUT MATCHING INVENTORY ITEM")
        if orphan_sales:
            for s in orphan_sales[:50]:
                lines.append(f"Sale {s.sale_id} on {s.sale_date}: {s.item_id} - {s.item_name} (Invoice {s.invoice_id or 'N/A'})")
            if len(orphan_sales) > 50:
                lines.append(f"... and {len(orphan_sales) - 50} more sales without matching items.")
            lines.append("")
            lines.append("Tip: Check sales_records.json and inventory.json for item ID mismatches.")
        else:
            lines.append("All sales are linked to existing inventory items.")
        
        return "\n".join(lines)
    
    def generate_stock_report(self):
        """Generate stock report"""
        summary = self.inventory_manager.get_inventory_summary()
        
        report = f"""
ASSISTEM - INVENTORY STOCK REPORT
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*80}

SUMMARY:
Total Items: {summary['total_items']}
Low Stock Items: {summary['low_stock_count']}
Expired Items: {summary['expired_count']}
Total Inventory Value: AED {summary['total_value']:,.2f}
Total Additional Costs: AED {summary['total_additional_costs']:,.2f}
Total Potential Sales: AED {summary['total_potential_sales']:,.2f}
Total Potential Profit: AED {summary['total_potential_profit']:,.2f}

STOCK DETAILS:
{'='*80}
"""
        for item in self.inventory_manager.items:
            status = "LOW STOCK" if item.is_low_stock() else "OK"
            if item.is_expired():
                status = "EXPIRED"
            
            total_cost = item.get_total_cost()
            cost_per_item = item.get_cost_per_item()
            profit_per_item = item.get_profit_per_item()
            
            report += f"""
Item ID: {item.item_id}
Name: {item.name}
Category: {item.category}
Date Entered: {item.created_date}
Quantity: {item.quantity} (Min: {item.min_stock})
Total Cost: AED {total_cost:.2f}
Cost per Item: AED {cost_per_item:.2f}
Selling Price: AED {item.selling_price:.2f}
Profit per Item: AED {profit_per_item:.2f}
Status: {status}
{'-'*50}
"""
        
        return report
    
    def generate_new_items_report(self, start_date, end_date):
        """Generate report for items entered within a date range"""
        try:
            start = datetime.strptime(start_date, "%Y-%m-%d").date()
            end = datetime.strptime(end_date, "%Y-%m-%d").date()
        except Exception:
            return "Invalid date format. Please use YYYY-MM-DD"

        new_items = []
        for item in self.inventory_manager.items:
            try:
                created = datetime.strptime(item.created_date, "%Y-%m-%d").date()
                if start <= created <= end:
                    new_items.append(item)
            except Exception:
                continue

        report = f"""
ASSISTEM - NEWLY ENTERED ITEMS REPORT
Period: {start_date} to {end_date}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*80}

SUMMARY:
Total New Items: {len(new_items)}
Total Value of New Items: AED {sum(i.total_cost for i in new_items):,.2f}

NEW ITEMS DETAILS:
{'='*80}
"""
        if not new_items:
            report += "\nNo new items entered in this period."
        else:
            for item in new_items:
                report += f"""
Date Entered: {item.created_date}
Item ID: {item.item_id}
Name: {item.name}
Category: {item.category}
Quantity: {item.quantity}
Total Cost: AED {item.total_cost:.2f}
Selling Price: AED {item.selling_price:.2f}
Supplier: {item.supplier}
{'-'*50}
"""
        return report

    def generate_per_item_report(self):
        selections = []
        try:
            indices = self.item_selector.curselection()
            selections = [self.item_selector.get(i) for i in indices]
        except Exception:
            pass
        if not selections:
            return "Please select at least one item"
        
        items_map = {}
        for raw in selections:
            try:
                item_id = raw.split('(')[-1].strip(')')
                item = next((i for i in self.inventory_manager.items if i.item_id == item_id), None)
                if item:
                    items_map[item_id] = item
            except Exception:
                continue
        if not items_map:
            return "Selected items could not be resolved"
        
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        self.report_text.delete('1.0', tk.END)
        self.report_text.insert(tk.END, "ASSISTEM - DETAILED ITEM ANALYSIS\n", "header")
        self.report_text.insert(tk.END, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}\n", "subhead")
        self.report_text.insert(tk.END, "="*60 + "\n\n")

        all_sales_data = self.inventory_manager.get_sales_report("2000-01-01", "2099-12-31")
        sales_data_period = self.inventory_manager.get_sales_report(start_date, end_date)

        first = True
        for item_id, item in items_map.items():
            if not first:
                self.report_text.insert(tk.END, "\n" + "="*80 + "\n\n")
            first = False

            self.report_text.insert(tk.END, "PRODUCT INFORMATION\n", "subhead")
            self.report_text.insert(tk.END, f"Name:            {item.name}\n")
            self.report_text.insert(tk.END, f"Item ID:         {item.item_id}\n")
            self.report_text.insert(tk.END, f"Category:       {item.category}\n")
            self.report_text.insert(tk.END, f"Supplier:       {item.supplier}\n")
            self.report_text.insert(tk.END, f"Batch Number:   {item.batch_number or 'N/A'}\n")
            self.report_text.insert(tk.END, f"Expiry Date:    {item.expiry_date or 'N/A'}\n")
            self.report_text.insert(tk.END, f"Description:    {item.description or 'No description provided'}\n")
            self.report_text.insert(tk.END, f"Date Entered:   {item.created_date}\n")
            self.report_text.insert(tk.END, f"Added Via:      {getattr(item, 'addition_source', 'Manual Entry')}\n", "highlight")
            self.report_text.insert(tk.END, f"Last Updated:   {item.last_updated}\n\n")

            status_tag = "success" if not item.is_low_stock() else "error"
            self.report_text.insert(tk.END, "STOCK STATUS\n", "subhead")
            self.report_text.insert(tk.END, f"Current Quantity: {item.quantity}\n")
            self.report_text.insert(tk.END, f"Min Stock Level:  {item.min_stock}\n")
            self.report_text.insert(tk.END, f"Status:           {'LOW STOCK' if item.is_low_stock() else 'SUFFICIENT'}\n\n", status_tag)

            self.report_text.insert(tk.END, "FINANCIAL SUMMARY\n", "subhead")
            self.report_text.insert(tk.END, f"Initial/Total Cost:    AED {item.total_cost:.2f}\n")
            self.report_text.insert(tk.END, f"Unit Cost (Avg):       AED {item.get_cost_per_item():.2f}\n")
            self.report_text.insert(tk.END, f"Selling Price:         AED {item.selling_price:.2f}\n")
            self.report_text.insert(tk.END, f"Profit Per Unit:       AED {item.get_profit_per_item():.2f}\n")
            self.report_text.insert(tk.END, f"Potential Profit:      AED {(item.get_profit_per_item() * item.quantity):.2f}\n")
            
            if item.costs:
                self.report_text.insert(tk.END, "\nCost Breakdown:\n")
                for cost in item.costs:
                    self.report_text.insert(tk.END, f"- {cost['date']}: {cost['cost_type']} (AED {cost['amount']:.2f}) - {cost['description']}\n")
            self.report_text.insert(tk.END, "\n")

            item_sales_all = [s for s in all_sales_data['sales'] if s.item_id == item_id]
            item_sales_period = [s for s in sales_data_period['sales'] if s.item_id == item_id]

            self.report_text.insert(tk.END, "LIFETIME SALES SUMMARY\n", "subhead")
            if not item_sales_all:
                self.report_text.insert(tk.END, "No lifetime sales recorded.\n\n")
            else:
                self.report_text.insert(tk.END, f"Total Sold (All-time): {sum(s.quantity for s in item_sales_all)} units\n")
                self.report_text.insert(tk.END, f"Total Revenue (All-time): AED {sum(s.total_amount for s in item_sales_all):.2f}\n")
                self.report_text.insert(tk.END, f"Total Profit (All-time):  AED {sum(s.total_profit for s in item_sales_all):.2f}\n\n")

            self.report_text.insert(tk.END, f"PERIOD SALES ({start_date} to {end_date})\n", "subhead")
            if not item_sales_period:
                self.report_text.insert(tk.END, "No sales recorded for this specific period.\n\n")
            else:
                self.report_text.insert(tk.END, f"Sold in Period:    {sum(s.quantity for s in item_sales_period)} units\n")
                self.report_text.insert(tk.END, f"Revenue in Period: AED {sum(s.total_amount for s in item_sales_period):.2f}\n")
                self.report_text.insert(tk.END, f"Profit in Period:  AED {sum(s.total_profit for s in item_sales_period):.2f}\n\n")
                
                self.report_text.insert(tk.END, "Sales Details for Period:\n")
                self.report_text.insert(tk.END, f"{'Date':<12} | {'Invoice':<10} | {'Customer':<20} | {'Qty':<4} | {'Revenue':<10} | {'Profit':<10}\n")
                self.report_text.insert(tk.END, "-"*80 + "\n")
                for sale in item_sales_period:
                    cust_name = (sale.customer_name[:17] + '...') if len(sale.customer_name) > 20 else sale.customer_name
                    inv_id = sale.invoice_id or 'N/A'
                    self.report_text.insert(tk.END, f"{sale.sale_date:<12} | {inv_id:<10} | {cust_name:<20} | {sale.quantity:<4} | {sale.total_amount:<10.2f} | {sale.total_profit:<10.2f}\n")
                self.report_text.insert(tk.END, "\n")

        return ""

    def generate_low_stock_report(self):
        """Generate low stock alert report"""
        low_stock_items = self.inventory_manager.get_low_stock_items()
        
        report = f"""
ASSISTEM - LOW STOCK ALERT REPORT
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*80}

ITEMS NEEDING RESTOCK:
{'='*80}
"""
        if not low_stock_items:
            report += "\nNo low stock items found. All items are sufficiently stocked."
        else:
            for item in low_stock_items:
                report += f"""
Item ID: {item.item_id}
Name: {item.name}
Category: {item.category}
Current Stock: {item.quantity}
Minimum Required: {item.min_stock}
Shortage: {item.min_stock - item.quantity}
Supplier: {item.supplier}
Cost per Item: AED {item.get_cost_per_item():.2f}
{'-'*50}
"""
        
        return report
    
    def generate_cost_analysis_report(self):
        """Generate cost analysis report with profit calculations"""
        report = f"""
ASSISTEM - COST AND PROFIT ANALYSIS REPORT
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}
{'='*90}

COST AND PROFIT BREAKDOWN BY ITEM:
{'='*90}
"""
        for item in self.inventory_manager.items:
            cost_breakdown = item.get_cost_breakdown()
            
            report += f"""
Item: {item.name} ({item.item_id})
Quantity: {item.quantity}
Total Cost: AED {cost_breakdown['total_cost']:.2f}
Cost per Item: AED {cost_breakdown['cost_per_item']:.2f}
Selling Price: AED {cost_breakdown['selling_price_per_item']:.2f}
Profit per Item: AED {cost_breakdown['profit_per_item']:.2f}
Total Profit Potential: AED {cost_breakdown['total_profit_potential']:.2f}
"""
            
            if item.costs:
                report += "Additional Cost Details:\n"
                for cost in item.costs:
                    report += f"  - {cost['cost_type']}: AED {cost['amount']:.2f} ({cost['description']})\n"
            
            report += f"{'-'*70}"
        
        # Summary
        summary = self.inventory_manager.get_inventory_summary()
        avg_profit_margin = (summary['total_potential_profit'] / summary['total_potential_sales'] * 100) if summary['total_potential_sales'] > 0 else 0
        
        report += f"""

SUMMARY:
Total Items: {summary['total_items']}
Total Inventory Value: AED {summary['total_value']:,.2f}
Total Additional Costs: AED {summary['total_additional_costs']:,.2f}
Total Potential Sales: AED {summary['total_potential_sales']:,.2f}
Total Potential Profit: AED {summary['total_potential_profit']:,.2f}
Average Profit Margin: {avg_profit_margin:.2f}%
"""
        
        return report
    
    def export_report(self):
        try:
            content = self.report_text.get('1.0', tk.END)
            if not content.strip():
                messagebox.showwarning("Warning", "No report content to export")
                return
            
            # Create reports folder
            reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)
            
            logo_path = resolve_logo_path(self.inventory_manager.data_folder)
            logo_tag = ""
            if logo_path and os.path.exists(logo_path):
                logo_url = logo_path.replace("\\", "/")
                logo_tag = f'<img src="file://{logo_url}" alt="HopePharma Logo" class="logo" />'
            
            # Export as HTML
            filename = f"Inventory_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
            filepath = os.path.join(reports_folder, filename)
            
            html_content = f"""
            <!DOCTYPE html>
            <html>
            <head>
                <title>Inventory Report - Assistem</title>
                <style>
                    body {{ font-family: Arial, sans-serif; margin: 20px; }}
                    .header {{ text-align: center; border-bottom: 2px solid #2c5aa0; padding-bottom: 20px; }}
                    .company-name {{ font-size: 24px; font-weight: bold; color: #2c5aa0; }}
                    .logo {{ height: 80px; margin-bottom: 10px; }}
                    .content {{ font-family: 'Courier New', monospace; font-size: 12px; white-space: pre-wrap; line-height: 1.2; }}
                    .footer {{ margin-top: 30px; text-align: center; color: #666; font-size: 11px; }}
                </style>
            </head>
            <body>
                <div class="header">
                    {logo_tag}
                    <div class="company-name">HopePharma</div>
                    <div>Inventory Management Report</div>
                </div>
                <div class="content">{content}</div>
                <div class="footer">
                    Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}
                </div>
            </body>
            </html>
            """
            
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            # Open in browser
            webbrowser.open('file://' + os.path.abspath(filepath))
            messagebox.showinfo("Success", f"Report exported to:\n{filepath}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export report: {e}")

    def export_to_pdf(self):
        report_type = self.report_type.get()
        start_date = self.start_date.get()
        end_date = self.end_date.get()
        
        try:
            # Create reports folder
            reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)
            
            filename = f"Inventory_{report_type.title()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            filepath = os.path.join(reports_folder, filename)
            
            doc = SimpleDocTemplate(filepath, pagesize=landscape(A4), topMargin=0.5*inch, bottomMargin=0.5*inch)
            elements = []

            # Styles
            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'ReportTitle',
                parent=styles['Heading1'],
                fontSize=20,
                textColor=colors.HexColor("#2c3e50"),
                alignment=1, # Center
                spaceAfter=10
            )
            subtitle_style = ParagraphStyle(
                'ReportSubtitle',
                parent=styles['Normal'],
                fontSize=10,
                textColor=colors.grey,
                alignment=1, # Center
                spaceAfter=20
            )
            section_style = ParagraphStyle(
                'SectionHeader',
                parent=styles['Heading2'],
                fontSize=14,
                textColor=colors.HexColor("#2980b9"),
                spaceBefore=15,
                spaceAfter=10,
                borderPadding=5,
                borderWidth=0,
                backColor=colors.HexColor("#ecf0f1")
            )
            summary_style = ParagraphStyle(
                'SummaryText',
                parent=styles['Normal'],
                fontSize=11,
                leading=14,
                spaceAfter=5
            )

            logo_path = resolve_logo_path(self.inventory_manager.data_folder)
            if logo_path and os.path.exists(logo_path):
                try:
                    logo_img = RLImage(logo_path, width=1.2*inch, height=1.2*inch)
                    elements.append(logo_img)
                    elements.append(Spacer(1, 0.15*inch))
                except Exception:
                    pass
            
            # Report Header
            title_text = "INVENTORY SYSTEM REPORT"
            if report_type == "full_summary":
                title_text = "FULL INVENTORY & SALES SUMMARY"
            elif report_type == "system_health":
                title_text = "INVENTORY SYSTEM HEALTH CHECK"
            elif report_type == "movement":
                title_text = "INVENTORY MOVEMENT REPORT"
            elif report_type == "sales":
                title_text = "SALES & PROFIT REPORT"
            elif report_type == "stock":
                title_text = "FULL INVENTORY STOCK REPORT"
            elif report_type == "low_stock":
                title_text = "LOW STOCK ALERT REPORT"
            elif report_type == "cost_analysis":
                title_text = "COST & PROFIT ANALYSIS REPORT"
            elif report_type == "new_items":
                title_text = "NEWLY ENTERED ITEMS REPORT"
            elif report_type == "per_item":
                title_text = "DETAILED ITEM REPORT"
            elif report_type == "quick_summary":
                title_text = "INVENTORY SUMMARY REPORT"
            elif report_type == "inventory_management":
                title_text = "INVENTORY MANAGEMENT REPORT"
            
            elements.append(Paragraph(title_text, title_style))
            elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subtitle_style))
            if report_type in ["sales", "new_items"]:
                elements.append(Paragraph(f"Period: {start_date} to {end_date}", subtitle_style))
            
            elements.append(Spacer(1, 0.2*inch))

            # Helper for table styling
            def get_table_style(header_color="#34495e"):
                return TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(header_color)),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                    ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ])

            # Generate content based on report type
            if report_type in ("custom_multi", "quick_summary", "inventory_management"):
                content = self.report_text.get('1.0', tk.END)
                lines = content.splitlines()
                for line in lines:
                    if line.strip():
                        elements.append(Paragraph(line.replace("  ", "&nbsp;&nbsp;"), summary_style))
                    else:
                        elements.append(Spacer(1, 0.08*inch))
            elif report_type == "full_summary":
                inv_summary = self.inventory_manager.get_inventory_summary()
                sales_summary = self.inventory_manager.get_sales_report(start_date, end_date)
                
                elements.append(Paragraph("INVENTORY SNAPSHOT", section_style))
                elements.append(Paragraph(f"Total Items: {inv_summary['total_items']}", summary_style))
                elements.append(Paragraph(f"Low Stock Items: {inv_summary['low_stock_count']}", summary_style))
                elements.append(Paragraph(f"Expired Items: {inv_summary['expired_count']}", summary_style))
                elements.append(Paragraph(f"Total Inventory Value (Cost): AED {inv_summary['total_value']:,.2f}", summary_style))
                elements.append(Paragraph(f"Total Additional Costs: AED {inv_summary['total_additional_costs']:,.2f}", summary_style))
                elements.append(Paragraph(f"Potential Sales Value: AED {inv_summary['total_potential_sales']:,.2f}", summary_style))
                elements.append(Paragraph(f"Potential Profit: AED {inv_summary['total_potential_profit']:,.2f}", summary_style))
                
                elements.append(Paragraph("SALES PERFORMANCE (PERIOD)", section_style))
                elements.append(Paragraph(f"Period: {start_date} to {end_date}", summary_style))
                elements.append(Paragraph(f"Total Sales: AED {sales_summary['total_sales']:,.2f}", summary_style))
                elements.append(Paragraph(f"Total Cost of Goods Sold: AED {sales_summary['total_cost']:,.2f}", summary_style))
                elements.append(Paragraph(f"Total Profit: AED {sales_summary['total_profit']:,.2f}", summary_style))
                elements.append(Paragraph(f"Profit Margin: {sales_summary['profit_margin']:.2f}%", summary_style))
                elements.append(Paragraph(f"Total Quantity Sold: {sales_summary['total_quantity']}", summary_style))
                elements.append(Paragraph(f"Number of Sales: {len(sales_summary['sales'])}", summary_style))

            elif report_type == "system_health":
                lines = self.generate_system_health_report(start_date, end_date).splitlines()
                for line in lines:
                    if line.startswith("ASSISTEM") or line.startswith("LOW STOCK ITEMS") or line.startswith("EXPIRED ITEMS") or line.startswith("ITEMS WITH"):
                        elements.append(Paragraph(line, section_style))
                    else:
                        if line.strip():
                            elements.append(Paragraph(line.replace("  ", "&nbsp;&nbsp;"), summary_style))

            elif report_type == "movement":
                data = self.inventory_manager.get_sales_report(start_date, end_date)
                sales = data['sales']
                movement = {}
                for s in sales:
                    m = movement.setdefault(s.item_id, {
                        'item_name': s.item_name,
                        'qty': 0,
                        'promo': getattr(s, 'promotion_quantity', 0),
                        'revenue': 0.0,
                        'cost': 0.0,
                        'profit': 0.0,
                    })
                    m['qty'] += s.quantity
                    m['revenue'] += s.total_amount
                    m['cost'] += s.total_cost
                    m['profit'] += s.total_profit
                
                elements.append(Paragraph("SUMMARY", section_style))
                elements.append(Paragraph(f"Total Sales: AED {data['total_sales']:,.2f}", summary_style))
                elements.append(Paragraph(f"Total Cost of Goods Sold: AED {data['total_cost']:,.2f}", summary_style))
                elements.append(Paragraph(f"Total Profit: AED {data['total_profit']:,.2f}", summary_style))
                elements.append(Paragraph(f"Total Quantity Sold: {data['total_quantity']}", summary_style))
                elements.append(Paragraph(f"Number of Sales: {len(sales)}", summary_style))
                
                elements.append(Paragraph("MOVEMENT BY ITEM", section_style))
                table_data = [["ID", "Item Name", "Qty Sold", "Revenue (AED)", "Profit (AED)"]]
                for item in self.inventory_manager.items:
                    m = movement.get(item.item_id)
                    if not m:
                        continue
                    table_data.append([
                        item.item_id,
                        item.name[:25] + '..' if len(item.name) > 25 else item.name,
                        m['qty'],
                        f"{m['revenue']:.2f}",
                        f"{m['profit']:.2f}",
                    ])
                if len(table_data) == 1:
                    elements.append(Paragraph("No movement detected for this period.", summary_style))
                else:
                    t = Table(table_data, colWidths=[0.9*inch, 2.2*inch, 0.8*inch, 1.1*inch, 1.1*inch])
                    t.setStyle(get_table_style("#8e44ad"))
                    elements.append(t)

            elif report_type == "sales":
                data = self.inventory_manager.get_sales_report(start_date, end_date)
                elements.append(Paragraph("SUMMARY", section_style))
                elements.append(Paragraph(f"Total Sales: <b>AED {data['total_sales']:,.2f}</b>", summary_style))
                elements.append(Paragraph(f"Total Cost: AED {data['total_cost']:,.2f}", summary_style))
                elements.append(Paragraph(f"Total Profit: <b>AED {data['total_profit']:,.2f}</b>", summary_style))
                elements.append(Paragraph(f"Profit Margin: {data['profit_margin']:.2f}%", summary_style))
                
                elements.append(Paragraph("DETAILED SALES", section_style))
                table_data = [["Date", "Item", "Qty", "Sales (AED)", "Cost (AED)", "Profit (AED)"]]
                for s in data['sales']:
                    table_data.append([
                        s.sale_date,
                        s.item_name[:30] + '...' if len(s.item_name) > 30 else s.item_name,
                        s.quantity,
                        f"{s.total_amount:.2f}",
                        f"{s.total_cost:.2f}",
                        f"{s.total_profit:.2f}"
                    ])
                
                t = Table(table_data, colWidths=[0.8*inch, 2.2*inch, 0.5*inch, 1.0*inch, 1.0*inch, 1.0*inch])
                t.setStyle(get_table_style("#2980b9"))
                elements.append(t)

            elif report_type == "stock":
                summary = self.inventory_manager.get_inventory_summary()
                elements.append(Paragraph("INVENTORY SUMMARY", section_style))
                elements.append(Paragraph(f"Total Unique Items: {summary['total_items']}", summary_style))
                elements.append(Paragraph(f"Total Value (at Cost): <b>AED {summary['total_value']:,.2f}</b>", summary_style))
                elements.append(Paragraph(f"Potential Profit: <b>AED {summary['total_potential_profit']:,.2f}</b>", summary_style))
                
                elements.append(Paragraph("STOCK DETAILS", section_style))
                table_data = [["ID", "Item Name", "Date Entered", "Qty", "Cost", "Price", "Status"]]
                for item in self.inventory_manager.items:
                    status = "Low" if item.is_low_stock() else "OK"
                    if item.is_expired(): status = "Expired"
                    
                    table_data.append([
                        item.item_id,
                        item.name[:25] + '..' if len(item.name) > 25 else item.name,
                        item.created_date,
                        item.quantity,
                        f"{item.get_cost_per_item():.2f}",
                        f"{item.selling_price:.2f}",
                        status
                    ])
                
                t = Table(table_data, colWidths=[0.8*inch, 2.0*inch, 0.9*inch, 0.5*inch, 0.8*inch, 0.8*inch, 0.7*inch])
                t.setStyle(get_table_style("#2c3e50"))
                elements.append(t)

            elif report_type == "new_items":
                new_items = []
                start = datetime.strptime(start_date, "%Y-%m-%d").date()
                end = datetime.strptime(end_date, "%Y-%m-%d").date()
                for item in self.inventory_manager.items:
                    try:
                        created = datetime.strptime(item.created_date, "%Y-%m-%d").date()
                        if start <= created <= end: new_items.append(item)
                    except: continue
                
                elements.append(Paragraph(f"New Items Entered: {len(new_items)}", section_style))
                if new_items:
                    table_data = [["Date", "ID", "Item Name", "Category", "Qty", "Total Cost"]]
                    for item in new_items:
                        table_data.append([
                            item.created_date,
                            item.item_id,
                            item.name[:25],
                            item.category[:15],
                            item.quantity,
                            f"{item.total_cost:.2f}"
                        ])
                    t = Table(table_data, colWidths=[1.0*inch, 0.8*inch, 2.2*inch, 1.2*inch, 0.5*inch, 1.0*inch])
                    t.setStyle(get_table_style("#27ae60"))
                    elements.append(t)
                else:
                    elements.append(Paragraph("No items entered during this period.", summary_style))

            elif report_type == "per_item":
                selected = self.item_selector.get()
                if not selected: raise ValueError("No item selected")
                item_id = selected.split('(')[-1].strip(')')
                item = next((i for i in self.inventory_manager.items if i.item_id == item_id), None)
                
                if item:
                    elements.append(Paragraph(f"ITEM: {item.name}", section_style))
                    
                    # Basic Info Table
                    info_data = [
                        ["Item ID", item.item_id, "Category", item.category],
                        ["Supplier", item.supplier, "Date Entered", item.created_date],
                        ["Batch #", item.batch_number or 'N/A', "Expiry Date", item.expiry_date or 'N/A'],
                        ["Current Qty", str(item.quantity), "Min Stock", str(item.min_stock)],
                        ["Addition Method", getattr(item, 'addition_source', 'Manual Entry'), "Last Updated", item.last_updated]
                    ]
                    t_info = Table(info_data, colWidths=[1.2*inch, 1.8*inch, 1.2*inch, 1.8*inch])
                    t_info.setStyle(TableStyle([
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
                        ('FONTNAME', (2,0), (2,-1), 'Helvetica-Bold'),
                        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#f8f9fa")),
                        ('BACKGROUND', (2,0), (2,-1), colors.HexColor("#f8f9fa")),
                    ]))
                    elements.append(t_info)
                    
                    # Description
                    if item.description:
                        elements.append(Paragraph("DESCRIPTION", section_style))
                        elements.append(Paragraph(item.description, summary_style))
                    
                    # Pricing & Profit Table
                    elements.append(Paragraph("PRICING & FINANCIALS", section_style))
                    pricing_data = [
                        ["Cost Per Item (Avg)", f"AED {item.get_cost_per_item():.2f}"],
                        ["Selling Price", f"AED {item.selling_price:.2f}"],
                        ["Profit Per Unit", f"AED {item.get_profit_per_item():.2f}"],
                        ["Total Inventory Value", f"AED {item.get_total_cost():.2f}"],
                        ["Potential Total Profit", f"AED {item.get_total_profit_potential():.2f}"]
                    ]
                    t_price = Table(pricing_data, colWidths=[2.5*inch, 2.5*inch])
                    t_price.setStyle(TableStyle([
                        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                        ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
                        ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#f8f9fa")),
                    ]))
                    elements.append(t_price)

                    # Additional Costs Breakdown
                    if item.costs:
                        elements.append(Paragraph("ADDITIONAL COSTS BREAKDOWN", section_style))
                        cost_data = [["Date", "Type", "Amount (AED)", "Description"]]
                        for c in item.costs:
                            cost_data.append([c['date'], c['cost_type'], f"{c['amount']:.2f}", c['description']])
                        t_costs = Table(cost_data, colWidths=[1.0*inch, 1.2*inch, 1.2*inch, 2.6*inch])
                        t_costs.setStyle(get_table_style("#e67e22"))
                        elements.append(t_costs)
                    
                    # Sales History
                    all_sales_data = self.inventory_manager.get_sales_report("2000-01-01", "2099-12-31")
                    item_sales_all = [s for s in all_sales_data['sales'] if s.item_id == item_id]
                    
                    elements.append(Paragraph("LIFETIME SALES SUMMARY", section_style))
                    if item_sales_all:
                        lifetime_data = [
                            ["Total Units Sold", str(sum(s.quantity for s in item_sales_all))],
                            ["Total Lifetime Revenue", f"AED {sum(s.total_amount for s in item_sales_all):.2f}"],
                            ["Total Lifetime Profit", f"AED {sum(s.total_profit for s in item_sales_all):.2f}"]
                        ]
                        t_lifetime = Table(lifetime_data, colWidths=[2.5*inch, 2.5*inch])
                        t_lifetime.setStyle(TableStyle([
                            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                            ('FONTNAME', (0,0), (0,-1), 'Helvetica-Bold'),
                            ('BACKGROUND', (0,0), (0,-1), colors.HexColor("#f8f9fa")),
                        ]))
                        elements.append(t_lifetime)
                    else:
                        elements.append(Paragraph("No lifetime sales recorded.", summary_style))

                    # Period Sales Details
                    sales_data = self.inventory_manager.get_sales_report(start_date, end_date)
                    item_sales_period = [s for s in sales_data['sales'] if s.item_id == item_id]
                    elements.append(Paragraph(f"PERIOD SALES DETAILS ({start_date} to {end_date})", section_style))
                    if item_sales_period:
                        h_data = [["Date", "Invoice", "Customer", "Qty", "Total Sales", "Profit"]]
                        for s in item_sales_period:
                            h_data.append([s.sale_date, s.invoice_id or 'N/A', s.customer_name[:15], s.quantity, f"{s.total_amount:.2f}", f"{s.total_profit:.2f}"])
                        t_h = Table(h_data, colWidths=[0.9*inch, 1.0*inch, 1.6*inch, 0.5*inch, 1.0*inch, 1.0*inch])
                        t_h.setStyle(get_table_style("#8e44ad"))
                        elements.append(t_h)
                    else:
                        elements.append(Paragraph("No sales history found for this period.", summary_style))

            elif report_type == "low_stock":
                low_items = self.inventory_manager.get_low_stock_items()
                elements.append(Paragraph("LOW STOCK ALERT", section_style))
                if low_items:
                    table_data = [["ID", "Item Name", "Qty", "Min Stock", "Shortage", "Supplier"]]
                    for item in low_items:
                        table_data.append([
                            item.item_id,
                            item.name[:25],
                            item.quantity,
                            item.min_stock,
                            item.min_stock - item.quantity,
                            item.supplier[:20]
                        ])
                    t = Table(table_data, colWidths=[0.8*inch, 2.5*inch, 0.6*inch, 0.8*inch, 0.8*inch, 1.5*inch])
                    t.setStyle(get_table_style("#c0392b"))
                    elements.append(t)
                else:
                    elements.append(Paragraph("All items are sufficiently stocked.", summary_style))

            elif report_type == "cost_analysis":
                elements.append(Paragraph("ITEM COST & MARGIN ANALYSIS", section_style))
                table_data = [["Item Name", "Cost/Unit", "Price/Unit", "Profit/Unit", "Margin"]]
                for item in self.inventory_manager.items:
                    profit = item.get_profit_per_item()
                    margin = (profit / item.selling_price * 100) if item.selling_price > 0 else 0
                    table_data.append([
                        item.name[:30],
                        f"{item.get_cost_per_item():.2f}",
                        f"{item.selling_price:.2f}",
                        f"{profit:.2f}",
                        f"{margin:.1f}%"
                    ])
                t = Table(table_data, colWidths=[2.5*inch, 1.0*inch, 1.0*inch, 1.0*inch, 1.0*inch])
                t.setStyle(get_table_style("#f39c12"))
                elements.append(t)

            # Finalize
            elements.append(Spacer(1, 0.5*inch))
            footer = Paragraph("ASSISTEM Inventory Management - Professional Report", subtitle_style)
            elements.append(footer)
            
            doc.build(elements)
            webbrowser.open('file://' + os.path.abspath(filepath))
            messagebox.showinfo("Success", f"Professional PDF generated:\n{filepath}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to generate professional PDF: {e}")
            import traceback
            traceback.print_exc()
    
    def print_report(self):
        """Print the current report"""
        self.export_report()  # For now, export as HTML which can be printed

class SupplierOutboundReportDialog:
    def __init__(self, parent, inventory_manager, default_supplier="All Suppliers"):
        self.inventory_manager = inventory_manager
        self.default_supplier = default_supplier or "All Suppliers"
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Supplier Outbound Report")
        self.dialog.geometry("1200x820")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        self._rows_cache = []
        self.setup_ui()
        try:
            self.dialog.after(0, self.generate_report)
        except Exception:
            pass

    def _parse_date(self, value):
        value = str(value or "").strip()
        if not value:
            return None
        try:
            return datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            return value

    def setup_ui(self):
        main = ttk.Frame(self.dialog, padding=10)
        main.pack(fill='both', expand=True)

        header = ttk.Frame(main)
        header.pack(fill='x', pady=(0, 8))
        ttk.Label(header, text="Supplier Outbound Report", font=('Helvetica', 16, 'bold')).pack(side='left')

        controls = ttk.LabelFrame(main, text="Filters", padding=10)
        controls.pack(fill='x', pady=(0, 10))

        suppliers = sorted({(getattr(i, 'supplier', '') or '').strip() or 'Unknown Supplier' for i in (self.inventory_manager.items or [])})
        supplier_values = ['All Suppliers'] + suppliers

        ttk.Label(controls, text="Supplier:").grid(row=0, column=0, sticky='w')
        self.supplier_var = tk.StringVar(value=self.default_supplier if self.default_supplier in supplier_values else 'All Suppliers')
        self.supplier_combo = ttk.Combobox(controls, textvariable=self.supplier_var, values=supplier_values, width=34, state='readonly')
        self.supplier_combo.grid(row=0, column=1, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="Start Date:").grid(row=0, column=2, sticky='w')
        self.start_date_var = tk.StringVar(value="")
        ttk.Entry(controls, textvariable=self.start_date_var, width=12).grid(row=0, column=3, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="End Date:").grid(row=0, column=4, sticky='w')
        self.end_date_var = tk.StringVar(value="")
        ttk.Entry(controls, textvariable=self.end_date_var, width=12).grid(row=0, column=5, sticky='w', padx=(6, 18))

        btns = ttk.Frame(controls)
        btns.grid(row=0, column=6, sticky='e')
        ttk.Button(btns, text="Generate", command=self.generate_report).pack(side='left', padx=4)
        ttk.Button(btns, text="Export PDF", command=self.export_pdf).pack(side='left', padx=4)
        ttk.Button(btns, text="Close", command=self.dialog.destroy).pack(side='left', padx=4)
        controls.columnconfigure(6, weight=1)

        results = ttk.LabelFrame(main, text="Outbound Activity", padding=10)
        results.pack(fill='both', expand=True)

        columns = ('Supplier', 'Item ID', 'Item Name', 'Batch', 'Stored', 'Entry Date', 'Sold Date', 'Sold To', 'Sold Qty', 'Current Qty', 'Invoice ID')
        self.tree = ttk.Treeview(results, columns=columns, show='headings')
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, stretch=True)
        self.tree.column('Supplier', width=170)
        self.tree.column('Item Name', width=220)
        self.tree.column('Sold To', width=170)
        self.tree.column('Invoice ID', width=130)

        yscroll = ttk.Scrollbar(results, orient='vertical', command=self.tree.yview)
        xscroll = ttk.Scrollbar(results, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(side='left', fill='both', expand=True)
        yscroll.pack(side='right', fill='y')
        xscroll.pack(side='bottom', fill='x')

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w').pack(fill='x', pady=(8, 0))

    def generate_report(self):
        start = self._parse_date(self.start_date_var.get())
        end = self._parse_date(self.end_date_var.get())
        supplier = self.supplier_var.get()
        rows = self.inventory_manager.get_supplier_activity_rows(start_date=start, end_date=end, supplier_filter=supplier)
        self._rows_cache = rows

        for child in self.tree.get_children():
            self.tree.delete(child)

        for row in rows:
            self.tree.insert('', 'end', values=(
                row.get('supplier', ''),
                row.get('item_id', ''),
                row.get('item_name', ''),
                row.get('batch_number', ''),
                "Yes" if row.get('is_storage_item') else "No",
                row.get('entry_date', ''),
                row.get('sold_date', ''),
                row.get('sold_to', ''),
                row.get('sold_qty', 0),
                row.get('current_qty', 0),
                row.get('invoice_id', ''),
            ))

        self.status_var.set(f"Rows: {len(rows)}")

    def export_pdf(self):
        rows = list(self._rows_cache or [])
        if not rows:
            messagebox.showwarning("Export PDF", "No outbound rows to export.")
            return

        reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
        os.makedirs(reports_folder, exist_ok=True)
        filepath = os.path.join(reports_folder, f"Supplier_Outbound_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('OutboundTitle', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor("#2c3e50"))
        subtitle_style = ParagraphStyle('OutboundSub', parent=styles['Normal'], fontSize=9, alignment=1, textColor=colors.grey)
        doc = SimpleDocTemplate(filepath, pagesize=landscape(A4), leftMargin=0.35 * inch, rightMargin=0.35 * inch, topMargin=0.4 * inch, bottomMargin=0.4 * inch)

        elements = []
        elements.extend(build_pdf_logo_flowables(self.inventory_manager.data_folder, width=1.0 * inch, height=1.0 * inch, spacer_height=0.08 * inch))
        elements.extend([
            Paragraph("Supplier Outbound Report", title_style),
            Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_style),
            Spacer(1, 10),
        ])

        grouped = {}
        for row in rows:
            grouped.setdefault(row.get('supplier', 'Unknown Supplier'), []).append(row)

        for supplier, supplier_rows in grouped.items():
            elements.append(Paragraph(f"Supplier: {supplier}", ParagraphStyle('OutboundH2', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#0b5394'))))
            data = [["Item", "Entry", "Sold", "Sold To", "Sold Qty", "Current Qty", "Invoice"]]
            for row in supplier_rows:
                item_name = row.get('item_name', '')
                if row.get('is_storage_item'):
                    item_name = f"[Stored] {item_name}"
                if row.get('batch_number'):
                    item_name = f"{item_name} ({row.get('batch_number')})"
                data.append([
                    item_name,
                    row.get('entry_date', ''),
                    row.get('sold_date', ''),
                    row.get('sold_to', ''),
                    str(row.get('sold_qty', 0)),
                    str(row.get('current_qty', 0)),
                    row.get('invoice_id', ''),
                ])
            table = Table(data, colWidths=[250, 75, 75, 150, 55, 65, 120])
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
            elements.append(Spacer(1, 12))

        try:
            doc.build(elements)
            try:
                webbrowser.open("file://" + os.path.abspath(filepath))
            except Exception:
                pass
            messagebox.showinfo("Export PDF", f"Outbound report exported to:\n{filepath}")
        except Exception as exc:
            messagebox.showerror("Export PDF", f"Failed to export PDF:\n{exc}")


class SupplierInboundReportDialog:
    def __init__(self, parent, inventory_manager, default_supplier_id=None):
        self.inventory_manager = inventory_manager
        self.default_supplier_id = default_supplier_id or ""
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Supplier Inbound Report")
        self.dialog.geometry("1250x820")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        self._rows_cache = []
        self._supplier_name_to_id = {'All Suppliers': ''}
        self.setup_ui()
        try:
            self.dialog.after(0, self.generate_report)
        except Exception:
            pass

    def _parse_date(self, value):
        value = str(value or "").strip()
        if not value:
            return None
        try:
            return datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            return value

    def setup_ui(self):
        main = ttk.Frame(self.dialog, padding=10)
        main.pack(fill='both', expand=True)

        self.tree_style_name = "SupplierInboundReport.Treeview"
        try:
            style = ttk.Style(self.dialog)
            style.configure(self.tree_style_name, rowheight=48)
        except Exception:
            self.tree_style_name = "Treeview"

        header = ttk.Frame(main)
        header.pack(fill='x', pady=(0, 8))
        ttk.Label(header, text="Supplier Inbound Report", font=('Helvetica', 16, 'bold')).pack(side='left')

        controls = ttk.LabelFrame(main, text="Filters", padding=10)
        controls.pack(fill='x', pady=(0, 10))

        supplier_records = self.inventory_manager.load_supplier_records()
        supplier_names = ['All Suppliers']
        for record in supplier_records:
            supplier_names.append(record.get('name', ''))
            self._supplier_name_to_id[record.get('name', '')] = record.get('supplier_id', '')

        default_name = 'All Suppliers'
        for record in supplier_records:
            if record.get('supplier_id') == self.default_supplier_id:
                default_name = record.get('name', 'All Suppliers')
                break

        ttk.Label(controls, text="Supplier:").grid(row=0, column=0, sticky='w')
        self.supplier_var = tk.StringVar(value=default_name)
        ttk.Combobox(controls, textvariable=self.supplier_var, values=supplier_names, width=34, state='readonly').grid(row=0, column=1, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="Start Date:").grid(row=0, column=2, sticky='w')
        self.start_date_var = tk.StringVar(value="")
        ttk.Entry(controls, textvariable=self.start_date_var, width=12).grid(row=0, column=3, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="End Date:").grid(row=0, column=4, sticky='w')
        self.end_date_var = tk.StringVar(value="")
        ttk.Entry(controls, textvariable=self.end_date_var, width=12).grid(row=0, column=5, sticky='w', padx=(6, 18))

        btns = ttk.Frame(controls)
        btns.grid(row=0, column=6, sticky='e')
        ttk.Button(btns, text="Generate", command=self.generate_report).pack(side='left', padx=4)
        ttk.Button(btns, text="Export PDF", command=self.export_pdf).pack(side='left', padx=4)
        ttk.Button(btns, text="Close", command=self.dialog.destroy).pack(side='left', padx=4)
        controls.columnconfigure(6, weight=1)

        results = ttk.LabelFrame(main, text="Inbound Activity", padding=10)
        results.pack(fill='both', expand=True)

        columns = ('Supplier', 'Invoice Date', 'Invoice No', 'File', 'Item', 'Qty', 'Unit Price', 'Line Total', 'Invoice Total')
        self.tree = ttk.Treeview(results, columns=columns, show='headings', style=self.tree_style_name)
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=110, stretch=True)
        self.tree.column('Supplier', width=170)
        self.tree.column('File', width=240)
        self.tree.column('Item', width=330)
        self.tree.column('Invoice No', width=130)

        yscroll = ttk.Scrollbar(results, orient='vertical', command=self.tree.yview)
        xscroll = ttk.Scrollbar(results, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(side='left', fill='both', expand=True)
        yscroll.pack(side='right', fill='y')
        xscroll.pack(side='bottom', fill='x')

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w').pack(fill='x', pady=(8, 0))

    def generate_report(self):
        start = self._parse_date(self.start_date_var.get())
        end = self._parse_date(self.end_date_var.get())
        supplier_name = self.supplier_var.get()
        supplier_id = self._supplier_name_to_id.get(supplier_name, '')
        rows = self.inventory_manager.get_supplier_inbound_rows(start_date=start, end_date=end, supplier_id=supplier_id or None)
        self._rows_cache = rows

        for child in self.tree.get_children():
            self.tree.delete(child)

        for row in rows:
            self.tree.insert('', 'end', values=(
                row.get('supplier_name', ''),
                row.get('document_date', ''),
                row.get('invoice_number', ''),
                self._wrap_tree_text(row.get('file_name', ''), 24),
                self._wrap_tree_text(row.get('item_description', ''), 34),
                row.get('quantity', 0),
                f"AED {row.get('unit_price', 0.0):.2f}",
                f"AED {row.get('line_total', 0.0):.2f}",
                f"AED {row.get('grand_total', 0.0):.2f}",
            ))

        self.status_var.set(f"Rows: {len(rows)}")

    def _wrap_tree_text(self, value, width):
        text = str(value or "").strip()
        if not text:
            return ""
        return textwrap.fill(text, width=width, break_long_words=False, break_on_hyphens=False)

    def _pdf_paragraph(self, value, style, width=34):
        text = str(value or "").strip()
        if not text:
            text = "-"
        text = textwrap.fill(text, width=width, break_long_words=False, break_on_hyphens=False)
        safe_text = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace("\n", "<br/>")
        return Paragraph(safe_text, style)

    def export_pdf(self):
        rows = list(self._rows_cache or [])
        if not rows:
            messagebox.showwarning("Export PDF", "No inbound rows to export.")
            return

        reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
        os.makedirs(reports_folder, exist_ok=True)
        filepath = os.path.join(reports_folder, f"Supplier_Inbound_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('InboundTitle', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor("#2c3e50"))
        subtitle_style = ParagraphStyle('InboundSub', parent=styles['Normal'], fontSize=9, alignment=1, textColor=colors.grey)
        cell_style = ParagraphStyle('InboundCell', parent=styles['BodyText'], fontSize=7.5, leading=9, wordWrap='CJK')
        doc = SimpleDocTemplate(filepath, pagesize=landscape(A4), leftMargin=0.35 * inch, rightMargin=0.35 * inch, topMargin=0.4 * inch, bottomMargin=0.4 * inch)

        elements = []
        elements.extend(build_pdf_logo_flowables(self.inventory_manager.data_folder, width=1.0 * inch, height=1.0 * inch, spacer_height=0.08 * inch))
        elements.extend([
            Paragraph("Supplier Inbound Report", title_style),
            Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_style),
            Spacer(1, 10),
        ])

        grouped = {}
        for row in rows:
            grouped.setdefault(row.get('supplier_name', 'Unknown Supplier'), []).append(row)

        for supplier, supplier_rows in grouped.items():
            elements.append(Paragraph(f"Supplier: {supplier}", ParagraphStyle('InboundH2', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#0b5394'))))
            data = [["Invoice Date", "Invoice No", "File", "Item", "Qty", "Unit Price", "Line Total", "Invoice Total"]]
            for row in supplier_rows:
                data.append([
                    self._pdf_paragraph(row.get('document_date', ''), cell_style, width=12),
                    self._pdf_paragraph(row.get('invoice_number', ''), cell_style, width=16),
                    self._pdf_paragraph(row.get('file_name', ''), cell_style, width=20),
                    self._pdf_paragraph(row.get('item_description', ''), cell_style, width=28),
                    str(row.get('quantity', 0)),
                    f"{row.get('unit_price', 0.0):.2f}",
                    f"{row.get('line_total', 0.0):.2f}",
                    f"{row.get('grand_total', 0.0):.2f}",
                ])
            table = Table(data, colWidths=[65, 85, 140, 270, 40, 65, 70, 75])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b5394')),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ]))
            elements.append(table)
            elements.append(Spacer(1, 12))

        try:
            doc.build(elements)
            try:
                webbrowser.open("file://" + os.path.abspath(filepath))
            except Exception:
                pass
            messagebox.showinfo("Export PDF", f"Inbound report exported to:\n{filepath}")
        except Exception as exc:
            messagebox.showerror("Export PDF", f"Failed to export PDF:\n{exc}")


class SupplierInvoiceItemDialog:
    def __init__(self, parent, item_data=None):
        self.result = None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Supplier Invoice Item")
        self.dialog.geometry("520x260")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(False, False)

        item_data = item_data or {}
        self.description_var = tk.StringVar(value=str(item_data.get('description') or ""))
        self.quantity_var = tk.StringVar(value=str(item_data.get('quantity') if item_data.get('quantity') is not None else "1"))
        self.unit_price_var = tk.StringVar(value=str(item_data.get('unit_price') if item_data.get('unit_price') is not None else "0"))
        self.total_var = tk.StringVar(value=str(item_data.get('total') if item_data.get('total') is not None else "0"))

        self.setup_ui()

    def setup_ui(self):
        main = ttk.Frame(self.dialog, padding=12)
        main.pack(fill='both', expand=True)

        ttk.Label(main, text="Description:").grid(row=0, column=0, sticky='nw')
        self.description_text = scrolledtext.ScrolledText(main, height=5, width=42)
        self.description_text.grid(row=0, column=1, sticky='ew', pady=(0, 10))
        self.description_text.insert("1.0", self.description_var.get())

        ttk.Label(main, text="Quantity:").grid(row=1, column=0, sticky='w')
        ttk.Entry(main, textvariable=self.quantity_var, width=18).grid(row=1, column=1, sticky='w', pady=(0, 8))

        ttk.Label(main, text="Unit Price:").grid(row=2, column=0, sticky='w')
        ttk.Entry(main, textvariable=self.unit_price_var, width=18).grid(row=2, column=1, sticky='w', pady=(0, 8))

        ttk.Label(main, text="Line Total:").grid(row=3, column=0, sticky='w')
        ttk.Entry(main, textvariable=self.total_var, width=18).grid(row=3, column=1, sticky='w')

        buttons = ttk.Frame(main)
        buttons.grid(row=4, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(buttons, text="Save", command=self.on_save).pack(side='left', padx=4)
        ttk.Button(buttons, text="Cancel", command=self.dialog.destroy).pack(side='left', padx=4)

        main.columnconfigure(1, weight=1)

    def on_save(self):
        description = self.description_text.get("1.0", tk.END).strip()
        if not description:
            messagebox.showwarning("Supplier Invoice Item", "Description is required.")
            return
        try:
            quantity = float(self.quantity_var.get().strip() or "0")
            unit_price = float(self.unit_price_var.get().strip() or "0")
            total = float(self.total_var.get().strip() or "0")
        except Exception:
            messagebox.showwarning("Supplier Invoice Item", "Quantity, unit price, and total must be valid numbers.")
            return
        self.result = {
            'description': description,
            'quantity': quantity,
            'unit_price': unit_price,
            'total': total,
            'taxable': False,
        }
        self.dialog.destroy()


class SupplierSummaryReportDialog:
    def __init__(self, parent, inventory_manager, default_supplier_id=None):
        self.inventory_manager = inventory_manager
        self.default_supplier_id = default_supplier_id or ""
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Supplier Summary Report")
        self.dialog.geometry("1280x820")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        self._supplier_name_to_id = {'All Suppliers': ''}
        self._summary_cache = {'supplier_summaries': [], 'invoice_rows': [], 'supplier_count': 0, 'invoice_count': 0, 'grand_total': 0.0}
        self.setup_ui()
        try:
            self.dialog.after(0, self.generate_report)
        except Exception:
            pass

    def _parse_date(self, value):
        value = str(value or "").strip()
        if not value:
            return None
        try:
            return datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")
        except Exception:
            return value

    def setup_ui(self):
        main = ttk.Frame(self.dialog, padding=10)
        main.pack(fill='both', expand=True)

        header = ttk.Frame(main)
        header.pack(fill='x', pady=(0, 8))
        ttk.Label(header, text="Supplier Summary Report", font=('Helvetica', 16, 'bold')).pack(side='left')

        controls = ttk.LabelFrame(main, text="Filters", padding=10)
        controls.pack(fill='x', pady=(0, 10))

        supplier_records = self.inventory_manager.load_supplier_records()
        supplier_names = ['All Suppliers']
        for record in supplier_records:
            supplier_names.append(record.get('name', ''))
            self._supplier_name_to_id[record.get('name', '')] = record.get('supplier_id', '')

        default_name = 'All Suppliers'
        for record in supplier_records:
            if record.get('supplier_id') == self.default_supplier_id:
                default_name = record.get('name', 'All Suppliers')
                break

        ttk.Label(controls, text="Supplier:").grid(row=0, column=0, sticky='w')
        self.supplier_var = tk.StringVar(value=default_name)
        ttk.Combobox(controls, textvariable=self.supplier_var, values=supplier_names, width=34, state='readonly').grid(row=0, column=1, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="Start Date:").grid(row=0, column=2, sticky='w')
        self.start_date_var = tk.StringVar(value="")
        ttk.Entry(controls, textvariable=self.start_date_var, width=12).grid(row=0, column=3, sticky='w', padx=(6, 18))

        ttk.Label(controls, text="End Date:").grid(row=0, column=4, sticky='w')
        self.end_date_var = tk.StringVar(value="")
        ttk.Entry(controls, textvariable=self.end_date_var, width=12).grid(row=0, column=5, sticky='w', padx=(6, 18))

        btns = ttk.Frame(controls)
        btns.grid(row=0, column=6, sticky='e')
        ttk.Button(btns, text="Generate", command=self.generate_report).pack(side='left', padx=4)
        ttk.Button(btns, text="Export PDF", command=self.export_pdf).pack(side='left', padx=4)
        ttk.Button(btns, text="Close", command=self.dialog.destroy).pack(side='left', padx=4)
        controls.columnconfigure(6, weight=1)

        top = ttk.PanedWindow(main, orient=tk.VERTICAL)
        top.pack(fill='both', expand=True)

        summary_frame = ttk.LabelFrame(top, text="Supplier Totals", padding=10)
        detail_frame = ttk.LabelFrame(top, text="Invoice Breakdown", padding=10)
        notes_frame = ttk.LabelFrame(top, text="Summary Text", padding=10)
        top.add(summary_frame, weight=2)
        top.add(detail_frame, weight=3)
        top.add(notes_frame, weight=2)

        summary_columns = ('Supplier', 'Invoices', 'First Invoice', 'Last Invoice', 'Total')
        self.summary_tree = ttk.Treeview(summary_frame, columns=summary_columns, show='headings', height=8)
        for col in summary_columns:
            self.summary_tree.heading(col, text=col)
            self.summary_tree.column(col, width=120, stretch=True)
        self.summary_tree.column('Supplier', width=240)
        self.summary_tree.column('Total', width=150)
        self.summary_tree.pack(side='left', fill='both', expand=True)
        summary_scroll = ttk.Scrollbar(summary_frame, orient='vertical', command=self.summary_tree.yview)
        self.summary_tree.configure(yscrollcommand=summary_scroll.set)
        summary_scroll.pack(side='right', fill='y')

        invoice_columns = ('Supplier', 'Invoice No', 'Invoice Date', 'Subtotal', 'Tax', 'Invoice Total', 'Items', 'File')
        self.invoice_tree = ttk.Treeview(detail_frame, columns=invoice_columns, show='headings', height=10)
        for col in invoice_columns:
            self.invoice_tree.heading(col, text=col)
            self.invoice_tree.column(col, width=110, stretch=True)
        self.invoice_tree.column('Supplier', width=220)
        self.invoice_tree.column('File', width=260)
        self.invoice_tree.column('Invoice No', width=150)
        self.invoice_tree.pack(side='left', fill='both', expand=True)
        invoice_yscroll = ttk.Scrollbar(detail_frame, orient='vertical', command=self.invoice_tree.yview)
        invoice_xscroll = ttk.Scrollbar(detail_frame, orient='horizontal', command=self.invoice_tree.xview)
        self.invoice_tree.configure(yscrollcommand=invoice_yscroll.set, xscrollcommand=invoice_xscroll.set)
        self.invoice_tree.pack(side='left', fill='both', expand=True)
        invoice_yscroll.pack(side='right', fill='y')
        invoice_xscroll.pack(side='bottom', fill='x')

        self.summary_text = scrolledtext.ScrolledText(notes_frame, height=10, width=80)
        self.summary_text.pack(fill='both', expand=True)

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(main, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w').pack(fill='x', pady=(8, 0))

    def generate_report(self):
        start = self._parse_date(self.start_date_var.get())
        end = self._parse_date(self.end_date_var.get())
        supplier_name = self.supplier_var.get()
        supplier_id = self._supplier_name_to_id.get(supplier_name, '')
        report = self.inventory_manager.get_supplier_invoice_summary(start_date=start, end_date=end, supplier_id=supplier_id or None)
        self._summary_cache = report

        for child in self.summary_tree.get_children():
            self.summary_tree.delete(child)
        for child in self.invoice_tree.get_children():
            self.invoice_tree.delete(child)

        for summary in report.get('supplier_summaries', []):
            invoice_rows = summary.get('invoice_rows') or []
            first_invoice = invoice_rows[0].get('grand_total', 0.0) if invoice_rows else 0.0
            second_invoice = invoice_rows[1].get('grand_total', 0.0) if len(invoice_rows) > 1 else None
            first_text = f"AED {first_invoice:.2f}" if invoice_rows else "-"
            if second_invoice is not None:
                first_text = f"{first_text} | 2nd: AED {second_invoice:.2f}"
            self.summary_tree.insert('', 'end', values=(
                summary.get('supplier_name', ''),
                summary.get('invoice_count', 0),
                first_text,
                invoice_rows[-1].get('document_date', '-') if invoice_rows else '-',
                f"AED {summary.get('total_amount', 0.0):.2f}",
            ))

        for row in report.get('invoice_rows', []):
            self.invoice_tree.insert('', 'end', values=(
                row.get('supplier_name', ''),
                row.get('invoice_number', ''),
                row.get('document_date', ''),
                f"AED {row.get('subtotal', 0.0):.2f}",
                f"AED {row.get('tax_amount', 0.0):.2f}",
                f"AED {row.get('grand_total', 0.0):.2f}",
                row.get('item_count', 0),
                row.get('file_name', ''),
            ))

        self.summary_text.delete("1.0", tk.END)
        lines = []
        for summary in report.get('supplier_summaries', []):
            invoice_rows = summary.get('invoice_rows') or []
            lines.append(f"Supplier: {summary.get('supplier_name', 'Unknown Supplier')}")
            lines.append(f"Number of invoices: {summary.get('invoice_count', 0)}")
            if invoice_rows:
                for index, invoice_row in enumerate(invoice_rows, start=1):
                    lines.append(
                        f"Invoice {index}: {invoice_row.get('invoice_number', '') or invoice_row.get('document_id', '')} | "
                        f"Date: {invoice_row.get('document_date', '') or '-'} | "
                        f"Cost: AED {invoice_row.get('grand_total', 0.0):.2f}"
                    )
            else:
                lines.append("No invoices found.")
            lines.append(f"Total for this supplier: AED {summary.get('total_amount', 0.0):.2f}")
            lines.append("-" * 72)
        if not lines:
            lines.append("No supplier invoices found for the selected filters.")
        lines.append(f"Suppliers: {report.get('supplier_count', 0)}")
        lines.append(f"Invoices: {report.get('invoice_count', 0)}")
        lines.append(f"Grand total: AED {report.get('grand_total', 0.0):.2f}")
        self.summary_text.insert("1.0", "\n".join(lines))

        self.status_var.set(
            f"Suppliers: {report.get('supplier_count', 0)} | "
            f"Invoices: {report.get('invoice_count', 0)} | "
            f"Grand Total: AED {report.get('grand_total', 0.0):.2f}"
        )

    def export_pdf(self):
        report = dict(self._summary_cache or {})
        supplier_summaries = list(report.get('supplier_summaries') or [])
        invoice_rows = list(report.get('invoice_rows') or [])
        if not supplier_summaries and not invoice_rows:
            messagebox.showwarning("Export PDF", "No supplier summary rows to export.")
            return

        reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
        os.makedirs(reports_folder, exist_ok=True)
        filepath = os.path.join(reports_folder, f"Supplier_Summary_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

        styles = getSampleStyleSheet()
        title_style = ParagraphStyle('SupplierSummaryTitle', parent=styles['Heading1'], fontSize=16, alignment=1, textColor=colors.HexColor("#2c3e50"))
        subtitle_style = ParagraphStyle('SupplierSummarySub', parent=styles['Normal'], fontSize=9, alignment=1, textColor=colors.grey)
        section_style = ParagraphStyle('SupplierSummarySection', parent=styles['Heading2'], fontSize=12, textColor=colors.HexColor('#0b5394'))
        body_style = ParagraphStyle('SupplierSummaryBody', parent=styles['BodyText'], fontSize=8, leading=10)
        doc = SimpleDocTemplate(filepath, pagesize=landscape(A4), leftMargin=0.35 * inch, rightMargin=0.35 * inch, topMargin=0.4 * inch, bottomMargin=0.4 * inch)

        supplier_label = self.supplier_var.get().strip() or "All Suppliers"
        start_label = self.start_date_var.get().strip() or "Any"
        end_label = self.end_date_var.get().strip() or "Any"

        elements = []
        elements.extend(build_pdf_logo_flowables(self.inventory_manager.data_folder, width=1.0 * inch, height=1.0 * inch, spacer_height=0.08 * inch))
        elements.extend([
            Paragraph("Supplier Summary Report", title_style),
            Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}", subtitle_style),
            Paragraph(f"Filters: Supplier = {supplier_label} | Start = {start_label} | End = {end_label}", subtitle_style),
            Spacer(1, 10),
        ])

        totals_table = [["Supplier", "Invoices", "First Invoice", "Second Invoice", "Total"]]
        for summary in supplier_summaries:
            rows = summary.get('invoice_rows') or []
            first_invoice = f"AED {float(rows[0].get('grand_total', 0.0)):.2f}" if rows else "-"
            second_invoice = f"AED {float(rows[1].get('grand_total', 0.0)):.2f}" if len(rows) > 1 else "-"
            totals_table.append([
                summary.get('supplier_name', ''),
                str(summary.get('invoice_count', 0)),
                first_invoice,
                second_invoice,
                f"AED {float(summary.get('total_amount', 0.0)):.2f}",
            ])

        elements.append(Paragraph("Supplier Totals", section_style))
        totals = Table(totals_table, colWidths=[3.2 * inch, 0.9 * inch, 1.3 * inch, 1.3 * inch, 1.5 * inch])
        totals.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b5394')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(totals)
        elements.append(Spacer(1, 12))

        elements.append(Paragraph("Invoice Breakdown", section_style))
        detail_table = [["Supplier", "Invoice No", "Date", "Subtotal", "Tax", "Invoice Total", "Items", "File"]]
        for row in invoice_rows:
            detail_table.append([
                Paragraph(str(row.get('supplier_name', '') or "-"), body_style),
                Paragraph(str(row.get('invoice_number', '') or "-"), body_style),
                Paragraph(str(row.get('document_date', '') or "-"), body_style),
                Paragraph(f"AED {float(row.get('subtotal', 0.0)):.2f}", body_style),
                Paragraph(f"AED {float(row.get('tax_amount', 0.0)):.2f}", body_style),
                Paragraph(f"AED {float(row.get('grand_total', 0.0)):.2f}", body_style),
                Paragraph(str(row.get('item_count', 0)), body_style),
                Paragraph(str(row.get('file_name', '') or "-"), body_style),
            ])
        details = Table(detail_table, colWidths=[2.0 * inch, 1.4 * inch, 0.9 * inch, 1.0 * inch, 0.9 * inch, 1.1 * inch, 0.6 * inch, 2.6 * inch])
        details.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#0b5394')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('GRID', (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ('FONTSIZE', (0, 1), (-1, -1), 7.5),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 4),
            ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        elements.append(details)
        elements.append(Spacer(1, 12))

        elements.append(Paragraph("Summary", section_style))
        elements.append(Paragraph(f"Suppliers: {report.get('supplier_count', 0)}", body_style))
        elements.append(Paragraph(f"Invoices: {report.get('invoice_count', 0)}", body_style))
        elements.append(Paragraph(f"Grand Total: AED {float(report.get('grand_total', 0.0)):.2f}", body_style))

        try:
            doc.build(elements)
            try:
                webbrowser.open("file://" + os.path.abspath(filepath))
            except Exception:
                pass
            messagebox.showinfo("Export PDF", f"Supplier summary report exported to:\n{filepath}")
        except Exception as exc:
            messagebox.showerror("Export PDF", f"Failed to export PDF:\n{exc}")


class SuppliersDialog:
    def __init__(self, parent, inventory_manager, on_data_changed=None):
        self.inventory_manager = inventory_manager
        self.on_data_changed = on_data_changed
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Suppliers")
        self.dialog.geometry("1380x860")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)

        self.selected_supplier_id = ""
        self.selected_document_id = ""
        self.status_var = tk.StringVar(value="Ready")

        self.supplier_id_var = tk.StringVar(value="")
        self.name_var = tk.StringVar(value="")
        self.contact_var = tk.StringVar(value="")
        self.phone_var = tk.StringVar(value="")
        self.email_var = tk.StringVar(value="")
        self.address_var = tk.StringVar(value="")
        self.doc_invoice_number_var = tk.StringVar(value="")
        self.doc_invoice_date_var = tk.StringVar(value="")
        self.doc_due_date_var = tk.StringVar(value="")
        self.doc_subtotal_var = tk.StringVar(value="0.00")
        self.doc_tax_var = tk.StringVar(value="0.00")
        self.doc_total_var = tk.StringVar(value="0.00")

        self.setup_ui()
        self.refresh_supplier_list()

    def setup_ui(self):
        main = ttk.Frame(self.dialog)
        main.pack(fill='both', expand=True)

        scroll_frame = ScrollableFrame(main)
        scroll_frame.pack(fill='both', expand=True)
        content = ttk.Frame(scroll_frame.scrollable_frame, padding=10)
        content.pack(fill='both', expand=True)

        header = ttk.Frame(content)
        header.pack(fill='x', pady=(0, 8))
        ttk.Label(header, text="Suppliers Dashboard", font=('Helvetica', 18, 'bold')).pack(side='left')
        ttk.Button(header, text="Summary", command=self.open_summary_report).pack(side='right', padx=4)
        ttk.Button(header, text="Inbound Report", command=self.open_inbound_report).pack(side='right', padx=4)
        ttk.Button(header, text="Outbound Report", command=self.open_outbound_report).pack(side='right', padx=4)
        ttk.Button(header, text="Close", command=self.dialog.destroy).pack(side='right', padx=4)

        body = ttk.PanedWindow(content, orient=tk.HORIZONTAL)
        body.pack(fill='both', expand=True)

        left = ttk.Frame(body, padding=(0, 0, 8, 0))
        right = ttk.Frame(body)
        body.add(left, weight=1)
        body.add(right, weight=3)

        suppliers_frame = ttk.LabelFrame(left, text="Suppliers", padding=8)
        suppliers_frame.pack(fill='both', expand=True)

        self.suppliers_tree = ttk.Treeview(suppliers_frame, columns=('Name', 'Contact', 'Phone', 'Email'), show='headings', height=18)
        for col in ('Name', 'Contact', 'Phone', 'Email'):
            self.suppliers_tree.heading(col, text=col)
            self.suppliers_tree.column(col, width=130, stretch=True)
        self.suppliers_tree.column('Name', width=160)
        self.suppliers_tree.pack(side='left', fill='both', expand=True)
        suppliers_scroll = ttk.Scrollbar(suppliers_frame, orient='vertical', command=self.suppliers_tree.yview)
        self.suppliers_tree.configure(yscrollcommand=suppliers_scroll.set)
        suppliers_scroll.pack(side='right', fill='y')
        self.suppliers_tree.bind('<<TreeviewSelect>>', self.on_supplier_select)

        supplier_btns = ttk.Frame(left)
        supplier_btns.pack(fill='x', pady=(8, 0))
        ttk.Button(supplier_btns, text="New Supplier", command=self.clear_supplier_form).pack(side='left', padx=4)
        ttk.Button(supplier_btns, text="Save Supplier", command=self.save_supplier).pack(side='left', padx=4)
        ttk.Button(supplier_btns, text="Delete Supplier", command=self.delete_supplier).pack(side='left', padx=4)

        details_frame = ttk.LabelFrame(right, text="Supplier Details", padding=10)
        details_frame.pack(fill='x')

        ttk.Label(details_frame, text="Supplier ID:").grid(row=0, column=0, sticky='w')
        ttk.Entry(details_frame, textvariable=self.supplier_id_var, width=20, state='readonly').grid(row=0, column=1, sticky='w', padx=(6, 18))
        ttk.Label(details_frame, text="Name:").grid(row=0, column=2, sticky='w')
        ttk.Entry(details_frame, textvariable=self.name_var, width=34).grid(row=0, column=3, sticky='ew', padx=(6, 0))

        ttk.Label(details_frame, text="Contact Person:").grid(row=1, column=0, sticky='w', pady=(8, 0))
        ttk.Entry(details_frame, textvariable=self.contact_var, width=24).grid(row=1, column=1, sticky='ew', padx=(6, 18), pady=(8, 0))
        ttk.Label(details_frame, text="Phone:").grid(row=1, column=2, sticky='w', pady=(8, 0))
        ttk.Entry(details_frame, textvariable=self.phone_var, width=24).grid(row=1, column=3, sticky='ew', padx=(6, 0), pady=(8, 0))

        ttk.Label(details_frame, text="Email:").grid(row=2, column=0, sticky='w', pady=(8, 0))
        ttk.Entry(details_frame, textvariable=self.email_var, width=24).grid(row=2, column=1, sticky='ew', padx=(6, 18), pady=(8, 0))
        ttk.Label(details_frame, text="Address:").grid(row=2, column=2, sticky='w', pady=(8, 0))
        ttk.Entry(details_frame, textvariable=self.address_var, width=34).grid(row=2, column=3, sticky='ew', padx=(6, 0), pady=(8, 0))

        ttk.Label(details_frame, text="Notes:").grid(row=3, column=0, sticky='nw', pady=(8, 0))
        self.notes_text = scrolledtext.ScrolledText(details_frame, height=4, width=60)
        self.notes_text.grid(row=3, column=1, columnspan=3, sticky='nsew', padx=(6, 0), pady=(8, 0))
        details_frame.columnconfigure(1, weight=1)
        details_frame.columnconfigure(3, weight=1)

        uploads_frame = ttk.LabelFrame(right, text="Uploaded Supplier Invoices", padding=10)
        uploads_frame.pack(fill='both', expand=True, pady=(10, 0))

        upload_btns = ttk.Frame(uploads_frame)
        upload_btns.pack(fill='x', pady=(0, 8))
        ttk.Button(upload_btns, text="Upload Invoice", command=self.upload_invoice).pack(side='left', padx=4)
        ttk.Button(upload_btns, text="Save Changes", command=self.save_selected_document_changes).pack(side='left', padx=4)
        ttk.Button(upload_btns, text="Open File", command=self.open_selected_document).pack(side='left', padx=4)
        ttk.Button(upload_btns, text="Delete Upload", command=self.delete_selected_document).pack(side='left', padx=4)
        ttk.Button(upload_btns, text="Refresh", command=self.refresh_documents).pack(side='left', padx=4)

        self.uploads_tree = ttk.Treeview(
            uploads_frame,
            columns=('Uploaded', 'File', 'Type', 'Invoice No', 'Invoice Date', 'Items', 'Total', 'Status'),
            show='headings',
            height=10
        )
        for col in ('Uploaded', 'File', 'Type', 'Invoice No', 'Invoice Date', 'Items', 'Total', 'Status'):
            self.uploads_tree.heading(col, text=col)
            self.uploads_tree.column(col, width=100, stretch=True)
        self.uploads_tree.column('File', width=200)
        self.uploads_tree.column('Invoice No', width=120)
        self.uploads_tree.column('Total', width=90)
        self.uploads_tree.pack(fill='both', expand=True)
        self.uploads_tree.bind('<<TreeviewSelect>>', self.on_document_select)

        editor_frame = ttk.LabelFrame(uploads_frame, text="Edit Selected Invoice", padding=10)
        editor_frame.pack(fill='x', pady=(10, 0))

        ttk.Label(editor_frame, text="Invoice No:").grid(row=0, column=0, sticky='w')
        ttk.Entry(editor_frame, textvariable=self.doc_invoice_number_var, width=24).grid(row=0, column=1, sticky='w', padx=(6, 18))
        ttk.Label(editor_frame, text="Invoice Date:").grid(row=0, column=2, sticky='w')
        ttk.Entry(editor_frame, textvariable=self.doc_invoice_date_var, width=16).grid(row=0, column=3, sticky='w', padx=(6, 18))
        ttk.Label(editor_frame, text="Due Date:").grid(row=0, column=4, sticky='w')
        ttk.Entry(editor_frame, textvariable=self.doc_due_date_var, width=16).grid(row=0, column=5, sticky='w', padx=(6, 0))

        ttk.Label(editor_frame, text="Subtotal:").grid(row=1, column=0, sticky='w', pady=(8, 0))
        ttk.Entry(editor_frame, textvariable=self.doc_subtotal_var, width=16).grid(row=1, column=1, sticky='w', padx=(6, 18), pady=(8, 0))
        ttk.Label(editor_frame, text="Tax:").grid(row=1, column=2, sticky='w', pady=(8, 0))
        ttk.Entry(editor_frame, textvariable=self.doc_tax_var, width=16).grid(row=1, column=3, sticky='w', padx=(6, 18), pady=(8, 0))
        ttk.Label(editor_frame, text="Grand Total:").grid(row=1, column=4, sticky='w', pady=(8, 0))
        ttk.Entry(editor_frame, textvariable=self.doc_total_var, width=16).grid(row=1, column=5, sticky='w', padx=(6, 0), pady=(8, 0))

        ttk.Label(editor_frame, text="Invoice Notes:").grid(row=2, column=0, sticky='nw', pady=(8, 0))
        self.doc_notes_text = scrolledtext.ScrolledText(editor_frame, height=4, width=70)
        self.doc_notes_text.grid(row=2, column=1, columnspan=5, sticky='ew', padx=(6, 0), pady=(8, 0))
        for idx in range(1, 6):
            editor_frame.columnconfigure(idx, weight=1)

        lower = ttk.PanedWindow(uploads_frame, orient=tk.HORIZONTAL)
        lower.pack(fill='both', expand=True, pady=(10, 0))

        items_frame = ttk.LabelFrame(lower, text="Extracted Items", padding=8)
        summary_frame = ttk.LabelFrame(lower, text="Extraction Summary", padding=8)
        lower.add(items_frame, weight=2)
        lower.add(summary_frame, weight=1)

        self.items_tree = ttk.Treeview(items_frame, columns=('Description', 'Qty', 'Unit Price', 'Total'), show='headings', height=8)
        for col in ('Description', 'Qty', 'Unit Price', 'Total'):
            self.items_tree.heading(col, text=col)
            self.items_tree.column(col, width=100, stretch=True)
        self.items_tree.column('Description', width=260)
        self.items_tree.pack(fill='both', expand=True)
        self.items_tree.bind('<Double-1>', lambda _event: self.edit_selected_invoice_item())

        item_btns = ttk.Frame(items_frame)
        item_btns.pack(fill='x', pady=(8, 0))
        ttk.Button(item_btns, text="Add Item", command=self.add_invoice_item).pack(side='left', padx=4)
        ttk.Button(item_btns, text="Edit Item", command=self.edit_selected_invoice_item).pack(side='left', padx=4)
        ttk.Button(item_btns, text="Delete Item", command=self.delete_selected_invoice_item).pack(side='left', padx=4)

        self.summary_text = scrolledtext.ScrolledText(summary_frame, height=10, width=36)
        self.summary_text.pack(fill='both', expand=True)

        ttk.Label(self.dialog, textvariable=self.status_var, relief=tk.SUNKEN, anchor='w').pack(fill='x', side='bottom')

    def _get_selected_supplier_id(self):
        selection = self.suppliers_tree.selection()
        if selection:
            return selection[0]
        return self.selected_supplier_id or self.supplier_id_var.get().strip()

    def _set_status(self, message):
        self.status_var.set(message)

    def _clear_document_editor(self):
        self.doc_invoice_number_var.set("")
        self.doc_invoice_date_var.set("")
        self.doc_due_date_var.set("")
        self.doc_subtotal_var.set("0.00")
        self.doc_tax_var.set("0.00")
        self.doc_total_var.set("0.00")
        self.doc_notes_text.delete("1.0", tk.END)

    def _short_display_text(self, value, width=36):
        text = str(value or "").strip()
        if len(text) <= width:
            return text
        return text[: max(0, width - 3)] + "..."

    def _get_selected_document_record(self):
        document_id = self.selected_document_id
        if not document_id:
            selection = self.uploads_tree.selection()
            if selection:
                document_id = selection[0]
        if not document_id:
            return None
        return self.inventory_manager.get_supplier_invoice_import(document_id)

    def _get_current_invoice_items(self):
        items = []
        for row_id in self.items_tree.get_children():
            values = self.items_tree.item(row_id).get('values') or []
            if len(values) < 4:
                continue
            try:
                quantity = float(str(values[1]).replace("AED", "").strip())
            except Exception:
                quantity = 0.0
            try:
                unit_price = float(str(values[2]).replace("AED", "").strip())
            except Exception:
                unit_price = 0.0
            try:
                total = float(str(values[3]).replace("AED", "").strip())
            except Exception:
                total = 0.0
            items.append({
                'description': str(values[0]).strip(),
                'quantity': quantity,
                'unit_price': unit_price,
                'total': total,
                'taxable': False,
            })
        return items

    def _recalculate_document_totals_from_items(self):
        items = self._get_current_invoice_items()
        subtotal = sum(float(item.get('total', 0.0)) for item in items)
        tax = 0.0
        try:
            tax = float(self.doc_tax_var.get().strip() or "0")
        except Exception:
            tax = 0.0
        self.doc_subtotal_var.set(f"{subtotal:.2f}")
        self.doc_total_var.set(f"{subtotal + tax:.2f}")

    def clear_supplier_form(self):
        self.selected_supplier_id = ""
        self.selected_document_id = ""
        self.supplier_id_var.set("")
        self.name_var.set("")
        self.contact_var.set("")
        self.phone_var.set("")
        self.email_var.set("")
        self.address_var.set("")
        self.notes_text.delete("1.0", tk.END)
        self._clear_document_editor()
        for row in self.suppliers_tree.selection():
            self.suppliers_tree.selection_remove(row)
        self.refresh_documents()
        self._set_status("Ready to add a new supplier.")

    def refresh_supplier_list(self):
        records = self.inventory_manager.load_supplier_records()
        current_id = self._get_selected_supplier_id()
        for child in self.suppliers_tree.get_children():
            self.suppliers_tree.delete(child)
        for record in records:
            supplier_id = record.get('supplier_id', '')
            self.suppliers_tree.insert('', 'end', iid=supplier_id, values=(
                record.get('name', ''),
                record.get('contact_person', ''),
                record.get('phone', ''),
                record.get('email', ''),
            ))
        if current_id and self.suppliers_tree.exists(current_id):
            self.suppliers_tree.selection_set(current_id)
            self.suppliers_tree.focus(current_id)
            self.suppliers_tree.see(current_id)
            self.load_supplier_into_form(current_id)
        elif records:
            first_id = records[0].get('supplier_id', '')
            if first_id and self.suppliers_tree.exists(first_id):
                self.suppliers_tree.selection_set(first_id)
                self.load_supplier_into_form(first_id)
        else:
            self.clear_supplier_form()
        self._set_status(f"Loaded {len(records)} suppliers.")

    def load_supplier_into_form(self, supplier_id):
        record = self.inventory_manager.get_supplier_record(supplier_id)
        if not record:
            return
        self.selected_supplier_id = supplier_id
        self.supplier_id_var.set(record.get('supplier_id', ''))
        self.name_var.set(record.get('name', ''))
        self.contact_var.set(record.get('contact_person', ''))
        self.phone_var.set(record.get('phone', ''))
        self.email_var.set(record.get('email', ''))
        self.address_var.set(record.get('address', ''))
        self.notes_text.delete("1.0", tk.END)
        self.notes_text.insert("1.0", record.get('notes', ''))
        self.refresh_documents()

    def on_supplier_select(self, event=None):
        selection = self.suppliers_tree.selection()
        if not selection:
            return
        self.load_supplier_into_form(selection[0])

    def save_supplier(self):
        name = self.name_var.get().strip()
        notes = self.notes_text.get("1.0", tk.END).strip()
        if self.selected_supplier_id:
            ok, message = self.inventory_manager.update_supplier_record(
                self.selected_supplier_id,
                name,
                contact_person=self.contact_var.get().strip(),
                phone=self.phone_var.get().strip(),
                email=self.email_var.get().strip(),
                address=self.address_var.get().strip(),
                notes=notes,
            )
        else:
            ok, message, record = self.inventory_manager.add_supplier_record(
                name,
                contact_person=self.contact_var.get().strip(),
                phone=self.phone_var.get().strip(),
                email=self.email_var.get().strip(),
                address=self.address_var.get().strip(),
                notes=notes,
            )
            if ok and record:
                self.selected_supplier_id = record.get('supplier_id', '')
        if not ok:
            messagebox.showerror("Suppliers", message)
            return
        self.refresh_supplier_list()
        self._set_status(message)
        if callable(self.on_data_changed):
            self.on_data_changed()

    def delete_supplier(self):
        supplier_id = self._get_selected_supplier_id()
        if not supplier_id:
            messagebox.showwarning("Suppliers", "Please select a supplier first.")
            return
        record = self.inventory_manager.get_supplier_record(supplier_id)
        if not record:
            messagebox.showerror("Suppliers", "Supplier not found.")
            return
        if not messagebox.askyesno("Delete Supplier", f"Delete supplier '{record.get('name', '')}'?\n\nThis only works when no inventory items or uploaded invoices are still linked."):
            return
        ok, message = self.inventory_manager.delete_supplier_record(supplier_id)
        if not ok:
            messagebox.showerror("Suppliers", message)
            return
        self.clear_supplier_form()
        self.refresh_supplier_list()
        self._set_status(message)
        if callable(self.on_data_changed):
            self.on_data_changed()

    def refresh_documents(self):
        supplier_id = self._get_selected_supplier_id()
        documents = self.inventory_manager.load_supplier_invoice_imports()
        if supplier_id:
            documents = [doc for doc in documents if doc.get('supplier_id') == supplier_id]
        current_document_id = self.selected_document_id

        for child in self.uploads_tree.get_children():
            self.uploads_tree.delete(child)
        for child in self.items_tree.get_children():
            self.items_tree.delete(child)
        self.summary_text.delete("1.0", tk.END)
        self.selected_document_id = ""
        self._clear_document_editor()

        for doc in documents:
            self.uploads_tree.insert('', 'end', iid=doc.get('document_id', ''), values=(
                doc.get('upload_date', ''),
                self._short_display_text(doc.get('file_name', ''), 34),
                doc.get('file_type', ''),
                self._short_display_text(doc.get('invoice_number', ''), 18),
                doc.get('document_date', ''),
                len(doc.get('items') or []),
                f"AED {float(doc.get('grand_total', 0.0)):.2f}",
                doc.get('status', ''),
            ))
        if current_document_id and self.uploads_tree.exists(current_document_id):
            self.uploads_tree.selection_set(current_document_id)
            self.uploads_tree.focus(current_document_id)
            self.on_document_select()
        self._set_status(f"Loaded {len(documents)} uploaded invoices.")

    def on_document_select(self, event=None):
        selection = self.uploads_tree.selection()
        if not selection:
            return
        document_id = selection[0]
        self.selected_document_id = document_id
        record = next((doc for doc in self.inventory_manager.load_supplier_invoice_imports() if doc.get('document_id') == document_id), None)
        if not record:
            return

        self.doc_invoice_number_var.set(record.get('invoice_number', ''))
        self.doc_invoice_date_var.set(record.get('document_date', ''))
        self.doc_due_date_var.set(record.get('due_date', ''))
        self.doc_subtotal_var.set(f"{float(record.get('subtotal', 0.0)):.2f}")
        self.doc_tax_var.set(f"{float(record.get('tax_amount', 0.0)):.2f}")
        self.doc_total_var.set(f"{float(record.get('grand_total', 0.0)):.2f}")
        self.doc_notes_text.delete("1.0", tk.END)
        self.doc_notes_text.insert("1.0", record.get('notes', ''))

        for child in self.items_tree.get_children():
            self.items_tree.delete(child)
        for index, item in enumerate(record.get('items') or [], start=1):
            self.items_tree.insert('', 'end', iid=f"item_{index}", values=(
                item.get('description', ''),
                item.get('quantity', 0),
                f"AED {float(item.get('unit_price', 0.0)):.2f}",
                f"AED {float(item.get('total', 0.0)):.2f}",
            ))

        summary_lines = [
            f"Supplier: {record.get('supplier_name', '')}",
            f"File: {record.get('file_name', '')}",
            f"Stored Path: {record.get('stored_path', '')}",
            f"Invoice Number: {record.get('invoice_number', '')}",
            f"Invoice Date: {record.get('document_date', '')}",
            f"Due Date: {record.get('due_date', '')}",
            f"Subtotal: AED {float(record.get('subtotal', 0.0)):.2f}",
            f"Tax: AED {float(record.get('tax_amount', 0.0)):.2f}",
            f"Grand Total: AED {float(record.get('grand_total', 0.0)):.2f}",
            f"Status: {record.get('status', '')}",
            "",
            "Warnings:",
        ]
        warnings = record.get('warnings') or []
        if warnings:
            summary_lines.extend(f"- {warning}" for warning in warnings)
        else:
            summary_lines.append("- None")
        summary_lines.extend(["", "Errors:"])
        errors = record.get('errors') or []
        if errors:
            summary_lines.extend(f"- {error}" for error in errors)
        else:
            summary_lines.append("- None")
        if record.get('notes'):
            summary_lines.extend(["", "Notes:", record.get('notes', '')])

        self.summary_text.delete("1.0", tk.END)
        self.summary_text.insert("1.0", "\n".join(summary_lines))

    def save_selected_document_changes(self):
        record = self._get_selected_document_record()
        if not record:
            messagebox.showwarning("Suppliers", "Please select an uploaded invoice first.")
            return

        try:
            subtotal = float(self.doc_subtotal_var.get().strip() or "0")
            tax_amount = float(self.doc_tax_var.get().strip() or "0")
            grand_total = float(self.doc_total_var.get().strip() or "0")
        except Exception:
            messagebox.showwarning("Suppliers", "Subtotal, tax, and grand total must be valid numbers.")
            return

        ok, message, updated_record = self.inventory_manager.update_supplier_invoice_import(
            record.get('document_id'),
            invoice_number=self.doc_invoice_number_var.get().strip(),
            document_date=self.doc_invoice_date_var.get().strip(),
            due_date=self.doc_due_date_var.get().strip(),
            subtotal=subtotal,
            tax_amount=tax_amount,
            grand_total=grand_total,
            notes=self.doc_notes_text.get("1.0", tk.END).strip(),
            items=self._get_current_invoice_items(),
        )
        if not ok:
            messagebox.showerror("Suppliers", message)
            return
        self.selected_document_id = updated_record.get('document_id', '') if updated_record else record.get('document_id')
        self.refresh_documents()
        self._set_status(message)
        if callable(self.on_data_changed):
            self.on_data_changed()
        messagebox.showinfo("Suppliers", message)

    def add_invoice_item(self):
        if not self._get_selected_document_record():
            messagebox.showwarning("Suppliers", "Please select an uploaded invoice first.")
            return
        dialog = SupplierInvoiceItemDialog(self.dialog)
        self.dialog.wait_window(dialog.dialog)
        if not dialog.result:
            return
        next_id = f"item_{len(self.items_tree.get_children()) + 1}"
        self.items_tree.insert('', 'end', iid=next_id, values=(
            dialog.result.get('description', ''),
            dialog.result.get('quantity', 0),
            f"AED {float(dialog.result.get('unit_price', 0.0)):.2f}",
            f"AED {float(dialog.result.get('total', 0.0)):.2f}",
        ))
        self._recalculate_document_totals_from_items()

    def edit_selected_invoice_item(self):
        selection = self.items_tree.selection()
        if not selection:
            messagebox.showwarning("Suppliers", "Please select an invoice item first.")
            return
        values = self.items_tree.item(selection[0]).get('values') or []
        item_data = {
            'description': values[0] if len(values) > 0 else "",
            'quantity': str(values[1]).replace("AED", "").strip() if len(values) > 1 else "0",
            'unit_price': str(values[2]).replace("AED", "").strip() if len(values) > 2 else "0",
            'total': str(values[3]).replace("AED", "").strip() if len(values) > 3 else "0",
        }
        dialog = SupplierInvoiceItemDialog(self.dialog, item_data=item_data)
        self.dialog.wait_window(dialog.dialog)
        if not dialog.result:
            return
        self.items_tree.item(selection[0], values=(
            dialog.result.get('description', ''),
            dialog.result.get('quantity', 0),
            f"AED {float(dialog.result.get('unit_price', 0.0)):.2f}",
            f"AED {float(dialog.result.get('total', 0.0)):.2f}",
        ))
        self._recalculate_document_totals_from_items()

    def delete_selected_invoice_item(self):
        selection = self.items_tree.selection()
        if not selection:
            messagebox.showwarning("Suppliers", "Please select an invoice item first.")
            return
        for row_id in selection:
            self.items_tree.delete(row_id)
        self._recalculate_document_totals_from_items()

    def upload_invoice(self):
        supplier_id = self._get_selected_supplier_id()
        if not supplier_id:
            messagebox.showwarning("Suppliers", "Please select or save a supplier before uploading an invoice.")
            return
        file_path = filedialog.askopenfilename(
            title="Select Supplier Invoice",
            filetypes=[("Invoice files", "*.pdf *.xlsx *.xls"), ("PDF files", "*.pdf"), ("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        ok, message, record = self.inventory_manager.save_supplier_invoice_import(supplier_id, file_path)
        if not ok:
            messagebox.showerror("Suppliers", message)
            return
        self.refresh_documents()
        if record and self.uploads_tree.exists(record.get('document_id', '')):
            self.uploads_tree.selection_set(record.get('document_id', ''))
            self.on_document_select()
        messagebox.showinfo("Suppliers", message)
        self._set_status(message)
        if callable(self.on_data_changed):
            self.on_data_changed()

    def open_selected_document(self):
        document_id = self.selected_document_id
        if not document_id:
            selection = self.uploads_tree.selection()
            if selection:
                document_id = selection[0]
        if not document_id:
            messagebox.showwarning("Suppliers", "Please select an uploaded invoice first.")
            return
        record = next((doc for doc in self.inventory_manager.load_supplier_invoice_imports() if doc.get('document_id') == document_id), None)
        if not record or not record.get('stored_path'):
            messagebox.showerror("Suppliers", "The stored file could not be found.")
            return
        try:
            webbrowser.open("file://" + os.path.abspath(record.get('stored_path')))
        except Exception as exc:
            messagebox.showerror("Suppliers", f"Could not open file:\n{exc}")

    def delete_selected_document(self):
        document_id = self.selected_document_id
        if not document_id:
            selection = self.uploads_tree.selection()
            if selection:
                document_id = selection[0]
        if not document_id:
            messagebox.showwarning("Suppliers", "Please select an uploaded invoice first.")
            return
        if not messagebox.askyesno("Delete Upload", "Delete the selected uploaded invoice and its stored file?"):
            return
        ok, message = self.inventory_manager.delete_supplier_invoice_import(document_id)
        if not ok:
            messagebox.showerror("Suppliers", message)
            return
        self.refresh_documents()
        self._set_status(message)

    def open_inbound_report(self):
        SupplierInboundReportDialog(self.dialog, self.inventory_manager, default_supplier_id=self._get_selected_supplier_id())

    def open_outbound_report(self):
        supplier_name = self.name_var.get().strip() or "All Suppliers"
        SupplierOutboundReportDialog(self.dialog, self.inventory_manager, default_supplier=supplier_name if self._get_selected_supplier_id() else "All Suppliers")

    def open_summary_report(self):
        SupplierSummaryReportDialog(self.dialog, self.inventory_manager, default_supplier_id=self._get_selected_supplier_id())


class DeliveryNoteDialog:
    def __init__(self, parent, inventory_manager):
        self.inventory_manager = inventory_manager
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Create Delivery Note")
        self.dialog.geometry("1000x800")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        self.items = []
        self.last_generated_pdf_path = None
        self._setup_ui()

    def _setup_ui(self):
        attach_clock_label(self.dialog)
        bottom_bar = ttk.Frame(self.dialog, padding="10")
        bottom_bar.pack(fill="x", side="bottom")
        ttk.Button(bottom_bar, text="Print", command=self.print_delivery_note).pack(side="right", padx=(6, 0))
        ttk.Button(bottom_bar, text="Create Delivery Note", command=self.create_delivery_note).pack(side="right")
        ttk.Button(bottom_bar, text="Cancel", command=self.dialog.destroy).pack(side="right", padx=(0, 6))

        content_container = ttk.Frame(self.dialog)
        content_container.pack(fill="both", expand=True)
        canvas = tk.Canvas(content_container, highlightthickness=0)
        yscroll = ttk.Scrollbar(content_container, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=yscroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        main = ttk.Frame(canvas, padding="10")
        canvas_window = canvas.create_window((0, 0), window=main, anchor="nw")

        def _on_main_configure(_event=None):
            try:
                canvas.configure(scrollregion=canvas.bbox("all"))
            except Exception:
                pass

        def _on_canvas_configure(event):
            try:
                canvas.itemconfigure(canvas_window, width=event.width)
            except Exception:
                pass

        main.bind("<Configure>", _on_main_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            try:
                delta = int(-1 * (event.delta / 120))
            except Exception:
                delta = -1 if getattr(event, "delta", 0) > 0 else 1
            try:
                canvas.yview_scroll(delta, "units")
            except Exception:
                pass

        def _bind_mousewheel(_event=None):
            try:
                canvas.bind_all("<MouseWheel>", _on_mousewheel)
            except Exception:
                pass

        def _unbind_mousewheel(_event=None):
            try:
                canvas.unbind_all("<MouseWheel>")
            except Exception:
                pass

        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)

        top = ttk.Frame(main)
        top.pack(fill="x")

        left_header = ttk.Frame(top)
        left_header.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ttk.Label(left_header, text="HopePharma", font=("Helvetica", 22, "bold"), foreground="#2c5aa0").pack(anchor="w")
        ttk.Label(left_header, text="Delivery Note", font=("Helvetica", 14, "bold"), foreground="#0b5394").pack(anchor="w", pady=(4,0))

        right_header = ttk.LabelFrame(top, text="Order Details", padding="8")
        right_header.pack(side="right", fill="x")

        self.order_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.order_no = tk.StringVar()
        self.delivery_note_no = tk.StringVar(value=datetime.now().strftime("DN%y%m%d%H%M%S"))
        self.customer_id = tk.StringVar()
        self.dispatch_date = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.delivery_method = tk.StringVar()

        labels = ["Order Date", "Order #", "Delivery Note #", "Customer ID", "Dispatch Date", "Delivery Method"]
        vars_ = [self.order_date, self.order_no, self.delivery_note_no, self.customer_id, self.dispatch_date, self.delivery_method]
        for i, (lbl, var) in enumerate(zip(labels, vars_)):
            ttk.Label(right_header, text=lbl + ":").grid(row=i, column=0, sticky="w", pady=2, padx=(0,6))
            ttk.Entry(right_header, textvariable=var, width=28).grid(row=i, column=1, sticky="w", pady=2)

        addr_frame = ttk.Frame(main)
        addr_frame.pack(fill="x", pady=10)

        ship = ttk.LabelFrame(addr_frame, text="Shipping Address", padding="8")
        ship.pack(side="left", fill="both", expand=True)

        self.ship_name = tk.StringVar()
        self.ship_company = tk.StringVar()
        self.ship_street = tk.StringVar()
        self.ship_cityzip = tk.StringVar()
        self.ship_phone = tk.StringVar()

        ship_fields = [("Name", self.ship_name), ("Company Name", self.ship_company), ("Street Address", self.ship_street),
                       ("City, ST  ZIP Code", self.ship_cityzip), ("Phone", self.ship_phone)]
        
        for i, (lbl, var) in enumerate(ship_fields):
            ttk.Label(ship, text=lbl + ":").grid(row=i, column=0, sticky="w", pady=2)
            ttk.Entry(ship, textvariable=var, width=50).grid(row=i, column=1, sticky="w", pady=2)

        items_frame = ttk.LabelFrame(main, text="Items", padding="8")
        items_frame.pack(fill="both", expand=True)
        cols = ("Item #", "Description", "Batch Number", "Quantity")
        self.tree = ttk.Treeview(items_frame, columns=cols, show="headings", height=10)
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120 if c != "Description" else 300)
        yscroll = ttk.Scrollbar(items_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")

        add_frame = ttk.LabelFrame(main, text="Add Item", padding="8")
        add_frame.pack(fill="x", pady=6)
        self.line_item = tk.StringVar()
        self.line_desc = tk.StringVar()
        self.line_batch = tk.StringVar()
        self.line_qty = tk.StringVar(value="1")
        
        # Row 1
        row1 = ttk.Frame(add_frame)
        row1.pack(fill="x", pady=(0, 5))
        
        ttk.Button(row1, text="Select Item", command=self.open_item_selector).pack(side="left", padx=(0, 10))
        
        ttk.Label(row1, text="Item #:").pack(side="left")
        self.e_item = ttk.Entry(row1, textvariable=self.line_item, width=15)
        self.e_item.pack(side="left", padx=(5, 15))
        
        ttk.Label(row1, text="Description:").pack(side="left")
        self.e_desc = ttk.Entry(row1, textvariable=self.line_desc, width=40)
        self.e_desc.pack(side="left", fill="x", expand=True, padx=5)
        
        # Row 2
        row2 = ttk.Frame(add_frame)
        row2.pack(fill="x")
        
        ttk.Label(row2, text="Batch Number:").pack(side="left")
        self.e_batch = ttk.Entry(row2, textvariable=self.line_batch, width=15)
        self.e_batch.pack(side="left", padx=(5, 15))
        
        ttk.Label(row2, text="Quantity:").pack(side="left")
        self.e_qty = ttk.Entry(row2, textvariable=self.line_qty, width=10)
        self.e_qty.pack(side="left", padx=(5, 15))
        
        ttk.Button(row2, text="Add Line (Enter)", command=self.add_line).pack(side="left", padx=5)
        ttk.Button(row2, text="Remove Selected", command=self.remove_selected).pack(side="left", padx=5)
        
        for e in (self.e_item, self.e_desc, self.e_batch, self.e_qty):
            e.bind("<Return>", lambda event: self.add_line())


    def open_item_selector(self):
        try:
            selector = tk.Toplevel(self.dialog)
            selector.title("Select Inventory Item")
            selector.geometry("600x400")
            selector.transient(self.dialog)
            selector.grab_set()
            
            # Search
            search_frame = ttk.Frame(selector, padding="5")
            search_frame.pack(fill="x")
            ttk.Label(search_frame, text="Search:").pack(side="left")
            search_var = tk.StringVar()
            search_entry = ttk.Entry(search_frame, textvariable=search_var)
            search_entry.pack(side="left", fill="x", expand=True, padx=5)
            search_entry.focus_set()
            
            # List
            columns = ("ID", "Name", "Stock", "Price")
            tree = ttk.Treeview(selector, columns=columns, show="headings")
            for col in columns:
                tree.heading(col, text=col)
                tree.column(col, width=80 if col != "Name" else 250)
            
            yscroll = ttk.Scrollbar(selector, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=yscroll.set)
            
            tree.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            yscroll.pack(side="right", fill="y", pady=5)
            
            def load_items(query=""):
                for item in tree.get_children():
                    tree.delete(item)
                query = query.lower()
                for item in self.inventory_manager.items:
                    if query in item.name.lower() or query in str(item.item_id).lower():
                        tree.insert("", "end", values=(
                            item.item_id, 
                            item.name, 
                            item.quantity,
                            f"{item.selling_price:.2f}" if item.selling_price else ""
                        ))
            
            load_items()
            
            def on_search(*args):
                load_items(search_var.get())
                
            search_var.trace("w", on_search)
            
            def on_select(event=None):
                sel = tree.selection()
                if not sel:
                    return
                item_vals = tree.item(sel[0])["values"]
                self.line_item.set(item_vals[0])
                self.line_desc.set(item_vals[1])
                selector.destroy()
                try:
                    self.e_batch.focus_set()
                except Exception:
                    pass
                
            tree.bind("<Double-1>", on_select)
            tree.bind("<Return>", on_select)
            ttk.Button(selector, text="Select", command=on_select).pack(side="bottom", pady=5)
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not open item selector: {e}")

    def add_line(self):
        try:
            qty = int(self.line_qty.get() or "0")
        except Exception:
            messagebox.showwarning("Warning", "Enter valid quantity")
            return
        if not self.line_desc.get().strip():
            messagebox.showwarning("Warning", "Enter description")
            return
        batch = self.line_batch.get().strip()
        vals = (self.line_item.get().strip(), self.line_desc.get().strip(), batch, qty)
        self.items.append(vals)
        self.tree.insert("", "end", values=vals)
        self.line_item.set("")
        self.line_desc.set("")
        self.line_batch.set("")
        self.line_qty.set("1")

    def remove_selected(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        self.tree.delete(sel[0])
        if 0 <= idx < len(self.items):
            del self.items[idx]

    def _generate_pdf(self, open_after=True):
        if not self.items:
            messagebox.showwarning("Warning", "Add at least one item")
            return None
        notes_folder = os.path.join(self.inventory_manager.data_folder, "DeliveryNotes")
        if not os.path.exists(notes_folder):
            os.makedirs(notes_folder)
        fname = f"DeliveryNote_{self.delivery_note_no.get()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        path = os.path.join(notes_folder, fname)
        doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=0.5 * inch, rightMargin=0.5 * inch, topMargin=0.5 * inch, bottomMargin=0.5 * inch)
        elements = []
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("DNTitle", parent=styles["Heading1"], fontSize=22, textColor=colors.HexColor("#0b5394"), spaceAfter=6)
        sub_style = ParagraphStyle("DNSub", parent=styles["Normal"], fontSize=9, textColor=colors.grey, spaceAfter=6)
        hdr_style = ParagraphStyle("Hdr", parent=styles["Normal"], fontSize=10, textColor=colors.white, alignment=0)
        val_style = ParagraphStyle("Val", parent=styles["Normal"], fontSize=10)

        logo_path = resolve_logo_path(self.inventory_manager.data_folder)
        left = []
        if logo_path and os.path.exists(logo_path):
            try:
                left.append(RLImage(logo_path, width=1.4 * inch, height=1.0 * inch))
            except Exception:
                pass
        left.append(Paragraph("HopePharma", title_style))
        left.append(Paragraph("Delivery Note", sub_style))
        left_box = KeepInFrame(3 * inch, 1.6 * inch, left)

        od = [["Order Date", self.order_date.get()], ["Order #", self.order_no.get()], ["Delivery Note #", self.delivery_note_no.get()],
              ["Customer ID", self.customer_id.get()], ["Despatch Date", self.dispatch_date.get()], ["Delivery Method", self.delivery_method.get()]]
        det = Table(od, colWidths=[1.6 * inch, 2.0 * inch])
        det.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#0b5394")),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#0b5394")),
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#0b5394")),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.white),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        header_tbl = Table([[left_box, det]], colWidths=[3.2 * inch, 3.4 * inch])
        elements.append(header_tbl)
        elements.append(Spacer(1, 0.2 * inch))

        ship_rows = [["Name", self.ship_name.get()], ["Company", self.ship_company.get()], ["Street", self.ship_street.get()],
                     ["City/ZIP", self.ship_cityzip.get()], ["Phone", self.ship_phone.get()]]
        
        ship_tbl = Table([["Shipping Address"]], colWidths=[7.0 * inch])
        ship_tbl.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b5394")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                                      ("ALIGN", (0, 0), (-1, -1), "LEFT")]))
        
        ship_vals = Table(ship_rows, colWidths=[1.5 * inch, 5.5 * inch])
        ship_vals.setStyle(TableStyle([("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey), ("BOX", (0, 0), (-1, -1), 0.25, colors.lightgrey)]))
        
        elements.append(ship_tbl)
        elements.append(ship_vals)
        elements.append(Spacer(1, 0.2 * inch))

        items_header = ["Item #", "Description", "Batch Number", "Quantity"]
        items_rows = [items_header]
        for it in self.items:
            # Wrap description in Paragraph to ensure it wraps within the column width
            # and doesn't overlap into the Batch Number column
            desc_para = Paragraph(str(it[1]), val_style)
            items_rows.append([it[0], desc_para, it[2], it[3]])
        # Adjusted column widths to prevent description from overlapping batch number
        items_table = Table(items_rows, colWidths=[1.0 * inch, 3.2 * inch, 1.8 * inch, 1.2 * inch])
        items_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17395d")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("ALIGN", (2, 1), (-1, -1), "CENTER"),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.lightgrey),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#17395d")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("WORDWRAP", (0, 0), (-1, -1), True),
        ]))
        elements.append(items_table)
        elements.append(Spacer(1, 0.2 * inch))

        notes = Paragraph("Thank you for your business.", sub_style)
        elements.append(notes)
        doc.build(elements)
        self.last_generated_pdf_path = path
        if open_after:
            try:
                webbrowser.open("file://" + os.path.abspath(path))
            except Exception:
                pass
        return path

    def create_delivery_note(self):
        path = self._generate_pdf(open_after=False)
        if path:
            # Save delivery note to system
            note_data = {
                "delivery_note_number": self.delivery_note_no.get(),
                "order_date": self.order_date.get(),
                "order_number": self.order_no.get(),
                "customer_id": self.customer_id.get(),
                "despatch_date": self.dispatch_date.get(),
                "delivery_method": self.delivery_method.get(),
                "shipping_address": {
                    "name": self.ship_name.get(),
                    "company": self.ship_company.get(),
                    "street": self.ship_street.get(),
                    "city_zip": self.ship_cityzip.get(),
                    "phone": self.ship_phone.get()
                },
                "items": [
                    {
                        "item_id": it[0],
                        "description": it[1],
                        "batch_number": it[2],
                        "quantity": it[3]
                    }
                    for it in self.items
                ],
                "generated_pdf_path": path,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            
            if self.inventory_manager.save_delivery_note(note_data):
                messagebox.showinfo("Success", f"Delivery Note created and saved to system:\n{path}")
            else:
                messagebox.showinfo("Success", f"Delivery Note created (PDF only, save failed):\n{path}")

    def print_delivery_note(self):
        path = self.last_generated_pdf_path
        if not path or not os.path.exists(path):
            path = self._generate_pdf(open_after=True)
        if not path:
            return
        try:
            if shutil.which("lp"):
                subprocess.run(["lp", path], check=False)
                messagebox.showinfo("Print", "Sent to printer.")
            else:
                webbrowser.open("file://" + os.path.abspath(path))
                messagebox.showinfo("Print", "Opened PDF. Use your PDF viewer to print.")
        except Exception:
            try:
                webbrowser.open("file://" + os.path.abspath(path))
            except Exception:
                pass

class InventorySystem:
    def __init__(self, parent, data_folder, invoice_manager=None):
        self.parent = parent
        self.data_folder = data_folder
        self.invoice_manager = invoice_manager
        self.inventory_manager = InventoryManager(data_folder)
        
        # Track file modification times for auto-refresh
        self.last_sync_times = {}
        self.files_to_watch = [
            self.inventory_manager.inventory_file,
            self.inventory_manager.sales_file,
            self.inventory_manager.invoices_file,
            self.inventory_manager.categories_file,
            self.inventory_manager.cost_types_file,
            self.inventory_manager.suppliers_file,
            self.inventory_manager.supplier_records_file,
            self.inventory_manager.supplier_invoice_imports_file,
        ]
        
        self.setup_ui()
        self.start_auto_refresh()
    
    def start_auto_refresh(self):
        """Start the auto-refresh cycle"""
        try:
            if not self.parent.winfo_exists():
                return
            self.check_for_updates()
            # Check every 5 seconds
            self.parent.after(5000, self.start_auto_refresh)
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
                print("DEBUG: Changes detected in data files, refreshing inventory UI...")
                # Reload data in manager
                self.inventory_manager.items = self.inventory_manager.load_inventory()
                self.inventory_manager.sales = self.inventory_manager.load_sales()
                self.inventory_manager.categories = self.inventory_manager.load_categories()
                self.inventory_manager.cost_types = self.inventory_manager.load_cost_types()
                # Refresh UI
                self.refresh_all()
        except Exception as e:
            print(f"Auto-refresh error: {e}")
    
    def setup_ui(self):
        """Setup the inventory system UI"""
        # Clear existing widgets
        for widget in self.parent.winfo_children():
            widget.destroy()
        
        # Back to main menu button
        back_frame = ttk.Frame(self.parent)
        back_frame.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(back_frame, text="← Back to Main Menu", 
                  command=self.show_main_menu).pack(anchor='w')
        
        # Main frame
        main_frame = ttk.Frame(self.parent, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Header
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill='x', pady=(0, 20))

        title_frame = ttk.Frame(header_frame)
        title_frame.pack()

        self.main_logo_image = None
        try:
            logo_path = resolve_logo_path(self.data_folder)
            if logo_path and os.path.exists(logo_path):
                img = Image.open(logo_path)
                img = img.resize((48, 48), Image.LANCZOS)
                self.main_logo_image = ImageTk.PhotoImage(img)
        except Exception:
            self.main_logo_image = None

        if self.main_logo_image:
            tk.Label(title_frame, image=self.main_logo_image).pack(side='left', padx=(0, 10))

        ttk.Label(title_frame, text="Inventory Management System",
                 font=('Helvetica', 20, 'bold')).pack(side='left')

        attach_clock_label(header_frame)
        
        # Dashboard - store reference to dashboard frame
        self.dashboard_frame = ttk.LabelFrame(main_frame, text="Inventory Dashboard", padding="15")
        self.dashboard_frame.pack(fill='x', pady=(0, 20))
        
        # Dashboard metrics
        self.refresh_dashboard()
        
        # Search and controls
        controls_frame = ttk.Frame(main_frame)
        controls_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(controls_frame, text="Search:").pack(side='left', padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(controls_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side='left', padx=(0, 10))
        self.search_entry.bind('<KeyRelease>', self.on_search)

        # Advanced search options (multi-field)
        self.search_by_id = tk.BooleanVar(value=True)
        self.search_by_name = tk.BooleanVar(value=True)
        self.search_by_supplier = tk.BooleanVar(value=True)
        self.search_by_batch = tk.BooleanVar(value=False)
        self.search_by_date = tk.BooleanVar(value=False)

        search_options_frame = ttk.Frame(main_frame)
        search_options_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(search_options_frame, text="Search in:").pack(side='left', padx=(0, 5))
        ttk.Checkbutton(search_options_frame, text="ID", variable=self.search_by_id, command=self.on_search).pack(side='left')
        ttk.Checkbutton(search_options_frame, text="Name", variable=self.search_by_name, command=self.on_search).pack(side='left', padx=(5, 0))
        ttk.Checkbutton(search_options_frame, text="Supplier", variable=self.search_by_supplier, command=self.on_search).pack(side='left', padx=(5, 0))
        ttk.Checkbutton(search_options_frame, text="Batch No", variable=self.search_by_batch, command=self.on_search).pack(side='left', padx=(5, 0))
        ttk.Checkbutton(search_options_frame, text="Date", variable=self.search_by_date, command=self.on_search).pack(side='left', padx=(5, 10))

        ttk.Label(search_options_frame, text="Date from (YYYY-MM-DD):").pack(side='left', padx=(10, 5))
        self.search_date_from = tk.StringVar()
        self.search_date_to = tk.StringVar()
        self.search_date_from_entry = ttk.Entry(search_options_frame, textvariable=self.search_date_from, width=10)
        self.search_date_from_entry.pack(side='left')
        self.search_date_from_entry.bind('<KeyRelease>', self.on_search)
        ttk.Label(search_options_frame, text="to").pack(side='left', padx=(5, 5))
        self.search_date_to_entry = ttk.Entry(search_options_frame, textvariable=self.search_date_to, width=10)
        self.search_date_to_entry.pack(side='left')
        self.search_date_to_entry.bind('<KeyRelease>', self.on_search)
        
        # Sort options
        ttk.Label(controls_frame, text="Sort by:").pack(side='left', padx=(20, 5))
        self.sort_var = tk.StringVar(value="name_asc")
        sort_combo = ttk.Combobox(controls_frame, textvariable=self.sort_var, 
                                 values=["Name (A-Z)", "Name (Z-A)", "Quantity (High-Low)", "Quantity (Low-High)", 
                                        "Price (High-Low)", "Price (Low-High)", "Category (A-Z)"],
                                 state="readonly", width=15)
        sort_combo.pack(side='left', padx=(0, 10))
        sort_combo.bind('<<ComboboxSelected>>', self.on_sort_change)
        
        # Buttons arranged in two rows (6 + 6)
        buttons_container = ttk.Frame(controls_frame)
        buttons_container.pack(side='right')
        row1 = ttk.Frame(buttons_container)
        row1.pack(fill='x')
        row2 = ttk.Frame(buttons_container)
        row2.pack(fill='x', pady=(4, 0))
        
        ttk.Button(row1, text="Add Item", command=self.add_item).pack(side='left', padx=5)
        ttk.Button(row1, text="Sell Item", command=self.sell_item).pack(side='left', padx=5)
        ttk.Button(row1, text="Edit Item", command=self.edit_item).pack(side='left', padx=5)
        ttk.Button(row1, text="Manage Costs", command=self.manage_costs).pack(side='left', padx=5)
        ttk.Button(row1, text="Delete Item", command=self.delete_item).pack(side='left', padx=5)
        ttk.Button(row1, text="Inventory Invoices", command=self.show_inventory_invoices).pack(side='left', padx=5)
        
        ttk.Button(row2, text="Saved Sales", command=self.show_saved_sales).pack(side='left', padx=5)
        ttk.Button(row2, text="Delivery Note", command=self.show_delivery_note).pack(side='left', padx=5)
        ttk.Button(row2, text="Suppliers", command=self.show_suppliers).pack(side='left', padx=5)
        ttk.Button(row2, text="Supplier Report", command=self.show_reports).pack(side='left', padx=5)
        ttk.Button(row2, text="Import Excel", command=self.import_excel).pack(side='left', padx=5)
        ttk.Button(row2, text="Import Invoice Excel", command=self.import_invoice_excel).pack(side='left', padx=5)
        ttk.Button(row2, text="Export Excel", command=self.export_inventory_view_excel).pack(side='left', padx=5)
        ttk.Button(row2, text="Product List", command=self.export_product_status_list_excel).pack(side='left', padx=5)
        ttk.Button(row2, text="Print View", command=self.print_inventory_view).pack(side='left', padx=5)
        ttk.Button(row2, text="Refresh", command=self.refresh_all).pack(side='left', padx=5)
        ttk.Button(row2, text="Help", command=self.show_help).pack(side='left', padx=5)
        
        # Inventory list
        list_frame = ttk.LabelFrame(main_frame, text="Inventory Items", padding="10")
        list_frame.pack(fill='both', expand=True)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill='both', expand=True)
        
        columns = ('ID', 'Name', 'Category', 'Cost/Item', 'Price', 'Quantity', 'BatchNo', 'ExpiryDate', 'Supplier', 'DateEntry', 'InvoiceDate', 'Status')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15, selectmode='extended')
        
        # Define headings
        self.tree.heading('ID', text='Item ID')
        self.tree.heading('Name', text='Item Name')
        self.tree.heading('Category', text='Category')
        self.tree.heading('Cost/Item', text='Cost/Item (AED)')
        self.tree.heading('Price', text='Price (AED)')
        self.tree.heading('Quantity', text='Quantity')
        self.tree.heading('BatchNo', text='Batch No')
        self.tree.heading('ExpiryDate', text='Expiry Date')
        self.tree.heading('Supplier', text='Supplier')
        self.tree.heading('DateEntry', text='Date of Entry')
        self.tree.heading('InvoiceDate', text='Invoice Date')
        self.tree.heading('Status', text='Status')
        
        # Define columns
        self.tree.column('ID', width=100)
        self.tree.column('Name', width=150)
        self.tree.column('Category', width=120)
        self.tree.column('Cost/Item', width=100)
        self.tree.column('Price', width=80)
        self.tree.column('Quantity', width=70)
        self.tree.column('BatchNo', width=90)
        self.tree.column('ExpiryDate', width=100)
        self.tree.column('Supplier', width=120)
        self.tree.column('DateEntry', width=100)
        self.tree.column('InvoiceDate', width=100)
        self.tree.column('Status', width=100)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Double click to edit
        self.tree.bind('<Double-1>', lambda e: self.edit_item())
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                              relief=tk.SUNKEN, font=('Helvetica', 10))
        status_bar.pack(fill='x', side='bottom', pady=(10, 0))
        
        # Load initial data
        self.refresh_items()
    
    def show_help(self):
        messagebox.showinfo("Help", "Inventory Management System\n\nUse the buttons to manage inventory items, sales, and reports.\nFor support, contact the administrator.")
    
    def refresh_dashboard(self):
        """Refresh the dashboard metrics"""
        if not hasattr(self, 'dashboard_frame'):
            return
            
        # Clear existing metrics
        for widget in self.dashboard_frame.winfo_children():
            widget.destroy()
        
        # Get current metrics
        metrics = self.inventory_manager.get_inventory_summary()
        
        metrics_frame = ttk.Frame(self.dashboard_frame)
        metrics_frame.pack(fill='x')
        
        # Total Items
        ttk.Label(metrics_frame, text=f"Total Items\n{metrics['total_items']}", 
                 font=('Helvetica', 14, 'bold'), foreground='#2c5aa0').grid(row=0, column=0, padx=20, pady=10)
        
        # Low Stock
        low_stock_color = 'red' if metrics['low_stock_count'] > 0 else 'green'
        ttk.Label(metrics_frame, text=f"Low Stock\n{metrics['low_stock_count']}", 
                 font=('Helvetica', 14, 'bold'), foreground=low_stock_color).grid(row=0, column=1, padx=20, pady=10)
        
        # Expired Items
        expired_color = 'red' if metrics['expired_count'] > 0 else 'green'
        ttk.Label(metrics_frame, text=f"Expired\n{metrics['expired_count']}", 
                 font=('Helvetica', 14, 'bold'), foreground=expired_color).grid(row=0, column=2, padx=20, pady=10)
        
        # Total Value
        ttk.Label(metrics_frame, text=f"Total Value\nAED {metrics['total_value']:,.2f}", 
                 font=('Helvetica', 14, 'bold'), foreground='#27ae60').grid(row=0, column=3, padx=20, pady=10)
        
        # Potential Profit
        ttk.Label(metrics_frame, text=f"Potential Profit\nAED {metrics['total_potential_profit']:,.2f}", 
                 font=('Helvetica', 14, 'bold'), foreground='#e67e22').grid(row=0, column=4, padx=20, pady=10)
    
    def refresh_all(self):
        """Refresh both dashboard and items"""
        self.refresh_dashboard()
        self.refresh_items()
    
    def import_excel(self):
        """Handle Excel import process"""
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Error", "The 'openpyxl' library is required for Excel import.\nPlease install it using: pip install openpyxl")
            return

        file_path = filedialog.askopenfilename(
            title="Select Excel Inventory File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        
        if not file_path:
            return
            
        # Show confirmation/guide
        guide_msg = "Excel Import Guide:\n\n" \
                    "- First row should be headers\n" \
                    "- Required: 'Item Name'\n" \
                    "- Recommended: 'Quantity', 'Cost Per Item', 'Selling Price'\n" \
                    "- Optional: 'Category', 'Supplier', 'Batch Number', 'Expiry Date'\n\n" \
                    "The system will try to match headers automatically.\n" \
                    "Continue with import?"
        
        if not messagebox.askyesno("Confirm Import", guide_msg):
            return
            
        success, message = self.inventory_manager.import_from_excel(file_path)
        
        if success:
            messagebox.showinfo("Success", message)
            self.refresh_all()
        else:
            messagebox.showerror("Error", message)

    def import_invoice_excel(self):
        """Import an invoice Excel sheet and create an invoice in the system"""
        if not OPENPYXL_AVAILABLE:
            messagebox.showerror("Error", "The 'openpyxl' library is required for Excel import.\nPlease install it using: pip install openpyxl")
            return

        file_path = filedialog.askopenfilename(
            title="Select Excel Invoice File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        invoice_dict, warnings, errors = self.inventory_manager.parse_invoice_excel(file_path)
        if not invoice_dict:
            msg = "Could not import invoice from this Excel file."
            if warnings:
                msg += "\n\n" + "\n".join(warnings[:10])
            if errors:
                msg += "\n\n" + "\n".join(errors[:10])
            messagebox.showerror("Error", msg)
            return

        if warnings:
            warn_msg = "Some information could not be detected, but the invoice can still be imported:\n\n" + "\n".join(warnings[:10]) + "\n\nContinue?"
            if not messagebox.askyesno("Import with Warnings", warn_msg):
                return

        try:
            if self.invoice_manager and hasattr(self.invoice_manager, 'generate_invoice_id'):
                invoice_id = self.invoice_manager.generate_invoice_id()
            else:
                invoice_id = self.inventory_manager.generate_invoice_id()
            invoice_dict['invoice_id'] = invoice_id
        except Exception:
            invoice_dict['invoice_id'] = self.inventory_manager.generate_invoice_id()

        saved_ok = False
        try:
            if self.invoice_manager and hasattr(self.invoice_manager, 'add_invoice_from_dict'):
                saved_ok = bool(self.invoice_manager.add_invoice_from_dict(invoice_dict))
            else:
                saved_ok = bool(self.inventory_manager._append_invoice(invoice_dict))
        except Exception:
            try:
                saved_ok = bool(self.inventory_manager._append_invoice(invoice_dict))
            except Exception:
                saved_ok = False

        if not saved_ok:
            messagebox.showerror("Error", "Failed to save the imported invoice.")
            return

        if errors:
            report_file = os.path.join(os.path.dirname(file_path), "import_invoice_problems.txt")
            try:
                with open(report_file, 'w') as f:
                    f.write("INVOICE IMPORT PROBLEM REPORT\n")
                    f.write("=============================\n\n")
                    f.write("\n".join(errors))
            except Exception:
                report_file = None
        else:
            report_file = None

        msg = f"Invoice imported successfully.\n\nInvoice ID: {invoice_dict.get('invoice_id')}\nClient: {invoice_dict.get('client_name')}\nTotal: AED {float(invoice_dict.get('grand_total', 0.0)):.2f}"
        if report_file:
            msg += f"\n\nSome rows had problems. Report saved to:\n{report_file}"

        try:
            if messagebox.askyesno("Success", msg + "\n\nGenerate PDF now?"):
                from hope_pharma_complete import EnhancedPDFGenerator
                out_folder = Path(self.inventory_manager.data_folder) / "HopePharmaInvoices"
                try:
                    out_folder.mkdir(parents=True, exist_ok=True)
                except Exception:
                    pass
                pdf_path = EnhancedPDFGenerator(output_folder=str(out_folder)).generate_invoice_pdf(invoice_dict)
                try:
                    webbrowser.open("file://" + os.path.abspath(pdf_path))
                except Exception:
                    pass
            else:
                messagebox.showinfo("Success", msg)
        except Exception:
            messagebox.showinfo("Success", msg)
    
    def print_inventory_view(self):
        try:
            rows = self.tree.get_children()
            if not rows:
                messagebox.showwarning("Print Inventory", "No items to print")
                return

            reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)

            filename = f"Inventory_View_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            filepath = os.path.join(reports_folder, filename)

            doc = SimpleDocTemplate(
                filepath,
                pagesize=landscape(A4),
                leftMargin=0.3 * inch,
                rightMargin=0.3 * inch,
                topMargin=0.4 * inch,
                bottomMargin=0.4 * inch,
            )
            elements = []

            styles = getSampleStyleSheet()
            title_style = ParagraphStyle(
                'InvViewTitle',
                parent=styles['Heading1'],
                fontSize=18,
                textColor=colors.HexColor("#2c3e50"),
                alignment=1,
                spaceAfter=8
            )
            subtitle_style = ParagraphStyle(
                'InvViewSubtitle',
                parent=styles['Normal'],
                fontSize=9,
                textColor=colors.grey,
                alignment=1,
                spaceAfter=16
            )

            logo_path = resolve_logo_path(self.inventory_manager.data_folder)
            if logo_path and os.path.exists(logo_path):
                try:
                    logo_img = RLImage(logo_path, width=1.0*inch, height=1.0*inch)
                    elements.append(logo_img)
                    elements.append(Spacer(1, 0.1*inch))
                except Exception:
                    pass

            elements.append(Paragraph("Inventory Management System", title_style))
            elements.append(Paragraph(f"Inventory View Snapshot - Generated {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", subtitle_style))

            columns = self.tree["columns"]
            header = []
            for c in columns:
                try:
                    header.append(self.tree.heading(c)["text"])
                except Exception:
                    header.append(str(c))

            data = [header]
            cell_style = ParagraphStyle(
                'InvCell',
                parent=styles['Normal'],
                fontSize=6,
                leading=7,
                alignment=0,
            )
            for item_id in rows:
                vals = self.tree.item(item_id)["values"]
                row = []
                for col_id, v in zip(columns, vals):
                    text = str(v)
                    if col_id in ('Name', 'Category', 'Supplier'):
                        row.append(Paragraph(text.replace('\n', ' '), cell_style))
                    else:
                        row.append(text)
                data.append(row)

            col_count = len(header)
            total_width = 10.5 * inch

            base_widths = {
                'ID': 0.5,
                'Name': 1.4,
                'Category': 0.9,
                'Cost/Item': 0.7,
                'Price': 0.7,
                'Quantity': 0.5,
                'BatchNo': 0.7,
                'ExpiryDate': 0.8,
                'Supplier': 1.0,
                'DateEntry': 0.6,
                'InvoiceDate': 0.6,
                'Status': 0.4,
            }

            base_total = sum(base_widths.get(c, 0.6) for c in columns) or 1.0
            scale = total_width / base_total
            col_widths = [base_widths.get(c, 0.6) * scale for c in columns]

            table = Table(data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#34495e")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('GRID', (0, 0), (-1, -1), 0.25, colors.grey),
                ('FONTSIZE', (0, 1), (-1, -1), 6),
                ('ALIGN', (0, 1), (-1, -1), 'LEFT'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ]))

            elements.append(table)
            doc.build(elements)

            webbrowser.open('file://' + os.path.abspath(filepath))
            messagebox.showinfo("Print Inventory", f"Inventory view exported to PDF:\n{filepath}\n\nYou can print it from your PDF viewer.")
        except Exception as e:
            messagebox.showerror("Print Inventory", f"Failed to print inventory view: {e}")

    def export_inventory_view_excel(self):
        try:
            try:
                import openpyxl
                from openpyxl.styles import Font, Alignment, PatternFill
            except Exception:
                messagebox.showerror("Export Excel", "Excel export is not available in this build.")
                return

            rows = self.tree.get_children()
            if not rows:
                messagebox.showwarning("Export Excel", "No items to export")
                return

            reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)

            filename = f"Inventory_View_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = os.path.join(reports_folder, filename)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Inventory View"

            columns = list(self.tree["columns"])
            headers = []
            for c in columns:
                try:
                    headers.append(self.tree.heading(c)["text"] or str(c))
                except Exception:
                    headers.append(str(c))

            ws.append(headers)
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill("solid", fgColor="1F2937")
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for item_id in rows:
                values = list(self.tree.item(item_id).get("values") or [])
                ws.append([str(v) if v is not None else "" for v in values])

            ws.freeze_panes = "A2"

            for col_idx, header in enumerate(headers, start=1):
                max_len = len(str(header))
                for row_idx in range(2, min(ws.max_row, 500) + 1):
                    v = ws.cell(row=row_idx, column=col_idx).value
                    if v is None:
                        continue
                    max_len = max(max_len, len(str(v)))
                ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(max(10, max_len + 2), 48)

            wb.save(filepath)
            try:
                webbrowser.open('file://' + os.path.abspath(filepath))
            except Exception:
                pass
            messagebox.showinfo("Export Excel", f"Inventory view exported to Excel:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Export Excel", f"Failed to export inventory view:\n{e}")

    def export_product_status_list_excel(self):
        try:
            try:
                import openpyxl
                from openpyxl.styles import Font, Alignment, PatternFill
            except Exception:
                messagebox.showerror("Product List", "Excel export is not available in this build.")
                return

            items = list(getattr(self.inventory_manager, "items", []) or [])
            if not items:
                messagebox.showwarning("Product List", "No products found")
                return

            status_order = {"marketed": 3, "registered": 2, "stored": 1}
            products = {}
            for item in items:
                sku = str(getattr(item, "sku", "") or "").strip()
                code = str(getattr(item, "item_id", "") or "").strip()
                name = str(getattr(item, "name", "") or "").strip()
                key = (sku or code or name).strip().lower()
                if not key:
                    continue
                current = products.get(key)
                status = str(getattr(item, "product_status", "") or "Stored").strip() or "Stored"
                if current:
                    prev = str(current.get("Status") or "Stored").strip() or "Stored"
                    if status_order.get(status.lower(), 0) > status_order.get(prev.lower(), 0):
                        current["Status"] = status
                    continue
                products[key] = {
                    "Item Code": code,
                    "Product Name": name,
                    "Brand": str(getattr(item, "brand", "") or "").strip(),
                    "SKU": sku,
                    "Status": status,
                }

            rows = sorted(products.values(), key=lambda r: (str(r.get("Product Name") or "").lower(), str(r.get("Item Code") or "").lower()))
            reports_folder = os.path.join(self.inventory_manager.data_folder, "InventoryReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)

            filename = f"Product_List_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = os.path.join(reports_folder, filename)

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Products"
            headers = ["Item Code", "Product Name", "Brand", "SKU", "Status"]
            ws.append(headers)
            header_font = Font(bold=True, color="FFFFFF")
            header_fill = PatternFill("solid", fgColor="1F2937")
            for col_idx in range(1, len(headers) + 1):
                cell = ws.cell(row=1, column=col_idx)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for row in rows:
                ws.append([row.get(h, "") for h in headers])

            ws.freeze_panes = "A2"
            for col_idx, header in enumerate(headers, start=1):
                max_len = len(str(header))
                for row_idx in range(2, min(ws.max_row, 800) + 1):
                    v = ws.cell(row=row_idx, column=col_idx).value
                    if v is None:
                        continue
                    max_len = max(max_len, len(str(v)))
                ws.column_dimensions[openpyxl.utils.get_column_letter(col_idx)].width = min(max(12, max_len + 2), 60)

            wb.save(filepath)
            try:
                webbrowser.open('file://' + os.path.abspath(filepath))
            except Exception:
                pass
            messagebox.showinfo("Product List", f"Product list exported to Excel:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Product List", f"Failed to export product list:\n{e}")
    
    def show_main_menu(self):
        """Show main menu - to be implemented by main app"""
        # This will be connected to the main app's show_main_menu method
        if hasattr(self, 'show_main_menu_callback'):
            self.show_main_menu_callback()
    
    def refresh_items(self):
        """Refresh the inventory items list"""
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        items = self.inventory_manager.items
        items = self.sort_items(items, self.sort_var.get())
        
        for item in items:
            name_display = item.name
            if getattr(item, 'is_storage_item', False):
                name_display = f"📦 {item.name} (Stored)"
                
            if item.is_expired():
                status = "EXPIRED"
            elif item.is_low_stock():
                status = "LOW STOCK"
            else:
                status = "OK"
                
            cost_per_item = item.get_cost_per_item()
                
            self.tree.insert('', tk.END, values=(
                item.item_id,
                name_display,
                item.category,
                f"AED {cost_per_item:.2f}",
                f"AED {item.selling_price:.2f}" if item.selling_price else "Not set",
                item.quantity,
                item.batch_number or "",
                item.expiry_date or "",
                item.supplier,
                item.created_date,
                getattr(item, "invoice_date", None) or "",
                status
            ))
        
        self.status_var.set(f"Loaded {len(items)} inventory items")
    
    def show_delivery_note(self):
        dialog = DeliveryNoteDialog(self.parent, self.inventory_manager)
        try:
            self.parent.wait_window(dialog.dialog)
        except Exception:
            pass
        self.refresh_all()
    def sort_items(self, items, sort_option):
        """Sort items based on selected option"""
        if sort_option == "Name (A-Z)":
            return sorted(items, key=lambda x: x.name.lower())
        elif sort_option == "Name (Z-A)":
            return sorted(items, key=lambda x: x.name.lower(), reverse=True)
        elif sort_option == "Quantity (High-Low)":
            return sorted(items, key=lambda x: x.quantity, reverse=True)
        elif sort_option == "Quantity (Low-High)":
            return sorted(items, key=lambda x: x.quantity)
        elif sort_option == "Price (High-Low)":
            return sorted(items, key=lambda x: x.selling_price, reverse=True)
        elif sort_option == "Price (Low-High)":
            return sorted(items, key=lambda x: x.selling_price)
        elif sort_option == "Category (A-Z)":
            return sorted(items, key=lambda x: x.category.lower())
        else:
            return items
    
    def on_sort_change(self, event=None):
        """Handle sort change"""
        self.refresh_items()
    
    def on_search(self, event=None):
        """Handle search functionality with multi-field filters"""
        query_raw = self.search_var.get().strip()
        query = query_raw.lower()

        # Determine which fields are enabled
        by_id = getattr(self, "search_by_id", None)
        by_name = getattr(self, "search_by_name", None)
        by_supplier = getattr(self, "search_by_supplier", None)
        by_batch = getattr(self, "search_by_batch", None)
        by_date = getattr(self, "search_by_date", None)

        # If no query and no date filter, show all
        date_from_str = getattr(self, "search_date_from", tk.StringVar()).get().strip() if by_date else ""
        date_to_str = getattr(self, "search_date_to", tk.StringVar()).get().strip() if by_date else ""
        if not query and (not by_date or (not date_from_str and not date_to_str)):
            self.refresh_items()
            return

        # Prepare date range
        start_date = None
        end_date = None
        if by_date and (date_from_str or date_to_str):
            try:
                if date_from_str:
                    start_date = datetime.strptime(date_from_str, "%Y-%m-%d").date()
                if date_to_str:
                    end_date = datetime.strptime(date_to_str, "%Y-%m-%d").date()
            except Exception:
                start_date = None
                end_date = None

        for item in self.tree.get_children():
            self.tree.delete(item)

        results = []
        for item in self.inventory_manager.items:
            text_match = True
            if query:
                text_match = False
                # If no advanced flags available, fall back to manager search
                if not any([by_id, by_name, by_supplier, by_batch]):
                    base_results = self.inventory_manager.search_items(query_raw)
                    results = base_results
                    break
                if by_id and by_id.get() and query in str(item.item_id).lower():
                    text_match = True
                if by_name and by_name.get() and query in (item.name or "").lower():
                    text_match = True
                if by_supplier and by_supplier.get() and query in (item.supplier or "").lower():
                    text_match = True
                if by_batch and by_batch.get() and query in (item.batch_number or "").lower():
                    text_match = True

            date_match = True
            if by_date and by_date.get() and (start_date or end_date):
                date_match = False
                created_str = getattr(item, "created_date", "") or ""
                try:
                    if created_str:
                        created_d = datetime.strptime(created_str, "%Y-%m-%d").date()
                        if (start_date is None or created_d >= start_date) and (end_date is None or created_d <= end_date):
                            date_match = True
                except Exception:
                    date_match = False

            if text_match and date_match:
                results.append(item)

        results = self.sort_items(results, self.sort_var.get())

        for item in results:
            if item.is_expired():
                status = "EXPIRED"
            elif item.is_low_stock():
                status = "LOW STOCK"
            else:
                status = "OK"

            cost_per_item = item.get_cost_per_item()

            self.tree.insert('', tk.END, values=(
                item.item_id,
                item.name,
                item.category,
                f"AED {cost_per_item:.2f}",
                f"AED {item.selling_price:.2f}" if item.selling_price else "Not set",
                item.quantity,
                item.batch_number or "",
                item.expiry_date or "",
                item.supplier,
                item.created_date,
                getattr(item, "invoice_date", None) or "",
                status
            ))

        if results:
            if by_date and by_date.get() and (start_date or end_date):
                period_text = ""
                if start_date and end_date:
                    period_text = f" between {start_date} and {end_date}"
                elif start_date:
                    period_text = f" from {start_date}"
                elif end_date:
                    period_text = f" up to {end_date}"
                if query_raw:
                    self.status_var.set(f"Found {len(results)} items for '{query_raw}'{period_text}")
                else:
                    self.status_var.set(f"Found {len(results)} items{period_text}")
            else:
                self.status_var.set(f"Found {len(results)} items matching '{query_raw}'")
        else:
            self.status_var.set("No items found for current search filters")
    
    def get_selected_items(self, require_single=False):
        """Get currently selected inventory items"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warning", "Please select an item first")
            return []

        items = []
        for selection_id in selection:
            values = self.tree.item(selection_id).get('values') or []
            if not values:
                continue
            item = self.inventory_manager.get_item(values[0])
            if item:
                items.append(item)

        if require_single and len(items) != 1:
            messagebox.showwarning("Warning", "Please select exactly one item for this action")
            return []
        return items

    def get_selected_item(self):
        """Get a single selected item"""
        items = self.get_selected_items(require_single=True)
        return items[0] if items else None
    
    def add_item(self):
        """Open add item dialog"""
        dialog = AddItemDialog(self.parent, self.inventory_manager)
        self.parent.wait_window(dialog.dialog)
        self.refresh_all()
    
    def edit_item(self):
        """Edit selected item"""
        item = self.get_selected_item()
        if item:
            dialog = EditItemDialog(self.parent, self.inventory_manager, item.item_id)
            self.parent.wait_window(dialog.dialog)
            self.refresh_all()
    
    def manage_costs(self):
        """Manage additional costs for selected item"""
        item = self.get_selected_item()
        if item:
            dialog = ManageCostsDialog(self.parent, self.inventory_manager, item.item_id)
            self.parent.wait_window(dialog.dialog)
            self.refresh_all()
    
    def delete_item(self):
        """Delete selected items with confirmation"""
        items = self.get_selected_items()
        if not items:
            return

        if len(items) == 1:
            item_lines = [
                f"Item: {items[0].name}",
                f"Item ID: {items[0].item_id}",
                f"Quantity: {items[0].quantity}",
            ]
        else:
            item_lines = [f"Items selected: {len(items)}"] + [f"- {item.name} ({item.item_id})" for item in items[:10]]
            if len(items) > 10:
                item_lines.append(f"- ... and {len(items) - 10} more")

        confirm_message = "Are you sure you want to delete the selected item(s)?\n\n" + "\n".join(item_lines) + "\n\nThis action cannot be undone!"
        if not messagebox.askyesno("Confirm Delete", confirm_message):
            return

        failed = []
        for item in items:
            if not self.inventory_manager.delete_item(item.item_id):
                failed.append(item.name)

        if failed:
            messagebox.showerror("Error", "Failed to delete:\n" + "\n".join(failed))
        else:
            messagebox.showinfo("Success", f"Deleted {len(items)} item(s) successfully.")
        self.refresh_all()
    
    def sell_item(self):
        """Sell selected item"""
        selected_items = self.get_selected_items(require_single=False)
        selected_ids = [item.item_id for item in selected_items] if selected_items else []
        dialog = SellItemDialog(
            self.parent,
            self.inventory_manager,
            self.invoice_manager,
            preselected_item_ids=selected_ids
        )
        self.parent.wait_window(dialog.dialog)
        self.refresh_all()
    
    def show_reports(self):
        """Show supplier outbound report."""
        dlg = SupplierOutboundReportDialog(self.parent, self.inventory_manager)
        try:
            self.parent.wait_window(dlg.dialog)
        except Exception:
            pass

    def show_suppliers(self):
        """Open the suppliers dashboard."""
        dlg = SuppliersDialog(self.parent, self.inventory_manager, on_data_changed=self.refresh_all)
        try:
            self.parent.wait_window(dlg.dialog)
        except Exception:
            pass
    
    def show_inventory_invoices(self):
        """Show inventory-only invoices"""
        InventoryInvoicesDialog(self.parent, self.inventory_manager, self.invoice_manager, on_data_changed=self.refresh_all)

    def show_saved_sales(self):
        """Show saved sales records to allow permanent cleanup"""
        InventorySalesDialog(self.parent, self.inventory_manager, on_data_changed=self.refresh_all)
