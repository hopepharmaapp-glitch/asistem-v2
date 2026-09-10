import copy
import csv
import json
import os
import shutil
import tkinter as tk
import webbrowser
from datetime import datetime, timedelta
from tkinter import filedialog, messagebox, ttk

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


COMPANY_TYPES = [
    "Warehouse only",
    "Distribution only",
    "Warehouse + Distribution",
    "Services only",
    "Full ERP",
]

MODULE_LABELS = {
    "dashboard": "Dashboard",
    "invoice_management": "Invoice Management",
    "purchases": "Purchases",
    "inventory": "Inventory",
    "warehouse_management": "Warehouse Management",
    "clients": "Clients",
    "suppliers": "Suppliers",
    "accounting": "Accounting",
    "vat": "VAT",
    "reports": "Reports",
    "temperature_log": "Temperature Log",
    "ecommerce_marketplace": "E-Commerce Marketplace",
    "employee_manager": "Employee Manager",
    "regulatory_documents": "Regulatory Documents",
    "delivery_notes": "Delivery Notes",
    "stock_transfers": "Stock Transfers",
    "batch_tracking": "Batch Tracking",
    "expiry_tracking": "Expiry Tracking",
    "barcode_scanning": "Barcode Scanning",
    "multi_warehouse": "Multi-Warehouse",
    "user_permissions": "User Permissions",
    "audit_log": "Audit Log",
}

MODULE_KEYS = list(MODULE_LABELS.keys())

DASHBOARD_WIDGET_LABELS = {
    "todays_sales": "Today's Sales",
    "purchases_today": "Purchases Today",
    "outstanding_receivables": "Outstanding Receivables",
    "temperature_alerts_card": "Temperature Alerts Card",
    "invoices": "Invoices",
    "outstanding_payables": "Outstanding Payables",
    "low_stock_products": "Low Stock Products Card",
    "expiring_products": "Expiring Products Card",
    "inventory_value": "Inventory Value",
    "total_stock_items": "Total Stock Items",
    "expired_items": "Expired Items",
    "pending_approvals": "Pending Approvals",
    "quick_actions": "Quick Actions",
    "low_stock_table": "Low Stock Products Table",
    "expiring_products_table": "Expiring Products Table",
    "notifications_panel": "Notifications",
    "temperature_alerts_panel": "Temperature Alerts Panel",
    "recent_activities": "Recent Activities",
    "recent_goods_received": "Recent Goods Received",
    "recent_stock_transfers": "Recent Stock Transfers",
}

DASHBOARD_WIDGET_KEYS = list(DASHBOARD_WIDGET_LABELS.keys())

WAREHOUSE_TYPES = [
    "Main",
    "Branch",
    "Cold Room",
    "Quarantine",
    "Returns",
    "Damaged",
    "Expired",
]

LOCATION_STATUSES = ["Active", "Blocked", "Full", "Quarantine", "Inactive"]
OUTBOUND_STATUSES = [
    "Draft",
    "Pending Approval",
    "Approved",
    "Reserved",
    "Picking",
    "Picked",
    "Packing",
    "Packed",
    "Ready for Dispatch",
    "Dispatched",
    "Delivered",
    "Returned",
    "Cancelled",
]

APPROVAL_TYPES = {
    "transfer": "Stock Transfer",
    "adjustment": "Stock Adjustment",
    "return": "Return",
    "stock_count": "Stock Count",
}


def default_enabled_modules(company_type="Full ERP"):
    selected = str(company_type or "Full ERP").strip()
    presets = {
        "Warehouse only": {
            "dashboard", "inventory", "warehouse_management", "suppliers", "reports",
            "temperature_log", "stock_transfers", "batch_tracking", "expiry_tracking",
            "audit_log",
        },
        "Distribution only": {
            "dashboard", "invoice_management", "purchases", "inventory", "clients",
            "suppliers", "reports", "delivery_notes", "batch_tracking",
            "expiry_tracking", "audit_log",
        },
        "Warehouse + Distribution": {
            "dashboard", "invoice_management", "purchases", "inventory",
            "warehouse_management", "clients", "suppliers", "reports",
            "temperature_log", "delivery_notes", "stock_transfers",
            "batch_tracking", "expiry_tracking", "audit_log",
        },
        "Services only": {
            "dashboard", "invoice_management", "clients", "accounting", "vat",
            "reports", "employee_manager", "audit_log",
        },
        "Full ERP": set(MODULE_KEYS),
    }
    enabled = presets.get(selected, presets["Full ERP"])
    return {key: key in enabled for key in MODULE_KEYS}


def default_company_settings():
    return {
        "default_currency": "AED",
        "vat_enabled": True,
        "vat_rate": 5.0,
        "invoice_prefix": "INV",
        "purchase_prefix": "PUR",
        "warehouse_prefix": "WH",
        "low_stock_alert_threshold": 10,
        "expiry_alert_days": 30,
        "temperature_monitoring_enabled": True,
        "batch_tracking_required": True,
        "barcode_required": False,
        "multi_warehouse_enabled": False,
        "approval_workflow_enabled": True,
        "role_based_access_enabled": True,
        "dashboard_widgets": {key: True for key in DASHBOARD_WIDGET_KEYS},
    }


def default_company_configuration(company_type="Full ERP"):
    settings = default_company_settings()
    if str(company_type).strip() == "Services only":
        settings["temperature_monitoring_enabled"] = False
        settings["batch_tracking_required"] = False
        settings["approval_workflow_enabled"] = False
    if str(company_type).strip() == "Warehouse + Distribution":
        settings["multi_warehouse_enabled"] = True
    return {
        "company_type": company_type,
        "enabled_modules": default_enabled_modules(company_type),
        "settings": settings,
    }


def normalize_company_record(record):
    base = copy.deepcopy(record or {})
    company_type = str(base.get("company_type") or "Full ERP").strip() or "Full ERP"
    module_cfg = base.get("enabled_modules")
    if not isinstance(module_cfg, dict):
        module_cfg = default_enabled_modules(company_type)
    else:
        normalized = default_enabled_modules(company_type)
        for key in MODULE_KEYS:
            if key in module_cfg:
                normalized[key] = bool(module_cfg.get(key))
        module_cfg = normalized
    base["company_type"] = company_type
    base["enabled_modules"] = module_cfg
    return base


def _read_json(path, default_value):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as handle:
                return json.load(handle)
    except Exception:
        pass
    return copy.deepcopy(default_value)


def _write_json(path, data):
    folder = os.path.dirname(path)
    if folder:
        os.makedirs(folder, exist_ok=True)
    with open(path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, indent=2, ensure_ascii=False)


def _safe_text(value):
    return str(value or "").strip()


def _safe_float(value, default=0.0):
    try:
        return float(value or 0.0)
    except Exception:
        return float(default)


def _safe_int(value, default=0):
    try:
        return int(float(value or 0))
    except Exception:
        return int(default)


def _fit_dialog(win, min_w=640, min_h=420):
    try:
        win.update_idletasks()
        screen_w = win.winfo_screenwidth()
        screen_h = win.winfo_screenheight()
        req_w = max(min_w, win.winfo_reqwidth() + 24)
        req_h = max(min_h, win.winfo_reqheight() + 24)
        width = min(req_w, int(screen_w * 0.92))
        height = min(req_h, int(screen_h * 0.9))
        width = max(min_w, width)
        height = max(min_h, height)
        x = max(0, int((screen_w - width) / 2))
        y = max(0, int((screen_h - height) / 2))
        win.geometry(f"{width}x{height}+{x}+{y}")
        win.minsize(min_w, min_h)
    except Exception:
        pass


class ScrollableSettingsFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        try:
            style = ttk.Style()
            canvas_bg = style.lookup("App.TFrame", "background") or style.lookup("TFrame", "background") or "#f8fafc"
        except Exception:
            canvas_bg = "#f8fafc"
        self.canvas = tk.Canvas(self, highlightthickness=0, borderwidth=0, background=canvas_bg)
        self.v_scroll = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scroll.set)
        self.scrollable_frame = ttk.Frame(self.canvas)
        self._canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")

        self.canvas.pack(side="left", fill="both", expand=True)
        self.v_scroll.pack(side="right", fill="y")

        self.scrollable_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        self.scrollable_frame.bind("<Enter>", self._bind_mousewheel)
        self.scrollable_frame.bind("<Leave>", self._unbind_mousewheel)

    def _on_frame_configure(self, _event=None):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        try:
            self.canvas.itemconfigure(self._canvas_window, width=event.width)
        except Exception:
            pass

    def _on_mousewheel(self, event):
        try:
            if getattr(event, "delta", 0):
                self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            elif getattr(event, "num", None) == 4:
                self.canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                self.canvas.yview_scroll(1, "units")
        except Exception:
            pass

    def _bind_mousewheel(self, _event=None):
        try:
            self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
            self.canvas.bind_all("<Button-4>", self._on_mousewheel)
            self.canvas.bind_all("<Button-5>", self._on_mousewheel)
        except Exception:
            pass

    def _unbind_mousewheel(self, _event=None):
        try:
            self.canvas.unbind_all("<MouseWheel>")
            self.canvas.unbind_all("<Button-4>")
            self.canvas.unbind_all("<Button-5>")
        except Exception:
            pass


def ensure_company_extension_storage(company_folder, company_name=""):
    folder = str(company_folder or os.getcwd())
    os.makedirs(folder, exist_ok=True)
    app_settings_path = os.path.join(folder, "app_settings.json")
    app_settings = _read_json(app_settings_path, {})
    migration_key = "warehouse_extension_v1"

    backups_dir = os.path.join(folder, "backups")
    os.makedirs(backups_dir, exist_ok=True)
    if not app_settings.get(migration_key):
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        snapshot_dir = os.path.join(backups_dir, f"warehouse_extension_pre_migration_{stamp}")
        os.makedirs(snapshot_dir, exist_ok=True)
        for filename in ("app_settings.json", "inventory.json", "purchases_data.json", "companies.json"):
            src = os.path.join(folder, filename)
            if os.path.exists(src):
                try:
                    shutil.copy2(src, os.path.join(snapshot_dir, filename))
                except Exception:
                    pass

    company_cfg = copy.deepcopy(app_settings.get("company_configuration") or {})
    company_type = _safe_text(company_cfg.get("company_type") or "Full ERP") or "Full ERP"
    default_cfg = default_company_configuration(company_type)
    merged_modules = default_cfg["enabled_modules"]
    merged_modules.update(company_cfg.get("enabled_modules") or {})
    merged_settings = default_cfg["settings"]
    merged_settings.update(company_cfg.get("settings") or {})
    dashboard_widgets = {key: True for key in DASHBOARD_WIDGET_KEYS}
    dashboard_widgets.update((merged_settings.get("dashboard_widgets") or {}))
    merged_settings["dashboard_widgets"] = {key: bool(dashboard_widgets.get(key, True)) for key in DASHBOARD_WIDGET_KEYS}
    app_settings["company_configuration"] = {
        "company_type": company_type,
        "enabled_modules": {key: bool(merged_modules.get(key, False)) for key in MODULE_KEYS},
        "settings": merged_settings,
    }
    app_settings[migration_key] = True
    _write_json(app_settings_path, app_settings)

    default_warehouse_code = f"{merged_settings.get('warehouse_prefix', 'WH')}-001"
    data_files = {
        "warehouses.json": [{
            "warehouse_code": default_warehouse_code,
            "warehouse_name": "Main Warehouse",
            "address": "",
            "emirate_country": "UAE",
            "contact_person": "",
            "phone": "",
            "email": "",
            "status": "Active",
            "temperature_controlled": True,
            "warehouse_type": "Main",
            "archived": False,
        }],
        "warehouse_locations.json": [],
        "goods_receipts.json": [],
        "stock_movements.json": [],
        "outbounds.json": [],
        "stock_transfers.json": [],
        "stock_adjustments.json": [],
        "stock_returns.json": [],
        "stock_counts.json": [],
        "warehouse_approvals.json": [],
    }
    for filename, default_data in data_files.items():
        path = os.path.join(folder, filename)
        if not os.path.exists(path):
            _write_json(path, default_data)
    audit_path = os.path.join(folder, "audit_log.jsonl")
    if not os.path.exists(audit_path):
        with open(audit_path, "a", encoding="utf-8"):
            pass
    return app_settings["company_configuration"]


def load_company_configuration(company_folder):
    settings_path = os.path.join(str(company_folder or os.getcwd()), "app_settings.json")
    app_settings = _read_json(settings_path, {})
    cfg = app_settings.get("company_configuration") or {}
    company_type = _safe_text(cfg.get("company_type") or "Full ERP") or "Full ERP"
    merged = default_company_configuration(company_type)
    merged["enabled_modules"].update(cfg.get("enabled_modules") or {})
    merged["settings"].update(cfg.get("settings") or {})
    dashboard_widgets = {key: True for key in DASHBOARD_WIDGET_KEYS}
    dashboard_widgets.update((merged["settings"].get("dashboard_widgets") or {}))
    merged["settings"]["dashboard_widgets"] = {key: bool(dashboard_widgets.get(key, True)) for key in DASHBOARD_WIDGET_KEYS}
    return merged


def save_company_configuration(company_folder, company_type, enabled_modules, settings_updates):
    folder = str(company_folder or os.getcwd())
    ensure_company_extension_storage(folder)
    settings_path = os.path.join(folder, "app_settings.json")
    app_settings = _read_json(settings_path, {})
    merged = default_company_configuration(company_type)
    merged["enabled_modules"].update(enabled_modules or {})
    merged["settings"].update(settings_updates or {})
    dashboard_widgets = {key: True for key in DASHBOARD_WIDGET_KEYS}
    dashboard_widgets.update((merged["settings"].get("dashboard_widgets") or {}))
    merged["settings"]["dashboard_widgets"] = {key: bool(dashboard_widgets.get(key, True)) for key in DASHBOARD_WIDGET_KEYS}
    merged["enabled_modules"] = {key: bool(merged["enabled_modules"].get(key, False)) for key in MODULE_KEYS}
    app_settings["company_configuration"] = merged
    _write_json(settings_path, app_settings)
    return merged


class WarehouseManager:
    def __init__(self, company_folder, current_company=""):
        self.company_folder = str(company_folder or os.getcwd())
        self.current_company = _safe_text(current_company)
        self.paths = {
            "warehouses": os.path.join(self.company_folder, "warehouses.json"),
            "locations": os.path.join(self.company_folder, "warehouse_locations.json"),
            "receipts": os.path.join(self.company_folder, "goods_receipts.json"),
            "movements": os.path.join(self.company_folder, "stock_movements.json"),
            "outbounds": os.path.join(self.company_folder, "outbounds.json"),
            "transfers": os.path.join(self.company_folder, "stock_transfers.json"),
            "adjustments": os.path.join(self.company_folder, "stock_adjustments.json"),
            "returns": os.path.join(self.company_folder, "stock_returns.json"),
            "counts": os.path.join(self.company_folder, "stock_counts.json"),
            "approvals": os.path.join(self.company_folder, "warehouse_approvals.json"),
            "settings": os.path.join(self.company_folder, "app_settings.json"),
            "audit": os.path.join(self.company_folder, "audit_log.jsonl"),
        }
        self.configuration = ensure_company_extension_storage(self.company_folder, self.current_company)

    def _load_list(self, key):
        return _read_json(self.paths[key], [])

    def _save_list(self, key, rows):
        _write_json(self.paths[key], rows)
        return True

    def get_configuration(self):
        self.configuration = load_company_configuration(self.company_folder)
        return copy.deepcopy(self.configuration)

    def update_configuration(self, company_type, enabled_modules, settings_updates):
        self.configuration = save_company_configuration(self.company_folder, company_type, enabled_modules, settings_updates)
        return copy.deepcopy(self.configuration)

    def _generate_ref(self, prefix, rows, key_name):
        prefix = _safe_text(prefix) or "REF"
        next_num = len(rows) + 1
        return f"{prefix}-{datetime.now().strftime('%Y%m%d')}-{next_num:04d}"

    def _append_audit(self, action, module_name, reference_id, old_value=None, new_value=None, user_name="System"):
        entry = {
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "user": _safe_text(user_name) or "System",
            "action": action,
            "old_value": old_value or {},
            "new_value": new_value or {},
            "company": self.current_company,
            "module": module_name,
            "reference_id": reference_id,
        }
        with open(self.paths["audit"], "a", encoding="utf-8") as handle:
            handle.write(json.dumps(entry, ensure_ascii=False) + "\n")

    def clear_audit_log(self):
        with open(self.paths["audit"], "w", encoding="utf-8") as handle:
            handle.write("")

    def _transfer_snapshot_from_item(self, item):
        if not item:
            return {}
        return {
            "item_description": _safe_text(getattr(item, "description", "")),
            "sku": _safe_text(getattr(item, "sku", "")),
            "brand": _safe_text(getattr(item, "brand", "")),
            "manufacture_date": _safe_text(getattr(item, "manufacture_date", "")),
            "expiry_date": _safe_text(getattr(item, "expiry_date", "")),
            "lot_number": _safe_text(getattr(item, "lot_number", "")),
            "unit_of_measure": _safe_text(getattr(item, "unit_of_measure", "")),
            "pack_size": _safe_text(getattr(item, "pack_size", "")),
            "stored_for_client": _safe_text(getattr(item, "client_name", "")),
            "source_supplier": _safe_text(getattr(item, "supplier_name", "") or getattr(item, "supplier", "")),
        }

    def list_warehouses(self, include_archived=False):
        rows = self._load_list("warehouses")
        if include_archived:
            return rows
        return [row for row in rows if not bool(row.get("archived"))]

    def save_warehouse(self, payload, user_name="System"):
        rows = self._load_list("warehouses")
        warehouse_code = _safe_text(payload.get("warehouse_code"))
        if not warehouse_code:
            prefix = self.get_configuration()["settings"].get("warehouse_prefix", "WH")
            warehouse_code = f"{prefix}-{len(rows) + 1:03d}"
            payload["warehouse_code"] = warehouse_code
        existing = next((row for row in rows if _safe_text(row.get("warehouse_code")) == warehouse_code), None)
        record = {
            "warehouse_code": warehouse_code,
            "warehouse_name": _safe_text(payload.get("warehouse_name")),
            "address": _safe_text(payload.get("address")),
            "emirate_country": _safe_text(payload.get("emirate_country")),
            "contact_person": _safe_text(payload.get("contact_person")),
            "phone": _safe_text(payload.get("phone")),
            "email": _safe_text(payload.get("email")),
            "status": _safe_text(payload.get("status") or "Active"),
            "temperature_controlled": bool(payload.get("temperature_controlled")),
            "warehouse_type": _safe_text(payload.get("warehouse_type") or "Main"),
            "archived": bool(payload.get("archived")),
        }
        if not record["warehouse_name"]:
            raise ValueError("Warehouse name is required")
        old_value = copy.deepcopy(existing) if existing else {}
        if existing:
            existing.update(record)
        else:
            rows.append(record)
        self._save_list("warehouses", rows)
        self._append_audit("warehouse_saved", "warehouse_management", warehouse_code, old_value, record, user_name)
        return record

    def archive_warehouse(self, warehouse_code, user_name="System"):
        rows = self._load_list("warehouses")
        for row in rows:
            if _safe_text(row.get("warehouse_code")) == _safe_text(warehouse_code):
                old_value = copy.deepcopy(row)
                row["archived"] = True
                row["status"] = "Archived"
                self._save_list("warehouses", rows)
                self._append_audit("warehouse_archived", "warehouse_management", warehouse_code, old_value, row, user_name)
                return True
        return False

    def delete_warehouse(self, warehouse_code, user_name="System"):
        warehouse_code = _safe_text(warehouse_code)
        rows = self._load_list("warehouses")
        row = next((x for x in rows if _safe_text(x.get("warehouse_code")) == warehouse_code), None)
        if not row:
            raise ValueError("Warehouse not found")
        for location in self.list_locations():
            if _safe_text(location.get("warehouse_code")) == warehouse_code:
                raise ValueError("Cannot delete warehouse because it still has locations assigned")
        for receipt in self.list_goods_receipts():
            if _safe_text(receipt.get("warehouse_code")) == warehouse_code:
                raise ValueError("Cannot delete warehouse because goods receipts are linked to it")
        for transfer in self.list_transfers():
            if _safe_text(transfer.get("warehouse_from")) == warehouse_code or _safe_text(transfer.get("warehouse_to")) == warehouse_code:
                raise ValueError("Cannot delete warehouse because transfers are linked to it")
        for adjustment in self.list_adjustments():
            if _safe_text(adjustment.get("warehouse_code")) == warehouse_code:
                raise ValueError("Cannot delete warehouse because adjustments are linked to it")
        for count in self.list_counts():
            if _safe_text(count.get("warehouse_code")) == warehouse_code:
                raise ValueError("Cannot delete warehouse because stock counts are linked to it")
        invm = self._get_inventory_manager()
        for item in invm.items:
            if _safe_text(getattr(item, "warehouse_code", "")) == warehouse_code:
                raise ValueError("Cannot delete warehouse because inventory is stored in it")
        old_value = copy.deepcopy(row)
        rows = [x for x in rows if _safe_text(x.get("warehouse_code")) != warehouse_code]
        self._save_list("warehouses", rows)
        self._append_audit("warehouse_deleted", "warehouse_management", warehouse_code, old_value, {}, user_name)
        return True

    def list_locations(self):
        return self._load_list("locations")

    def _location_owner_name(self, record):
        return _safe_text((record or {}).get("owner_company") or (record or {}).get("client_name"))

    def _location_matches_hint(self, record, hint):
        probe = _safe_text(hint).lower()
        if not probe:
            return True
        values = [
            record.get("location_code"),
            record.get("bin"),
            record.get("shelf"),
            f"{_safe_text(record.get('rack'))}/{_safe_text(record.get('shelf'))}",
            f"{_safe_text(record.get('rack'))}/{_safe_text(record.get('shelf'))}/{_safe_text(record.get('bin'))}",
            f"{_safe_text(record.get('zone'))}/{_safe_text(record.get('aisle'))}/{_safe_text(record.get('rack'))}/{_safe_text(record.get('shelf'))}/{_safe_text(record.get('bin'))}",
        ]
        for value in values:
            text = _safe_text(value).lower()
            if text and text == probe:
                return True
        return False

    def _resolve_shelf_location(self, warehouse_code, client_name="", location_hint=""):
        warehouse = _safe_text(warehouse_code).lower()
        hint = _safe_text(location_hint)
        client = _safe_text(client_name)
        rows = [
            row for row in self.list_locations()
            if _safe_text(row.get("warehouse_code")).lower() == warehouse
            and _safe_text(row.get("status") or "Active").lower() != "inactive"
        ]
        if not rows:
            return None
        if client:
            owned = [
                row for row in rows
                if self._location_owner_name(row).lower() == client.lower()
            ]
            if hint:
                owned = [row for row in owned if self._location_matches_hint(row, hint)]
            if len(owned) == 1:
                return owned[0]
            if len(owned) > 1:
                raise ValueError(
                    f"Multiple shelf locations are assigned to {client}. Please enter the shelf, bin, or location code."
                )
            warehouse_has_owned_shelves = any(self._location_owner_name(row) for row in rows)
            if warehouse_has_owned_shelves:
                raise ValueError(
                    f"No shelf is assigned to {client} in warehouse {warehouse_code}."
                )
        if hint:
            hinted = [row for row in rows if self._location_matches_hint(row, hint)]
            if len(hinted) == 1:
                return hinted[0]
            if len(hinted) > 1:
                raise ValueError("Location hint matches more than one shelf. Please use the exact location code.")
        return None

    def save_location(self, payload, user_name="System"):
        rows = self._load_list("locations")
        location_code = _safe_text(payload.get("location_code"))
        if not location_code:
            location_code = f"LOC-{len(rows) + 1:04d}"
        record = {
            "location_code": location_code,
            "warehouse_code": _safe_text(payload.get("warehouse_code")),
            "zone": _safe_text(payload.get("zone")),
            "aisle": _safe_text(payload.get("aisle")),
            "rack": _safe_text(payload.get("rack")),
            "shelf": _safe_text(payload.get("shelf")),
            "bin": _safe_text(payload.get("bin")),
            "owner_company": _safe_text(payload.get("owner_company") or payload.get("client_name")),
            "capacity": _safe_float(payload.get("capacity"), 0.0),
            "status": _safe_text(payload.get("status") or "Active"),
            "product_code": _safe_text(payload.get("product_code")),
        }
        if not record["warehouse_code"]:
            raise ValueError("Warehouse is required")
        for row in rows:
            row_code = _safe_text(row.get("location_code"))
            if row_code == location_code:
                continue
            same_spot = (
                _safe_text(row.get("warehouse_code")).lower() == record["warehouse_code"].lower()
                and _safe_text(row.get("zone")).lower() == record["zone"].lower()
                and _safe_text(row.get("aisle")).lower() == record["aisle"].lower()
                and _safe_text(row.get("rack")).lower() == record["rack"].lower()
                and _safe_text(row.get("shelf")).lower() == record["shelf"].lower()
                and _safe_text(row.get("bin")).lower() == record["bin"].lower()
            )
            if same_spot:
                raise ValueError("This exact warehouse/rack/shelf/bin location already exists.")
            same_shelf = (
                _safe_text(row.get("warehouse_code")).lower() == record["warehouse_code"].lower()
                and _safe_text(row.get("zone")).lower() == record["zone"].lower()
                and _safe_text(row.get("aisle")).lower() == record["aisle"].lower()
                and _safe_text(row.get("rack")).lower() == record["rack"].lower()
                and _safe_text(row.get("shelf")).lower() == record["shelf"].lower()
            )
            other_owner = self._location_owner_name(row)
            if same_shelf and record["owner_company"] and other_owner and other_owner.lower() != record["owner_company"].lower():
                raise ValueError(
                    f"Shelf {record['shelf']} on rack {record['rack']} is already assigned to {other_owner}."
                )
        old_value = {}
        updated = False
        for row in rows:
            if _safe_text(row.get("location_code")) == location_code:
                old_value = copy.deepcopy(row)
                row.update(record)
                updated = True
                break
        if not updated:
            rows.append(record)
        self._save_list("locations", rows)
        self._append_audit("location_saved", "warehouse_management", location_code, old_value, record, user_name)
        return record

    def delete_location(self, location_code, user_name="System"):
        location_code = _safe_text(location_code)
        rows = self._load_list("locations")
        row = next((x for x in rows if _safe_text(x.get("location_code")) == location_code), None)
        if not row:
            raise ValueError("Location not found")
        invm = self._get_inventory_manager()
        for item in invm.items:
            if _safe_text(getattr(item, "bin_location", "")) == location_code:
                raise ValueError("Cannot delete location because inventory is stored in it")
        for receipt in self.list_goods_receipts():
            if _safe_text(receipt.get("location_code")) == location_code or _safe_text(receipt.get("bin_location")) == location_code:
                raise ValueError("Cannot delete location because goods receipts are linked to it")
        for transfer in self.list_transfers():
            if _safe_text(transfer.get("bin_from")) == location_code or _safe_text(transfer.get("bin_to")) == location_code:
                raise ValueError("Cannot delete location because transfers are linked to it")
        for adjustment in self.list_adjustments():
            if _safe_text(adjustment.get("bin_location")) == location_code:
                raise ValueError("Cannot delete location because adjustments are linked to it")
        for count in self.list_counts():
            if _safe_text(count.get("bin_location")) == location_code:
                raise ValueError("Cannot delete location because stock counts are linked to it")
        old_value = copy.deepcopy(row)
        rows = [x for x in rows if _safe_text(x.get("location_code")) != location_code]
        self._save_list("locations", rows)
        self._append_audit("location_deleted", "warehouse_management", location_code, old_value, {}, user_name)
        return True

    def list_goods_receipts(self):
        return self._load_list("receipts")

    def list_transfers(self):
        return self._load_list("transfers")

    def list_adjustments(self):
        return self._load_list("adjustments")

    def list_returns(self):
        return self._load_list("returns")

    def list_counts(self):
        return self._load_list("counts")

    def list_movements(self):
        return self._load_list("movements")

    def get_goods_receipt(self, grn_id):
        return next((row for row in self.list_goods_receipts() if _safe_text(row.get("grn_id")) == _safe_text(grn_id)), None)

    def get_transfer(self, transfer_id):
        return next((row for row in self.list_transfers() if _safe_text(row.get("transfer_id")) == _safe_text(transfer_id)), None)

    def get_adjustment(self, adjustment_id):
        return next((row for row in self.list_adjustments() if _safe_text(row.get("adjustment_id")) == _safe_text(adjustment_id)), None)

    def get_stock_count(self, count_id):
        return next((row for row in self.list_counts() if _safe_text(row.get("count_id")) == _safe_text(count_id)), None)

    def list_outbounds(self):
        return self._load_list("outbounds")

    def get_outbound(self, outbound_no):
        return next((row for row in self.list_outbounds() if _safe_text(row.get("outbound_no")) == _safe_text(outbound_no)), None)

    def _parse_iso_date(self, value):
        text = _safe_text(value)
        if not text:
            return None
        try:
            return datetime.strptime(text, "%Y-%m-%d").date()
        except Exception:
            return None

    def _expiry_sort_key(self, item):
        expiry = self._parse_iso_date(getattr(item, "expiry_date", ""))
        return (expiry is None, expiry or datetime.max.date(), _safe_text(getattr(item, "batch_number", "")), _safe_text(getattr(item, "item_id", "")))

    def _item_matches_query(self, item, product_name="", sku="", item_id=""):
        product_probe = _safe_text(product_name).lower()
        sku_probe = _safe_text(sku).lower()
        item_probe = _safe_text(item_id).lower()
        if item_probe and _safe_text(getattr(item, "item_id", "")).lower() != item_probe:
            return False
        if product_probe:
            haystack = " ".join([
                _safe_text(getattr(item, "name", "")),
                _safe_text(getattr(item, "description", "")),
                _safe_text(getattr(item, "brand", "")),
            ]).lower()
            if product_probe not in haystack:
                return False
        if sku_probe:
            sku_haystack = " ".join([
                _safe_text(getattr(item, "sku", "")),
                _safe_text(getattr(item, "item_id", "")),
            ]).lower()
            if sku_probe not in sku_haystack:
                return False
        return True

    def _fefo_inventory_candidates(
        self,
        client_owner="",
        warehouse_code="",
        bin_location="",
        product_name="",
        sku="",
        item_id="",
        batch_number="",
        lot_number="",
        include_expired=False,
    ):
        invm = self._get_inventory_manager()
        client_probe = _safe_text(client_owner).lower()
        wh_probe = _safe_text(warehouse_code).lower()
        bin_probe = _safe_text(bin_location).lower()
        batch_probe = _safe_text(batch_number).lower()
        lot_probe = _safe_text(lot_number).lower()
        rows = []
        for item in invm.items:
            if client_probe and _safe_text(getattr(item, "client_name", "")).lower() != client_probe:
                continue
            if wh_probe and _safe_text(getattr(item, "warehouse_code", "")).lower() != wh_probe:
                continue
            if bin_probe and _safe_text(getattr(item, "bin_location", "")).lower() != bin_probe:
                continue
            if batch_probe and _safe_text(getattr(item, "batch_number", "")).lower() != batch_probe:
                continue
            if lot_probe and _safe_text(getattr(item, "lot_number", "")).lower() != lot_probe:
                continue
            if not self._item_matches_query(item, product_name=product_name, sku=sku, item_id=item_id):
                continue
            if bool(getattr(item, "quarantine", False)):
                continue
            available_qty = _safe_float(item.get_available_quantity() if hasattr(item, "get_available_quantity") else getattr(item, "quantity", 0), 0.0)
            if available_qty <= 0:
                continue
            expired = bool(item.is_expired()) if hasattr(item, "is_expired") else False
            if expired and not include_expired:
                continue
            rows.append(item)
        rows.sort(key=self._expiry_sort_key)
        return rows

    def _build_outbound_item_line(self, item, requested_qty):
        requested_qty = _safe_float(requested_qty, 0.0)
        available_qty = _safe_float(item.get_available_quantity() if hasattr(item, "get_available_quantity") else getattr(item, "quantity", 0), 0.0)
        cost_per_item = item.get_cost_per_item() if hasattr(item, "get_cost_per_item") else 0.0
        return {
            "item_id": _safe_text(getattr(item, "item_id", "")),
            "product_name": _safe_text(getattr(item, "name", "")),
            "brand": _safe_text(getattr(item, "brand", "")),
            "sku": _safe_text(getattr(item, "sku", "")),
            "batch_number": _safe_text(getattr(item, "batch_number", "")),
            "lot_number": _safe_text(getattr(item, "lot_number", "")),
            "manufacture_date": _safe_text(getattr(item, "manufacture_date", "")),
            "expiry_date": _safe_text(getattr(item, "expiry_date", "")),
            "unit_pack_size": " / ".join([part for part in [_safe_text(getattr(item, "unit_of_measure", "")), _safe_text(getattr(item, "pack_size", ""))] if part]),
            "unit_of_measure": _safe_text(getattr(item, "unit_of_measure", "")),
            "pack_size": _safe_text(getattr(item, "pack_size", "")),
            "warehouse_code": _safe_text(getattr(item, "warehouse_code", "")),
            "bin_location": _safe_text(getattr(item, "bin_location", "")),
            "available_quantity": available_qty,
            "requested_quantity": requested_qty,
            "reserved_quantity": 0.0,
            "picked_quantity": 0.0,
            "packed_quantity": 0.0,
            "dispatched_quantity": 0.0,
            "returned_quantity": 0.0,
            "status": "Draft",
            "item_description": _safe_text(getattr(item, "description", "")),
            "client_owner": _safe_text(getattr(item, "client_name", "")),
            "supplier_name": _safe_text(getattr(item, "supplier_name", "") or getattr(item, "supplier", "")),
            "temperature_controlled": bool(getattr(item, "temperature_controlled", False)),
            "source_item_id": _safe_text(getattr(item, "item_id", "")),
            "unit_cost_snapshot": _safe_float(cost_per_item, 0.0),
        }

    def _refresh_outbound_line_stock(self, line, item):
        refreshed = dict(line or {})
        if item:
            refreshed["available_quantity"] = _safe_float(item.get_available_quantity() if hasattr(item, "get_available_quantity") else getattr(item, "quantity", 0), 0.0)
            refreshed["warehouse_code"] = _safe_text(getattr(item, "warehouse_code", refreshed.get("warehouse_code")))
            refreshed["bin_location"] = _safe_text(getattr(item, "bin_location", refreshed.get("bin_location")))
            refreshed["status"] = _safe_text(refreshed.get("status") or "Draft")
        return refreshed

    def _find_outbound_source_item(self, invm, line):
        return invm.get_item(_safe_text((line or {}).get("source_item_id") or (line or {}).get("item_id")))

    def _outbound_status_allowed(self, current_status, allowed_statuses):
        return _safe_text(current_status) in {str(x) for x in (allowed_statuses or [])}

    def _get_outbound_ref_prefix(self):
        return "OUT"

    def _get_inventory_manager(self):
        from inventory_system import InventoryManager
        return InventoryManager(self.company_folder)

    def _find_inventory_item(self, invm, product_name, batch_number="", warehouse_code="", bin_location=""):
        pname = _safe_text(product_name).lower()
        batch = _safe_text(batch_number).lower()
        wh = _safe_text(warehouse_code).lower()
        bin_code = _safe_text(bin_location).lower()
        for item in invm.items:
            if _safe_text(getattr(item, "name", "")).lower() != pname:
                continue
            if batch and _safe_text(getattr(item, "batch_number", "")).lower() != batch:
                continue
            if wh and _safe_text(getattr(item, "warehouse_code", "")).lower() != wh:
                continue
            if bin_code and _safe_text(getattr(item, "bin_location", "")).lower() != bin_code:
                continue
            return item
        return None

    def _log_movement(self, movement_type, payload, user_name="System"):
        rows = self._load_list("movements")
        movement_id = self._generate_ref("MOV", rows, "movement_id")
        record = {
            "movement_id": movement_id,
            "movement_type": movement_type,
            "reference_id": _safe_text(payload.get("reference_id")),
            "product_name": _safe_text(payload.get("product_name")),
            "item_id": _safe_text(payload.get("item_id")),
            "batch_number": _safe_text(payload.get("batch_number")),
            "quantity": _safe_float(payload.get("quantity"), 0.0),
            "warehouse_from": _safe_text(payload.get("warehouse_from")),
            "warehouse_to": _safe_text(payload.get("warehouse_to")),
            "bin_from": _safe_text(payload.get("bin_from")),
            "bin_to": _safe_text(payload.get("bin_to")),
            "unit_cost": _safe_float(payload.get("unit_cost"), 0.0),
            "movement_date": _safe_text(payload.get("movement_date") or datetime.now().strftime("%Y-%m-%d")),
            "user": _safe_text(user_name) or "System",
            "notes": _safe_text(payload.get("notes")),
        }
        rows.append(record)
        self._save_list("movements", rows)
        return record

    def _delete_movement_records(self, reference_id, movement_type=None):
        rows = self._load_list("movements")
        kept = []
        removed = []
        ref = _safe_text(reference_id)
        move_type = _safe_text(movement_type)
        for row in rows:
            same_ref = _safe_text(row.get("reference_id")) == ref
            same_type = True if not move_type else _safe_text(row.get("movement_type")) == move_type
            if same_ref and same_type:
                removed.append(row)
            else:
                kept.append(row)
        if len(kept) != len(rows):
            self._save_list("movements", kept)
        return removed

    def _delete_or_update_item_quantity(self, invm, item, new_qty, cost_per_item):
        new_qty = _safe_float(new_qty, 0.0)
        if new_qty <= 0:
            invm.delete_item(item.item_id)
            return None
        new_total_cost = max(0.0, new_qty * _safe_float(cost_per_item, 0.0))
        invm.update_item(item.item_id, quantity=new_qty, total_cost=new_total_cost)
        return invm.get_item(item.item_id)

    def receive_goods(self, payload, user_name="System"):
        invm = self._get_inventory_manager()
        receipts = self._load_list("receipts")
        grn_prefix = self.get_configuration()["settings"].get("purchase_prefix", "GRN")
        grn_id = self._generate_ref(grn_prefix, receipts, "grn_id")
        qty = _safe_float(payload.get("quantity_received"), 0.0)
        unit_cost = _safe_float(payload.get("unit_cost"), 0.0)
        product_name = _safe_text(payload.get("product_name"))
        warehouse_code = _safe_text(payload.get("warehouse_code"))
        client_name = _safe_text(
            payload.get("client_name") or payload.get("owner_company") or payload.get("company_name")
        )
        bin_location = _safe_text(payload.get("bin_location"))
        if not product_name or qty <= 0:
            raise ValueError("Product and quantity are required")
        location_record = self._resolve_shelf_location(warehouse_code, client_name=client_name, location_hint=bin_location)
        if location_record:
            if not client_name:
                client_name = self._location_owner_name(location_record)
            bin_location = _safe_text(
                location_record.get("location_code") or location_record.get("bin") or bin_location
            )
        existing = self._find_inventory_item(
            invm,
            product_name=product_name,
            batch_number=payload.get("batch_number"),
            warehouse_code=warehouse_code,
            bin_location=bin_location,
        )
        old_item = existing.to_dict() if existing else {}
        total_cost = qty * unit_cost
        if existing:
            new_qty = _safe_float(getattr(existing, "quantity", 0), 0.0) + qty
            new_total_cost = _safe_float(getattr(existing, "total_cost", 0), 0.0) + total_cost
            invm.update_item(
                existing.item_id,
                quantity=new_qty,
                total_cost=new_total_cost,
                supplier=_safe_text(payload.get("supplier")) or getattr(existing, "supplier", ""),
                expiry_date=_safe_text(payload.get("expiry_date")) or getattr(existing, "expiry_date", ""),
                manufacture_date=_safe_text(payload.get("manufacture_date")) or getattr(existing, "manufacture_date", ""),
                warehouse_code=warehouse_code or getattr(existing, "warehouse_code", ""),
                bin_location=bin_location or getattr(existing, "bin_location", ""),
                lot_number=_safe_text(payload.get("lot_number")) or getattr(existing, "lot_number", ""),
                barcode=_safe_text(payload.get("barcode")) or getattr(existing, "barcode", ""),
                client_name=client_name or getattr(existing, "client_name", ""),
                status="Available",
            )
            item_record = invm.get_item(existing.item_id)
        else:
            item_record = invm.add_item(
                name=product_name,
                category=_safe_text(payload.get("category")) or "Pharmaceutical",
                description=_safe_text(payload.get("notes")) or product_name,
                total_cost=total_cost,
                selling_price=_safe_float(payload.get("selling_price"), unit_cost),
                quantity=qty,
                min_stock=_safe_float(payload.get("min_stock"), self.get_configuration()["settings"].get("low_stock_alert_threshold", 10)),
                supplier=_safe_text(payload.get("supplier")),
                batch_number=_safe_text(payload.get("batch_number")),
                expiry_date=_safe_text(payload.get("expiry_date")),
                created_date=_safe_text(payload.get("received_date")) or datetime.now().strftime("%Y-%m-%d"),
                invoice_date=_safe_text(payload.get("purchase_invoice_reference")),
                addition_source="Goods Receipt",
                manufacture_date=_safe_text(payload.get("manufacture_date")),
                package_type=_safe_text(payload.get("package_type")),
                brand=_safe_text(payload.get("brand")),
                client_name=client_name,
                sku=_safe_text(payload.get("sku")) or _safe_text(payload.get("product_code")),
                supplier_name=_safe_text(payload.get("supplier")),
                unit_of_measure=_safe_text(payload.get("unit_of_measure")) or "PCS",
                pack_size=_safe_text(payload.get("pack_size")),
                vat_category=_safe_text(payload.get("vat_category")) or "Standard",
                barcode=_safe_text(payload.get("barcode")),
                lot_number=_safe_text(payload.get("lot_number")),
                warehouse_code=warehouse_code,
                bin_location=bin_location,
                reserved_quantity=0,
                max_stock=_safe_float(payload.get("max_stock"), 0.0),
                reorder_point=_safe_float(payload.get("reorder_point"), 0.0),
                reorder_quantity=_safe_float(payload.get("reorder_quantity"), 0.0),
                status="Available",
                temperature_controlled=bool(payload.get("temperature_controlled")),
                quarantine=bool(payload.get("quarantine")),
            )
        if not item_record:
            raise RuntimeError("Failed to receive stock into inventory")
        receipt = {
            "grn_id": grn_id,
            "client_name": client_name,
            "owner_company": client_name,
            "supplier": _safe_text(payload.get("supplier")),
            "purchase_invoice_reference": _safe_text(payload.get("purchase_invoice_reference")),
            "purchase_id": _safe_text(payload.get("purchase_id")),
            "received_by": _safe_text(user_name),
            "received_date": _safe_text(payload.get("received_date")) or datetime.now().strftime("%Y-%m-%d"),
            "product_name": product_name,
            "item_id": _safe_text(getattr(item_record, "item_id", "")),
            "batch_number": _safe_text(payload.get("batch_number")),
            "expiry_date": _safe_text(payload.get("expiry_date")),
            "manufacture_date": _safe_text(payload.get("manufacture_date")),
            "quantity_received": qty,
            "unit_cost": unit_cost,
            "warehouse_code": warehouse_code,
            "bin_location": bin_location,
            "location_code": _safe_text((location_record or {}).get("location_code")),
            "rack": _safe_text((location_record or {}).get("rack")),
            "shelf": _safe_text((location_record or {}).get("shelf")),
            "attachment": _safe_text(payload.get("attachment")),
            "notes": _safe_text(payload.get("notes")),
            "status": "Received",
        }
        receipts.append(receipt)
        self._save_list("receipts", receipts)
        self._log_movement("Goods Receipt", {
            "reference_id": grn_id,
            "product_name": product_name,
            "item_id": receipt["item_id"],
            "batch_number": receipt["batch_number"],
            "quantity": qty,
            "warehouse_to": warehouse_code,
            "bin_to": bin_location,
            "unit_cost": unit_cost,
            "movement_date": receipt["received_date"],
            "notes": receipt["notes"],
        }, user_name)
        self._append_audit("goods_received", "warehouse_management", grn_id, old_item, receipt, user_name)
        return receipt

    def update_goods_receipt(self, grn_id, payload, user_name="System"):
        receipts = self._load_list("receipts")
        row = next((x for x in receipts if _safe_text(x.get("grn_id")) == _safe_text(grn_id)), None)
        if not row:
            raise ValueError("Goods receipt not found")
        invm = self._get_inventory_manager()
        item = invm.get_item(_safe_text(row.get("item_id")))
        if not item:
            raise ValueError("Linked inventory item not found")

        old_receipt = copy.deepcopy(row)
        old_item = item.to_dict()
        product_name = _safe_text(payload.get("product_name"))
        warehouse_code = _safe_text(payload.get("warehouse_code"))
        client_name = _safe_text(
            payload.get("client_name") or payload.get("owner_company") or payload.get("company_name")
        )
        qty = _safe_float(payload.get("quantity_received"), row.get("quantity_received"))
        unit_cost = _safe_float(payload.get("unit_cost"), row.get("unit_cost"))
        if not product_name or qty <= 0:
            raise ValueError("Product and quantity are required")

        bin_location = _safe_text(payload.get("bin_location"))
        location_record = self._resolve_shelf_location(warehouse_code, client_name=client_name, location_hint=bin_location)
        if location_record:
            if not client_name:
                client_name = self._location_owner_name(location_record)
            bin_location = _safe_text(
                location_record.get("location_code") or location_record.get("bin") or bin_location
            )

        previous_qty = _safe_float(row.get("quantity_received"), 0.0)
        previous_total_cost = previous_qty * _safe_float(row.get("unit_cost"), 0.0)
        updated_item_qty = _safe_float(getattr(item, "quantity", 0), 0.0) - previous_qty + qty
        updated_item_total_cost = _safe_float(getattr(item, "total_cost", 0), 0.0) - previous_total_cost + (qty * unit_cost)
        if updated_item_qty < 0:
            raise ValueError("Edited receipt would make inventory quantity negative")
        if updated_item_total_cost < 0:
            updated_item_total_cost = 0.0

        invm.update_item(
            item.item_id,
            name=product_name,
            description=_safe_text(payload.get("notes")) or product_name,
            quantity=updated_item_qty,
            total_cost=updated_item_total_cost,
            supplier=_safe_text(payload.get("supplier")) or getattr(item, "supplier", ""),
            batch_number=_safe_text(payload.get("batch_number")) or getattr(item, "batch_number", ""),
            expiry_date=_safe_text(payload.get("expiry_date")) or getattr(item, "expiry_date", ""),
            manufacture_date=_safe_text(payload.get("manufacture_date")) or getattr(item, "manufacture_date", ""),
            invoice_date=_safe_text(payload.get("purchase_invoice_reference") or payload.get("purchase_id")) or getattr(item, "invoice_date", ""),
            package_type=_safe_text(payload.get("package_type")) or getattr(item, "package_type", ""),
            brand=_safe_text(payload.get("brand")) or getattr(item, "brand", ""),
            client_name=client_name or getattr(item, "client_name", ""),
            sku=_safe_text(payload.get("sku") or payload.get("product_code")) or getattr(item, "sku", ""),
            supplier_name=_safe_text(payload.get("supplier")) or getattr(item, "supplier_name", ""),
            unit_of_measure=_safe_text(payload.get("unit_of_measure")) or getattr(item, "unit_of_measure", ""),
            pack_size=_safe_text(payload.get("pack_size")) or getattr(item, "pack_size", ""),
            vat_category=_safe_text(payload.get("vat_category")) or getattr(item, "vat_category", ""),
            barcode=_safe_text(payload.get("barcode")) or getattr(item, "barcode", ""),
            lot_number=_safe_text(payload.get("lot_number")) or getattr(item, "lot_number", ""),
            warehouse_code=warehouse_code or getattr(item, "warehouse_code", ""),
            bin_location=bin_location or getattr(item, "bin_location", ""),
            status="Available",
        )

        row.update({
            "client_name": client_name,
            "owner_company": client_name,
            "supplier": _safe_text(payload.get("supplier")),
            "purchase_invoice_reference": _safe_text(payload.get("purchase_invoice_reference")),
            "purchase_id": _safe_text(payload.get("purchase_id")),
            "product_name": product_name,
            "batch_number": _safe_text(payload.get("batch_number")),
            "expiry_date": _safe_text(payload.get("expiry_date")),
            "manufacture_date": _safe_text(payload.get("manufacture_date")),
            "quantity_received": qty,
            "unit_cost": unit_cost,
            "warehouse_code": warehouse_code,
            "bin_location": bin_location,
            "location_code": _safe_text((location_record or {}).get("location_code")) or _safe_text(row.get("location_code")),
            "rack": _safe_text((location_record or {}).get("rack")) or _safe_text(row.get("rack")),
            "shelf": _safe_text((location_record or {}).get("shelf")) or _safe_text(row.get("shelf")),
            "attachment": _safe_text(payload.get("attachment")),
            "notes": _safe_text(payload.get("notes")),
        })
        self._save_list("receipts", receipts)
        self._append_audit("goods_receipt_updated", "warehouse_management", grn_id, old_receipt, row, user_name)
        self._append_audit("inventory_receipt_updated", "inventory", item.item_id, old_item, invm.get_item(item.item_id).to_dict(), user_name)
        return row

    def delete_goods_receipt(self, grn_id, user_name="System"):
        receipts = self._load_list("receipts")
        row = next((x for x in receipts if _safe_text(x.get("grn_id")) == _safe_text(grn_id)), None)
        if not row:
            raise ValueError("Goods receipt not found")
        invm = self._get_inventory_manager()
        item = invm.get_item(_safe_text(row.get("item_id")))
        if not item:
            raise ValueError("Linked inventory item not found")
        qty = _safe_float(row.get("quantity_received"), 0.0)
        current_qty = _safe_float(getattr(item, "quantity", 0), 0.0)
        if current_qty < qty:
            raise ValueError("Cannot delete this receipt because the linked stock has already been used")
        old_receipt = copy.deepcopy(row)
        old_item = item.to_dict()
        cost_per_item = item.get_cost_per_item() if hasattr(item, "get_cost_per_item") else _safe_float(row.get("unit_cost"), 0.0)
        updated_item = self._delete_or_update_item_quantity(invm, item, current_qty - qty, cost_per_item)
        receipts = [x for x in receipts if _safe_text(x.get("grn_id")) != _safe_text(grn_id)]
        self._save_list("receipts", receipts)
        self._delete_movement_records(grn_id, "Goods Receipt")
        self._append_audit("goods_receipt_deleted", "warehouse_management", grn_id, old_receipt, {}, user_name)
        self._append_audit(
            "inventory_receipt_deleted",
            "inventory",
            old_item.get("item_id"),
            old_item,
            updated_item.to_dict() if updated_item else {},
            user_name,
        )
        return True

    def create_transfer(self, payload, user_name="System"):
        transfers = self._load_list("transfers")
        transfer_id = self._generate_ref("TRF", transfers, "transfer_id")
        transfer_type = _safe_text(payload.get("transfer_type") or "Internal Transfer") or "Internal Transfer"
        record = {
            "transfer_id": transfer_id,
            "transfer_type": transfer_type,
            "requested_by": _safe_text(user_name),
            "requested_date": datetime.now().strftime("%Y-%m-%d"),
            "item_id": _safe_text(payload.get("item_id")),
            "product_name": _safe_text(payload.get("product_name")),
            "batch_number": _safe_text(payload.get("batch_number")),
            "quantity": _safe_float(payload.get("quantity"), 0.0),
            "warehouse_from": _safe_text(payload.get("warehouse_from")),
            "bin_from": _safe_text(payload.get("bin_from")),
            "warehouse_to": _safe_text(payload.get("warehouse_to")),
            "bin_to": _safe_text(payload.get("bin_to")),
            "destination_client": _safe_text(payload.get("destination_client")),
            "destination_contact": _safe_text(payload.get("destination_contact")),
            "destination_address": _safe_text(payload.get("destination_address")),
            "transfer_description": _safe_text(payload.get("transfer_description") or payload.get("item_description")),
            "sign_1": _safe_text(payload.get("sign_1")),
            "sign_2": _safe_text(payload.get("sign_2")),
            "notes": _safe_text(payload.get("notes")),
            "status": "Requested",
            "approved_by": "",
            "approved_date": "",
            "received_by": "",
            "received_date": "",
        }
        record.update(self._transfer_snapshot_from_item(payload.get("_source_item")))
        if not record["item_id"] or record["quantity"] <= 0:
            raise ValueError("Item and quantity are required")
        if transfer_type == "Client Transfer" and not record["destination_client"]:
            raise ValueError("Client name is required for client transfers")
        if transfer_type != "Client Transfer" and not record["warehouse_to"]:
            raise ValueError("Destination warehouse is required for internal transfers")
        transfers.append(record)
        self._save_list("transfers", transfers)
        self._append_audit("transfer_requested", "warehouse_management", transfer_id, {}, record, user_name)
        return record

    def update_transfer(self, transfer_id, payload, user_name="System"):
        transfers = self._load_list("transfers")
        row = next((x for x in transfers if _safe_text(x.get("transfer_id")) == _safe_text(transfer_id)), None)
        if not row:
            raise ValueError("Transfer not found")
        if _safe_text(row.get("status")) not in {"Requested", "Pending"}:
            raise ValueError("Only requested transfers can be edited")
        old_value = copy.deepcopy(row)
        transfer_type = _safe_text(payload.get("transfer_type") or row.get("transfer_type") or "Internal Transfer") or "Internal Transfer"
        row.update({
            "transfer_type": transfer_type,
            "item_id": _safe_text(payload.get("item_id")),
            "product_name": _safe_text(payload.get("product_name")),
            "batch_number": _safe_text(payload.get("batch_number")),
            "quantity": _safe_float(payload.get("quantity"), 0.0),
            "warehouse_from": _safe_text(payload.get("warehouse_from")),
            "bin_from": _safe_text(payload.get("bin_from")),
            "warehouse_to": _safe_text(payload.get("warehouse_to")),
            "bin_to": _safe_text(payload.get("bin_to")),
            "destination_client": _safe_text(payload.get("destination_client")),
            "destination_contact": _safe_text(payload.get("destination_contact")),
            "destination_address": _safe_text(payload.get("destination_address")),
            "transfer_description": _safe_text(payload.get("transfer_description") or payload.get("item_description")),
            "sign_1": _safe_text(payload.get("sign_1")),
            "sign_2": _safe_text(payload.get("sign_2")),
            "notes": _safe_text(payload.get("notes")),
        })
        row.update(self._transfer_snapshot_from_item(payload.get("_source_item")))
        if not row["item_id"] or row["quantity"] <= 0:
            raise ValueError("Item and quantity are required")
        if transfer_type == "Client Transfer" and not row["destination_client"]:
            raise ValueError("Client name is required for client transfers")
        if transfer_type != "Client Transfer" and not row["warehouse_to"]:
            raise ValueError("Destination warehouse is required for internal transfers")
        self._save_list("transfers", transfers)
        self._append_audit("transfer_updated", "warehouse_management", transfer_id, old_value, row, user_name)
        return row

    def delete_transfer(self, transfer_id, user_name="System"):
        transfers = self._load_list("transfers")
        row = next((x for x in transfers if _safe_text(x.get("transfer_id")) == _safe_text(transfer_id)), None)
        if not row:
            raise ValueError("Transfer not found")
        status = _safe_text(row.get("status"))
        old_transfer = copy.deepcopy(row)
        if status == "Received":
            invm = self._get_inventory_manager()
            source_item = invm.get_item(_safe_text(row.get("item_id")))
            if not source_item:
                raise ValueError("Source inventory item not found for reversing transfer")
            qty = _safe_float(row.get("quantity"), 0.0)
            old_source = source_item.to_dict()
            cost_per_item = source_item.get_cost_per_item() if hasattr(source_item, "get_cost_per_item") else 0.0
            if _safe_text(row.get("transfer_type")) == "Client Transfer":
                source_qty = _safe_float(getattr(source_item, "quantity", 0), 0.0) + qty
                invm.update_item(source_item.item_id, quantity=source_qty, total_cost=max(0.0, source_qty * _safe_float(cost_per_item, 0.0)))
            else:
                destination_item = self._find_inventory_item(
                    invm,
                    product_name=row.get("product_name"),
                    batch_number=row.get("batch_number"),
                    warehouse_code=row.get("warehouse_to"),
                    bin_location=row.get("bin_to"),
                )
                if not destination_item:
                    raise ValueError("Destination inventory item not found for reversing transfer")
                dest_qty = _safe_float(getattr(destination_item, "quantity", 0), 0.0)
                if dest_qty < qty:
                    raise ValueError("Cannot delete this transfer because destination stock has already been used")
                old_destination = destination_item.to_dict()
                cost_per_item = destination_item.get_cost_per_item() if hasattr(destination_item, "get_cost_per_item") else cost_per_item
                self._delete_or_update_item_quantity(invm, destination_item, dest_qty - qty, cost_per_item)
                source_qty = _safe_float(getattr(source_item, "quantity", 0), 0.0) + qty
                invm.update_item(source_item.item_id, quantity=source_qty, total_cost=max(0.0, source_qty * _safe_float(cost_per_item, 0.0)))
                dest_after = invm.get_item(old_destination.get("item_id"))
                self._append_audit("inventory_transfer_reversed_destination", "inventory", old_destination.get("item_id"), old_destination, dest_after.to_dict() if dest_after else {}, user_name)
            self._append_audit("inventory_transfer_reversed_source", "inventory", source_item.item_id, old_source, invm.get_item(source_item.item_id).to_dict(), user_name)
            movement_type = "Client Transfer" if _safe_text(row.get("transfer_type")) == "Client Transfer" else "Transfer"
            self._delete_movement_records(transfer_id, movement_type)
        elif status not in {"Requested", "Pending", "Approved", "Dispatched"}:
            raise ValueError("This transfer cannot be deleted")
        transfers = [x for x in transfers if _safe_text(x.get("transfer_id")) != _safe_text(transfer_id)]
        self._save_list("transfers", transfers)
        self._append_audit("transfer_deleted", "warehouse_management", transfer_id, old_transfer, {}, user_name)
        return True

    def approve_transfer(self, transfer_id, user_name="System"):
        transfers = self._load_list("transfers")
        for row in transfers:
            if _safe_text(row.get("transfer_id")) == _safe_text(transfer_id):
                old_value = copy.deepcopy(row)
                row["status"] = "Approved"
                row["approved_by"] = _safe_text(user_name)
                row["approved_date"] = datetime.now().strftime("%Y-%m-%d")
                self._save_list("transfers", transfers)
                self._append_audit("transfer_approved", "warehouse_management", transfer_id, old_value, row, user_name)
                return row
        raise ValueError("Transfer not found")

    def receive_transfer(self, transfer_id, user_name="System"):
        invm = self._get_inventory_manager()
        transfers = self._load_list("transfers")
        transfer = next((row for row in transfers if _safe_text(row.get("transfer_id")) == _safe_text(transfer_id)), None)
        if not transfer:
            raise ValueError("Transfer not found")
        if _safe_text(transfer.get("status")) not in {"Approved", "Dispatched"}:
            raise ValueError("Transfer must be approved before receiving")
        source_item = invm.get_item(_safe_text(transfer.get("item_id")))
        if not source_item:
            raise ValueError("Source inventory item not found")
        qty = _safe_float(transfer.get("quantity"), 0.0)
        if _safe_float(getattr(source_item, "quantity", 0), 0.0) < qty:
            raise ValueError("Insufficient source stock for transfer")
        old_source = source_item.to_dict()
        new_source_qty = _safe_float(getattr(source_item, "quantity", 0), 0.0) - qty
        cost_per_item = source_item.get_cost_per_item() if hasattr(source_item, "get_cost_per_item") else 0.0
        invm.update_item(source_item.item_id, quantity=new_source_qty, total_cost=max(0.0, new_source_qty * cost_per_item))
        if _safe_text(transfer.get("transfer_type")) != "Client Transfer":
            destination = self._find_inventory_item(
                invm,
                product_name=getattr(source_item, "name", ""),
                batch_number=getattr(source_item, "batch_number", ""),
                warehouse_code=transfer.get("warehouse_to"),
                bin_location=transfer.get("bin_to"),
            )
            if destination:
                new_qty = _safe_float(getattr(destination, "quantity", 0), 0.0) + qty
                invm.update_item(
                    destination.item_id,
                    quantity=new_qty,
                    total_cost=_safe_float(getattr(destination, "total_cost", 0), 0.0) + (qty * cost_per_item),
                    warehouse_code=transfer.get("warehouse_to"),
                    bin_location=transfer.get("bin_to"),
                )
            else:
                invm.add_item(
                    name=getattr(source_item, "name", ""),
                    category=getattr(source_item, "category", ""),
                    description=getattr(source_item, "description", ""),
                    total_cost=qty * cost_per_item,
                    selling_price=_safe_float(getattr(source_item, "selling_price", 0), 0.0),
                    quantity=qty,
                    min_stock=_safe_float(getattr(source_item, "min_stock", 0), 0.0),
                    supplier=_safe_text(getattr(source_item, "supplier", "")),
                    batch_number=_safe_text(getattr(source_item, "batch_number", "")),
                    expiry_date=_safe_text(getattr(source_item, "expiry_date", "")),
                    created_date=datetime.now().strftime("%Y-%m-%d"),
                    addition_source="Warehouse Transfer",
                    manufacture_date=_safe_text(getattr(source_item, "manufacture_date", "")),
                    package_type=_safe_text(getattr(source_item, "package_type", "")),
                    brand=_safe_text(getattr(source_item, "brand", "")),
                    client_name=_safe_text(getattr(source_item, "client_name", "")),
                    sku=_safe_text(getattr(source_item, "sku", "")),
                    supplier_name=_safe_text(getattr(source_item, "supplier_name", "")),
                    unit_of_measure=_safe_text(getattr(source_item, "unit_of_measure", "")),
                    pack_size=_safe_text(getattr(source_item, "pack_size", "")),
                    vat_category=_safe_text(getattr(source_item, "vat_category", "")),
                    barcode=_safe_text(getattr(source_item, "barcode", "")),
                    lot_number=_safe_text(getattr(source_item, "lot_number", "")),
                    warehouse_code=_safe_text(transfer.get("warehouse_to")),
                    bin_location=_safe_text(transfer.get("bin_to")),
                    status="Available",
                    temperature_controlled=bool(getattr(source_item, "temperature_controlled", False)),
                    quarantine=bool(getattr(source_item, "quarantine", False)),
                )
        old_transfer = copy.deepcopy(transfer)
        transfer["status"] = "Received"
        transfer["received_by"] = _safe_text(user_name)
        transfer["received_date"] = datetime.now().strftime("%Y-%m-%d")
        self._save_list("transfers", transfers)
        movement_type = "Client Transfer" if _safe_text(transfer.get("transfer_type")) == "Client Transfer" else "Transfer"
        self._log_movement(movement_type, {
            "reference_id": transfer_id,
            "product_name": transfer.get("product_name") or getattr(source_item, "name", ""),
            "item_id": transfer.get("item_id"),
            "batch_number": transfer.get("batch_number"),
            "quantity": qty,
            "warehouse_from": transfer.get("warehouse_from"),
            "warehouse_to": transfer.get("warehouse_to"),
            "bin_from": transfer.get("bin_from"),
            "bin_to": transfer.get("bin_to"),
            "unit_cost": cost_per_item,
            "notes": transfer.get("notes"),
        }, user_name)
        self._append_audit("transfer_received", "warehouse_management", transfer_id, old_transfer, transfer, user_name)
        self._append_audit("inventory_transfer_source_update", "inventory", source_item.item_id, old_source, invm.get_item(source_item.item_id).to_dict(), user_name)
        return transfer

    def export_transfer_pdf(self, transfer_id, open_after=True):
        transfer = self.get_transfer(transfer_id)
        if not transfer:
            raise ValueError("Transfer not found")
        from inventory_system import build_pdf_logo_flowables

        folder = os.path.join(self.company_folder, "GoodsTransfers")
        os.makedirs(folder, exist_ok=True)
        filename = f"GoodsTransfer_{_safe_text(transfer_id)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        path = os.path.join(folder, filename)

        doc = SimpleDocTemplate(
            path,
            pagesize=A4,
            leftMargin=0.45 * inch,
            rightMargin=0.45 * inch,
            topMargin=0.45 * inch,
            bottomMargin=0.45 * inch,
        )
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("TransferTitle", parent=styles["Heading1"], fontSize=18, textColor=colors.HexColor("#0b5394"), alignment=1, spaceAfter=6)
        sub_style = ParagraphStyle("TransferSub", parent=styles["Normal"], fontSize=9, textColor=colors.grey, alignment=1, spaceAfter=6)
        label_style = ParagraphStyle("TransferLabel", parent=styles["Normal"], fontSize=9, textColor=colors.white)
        value_style = ParagraphStyle("TransferValue", parent=styles["Normal"], fontSize=9, leading=12)

        elements = []
        elements.extend(build_pdf_logo_flowables(self.company_folder, width=1.0 * inch, height=0.95 * inch, spacer_height=0.05 * inch))
        elements.append(Paragraph("Goods Transfer Note", title_style))
        elements.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", sub_style))
        elements.append(Spacer(1, 0.12 * inch))

        destination_text = _safe_text(transfer.get("destination_client")) if _safe_text(transfer.get("transfer_type")) == "Client Transfer" else _safe_text(transfer.get("warehouse_to"))
        header_rows = [
            ["Transfer No.", _safe_text(transfer.get("transfer_id")), "Type", _safe_text(transfer.get("transfer_type") or "Internal Transfer")],
            ["Requested Date", _safe_text(transfer.get("requested_date")), "Requested By", _safe_text(transfer.get("requested_by"))],
            ["Status", _safe_text(transfer.get("status")), "Received Date", _safe_text(transfer.get("received_date"))],
            ["From Warehouse", _safe_text(transfer.get("warehouse_from")), "From Bin", _safe_text(transfer.get("bin_from"))],
            ["Destination", destination_text, "Destination Bin", _safe_text(transfer.get("bin_to"))],
            ["Client Contact", _safe_text(transfer.get("destination_contact")), "Client Address", _safe_text(transfer.get("destination_address"))],
        ]
        header_table = Table(header_rows, colWidths=[1.25 * inch, 2.0 * inch, 1.25 * inch, 2.45 * inch])
        header_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#17395d")),
            ("BACKGROUND", (2, 0), (2, -1), colors.HexColor("#17395d")),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.white),
            ("TEXTCOLOR", (2, 0), (2, -1), colors.white),
            ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#17395d")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        elements.append(header_table)
        elements.append(Spacer(1, 0.16 * inch))

        item_rows = [[
            Paragraph("Item Code", label_style),
            Paragraph("Description", label_style),
            Paragraph("Batch", label_style),
            Paragraph("MFG", label_style),
            Paragraph("EXP", label_style),
            Paragraph("Qty", label_style),
        ]]
        item_rows.append([
            Paragraph(_safe_text(transfer.get("item_id") or transfer.get("sku")), value_style),
            Paragraph(_safe_text(transfer.get("transfer_description") or transfer.get("item_description") or transfer.get("product_name")), value_style),
            Paragraph(_safe_text(transfer.get("batch_number")), value_style),
            Paragraph(_safe_text(transfer.get("manufacture_date")), value_style),
            Paragraph(_safe_text(transfer.get("expiry_date")), value_style),
            Paragraph(str(transfer.get("quantity") or ""), value_style),
        ])
        item_table = Table(item_rows, colWidths=[1.1 * inch, 2.75 * inch, 1.1 * inch, 0.9 * inch, 0.9 * inch, 0.75 * inch])
        item_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17395d")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#17395d")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        elements.append(item_table)
        elements.append(Spacer(1, 0.16 * inch))

        detail_rows = [
            ["Product Name", _safe_text(transfer.get("product_name"))],
            ["Brand", _safe_text(transfer.get("brand"))],
            ["SKU", _safe_text(transfer.get("sku"))],
            ["Lot Number", _safe_text(transfer.get("lot_number"))],
            ["Unit / Pack Size", f"{_safe_text(transfer.get('unit_of_measure'))} / {_safe_text(transfer.get('pack_size'))}".strip(" /")],
            ["Stored For Client", _safe_text(transfer.get("stored_for_client"))],
            ["Supplier", _safe_text(transfer.get("source_supplier"))],
            ["Notes", _safe_text(transfer.get("notes"))],
        ]
        detail_table = Table(detail_rows, colWidths=[1.65 * inch, 5.0 * inch])
        detail_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#0b5394")),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.white),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#0b5394")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        elements.append(detail_table)
        elements.append(Spacer(1, 0.28 * inch))

        signature_table = Table(
            [[
                Paragraph(f"Sign 1<br/><br/>_________________________<br/>{_safe_text(transfer.get('sign_1')) or 'Name / Signature'}", value_style),
                Paragraph(f"Sign 2<br/><br/>_________________________<br/>{_safe_text(transfer.get('sign_2')) or 'Name / Signature'}", value_style),
            ]],
            colWidths=[3.2 * inch, 3.2 * inch],
        )
        signature_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
        ]))
        elements.append(signature_table)

        doc.build(elements)
        if open_after:
            try:
                webbrowser.open("file://" + os.path.abspath(path))
            except Exception:
                pass
        return path

    def create_outbound(self, payload, user_name="System"):
        outbounds = self._load_list("outbounds")
        outbound_no = self._generate_ref(self._get_outbound_ref_prefix(), outbounds, "outbound_no")
        requested_qty = _safe_float(payload.get("requested_quantity"), 0.0)
        if requested_qty <= 0:
            raise ValueError("Requested quantity is required")
        allow_expired = bool(payload.get("allow_expired_override"))
        candidates = self._fefo_inventory_candidates(
            client_owner=payload.get("client_owner"),
            warehouse_code=payload.get("warehouse_code"),
            bin_location=payload.get("bin_location"),
            product_name=payload.get("product_name"),
            sku=payload.get("sku"),
            item_id=payload.get("item_id"),
            batch_number=payload.get("batch_number"),
            lot_number=payload.get("lot_number"),
            include_expired=allow_expired,
        )
        if not candidates:
            raise ValueError("No available stock matches the outbound request")
        source_item = candidates[0]
        available_qty = _safe_float(source_item.get_available_quantity() if hasattr(source_item, "get_available_quantity") else getattr(source_item, "quantity", 0), 0.0)
        if requested_qty > available_qty:
            raise ValueError(f"Requested quantity exceeds available quantity ({available_qty:g})")
        if hasattr(source_item, "is_expired") and source_item.is_expired() and not allow_expired:
            raise ValueError("Expired stock cannot be dispatched without admin override")
        line = self._build_outbound_item_line(source_item, requested_qty)
        expiry_days = _safe_int(self.get_configuration()["settings"].get("expiry_alert_days"), 30)
        near_expiry = bool(source_item.is_expiring_soon(expiry_days)) if hasattr(source_item, "is_expiring_soon") else False
        line["near_expiry_warning"] = near_expiry
        notes = _safe_text(payload.get("notes"))
        if near_expiry:
            notes = (notes + "\nNear-expiry stock warning.").strip()
        record = {
            "outbound_no": outbound_no,
            "delivery_note_no": "",
            "date": _safe_text(payload.get("date")) or datetime.now().strftime("%Y-%m-%d"),
            "client_owner": _safe_text(payload.get("client_owner") or getattr(source_item, "client_name", "")),
            "destination": _safe_text(payload.get("destination")),
            "warehouse_code": _safe_text(payload.get("warehouse_code") or getattr(source_item, "warehouse_code", "")),
            "bin_location": _safe_text(payload.get("bin_location") or getattr(source_item, "bin_location", "")),
            "status": "Draft",
            "requested_by": _safe_text(user_name),
            "approved_by": "",
            "picked_by": "",
            "packed_by": "",
            "dispatched_by": "",
            "received_by": "",
            "requested_date": datetime.now().strftime("%Y-%m-%d"),
            "approved_date": "",
            "picked_date": "",
            "packed_date": "",
            "dispatched_date": "",
            "received_date": "",
            "notes": notes,
            "admin_override_expired": allow_expired,
            "package_number": _safe_text(payload.get("package_number")),
            "items": [line],
        }
        if not record["client_owner"]:
            raise ValueError("Client/owner is required")
        if not record["destination"]:
            raise ValueError("Destination is required")
        outbounds.append(record)
        self._save_list("outbounds", outbounds)
        self._append_audit("outbound_created", "outbound", outbound_no, {}, record, user_name)
        return record

    def update_outbound(self, outbound_no, payload, user_name="System"):
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if not self._outbound_status_allowed(row.get("status"), {"Draft", "Pending Approval", "Approved"}):
            raise ValueError("Only draft, pending approval, or approved outbound requests can be edited")
        existing_line = (row.get("items") or [{}])[0]
        invm = self._get_inventory_manager()
        source_item = invm.get_item(_safe_text(payload.get("item_id") or existing_line.get("source_item_id") or existing_line.get("item_id")))
        if not source_item:
            candidates = self._fefo_inventory_candidates(
                client_owner=payload.get("client_owner") or row.get("client_owner"),
                warehouse_code=payload.get("warehouse_code") or row.get("warehouse_code"),
                bin_location=payload.get("bin_location") or row.get("bin_location"),
                product_name=payload.get("product_name") or existing_line.get("product_name"),
                sku=payload.get("sku") or existing_line.get("sku"),
                batch_number=payload.get("batch_number") or existing_line.get("batch_number"),
                lot_number=payload.get("lot_number") or existing_line.get("lot_number"),
                include_expired=bool(payload.get("allow_expired_override") or row.get("admin_override_expired")),
            )
            source_item = candidates[0] if candidates else None
        if not source_item:
            raise ValueError("Matching outbound stock item could not be found")
        requested_qty = _safe_float(payload.get("requested_quantity"), existing_line.get("requested_quantity"))
        if requested_qty <= 0:
            raise ValueError("Requested quantity is required")
        if requested_qty > _safe_float(source_item.get_available_quantity() if hasattr(source_item, "get_available_quantity") else getattr(source_item, "quantity", 0), 0.0):
            raise ValueError("Requested quantity exceeds available quantity")
        old_value = copy.deepcopy(row)
        row.update({
            "date": _safe_text(payload.get("date") or row.get("date")),
            "client_owner": _safe_text(payload.get("client_owner") or row.get("client_owner") or getattr(source_item, "client_name", "")),
            "destination": _safe_text(payload.get("destination") or row.get("destination")),
            "warehouse_code": _safe_text(payload.get("warehouse_code") or getattr(source_item, "warehouse_code", "")),
            "bin_location": _safe_text(payload.get("bin_location") or getattr(source_item, "bin_location", "")),
            "notes": _safe_text(payload.get("notes") or row.get("notes")),
            "admin_override_expired": bool(payload.get("allow_expired_override") or row.get("admin_override_expired")),
            "package_number": _safe_text(payload.get("package_number") or row.get("package_number")),
        })
        row["items"] = [self._build_outbound_item_line(source_item, requested_qty)]
        self._save_list("outbounds", outbounds)
        self._append_audit("outbound_updated", "outbound", outbound_no, old_value, row, user_name)
        return row

    def delete_outbound(self, outbound_no, user_name="System"):
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if _safe_text(row.get("status")) in {"Dispatched", "Delivered", "Returned"}:
            raise ValueError("Dispatched, delivered, or returned outbound requests cannot be deleted")
        if _safe_text(row.get("status")) in {"Reserved", "Picking", "Picked", "Packing", "Packed", "Ready for Dispatch"}:
            self.cancel_outbound(outbound_no, user_name=user_name, delete_mode=True)
            outbounds = self._load_list("outbounds")
            row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        old_value = copy.deepcopy(row)
        outbounds = [x for x in outbounds if _safe_text(x.get("outbound_no")) != _safe_text(outbound_no)]
        self._save_list("outbounds", outbounds)
        self._append_audit("outbound_deleted", "outbound", outbound_no, old_value, {}, user_name)
        return True

    def _update_outbound_status(self, outbound_no, new_status, user_name="System", actor_field="", extra_updates=None, audit_action=None, allowed_statuses=None):
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if allowed_statuses and not self._outbound_status_allowed(row.get("status"), allowed_statuses):
            raise ValueError(f"Outbound status must be one of: {', '.join(allowed_statuses)}")
        old_value = copy.deepcopy(row)
        row["status"] = new_status
        stamp = datetime.now().strftime("%Y-%m-%d")
        if actor_field:
            row[actor_field] = _safe_text(user_name)
            row[f"{actor_field.split('_by')[0]}_date"] = stamp
        if extra_updates:
            row.update(extra_updates)
        self._save_list("outbounds", outbounds)
        self._append_audit(audit_action or f"outbound_{new_status.lower().replace(' ', '_')}", "outbound", outbound_no, old_value, row, user_name)
        return row

    def submit_outbound_for_approval(self, outbound_no, user_name="System"):
        return self._update_outbound_status(outbound_no, "Pending Approval", user_name=user_name, allowed_statuses={"Draft"}, audit_action="outbound_submitted")

    def approve_outbound(self, outbound_no, user_name="System"):
        return self._update_outbound_status(outbound_no, "Approved", user_name=user_name, actor_field="approved_by", allowed_statuses={"Pending Approval"}, audit_action="outbound_approved")

    def reserve_outbound(self, outbound_no, user_name="System"):
        invm = self._get_inventory_manager()
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if not self._outbound_status_allowed(row.get("status"), {"Approved"}):
            raise ValueError("Only approved outbound requests can be reserved")
        line = (row.get("items") or [{}])[0]
        item = self._find_outbound_source_item(invm, line)
        if not item:
            raise ValueError("Source stock item not found")
        reserve_qty = _safe_float(line.get("requested_quantity"), 0.0)
        available_qty = _safe_float(item.get_available_quantity() if hasattr(item, "get_available_quantity") else getattr(item, "quantity", 0), 0.0)
        if reserve_qty > available_qty:
            raise ValueError(f"Cannot reserve more than available quantity ({available_qty:g})")
        old_value = copy.deepcopy(row)
        new_reserved_total = _safe_float(getattr(item, "reserved_quantity", 0), 0.0) + reserve_qty
        invm.update_item(item.item_id, reserved_quantity=new_reserved_total)
        item = invm.get_item(item.item_id)
        refreshed_line = self._refresh_outbound_line_stock(line, item)
        refreshed_line["reserved_quantity"] = reserve_qty
        refreshed_line["status"] = "Reserved"
        row["items"] = [refreshed_line]
        row["status"] = "Reserved"
        self._save_list("outbounds", outbounds)
        self._log_movement("Outbound Reserve", {
            "reference_id": outbound_no,
            "product_name": refreshed_line.get("product_name"),
            "item_id": refreshed_line.get("item_id"),
            "batch_number": refreshed_line.get("batch_number"),
            "quantity": reserve_qty,
            "warehouse_from": refreshed_line.get("warehouse_code"),
            "bin_from": refreshed_line.get("bin_location"),
            "notes": row.get("notes"),
        }, user_name)
        self._append_audit("outbound_reserved", "outbound", outbound_no, old_value, row, user_name)
        return row

    def start_outbound_picking(self, outbound_no, user_name="System"):
        return self._update_outbound_status(outbound_no, "Picking", user_name=user_name, allowed_statuses={"Reserved"}, audit_action="outbound_picking_started")

    def confirm_outbound_picked(self, outbound_no, picked_quantity, user_name="System"):
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if not self._outbound_status_allowed(row.get("status"), {"Reserved", "Picking"}):
            raise ValueError("Outbound must be reserved or in picking to confirm picked quantity")
        line = dict((row.get("items") or [{}])[0])
        picked_qty = _safe_float(picked_quantity, 0.0)
        if picked_qty <= 0 or picked_qty > _safe_float(line.get("reserved_quantity"), 0.0):
            raise ValueError("Picked quantity must be greater than zero and not exceed reserved quantity")
        old_value = copy.deepcopy(row)
        line["picked_quantity"] = picked_qty
        line["status"] = "Picked"
        row["items"] = [line]
        row["status"] = "Picked"
        row["picked_by"] = _safe_text(user_name)
        row["picked_date"] = datetime.now().strftime("%Y-%m-%d")
        self._save_list("outbounds", outbounds)
        self._append_audit("outbound_picked", "outbound", outbound_no, old_value, row, user_name)
        return row

    def start_outbound_packing(self, outbound_no, user_name="System"):
        return self._update_outbound_status(outbound_no, "Packing", user_name=user_name, allowed_statuses={"Picked"}, audit_action="outbound_packing_started")

    def confirm_outbound_packed(self, outbound_no, packed_quantity, package_number="", user_name="System"):
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if not self._outbound_status_allowed(row.get("status"), {"Picked", "Packing"}):
            raise ValueError("Outbound must be picked or in packing to confirm packed quantity")
        line = dict((row.get("items") or [{}])[0])
        packed_qty = _safe_float(packed_quantity, 0.0)
        if packed_qty <= 0 or packed_qty > _safe_float(line.get("picked_quantity"), 0.0):
            raise ValueError("Packed quantity must be greater than zero and not exceed picked quantity")
        old_value = copy.deepcopy(row)
        line["packed_quantity"] = packed_qty
        line["status"] = "Packed"
        row["items"] = [line]
        row["status"] = "Packed"
        row["packed_by"] = _safe_text(user_name)
        row["packed_date"] = datetime.now().strftime("%Y-%m-%d")
        row["package_number"] = _safe_text(package_number) or _safe_text(row.get("package_number"))
        self._save_list("outbounds", outbounds)
        self._append_audit("outbound_packed", "outbound", outbound_no, old_value, row, user_name)
        return row

    def mark_outbound_ready_for_dispatch(self, outbound_no, user_name="System"):
        return self._update_outbound_status(outbound_no, "Ready for Dispatch", user_name=user_name, allowed_statuses={"Packed"}, audit_action="outbound_ready_for_dispatch")

    def dispatch_outbound(self, outbound_no, dispatched_quantity=None, user_name="System"):
        invm = self._get_inventory_manager()
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if not self._outbound_status_allowed(row.get("status"), {"Ready for Dispatch", "Packed"}):
            raise ValueError("Outbound must be packed or ready for dispatch")
        line = dict((row.get("items") or [{}])[0])
        item = self._find_outbound_source_item(invm, line)
        if not item:
            raise ValueError("Source stock item not found")
        ship_qty = _safe_float(dispatched_quantity or line.get("packed_quantity") or line.get("picked_quantity") or line.get("reserved_quantity"), 0.0)
        if ship_qty <= 0:
            raise ValueError("Dispatched quantity must be greater than zero")
        reserved_qty = _safe_float(getattr(item, "reserved_quantity", 0), 0.0)
        on_hand_qty = _safe_float(getattr(item, "quantity", 0), 0.0)
        if ship_qty > reserved_qty:
            raise ValueError("Dispatched quantity cannot exceed reserved quantity")
        if ship_qty > on_hand_qty:
            raise ValueError("Insufficient on-hand quantity to dispatch")
        old_value = copy.deepcopy(row)
        old_item = item.to_dict()
        cost_per_item = item.get_cost_per_item() if hasattr(item, "get_cost_per_item") else _safe_float(line.get("unit_cost_snapshot"), 0.0)
        new_qty = on_hand_qty - ship_qty
        new_reserved = max(0.0, reserved_qty - ship_qty)
        invm.update_item(item.item_id, quantity=new_qty, reserved_quantity=new_reserved, total_cost=max(0.0, new_qty * cost_per_item))
        item = invm.get_item(item.item_id)
        line = self._refresh_outbound_line_stock(line, item)
        line["reserved_quantity"] = new_reserved
        line["dispatched_quantity"] = ship_qty
        line["status"] = "Dispatched"
        row["items"] = [line]
        row["status"] = "Dispatched"
        row["dispatched_by"] = _safe_text(user_name)
        row["dispatched_date"] = datetime.now().strftime("%Y-%m-%d")
        if not _safe_text(row.get("delivery_note_no")):
            row["delivery_note_no"] = self._generate_ref("DEL", self.list_outbounds(), "delivery_note_no")
        self._save_list("outbounds", outbounds)
        self._log_movement("Outbound Dispatch", {
            "reference_id": outbound_no,
            "product_name": line.get("product_name"),
            "item_id": line.get("item_id"),
            "batch_number": line.get("batch_number"),
            "quantity": ship_qty,
            "warehouse_from": row.get("warehouse_code"),
            "warehouse_to": row.get("destination"),
            "bin_from": row.get("bin_location"),
            "unit_cost": cost_per_item,
            "notes": row.get("notes"),
        }, user_name)
        self._append_audit("outbound_dispatched", "outbound", outbound_no, old_value, row, user_name)
        self._append_audit("inventory_outbound_dispatched", "inventory", old_item.get("item_id"), old_item, item.to_dict() if item else {}, user_name)
        return row

    def mark_outbound_delivered(self, outbound_no, received_by="", user_name="System"):
        extra = {"received_by": _safe_text(received_by) or _safe_text(user_name), "received_date": datetime.now().strftime("%Y-%m-%d")}
        return self._update_outbound_status(outbound_no, "Delivered", user_name=user_name, allowed_statuses={"Dispatched"}, extra_updates=extra, audit_action="outbound_delivered")

    def cancel_outbound(self, outbound_no, user_name="System", delete_mode=False):
        invm = self._get_inventory_manager()
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if _safe_text(row.get("status")) in {"Dispatched", "Delivered", "Returned"}:
            raise ValueError("Dispatched, delivered, or returned outbounds cannot be cancelled")
        line = dict((row.get("items") or [{}])[0])
        item = self._find_outbound_source_item(invm, line)
        old_value = copy.deepcopy(row)
        if item and _safe_float(line.get("reserved_quantity"), 0.0) > 0:
            release_qty = min(_safe_float(line.get("reserved_quantity"), 0.0), _safe_float(getattr(item, "reserved_quantity", 0), 0.0))
            invm.update_item(item.item_id, reserved_quantity=max(0.0, _safe_float(getattr(item, "reserved_quantity", 0), 0.0) - release_qty))
            line["available_quantity"] = _safe_float(invm.get_item(item.item_id).get_available_quantity(), 0.0)
        line["reserved_quantity"] = 0.0
        line["status"] = "Cancelled"
        row["items"] = [line]
        row["status"] = "Cancelled"
        self._save_list("outbounds", outbounds)
        self._delete_movement_records(outbound_no, "Outbound Reserve")
        self._append_audit("outbound_cancelled", "outbound", outbound_no, old_value, row, user_name)
        if delete_mode:
            return row
        return row

    def return_outbound(self, outbound_no, return_quantity, warehouse_code="", bin_location="", quarantine=False, notes="", user_name="System"):
        invm = self._get_inventory_manager()
        outbounds = self._load_list("outbounds")
        row = next((x for x in outbounds if _safe_text(x.get("outbound_no")) == _safe_text(outbound_no)), None)
        if not row:
            raise ValueError("Outbound request not found")
        if not self._outbound_status_allowed(row.get("status"), {"Dispatched", "Delivered"}):
            raise ValueError("Only dispatched or delivered outbounds can be returned")
        line = dict((row.get("items") or [{}])[0])
        qty = _safe_float(return_quantity, 0.0)
        if qty <= 0 or qty > _safe_float(line.get("dispatched_quantity"), 0.0):
            raise ValueError("Return quantity must be greater than zero and not exceed dispatched quantity")
        target_wh = _safe_text(warehouse_code) or _safe_text(row.get("warehouse_code"))
        target_bin = _safe_text(bin_location) or _safe_text(row.get("bin_location"))
        existing = self._find_inventory_item(
            invm,
            product_name=line.get("product_name"),
            batch_number=line.get("batch_number"),
            warehouse_code=target_wh,
            bin_location=target_bin,
        )
        old_value = copy.deepcopy(row)
        cost_per_item = _safe_float(line.get("unit_cost_snapshot"), 0.0)
        if existing:
            new_qty = _safe_float(getattr(existing, "quantity", 0), 0.0) + qty
            invm.update_item(existing.item_id, quantity=new_qty, total_cost=max(0.0, new_qty * cost_per_item), quarantine=bool(quarantine))
        else:
            invm.add_item(
                name=line.get("product_name"),
                category="Pharmaceutical",
                description=line.get("item_description"),
                total_cost=qty * cost_per_item,
                selling_price=0.0,
                quantity=qty,
                min_stock=0,
                supplier=line.get("supplier_name"),
                batch_number=line.get("batch_number"),
                expiry_date=line.get("expiry_date"),
                created_date=datetime.now().strftime("%Y-%m-%d"),
                addition_source="Outbound Return",
                manufacture_date=line.get("manufacture_date"),
                package_type="",
                brand=line.get("brand"),
                client_name=row.get("client_owner"),
                sku=line.get("sku"),
                supplier_name=line.get("supplier_name"),
                unit_of_measure=line.get("unit_of_measure"),
                pack_size=line.get("pack_size"),
                lot_number=line.get("lot_number"),
                warehouse_code=target_wh,
                bin_location=target_bin,
                status="Available",
                quarantine=bool(quarantine),
            )
        line["returned_quantity"] = qty
        line["status"] = "Returned"
        row["items"] = [line]
        row["status"] = "Returned"
        row["notes"] = "\n".join([part for part in [_safe_text(row.get("notes")), _safe_text(notes)] if part]).strip()
        self._save_list("outbounds", outbounds)
        self._log_movement("Outbound Return", {
            "reference_id": outbound_no,
            "product_name": line.get("product_name"),
            "item_id": line.get("item_id"),
            "batch_number": line.get("batch_number"),
            "quantity": qty,
            "warehouse_to": target_wh,
            "bin_to": target_bin,
            "unit_cost": cost_per_item,
            "notes": notes or "Returned outbound goods",
        }, user_name)
        self._append_audit("outbound_returned", "outbound", outbound_no, old_value, row, user_name)
        return row

    def get_outbound_report_rows(self, filters=None):
        filters = filters or {}
        rows = []
        date_from = self._parse_iso_date(filters.get("date_from"))
        date_to = self._parse_iso_date(filters.get("date_to"))
        for outbound in self.list_outbounds():
            outbound_date = self._parse_iso_date(outbound.get("date"))
            if date_from and outbound_date and outbound_date < date_from:
                continue
            if date_to and outbound_date and outbound_date > date_to:
                continue
            probes = {
                "client_owner": _safe_text(filters.get("client_owner")).lower(),
                "destination": _safe_text(filters.get("destination")).lower(),
                "warehouse_code": _safe_text(filters.get("warehouse_code")).lower(),
                "status": _safe_text(filters.get("status")).lower(),
                "requested_by": _safe_text(filters.get("requested_by")).lower(),
                "dispatched_by": _safe_text(filters.get("dispatched_by")).lower(),
            }
            if probes["client_owner"] and probes["client_owner"] not in _safe_text(outbound.get("client_owner")).lower():
                continue
            if probes["destination"] and probes["destination"] not in _safe_text(outbound.get("destination")).lower():
                continue
            if probes["warehouse_code"] and probes["warehouse_code"] not in _safe_text(outbound.get("warehouse_code")).lower():
                continue
            if probes["status"] and probes["status"] != _safe_text(outbound.get("status")).lower():
                continue
            if probes["requested_by"] and probes["requested_by"] not in _safe_text(outbound.get("requested_by")).lower():
                continue
            if probes["dispatched_by"] and probes["dispatched_by"] not in _safe_text(outbound.get("dispatched_by")).lower():
                continue
            for line in outbound.get("items") or []:
                sku_probe = _safe_text(filters.get("sku")).lower()
                product_probe = _safe_text(filters.get("product_name")).lower()
                batch_probe = _safe_text(filters.get("batch_number")).lower()
                expiry_probe = _safe_text(filters.get("expiry_date")).lower()
                if sku_probe and sku_probe not in _safe_text(line.get("sku")).lower():
                    continue
                if product_probe and product_probe not in _safe_text(line.get("product_name")).lower():
                    continue
                if batch_probe and batch_probe not in _safe_text(line.get("batch_number")).lower():
                    continue
                if expiry_probe and expiry_probe not in _safe_text(line.get("expiry_date")).lower():
                    continue
                rows.append({
                    "Outbound No.": outbound.get("outbound_no"),
                    "Transfer/Delivery No.": outbound.get("delivery_note_no"),
                    "Date": outbound.get("date"),
                    "Client/Owner": outbound.get("client_owner"),
                    "Destination": outbound.get("destination"),
                    "Warehouse": outbound.get("warehouse_code"),
                    "Product": line.get("product_name"),
                    "SKU": line.get("sku"),
                    "Batch Number": line.get("batch_number"),
                    "Lot Number": line.get("lot_number"),
                    "Manufacturing Date": line.get("manufacture_date"),
                    "Expiry Date": line.get("expiry_date"),
                    "Quantity Shipped": line.get("dispatched_quantity") or line.get("packed_quantity") or line.get("picked_quantity") or line.get("requested_quantity"),
                    "Unit/Pack Size": line.get("unit_pack_size"),
                    "Status": outbound.get("status"),
                    "Requested By": outbound.get("requested_by"),
                    "Picked By": outbound.get("picked_by"),
                    "Packed By": outbound.get("packed_by"),
                    "Dispatched By": outbound.get("dispatched_by"),
                    "Received By": outbound.get("received_by"),
                    "Notes": outbound.get("notes"),
                })
        return rows

    def _export_outbound_rows_csv(self, rows, filename_prefix="OutboundReport", open_after=True):
        folder = os.path.join(self.company_folder, "OutboundReports")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
        with open(path, "w", encoding="utf-8", newline="") as handle:
            if rows:
                writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
                writer.writeheader()
                writer.writerows(rows)
            else:
                handle.write("No data\n")
        if open_after:
            try:
                webbrowser.open("file://" + os.path.abspath(path))
            except Exception:
                pass
        return path

    def _export_outbound_rows_xlsx(self, rows, filename_prefix="OutboundReport", open_after=True):
        from openpyxl import Workbook

        folder = os.path.join(self.company_folder, "OutboundReports")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "Outbound Report"
        headers = list(rows[0].keys()) if rows else ["No Data"]
        ws.append(headers)
        for row in rows:
            ws.append([row.get(h, "") for h in headers])
        wb.save(path)
        if open_after:
            try:
                webbrowser.open("file://" + os.path.abspath(path))
            except Exception:
                pass
        return path

    def _export_outbound_rows_pdf(self, rows, title, filename_prefix, open_after=True):
        from inventory_system import build_pdf_logo_flowables

        folder = os.path.join(self.company_folder, "OutboundReports")
        os.makedirs(folder, exist_ok=True)
        path = os.path.join(folder, f"{filename_prefix}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=0.35 * inch, rightMargin=0.35 * inch, topMargin=0.4 * inch, bottomMargin=0.4 * inch)
        styles = getSampleStyleSheet()
        head = ParagraphStyle("OutboundHead", parent=styles["Heading1"], alignment=1, textColor=colors.HexColor("#0b5394"), fontSize=16, spaceAfter=8)
        body = []
        body.extend(build_pdf_logo_flowables(self.company_folder, width=1.0 * inch, height=0.95 * inch, spacer_height=0.04 * inch))
        body.append(Paragraph(title, head))
        body.append(Paragraph(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", styles["Normal"]))
        body.append(Spacer(1, 0.12 * inch))
        headers = list(rows[0].keys()) if rows else ["Message"]
        table_rows = [headers]
        if rows:
            for row in rows:
                table_rows.append([_safe_text(row.get(col)) for col in headers])
        else:
            table_rows.append(["No data available"])
        table = Table(table_rows, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#17395d")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#17395d")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
        ]))
        body.append(table)
        doc.build(body)
        if open_after:
            try:
                webbrowser.open("file://" + os.path.abspath(path))
            except Exception:
                pass
        return path

    def _export_outbound_document_pdf(self, outbound_no, title, headers, item_headers, item_rows, signature_labels=None, open_after=True):
        outbound = self.get_outbound(outbound_no)
        if not outbound:
            raise ValueError("Outbound request not found")
        from inventory_system import build_pdf_logo_flowables

        folder = os.path.join(self.company_folder, "OutboundDocuments")
        os.makedirs(folder, exist_ok=True)
        safe_title = title.replace(" ", "")
        path = os.path.join(folder, f"{safe_title}_{_safe_text(outbound_no)}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        doc = SimpleDocTemplate(path, pagesize=A4, leftMargin=0.45 * inch, rightMargin=0.45 * inch, topMargin=0.45 * inch, bottomMargin=0.45 * inch)
        styles = getSampleStyleSheet()
        title_style = ParagraphStyle("OutboundDocTitle", parent=styles["Heading1"], alignment=1, fontSize=18, textColor=colors.HexColor("#0b5394"), spaceAfter=8)
        label_style = ParagraphStyle("OutboundDocLabel", parent=styles["Normal"], fontSize=9, textColor=colors.white)
        value_style = ParagraphStyle("OutboundDocValue", parent=styles["Normal"], fontSize=9, leading=11)
        body = []
        body.extend(build_pdf_logo_flowables(self.company_folder, width=1.0 * inch, height=0.95 * inch, spacer_height=0.04 * inch))
        body.append(Paragraph(title, title_style))
        body.append(Spacer(1, 0.08 * inch))
        header_rows = [[key, _safe_text(value)] for key, value in headers]
        header_table = Table(header_rows, colWidths=[1.8 * inch, 4.5 * inch])
        header_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#17395d")),
            ("TEXTCOLOR", (0, 0), (0, -1), colors.white),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#17395d")),
        ]))
        body.append(header_table)
        body.append(Spacer(1, 0.14 * inch))
        rows = [[Paragraph(h, label_style) for h in item_headers]]
        for raw in item_rows:
            rows.append([Paragraph(_safe_text(val), value_style) for val in raw])
        item_table = Table(rows, repeatRows=1)
        item_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0b5394")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("INNERGRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cbd5e1")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#0b5394")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        body.append(item_table)
        if signature_labels:
            body.append(Spacer(1, 0.24 * inch))
            sig_table = Table([[Paragraph(f"{sig}<br/><br/>_________________________", value_style) for sig in signature_labels]])
            body.append(sig_table)
        doc.build(body)
        if open_after:
            try:
                webbrowser.open("file://" + os.path.abspath(path))
            except Exception:
                pass
        return path

    def export_pick_list_pdf(self, outbound_no, open_after=True):
        outbound = self.get_outbound(outbound_no)
        if not outbound:
            raise ValueError("Outbound request not found")
        line = (outbound.get("items") or [{}])[0]
        return self._export_outbound_document_pdf(
            outbound_no,
            "Pick List",
            [
                ("Outbound No.", outbound.get("outbound_no")),
                ("Client", outbound.get("client_owner")),
                ("Warehouse", outbound.get("warehouse_code")),
                ("Bin Location", outbound.get("bin_location")),
            ],
            ["Product", "Batch", "Expiry", "Quantity To Pick"],
            [[line.get("product_name"), line.get("batch_number"), line.get("expiry_date"), line.get("requested_quantity")]],
            signature_labels=["Picker Signature"],
            open_after=open_after,
        )

    def export_packing_list_pdf(self, outbound_no, open_after=True):
        outbound = self.get_outbound(outbound_no)
        if not outbound:
            raise ValueError("Outbound request not found")
        line = (outbound.get("items") or [{}])[0]
        return self._export_outbound_document_pdf(
            outbound_no,
            "Packing List",
            [
                ("Outbound No.", outbound.get("outbound_no")),
                ("Client", outbound.get("client_owner")),
                ("Packed By", outbound.get("packed_by")),
                ("Package Number", outbound.get("package_number")),
            ],
            ["Product", "Batch", "Quantity Packed"],
            [[line.get("product_name"), line.get("batch_number"), line.get("packed_quantity")]],
            signature_labels=["Packed By Signature"],
            open_after=open_after,
        )

    def export_outbound_delivery_note_pdf(self, outbound_no, open_after=True):
        outbound = self.get_outbound(outbound_no)
        if not outbound:
            raise ValueError("Outbound request not found")
        line = (outbound.get("items") or [{}])[0]
        return self._export_outbound_document_pdf(
            outbound_no,
            "Goods Transfer Note / Delivery Note",
            [
                ("Transfer/Delivery No.", outbound.get("delivery_note_no") or outbound.get("outbound_no")),
                ("Outbound No.", outbound.get("outbound_no")),
                ("Client", outbound.get("client_owner")),
                ("Destination", outbound.get("destination")),
                ("Warehouse", outbound.get("warehouse_code")),
                ("Date", outbound.get("dispatched_date") or outbound.get("date")),
            ],
            ["Product", "SKU", "Batch", "MFG Date", "EXP Date", "Quantity", "Unit/Pack Size", "Notes"],
            [[
                line.get("product_name"),
                line.get("sku"),
                line.get("batch_number"),
                line.get("manufacture_date"),
                line.get("expiry_date"),
                line.get("dispatched_quantity") or line.get("packed_quantity") or line.get("requested_quantity"),
                line.get("unit_pack_size"),
                outbound.get("notes"),
            ]],
            signature_labels=["Prepared By", "Received By"],
            open_after=open_after,
        )

    def create_adjustment(self, payload, user_name="System"):
        adjustments = self._load_list("adjustments")
        adjustment_id = self._generate_ref("ADJ", adjustments, "adjustment_id")
        record = {
            "adjustment_id": adjustment_id,
            "item_id": _safe_text(payload.get("item_id")),
            "product_name": _safe_text(payload.get("product_name")),
            "batch_number": _safe_text(payload.get("batch_number")),
            "quantity_change": _safe_float(payload.get("quantity_change"), 0.0),
            "adjustment_type": _safe_text(payload.get("adjustment_type") or "Correction"),
            "reason": _safe_text(payload.get("reason")),
            "warehouse_code": _safe_text(payload.get("warehouse_code")),
            "bin_location": _safe_text(payload.get("bin_location")),
            "notes": _safe_text(payload.get("notes")),
            "requested_by": _safe_text(user_name),
            "requested_date": datetime.now().strftime("%Y-%m-%d"),
            "status": "Pending",
        }
        if not record["item_id"] or not record["reason"]:
            raise ValueError("Item and reason are required")
        adjustments.append(record)
        self._save_list("adjustments", adjustments)
        self._append_audit("adjustment_requested", "warehouse_management", adjustment_id, {}, record, user_name)
        return record

    def update_adjustment(self, adjustment_id, payload, user_name="System"):
        adjustments = self._load_list("adjustments")
        row = next((x for x in adjustments if _safe_text(x.get("adjustment_id")) == _safe_text(adjustment_id)), None)
        if not row:
            raise ValueError("Adjustment not found")
        if _safe_text(row.get("status")) != "Pending":
            raise ValueError("Only pending adjustments can be edited")
        old_value = copy.deepcopy(row)
        row.update({
            "item_id": _safe_text(payload.get("item_id")),
            "product_name": _safe_text(payload.get("product_name")),
            "batch_number": _safe_text(payload.get("batch_number")),
            "quantity_change": _safe_float(payload.get("quantity_change"), 0.0),
            "adjustment_type": _safe_text(payload.get("adjustment_type") or "Correction"),
            "reason": _safe_text(payload.get("reason")),
            "warehouse_code": _safe_text(payload.get("warehouse_code")),
            "bin_location": _safe_text(payload.get("bin_location")),
            "notes": _safe_text(payload.get("notes")),
        })
        if not row["item_id"] or not row["reason"]:
            raise ValueError("Item and reason are required")
        self._save_list("adjustments", adjustments)
        self._append_audit("adjustment_updated", "warehouse_management", adjustment_id, old_value, row, user_name)
        return row

    def delete_adjustment(self, adjustment_id, user_name="System"):
        adjustments = self._load_list("adjustments")
        row = next((x for x in adjustments if _safe_text(x.get("adjustment_id")) == _safe_text(adjustment_id)), None)
        if not row:
            raise ValueError("Adjustment not found")
        old_row = copy.deepcopy(row)
        status = _safe_text(row.get("status"))
        if status == "Approved":
            invm = self._get_inventory_manager()
            item = invm.get_item(_safe_text(row.get("item_id")))
            if not item:
                raise ValueError("Inventory item not found")
            qty_change = _safe_float(row.get("quantity_change"), 0.0)
            current_qty = _safe_float(getattr(item, "quantity", 0), 0.0)
            reverted_qty = current_qty - qty_change
            if reverted_qty < 0:
                raise ValueError("Cannot delete this adjustment because later stock usage depends on it")
            old_item = item.to_dict()
            cost_per_item = item.get_cost_per_item() if hasattr(item, "get_cost_per_item") else 0.0
            updated_item = self._delete_or_update_item_quantity(invm, item, reverted_qty, cost_per_item)
            self._delete_movement_records(adjustment_id, "Adjustment")
            self._append_audit("inventory_adjustment_deleted", "inventory", old_item.get("item_id"), old_item, updated_item.to_dict() if updated_item else {}, user_name)
        elif status != "Pending":
            raise ValueError("This adjustment cannot be deleted")
        adjustments = [x for x in adjustments if _safe_text(x.get("adjustment_id")) != _safe_text(adjustment_id)]
        self._save_list("adjustments", adjustments)
        self._append_audit("adjustment_deleted", "warehouse_management", adjustment_id, old_row, {}, user_name)
        return True

    def approve_adjustment(self, adjustment_id, user_name="System"):
        invm = self._get_inventory_manager()
        adjustments = self._load_list("adjustments")
        row = next((x for x in adjustments if _safe_text(x.get("adjustment_id")) == _safe_text(adjustment_id)), None)
        if not row:
            raise ValueError("Adjustment not found")
        item = invm.get_item(_safe_text(row.get("item_id")))
        if not item:
            raise ValueError("Inventory item not found")
        qty_change = _safe_float(row.get("quantity_change"), 0.0)
        new_qty = _safe_float(getattr(item, "quantity", 0), 0.0) + qty_change
        if new_qty < 0:
            raise ValueError("Adjustment would make stock negative")
        old_item = item.to_dict()
        cost_per_item = item.get_cost_per_item() if hasattr(item, "get_cost_per_item") else 0.0
        invm.update_item(item.item_id, quantity=new_qty, total_cost=max(0.0, new_qty * cost_per_item))
        old_row = copy.deepcopy(row)
        row["status"] = "Approved"
        row["approved_by"] = _safe_text(user_name)
        row["approved_date"] = datetime.now().strftime("%Y-%m-%d")
        self._save_list("adjustments", adjustments)
        self._log_movement("Adjustment", {
            "reference_id": adjustment_id,
            "product_name": row.get("product_name") or getattr(item, "name", ""),
            "item_id": item.item_id,
            "batch_number": row.get("batch_number"),
            "quantity": qty_change,
            "warehouse_to": row.get("warehouse_code"),
            "bin_to": row.get("bin_location"),
            "unit_cost": cost_per_item,
            "notes": row.get("reason"),
        }, user_name)
        self._append_audit("adjustment_approved", "warehouse_management", adjustment_id, old_row, row, user_name)
        self._append_audit("inventory_adjusted", "inventory", item.item_id, old_item, invm.get_item(item.item_id).to_dict(), user_name)
        return row

    def create_stock_count(self, payload, user_name="System"):
        counts = self._load_list("counts")
        count_id = self._generate_ref("CNT", counts, "count_id")
        record = {
            "count_id": count_id,
            "item_id": _safe_text(payload.get("item_id")),
            "product_name": _safe_text(payload.get("product_name")),
            "warehouse_code": _safe_text(payload.get("warehouse_code")),
            "bin_location": _safe_text(payload.get("bin_location")),
            "system_quantity": _safe_float(payload.get("system_quantity"), 0.0),
            "actual_quantity": _safe_float(payload.get("actual_quantity"), 0.0),
            "variance": _safe_float(payload.get("actual_quantity"), 0.0) - _safe_float(payload.get("system_quantity"), 0.0),
            "count_type": _safe_text(payload.get("count_type") or "Physical"),
            "notes": _safe_text(payload.get("notes")),
            "requested_by": _safe_text(user_name),
            "requested_date": datetime.now().strftime("%Y-%m-%d"),
            "status": "Pending",
        }
        counts.append(record)
        self._save_list("counts", counts)
        self._append_audit("stock_count_requested", "warehouse_management", count_id, {}, record, user_name)
        return record

    def update_stock_count(self, count_id, payload, user_name="System"):
        counts = self._load_list("counts")
        row = next((x for x in counts if _safe_text(x.get("count_id")) == _safe_text(count_id)), None)
        if not row:
            raise ValueError("Stock count not found")
        if _safe_text(row.get("status")) != "Pending":
            raise ValueError("Only pending stock counts can be edited")
        old_value = copy.deepcopy(row)
        system_qty = _safe_float(payload.get("system_quantity"), row.get("system_quantity"))
        actual_qty = _safe_float(payload.get("actual_quantity"), row.get("actual_quantity"))
        row.update({
            "item_id": _safe_text(payload.get("item_id")),
            "product_name": _safe_text(payload.get("product_name")),
            "warehouse_code": _safe_text(payload.get("warehouse_code")),
            "bin_location": _safe_text(payload.get("bin_location")),
            "system_quantity": system_qty,
            "actual_quantity": actual_qty,
            "variance": actual_qty - system_qty,
            "count_type": _safe_text(payload.get("count_type") or "Physical"),
            "notes": _safe_text(payload.get("notes")),
        })
        self._save_list("counts", counts)
        self._append_audit("stock_count_updated", "warehouse_management", count_id, old_value, row, user_name)
        return row

    def delete_stock_count(self, count_id, user_name="System"):
        counts = self._load_list("counts")
        row = next((x for x in counts if _safe_text(x.get("count_id")) == _safe_text(count_id)), None)
        if not row:
            raise ValueError("Stock count not found")
        old_row = copy.deepcopy(row)
        status = _safe_text(row.get("status"))
        if status == "Approved":
            invm = self._get_inventory_manager()
            item = invm.get_item(_safe_text(row.get("item_id")))
            if not item:
                raise ValueError("Inventory item not found")
            old_item = item.to_dict()
            restored_qty = _safe_float(row.get("system_quantity"), 0.0)
            cost_per_item = item.get_cost_per_item() if hasattr(item, "get_cost_per_item") else 0.0
            updated_item = self._delete_or_update_item_quantity(invm, item, restored_qty, cost_per_item)
            self._delete_movement_records(count_id, "Stock Count")
            self._append_audit("inventory_stock_count_deleted", "inventory", old_item.get("item_id"), old_item, updated_item.to_dict() if updated_item else {}, user_name)
        elif status != "Pending":
            raise ValueError("This stock count cannot be deleted")
        counts = [x for x in counts if _safe_text(x.get("count_id")) != _safe_text(count_id)]
        self._save_list("counts", counts)
        self._append_audit("stock_count_deleted", "warehouse_management", count_id, old_row, {}, user_name)
        return True

    def approve_stock_count(self, count_id, user_name="System"):
        invm = self._get_inventory_manager()
        counts = self._load_list("counts")
        row = next((x for x in counts if _safe_text(x.get("count_id")) == _safe_text(count_id)), None)
        if not row:
            raise ValueError("Stock count not found")
        item = invm.get_item(_safe_text(row.get("item_id")))
        if not item:
            raise ValueError("Inventory item not found")
        actual_qty = _safe_float(row.get("actual_quantity"), 0.0)
        old_item = item.to_dict()
        cost_per_item = item.get_cost_per_item() if hasattr(item, "get_cost_per_item") else 0.0
        invm.update_item(item.item_id, quantity=actual_qty, total_cost=max(0.0, actual_qty * cost_per_item))
        old_row = copy.deepcopy(row)
        row["status"] = "Approved"
        row["approved_by"] = _safe_text(user_name)
        row["approved_date"] = datetime.now().strftime("%Y-%m-%d")
        self._save_list("counts", counts)
        self._log_movement("Stock Count", {
            "reference_id": count_id,
            "product_name": row.get("product_name") or getattr(item, "name", ""),
            "item_id": item.item_id,
            "quantity": row.get("variance"),
            "warehouse_to": row.get("warehouse_code"),
            "bin_to": row.get("bin_location"),
            "unit_cost": cost_per_item,
            "notes": row.get("notes"),
        }, user_name)
        self._append_audit("stock_count_approved", "warehouse_management", count_id, old_row, row, user_name)
        self._append_audit("inventory_count_adjusted", "inventory", item.item_id, old_item, invm.get_item(item.item_id).to_dict(), user_name)
        return row

    def get_dashboard_metrics(self):
        from inventory_system import InventoryManager
        invm = InventoryManager(self.company_folder)
        items = list(invm.items or [])
        total_value = sum(_safe_float(getattr(item, "total_cost", 0), 0.0) for item in items)
        expired = []
        near_expiry = []
        dead_stock = []
        threshold_days = _safe_int(self.get_configuration()["settings"].get("expiry_alert_days"), 30)
        ninety_days_ago = datetime.now().date() - timedelta(days=90)
        for item in items:
            if hasattr(item, "is_expired") and item.is_expired():
                expired.append(item)
            if hasattr(item, "is_expiring_soon") and item.is_expiring_soon(threshold_days):
                near_expiry.append(item)
            try:
                updated = datetime.strptime(_safe_text(getattr(item, "last_updated", "")), "%Y-%m-%d").date()
                if updated <= ninety_days_ago and _safe_float(getattr(item, "quantity", 0), 0.0) > 0:
                    dead_stock.append(item)
            except Exception:
                pass
        receipts = list(reversed(self.list_goods_receipts()[-5:]))
        transfers = list(reversed(self.list_transfers()[-5:]))
        pending_approvals = sum(
            1 for row in (self.list_transfers() + self.list_adjustments() + self.list_counts())
            if _safe_text(row.get("status")) == "Pending"
        )
        stock_by_warehouse = {}
        for item in items:
            wh = _safe_text(getattr(item, "warehouse_code", "")) or "Unassigned"
            stock_by_warehouse[wh] = stock_by_warehouse.get(wh, 0.0) + _safe_float(getattr(item, "quantity", 0), 0.0)
        return {
            "total_inventory_value": total_value,
            "total_stock_items": len(items),
            "expired_items": len(expired),
            "near_expiry_items": len(near_expiry),
            "dead_stock_items": len(dead_stock),
            "recent_goods_received": receipts,
            "recent_stock_transfers": transfers,
            "pending_approvals": pending_approvals,
            "stock_by_warehouse": stock_by_warehouse,
        }


class CompanyConfigurationFrame(ttk.Frame):
    def __init__(self, parent, initial_config=None):
        super().__init__(parent)
        self._module_vars = {}
        initial = copy.deepcopy(initial_config or default_company_configuration())
        self.company_type_var = tk.StringVar(value=initial.get("company_type", "Full ERP"))
        settings = initial.get("settings") or {}
        enabled_modules = initial.get("enabled_modules") or default_enabled_modules(self.company_type_var.get())
        dashboard_widgets = settings.get("dashboard_widgets") or {key: True for key in DASHBOARD_WIDGET_KEYS}

        top = ttk.LabelFrame(self, text="Company Configuration", padding=12)
        top.pack(fill="both", expand=True)
        ttk.Label(top, text="Company Type").grid(row=0, column=0, sticky="w", pady=4)
        type_combo = ttk.Combobox(top, textvariable=self.company_type_var, values=COMPANY_TYPES, state="readonly")
        type_combo.grid(row=0, column=1, sticky="ew", pady=4)
        top.columnconfigure(1, weight=1)
        type_combo.bind("<<ComboboxSelected>>", self._apply_company_type_preset)

        modules_frame = ttk.LabelFrame(top, text="Enabled Modules", padding=10)
        modules_frame.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 10))
        for idx, key in enumerate(MODULE_KEYS):
            var = tk.BooleanVar(value=bool(enabled_modules.get(key, False)))
            self._module_vars[key] = var
            ttk.Checkbutton(modules_frame, text=MODULE_LABELS[key], variable=var).grid(
                row=idx // 2, column=idx % 2, sticky="w", padx=6, pady=2
            )
        modules_frame.columnconfigure(0, weight=1)
        modules_frame.columnconfigure(1, weight=1)

        settings_frame = ttk.LabelFrame(top, text="Settings", padding=10)
        settings_frame.grid(row=2, column=0, columnspan=2, sticky="nsew")
        top.rowconfigure(2, weight=1)

        self.setting_vars = {
            "default_currency": tk.StringVar(value=_safe_text(settings.get("default_currency") or "AED")),
            "vat_enabled": tk.BooleanVar(value=bool(settings.get("vat_enabled", True))),
            "vat_rate": tk.StringVar(value=str(settings.get("vat_rate", 5))),
            "invoice_prefix": tk.StringVar(value=_safe_text(settings.get("invoice_prefix") or "INV")),
            "purchase_prefix": tk.StringVar(value=_safe_text(settings.get("purchase_prefix") or "PUR")),
            "warehouse_prefix": tk.StringVar(value=_safe_text(settings.get("warehouse_prefix") or "WH")),
            "low_stock_alert_threshold": tk.StringVar(value=str(settings.get("low_stock_alert_threshold", 10))),
            "expiry_alert_days": tk.StringVar(value=str(settings.get("expiry_alert_days", 30))),
            "temperature_monitoring_enabled": tk.BooleanVar(value=bool(settings.get("temperature_monitoring_enabled", True))),
            "batch_tracking_required": tk.BooleanVar(value=bool(settings.get("batch_tracking_required", True))),
            "barcode_required": tk.BooleanVar(value=bool(settings.get("barcode_required", False))),
            "multi_warehouse_enabled": tk.BooleanVar(value=bool(settings.get("multi_warehouse_enabled", False))),
            "approval_workflow_enabled": tk.BooleanVar(value=bool(settings.get("approval_workflow_enabled", True))),
            "role_based_access_enabled": tk.BooleanVar(value=bool(settings.get("role_based_access_enabled", True))),
        }
        self.dashboard_widget_vars = {
            key: tk.BooleanVar(value=bool(dashboard_widgets.get(key, True)))
            for key in DASHBOARD_WIDGET_KEYS
        }
        rows = [
            ("Default Currency", "default_currency"),
            ("VAT Rate", "vat_rate"),
            ("Invoice Prefix", "invoice_prefix"),
            ("Purchase Prefix", "purchase_prefix"),
            ("Warehouse Prefix", "warehouse_prefix"),
            ("Low Stock Alert Threshold", "low_stock_alert_threshold"),
            ("Expiry Alert Days", "expiry_alert_days"),
        ]
        for idx, (label, key) in enumerate(rows):
            ttk.Label(settings_frame, text=label).grid(row=idx, column=0, sticky="w", pady=3, padx=(0, 10))
            ttk.Entry(settings_frame, textvariable=self.setting_vars[key]).grid(row=idx, column=1, sticky="ew", pady=3)
        check_rows = [
            ("VAT Enabled", "vat_enabled"),
            ("Temperature Monitoring Enabled", "temperature_monitoring_enabled"),
            ("Batch Tracking Required", "batch_tracking_required"),
            ("Barcode Required", "barcode_required"),
            ("Multi-Warehouse Enabled", "multi_warehouse_enabled"),
            ("Approval Workflow Enabled", "approval_workflow_enabled"),
            ("Role-Based Access Enabled", "role_based_access_enabled"),
        ]
        for idx, (label, key) in enumerate(check_rows):
            ttk.Checkbutton(settings_frame, text=label, variable=self.setting_vars[key]).grid(
                row=idx, column=2, sticky="w", padx=(20, 0), pady=3
            )
        settings_frame.columnconfigure(1, weight=1)

        widget_frame = ttk.LabelFrame(top, text="Dashboard Widgets", padding=10)
        widget_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(10, 0))
        for idx, key in enumerate(DASHBOARD_WIDGET_KEYS):
            ttk.Checkbutton(
                widget_frame,
                text=DASHBOARD_WIDGET_LABELS[key],
                variable=self.dashboard_widget_vars[key],
            ).grid(row=idx // 2, column=idx % 2, sticky="w", padx=6, pady=2)
        widget_frame.columnconfigure(0, weight=1)
        widget_frame.columnconfigure(1, weight=1)

    def _apply_company_type_preset(self, _event=None):
        defaults = default_enabled_modules(self.company_type_var.get())
        for key, var in self._module_vars.items():
            var.set(bool(defaults.get(key, False)))
        cfg = default_company_configuration(self.company_type_var.get())
        for key, var in self.setting_vars.items():
            if key in cfg["settings"]:
                value = cfg["settings"][key]
                if isinstance(var, tk.BooleanVar):
                    var.set(bool(value))
                else:
                    var.set(str(value))
        for key, var in self.dashboard_widget_vars.items():
            var.set(bool((cfg["settings"].get("dashboard_widgets") or {}).get(key, True)))

    def get_configuration(self):
        settings = {}
        for key, var in self.setting_vars.items():
            value = var.get()
            if key in {"vat_rate"}:
                settings[key] = _safe_float(value, 0.0)
            elif key in {"low_stock_alert_threshold", "expiry_alert_days"}:
                settings[key] = _safe_int(value, 0)
            elif isinstance(var, tk.BooleanVar):
                settings[key] = bool(value)
            else:
                settings[key] = _safe_text(value)
        settings["dashboard_widgets"] = {
            key: bool(var.get()) for key, var in self.dashboard_widget_vars.items()
        }
        return {
            "company_type": _safe_text(self.company_type_var.get()) or "Full ERP",
            "enabled_modules": {key: bool(var.get()) for key, var in self._module_vars.items()},
            "settings": settings,
        }


class CompanySettingsDialog:
    def __init__(self, parent, company_folder, company_name, current_config, on_save):
        self.company_folder = company_folder
        self.company_name = company_name
        self.on_save = on_save
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Company Settings")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        self.dialog.grab_set()
        frame = ttk.Frame(self.dialog, padding=14)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"Company Settings - {company_name}", font=("Helvetica", 16, "bold")).pack(anchor="w", pady=(0, 10))
        scroll_wrap = ScrollableSettingsFrame(frame)
        scroll_wrap.pack(fill="both", expand=True)
        self.config_frame = CompanyConfigurationFrame(scroll_wrap.scrollable_frame, initial_config=current_config)
        self.config_frame.pack(fill="both", expand=True)
        btns = ttk.Frame(frame)
        btns.pack(fill="x", pady=(10, 0))
        ttk.Button(btns, text="Save", command=self._save).pack(side="left")
        ttk.Button(btns, text="Cancel", command=self.dialog.destroy).pack(side="left", padx=8)
        _fit_dialog(self.dialog, 900, 720)

    def _save(self):
        payload = self.config_frame.get_configuration()
        self.on_save(payload)
        self.dialog.destroy()


class WarehouseManagementDialog:
    def __init__(self, parent, manager, company_name="", user_role="Staff"):
        self.parent = parent
        self.manager = manager
        self.company_name = company_name
        self.user_role = user_role
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Warehouse Management")
        self.dialog.transient(parent)
        self.dialog.resizable(True, True)
        self.dialog.grab_set()
        self._build_ui()
        self.refresh_all()
        _fit_dialog(self.dialog, 1180, 760)

    def _build_ui(self):
        root = ttk.Frame(self.dialog, padding=12)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="Warehouse Management", font=("Helvetica", 18, "bold")).pack(anchor="w")
        ttk.Label(root, text="Warehouses, locations, GRN, transfers, adjustments, returns, and stock count").pack(anchor="w", pady=(2, 10))
        notebook = ttk.Notebook(root)
        notebook.pack(fill="both", expand=True)
        self.tabs = {}
        for key, title in (
            ("warehouses", "Warehouses"),
            ("locations", "Locations"),
            ("receipts", "Goods Receipts"),
            ("outbounds", "Outbound"),
            ("transfers", "Transfers"),
            ("adjustments", "Adjustments"),
            ("counts", "Stock Count"),
        ):
            frame = ttk.Frame(notebook, padding=8)
            notebook.add(frame, text=title)
            self.tabs[key] = frame
        self._build_warehouses_tab()
        self._build_locations_tab()
        self._build_receipts_tab()
        self._build_outbounds_tab()
        self._build_transfers_tab()
        self._build_adjustments_tab()
        self._build_counts_tab()

    def _tree_with_scroll(self, parent, columns):
        frame = ttk.Frame(parent)
        frame.pack(fill="both", expand=True)
        tree = ttk.Treeview(frame, columns=columns, show="headings", height=14)
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=140, stretch=True)
        yscroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        return tree

    def _build_warehouses_tab(self):
        top = ttk.Frame(self.tabs["warehouses"])
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="Add Warehouse", command=self._add_warehouse).pack(side="left")
        ttk.Button(top, text="Edit Warehouse", command=self._edit_warehouse).pack(side="left", padx=6)
        ttk.Button(top, text="Delete Warehouse", command=self._delete_warehouse).pack(side="left", padx=6)
        ttk.Button(top, text="Archive Warehouse", command=self._archive_warehouse).pack(side="left", padx=6)
        self.warehouse_tree = self._tree_with_scroll(
            self.tabs["warehouses"],
            ("Code", "Name", "Type", "Status", "Temperature Controlled", "Contact"),
        )

    def _build_locations_tab(self):
        top = ttk.Frame(self.tabs["locations"])
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="Add Location", command=self._add_location).pack(side="left")
        ttk.Button(top, text="Edit Location", command=self._edit_location).pack(side="left", padx=6)
        ttk.Button(top, text="Delete Location", command=self._delete_location).pack(side="left", padx=6)
        self.location_tree = self._tree_with_scroll(
            self.tabs["locations"],
            ("Code", "Warehouse", "Zone", "Aisle", "Rack", "Shelf", "Company", "Bin", "Capacity", "Status", "Product"),
        )

    def _build_receipts_tab(self):
        top = ttk.Frame(self.tabs["receipts"])
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="New GRN", command=self._new_receipt).pack(side="left")
        ttk.Button(top, text="Edit GRN", command=self._edit_receipt).pack(side="left", padx=6)
        ttk.Button(top, text="Delete GRN", command=self._delete_receipt).pack(side="left", padx=6)
        self.receipt_tree = self._tree_with_scroll(
            self.tabs["receipts"],
            ("GRN", "Company", "Supplier", "Purchase Ref", "Product", "Batch", "Qty", "Warehouse", "Shelf/Bin", "Date", "Status"),
        )

    def _build_transfers_tab(self):
        top = ttk.Frame(self.tabs["transfers"])
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="Request Transfer", command=self._request_transfer).pack(side="left")
        ttk.Button(top, text="Edit Transfer", command=self._edit_transfer).pack(side="left", padx=6)
        ttk.Button(top, text="Delete Transfer", command=self._delete_transfer).pack(side="left", padx=6)
        ttk.Button(top, text="Approve Transfer", command=self._approve_transfer).pack(side="left", padx=6)
        ttk.Button(top, text="Receive Transfer", command=self._receive_transfer).pack(side="left", padx=6)
        ttk.Button(top, text="Goods Transfer PDF", command=self._export_transfer_pdf).pack(side="left", padx=6)
        self.transfer_tree = self._tree_with_scroll(
            self.tabs["transfers"],
            ("Transfer ID", "Type", "Item", "Batch", "Qty", "From", "To", "Status", "Requested By"),
        )

    def _build_outbounds_tab(self):
        self.outbound_filter_vars = {
            "date_from": tk.StringVar(),
            "date_to": tk.StringVar(),
            "client_owner": tk.StringVar(),
            "destination": tk.StringVar(),
            "warehouse_code": tk.StringVar(),
            "product_name": tk.StringVar(),
            "sku": tk.StringVar(),
            "batch_number": tk.StringVar(),
            "expiry_date": tk.StringVar(),
            "status": tk.StringVar(),
            "requested_by": tk.StringVar(),
            "dispatched_by": tk.StringVar(),
        }
        self.outbound_filter_widgets = {}
        filter_wrap = ttk.LabelFrame(self.tabs["outbounds"], text="Filters")
        filter_wrap.pack(fill="x", pady=(0, 8))
        choices = self._outbound_filter_choices()
        warehouse_codes = choices.get("warehouse_code") or [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        filter_specs = [
            ("Date From", "date_from", "entry", []),
            ("Date To", "date_to", "entry", []),
            ("Client", "client_owner", "combo", choices.get("client_owner") or []),
            ("Destination", "destination", "combo", choices.get("destination") or []),
            ("Warehouse", "warehouse_code", "combo", warehouse_codes),
            ("Product", "product_name", "combo", choices.get("product_name") or []),
            ("SKU", "sku", "combo", choices.get("sku") or []),
            ("Batch", "batch_number", "combo", choices.get("batch_number") or []),
            ("Expiry", "expiry_date", "entry", []),
            ("Status", "status", "combo", [""] + OUTBOUND_STATUSES),
            ("Requested By", "requested_by", "combo", choices.get("requested_by") or []),
            ("Dispatched By", "dispatched_by", "combo", choices.get("dispatched_by") or []),
        ]
        for idx, (label, key, kind, values) in enumerate(filter_specs):
            row = idx // 4
            col = (idx % 4) * 2
            ttk.Label(filter_wrap, text=label).grid(row=row, column=col, sticky="w", padx=(6, 6), pady=4)
            if kind == "combo":
                widget = ttk.Combobox(filter_wrap, textvariable=self.outbound_filter_vars[key], values=values, state="readonly")
                widget.grid(row=row, column=col + 1, sticky="ew", padx=(0, 12), pady=4)
                self.outbound_filter_widgets[key] = widget
                widget.bind("<<ComboboxSelected>>", self._on_outbound_filter_changed)
            else:
                ttk.Entry(filter_wrap, textvariable=self.outbound_filter_vars[key]).grid(row=row, column=col + 1, sticky="ew", padx=(0, 12), pady=4)
        for col in range(1, 8, 2):
            filter_wrap.columnconfigure(col, weight=1)
        filter_btns = ttk.Frame(filter_wrap)
        filter_btns.grid(row=3, column=0, columnspan=8, sticky="e", padx=6, pady=(4, 6))
        ttk.Button(filter_btns, text="Apply Filters", command=self._populate_outbounds).pack(side="left")
        ttk.Button(filter_btns, text="Reset Filters", command=self._reset_outbound_filters).pack(side="left", padx=6)
        self._refresh_outbound_filter_options()

        action_top = ttk.Frame(self.tabs["outbounds"])
        action_top.pack(fill="x", pady=(0, 6))
        for text, cmd in (
            ("New Outbound", self._new_outbound),
            ("Edit Outbound", self._edit_outbound),
            ("Delete Outbound", self._delete_outbound),
            ("Submit Approval", self._submit_outbound),
            ("Approve", self._approve_outbound),
            ("Reserve", self._reserve_outbound),
            ("Start Picking", self._start_outbound_picking),
            ("Confirm Picked", self._confirm_outbound_picked),
            ("Start Packing", self._start_outbound_packing),
            ("Confirm Packed", self._confirm_outbound_packed),
            ("Ready", self._mark_outbound_ready),
            ("Dispatch", self._dispatch_outbound),
            ("Deliver", self._deliver_outbound),
            ("Cancel", self._cancel_outbound),
            ("Return", self._return_outbound),
        ):
            ttk.Button(action_top, text=text, command=cmd).pack(side="left", padx=(0, 4))

        doc_top = ttk.Frame(self.tabs["outbounds"])
        doc_top.pack(fill="x", pady=(0, 8))
        for text, cmd in (
            ("Pick List PDF", self._export_pick_list_pdf),
            ("Packing List PDF", self._export_packing_list_pdf),
            ("Delivery Note PDF", self._export_outbound_delivery_pdf),
            ("Report PDF", self._export_outbound_report_pdf),
            ("Report Excel", self._export_outbound_report_excel),
            ("Report CSV", self._export_outbound_report_csv),
            ("Print Preview", self._print_preview_outbound_report),
        ):
            ttk.Button(doc_top, text=text, command=cmd).pack(side="left", padx=(0, 4))

        self.outbound_tree = self._tree_with_scroll(
            self.tabs["outbounds"],
            ("Outbound No.", "Date", "Client", "Destination", "Warehouse", "Status", "Requested By", "Approved By", "Picked By", "Packed By", "Dispatched By", "Received By"),
        )
        self.outbound_tree.bind("<<TreeviewSelect>>", self._on_outbound_selected)
        item_label = ttk.Label(self.tabs["outbounds"], text="Outbound Items", font=("Helvetica", 11, "bold"))
        item_label.pack(anchor="w", pady=(8, 4))
        self.outbound_item_tree = self._tree_with_scroll(
            self.tabs["outbounds"],
            ("Item code", "Product name", "Brand", "SKU", "Batch number", "Lot number", "Manufacturing date", "Expiry date", "Unit/pack size", "Warehouse", "Bin location", "Available quantity", "Requested quantity", "Reserved quantity", "Picked quantity", "Packed quantity", "Dispatched quantity", "Status"),
        )

    def _build_adjustments_tab(self):
        top = ttk.Frame(self.tabs["adjustments"])
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="New Adjustment", command=self._new_adjustment).pack(side="left")
        ttk.Button(top, text="Edit Adjustment", command=self._edit_adjustment).pack(side="left", padx=6)
        ttk.Button(top, text="Delete Adjustment", command=self._delete_adjustment).pack(side="left", padx=6)
        ttk.Button(top, text="Approve Adjustment", command=self._approve_adjustment).pack(side="left", padx=6)
        self.adjustment_tree = self._tree_with_scroll(
            self.tabs["adjustments"],
            ("Adjustment ID", "Item", "Qty Change", "Type", "Warehouse", "Reason", "Status"),
        )

    def _build_counts_tab(self):
        top = ttk.Frame(self.tabs["counts"])
        top.pack(fill="x", pady=(0, 8))
        ttk.Button(top, text="New Count", command=self._new_count).pack(side="left")
        ttk.Button(top, text="Edit Count", command=self._edit_count).pack(side="left", padx=6)
        ttk.Button(top, text="Delete Count", command=self._delete_count).pack(side="left", padx=6)
        ttk.Button(top, text="Approve Count", command=self._approve_count).pack(side="left", padx=6)
        self.count_tree = self._tree_with_scroll(
            self.tabs["counts"],
            ("Count ID", "Item", "Warehouse", "System Qty", "Actual Qty", "Variance", "Status"),
        )

    def _require_admin(self):
        if _safe_text(self.user_role) != "Admin":
            messagebox.showerror("Error", "Admin approval is required for this action")
            return False
        return True

    def refresh_all(self):
        self._populate_warehouses()
        self._populate_locations()
        self._populate_receipts()
        self._populate_outbounds()
        self._populate_transfers()
        self._populate_adjustments()
        self._populate_counts()

    def _populate_warehouses(self):
        self._clear_tree(self.warehouse_tree)
        for row in self.manager.list_warehouses():
            self.warehouse_tree.insert("", "end", values=(
                row.get("warehouse_code"), row.get("warehouse_name"), row.get("warehouse_type"),
                row.get("status"), "Yes" if row.get("temperature_controlled") else "No",
                row.get("contact_person"),
            ))

    def _populate_locations(self):
        self._clear_tree(self.location_tree)
        for row in self.manager.list_locations():
            self.location_tree.insert("", "end", values=(
                row.get("location_code"), row.get("warehouse_code"), row.get("zone"),
                row.get("aisle"), row.get("rack"), row.get("shelf"),
                row.get("owner_company") or row.get("client_name"), row.get("bin"),
                row.get("capacity"), row.get("status"), row.get("product_code"),
            ))

    def _populate_receipts(self):
        self._clear_tree(self.receipt_tree)
        for row in reversed(self.manager.list_goods_receipts()):
            self.receipt_tree.insert("", "end", values=(
                row.get("grn_id"), row.get("client_name") or row.get("owner_company"), row.get("supplier"),
                row.get("purchase_invoice_reference") or row.get("purchase_id"),
                row.get("product_name"), row.get("batch_number"), row.get("quantity_received"),
                row.get("warehouse_code"), row.get("bin_location"), row.get("received_date"), row.get("status"),
            ))

    def _populate_transfers(self):
        self._clear_tree(self.transfer_tree)
        for row in reversed(self.manager.list_transfers()):
            destination = _safe_text(row.get("destination_client")) if _safe_text(row.get("transfer_type")) == "Client Transfer" else f"{row.get('warehouse_to')} / {row.get('bin_to')}"
            self.transfer_tree.insert("", "end", values=(
                row.get("transfer_id"), row.get("transfer_type") or "Internal Transfer", row.get("product_name") or row.get("item_id"),
                row.get("batch_number"), row.get("quantity"),
                f"{row.get('warehouse_from')} / {row.get('bin_from')}",
                destination,
                row.get("status"), row.get("requested_by"),
            ))

    def _populate_outbounds(self):
        self._clear_tree(self.outbound_tree)
        filters = {key: var.get() for key, var in getattr(self, "outbound_filter_vars", {}).items()}
        for row in reversed(self.manager.list_outbounds()):
            probe_client = _safe_text(filters.get("client_owner")).lower()
            probe_dest = _safe_text(filters.get("destination")).lower()
            probe_wh = _safe_text(filters.get("warehouse_code")).lower()
            probe_status = _safe_text(filters.get("status")).lower()
            probe_req = _safe_text(filters.get("requested_by")).lower()
            probe_dis = _safe_text(filters.get("dispatched_by")).lower()
            probe_from = self.manager._parse_iso_date(filters.get("date_from"))
            probe_to = self.manager._parse_iso_date(filters.get("date_to"))
            row_date = self.manager._parse_iso_date(row.get("date"))
            if probe_from and row_date and row_date < probe_from:
                continue
            if probe_to and row_date and row_date > probe_to:
                continue
            if probe_client and probe_client not in _safe_text(row.get("client_owner")).lower():
                continue
            if probe_dest and probe_dest not in _safe_text(row.get("destination")).lower():
                continue
            if probe_wh and probe_wh not in _safe_text(row.get("warehouse_code")).lower():
                continue
            if probe_status and probe_status != _safe_text(row.get("status")).lower():
                continue
            if probe_req and probe_req not in _safe_text(row.get("requested_by")).lower():
                continue
            if probe_dis and probe_dis not in _safe_text(row.get("dispatched_by")).lower():
                continue
            line = (row.get("items") or [{}])[0]
            if _safe_text(filters.get("product_name")).lower() and _safe_text(filters.get("product_name")).lower() not in _safe_text(line.get("product_name")).lower():
                continue
            if _safe_text(filters.get("sku")).lower() and _safe_text(filters.get("sku")).lower() not in _safe_text(line.get("sku")).lower():
                continue
            if _safe_text(filters.get("batch_number")).lower() and _safe_text(filters.get("batch_number")).lower() not in _safe_text(line.get("batch_number")).lower():
                continue
            if _safe_text(filters.get("expiry_date")).lower() and _safe_text(filters.get("expiry_date")).lower() not in _safe_text(line.get("expiry_date")).lower():
                continue
            self.outbound_tree.insert("", "end", values=(
                row.get("outbound_no"), row.get("date"), row.get("client_owner"), row.get("destination"),
                row.get("warehouse_code"), row.get("status"), row.get("requested_by"), row.get("approved_by"),
                row.get("picked_by"), row.get("packed_by"), row.get("dispatched_by"), row.get("received_by"),
            ))
        self._clear_tree(self.outbound_item_tree)

    def _on_outbound_selected(self, _event=None):
        self._clear_tree(self.outbound_item_tree)
        outbound_no = self._selected_value(self.outbound_tree, 0)
        if not outbound_no:
            return
        row = self.manager.get_outbound(outbound_no)
        if not row:
            return
        for line in row.get("items") or []:
            self.outbound_item_tree.insert("", "end", values=(
                line.get("item_id"), line.get("product_name"), line.get("brand"), line.get("sku"),
                line.get("batch_number"), line.get("lot_number"), line.get("manufacture_date"), line.get("expiry_date"),
                line.get("unit_pack_size"), line.get("warehouse_code"), line.get("bin_location"),
                line.get("available_quantity"), line.get("requested_quantity"), line.get("reserved_quantity"),
                line.get("picked_quantity"), line.get("packed_quantity"), line.get("dispatched_quantity"), line.get("status"),
            ))

    def _reset_outbound_filters(self):
        for var in self.outbound_filter_vars.values():
            var.set("")
        self._refresh_outbound_filter_options()
        self._populate_outbounds()

    def _on_outbound_filter_changed(self, _event=None):
        self._refresh_outbound_filter_options()
        self._populate_outbounds()

    def _outbound_filter_snapshot(self):
        return {key: var.get() for key, var in getattr(self, "outbound_filter_vars", {}).items()}

    def _outbound_filter_choices(self):
        invm = self.manager._get_inventory_manager()
        outbounds = self.manager.list_outbounds()
        clients = set()
        destinations = set()
        warehouses = set()
        requested_by = set()
        dispatched_by = set()
        for item in getattr(invm, "items", []) or []:
            clients.add(_safe_text(getattr(item, "client_name", "")))
            warehouses.add(_safe_text(getattr(item, "warehouse_code", "")))
        for row in outbounds:
            clients.add(_safe_text(row.get("client_owner")))
            destinations.add(_safe_text(row.get("destination")))
            warehouses.add(_safe_text(row.get("warehouse_code")))
            requested_by.add(_safe_text(row.get("requested_by")))
            dispatched_by.add(_safe_text(row.get("dispatched_by")))
        wh_codes = [_safe_text(x.get("warehouse_code")) for x in self.manager.list_warehouses()]
        warehouses.update([_safe_text(x) for x in wh_codes if _safe_text(x)])

        return {
            "client_owner": [""] + sorted({x for x in clients if x}),
            "destination": [""] + sorted({x for x in destinations if x}),
            "warehouse_code": [""] + sorted({x for x in warehouses if x}),
            "requested_by": [""] + sorted({x for x in requested_by if x}),
            "dispatched_by": [""] + sorted({x for x in dispatched_by if x}),
        }

    def _outbound_inventory_choices(self, filters):
        invm = self.manager._get_inventory_manager()
        client_probe = _safe_text(filters.get("client_owner")).lower()
        wh_probe = _safe_text(filters.get("warehouse_code")).lower()
        product_probe = _safe_text(filters.get("product_name")).lower()
        sku_probe = _safe_text(filters.get("sku")).lower()
        batch_probe = _safe_text(filters.get("batch_number")).lower()
        rows = []
        for item in getattr(invm, "items", []) or []:
            if client_probe and _safe_text(getattr(item, "client_name", "")).lower() != client_probe:
                continue
            if wh_probe and _safe_text(getattr(item, "warehouse_code", "")).lower() != wh_probe:
                continue
            if product_probe and _safe_text(getattr(item, "name", "")).lower() != product_probe:
                continue
            if sku_probe and (_safe_text(getattr(item, "sku", "")) or _safe_text(getattr(item, "item_id", ""))).lower() != sku_probe:
                continue
            if batch_probe and _safe_text(getattr(item, "batch_number", "")).lower() != batch_probe:
                continue
            rows.append(item)
        return rows

    def _refresh_outbound_filter_options(self):
        base = self._outbound_filter_choices()
        current = self._outbound_filter_snapshot()
        inv_items = self._outbound_inventory_choices(current)
        products = { _safe_text(getattr(it, "name", "")) for it in inv_items }
        skus = { _safe_text(getattr(it, "sku", "")) or _safe_text(getattr(it, "item_id", "")) for it in inv_items }
        batches = { _safe_text(getattr(it, "batch_number", "")) for it in inv_items }
        base["product_name"] = [""] + sorted({x for x in products if x})
        base["sku"] = [""] + sorted({x for x in skus if x})
        base["batch_number"] = [""] + sorted({x for x in batches if x})

        for key, widget in (getattr(self, "outbound_filter_widgets", {}) or {}).items():
            if key in base and widget:
                widget.configure(values=base[key])
                if _safe_text(self.outbound_filter_vars[key].get()) and self.outbound_filter_vars[key].get() not in base[key]:
                    self.outbound_filter_vars[key].set("")

    def _populate_adjustments(self):
        self._clear_tree(self.adjustment_tree)
        for row in reversed(self.manager.list_adjustments()):
            self.adjustment_tree.insert("", "end", values=(
                row.get("adjustment_id"), row.get("product_name") or row.get("item_id"),
                row.get("quantity_change"), row.get("adjustment_type"),
                row.get("warehouse_code"), row.get("reason"), row.get("status"),
            ))

    def _populate_counts(self):
        self._clear_tree(self.count_tree)
        for row in reversed(self.manager.list_counts()):
            self.count_tree.insert("", "end", values=(
                row.get("count_id"), row.get("product_name") or row.get("item_id"),
                row.get("warehouse_code"), row.get("system_quantity"),
                row.get("actual_quantity"), row.get("variance"), row.get("status"),
            ))

    def _clear_tree(self, tree):
        for child in tree.get_children():
            tree.delete(child)

    def _selected_value(self, tree, index=0):
        sel = tree.selection()
        if not sel:
            return ""
        values = tree.item(sel[0]).get("values") or []
        return values[index] if index < len(values) else ""

    def _warehouse_form(self, initial=None):
        rec = initial or {}
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Warehouse")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frame = ttk.Frame(dlg, padding=12)
        frame.pack(fill="both", expand=True)
        vars_ = {
            "warehouse_code": tk.StringVar(value=_safe_text(rec.get("warehouse_code"))),
            "warehouse_name": tk.StringVar(value=_safe_text(rec.get("warehouse_name"))),
            "address": tk.StringVar(value=_safe_text(rec.get("address"))),
            "emirate_country": tk.StringVar(value=_safe_text(rec.get("emirate_country"))),
            "contact_person": tk.StringVar(value=_safe_text(rec.get("contact_person"))),
            "phone": tk.StringVar(value=_safe_text(rec.get("phone"))),
            "email": tk.StringVar(value=_safe_text(rec.get("email"))),
            "status": tk.StringVar(value=_safe_text(rec.get("status") or "Active")),
            "temperature_controlled": tk.BooleanVar(value=bool(rec.get("temperature_controlled"))),
            "warehouse_type": tk.StringVar(value=_safe_text(rec.get("warehouse_type") or "Main")),
        }
        rows = [
            ("Warehouse Code", "warehouse_code"),
            ("Warehouse Name", "warehouse_name"),
            ("Address", "address"),
            ("Emirate/Country", "emirate_country"),
            ("Contact Person", "contact_person"),
            ("Phone", "phone"),
            ("Email", "email"),
        ]
        for idx, (label, key) in enumerate(rows):
            ttk.Label(frame, text=label).grid(row=idx, column=0, sticky="w", pady=4, padx=(0, 8))
            ttk.Entry(frame, textvariable=vars_[key]).grid(row=idx, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Status").grid(row=7, column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Combobox(frame, textvariable=vars_["status"], values=["Active", "Inactive", "Archived"], state="readonly").grid(row=7, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Warehouse Type").grid(row=8, column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Combobox(frame, textvariable=vars_["warehouse_type"], values=WAREHOUSE_TYPES, state="readonly").grid(row=8, column=1, sticky="ew", pady=4)
        ttk.Checkbutton(frame, text="Temperature Controlled", variable=vars_["temperature_controlled"]).grid(row=9, column=1, sticky="w", pady=4)
        frame.columnconfigure(1, weight=1)
        result = {}

        def save():
            result.update({key: var.get() for key, var in vars_.items()})
            dlg.destroy()

        def clear():
            for var in vars_.values():
                try:
                    if isinstance(var, tk.BooleanVar):
                        var.set(False)
                    else:
                        var.set("")
                except Exception:
                    pass

        btns = ttk.Frame(frame)
        btns.grid(row=10, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Clear", command=clear).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Save", command=save).pack(side="left")
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        _fit_dialog(dlg, 560, 520)
        self.dialog.wait_window(dlg)
        return result

    def _location_form(self, initial=None):
        rec = initial or {}
        dlg = tk.Toplevel(self.dialog)
        dlg.title("Location")
        dlg.transient(self.dialog)
        dlg.grab_set()
        frame = ttk.Frame(dlg, padding=12)
        frame.pack(fill="both", expand=True)
        wh_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        vars_ = {
            "location_code": tk.StringVar(value=_safe_text(rec.get("location_code"))),
            "warehouse_code": tk.StringVar(value=_safe_text(rec.get("warehouse_code") or (wh_codes[0] if wh_codes else ""))),
            "zone": tk.StringVar(value=_safe_text(rec.get("zone"))),
            "aisle": tk.StringVar(value=_safe_text(rec.get("aisle"))),
            "rack": tk.StringVar(value=_safe_text(rec.get("rack"))),
            "shelf": tk.StringVar(value=_safe_text(rec.get("shelf"))),
            "owner_company": tk.StringVar(value=_safe_text(rec.get("owner_company") or rec.get("client_name"))),
            "bin": tk.StringVar(value=_safe_text(rec.get("bin"))),
            "capacity": tk.StringVar(value=str(rec.get("capacity") or "")),
            "status": tk.StringVar(value=_safe_text(rec.get("status") or "Active")),
            "product_code": tk.StringVar(value=_safe_text(rec.get("product_code"))),
        }
        labels = [
            ("Location Code", "location_code"),
            ("Warehouse", "warehouse_code"),
            ("Zone", "zone"),
            ("Aisle", "aisle"),
            ("Rack", "rack"),
            ("Shelf", "shelf"),
            ("Owner Company", "owner_company"),
            ("Bin", "bin"),
            ("Capacity", "capacity"),
            ("Assigned Product", "product_code"),
        ]
        for idx, (label, key) in enumerate(labels):
            ttk.Label(frame, text=label).grid(row=idx, column=0, sticky="w", pady=4, padx=(0, 8))
            if key == "warehouse_code":
                ttk.Combobox(frame, textvariable=vars_[key], values=wh_codes, state="readonly").grid(row=idx, column=1, sticky="ew", pady=4)
            else:
                ttk.Entry(frame, textvariable=vars_[key]).grid(row=idx, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="Status").grid(row=10, column=0, sticky="w", pady=4, padx=(0, 8))
        ttk.Combobox(frame, textvariable=vars_["status"], values=LOCATION_STATUSES, state="readonly").grid(row=10, column=1, sticky="ew", pady=4)
        frame.columnconfigure(1, weight=1)
        result = {}

        def save():
            result.update({key: var.get() for key, var in vars_.items()})
            dlg.destroy()

        def clear():
            for var in vars_.values():
                try:
                    if isinstance(var, tk.BooleanVar):
                        var.set(False)
                    else:
                        var.set("")
                except Exception:
                    pass

        btns = ttk.Frame(frame)
        btns.grid(row=11, column=0, columnspan=2, sticky="e", pady=(10, 0))
        ttk.Button(btns, text="Clear", command=clear).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Save", command=save).pack(side="left")
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        _fit_dialog(dlg, 560, 540)
        self.dialog.wait_window(dlg)
        return result

    def _record_form(self, title, fields, min_size=(640, 520)):
        dlg = tk.Toplevel(self.dialog)
        dlg.title(title)
        dlg.transient(self.dialog)
        dlg.grab_set()
        frame = ttk.Frame(dlg, padding=12)
        frame.pack(fill="both", expand=True)
        vars_ = {}
        for idx, spec in enumerate(fields):
            key = spec["key"]
            kind = spec.get("kind", "entry")
            vars_[key] = spec["var"]
            ttk.Label(frame, text=spec["label"]).grid(row=idx, column=0, sticky="w", pady=4, padx=(0, 8))
            if kind == "combo":
                ttk.Combobox(frame, textvariable=vars_[key], values=spec.get("values", []), state="readonly").grid(row=idx, column=1, sticky="ew", pady=4)
            elif kind == "check":
                ttk.Checkbutton(frame, variable=vars_[key]).grid(row=idx, column=1, sticky="w", pady=4)
            else:
                ttk.Entry(frame, textvariable=vars_[key]).grid(row=idx, column=1, sticky="ew", pady=4)
        frame.columnconfigure(1, weight=1)
        result = {}

        def save():
            result.update({key: var.get() for key, var in vars_.items()})
            dlg.destroy()

        def clear():
            for key, var in vars_.items():
                spec = next((field for field in fields if field["key"] == key), {})
                kind = spec.get("kind", "entry")
                try:
                    if kind == "check":
                        var.set(False)
                    else:
                        var.set("")
                except Exception:
                    pass

        btns = ttk.Frame(frame)
        btns.grid(row=len(fields), column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(btns, text="Clear", command=clear).pack(side="left", padx=(0, 8))
        ttk.Button(btns, text="Save", command=save).pack(side="left")
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="left", padx=8)
        _fit_dialog(dlg, min_size[0], min_size[1])
        self.dialog.wait_window(dlg)
        return result

    def _inventory_item_choices(self):
        from inventory_system import InventoryManager
        invm = InventoryManager(self.manager.company_folder)
        choices = [f"{item.item_id} - {item.name}" for item in invm.items]
        return invm, choices

    def _selected_outbound_no(self):
        return self._selected_value(self.outbound_tree, 0)

    def _outbound_form(self, initial=None):
        rec = initial or {}
        line = (rec.get("items") or [{}])[0]
        invm, item_choices = self._inventory_item_choices()
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        current_item = invm.get_item(_safe_text(line.get("source_item_id") or line.get("item_id")))
        current_item_label = f"{current_item.item_id} - {current_item.name}" if current_item else (item_choices[0] if item_choices else "")
        fields = [
            {"label": "Date", "key": "date", "var": tk.StringVar(value=_safe_text(rec.get("date") or datetime.now().strftime("%Y-%m-%d")))},
            {"label": "Client / Owner", "key": "client_owner", "var": tk.StringVar(value=_safe_text(rec.get("client_owner") or line.get("client_owner")))},
            {"label": "Destination", "key": "destination", "var": tk.StringVar(value=_safe_text(rec.get("destination")))},
            {"label": "Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=_safe_text(rec.get("warehouse_code") or line.get("warehouse_code") or (warehouse_codes[0] if warehouse_codes else ""))), "kind": "combo", "values": warehouse_codes},
            {"label": "Bin Location", "key": "bin_location", "var": tk.StringVar(value=_safe_text(rec.get("bin_location") or line.get("bin_location")))},
            {"label": "Inventory Item", "key": "item_id", "var": tk.StringVar(value=current_item_label), "kind": "combo", "values": item_choices},
            {"label": "Product Search", "key": "product_name", "var": tk.StringVar(value=_safe_text(line.get("product_name")))},
            {"label": "SKU", "key": "sku", "var": tk.StringVar(value=_safe_text(line.get("sku")))},
            {"label": "Batch Number", "key": "batch_number", "var": tk.StringVar(value=_safe_text(line.get("batch_number")))},
            {"label": "Lot Number", "key": "lot_number", "var": tk.StringVar(value=_safe_text(line.get("lot_number")))},
            {"label": "Requested Quantity", "key": "requested_quantity", "var": tk.StringVar(value=str(line.get("requested_quantity") or ""))},
            {"label": "Admin Override Expired", "key": "allow_expired_override", "var": tk.BooleanVar(value=bool(rec.get("admin_override_expired"))), "kind": "check"},
            {"label": "Package Number", "key": "package_number", "var": tk.StringVar(value=_safe_text(rec.get("package_number")))},
            {"label": "Notes", "key": "notes", "var": tk.StringVar(value=_safe_text(rec.get("notes")))},
        ]
        result = self._record_form("Outbound Request", fields, min_size=(820, 760))
        if not result:
            return {}
        raw_item = _safe_text(result.get("item_id"))
        item_id = raw_item.split(" - ", 1)[0] if raw_item else ""
        if item_id:
            result["item_id"] = item_id
        return result

    def _simple_outbound_qty_form(self, title, qty_label="", include_package=False, include_received=False, include_return=False, include_quantity=True):
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        fields = []
        if include_quantity:
            fields.append({"label": qty_label, "key": "quantity", "var": tk.StringVar()})
        if include_package:
            fields.append({"label": "Package Number", "key": "package_number", "var": tk.StringVar()})
        if include_received:
            fields.append({"label": "Received By", "key": "received_by", "var": tk.StringVar()})
        if include_return:
            fields.extend([
                {"label": "Return Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=warehouse_codes[0] if warehouse_codes else ""), "kind": "combo", "values": warehouse_codes},
                {"label": "Return Bin", "key": "bin_location", "var": tk.StringVar()},
                {"label": "Quarantine", "key": "quarantine", "var": tk.BooleanVar(value=False), "kind": "check"},
                {"label": "Notes", "key": "notes", "var": tk.StringVar()},
            ])
        return self._record_form(title, fields, min_size=(560, 420))

    def _new_outbound(self):
        result = self._outbound_form()
        if result:
            self.manager.create_outbound(result, user_name=self.user_role)
            self.refresh_all()

    def _edit_outbound(self):
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        rec = self.manager.get_outbound(outbound_no)
        if not rec:
            messagebox.showerror("Error", "Selected outbound request was not found")
            return
        result = self._outbound_form(rec)
        if result:
            self.manager.update_outbound(outbound_no, result, user_name=self.user_role)
            self.refresh_all()

    def _delete_outbound(self):
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete outbound request {outbound_no}?"):
            return
        try:
            self.manager.delete_outbound(outbound_no, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _submit_outbound(self):
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.submit_outbound_for_approval(outbound_no, user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _approve_outbound(self):
        if not self._require_admin():
            return
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.approve_outbound(outbound_no, user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _reserve_outbound(self):
        if not self._require_admin():
            return
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.reserve_outbound(outbound_no, user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _start_outbound_picking(self):
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.start_outbound_picking(outbound_no, user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _confirm_outbound_picked(self):
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        result = self._simple_outbound_qty_form("Confirm Picked Quantity", "Picked Quantity")
        if result:
            try:
                self.manager.confirm_outbound_picked(outbound_no, result.get("quantity"), user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _start_outbound_packing(self):
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.start_outbound_packing(outbound_no, user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _confirm_outbound_packed(self):
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        result = self._simple_outbound_qty_form("Confirm Packed Quantity", "Packed Quantity", include_package=True)
        if result:
            try:
                self.manager.confirm_outbound_packed(outbound_no, result.get("quantity"), package_number=result.get("package_number"), user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _mark_outbound_ready(self):
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.mark_outbound_ready_for_dispatch(outbound_no, user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _dispatch_outbound(self):
        if not self._require_admin():
            return
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        result = self._simple_outbound_qty_form("Dispatch Outbound", "Dispatched Quantity")
        if result:
            try:
                self.manager.dispatch_outbound(outbound_no, dispatched_quantity=result.get("quantity"), user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _deliver_outbound(self):
        if not self._require_admin():
            return
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        result = self._simple_outbound_qty_form("Mark Delivered", include_received=True, include_quantity=False)
        if result:
            try:
                self.manager.mark_outbound_delivered(outbound_no, received_by=result.get("received_by"), user_name=self.user_role)
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _cancel_outbound(self):
        if not self._require_admin():
            return
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        if not messagebox.askyesno("Confirm Cancel", f"Cancel outbound request {outbound_no}?"):
            return
        try:
            self.manager.cancel_outbound(outbound_no, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _return_outbound(self):
        if not self._require_admin():
            return
        outbound_no = self._selected_outbound_no()
        if not outbound_no:
            messagebox.showwarning("Warning", "Select an outbound request first")
            return
        result = self._simple_outbound_qty_form("Return Outbound Goods", "Return Quantity", include_return=True)
        if result:
            try:
                self.manager.return_outbound(
                    outbound_no,
                    result.get("quantity"),
                    warehouse_code=result.get("warehouse_code"),
                    bin_location=result.get("bin_location"),
                    quarantine=bool(result.get("quarantine")),
                    notes=result.get("notes"),
                    user_name=self.user_role,
                )
                self.refresh_all()
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _export_pick_list_pdf(self):
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.export_pick_list_pdf(outbound_no, open_after=True)
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _export_packing_list_pdf(self):
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.export_packing_list_pdf(outbound_no, open_after=True)
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _export_outbound_delivery_pdf(self):
        outbound_no = self._selected_outbound_no()
        if outbound_no:
            try:
                self.manager.export_outbound_delivery_note_pdf(outbound_no, open_after=True)
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _outbound_report_rows(self):
        filters = {key: var.get() for key, var in getattr(self, "outbound_filter_vars", {}).items()}
        return self.manager.get_outbound_report_rows(filters)

    def _export_outbound_report_pdf(self):
        try:
            self.manager._export_outbound_rows_pdf(self._outbound_report_rows(), "Outbound Report", "OutboundReport", open_after=True)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _export_outbound_report_excel(self):
        try:
            self.manager._export_outbound_rows_xlsx(self._outbound_report_rows(), "OutboundReport", open_after=True)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _export_outbound_report_csv(self):
        try:
            self.manager._export_outbound_rows_csv(self._outbound_report_rows(), "OutboundReport", open_after=True)
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _print_preview_outbound_report(self):
        self._export_outbound_report_pdf()

    def _add_warehouse(self):
        result = self._warehouse_form()
        if result:
            self.manager.save_warehouse(result, user_name=self.user_role)
            self.refresh_all()

    def _edit_warehouse(self):
        code = self._selected_value(self.warehouse_tree, 0)
        if not code:
            messagebox.showwarning("Warning", "Select a warehouse first")
            return
        rec = next((row for row in self.manager.list_warehouses(include_archived=True) if row.get("warehouse_code") == code), None)
        result = self._warehouse_form(rec)
        if result:
            self.manager.save_warehouse(result, user_name=self.user_role)
            self.refresh_all()

    def _archive_warehouse(self):
        code = self._selected_value(self.warehouse_tree, 0)
        if not code:
            messagebox.showwarning("Warning", "Select a warehouse first")
            return
        self.manager.archive_warehouse(code, user_name=self.user_role)
        self.refresh_all()

    def _delete_warehouse(self):
        code = self._selected_value(self.warehouse_tree, 0)
        if not code:
            messagebox.showwarning("Warning", "Select a warehouse first")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete warehouse {code}?"):
            return
        try:
            self.manager.delete_warehouse(code, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _add_location(self):
        result = self._location_form()
        if result:
            self.manager.save_location(result, user_name=self.user_role)
            self.refresh_all()

    def _edit_location(self):
        code = self._selected_value(self.location_tree, 0)
        if not code:
            messagebox.showwarning("Warning", "Select a location first")
            return
        rec = next((row for row in self.manager.list_locations() if row.get("location_code") == code), None)
        result = self._location_form(rec)
        if result:
            self.manager.save_location(result, user_name=self.user_role)
            self.refresh_all()

    def _delete_location(self):
        code = self._selected_value(self.location_tree, 0)
        if not code:
            messagebox.showwarning("Warning", "Select a location first")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete location {code}?"):
            return
        try:
            self.manager.delete_location(code, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _new_receipt(self):
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        fields = [
            {"label": "Company / Client", "key": "client_name", "var": tk.StringVar()},
            {"label": "Supplier", "key": "supplier", "var": tk.StringVar()},
            {"label": "Purchase ID / Reference", "key": "purchase_id", "var": tk.StringVar()},
            {"label": "Product Name", "key": "product_name", "var": tk.StringVar()},
            {"label": "SKU / Product Code", "key": "product_code", "var": tk.StringVar()},
            {"label": "Brand", "key": "brand", "var": tk.StringVar()},
            {"label": "Batch Number", "key": "batch_number", "var": tk.StringVar()},
            {"label": "Lot Number", "key": "lot_number", "var": tk.StringVar()},
            {"label": "Manufacture Date", "key": "manufacture_date", "var": tk.StringVar()},
            {"label": "Expiry Date", "key": "expiry_date", "var": tk.StringVar()},
            {"label": "Quantity Received", "key": "quantity_received", "var": tk.StringVar()},
            {"label": "Unit Cost", "key": "unit_cost", "var": tk.StringVar()},
            {"label": "Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=warehouse_codes[0] if warehouse_codes else ""), "kind": "combo", "values": warehouse_codes},
            {"label": "Shelf / Bin / Location Code", "key": "bin_location", "var": tk.StringVar()},
            {"label": "Notes", "key": "notes", "var": tk.StringVar()},
        ]
        result = self._record_form("New Goods Receipt", fields, min_size=(720, 700))
        if result:
            self.manager.receive_goods(result, user_name=self.user_role)
            self.refresh_all()

    def _edit_receipt(self):
        grn_id = self._selected_value(self.receipt_tree, 0)
        if not grn_id:
            messagebox.showwarning("Warning", "Select a goods receipt first")
            return
        rec = self.manager.get_goods_receipt(grn_id)
        if not rec:
            messagebox.showerror("Error", "Selected goods receipt could not be found")
            return
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        fields = [
            {"label": "Company / Client", "key": "client_name", "var": tk.StringVar(value=_safe_text(rec.get("client_name") or rec.get("owner_company")))},
            {"label": "Supplier", "key": "supplier", "var": tk.StringVar(value=_safe_text(rec.get("supplier")))},
            {"label": "Purchase Invoice Reference", "key": "purchase_invoice_reference", "var": tk.StringVar(value=_safe_text(rec.get("purchase_invoice_reference")))},
            {"label": "Purchase ID / Reference", "key": "purchase_id", "var": tk.StringVar(value=_safe_text(rec.get("purchase_id")))},
            {"label": "Product Name", "key": "product_name", "var": tk.StringVar(value=_safe_text(rec.get("product_name")))},
            {"label": "SKU / Product Code", "key": "product_code", "var": tk.StringVar()},
            {"label": "Brand", "key": "brand", "var": tk.StringVar()},
            {"label": "Batch Number", "key": "batch_number", "var": tk.StringVar(value=_safe_text(rec.get("batch_number")))},
            {"label": "Lot Number", "key": "lot_number", "var": tk.StringVar()},
            {"label": "Manufacture Date", "key": "manufacture_date", "var": tk.StringVar(value=_safe_text(rec.get("manufacture_date")))},
            {"label": "Expiry Date", "key": "expiry_date", "var": tk.StringVar(value=_safe_text(rec.get("expiry_date")))},
            {"label": "Quantity Received", "key": "quantity_received", "var": tk.StringVar(value=str(rec.get("quantity_received") or ""))},
            {"label": "Unit Cost", "key": "unit_cost", "var": tk.StringVar(value=str(rec.get("unit_cost") or ""))},
            {"label": "Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=_safe_text(rec.get("warehouse_code"))), "kind": "combo", "values": warehouse_codes},
            {"label": "Shelf / Bin / Location Code", "key": "bin_location", "var": tk.StringVar(value=_safe_text(rec.get("bin_location") or rec.get("location_code")))},
            {"label": "Notes", "key": "notes", "var": tk.StringVar(value=_safe_text(rec.get("notes")))},
        ]
        result = self._record_form("Edit Goods Receipt", fields, min_size=(720, 720))
        if result:
            self.manager.update_goods_receipt(grn_id, result, user_name=self.user_role)
            self.refresh_all()

    def _delete_receipt(self):
        grn_id = self._selected_value(self.receipt_tree, 0)
        if not grn_id:
            messagebox.showwarning("Warning", "Select a goods receipt first")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete goods receipt {grn_id}?\nThis will reverse the received stock."):
            return
        try:
            self.manager.delete_goods_receipt(grn_id, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _request_transfer(self):
        invm, item_names = self._inventory_item_choices()
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        fields = [
            {"label": "Inventory Item", "key": "item_id", "var": tk.StringVar(value=item_names[0] if item_names else ""), "kind": "combo", "values": item_names},
            {"label": "Transfer Type", "key": "transfer_type", "var": tk.StringVar(value="Client Transfer"), "kind": "combo", "values": ["Client Transfer", "Internal Transfer"]},
            {"label": "Quantity", "key": "quantity", "var": tk.StringVar()},
            {"label": "From Warehouse", "key": "warehouse_from", "var": tk.StringVar(value=warehouse_codes[0] if warehouse_codes else ""), "kind": "combo", "values": warehouse_codes},
            {"label": "From Bin", "key": "bin_from", "var": tk.StringVar()},
            {"label": "To Warehouse", "key": "warehouse_to", "var": tk.StringVar(value=warehouse_codes[0] if warehouse_codes else ""), "kind": "combo", "values": warehouse_codes},
            {"label": "To Bin", "key": "bin_to", "var": tk.StringVar()},
            {"label": "Client Name", "key": "destination_client", "var": tk.StringVar()},
            {"label": "Client Contact", "key": "destination_contact", "var": tk.StringVar()},
            {"label": "Client Address", "key": "destination_address", "var": tk.StringVar()},
            {"label": "Description", "key": "transfer_description", "var": tk.StringVar()},
            {"label": "Sign 1", "key": "sign_1", "var": tk.StringVar()},
            {"label": "Sign 2", "key": "sign_2", "var": tk.StringVar()},
            {"label": "Notes", "key": "notes", "var": tk.StringVar()},
        ]
        result = self._record_form("Request Transfer", fields, min_size=(760, 700))
        if result:
            raw = _safe_text(result.get("item_id"))
            item_id = raw.split(" - ", 1)[0]
            item = invm.get_item(item_id)
            result["item_id"] = item_id
            result["product_name"] = getattr(item, "name", item_id) if item else item_id
            result["batch_number"] = getattr(item, "batch_number", "") if item else ""
            result["transfer_description"] = _safe_text(result.get("transfer_description")) or _safe_text(getattr(item, "description", ""))
            result["_source_item"] = item
            self.manager.create_transfer(result, user_name=self.user_role)
            self.refresh_all()

    def _edit_transfer(self):
        transfer_id = self._selected_value(self.transfer_tree, 0)
        if not transfer_id:
            messagebox.showwarning("Warning", "Select a transfer first")
            return
        rec = self.manager.get_transfer(transfer_id)
        if not rec:
            messagebox.showerror("Error", "Selected transfer could not be found")
            return
        invm, item_names = self._inventory_item_choices()
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        current_item = invm.get_item(_safe_text(rec.get("item_id")))
        current_item_label = f"{current_item.item_id} - {current_item.name}" if current_item else _safe_text(rec.get("item_id"))
        fields = [
            {"label": "Inventory Item", "key": "item_id", "var": tk.StringVar(value=current_item_label), "kind": "combo", "values": item_names},
            {"label": "Transfer Type", "key": "transfer_type", "var": tk.StringVar(value=_safe_text(rec.get("transfer_type") or "Client Transfer")), "kind": "combo", "values": ["Client Transfer", "Internal Transfer"]},
            {"label": "Quantity", "key": "quantity", "var": tk.StringVar(value=str(rec.get("quantity") or ""))},
            {"label": "From Warehouse", "key": "warehouse_from", "var": tk.StringVar(value=_safe_text(rec.get("warehouse_from"))), "kind": "combo", "values": warehouse_codes},
            {"label": "From Bin", "key": "bin_from", "var": tk.StringVar(value=_safe_text(rec.get("bin_from")))},
            {"label": "To Warehouse", "key": "warehouse_to", "var": tk.StringVar(value=_safe_text(rec.get("warehouse_to"))), "kind": "combo", "values": warehouse_codes},
            {"label": "To Bin", "key": "bin_to", "var": tk.StringVar(value=_safe_text(rec.get("bin_to")))},
            {"label": "Client Name", "key": "destination_client", "var": tk.StringVar(value=_safe_text(rec.get("destination_client")))},
            {"label": "Client Contact", "key": "destination_contact", "var": tk.StringVar(value=_safe_text(rec.get("destination_contact")))},
            {"label": "Client Address", "key": "destination_address", "var": tk.StringVar(value=_safe_text(rec.get("destination_address")))},
            {"label": "Description", "key": "transfer_description", "var": tk.StringVar(value=_safe_text(rec.get("transfer_description") or rec.get("item_description")))},
            {"label": "Sign 1", "key": "sign_1", "var": tk.StringVar(value=_safe_text(rec.get("sign_1")))},
            {"label": "Sign 2", "key": "sign_2", "var": tk.StringVar(value=_safe_text(rec.get("sign_2")))},
            {"label": "Notes", "key": "notes", "var": tk.StringVar(value=_safe_text(rec.get("notes")))},
        ]
        result = self._record_form("Edit Transfer", fields, min_size=(760, 700))
        if result:
            raw = _safe_text(result.get("item_id"))
            item_id = raw.split(" - ", 1)[0]
            item = invm.get_item(item_id)
            result["item_id"] = item_id
            result["product_name"] = getattr(item, "name", item_id) if item else item_id
            result["batch_number"] = getattr(item, "batch_number", "") if item else ""
            result["transfer_description"] = _safe_text(result.get("transfer_description")) or _safe_text(getattr(item, "description", "")) or _safe_text(rec.get("transfer_description"))
            result["_source_item"] = item
            self.manager.update_transfer(transfer_id, result, user_name=self.user_role)
            self.refresh_all()

    def _delete_transfer(self):
        transfer_id = self._selected_value(self.transfer_tree, 0)
        if not transfer_id:
            messagebox.showwarning("Warning", "Select a transfer first")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete transfer {transfer_id}?\nIf already received, stock will be reversed."):
            return
        try:
            self.manager.delete_transfer(transfer_id, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _approve_transfer(self):
        if not self._require_admin():
            return
        transfer_id = self._selected_value(self.transfer_tree, 0)
        if transfer_id:
            self.manager.approve_transfer(transfer_id, user_name=self.user_role)
            self.refresh_all()

    def _receive_transfer(self):
        if not self._require_admin():
            return
        transfer_id = self._selected_value(self.transfer_tree, 0)
        if transfer_id:
            self.manager.receive_transfer(transfer_id, user_name=self.user_role)
            self.refresh_all()

    def _export_transfer_pdf(self):
        transfer_id = self._selected_value(self.transfer_tree, 0)
        if not transfer_id:
            messagebox.showwarning("Warning", "Select a transfer first")
            return
        try:
            path = self.manager.export_transfer_pdf(transfer_id, open_after=True)
            messagebox.showinfo("Success", f"Goods transfer PDF generated:\n{path}")
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _new_adjustment(self):
        invm, item_names = self._inventory_item_choices()
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        fields = [
            {"label": "Inventory Item", "key": "item_id", "var": tk.StringVar(value=item_names[0] if item_names else ""), "kind": "combo", "values": item_names},
            {"label": "Quantity Change (+/-)", "key": "quantity_change", "var": tk.StringVar()},
            {"label": "Adjustment Type", "key": "adjustment_type", "var": tk.StringVar(value="Correction"), "kind": "combo", "values": ["Increase", "Decrease", "Damage", "Expiry", "Lost", "Correction"]},
            {"label": "Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=warehouse_codes[0] if warehouse_codes else ""), "kind": "combo", "values": warehouse_codes},
            {"label": "Bin Location", "key": "bin_location", "var": tk.StringVar()},
            {"label": "Reason", "key": "reason", "var": tk.StringVar()},
            {"label": "Notes", "key": "notes", "var": tk.StringVar()},
        ]
        result = self._record_form("New Adjustment", fields, min_size=(680, 540))
        if result:
            raw = _safe_text(result.get("item_id"))
            item_id = raw.split(" - ", 1)[0]
            item = invm.get_item(item_id)
            result["item_id"] = item_id
            result["product_name"] = getattr(item, "name", item_id) if item else item_id
            result["batch_number"] = getattr(item, "batch_number", "") if item else ""
            self.manager.create_adjustment(result, user_name=self.user_role)
            self.refresh_all()

    def _edit_adjustment(self):
        adjustment_id = self._selected_value(self.adjustment_tree, 0)
        if not adjustment_id:
            messagebox.showwarning("Warning", "Select an adjustment first")
            return
        rec = self.manager.get_adjustment(adjustment_id)
        if not rec:
            messagebox.showerror("Error", "Selected adjustment could not be found")
            return
        invm, item_names = self._inventory_item_choices()
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        current_item = invm.get_item(_safe_text(rec.get("item_id")))
        current_item_label = f"{current_item.item_id} - {current_item.name}" if current_item else _safe_text(rec.get("item_id"))
        fields = [
            {"label": "Inventory Item", "key": "item_id", "var": tk.StringVar(value=current_item_label), "kind": "combo", "values": item_names},
            {"label": "Quantity Change (+/-)", "key": "quantity_change", "var": tk.StringVar(value=str(rec.get("quantity_change") or ""))},
            {"label": "Adjustment Type", "key": "adjustment_type", "var": tk.StringVar(value=_safe_text(rec.get("adjustment_type") or "Correction")), "kind": "combo", "values": ["Increase", "Decrease", "Damage", "Expiry", "Lost", "Correction"]},
            {"label": "Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=_safe_text(rec.get("warehouse_code"))), "kind": "combo", "values": warehouse_codes},
            {"label": "Bin Location", "key": "bin_location", "var": tk.StringVar(value=_safe_text(rec.get("bin_location")))},
            {"label": "Reason", "key": "reason", "var": tk.StringVar(value=_safe_text(rec.get("reason")))},
            {"label": "Notes", "key": "notes", "var": tk.StringVar(value=_safe_text(rec.get("notes")))},
        ]
        result = self._record_form("Edit Adjustment", fields, min_size=(680, 540))
        if result:
            raw = _safe_text(result.get("item_id"))
            item_id = raw.split(" - ", 1)[0]
            item = invm.get_item(item_id)
            result["item_id"] = item_id
            result["product_name"] = getattr(item, "name", item_id) if item else item_id
            result["batch_number"] = getattr(item, "batch_number", "") if item else ""
            self.manager.update_adjustment(adjustment_id, result, user_name=self.user_role)
            self.refresh_all()

    def _delete_adjustment(self):
        adjustment_id = self._selected_value(self.adjustment_tree, 0)
        if not adjustment_id:
            messagebox.showwarning("Warning", "Select an adjustment first")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete adjustment {adjustment_id}?\nIf already approved, stock will be reversed."):
            return
        try:
            self.manager.delete_adjustment(adjustment_id, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _approve_adjustment(self):
        if not self._require_admin():
            return
        adjustment_id = self._selected_value(self.adjustment_tree, 0)
        if adjustment_id:
            self.manager.approve_adjustment(adjustment_id, user_name=self.user_role)
            self.refresh_all()

    def _new_count(self):
        invm, item_names = self._inventory_item_choices()
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        fields = [
            {"label": "Inventory Item", "key": "item_id", "var": tk.StringVar(value=item_names[0] if item_names else ""), "kind": "combo", "values": item_names},
            {"label": "Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=warehouse_codes[0] if warehouse_codes else ""), "kind": "combo", "values": warehouse_codes},
            {"label": "Bin Location", "key": "bin_location", "var": tk.StringVar()},
            {"label": "Actual Quantity", "key": "actual_quantity", "var": tk.StringVar()},
            {"label": "Count Type", "key": "count_type", "var": tk.StringVar(value="Physical"), "kind": "combo", "values": ["Physical", "Cycle"]},
            {"label": "Notes", "key": "notes", "var": tk.StringVar()},
        ]
        result = self._record_form("New Stock Count", fields, min_size=(680, 500))
        if result:
            raw = _safe_text(result.get("item_id"))
            item_id = raw.split(" - ", 1)[0]
            item = invm.get_item(item_id)
            system_qty = _safe_float(getattr(item, "quantity", 0), 0.0) if item else 0.0
            result["item_id"] = item_id
            result["product_name"] = getattr(item, "name", item_id) if item else item_id
            result["system_quantity"] = system_qty
            self.manager.create_stock_count(result, user_name=self.user_role)
            self.refresh_all()

    def _edit_count(self):
        count_id = self._selected_value(self.count_tree, 0)
        if not count_id:
            messagebox.showwarning("Warning", "Select a stock count first")
            return
        rec = self.manager.get_stock_count(count_id)
        if not rec:
            messagebox.showerror("Error", "Selected stock count could not be found")
            return
        invm, item_names = self._inventory_item_choices()
        warehouse_codes = [row.get("warehouse_code") for row in self.manager.list_warehouses()]
        current_item = invm.get_item(_safe_text(rec.get("item_id")))
        current_item_label = f"{current_item.item_id} - {current_item.name}" if current_item else _safe_text(rec.get("item_id"))
        fields = [
            {"label": "Inventory Item", "key": "item_id", "var": tk.StringVar(value=current_item_label), "kind": "combo", "values": item_names},
            {"label": "Warehouse", "key": "warehouse_code", "var": tk.StringVar(value=_safe_text(rec.get("warehouse_code"))), "kind": "combo", "values": warehouse_codes},
            {"label": "Bin Location", "key": "bin_location", "var": tk.StringVar(value=_safe_text(rec.get("bin_location")))},
            {"label": "Actual Quantity", "key": "actual_quantity", "var": tk.StringVar(value=str(rec.get("actual_quantity") or ""))},
            {"label": "Count Type", "key": "count_type", "var": tk.StringVar(value=_safe_text(rec.get("count_type") or "Physical")), "kind": "combo", "values": ["Physical", "Cycle"]},
            {"label": "Notes", "key": "notes", "var": tk.StringVar(value=_safe_text(rec.get("notes")))},
        ]
        result = self._record_form("Edit Stock Count", fields, min_size=(680, 500))
        if result:
            raw = _safe_text(result.get("item_id"))
            item_id = raw.split(" - ", 1)[0]
            item = invm.get_item(item_id)
            system_qty = _safe_float(getattr(item, "quantity", 0), 0.0) if item else _safe_float(rec.get("system_quantity"), 0.0)
            result["item_id"] = item_id
            result["product_name"] = getattr(item, "name", item_id) if item else item_id
            result["system_quantity"] = system_qty
            self.manager.update_stock_count(count_id, result, user_name=self.user_role)
            self.refresh_all()

    def _delete_count(self):
        count_id = self._selected_value(self.count_tree, 0)
        if not count_id:
            messagebox.showwarning("Warning", "Select a stock count first")
            return
        if not messagebox.askyesno("Confirm Delete", f"Delete stock count {count_id}?\nIf already approved, stock will be restored to the system quantity."):
            return
        try:
            self.manager.delete_stock_count(count_id, user_name=self.user_role)
            self.refresh_all()
        except Exception as exc:
            messagebox.showerror("Error", str(exc))

    def _approve_count(self):
        if not self._require_admin():
            return
        count_id = self._selected_value(self.count_tree, 0)
        if count_id:
            self.manager.approve_stock_count(count_id, user_name=self.user_role)
            self.refresh_all()
