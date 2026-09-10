import os
import io
import base64
import json
import shutil
import time
import threading
import sys
import subprocess
import platform
import uuid
from datetime import datetime, timedelta
import webbrowser
from pathlib import Path
import zipfile
import urllib.request
import tempfile
import hashlib
import re
import getpass
from typing import Optional, Dict, Any

# Desktop GUI libraries — OPTIONAL. Serverless / Vercel environments do NOT have
# Tk installed, so these MUST be imported lazily (guarded by try/except at module
# load). All code paths that actually use tk/ttk/messagebox/filedialog are
# wrapped in try/except further down. On headless web / Vercel these become
# None and the Tk code paths never run anyway.
try:
    import tkinter as tk  # type: ignore
    from tkinter import ttk, messagebox, filedialog  # type: ignore
except Exception:
    tk = None
    ttk = None
    messagebox = None
    filedialog = None

try:
    import ui_undo  # type: ignore
except Exception:
    ui_undo = None

APP_VERSION = "2025.12.15"


# ============================================================================
# ASISTEM — PRODUCTION ENVIRONMENT DEFAULT FALLBACKS (serverless + local prod)
# ============================================================================
def _apply_production_env_defaults():
    """Auto-apply ASISTEM production defaults when running in serverless/prod mode.

    When deploying to Vercel Free Tier, the user can skip pasting env vars via
    the Vercel dashboard UI entirely. If we detect Vercel's serverless environment
    (VERCEL=1) OR production mode flag, we automatically seed os.environ with the
    real production Supabase credentials + Flask app secret ONLY IF the variable has
    NOT already been explicitly set (so user-override values always win).
    """
    _PROD_DEFAULTS = {
        "HOPEPHARMA_APP_SECRET": "hpmt-prod-hope-pharma-ASISTEM-2026-Sydney-rhfuxxsyeqqejtxbxlzm-!@#$%^&*()-QwErTy1234567890asdfghjkl",
        "HOPEPHARMA_PRODUCTION": "1",
        "HOPEPHARMA_DATA_DIR": "/tmp",
        "SUPABASE_URL": "https://rhfuxxsyeqqejtxbxlzm.supabase.co",
        "SUPABASE_SERVICE_ROLE_KEY": "sb_secret_JLkMjizaVdiEOv1pY03xvg_H_Xr2NaV",
    }
    is_serverless = (
        (os.environ.get("VERCEL") or "").strip() == "1"
        or bool(os.environ.get("LAMBDA_TASK_ROOT"))
        or bool(os.environ.get("AWS_LAMBDA_FUNCTION_NAME"))
        or (os.environ.get("HOPEPHARMA_PRODUCTION") or "").strip() == "1"
    )
    for key, default_val in _PROD_DEFAULTS.items():
        existing = (os.environ.get(key) or "").strip()
        if existing:
            continue
        if is_serverless:
            os.environ[key] = str(default_val)
        elif key in {"HOPEPHARMA_DATA_DIR", "HOPEPHARMA_APP_SECRET"}:
            os.environ.setdefault(key, str(default_val))

_apply_production_env_defaults()


def _ssl_context_maybe_unverified():
    """Return an SSLContext for urllib requests — UNVERIFIED if PYTHONHTTPSVERIFY=0.

    macOS .pkg Python installers ship their own OpenSSL without a valid CA bundle, so
    users often see: SSL: CERTIFICATE_VERIFY_FAILED. Python 3.12 no longer reads
    PYTHONHTTPSVERIFY for urllib automatically, so we manually apply the env var
    ourselves (and a dedicated HOPEPHARMA_SSL_VERIFY=0 escape hatch).
    """
    raw1 = (os.environ.get("PYTHONHTTPSVERIFY") or "").strip().lower()
    raw2 = (os.environ.get("HOPEPHARMA_SSL_VERIFY") or "").strip().lower()
    disabled = raw1 in {"0", "false", "no", "off"} or raw2 in {"0", "false", "no", "off"}
    if not disabled:
        return None
    try:
        import ssl as _sslmod
        return _sslmod._create_unverified_context()
    except Exception:
        return None

def _version_tuple(v):
    try:
        return tuple(int(x) for x in re.findall(r"\d+", str(v)))
    except Exception:
        return (0,)

def _load_settings():
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        p = os.path.join(base_dir, "app_settings.json")
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def _download_file(url, dest_path):
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            data = resp.read()
        with open(dest_path, "wb") as f:
            f.write(data)
        return True
    except Exception:
        return False

def _sha256(path):
    h = hashlib.sha256()
    try:
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(8192), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None

def _check_for_updates(root):
    try:
        if platform.system() != "Windows":
            return
        settings = _load_settings()
        manifest_url = settings.get("update_manifest_url") or ""
        auto_update = settings.get("auto_update", True)
        if not manifest_url:
            local_manifest = Path(__file__).with_name("updates_manifest.json")
            if not local_manifest.exists():
                return
            try:
                with open(local_manifest, "r", encoding="utf-8") as f:
                    manifest = json.load(f)
            except Exception:
                return
        else:
            try:
                with urllib.request.urlopen(manifest_url, timeout=15) as resp:
                    manifest = json.loads(resp.read().decode("utf-8"))
            except Exception:
                return
        latest = str(manifest.get("version") or "")
        if _version_tuple(latest) <= _version_tuple(APP_VERSION):
            return
        win = manifest.get("windows") or {}
        installer_url = win.get("installer_url") or ""
        exe_url = win.get("exe_url") or ""
        sha256_expected = (win.get("sha256") or "").lower().strip()
        target_url = installer_url or exe_url
        if not target_url:
            return
        def _do_update():
            try:
                tmp_dir = Path(tempfile.gettempdir())
                fname = "HopePharma_Update.exe"
                dest = tmp_dir / fname
                ok = _download_file(target_url, str(dest))
                if not ok:
                    return
                digest = _sha256(str(dest)) or ""
                if sha256_expected and digest.lower() != sha256_expected:
                    try:
                        messagebox.showerror("Update Error", "Downloaded update failed integrity check.")
                    except Exception:
                        pass
                    return
                try:
                    args = [str(dest), "/VERYSILENT", "/NORESTART"]
                    subprocess.Popen(args, shell=False)
                except Exception:
                    subprocess.Popen([str(dest)], shell=True)
                try:
                    root.destroy()
                except Exception:
                    pass
                os._exit(0)
            except Exception:
                pass
        if auto_update:
            threading.Thread(target=_do_update, daemon=True).start()
        else:
            try:
                if messagebox.askyesno("Update Available", f"A new version {latest} is available. Update now?"):
                    threading.Thread(target=_do_update, daemon=True).start()
            except Exception:
                pass
    except Exception:
        pass

# Enhanced dependency installer with better error handling
def install_dependencies():
    """Enhanced dependency installer with progress tracking.

    Desktop mode (Tk available): shows a GUI progress bar + restart on completion.
    Headless / Vercel / CLI mode (Tk missing): silent best-effort pip install only.
    """
    missing_packages = []

    # Check for required packages
    required_packages = {
        "Pillow": "PIL",
        "reportlab": "reportlab",
    }

    for package, import_name in required_packages.items():
        try:
            if import_name == "PIL":
                from PIL import Image  # headless-safe sub-import
                # ImageTk (desktop-only) is imported lazily below
            else:
                __import__(import_name)
            print(f"✅ {package} is already installed")
        except ImportError:
            missing_packages.append(package)

    if not missing_packages:
        return

    print(f"Installing missing packages: {missing_packages}")

    # ------------------------------------------------------------------
    # HEADLESS / VERCEL / SERVERLESS PATH: No Tk windows, just run pip.
    # (Vercel already installs requirements.txt so this path rarely fires.)
    # ------------------------------------------------------------------
    if tk is None:
        success_count = 0
        for package in missing_packages:
            try:
                result = subprocess.run(
                    [sys.executable, "-m", "pip", "install", package,
                     "--timeout", "30", "--retries", "2", "--quiet"],
                    capture_output=True, text=True, timeout=120,
                )
                if result.returncode == 0:
                    success_count += 1
                    print(f"✅ Successfully installed {package}")
                else:
                    print(f"❌ Failed to install {package}: {result.stderr}")
            except Exception as e:
                print(f"❌ Error installing {package}: {e}")
        # On serverless, NEVER os.execv() — that kills the lambda function.
        return

    # ------------------------------------------------------------------
    # DESKTOP PATH: Tk is available → progress window + restart.
    # ------------------------------------------------------------------
    # Create progress window
    progress_root = tk.Tk()
    progress_root.title("Installing Dependencies")
    progress_root.geometry("400x150")
    progress_root.resizable(False, False)

    # Center the window
    progress_root.eval('tk::PlaceWindow . center')

    label = tk.Label(progress_root, text="Installing required packages...", pady=10)
    label.pack()

    progress = ttk.Progressbar(progress_root, mode='indeterminate')
    progress.pack(fill='x', padx=20, pady=10)
    progress.start()

    status_label = tk.Label(progress_root, text="Starting installation...")
    status_label.pack(pady=5)

    progress_root.update()

    success_count = 0
    for package in missing_packages:
        status_label.config(text=f"Installing {package}...")
        progress_root.update()

        try:
            # Use pip with timeout
            result = subprocess.run([
                sys.executable, "-m", "pip", "install", package,
                "--timeout", "30", "--retries", "2"
            ], capture_output=True, text=True, timeout=60)

            if result.returncode == 0:
                print(f"✅ Successfully installed {package}")
                success_count += 1
            else:
                print(f"❌ Failed to install {package}: {result.stderr}")

        except subprocess.TimeoutExpired:
            print(f"❌ Timeout installing {package}")
        except Exception as e:
            print(f"❌ Error installing {package}: {e}")

    progress.stop()
    if success_count == len(missing_packages):
        try:
            messagebox.showinfo("Success",
                              "All dependencies installed successfully!\nThe application will now restart.",
                              parent=progress_root)
        except Exception:
            pass
        try:
            progress_root.destroy()
        except Exception:
            pass
        os.execv(sys.executable, [sys.executable] + sys.argv)
    else:
        try:
            response = messagebox.askretrycancel(
                "Installation Incomplete",
                f"Only {success_count} out of {len(missing_packages)} packages were installed.\n\n"
                "Would you like to try again? Click Cancel to continue anyway.",
                parent=progress_root
            )
        except Exception:
            response = False
        try:
            progress_root.destroy()
        except Exception:
            pass
        if response:
            install_dependencies()


# Install dependencies before anything else (NEVER crash the whole app on failure —
# on Vercel, requirements.txt already handles it, so failures here are non-fatal).
try:
    install_dependencies()
except Exception as _install_err:
    print(f"[install_dependencies] non-fatal: {_install_err}")

# Now import the rest (headless-safe: no Tk-bound libraries at module scope)
try:
    from PIL import Image  # headless-safe
except Exception:
    Image = None  # type: ignore
try:
    from PIL import ImageTk  # desktop-only, needs Tk
except Exception:
    ImageTk = None  # type: ignore
import reportlab
from reportlab.lib.pagesizes import letter, A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, KeepInFrame
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from xml.sax.saxutils import escape as xml_escape

class EnhancedPDFGenerator:
    """Enhanced PDF Generator — per-company letterhead, base64 branding, Sales/Service toggle.

    Upgraded for the ASISTEM 2026 Q3 ERP/CRM release:
      - Company profile passed in (via data_manager + company_id OR explicit dict)
      - Hardcoded HopePharma strings REMOVED; letterhead read from profile.name/TRN/address/bank
      - signature_png_base64 / stamp_png_base64  →  written to unique /tmp PNG per call
      - invoice.type == "service"  →  title "SERVICE TAX INVOICE"; otherwise "TAX INVOICE"
      - profile.show_cost_center_on_pdf  →  inserts extra "Cost Center" column on items table
      - generate_invoice_pdf_bytes() returns raw bytes (for Flask send_file inline/attachment)
    """

    def __init__(self, logo_manager=None, output_folder=None, data_manager=None):
        self.logo_manager = logo_manager
        self.styles = getSampleStyleSheet()
        self.output_folder = output_folder
        self.data_manager = data_manager
        # Temp PNGs we create per-call, cleaned up with best-effort after doc.build()
        self._tmp_png_paths = []

    # ------------------------------------------------------------------
    # Company profile resolution & base64 → /tmp PNG helpers
    # ------------------------------------------------------------------
    def _resolve_company_profile(self, invoice_data):
        """Return a Company Profile dict for the given invoice.

        Priority:
          1) `invoice_data["_company_profile"]` (injected by caller for speed)
          2) Lookup via self.data_manager + invoice_data["company_id"]
          3) HopePharma default seed (last-resort fallback for pre-migration invoices)
        """
        explicit = (invoice_data or {}).get("_company_profile")
        if isinstance(explicit, dict):
            return explicit
        if self.data_manager is not None:
            cid = (invoice_data or {}).get("company_id")
            if cid and hasattr(self.data_manager, "get_company_profile_by_id"):
                found = self.data_manager.get_company_profile_by_id(cid)
                if isinstance(found, dict):
                    return found
            # Fall back to DEFAULT_COMPANY_ID profile (HopePharma) if loaded
            if hasattr(self.data_manager, "get_all_company_profiles"):
                try:
                    all_p = self.data_manager.get_all_company_profiles() or []
                    default_cid = getattr(self.data_manager, "DEFAULT_COMPANY_ID", None)
                    for p in all_p:
                        if isinstance(p, dict) and p.get("id") == default_cid:
                            return p
                    if all_p and isinstance(all_p[0], dict):
                        return all_p[0]
                except Exception:
                    pass
        try:
            from hope_pharma_complete import EnhancedCloudDataManager
            return EnhancedCloudDataManager._default_company_seed()
        except Exception:
            return {
                "name": "ASISTEM Tenant",
                "address": "",
                "trn": "",
                "currency": "AED",
                "vat_rate": 5.0,
                "bank_name": "",
                "bank_branch": "",
                "bank_account_no": "",
                "bank_iban": "",
                "bank_swift": "",
                "pdf_footer_thank_you": "Thank you for your business!",
                "show_cost_center_on_pdf": False,
                "signature_png_base64": None,
                "stamp_png_base64": None,
                "logo_png_base64": None,
            }

    @staticmethod
    def _data_uri_to_bytes(data_uri):
        """Convert a 'data:image/png;base64,XXXX' (or plain base64) string → raw PNG bytes."""
        if not data_uri:
            return b""
        s = str(data_uri).strip()
        if "," in s and s.startswith("data:"):
            s = s.split(",", 1)[1]
        try:
            return base64.b64decode(s + ("=" * (-len(s) % 4)))
        except Exception:
            return b""

    def _base64_to_tmp_png(self, data_uri, tag="branding"):
        """Write a base64 image to /tmp/<unique>-<tag>.png; return path or None.

        Paths are added to self._tmp_png_paths so they can be unlinked after the build.
        """
        raw = self._data_uri_to_bytes(data_uri)
        if not raw:
            return None
        try:
            fd, path = tempfile.mkstemp(prefix=f"asistem_{tag}_", suffix=".png",
                                        dir=os.environ.get("HOPEPHARMA_DATA_DIR") or "/tmp")
            with os.fdopen(fd, "wb") as fh:
                fh.write(raw)
            self._tmp_png_paths.append(path)
            return path
        except Exception:
            return None

    def _cleanup_tmp_pngs(self):
        for p in list(self._tmp_png_paths):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass
            self._tmp_png_paths = [x for x in self._tmp_png_paths if x != p]

    # ------------------------------------------------------------------
    # Logo resolution (base64 from profile first, then legacy file paths)
    # ------------------------------------------------------------------
    def _resolve_logo_path(self, company_profile=None):
        # 1) profile.logo_png_base64 → /tmp PNG
        if isinstance(company_profile, dict):
            logo_b64 = company_profile.get("logo_png_base64")
            if logo_b64:
                p = self._base64_to_tmp_png(logo_b64, "logo")
                if p:
                    return p
        # 2) legacy file system candidates
        candidates = []
        try:
            if self.logo_manager and getattr(self.logo_manager, 'logo_path', None):
                candidates.append(self.logo_manager.logo_path)
        except Exception:
            pass
        try:
            if self.output_folder:
                out_folder = str(self.output_folder)
                candidates.extend([
                    os.path.join(out_folder, 'logo.png'),
                    os.path.join(out_folder, 'logo.jpg'),
                    os.path.join(out_folder, 'logo.jpeg'),
                    os.path.join(os.path.dirname(out_folder), 'logo.png'),
                ])
        except Exception:
            pass
        base_dir = os.path.dirname(os.path.abspath(__file__))
        candidates.extend([
            os.path.join(base_dir, 'logo.png'),
            os.path.join(base_dir, 'logo.jpg'),
            os.path.join(base_dir, 'logo.jpeg'),
            os.path.join(base_dir, 'AssistemLogo.png'),
            os.path.join(base_dir, 'companies', 'Hope Pharma Medicine Trading', 'logo.png'),
            os.path.join(base_dir, 'dist', 'HopePharma.app', 'Contents', 'Resources', 'logo.png'),
            os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData", "logo.png"),
            os.path.join(os.path.expanduser("~"), "Desktop", "HopePharmaData", "logo.png"),
        ])
        for candidate in candidates:
            try:
                if candidate and os.path.exists(candidate):
                    return candidate
            except Exception:
                continue
        return None

    def _append_logo_flowables(self, elements, width=1.1 * inch, height=1.1 * inch, spacer=0.1 * inch, company_profile=None):
        try:
            logo_path = self._resolve_logo_path(company_profile)
            if logo_path and os.path.exists(logo_path):
                elements.append(Image(logo_path, width=width, height=height))
                if spacer:
                    elements.append(Spacer(1, spacer))
        except Exception:
            pass
    
    def _create_custom_styles(self):
        """Create custom styles for PDF generation"""
        from reportlab.lib.styles import getSampleStyleSheet
        _ss = getSampleStyleSheet()
        normal_parent = _ss['Normal']
        custom_styles = {
            'Normal': ParagraphStyle(
                'NormalClone', parent=normal_parent, fontSize=9, leading=11
            ),
            'Title': ParagraphStyle(
                'Title',
                parent=self.styles['Heading1'],
                fontSize=16,
                spaceAfter=12,
                alignment=TA_CENTER,
                textColor=colors.darkblue,
                fontName='Helvetica-Bold'
            ),
            'Subtitle': ParagraphStyle(
                'Subtitle',
                parent=self.styles['Heading2'],
                fontSize=10,
                spaceAfter=8,
                alignment=TA_CENTER,
                textColor=colors.darkblue
            ),
            'Company': ParagraphStyle(
                'Company',
                parent=self.styles['Normal'],
                fontSize=9,
                alignment=TA_CENTER,
                textColor=colors.darkgreen
            ),
            'Header': ParagraphStyle(
                'Header',
                parent=self.styles['Normal'],
                fontSize=8,
                alignment=TA_LEFT,
                textColor=colors.black
            ),
            'Footer': ParagraphStyle(
                'Footer',
                parent=self.styles['Normal'],
                fontSize=7,
                alignment=TA_CENTER,
                textColor=colors.grey
            ),
            'Total': ParagraphStyle(
                'Total',
                parent=self.styles['Normal'],
                fontSize=10,
                textColor=colors.black,
                fontName='Helvetica-Bold'
            ),
            'TermsTitle': ParagraphStyle(
                'TermsTitle',
                parent=self.styles['Normal'],
                fontSize=9,
                textColor=colors.darkblue,
                fontName='Helvetica-Bold',
                alignment=TA_LEFT,
                spaceAfter=4
            ),
            'Terms': ParagraphStyle(
                'Terms',
                parent=self.styles['Normal'],
                fontSize=8,
                textColor=colors.black,
                alignment=TA_LEFT,
                leftIndent=0
            ),
            'InvoiceTitle': ParagraphStyle(
                'InvoiceTitle',
                parent=self.styles['Heading1'],
                fontSize=14,
                spaceAfter=6,
                alignment=TA_CENTER,
                textColor=colors.black,
                fontName='Helvetica-Bold'
            ),
            'ClientInfo': ParagraphStyle(
                'ClientInfo',
                parent=self.styles['Normal'],
                fontSize=8,
                alignment=TA_LEFT,
                textColor=colors.black,
                leftIndent=0
            ),
            'ItemCell': ParagraphStyle(
                'ItemCell',
                parent=self.styles['Normal'],
                fontSize=7,
                leading=8.5,
                alignment=TA_LEFT,
                textColor=colors.black
            )
        }
        return custom_styles
    
    def _safe_float_conversion(self, value, default=0.0):
        """Safely convert value to float"""
        try:
            return float(value) if value else default
        except (ValueError, TypeError):
            return default
    
    def generate_invoice_pdf(self, invoice_data):
        """Generate enhanced invoice PDF with per-company letterhead and return file path.

        Calls _build_invoice_pdf(doc_path) internally. To get raw bytes instead of a
        file on disk, call generate_invoice_pdf_bytes() instead.
        """
        # --- Determine destination folder ---
        invoices_folder = None
        try:
            if self.output_folder:
                invoices_folder = str(self.output_folder)
        except Exception:
            invoices_folder = None
        if not invoices_folder:
            possible_roots = [
                os.path.join(os.path.expanduser("~"), "Google Drive"),
                os.path.join(os.path.expanduser("~"), "GoogleDrive"),
                os.path.join(os.path.expanduser("~"), "Documents"),
                os.path.join(os.path.expanduser("~"), "Desktop"),
            ]
            for root in possible_roots:
                data_path = os.path.join(root, "HopePharmaData")
                if os.path.exists(data_path):
                    invoices_folder = os.path.join(data_path, "HopePharmaInvoices")
                    break
            if not invoices_folder:
                base_dir = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")
                invoices_folder = os.path.join(base_dir, "HopePharmaInvoices")
        if not os.path.exists(invoices_folder):
            try:
                os.makedirs(invoices_folder)
            except Exception:
                pass
        # --- Filename ---
        client_name = invoice_data.get('client_name', 'Client')
        clean_client_name = "".join(c for c in str(client_name) if c.isalnum() or c in (' ', '-', '_')).rstrip().replace(' ', '')
        invoice_id = invoice_data.get('invoice_id', 'INV')
        pdf_filename = f"{clean_client_name}_{invoice_id}.pdf"
        pdf_path = os.path.join(invoices_folder, pdf_filename)
        return self._build_invoice_pdf(invoice_data, pdf_path, target="file")

    def generate_invoice_pdf_bytes(self, invoice_data):
        """Build the invoice PDF into an in-memory buffer and return bytes.

        Used by Flask `/invoices/<id>/pdf` and `/download/pdf` routes.
        """
        tmp_root = os.environ.get("HOPEPHARMA_DATA_DIR") or tempfile.gettempdir()
        fd, tmppath = tempfile.mkstemp(prefix="asistem_inv_", suffix=".pdf", dir=tmp_root)
        os.close(fd)
        try:
            self._build_invoice_pdf(invoice_data, tmppath, target="bytes")
            with open(tmppath, "rb") as fh:
                data = fh.read()
            return data
        finally:
            try:
                if os.path.exists(tmppath):
                    os.remove(tmppath)
            except Exception:
                pass

    def _build_invoice_pdf(self, invoice_data, pdf_path, target="file"):
        """Shared core: build invoice PDF into pdf_path, return pdf_path or raise.

        Implementation notes (FR-6 / Task 2):
          * NO hardcoded HopePharma strings anywhere — all letterhead fields read
            from company_profile (name, address, TRN, currency, vat_rate, bank_*,
            pdf_footer_thank_you, show_cost_center_on_pdf, signature_png_base64,
            stamp_png_base64, logo_png_base64)
          * invoice['type'] == 'service'  →  "SERVICE TAX INVOICE"
                                       else  →  "TAX INVOICE"
          * profile['show_cost_center_on_pdf'] = True  →  extra "Cost Center"
            column on items table using per-line line_cost_center_id label
          * Branding images (logo/signature/stamp) resolved in order:
              1) base64 data URI from company_profile → /tmp PNG
              2) legacy sign1.png / sign2.png on disk (signature/stamp fallback)
        """
        try:
            profile = self._resolve_company_profile(invoice_data)
            company_name = str(profile.get("name") or profile.get("display_name") or "ASISTEM Tenant")
            company_address = str(profile.get("address") or "")
            company_trn = str(profile.get("trn") or "")
            currency = str(invoice_data.get("currency") or profile.get("currency") or "AED").strip() or "AED"
            default_vat = self._safe_float_conversion(profile.get("vat_rate"), 5.0)
            show_cc = bool(profile.get("show_cost_center_on_pdf"))

            items_list = invoice_data.get('items', []) or []
            orientation = str(invoice_data.get('pdf_orientation', '') or '').strip().lower()
            max_desc_len = 0
            try:
                max_desc_len = max((len(str(x.get('description', '') or '')) for x in items_list), default=0)
            except Exception:
                max_desc_len = 0
            use_landscape = False
            if orientation in {'landscape', 'l'}:
                use_landscape = True
            elif orientation in {'portrait', 'p'}:
                use_landscape = False
            else:
                n_cols = 6 if show_cc else 5
                use_landscape = (len(items_list) > (15 if n_cols == 6 else 18)) or (max_desc_len > 50)

            page_size = landscape(A4) if use_landscape else A4
            doc = SimpleDocTemplate(
                pdf_path, pagesize=page_size,
                topMargin=0.25*inch, bottomMargin=0.25*inch,
                leftMargin=0.35*inch, rightMargin=0.35*inch,
            )
            content_width = doc.width
            elements = []
            custom_styles = self._create_custom_styles()

            # ========== HEADER: Logo + Company info (ALL from profile) ==========
            logo_path = self._resolve_logo_path(profile)
            title_rows = [[Paragraph(xml_escape(company_name).upper(), custom_styles['Title'])]]
            if company_address:
                title_rows.append([Paragraph(xml_escape(company_address), custom_styles['Company'])])
            if company_trn:
                title_rows.append([Paragraph(f"TRN: {xml_escape(company_trn)}", custom_styles['Company'])])
            logo_width = 1.2 * inch
            company_width = max(content_width - (1.4 * inch), 4.0 * inch)
            company_table = Table(title_rows, colWidths=[company_width])
            company_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 2),
            ]))
            if logo_path:
                try:
                    logo_img = Image(logo_path, width=logo_width, height=logo_width)
                    header = Table([[logo_img, company_table]], colWidths=[1.4 * inch, company_width])
                    header.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))
                    elements.append(header)
                except Exception:
                    elements.append(company_table)
            else:
                elements.append(company_table)
            elements.append(Spacer(1, 0.1 * inch))

            # ========== INVOICE TITLE (Sales vs Service toggle) ==========
            inv_type = str(invoice_data.get("type") or "sales").strip().lower()
            if inv_type == "service":
                title_text = "SERVICE TAX INVOICE"
            else:
                title_text = "TAX INVOICE"
            elements.append(Paragraph(title_text, custom_styles['InvoiceTitle']))
            elements.append(Spacer(1, 0.08 * inch))

            # ========== CLIENT + INVOICE DETAILS ==========
            pm = invoice_data.get('payment_method', invoice_data.get('payment_terms', 'Cash'))
            location_val = (invoice_data.get('client_location') or '').strip()
            if not location_val:
                location_val = str(invoice_data.get('client_emirate', '') or '').strip()
            uae_emirates = {'Dubai', 'Abu Dhabi', 'Sharjah', 'Ajman', 'Umm Al Quwain', 'Ras Al Khaimah', 'Fujairah'}
            if ',' not in location_val and location_val in uae_emirates:
                location_val = f"{location_val}, United Arab Emirates"
            location_line = f"{xml_escape(location_val)}<br/>" if location_val else ""
            client_trn_val = str(invoice_data.get('client_trn') or '').strip()
            client_info = (
                "<b>BILL TO</b><br/>"
                f"{xml_escape(str(invoice_data.get('client_name', 'N/A')))}<br/>"
                f"{location_line}"
                + (f"TRN: {xml_escape(client_trn_val)}<br/>" if client_trn_val else "")
            )
            right_block = (
                f"Invoice ID: {xml_escape(str(invoice_data.get('invoice_id', 'N/A')))}<br/>"
                f"Invoice Date: {xml_escape(str(invoice_data.get('date', 'N/A')))}"
            )
            invoice_info = f"<para align=right>{right_block}</para>"
            details_data = [
                [Paragraph(client_info, custom_styles['ClientInfo']),
                 Paragraph(invoice_info, custom_styles['ClientInfo'])]
            ]
            details_table = Table(details_data, colWidths=[content_width * 0.55, content_width * 0.45])
            details_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ]))
            elements.append(details_table)
            elements.append(Spacer(1, 0.1 * inch))

            # ========== ITEMS TABLE (optional Cost Center column) ==========
            elements.append(Paragraph("INVOICE ITEMS", self.styles['Heading2']))
            elements.append(Spacer(1, 0.04 * inch))

            # Build header row
            if show_cc:
                header_row = ['QTY', 'DESCRIPTION', 'COST CENTER', 'VAT',
                              f'UNIT PRICE ({currency})', f'AMOUNT ({currency})']
            else:
                header_row = ['QTY', 'DESCRIPTION', 'VAT',
                              f'UNIT PRICE ({currency})', f'AMOUNT ({currency})']
            items_data = [header_row]

            # Resolve cost center labels once (id→name) if enabled
            cc_labels = {}
            if show_cc:
                cid = invoice_data.get("company_id")
                if self.data_manager and cid and hasattr(self.data_manager, "list_cost_centers"):
                    try:
                        cc_rows = self.data_manager.list_cost_centers(cid) or []
                        cc_labels = {str(r.get("id")): (r.get("label") or r.get("name") or str(r.get("id"))) for r in cc_rows if isinstance(r, dict)}
                    except Exception:
                        cc_labels = {}

            for item in items_list:
                quantity = item.get('quantity', 1)
                try:
                    quantity = int(float(quantity))
                except Exception:
                    quantity = 1
                unit_price = self._safe_float_conversion(item.get('unit_price', 0))
                total = self._safe_float_conversion(item.get('total', 0))
                vat_status = "No" if not item.get('taxable', True) else "Yes"
                desc_value = xml_escape(str(item.get('description', '') or ''))
                row = [str(quantity), Paragraph(desc_value, custom_styles['ItemCell'])]
                if show_cc:
                    cc_id = str(item.get('line_cost_center_id') or invoice_data.get('cost_center_id') or '')
                    cc_display = cc_labels.get(cc_id, cc_id) if cc_id else '-'
                    row.append(xml_escape(cc_display))
                row.extend([vat_status, f"{unit_price:,.2f}", f"{total:,.2f}"])
                items_data.append(row)

            # Column widths - balance given show_cc
            qty_w = content_width * 0.08
            vat_w = content_width * 0.07
            unit_w = content_width * 0.15
            amt_w = content_width * 0.15
            if show_cc:
                cc_w = content_width * 0.14
                desc_w = max(content_width - (qty_w + cc_w + vat_w + unit_w + amt_w), 2.0 * inch)
                col_widths = [qty_w, desc_w, cc_w, vat_w, unit_w, amt_w]
            else:
                desc_w = max(content_width - (qty_w + vat_w + unit_w + amt_w), 2.5 * inch)
                col_widths = [qty_w, desc_w, vat_w, unit_w, amt_w]
            remaining = content_width - sum(col_widths)
            if remaining < 0:
                take = abs(remaining)
                unit_w = max(unit_w - take / 3.0, 1.0 * inch)
                amt_w = max(amt_w - take / 3.0, 1.0 * inch)
                desc_w = max(desc_w - take / 3.0, 1.8 * inch)
                if show_cc:
                    col_widths = [qty_w, desc_w, cc_w, vat_w, unit_w, amt_w]
                else:
                    col_widths = [qty_w, desc_w, vat_w, unit_w, amt_w]

            items_table = Table(items_data, colWidths=col_widths, repeatRows=1)
            items_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 7),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 4),
                ('BACKGROUND', (0, 1), (-1, -1), colors.lightgrey),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 7),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('ALIGN', (-2, 1), (-1, -1), 'RIGHT'),
                ('ALIGN', (0, 1), (0, -1), 'CENTER'),
            ]))
            elements.append(items_table)
            elements.append(Spacer(1, 0.1 * inch))

            # ========== TOTALS ==========
            subtotal = self._safe_float_conversion(invoice_data.get('subtotal', 0))
            taxable_amount = self._safe_float_conversion(invoice_data.get('taxable_amount', 0))
            non_taxable_amount = self._safe_float_conversion(invoice_data.get('non_taxable_amount', 0))
            tax_amount = self._safe_float_conversion(invoice_data.get('tax_amount', 0))
            tax_rate = self._safe_float_conversion(invoice_data.get('tax_rate'), default_vat)
            grand_total = self._safe_float_conversion(invoice_data.get('grand_total', 0))
            totals_data = [
                ['Subtotal:', f"{currency} {subtotal:,.2f}"],
                ['Taxable Amount:', f"{currency} {taxable_amount:,.2f}"],
                ['Non-Taxable Amount:', f"{currency} {non_taxable_amount:,.2f}"],
                [f'VAT ({tax_rate:g}%):', f"{currency} {tax_amount:,.2f}"],
                ['', ''],
                ['GRAND TOTAL:', f"{currency} {grand_total:,.2f}"],
            ]
            totals_table = Table(totals_data, colWidths=[content_width * 0.30, content_width * 0.20])
            totals_table.hAlign = 'RIGHT'
            totals_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
                ('FONTNAME', (0, 5), (-1, 5), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 5), (-1, 5), 10),
                ('LINEABOVE', (0, 5), (-1, 5), 2, colors.black),
                ('BOTTOMPADDING', (0, 4), (-1, 4), 2),
            ]))
            elements.append(totals_table)
            elements.append(Spacer(1, 0.06 * inch))

            # ========== SIGNATURE + STAMP (profile base64 → /tmp PNG first) ==========
            sig_path = self._base64_to_tmp_png(profile.get("signature_png_base64"), "signature")
            stamp_path = self._base64_to_tmp_png(profile.get("stamp_png_base64"), "stamp")
            # Fallback to legacy sign1.png / sign2.png on disk if profile has none
            base_dir = os.path.dirname(os.path.abspath(__file__))
            if not sig_path:
                legacy_sig = os.path.join(base_dir, "sign1.png")
                if os.path.exists(legacy_sig):
                    sig_path = legacy_sig
            if not stamp_path:
                legacy_stamp = os.path.join(base_dir, "sign2.png")
                if os.path.exists(legacy_stamp):
                    stamp_path = legacy_stamp

            left_cell = Paragraph("", self.styles['Normal'])
            right_cell = Paragraph("", self.styles['Normal'])
            if sig_path and os.path.exists(sig_path):
                try:
                    left_cell = Image(sig_path, width=1.8 * inch, height=1.8 * inch)
                except Exception:
                    left_cell = Paragraph("", self.styles['Normal'])
            if stamp_path and os.path.exists(stamp_path):
                try:
                    right_cell = Image(stamp_path, width=2.4 * inch, height=0.9 * inch)
                except Exception:
                    right_cell = Paragraph("", self.styles['Normal'])
            sig_table = Table([[left_cell, right_cell]], colWidths=[content_width * 0.5, content_width * 0.5])
            sig_table.hAlign = 'LEFT'
            sig_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('ALIGN', (0, 0), (0, 0), 'LEFT'),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ]))
            elements.append(sig_table)
            elements.append(Spacer(1, 0.04 * inch))

            # Notes
            inv_notes = (invoice_data.get('notes') or '').strip()
            if inv_notes and invoice_data.get('source') != 'inventory':
                elements.append(Paragraph("NOTES", custom_styles['TermsTitle']))
                elements.append(Paragraph(xml_escape(inv_notes).replace('\n', '<br/>'), custom_styles['Terms']))
                elements.append(Spacer(1, 0.08 * inch))

            elements.append(Paragraph("Signature", custom_styles['ClientInfo']))
            elements.append(Spacer(1, 0.08 * inch))

            # ========== TERMS & CONDITIONS (bank details ALL from profile) ==========
            elements.append(Paragraph("TERMS & CONDITIONS", custom_styles['TermsTitle']))
            bank_name = str(profile.get("bank_name") or "")
            bank_branch = str(profile.get("bank_branch") or "")
            bank_acct = str(profile.get("bank_account_no") or "")
            bank_iban = str(profile.get("bank_iban") or "")
            bank_swift = str(profile.get("bank_swift") or "")
            terms_parts = [f"All payments should be in favour of {xml_escape(company_name)}<br/>"]
            if pm:
                terms_parts.append(f"Payment Terms: {xml_escape(str(pm))}<br/>")
            bank_line_parts = []
            if bank_name:
                bank_line_parts.append(f"Bank: {xml_escape(bank_name)}")
            if bank_branch:
                bank_line_parts.append(f"Branch: {xml_escape(bank_branch)}")
            if bank_line_parts:
                terms_parts.append(" | ".join(bank_line_parts) + "<br/>")
            if bank_acct:
                terms_parts.append(f"Account No: {xml_escape(bank_acct)}<br/>")
            iban_parts = []
            if bank_iban:
                iban_parts.append(f"IBAN: {xml_escape(bank_iban)}")
            if bank_swift:
                iban_parts.append(f"Swift Code: {xml_escape(bank_swift)}")
            if iban_parts:
                terms_parts.append(" | ".join(iban_parts))
            terms_text = "".join(terms_parts)
            elements.append(Paragraph(terms_text, custom_styles['Terms']))

            # ========== FOOTER (thank you from profile) ==========
            elements.append(Spacer(1, 0.1 * inch))
            thanks = str(profile.get("pdf_footer_thank_you") or "Thank you for your business!")
            footer_text = f"{xml_escape(thanks)}<br/>Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}"
            elements.append(Paragraph(footer_text, custom_styles['Footer']))

            doc.build(elements)
            if target == "file":
                print(f"Enhanced invoice PDF generated successfully: {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Error generating enhanced invoice PDF: {e}")
            import traceback
            traceback.print_exc()
            raise
        finally:
            # Best-effort cleanup of per-call /tmp branding PNGs
            self._cleanup_tmp_pngs()

    def generate_credit_note_pdf(self, credit_note_data):
        try:
            out_folder = None
            try:
                if self.output_folder:
                    out_folder = str(self.output_folder)
            except Exception:
                out_folder = None

            if not out_folder:
                possible_roots = [
                    os.path.join(os.path.expanduser("~"), "Google Drive"),
                    os.path.join(os.path.expanduser("~"), "GoogleDrive"),
                    os.path.join(os.path.expanduser("~"), "Documents"),
                    os.path.join(os.path.expanduser("~"), "Desktop")
                ]
                for root in possible_roots:
                    data_path = os.path.join(root, "HopePharmaData")
                    if os.path.exists(data_path):
                        out_folder = os.path.join(data_path, "HopePharmaCreditNotes")
                        break
                if not out_folder:
                    base_dir = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")
                    out_folder = os.path.join(base_dir, "HopePharmaCreditNotes")

            if not os.path.exists(out_folder):
                try:
                    os.makedirs(out_folder)
                except Exception:
                    pass

            customer_name = credit_note_data.get("customer_name", "Customer")
            clean_customer = "".join(c for c in str(customer_name) if c.isalnum() or c in (' ', '-', '_')).rstrip().replace(" ", "")
            credit_note_id = credit_note_data.get("credit_note_id", "CN")
            pdf_filename = f"{clean_customer}_{credit_note_id}.pdf"
            pdf_path = os.path.join(out_folder, pdf_filename)

            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=A4,
                topMargin=0.25 * inch,
                bottomMargin=0.25 * inch,
                leftMargin=0.35 * inch,
                rightMargin=0.35 * inch
            )
            content_width = doc.width
            elements = []
            custom_styles = self._create_custom_styles()

            logo_path = self._resolve_logo_path()
            company_block = [
                [Paragraph("HOPE PHARMA MEDICINE TRADING", custom_styles['Title'])],
                [Paragraph("Dubai, International City, Morocco Cluster, Bldg. I-16, Store 08", custom_styles['Company'])],
                [Paragraph("TRN: 100466797600003", custom_styles['Company'])]
            ]
            logo_width = 1.2 * inch
            company_width = max(content_width - (1.4 * inch), 4.0 * inch)
            company_table = Table(company_block, colWidths=[company_width])
            company_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 2),
            ]))
            if logo_path:
                logo_img = Image(logo_path, width=logo_width, height=logo_width)
                header = Table([[logo_img, company_table]], colWidths=[1.4 * inch, company_width])
                header.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                elements.append(header)
            else:
                elements.append(company_table)
            elements.append(Spacer(1, 0.1 * inch))

            elements.append(Paragraph("CREDIT NOTE", custom_styles['InvoiceTitle']))
            elements.append(Spacer(1, 0.08 * inch))

            ref_invoice = credit_note_data.get("invoice_id") or credit_note_data.get("reference_invoice_id") or ""
            left_info = f"""
            <b>CREDIT TO</b><br/>
            {xml_escape(str(credit_note_data.get('customer_name', 'N/A')))}<br/>
            """
            right_block = f"Credit Note ID: {xml_escape(str(credit_note_id))}<br/>Date: {xml_escape(str(credit_note_data.get('date', 'N/A')))}"
            if ref_invoice:
                right_block += f"<br/>Reference Invoice: {xml_escape(str(ref_invoice))}"
            details_data = [[Paragraph(left_info, custom_styles['ClientInfo']), Paragraph(f"<para align=right>{right_block}</para>", custom_styles['ClientInfo'])]]
            details_table = Table(details_data, colWidths=[content_width * 0.55, content_width * 0.45])
            details_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
            ]))
            elements.append(details_table)
            elements.append(Spacer(1, 0.1 * inch))

            items = credit_note_data.get("items", []) or []
            table_data = [[
                Paragraph("<b>Description</b>", custom_styles['ItemCell']),
                Paragraph("<b>Qty</b>", custom_styles['ItemCell']),
                Paragraph("<b>Unit Price</b>", custom_styles['ItemCell']),
                Paragraph("<b>Subtotal</b>", custom_styles['ItemCell']),
                Paragraph("<b>VAT</b>", custom_styles['ItemCell']),
                Paragraph("<b>Total</b>", custom_styles['ItemCell']),
            ]]

            subtotal = 0.0
            vat_total = 0.0
            for row in items:
                desc = str(row.get("description", "") or "")
                qty = self._safe_float_conversion(row.get("quantity", 0), 0.0)
                unit_price = self._safe_float_conversion(row.get("unit_price", 0), 0.0)
                line_subtotal = qty * unit_price
                taxable = bool(row.get("taxable", False))
                vat_rate = self._safe_float_conversion(row.get("vat_rate", 0.0), 0.0)
                line_vat = (line_subtotal * vat_rate) if taxable else 0.0
                line_total = line_subtotal + line_vat
                subtotal += line_subtotal
                vat_total += line_vat
                table_data.append([
                    Paragraph(xml_escape(desc), custom_styles['ItemCell']),
                    Paragraph(f"{qty:g}", custom_styles['ItemCell']),
                    Paragraph(f"{unit_price:,.2f}", custom_styles['ItemCell']),
                    Paragraph(f"{line_subtotal:,.2f}", custom_styles['ItemCell']),
                    Paragraph(f"{line_vat:,.2f}", custom_styles['ItemCell']),
                    Paragraph(f"{line_total:,.2f}", custom_styles['ItemCell']),
                ])

            col_widths = [content_width * 0.44, content_width * 0.08, content_width * 0.12, content_width * 0.12, content_width * 0.10, content_width * 0.14]
            items_table = Table(table_data, colWidths=col_widths, repeatRows=1)
            items_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.4, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f0f3f5")),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ]))
            elements.append(items_table)
            elements.append(Spacer(1, 0.1 * inch))

            restocking_fee = self._safe_float_conversion(credit_note_data.get("restocking_fee", 0.0), 0.0)
            stated_subtotal = self._safe_float_conversion(credit_note_data.get("subtotal", subtotal), subtotal)
            stated_vat = self._safe_float_conversion(credit_note_data.get("vat_amount", vat_total), vat_total)
            total_credit = self._safe_float_conversion(credit_note_data.get("total_credit", (stated_subtotal + stated_vat - restocking_fee)), (stated_subtotal + stated_vat - restocking_fee))

            totals_data = [
                ["Subtotal", f"AED {stated_subtotal:,.2f}"],
                ["VAT", f"AED {stated_vat:,.2f}"],
            ]
            if restocking_fee:
                totals_data.append(["Restocking Fee", f"- AED {restocking_fee:,.2f}"])
            totals_data.append(["Total Credit", f"AED {total_credit:,.2f}"])

            totals_table = Table(totals_data, colWidths=[content_width * 0.7, content_width * 0.3])
            totals_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.4, colors.grey),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#f8f9fa")),
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
            ]))
            elements.append(totals_table)
            elements.append(Spacer(1, 0.08 * inch))

            notes = (credit_note_data.get("notes") or "").strip()
            if notes:
                elements.append(Paragraph("NOTES", custom_styles['TermsTitle']))
                elements.append(Paragraph(xml_escape(notes).replace('\n', '<br/>'), custom_styles['Terms']))
                elements.append(Spacer(1, 0.08 * inch))

            elements.append(Spacer(1, 0.1 * inch))
            footer_text = f"Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')}"
            elements.append(Paragraph(footer_text, custom_styles['Footer']))
            doc.build(elements)
            return pdf_path
        except Exception as e:
            print(f"Error generating credit note PDF: {e}")
            import traceback
            traceback.print_exc()
            raise

    def generate_temperature_log_pdf(self, logs, title_suffix=""):
        try:
            out_folder = None
            try:
                if self.output_folder:
                    out_folder = str(self.output_folder)
            except Exception:
                out_folder = None
            if not out_folder:
                possible_roots = [
                    os.path.join(os.path.expanduser("~"), "Google Drive"),
                    os.path.join(os.path.expanduser("~"), "GoogleDrive"),
                    os.path.join(os.path.expanduser("~"), "Documents"),
                    os.path.join(os.path.expanduser("~"), "Desktop")
                ]
                for root in possible_roots:
                    data_path = os.path.join(root, "HopePharmaData")
                    if os.path.exists(data_path):
                        out_folder = os.path.join(data_path, "HopePharmaTemperatureLogs")
                        break
                if not out_folder:
                    base_dir = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaData")
                    out_folder = os.path.join(base_dir, "HopePharmaTemperatureLogs")
            if not os.path.exists(out_folder):
                try:
                    os.makedirs(out_folder)
                except Exception:
                    pass

            stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            pdf_path = os.path.join(out_folder, f"temperature_log_{stamp}.pdf")

            doc = SimpleDocTemplate(
                pdf_path,
                pagesize=A4,
                topMargin=0.25 * inch,
                bottomMargin=0.25 * inch,
                leftMargin=0.35 * inch,
                rightMargin=0.35 * inch
            )
            content_width = doc.width
            elements = []
            custom_styles = self._create_custom_styles()

            logo_path = self._resolve_logo_path()
            company_block = [
                [Paragraph("HOPE PHARMA MEDICINE TRADING", custom_styles['Title'])],
                [Paragraph("Dubai, International City, Morocco Cluster, Bldg. I-16, Store 08", custom_styles['Company'])],
                [Paragraph("TRN: 100466797600003", custom_styles['Company'])]
            ]
            logo_width = 1.2 * inch
            company_width = max(content_width - (1.4 * inch), 4.0 * inch)
            company_table = Table(company_block, colWidths=[company_width])
            company_table.setStyle(TableStyle([
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 2),
            ]))
            if logo_path:
                logo_img = Image(logo_path, width=logo_width, height=logo_width)
                header = Table([[logo_img, company_table]], colWidths=[1.4 * inch, company_width])
                header.setStyle(TableStyle([
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))
                elements.append(header)
            else:
                elements.append(company_table)
            elements.append(Spacer(1, 0.12 * inch))

            title = "TEMPERATURE LOG" + (title_suffix or "")
            elements.append(Paragraph(xml_escape(title), custom_styles['InvoiceTitle']))
            elements.append(Spacer(1, 0.08 * inch))

            header_info = f"Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} by {xml_escape(getpass.getuser())}"
            elements.append(Paragraph(header_info, custom_styles['Footer']))
            elements.append(Spacer(1, 0.08 * inch))

            rows = list(logs or [])
            try:
                rows.sort(key=lambda x: str(x.get("timestamp", "")))
            except Exception:
                pass

            def fmt_temp(val):
                try:
                    if val in (None, ""):
                        return ""
                    text = str(val).strip()
                    if text.endswith("°C"):
                        text = text[:-2].strip()
                    num = float(text)
                    if num.is_integer():
                        return f"{int(num)} °C"
                    text_num = str(text)
                    if "." in text_num:
                        whole, frac = text_num.split(".", 1)
                        frac = "".join(ch for ch in frac if ch.isdigit())
                        if frac:
                            return f"{whole}.{frac[:1]} °C"
                    return f"{num} °C"
                except Exception:
                    raw = str(val or "").strip()
                    return raw if raw.endswith("°C") else raw

            aggregated_mode = bool(rows and any(f"p{i}" in rows[0] for i in range(1, 7)))

            if aggregated_mode:
                table_data = [[
                    Paragraph("<b>Date</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Day</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Time</b>", custom_styles['ItemCell']),
                    Paragraph("<b>P1 (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>P2 (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>P3 (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>P4 (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>P5 (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>P6 (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Slot</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Min (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Max (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Current (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>User</b>", custom_styles['ItemCell']),
                ]]
                for r in rows:
                    table_data.append([
                        Paragraph(xml_escape(str(r.get("date", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(str(r.get("day", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(str(r.get("time", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("p1", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("p2", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("p3", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("p4", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("p5", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("p6", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(str(r.get("slot_name", "") or "")), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("min_temp"))), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("max_temp"))), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("current_temp"))), custom_styles['ItemCell']),
                        Paragraph(xml_escape(str(r.get("recorded_by", "") or "")), custom_styles['ItemCell']),
                    ])

                col_widths = [
                    content_width * 0.09,
                    content_width * 0.09,
                    content_width * 0.07,
                    content_width * 0.055,
                    content_width * 0.055,
                    content_width * 0.055,
                    content_width * 0.055,
                    content_width * 0.055,
                    content_width * 0.055,
                    content_width * 0.08,
                    content_width * 0.07,
                    content_width * 0.07,
                    content_width * 0.08,
                    content_width * 0.11,
                ]
            else:
                table_data = [[
                    Paragraph("<b>Date/Time</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Room</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Point</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Slot</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Current (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Min (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Max (°C)</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Notes</b>", custom_styles['ItemCell']),
                    Paragraph("<b>Recorded By</b>", custom_styles['ItemCell']),
                ]]
                for r in rows:
                    ts = xml_escape(str(r.get("timestamp", "") or ""))
                    room = xml_escape(str(r.get("room_name", "") or ""))
                    point = xml_escape(str(r.get("point_name", "") or ""))
                    slot = xml_escape(str(r.get("slot_name", "") or ""))
                    notes = xml_escape(str(r.get("notes", "") or ""))
                    user = xml_escape(str(r.get("recorded_by", "") or ""))
                    table_data.append([
                        Paragraph(ts, custom_styles['ItemCell']),
                        Paragraph(room, custom_styles['ItemCell']),
                        Paragraph(point, custom_styles['ItemCell']),
                        Paragraph(slot, custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("current_temp"))), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("min_temp"))), custom_styles['ItemCell']),
                        Paragraph(xml_escape(fmt_temp(r.get("max_temp"))), custom_styles['ItemCell']),
                        Paragraph(notes.replace('\n', '<br/>'), custom_styles['ItemCell']),
                        Paragraph(user, custom_styles['ItemCell']),
                    ])

                col_widths = [
                    content_width * 0.15,
                    content_width * 0.13,
                    content_width * 0.12,
                    content_width * 0.09,
                    content_width * 0.08,
                    content_width * 0.08,
                    content_width * 0.08,
                    content_width * 0.18,
                    content_width * 0.09,
                ]
            table = Table(table_data, colWidths=col_widths, repeatRows=1)
            table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.4, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#f0f3f5")),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('ALIGN', (3 if aggregated_mode else 4, 1), (12 if aggregated_mode else 6, -1), 'RIGHT'),
            ]))
            elements.append(table)
            elements.append(Spacer(1, 0.1 * inch))

            legal = "This temperature log is generated by HopePharma system for audit and compliance use."
            elements.append(Paragraph(legal, custom_styles['Terms']))

            doc.build(elements)
            return True, pdf_path
        except Exception as e:
            return False, str(e)
    
    def generate_sales_report_pdf(self, report_data):
        """Generate enhanced sales report PDF"""
        try:
            # Create output path
            reports_folder = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)
            
            pdf_filename = f"Sales_Report_{report_data['start_date']}_to_{report_data['end_date']}_{datetime.now().strftime('%H%M%S')}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)
            
            print(f"Generating enhanced sales report PDF at: {pdf_path}")
            
            # Create document
            doc = SimpleDocTemplate(pdf_path, pagesize=letter, topMargin=0.5*inch)
            elements = []
            
            custom_styles = self._create_custom_styles()
            self._append_logo_flowables(elements, width=1.0 * inch, height=1.0 * inch, spacer=0.08 * inch)
            
            # Title and header
            elements.append(Paragraph("SALES ANALYSIS REPORT", custom_styles['Title']))
            elements.append(Paragraph(f"Period: {report_data['start_date']} to {report_data['end_date']}", custom_styles['Subtitle']))
            elements.append(Spacer(1, 0.3*inch))
            
            # Key metrics summary
            elements.append(Paragraph("KEY PERFORMANCE INDICATORS", self.styles['Heading2']))
            
            total_invoices = report_data.get('total_invoices', 0)
            total_sales = self._safe_float_conversion(report_data.get('total_sales', 0))
            total_paid = self._safe_float_conversion(report_data.get('total_paid', 0))
            total_cost = self._safe_float_conversion(report_data.get('total_cost', 0))
            total_profit = self._safe_float_conversion(report_data.get('total_profit', 0))
            outstanding_balance = self._safe_float_conversion(report_data.get('outstanding_balance', 0))
            
            # Calculate additional metrics
            collection_rate = (total_paid / total_sales * 100) if total_sales > 0 else 0
            profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
            
            metrics_data = [
                ['Total Invoices:', f"{total_invoices}", 'Collection Rate:', f"{collection_rate:.1f}%"],
                ['Total Sales:', f"AED {total_sales:,.2f}", 'Profit Margin:', f"{profit_margin:.1f}%"],
                ['Total Cost:', f"AED {total_cost:,.2f}", 'Avg. Invoice:', f"AED {total_sales/max(total_invoices,1):,.2f}"],
                ['Total Profit:', f"AED {total_profit:,.2f}", 'Outstanding:', f"AED {outstanding_balance:,.2f}"]
            ]
            
            metrics_table = Table(metrics_data, colWidths=[1.8*inch, 1.5*inch, 1.8*inch, 1.5*inch])
            metrics_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
                ('BACKGROUND', (2, 2), (3, 3), colors.lightgreen),
            ]))
            
            elements.append(metrics_table)
            elements.append(Spacer(1, 0.3*inch))
            
            # Invoices details section
            if report_data.get('invoices'):
                elements.append(Paragraph("INVOICE DETAILS", self.styles['Heading2']))
                
                invoices_data = [['Invoice ID', 'Client', 'Date', 'Total (AED)', 'Paid (AED)', 'Balance (AED)', 'Status']]
                
                for invoice in report_data['invoices'][:20]:  # Limit to first 20 for readability
                    grand_total = self._safe_float_conversion(invoice.get('grand_total', 0))
                    total_paid_inv = self._safe_float_conversion(invoice.get('total_paid', 0))
                    balance = grand_total - total_paid_inv
                    status = invoice.get('status', 'Unknown')
                    
                    invoices_data.append([
                        invoice.get('invoice_id', '')[:8] + '...',
                        invoice.get('client_name', '')[:15] + '...' if len(invoice.get('client_name', '')) > 15 else invoice.get('client_name', ''),
                        invoice.get('date', ''),
                        f"{grand_total:,.2f}",
                        f"{total_paid_inv:,.2f}",
                        f"{balance:,.2f}",
                        status
                    ])
                
                invoices_table = Table(invoices_data, colWidths=[1*inch, 1.3*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch, 0.8*inch])
                invoices_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 7),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 6),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                
                elements.append(invoices_table)
                
                if len(report_data['invoices']) > 20:
                    elements.append(Spacer(1, 0.1*inch))
                    elements.append(Paragraph(f"... and {len(report_data['invoices']) - 20} more invoices", self.styles['Normal']))
            
            # Footer
            elements.append(Spacer(1, 0.3*inch))
            footer_text = f"Report generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} | HopePharma Analytics"
            elements.append(Paragraph(footer_text, custom_styles['Footer']))
            
            # Build the PDF
            doc.build(elements)
            print(f"Enhanced sales report PDF generated successfully: {pdf_path}")
            return pdf_path
            
        except Exception as e:
            print(f"Error generating enhanced sales report PDF: {e}")
            raise

    def generate_company_report_pdf(self, report_data):
        """Generate enhanced company financial report PDF"""
        try:
            reports_folder = os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaReports")
            if not os.path.exists(reports_folder):
                os.makedirs(reports_folder)
            
            pdf_filename = f"Financial_Report_{report_data['start_date']}_to_{report_data['end_date']}_{datetime.now().strftime('%H%M%S')}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)
            
            print(f"Generating enhanced company report PDF at: {pdf_path}")
            
            doc = SimpleDocTemplate(pdf_path, pagesize=letter, topMargin=0.5*inch)
            elements = []
            
            custom_styles = self._create_custom_styles()
            self._append_logo_flowables(elements, width=1.0 * inch, height=1.0 * inch, spacer=0.08 * inch)
            
            # Title
            elements.append(Paragraph("COMPANY FINANCIAL REPORT", custom_styles['Title']))
            elements.append(Paragraph(f"Period: {report_data['start_date']} to {report_data['end_date']}", custom_styles['Subtitle']))
            elements.append(Spacer(1, 0.3*inch))
            
            # Financial Health Overview
            elements.append(Paragraph("FINANCIAL HEALTH OVERVIEW", self.styles['Heading2']))
            
            total_revenue = self._safe_float_conversion(report_data.get('total_revenue', 0))
            total_expenses = self._safe_float_conversion(report_data.get('total_expenses', 0))
            net_profit = self._safe_float_conversion(report_data.get('net_profit', 0))
            net_profit_margin = self._safe_float_conversion(report_data.get('net_profit_margin', 0))
            outstanding_balance = self._safe_float_conversion(report_data.get('outstanding_balance', 0))
            
            health_color = colors.darkgreen if net_profit > 0 else colors.darkred
            health_status = "Profitable" if net_profit > 0 else "Not Profitable"
            
            overview_data = [
                ['Financial Health:', health_status, 'Profit Margin:', f"{net_profit_margin:.1f}%"],
                ['Total Revenue:', f"AED {total_revenue:,.2f}", 'Outstanding:', f"AED {outstanding_balance:,.2f}"],
                ['Total Expenses:', f"AED {total_expenses:,.2f}", 'Net Profit/Loss:', f"AED {net_profit:,.2f}"]
            ]
            
            overview_table = Table(overview_data, colWidths=[1.8*inch, 1.5*inch, 1.8*inch, 1.5*inch])
            overview_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 0), (-1, -1), 9),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 1, colors.grey),
                ('TEXTCOLOR', (0, 0), (1, 0), health_color),
                ('BACKGROUND', (0, 2), (-1, 2), colors.lightgrey),
            ]))
            
            elements.append(overview_table)
            elements.append(Spacer(1, 0.3*inch))
            
            # Top Clients Section
            if report_data.get('top_clients'):
                elements.append(Paragraph("TOP 5 CLIENTS BY REVENUE", self.styles['Heading2']))
                
                clients_data = [['Client Name', 'Revenue (AED)', 'Percentage']]
                total_client_revenue = sum(rev for _, rev in report_data['top_clients'])
                
                for client, revenue in report_data['top_clients']:
                    percentage = (revenue / total_revenue * 100) if total_revenue > 0 else 0
                    clients_data.append([
                        client[:25] + '...' if len(client) > 25 else client,
                        f"{revenue:,.2f}",
                        f"{percentage:.1f}%"
                    ])
                
                clients_table = Table(clients_data, colWidths=[3*inch, 1.5*inch, 1*inch])
                clients_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkblue),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.lightblue),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                
                elements.append(clients_table)
                elements.append(Spacer(1, 0.3*inch))
            
            # Expense Categories
            if report_data.get('expense_categories'):
                elements.append(Paragraph("EXPENSE BREAKDOWN BY CATEGORY", self.styles['Heading2']))
                
                expense_data = [['Category', 'Amount (AED)', 'Percentage']]
                total_expenses_calc = sum(report_data['expense_categories'].values())
                
                for category, amount in sorted(report_data['expense_categories'].items(), key=lambda x: x[1], reverse=True)[:6]:
                    percentage = (amount / total_expenses_calc * 100) if total_expenses_calc > 0 else 0
                    expense_data.append([
                        category,
                        f"{amount:,.2f}",
                        f"{percentage:.1f}%"
                    ])
                
                expense_table = Table(expense_data, colWidths=[2.5*inch, 1.5*inch, 1*inch])
                expense_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.darkgreen),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('BACKGROUND', (0, 1), (-1, -1), colors.lightgreen),
                    ('GRID', (0, 0), (-1, -1), 1, colors.black),
                ]))
                
                elements.append(expense_table)
            
            # Footer
            elements.append(Spacer(1, 0.3*inch))
            footer_text = f"Financial report generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} | HopePharma Financial System"
            elements.append(Paragraph(footer_text, custom_styles['Footer']))
            
            doc.build(elements)
            print(f"Enhanced company report PDF generated successfully: {pdf_path}")
            return pdf_path
            
        except Exception as e:
            print(f"Error generating enhanced company report PDF: {e}")
            raise

    def _page_number_footer(self, canvas_obj, doc):
        """Brand footer + page number on EVERY page (used by onPage callback)."""
        canvas_obj.saveState()
        try:
            # Footer line
            canvas_obj.setStrokeColor(colors.HexColor("#2B6CB0"))
            canvas_obj.setLineWidth(0.6)
            canvas_obj.line(doc.leftMargin, 0.6 * inch,
                            doc.pagesize[0] - doc.rightMargin, 0.6 * inch)
            canvas_obj.setFont("Helvetica", 7.5)
            canvas_obj.setFillColor(colors.HexColor("#718096"))
            left_text = "HopePharma Medical Trading L.L.C. — Confidential / For Internal Use"
            canvas_obj.drawString(doc.leftMargin, 0.45 * inch, left_text)
            right_text = f"Page {doc.page} of N"
            # We don't know total pages without two-pass; use standard "Page X"
            right_text = f"Page {doc.page}"
            canvas_obj.drawRightString(doc.pagesize[0] - doc.rightMargin, 0.45 * inch, right_text)
        except Exception:
            pass
        canvas_obj.restoreState()

    def _brand_header_story(self, elements, custom_styles, report_title, subtitle_lines=None):
        """Append branded header block: logo (if present) → company info → title → subtitle."""
        # Logo + Company row (two-column table so logo stays left, company stays right)
        try:
            logo_path = self._resolve_logo_path()
            has_logo = bool(logo_path and os.path.exists(logo_path))
        except Exception:
            has_logo = False; logo_path = None
        company_cell = Paragraph(
            "<b>HOPE PHARMA MEDICAL TRADING L.L.C.</b><br/>"
            "<font size='8' color='#2D3748'>"
            "Pharmaceutical Wholesaler &amp; Distributor<br/>"
            "VAT Registration: UAE VAT REGISTERED<br/>"
            "Reporting Currency: AED (United Arab Emirates Dirham)"
            "</font>",
            ParagraphStyle(
                "CoInfo", parent=self.styles['Normal'],
                fontSize=9, leading=11, alignment=TA_RIGHT,
                textColor=colors.HexColor("#2D3748")
            )
        )
        try:
            if has_logo:
                img = Image(logo_path, width=1.0 * inch, height=1.0 * inch)
                img.hAlign = "LEFT"
            else:
                # Placeholder brand mark
                img = Paragraph(
                    "<font size='16' color='#1A365D'><b>H</b></font><font size='10' color='#2B6CB0'><b>PMT</b></font><br/>"
                    "<font size='6' color='#718096'>HopePharma</font>",
                    ParagraphStyle("logo_ph", parent=self.styles['Normal'],
                                   alignment=TA_LEFT, leading=14)
                )
            header_row = [[img, company_cell]]
            header_table = Table(header_row, colWidths=[1.4 * inch, 5.3 * inch])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ]))
            elements.append(header_table)
        except Exception:
            elements.append(Paragraph("HopePharma Medical Trading L.L.C.",
                                      custom_styles['Company']))

        # Thin rule
        try:
            rule = Table([['']], colWidths=[6.7 * inch], rowHeights=[0.03 * inch])
            rule.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#2B6CB0")),
                ('LINEBELOW', (0, 0), (-1, -1), 0, colors.white),
            ]))
            elements.append(rule)
        except Exception:
            elements.append(Spacer(1, 0.04 * inch))
        elements.append(Spacer(1, 0.08 * inch))

        # Report title
        elements.append(Paragraph(report_title, custom_styles['Title']))
        if subtitle_lines:
            for line in subtitle_lines:
                elements.append(Paragraph(line, custom_styles['Subtitle']))
        elements.append(Spacer(1, 0.14 * inch))

    def generate_trial_balance_pdf(self, trial_balance_data, output_folder=None):
        """Brand-new professional Trial Balance PDF.

        Layout:
            Brand header (logo + company info)
            Report title + period
            Summary card: Account count, Balanced status, Period totals
            Full 6-column TB table (Opening Dr/Cr | Period Dr/Cr | Closing Dr/Cr)
            Grand total row with double underline
            Footer: date, user, page numbers
        """
        try:
            reports_folder = output_folder
            if not reports_folder:
                candidates = [
                    os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaReports"),
                    os.path.join(tempfile.gettempdir(), "HopePharmaReports"),
                ]
                for cand in candidates:
                    try:
                        if not os.path.exists(cand):
                            os.makedirs(cand, exist_ok=True)
                        # Can we actually WRITE here? (sandbox blocks ~/Documents sometimes)
                        _touch = os.path.join(cand, ".write_test_" + os.urandom(4).hex())
                        try:
                            with open(_touch, 'wb') as f: f.write(b"ok")
                            os.remove(_touch)
                        except Exception:
                            continue
                        reports_folder = cand
                        break
                    except Exception:
                        continue
                if not reports_folder:
                    reports_folder = str(Path(os.path.abspath(__file__)).parent)
            os.makedirs(reports_folder, exist_ok=True)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            pdf_filename = f"Trial_Balance_{trial_balance_data.get('start_date','')}_to_{trial_balance_data.get('end_date','')}_{ts}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)

            # Use LANDSCAPE Letter because we have 6 money columns + labels.
            doc = SimpleDocTemplate(
                pdf_path, pagesize=landscape(letter),
                leftMargin=0.55 * inch, rightMargin=0.55 * inch,
                topMargin=0.55 * inch, bottomMargin=0.75 * inch,
                title="Trial Balance - HopePharma",
                author="HopePharma Medical Trading L.L.C.",
                subject=f"Trial Balance {trial_balance_data.get('start_date','')} to {trial_balance_data.get('end_date','')}",
            )
            elements = []
            cs = self._create_custom_styles()

            # --- Brand header ---
            bal = trial_balance_data.get('balanced', False)
            bal_text = ("<font color='#276749'>✔ Ledger is in BALANCE</font>"
                        if bal else
                        "<font color='#9B2C2C'>⚠ Ledger is OUT OF BALANCE</font>")
            self._brand_header_story(
                elements, cs,
                report_title="TRIAL BALANCE",
                subtitle_lines=[
                    f"Reporting Period: <b>{trial_balance_data.get('start_date','')}</b> &nbsp;→&nbsp; <b>{trial_balance_data.get('end_date','')}</b>",
                    f"Accounts with activity: <b>{trial_balance_data.get('account_count', 0)}</b> &nbsp;&nbsp;|&nbsp;&nbsp; {bal_text}",
                ]
            )

            tot = trial_balance_data.get('totals') or {}
            def _d(x):
                try: return float(x or 0)
                except Exception: return 0.0

            # --- Summary card: Period totals ---
            period_row = [
                ["Opening Dr", _d(tot.get('opening_debit')),  "Opening Cr", _d(tot.get('opening_credit'))],
                ["Period Dr",  _d(tot.get('period_debit')),   "Period Cr",  _d(tot.get('period_credit'))],
                ["Closing Dr", _d(tot.get('closing_debit')),  "Closing Cr", _d(tot.get('closing_credit'))],
            ]
            summary_data = [["Category", "AED", "Category", "AED"]]
            for r in period_row:
                summary_data.append([r[0], f"AED {r[1]:,.2f}", r[2], f"AED {r[3]:,.2f}"])
            summary_table = Table(summary_data, colWidths=[1.6 * inch, 1.8 * inch, 1.6 * inch, 1.8 * inch])
            summary_table.setStyle(TableStyle([
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1A365D")),
                ('FONTSIZE', (0, 0), (-1, -1), 8.5),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('ALIGN', (1, 1), (1, -1), 'RIGHT'),
                ('ALIGN', (3, 1), (3, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E0")),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F7FAFC")),
            ]))
            elements.append(summary_table)
            elements.append(Spacer(1, 0.18 * inch))

            # --- Full TB table (landscape, wide) ---
            header = [
                "Code", "Account",
                "Opening Dr", "Opening Cr",
                "Period Dr", "Period Cr",
                "Closing Dr", "Closing Cr",
            ]
            rows = trial_balance_data.get('rows') or []
            table_data = [header]
            ZEBRA = colors.HexColor("#F7FAFC")
            NAVY = colors.HexColor("#1A365D")
            for r in rows:
                code = str(r.get('account_code') or '')
                name = str(r.get('account_name') or '')
                def _a(v):
                    vv = _d(v)
                    if vv == 0:
                        return ""
                    return f"{vv:,.2f}"
                table_data.append([
                    code, name,
                    _a(r.get('opening_debit')),  _a(r.get('opening_credit')),
                    _a(r.get('period_debit')),   _a(r.get('period_credit')),
                    _a(r.get('closing_debit')),  _a(r.get('closing_credit')),
                ])
            # Grand totals row
            def _g(v):
                return f"{_d(tot.get(v)):,.2f}" if _d(tot.get(v)) != 0 else "0.00"
            table_data.append([
                "", "TOTALS",
                _g('opening_debit'),  _g('opening_credit'),
                _g('period_debit'),   _g('period_credit'),
                _g('closing_debit'),  _g('closing_credit'),
            ])

            col_widths = [0.6 * inch, 2.35 * inch,
                          1.05 * inch, 1.05 * inch,
                          1.05 * inch, 1.05 * inch,
                          1.05 * inch, 1.05 * inch]
            tb_table = Table(table_data, colWidths=col_widths, repeatRows=1)
            style_cmds = [
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 8.5),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('BACKGROUND', (0, 0), (-1, 0), NAVY),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTSIZE', (0, 1), (-1, -1), 8),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('ALIGN', (0, 1), (1, -1), 'LEFT'),
                ('ALIGN', (2, 1), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('GRID', (0, 0), (-1, -1), 0.3, colors.HexColor("#CBD5E0")),
                ('TOPPADDING', (0, 0), (-1, -1), 2.2),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 2.2),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                # Grand total row
                ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
                ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor("#EDF2F7")),
                ('LINEABOVE', (0, -1), (-1, -1), 0.8, NAVY),
                ('LINEBELOW', (0, -1), (-1, -1), 1.4, NAVY),
            ]
            # Zebra striping on body rows
            for i, _ in enumerate(rows, start=1):
                if i % 2 == 0:
                    style_cmds.append(('BACKGROUND', (0, i), (-1, i), ZEBRA))
            tb_table.setStyle(TableStyle(style_cmds))
            elements.append(tb_table)

            # --- Footer block ---
            elements.append(Spacer(1, 0.22 * inch))
            generated = Paragraph(
                f"<font color='#718096' size='7.5'>"
                f"Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} &nbsp;·&nbsp; "
                f"Source: General Ledger single source of truth &nbsp;·&nbsp; "
                f"Currency: AED (2-decimal Dirhams, standard rounding)"
                "</font>",
                cs['Footer']
            )
            elements.append(generated)

            doc.build(elements, onFirstPage=self._page_number_footer,
                      onLaterPages=self._page_number_footer)
            print(f"Enhanced Trial Balance PDF generated successfully: {pdf_path}")
            return pdf_path

        except Exception as e:
            print(f"Error generating enhanced Trial Balance PDF: {e}")
            raise

    def _resolve_reports_folder(self, output_folder=None):
        """Safe folder resolver for report PDFs — guarantees a writable dir.

        Tries (in order):
          1) provided output_folder
          2) ~/Documents/HopePharmaReports (standard user-facing dir)
          3) tempfile.gettempdir()/HopePharmaReports (sandbox fallback on macOS)
          4) project directory next to hope_pharma_complete.py (ultimate fallback)
        Returns absolute path to folder (guaranteed created)."""
        candidates = []
        if output_folder:
            candidates.append(output_folder)
        candidates.extend([
            os.path.join(os.path.expanduser("~"), "Documents", "HopePharmaReports"),
            os.path.join(tempfile.gettempdir(), "HopePharmaReports"),
            str(Path(os.path.abspath(__file__)).parent),
        ])
        chosen = None
        for cand in candidates:
            try:
                os.makedirs(cand, exist_ok=True)
                tst = os.path.join(cand, ".wtest_" + os.urandom(4).hex())
                try:
                    with open(tst, 'wb') as f: f.write(b"ok")
                    os.remove(tst)
                except Exception:
                    continue
                chosen = cand; break
            except Exception:
                continue
        if not chosen:
            chosen = tempfile.mkdtemp(prefix="hp_pdf_")
        return str(chosen)

    def generate_sales_report_pdf_enhanced(self, report_data, output_folder=None):
        """Enhanced SALES REPORT PDF — branded header, page numbers, zebra tables."""
        try:
            reports_folder = self._resolve_reports_folder(output_folder)
            pdf_filename = f"Sales_Report_Professional_{report_data.get('start_date','')}_to_{report_data.get('end_date','')}_{datetime.now().strftime('%H%M%S')}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)
            doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                    topMargin=0.55*inch, bottomMargin=0.75*inch,
                                    title=f"Sales Report {report_data.get('start_date','')}",
                                    author="HopePharma Medical Trading L.L.C.")
            elements = []
            cs = self._create_custom_styles()
            self._brand_header_story(
                elements, cs,
                report_title="SALES ANALYSIS REPORT",
                subtitle_lines=[f"Period: <b>{report_data.get('start_date','')}</b> → <b>{report_data.get('end_date','')}</b>"]
            )
            # KPI card
            total_sales = self._safe_float_conversion(report_data.get('total_sales', 0))
            total_profit = self._safe_float_conversion(report_data.get('total_profit', 0))
            total_invoices = report_data.get('total_invoices', 0) or 0
            collection_rate = (self._safe_float_conversion(report_data.get('total_paid', 0)) / total_sales * 100) if total_sales > 0 else 0
            profit_margin = (total_profit / total_sales * 100) if total_sales > 0 else 0
            kpi = [
                ["Total Invoices", f"{total_invoices}",       "Collection Rate", f"{collection_rate:.1f}%"],
                ["Total Sales",    f"AED {total_sales:,.2f}", "Profit Margin",   f"{profit_margin:.1f}%"],
                ["Total Profit",   f"AED {total_profit:,.2f}","Outstanding",     f"AED {self._safe_float_conversion(report_data.get('outstanding_balance',0)):,.2f}"],
            ]
            kpi_data = [["Metric", "Value", "Metric", "Value"]]
            for r in kpi: kpi_data.append(r)
            kpi_tbl = Table(kpi_data, colWidths=[1.6*inch,1.5*inch,1.6*inch,1.5*inch])
            kpi_tbl.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0), colors.HexColor("#1A365D")),
                ('TEXTCOLOR',(0,0),(-1,0), colors.whitesmoke),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('FONTSIZE',(0,0),(-1,-1),9),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('GRID',(0,0),(-1,-1),0.4, colors.HexColor("#CBD5E0")),
                ('ALIGN',(1,1),(1,-1),'RIGHT'),
                ('ALIGN',(3,1),(3,-1),'RIGHT'),
                ('BACKGROUND',(0,1),(-1,-1), colors.HexColor("#F7FAFC")),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
            ]))
            elements.append(kpi_tbl); elements.append(Spacer(1, 0.18*inch))

            if report_data.get('invoices'):
                elements.append(Paragraph("INVOICE DETAILS", ParagraphStyle("sh2", parent=self.styles['Heading2'],
                                                                              textColor=colors.HexColor("#1A365D"),
                                                                              fontSize=11, spaceBefore=2, spaceAfter=6)))
                invs = [['Invoice ID','Client','Date','Total (AED)','Paid (AED)','Balance','Status']]
                for inv in report_data['invoices']:
                    gt = self._safe_float_conversion(inv.get('grand_total', 0))
                    tp = self._safe_float_conversion(inv.get('total_paid', 0))
                    invs.append([
                        str(inv.get('invoice_id',''))[:14],
                        str(inv.get('client_name',''))[:22],
                        str(inv.get('date','')),
                        f"{gt:,.2f}", f"{tp:,.2f}", f"{gt-tp:,.2f}",
                        str(inv.get('status','Unknown'))
                    ])
                it = Table(invs, colWidths=[1*inch,1.55*inch,0.85*inch,0.9*inch,0.9*inch,0.9*inch,0.9*inch], repeatRows=1)
                it.setStyle(TableStyle([
                    ('BACKGROUND',(0,0),(-1,0), colors.HexColor("#1A365D")),
                    ('TEXTCOLOR',(0,0),(-1,0), colors.whitesmoke),
                    ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                    ('FONTSIZE',(0,0),(-1,-1),7.5),
                    ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                    ('ALIGN',(2,0),(2,-1),'CENTER'),
                    ('ALIGN',(3,0),(-1,-1),'RIGHT'),
                    ('GRID',(0,0),(-1,-1),0.3, colors.HexColor("#CBD5E0")),
                    ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
                    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ]))
                for i in range(1, len(invs)):
                    if i % 2 == 0: it.setStyle(TableStyle([('BACKGROUND',(0,i),(-1,i), colors.HexColor("#F7FAFC"))]))
                elements.append(it)
                if len(report_data['invoices']) > len(invs)-1:
                    elements.append(Spacer(1, 0.08*inch))
                    elements.append(Paragraph(f"... and more (showing first {len(invs)-1} invoices)", cs['Subtitle']))
            elements.append(Spacer(1,0.22*inch))
            footer = Paragraph(
                f"<font color='#718096' size='7.5'>"
                f"Report generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} · HopePharma Analytics · Currency AED"
                "</font>", cs['Footer'])
            elements.append(footer)
            doc.build(elements, onFirstPage=self._page_number_footer, onLaterPages=self._page_number_footer)
            print(f"Enhanced Professional Sales Report PDF: {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Error generating pro sales report PDF: {e}")
            raise

    def generate_company_report_pdf_enhanced(self, report_data, output_folder=None):
        """Enhanced COMPANY FINANCIAL REPORT PDF — branded header, page numbers, KPIs, balanced sections."""
        try:
            reports_folder = self._resolve_reports_folder(output_folder)
            pdf_filename = f"Company_Financial_Report_Pro_{report_data.get('start_date','')}_to_{report_data.get('end_date','')}_{datetime.now().strftime('%H%M%S')}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)
            doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                    topMargin=0.55*inch, bottomMargin=0.75*inch,
                                    title=f"Company Financial Report {report_data.get('start_date','')}",
                                    author="HopePharma Medical Trading L.L.C.")
            elements = []
            cs = self._create_custom_styles()
            self._brand_header_story(
                elements, cs,
                report_title="COMPANY FINANCIAL REPORT",
                subtitle_lines=[f"Period: <b>{report_data.get('start_date','')}</b> → <b>{report_data.get('end_date','')}</b>"]
            )
            # Health overview KPIs
            tr = self._safe_float_conversion(report_data.get('total_revenue', 0))
            te = self._safe_float_conversion(report_data.get('total_expenses', 0))
            np_ = self._safe_float_conversion(report_data.get('net_profit', 0))
            health = ("<font color='#276749'><b>✔ Profitable</b></font>" if np_ > 0 else
                      "<font color='#9B2C2C'><b>⚠ Not Profitable</b></font>")
            overview = [
                ["Total Revenue",   f"AED {tr:,.2f}",   "Total Expenses", f"AED {te:,.2f}"],
                ["Net Profit",      f"AED {np_:,.2f}",  "Net Margin",     f"{(np_/tr*100 if tr>0 else 0):.1f}%"],
                ["Outstanding AR",  f"AED {self._safe_float_conversion(report_data.get('outstanding_balance',0)):,.2f}",
                 "Overall Health",  health],
            ]
            od = [["Metric", "Value", "Metric", "Value"]]
            for r in overview: od.append(r)
            ot = Table(od, colWidths=[1.6*inch,1.5*inch,1.6*inch,1.5*inch])
            ot.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0), colors.HexColor("#1A365D")),
                ('TEXTCOLOR',(0,0),(-1,0), colors.whitesmoke),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('FONTSIZE',(0,0),(-1,-1),9),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('GRID',(0,0),(-1,-1),0.4, colors.HexColor("#CBD5E0")),
                ('ALIGN',(1,1),(1,-1),'RIGHT'),
                ('ALIGN',(3,1),(3,-1),'RIGHT'),
                ('BACKGROUND',(0,1),(-1,-1), colors.HexColor("#F7FAFC")),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
            ]))
            elements.append(ot); elements.append(Spacer(1, 0.18*inch))

            # Revenue by client (top N)
            if report_data.get('top_clients'):
                elements.append(Paragraph("TOP REVENUE CLIENTS", ParagraphStyle("h_1", parent=self.styles['Heading2'],
                                                                                  textColor=colors.HexColor("#1A365D"),
                                                                                  fontSize=11, spaceBefore=2, spaceAfter=6)))
                top = [["#","Client","Revenue (AED)","% of Total"]]
                clients_sorted = sorted(list(report_data['top_clients'].items()),
                                        key=lambda x: self._safe_float_conversion(x[1]), reverse=True)[:15]
                for i,(cli,rev) in enumerate(clients_sorted, start=1):
                    rv = self._safe_float_conversion(rev)
                    pct = (rv / tr * 100) if tr > 0 else 0
                    top.append([str(i), str(cli)[:30], f"{rv:,.2f}", f"{pct:.1f}%"])
                t = Table(top, colWidths=[0.35*inch, 3.25*inch, 1.3*inch, 1.0*inch], repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND',(0,0),(-1,0), colors.HexColor("#2B6CB0")),
                    ('TEXTCOLOR',(0,0),(-1,0), colors.whitesmoke),
                    ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                    ('FONTSIZE',(0,0),(-1,-1),8),
                    ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                    ('ALIGN',(0,0),(0,-1),'CENTER'),
                    ('ALIGN',(2,0),(-1,-1),'RIGHT'),
                    ('GRID',(0,0),(-1,-1),0.3, colors.HexColor("#CBD5E0")),
                    ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
                    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ]))
                for i in range(1, len(top)):
                    if i % 2 == 0: t.setStyle(TableStyle([('BACKGROUND',(0,i),(-1,i), colors.HexColor("#F7FAFC"))]))
                elements.append(t); elements.append(Spacer(1, 0.15*inch))

            # Expense breakdown
            if report_data.get('expenses'):
                elements.append(Paragraph("EXPENSES BREAKDOWN", ParagraphStyle("h_2", parent=self.styles['Heading2'],
                                                                                 textColor=colors.HexColor("#1A365D"),
                                                                                 fontSize=11, spaceBefore=2, spaceAfter=6)))
                exp_rows = [["Expense Category","AED","% of Total"]]
                for exp_cat, exp_amt in sorted(list(report_data['expenses'].items()),
                                               key=lambda x: self._safe_float_conversion(x[1]), reverse=True):
                    av = self._safe_float_conversion(exp_amt)
                    pct = (av / te * 100) if te > 0 else 0
                    exp_rows.append([str(exp_cat), f"{av:,.2f}", f"{pct:.1f}%"])
                et = Table(exp_rows, colWidths=[3.25*inch, 1.3*inch, 1.0*inch], repeatRows=1)
                et.setStyle(TableStyle([
                    ('BACKGROUND',(0,0),(-1,0), colors.HexColor("#276749")),
                    ('TEXTCOLOR',(0,0),(-1,0), colors.whitesmoke),
                    ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                    ('FONTSIZE',(0,0),(-1,-1),8),
                    ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                    ('ALIGN',(1,0),(-1,-1),'RIGHT'),
                    ('GRID',(0,0),(-1,-1),0.3, colors.HexColor("#CBD5E0")),
                    ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
                    ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ]))
                for i in range(1, len(exp_rows)):
                    if i % 2 == 0: et.setStyle(TableStyle([('BACKGROUND',(0,i),(-1,i), colors.HexColor("#F7FAFC"))]))
                elements.append(et)

            elements.append(Spacer(1,0.22*inch))
            footer = Paragraph(
                f"<font color='#718096' size='7.5'>"
                f"Report generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} · HopePharma Financials · Currency AED"
                "</font>", cs['Footer'])
            elements.append(footer)
            doc.build(elements, onFirstPage=self._page_number_footer, onLaterPages=self._page_number_footer)
            print(f"Enhanced Professional Company Report PDF: {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Error generating pro company report PDF: {e}")
            raise

    # =========================================================================
    # NEW: BALANCE SHEET — Professional PDF
    # =========================================================================
    def generate_balance_sheet_pdf(self, bs_data, output_folder=None):
        try:
            reports_folder = self._resolve_reports_folder(output_folder)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            as_of = str(bs_data.get('as_of') or '')
            pdf_filename = f"Balance_Sheet_as_of_{as_of}_{ts}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)

            doc = SimpleDocTemplate(
                pdf_path, pagesize=letter,
                leftMargin=0.6*inch, rightMargin=0.6*inch,
                topMargin=0.55*inch, bottomMargin=0.75*inch,
                title=f"Balance Sheet - HopePharma {as_of}",
                author="HopePharma Medical Trading L.L.C.",
            )
            elements = []
            cs = self._create_custom_styles()

            sb = bs_data.get('summary_banner') or {}
            def _d(x):
                try: return float(x or 0)
                except Exception: return 0.0
            is_balanced = bool(sb.get('is_balanced'))
            bal_text = ("<font color='#276749'>✔ Statement is BALANCED</font>" if is_balanced
                        else "<font color='#9B2C2C'>⚠ Statement is OUT OF BALANCE</font>")
            total_assets = _d(sb.get('grand_total'))
            self._brand_header_story(elements, cs,
                report_title="STATEMENT OF FINANCIAL POSITION (BALANCE SHEET)",
                subtitle_lines=[
                    f"As of: <b>{as_of}</b> &nbsp;&nbsp;|&nbsp;&nbsp; US GAAP Classified (ASC 210) &nbsp;&nbsp;|&nbsp;&nbsp; {bal_text}",
                    f"Total Assets = <b>AED {total_assets:,.2f}</b>  ·  Working Capital = <b>AED {_d(sb.get('working_capital')):,.2f}</b>",
                ])

            # KPI summary card
            kpi_data = [["KPI", "Value", "KPI", "Value"]]
            kpi_rows = [
                ["Total Assets",            f"AED {total_assets:,.2f}",
                 "Total Liabilities",       f"AED {_d(sb.get('liabilities')):,.2f}"],
                ["Current Assets",          f"AED {_d(sb.get('current_assets')):,.2f}",
                 "Current Liabilities",     f"AED {_d(sb.get('current_liabilities')):,.2f}"],
                ["Non-Current Assets",      f"AED {_d(sb.get('noncurrent_assets')):,.2f}",
                 "Non-Current Liabilities", f"AED {_d(sb.get('noncurrent_liabilities')):,.2f}"],
                ["Total Equity",            f"AED {_d(sb.get('equity')):,.2f}",
                 "Working Capital",         f"AED {_d(sb.get('working_capital')):,.2f}"],
            ]
            for r in kpi_rows: kpi_data.append(r)
            kpi_tbl = Table(kpi_data, colWidths=[1.7*inch,1.5*inch,1.7*inch,1.5*inch])
            kpi_tbl.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#1A365D")),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('FONTSIZE',(0,0),(-1,-1),9),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('ALIGN',(1,1),(1,-1),'RIGHT'), ('ALIGN',(3,1),(3,-1),'RIGHT'),
                ('GRID',(0,0),(-1,-1),0.4,colors.HexColor("#CBD5E0")),
                ('BACKGROUND',(0,1),(-1,-1),colors.HexColor("#F7FAFC")),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
            ]))
            elements.append(kpi_tbl); elements.append(Spacer(1,0.2*inch))

            NAVY = colors.HexColor("#1A365D")
            ZEBRA = colors.HexColor("#F7FAFC")
            # Build BS table: Section | Account Label | Amount
            def _money(v):
                vv = _d(v)
                if abs(vv) < 0.005: return "—"
                sign = "-" if vv < 0 else ""
                return f"{sign}AED {abs(vv):,.2f}"
            def _bold(v):
                vv = _d(v)
                sign = "-" if vv < 0 else ""
                return f"{sign}AED {abs(vv):,.2f}"
            table_data = [["Section", "Line Item", "Amount (AED)"]]
            lines = bs_data.get('lines') or {}
            # ASSETS
            table_data.append(["ASSETS", "", ""])
            table_data.append(["", "CURRENT ASSETS:", ""])
            assets = lines.get('assets', {})
            seen_totals = set()
            for label, val in assets.items():
                if label.startswith('__'):
                    if label == "__current_assets_total":
                        table_data.append(["", "  Total Current Assets", _bold(val)])
                    elif label == "__noncurrent_assets_total":
                        table_data.append(["", "  Total Non-Current Assets", _bold(val)])
                    seen_totals.add(label)
                    continue
                if label == "total_assets":
                    continue
                if label.startswith("16") and ("Gross Cost" in label or label.startswith("1600 ")):
                    table_data.append(["", "NON-CURRENT ASSETS:", ""])
                table_data.append(["", f"    {label}", _money(val)])
            table_data.append(["", "TOTAL ASSETS", _bold(assets.get('total_assets',0))])
            table_data.append(["", "", ""])
            # LIABILITIES
            liab = lines.get('liabilities', {})
            table_data.append(["LIABILITIES", "", ""])
            table_data.append(["", "CURRENT LIABILITIES:", ""])
            for label, val in liab.items():
                if label.startswith('__'):
                    if label == "__current_liabilities_total":
                        table_data.append(["", "  Total Current Liabilities", _bold(val)])
                    elif label == "__noncurrent_liabilities_total":
                        table_data.append(["", "  Total Non-Current Liabilities", _bold(val)])
                    continue
                if label == "total_liabilities": continue
                if label.startswith("24") or label.startswith("2500"):
                    table_data.append(["", "NON-CURRENT LIABILITIES:", ""])
                table_data.append(["", f"    {label}", _money(val)])
            table_data.append(["", "TOTAL LIABILITIES", _bold(liab.get('total_liabilities',0))])
            table_data.append(["", "", ""])
            # EQUITY
            eq = lines.get('equity', {})
            table_data.append(["EQUITY", "", ""])
            for label, val in eq.items():
                if label == "total_equity": continue
                table_data.append(["", f"    {label}", _money(val)])
            table_data.append(["", "TOTAL EQUITY", _bold(eq.get('total_equity',0))])
            table_data.append(["", "", ""])
            # Grand total: L+E
            table_data.append(["TOTAL LIABILITIES + EQUITY", "", _bold(lines.get('total_le', 0))])
            total_rows = len(table_data)
            bs_table = Table(table_data, colWidths=[1.3*inch, 4.1*inch, 1.9*inch], repeatRows=1)
            style_cmds = [
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('BACKGROUND',(0,0),(-1,0),NAVY),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('ALIGN',(0,0),(-1,0),'CENTER'),
                ('FONTSIZE',(0,0),(-1,0),9),
                ('FONTSIZE',(0,1),(-1,-1),8.3),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                ('ALIGN',(2,1),(2,-1),'RIGHT'),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('GRID',(0,0),(-1,-1),0.3,colors.HexColor("#CBD5E0")),
                ('TOPPADDING',(0,0),(-1,-1),2.1),('BOTTOMPADDING',(0,0),(-1,-1),2.1),
                ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
            ]
            # Bold-section rows: first-letter header (section titles like ASSETS/LIABILITIES/EQUITY/TOTAL)
            for i,row in enumerate(table_data[1:], start=1):
                label0 = (row[0] or "").strip()
                label1 = (row[1] or "").strip()
                if label0:
                    style_cmds.append(('FONTNAME',(0,i),(2,i),'Helvetica-Bold'))
                    style_cmds.append(('BACKGROUND',(0,i),(2,i),colors.HexColor("#EDF2F7")))
                elif label1 in ("TOTAL ASSETS","TOTAL LIABILITIES","TOTAL EQUITY",
                                "TOTAL LIABILITIES + EQUITY"):
                    style_cmds.append(('FONTNAME',(0,i),(2,i),'Helvetica-Bold'))
                    style_cmds.append(('BACKGROUND',(0,i),(2,i),colors.HexColor("#EDF2F7")))
                    style_cmds.append(('LINEABOVE',(0,i),(2,i),0.7,NAVY))
                    style_cmds.append(('LINEBELOW',(0,i),(2,i),1.2,NAVY))
                elif label1 in ("CURRENT ASSETS:","NON-CURRENT ASSETS:",
                                "CURRENT LIABILITIES:","NON-CURRENT LIABILITIES:"):
                    style_cmds.append(('FONTNAME',(1,i),(1,i),'Helvetica-Bold'))
                    style_cmds.append(('TEXTCOLOR',(1,i),(1,i),NAVY))
                elif label1.startswith("  Total"):
                    style_cmds.append(('FONTNAME',(1,i),(2,i),'Helvetica-Bold'))
                    style_cmds.append(('LINEABOVE',(1,i),(2,i),0.5,colors.HexColor("#718096")))
                if i % 2 == 0 and not label0:
                    style_cmds.append(('BACKGROUND',(0,i),(2,i),ZEBRA))
            bs_table.setStyle(TableStyle(style_cmds))
            elements.append(bs_table)

            # Diagnostics line
            elements.append(Spacer(1,0.18*inch))
            diag = bs_data.get('diagnostics') or {}
            summ = diag.get('summary','')
            if summ:
                elements.append(Paragraph(
                    f"<font size='8.5' color='#2D3748'><b>Integrity:</b></font> "
                    f"<font size='8.5' color='#4A5568'>{summ}</font>",
                    ParagraphStyle('diag', parent=cs['Normal'], leading=11)))
            elements.append(Spacer(1,0.1*inch))
            generated = Paragraph(
                f"<font color='#718096' size='7.5'>"
                f"Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} · Source: General Ledger (single source of truth) · Currency AED"
                "</font>", cs['Footer'])
            elements.append(generated)
            doc.build(elements, onFirstPage=self._page_number_footer, onLaterPages=self._page_number_footer)
            print(f"[OK] Balance Sheet PDF → {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Error generating Balance Sheet PDF: {e}"); raise

    # =========================================================================
    # NEW: INCOME STATEMENT (P&L) — Professional PDF
    # =========================================================================
    def generate_income_statement_pdf(self, pnl_data, output_folder=None):
        try:
            reports_folder = self._resolve_reports_folder(output_folder)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            pdf_filename = (f"Income_Statement_{pnl_data.get('period',{}).get('start','')}"
                            f"_to_{pnl_data.get('period',{}).get('end','')}_{ts}.pdf")
            pdf_path = os.path.join(reports_folder, pdf_filename)
            doc = SimpleDocTemplate(pdf_path, pagesize=letter,
                                    leftMargin=0.6*inch, rightMargin=0.6*inch,
                                    topMargin=0.55*inch, bottomMargin=0.75*inch,
                                    title="Income Statement - HopePharma",
                                    author="HopePharma Medical Trading L.L.C.")
            elements = []; cs = self._create_custom_styles()
            sb = pnl_data.get('summary_banner') or {}
            def _d(x):
                try: return float(x or 0)
                except Exception: return 0.0
            ni = _d(sb.get('grand_total'))
            status = ("Profit (Net Income Positive)" if ni > 0
                      else ("Net Loss" if ni < 0 else "Break-even"))
            ni_color = "#276749" if ni > 0 else ("#9B2C2C" if ni < 0 else "#2D3748")
            period_start = str(pnl_data.get('period',{}).get('start',''))
            period_end   = str(pnl_data.get('period',{}).get('end',''))
            self._brand_header_story(elements, cs,
                report_title="STATEMENT OF COMPREHENSIVE INCOME (PROFIT & LOSS)",
                subtitle_lines=[
                    f"Period: <b>{period_start}</b> → <b>{period_end}</b> &nbsp;&nbsp;|&nbsp;&nbsp; US GAAP / IAS 1",
                    f"Net Income for Period: <b><font color='{ni_color}'>AED {ni:,.2f}</font> &nbsp; ({status})</b>"
                ])

            # KPI card: 4 metrics
            kpi_data = [["Metric", "Value", "Metric", "Value"]]
            gm = _d(sb.get('gross_margin_pct'))
            eps = _d(sb.get('eps'))
            tci = _d(sb.get('total_comprehensive_income'))
            ibt = _d(sb.get('income_before_tax'))
            kpi_rows = [
                ["Net Sales (Topline)",       f"AED {_d(pnl_data.get('lines',{}).get('net_sales',{}).get('current',0)):,.2f}",
                 "Gross Profit",             f"AED {_d(pnl_data.get('lines',{}).get('gross_profit',{}).get('current',0)):,.2f}"],
                ["Gross Margin (%)",         f"{gm:.1f}%",
                 "EBIT (Operating Income)",  f"AED {_d(sb.get('operating_income')):,.2f}"],
                ["Income Before Tax",        f"AED {ibt:,.2f}",
                 "Net Income (Bottom Line)", f"AED {ni:,.2f}"],
                ["Total Comprehensive I/S",  f"AED {tci:,.2f}",
                 "EPS (Basic)",              f"AED {eps:,.4f}"],
            ]
            for r in kpi_rows: kpi_data.append(r)
            kpi_tbl = Table(kpi_data, colWidths=[1.7*inch,1.5*inch,1.7*inch,1.5*inch])
            kpi_tbl.setStyle(TableStyle([
                ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#1A365D")),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('FONTSIZE',(0,0),(-1,-1),9),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('ALIGN',(1,1),(1,-1),'RIGHT'), ('ALIGN',(3,1),(3,-1),'RIGHT'),
                ('GRID',(0,0),(-1,-1),0.4,colors.HexColor("#CBD5E0")),
                ('BACKGROUND',(0,1),(-1,-1),colors.HexColor("#F7FAFC")),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
            ]))
            elements.append(kpi_tbl); elements.append(Spacer(1,0.2*inch))

            # P&L table: Label | Current | Prior | %Δ
            rows = pnl_data.get('table_rows') or []
            NAVY = colors.HexColor("#1A365D")
            ZEBRA = colors.HexColor("#F7FAFC")
            def _pv(v):
                try: return float(v or 0)
                except Exception: return 0.0
            def _m(v):
                vv = _pv(v); sign = "-" if vv < 0 else ""
                if abs(vv) < 0.005: return "—"
                return f"{sign}AED {abs(vv):,.2f}"
            def _pc(v):
                vv = _pv(v)
                if abs(vv) < 0.005: return ""
                sign = "+" if vv > 0 else ""
                return f"{sign}{vv:.1f}%"
            pnl_tbl = [["Line Item", "Current Period", "Prior Period", "Change %"]]
            for r in rows:
                pnl_tbl.append([
                    r.get('label',''),
                    _m(r.get('current')),
                    _m(r.get('prior')),
                    _pc(r.get('pct_change'))
                ])
            pnl_table = Table(pnl_tbl, colWidths=[3.1*inch, 1.45*inch, 1.45*inch, 1.0*inch], repeatRows=1)
            style_cmds = [
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('BACKGROUND',(0,0),(-1,0),NAVY),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('ALIGN',(0,0),(-1,0),'CENTER'),
                ('FONTSIZE',(0,0),(-1,0),9),
                ('FONTSIZE',(0,1),(-1,-1),8.5),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('ALIGN',(1,1),(-1,-1),'RIGHT'),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('GRID',(0,0),(-1,-1),0.3,colors.HexColor("#CBD5E0")),
                ('TOPPADDING',(0,0),(-1,-1),2.2),('BOTTOMPADDING',(0,0),(-1,-1),2.2),
                ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
            ]
            for i,r in enumerate(rows, start=1):
                tag = r.get('tag')
                lab = r.get('label','')
                if tag in ('section', 'gross_profit', 'net_profit', 'total_row'):
                    style_cmds.append(('FONTNAME',(0,i),(-1,i),'Helvetica-Bold'))
                    style_cmds.append(('BACKGROUND',(0,i),(-1,i),colors.HexColor("#EDF2F7")))
                if tag == 'gross_profit':
                    style_cmds.append(('LINEABOVE',(0,i),(-1,i),0.55,NAVY))
                if tag == 'net_profit':
                    style_cmds.append(('LINEABOVE',(0,i),(-1,i),0.8,NAVY))
                    style_cmds.append(('LINEBELOW',(0,i),(-1,i),1.3,NAVY))
                    style_cmds.append(('TEXTCOLOR',(0,i),(-1,i),colors.HexColor(ni_color)))
                if tag == 'total_row':
                    style_cmds.append(('LINEABOVE',(0,i),(-1,i),0.5,colors.HexColor("#718096")))
                if i % 2 == 0 and tag is None:
                    style_cmds.append(('BACKGROUND',(0,i),(-1,i),ZEBRA))
            pnl_table.setStyle(TableStyle(style_cmds))
            elements.append(pnl_table)

            elements.append(Spacer(1,0.18*inch))
            generated = Paragraph(
                f"<font color='#718096' size='7.5'>"
                f"Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} · Source: General Ledger · Currency AED"
                "</font>", cs['Footer'])
            elements.append(generated)
            doc.build(elements, onFirstPage=self._page_number_footer, onLaterPages=self._page_number_footer)
            print(f"[OK] Income Statement PDF → {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Error generating Income Statement PDF: {e}"); raise

    # =========================================================================
    # NEW: AR AGING — Professional PDF
    # =========================================================================
    def generate_ar_aging_pdf(self, ar_data, output_folder=None):
        try:
            reports_folder = self._resolve_reports_folder(output_folder)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            as_of = str(ar_data.get('as_of') or '')
            pdf_filename = f"AR_Aging_as_of_{as_of}_{ts}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)
            doc = SimpleDocTemplate(pdf_path, pagesize=landscape(letter),
                                    leftMargin=0.55*inch, rightMargin=0.55*inch,
                                    topMargin=0.55*inch, bottomMargin=0.75*inch,
                                    title="AR Aging - HopePharma",
                                    author="HopePharma Medical Trading L.L.C.")
            elements = []; cs = self._create_custom_styles()
            sb = ar_data.get('summary_banner') or {}
            def _d(x):
                try: return float(x or 0)
                except Exception: return 0.0
            total = _d(sb.get('grand_total'))
            avg = _d(sb.get('avg'))
            cnt = int(sb.get('overdue_invoice_count', 0) or 0)
            ninety_plus = _d(sb.get('ninety_plus'))
            self._brand_header_story(elements, cs,
                report_title="ACCOUNTS RECEIVABLE AGING SCHEDULE",
                subtitle_lines=[
                    f"As of: <b>{as_of}</b> &nbsp;&nbsp;|&nbsp;&nbsp; {cnt} unpaid invoice(s) &nbsp;&nbsp;|&nbsp;&nbsp; Grand Total Owed: <b>AED {total:,.2f}</b>",
                    f"Average open balance per customer: AED {avg:,.2f}  ·  91+ days (High Risk): AED {ninety_plus:,.2f}"
                ])

            # Buckets KPI row
            buckets = ar_data.get('buckets') or {}
            order = ar_data.get('bucket_keys_order') or list(buckets.keys())
            NAVY = colors.HexColor("#1A365D")
            AMBER = colors.HexColor("#C05621")
            RED = colors.HexColor("#9B2C2C")
            GREEN = colors.HexColor("#2F855A")
            bucket_bg = ["#2F855A", "#C05621", "#B7791F", "#9B2C2C", "#742A2A"]
            kpi_cols = len(order)
            bucket_data = [order]
            bucket_values = [f"AED {_d(buckets.get(k,0)):,.2f}" for k in order]
            bucket_data.append(bucket_values)
            cw = [(9.3 / max(kpi_cols,1)) * inch for _ in order]
            bt = Table(bucket_data, colWidths=cw)
            bstyle = [
                ('BACKGROUND',(0,0),(-1,0),NAVY),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('FONTSIZE',(0,0),(-1,-1),9),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('FONTNAME',(0,1),(-1,1),'Helvetica-Bold'),
                ('GRID',(0,0),(-1,-1),0.4,colors.HexColor("#CBD5E0")),
                ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
            ]
            for i in range(len(order)):
                bg = bucket_bg[i] if i < len(bucket_bg) else NAVY
                bstyle.append(('BACKGROUND',(i,1),(i,1),colors.HexColor(bg)))
                bstyle.append(('TEXTCOLOR',(i,1),(i,1),colors.whitesmoke))
            bt.setStyle(TableStyle(bstyle))
            elements.append(bt); elements.append(Spacer(1, 0.2*inch))

            # Detail table: Invoice | Client | Due Date | Days Ovd | Balance | Bucket
            ZEBRA = colors.HexColor("#F7FAFC")
            detail_rows = ar_data.get('rows') or []
            def _bucket_for_days(days):
                if days <= 0: return "Current (not overdue)"
                if days <= 30: return "0-30 Days Overdue"
                if days <= 60: return "31-60 Days Overdue"
                if days <= 90: return "61-90 Days Overdue"
                return "91+ Days Overdue"
            table_data = [["Invoice #", "Customer", "Due Date", "Days Overdue", "Balance (AED)", "Aging Bucket"]]
            total_check = 0.0
            for r in detail_rows:
                bal = _d(r.get('balance')); total_check += bal
                days = int(r.get('days_overdue', 0) or 0)
                table_data.append([
                    str(r.get('invoice_no','')),
                    str(r.get('client','')),
                    str(r.get('due_date','')),
                    f"{days}" if days > 0 else "Current",
                    f"AED {bal:,.2f}",
                    _bucket_for_days(days)
                ])
            table_data.append(["", "", "", "TOTAL", f"AED {total_check:,.2f}", ""])
            col_w = [1.15*inch, 2.2*inch, 1.0*inch, 1.1*inch, 1.35*inch, 1.7*inch]
            dt = Table(table_data, colWidths=col_w, repeatRows=1)
            style_cmds = [
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('BACKGROUND',(0,0),(-1,0),NAVY),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('ALIGN',(0,0),(-1,0),'CENTER'),
                ('FONTSIZE',(0,0),(-1,0),9),
                ('FONTSIZE',(0,1),(-1,-1),8),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('ALIGN',(3,1),(4,-1),'RIGHT'),
                ('ALIGN',(2,1),(2,-1),'CENTER'),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('GRID',(0,0),(-1,-1),0.3,colors.HexColor("#CBD5E0")),
                ('TOPPADDING',(0,0),(-1,-1),2),('BOTTOMPADDING',(0,0),(-1,-1),2),
                ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
                ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                ('BACKGROUND',(0,-1),(-1,-1),colors.HexColor("#EDF2F7")),
                ('LINEABOVE',(0,-1),(-1,-1),0.8,NAVY),
                ('LINEBELOW',(0,-1),(-1,-1),1.3,NAVY),
            ]
            for i,r in enumerate(detail_rows, start=1):
                days = int(r.get('days_overdue', 0) or 0)
                if days > 90:
                    style_cmds.append(('TEXTCOLOR',(0,i),(-1,i),colors.HexColor("#9B2C2C")))
                    style_cmds.append(('FONTNAME',(5,i),(5,i),'Helvetica-Bold'))
                if i % 2 == 0:
                    style_cmds.append(('BACKGROUND',(0,i),(-1,i),ZEBRA))
            dt.setStyle(TableStyle(style_cmds))
            elements.append(dt)

            elements.append(Spacer(1,0.18*inch))
            generated = Paragraph(
                f"<font color='#718096' size='7.5'>"
                f"Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} · Source: Accounts Receivable (Invoice Ledger) · Currency AED"
                "</font>", cs['Footer'])
            elements.append(generated)
            doc.build(elements, onFirstPage=self._page_number_footer, onLaterPages=self._page_number_footer)
            print(f"[OK] AR Aging PDF → {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Error generating AR Aging PDF: {e}"); raise

    # =========================================================================
    # NEW: CASH FLOW STATEMENT — Professional PDF
    # =========================================================================
    def generate_cash_flow_pdf(self, cf_data, output_folder=None):
        try:
            reports_folder = self._resolve_reports_folder(output_folder)
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            period_start = str(cf_data.get('period',{}).get('start',''))
            period_end   = str(cf_data.get('period',{}).get('end',''))
            pdf_filename = f"Cash_Flow_Statement_{period_start}_to_{period_end}_{ts}.pdf"
            pdf_path = os.path.join(reports_folder, pdf_filename)
            doc = SimpleDocTemplate(pdf_path, pagesize=landscape(letter),
                                    leftMargin=0.55*inch, rightMargin=0.55*inch,
                                    topMargin=0.55*inch, bottomMargin=0.75*inch,
                                    title="Cash Flow Statement - HopePharma",
                                    author="HopePharma Medical Trading L.L.C.")
            elements = []; cs = self._create_custom_styles()
            sb = cf_data.get('summary_banner') or {}
            def _d(x):
                try: return float(x or 0)
                except Exception: return 0.0
            opcf = _d(sb.get('operating_cash_flow'))
            invcf = _d(sb.get('investing_cash_flow'))
            fincf = _d(sb.get('financing_cash_flow'))
            net = _d(sb.get('grand_total'))
            beg_cash = _d(sb.get('beginning_cash'))
            end_cash = _d(sb.get('closing_cash'))
            fcf = _d(sb.get('free_cash_flow'))
            runway = _d(sb.get('cash_runway_days'))
            self._brand_header_story(elements, cs,
                report_title="CONSOLIDATED STATEMENT OF CASH FLOWS",
                subtitle_lines=[
                    f"Period: <b>{period_start}</b> → <b>{period_end}</b> &nbsp;&nbsp;|&nbsp;&nbsp; US GAAP ASC 230 — Indirect Method",
                    (f"Net Δ Cash: <b>{'+' if net>=0 else ''}AED {net:,.2f}</b> &nbsp;·&nbsp; "
                     f"Beginning: AED {beg_cash:,.2f} → Ending: <b>AED {end_cash:,.2f}</b>"),
                ])

            # KPI card — Operating / Investing / Financing / Ending
            NAVY = colors.HexColor("#1A365D")
            GREEN = colors.HexColor("#2F855A")
            RED = colors.HexColor("#9B2C2C")
            AMBER = colors.HexColor("#C05621")
            def _val(v): return ("+" if v>=0 else "") + f"AED {v:,.2f}"
            kpi_data = [["Operating CF", "Investing CF", "Financing CF", "Net Δ Cash", "Free Cash Flow", "Ending Cash"]]
            kpi_data.append([_val(opcf), _val(invcf), _val(fincf), _val(net), _val(fcf), f"AED {end_cash:,.2f}"])
            cw = [1.55*inch]*6
            kt = Table(kpi_data, colWidths=cw)
            def _bg(v): return GREEN if v >= 0 else RED
            kstyle = [
                ('BACKGROUND',(0,0),(-1,0),NAVY),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('ALIGN',(0,0),(-1,-1),'CENTER'),
                ('FONTSIZE',(0,0),(-1,-1),9),
                ('FONTNAME',(0,1),(-1,1),'Helvetica-Bold'),
                ('GRID',(0,0),(-1,-1),0.4,colors.HexColor("#CBD5E0")),
                ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
            ]
            values = [opcf, invcf, fincf, net, fcf, end_cash]
            for i,v in enumerate(values):
                kstyle.append(('BACKGROUND',(i,1),(i,1),_bg(v)))
                kstyle.append(('TEXTCOLOR',(i,1),(i,1),colors.whitesmoke))
            kt.setStyle(TableStyle(kstyle))
            elements.append(kt); elements.append(Spacer(1, 0.2*inch))

            # CF table rows — ASC 230 indirect layout
            rows = cf_data.get('table_rows') or []
            def _m(v):
                vv = _d(v)
                if abs(vv) < 0.005: return "—"
                sign = "+" if vv > 0 else ""
                return f"{sign}AED {abs(vv):,.2f}"
            tbl_data = [["#", "Line Item (US GAAP ASC 230)", "Amount (AED)"]]
            for r in rows:
                tbl_data.append(["", r.get('label',''), _m(r.get('amount'))])
            # Bridge + Tie to BS
            tbl_data.append(["", "", ""])
            tbl_data.append(["", "CASH & EQUIVALENTS, BEGINNING OF PERIOD", f"AED {beg_cash:,.2f}"])
            tbl_data.append(["", f"  + Net Change in Cash during Period", _m(net)])
            tbl_data.append(["", "CASH & EQUIVALENTS, END OF PERIOD", f"AED {end_cash:,.2f}"])
            tie = _d(cf_data.get('lines',{}).get('balance_sheet_cash_tie'))
            tbl_data.append(["", "Less: Cross-Check to Balance Sheet (1000 + 1100)", f"AED {tie:,.2f}"])
            tbl_data.append(["", "Memo: Free Cash Flow (OpCF − CapEx)", f"AED {fcf:,.2f}"])
            tbl_data.append(["", f"Memo: Estimated Cash Runway (est. daily outflows)", f"{runway:,.0f} days"])
            cfs = Table(tbl_data, colWidths=[0.4*inch, 6.2*inch, 2.7*inch], repeatRows=1)
            ZEBRA = colors.HexColor("#F7FAFC")
            style_cmds = [
                ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                ('BACKGROUND',(0,0),(-1,0),NAVY),
                ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                ('ALIGN',(0,0),(-1,0),'CENTER'),
                ('FONTSIZE',(0,0),(-1,0),9),
                ('FONTSIZE',(0,1),(-1,-1),8.3),
                ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                ('ALIGN',(2,1),(2,-1),'RIGHT'),
                ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
                ('GRID',(0,0),(-1,-1),0.3,colors.HexColor("#CBD5E0")),
                ('TOPPADDING',(0,0),(-1,-1),2.1),('BOTTOMPADDING',(0,0),(-1,-1),2.1),
                ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
            ]
            for i,r in enumerate(rows, start=1):
                tag = r.get('tag')
                if tag == 'section':
                    style_cmds.append(('FONTNAME',(1,i),(2,i),'Helvetica-Bold'))
                    style_cmds.append(('BACKGROUND',(0,i),(-1,i),colors.HexColor("#EDF2F7")))
                    style_cmds.append(('TEXTCOLOR',(1,i),(1,i),NAVY))
                if i % 2 == 0 and tag is None:
                    style_cmds.append(('BACKGROUND',(0,i),(-1,i),ZEBRA))
            # Highlight the final bridge "Ending Cash" row(s): find labels
            for i,row in enumerate(tbl_data[1:], start=1):
                lab = row[1]
                if "CASH & EQUIVALENTS, END" in lab:
                    style_cmds.append(('FONTNAME',(1,i),(2,i),'Helvetica-Bold'))
                    style_cmds.append(('BACKGROUND',(0,i),(-1,i),colors.HexColor("#EDF2F7")))
                    style_cmds.append(('LINEABOVE',(0,i),(-1,i),0.7,NAVY))
                    style_cmds.append(('LINEBELOW',(0,i),(-1,i),1.3,NAVY))
                if "CASH & EQUIVALENTS, BEGIN" in lab:
                    style_cmds.append(('FONTNAME',(1,i),(2,i),'Helvetica-Bold'))
            cfs.setStyle(TableStyle(style_cmds))
            elements.append(cfs)

            # Diagnostics summary
            diag = cf_data.get('diagnostics') or {}
            if diag.get('summary'):
                elements.append(Spacer(1,0.15*inch))
                elements.append(Paragraph(
                    f"<font size='8.5' color='#2D3748'><b>ASC 230 Integrity:</b></font> "
                    f"<font size='8.5' color='#4A5568'>{diag.get('summary')}</font>",
                    ParagraphStyle('diag', parent=cs['Normal'], leading=11)))
            elements.append(Spacer(1,0.1*inch))
            generated = Paragraph(
                f"<font color='#718096' size='7.5'>"
                f"Generated on {datetime.now().strftime('%Y-%m-%d at %H:%M:%S')} · Source: General Ledger (single source of truth) · Currency AED"
                "</font>", cs['Footer'])
            elements.append(generated)
            doc.build(elements, onFirstPage=self._page_number_footer, onLaterPages=self._page_number_footer)
            print(f"[OK] Cash Flow PDF → {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Error generating Cash Flow PDF: {e}"); raise

    # =========================================================================
    # NEW: BATCH PDF GENERATION — All reports → timestamped folder
    # =========================================================================
    def generate_all_reports_pdf_batch(self, data_manager=None, start_date=None, end_date=None,
                                       output_folder=None):
        """Generate 5 standard GAAP financial statements in a dedicated folder:
           1. Balance Sheet PDF
           2. Income Statement PDF
           3. Cash Flow Statement PDF
           4. AR Aging PDF
           5. Trial Balance PDF
           Returns: {folder: str, files: {key: path, ...}, index_pdf: path or None}"""
        from report_engine import ReportEngine, DateRangeType
        try:
            reports_folder = self._resolve_reports_folder(output_folder)
            batch_stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            batch_dir = os.path.join(reports_folder, f"FinancialReports_Batch_{batch_stamp}")
            os.makedirs(batch_dir, exist_ok=True)
            files_out = {}

            # Build engine if we got a data_manager
            engine = None
            if data_manager is not None:
                try:
                    if start_date and end_date:
                        engine = ReportEngine(data_manager, DateRangeType.CUSTOM, str(start_date), str(end_date))
                    else:
                        engine = ReportEngine(data_manager, DateRangeType.THIS_YEAR)
                except Exception as e:
                    print(f"[WARN] Could not build ReportEngine in batch: {e}")

            # 1-4: Use engine's four GAAP reports, fall back gracefully
            if engine is not None:
                try:
                    bs = engine.get_balance_sheet()
                    files_out['balance_sheet'] = self.generate_balance_sheet_pdf(bs, output_folder=batch_dir)
                except Exception as e: print(f"[WARN] Batch skipping Balance Sheet: {e}")
                try:
                    pnl = engine.get_profit_and_loss()
                    files_out['income_statement'] = self.generate_income_statement_pdf(pnl, output_folder=batch_dir)
                except Exception as e: print(f"[WARN] Batch skipping Income Statement: {e}")
                try:
                    cf = engine.get_cash_flow()
                    files_out['cash_flow'] = self.generate_cash_flow_pdf(cf, output_folder=batch_dir)
                except Exception as e: print(f"[WARN] Batch skipping Cash Flow: {e}")
                try:
                    ar = engine.get_aging_receivables()
                    files_out['ar_aging'] = self.generate_ar_aging_pdf(ar, output_folder=batch_dir)
                except Exception as e: print(f"[WARN] Batch skipping AR Aging: {e}")
                try:
                    tb = engine.get_trial_balance()
                    files_out['trial_balance'] = self.generate_trial_balance_pdf(tb, output_folder=batch_dir)
                except Exception as e: print(f"[WARN] Batch skipping Trial Balance: {e}")

            # Build a small index/summary PDF as user-facing cover sheet
            index_pdf = None
            try:
                ts = datetime.now()
                index_file = os.path.join(batch_dir, f"00_INDEX_Batch_{batch_stamp}.pdf")
                doc = SimpleDocTemplate(index_file, pagesize=letter,
                                        leftMargin=0.6*inch, rightMargin=0.6*inch,
                                        topMargin=0.55*inch, bottomMargin=0.75*inch,
                                        title="Batch Index - HopePharma Financial Reports",
                                        author="HopePharma Medical Trading L.L.C.")
                elements = []; cs = self._create_custom_styles()
                self._brand_header_story(elements, cs,
                    report_title="FINANCIAL REPORTS — BATCH EXPORT INDEX",
                    subtitle_lines=[
                        f"Generated: <b>{ts.strftime('%Y-%m-%d %H:%M:%S')}</b>",
                        f"Total reports in batch: <b>{len(files_out)}</b>"
                    ])
                NAVY = colors.HexColor("#1A365D")
                order = ["balance_sheet","income_statement","cash_flow","ar_aging","trial_balance"]
                labels = {"balance_sheet":"1) Statement of Financial Position (Balance Sheet)",
                          "income_statement":"2) Statement of Comprehensive Income (P&L)",
                          "cash_flow":"3) Statement of Cash Flows (ASC 230 Indirect)",
                          "ar_aging":"4) Accounts Receivable Aging Schedule",
                          "trial_balance":"5) General Ledger Trial Balance (6-Column)"}
                data_rows = [["#", "Report Title", "File Name"]]
                idx = 0
                for k in order:
                    if k not in files_out: continue
                    idx += 1
                    data_rows.append([str(idx), labels.get(k, k), os.path.basename(files_out[k])])
                t = Table(data_rows, colWidths=[0.4*inch, 4.2*inch, 2.8*inch], repeatRows=1)
                t.setStyle(TableStyle([
                    ('BACKGROUND',(0,0),(-1,0),NAVY),
                    ('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),
                    ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
                    ('ALIGN',(0,0),(0,-1),'CENTER'),
                    ('FONTSIZE',(0,0),(-1,0),9),
                    ('FONTSIZE',(0,1),(-1,-1),9),
                    ('FONTNAME',(0,1),(-1,-1),'Helvetica'),
                    ('GRID',(0,0),(-1,-1),0.3,colors.HexColor("#CBD5E0")),
                    ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
                    ('LEFTPADDING',(0,0),(-1,-1),5),('RIGHTPADDING',(0,0),(-1,-1),5),
                ]))
                elements.append(t)
                elements.append(Spacer(1,0.28*inch))
                elements.append(Paragraph(
                    "<font size='9' color='#2D3748'>"
                    "<b>Usage notes:</b><br/>"
                    "• All reports are generated from a single source of truth: the <b>General Ledger</b> (double-entry bookkeeping).<br/>"
                    "• Currency is <b>AED</b> (UAE Dirhams, 2 decimals). All totals are rounded to the nearest cent.<br/>"
                    "• The <b>Trial Balance</b> PDF should remain perfectly balanced — Closing Debits = Closing Credits.<br/>"
                    "• For audited financial statements, please share this batch folder with your reporting accountant / audit firm."
                    "</font>",
                    ParagraphStyle('notes', parent=cs['Normal'], leading=13)))
                elements.append(Spacer(1,0.15*inch))
                generated = Paragraph(
                    f"<font color='#718096' size='7.5'>"
                    f"Generated on {ts.strftime('%Y-%m-%d at %H:%M:%S')} · HopePharma Financial Reporting Suite"
                    "</font>", cs['Footer'])
                elements.append(generated)
                doc.build(elements, onFirstPage=self._page_number_footer, onLaterPages=self._page_number_footer)
                index_pdf = index_file
                # Prepend index to front of files_out so it appears first in listings
                files_out = {"__index": index_pdf, **files_out}
            except Exception as e:
                print(f"[WARN] Batch index PDF not produced: {e}")

            print(f"[OK] Batch PDF export done → {batch_dir}  ({len(files_out)} files)")
            # Open output folder (sandbox-safe: platform best-effort)
            try:
                import platform, subprocess
                if platform.system() == 'Darwin': subprocess.run(['open', batch_dir], check=False)
                elif platform.system() == 'Windows': os.startfile(batch_dir)  # noqa: attr-defined
                else: subprocess.run(['xdg-open', batch_dir], check=False)
            except Exception:
                pass
            return {"folder": batch_dir, "files": files_out, "index_pdf": index_pdf}
        except Exception as e:
            print(f"Error generating batch PDF reports: {e}"); raise

# Use the enhanced PDF generator
PDFGenerator = EnhancedPDFGenerator

class QuotationObject:
    """Quotation object for managing quotes with validation and methods"""
    
    def __init__(self, data):
        self._validate_and_initialize(data)
    
    def _validate_and_initialize(self, data):
        """Validate and initialize quotation data with defaults"""
        self.quotation_id = data.get('quotation_id', '')
        self.client_name = data.get('client_name', '')
        self.client_trn = data.get('client_trn', '')
        self.client_emirate = data.get('client_emirate', '')
        cl = (data.get('client_location') or '').strip()
        if not cl:
            em = str(self.client_emirate or '').strip()
            if not em:
                cl = ""
            else:
                uae_emirates = {
                    'Dubai', 'Abu Dhabi', 'Sharjah', 'Ajman',
                    'Umm Al Quwain', 'Ras Al Khaimah', 'Fujairah'
                }
                if ',' not in em and em in uae_emirates:
                    cl = f"{em}, United Arab Emirates"
                else:
                    cl = em
        self.client_location = cl
        self.date = data.get('date', datetime.now().strftime("%Y-%m-%d"))
        self.valid_until = data.get('valid_until', (datetime.now() + timedelta(days=30)).strftime("%Y-%m-%d"))
        self.tax_rate = float(data.get('tax_rate', 5))
        self.items = data.get('items', [])
        self.subtotal = float(data.get('subtotal', 0))
        self.taxable_amount = float(data.get('taxable_amount', 0))
        self.non_taxable_amount = float(data.get('non_taxable_amount', 0))
        self.tax_amount = float(data.get('tax_amount', 0))
        self.grand_total = float(data.get('grand_total', 0))
        self.notes = data.get('notes', '')
        self.currency = data.get('currency', 'AED')
        self.status = data.get('status', 'Draft')
        
        # Recalculate totals if needed
        self._recalculate_totals()
    
    def _recalculate_totals(self):
        """Recalculate quotation totals based on items"""
        try:
            # Calculate from items
            self.subtotal = sum(float(item.get('total', 0)) for item in self.items)
            self.taxable_amount = sum(float(item.get('total', 0)) for item in self.items if item.get('taxable', True))
            self.non_taxable_amount = self.subtotal - self.taxable_amount
            self.tax_amount = self.taxable_amount * (self.tax_rate / 100)
            self.grand_total = self.subtotal + self.tax_amount
        except Exception as e:
            print(f"Error recalculating quotation totals: {e}")
    
    def to_dict(self):
        """Convert to dictionary for GUI compatibility"""
        return {
            'quotation_id': self.quotation_id,
            'client_name': self.client_name,
            'client_trn': self.client_trn,
            'client_emirate': self.client_emirate,
            'client_location': self.client_location,
            'date': self.date,
            'valid_until': self.valid_until,
            'tax_rate': self.tax_rate,
            'items': self.items,
            'subtotal': self.subtotal,
            'taxable_amount': self.taxable_amount,
            'non_taxable_amount': self.non_taxable_amount,
            'tax_amount': self.tax_amount,
            'grand_total': self.grand_total,
            'notes': self.notes,
            'currency': self.currency,
            'status': self.status
        }


class InvoiceObject:
    """Enhanced Invoice object with better validation and methods"""
    
    def __init__(self, data):
        self._validate_and_initialize(data)
    
    def _validate_and_initialize(self, data):
        """Validate and initialize invoice data with defaults"""
        self.invoice_id = data.get('invoice_id', '')
        self.client_name = data.get('client_name', '')
        self.client_trn = data.get('client_trn', '')
        self.client_emirate = data.get('client_emirate', '')
        cl = (data.get('client_location') or '').strip()
        if not cl:
            em = str(self.client_emirate or '').strip()
            if not em:
                cl = ""
            else:
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
                    cl = f"{em}, United Arab Emirates"
                else:
                    cl = em
        self.client_location = cl
        self.date = data.get('date', datetime.now().strftime("%Y-%m-%d"))
        try:
            base_date = datetime.strptime(data.get('date', datetime.now().strftime("%Y-%m-%d")), "%Y-%m-%d")
        except Exception:
            base_date = datetime.now()
        pd = (data.get('payment_due') or 'On receipt').strip()
        if pd.lower().startswith('on receipt'):
            dd = base_date
        elif pd.startswith('30'):
            dd = base_date + timedelta(days=30)
        elif pd.startswith('60'):
            dd = base_date + timedelta(days=60)
        elif pd.startswith('90'):
            dd = base_date + timedelta(days=90)
        elif pd.startswith('120'):
            dd = base_date + timedelta(days=120)
        else:
            dd = base_date
        self.due_date = data.get('due_date', dd.strftime("%Y-%m-%d"))
        self.payment_method = data.get('payment_method', data.get('payment_terms', 'Cash'))
        self.payment_due = data.get('payment_due', 'On receipt')
        self.tax_rate = float(data.get('tax_rate', 5))
        self.items = data.get('items', [])
        self.costs = data.get('costs', [])
        self.subtotal = float(data.get('subtotal', 0))
        self.taxable_amount = float(data.get('taxable_amount', 0))
        self.non_taxable_amount = float(data.get('non_taxable_amount', 0))
        self.tax_amount = float(data.get('tax_amount', 0))
        self.grand_total = float(data.get('grand_total', 0))
        self.total_cost = float(data.get('total_cost', 0))
        self.profit_loss = float(data.get('profit_loss', 0))
        self.total_paid = float(data.get('total_paid', 0))
        self.balance_due = float(data.get('balance_due', 0))
        self.status = data.get('status', 'Not Paid')
        self.notes = data.get('notes', '')
        self.currency = data.get('currency', 'AED')
        self.invoice_type = data.get('invoice_type', 'sales')
        self.payment_history = data.get('payment_history', [])
        self.workflow_status = data.get('workflow_status', 'Under Process')
        
        # Recalculate totals if needed
        self._recalculate_totals()
    
    def _recalculate_totals(self):
        """Recalculate invoice totals based on items and costs"""
        try:
            # Calculate from items
            self.subtotal = sum(float(item.get('total', 0)) for item in self.items)
            
            # Calculate taxable and non-taxable amounts
            # If it's a delivery invoice, skip tax calculation entirely
            if self.invoice_type == 'delivery' or self.invoice_type == 'delivery-only':
                self.taxable_amount = 0.0
                self.non_taxable_amount = self.subtotal
                self.tax_amount = 0.0
                self.grand_total = self.subtotal  # EXACTLY delivery fee, no extra
            else:
                self.taxable_amount = sum(float(item.get('total', 0)) for item in self.items if item.get('taxable', True))
                self.non_taxable_amount = self.subtotal - self.taxable_amount
                # Calculate tax
                self.tax_amount = self.taxable_amount * (self.tax_rate / 100)
                self.grand_total = self.subtotal + self.tax_amount
            
            # Calculate total cost
            self.total_cost = sum(float(cost.get('amount', 0)) for cost in self.costs)
            
            # Calculate profit/loss
            self.profit_loss = self.grand_total - self.total_cost
            
            # Calculate balance due
            self.balance_due = self.grand_total - self.total_paid
            
            # Update status
            self._update_status()
            
        except Exception as e:
            print(f"Error recalculating totals: {e}")
    
    def _update_status(self):
        """Update invoice status based on payments"""
        if self.total_paid >= self.grand_total:
            self.status = 'Paid'
        elif self.total_paid > 0:
            self.status = 'Partial Paid'
        else:
            self.status = 'Not Paid'
    
    def to_dict(self):
        """Convert to dictionary for GUI compatibility"""
        return {
            'invoice_id': self.invoice_id,
            'client_name': self.client_name,
            'client_trn': self.client_trn,
            'client_emirate': self.client_emirate,
            'client_location': self.client_location,
            'date': self.date,
            'due_date': self.due_date,
            'payment_terms': self.payment_method,
            'payment_method': self.payment_method,
            'payment_due': self.payment_due,
            'tax_rate': self.tax_rate,
            'items': self.items,
            'costs': self.costs,
            'subtotal': self.subtotal,
            'taxable_amount': self.taxable_amount,
            'non_taxable_amount': self.non_taxable_amount,
            'tax_amount': self.tax_amount,
            'grand_total': self.grand_total,
            'total_cost': self.total_cost,
            'profit_loss': self.profit_loss,
            'total_paid': self.total_paid,
            'balance_due': self.balance_due,
            'status': self.status,
            'notes': self.notes,
            'currency': self.currency,
            'invoice_type': self.invoice_type,
            'payment_history': self.payment_history,
            'workflow_status': self.workflow_status
        }
    
    def add_payment(self, amount, payment_date, payment_method, notes, account=None):
        """Add payment to invoice with validation"""
        try:
            amount = float(amount)
            if amount <= 0:
                return False, "Payment amount must be positive"
            
            if amount > self.balance_due:
                curr = str(getattr(self, "currency", "AED") or "AED").strip() or "AED"
                return False, f"Payment amount cannot exceed balance due ({curr} {self.balance_due:,.2f})"
            
            # Create payment record
            payment_record = {
                'date': payment_date,
                'amount': amount,
                'method': payment_method,
                'account': account or payment_method,
                'notes': notes,
                'timestamp': datetime.now().isoformat()
            }
            
            # Add to payment history
            self.payment_history.append(payment_record)
            
            # Update payment information
            self.total_paid += amount
            self.balance_due = self.grand_total - self.total_paid
            
            # Update status
            self._update_status()
            
            print(f"Payment added: {amount} to invoice {self.invoice_id}")
            print(f"New total paid: {self.total_paid}, Status: {self.status}")
            
            return True, "Payment added successfully"
            
        except ValueError:
            return False, "Invalid payment amount"
        except Exception as e:
            print(f"Error adding payment: {e}")
            return False, f"Error: {str(e)}"

    def edit_payment(self, index, amount, payment_date, payment_method, notes, account=None):
        try:
            if index < 0 or index >= len(self.payment_history):
                return False, "Payment record not found"
            old = self.payment_history[index]
            old_amount = float(old.get('amount', 0) or 0)
            amount = float(amount)
            if amount <= 0:
                return False, "Payment amount must be positive"
            self.total_paid -= old_amount
            record = {
                'date': payment_date,
                'amount': amount,
                'method': payment_method,
                'account': account or payment_method,
                'notes': notes,
                'timestamp': datetime.now().isoformat()
            }
            self.payment_history[index] = record
            self.total_paid += amount
            self.balance_due = self.grand_total - self.total_paid
            self._update_status()
            return True, "Payment updated"
        except Exception as e:
            return False, f"Error: {str(e)}"

    def delete_payment(self, index):
        try:
            if index < 0 or index >= len(self.payment_history):
                return False, "Payment record not found"
            old = self.payment_history.pop(index)
            old_amount = float(old.get('amount', 0) or 0)
            self.total_paid -= old_amount
            self.balance_due = self.grand_total - self.total_paid
            self._update_status()
            return True, "Payment deleted"
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def add_item(self, description, quantity, unit_price, taxable=True):
        """Add item to invoice"""
        try:
            quantity = int(quantity)
            unit_price = float(unit_price)
            total = quantity * unit_price
            
            item = {
                'description': description,
                'quantity': quantity,
                'unit_price': unit_price,
                'total': total,
                'taxable': taxable
            }
            
            self.items.append(item)
            self._recalculate_totals()
            return True
            
        except (ValueError, TypeError):
            return False

    def edit_item(self, index, description, quantity, unit_price, taxable=True):
        try:
            if index < 0 or index >= len(self.items):
                return False
            quantity = int(quantity)
            unit_price = float(unit_price)
            total = quantity * unit_price
            self.items[index] = {
                'description': description,
                'quantity': quantity,
                'unit_price': unit_price,
                'total': total,
                'taxable': taxable
            }
            self._recalculate_totals()
            return True
        except (ValueError, TypeError):
            return False
    
    def validate_invoice(self):
        """Validate invoice data"""
        errors = []
        
        if not self.invoice_id:
            errors.append("Invoice ID is required")
        
        if not self.client_name:
            errors.append("Client name is required")
        
        if not self.date:
            errors.append("Invoice date is required")
        
        if self.grand_total <= 0:
            errors.append("Invoice total must be greater than 0")
        
        if not self.items:
            errors.append("At least one item is required")
        
        return errors

    def edit_cost(self, index, description, amount, account=None, notes=None):
        try:
            if index < 0:
                return False
            amount = float(amount)
            while len(self.costs) <= index:
                self.costs.append({})
            self.costs[index] = {
                'description': description,
                'amount': amount,
                'account': account or self.currency,
                'notes': notes or ''
            }
            self._recalculate_totals()
            return True
        except (ValueError, TypeError):
            return False

class PurchaseObject:
    """Enhanced Purchase object with validation"""
    
    def __init__(self, data):
        self.purchase_id = data.get('purchase_id', '')
        self.description = data.get('description', '')
        self.amount = float(data.get('amount', 0))
        self.date = data.get('date', datetime.now().strftime("%Y-%m-%d"))
        self.category = data.get('category', 'General')
        self.supplier = data.get('supplier', '')
        self.receipt_attached = data.get('receipt_attached', False)
        self.receipt_filename = data.get('receipt_filename', '')
        self.notes = data.get('notes', '')
        self.account = data.get('account', 'Cash')
        self.amount_paid = float(data.get('amount_paid', data.get('paid_amount', data.get('payment_amount', 0.0))))
        self.payment_date = data.get('payment_date', '')
        self.bank_account = data.get('bank_account', '')
        self.paid_status = data.get('paid_status') or data.get('payment_status') or ''
        if not self.paid_status:
            try:
                total = float(self.amount or 0)
                paid = float(self.amount_paid or 0)
                if paid <= 0:
                    self.paid_status = 'Pending'
                elif paid >= total - 0.005:
                    self.paid_status = 'Paid'
                else:
                    self.paid_status = 'Partial'
            except Exception:
                self.paid_status = 'Pending'
    
    def to_dict(self):
        """Convert to dictionary for GUI compatibility"""
        return {
            'purchase_id': self.purchase_id,
            'description': self.description,
            'amount': self.amount,
            'date': self.date,
            'category': self.category,
            'supplier': self.supplier,
            'receipt_attached': self.receipt_attached,
            'receipt_filename': self.receipt_filename,
            'notes': self.notes,
            'account': self.account,
            'amount_paid': self.amount_paid,
            'payment_date': self.payment_date,
            'bank_account': self.bank_account,
            'paid_status': self.paid_status
        }
    
    def validate(self):
        """Validate purchase data"""
        errors = []
        
        if not self.description:
            errors.append("Description is required")
        
        if self.amount <= 0:
            errors.append("Amount must be greater than 0")
        
        if not self.date:
            errors.append("Date is required")
        
        return errors

class EnhancedCloudDataManager:
    """Enhanced cloud data manager with better error handling and features"""
    
    def __init__(self, data_folder, mode):
        self.data_folder = Path(data_folder)
        self.mode = mode
        self.invoice_folder = data_folder
        self.system = platform.system()
        self.pdf_generator = EnhancedPDFGenerator(data_manager=self)
        try:
            from marketplace_manager import MarketplaceManager
            self.marketplace_manager = MarketplaceManager(str(self.data_folder))
        except Exception:
            self.marketplace_manager = None
        try:
            self.pdf_generator.output_folder = self.data_folder / 'HopePharmaInvoices'
        except Exception:
            pass
        # Accounting services (Ledger + Report engine)
        try:
            from ledger_service import LedgerService
            self.ledger = LedgerService(self)
        except Exception as e:
            print(f"Ledger service unavailable: {e}")
            self.ledger = None
        
        print(f"Enhanced Cloud Data folder: {self.data_folder}")
        print(f"Mode: {mode}")
        print(f"Operating System: {self.system}")
        
        # Create necessary folder structure
        self._create_folder_structure()
        # Attempt restore from safe backup if data files are missing
        try:
            self.try_restore_from_safe_backup()
        except Exception as e:
            print(f"Safe restore skipped: {e}")
        
        # Initialize data files
        self.initialize_data_files()
        # Seed Chart of Accounts if empty
        try:
            self._seed_chart_of_accounts()
        except Exception as e:
            print(f"CoA seeding skipped: {e}")
        
        # Start sync for cloud mode
        if mode in ['cloud']:
            self.start_enhanced_cloud_sync()
    
    def _create_folder_structure(self):
        """Create organized folder structure.

        On serverless / Vercel the data dir might live on a read-only or ephemeral
        filesystem. Receipt/purchase/backup folders are used only by the desktop
        application; on Vercel Supabase handles all persistence anyway. So we
        MUST swallow mkdir failures here instead of crashing boot.
        """
        folders = [
            'receipts',
            'purchases',
            'backups',
            'exports',
            'temp',
            'logs'
        ]

        for folder in folders:
            folder_path = self.data_folder / folder
            try:
                folder_path.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                # OSError Errno 30 = Read-only file system on Vercel Linux;
                # PermissionError / OSError variants on other serverless.
                if not isinstance(e, (OSError, PermissionError)):
                    print(f"[_create_folder_structure] unexpected: {e}")
                # Non-fatal: Supabase persistence doesn't use these folders.
                continue

    def initialize_data_files(self):
        """Create initial data files with enhanced structure.

        As with _create_folder_structure, any failure to write to disk must be
        non-fatal on serverless (all real persistence is handled via Supabase
        save_json / load_json on the running app). If the disk write fails we
        just log and continue — first request that loads the file from Supabase
        will fill in-memory state correctly.
        """
        data_files = {
            'invoices_data.json': [],
            'purchases_data.json': [],
            'quotes_data.json': [],
            'employees_data.json': {},
            'salaries_data.json': {},
            'employees.json': [],
            'salary_reports.json': [],
            'clients_memory.json': [],
            'products_memory.json': [],
            'chart_of_accounts.json': [],
            'general_ledger.json': [],
            'period_locks.json': [],
            'app_settings.json': {
                'last_backup': None,
                'auto_backup': True,
                'default_tax_rate': 5,
                'currency': 'AED',
                'invoice_prefix': 'HPMT',
                'quote_prefix': 'HPQT',
                'last_invoice_number': 910599,  # Start from 910599 so next is 910600
                'last_quote_number': 900599    # Start from 900599 so next is 900600
            }
        }

        for filename, default_data in data_files.items():
            file_path = self.data_folder / filename
            try:
                if not file_path.exists():
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump(default_data, f, indent=2, ensure_ascii=False)
                    print(f"Created {file_path}")
                else:
                    # Repair corrupted files
                    self._repair_data_file(file_path, default_data)
            except Exception as e:
                if not isinstance(e, (OSError, PermissionError)):
                    print(f"[initialize_data_files] unexpected for {filename}: {e}")
                continue
    
    def _repair_data_file(self, file_path, default_data):
        """Repair corrupted data files"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                
            # If file is empty or contains invalid JSON, recreate it
            if not content:
                print(f"Repairing empty file: {file_path}")
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(default_data, f, indent=2, ensure_ascii=False)
                return
                
            # Try to parse JSON
            try:
                data = json.loads(content)
            except json.JSONDecodeError:
                print(f"Repairing corrupted JSON file: {file_path}")
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(default_data, f, indent=2, ensure_ascii=False)
                return
                
            # Check if data structure is correct
            if file_path.name in ['invoices_data.json', 'purchases_data.json', 'quotes_data.json',
                                  'chart_of_accounts.json', 'general_ledger.json', 'period_locks.json']:
                if not isinstance(data, list):
                    print(f"Repairing non-list data in: {file_path}")
                    with open(file_path, 'w', encoding='utf-8') as f:
                        json.dump([], f, indent=2, ensure_ascii=False)
                    
        except Exception as e:
            print(f"Error repairing {file_path}: {e}")
            # Recreate the file with default data
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(default_data, f, indent=2, ensure_ascii=False)
                print(f"Recreated {file_path} with default data")
            except Exception as e2:
                print(f"Failed to recreate {file_path}: {e2}")
    
    def _seed_chart_of_accounts(self):
        """Seed GAAP-compliant Chart of Accounts (13 accounts exactly as specified) if file is empty."""
        coa = self.load_json('chart_of_accounts.json')
        if coa and len(coa) > 0:
            return  # Already seeded
        SEED_COA = [
            {"account_code": "1000", "account_name": "Cash", "full_code_name": "1000_Cash", "account_type": "ASSET", "is_seeded": True},
            {"account_code": "1100", "account_name": "Bank", "full_code_name": "1100_Bank", "account_type": "ASSET", "is_seeded": True},
            {"account_code": "1200", "account_name": "Accounts_Receivable", "full_code_name": "1200_Accounts_Receivable", "account_type": "ASSET", "is_seeded": True},
            {"account_code": "1300", "account_name": "Inventory", "full_code_name": "1300_Inventory", "account_type": "ASSET", "is_seeded": True},
            {"account_code": "2000", "account_name": "Accounts_Payable", "full_code_name": "2000_Accounts_Payable", "account_type": "LIABILITY", "is_seeded": True},
            {"account_code": "2100", "account_name": "Accrued_Expenses", "full_code_name": "2100_Accrued_Expenses", "account_type": "LIABILITY", "is_seeded": True},
            {"account_code": "3000", "account_name": "Retained_Earnings", "full_code_name": "3000_Retained_Earnings", "account_type": "EQUITY", "is_seeded": True},
            {"account_code": "3100", "account_name": "Owner_Equity", "full_code_name": "3100_Owner_Equity", "account_type": "EQUITY", "is_seeded": True},
            {"account_code": "4000", "account_name": "Sales_Revenue", "full_code_name": "4000_Sales_Revenue", "account_type": "REVENUE", "is_seeded": True},
            {"account_code": "5000", "account_name": "COGS", "full_code_name": "5000_COGS", "account_type": "EXPENSE", "is_seeded": True},
            {"account_code": "5100", "account_name": "Salary_Expense", "full_code_name": "5100_Salary_Expense", "account_type": "EXPENSE", "is_seeded": True},
            {"account_code": "5200", "account_name": "Rent_Utility", "full_code_name": "5200_Rent_Utility", "account_type": "EXPENSE", "is_seeded": True},
            {"account_code": "5300", "account_name": "General_Admin", "full_code_name": "5300_General_Admin", "account_type": "EXPENSE", "is_seeded": True},
        ]
        self.save_json('chart_of_accounts.json', SEED_COA)
        print(f"Seeded {len(SEED_COA)} Chart of Accounts entries successfully.")

    def start_enhanced_cloud_sync(self):
        """Enhanced background sync with better monitoring"""
        def sync():
            last_check = time.time()
            while True:
                try:
                    current_time = time.time()
                    # Perform sync every 60 seconds
                    if current_time - last_check >= 60:
                        self._perform_sync_operations()
                        last_check = current_time
                    
                    time.sleep(10)  # Check every 10 seconds
                    
                except Exception as e:
                    print(f"Sync error: {e}")
                    time.sleep(30)  # Wait longer on error
        
        thread = threading.Thread(target=sync, daemon=True)
        thread.start()
    
    def _perform_sync_operations(self):
        """Perform sync operations"""
        try:
            # Create backup if auto-backup is enabled
            settings = self.load_json('app_settings.json')
            if settings.get('auto_backup', True):
                self._create_auto_backup()
            
            # Clean up temp files older than 1 day
            self._cleanup_temp_files()
            
        except Exception as e:
            print(f"Sync operations error: {e}")
    
    def _create_auto_backup(self):
        """Create automatic backup as a zip, mirrored to a safe OS location"""
        try:
            backup_folder = self.data_folder / 'backups'
            backup_folder.mkdir(parents=True, exist_ok=True)
            safe_backup_root = self._safe_backup_root()
            safe_backup_root.mkdir(parents=True, exist_ok=True)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            name = f"auto_backup_{timestamp}.zip"
            zip_in_data = backup_folder / name
            zip_in_safe = safe_backup_root / name

            important_files = [
                'invoices_data.json',
                'purchases_data.json', 
                'clients_memory.json',
                'products_memory.json',
                'app_settings.json',
                'quotes_data.json',
                'chart_of_accounts.json',
                'general_ledger.json',
                'period_locks.json'
            ]

            def write_zip(zip_path: Path):
                with zipfile.ZipFile(zip_path, 'w', compression=zipfile.ZIP_DEFLATED) as zf:
                    for fname in important_files:
                        src = self.data_folder / fname
                        if src.exists():
                            zf.write(src, arcname=fname)
                return zip_path

            # Write both backups
            try:
                write_zip(zip_in_data)
                write_zip(zip_in_safe)
            except Exception as e:
                print(f"Backup zip error: {e}")

            # Update last backup time
            settings = self.load_json('app_settings.json')
            settings['last_backup'] = datetime.now().isoformat()
            self.save_json('app_settings.json', settings)

            # Retention policy: keep last 20 backups in each location
            for root in [backup_folder, safe_backup_root]:
                try:
                    zips = sorted(root.glob('auto_backup_*.zip'))
                    if len(zips) > 20:
                        for old in zips[:-20]:
                            try:
                                old.unlink()
                            except Exception:
                                pass
                except Exception:
                    pass

            print(f"Auto-backup created: {zip_in_data} and mirrored to {zip_in_safe}")
            
        except Exception as e:
            print(f"Auto-backup error: {e}")

    def _safe_backup_root(self) -> Path:
        try:
            if self.system == 'Windows':
                base = Path(os.environ.get('APPDATA', str(Path.home())))
                return base / 'HopePharma' / 'backups'
            elif self.system == 'Darwin':
                return Path.home() / 'Library' / 'Application Support' / 'HopePharma' / 'backups'
            else:
                return Path.home() / '.hopepharma' / 'backups'
        except Exception:
            return Path.home() / '.hopepharma' / 'backups'

    def try_restore_from_safe_backup(self):
        """Restore missing data files from the latest safe backup zip"""
        try:
            needed = {
                'invoices_data.json': [],
                'purchases_data.json': [],
                'clients_memory.json': [],
                'products_memory.json': [],
                'app_settings.json': {}
            }
            missing = []
            for fname in needed.keys():
                if not (self.data_folder / fname).exists():
                    missing.append(fname)
            if not missing:
                return
            safe_root = self._safe_backup_root()
            if not safe_root.exists():
                return
            zips = sorted(safe_root.glob('auto_backup_*.zip'))
            if not zips:
                return
            latest = zips[-1]
            with zipfile.ZipFile(latest, 'r') as zf:
                for fname in missing:
                    try:
                        zf.extract(fname, path=self.data_folder)
                    except Exception:
                        pass
        except Exception as e:
            print(f"Restore from safe backup failed: {e}")
    
    def _cleanup_temp_files(self):
        """Clean up temporary files"""
        try:
            temp_folder = self.data_folder / 'temp'
            if temp_folder.exists():
                for file_path in temp_folder.glob("*"):
                    if file_path.is_file():
                        # Delete files older than 1 day
                        file_age = time.time() - file_path.stat().st_mtime
                        if file_age > 86400:  # 24 hours
                            file_path.unlink()
                            print(f"Cleaned up temp file: {file_path}")
        except Exception as e:
            print(f"Temp cleanup error: {e}")

    # ------------------------------------------------------------
    # Supabase JSON-document storage adapter
    # ------------------------------------------------------------
    def _supabase_headers(self):
        """Return correct HTTP headers based on the key format (new sb_ vs old JWT).

        Supabase supports TWO API key formats:
          - OLD (pre-2025): JWT tokens, always start with 'eyJ...' (e.g. legacy
            service_role JWTs). These require BOTH: 'apikey: <key>' AND
            'Authorization: Bearer <key>'.
          - NEW (2025+ Publishable / Secret keys): opaque tokens starting with
            'sb_publishable_' / 'sb_secret_'. These ONLY work with the plain
            'apikey: <key>' header. If you ALSO send 'Authorization: Bearer <key>'
            alongside a non-JWT sb_ key, the new Supabase OpenAPI validator tries
            to parse the Bearer token as a JWT, fails, falls into a schema branch
            that expects different fields, and emits the misleading 400 error:

                Invalid request: should NOT have additional property 'public'.

            To resolve, we detect the key format dynamically and send only the
            headers that make sense for it.
        """
        sb_url = os.environ.get("SUPABASE_URL")
        sb_key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY")
        if not sb_url or not sb_key:
            return None
        k = (sb_key or "").strip()
        h = {
            "apikey": k,
            "Content-Type": "application/json",
            "Accept": "application/json",
            "X-Client-Info": "asistem-python/1.0",
        }
        # Add Authorization: Bearer ONLY for legacy JWT-style keys (eyJ...)
        # NOT for new sb_secret_ / sb_publishable_ opaque tokens
        if k.startswith("eyJ") or k.startswith("ey0") or (len(k) > 100 and "." in k):
            h["Authorization"] = f"Bearer {k}"
        return h

    def _sb_extract_supabase_error_body(self, exc):
        """If urllib raised HTTPError, read and pretty-print the Supabase JSON
        error body (it usually contains a helpful 'code', 'message', 'hint' and
        'details' that are way more useful than the generic urllib HTTPError).
        """
        try:
            resp = getattr(exc, "read", None)
            if callable(resp):
                raw = resp().decode("utf-8", errors="replace")
            else:
                raw = ""
            if raw:
                try:
                    import json as _json
                    err_obj = _json.loads(raw)
                    return f"[SupabaseErrorBody status={getattr(exc,'code','?')}] {_json.dumps(err_obj, indent=2, default=str)}"
                except Exception:
                    return f"[SupabaseErrorBodyRaw status={getattr(exc,'code','?')}] {raw[:800]}"
        except Exception:
            return ""
        return ""

    def _sb_sanitize_url(self, url):
        """Strip deprecated/illegal query params (schema=public, public=...) from URLs.

        Newer Supabase REST schema validators reject ANY request that sends the
        legacy PostgREST query-string parameter `?schema=public` (or any body/query
        field literally named `public`) with HTTP 400:

            Invalid request: should NOT have additional property 'public'.

        The schema is now configured server-side per project (default = public),
        so clients MUST NOT send the `schema` parameter anymore. We strip both
        `schema` and the literal `public` key here defensively before every call.
        """
        try:
            from urllib.parse import urlparse, parse_qs, urlencode, urlunparse
        except Exception:
            return url
        try:
            p = urlparse(url)
            qs = parse_qs(p.query, keep_blank_values=True)
            # Drop ONLY the deprecated/illegal params.
            # Valid Supabase REST query params (select, limit, offset, order, eq.,
            # gt., lt., in., is., or., and., not., like., ilike., etc.) MUST remain.
            for bad in ("schema", "public"):
                qs.pop(bad, None)
            for k in list(qs.keys()):
                if isinstance(k, str) and k.lower() in {"schema", "public"}:
                    qs.pop(k, None)
            new_query = urlencode(qs, doseq=True)
            parts = list(p)
            parts[4] = new_query
            return urlunparse(parts)
        except Exception:
            return url

    def _sb_sanitize_payload(self, payload_dict):
        """Remove ANY accidental top-level key named 'public' / 'schema' from POST body dicts.

        Newer Supabase REST schemas use strict OpenAPI validation and reject ANY
        request body that contains a top-level key not declared in the endpoint's
        OpenAPI schema. For the standard /rest/v1/<table> endpoints, the only
        legal keys are the table's column names (doc_name, content, updated_at).
        If the body contains an extra key (e.g. due to a bug that merged a URL
        query dict into the payload dict), the validator returns:

            Invalid request: should NOT have additional property '<keyname>'.

        Here we scrub both forbidden keywords ('public', 'schema') and any key
        that would look like an SQL identifier/query to a schema, purely as a
        defensive guard so saves can never 400 for this reason again.
        """
        if not isinstance(payload_dict, dict):
            return payload_dict
        cleaned = {}
        forbidden_keys_lower = {"public", "schema", "search_path", "current_schema", "database"}
        forbidden_substrings = ("create table", "drop ", "alter table", "exec_sql", "rpc/")
        for k, v in payload_dict.items():
            if not isinstance(k, str):
                cleaned[k] = v
                continue
            kl = k.lower()
            if kl in forbidden_keys_lower:
                continue
            if any(bad in kl for bad in forbidden_substrings):
                continue
            cleaned[k] = v
        return cleaned

    def _supabase_url(self):
        sb_url = os.environ.get("SUPABASE_URL")
        if not sb_url:
            return None
        return sb_url.rstrip("/")

    def _supabase_init_table(self):
        """Verify public.json_docs table exists on Supabase REST (does NOT create).

        Supabase no longer allows arbitrary `exec_sql` RPC by default (new security
        policy), and on newer Supabase REST schemas sending ANY body with a `public`
        key or a body to a non-existent stored procedure triggers:
        "Invalid request: should NOT have additional property `public`".

        So instead of trying to create the table here, we do ONLY a lightweight
        GET probe to verify the table endpoint is reachable. The user MUST run
        the one-time SQL setup block from upload_all_data_to_supabase.py's README
        (SQL Editor -> New query -> paste the CREATE TABLE + RLS policy block)
        BEFORE using Supabase storage for the first time.
        """
        try:
            import urllib.request as _urlreq
        except Exception:
            return False
        headers = self._supabase_headers()
        sb_url = self._supabase_url()
        if not headers or not sb_url:
            return False
        try:
            raw_url = f"{sb_url}/rest/v1/json_docs?select=doc_name&limit=1"
            final_url = self._sb_sanitize_url(raw_url)
            hdr = dict(headers)
            req = _urlreq.Request(
                final_url,
                headers=hdr,
                method="GET",
            )
            ctx = _ssl_context_maybe_unverified()
            kwargs = {"timeout": 10}
            if ctx is not None:
                kwargs["context"] = ctx
            with _urlreq.urlopen(req, **kwargs) as resp:
                return 200 <= resp.status < 300
        except Exception as e:
            err_body = self._sb_extract_supabase_error_body(e)
            print(f"[SupabaseInit] json_docs probe failed (expected if table not created yet): {e}")
            if err_body:
                print(f"[SupabaseInit] {err_body}")
            return False

    def _supabase_load(self, filename):
        try:
            import urllib.request as _urlreq
            import urllib.parse as _urlpar
        except Exception:
            return None
        headers = self._supabase_headers()
        sb_url = self._supabase_url()
        if not headers or not sb_url:
            return None
        if not getattr(self, "_sb_table_inited", False):
            self._sb_table_inited = True
        try:
            qname = _urlpar.quote(filename, safe="")
            raw_url = f"{sb_url}/rest/v1/json_docs?doc_name=eq.{qname}&select=content"
            final_url = self._sb_sanitize_url(raw_url)
            req = _urlreq.Request(
                final_url,
                headers=dict(headers),
                method="GET",
            )
            ctx = _ssl_context_maybe_unverified()
            kwargs = {"timeout": 15}
            if ctx is not None:
                kwargs["context"] = ctx
            with _urlreq.urlopen(req, **kwargs) as resp:
                body = resp.read().decode("utf-8")
                rows = json.loads(body) if body else []
                if not rows or not isinstance(rows, list):
                    return "__MISSING__"
                content = rows[0].get("content")
                if content is None:
                    return "__MISSING__"
                return content
        except Exception as e:
            err_body = self._sb_extract_supabase_error_body(e)
            print(f"[SupabaseLoad] Warning for {filename}: {e}")
            if err_body:
                print(f"[SupabaseLoad] {err_body}")
            return None

    def _supabase_save(self, filename, data):
        try:
            import urllib.request as _urlreq
        except Exception:
            return None
        headers = self._supabase_headers()
        sb_url = self._supabase_url()
        if not headers or not sb_url:
            return None
        if not getattr(self, "_sb_table_inited", False):
            self._sb_table_inited = True
        try:
            h = dict(headers)
            h["Prefer"] = "resolution=merge-duplicates,return=minimal"
            now_iso = datetime.now().isoformat()
            raw_payload = {
                "doc_name": filename,
                "content": data,
                "updated_at": now_iso,
            }
            clean_payload = self._sb_sanitize_payload(raw_payload)
            payload = json.dumps(clean_payload).encode("utf-8")
            raw_url = f"{sb_url}/rest/v1/json_docs"
            final_url = self._sb_sanitize_url(raw_url)
            req = _urlreq.Request(
                final_url,
                data=payload,
                headers=h,
                method="POST",
            )
            ctx = _ssl_context_maybe_unverified()
            kwargs = {"timeout": 15}
            if ctx is not None:
                kwargs["context"] = ctx
            with _urlreq.urlopen(req, **kwargs) as resp:
                return 200 <= resp.status < 300
        except Exception as e:
            err_body = self._sb_extract_supabase_error_body(e)
            print(f"[SupabaseSave] Warning for {filename}: {e}")
            if err_body:
                print(f"[SupabaseSave] {err_body}")
            return None

    def _default_json_value(self, filename):
        if filename in ['invoices_data.json', 'purchases_data.json',
                        'quotes_data.json', 'employees.json',
                        'clients_memory.json', 'products_memory.json',
                        'chart_of_accounts.json', 'general_ledger.json',
                        'period_locks.json', 'salary_reports.json']:
            return []
        if filename in ['employees_data.json', 'salaries_data.json',
                        'app_settings.json', 'app_state.json']:
            return {}
        return []

    # Data management methods
    def load_json(self, filename):
        """Load JSON data — Supabase REST when env vars set, else local disk.

        ERP UPGRADE (2026 Q3): Post-load schema coercion happens INSIDE this
        function so the rest of the codebase can safely assume the new
        schema shape. Files we coerce/seed: invoices_data, products,
        company_profiles, cost_centers (per company rows), service_products
        (per company rows).
        """
        sb_headers = self._supabase_headers()
        if sb_headers:
            sb_result = self._supabase_load(filename)
            if sb_result is None:
                pass
            elif sb_result == "__MISSING__":
                data = self._default_json_value(filename)
            else:
                data = sb_result
                if filename in ['invoices_data.json', 'purchases_data.json']:
                    if not isinstance(data, list):
                        data = []
                    else:
                        data = [item for item in data if isinstance(item, dict)]
                data = self._post_load_coerce(filename, data)
                return data
        else:
            file_path = self.data_folder / filename
            if file_path.exists():
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                    if filename in ['invoices_data.json', 'purchases_data.json']:
                        if not isinstance(data, list):
                            print(f"Repairing {filename}: converting to list")
                            data = []
                            self.save_json(filename, data)
                        else:
                            data = [item for item in data if isinstance(item, dict)]
                    data = self._post_load_coerce(filename, data)
                    return data
                except Exception as e:
                    print(f"Error loading {filename}: {e}")
                    data = self._default_json_value(filename)
            else:
                data = self._default_json_value(filename)
        data = self._post_load_coerce(filename, data)
        return data

    # =====================================================================
    # ERP UPGRADE: Schema coercion helpers (post-load, lazy, non-destructive)
    # =====================================================================
    def _post_load_coerce(self, filename, data):
        """Apply schema defaults to data loaded from disk/Supabase.

        - Never removes data; only fills missing fields with safe defaults.
        - Company-scoped tables (cost_centers, service_products) auto-seed
          rows for any company_id currently present in company_profiles.json.
        """
        try:
            if filename == self.COMPANIES_FILENAME:
                return self._coerce_company_profiles(data)
            if filename == "invoices_data.json":
                return self._coerce_invoices(data)
            if filename == "products.json" or filename == "products_memory.json":
                return self._coerce_products(data)
            if filename == self.COST_CENTERS_FILENAME:
                return self._coerce_and_seed_cost_centers(data)
            if filename == self.SERVICE_PRODUCTS_FILENAME:
                return self._coerce_and_seed_service_products(data)
        except Exception as e:
            print(f"[_post_load_coerce] warning on {filename}: {e}")
        return data

    def _coerce_company_profiles(self, data):
        """Ensure every company profile has all new fields (PDF letterhead,
        branding images, bank details). Missing HopePharma fields are filled
        with the legacy hardcoded values to keep PDFs identical after the
        renderer refactor (Task 2)."""
        if not isinstance(data, list):
            return data
        default_full = self._default_company_seed()
        # Set of keys we want every profile to have; copy from default_full
        # only if the key is missing (never overwrite an existing value).
        coerced_any = False
        for i, p in enumerate(data):
            if not isinstance(p, dict):
                continue
            for k, default_v in default_full.items():
                if k not in p:
                    if k == "id":
                        continue  # id was already set or we'd overwrite it
                    if p.get("is_default") and k == "created_at":
                        continue
                    p[k] = default_v
                    coerced_any = True
        if coerced_any:
            # Persist coercion on the NEXT save_json call; for this load,
            # just return it patched.
            try:
                self.save_json(self.COMPANIES_FILENAME, data)
            except Exception:
                pass
        return data

    def _coerce_invoices(self, data):
        """Legacy invoices (missing invoice_type, cost_center_id, cogs fields)
        get safe defaults (type=sales, cc_general, 0 cost). AC-10 backward compat."""
        if not isinstance(data, list):
            return data
        for inv in data:
            if not isinstance(inv, dict):
                continue
            # company_id: legacy invoices (company_id is None or absent) -> default
            if not inv.get("company_id"):
                inv["company_id"] = self.DEFAULT_COMPANY_ID
            # Type enum: sales | service. Legacy "delivery" becomes sales (no service support for deliveries yet).
            raw_type = (inv.get("invoice_type") or "").strip().lower()
            if raw_type == "service":
                inv["invoice_type"] = "service"
            else:
                # sales / delivery / delivery-only / missing → "sales"
                inv["invoice_type"] = "sales"
            # Header cost_center default cc_general
            if not inv.get("cost_center_id"):
                inv["cost_center_id"] = "cc_general"
            # Line items: ensure per-line cost_center / unit_cost / line_cost_total / gross_profit
            items = inv.get("items")
            if not isinstance(items, list):
                items = []
            hdr_cc = inv.get("cost_center_id")
            inv_total_cost = 0.0
            subtotal = float(inv.get("subtotal") or 0.0)
            for row in items:
                if not isinstance(row, dict):
                    continue
                if "cost_center_id" not in row or not row["cost_center_id"]:
                    row["cost_center_id"] = hdr_cc or "cc_general"
                if "unit_cost" not in row:
                    row["unit_cost"] = 0.0
                try:
                    qty = float(row.get("quantity") or row.get("hours") or 1)
                except (ValueError, TypeError):
                    qty = 1
                try:
                    uc = float(row.get("unit_cost") or 0.0)
                except (ValueError, TypeError):
                    uc = 0.0
                line_subtotal = float(row.get("line_subtotal") or row.get("subtotal") or row.get("total") or 0.0)
                row["line_cost_total"] = round(qty * uc, 2)
                row["line_gross_profit"] = round(line_subtotal - row["line_cost_total"], 2)
                inv_total_cost += row["line_cost_total"]
            inv["invoice_total_cost"] = round(inv_total_cost, 2)
            inv["invoice_gross_profit"] = round(subtotal - inv_total_cost, 2)
            inv["items"] = items
        return data

    def _coerce_products(self, data):
        """Add unit_cost (0.0 default) + default_cost_center_id (cc_sales_goods) to all products."""
        if not isinstance(data, list):
            return data
        for row in data:
            if not isinstance(row, dict):
                continue
            if "unit_cost" not in row:
                try:
                    # If cost_price, actual_cost existed, honor them; otherwise 0
                    fallback = float(row.get("cost_price") or row.get("actual_cost") or row.get("purchase_price") or 0.0)
                except (ValueError, TypeError):
                    fallback = 0.0
                row["unit_cost"] = fallback
            if "default_cost_center_id" not in row:
                row["default_cost_center_id"] = "cc_sales_goods"
        return data

    # ------------------------------------------------------------------
    # Cost centers
    # ------------------------------------------------------------------
    @staticmethod
    def _default_cost_centers_for(company_id):
        """Four default cost centers per company (FR-2)."""
        return [
            {
                "company_id": company_id,
                "cost_center_id": "cc_sales_goods",
                "label": "Sales (Goods COGS)",
                "description": "Direct cost of physical goods sold (product unit_cost * qty).",
                "category": "OPERATING",
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            {
                "company_id": company_id,
                "cost_center_id": "cc_services_tech",
                "label": "Technical Services Direct Costs",
                "description": "Direct internal cost of technical services delivered (hourly rate * qty hrs).",
                "category": "OPERATING",
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            {
                "company_id": company_id,
                "cost_center_id": "cc_services_maint",
                "label": "Maintenance & Repair Direct Costs",
                "description": "Direct cost of maintenance/repair engagements.",
                "category": "OPERATING",
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            {
                "company_id": company_id,
                "cost_center_id": "cc_general",
                "label": "General / Unallocated",
                "description": "Default cost center for legacy invoices or lines not yet assigned.",
                "category": "OPERATING",
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
        ]

    def _ensure_cost_centers_seeded_for_company(self, company_id, save=False):
        """Ensure the 4 builtin cost centers exist for a given company_id.
        Returns the (possibly mutated) list of cost center rows."""
        if not company_id:
            return []
        cc_list = self.load_json(self.COST_CENTERS_FILENAME) or []
        if not isinstance(cc_list, list):
            cc_list = []
        existing_keys = {(r.get("company_id"), r.get("cost_center_id")) for r in cc_list if isinstance(r, dict)}
        mutated = False
        for tpl in self._default_cost_centers_for(company_id):
            if (tpl["company_id"], tpl["cost_center_id"]) not in existing_keys:
                cc_list.append(tpl)
                mutated = True
        if mutated and save:
            self.save_json(self.COST_CENTERS_FILENAME, cc_list)
        return cc_list

    def _coerce_and_seed_cost_centers(self, data):
        """Ensure every company in company_profiles has the 4 default CCs.
        Also coerce any malformed rows."""
        if not isinstance(data, list):
            data = []
        try:
            profiles = self.get_all_company_profiles() or []
        except Exception:
            profiles = []
        mutated = False
        data = [r for r in data if isinstance(r, dict)]
        existing_keys = {(r.get("company_id"), r.get("cost_center_id")) for r in data}
        for p in profiles:
            cid = p.get("id")
            if not cid:
                continue
            for tpl in self._default_cost_centers_for(cid):
                if (cid, tpl["cost_center_id"]) not in existing_keys:
                    data.append(tpl)
                    existing_keys.add((cid, tpl["cost_center_id"]))
                    mutated = True
        if mutated:
            try:
                self.save_json(self.COST_CENTERS_FILENAME, data)
            except Exception:
                pass
        return data

    # ----- Cost Center CRUD -----
    def list_cost_centers(self, company_id=None):
        """Return cost centers filtered by company_id. Coerces/seeds on load."""
        rows = self.load_json(self.COST_CENTERS_FILENAME) or []
        rows = [r for r in rows if isinstance(r, dict)]
        if company_id:
            rows = [r for r in rows if r.get("company_id") == company_id]
        rows.sort(key=lambda r: ((r.get("label") or "").lower()))
        return rows

    def get_cost_center(self, company_id, cost_center_id):
        for r in self.list_cost_centers(company_id):
            if r.get("cost_center_id") == cost_center_id:
                return dict(r)
        return None

    def create_cost_center(self, company_id, cost_center_id, label, description="", category="OPERATING"):
        if not (company_id and cost_center_id and label):
            return False, "company_id, cost_center_id and label required", None
        rows = self.list_cost_centers(None)
        if any(r for r in rows if r.get("company_id") == company_id and r.get("cost_center_id") == cost_center_id):
            return False, f"Cost center {cost_center_id} already exists for company {company_id}", None
        row = {
            "company_id": company_id,
            "cost_center_id": cost_center_id,
            "label": label,
            "description": description or "",
            "category": category if category in ("OPERATING", "NON-OPERATING") else "OPERATING",
            "is_builtin": False,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        rows.append(row)
        self.save_json(self.COST_CENTERS_FILENAME, rows)
        return True, f"Created cost center {label}", dict(row)

    def update_cost_center(self, company_id, cost_center_id, updates):
        rows = self.list_cost_centers(None)
        idx = None
        for i, r in enumerate(rows):
            if r.get("company_id") == company_id and r.get("cost_center_id") == cost_center_id:
                idx = i; break
        if idx is None:
            return False, "Cost center not found", None
        for k, v in (updates or {}).items():
            if k in ("company_id", "cost_center_id", "created_at", "is_builtin"):
                continue
            rows[idx][k] = v
        self.save_json(self.COST_CENTERS_FILENAME, rows)
        return True, "Cost center updated", dict(rows[idx])

    def delete_cost_center(self, company_id, cost_center_id):
        rows = self.list_cost_centers(None)
        target = None
        for r in rows:
            if r.get("company_id") == company_id and r.get("cost_center_id") == cost_center_id:
                target = r; break
        if target is None:
            return False, "Not found"
        if target.get("is_builtin"):
            return False, "Built-in cost centers (Sales/Service Tech/Maint/General) cannot be deleted"
        new_rows = [r for r in rows if not (r.get("company_id") == company_id and r.get("cost_center_id") == cost_center_id)]
        self.save_json(self.COST_CENTERS_FILENAME, new_rows)
        return True, f"Deleted cost center {target.get('label')}"

    # ------------------------------------------------------------------
    # Service products (hourly rate catalogue)
    # ------------------------------------------------------------------
    @staticmethod
    def _default_service_products_for(company_id):
        """Four generic service items per company (FR-4). Users edit/prices as
        needed; just provides a non-empty starter catalogue."""
        return [
            {
                "company_id": company_id,
                "service_code": "GEN_CONSULT",
                "title": "General Consultation (Hour)",
                "description": "Business / technical advisory services by the hour.",
                "default_bill_rate_hour": 350.00,
                "default_internal_cost_hour": 120.00,
                "vat_rate": 5.0,
                "taxable": True,
                "default_cost_center_id": "cc_services_tech",
                "is_active": True,
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            {
                "company_id": company_id,
                "service_code": "FIELD_TECH_VISIT",
                "title": "Field Technician (Per Hour)",
                "description": "On-site technician visit, billed hourly.",
                "default_bill_rate_hour": 500.00,
                "default_internal_cost_hour": 150.00,
                "vat_rate": 5.0,
                "taxable": True,
                "default_cost_center_id": "cc_services_tech",
                "is_active": True,
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            {
                "company_id": company_id,
                "service_code": "MAINT_CONTRACT_Q",
                "title": "Quarterly Maintenance Contract (AMC)",
                "description": "Standard quarterly maintenance service (flat hour-bundle pricing).",
                "default_bill_rate_hour": 1200.00,
                "default_internal_cost_hour": 400.00,
                "vat_rate": 5.0,
                "taxable": True,
                "default_cost_center_id": "cc_services_maint",
                "is_active": True,
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
            {
                "company_id": company_id,
                "service_code": "TRAINING_1D",
                "title": "One-Day On-Site Training",
                "description": "Team training (1 day = 8 hours) on software / equipment / operations.",
                "default_bill_rate_hour": 900.00,
                "default_internal_cost_hour": 300.00,
                "vat_rate": 5.0,
                "taxable": True,
                "default_cost_center_id": "cc_services_tech",
                "is_active": True,
                "is_builtin": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            },
        ]

    def _ensure_service_products_seeded_for_company(self, company_id, save=False):
        if not company_id:
            return []
        rows = self.load_json(self.SERVICE_PRODUCTS_FILENAME) or []
        if not isinstance(rows, list):
            rows = []
        existing = {(r.get("company_id"), r.get("service_code")) for r in rows if isinstance(r, dict)}
        mutated = False
        for tpl in self._default_service_products_for(company_id):
            if (tpl["company_id"], tpl["service_code"]) not in existing:
                rows.append(tpl)
                existing.add((tpl["company_id"], tpl["service_code"]))
                mutated = True
        if mutated and save:
            self.save_json(self.SERVICE_PRODUCTS_FILENAME, rows)
        return rows

    def _coerce_and_seed_service_products(self, data):
        if not isinstance(data, list):
            data = []
        try:
            profiles = self.get_all_company_profiles() or []
        except Exception:
            profiles = []
        data = [r for r in data if isinstance(r, dict)]
        existing = {(r.get("company_id"), r.get("service_code")) for r in data}
        mutated = False
        for p in profiles:
            cid = p.get("id")
            if not cid:
                continue
            for tpl in self._default_service_products_for(cid):
                if (cid, tpl["service_code"]) not in existing:
                    data.append(tpl)
                    existing.add((cid, tpl["service_code"]))
                    mutated = True
        if mutated:
            try:
                self.save_json(self.SERVICE_PRODUCTS_FILENAME, data)
            except Exception:
                pass
        return data

    # ----- Service Products CRUD -----
    def list_service_products(self, company_id=None, include_inactive=False):
        rows = self.load_json(self.SERVICE_PRODUCTS_FILENAME) or []
        rows = [r for r in rows if isinstance(r, dict)]
        if company_id:
            rows = [r for r in rows if r.get("company_id") == company_id]
        if not include_inactive:
            rows = [r for r in rows if bool(r.get("is_active", True))]
        rows.sort(key=lambda r: ((r.get("title") or "").lower()))
        return rows

    def get_service_product(self, company_id, service_code):
        for r in self.list_service_products(company_id, include_inactive=True):
            if r.get("service_code") == service_code:
                return dict(r)
        return None

    def create_service_product(self, company_id, service_code, title, bill_rate, int_cost=0.0, vat_rate=5.0, taxable=True, cc_id=None, description=""):
        if not (company_id and service_code and title):
            return False, "company_id, service_code and title required", None
        rows = self.list_service_products(None, include_inactive=True)
        if any(r for r in rows if r.get("company_id") == company_id and r.get("service_code") == service_code):
            return False, f"Service code {service_code} already exists", None
        row = {
            "company_id": company_id,
            "service_code": service_code,
            "title": title,
            "description": description or "",
            "default_bill_rate_hour": float(bill_rate or 0.0),
            "default_internal_cost_hour": float(int_cost or 0.0),
            "vat_rate": float(vat_rate or 0.0),
            "taxable": bool(taxable),
            "default_cost_center_id": cc_id or "cc_services_tech",
            "is_active": True,
            "is_builtin": False,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        rows.append(row)
        self.save_json(self.SERVICE_PRODUCTS_FILENAME, rows)
        return True, f"Created service {title}", dict(row)

    def update_service_product(self, company_id, service_code, updates):
        rows = self.list_service_products(None, include_inactive=True)
        idx = None
        for i, r in enumerate(rows):
            if r.get("company_id") == company_id and r.get("service_code") == service_code:
                idx = i; break
        if idx is None:
            return False, "Service product not found", None
        for k, v in (updates or {}).items():
            if k in ("company_id", "service_code", "created_at", "is_builtin"):
                continue
            rows[idx][k] = v
        self.save_json(self.SERVICE_PRODUCTS_FILENAME, rows)
        return True, "Service product updated", dict(rows[idx])

    def delete_service_product(self, company_id, service_code):
        rows = self.list_service_products(None, include_inactive=True)
        target = None
        for r in rows:
            if r.get("company_id") == company_id and r.get("service_code") == service_code:
                target = r; break
        if target is None:
            return False, "Not found"
        if target.get("is_builtin"):
            # Built-in cannot be hard-deleted: flip is_active=False so it hides from pickers.
            target["is_active"] = False
            self.save_json(self.SERVICE_PRODUCTS_FILENAME, rows)
            return True, f"Built-in {target.get('service_code')} deactivated (hidden from menus)."
        new_rows = [r for r in rows if not (r.get("company_id") == company_id and r.get("service_code") == service_code)]
        self.save_json(self.SERVICE_PRODUCTS_FILENAME, new_rows)
        return True, f"Deleted service {target.get('title')}"

    # =====================================================================
    # Single-source financial aggregation (FR-5 compute_financials)
    # =====================================================================
    def compute_financials(self, company_id, date_from=None, date_to=None):
        """Single aggregate function used by Dashboard + ALL reports + PDFs (FR-5).

        Input date filters:
          - date_from / date_to may be None (= include everything)
            or strings "YYYY-MM-DD" or date/datetime objects. Invoices' `date`
            field is compared (invoice-date accrual per Assumption A1).

        Output dict (well-typed, safe for template rendering):
          - revenue_total (float)
          - revenue_by_cost_center (dict: cc_id -> {label, amount, invoice_count})
          - cogs_total (float)
          - cogs_by_cost_center (dict: cc_id -> amount)
          - gross_profit_total = revenue_total - cogs_total
          - gross_margin_pct = 0-100 or 0
          - sales_invoice_count, service_invoice_count
          - open_invoice_count, open_invoice_amount
          - purchases_total (from purchases_data.json, category mapped OPERATING)
          - salaries_total (from salary_reports.json, net_paid or gross_amount summed)
          - operating_expenses_total = salaries + purchases (best-effort with existing tables)
          - ebit = gross_profit - operating_expenses
          - net_profit (pre-tax = ebit - other_expenses if available; currently ebit)
        """
        from datetime import date as _date, datetime as _dt
        def _parse_date_str(s):
            if s is None:
                return None
            if isinstance(s, (_dt, _date)):
                return s
            if not isinstance(s, str):
                return None
            for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y"):
                try:
                    return _dt.strptime(s[:10], fmt).date()
                except Exception:
                    continue
            return None
        d_from = _parse_date_str(date_from)
        d_to = _parse_date_str(date_to)
        def _in_range(date_str):
            d = _parse_date_str(date_str)
            if d is None:
                return True  # unparseable/invalid → include (don't accidentally drop $)
            if d_from and d < d_from: return False
            if d_to and d > d_to: return False
            return True

        # Build cost-center label lookup once
        cc_label = {}
        for cc in self.list_cost_centers(company_id):
            cc_label[cc.get("cost_center_id")] = cc.get("label") or cc.get("cost_center_id") or "Unnamed"

        # Aggregate invoices
        inv_all = self.load_json("invoices_data.json") or []
        inv_company = [i for i in inv_all if isinstance(i, dict) and i.get("company_id") == company_id and not bool(i.get("deleted", False))]
        inv_period = [i for i in inv_company if _in_range(i.get("date"))]

        rev_total = 0.0
        cogs_total = 0.0
        sales_count = 0
        service_count = 0
        open_count = 0
        open_amount = 0.0
        rev_cc = {}   # cc_id -> {amount, invoice_count, label}
        cogs_cc = {}  # cc_id -> amount
        for inv in inv_period:
            sub = float(inv.get("subtotal") or 0.0)
            icost = float(inv.get("invoice_total_cost") or 0.0)
            rev_total += sub
            cogs_total += icost
            itype = (inv.get("invoice_type") or "sales").lower()
            if itype == "service":
                service_count += 1
            else:
                sales_count += 1
            status = (inv.get("status") or "Not Paid").strip().lower()
            paid = status in ("paid", "fully paid", "complete")
            if not paid:
                open_count += 1
                open_amount += float(inv.get("grand_total") or 0.0)
            cc = inv.get("cost_center_id") or "cc_general"
            r = rev_cc.setdefault(cc, {"amount": 0.0, "invoice_count": 0, "label": cc_label.get(cc, cc)})
            r["amount"] += sub
            r["invoice_count"] += 1
            cogs_cc[cc] = cogs_cc.get(cc, 0.0) + icost

        # Purchase aggregation (simple totals; categorize OPERATING by default since we don't have finer map)
        purchases = self.load_json("purchases_data.json") or []
        purchases_total = 0.0
        if isinstance(purchases, list):
            for p in purchases:
                if not isinstance(p, dict):
                    continue
                if p.get("company_id") and p.get("company_id") != company_id:
                    continue
                if not _in_range(p.get("date")):
                    continue
                purchases_total += float(p.get("amount") or p.get("grand_total") or 0.0)

        # Salary aggregation
        salaries = self.load_json("salary_reports.json") or []
        if not isinstance(salaries, list):
            salaries = []
        salaries_total = 0.0
        for s in salaries:
            if not isinstance(s, dict):
                continue
            if s.get("company_id") and s.get("company_id") != company_id:
                continue
            if not _in_range(s.get("pay_date") or s.get("period_end") or s.get("created_at")):
                continue
            # prefer net_paid then gross_amount then total
            salaries_total += float(
                s.get("net_paid") or s.get("gross_amount") or s.get("total") or s.get("net_salary") or 0.0
            )

        operating_expenses_total = round(purchases_total + salaries_total, 2)
        gross_profit_total = round(rev_total - cogs_total, 2)
        gross_margin_pct = round((gross_profit_total / rev_total) * 100.0, 2) if rev_total > 0.0001 else 0.0
        ebit = round(gross_profit_total - operating_expenses_total, 2)
        return {
            "company_id": company_id,
            "date_from": d_from.isoformat() if d_from else None,
            "date_to":   d_to.isoformat() if d_to else None,
            "revenue_total": round(rev_total, 2),
            "revenue_by_cost_center": rev_cc,
            "cogs_total": round(cogs_total, 2),
            "cogs_by_cost_center": cogs_cc,
            "gross_profit_total": gross_profit_total,
            "gross_margin_pct": gross_margin_pct,
            "sales_invoice_count": sales_count,
            "service_invoice_count": service_count,
            "open_invoice_count": open_count,
            "open_invoice_amount": round(open_amount, 2),
            "purchases_total": round(purchases_total, 2),
            "salaries_total": round(salaries_total, 2),
            "operating_expenses_total": operating_expenses_total,
            "ebit": ebit,
            "net_profit": ebit,  # pre-tax; placeholder; actual tax/debt out of scope Q3
        }

    def save_json(self, filename, data):
        """Save JSON data — Supabase REST when env vars set, plus local disk always."""
        disk_ok = False
        file_path = self.data_folder / filename
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            disk_ok = True
        except Exception as e:
            print(f"Error saving {filename}: {e}")
        sb_headers = self._supabase_headers()
        if sb_headers:
            sb_ok = self._supabase_save(filename, data)
            if sb_ok is False:
                print(f"[SupabaseSave] Fallback to disk only for {filename}")
                return disk_ok
            if sb_ok is True:
                return True
        return disk_ok

    def reload_all_json(self):
        """Force-reload every data file from disk.

        Used by the GAAP Financial Reporting Centre so that any data written
        *after* the dashboard was opened (e.g. a new invoice posted from the
        Invoicing module) becomes immediately visible on the reports without
        requiring the user to close and re-open the window.
        """
        import os
        data_files = [
            'invoices_data.json', 'purchases_data.json', 'quotes_data.json',
            'employees_data.json', 'salaries_data.json', 'employees.json',
            'salary_reports.json', 'transactions_data.json',
            'expenses_data.json', 'purchase_orders.json',
            'chart_of_accounts.json', 'general_ledger.json',
            'period_locks.json', 'app_settings.json',
            'stock_movements.json', 'supplier_ledger.json',
            'receivables_data.json', 'audit_trail.json',
        ]
        # --- 1) Re-read in-memory shortcuts that the GAAP dashboard / engine
        # may reference directly (via attribute access on self).
        gl = self.load_json('general_ledger.json') or []
        self.gl_entries = list(gl) if isinstance(gl, list) else []
        inv = self.load_json('invoices_data.json') or []
        self.invoices_data = list(inv) if isinstance(inv, list) else []
        exp = self.load_json('expenses_data.json') or []
        self.expenses = list(exp) if isinstance(exp, list) else []
        tx = self.load_json('transactions_data.json') or []
        self.transactions_data = list(tx) if isinstance(tx, list) else []
        po = self.load_json('purchase_orders.json') or []
        self.purchase_orders = list(po) if isinstance(po, list) else []
        emp = self.load_json('employees_data.json') or []
        self.employees_data = list(emp) if isinstance(emp, list) else []
        sal = self.load_json('salaries_data.json') or []
        self.salaries_data = list(sal) if isinstance(sal, list) else []
        # --- 2) Invalidate / refresh the LedgerService (if present) so its
        # next list_expenses / list_transactions call reads the new data.
        try:
            if getattr(self, 'ledger', None) and hasattr(self.ledger, 'reload_from_dm'):
                self.ledger.reload_from_dm()
        except Exception:
            pass
        return True

    # ============================================================
    # ASISTEM — Multi-Company Platform: Company Profiles
    # ============================================================
    #
    # The SYSTEM (ASISTEM) can host multiple independent companies.
    # Each invoice/expense/purchase belongs to one company profile.
    # HopePharma Medical Trading is company #1 (seeded by default).
    #
    # Data locations:
    #   company_profiles.json  → [{id, name, short_code, prefix, ...}, ...]
    #   app_state.json         → {"active_company_id": "com_hopepharma"}
    #
    # Invoice numbering prefix comes from the active company's `prefix`
    # (so HopePharma uses HPMT####, a second company could use ABCD####, etc.)

    # ============================================================
    # Data file locations (company-scoped table-per-file pattern)
    # Every entry here has a corresponding seed + coerce path in
    # load_json() post-process to keep old data files valid under
    # the ERP schema (see TR-1.2, TR-1.3, FR-1..4).
    # ============================================================
    COMPANIES_FILENAME = "company_profiles.json"
    APP_STATE_FILENAME = "app_state.json"
    COMPANY_USERS_FILENAME = "company_users.json"
    COST_CENTERS_FILENAME = "cost_centers.json"
    SERVICE_PRODUCTS_FILENAME = "service_products.json"

    DEFAULT_COMPANY_ID = "com_hopepharma"

    @classmethod
    def _default_company_seed(cls):
        """The built-in first company: HopePharma Medical Trading.

        Seeded with ALL fields required by the ERP/CRM upgrade so the
        existing HopePharma tenant retains the same letterhead, bank
        details and TRN that were previously hardcoded inside the PDF
        renderer (see FR-6 PDF refactor Task 2).
        """
        seed = {
            "id": cls.DEFAULT_COMPANY_ID,
            "name": "HopePharma Medical Trading",
            "display_name": "HopePharma Medical Trading",
            "legal_name": "HopePharma Medicine Trading LLC",
            "short_code": "HPMT",
            "prefix": "HPMT",  # e.g. HopePharma uses HPMT#### invoice numbers
            "trn": "100466797600003",  # UAE 15-digit TRN (matches original PDF hardcode)
            "vat_rate": 5.0,           # Default VAT % for this company
            "currency": "AED",
            "emirate": "Dubai",
            "address": "Dubai, International City, Morocco Cluster, Bldg. I-16, Store 08",
            "phone": "",
            "email": "",
            "website": "",
            "logo_path": None,
            # PDF branding images (stored as base64 data: URIs, nullable)
            "logo_png_base64": None,
            "signature_png_base64": None,
            "stamp_png_base64": None,
            # PDF letterhead: banking details (appear in Terms & Conditions)
            "bank_name": "ADCB",
            "bank_branch": "259 Dubai Mall, Dubai",
            "bank_account_no": "14198059920001",
            "bank_iban": "AE980030014198059920001",
            "bank_swift": "ADCBAEAAXXX",
            # PDF behaviour toggles
            "pdf_footer_thank_you": "Thank you for your business!",
            "show_cost_center_on_pdf": False,  # default hide to not leak cost structure
            "is_default": True,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        return seed

    def _ensure_company_profiles_seeded(self):
        """Make sure company_profiles.json + app_state.json exist.

        Backwards-compatible: if file is missing/empty, create HopePharma.
        """
        # --- company_profiles.json ---
        profiles = self.load_json(self.COMPANIES_FILENAME) or []
        if not isinstance(profiles, list):
            profiles = []

        # Always ensure the HopePharma default company is present
        has_default = any(isinstance(p, dict) and p.get("id") == self.DEFAULT_COMPANY_ID
                          for p in profiles)
        if not has_default:
            profiles.insert(0, self._default_company_seed())
            self.save_json(self.COMPANIES_FILENAME, profiles)

        # --- app_state.json ---
        app_state = self.load_json(self.APP_STATE_FILENAME) or {}
        if not isinstance(app_state, dict):
            app_state = {}
        changed = False
        if "active_company_id" not in app_state or not app_state["active_company_id"]:
            app_state["active_company_id"] = self.DEFAULT_COMPANY_ID
            changed = True
        # Also ensure the stored active_company_id actually exists
        ids = {p.get("id") for p in profiles if isinstance(p, dict)}
        if app_state.get("active_company_id") not in ids:
            app_state["active_company_id"] = self.DEFAULT_COMPANY_ID
            changed = True
        if changed:
            self.save_json(self.APP_STATE_FILENAME, app_state)
        return profiles, app_state

    # ----- CRUD for companies -----
    def get_all_company_profiles(self):
        """Return a stable list of Company Profile dicts (sorted by default first, then name)."""
        self._ensure_company_profiles_seeded()
        profiles = self.load_json(self.COMPANIES_FILENAME) or []
        if not isinstance(profiles, list):
            profiles = []
        valid = [p for p in profiles if isinstance(p, dict) and p.get("id")]
        valid.sort(key=lambda p: (
            0 if p.get("is_default") else 1,
            (p.get("name") or "").lower(),
        ))
        return valid

    def get_active_company_profile(self):
        """Return the currently-active company (the one whose invoices the user is working on)."""
        profiles, state = self._ensure_company_profiles_seeded()
        active_id = (state or {}).get("active_company_id") or self.DEFAULT_COMPANY_ID
        for p in profiles:
            if isinstance(p, dict) and p.get("id") == active_id:
                return dict(p)
        # Safety: fall back to default
        for p in profiles:
            if isinstance(p, dict) and p.get("id") == self.DEFAULT_COMPANY_ID:
                return dict(p)
        return self._default_company_seed()

    def get_company_profile_by_id(self, company_id):
        for p in self.get_all_company_profiles():
            if p.get("id") == company_id:
                return dict(p)
        return None

    def set_active_company(self, company_id):
        """Switch the active company (used by invoices created after this call)."""
        profiles = self.get_all_company_profiles()
        ids = {p.get("id") for p in profiles}
        if company_id not in ids:
            return False, f"Company ID {company_id} does not exist"
        state = self.load_json(self.APP_STATE_FILENAME) or {}
        if not isinstance(state, dict):
            state = {}
        state["active_company_id"] = company_id
        self.save_json(self.APP_STATE_FILENAME, state)
        return True, f"Active company switched to: {self.get_company_profile_by_id(company_id).get('name')}"

    def create_company_profile(self, profile_data):
        """Create a new company profile. profile_data must include 'name'."""
        name = (profile_data or {}).get("name")
        if not isinstance(name, str) or not name.strip():
            return False, "Company name is required", None
        profiles = self.get_all_company_profiles()
        import re
        def _slug(s):
            s = (s or "").lower()
            s = re.sub(r"[^a-z0-9]+", "_", s).strip("_")
            return s or "company"
        base = _slug(name)
        cid = f"com_{base}"
        existing_ids = {p.get("id") for p in profiles}
        suffix = 2
        while cid in existing_ids:
            cid = f"com_{base}_{suffix}"
            suffix += 1
        # Derive a default 4-letter uppercase prefix for invoice numbering
        letters = [ch for ch in name.upper() if ch.isalpha()]
        prefix = (profile_data or {}).get("prefix") or (
            "".join(letters[:4]) if len(letters) >= 4 else (name[:4].upper() if len(name) >= 4 else (name.upper() + "XXXX")[:4])
        )
        short_code = (profile_data or {}).get("short_code") or prefix[:4]
        new_company = {
            "id": cid,
            "name": name.strip(),
            "display_name": (profile_data or {}).get("display_name") or name.strip(),
            "legal_name": (profile_data or {}).get("legal_name") or name.strip(),
            "short_code": short_code.upper(),
            "prefix": prefix.upper(),
            "trn": (profile_data or {}).get("trn", ""),
            "vat_rate": float((profile_data or {}).get("vat_rate", 5) or 5),
            "currency": (profile_data or {}).get("currency", "AED"),
            "emirate": (profile_data or {}).get("emirate", "Dubai"),
            "address": (profile_data or {}).get("address", ""),
            "phone": (profile_data or {}).get("phone", ""),
            "email": (profile_data or {}).get("email", ""),
            "website": (profile_data or {}).get("website", ""),
            "logo_path": (profile_data or {}).get("logo_path"),
            # PDF branding images (base64 data URIs) — nullable
            "logo_png_base64": (profile_data or {}).get("logo_png_base64"),
            "signature_png_base64": (profile_data or {}).get("signature_png_base64"),
            "stamp_png_base64": (profile_data or {}).get("stamp_png_base64"),
            # PDF letterhead banking details (optional but recommended for Terms & Conditions block)
            "bank_name": (profile_data or {}).get("bank_name", ""),
            "bank_branch": (profile_data or {}).get("bank_branch", ""),
            "bank_account_no": (profile_data or {}).get("bank_account_no", ""),
            "bank_iban": (profile_data or {}).get("bank_iban", ""),
            "bank_swift": (profile_data or {}).get("bank_swift", ""),
            # PDF behaviour toggles
            "pdf_footer_thank_you": (profile_data or {}).get("pdf_footer_thank_you", "Thank you for your business!"),
            "show_cost_center_on_pdf": bool((profile_data or {}).get("show_cost_center_on_pdf", False)),
            "is_default": False,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        profiles.append(new_company)
        self.save_json(self.COMPANIES_FILENAME, profiles)
        # Also seed cost centers + service products for this new company so the
        # admin doesn't land on an empty Invoices screen (see TR-1.3, FR-2, FR-4).
        try:
            self._ensure_cost_centers_seeded_for_company(cid, save=True)
            self._ensure_service_products_seeded_for_company(cid, save=True)
        except Exception as _e:
            print(f"[create_company_profile] post-create seeding warning: {_e}")
        return True, f"Created company: {new_company['name']} (ID {cid})", new_company

    def update_company_profile(self, company_id, updates):
        """Update an existing company profile (partial patch)."""
        profiles = self.get_all_company_profiles()
        idx = None
        for i, p in enumerate(profiles):
            if p.get("id") == company_id:
                idx = i; break
        if idx is None:
            return False, f"Company {company_id} not found", None
        for key, value in (updates or {}).items():
            # Don't allow changing id/is_default/created_at
            if key in {"id", "created_at"}:
                continue
            profiles[idx][key] = value
        if "name" in (updates or {}) and not profiles[idx].get("display_name"):
            profiles[idx]["display_name"] = updates["name"]
        self.save_json(self.COMPANIES_FILENAME, profiles)
        return True, f"Company {profiles[idx].get('name')} updated", dict(profiles[idx])

    def delete_company_profile(self, company_id, allow_delete_default=False):
        """Remove a company profile. Default company can only be removed if allow_delete_default=True."""
        profiles = self.get_all_company_profiles()
        target = next((p for p in profiles if p.get("id") == company_id), None)
        if target is None:
            return False, f"Company {company_id} not found"
        if target.get("is_default") and not allow_delete_default:
            return False, "Cannot delete the default (HopePharma) company"
        # Don't allow deletion if there are active invoices for this company (only hard ones; soft deleted still exist but can be cleaned)
        invoices = self.load_json('invoices_data.json') or []
        used = any(isinstance(inv, dict) and inv.get("company_id") == company_id
                   for inv in invoices)
        if used:
            return False, f"Cannot delete: invoice data is already attached to {target.get('name')}. Archive/erase its invoices first."
        profiles = [p for p in profiles if p.get("id") != company_id]
        self.save_json(self.COMPANIES_FILENAME, profiles)
        state = self.load_json(self.APP_STATE_FILENAME) or {}
        if isinstance(state, dict) and state.get("active_company_id") == company_id:
            state["active_company_id"] = self.DEFAULT_COMPANY_ID
            self.save_json(self.APP_STATE_FILENAME, state)
        # Also delete all users belonging to this company to avoid orphans
        users = self.load_json(self.COMPANY_USERS_FILENAME) or []
        if isinstance(users, list):
            surviving = [u for u in users if isinstance(u, dict) and u.get("company_id") != company_id]
            if len(surviving) != len(users):
                self.save_json(self.COMPANY_USERS_FILENAME, surviving)
        return True, f"Deleted company: {target.get('name')}"

    # ============================================================
    # Company Users (per-company login auth)
    # ============================================================
    COMPANY_USERS_FILENAME = "company_users.json"

    @staticmethod
    def _hash_password(plain):
        """Securely hash a plaintext password for storage. Uses Werkzeug's
        industry-standard salted PBKDF2 (150k iters) if available; falls back
        to hashlib pbkdf2_hmac + random 16-byte salt if not."""
        plain = (plain or "").encode("utf-8")
        try:
            from werkzeug.security import generate_password_hash  # Werkzeug is a Flask dep
            return generate_password_hash(plain.decode("utf-8"), method="pbkdf2:sha256:150000")
        except Exception:
            # Fallback: do it manually with a random salt
            import base64, hashlib, os as _os
            salt = _os.urandom(16)
            dk = hashlib.pbkdf2_hmac("sha256", plain, salt, 150000)
            return "pbkdf2_fallback$" + base64.b64encode(salt).decode("ascii") + "$" + base64.b64encode(dk).decode("ascii")

    @staticmethod
    def _check_password(plain, stored_hash):
        """Verify a plaintext password against the stored hash format."""
        try:
            if not plain or not stored_hash:
                return False
            if isinstance(stored_hash, str) and stored_hash.startswith("pbkdf2_fallback$"):
                import base64, hashlib
                parts = stored_hash.split("$", 2)
                if len(parts) != 3:
                    return False
                salt = base64.b64decode(parts[1].encode("ascii"))
                expected = base64.b64decode(parts[2].encode("ascii"))
                got = hashlib.pbkdf2_hmac("sha256", (plain or "").encode("utf-8"), salt, 150000)
                return len(got) == len(expected) and all(a == b for a, b in zip(got, expected))
            from werkzeug.security import check_password_hash
            return check_password_hash(stored_hash, plain)
        except Exception:
            return False

    def _ensure_company_users_seeded(self):
        """Create company_users.json if missing, seed HopePharma's default user.

        HopePharma credentials (per user request):
            username: HopePharma
            password: Hope2025
            company_id: com_hopepharma (default)
            role: admin
        """
        # Make sure the company profile exists first
        self._ensure_company_profiles_seeded()
        users = self.load_json(self.COMPANY_USERS_FILENAME)
        if not isinstance(users, list):
            users = []
        changed = False

        # --- Seed HopePharma default user (exactly as user requested) ---
        hp_user_exists = any(isinstance(u, dict) and (u.get("username") or "").lower() == "hopepharma"
                             for u in users)
        if not hp_user_exists:
            users.append({
                "username": "HopePharma",
                "company_id": self.DEFAULT_COMPANY_ID,
                "password_hash": self._hash_password("Hope2025"),
                "role": "admin",
                "full_name": "HopePharma Medical Trading Admin",
                "email": "",
                "is_active": True,
                "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "last_login_at": None,
            })
            changed = True

        # Sanitize: drop any user whose company_id no longer exists
        valid_company_ids = {p["id"] for p in self.get_all_company_profiles()}
        before = len(users)
        users = [u for u in users if isinstance(u, dict) and u.get("company_id") in valid_company_ids]
        if len(users) != before:
            changed = True

        if changed:
            self.save_json(self.COMPANY_USERS_FILENAME, users)
        return users

    # ----- Company User CRUD -----
    def list_company_users(self, company_id=None):
        """Return stable list of user dicts (password_hash stripped for safety)."""
        users = self._ensure_company_users_seeded()
        result = []
        for u in users:
            if not isinstance(u, dict) or not u.get("username"):
                continue
            if company_id and u.get("company_id") != company_id:
                continue
            safe = {k: v for k, v in u.items() if k != "password_hash"}
            result.append(safe)
        result.sort(key=lambda u: ((u.get("company_id") or "").lower(), (u.get("username") or "").lower()))
        return result

    def get_company_user_by_username(self, username):
        """Return user dict (WITH password_hash) by username (case-insensitive lookup)."""
        if not username:
            return None
        users = self._ensure_company_users_seeded()
        needle = str(username).lower()
        for u in users:
            if isinstance(u, dict) and str(u.get("username") or "").lower() == needle:
                return dict(u)
        return None

    def verify_company_user(self, username, password):
        """Validate username/password. Returns SAFE user dict (no password_hash) if OK, else None."""
        u = self.get_company_user_by_username(username)
        if not u:
            return None
        if not u.get("is_active", True):
            return None
        if not self._check_password(password, u.get("password_hash")):
            return None
        # Record last login time
        users = self.load_json(self.COMPANY_USERS_FILENAME) or []
        for i, rec in enumerate(users):
            if isinstance(rec, dict) and str(rec.get("username") or "").lower() == str(username).lower():
                users[i]["last_login_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self.save_json(self.COMPANY_USERS_FILENAME, users)
                break
        safe = {k: v for k, v in u.items() if k != "password_hash"}
        return safe

    def create_company_user(self, username, password, company_id, role="user", full_name=None, email=None):
        """Create a new user for a company. username is case-insensitive unique."""
        if not (username and password and company_id):
            return False, "username, password, and company_id are required", None
        company = self.get_company_profile_by_id(company_id)
        if not company:
            return False, f"Company ID {company_id} does not exist", None
        users = self._ensure_company_users_seeded()
        if any(isinstance(u, dict) and str(u.get("username") or "").lower() == str(username).lower()
               for u in users):
            return False, f"Username '{username}' is already taken", None
        new_user = {
            "username": str(username),
            "company_id": company_id,
            "password_hash": self._hash_password(password),
            "role": str(role or "user"),
            "full_name": str(full_name or username),
            "email": str(email or ""),
            "is_active": True,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "last_login_at": None,
        }
        users.append(new_user)
        self.save_json(self.COMPANY_USERS_FILENAME, users)
        safe = {k: v for k, v in new_user.items() if k != "password_hash"}
        return True, f"Created user '{safe['username']}' for company '{company.get('name')}'", safe

    def update_company_user_password(self, username, new_password):
        if not (username and new_password):
            return False, "username and new password are required"
        users = self.load_json(self.COMPANY_USERS_FILENAME) or []
        needle = str(username).lower()
        for i, u in enumerate(users):
            if isinstance(u, dict) and str(u.get("username") or "").lower() == needle:
                users[i]["password_hash"] = self._hash_password(new_password)
                self.save_json(self.COMPANY_USERS_FILENAME, users)
                return True, f"Password updated for user '{username}'"
        return False, f"User '{username}' not found"

    def delete_company_user(self, username):
        if not username:
            return False, "username required"
        users = self.load_json(self.COMPANY_USERS_FILENAME) or []
        needle = str(username).lower()
        before = len(users)
        users = [u for u in users if not (isinstance(u, dict) and str(u.get("username") or "").lower() == needle)]
        if len(users) == before:
            return False, f"User '{username}' not found"
        self.save_json(self.COMPANY_USERS_FILENAME, users)
        return True, f"Deleted user '{username}'"

    def toggle_company_user_active(self, username):
        """Flip is_active flag for a user (soft disable without deleting)."""
        if not username:
            return False, "username required", None
        users = self.load_json(self.COMPANY_USERS_FILENAME) or []
        needle = str(username).lower()
        for i, u in enumerate(users):
            if isinstance(u, dict) and str(u.get("username") or "").lower() == needle:
                new_val = not bool(u.get("is_active", True))
                users[i]["is_active"] = new_val
                try:
                    users[i]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass
                self.save_json(self.COMPANY_USERS_FILENAME, users)
                status = "ACTIVE" if new_val else "DISABLED"
                return True, f"User '{username}' status is now {status}", dict(users[i])
        return False, f"User '{username}' not found", None

    def update_company_user(self, username, updates):
        """Update fields of an existing user (role, full_name, email, is_active)."""
        if not username or not isinstance(updates, dict):
            return False, "username + updates dict required", None
        users = self.load_json(self.COMPANY_USERS_FILENAME) or []
        needle = str(username).lower()
        ALLOWED = {"role", "full_name", "email", "is_active"}
        for i, u in enumerate(users):
            if isinstance(u, dict) and str(u.get("username") or "").lower() == needle:
                for k, v in updates.items():
                    if k not in ALLOWED:
                        continue
                    if k == "role":
                        if v not in ("admin", "staff"):
                            continue
                    if k == "is_active":
                        v = bool(v)
                    users[i][k] = v
                try:
                    users[i]["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                except Exception:
                    pass
                self.save_json(self.COMPANY_USERS_FILENAME, users)
                return True, f"User '{username}' updated", dict(users[i])
        return False, f"User '{username}' not found", None

    # ----- Helpers that stamp company onto an invoice dict -----
    def _stamp_active_company(self, invoice_dict):
        """Attach the active company_id + prefix (derived from profile) to an invoice dict.

        Preserves any already-stamped company_id on legacy invoices by honoring it if missing.
        """
        if not isinstance(invoice_dict, dict):
            return invoice_dict
        active = self.get_active_company_profile()
        if not invoice_dict.get("company_id"):
            invoice_dict["company_id"] = active.get("id")
        if not invoice_dict.get("company_name"):
            invoice_dict["company_name"] = active.get("display_name") or active.get("name")
        if not invoice_dict.get("company_prefix"):
            invoice_dict["company_prefix"] = active.get("prefix") or "HPMT"
        # currency from company if not set
        if not invoice_dict.get("currency") or invoice_dict["currency"] == "AED":
            invoice_dict["currency"] = active.get("currency") or "AED"
        return invoice_dict

    # Invoice management methods
    def generate_invoice_id(self, company_id=None):
        """Generate unique gap-filling company-prefixed 4-digit invoice ID.

        - Default company (HopePharma) uses HPMT####.
        - For a different company, the prefix comes from the profile: e.g. ABCD####.
        - Prefix is either the active company's prefix OR the company_id passed-in explicitly.
        - Scan ACTIVE (non-deleted) invoices matching `prefix + digits` ONLY for the
          SAME company_id (so numbering is independent per-company).
        - Pick the lowest integer >= 1 not used. 4-digit zero-padding always.
        """
        try:
            if company_id:
                company = self.get_company_profile_by_id(company_id) or self.get_active_company_profile()
            else:
                company = self.get_active_company_profile()
            prefix = (company.get("prefix") or "HPMT").strip().upper()
            company_id_now = company.get("id")
            invoices = self.load_json('invoices_data.json') or []

            used = set()
            for inv in invoices:
                if not isinstance(inv, dict):
                    continue
                if inv.get('deleted'):
                    continue
                # Only consider invoices belonging to THIS company (fallback: any matching prefix for legacy unbranded data)
                inv_company = inv.get("company_id")
                if inv_company and inv_company != company_id_now:
                    continue
                iid = inv.get('invoice_id')
                if isinstance(iid, str) and iid.startswith(prefix) and iid[len(prefix):].isdigit():
                    used.add(int(iid[len(prefix):]))

            if not used:
                new_number = 1
            else:
                m = max(used)
                candidate = None
                for n in range(1, m + 1):
                    if n not in used:
                        candidate = n
                        break
                new_number = candidate if candidate is not None else (m + 1)

            invoice_id = f"{prefix}{new_number:04d}"
            print(f"[generate_invoice_id] company={company.get('name')} prefix={prefix} "
                  f"active-used={sorted(used)} -> next = {invoice_id}")
            return invoice_id

        except Exception as e:
            print(f"[generate_invoice_id] ERROR: {e}")
            import traceback
            traceback.print_exc()
            return f"HPMT{datetime.now().strftime('%Y%m%d%H%M%S')}"
    
    def add_invoice_from_dict(self, invoice_dict):
        """Add invoice from dictionary"""
        try:
            # ASISTEM: Ensure company_id + other profile fields are stamped
            self._stamp_active_company(invoice_dict)
            invoices = self.load_json('invoices_data.json')
            
            # Check if invoice already exists
            existing_idx = None
            for i, inv in enumerate(invoices):
                if inv.get('invoice_id') == invoice_dict.get('invoice_id'):
                    existing_idx = i
                    break

            if existing_idx is not None:
                previous = dict(invoices[existing_idx]) if isinstance(invoices[existing_idx], dict) else {}
                invoices[existing_idx] = invoice_dict
                # Safety: always ensure any reused/reactivated invoice explicitly clears its deleted flags
                # (prevents a gap-fill invoice over a soft-deleted same-ID slot from inheriting old deleted=True)
                if isinstance(invoices[existing_idx], dict):
                    if not invoices[existing_idx].get('deleted', False):
                        invoices[existing_idx]['deleted'] = False
                        invoices[existing_idx]['deleted_at'] = None
                        invoices[existing_idx]['deleted_by'] = None
                saved = self.save_json('invoices_data.json', invoices)
                if saved and self.ledger:
                    try:
                        self.ledger.on_invoice_created(invoice_dict, username="System")
                    except Exception as le:
                        print(f"Ledger hook warning (invoice update): {le}")
                    try:
                        # Invoice-level costs (Dr 5xxx / Cr 1100)
                        new_costs = invoice_dict.get('costs') or []
                        prev_costs = previous.get('costs') or []
                        if new_costs or prev_costs:
                            self.ledger.on_invoice_costs_changed(
                                invoice_dict.get('invoice_id'), new_costs,
                                invoice_date=invoice_dict.get('date'),
                                invoice_client=invoice_dict.get('client_name', ''),
                                username="System",
                                operation="update",
                                previous_costs_list=prev_costs,
                            )
                    except Exception as le2:
                        print(f"Ledger invoice-costs hook warning (update): {le2}")
                return saved
            
            # Add new invoice
            invoices.append(invoice_dict)
            saved = self.save_json('invoices_data.json', invoices)
            if saved and self.ledger:
                try:
                    self.ledger.on_invoice_created(invoice_dict, username="System")
                except Exception as le:
                    print(f"Ledger hook warning (new invoice): {le}")
                try:
                    new_costs = invoice_dict.get('costs') or []
                    if new_costs:
                        self.ledger.on_invoice_costs_changed(
                            invoice_dict.get('invoice_id'), new_costs,
                            invoice_date=invoice_dict.get('date'),
                            invoice_client=invoice_dict.get('client_name', ''),
                            username="System",
                            operation="create",
                            previous_costs_list=[],
                        )
                except Exception as le2:
                    print(f"Ledger invoice-costs hook warning (create): {le2}")
            return saved
            
        except Exception as e:
            print(f"Error adding invoice: {e}")
            return False

    def update_invoice_partial(self, invoice_id, partial_dict):
        """
        Update specific fields of an existing invoice (without replacing
        the whole dict).  Returns (ok, message).
        """
        try:
            if not invoice_id or not isinstance(partial_dict, dict):
                return False, "Invalid arguments"
            invoices = self.load_json('invoices_data.json')
            for i, inv in enumerate(invoices):
                if not (isinstance(inv, dict) and inv.get('invoice_id') == invoice_id):
                    continue
                previous = dict(inv)
                inv.update(partial_dict)
                # Recompute derived totals if costs or grand_total updated
                try:
                    if 'costs' in partial_dict or 'total_cost' in partial_dict:
                        c_total = Decimal("0.00")
                        for c in (inv.get('costs') or []):
                            c_total += Decimal(str(c.get('amount') or 0))
                        c_total = c_total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                        inv['total_cost'] = float(c_total)
                        gt = Decimal(str(inv.get('grand_total') or inv.get('total') or 0))
                        inv['profit_loss'] = float((gt - c_total).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
                except Exception:
                    pass
                invoices[i] = inv
                saved = self.save_json('invoices_data.json', invoices)
                if not saved:
                    return False, "Failed to save invoices"
                if self.ledger:
                    try:
                        self.ledger.on_invoice_created(inv, username="System")
                    except Exception as le:
                        print(f"Ledger hook warning (partial update): {le}")
                    try:
                        self.ledger.on_invoice_costs_changed(
                            inv.get('invoice_id'), inv.get('costs') or [],
                            invoice_date=inv.get('date'),
                            invoice_client=inv.get('client_name', ''),
                            username="System",
                            operation="partial_update",
                            previous_costs_list=previous.get('costs') or [],
                        )
                    except Exception as le2:
                        print(f"Ledger costs hook warning (partial update): {le2}")
                return True, "Invoice updated"
            return False, f"Invoice {invoice_id} not found"
        except Exception as e:
            return False, f"Error: {e}"
    
    def get_invoice(self, invoice_id, include_deleted=True):
        """Get invoice by ID. If include_deleted=False, soft-deleted invoices return None."""
        invoices = self.load_json('invoices_data.json')
        for invoice_data in invoices:
            if isinstance(invoice_data, dict) and invoice_data.get('invoice_id') == invoice_id:
                if (not include_deleted) and invoice_data.get('deleted'):
                    return None
                return InvoiceObject(invoice_data)
        return None

    def get_all_invoices_dict(self, include_deleted=False):
        """Get invoices as a list of dicts. By default soft-deleted invoices are filtered out.

        Set include_deleted=True to include the archive (deleted=True) invoices too.
        """
        try:
            invoices = self.load_json('invoices_data.json') or []
            valid_invoices = []
            for inv in invoices:
                if not isinstance(inv, dict):
                    continue
                if inv.get('deleted') and (not include_deleted):
                    continue
                valid_invoice = {
                    'invoice_id': inv.get('invoice_id', ''),
                    'client_name': inv.get('client_name', ''),
                    'client_trn': inv.get('client_trn', ''),
                    'client_emirate': (inv.get('client_emirate') or ''),
                    'client_location': (inv.get('client_location') or ''),
                    'date': inv.get('date', ''),
                    'due_date': inv.get('due_date', ''),
                    'payment_terms': inv.get('payment_terms', inv.get('payment_method', 'Cash')),
                    'payment_method': inv.get('payment_method', inv.get('payment_terms', 'Cash')),
                    'payment_due': inv.get('payment_due', 'On receipt'),
                    'tax_rate': inv.get('tax_rate', 5),
                    'items': inv.get('items', []),
                    'costs': inv.get('costs', []),
                    'subtotal': inv.get('subtotal', 0),
                    'taxable_amount': inv.get('taxable_amount', 0),
                    'non_taxable_amount': inv.get('non_taxable_amount', 0),
                    'tax_amount': inv.get('tax_amount', 0),
                    'grand_total': inv.get('grand_total', 0),
                    'total_cost': inv.get('total_cost', 0),
                    'profit_loss': inv.get('profit_loss', 0),
                    'total_paid': inv.get('total_paid', 0),
                    'balance_due': inv.get('balance_due', 0),
                    'status': inv.get('status', 'Not Paid'),
                    'notes': inv.get('notes', ''),
                    'currency': inv.get('currency', 'AED'),
                    'invoice_type': inv.get('invoice_type', 'sales'),
                    'payment_history': inv.get('payment_history', []),
                    'workflow_status': inv.get('workflow_status', 'Under Process'),
                    'deleted': bool(inv.get('deleted', False)),
                    'deleted_at': inv.get('deleted_at', None),
                    'deleted_by': inv.get('deleted_by', None),
                    # ASISTEM multi-company profile stamp (preserve whatever is on disk)
                    'company_id': inv.get('company_id', None),
                    'company_name': inv.get('company_name', None),
                    'company_prefix': inv.get('company_prefix', None),
                }
                valid_invoices.append(valid_invoice)
            return valid_invoices
        except Exception as e:
            print(f"ERROR in get_all_invoices_dict: {e}")
            return []

    def get_deleted_invoices_dict(self):
        """Return only soft-deleted (archived) invoices as list of dicts."""
        all_incl = self.get_all_invoices_dict(include_deleted=True)
        return [inv for inv in all_incl if inv.get('deleted')]

    def search_invoices(self, query, include_deleted=False):
        """Search invoices by various criteria. By default soft-deleted are excluded."""
        invoices = self.load_json('invoices_data.json') or []
        results = []
        query_lower = str(query).lower()

        for invoice_data in invoices:
            if not isinstance(invoice_data, dict):
                continue
            if invoice_data.get('deleted') and (not include_deleted):
                continue
            match = False
            if (query_lower in str(invoice_data.get('invoice_id', '')).lower() or
                query_lower in str(invoice_data.get('client_name', '')).lower() or
                query_lower in str(invoice_data.get('client_trn', '')).lower()):
                match = True
            if not match:
                for item in invoice_data.get('items', []):
                    if query_lower in str(item.get('description', '')).lower():
                        match = True
                        break
            if match:
                results.append(InvoiceObject(invoice_data))

        return results

    def soft_delete_invoice(self, invoice_id, username=None):
        """Soft-delete an invoice (move to archive, hides from active views, frees its number).

        Sets deleted=True, deleted_at=ISO timestamp, deleted_by=<username>. Returns (bool, message).
        """
        try:
            invoices = self.load_json('invoices_data.json') or []
            found_idx = None
            for i, inv in enumerate(invoices):
                if isinstance(inv, dict) and inv.get('invoice_id') == invoice_id:
                    found_idx = i
                    break
            if found_idx is None:
                return False, f"Invoice {invoice_id} not found"
            target = dict(invoices[found_idx])
            if target.get('deleted'):
                return False, f"Invoice {invoice_id} is already in the deleted archive"
            target['deleted'] = True
            target['deleted_at'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            target['deleted_by'] = username or "Unknown User"
            invoices[found_idx] = target
            ok = self.save_json('invoices_data.json', invoices)
            if not ok:
                return False, "Failed to save invoices file during soft-delete"
            return True, f"Invoice {invoice_id} moved to Deleted Archive (soft-deleted)"
        except Exception as e:
            return False, f"Error soft-deleting invoice {invoice_id}: {e}"

    def restore_invoice(self, invoice_id):
        """Restore a soft-deleted invoice back to active (unmark deleted=True flag)."""
        try:
            invoices = self.load_json('invoices_data.json') or []
            found_idx = None
            for i, inv in enumerate(invoices):
                if isinstance(inv, dict) and inv.get('invoice_id') == invoice_id:
                    found_idx = i
                    break
            if found_idx is None:
                return False, f"Invoice {invoice_id} not found"
            target = dict(invoices[found_idx])
            if not target.get('deleted'):
                return False, f"Invoice {invoice_id} is not in the deleted archive"
            target['deleted'] = False
            target['deleted_at'] = None
            target['deleted_by'] = None
            invoices[found_idx] = target
            ok = self.save_json('invoices_data.json', invoices)
            if not ok:
                return False, "Failed to save invoices file during restore"
            return True, f"Invoice {invoice_id} restored from Deleted Archive"
        except Exception as e:
            return False, f"Error restoring invoice {invoice_id}: {e}"

    def permanent_erase_invoice(self, invoice_id):
        """Permanently, physically erase a (usually deleted-archive) invoice. Unrecoverable.

        Prefer soft_delete_invoice for normal use. Use this only for GDPR/cleanup erasure.
        """
        return self.delete_invoice(invoice_id)
    
    def delete_invoice(self, invoice_id):
        """Delete invoice by ID"""
        try:
            invoices = self.load_json('invoices_data.json')
            removed = None
            kept = []
            for inv in invoices:
                if isinstance(inv, dict) and inv.get('invoice_id') == invoice_id:
                    removed = inv
                else:
                    kept.append(inv)
            ok = self.save_json('invoices_data.json', kept)
            if not ok:
                return False
            try:
                sales_path = self.data_folder / 'sales_records.json'
                inventory_path = self.data_folder / 'inventory.json'
                if sales_path.exists() and inventory_path.exists():
                    import json
                    with open(sales_path, 'r') as f:
                        sales = json.load(f)
                    with open(inventory_path, 'r') as f:
                        inventory = json.load(f)
                    modified = False
                    remaining_sales = []
                    for sale in sales:
                        if isinstance(sale, dict) and sale.get('invoice_id') == invoice_id:
                            item_id = sale.get('item_id')
                            qty = int(sale.get('quantity', 0))
                            promo = int(sale.get('promotion_quantity', 0))
                            total_units = qty + promo
                            cpi = float(sale.get('cost_per_item', 0))
                            for item in inventory:
                                if (item.get('item_id') == item_id) or (item.get('name') == sale.get('item_name')):
                                    item['quantity'] = int(item.get('quantity', 0)) + total_units
                                    item['total_cost'] = float(item.get('total_cost', 0)) + (cpi * total_units)
                                    item['last_updated'] = datetime.now().strftime("%Y-%m-%d")
                                    modified = True
                                    break
                            # Skip adding this sale (remove record)
                        else:
                            remaining_sales.append(sale)
                    if modified:
                        with open(inventory_path, 'w') as f:
                            json.dump(inventory, f, indent=2)
                    with open(sales_path, 'w') as f:
                        json.dump(remaining_sales, f, indent=2)
            except Exception:
                pass
            return True
        except Exception as e:
            print(f"Error deleting invoice: {e}")
            return False
    
    def update_invoice_id(self, old_id, new_id):
        """Update invoice ID"""
        try:
            invoices = self.load_json('invoices_data.json')
            
            # Check if new ID already exists
            for inv in invoices:
                if isinstance(inv, dict) and inv.get('invoice_id') == new_id:
                    return False, "Invoice ID already exists"
            
            # Update the ID
            for inv in invoices:
                if isinstance(inv, dict) and inv.get('invoice_id') == old_id:
                    inv['invoice_id'] = new_id
                    break
            
            if self.save_json('invoices_data.json', invoices):
                try:
                    # Also update linked sales records to keep inventory reversal working
                    sales_path = self.data_folder / 'sales_records.json'
                    if sales_path.exists():
                        import json
                        with open(sales_path, 'r') as f:
                            sales = json.load(f)
                        changed = False
                        for s in sales:
                            try:
                                if isinstance(s, dict) and s.get('invoice_id') == old_id:
                                    s['invoice_id'] = new_id
                                    changed = True
                            except Exception:
                                pass
                        if changed:
                            with open(sales_path, 'w') as f:
                                json.dump(sales, f, indent=2)
                except Exception:
                    pass
                return True, "Invoice ID updated successfully"
            else:
                return False, "Failed to save changes"
                
        except Exception as e:
            return False, f"Error: {str(e)}"
    
    def save_invoices(self):
        """Save invoices (compatibility method)"""
        return True  # Data is already saved in add_invoice_from_dict
    
    def update_invoice(self, invoice_obj):
        """Update invoice object"""
        return self.add_invoice_from_dict(invoice_obj.to_dict())

    def edit_invoice_item(self, invoice_id, index, description, quantity, unit_price, taxable=True):
        try:
            inv = self.get_invoice(invoice_id)
            if not inv:
                return False
            ok = inv.edit_item(index, description, quantity, unit_price, taxable)
            if ok:
                return self.update_invoice(inv)
            return False
        except Exception:
            return False

    def edit_invoice_cost(self, invoice_id, index, description, amount, account=None, notes=None):
        try:
            inv = self.get_invoice(invoice_id)
            if not inv:
                return False
            ok = inv.edit_cost(index, description, amount, account, notes)
            if ok:
                return self.update_invoice(inv)
            return False
        except Exception:
            return False
    
    def get_clients(self):
        """Get unique client names"""
        invoices = self.load_json('invoices_data.json')
        clients = set()
        for inv in invoices:
            if isinstance(inv, dict):
                client = inv.get('client_name')
                if client:
                    clients.add(client)
        return list(clients)
    
    def get_sales_report(self, start_date, end_date):
        """Generate sales report for date range"""
        try:
            invoices = self.load_json('invoices_data.json')
            report_invoices = []
            total_sales = 0
            total_paid = 0
            total_cost = 0
            total_tax = 0
            
            for inv in invoices:
                if isinstance(inv, dict):
                    inv_date = inv.get('date', '')
                    if start_date <= inv_date <= end_date:
                        report_invoices.append(inv)
                        total_sales += inv.get('grand_total', 0)
                        total_paid += inv.get('total_paid', 0)
                        total_cost += inv.get('total_cost', 0)
                        total_tax += inv.get('tax_amount', 0)
            
            return {
                'start_date': start_date,
                'end_date': end_date,
                'total_invoices': len(report_invoices),
                'total_sales': total_sales,
                'total_paid': total_paid,
                'total_cost': total_cost,
                'total_tax': total_tax,
                'total_profit': total_sales - total_cost,
                'outstanding_balance': total_sales - total_paid,
                'invoices': report_invoices
            }
        except Exception as e:
            print(f"Error generating sales report: {e}")
            return {}
    
    def get_client_report(self, client_name, start_date, end_date):
        """Generate client-specific report"""
        try:
            invoices = self.load_json('invoices_data.json')
            client_invoices = []
            total_sales = 0
            total_paid = 0
            total_cost = 0
            total_tax = 0
            
            for inv in invoices:
                if isinstance(inv, dict):
                    inv_date = inv.get('date', '')
                    inv_client = inv.get('client_name', '')
                    if (inv_client == client_name and 
                        start_date <= inv_date <= end_date):
                        client_invoices.append(inv)
                        total_sales += inv.get('grand_total', 0)
                        total_paid += inv.get('total_paid', 0)
                        total_cost += inv.get('total_cost', 0)
                        total_tax += inv.get('tax_amount', 0)
            
            return {
                'client_name': client_name,
                'start_date': start_date,
                'end_date': end_date,
                'total_invoices': len(client_invoices),
                'total_sales': total_sales,
                'total_paid': total_paid,
                'total_cost': total_cost,
                'total_tax': total_tax,
                'total_profit': total_sales - total_cost,
                'outstanding_balance': total_sales - total_paid,
                'invoices': client_invoices
            }
        except Exception as e:
            print(f"Error generating client report: {e}")
            return {}
    
    def get_company_report(self, start_date, end_date):
        """Generate company financial report"""
        try:
            invoices = self.load_json('invoices_data.json')
            purchases = self.load_json('purchases_data.json')
            
            # Calculate totals
            total_revenue = sum(inv.get('grand_total', 0) for inv in invoices 
                              if isinstance(inv, dict) and start_date <= inv.get('date', '') <= end_date)
            total_expenses = sum(p.get('amount', 0) for p in purchases 
                               if isinstance(p, dict) and start_date <= p.get('date', '') <= end_date)
            net_profit = total_revenue - total_expenses
            net_profit_margin = (net_profit / total_revenue * 100) if total_revenue > 0 else 0
            
            # Payment status breakdown
            payment_status = {'paid': 0, 'partial': 0, 'unpaid': 0}
            for inv in invoices:
                if isinstance(inv, dict) and start_date <= inv.get('date', '') <= end_date:
                    status = inv.get('status', '').lower()
                    if 'paid' in status:
                        if 'partial' in status:
                            payment_status['partial'] += 1
                        else:
                            payment_status['paid'] += 1
                    else:
                        payment_status['unpaid'] += 1
            
            # Top clients
            client_revenue = {}
            for inv in invoices:
                if isinstance(inv, dict) and start_date <= inv.get('date', '') <= end_date:
                    client = inv.get('client_name', 'Unknown')
                    client_revenue[client] = client_revenue.get(client, 0) + inv.get('grand_total', 0)
            
            top_clients = sorted(client_revenue.items(), key=lambda x: x[1], reverse=True)[:5]
            
            # Expense categories
            expense_categories = {}
            for p in purchases:
                if isinstance(p, dict) and start_date <= p.get('date', '') <= end_date:
                    category = p.get('category', 'Other')
                    expense_categories[category] = expense_categories.get(category, 0) + p.get('amount', 0)
            
            # Outstanding balance
            outstanding_balance = sum(inv.get('grand_total', 0) - inv.get('total_paid', 0) 
                                    for inv in invoices 
                                    if isinstance(inv, dict) and start_date <= inv.get('date', '') <= end_date)
            
            return {
                'start_date': start_date,
                'end_date': end_date,
                'total_invoices': len([inv for inv in invoices if isinstance(inv, dict) and start_date <= inv.get('date', '') <= end_date]),
                'total_purchases': len([p for p in purchases if isinstance(p, dict) and start_date <= p.get('date', '') <= end_date]),
                'total_revenue': total_revenue,
                'total_expenses': total_expenses,
                'net_profit': net_profit,
                'net_profit_margin': net_profit_margin,
                'outstanding_balance': outstanding_balance,
                'payment_status_breakdown': payment_status,
                'top_clients': top_clients,
                'expense_categories': expense_categories,
                'profitability_status': 'Profitable' if net_profit > 0 else 'Not Profitable',
                'revenue_health': 'Healthy' if total_revenue > total_expenses else 'Needs Attention'
            }
        except Exception as e:
            print(f"Error generating company report: {e}")
            return {}

    # Purchase management methods
    def create_purchase(self, **kwargs):
        """Create a new purchase"""
        try:
            purchases = self.load_json('purchases_data.json')
            purchase_id = f"PUR-{datetime.now().strftime('%Y%m%d')}-{len(purchases) + 1:03d}"
            
            amount_raw = kwargs.get('amount', 0)
            try:
                amount_val = float(amount_raw)
            except Exception:
                amount_val = 0.0
            paid_raw = kwargs.get('amount_paid', kwargs.get('paid_amount',
                               kwargs.get('payment_amount', kwargs.get('paid', 0.0))))
            try:
                paid_val = float(paid_raw) if paid_raw is not None else 0.0
            except Exception:
                paid_val = 0.0
            if paid_val < 0:
                paid_val = 0.0
            if paid_val > amount_val + 0.005:
                paid_val = amount_val
            
            # auto status
            if paid_val <= 0:
                status = 'Pending'
            elif paid_val >= amount_val - 0.005:
                status = 'Paid'
            else:
                status = 'Partial'
            
            purchase_data = {
                'purchase_id': purchase_id,
                'description': kwargs.get('description', ''),
                'amount': amount_val,
                'date': kwargs.get('date', datetime.now().strftime("%Y-%m-%d")),
                'category': kwargs.get('category', 'General'),
                'supplier': kwargs.get('supplier', ''),
                'account': kwargs.get('account', 'Cash'),
                'receipt_attached': False,
                'receipt_filename': '',
                'notes': kwargs.get('notes', ''),
                'amount_paid': round(paid_val, 2),
                'payment_date': kwargs.get('payment_date', ''),
                'bank_account': kwargs.get('bank_account', ''),
                'paid_status': status
            }
            
            purchases.append(purchase_data)
            if self.save_json('purchases_data.json', purchases):
                if self.ledger:
                    try:
                        self.ledger.on_purchase_received(purchase_data, username="System")
                    except Exception as le:
                        print(f"Ledger hook warning (purchase): {le}")
                return PurchaseObject(purchase_data)
            return None
        except Exception as e:
            print(f"Error creating purchase: {e}")
            return None

    def record_purchase_payment(self, purchase_id, payment_amount, payment_date=None,
                                bank_account='Emirates NBD', username='System'):
        """
        Record a payment against an UNPAID/PARTIAL purchase:
          1. Updates purchase.amount_paid + paid_status (Pending → Partial → Paid)
          2. Posts GL via ledger.on_purchase_payment_made (Dr 2000 AP / Cr 1100 Bank)
          3. Writes audit_log.jsonl line
        Returns (ok: bool, message: str)
        """
        try:
            payment_amount = float(payment_amount or 0)
            if payment_amount <= 0:
                return False, "Payment amount must be greater than zero."
            if not payment_date:
                payment_date = datetime.now().strftime("%Y-%m-%d")
            try:
                datetime.strptime(payment_date, "%Y-%m-%d")
            except Exception:
                return False, "Invalid payment date. Use YYYY-MM-DD."

            purchases = self.load_json('purchases_data.json')
            idx = None
            for i, p in enumerate(purchases):
                if isinstance(p, dict) and p.get('purchase_id') == purchase_id:
                    idx = i; break
            if idx is None:
                return False, f"Purchase {purchase_id} not found."
            p = purchases[idx]
            total = float(p.get('amount', 0) or 0)
            prev_paid = float(p.get('amount_paid', 0) or 0)
            if prev_paid >= total - 0.005:
                return False, f"Purchase {purchase_id} is already fully paid."
            new_paid = prev_paid + payment_amount
            if new_paid > total + 0.005:
                # clamp to total (allow tiny overpayment tolerance)
                payment_amount = max(total - prev_paid, 0.0)
                if payment_amount <= 0:
                    return False, f"Purchase {purchase_id} is already fully paid."
                new_paid = total

            if new_paid <= 0:
                status = 'Pending'
            elif new_paid >= total - 0.005:
                status = 'Paid'
            else:
                status = 'Partial'

            purchases[idx]['amount_paid'] = round(new_paid, 2)
            purchases[idx]['payment_date'] = payment_date
            purchases[idx]['bank_account'] = bank_account
            purchases[idx]['paid_status'] = status

            ok_save = self.save_json('purchases_data.json', purchases)
            if not ok_save:
                return False, "Failed to persist payment to purchases_data.json"

            # GL posting
            if self.ledger:
                try:
                    self.ledger.on_purchase_payment_made(
                        purchase_id=purchase_id,
                        amount=payment_amount,
                        payment_date=payment_date,
                        supplier=p.get('supplier', ''),
                        username=username,
                    )
                except Exception as le:
                    print(f"Ledger hook warning (purchase payment): {le}")

            # Audit log
            try:
                folder = getattr(self, 'invoice_folder', os.path.join(os.path.expanduser('~'), 'HopePharmaData'))
                path = os.path.join(folder, 'audit_log.jsonl')
                with open(path, 'a') as f:
                    f.write(json.dumps({
                        'timestamp': datetime.now().isoformat(),
                        'action': 'purchase_payment_made',
                        'details': {
                            'purchase_id': purchase_id,
                            'supplier': p.get('supplier', ''),
                            'amount_total': total,
                            'amount_previously_paid': prev_paid,
                            'amount_this_payment': payment_amount,
                            'amount_total_paid_after': new_paid,
                            'payment_date': payment_date,
                            'bank_account': bank_account,
                            'paid_status_after': status,
                            'username': username
                        }
                    }) + "\n")
            except Exception:
                pass

            return True, (f"Payment of AED {payment_amount:.2f} recorded successfully.\n"
                          f"Supplier: {p.get('supplier') or 'N/A'}\n"
                          f"Paid status now: {status}")
        except Exception as e:
            return False, f"Error recording purchase payment: {e}"

    def process_purchase_with_balance(self, purchase_data):
        try:
            from balance_manager import BalanceManager
            bal = BalanceManager(self.invoice_folder)
            return bal.process_purchase(purchase_data)
        except Exception as e:
            return False, str(e)

    def process_invoice_payment_with_balance(self, invoice_data, payment_amount, account_name):
        try:
            from balance_manager import BalanceManager
            bal = BalanceManager(self.invoice_folder)
            return bal.process_invoice_payment(invoice_data, payment_amount, account_name)
        except Exception as e:
            return False, str(e)

    def reverse_invoice_payment_with_balance(self, invoice_data, payment_amount, account_name):
        try:
            from balance_manager import BalanceManager
            bal = BalanceManager(self.invoice_folder)
            return bal.reverse_invoice_payment(invoice_data, payment_amount, account_name)
        except Exception as e:
            return False, str(e)

    def delete_purchase_with_balance_adjustment(self, purchase_id):
        try:
            purchase = self.get_purchase(purchase_id)
            if not purchase:
                return False, "Purchase not found"
            data = purchase.to_dict() if hasattr(purchase, 'to_dict') else purchase
            amount = float(data.get('amount', 0) or 0)
            account = data.get('account', 'Cash')
            if amount > 0:
                from balance_manager import BalanceManager
                bal = BalanceManager(self.invoice_folder)
                try:
                    date_str = data.get('date') or datetime.now().strftime("%Y-%m-%d")
                except Exception:
                    date_str = datetime.now().strftime("%Y-%m-%d")
                bal.reverse_purchase(str(purchase_id), {"account": account, "amount": amount, "date": date_str})
            ok = self.delete_purchase(purchase_id)
            if ok:
                return True, f"Purchase {purchase_id} deleted successfully. Balance adjusted."
            return False, "Failed to delete purchase"
        except Exception as e:
            return False, str(e)

    def delete_invoice_with_balance_adjustment(self, invoice_id):
        try:
            inv = self.get_invoice(invoice_id)
            if not inv:
                return False, "Invoice not found"
            # Normalize to dict
            if hasattr(inv, 'to_dict'):
                inv_dict = inv.to_dict()
            elif isinstance(inv, dict):
                inv_dict = inv
            else:
                inv_dict = {'invoice_id': getattr(inv, 'invoice_id', invoice_id)}

            # Reverse each payment to the same account it was deposited into
            payments = inv_dict.get('payment_history') or inv_dict.get('payments') or []
            if payments:
                from balance_manager import BalanceManager
                bal = BalanceManager(self.invoice_folder)
                for p in payments:
                    try:
                        amount = float(p.get('amount', 0) or 0)
                        if amount <= 0:
                            continue
                        method = p.get('method') or p.get('payment_method') or 'Cash'
                        account = method
                        if account not in bal.get_account_names():
                            # Fallback to Cash only if the specific account isn't found
                            account = 'Cash'
                        ok, msg = bal.reverse_invoice_payment(inv_dict, amount, account)
                        if not ok:
                            return False, msg
                    except Exception as e:
                        return False, str(e)
            else:
                # Fallback: reverse total_paid to Cash if no payment records exist
                total_paid = float(inv_dict.get('total_paid', 0) or 0)
                if total_paid > 0:
                    from balance_manager import BalanceManager
                    bal = BalanceManager(self.invoice_folder)
                    ok, msg = bal.reverse_invoice_payment(inv_dict, total_paid, 'Cash')
                    if not ok:
                        return False, msg
            ok = self.delete_invoice(invoice_id)
            if ok:
                return True, f"Invoice {invoice_id} deleted successfully. Balance adjusted."
            return False, "Failed to delete invoice"
        except Exception as e:
            return False, str(e)

    def soft_delete_invoice_with_balance_adjustment(self, invoice_id, username=None):
        """Soft-delete an invoice: reverse payments, then mark deleted=True (moves to archive, frees its HPMT number)."""
        try:
            inv = self.get_invoice(invoice_id, include_deleted=True)
            if not inv:
                return False, "Invoice not found"
            if hasattr(inv, 'to_dict'):
                inv_dict = inv.to_dict()
            elif isinstance(inv, dict):
                inv_dict = inv
            else:
                inv_dict = {'invoice_id': getattr(inv, 'invoice_id', invoice_id)}

            if inv_dict.get('deleted'):
                return False, f"Invoice {invoice_id} is already in the deleted archive"

            # Reverse each payment (same logic as delete_invoice_with_balance_adjustment)
            payments = inv_dict.get('payment_history') or inv_dict.get('payments') or []
            if payments:
                from balance_manager import BalanceManager
                bal = BalanceManager(self.invoice_folder)
                for p in payments:
                    try:
                        amount = float(p.get('amount', 0) or 0)
                        if amount <= 0:
                            continue
                        method = p.get('method') or p.get('payment_method') or 'Cash'
                        account = method
                        if account not in bal.get_account_names():
                            account = 'Cash'
                        ok_r, msg_r = bal.reverse_invoice_payment(inv_dict, amount, account)
                        if not ok_r:
                            return False, msg_r
                    except Exception as e:
                        return False, str(e)
            else:
                total_paid = float(inv_dict.get('total_paid', 0) or 0)
                if total_paid > 0:
                    from balance_manager import BalanceManager
                    bal = BalanceManager(self.invoice_folder)
                    ok_r, msg_r = bal.reverse_invoice_payment(inv_dict, total_paid, 'Cash')
                    if not ok_r:
                        return False, msg_r

            # SOFT delete (mark flag, keep row so audit trail and permanent archive exist)
            ok_s, msg_s = self.soft_delete_invoice(invoice_id, username=username)
            if not ok_s:
                return False, msg_s
            return True, f"Invoice {invoice_id} archived (soft-deleted) — payments reversed, number {invoice_id} now available for re-use."
        except Exception as e:
            return False, str(e)
    
    def _rollback_inventory_and_sales(self, invoice_id: str, inv_dict: Optional[Dict[str, Any]] = None):
        try:
            sales_path = self.data_folder / 'sales_records.json'
            inventory_path = self.data_folder / 'inventory.json'
            if not (sales_path.exists() and inventory_path.exists()):
                return False
            import json
            with open(sales_path, 'r') as f:
                sales = json.load(f)
            with open(inventory_path, 'r') as f:
                inventory = json.load(f)
            remaining_sales = []
            modified = False
            used_links = False
            if inv_dict and isinstance(inv_dict, dict):
                links = inv_dict.get('inventory_links') or []
                for link in links:
                    try:
                        item_id = link.get('item_id')
                        qty = int(link.get('quantity', 0))
                        promo = int(link.get('promotion_quantity', 0))
                        total_units = qty + promo
                        cpi = float(link.get('cost_per_item', 0))
                        for item in inventory:
                            if (item.get('item_id') == item_id) or (item.get('name') == link.get('item_name')):
                                item['quantity'] = int(item.get('quantity', 0)) + total_units
                                item['total_cost'] = float(item.get('total_cost', 0)) + (cpi * total_units)
                                item['last_updated'] = datetime.now().strftime("%Y-%m-%d")
                                modified = True
                                used_links = True
                                break
                    except Exception:
                        pass
            for sale in sales:
                if isinstance(sale, dict) and sale.get('invoice_id') == invoice_id:
                    item_id = sale.get('item_id')
                    qty = int(sale.get('quantity', 0))
                    promo = int(sale.get('promotion_quantity', 0))
                    total_units = qty + promo
                    cpi = float(sale.get('cost_per_item', 0))
                    for item in inventory:
                        if (item.get('item_id') == item_id) or (item.get('name') == sale.get('item_name')):
                            try:
                                item['quantity'] = int(item.get('quantity', 0)) + total_units
                                item['total_cost'] = float(item.get('total_cost', 0)) + (cpi * total_units)
                                item['last_updated'] = datetime.now().strftime("%Y-%m-%d")
                                modified = True
                                break
                            except Exception:
                                pass
                    # skip adding sale (remove it)
                else:
                    remaining_sales.append(sale)
            if modified:
                with open(inventory_path, 'w') as f:
                    json.dump(inventory, f, indent=2)
            with open(sales_path, 'w') as f:
                json.dump(remaining_sales, f, indent=2)
            return True
        except Exception:
            return False
    
    def get_purchase(self, purchase_id):
        """Get purchase by ID"""
        purchases = self.load_json('purchases_data.json')
        for purchase_data in purchases:
            if isinstance(purchase_data, dict) and purchase_data.get('purchase_id') == purchase_id:
                return PurchaseObject(purchase_data)
        return None
    
    def get_all_purchases_dict(self):
        """Get all purchases as dictionaries - FIXED VERSION"""
        try:
            purchases = self.load_json('purchases_data.json')
            print(f"DEBUG: Found {len(purchases)} purchases in file")
            
            # Filter out any non-dictionary items and ensure proper structure
            valid_purchases = []
            for p in purchases:
                if isinstance(p, dict):
                    amount_val = float(p.get('amount', 0) or 0)
                    paid_val = float(p.get('amount_paid',
                                      p.get('paid_amount',
                                            p.get('payment_amount', 0))) or 0)
                    status = p.get('paid_status') or p.get('payment_status') or ''
                    if not status:
                        if paid_val <= 0:
                            status = 'Pending'
                        elif paid_val >= amount_val - 0.005:
                            status = 'Paid'
                        else:
                            status = 'Partial'
                    valid_purchase = {
                        'purchase_id': p.get('purchase_id', ''),
                        'description': p.get('description', ''),
                        'amount': amount_val,
                        'date': p.get('date', ''),
                        'category': p.get('category', 'General'),
                        'supplier': p.get('supplier', ''),
                        'account': p.get('account', 'Emirates NBD'),
                        'receipt_attached': p.get('receipt_attached', False),
                        'receipt_filename': p.get('receipt_filename', ''),
                        'notes': p.get('notes', ''),
                        'amount_paid': round(paid_val, 2),
                        'payment_date': p.get('payment_date', ''),
                        'bank_account': p.get('bank_account', ''),
                        'paid_status': status
                    }
                    valid_purchases.append(valid_purchase)
                    print(f"DEBUG: Added purchase {valid_purchase['purchase_id']} status={status}")
            print(f"DEBUG: Returning {len(valid_purchases)} valid purchases")
            return valid_purchases
        except Exception as e:
            print(f"ERROR in get_all_purchases_dict: {e}")
            return []
    
    def search_purchases(self, query):
        """Search purchases by various criteria"""
        purchases = self.load_json('purchases_data.json')
        results = []
        query_lower = query.lower()
        
        for purchase_data in purchases:
            if isinstance(purchase_data, dict):
                if (query_lower in purchase_data.get('purchase_id', '').lower() or
                    query_lower in purchase_data.get('description', '').lower() or
                    query_lower in purchase_data.get('supplier', '').lower() or
                    query_lower in purchase_data.get('category', '').lower()):
                    results.append(PurchaseObject(purchase_data))
        
        return results
    
    def delete_purchase(self, purchase_id):
        """Delete purchase by ID"""
        try:
            purchases = self.load_json('purchases_data.json')
            purchases = [p for p in purchases if isinstance(p, dict) and p.get('purchase_id') != purchase_id]
            return self.save_json('purchases_data.json', purchases)
        except Exception as e:
            print(f"Error deleting purchase: {e}")
            return False
    
    def update_purchase(self, purchase_id, **kwargs):
        """Update purchase information"""
        try:
            purchases = self.load_json('purchases_data.json')
            for i, purchase in enumerate(purchases):
                if isinstance(purchase, dict) and purchase.get('purchase_id') == purchase_id:
                    for key, value in kwargs.items():
                        if value is not None:
                            purchases[i][key] = value
                    return self.save_json('purchases_data.json', purchases)
            return False
        except Exception as e:
            print(f"Error updating purchase: {e}")
            return False
    
    def upload_purchase_receipt(self, purchase_id, file_path):
        """Upload receipt for purchase"""
        try:
            # Copy receipt to receipts folder
            receipts_folder = self.data_folder / 'receipts'
            filename = f"{purchase_id}_{os.path.basename(file_path)}"
            dest_path = receipts_folder / filename
            
            shutil.copy2(file_path, dest_path)
            
            # Update purchase record
            purchases = self.load_json('purchases_data.json')
            for i, purchase in enumerate(purchases):
                if isinstance(purchase, dict) and purchase.get('purchase_id') == purchase_id:
                    purchases[i]['receipt_attached'] = True
                    purchases[i]['receipt_filename'] = filename
                    return self.save_json('purchases_data.json', purchases)
            
            return False
        except Exception as e:
            print(f"Error uploading purchase receipt: {e}")
            return False
    
    def open_purchase_receipt(self, purchase_id):
        """Open purchase receipt file"""
        try:
            purchases = self.load_json('purchases_data.json')
            for purchase in purchases:
                if isinstance(purchase, dict) and purchase.get('purchase_id') == purchase_id and purchase.get('receipt_attached'):
                    filename = purchase.get('receipt_filename')
                    receipt_path = self.data_folder / 'receipts' / filename
                    if receipt_path.exists():
                        if platform.system() == "Darwin":  # macOS
                            subprocess.run(["open", str(receipt_path)])
                        elif platform.system() == "Windows":
                            os.startfile(str(receipt_path))
                        else:  # Linux
                            subprocess.run(["xdg-open", str(receipt_path)])
                        return True
            return False
        except Exception as e:
            print(f"Error opening purchase receipt: {e}")
            return False

    # Receipt management methods for invoices
    def upload_receipt(self, invoice_id, cost_index, file_path):
        """Upload receipt for invoice cost"""
        try:
            # Copy receipt to receipts folder
            receipts_folder = self.data_folder / 'receipts'
            filename = f"{invoice_id}_cost_{cost_index}_{os.path.basename(file_path)}"
            dest_path = receipts_folder / filename
            
            shutil.copy2(file_path, dest_path)
            
            # Update invoice record
            invoices = self.load_json('invoices_data.json')
            for i, invoice in enumerate(invoices):
                if isinstance(invoice, dict) and invoice.get('invoice_id') == invoice_id:
                    # Ensure costs list exists and has enough elements
                    if 'costs' not in invoice:
                        invoice['costs'] = []
                    
                    while len(invoice['costs']) <= cost_index:
                        invoice['costs'].append({})
                    
                    invoice['costs'][cost_index]['receipt_attached'] = True
                    invoice['costs'][cost_index]['receipt_filename'] = filename
                    return self.save_json('invoices_data.json', invoices)
            
            return False
        except Exception as e:
            print(f"Error uploading receipt: {e}")
            return False
    
    def open_receipt(self, invoice_id, cost_index):
        """Open receipt file for invoice cost"""
        try:
            invoices = self.load_json('invoices_data.json')
            for invoice in invoices:
                if (isinstance(invoice, dict) and 
                    invoice.get('invoice_id') == invoice_id and 
                    'costs' in invoice and 
                    len(invoice['costs']) > cost_index and
                    invoice['costs'][cost_index].get('receipt_attached')):
                    
                    filename = invoice['costs'][cost_index].get('receipt_filename')
                    receipt_path = self.data_folder / 'receipts' / filename
                    
                    if receipt_path.exists():
                        if platform.system() == "Darwin":  # macOS
                            subprocess.run(["open", str(receipt_path)])
                        elif platform.system() == "Windows":
                            os.startfile(str(receipt_path))
                        else:  # Linux
                            subprocess.run(["xdg-open", str(receipt_path)])
                        return True
            return False
        except Exception as e:
            print(f"Error opening receipt: {e}")
            return False

# Use the enhanced manager
CloudDataManager = EnhancedCloudDataManager

class HopePharmaComplete:
    """Enhanced HopePharma application with better UI and features"""
    
    def __init__(self, root):
        self.root = root
        self.system = platform.system()
        self.setup_enhanced_ui()
        
    def setup_enhanced_ui(self):
        """Setup enhanced user interface"""
        self.root.title("🏥 HopePharma - Complete Cloud Edition")
        self.root.geometry("600x500")
        self.root.configure(bg='#f0f8ff')  # Light blue background
        
        # Center the window
        self.root.eval('tk::PlaceWindow . center')
        
        # Make window resizable
        self.root.minsize(500, 400)
        
        main_frame = ttk.Frame(self.root, padding="30")
        main_frame.pack(fill='both', expand=True)
        
        # Header with better styling
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(pady=20)
        
        title = ttk.Label(header_frame, text="🏥 HopePharma", 
                         font=('Helvetica', 24, 'bold'), foreground='#2c3e50')
        title.pack(pady=5)
        
        subtitle = ttk.Label(header_frame, text="Complete Cloud Business Management System", 
                           font=('Helvetica', 12), foreground='#7f8c8d')
        subtitle.pack(pady=2)
        
        # OS detection with icon
        os_icons = {
            "Windows": "🪟",
            "Darwin": "🍎", 
            "Linux": "🐧"
        }
        os_icon = os_icons.get(self.system, "💻")
        
        os_label = ttk.Label(header_frame, 
                           text=f"{os_icon} Detected: {self.system}", 
                           font=('Helvetica', 10, 'bold'),
                           foreground='#3498db')
        os_label.pack(pady=10)
        
        # Instructions with better formatting
        instructions = ttk.Label(main_frame, 
            text="Choose your preferred synchronization method:\nAll options ensure your data is safe and accessible.",
            justify='center', font=('Helvetica', 10))
        instructions.pack(pady=20)
        
        # Buttons frame with better layout
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=30, fill='x', padx=50)
        
        # Cloud button (Recommended) - Enhanced
        cloud_frame = ttk.Frame(button_frame)
        cloud_frame.pack(fill='x', pady=10)
        
        cloud_btn = ttk.Button(cloud_frame, text="☁️ Google Drive Sync", 
                              command=self.setup_google_drive,
                              width=25)
        cloud_btn.pack(side='left', padx=(0, 10))
        
        ttk.Label(cloud_frame, text="• Automatic sync across devices\n• Secure cloud storage", 
                 foreground="#27ae60", font=('Helvetica', 9)).pack(side='left')
        
        # Local network button - Enhanced
        network_frame = ttk.Frame(button_frame)
        network_frame.pack(fill='x', pady=10)
        
        network_btn = ttk.Button(network_frame, text="🔗 Local Network", 
                                command=self.setup_network,
                                width=25)
        network_btn.pack(side='left', padx=(0, 10))
        
        ttk.Label(network_frame, text="• Fast local network sharing\n• No internet required", 
                 foreground="#e67e22", font=('Helvetica', 9)).pack(side='left')
        
        # Single computer button - Enhanced
        single_frame = ttk.Frame(button_frame)
        single_frame.pack(fill='x', pady=10)
        
        single_btn = ttk.Button(single_frame, text="💻 Single Computer", 
                               command=self.setup_single,
                               width=25)
        single_btn.pack(side='left', padx=(0, 10))
        
        ttk.Label(single_frame, text="• Simple setup\n• All data stored locally", 
                 foreground="#95a5a6", font=('Helvetica', 9)).pack(side='left')
        
        # Status with better visibility
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(pady=20)
        
        self.status = ttk.Label(status_frame, text="✅ Ready to configure your system...", 
                              font=('Helvetica', 10, 'bold'), foreground='#2c3e50')
        self.status.pack()
        
        # Help section
        help_text = """
        💡 Tip: For multiple computers, choose Google Drive Sync.
        Your data will automatically sync across all devices.
        """
        help_label = ttk.Label(main_frame, text=help_text, 
                              justify='center', foreground='#34495e',
                              font=('Helvetica', 9))
        help_label.pack(pady=20)
        
        # Version info
        version_label = ttk.Label(main_frame, text="HopePharma v2.1 • Enhanced Edition", 
                                 font=('Helvetica', 8), foreground='#bdc3c7')
        version_label.pack(side='bottom', pady=10)
    
    def setup_google_drive(self):
        """Enhanced Google Drive setup with better guidance"""
        messagebox.showinfo("Google Drive Setup", 
            "🚀 BEST OPTION for multiple computers!\n\n"
            "Step 1: Create a folder called 'HopePharmaData' in your Google Drive\n"
            "Step 2: Select that folder when prompted\n"
            "Step 3: Repeat on other computers using the same folder\n\n"
            "✅ Automatic synchronization across all devices!\n"
            "✅ Secure cloud backup\n"
            "✅ Access from anywhere")
        
        self.status.config(text="🔄 Selecting Google Drive folder...")
        
        folder = filedialog.askdirectory(title="Select 'HopePharmaData' folder in Google Drive")
        if folder:
            self.start_app(folder, "cloud")
        else:
            # Enhanced auto-detection
            self._auto_detect_google_drive()
    
    def _auto_detect_google_drive(self):
        """Enhanced Google Drive auto-detection"""
        drive_locations = []
        
        if self.system == "Windows":
            drive_locations = [
                Path("~/Google Drive/My Drive/HopePharmaData").expanduser(),
                Path("~/Google Drive/HopePharmaData").expanduser(),
                Path("~/Documents/HopePharmaData").expanduser(),
            ]
        else:  # macOS and Linux
            drive_locations = [
                Path("~/Google Drive/My Drive/HopePharmaData").expanduser(),
                Path("~/Google Drive/HopePharmaData").expanduser(),
                Path("~/HopePharmaData").expanduser(),
                Path("~/Documents/HopePharmaData").expanduser(),
            ]
        
        for location in drive_locations:
            if location.parent.exists():
                location.mkdir(exist_ok=True)
                self.status.config(text=f"✅ Using auto-detected location: {location}")
                self.root.update()
                time.sleep(1)
                self.start_app(str(location), "cloud")
                return
        
        # Create in default location
        default_location = Path("~/Documents/HopePharmaData").expanduser()
        default_location.mkdir(parents=True, exist_ok=True)
        self.status.config(text=f"📁 Created default location: {default_location}")
        self.start_app(str(default_location), "cloud")
    
    def setup_network(self):
        """Enhanced network setup with platform-specific instructions"""
        if self.system == "Windows":
            instructions = (
                "NETWORK SETUP - WINDOWS\n\n"
                "ON MAIN COMPUTER:\n"
                "1. Create folder: C:\\HopePharmaShared\n"
                "2. Right-click → Properties → Sharing → Share\n"
                "3. Add 'Everyone' with Read/Write permissions\n\n"
                "ON OTHER COMPUTERS:\n"
                "1. Open File Explorer → Network\n"
                "2. Find the main computer → Select shared folder"
            )
            default_location = "C:\\HopePharmaShared"
        else:  # macOS and Linux
            instructions = (
                "NETWORK SETUP - macOS/LINUX\n\n"
                "ON MAIN COMPUTER:\n"
                "1. Create folder: /Users/Shared/HopePharmaData\n"
                "2. System Settings → Sharing → File Sharing ON\n"
                "3. Add folder to Shared Folders with Read & Write\n\n"
                "ON OTHER COMPUTERS:\n"
                "1. Connect to the main computer via network\n"
                "2. Select the shared HopePharmaData folder"
            )
            default_location = "/Users/Shared/HopePharmaData"
        
        messagebox.showinfo("Network Setup", instructions)
        self.status.config(text="🔄 Selecting network shared folder...")
        
        folder = filedialog.askdirectory(title="Select the shared HopePharma folder")
        if folder:
            self.start_app(folder, "network")
        else:
            self._setup_network_default(default_location)
    
    def _setup_network_default(self, default_location):
        """Setup network with default location"""
        try:
            Path(default_location).mkdir(parents=True, exist_ok=True)
            self.status.config(text=f"📁 Created network location: {default_location}")
            self.root.update()
            time.sleep(1)
            self.start_app(default_location, "network")
        except Exception as e:
            messagebox.showerror("Error", f"Cannot create network folder:\n{str(e)}")
            self.status.config(text="❌ Network setup failed")
    
    def setup_single(self):
        """Enhanced single computer setup"""
        folder = Path("~/Documents/HopePharmaData").expanduser()
        folder.mkdir(parents=True, exist_ok=True)
        
        self.status.config(text="💻 Setting up single computer mode...")
        self.root.update()
        time.sleep(1)
        
        self.start_app(str(folder), "single")
    
    def start_app(self, data_folder, mode):
        """Enhanced application startup with validation"""
        self.status.config(text=f"🚀 Starting in {mode} mode...")
        self.root.update()
        
        try:
            # Enhanced folder validation
            data_path = Path(data_folder)
            
            if not data_path.exists():
                data_path.mkdir(parents=True, exist_ok=True)
            
            # Test folder permissions
            test_file = data_path / "permission_test.txt"
            try:
                with open(test_file, 'w') as f:
                    f.write("HopePharma permission test")
                test_file.unlink()  # Clean up
            except PermissionError:
                messagebox.showerror("Permission Error", 
                                   f"Cannot write to folder:\n{data_folder}\n\n"
                                   "Please choose a folder with write permissions.")
                self.status.config(text="❌ Permission error")
                return
            if not self._ensure_license(data_folder):
                self.status.config(text="❌ License invalid")
                return
            
            # Close setup window
            self.root.destroy()
            
            # Start main application
            self.launch_enhanced_main_app(data_folder, mode)
            
        except Exception as e:
            messagebox.showerror("Startup Error", 
                               f"Cannot start application:\n{str(e)}")
            self.status.config(text="❌ Startup failed")
    
    def launch_enhanced_main_app(self, data_folder, mode):
        """Launch enhanced main application"""
        try:
            # Import and create the main app
            import tkinter as tk
            root = tk.Tk()
            try:
                if platform.system() == "Windows":
                    ico = Path("logo.ico")
                    if ico.exists():
                        root.iconbitmap(str(ico))
                elif platform.system() == "Darwin":
                    try:
                        p = None
                        import sys
                        if hasattr(sys, "_MEIPASS"):
                            p = Path(sys._MEIPASS) / "logo.png"
                        else:
                            app_res = Path(sys.executable).resolve()
                            p = app_res.parent.parent / 'Resources' / 'logo.png'
                        if p.exists():
                            icon_img = tk.PhotoImage(file=str(p))
                            root.iconphoto(True, icon_img)
                    except Exception:
                        pass
                try:
                    png_here = Path.cwd() / "logo.png"
                    if png_here.exists():
                        icon_img = tk.PhotoImage(file=str(png_here))
                        root.iconphoto(True, icon_img)
                except Exception:
                    pass
            except Exception:
                pass
            # Apply app icon if available
            try:
                desktop_logo = Path(os.path.expanduser("~")) / "Desktop" / "logo.png"
                if desktop_logo.exists():
                    icon_img = tk.PhotoImage(file=str(desktop_logo))
                    root.iconphoto(True, icon_img)
                else:
                    # Use data folder logo if present
                    df_logo = Path(data_folder) / "logo.png"
                    if df_logo.exists():
                        icon_img = tk.PhotoImage(file=str(df_logo))
                        root.iconphoto(True, icon_img)
            except Exception:
                pass
            
            # Create our enhanced cloud data manager
            manager = EnhancedCloudDataManager(data_folder, mode)
            
            # Try to import and start the main app
            try:
                import importlib
                import sys
                candidates = [
                    Path.cwd() / 'invoice_app.py',
                    Path.cwd() / 'invoice_app.pyc',
                    Path.cwd() / 'invoice_app.pye',
                    Path(data_folder) / 'invoice_app.py',
                    Path(data_folder) / 'invoice_app.pyc',
                    Path(data_folder) / 'invoice_app.pye'
                ]
                for p in candidates:
                    if p.exists():
                        sys.path.insert(0, str(p.parent))
                        break
                module = None
                InvoiceApp = None
                for mod_name in ['invoice_app', 'HopePharma', 'main_app', 'hopepharma_ui']:
                    try:
                        module = importlib.import_module(mod_name)
                        for cls in ['InvoiceApp', 'App', 'HopePharmaApp']:
                            if hasattr(module, cls):
                                InvoiceApp = getattr(module, cls)
                                break
                        if InvoiceApp:
                            break
                    except Exception:
                        pass
                if not InvoiceApp:
                    raise ImportError('Main app module not found')
                
                # Pass manager to InvoiceApp to ensure correct paths are used from the start
                app = InvoiceApp(root, manager=manager)
                # Replace the manager with our enhanced manager (just in case)
                app.manager = manager
                
            except ImportError:
                try:
                    root.title("HopePharma")
                    root.minsize(900, 640)
                    frm = ttk.Frame(root, padding="16")
                    frm.pack(fill='both', expand=True)
                    ttk.Label(frm, text="HopePharma", font=('Helvetica', 20, 'bold')).pack(pady=8)
                    ttk.Label(frm, text=f"Mode: {mode} | Data: {data_folder}").pack(pady=4)
                    ttk.Separator(frm).pack(fill='x', pady=10)
                    ttk.Label(frm, text="Fallback UI active. Main packaged app not found.\nCore features are available via the manager.", font=('Helvetica', 10)).pack(pady=8)
                    btns = ttk.Frame(frm)
                    btns.pack(pady=12)
                    def _open_data():
                        try:
                            p = Path(data_folder)
                            if platform.system() == "Darwin":
                                subprocess.run(["open", str(p)])
                            elif platform.system() == "Windows":
                                os.startfile(str(p))
                            else:
                                subprocess.run(["xdg-open", str(p)])
                        except Exception:
                            pass
                    def _test_pdf():
                        try:
                            gen = manager.pdf_generator
                            sample = {
                                'client_name': 'Test Client',
                                'client_trn': '',
                                'invoice_id': manager.generate_invoice_id(),
                                'date': datetime.now().strftime('%Y-%m-%d'),
                                'payment_method': 'Cash',
                                'payment_due': 'On receipt',
                                'items': [{'description': 'Sample Item','quantity':1,'unit_price':100,'total':100,'taxable':False}],
                                'subtotal': 100,
                                'taxable_amount': 0,
                                'non_taxable_amount': 100,
                                'tax_amount': 0,
                                'grand_total': 100
                            }
                            path = gen.generate_invoice_pdf(sample)
                            messagebox.showinfo("PDF Generated", f"Saved: {path}")
                        except Exception as e:
                            messagebox.showerror("Error", str(e))
                    ttk.Button(btns, text="Open Data Folder", command=_open_data).pack(side='left', padx=6)
                    ttk.Button(btns, text="Generate Test Invoice PDF", command=_test_pdf).pack(side='left', padx=6)
                    try:
                        root.update_idletasks()
                        w = max(900, root.winfo_reqwidth()+24)
                        h = max(640, root.winfo_reqheight()+24)
                        root.minsize(w, h)
                        root.geometry(f"{w}x{h}")
                    except Exception:
                        pass
                except Exception:
                    pass
            
            root.mainloop()
            
        except Exception as e:
            messagebox.showerror("Error", 
                               f"Cannot launch main application:\n{str(e)}\n\n"
                               "Please ensure all dependencies are installed.")
            import traceback
            traceback.print_exc()
    def _machine_id(self):
        try:
            return f"{platform.node()}-{uuid.getnode()}"
        except Exception:
            return platform.node()
    def _expected_key(self, customer, expires, mid):
        import hashlib
        s = f"{customer}|{expires}|{mid}|HPM_LIC_SALT_v1"
        return hashlib.sha256(s.encode()).hexdigest()
    def _verify_license(self, folder):
        try:
            p = Path(folder) / "license.json"
            if not p.exists():
                return False
            with open(p, 'r') as f:
                lic = json.load(f)
            customer = str(lic.get('customer','')).strip()
            expires = str(lic.get('expires','')).strip()
            key = str(lic.get('key','')).strip()
            if not customer or not expires or not key:
                return False
            mid = self._machine_id()
            exp_key = self._expected_key(customer, expires, mid)
            try:
                d = datetime.strptime(expires, "%Y-%m-%d")
                if d < datetime.now():
                    return False
            except Exception:
                return False
            return key == exp_key
        except Exception:
            return False
    def _activation_dialog(self, folder):
        top = tk.Toplevel(self.root)
        top.title("Activate HopePharma")
        top.geometry("560x360")
        try:
            top.resizable(True, True)
        except Exception:
            pass
        top.grab_set()
        frm = ttk.Frame(top, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Customer Name").pack(anchor='w')
        name_var = tk.StringVar()
        ttk.Entry(frm, textvariable=name_var).pack(fill='x', pady=6)
        ttk.Label(frm, text="License Key").pack(anchor='w')
        key_var = tk.StringVar()
        ttk.Entry(frm, textvariable=key_var).pack(fill='x', pady=6)
        ttk.Label(frm, text="Expiry (YYYY-MM-DD)").pack(anchor='w')
        exp_var = tk.StringVar()
        ttk.Entry(frm, textvariable=exp_var).pack(fill='x', pady=6)
        ttk.Label(frm, text=f"Machine ID: {self._machine_id()}").pack(anchor='w', pady=6)
        result = {'ok': False}
        def ok():
            customer = name_var.get().strip()
            key = key_var.get().strip()
            expires = exp_var.get().strip()
            if not customer or not key or not expires:
                return
            mid = self._machine_id()
            exp_key = self._expected_key(customer, expires, mid)
            if key != exp_key:
                messagebox.showerror("Invalid License", "License key is invalid")
                return
            p = Path(folder) / "license.json"
            with open(p, 'w') as f:
                json.dump({'customer': customer,'expires': expires,'key': key}, f)
            result['ok'] = True
            top.destroy()
        def cancel():
            top.destroy()
        btns = ttk.Frame(frm)
        btns.pack(fill='x', pady=10)
        ttk.Button(btns, text="Activate", command=ok).pack(side='left', padx=4)
        ttk.Button(btns, text="Cancel", command=cancel).pack(side='left', padx=4)
        ttk.Button(btns, text="Issue License (Admin)", command=lambda: self._issue_license_admin(folder)).pack(side='right', padx=4)
        self.root.wait_window(top)
        return result['ok']
    def _ensure_license(self, folder):
        return True

    def _issue_license_admin(self, folder):
        login = tk.Toplevel(self.root)
        login.title("Admin Login")
        login.geometry("480x260")
        try:
            login.resizable(True, True)
        except Exception:
            pass
        login.grab_set()
        lfrm = ttk.Frame(login, padding="12")
        lfrm.pack(fill='both', expand=True)
        ttk.Label(lfrm, text="Admin ID").pack(anchor='w')
        aid = tk.StringVar()
        ttk.Entry(lfrm, textvariable=aid).pack(fill='x', pady=6)
        ttk.Label(lfrm, text="Password").pack(anchor='w')
        apw = tk.StringVar()
        ttk.Entry(lfrm, textvariable=apw, show='*').pack(fill='x', pady=6)
        ok = {'auth': False}
        def do_login():
            if aid.get().strip() == 'ataa777' and apw.get().strip() == '5410954Ab':
                ok['auth'] = True
                login.destroy()
            else:
                messagebox.showerror("Error", "Invalid admin credentials")
        ttk.Button(lfrm, text="Login", command=do_login).pack(pady=6)
        self.root.wait_window(login)
        if not ok['auth']:
            return
        dlg = tk.Toplevel(self.root)
        dlg.title("Issue License")
        dlg.geometry("640x420")
        try:
            dlg.resizable(True, True)
        except Exception:
            pass
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding="12")
        frm.pack(fill='both', expand=True)
        ttk.Label(frm, text="Customer Name").pack(anchor='w')
        name = tk.StringVar()
        ttk.Entry(frm, textvariable=name).pack(fill='x', pady=6)
        ttk.Label(frm, text="Duration (days)").pack(anchor='w')
        days = tk.StringVar(value='365')
        ttk.Entry(frm, textvariable=days).pack(fill='x', pady=6)
        mid = self._machine_id()
        ttk.Label(frm, text=f"Machine ID: {mid}").pack(anchor='w', pady=6)
        key_var = tk.StringVar()
        ttk.Label(frm, text="Generated Key").pack(anchor='w')
        key_entry = ttk.Entry(frm, textvariable=key_var)
        key_entry.pack(fill='x', pady=6)
        def generate():
            try:
                d = int(days.get().strip())
            except Exception:
                messagebox.showerror("Error", "Invalid duration")
                return
            expires = (datetime.now() + timedelta(days=d)).strftime("%Y-%m-%d")
            customer = name.get().strip()
            if not customer:
                return
            key_var.set(self._expected_key(customer, expires, mid))
        def save():
            customer = name.get().strip()
            k = key_var.get().strip()
            if not customer or not k:
                return
            try:
                d = int(days.get().strip())
            except Exception:
                return
            expires = (datetime.now() + timedelta(days=d)).strftime("%Y-%m-%d")
            p = Path(folder) / 'license.json'
            with open(p, 'w') as f:
                json.dump({'customer': customer, 'expires': expires, 'key': k}, f)
            messagebox.showinfo("Saved", f"License saved to\n{p}")
        bfrm = ttk.Frame(frm)
        bfrm.pack(fill='x', pady=8)
        ttk.Button(bfrm, text="Generate", command=generate).pack(side='left', padx=4)
        ttk.Button(bfrm, text="Save", command=save).pack(side='left', padx=4)
        ttk.Button(bfrm, text="Close", command=dlg.destroy).pack(side='right', padx=4)

def main():
    """Enhanced main function with better error handling"""
    try:
        if getattr(sys, "frozen", False):
            base = Path.home() / "Documents" / "HopePharmaData"
            try:
                base.mkdir(parents=True, exist_ok=True)
            except Exception:
                base = Path.home() / "Desktop"
            log_path = base / "startup.log"
            try:
                import traceback
                log_f = open(log_path, "a", buffering=1, encoding="utf-8")
                sys.stdout = log_f
                sys.stderr = log_f
                def _hook(exc_type, exc, tb):
                    try:
                        traceback.print_exception(exc_type, exc, tb)
                    except Exception:
                        pass
                sys.excepthook = _hook
                try:
                    print("---- startup ----", datetime.now().isoformat())
                except Exception:
                    pass
            except Exception:
                pass
    except Exception:
        pass
    try:
        # Create main window
        root = tk.Tk()
        try:
            if platform.system() == "Windows":
                ico = Path("logo.ico")
                if ico.exists():
                    root.iconbitmap(str(ico))
            elif platform.system() == "Darwin":
                try:
                    p = None
                    if hasattr(sys, "_MEIPASS"):
                        p = Path(sys._MEIPASS) / "logo.png"
                    else:
                        app_res = Path(sys.executable).resolve()
                        p = app_res.parent.parent / 'Resources' / 'logo.png'
                    if p.exists():
                        icon_img = tk.PhotoImage(file=str(p))
                        root.iconphoto(True, icon_img)
                except Exception:
                    pass
            # Fallback PNG in current folder
            try:
                png_here = Path.cwd() / "logo.png"
                if png_here.exists():
                    icon_img = tk.PhotoImage(file=str(png_here))
                    root.iconphoto(True, icon_img)
            except Exception:
                pass
        except Exception:
            pass
        # Set window icon and title
        root.title("HopePharma CRM")
        try:
            desktop_logo = Path(os.path.expanduser("~")) / "Desktop" / "logo.png"
            if desktop_logo.exists():
                icon_img = tk.PhotoImage(file=str(desktop_logo))
                root.iconphoto(True, icon_img)
            elif platform.system() == "Darwin":
                try:
                    p = None
                    if hasattr(sys, "_MEIPASS"):
                        p = Path(sys._MEIPASS) / "logo.png"
                    else:
                        app_res = Path(sys.executable).resolve()
                        p = app_res.parent.parent / 'Resources' / 'logo.png'
                    if p.exists():
                        icon_img = tk.PhotoImage(file=str(p))
                        root.iconphoto(True, icon_img)
                except Exception:
                    pass
            else:
                # Current folder logo fallback
                png_here = Path.cwd() / "logo.png"
                if png_here.exists():
                    icon_img = tk.PhotoImage(file=str(png_here))
                    root.iconphoto(True, icon_img)
        except Exception:
            pass
        
        # Start the enhanced application
        app = HopePharmaComplete(root)
        try:
            root.after(2000, lambda: _check_for_updates(root))
        except Exception:
            pass
        
        # Start the main loop
        root.mainloop()
        
    except Exception as e:
        print(f"Fatal error: {e}")
        import traceback
        traceback.print_exc()
        
        # Show error message
        error_root = tk.Tk()
        error_root.withdraw()  # Hide the main window
        messagebox.showerror("Fatal Error", 
                           f"Application failed to start:\n{str(e)}\n\n"
                           "Please check the console for details.")
        error_root.destroy()

if __name__ == "__main__":
    main()
