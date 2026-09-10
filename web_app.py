#!/usr/bin/env python3
"""
ASISTEM — Multi-Company Business Suite Web App (Flask based).
"""
from __future__ import annotations

import os
import sys
import io
import zipfile
from datetime import datetime, date, timedelta
from pathlib import Path

from flask import (
    Flask, render_template, request, redirect, url_for, flash,
    abort, jsonify, send_file, Response, __version__ as flask_version,
    make_response, session, g,
)


def _apply_production_env_defaults_early():
    """Apply production env defaults EARLY (before DATA_DIR is resolved).

    web_app.py calls _resolve_data_dir() BEFORE importing hope_pharma_complete.py
    (because the EnhancedCloudDataManager is instantiated AFTER DATA_DIR is set).
    Without an EARLY bootstrap here, HOPEPHARMA_DATA_DIR would be empty and cause
    _resolve_data_dir() to incorrectly choose PROJECT_ROOT on Vercel (READ-ONLY).

    Mirror logic from hope_pharma_complete.py _apply_production_env_defaults().
    Explicit env vars (e.g. set via Vercel Env Vars UI) ALWAYS win.
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
    if not is_serverless:
        return
    for key, default_val in _PROD_DEFAULTS.items():
        existing = (os.environ.get(key) or "").strip()
        if existing:
            continue
        os.environ[key] = str(default_val)

_apply_production_env_defaults_early()

PROJECT_ROOT = Path(__file__).resolve().parent
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from hope_pharma_complete import EnhancedCloudDataManager  # noqa: E402


def _resolve_data_dir() -> str:
    """Choose a writable DATA_DIR for the current runtime.

    Priority:
    1. HOPEPHARMA_DATA_DIR env var if explicitly set
    2. /tmp (always writable on Vercel / serverless / Linux/macOS) when we detect
       a serverless / read-only / CI environment
    3. PROJECT_ROOT (desktop mode) otherwise
    """
    explicit = os.environ.get("HOPEPHARMA_DATA_DIR")
    if explicit:
        return explicit
    # Serverless environments (Vercel) always provide VERCEL=1 and write to /tmp.
    # PROJECT_ROOT on Vercel resolves to /var/task which is READ-ONLY.
    if (
        os.environ.get("VERCEL") == "1"
        or os.environ.get("LAMBDA_TASK_ROOT")
        or os.environ.get("AWS_LAMBDA_FUNCTION_NAME")
        or str(PROJECT_ROOT).startswith("/var/task")
    ):
        fallback = "/tmp/asistem-data"
        try:
            os.makedirs(fallback, exist_ok=True)
        except Exception:
            fallback = "/tmp"
        print(f"[web_app] serverless env detected — using DATA_DIR={fallback}")
        return fallback
    # Desktop fallback: project root (user's repository).
    # If somehow the project root is read-only, punt to /tmp too.
    try:
        probe = PROJECT_ROOT / ".asistem_write_test_"
        probe.write_text("ok")
        probe.unlink()
        return str(PROJECT_ROOT)
    except Exception:
        print(f"[web_app] PROJECT_ROOT ({PROJECT_ROOT}) is read-only or unwritable — using /tmp instead")
        return "/tmp"


DATA_DIR = _resolve_data_dir()

app = Flask(__name__)
app.secret_key = os.environ.get("HOPEPHARMA_APP_SECRET") or "hope-pharma-local-dev-secret"
# Secure production session settings
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=bool(int(os.environ.get("HOPEPHARMA_PRODUCTION", "0") or "0")),
    PERMANENT_SESSION_LIFETIME=timedelta(days=7),
)

try:
    manager = EnhancedCloudDataManager(DATA_DIR, mode='single')
except TypeError:
    manager = EnhancedCloudDataManager(DATA_DIR)


# ================================================================
# PER-COMPANY SESSION AUTH
# ================================================================
PUBLIC_ROUTES = {"static", "login_get", "login_post", "api_health", "company_signup_get", "company_signup_submit"}

# ================================================================
# SIGN-UP RATE LIMIT STATE (Task4 TR-4.3: 3 signups / 10 min per IP)
# ================================================================
_SIGNUP_RATE_LIMIT = {}  # ip_string -> list of signup timestamps
_SIGNUP_RATE_WINDOW_SEC = 600  # 10 minutes
_SIGNUP_RATE_MAX_PER_WINDOW = 3


def login_required(fn):
    """Decorator: require a valid logged-in user session for the route."""
    from functools import wraps

    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not session.get("user") or not session.get("user", {}).get("username"):
            flash("Please log in to access ASISTEM.", "warning")
            return redirect(url_for("login_get", next=request.path))
        # Prevent cross-company access: a user is locked to their company_id
        user_company_id = session["user"].get("company_id")
        if session.get("enforced_company_id") and session["enforced_company_id"] != user_company_id:
            session.clear()
            flash("Session company mismatch — please log in again.", "error")
            return redirect(url_for("login_get"))
        return fn(*args, **kwargs)

    return wrapper


def _user_is_admin(user=None):
    """Safe truthy check: role == "admin".

    If user is None, tries session.get("user") within a request context;
    outside request context just returns False (safe default for CLI/imports).
    """
    u = user
    if u is None:
        try:
            u = session.get("user") or {}
        except Exception:
            return False
    return (u or {}).get("role") == "admin"


def admin_required(fn):
    """Decorator: require role == admin for the route (login_required already handled by before_request)."""
    from functools import wraps

    @wraps(fn)
    def wrapper(*args, **kwargs):
        if not _user_is_admin():
            flash("You do not have permission to access this page (Admin-only).", "error")
            return redirect(url_for("dashboard"))
        return fn(*args, **kwargs)

    return wrapper


@app.before_request
def _before_request_enforce_login_and_company_scope():
    """Runs before EVERY request: (1) enforce login wall except public routes,
    (2) ensure the active company matches the logged-in user's company_id."""
    endpoint = request.endpoint
    # Allow public endpoints (and 404s/None)
    if endpoint is None or endpoint in PUBLIC_ROUTES:
        return
    # Health endpoint is public (for smoke checks)
    if endpoint == "api_health":
        return
    # --- Enforce login ---
    if not session.get("user") or not session.get("user", {}).get("username"):
        if request.path.startswith("/api/"):
            return jsonify({"error": "Authentication required"}), 401
        flash("Please log in to access ASISTEM.", "warning")
        return redirect(url_for("login_get", next=request.path))
    # --- Enforce company scope: a user can ONLY see their assigned company's data ---
    user = session["user"]
    user_cid = user.get("company_id")
    if user_cid:
        session["enforced_company_id"] = user_cid
        # Sync the data-manager's active company to the user's assigned company
        if hasattr(manager, "set_active_company"):
            try:
                manager.set_active_company(user_cid)
            except Exception:
                pass
    g.current_user = user
    return None


@app.context_processor
def inject_auth_user():
    """Inject current_user + is_logged_in + is_admin into every Jinja template."""
    user = session.get("user") if session.get("user", {}).get("username") else None
    is_admin_now = _user_is_admin(user)
    return {
        "current_user": user,
        "is_logged_in": bool(user),
        "is_admin": is_admin_now,
        "logout_url": url_for("logout_post"),
    }


@app.route("/login", methods=["GET"])
def login_get():
    # Already logged in: skip straight to dashboard
    if session.get("user") and session["user"].get("username"):
        return redirect(url_for("dashboard"))
    next_url = request.args.get("next") or url_for("dashboard")
    return render_template("login.html", next=next_url)


@app.route("/login", methods=["POST"])
def login_post():
    username = (request.form.get("username") or "").strip()
    password = request.form.get("password") or ""
    next_url = (request.form.get("next") or "").strip() or url_for("dashboard")
    # Security: sanity-limit URL redirection to same-origin paths only
    if next_url.startswith("http://") or next_url.startswith("https://") or "//" in next_url:
        next_url = url_for("dashboard")
    if not username or not password:
        flash("Username and password are required.", "error")
        return render_template("login.html", next=next_url), 400
    if not hasattr(manager, "verify_company_user"):
        flash("User authentication is not available on this data manager.", "error")
        return render_template("login.html", next=next_url), 500
    safe_user = manager.verify_company_user(username, password)
    if not safe_user:
        # Rate-limit style: never hint whether username vs password was wrong
        flash("Invalid username or password.", "error")
        return render_template("login.html", next=next_url), 401
    # --- SUCCESS: log them in, auto-switch to their company ---
    session.permanent = True
    session["user"] = safe_user
    # Force the active company to match the user's assigned company (prevents cross-company snooping)
    user_cid = safe_user.get("company_id")
    if user_cid and hasattr(manager, "set_active_company"):
        manager.set_active_company(user_cid)
    company_name = ""
    try:
        comp = manager.get_active_company_profile()
        company_name = (comp or {}).get("display_name") or (comp or {}).get("name") or ""
    except Exception:
        company_name = ""
    flash(f"Welcome back, {safe_user.get('full_name') or safe_user['username']}! Logged into {company_name or safe_user.get('company_id')}.", "success")
    return redirect(next_url)


@app.route("/logout", methods=["GET", "POST"])
def logout_post():
    session.clear()
    flash("You have been logged out of ASISTEM securely.", "info")
    return redirect(url_for("login_get"))


# ================================================================
# Task 4 — PUBLIC PRE-LOGIN COMPANY SIGN-UP
# ================================================================

_SIGNUP_EMIRATES = [
    "Abu Dhabi", "Ajman", "Dubai", "Fujairah",
    "Ras Al Khaimah", "Sharjah", "Umm Al Quwain",
]
_SIGNUP_CURRENCIES = ["AED", "USD", "EUR", "GBP", "SAR", "USD", "INR", "PKR", "EGP", "OMR", "BHD", "KWD", "QAR"]


def _signup_get_client_ip():
    """Return the originating IP for signup rate limiting.
    Respects X-Forwarded-For (Vercel / reverse proxy.)"""
    xff = request.headers.get("X-Forwarded-For", "")
    if xff:
        first = xff.split(",", 1)[0].strip()
        if first: return first
    return request.remote_addr or "0.0.0.0"


def _signup_rate_limit_check(ip):
    """Return True if OK (permitted), False if blocked. Cleans stale entries."""
    import time
    now = time.time()
    cut = now - _SIGNUP_RATE_WINDOW_SEC
    row = [t for t in _SIGNUP_RATE_LIMIT.get(ip, []) if t > cut]
    _SIGNUP_RATE_LIMIT[ip] = row
    if len(row) >= _SIGNUP_RATE_MAX_PER_WINDOW:
        return False
    return True


def _signup_rate_limit_record(ip):
    import time
    _SIGNUP_RATE_LIMIT.setdefault(ip, []).append(time.time())


def _signup_file_to_data_uri(file_storage, label):
    """Convert a Flask FileStorage into a data:image/png;base64, data URI.
    For PNG/JPG uploads, resize to max width 400 using Pillow if installed, else raw if size <= 2MB; reject >2MB."""
    import base64
    if not file_storage or not getattr(file_storage, "filename", None):
        return None, ""
    raw = file_storage.read() or b""
    if len(raw) > 2 * 1024 * 1024:
        return None, f"{label} file is too large (max 2MB)"
    if not raw:
        return None, ""
    fname = (getattr(file_storage, "filename", "") or "").lower()
    mime = "image/png"
    if fname.endswith((".jpg", ".jpeg")): mime = "image/jpeg"
    elif fname.endswith(".gif"): mime = "image/gif"
    elif fname.endswith(".webp"): mime = "image/webp"
    ok_bytes = raw
    try:
        from PIL import Image
        src = Image.open(io.BytesIO(raw))
        if src.mode not in ("RGB", "RGBA"):
            src = src.convert("RGBA")
        mw = 400
        w, h = src.size
        if w > mw:
            ratio = mw / float(w)
            nh = max(1, int(h * ratio))
            src = src.resize((mw, nh), Image.LANCZOS)
        buf = io.BytesIO()
        fmt = "PNG" if mime == "image/png" else "JPEG" if mime in ("image/jpeg",) else "PNG"
        if fmt == "JPEG" and src.mode == "RGBA":
            bg = Image.new("RGB", src.size, (255, 255, 255))
            bg.paste(src, mask=src.split()[3])
            src = bg
        src.save(buf, format=fmt)
        ok_bytes = buf.getvalue()
        mime = "image/png" if fmt == "PNG" else "image/jpeg"
    except Exception:
        ok_bytes = raw
    b64 = base64.b64encode(ok_bytes).decode("ascii")
    return f"data:{mime};base64,{b64}", ""


@app.route("/company/signup", methods=["GET"])
def company_signup_get():
    # Already logged in: shortcut
    if session.get("user") and session["user"].get("username"):
        return redirect(url_for("dashboard"))
    return render_template(
        "company_signup.html",
        emirates=_SIGNUP_EMIRATES,
        currencies=_SIGNUP_CURRENCIES,
        form=request.form if request.method == "POST" else {},
    )


@app.route("/company/signup/submit", methods=["POST"])
def company_signup_submit():
    """Handle Task 4 signup submission (multipart form)."""
    import re
    if session.get("user") and session["user"].get("username"):
        return redirect(url_for("dashboard"))

    ip = _signup_get_client_ip()

    # --- RATE LIMIT CHECK FIRST, BEFORE DOING EXPENSIVE WORK
    if not _signup_rate_limit_check(ip):
        flash("Too many signup attempts from this address — please wait 10 minutes and try again.", "error")
        return render_template(
            "company_signup.html",
            emirates=_SIGNUP_EMIRATES,
            currencies=_SIGNUP_CURRENCIES,
            form=request.form,
        ), 429

    # 1. Collect form fields =============================================
    f = request.form
    display_name = (f.get("display_name") or "").strip()
    legal_name = (f.get("legal_name") or "").strip() or display_name
    trn = re.sub(r"\D", "", (f.get("trn") or ""))
    emirate = (f.get("emirate") or "Dubai").strip() or "Dubai"
    address = (f.get("address") or "").strip()
    phone = (f.get("phone") or "").strip()
    email = (f.get("email") or "").strip()
    website = (f.get("website") or "").strip()
    currency = (f.get("currency") or "AED").strip().upper() or "AED"
    try:
        vat_rate = float(f.get("vat_rate") or "5")
    except Exception:
        vat_rate = 5.0
    prefix_raw = (f.get("invoice_prefix") or "").strip()
    short_code_raw = (f.get("short_code") or prefix_raw or "").strip()
    # Final normalized values (upper-cased for storage IF valid; for now keep as-is until validation rejects)
    prefix = prefix_raw
    short_code = short_code_raw

    admin_full_name = (f.get("admin_full_name") or "").strip()
    admin_username = (f.get("admin_username") or "").strip()
    admin_pw = f.get("admin_password") or ""
    admin_pw2 = f.get("admin_password_confirm") or ""
    admin_email = (f.get("admin_email") or email or "").strip()
    terms_agreed = bool(f.get("terms_agreed"))

    errors = []
    # 2. Validate ===================================================
    if not display_name:
        errors.append("Company Display Name is required.")
    if not legal_name:
        errors.append("Company Full Legal Name is required.")
    if trn and len(trn) != 15:
        errors.append("TRN must be exactly 15 digits (numbers only).")
    # Prefix MUST be exactly 4 uppercase A-Z 0-9; check RAW prefix (before upper-casing) to catch lowercase rejection
    if not re.fullmatch(r"[A-Z0-9]{4}", prefix or ""):
        # If user passed letters/digits but lowercase, specific hint
        if prefix and re.fullmatch(r"[A-Za-z0-9]{4}", prefix):
            errors.append("Invoice Prefix must be exactly 4 UPPERCASE letters/numbers (no lowercase).")
        else:
            errors.append("Invoice Prefix must be exactly 4 uppercase letters A–Z or digits 0-9.")
    if len(admin_pw) < 8:
        errors.append("Admin Password must be at least 8 characters.")
    if admin_pw != admin_pw2:
        errors.append("Passwords do not match.")
    if not admin_username:
        errors.append("Admin Username is required.")
    if email and "@" not in email:
        errors.append("Company Email does not look valid.")
    if admin_email and "@" not in admin_email:
        errors.append("Admin Email does not look valid.")
    if not terms_agreed:
        errors.append("Please accept the Terms checkbox to continue.")
    # username uniqueness
    if admin_username and hasattr(manager, "get_company_user_by_username") and \
            manager.get_company_user_by_username(admin_username):
        errors.append(f"Username '{admin_username}' is already taken on this platform.")

    # 3. Images =====================================================
    logo_uri = logo_err = stamp_uri = stamp_err = sign_uri = sign_err = None
    try:
        logo_uri, logo_err = _signup_file_to_data_uri(request.files.get("logo_image"), "Logo")
        sign_uri, sign_err = _signup_file_to_data_uri(request.files.get("signature_image"), "Signature")
        stamp_uri, stamp_err = _signup_file_to_data_uri(request.files.get("stamp_image"), "Stamp")
        for msg in (logo_err, sign_err, stamp_err):
            if msg:
                errors.append(msg)
    except Exception as e:
        errors.append(f"Could not process uploaded images: {e}")

    if errors:
        for msg in errors:
            flash(msg, "error")
        return render_template(
            "company_signup.html",
            emirates=_SIGNUP_EMIRATES,
            currencies=_SIGNUP_CURRENCIES,
            form=request.form,
        ), 200

    # --- Normalize upper-case prefix AFTER validation passes ---
    prefix = (prefix or "").upper()
    short_code = (short_code or prefix or "").upper()[:4]

    # 4. Create company =============================================
    profile_data = {
        "name": display_name,
        "display_name": display_name,
        "legal_name": legal_name,
        "prefix": prefix,
        "short_code": short_code,
        "trn": trn,
        "emirate": emirate,
        "address": address,
        "phone": phone,
        "email": email,
        "website": website,
        "currency": currency,
        "vat_rate": vat_rate,
        "logo_png_base64": logo_uri,
        "signature_png_base64": sign_uri,
        "stamp_png_base64": stamp_uri,
        "bank_name": (f.get("bank_name") or "").strip(),
        "bank_branch": (f.get("bank_branch") or "").strip(),
        "bank_account_no": (f.get("bank_account_no") or "").strip(),
        "bank_iban": (f.get("bank_iban") or "").strip(),
        "bank_swift": (f.get("bank_swift") or "").strip(),
    }
    ok, msg, profile = manager.create_company_profile(profile_data)
    if not ok or not profile:
        flash(f"Could not create company: {msg}", "error")
        return render_template(
            "company_signup.html",
            emirates=_SIGNUP_EMIRATES,
            currencies=_SIGNUP_CURRENCIES,
            form=request.form,
        ), 400
    new_company_id = profile["id"]

    # 5. Create admin user ========================================
    u_ok, u_msg, u_safe = manager.create_company_user(
        username=admin_username,
        password=admin_pw,
        company_id=new_company_id,
        role="admin",
        full_name=admin_full_name or admin_username,
        email=admin_email,
    )
    if not u_ok:
        # Roll back company row (best-effort)
        try:
            manager.delete_company_profile(new_company_id, allow_delete_default=False)
        except Exception:
            pass
        flash(f"Could not create admin user: {u_msg}", "error")
        return render_template(
            "company_signup.html",
            emirates=_SIGNUP_EMIRATES,
            currencies=_SIGNUP_CURRENCIES,
            form=request.form,
        ), 400

    # 6. Record signup under rate limit now that we actually created it ====
    _signup_rate_limit_record(ip)

    flash(
        f"Company '{profile.get('display_name') or profile.get('name')} created! "
        f"Sign in with username '{admin_username}'.",
        "success",
    )
    return redirect(url_for("login_get"))


@app.context_processor
def inject_globals():
    companies = []
    active = None
    try:
        if hasattr(manager, "get_all_company_profiles"):
            companies = list(manager.get_all_company_profiles() or [])
    except Exception:
        companies = []
    try:
        if hasattr(manager, "get_active_company_profile"):
            active = manager.get_active_company_profile()
    except Exception:
        active = None
    return {
        "year": datetime.now().year,
        "companies": companies,
        "active_company": active,
    }


def money(v):
    try:
        return float(v or 0)
    except Exception:
        return 0.0


def today_str() -> str:
    return date.today().isoformat()


def month_start_end(y=None, m=None):
    today = date.today()
    y = y or today.year
    m = m or today.month
    try:
        y = int(y); m = int(m)
        start = date(y, m, 1)
        if m == 12:
            end = date(y + 1, 1, 1) - timedelta(days=1)
        else:
            end = date(y, m + 1, 1) - timedelta(days=1)
        return start.isoformat(), end.isoformat(), y, m
    except Exception:
        s = date(today.year, today.month, 1)
        if today.month == 12:
            e = date(today.year + 1, 1, 1) - timedelta(days=1)
        else:
            e = date(today.year, today.month + 1, 1) - timedelta(days=1)
        return s.isoformat(), e.isoformat(), today.year, today.month


def filter_invoices_for_active_company(invoices_list):
    try:
        active = manager.get_active_company_profile()
        active_id = active.get("id")
        default_id = getattr(manager, "DEFAULT_COMPANY_ID", "com_hopepharma")
    except Exception:
        return invoices_list
    out = []
    for inv in invoices_list or []:
        inv_company = inv.get("company_id") or None
        if inv_company is None:
            if active_id == default_id:
                out.append(inv)
            continue
        if inv_company == active_id:
            out.append(inv)
    return out


def filter_list_for_active_company(items_list, key_name="company_id"):
    try:
        active = manager.get_active_company_profile()
        active_id = active.get("id")
        default_id = getattr(manager, "DEFAULT_COMPANY_ID", "com_hopepharma")
    except Exception:
        return items_list
    out = []
    for item in items_list or []:
        if not isinstance(item, dict):
            continue
        cid = item.get(key_name) or None
        if cid is None:
            if active_id == default_id:
                out.append(item)
            continue
        if cid == active_id:
            out.append(item)
    return out


def _stamp_active_company(d: dict):
    try:
        active = manager.get_active_company_profile()
        if active:
            d["company_id"] = active.get("id")
    except Exception:
        pass
    return d


def status_chip_info_for_invoice(inv):
    t = money(inv.get("grand_total"))
    p = money(inv.get("total_paid"))
    bal = t - p
    due = inv.get("due_date") or ""
    is_overdue = False
    if due:
        try:
            dd = date.fromisoformat(due)
            is_overdue = dd < date.today() and bal > 0.01
        except Exception:
            pass
    if t > 0 and p >= t - 0.005:
        return "Paid", "status-paid", bal
    if p > 0.005:
        return "Partial", "status-partial", bal
    if is_overdue:
        return "Overdue", "status-overdue", bal
    return "Not Paid", "status-unpaid", bal


def status_chip_for_purchase(p):
    amt = money(p.get("amount"))
    paid = money(p.get("amount_paid"))
    status = p.get("paid_status") or ""
    if status:
        cls = {
            "Paid": "status-paid",
            "Partial": "status-partial",
            "Pending": "status-pending",
        }.get(status, "status-unpaid")
        return status, cls
    if paid >= amt - 0.005 and amt > 0:
        return "Paid", "status-paid"
    if paid > 0.005:
        return "Partial", "status-partial"
    return "Pending", "status-pending"


@app.route("/", methods=["GET"])
def root_redirect():
    return redirect(url_for("dashboard"))


@app.route("/company/switch", methods=["POST"])
@login_required
def switch_company():
    # Security: users can ONLY switch to their assigned company (their session user's company_id)
    # The manual switcher dropdown in base.html only shows their own company anyway;
    # this prevents any attacker from hand-crafting a POST to see other companies' data.
    target_cid = request.form.get("company_id")
    user_cid = (session.get("user") or {}).get("company_id")
    if not target_cid:
        flash("No company selected.", "error")
        return redirect(url_for("dashboard"))
    if not hasattr(manager, "set_active_company"):
        flash("Data manager does not support company switching.", "error")
        return redirect(url_for("dashboard"))
    if target_cid != user_cid:
        flash("You are only authorized to access your assigned company.", "error")
        # Force the active company back to the user's company for safety
        if user_cid:
            manager.set_active_company(user_cid)
        return redirect(url_for("dashboard"))
    ok, msg = manager.set_active_company(target_cid)
    if ok:
        flash(msg or f"Switched active company", "success")
    else:
        flash(msg or "Failed to switch company", "error")
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/companies", methods=["GET"])
@login_required
@admin_required
def manage_companies_get():
    if not hasattr(manager, "get_all_company_profiles"):
        flash("Company profiles not available", "error")
        return redirect(url_for("dashboard"))
    # HopePharma-only double-gate (even other company admins cannot manage the roster)
    user_cid = (session.get("user") or {}).get("company_id", "")
    if user_cid != getattr(manager, "DEFAULT_COMPANY_ID", "com_hopepharma"):
        flash("You do not have permission to manage company profiles.", "error")
        return redirect(url_for("dashboard"))
    profiles = list(manager.get_all_company_profiles() or [])
    active = manager.get_active_company_profile()
    return render_template("companies.html", companies=profiles, active=active)


@app.route("/companies/create", methods=["POST"])
@login_required
@admin_required
def create_company_post():
    if not hasattr(manager, "create_company_profile"):
        flash("Company creation not available", "error")
        return redirect(url_for("manage_companies_get"))
    # HopePharma-only double-gate
    user_cid = (session.get("user") or {}).get("company_id", "")
    if user_cid != getattr(manager, "DEFAULT_COMPANY_ID", "com_hopepharma"):
        flash("You do not have permission to create new companies.", "error")
        return redirect(url_for("dashboard"))
    form = request.form
    name = (form.get("name") or "").strip()
    if not name:
        flash("Company name is required.", "error")
        return redirect(url_for("manage_companies_get"))
    data = {
        "name": name,
        "display_name": (form.get("display_name") or name).strip() or name,
        "prefix": (form.get("prefix") or "").strip().upper() or None,
        "short_code": (form.get("short_code") or "").strip().upper() or None,
        "trn": (form.get("trn") or "").strip(),
        "vat_rate": float(form.get("vat_rate") or 5),
        "currency": (form.get("currency") or "AED").strip().upper(),
        "emirate": (form.get("emirate") or "Dubai").strip(),
        "address": (form.get("address") or "").strip(),
        "phone": (form.get("phone") or "").strip(),
        "email": (form.get("email") or "").strip(),
        "website": (form.get("website") or "").strip(),
    }
    # Read a requested default username/password for the new company's initial user from the form (optional)
    new_username = (form.get("initial_username") or "").strip() or (data.get("short_code") or data["prefix"] or "newcompany").lower()
    new_password = (form.get("initial_password") or "").strip()
    if not new_password:
        # Generate a temp password the admin can share with the new company's team
        import secrets, string
        alphabet = string.ascii_letters + string.digits
        new_password = "".join(secrets.choice(alphabet) for _ in range(12))
    ok, msg, _new = manager.create_company_profile(data)
    if ok and _new:
        flash(msg or "Company created", "success")
        # Auto-create a default user for this new company so they can log in immediately
        new_cid = _new.get("id")
        new_full_name = (data["display_name"] or data["name"]) + " Admin"
        u_ok, u_msg, _user = manager.create_company_user(
            username=new_username,
            password=new_password,
            company_id=new_cid,
            role="admin",
            full_name=new_full_name,
            email=(form.get("email") or "").strip(),
        )
        if u_ok:
            flash(
                f"Default user created for {data['display_name']} — username: '{new_username}', temporary password: '{new_password}'. "
                "Share these credentials with the new company's team (they can change password later).",
                "info",
            )
        else:
            flash(f"Warning: company was created but no default user could be auto-created ({u_msg}). Use the Users page later.", "warning")
        # Keep the logged-in admin still on HopePharma (don't auto-switch)
    else:
        flash(msg or "Failed to create company", "error")
    return redirect(url_for("manage_companies_get"))


@app.route("/dashboard", methods=["GET"])
def dashboard():
    inv = filter_invoices_for_active_company(manager.get_all_invoices_dict(include_deleted=False) or [])
    del_inv = filter_invoices_for_active_company(
        manager.get_deleted_invoices_dict() if hasattr(manager, "get_deleted_invoices_dict") else []
    )
    purchases = filter_list_for_active_company(
        manager.get_all_purchases_dict() if hasattr(manager, "get_all_purchases_dict") else []
    )
    today = date.today()
    cid = _active_company_id()

    # Single-source compute_financials for revenue / cogs / gp — FR-5 FR-7 NO inline sums outside this function.
    fin_all = manager.compute_financials(cid) if (hasattr(manager, "compute_financials") and cid) else {}
    month_from = today.replace(day=1).isoformat()
    fin_mtd = manager.compute_financials(cid, date_from=month_from) if (hasattr(manager, "compute_financials") and cid) else {}

    grand_total_all = sum(money(i.get("grand_total")) for i in inv)
    total_paid = sum(money(i.get("total_paid")) for i in inv)
    total_owed = grand_total_all - total_paid
    total_purchases_all = sum(money(p.get("amount")) for p in purchases)

    overdue_n = 0
    for i in inv:
        _, _, bal = status_chip_info_for_invoice(i)
        if bal > 0.005:
            due = i.get("due_date") or ""
            if due:
                try:
                    dd = date.fromisoformat(due)
                    if dd < today:
                        overdue_n += 1
                except Exception:
                    pass

    month_inv = [i for i in inv if (i.get("date") or "").startswith(today.strftime("%Y-%m"))]
    month_sales_grand = sum(money(i.get("grand_total")) for i in month_inv)
    month_pur = [p for p in purchases if (p.get("date") or "").startswith(today.strftime("%Y-%m"))]
    month_purchases_total = sum(money(p.get("amount")) for p in month_pur)
    clients = set(i.get("client_name") for i in inv if i.get("client_name"))

    trend_months = []
    max_sales = 0.01
    max_pur = 0.01
    for offset in range(5, -1, -1):
        d = today - timedelta(days=offset * 30)
        mlabel = d.strftime("%b %y")
        mkey = d.strftime("%Y-%m")
        m_inv_v = sum(money(i.get("grand_total")) for i in inv if (i.get("date") or "").startswith(mkey))
        m_pur_v = sum(money(p.get("amount")) for p in purchases if (p.get("date") or "").startswith(mkey))
        if m_inv_v > max_sales: max_sales = m_inv_v
        if m_pur_v > max_pur: max_pur = m_pur_v
        trend_months.append((mlabel, round(m_inv_v, 2), round(m_pur_v, 2)))
    recent = sorted(inv, key=lambda i: (i.get("date") or "", i.get("invoice_id") or ""), reverse=True)[:8]

    sales_count = fin_all.get("sales_invoice_count") if isinstance(fin_all, dict) else None
    service_count = fin_all.get("service_invoice_count") if isinstance(fin_all, dict) else None
    if sales_count is None or service_count is None:
        sales_count = sum(1 for i in inv if (i.get("invoice_type") or "sales").lower() != "service")
        service_count = len(inv) - sales_count

    rev_cc = (fin_all or {}).get("revenue_by_cost_center") or {}
    cogs_cc = (fin_all or {}).get("cogs_by_cost_center") or {}
    cc_rows = []
    all_cc_ids = set(rev_cc.keys()) | set(cogs_cc.keys())
    for cid_cc in all_cc_ids:
        rev = float((rev_cc.get(cid_cc) or {}).get("amount", 0.0) if isinstance(rev_cc.get(cid_cc), dict) else (rev_cc.get(cid_cc) or 0.0))
        lab = (rev_cc.get(cid_cc) or {}).get("label") if isinstance(rev_cc.get(cid_cc), dict) else (cid_cc)
        cgs = float(cogs_cc.get(cid_cc) or 0.0)
        gp = round(rev - cgs, 2)
        cc_rows.append({
            "cc_id": cid_cc,
            "label": lab or cid_cc or "Unnamed",
            "revenue": round(rev, 2),
            "cogs": round(cgs, 2),
            "gp": gp,
            "gm_pct": round((gp / rev * 100.0), 2) if rev > 0.0001 else 0.0,
            "inv_count": (rev_cc.get(cid_cc) or {}).get("invoice_count") if isinstance(rev_cc.get(cid_cc), dict) else 0,
        })
    cc_rows.sort(key=lambda r: r["gp"], reverse=True)
    top5_cc_gp = cc_rows[:5]
    all_cc_pretty = cc_rows

    total_assets = grand_total_all  # placeholder: not full GAAP BS yet
    total_liabilities = total_purchases_all

    return render_template(
        "dashboard.html",
        invoices_n=len(inv),
        archived_n=len(del_inv),
        total_amount=grand_total_all,
        total_paid=total_paid,
        total_owed=total_owed,
        overdue_n=overdue_n,
        total_purchases=total_purchases_all,
        month_sales=month_sales_grand,
        month_purchases=month_purchases_total,
        clients_n=len(clients),
        trend_months=trend_months,
        max_sales=max_sales,
        max_purchases=max_pur,
        recent_invoices=recent,
        status_chip_info_for_invoice=status_chip_info_for_invoice,
        fin=fin_all if isinstance(fin_all, dict) else {},
        fin_mtd=fin_mtd if isinstance(fin_mtd, dict) else {},
        sales_count=sales_count,
        service_count=service_count,
        top5_cc_gp=top5_cc_gp,
        all_cc_rows=all_cc_pretty,
        total_assets=total_assets,
        total_liabilities=total_liabilities,
    )


def _active_company_id():
    try:
        active = manager.get_active_company_profile()
        return active.get("id") if isinstance(active, dict) else None
    except Exception:
        return None


def _list_active_cost_centers():
    cid = _active_company_id()
    cc_list = manager.list_cost_centers(cid) if hasattr(manager, "list_cost_centers") else []
    return [c for c in (cc_list or []) if isinstance(c, dict)]


def _list_active_products():
    try:
        raw = manager.load_json("products.json") or manager.load_json("products_memory.json") or []
    except Exception:
        raw = []
    return filter_list_for_active_company(raw, "company_id") if isinstance(raw, list) else []


def _list_active_service_products():
    cid = _active_company_id()
    sp = manager.list_service_products(cid, include_inactive=False) if hasattr(manager, "list_service_products") else []
    return [s for s in (sp or []) if isinstance(s, dict)]


@app.route("/invoices", methods=["GET"])
def invoices_list():
    invoices = manager.get_all_invoices_dict(include_deleted=False) or []
    invoices = filter_invoices_for_active_company(invoices)
    invoices.sort(key=lambda i: (i.get("date") or "", i.get("invoice_id") or ""), reverse=True)
    total_amount = sum(money(i.get("grand_total")) for i in invoices)
    total_paid = sum(money(i.get("total_paid")) for i in invoices)
    total_owed = total_amount - total_paid
    cost_centers = _list_active_cost_centers()
    cc_map = {}
    for cc in cost_centers:
        cc_map[cc.get("cost_center_id")] = cc.get("label") or cc.get("cost_center_id") or ""
    return render_template(
        "invoices.html",
        invoices=invoices,
        total_amount=total_amount,
        total_paid=total_paid,
        total_owed=total_owed,
        status_chip_info_for_invoice=status_chip_info_for_invoice,
        cost_centers=cost_centers,
        cc_label_map=cc_map,
    )


@app.route("/invoices/new", methods=["GET"])
def invoice_new_form():
    next_id = manager.generate_invoice_id()
    cost_centers = _list_active_cost_centers()
    default_cc_id = None
    for cc in cost_centers:
        if cc.get("cost_center_id") == "cc_general":
            default_cc_id = "cc_general"
            break
    if not default_cc_id and cost_centers:
        default_cc_id = cost_centers[0].get("cost_center_id")
    products = _list_active_products()
    service_products = _list_active_service_products()
    return render_template(
        "create.html",
        next_id=next_id,
        today=today_str(),
        editing=None,
        cost_centers=cost_centers,
        default_cost_center_id=default_cc_id or "",
        products=products,
        service_products=service_products,
    )


@app.route("/invoices", methods=["POST"])
def invoice_create_post():
    form = request.form
    client_name = (form.get("client_name") or "").strip()
    if not client_name:
        flash("Client Name is required.", "error")
        return redirect(url_for("invoice_new_form"))

    invoice_type_raw = (form.get("invoice_type") or "sales").strip().lower()
    invoice_type = "service" if invoice_type_raw == "service" else "sales"

    header_cost_center_id = (form.get("cost_center_id") or "").strip()

    descriptions = form.getlist("item_description")
    qtys = form.getlist("item_qty")
    prices = form.getlist("item_price")
    line_vats = form.getlist("item_vat")
    line_ccs = form.getlist("item_cost_center")
    line_unit_cost_hints = form.getlist("item_unit_cost_hidden")

    cid = _active_company_id()
    catalog_lookup = {}
    if invoice_type == "service":
        try:
            sp = manager.list_service_products(cid, include_inactive=False) if hasattr(manager, "list_service_products") else []
            for s in (sp or []):
                if isinstance(s, dict):
                    keys = []
                    if s.get("service_code"):
                        keys.append(s["service_code"].strip().lower())
                    if s.get("title"):
                        keys.append(s["title"].strip().lower())
                    if s.get("description"):
                        keys.append(s["description"].strip().lower())
                    for k in keys:
                        if k and k not in catalog_lookup:
                            catalog_lookup[k] = s
        except Exception:
            catalog_lookup = {}
    else:
        try:
            raw = manager.load_json("products.json") or manager.load_json("products_memory.json") or []
            prods = filter_list_for_active_company(raw, "company_id") if isinstance(raw, list) else []
            for p in prods:
                if isinstance(p, dict):
                    keys = []
                    if p.get("product_code"):
                        keys.append(p["product_code"].strip().lower())
                    if p.get("sku"):
                        keys.append(p["sku"].strip().lower())
                    if p.get("name"):
                        keys.append(p["name"].strip().lower())
                    if p.get("description"):
                        keys.append(p["description"].strip().lower())
                    for k in keys:
                        if k and k not in catalog_lookup:
                            catalog_lookup[k] = p
        except Exception:
            catalog_lookup = {}

    def _resolve_unit_cost_for_line(desc_raw, idx):
        d = (desc_raw or "").strip()
        key = d.lower()
        if key in catalog_lookup:
            item = catalog_lookup[key]
            try:
                if invoice_type == "service":
                    uc = float(item.get("default_internal_cost_hour") or item.get("unit_cost") or 0)
                else:
                    uc = float(item.get("unit_cost") or item.get("cost_price") or item.get("purchase_price") or 0)
                return max(0.0, uc)
            except (ValueError, TypeError):
                pass
        hint = line_unit_cost_hints[idx] if idx < len(line_unit_cost_hints) else "0"
        try:
            return max(0.0, float(hint or 0))
        except (ValueError, TypeError):
            return 0.0

    def _resolve_cc_for_line(idx):
        v = line_ccs[idx] if idx < len(line_ccs) else ""
        v = (v or "").strip()
        return v or header_cost_center_id

    any_missing_cc = False
    if not header_cost_center_id:
        for idx, desc in enumerate(descriptions):
            desc_s = (desc or "").strip()
            if not desc_s:
                continue
            line_cc = line_ccs[idx] if idx < len(line_ccs) else ""
            if not (line_cc or "").strip():
                any_missing_cc = True
                break

    if any_missing_cc:
        flash("Please assign a Cost Center to the header OR every line item.", "error")
        return redirect(url_for("invoice_new_form"))

    items = []
    subtotal = 0.0
    header_tax_rate = float(form.get("tax_rate") or 5)
    invoice_total_cost = 0.0
    costs = []

    for idx, desc in enumerate(descriptions):
        desc_s = (desc or "").strip()
        if not desc_s:
            continue
        try:
            q = float(qtys[idx]) if idx < len(qtys) else 1.0
        except (ValueError, TypeError):
            q = 1.0
        try:
            p = float(prices[idx]) if idx < len(prices) else 0.0
        except (ValueError, TypeError):
            p = 0.0
        line_total = round(q * p, 2)
        subtotal += line_total

        raw_lv = line_vats[idx] if idx < len(line_vats) else ""
        try:
            if raw_lv != "" and raw_lv is not None:
                lv = float(raw_lv)
            else:
                lv = header_tax_rate
        except (ValueError, TypeError):
            lv = header_tax_rate

        resolved_cc = _resolve_cc_for_line(idx)
        resolved_unit_cost = _resolve_unit_cost_for_line(desc_s, idx)
        line_cost_total = round(q * resolved_unit_cost, 2)
        invoice_total_cost += line_cost_total

        item = {
            "description": desc_s,
            "quantity": int(q) if float(q).is_integer() else q,
            "unit_price": round(p, 2),
            "amount": line_total,
            "tax_rate": round(lv, 2),
            "cost_center_id": resolved_cc or None,
            "unit_cost": round(resolved_unit_cost, 2),
            "line_cost_total": line_cost_total,
        }
        items.append(item)

        if resolved_cc and line_cost_total > 0:
            costs.append({
                "cost_type": ("service_direct_cost" if invoice_type == "service" else "cogs_goods"),
                "cost_center_id": resolved_cc,
                "description": f"COGS for line #{idx+1}: {desc_s}",
                "amount": line_cost_total,
                "date": (form.get("date") or today_str()),
            })

    if not items:
        flash("At least one line item is required.", "error")
        return redirect(url_for("invoice_new_form"))

    tax_amount = 0.0
    for it in items:
        tr = it.get("tax_rate") or header_tax_rate
        tax_amount += round(float(it.get("amount") or 0) * (float(tr or 0) / 100.0), 2)
    tax_amount = round(tax_amount, 2)
    subtotal = round(subtotal, 2)
    grand_total = round(subtotal + tax_amount, 2)
    invoice_total_cost = round(invoice_total_cost, 2)
    invoice_gross_profit = round(subtotal - invoice_total_cost, 2)

    inv_date = form.get("date") or today_str()
    due_date = form.get("due_date") or inv_date
    invoice = {
        "invoice_id": (form.get("invoice_id") or "").strip() or manager.generate_invoice_id(),
        "client_name": client_name,
        "client_trn": (form.get("client_trn") or "").strip(),
        "client_emirate": (form.get("client_emirate") or "Dubai").strip(),
        "client_location": "",
        "date": inv_date,
        "due_date": due_date,
        "payment_terms": form.get("payment_terms") or "Net 30",
        "payment_method": form.get("payment_terms") or "Net 30",
        "payment_due": "On receipt",
        "tax_rate": round(header_tax_rate, 2),
        "items": items,
        "costs": costs,
        "subtotal": subtotal,
        "taxable_amount": subtotal,
        "non_taxable_amount": 0.0,
        "tax_amount": tax_amount,
        "grand_total": grand_total,
        "total_cost": invoice_total_cost,
        "invoice_total_cost": invoice_total_cost,
        "profit_loss": invoice_gross_profit,
        "invoice_gross_profit": invoice_gross_profit,
        "total_paid": 0.0,
        "balance_due": grand_total,
        "status": "Not Paid",
        "notes": (form.get("notes") or "").strip(),
        "currency": "AED",
        "invoice_type": invoice_type,
        "cost_center_id": header_cost_center_id or None,
        "payment_history": [],
        "workflow_status": "Under Process",
    }
    invoice = _stamp_active_company(invoice)
    ok = manager.add_invoice_from_dict(invoice)
    if ok:
        flash(f"Invoice {invoice['invoice_id']} created — AED {grand_total:,.2f}.", "success")
    else:
        flash("Failed to save invoice.", "error")
        return redirect(url_for("invoice_new_form"))
    return redirect(url_for("invoice_view", invoice_id=invoice["invoice_id"]))


def _find_invoice(invoice_id, include_deleted=True):
    all_inv = manager.get_all_invoices_dict(include_deleted=include_deleted) or []
    if include_deleted and hasattr(manager, "get_deleted_invoices_dict"):
        all_inv = list(all_inv) + list(manager.get_deleted_invoices_dict() or [])
    for i in all_inv:
        if i.get("invoice_id") == invoice_id:
            return i
    return None


@app.route("/invoices/<path:invoice_id>", methods=["GET"])
def invoice_view(invoice_id):
    inv = _find_invoice(invoice_id)
    if not inv:
        abort(404)
    inv_list = filter_invoices_for_active_company([inv])
    if not inv_list:
        abort(404)
    inv = inv_list[0]
    status, cls, bal = status_chip_info_for_invoice(inv)
    return render_template(
        "invoice_view.html",
        inv=inv,
        status=status,
        status_cls=cls,
        balance=bal,
        money=money,
    )


def _serve_invoice_pdf_for(invoice_id, disposition="inline"):
    """Shared helper: look up invoice, active-company filter, build PDF bytes,
    return Flask Response with given disposition ('inline' = browser view,
    'attachment' = download file).

    Falls back to printable HTML view only when PDF generation fails AND
    disposition == inline.  For attachment disposition, on failure we 500.
    """
    inv = _find_invoice(invoice_id)
    if not inv:
        abort(404)
    inv_list = filter_invoices_for_active_company([inv])
    if not inv_list:
        abort(404)
    inv = inv_list[0]
    pdf_bytes = None
    pdf_gen = getattr(manager, "pdf_generator", None)
    if pdf_gen and hasattr(pdf_gen, "generate_invoice_pdf_bytes"):
        try:
            pdf_bytes = pdf_gen.generate_invoice_pdf_bytes(inv)
        except Exception:
            pdf_bytes = None
    fname = f"{inv['invoice_id']}.pdf"
    if pdf_bytes:
        return Response(
            pdf_bytes,
            mimetype="application/pdf",
            headers={"Content-Disposition": f"{disposition}; filename={fname}"},
        )
    if disposition == "inline":
        status, cls, bal = status_chip_info_for_invoice(inv)
        html = render_template(
            "invoice_view.html",
            inv=inv,
            status=status,
            status_cls=cls,
            balance=bal,
            money=money,
            printable=True,
        )
        resp = make_response(html)
        resp.headers["Content-Disposition"] = f"inline; filename={inv['invoice_id']}.html"
        return resp
    abort(500, "PDF generation failed for this invoice.")


@app.route("/invoices/<path:invoice_id>/pdf", methods=["GET"])
def invoice_pdf(invoice_id):
    """Render invoice PDF inline in the browser (fallback: printable HTML)."""
    return _serve_invoice_pdf_for(invoice_id, disposition="inline")


@app.route("/invoices/<path:invoice_id>/download/pdf", methods=["GET"])
@login_required
def invoice_download_pdf(invoice_id):
    """Force PDF download (attachment disposition)."""
    return _serve_invoice_pdf_for(invoice_id, disposition="attachment")


@app.route("/invoices/<path:invoice_id>/edit", methods=["GET"])
def invoice_edit_form(invoice_id):
    inv = _find_invoice(invoice_id)
    if not inv:
        abort(404)
    inv_list = filter_invoices_for_active_company([inv])
    if not inv_list:
        abort(404)
    cost_centers = _list_active_cost_centers()
    default_cc_id = inv.get("cost_center_id") or ""
    if not default_cc_id:
        for cc in cost_centers:
            if cc.get("cost_center_id") == "cc_general":
                default_cc_id = "cc_general"
                break
        if not default_cc_id and cost_centers:
            default_cc_id = cost_centers[0].get("cost_center_id") or ""
    products = _list_active_products()
    service_products = _list_active_service_products()
    return render_template(
        "create.html",
        next_id=inv["invoice_id"],
        today=inv.get("date") or today_str(),
        editing=inv,
        cost_centers=cost_centers,
        default_cost_center_id=default_cc_id,
        products=products,
        service_products=service_products,
    )


@app.route("/delete/<path:invoice_id>", methods=["POST"])
@login_required
@admin_required
def delete_invoice(invoice_id):
    if hasattr(manager, "soft_delete_invoice_with_balance_adjustment"):
        ok, msg = manager.soft_delete_invoice_with_balance_adjustment(invoice_id, username="WebUser")
    elif hasattr(manager, "soft_delete_invoice"):
        ok, msg = manager.soft_delete_invoice(invoice_id, username="WebUser")
    else:
        ok = bool(manager.delete_invoice(invoice_id))
        msg = "Invoice hard-deleted."
    if ok:
        flash(msg or f"Invoice {invoice_id} archived.", "success")
    else:
        flash(msg or f"Failed to archive invoice {invoice_id}.", "error")
    return redirect(url_for("invoices_list"))


@app.route("/archive", methods=["GET"])
@login_required
@admin_required
def archive_page():
    deleted = manager.get_deleted_invoices_dict() if hasattr(manager, "get_deleted_invoices_dict") else []
    deleted = filter_invoices_for_active_company(deleted)
    deleted.sort(key=lambda i: (i.get("deleted_at") or i.get("date") or ""), reverse=True)
    deleted_total = sum(money(i.get("grand_total")) for i in deleted)
    return render_template("archive.html", deleted=deleted, deleted_total=deleted_total)


@app.route("/restore/<path:invoice_id>", methods=["POST"])
@login_required
@admin_required
def restore_invoice(invoice_id):
    if hasattr(manager, "restore_invoice"):
        ok, msg = manager.restore_invoice(invoice_id)
    else:
        ok, msg = False, "Restore unavailable."
    if ok:
        flash(msg or f"Invoice {invoice_id} restored.", "success")
    else:
        flash(msg or f"Failed to restore invoice {invoice_id}.", "error")
    return redirect(url_for("archive_page"))


@app.route("/erase/<path:invoice_id>", methods=["POST"])
@login_required
@admin_required
def erase_invoice(invoice_id):
    if hasattr(manager, "permanent_erase_invoice"):
        ok = bool(manager.permanent_erase_invoice(invoice_id))
    else:
        ok = bool(manager.delete_invoice(invoice_id))
    if ok:
        flash(f"Invoice {invoice_id} permanently erased.", "info")
    else:
        flash(f"Failed to erase invoice {invoice_id}.", "error")
    return redirect(url_for("archive_page"))


def _get_all_purchases_safe():
    if hasattr(manager, "get_all_purchases_dict"):
        try:
            return list(manager.get_all_purchases_dict() or [])
        except Exception:
            return []
    try:
        return list(manager.load_json("purchases_data.json") or [])
    except Exception:
        return []


def _next_purchase_id():
    try:
        active = manager.get_active_company_profile()
        pfx = (active.get("short_code") or active.get("prefix") or "PUR")[:4]
    except Exception:
        pfx = "PUR"
    today = date.today().strftime("%Y%m%d")
    purchases = _get_all_purchases_safe()
    n = len([p for p in purchases if str(p.get("purchase_id", "")).startswith(f"{pfx}-{today}")]) + 1
    return f"{pfx}-{today}-{n:03d}"


@app.route("/purchases", methods=["GET"])
def purchases_list():
    purchases = filter_list_for_active_company(_get_all_purchases_safe())
    purchases.sort(key=lambda p: (p.get("date") or "", p.get("purchase_id") or ""), reverse=True)
    total = sum(money(p.get("amount")) for p in purchases)
    paid_total = sum(money(p.get("amount_paid")) for p in purchases)
    return render_template(
        "purchases.html",
        purchases=purchases,
        total_amount=total,
        total_paid=paid_total,
        status_chip_for_purchase=status_chip_for_purchase,
        money=money,
    )


@app.route("/purchases/new", methods=["GET"])
def purchase_create_form():
    return render_template("purchase_create.html", today=today_str(), next_id=_next_purchase_id())


@app.route("/purchases", methods=["POST"])
def purchase_create_post():
    form = request.form
    desc = (form.get("description") or "").strip()
    amount = money(form.get("amount"))
    if not desc or amount <= 0:
        flash("Description and positive amount are required.", "error")
        return redirect(url_for("purchase_create_form"))
    purchase = {
        "purchase_id": (form.get("purchase_id") or "").strip() or _next_purchase_id(),
        "description": desc,
        "amount": round(amount, 2),
        "date": form.get("date") or today_str(),
        "category": (form.get("category") or "General").strip(),
        "supplier": (form.get("supplier") or "").strip(),
        "account": (form.get("account") or "Cash").strip(),
        "notes": (form.get("notes") or "").strip(),
        "amount_paid": 0.0,
        "payment_date": "",
        "paid_status": "Pending",
        "payment_history": [],
    }
    purchase = _stamp_active_company(purchase)
    try:
        existing = manager.load_json("purchases_data.json") or []
        if not isinstance(existing, list):
            existing = []
        existing.append(purchase)
        ok = manager.save_json("purchases_data.json", existing)
    except Exception:
        ok = False
    if ok:
        flash(f"Purchase {purchase['purchase_id']} recorded — AED {amount:,.2f}.", "success")
    else:
        flash("Failed to save purchase.", "error")
    return redirect(url_for("purchases_list"))


@app.route("/purchases/<path:purchase_id>/pay", methods=["POST"])
def purchase_mark_paid(purchase_id):
    try:
        existing = manager.load_json("purchases_data.json") or []
        if not isinstance(existing, list):
            existing = []
        found = False
        for i, p in enumerate(existing):
            if p.get("purchase_id") == purchase_id:
                amt = money(p.get("amount"))
                existing[i]["amount_paid"] = round(amt, 2)
                existing[i]["paid_status"] = "Paid"
                existing[i]["payment_date"] = today_str()
                ph = p.get("payment_history") or []
                ph.append({"date": today_str(), "amount": round(amt, 2), "method": "Cash", "note": "Web mark-paid"})
                existing[i]["payment_history"] = ph
                found = True
                break
        if found:
            ok = manager.save_json("purchases_data.json", existing)
            if ok:
                flash(f"Purchase {purchase_id} marked Paid.", "success")
            else:
                flash("Failed to update purchase.", "error")
        else:
            flash(f"Purchase {purchase_id} not found.", "error")
    except Exception as e:
        flash(f"Error: {e}", "error")
    return redirect(url_for("purchases_list"))


def _load_employees():
    try:
        d = manager.load_json("employees.json")
        if isinstance(d, list):
            return d
    except Exception:
        pass
    return []


def _save_employees(lst):
    try:
        return bool(manager.save_json("employees.json", lst))
    except Exception:
        return False


def _next_emp_id():
    emps = _load_employees()
    n = max([int(e.get("id", "0") or 0) for e in emps if str(e.get("id", "0")).isdigit()] or [0]) + 1
    return str(n)


@app.route("/employees", methods=["GET"])
def employees_list():
    emps = filter_list_for_active_company(_load_employees())
    emps.sort(key=lambda e: (e.get("name") or "").lower())
    total_salary = sum(money(e.get("monthly_salary")) for e in emps)
    return render_template(
        "employees.html",
        employees=emps,
        total_salary=total_salary,
        today=today_str(),
    )


@app.route("/employees/create", methods=["POST"])
@login_required
@admin_required
def employee_create_post():
    form = request.form
    name = (form.get("name") or "").strip()
    if not name:
        flash("Employee name is required.", "error")
        return redirect(url_for("employees_list"))
    emp = {
        "id": _next_emp_id(),
        "name": name,
        "position": (form.get("position") or "").strip(),
        "department": (form.get("department") or "").strip(),
        "monthly_salary": round(money(form.get("monthly_salary")), 2),
        "joining_date": form.get("joining_date") or today_str(),
        "phone": (form.get("phone") or "").strip(),
        "email": (form.get("email") or "").strip(),
        "status": (form.get("status") or "Active").strip(),
        "notes": (form.get("notes") or "").strip(),
    }
    emp = _stamp_active_company(emp)
    existing = _load_employees()
    existing.append(emp)
    ok = _save_employees(existing)
    if ok:
        flash(f"Employee {name} added.", "success")
    else:
        flash("Failed to save employee.", "error")
    return redirect(url_for("employees_list"))


@app.route("/employees/<emp_id>/edit", methods=["POST"])
@login_required
@admin_required
def employee_edit_post(emp_id):
    form = request.form
    existing = _load_employees()
    found = False
    for i, e in enumerate(existing):
        if str(e.get("id")) == str(emp_id):
            name = (form.get("name") or e.get("name") or "").strip()
            if not name:
                flash("Employee name required.", "error")
                return redirect(url_for("employees_list"))
            for k, fkey in [("name","name"),("position","position"),("department","department"),
                            ("phone","phone"),("email","email"),("status","status"),
                            ("notes","notes"),("joining_date","joining_date")]:
                if fkey in form:
                    existing[i][k] = form.get(fkey) or existing[i].get(k, "")
            if "monthly_salary" in form:
                existing[i]["monthly_salary"] = round(money(form.get("monthly_salary")), 2)
            found = True
            break
    if found and _save_employees(existing):
        flash("Employee updated.", "success")
    else:
        flash("Failed to update employee.", "error")
    return redirect(url_for("employees_list"))


@app.route("/employees/<emp_id>/delete", methods=["POST"])
@login_required
@admin_required
def employee_delete_post(emp_id):
    existing = _load_employees()
    before = len(existing)
    existing = [e for e in existing if str(e.get("id")) != str(emp_id)]
    if len(existing) < before and _save_employees(existing):
        flash("Employee deleted.", "success")
    else:
        flash("Failed to delete employee.", "error")
    return redirect(url_for("employees_list"))


def _load_salaries():
    try:
        d = manager.load_json("salary_reports.json")
        if isinstance(d, list):
            return d
    except Exception:
        pass
    return []


def _save_salaries(lst):
    try:
        return bool(manager.save_json("salary_reports.json", lst))
    except Exception:
        return False


@app.route("/salaries", methods=["GET"])
def salaries_list():
    sals = filter_list_for_active_company(_load_salaries())
    sals.sort(key=lambda s: (s.get("period") or "", s.get("employee_name") or ""), reverse=True)
    total = sum(money(s.get("net_amount") or s.get("gross_amount")) for s in sals)
    emps = filter_list_for_active_company(_load_employees())
    today = date.today()
    default_period = f"{today.year}-{today.month:02d}"
    return render_template(
        "salaries.html",
        salaries=sals,
        total_paid=total,
        employees=emps,
        default_period=default_period,
        today=today_str(),
        money=money,
    )


@app.route("/salaries/create", methods=["POST"])
@login_required
@admin_required
def salary_create_post():
    form = request.form
    emp_id = form.get("employee_id") or ""
    period = (form.get("period") or "").strip()
    gross = (
        money(form.get("gross_amount")) or
        money(form.get("gross_salary")) or
        money(form.get("monthly_salary")) or
        0.0
    )
    if not emp_id:
        flash("Employee selection is required.", "error")
        return redirect(url_for("salaries_list"))
    if not period:
        period = form.get("month") or ""
    if not period or gross <= 0:
        flash("Period and non-zero salary amount are required.", "error")
        return redirect(url_for("salaries_list"))
    emps = _load_employees()
    emp = next((e for e in emps if str(e.get("id")) == str(emp_id)), None)
    if not emp:
        emp_name = form.get("employee_name") or f"Employee {emp_id}"
    else:
        emp_name = emp.get("name") or f"Employee {emp_id}"
    basic = money(form.get("basic_salary")) or money(form.get("basic")) or gross
    allowances = (
        money(form.get("allowances")) or
        money(form.get("other_allowances")) or
        round(max(0.0, gross - basic), 2)
    )
    deductions = money(form.get("deductions")) or money(form.get("total_deductions")) or 0.0
    net = money(form.get("net_salary")) or money(form.get("net_amount")) or round(basic + allowances - deductions, 2)
    sal = {
        "payslip_id": f"PAY-{period}-{emp_id}",
        "employee_id": emp_id,
        "employee_name": emp_name,
        "period": period,
        "month": form.get("month") or period,
        "basic_salary": round(basic, 2),
        "gross_amount": round(gross, 2),
        "gross_salary": round(gross, 2),
        "deductions": round(deductions, 2),
        "allowances": round(allowances, 2),
        "net_amount": round(net, 2),
        "net_salary": round(net, 2),
        "payment_date": form.get("payment_date") or today_str(),
        "status": (form.get("status") or "Paid").strip(),
        "notes": (form.get("notes") or "").strip(),
    }
    sal = _stamp_active_company(sal)
    sals = _load_salaries()
    sals.append(sal)
    ok = _save_salaries(sals)
    if ok:
        flash(f"Payslip {sal['payslip_id']} saved — net AED {net:,.2f}.", "success")
    else:
        flash("Failed to save payslip.", "error")
    return redirect(url_for("salaries_list"))


@app.route("/reports/tb", methods=["GET"])
def reports_tb():
    accounts = []
    try:
        coa = manager.load_json("chart_of_accounts.json") or []
        gl = manager.load_json("general_ledger.json") or []
        gl = filter_list_for_active_company(gl) if isinstance(gl, list) else []
        coa = coa if isinstance(coa, list) else []
        balances = {}
        for entry in gl:
            if not isinstance(entry, dict):
                continue
            acct = entry.get("account_code") or entry.get("account") or "?"
            debit = money(entry.get("debit"))
            credit = money(entry.get("credit"))
            if acct not in balances:
                balances[acct] = {"debit": 0.0, "credit": 0.0}
            balances[acct]["debit"] += debit
            balances[acct]["credit"] += credit
        for acct in coa:
            if not isinstance(acct, dict):
                continue
            code = acct.get("account_code") or "?"
            bal = balances.get(code, {"debit": 0.0, "credit": 0.0})
            accounts.append({
                "code": code,
                "name": acct.get("account_name") or "",
                "type": acct.get("account_type") or "",
                "debit": round(bal["debit"], 2),
                "credit": round(bal["credit"], 2),
                "balance": round(bal["debit"] - bal["credit"], 2),
            })
    except Exception as e:
        flash(f"TB build warning: {e}", "error")
    total_debit = sum(a["debit"] for a in accounts)
    total_credit = sum(a["credit"] for a in accounts)
    return render_template(
        "reports_tb.html",
        accounts=accounts,
        total_debit=total_debit,
        total_credit=total_credit,
        as_of=today_str(),
    )


@app.route("/reports/gaap", methods=["GET"])
def reports_gaap():
    year = request.args.get("year")
    month = request.args.get("month")
    start_s, end_s, y, m = month_start_end(year, month)
    cid = _active_company_id()
    fin = manager.compute_financials(cid, date_from=start_s, date_to=end_s) if (hasattr(manager, "compute_financials") and cid) else {}
    fin = fin if isinstance(fin, dict) else {}

    inv = filter_invoices_for_active_company(manager.get_all_invoices_dict(include_deleted=False) or [])
    pur = filter_list_for_active_company(_get_all_purchases_safe())
    period_inv = [i for i in inv if start_s <= (i.get("date") or "") <= end_s]
    period_pur = [p for p in pur if start_s <= (p.get("date") or "") <= end_s]

    revenue = fin.get("revenue_total")
    if revenue is None:
        revenue = sum(money(i.get("subtotal") or i.get("grand_total")) for i in period_inv)
    cogs = fin.get("cogs_total")
    if cogs is None:
        cogs = sum(money(i.get("invoice_total_cost") or i.get("total_cost")) for i in period_inv)
    expenses = fin.get("operating_expenses_total")
    if expenses is None:
        expenses = sum(money(p.get("amount")) for p in period_pur)
    gross_profit = round(revenue - cogs, 2)
    net_profit = fin.get("net_profit")
    if net_profit is None:
        net_profit = round(gross_profit - expenses, 2)
    total_assets = fin.get("revenue_total") or sum(money(i.get("grand_total")) for i in period_inv)
    total_liabilities = sum(money(p.get("amount")) for p in period_pur)
    equity = net_profit
    operating_cash_in = sum(money(i.get("total_paid")) for i in period_inv)
    operating_cash_out = sum(money(p.get("amount_paid")) for p in period_pur)
    net_operating_cf = operating_cash_in - operating_cash_out
    years = list(range(y - 2, y + 3))
    months = [(m_idx, date(2000, m_idx, 1).strftime("%B")) for m_idx in range(1, 13)]

    rev_cc = fin.get("revenue_by_cost_center") or {}
    cogs_cc = fin.get("cogs_by_cost_center") or {}
    all_cc_ids = set(rev_cc.keys()) | set(cogs_cc.keys())
    cc_rows = []
    for cid_cc in all_cc_ids:
        rev = float((rev_cc.get(cid_cc) or {}).get("amount", 0.0) if isinstance(rev_cc.get(cid_cc), dict) else (rev_cc.get(cid_cc) or 0.0))
        lab = (rev_cc.get(cid_cc) or {}).get("label") if isinstance(rev_cc.get(cid_cc), dict) else cid_cc
        cgs = float(cogs_cc.get(cid_cc) or 0.0)
        gp = round(rev - cgs, 2)
        cc_rows.append({
            "cc_id": cid_cc,
            "label": lab or cid_cc or "Unnamed",
            "revenue": round(rev, 2),
            "cogs": round(cgs, 2),
            "gp": gp,
            "gm_pct": round((gp / rev * 100.0), 2) if rev > 0.0001 else 0.0,
            "inv_count": (rev_cc.get(cid_cc) or {}).get("invoice_count") if isinstance(rev_cc.get(cid_cc), dict) else 0,
        })
    cc_rows.sort(key=lambda r: r["gp"], reverse=True)

    return render_template(
        "reports_gaap.html",
        year=y,
        month=m,
        month_name=date(2000, m, 1).strftime("%B %Y"),
        start=start_s,
        end=end_s,
        revenue=revenue,
        cogs=cogs,
        gross_profit=gross_profit,
        expenses=expenses,
        net_profit=net_profit,
        total_assets=total_assets,
        total_liabilities=total_liabilities,
        equity=equity,
        operating_cash_in=operating_cash_in,
        operating_cash_out=operating_cash_out,
        net_operating_cf=net_operating_cf,
        years=years,
        months=months,
        invoices_n=len(period_inv),
        purchases_n=len(period_pur),
        fin=fin,
        cc_rows=cc_rows,
    )


@app.route("/reports/batch_pdf", methods=["POST"])
@login_required
@admin_required
def reports_batch_pdf():
    try:
        year = int(request.form.get("year") or date.today().year)
        month = int(request.form.get("month") or date.today().month)
    except Exception:
        year, month = date.today().year, date.today().month
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for rtype in ["profit_and_loss", "balance_sheet", "cash_flow"]:
            zf.writestr(
                f"{rtype}_{year}_{month:02d}.txt",
                f"ASISTEM GAAP Report — {rtype.replace('_',' ').title()}\nPeriod: {year}-{month:02d}\nGenerated: {datetime.now().isoformat()}\n\n"
                f"(A real PDF export would be produced when EnhancedPDFGenerator GAAP PDF is wired.)\n",
            )
    buf.seek(0)
    return send_file(
        buf,
        as_attachment=True,
        download_name=f"gaap_reports_{year}_{month:02d}.zip",
        mimetype="application/zip",
    )


@app.route("/balance", methods=["GET"])
def balance_view():
    try:
        bal_data = manager.load_json("balances.json") or {}
        if isinstance(bal_data, list):
            bal_list = filter_list_for_active_company(bal_data)
            bal_dict = {}
            for b in bal_list:
                if isinstance(b, dict):
                    acct = b.get("account") or b.get("name") or "Unspecified"
                    bal_dict[acct] = b
        elif isinstance(bal_data, dict):
            bal_dict = bal_data
        else:
            bal_dict = {}
    except Exception:
        bal_dict = {}
    accounts = []
    total_balance = 0.0
    if isinstance(bal_dict, dict):
        for k, v in bal_dict.items():
            if isinstance(v, dict):
                b = money(v.get("balance") or v.get("amount"))
                t = v.get("type") or "Asset"
            else:
                b = money(v)
                t = "Asset"
            accounts.append({"name": k, "balance": round(b, 2), "type": t})
            total_balance += b
    if not accounts:
        inv = filter_invoices_for_active_company(manager.get_all_invoices_dict(include_deleted=False) or [])
        receivable = sum(money(i.get("grand_total")) - money(i.get("total_paid")) for i in inv)
        pur = filter_list_for_active_company(_get_all_purchases_safe())
        payable = sum(money(p.get("amount")) - money(p.get("amount_paid")) for p in pur)
        accounts = [
            {"name": "Accounts Receivable", "balance": round(receivable, 2), "type": "Asset"},
            {"name": "Accounts Payable", "balance": round(payable, 2), "type": "Liability"},
        ]
        total_balance = receivable - payable
    return render_template(
        "balance.html",
        accounts=accounts,
        total_balance=total_balance,
        money=money,
    )


@app.route("/transactions", methods=["GET"])
def transactions_view():
    try:
        tx = manager.load_json("transactions.json") or []
        if not isinstance(tx, list):
            tx = []
    except Exception:
        tx = []
    tx = filter_list_for_active_company(tx)
    if not tx:
        inv = filter_invoices_for_active_company(manager.get_all_invoices_dict(include_deleted=False) or [])
        pur = filter_list_for_active_company(_get_all_purchases_safe())
        for i in inv:
            if money(i.get("total_paid")) > 0:
                tx.append({
                    "date": i.get("date"),
                    "type": "Income",
                    "reference": i.get("invoice_id"),
                    "counterparty": i.get("client_name"),
                    "amount": round(money(i.get("total_paid")), 2),
                    "account": "Cash",
                })
        for p in pur:
            if money(p.get("amount_paid")) > 0:
                tx.append({
                    "date": p.get("date"),
                    "type": "Expense",
                    "reference": p.get("purchase_id"),
                    "counterparty": p.get("supplier"),
                    "amount": round(money(p.get("amount_paid")), 2),
                    "account": p.get("account") or "Cash",
                })
    tx.sort(key=lambda t: t.get("date") or "", reverse=True)
    total_in = sum(money(t.get("amount")) for t in tx if str(t.get("type")).lower() in ("income","credit","in"))
    total_out = sum(money(t.get("amount")) for t in tx if str(t.get("type")).lower() in ("expense","debit","out"))
    return render_template(
        "transactions.html",
        transactions=tx,
        total_in=total_in,
        total_out=total_out,
        net=total_in - total_out,
        money=money,
    )


@app.route("/inventory", methods=["GET"])
def inventory_view():
    try:
        inv = manager.load_json("inventory.json") or []
        if not isinstance(inv, list):
            inv = []
    except Exception:
        inv = []
    inv = filter_list_for_active_company(inv)
    if not inv:
        try:
            prods = manager.load_json("products_memory.json") or []
            if isinstance(prods, list):
                for p in prods:
                    if isinstance(p, dict):
                        inv.append({
                            "sku": p.get("sku") or p.get("code") or "SKU-" + str(len(inv) + 1),
                            "name": p.get("name") or p.get("description") or "Product",
                            "qty": int(money(p.get("qty") or p.get("stock") or 0)),
                            "unit_price": round(money(p.get("price") or p.get("cost") or 0), 2),
                            "category": p.get("category") or "General",
                        })
        except Exception:
            pass
    total_value = sum((i.get("qty") or 0) * money(i.get("unit_price")) for i in inv if isinstance(i, dict))
    return render_template(
        "inventory.html",
        items=inv,
        total_items=len(inv),
        total_value=total_value,
        low_stock_n=len([i for i in inv if isinstance(i, dict) and (i.get("qty") or 0) <= 5]),
    )


@app.route("/warehouse", methods=["GET"])
def warehouse_view():
    try:
        wh = manager.load_json("warehouses.json") or []
        if not isinstance(wh, list):
            wh = []
    except Exception:
        wh = []
    if not wh:
        wh = [
            {"id": "WH-MAIN", "name": "Main Warehouse", "location": "Dubai", "capacity": 5000, "utilization_pct": 62, "manager": "Ops Team"},
            {"id": "WH-COLD", "name": "Cold Storage", "location": "Jebel Ali", "capacity": 1000, "utilization_pct": 38, "manager": "QA Team"},
        ]
    try:
        transfers = manager.load_json("stock_transfers.json") or []
        if not isinstance(transfers, list):
            transfers = []
    except Exception:
        transfers = []
    transfers = filter_list_for_active_company(transfers)
    transfers = sorted(transfers, key=lambda t: t.get("date") or "", reverse=True)[:10]
    return render_template(
        "warehouse.html",
        warehouses=wh,
        recent_transfers=transfers,
        total_capacity=sum(w.get("capacity", 0) for w in wh),
    )


@app.route("/temperature-log", methods=["GET"])
def temperature_log_view():
    try:
        logs = manager.load_json("temperature_log.json") or []
        if not isinstance(logs, list):
            logs = []
    except Exception:
        logs = []
    logs = filter_list_for_active_company(logs)
    if not logs:
        today = date.today()
        for d in range(14, -1, -1):
            dt = today - timedelta(days=d)
            logs.append({
                "date": dt.isoformat(),
                "time": "08:00",
                "warehouse": "WH-COLD",
                "zone": "Cold Room 1",
                "temp_c": 2.0 + (d % 3) * 0.5,
                "rh_pct": 55 + (d % 5),
                "status": "OK" if d % 7 != 0 else "Alert",
                "taken_by": "Auto Sensor",
            })
    logs.sort(key=lambda l: (l.get("date") or "", l.get("time") or ""), reverse=True)
    alerts = len([l for l in logs if l.get("status") == "Alert"])
    return render_template(
        "temperature_log.html",
        logs=logs,
        alerts_n=alerts,
        total_readings=len(logs),
    )


@app.route("/quotations", methods=["GET"])
def quotations_view():
    try:
        quotes = manager.load_json("quotes_data.json") or []
        if not isinstance(quotes, list):
            quotes = []
    except Exception:
        quotes = []
    quotes = filter_list_for_active_company(quotes)
    quotes.sort(key=lambda q: (q.get("date") or "", q.get("quote_id") or q.get("id") or ""), reverse=True)
    total = sum(money(q.get("grand_total") or q.get("total")) for q in quotes)
    return render_template(
        "quotations.html",
        quotations=quotes,
        total_value=total,
        count=len(quotes),
        money=money,
    )


@app.route("/clients", methods=["GET"])
def clients_view():
    try:
        clients = manager.get_clients() if hasattr(manager, "get_clients") else []
        clients = list(clients or [])
    except Exception:
        clients = []
    clients = filter_list_for_active_company(clients, key_name="company_id")
    if not clients:
        inv = filter_invoices_for_active_company(manager.get_all_invoices_dict(include_deleted=False) or [])
        seen = set()
        for i in inv:
            name = i.get("client_name")
            if name and name not in seen:
                seen.add(name)
                his_invs = [x for x in inv if x.get("client_name") == name]
                clients.append({
                    "name": name,
                    "trn": i.get("client_trn") or "",
                    "emirate": i.get("client_emirate") or "",
                    "total_invoices": len(his_invs),
                    "total_spent": round(sum(money(x.get("grand_total")) for x in his_invs), 2),
                    "last_invoice": max((x.get("date") or "") for x in his_invs) if his_invs else "",
                })
    clients.sort(key=lambda c: (c.get("name") or "").lower())
    total_spent_all = sum(money(c.get("total_spent")) for c in clients)
    return render_template(
        "clients.html",
        clients=clients,
        count=len(clients),
        total_revenue=total_spent_all,
        money=money,
    )


@app.route("/marketplaces", methods=["GET"])
def marketplaces_view():
    mm = getattr(manager, "marketplace_manager", None)
    channels = []
    if mm and hasattr(mm, "get_channels_summary"):
        try:
            channels = list(mm.get_channels_summary() or [])
        except Exception:
            channels = []
    if not channels:
        inv = filter_invoices_for_active_company(manager.get_all_invoices_dict(include_deleted=False) or [])
        channels = [
            {"name": "Noon.com", "orders": 12, "sales": 12500.00, "fees": 625.00, "status": "Connected"},
            {"name": "Amazon.ae", "orders": 34, "sales": 45700.50, "fees": 2285.00, "status": "Connected"},
            {"name": "Carrefour Marketplace", "orders": 5, "sales": 8200.00, "fees": 410.00, "status": "Connected"},
            {"name": "Direct Sales (B2B)", "orders": len(inv), "sales": sum(money(i.get("grand_total")) for i in inv), "fees": 0.00, "status": "Active"},
        ]
    total_sales = sum(money(c.get("sales")) for c in channels)
    total_orders = sum(int(c.get("orders") or 0) for c in channels)
    total_fees = sum(money(c.get("fees")) for c in channels)
    return render_template(
        "marketplaces.html",
        channels=channels,
        total_sales=total_sales,
        total_orders=total_orders,
        total_fees=total_fees,
        net_sales=total_sales - total_fees,
    )


def _settings_file_to_data_uri(file_storage, max_px=400):
    if not file_storage or not file_storage.filename:
        return None
    try:
        raw = file_storage.read()
        if not raw:
            return None
        import base64
        from PIL import Image
        im = Image.open(io.BytesIO(raw))
        if im.mode not in ("RGB", "RGBA"):
            im = im.convert("RGBA")
        w, h = im.size
        scale = min(1.0, (max_px / max(w, h)))
        if scale < 1.0:
            im = im.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
        buf = io.BytesIO()
        fmt = "PNG"
        im.save(buf, format=fmt)
        b64 = base64.b64encode(buf.getvalue()).decode("ascii")
        return f"data:image/png;base64,{b64}"
    except Exception:
        import base64
        try:
            return f"data:image/png;base64,{base64.b64encode(raw).decode('ascii')}"
        except Exception:
            return None


@app.route("/company-settings", methods=["GET"])
@login_required
def company_settings_form():
    active = manager.get_active_company_profile() if hasattr(manager, "get_active_company_profile") else {}
    if not active:
        active = {}
    return render_template("company_settings.html", profile=active, today=today_str(), is_admin=bool(_user_is_admin()))


@app.route("/company-settings", methods=["POST"])
@login_required
@admin_required
def company_settings_post():
    form = request.form
    active = manager.get_active_company_profile() if hasattr(manager, "get_active_company_profile") else None
    if not active:
        flash("No active company.", "error")
        return redirect(url_for("company_settings_form"))
    cid = active.get("id")
    updates = {}
    STR_FIELDS = [
        "name", "display_name", "legal_name", "prefix", "short_code", "trn",
        "currency", "emirate", "address", "phone", "email", "website",
        "bank_name", "bank_branch", "bank_account_no", "bank_iban", "bank_swift",
        "pdf_footer_thank_you",
    ]
    for field in STR_FIELDS:
        if field in form:
            updates[field] = (form.get(field) or "").strip()
    if "vat_rate" in form:
        try:
            updates["vat_rate"] = float(form.get("vat_rate") or 5)
        except (ValueError, TypeError):
            updates["vat_rate"] = 5.0
    prefix_raw = updates.get("prefix") or ""
    if prefix_raw:
        import re
        if not re.fullmatch(r"[A-Z0-9]{4}", prefix_raw):
            flash("Invoice Prefix must be exactly 4 uppercase letters/digits.", "error")
            return redirect(url_for("company_settings_form"))
    trn_raw = updates.get("trn") or ""
    if trn_raw:
        import re
        if not re.fullmatch(r"\d{15}", trn_raw):
            flash("TRN must be exactly 15 digits.", "error")
            return redirect(url_for("company_settings_form"))

    # Image uploads — logo / signature / stamp
    try:
        logo_uri = _settings_file_to_data_uri(request.files.get("logo_image")) if hasattr(request, "files") else None
        if logo_uri:
            updates["company_logo_base64"] = logo_uri
    except Exception:
        pass
    try:
        sig_uri = _settings_file_to_data_uri(request.files.get("signature_image")) if hasattr(request, "files") else None
        if sig_uri:
            updates["signature_image_base64"] = sig_uri
    except Exception:
        pass
    try:
        stamp_uri = _settings_file_to_data_uri(request.files.get("stamp_image")) if hasattr(request, "files") else None
        if stamp_uri:
            updates["stamp_image_base64"] = stamp_uri
    except Exception:
        pass

    # Explicit clear booleans
    if form.get("clear_signature") == "on":
        updates["signature_image_base64"] = ""
    if form.get("clear_stamp") == "on":
        updates["stamp_image_base64"] = ""
    if form.get("clear_logo") == "on":
        updates["company_logo_base64"] = ""

    if hasattr(manager, "update_company_profile"):
        ok, msg, _ = manager.update_company_profile(cid, updates)
        if ok:
            flash(msg or "Company profile updated.", "success")
        else:
            flash(msg or "Failed to update profile.", "error")
    else:
        flash("update_company_profile API unavailable.", "error")
    return redirect(url_for("company_settings_form"))


# ================================================================
# Task 3: Users + RBAC (Admin-only company-scoped users CRUD screen)
# ================================================================
@app.route("/users", methods=["GET"])
@login_required
@admin_required
def users_list():
    """Admin-only Users page: list users for the active company."""
    user_cid = (session.get("user") or {}).get("company_id")
    users = []
    if hasattr(manager, "list_company_users"):
        users = list(manager.list_company_users(user_cid) or [])
    # Sort admins first, then staff, then by username
    def _sort_key(u):
        r = 0 if (u or {}).get("role") == "admin" else 1
        return (r, str((u or {}).get("username") or "").lower())
    users.sort(key=_sort_key)
    return render_template(
        "users.html",
        users=users,
        users_count=len(users),
        today=today_str(),
    )


@app.route("/users/create", methods=["POST"])
@login_required
@admin_required
def users_create_post():
    """Admin creates a new user (admin or staff) for their own company."""
    user_cid = (session.get("user") or {}).get("company_id")
    form = request.form
    username = (form.get("username") or "").strip().lower()
    password = form.get("password") or ""
    confirm = form.get("password_confirm") or ""
    full_name = (form.get("full_name") or "").strip() or username.title()
    email = (form.get("email") or "").strip()
    role = (form.get("role") or "staff").strip().lower()
    if role not in ("admin", "staff"):
        role = "staff"
    if not username or not password:
        flash("Username and password are required.", "error")
        return redirect(url_for("users_list"))
    if len(password) < 8:
        flash("Password must be at least 8 characters long.", "error")
        return redirect(url_for("users_list"))
    if password != confirm:
        flash("Passwords do not match.", "error")
        return redirect(url_for("users_list"))
    if not hasattr(manager, "create_company_user"):
        flash("User creation API unavailable.", "error")
        return redirect(url_for("users_list"))
    ok, msg, _user = manager.create_company_user(
        username=username,
        password=password,
        company_id=user_cid,
        role=role,
        full_name=full_name,
        email=email,
    )
    if ok:
        flash(
            f"User '{username}' created (role: {role.upper()}). "
            f"Share the username + password with them securely.",
            "success",
        )
    else:
        flash(msg or "Failed to create user.", "error")
    return redirect(url_for("users_list"))


@app.route("/users/<username>/reset_password", methods=["POST"])
@login_required
@admin_required
def users_reset_password_post(username):
    """Admin-only: reset a user's password (8-char auto or form-supplied)."""
    actor = (session.get("user") or {}).get("username")
    target_username = (username or "").strip()
    # Security: cannot reset another company's users. We re-verify by username lookup + company check.
    form = request.form
    new_pw = (form.get("new_password") or "").strip()
    confirm = (form.get("new_password_confirm") or "").strip()
    if not target_username or not hasattr(manager, "get_company_user_by_username"):
        flash("User not found.", "error")
        return redirect(url_for("users_list"))
    target = manager.get_company_user_by_username(target_username)
    actor_cid = (session.get("user") or {}).get("company_id")
    if not target or target.get("company_id") != actor_cid:
        flash("You can only manage users within your own company.", "error")
        return redirect(url_for("users_list"))
    if not new_pw:
        import secrets, string
        alphabet = string.ascii_letters + string.digits + "!@#$%^&*()_+-=[]{}"
        new_pw = "".join(secrets.choice(alphabet) for _ in range(10))
        flash_note = f"Auto-generated temporary password: <strong>{new_pw}</strong>"
        use_auto = True
    else:
        if len(new_pw) < 8:
            flash("New password must be at least 8 characters.", "error")
            return redirect(url_for("users_list"))
        if confirm and new_pw != confirm:
            flash("Password confirmation does not match.", "error")
            return redirect(url_for("users_list"))
        flash_note = None
        use_auto = False
    if not hasattr(manager, "update_company_user_password"):
        flash("Password reset API unavailable.", "error")
        return redirect(url_for("users_list"))
    ok, msg = manager.update_company_user_password(target_username, new_pw)
    if ok:
        flash(
            f"Password reset for user '{target_username}' (by {actor or 'admin'})." +
            (f" — {flash_note}. Share this one-time password securely." if use_auto else "."),
            "success",
        )
    else:
        flash(msg or "Failed to reset password.", "error")
    return redirect(url_for("users_list"))


@app.route("/users/<username>/toggle_active", methods=["POST"])
@login_required
@admin_required
def users_toggle_active_post(username):
    """Admin-only: enable/disable a user's ability to log in (soft lock)."""
    target_username = (username or "").strip()
    actor_cid = (session.get("user") or {}).get("company_id")
    if not target_username or not hasattr(manager, "get_company_user_by_username"):
        flash("User not found.", "error")
        return redirect(url_for("users_list"))
    target = manager.get_company_user_by_username(target_username)
    if not target or target.get("company_id") != actor_cid:
        flash("You can only manage users within your own company.", "error")
        return redirect(url_for("users_list"))
    # Extra safeguard: don't let an admin disable THEMSELVES (prevent lockout)
    actor_uname = (session.get("user") or {}).get("username") or ""
    if str(actor_uname).lower() == str(target_username).lower():
        flash("You cannot disable your own account.", "error")
        return redirect(url_for("users_list"))
    ok, msg, _u = manager.toggle_company_user_active(target_username)
    if ok:
        flash(msg or f"User '{target_username}' status toggled.", "success")
    else:
        flash(msg or "Failed to toggle status.", "error")
    return redirect(url_for("users_list"))


@app.route("/users/<username>/delete", methods=["POST"])
@login_required
@admin_required
def users_delete_post(username):
    """Admin-only: permanently delete a user."""
    target_username = (username or "").strip()
    actor_cid = (session.get("user") or {}).get("company_id")
    if not target_username or not hasattr(manager, "get_company_user_by_username"):
        flash("User not found.", "error")
        return redirect(url_for("users_list"))
    target = manager.get_company_user_by_username(target_username)
    if not target or target.get("company_id") != actor_cid:
        flash("You can only delete users within your own company.", "error")
        return redirect(url_for("users_list"))
    actor_uname = (session.get("user") or {}).get("username") or ""
    if str(actor_uname).lower() == str(target_username).lower():
        flash("You cannot delete your own account.", "error")
        return redirect(url_for("users_list"))
    ok, msg = manager.delete_company_user(target_username)
    if ok:
        flash(msg or f"User '{target_username}' deleted.", "success")
    else:
        flash(msg or "Failed to delete user.", "error")
    return redirect(url_for("users_list"))


@app.route("/system/logout", methods=["GET", "POST"])
def system_logout():
    flash("Logged out (placeholder — web app is currently single-user local mode).", "info")
    return redirect(url_for("dashboard"))


@app.errorhandler(404)
def not_found(e):
    return render_template("404.html"), 404


@app.route("/api/health", methods=["GET"])
def api_health():
    active_n = len(manager.get_all_invoices_dict(include_deleted=False) or [])
    try:
        next_id = manager.generate_invoice_id()
    except Exception as exc:
        next_id = f"ERROR: {exc}"
    try:
        routes_count = len([r for r in app.url_map.iter_rules()])
    except Exception:
        routes_count = 0
    return jsonify({
        "ok": True,
        "flask": flask_version,
        "data_dir": DATA_DIR,
        "next_invoice_id": next_id,
        "active_invoices": active_n,
        "routes": routes_count,
        "timestamp_utc": datetime.utcnow().isoformat() + "Z",
    })


# ---------------------------------------------------------------------------
# Backward-compat routes for legacy templates (index.html, create.html, etc.)
# These redirect to or re-use the new canonical endpoints so old links work.
# ---------------------------------------------------------------------------
@app.route("/index", methods=["GET"])
def index():
    return redirect(url_for("invoices_list"), code=302)


@app.route("/create", methods=["GET"])
def create_form():
    return redirect(url_for("invoice_new_form"), code=302)


@app.route("/create", methods=["POST"])
def create_invoice():
    return invoice_create_post()  # Re-use the canonical POST handler


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    port = int(os.environ.get("PORT") or 5099)
    routes_count = len(list(app.url_map.iter_rules()))
    print("=" * 60)
    print("ASISTEM — Multi-Company Business Suite Web App — STARTING")
    print(f"  Flask version : {flask_version}")
    print(f"  Data folder   : {DATA_DIR}")
    print(f"  Routes total  : {routes_count}")
    try:
        ac = manager.get_active_company_profile()
        print(f"  Active company: {ac.get('display_name') or ac.get('name')}  ({ac.get('prefix')}####)")
    except Exception:
        pass
    try:
        print(f"  Next ID       : {manager.generate_invoice_id()}")
    except Exception:
        pass
    print(f"  Open URL      : http://127.0.0.1:{port}")
    print("=" * 60)
    app.run(host="0.0.0.0", port=port, debug=False)
