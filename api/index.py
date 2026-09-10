"""Vercel serverless entrypoint — 2-PHASE BOOT LOADER.

PHASE 1 (always, immediately): Create a minimal "debug-shell" Flask app with
just /api/health and /api/debug routes. These never fail and answer in <5ms.
If you see the generic Vercel FUNCTION_INVOCATION_FAILED, it means Python
never even reached THIS file — check vercel.json / requirements.txt install.

PHASE 2 (on first non-debug request): try to import the real ASISTEM Flask app
from web_app.py. If it works — forward all future requests to the real app.
If it fails — render the full Python traceback IN THE BROWSER (not a generic
page) and also dump every env var (hiding secrets) so the user can paste a
diagnostic.

This two-phase approach is what guarantees "something always renders" on Vercel
Free Tier even if there's a bug during the heavy EnhancedCloudDataManager boot.
"""
import sys
import os
import traceback
import platform
from pathlib import Path
from datetime import datetime

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Always import Flask first — MUST succeed. If THIS fails, requirements.txt is broken.
from flask import Flask, jsonify, request, Response  # noqa: E402

# ---------------------------------------------------------------------------
# ⚠️ CRITICAL (Vercel static analyzer)
# Vercel CLI >=59 does Python static AST scanning of api/index.py BEFORE any code
# runs. It requires an EXPLICIT top-level binding: `app = Flask(...)` at module
# scope, visible within the first ~30 lines. Assigning `app = some_other_var`
# later (even at the bottom of the file) is MISSED and triggers:
#
#   Error: Found api/index.py but it does not define a top-level "app" Flask
#          instance. but found potential entrypoints: web_app.py (variable: app)
#
# To fix, we declare the Flask app HERE (line ~33) with the standard
# `app = Flask(__name__)` constructor form that Vercel's scanner looks for.
# Later code just mutates `app` (adds routes, swaps .wsgi_app, ...) — the
# module-level binding `app` never changes identity so the analyzer stays happy.
# ---------------------------------------------------------------------------
app = Flask(__name__)
app.config["JSON_SORT_KEYS"] = False

# ---------------------------------------------------------------------------
# PHASE 1: Debug shell routes (always registered, always work).
# ---------------------------------------------------------------------------
_BOOT = {
    "started_at": datetime.utcnow().isoformat() + "Z",
    "real_app_loaded": False,
    "real_app_error": None,
    "phase1_ok": True,
}


def _safe_env():
    """Return list of env vars (values hidden for secret-looking keys)."""
    out = {}
    for k, v in sorted(os.environ.items()):
        secret = any(x in k.lower() for x in ("key", "secret", "token", "password"))
        if k.lower() in {"path", "pythonpath", "home", "lang"}:
            out[k] = v
        elif secret:
            out[k] = "*" * 8 + f" ({len(v)} chars, set={bool(v)})"
        else:
            out[k] = v if len(v) < 120 else v[:117] + "..."
    return out


@app.route("/api/health")
def shell_health():
    return jsonify({
        "ok": True,
        "service": "asistem",
        "phase1": "ok",
        "real_app_loaded": _BOOT["real_app_loaded"],
        "real_app_error": bool(_BOOT["real_app_error"]),
        "now": datetime.utcnow().isoformat() + "Z",
        "python": platform.python_version(),
        "platform": platform.platform(),
    })


@app.route("/api/debug")
def shell_debug():
    payload = {
        "boot": _BOOT,
        "root": str(ROOT),
        "cwd": os.getcwd(),
        "dir_root_exists": ROOT.exists(),
        "files_root_exists": os.path.exists(str(ROOT / "web_app.py")),
        "files_api_exists": os.path.exists(str(ROOT / "api" / "index.py")),
        "python_version": sys.version,
        "argv": sys.argv,
        "sys_path": [p if len(p) < 120 else p[:117] + "..." for p in sys.path[:30]],
        "env": _safe_env(),
        "requirements_txt": (ROOT / "requirements.txt").read_text() if (ROOT / "requirements.txt").exists() else "(NOT FOUND)",
    }
    if _BOOT["real_app_error"]:
        payload["real_app_error_traceback"] = _BOOT["real_app_error"]["traceback"]
        payload["real_app_error_type"] = _BOOT["real_app_error"]["type"]
        payload["real_app_error_message"] = _BOOT["real_app_error"]["message"]
    return jsonify(payload)


# ---------------------------------------------------------------------------
# PHASE 2: Lazy-load the real ASISTEM app.
# ---------------------------------------------------------------------------
_REAL_APP = None  # type: ignore[var-annotated]


def _try_load_real_app():
    """Attempt to import web_app.app. Returns the app on success, None if failed."""
    global _REAL_APP
    if _REAL_APP is not None:
        return _REAL_APP
    if _BOOT["real_app_loaded"]:
        return _REAL_APP  # even if None (retry not worth it for serverless cold start)
    print(f"[boot phase2] attempting to import web_app from web_app.py ...")
    print(f"[boot phase2] HOPEPHARMA_PRODUCTION={os.environ.get('HOPEPHARMA_PRODUCTION')!r}")
    print(f"[boot phase2] HOPEPHARMA_DATA_DIR={os.environ.get('HOPEPHARMA_DATA_DIR')!r}")
    print(f"[boot phase2] SUPABASE_URL set={bool(os.environ.get('SUPABASE_URL'))}")
    print(f"[boot phase2] SUPABASE_SERVICE_ROLE_KEY set={bool(os.environ.get('SUPABASE_SERVICE_ROLE_KEY'))}")
    try:
        from web_app import app as _app  # noqa: E402
        _REAL_APP = _app
        _BOOT["real_app_loaded"] = True
        _BOOT["real_app_error"] = None
        print("[boot phase2] ✅ web_app imported OK — serving with real app.")
        return _REAL_APP
    except Exception as _e:
        tb_str = traceback.format_exc()
        print("=" * 80, file=sys.stderr)
        print("ASISTEM PHASE2 FAILURE: could not load web_app.app", file=sys.stderr)
        print(tb_str, file=sys.stderr)
        print("=" * 80, file=sys.stderr)
        _BOOT["real_app_loaded"] = False
        _BOOT["real_app_error"] = {
            "type": type(_e).__name__,
            "message": str(_e),
            "traceback": tb_str,
        }
        return None


def _make_error_html():
    err = _BOOT["real_app_error"] or {"type": "Unknown", "message": "n/a", "traceback": ""}
    safe_tb = err["traceback"].replace("<", "&lt;").replace(">", "&gt;")[:12000]
    safe_msg = err["message"].replace("<", "&lt;").replace(">", "&gt;")
    env_list = "".join(
        f"<li><code>{k}</code> = <code>{str(v).replace('<','&lt;')}</code></li>"
        for k, v in list(_safe_env().items())[:60]
    )
    return f"""<!doctype html><html><head><meta charset="utf-8">
<title>ASISTEM — Boot Failure Diagnostic</title>
<style>
body {{ font-family: -apple-system, Segoe UI, Roboto, sans-serif; margin:0; background:#f8fafc; color:#0f172a; padding:28px; }}
.card {{ background:#fff; border:1px solid #e2e8f0; border-radius:14px; padding:22px 26px; max-width:1000px; margin:0 auto 18px; box-shadow:0 1px 2px rgba(15,23,42,.04); }}
h1 {{ margin:0 0 6px; font-size:22px; color:#b91c1c; }}
h2 {{ margin:0 0 12px; font-size:16px; color:#1e40af; }}
pre {{ background:#fef2f2; border:1px solid #fecaca; padding:14px; border-radius:10px; overflow:auto; font-size:12.5px; line-height:1.55; }}
code {{ background:#f1f5f9; padding:2px 6px; border-radius:5px; font-size:12.5px; }}
ul.env {{ column-count:2; column-gap:24px; font-size:12px; line-height:1.7; }}
.footer {{ max-width:1000px; margin:0 auto; color:#64748b; font-size:12px; text-align:center; }}
.btn {{ display:inline-block; background:#1e40af; color:#fff; padding:10px 16px; border-radius:10px; text-decoration:none; font-weight:600; margin-top:10px; }}
.btn.s {{ background:#065f46; }}
</style></head><body>
<div class="card">
  <h1>⚠️ ASISTEM failed to start (Phase 2 boot)</h1>
  <p>The <em>minimal web server shell</em> is running (<strong>Phase 1 is OK</strong> — so routing, Flask, Vercel basics are all working).
  The heavy data layer (<code>web_app.py</code> → <code>EnhancedCloudDataManager</code>) crashed while importing. Here's why:</p>
</div>
<div class="card">
  <h2>🐛 Python Error</h2>
  <p><strong>Type:</strong> <code>{err['type']}</code></p>
  <p><strong>Message:</strong> <code>{safe_msg}</code></p>
  <h2 style="margin-top:22px;">📜 Full Traceback</h2>
  <pre>{safe_tb if safe_tb else '(empty — click the /api/debug JSON link for raw details)'}</pre>
  <p><a class="btn" href="/api/debug">View raw diagnostic JSON (/api/debug)</a>
     <a class="btn s" href="/api/health">Check /api/health</a></p>
</div>
<div class="card">
  <h2>🔧 Environment Variables (secrets masked)</h2>
  <ul class="env">{env_list}</ul>
</div>
<div class="footer">
  <p>ASISTEM bootloader · Python {platform.python_version()} · {platform.platform()}</p>
  <p>If you can, send this ENTIRE page screenshot to whoever manages your deployment.</p>
</div>
</body></html>"""


# ---------------------------------------------------------------------------
# WSGI-layer dispatcher. All request routing happens here so shell routes
# (/api/health, /api/debug) always answer even if phase 2 failed to load.
# ---------------------------------------------------------------------------
_ORIGINAL_SHELL_WSGI = app.wsgi_app


def _smart_wsgi_app(environ, start_response):
    path = environ.get("PATH_INFO", "/")
    # Debug shell routes always work (no heavy imports needed)
    if path in ("/api/health", "/api/debug"):
        return _ORIGINAL_SHELL_WSGI(environ, start_response)
    # All other paths → attempt phase2 lazy boot of real ASISTEM app
    real_app = _try_load_real_app()
    if real_app is not None:
        # Booted OK → delegate entirely to the real Flask app's WSGI handler
        return real_app.wsgi_app(environ, start_response)
    # Phase 2 failed: render our friendly diagnostic HTML error page instead of
    # a generic Vercel 500. /api/debug + /api/health are still reachable.
    resp = Response(_make_error_html(), status=500,
                    mimetype="text/html; charset=utf-8")
    return resp(environ, start_response)


app.wsgi_app = _smart_wsgi_app
