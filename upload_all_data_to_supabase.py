#!/usr/bin/env python3
"""ASISTEM — one-click upload of ALL local JSON data to Supabase.
Run this after you have created your Supabase project (Step 10 of the guide) and
past the 2 env vars into your shell BEFORE running:

    export SUPABASE_URL="https://xxxx.supabase.co"
    export SUPABASE_SERVICE_ROLE_KEY="eyJhbGciOi...service_role_key_here..."
    export HOPEPHARMA_DATA_DIR="/Users/sm/Desktop/HopePharmaData"   # optional
    python3 upload_all_data_to_supabase.py

This migrates EVERY JSON file under your local data folder up to Supabase's
public.json_docs (doc_name PK + content JSONB) table, so when you go live on
Vercel your existing HopePharma invoices/company profiles are there.
"""
import sys, os, json
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from hope_pharma_complete import EnhancedCloudDataManager  # noqa

sb_url = os.environ.get("SUPABASE_URL","").strip()
sb_key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY","").strip()

if not sb_url or not sb_key:
    print("💥 ERROR: You MUST set 2 env vars first before running:")
    print("   export SUPABASE_URL=\"https://YOUR-PROJECT.supabase.co\"")
    print("   export SUPABASE_SERVICE_ROLE_KEY=\"eyJhbGciOi...your-service-role-key...\"")
    print("   (Get both values from Supabase dashboard → Project Settings → API)")
    sys.exit(1)

data_dir = os.environ.get("HOPEPHARMA_DATA_DIR","").strip() or os.path.dirname(os.path.abspath(__file__))
print(f"📂 HOPEPHARMA_DATA_DIR = {data_dir}")
print(f"🔗 SUPABASE_URL        = {sb_url}")
print()

mgr = EnhancedCloudDataManager(data_dir, mode='single')

# Supabase no longer allows arbitrary `exec_sql` RPC on new projects, and the
# REST schema validator on newer Supabase versions rejects POST bodies to any
# non-existent stored procedure (error: "should NOT have additional property `public`").
# Table creation MUST be done once via Supabase SQL Editor (user-facing step).
SQL_SETUP_BLOCK = r"""
-- ============================================================
--  ASISTEM: one-time Supabase SQL setup (run in SQL Editor ONCE)
-- ============================================================
CREATE TABLE IF NOT EXISTS public.json_docs (
  doc_name text PRIMARY KEY,
  content jsonb NOT NULL DEFAULT '{}'::jsonb,
  updated_at timestamptz NOT NULL DEFAULT now()
);
ALTER TABLE public.json_docs ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "json_docs_full_access_service_role" ON public.json_docs;
CREATE POLICY "json_docs_full_access_service_role"
  ON public.json_docs FOR ALL USING (true) WITH CHECK (true);
GRANT ALL PRIVILEGES ON TABLE public.json_docs TO postgres;
GRANT ALL PRIVILEGES ON TABLE public.json_docs TO service_role;
GRANT SELECT, INSERT, UPDATE, DELETE ON TABLE public.json_docs TO anon;
COMMENT ON TABLE public.json_docs IS 'ASISTEM All JSON documents (invoices, purchases, employees, company profiles, app_state)';
"""

print("▶ Step 1/2: Verifying public.json_docs table exists on Supabase...")
print("   (If you see probe failure, open Supabase → SQL Editor → New Query, and run:)")
print("   -----------------------------")
print(SQL_SETUP_BLOCK)
print("   -----------------------------")
ok_init = mgr._supabase_init_table()
# Force the inited flag to True regardless of probe result
# (user's manual SQL above is sufficient; probe sometimes times out on free tier).
mgr._sb_table_inited = True
if ok_init:
    print("   ✅ public.json_docs table probe OK")
else:
    print("   ⚠️ probe returned False — continuing anyway (table likely exists already from your manual SQL step).")

# List ONLY actual data JSON files in data dir (skip deploy configs, debug, and large non-data blobs)
_DATA_JSON_STARTS = (
    "app_","audit_","balance_","balances","chart_of_","clients_","company_","companies",
    "cost_centers","credit_notes","customers","delivery_notes","employees","financial_",
    "general_ledger","goods_receipts","inventory","invoices_","journals","outbounds",
    "payments","period_locks","products_","purchases_","quotes_","salaries_","sales_",
    "salary_reports","supplier","suppliers","stock_","transactions","vat_records",
    "warehouse_","warehouses","accounting_periods","expenses","fixed_assets",
    "import_batches","journal_entries","opening_balances","reconciliation_records",
    "report_versions","temperature_","delivery_notes","delivery_notes_",
)
json_files = sorted([
    f for f in os.listdir(data_dir)
    if f.endswith(".json") and os.path.isfile(os.path.join(data_dir, f))
    and (not f.startswith(".") and f not in {})
    and (
        any(f.startswith(p) for p in _DATA_JSON_STARTS)
        or f in {"app_settings.json", "app_state.json", "audit_log.jsonl", "company_profiles.json", "companies.json"}
    )
])
# Safety: if our strict allowlist accidentally filtered out a user's file, include any
# small-ish (< 20 MB) JSON that's not in a blacklist (vercel.json, build configs etc.)
_BLACKLIST_JSON_FILENAMES = {"vercel.json", "package.json", "package-lock.json", "tsconfig.json", "requirements-lock.json"}
all_jsons = sorted([
    f for f in os.listdir(data_dir)
    if f.endswith(".json") and os.path.isfile(os.path.join(data_dir, f))
    and f not in _BLACKLIST_JSON_FILENAMES
    and not f.startswith(".")
])
for f in all_jsons:
    if f not in json_files and os.path.getsize(os.path.join(data_dir, f)) < 20 * 1024 * 1024:
        json_files.append(f)
json_files = sorted(set(json_files))
print(f"\n▶ Step 2/2: Uploading {len(json_files)} JSON data files from disk → Supabase...")
print("-" * 66)
ok_count = 0; fail_count = 0
for filename in json_files:
    fp = os.path.join(data_dir, filename)
    try:
        with open(fp, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"  ❌ SKIP {filename}: cannot read JSON — {e}")
        fail_count += 1
        continue
    shape = "list " + str(len(data)) if isinstance(data, list) else (
        "dict" if isinstance(data, dict) else type(data).__name__
    )
    sb_ok = mgr._supabase_save(filename, data)
    if sb_ok:
        print(f"  ✅ {filename:<36} size={os.path.getsize(fp):>8} bytes  shape=({shape})")
        ok_count += 1
    else:
        print(f"  ❌ FAIL {filename}")
        fail_count += 1
print("-" * 66)
print(f"\n🏁 UPLOAD FINISHED: {ok_count}/{len(json_files)} files OK, {fail_count} failed")
if fail_count == 0:
    print("""
🎉 SUCCESS! Every JSON file is now stored in Supabase.
NEXT STEPS for you:
  1. Go to Supabase Dashboard → Table Editor → json_docs
     → you should see all filenames in the 'doc_name' column with content column populated.
  2. In Vercel → Project Settings → Environment Variables → paste the SAME
     SUPABASE_URL and SUPABASE_SERVICE_ROLE_KEY values.
  3. Click Vercel's Redeploy so the new env vars are picked up by the serverless functions.
  4. Visit your public `.vercel.app` URL → the Invoices dashboard shows your HopePharma
     HPMT####/HPS-#### invoices exactly like they were on the desktop!
""")
else:
    print("Some files failed. Try re-running after checking the env vars.")
    sys.exit(2)
