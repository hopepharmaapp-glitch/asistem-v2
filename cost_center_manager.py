"""
CostCenterManager
=================

Manages reusable "Cost Centers" for HopePharma invoices.

A Cost Center is a named collection of rules that automatically apply
costs (fixed or percentage-based) to invoices, keyed off the item/service
names on each invoice line.  Once a Cost Center is set up, the user can
attach it to any invoice (new or historical) and the system will:

  * generate one `costs[]` entry per matched rule,
  * optionally post a GL double-entry (Dr expense 5xxx, Cr asset/liability
    based on the "account" selected on the rule / cost line),
  * allow bulk backfill of ALL historical invoices that match a filter and
    have empty/missing costs, with a preview + confirm step.

Stored in ``<data_folder>/cost_centers.json``.
"""

from __future__ import annotations

import json
import os
import re
from copy import deepcopy
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple


def _q(value: Any) -> Decimal:
    try:
        if value is None:
            return Decimal("0")
        if isinstance(value, Decimal):
            return value
        return Decimal(str(value))
    except (InvalidOperation, ValueError, TypeError):
        return Decimal("0")


def _aed(value: Any) -> Decimal:
    q = Decimal("0.01")
    try:
        return _q(value).quantize(q, rounding=ROUND_HALF_UP)
    except Exception:
        return Decimal("0.00").quantize(q, rounding=ROUND_HALF_UP)


def _norm(text: Any) -> str:
    if text is None:
        return ""
    s = str(text).strip().lower()
    s = re.sub(r"[^a-z0-9\u0600-\u06ff\s]", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ---------------------------------------------------------------------------
# CostCenterManager
# ---------------------------------------------------------------------------

class CostCenterManager:
    """Persisted store of cost centers + engine to apply rules to invoices."""

    FILE = "cost_centers.json"

    EXPENSE_OPTIONS = [
        ("5000  COGS (Cost of Goods Sold)",                           "5000"),
        ("5100  Salary / Wages",                                      "5100"),
        ("5200  Rent, Utilities (Electricity, Water, Internet)",      "5200"),
        ("5300  Office Supplies, Travel, Marketing, Other Admin",     "5300"),
        ("5400  Professional Fees, Audit, Legal",                     "5400"),
        ("5500  Delivery & Shipping / Logistics",                     "5500"),
        ("5600  Medical / Clinical / Device Service Fees",            "5600"),
        ("5700  Commissions Paid to Sales Agents",                    "5700"),
        ("5800  Warranty / Returns / After-Sales",                    "5800"),
        ("5900  Other Operating Expenses",                            "5900"),
    ]

    DEFAULT_ACCOUNT = "5000"

    def __init__(self, data_folder: str):
        self.data_folder = str(data_folder)
        self.path = os.path.join(self.data_folder, self.FILE)
        self._ensure()
        self.centers: List[Dict[str, Any]] = self._load()

    # ------------------------------------------------------------------ IO --

    def _ensure(self) -> None:
        if not os.path.exists(self.path):
            try:
                with open(self.path, "w", encoding="utf-8") as f:
                    json.dump([], f, indent=2)
            except Exception:
                pass

    def _load(self) -> List[Dict[str, Any]]:
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception:
            pass
        return []

    def save(self) -> None:
        def _default(o):
            if isinstance(o, Decimal):
                return float(o)
            if hasattr(o, "isoformat"):
                return str(o)
            raise TypeError(f"Unserializable {type(o).__name__}")
        tmp = self.path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.centers, f, indent=2, ensure_ascii=False, default=_default)
            os.replace(tmp, self.path)
        except Exception as exc:
            raise RuntimeError(f"Failed to save cost centers: {exc}") from exc

    # ---------------------------------------------------------------- CRUD --

    def list_centers(self) -> List[Dict[str, Any]]:
        return deepcopy(self.centers)

    def get_center(self, center_id: str) -> Optional[Dict[str, Any]]:
        for c in self.centers:
            if c.get("center_id") == center_id:
                return deepcopy(c)
        return None

    def get_center_by_name(self, name: str) -> Optional[Dict[str, Any]]:
        key = _norm(name)
        for c in self.centers:
            if _norm(c.get("name")) == key and key:
                return deepcopy(c)
        return None

    def _next_id(self) -> str:
        used = set()
        for c in self.centers:
            cid = c.get("center_id", "")
            m = re.search(r"CC-(\d+)", cid or "")
            if m:
                used.add(int(m.group(1)))
        n = 1
        while n in used:
            n += 1
        return f"CC-{n:03d}"

    def create_center(self, name: str, description: str = "",
                      default_expense_account: str = DEFAULT_ACCOUNT,
                      default_fund_account: str = "ADCB",
                      rules: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]:
        if not (name or "").strip():
            raise ValueError("Cost Center name is required")
        if self.get_center_by_name(name):
            raise ValueError(f"Cost Center with name '{name}' already exists")
        center = {
            "center_id": self._next_id(),
            "name": name.strip(),
            "description": (description or "").strip(),
            "default_expense_account": default_expense_account or self.DEFAULT_ACCOUNT,
            "default_fund_account": default_fund_account or "ADCB",
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "rules": self._sanitize_rules(rules or [], default_expense_account, default_fund_account),
        }
        self.centers.append(center)
        self.save()
        return deepcopy(center)

    def update_center(self, center_id: str, *, name: Optional[str] = None,
                      description: Optional[str] = None,
                      default_expense_account: Optional[str] = None,
                      default_fund_account: Optional[str] = None,
                      rules: Optional[List[Dict[str, Any]]] = None) -> Optional[Dict[str, Any]]:
        for idx, c in enumerate(self.centers):
            if c.get("center_id") != center_id:
                continue
            if name is not None:
                nm = name.strip()
                if not nm:
                    raise ValueError("Cost Center name cannot be empty")
                existing = self.get_center_by_name(nm)
                if existing and existing.get("center_id") != center_id:
                    raise ValueError(f"Another Cost Center already uses name '{nm}'")
                c["name"] = nm
            if description is not None:
                c["description"] = description.strip()
            if default_expense_account is not None:
                c["default_expense_account"] = default_expense_account or self.DEFAULT_ACCOUNT
            if default_fund_account is not None:
                c["default_fund_account"] = default_fund_account or "ADCB"
            if rules is not None:
                c["rules"] = self._sanitize_rules(
                    rules,
                    c.get("default_expense_account", self.DEFAULT_ACCOUNT),
                    c.get("default_fund_account", "ADCB"),
                )
            c["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self.centers[idx] = c
            self.save()
            return deepcopy(c)
        return None

    def delete_center(self, center_id: str) -> bool:
        for i, c in enumerate(self.centers):
            if c.get("center_id") == center_id:
                del self.centers[i]
                self.save()
                return True
        return False

    # --------------------------------------------------------------- Rules --

    def _sanitize_rules(self, rules: Iterable[Dict[str, Any]],
                        default_expense_account: str,
                        default_fund_account: str = "ADCB") -> List[Dict[str, Any]]:
        clean: List[Dict[str, Any]] = []
        default_expense_account = default_expense_account or self.DEFAULT_ACCOUNT
        default_fund_account = default_fund_account or "ADCB"
        for i, r in enumerate(rules or []):
            if not isinstance(r, dict):
                continue
            name = (r.get("name") or r.get("cost_name") or "").strip()
            if not name:
                name = f"Rule {i + 1}"
            match_mode = str(r.get("match_mode", "contains")).lower()
            if match_mode not in {"contains", "exact", "regex", "starts_with"}:
                match_mode = "contains"
            patterns_raw = r.get("patterns") or r.get("keywords") or []
            if isinstance(patterns_raw, str):
                patterns_raw = [patterns_raw]
            patterns = [str(p).strip() for p in patterns_raw if str(p).strip()]

            calc_type = str(r.get("calc_type", "percentage")).lower()
            if calc_type not in {"percentage", "fixed_per_line", "fixed_per_invoice"}:
                calc_type = "percentage"

            value_raw = r.get("value") or r.get("amount") or 0
            try:
                value = Decimal(str(value_raw))
            except Exception:
                value = Decimal("0")

            account = (r.get("expense_account") or r.get("account_code") or
                       default_expense_account).strip()
            # Allow either "5000  COGS ..." or raw "5000"
            if " " in account:
                head = account.split()[0]
                if re.fullmatch(r"\d{4}", head):
                    account = head

            fund_account = (r.get("fund_account") or r.get("account") or
                            r.get("payment_account") or default_fund_account).strip()
            if not fund_account:
                fund_account = "ADCB"

            rule_id = str(r.get("rule_id") or f"R{i + 1:03d}-{abs(hash(name)) % 1000:03d}")
            clean.append({
                "rule_id": rule_id,
                "name": name,
                "match_mode": match_mode,
                "patterns": patterns,
                "calc_type": calc_type,
                "value": _aed(value) if calc_type.startswith("fixed") else value,
                "expense_account": account,
                "fund_account": fund_account,
                "apply_to_taxable_only": bool(r.get("apply_to_taxable_only", False)),
                "apply_to_non_taxable_only": bool(r.get("apply_to_non_taxable_only", False)),
                "notes": (r.get("notes") or "").strip(),
            })
        return clean

    # ------------------------------------------------------------ Matching --

    def _line_matches(self, rule: Dict[str, Any], line_desc: Any) -> bool:
        if not line_desc:
            return False
        desc = _norm(line_desc)
        if not desc:
            return False
        patterns = rule.get("patterns") or []
        if not patterns:
            # No patterns -> matches EVERY line (catch-all for fixed per invoice defaults)
            return False
        mode = rule.get("match_mode", "contains")
        for p in patterns:
            if not p:
                continue
            pn = _norm(p)
            if not pn:
                continue
            if mode == "contains":
                if pn in desc:
                    return True
            elif mode == "exact":
                if pn == desc:
                    return True
            elif mode == "starts_with":
                if desc.startswith(pn):
                    return True
            elif mode == "regex":
                try:
                    if re.search(p, str(line_desc), flags=re.IGNORECASE):
                        return True
                except Exception:
                    if pn in desc:
                        return True
        return False

    # ------------------------------------------------------------- Apply --

    def compute_costs_for_invoice(self, center: Dict[str, Any],
                                  invoice_dict: Dict[str, Any]) -> Tuple[List[Dict[str, Any]], Decimal, List[str]]:
        """
        Apply a Cost Center's rules against an invoice.

        Returns (new_costs, total_cost, info_lines) where new_costs can be
        inserted directly into ``invoice_dict["costs"]`` after user confirm.
        """
        center = center or {}
        rules = center.get("rules") or []
        items = invoice_dict.get("items") or invoice_dict.get("line_items") or []
        inv_date = invoice_dict.get("date") or invoice_dict.get("invoice_date") or datetime.now().strftime("%Y-%m-%d")

        per_rule_total: Dict[str, Decimal] = {}
        per_rule_lines: Dict[str, int] = {}
        per_rule_first_fund: Dict[str, str] = {}
        per_rule_first_account: Dict[str, str] = {}
        per_rule_first_name: Dict[str, str] = {}

        matched_rule_ids = set()

        for line in items:
            if not isinstance(line, dict):
                continue
            desc = line.get("description") or line.get("name") or line.get("product") or ""
            qty = _q(line.get("quantity") or line.get("qty") or 0)
            unit_price = _q(line.get("unit_price") or line.get("price") or line.get("rate") or 0)
            line_total = _aed(qty * unit_price)
            if line_total <= 0:
                t2 = _aed(line.get("total") or line.get("amount") or 0)
                if t2 > 0:
                    line_total = t2
            taxable = bool(line.get("taxable", line.get("vat") if isinstance(line.get("vat"), bool) else (
                str(line.get("taxable") or line.get("vat") or "No").strip().lower() in {"yes", "true", "1", "y", "vat"}
            )))

            for rule in rules:
                taxable_only = bool(rule.get("apply_to_taxable_only", False))
                non_taxable_only = bool(rule.get("apply_to_non_taxable_only", False))
                if taxable_only and not taxable:
                    continue
                if non_taxable_only and taxable:
                    continue

                if not self._line_matches(rule, desc):
                    continue

                rid = rule["rule_id"]
                matched_rule_ids.add(rid)
                calc_type = rule.get("calc_type", "percentage")
                value = _q(rule.get("value", 0))

                if calc_type == "percentage":
                    # value treated as fraction (0.65 = 65%, 0.2 = 20%)
                    portion = _aed(line_total * value)
                elif calc_type == "fixed_per_line":
                    portion = _aed(value)
                elif calc_type == "fixed_per_invoice":
                    portion = Decimal("0")  # handled once below outside line loop
                else:
                    portion = Decimal("0")

                per_rule_total[rid] = per_rule_total.get(rid, Decimal("0")) + portion
                per_rule_lines[rid] = per_rule_lines.get(rid, 0) + 1
                per_rule_first_fund.setdefault(rid, rule.get("fund_account") or "ADCB")
                per_rule_first_account.setdefault(rid, rule.get("expense_account") or center.get("default_expense_account", self.DEFAULT_ACCOUNT))
                per_rule_first_name.setdefault(rid, rule.get("name") or rid)

        # Fixed-per-invoice rules: apply once IF any line matched (or, if no
        # patterns set, always apply as an invoice-level overhead)
        for rule in rules:
            if rule.get("calc_type") != "fixed_per_invoice":
                continue
            rid = rule["rule_id"]
            patterns = rule.get("patterns") or []
            should_apply = (rid in matched_rule_ids) or (len(patterns) == 0)
            if not should_apply:
                continue
            val = _aed(rule.get("value", 0))
            per_rule_total[rid] = per_rule_total.get(rid, Decimal("0")) + val
            per_rule_first_fund.setdefault(rid, rule.get("fund_account") or "ADCB")
            per_rule_first_account.setdefault(rid, rule.get("expense_account") or center.get("default_expense_account", self.DEFAULT_ACCOUNT))
            per_rule_first_name.setdefault(rid, rule.get("name") or rid)

        new_costs: List[Dict[str, Any]] = []
        total = Decimal("0.00")
        info: List[str] = []

        for rid, amt in per_rule_total.items():
            amt = _aed(amt)
            if amt <= 0:
                continue
            rule = next((r for r in rules if r.get("rule_id") == rid), None) or {}
            cost = {
                "description": per_rule_first_name.get(rid, rid),
                "amount": float(amt),
                "account": per_rule_first_fund.get(rid, "ADCB"),
                "expense_account": per_rule_first_account.get(rid, self.DEFAULT_ACCOUNT),
                "rule_id": rid,
                "center_id": center.get("center_id"),
                "center_name": center.get("name"),
                "date": inv_date,
                "matched_lines": per_rule_lines.get(rid, 0),
                "calc_type": rule.get("calc_type", "percentage"),
                "receipt_attached": False,
                "notes": rule.get("notes", ""),
            }
            new_costs.append(cost)
            total = _aed(total + amt)
            info.append(
                f"• {cost['description']}: AED {amt:,.2f} "
                f"({cost['calc_type']}; matched {cost['matched_lines']} line(s))"
            )

        return new_costs, total, info

    # ----------------------------------------------------------- Backfill --

    def backfill_preview(self, center_id: str, invoices: Iterable[Dict[str, Any]],
                         *, replace_existing_costs: bool = False,
                         min_date: Optional[str] = None,
                         max_date: Optional[str] = None,
                         invoice_ids: Optional[Iterable[str]] = None) -> Dict[str, Any]:
        """
        Dry-run a backfill: returns summary + per-invoice previews WITHOUT
        touching the invoices.  Use :meth:`apply_backfill` after confirm.
        """
        center = self.get_center(center_id)
        if not center:
            raise ValueError(f"Cost Center {center_id} not found")

        id_filter = None
        if invoice_ids:
            id_filter = {str(x) for x in invoice_ids if x}

        affected: List[Dict[str, Any]] = []
        total_all = Decimal("0.00")
        total_costs_added = 0
        skipped_no_lines = 0
        skipped_has_costs = 0
        skipped_dates = 0

        for inv in invoices:
            if not isinstance(inv, dict):
                continue
            inv_id = str(inv.get("invoice_id") or inv.get("id") or inv.get("invoice_no") or "")
            if not inv_id:
                continue
            if id_filter is not None and inv_id not in id_filter:
                continue
            d = inv.get("date") or inv.get("invoice_date") or ""
            if min_date and d < min_date:
                skipped_dates += 1
                continue
            if max_date and d > max_date:
                skipped_dates += 1
                continue
            items = inv.get("items") or inv.get("line_items") or []
            if not items:
                skipped_no_lines += 1
                continue
            existing = inv.get("costs") or []
            if existing and not replace_existing_costs:
                skipped_has_costs += 1
                continue

            costs, tot, info = self.compute_costs_for_invoice(center, inv)
            if not costs:
                skipped_no_lines += 1
                continue
            existing_total = Decimal("0.00")
            for ec in existing:
                existing_total = _aed(existing_total + (ec.get("amount") or 0))
            affected.append({
                "invoice_id": inv_id,
                "client": inv.get("client_name") or inv.get("customer") or "",
                "date": d,
                "grand_total": float(_aed(inv.get("grand_total") or inv.get("total") or 0)),
                "prev_cost_total": float(existing_total),
                "new_cost_total": float(tot),
                "new_costs": costs,
                "replace_existing": bool(replace_existing_costs),
                "num_costs": len(costs),
                "info_lines": info,
            })
            total_all = _aed(total_all + tot)
            total_costs_added += len(costs)

        affected.sort(key=lambda r: (r.get("date") or "", r.get("invoice_id") or ""))
        return {
            "center": center,
            "replace_existing_costs": bool(replace_existing_costs),
            "min_date": min_date,
            "max_date": max_date,
            "invoices_count": len(affected),
            "total_cost_lines": total_costs_added,
            "total_cost_amount": float(total_all),
            "skipped_no_lines": skipped_no_lines,
            "skipped_has_costs": skipped_has_costs,
            "skipped_outside_dates": skipped_dates,
            "preview_rows": affected,
        }

    def apply_backfill(self, preview: Dict[str, Any], invoices_store: Any,
                       *, save_method_name: str = "save_invoices",
                       update_method_name: str = "update_invoice") -> Dict[str, Any]:
        """
        Actually commit the backfill preview to the invoices list.

        ``invoices_store`` must expose either:
          * ``update_invoice(invoice_id, partial_dict)`` and ``load_invoices()``, OR
          * ``invoices`` attribute (a list) + ``save_invoices()``.
        """
        rows = preview.get("preview_rows") or []
        replace_existing = bool(preview.get("replace_existing_costs", False))
        applied = 0
        errors: List[Dict[str, Any]] = []

        # Strategy A: store has update_invoice() + load_invoices()
        update = getattr(invoices_store, update_method_name, None) if update_method_name else None
        load = getattr(invoices_store, "load_invoices", None)
        invoices_list = getattr(invoices_store, "invoices", None)
        save_inv = getattr(invoices_store, save_method_name, None) if save_method_name else None

        for r in rows:
            inv_id = r.get("invoice_id")
            if not inv_id:
                continue
            try:
                new_costs = r.get("new_costs") or []
                # Convert Decimal-heavy dicts → plain cost dicts for storage
                store_costs: List[Dict[str, Any]] = []
                for c in new_costs:
                    store_costs.append({
                        "description": c.get("description", ""),
                        "amount": float(c.get("amount") or 0),
                        "account": c.get("account") or c.get("fund_account") or "ADCB",
                        "expense_account": c.get("expense_account") or self.DEFAULT_ACCOUNT,
                        "rule_id": c.get("rule_id"),
                        "center_id": c.get("center_id"),
                        "center_name": c.get("center_name"),
                        "matched_lines": c.get("matched_lines", 0),
                        "calc_type": c.get("calc_type", "percentage"),
                        "receipt_attached": bool(c.get("receipt_attached", False)),
                        "receipt_filename": c.get("receipt_filename"),
                        "notes": c.get("notes", ""),
                    })
                if replace_existing:
                    costs_final = store_costs
                else:
                    # Merge with any existing (find inv to read them)
                    existing: List[Dict[str, Any]] = []
                    found = None
                    if load:
                        try:
                            all_inv = load() or []
                            for x in all_inv:
                                if str(x.get("invoice_id") or x.get("id") or "") == str(inv_id):
                                    found = x
                                    break
                        except Exception:
                            found = None
                    elif isinstance(invoices_list, list):
                        for x in invoices_list:
                            if str(x.get("invoice_id") or x.get("id") or "") == str(inv_id):
                                found = x
                                break
                    if found:
                        existing = list(found.get("costs") or [])
                    costs_final = list(existing) + store_costs

                new_total_cost = sum(_aed(c.get("amount") or 0) for c in costs_final)
                new_total_cost_f = float(new_total_cost)

                if update:
                    partial = {"costs": costs_final, "total_cost": new_total_cost_f}
                    update(inv_id, partial)
                elif isinstance(invoices_list, list) and save_inv:
                    for idx, x in enumerate(invoices_list):
                        if str(x.get("invoice_id") or x.get("id") or "") == str(inv_id):
                            invoices_list[idx]["costs"] = costs_final
                            invoices_list[idx]["total_cost"] = new_total_cost_f
                            gt = _aed(x.get("grand_total") or x.get("total") or 0)
                            invoices_list[idx]["profit_loss"] = float(_aed(gt - new_total_cost))
                            break
                    save_inv()
                else:
                    raise RuntimeError("Invoices store has no usable update method.")
                applied += 1
            except Exception as exc:
                errors.append({"invoice_id": inv_id, "error": str(exc)})
        return {
            "applied": applied,
            "errors": errors,
            "total": len(rows),
        }
