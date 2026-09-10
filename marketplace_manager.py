import json
import os
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple


class MarketplaceManager:
    def __init__(self, data_folder: str):
        self.data_folder = Path(data_folder)
        self.data_file = self.data_folder / "marketplace_data.json"
        self.data = {
            "platforms": [],
            "consignments": [],
            "monthly_settlements": [],
            "orders": [],
            "products": [],
            "listings": [],
            "rules": []
        }
        self._load()

    def _load(self) -> None:
        if self.data_file.exists():
            try:
                with open(self.data_file, "r", encoding="utf-8") as f:
                    raw = json.load(f)
                if isinstance(raw, dict):
                    self.data.update(raw)
            except Exception:
                pass

    def _save(self) -> None:
        try:
            with open(self.data_file, "w", encoding="utf-8") as f:
                json.dump(self.data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def add_platform(
        self,
        name: str,
        seller_id: str,
        api_credentials: Optional[Dict[str, Any]] = None,
        commission_rate: float = 0.0,
        payment_terms_days: int = 30,
        return_policy_days: int = 14
    ) -> Dict[str, Any]:
        if commission_rate > 1:
            commission_rate = commission_rate / 100.0
        platform = {
            "id": f"PLAT-{len(self.data['platforms'])+1:03d}",
            "name": name,
            "seller_id": seller_id,
            "api_credentials": api_credentials or {},
            "consignment_terms": {
                "commission_rate": commission_rate,
                "payment_terms_days": payment_terms_days,
                "return_policy_days": return_policy_days
            },
            "created_at": datetime.now().isoformat()
        }
        self.data["platforms"].append(platform)
        self._save()
        return platform

    def list_platforms(self) -> List[Dict[str, Any]]:
        return list(self.data.get("platforms", []))

    def get_platform(self, platform_id: str) -> Optional[Dict[str, Any]]:
        for p in self.data.get("platforms", []):
            if p.get("id") == platform_id or p.get("name") == platform_id:
                return p
        return None

    def update_platform(
        self,
        platform_id: str,
        name: Optional[str] = None,
        seller_id: Optional[str] = None,
        commission_rate: Optional[float] = None,
        payment_terms_days: Optional[int] = None,
        return_policy_days: Optional[int] = None
    ) -> bool:
        platforms = self.data.get("platforms", [])
        for p in platforms:
            if p.get("id") == platform_id:
                if name is not None:
                    p["name"] = name
                if seller_id is not None:
                    p["seller_id"] = seller_id
                terms = p.setdefault("consignment_terms", {})
                if commission_rate is not None:
                    terms["commission_rate"] = commission_rate / 100.0 if commission_rate > 1 else commission_rate
                if payment_terms_days is not None:
                    terms["payment_terms_days"] = payment_terms_days
                if return_policy_days is not None:
                    terms["return_policy_days"] = return_policy_days
                p["updated_at"] = datetime.now().isoformat()
                self._save()
                return True
        return False

    def delete_platform(self, platform_id: str) -> bool:
        platforms = self.data.get("platforms", [])
        before = len(platforms)
        platforms[:] = [p for p in platforms if p.get("id") != platform_id]
        if len(platforms) < before:
            self.data["platforms"] = platforms
            self._save()
            return True
        return False

    def create_consignment(
        self,
        platform_id: str,
        items: List[Dict[str, Any]],
        consignment_date: Optional[str] = None
    ) -> Dict[str, Any]:
        seq = len(self.data.get("consignments", [])) + 1
        cid = f"CONS-{datetime.now().strftime('%Y%m')}-{seq:03d}"
        consignment = {
            "consignment_id": cid,
            "platform_id": platform_id,
            "date": consignment_date or datetime.now().strftime("%Y-%m-%d"),
            "items": items,
            "status": "active",
            "summary": {
                "total_quantity": sum(int(i.get("quantity", 0)) for i in items),
                "current_stock": sum(int(i.get("quantity", 0)) for i in items),
                "sold_quantity": 0,
                "returned_quantity": 0
            }
        }
        self.data.setdefault("consignments", []).append(consignment)
        self._save()
        return consignment

    def list_consignments(self, platform_id: Optional[str] = None) -> List[Dict[str, Any]]:
        consignments = self.data.get("consignments", [])
        if platform_id:
            consignments = [c for c in consignments if c.get("platform_id") == platform_id]
        return consignments

    def get_consignment(self, consignment_id: str) -> Optional[Dict[str, Any]]:
        for c in self.data.get("consignments", []):
            if c.get("consignment_id") == consignment_id:
                return c
        return None

    def set_consignment_invoice(self, consignment_id: str, invoice_number: str) -> bool:
        for c in self.data.get("consignments", []):
            if c.get("consignment_id") == consignment_id:
                c["invoice_number"] = invoice_number
                c.setdefault("meta", {})["linked_invoice_created_at"] = datetime.now().isoformat()
                self._save()
                return True
        return False

    def delete_consignment(self, consignment_id: str) -> bool:
        consignments = self.data.get("consignments", [])
        before = len(consignments)
        consignments[:] = [c for c in consignments if c.get("consignment_id") != consignment_id]
        if len(consignments) < before:
            self.data["consignments"] = consignments
            self._save()
            return True
        return False

    def update_consignment_stock(
        self,
        consignment_id: str,
        sold_delta: int = 0,
        returned_delta: int = 0
    ) -> bool:
        for c in self.data.get("consignments", []):
            if c.get("consignment_id") == consignment_id:
                summary = c.setdefault("summary", {})
                sold = int(summary.get("sold_quantity", 0)) + sold_delta
                returned = int(summary.get("returned_quantity", 0)) + returned_delta
                total = int(summary.get("total_quantity", 0))
                current = total - sold + returned
                summary["sold_quantity"] = sold
                summary["returned_quantity"] = returned
                summary["current_stock"] = current
                self._save()
                return True
        return False

    def record_monthly_settlement(
        self,
        platform_id: str,
        month: str,
        total_sales: float,
        platform_fees: float,
        refunds: float,
        promotions: float,
        payment_date: Optional[str] = None,
        payment_status: str = "pending",
        invoice_generated: bool = False,
        invoice_number: Optional[str] = None,
        allocation_invoice_number: Optional[str] = None,
        payment_gross: Optional[float] = None,
        commission_rate: Optional[float] = None
    ) -> Dict[str, Any]:
        net_settlement = total_sales - platform_fees - refunds - promotions
        settlement = {
            "platform_id": platform_id,
            "month": month,
            "total_sales": total_sales,
            "platform_fees": platform_fees,
            "refunds": refunds,
            "promotions": promotions,
            "net_settlement": net_settlement,
            "payment_date": payment_date,
            "payment_status": payment_status,
            "invoice_generated": invoice_generated,
            "invoice_number": invoice_number,
            "allocation_invoice_number": allocation_invoice_number,
            "payment_gross": payment_gross,
            "commission_rate": commission_rate
        }
        self.data.setdefault("monthly_settlements", []).append(settlement)
        self._save()
        return settlement

    def list_settlements(
        self,
        platform_id: Optional[str] = None,
        month: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        settlements = self.data.get("monthly_settlements", [])
        result = []
        for s in settlements:
            if platform_id and s.get("platform_id") != platform_id:
                continue
            if month and s.get("month") != month:
                continue
            result.append(s)
        return result

    def set_settlement_payment_date(
        self,
        platform_id: str,
        month: str,
        payment_date: Optional[str] = None,
        status: Optional[str] = None
    ) -> bool:
        settlements = self.data.get("monthly_settlements", [])
        changed = False
        for s in settlements:
            if s.get("platform_id") == platform_id and s.get("month") == month:
                s["payment_date"] = payment_date or datetime.now().strftime("%Y-%m-%d")
                if status:
                    s["payment_status"] = status
                changed = True
        if changed:
            self._save()
        return changed

    def generate_platform_summary(self, platform_id: str) -> Dict[str, Any]:
        consignments = self.list_consignments(platform_id)
        settlements = self.list_settlements(platform_id)
        total_sales = sum(float(s.get("total_sales", 0) or 0) for s in settlements)
        net_settlement = sum(float(s.get("net_settlement", 0) or 0) for s in settlements)
        open_settlements = [s for s in settlements if s.get("payment_status") != "paid"]
        summary = {
            "platform_id": platform_id,
            "consignment_count": len(consignments),
            "settlement_count": len(settlements),
            "total_sales": total_sales,
            "net_settlement": net_settlement,
            "open_settlements": len(open_settlements)
        }
        return summary

    def register_product(
        self,
        product_id: str,
        name: str,
        internal_code: Optional[str] = None,
        default_price: float = 0.0,
        category: Optional[str] = None
    ) -> Dict[str, Any]:
        products = self.data.setdefault("products", [])
        for p in products:
            if p.get("product_id") == product_id:
                p.update(
                    {
                        "name": name,
                        "internal_code": internal_code,
                        "default_price": default_price,
                        "category": category,
                        "updated_at": datetime.now().isoformat()
                    }
                )
                self._save()
                return p
        product = {
            "product_id": product_id,
            "name": name,
            "internal_code": internal_code,
            "default_price": default_price,
            "category": category,
            "created_at": datetime.now().isoformat()
        }
        products.append(product)
        self._save()
        return product

    def list_products(self) -> List[Dict[str, Any]]:
        return list(self.data.get("products", []))

    def map_product_to_platform(
        self,
        platform_id: str,
        product_id: str,
        platform_sku: str,
        platform_price: float,
        status: str = "active"
    ) -> Dict[str, Any]:
        listings = self.data.setdefault("listings", [])
        for l in listings:
            if l.get("platform_id") == platform_id and l.get("platform_sku") == platform_sku:
                prices = l.setdefault("prices", [])
                prices.append(
                    {
                        "price": platform_price,
                        "date": datetime.now().isoformat()
                    }
                )
                l["status"] = status
                l["updated_at"] = datetime.now().isoformat()
                self._save()
                return l
        listing = {
            "platform_id": platform_id,
            "product_id": product_id,
            "platform_sku": platform_sku,
            "status": status,
            "prices": [
                {
                    "price": platform_price,
                    "date": datetime.now().isoformat()
                }
            ],
            "created_at": datetime.now().isoformat()
        }
        listings.append(listing)
        self._save()
        return listing

    def list_listings(self, platform_id: Optional[str] = None) -> List[Dict[str, Any]]:
        listings = self.data.get("listings", [])
        if platform_id:
            listings = [l for l in listings if l.get("platform_id") == platform_id]
        return listings

    def set_listing_status(self, platform_id: str, platform_sku: str, status: str) -> bool:
        listings = self.data.get("listings", [])
        for l in listings:
            if l.get("platform_id") == platform_id and l.get("platform_sku") == platform_sku:
                l["status"] = status
                l["updated_at"] = datetime.now().isoformat()
                self._save()
                return True
        return False

    def set_platform_price(
        self,
        platform_id: str,
        platform_sku: str,
        price: float
    ) -> bool:
        listings = self.data.get("listings", [])
        for l in listings:
            if l.get("platform_id") == platform_id and l.get("platform_sku") == platform_sku:
                prices = l.setdefault("prices", [])
                prices.append(
                    {
                        "price": price,
                        "date": datetime.now().isoformat()
                    }
                )
                l["updated_at"] = datetime.now().isoformat()
                self._save()
                return True
        return False

    def record_order(
        self,
        platform_id: str,
        order_id: str,
        order_date: str,
        status: str,
        items: List[Dict[str, Any]],
        customer: Dict[str, Any],
        total_amount: float,
        fees: float = 0.0,
        refunds: float = 0.0
    ) -> Dict[str, Any]:
        orders = self.data.setdefault("orders", [])
        order = {
            "platform_id": platform_id,
            "order_id": order_id,
            "order_date": order_date,
            "status": status,
            "items": items,
            "customer": customer,
            "total_amount": total_amount,
            "fees": fees,
            "refunds": refunds,
            "created_at": datetime.now().isoformat()
        }
        orders.append(order)
        self._save()
        return order

    def delete_order(self, order_id: str) -> bool:
        orders = self.data.get("orders", [])
        before = len(orders)
        orders[:] = [o for o in orders if o.get("order_id") != order_id]
        if len(orders) < before:
            self.data["orders"] = orders
            self._save()
            return True
        return False

    def update_order_status(self, order_id: str, status: str) -> bool:
        orders = self.data.get("orders", [])
        for o in orders:
            if o.get("order_id") == order_id:
                o["status"] = status
                o["updated_at"] = datetime.now().isoformat()
                self._save()
                return True
        return False

    def list_orders(
        self,
        platform_id: Optional[str] = None,
        status: Optional[str] = None,
        date_from: Optional[str] = None,
        date_to: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        orders = self.data.get("orders", [])
        result = []
        for o in orders:
            if platform_id and o.get("platform_id") != platform_id:
                continue
            if status and o.get("status") != status:
                continue
            d = o.get("order_date")
            if date_from and d and d < date_from:
                continue
            if date_to and d and d > date_to:
                continue
            result.append(o)
        return result

    def record_return(
        self,
        order_id: str,
        amount: float,
        reason: Optional[str] = None
    ) -> bool:
        orders = self.data.get("orders", [])
        for o in orders:
            if o.get("order_id") == order_id:
                returns = o.setdefault("returns", [])
                returns.append(
                    {
                        "amount": amount,
                        "reason": reason,
                        "date": datetime.now().isoformat()
                    }
                )
                prev_refunds = float(o.get("refunds", 0) or 0)
                o["refunds"] = prev_refunds + amount
                self._save()
                return True
        return False

    def mark_settlement_invoice_generated(
        self,
        platform_id: str,
        month: str,
        invoice_number: str
    ) -> bool:
        settlements = self.data.get("monthly_settlements", [])
        for s in settlements:
            if s.get("platform_id") == platform_id and s.get("month") == month:
                s["invoice_generated"] = True
                s["invoice_number"] = invoice_number
                self._save()
                return True
        return False

    def mark_settlement_paid(
        self,
        platform_id: str,
        month: str,
        payment_date: Optional[str] = None
    ) -> bool:
        settlements = self.data.get("monthly_settlements", [])
        for s in settlements:
            if s.get("platform_id") == platform_id and s.get("month") == month:
                s["payment_status"] = "paid"
                s["payment_date"] = payment_date or datetime.now().strftime("%Y-%m-%d")
                self._save()
                return True
        return False

    def get_settlement(self, platform_id: str, month: str) -> Optional[Dict[str, Any]]:
        for s in self.data.get("monthly_settlements", []):
            if s.get("platform_id") == platform_id and s.get("month") == month:
                return s
        return None

    def get_platform_commission_rate(self, platform_id: str) -> float:
        p = self.get_platform(platform_id) or {}
        try:
            terms = p.get("consignment_terms") or {}
            rate = float(terms.get("commission_rate", 0) or 0)
            if rate > 1:
                rate = rate / 100.0
            return max(0.0, min(rate, 1.0))
        except Exception:
            return 0.0

    def ensure_settlement_invoice(
        self,
        invoice_manager: Any,
        platform_id: str,
        month: str
    ):
        s = self.get_settlement(platform_id, month)
        if not s:
            return None
        inv_no = s.get("invoice_number")
        if inv_no:
            inv = invoice_manager.get_invoice(inv_no)
            if inv:
                return inv
        p = self.get_platform(platform_id) or {}
        client_name = p.get("name") or platform_id
        try:
            terms_days = int((p.get("consignment_terms") or {}).get("payment_terms_days", 30))
        except Exception:
            terms_days = 30
        today = datetime.now()
        due = today + timedelta(days=terms_days)
        items = [
            {
                "description": f"Platform sales {month}",
                "quantity": 1,
                "unit_price": float(s.get("total_sales", 0) or 0),
                "total": float(s.get("total_sales", 0) or 0),
                "taxable": False
            },
            {
                "description": "Platform fees",
                "quantity": 1,
                "unit_price": -float(s.get("platform_fees", 0) or 0),
                "total": -float(s.get("platform_fees", 0) or 0),
                "taxable": False
            },
            {
                "description": "Refunds",
                "quantity": 1,
                "unit_price": -float(s.get("refunds", 0) or 0),
                "total": -float(s.get("refunds", 0) or 0),
                "taxable": False
            },
            {
                "description": "Promotions",
                "quantity": 1,
                "unit_price": -float(s.get("promotions", 0) or 0),
                "total": -float(s.get("promotions", 0) or 0),
                "taxable": False
            }
        ]
        subtotal = sum(float(i.get("total", 0) or 0) for i in items)
        inv_dict = {
            "invoice_id": invoice_manager.generate_invoice_id(),
            "invoice_type": "platform_settlement",
            "client_name": client_name,
            "client_trn": "",
            "date": today.strftime("%Y-%m-%d"),
            "due_date": due.strftime("%Y-%m-%d"),
            "payment_terms": f"{terms_days} Days",
            "payment_method": "Bank Transfer",
            "payment_due": f"{terms_days} Days",
            "tax_rate": 0,
            "items": items,
            "costs": [],
            "subtotal": subtotal,
            "taxable_amount": 0,
            "non_taxable_amount": subtotal,
            "tax_amount": 0,
            "grand_total": subtotal,
            "total_cost": 0,
            "profit_loss": subtotal,
            "total_paid": 0,
            "balance_due": subtotal,
            "status": "Not Paid",
            "notes": f"Settlement {platform_id} {month}",
            "currency": "AED",
            "payment_history": [],
            "workflow_status": "Under Process",
            "platform_id": platform_id,
            "month": month
        }
        ok = invoice_manager.add_invoice_from_dict(inv_dict)
        if not ok:
            return None
        self.mark_settlement_invoice_generated(platform_id, month, inv_dict["invoice_id"])
        return invoice_manager.get_invoice(inv_dict["invoice_id"])

    def apply_settlement_payment(
        self,
        invoice_manager: Any,
        platform_id: str,
        month: str,
        amount: float,
        method: str,
        account: Optional[str] = None,
        notes: str = "",
        payment_date: Optional[str] = None
    ) -> Tuple[bool, str]:
        inv = self.ensure_settlement_invoice(invoice_manager, platform_id, month)
        if not inv:
            return False, "Settlement not found"
        date_str = payment_date or datetime.now().strftime("%Y-%m-%d")
        ok, msg = inv.add_payment(amount, date_str, method, notes, account or method)
        if not ok:
            return False, msg
        ok2 = invoice_manager.update_invoice(inv)
        if not ok2:
            return False, "Failed to update invoice"
        try:
            acc_name = account or method
            invoice_manager.process_invoice_payment_with_balance(inv.to_dict(), float(amount), acc_name)
        except Exception:
            pass
        if inv.status.lower() == "paid":
            self.mark_settlement_paid(platform_id, month, date_str)
        else:
            self.set_settlement_payment_date(platform_id, month, date_str, status="partial")
        return True, "Payment applied"

    def settlement_statement(
        self,
        invoice_manager: Any,
        platform_id: str,
        month: str
    ) -> Dict[str, Any]:
        inv = self.ensure_settlement_invoice(invoice_manager, platform_id, month)
        if not inv:
            return {}
        data = inv.to_dict()
        payments = data.get("payment_history", [])
        total_paid = float(data.get("total_paid", 0) or 0)
        grand_total = float(data.get("grand_total", 0) or 0)
        balance = grand_total - total_paid
        try:
            due_date = datetime.strptime(data.get("due_date") or data.get("date"), "%Y-%m-%d")
        except Exception:
            due_date = datetime.now()
        aging_days = (datetime.now() - due_date).days
        return {
            "platform_id": platform_id,
            "month": month,
            "invoice_id": data.get("invoice_id"),
            "original_amount": grand_total,
            "cumulative_payments": total_paid,
            "outstanding_balance": balance,
            "aging_days": aging_days,
            "payment_history": payments
        }

    def daily_sales_report(self, date: str) -> Dict[str, Any]:
        orders = self.list_orders(date_from=date, date_to=date)
        total_sales = sum(float(o.get("total_amount", 0) or 0) for o in orders)
        total_fees = sum(float(o.get("fees", 0) or 0) for o in orders)
        total_refunds = sum(float(o.get("refunds", 0) or 0) for o in orders)
        return {
            "date": date,
            "order_count": len(orders),
            "total_sales": total_sales,
            "total_fees": total_fees,
            "total_refunds": total_refunds,
            "net_sales": total_sales - total_fees - total_refunds
        }

    def platform_performance_report(self, month: Optional[str] = None) -> List[Dict[str, Any]]:
        result = []
        for p in self.list_platforms():
            pid = p.get("id")
            settlements = self.list_settlements(platform_id=pid, month=month)
            total_sales = sum(float(s.get("total_sales", 0) or 0) for s in settlements)
            net_settlement = sum(float(s.get("net_settlement", 0) or 0) for s in settlements)
            result.append(
                {
                    "platform_id": pid,
                    "platform_name": p.get("name"),
                    "total_sales": total_sales,
                    "net_settlement": net_settlement,
                    "settlement_count": len(settlements)
                }
            )
        return result

    def product_performance_report(
        self,
        platform_id: Optional[str] = None,
        month: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        orders = self.list_orders(platform_id=platform_id)
        data: Dict[str, Dict[str, Any]] = {}
        for o in orders:
            d = o.get("order_date", "")
            if month and not d.startswith(month):
                continue
            for item in o.get("items", []):
                pid = item.get("product_id")
                qty = int(item.get("quantity", 0))
                price = float(item.get("unit_price", 0) or 0)
                key = pid or ""
                rec = data.setdefault(
                    key,
                    {
                        "product_id": pid,
                        "quantity_sold": 0,
                        "sales": 0.0
                    }
                )
                rec["quantity_sold"] += qty
                rec["sales"] += qty * price
        return list(data.values())

    def settlement_status_report(self) -> List[Dict[str, Any]]:
        settlements = self.data.get("monthly_settlements", [])
        result: Dict[str, Dict[str, Any]] = {}
        for s in settlements:
            key = s.get("payment_status", "unknown")
            rec = result.setdefault(
                key,
                {
                    "status": key,
                    "count": 0,
                    "total_net_settlement": 0.0
                }
            )
            rec["count"] += 1
            rec["total_net_settlement"] += float(s.get("net_settlement", 0) or 0)
        return list(result.values())

    def delete_settlement(self, platform_id: str, month: str) -> bool:
        settlements = self.data.get("monthly_settlements", [])
        before = len(settlements)
        settlements[:] = [s for s in settlements if not (s.get("platform_id") == platform_id and s.get("month") == month)]
        if len(settlements) < before:
            self.data["monthly_settlements"] = settlements
            self._save()
            return True
        return False

    def inventory_turnover_report(self, platform_id: Optional[str] = None) -> List[Dict[str, Any]]:
        consignments = self.list_consignments(platform_id=platform_id)
        result = []
        for c in consignments:
            s = c.get("summary", {})
            total = int(s.get("total_quantity", 0))
            sold = int(s.get("sold_quantity", 0))
            current = int(s.get("current_stock", 0))
            result.append(
                {
                    "consignment_id": c.get("consignment_id"),
                    "platform_id": c.get("platform_id"),
                    "total_quantity": total,
                    "sold_quantity": sold,
                    "current_stock": current
                }
            )
        return result

    def return_analysis_report(
        self,
        platform_id: Optional[str] = None,
        month: Optional[str] = None
    ) -> List[Dict[str, Any]]:
        orders = self.list_orders(platform_id=platform_id)
        reasons: Dict[str, Dict[str, Any]] = {}
        for o in orders:
            d = o.get("order_date", "")
            if month and not d.startswith(month):
                continue
            for r in o.get("returns", []):
                reason = r.get("reason") or "unspecified"
                rec = reasons.setdefault(
                    reason,
                    {
                        "reason": reason,
                        "count": 0,
                        "amount": 0.0
                    }
                )
                rec["count"] += 1
                rec["amount"] += float(r.get("amount", 0) or 0)
        return list(reasons.values())

    def add_rule(self, rule_type: str, config: Dict[str, Any]) -> Dict[str, Any]:
        rules = self.data.setdefault("rules", [])
        rule = {
            "id": f"RULE-{len(rules)+1:03d}",
            "type": rule_type,
            "config": config,
            "created_at": datetime.now().isoformat()
        }
        rules.append(rule)
        self._save()
        return rule

    def list_rules(self) -> List[Dict[str, Any]]:
        return list(self.data.get("rules", []))
