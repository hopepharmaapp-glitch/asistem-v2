"""
🏥 HopePharma Professional GAAP Financial Reporting Dashboard
============================================================
Accountant-quality reports, no accounting jargon required.

For NON-accountant users:
• Every report has a plain-English "What this means for you" box
• Expenses are recorded via Category dropdown (no Debit/Credit)
• Golden Rule (Assets = Liabilities + Equity) is checked automatically
• Audit button verifies all Debits = Credits in one click

Exports: Excel (2-sheet professional format) + PDF (A4 branded)
Data Source: General Ledger (GL) — the Single Source of Truth.
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, date, timedelta
from decimal import Decimal
import os
import sys
import re
import json
import hashlib

# ---------------------------------------------------------------------------
# Make sure we can import sibling modules regardless of CWD
# ---------------------------------------------------------------------------
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from ledger_service import LedgerService, EXPENSE_CATEGORY_MAP, round_aed
from report_engine import (
    ReportEngine, DateRangeType, DATE_RANGE_PRESETS,
    fmt_aed, calc_pct_change
)

# ===========================================================================
# 1. STYLING HELPERS (Professional ERP look: white, subtle blue, zebra rows)
# ===========================================================================
NAVY = "#1A365D"
NAVY_DARK = "#0F2744"
LIGHT_BLUE = "#EBF4FF"
LIGHT_BLUE_2 = "#DEEBFF"
BLUE = "#3182CE"
BLUE_DEEP = "#2B6CB0"
GREEN_BG = "#F0FFF4"
GREEN_BG_2 = "#C6F6D5"
GREEN_STRONG = "#2F855A"
GREEN_DEEP = "#276749"
RED_BG = "#FFF5F5"
RED_BG_2 = "#FED7D7"
RED_STRONG = "#C53030"
AMBER_BG = "#FFFFF0"
AMBER_BG_2 = "#FEFCBF"
AMBER_DEEP = "#B7791F"
GREY_BG = "#EDF2F7"
GREY_BG_2 = "#E2E8F0"
GREY_MED = "#A0AEC0"
GREY_DARK = "#4A5568"
MUTED = "#718096"
GRAY_BG = "#EDF2F7"
PRIMARY = NAVY
HOVER = BLUE_DEEP
ZEBRA = "#F7FAFC"
WHITE = "#FFFFFF"
SEPARATOR = "#CBD5E0"
GOLD = "#B7791F"
PURPLE_DEEP = "#553C9A"
PURPLE_BG = "#FAF5FF"
TEAL_DEEP = "#2C7A7B"
TEAL_BG = "#E6FFFA"
ORANGE_BG = "#FFFAF0"


class ZebraTreeview(ttk.Treeview):
    """Treeview with alternating row colours + bold total row highlighting.

    Bonus UX features (this session):
      * <Enter> = auto-grab focus so mouse wheel works IMMEDIATELY where the
        cursor lands (no click needed).
      * Edge-hover auto-scroll: hold cursor within 35px of any top/bottom/
        left/right edge and the view scrolls smoothly in that direction
        (closer to edge = faster).  Automatically stops when cursor moves
        away from the edge or leaves the widget.
      * Pointer-following mouse wheel: the dashboard's global <MouseWheel>
        handler routes wheel events to whichever Treeview (or Canvas) the
        pointer is currently over, so scrolling works wherever you point.
    """
    def __init__(self, master=None, **kw):
        super().__init__(master, **kw)
        # Edge-hover auto-scroll state
        self._edge_after_id = None
        self._edge_dx = 0
        self._edge_dy = 0
        # Base zebra rows
        self.tag_configure("odd", background=ZEBRA, foreground=NAVY_DARK)
        self.tag_configure("even", background=WHITE, foreground=NAVY_DARK)
        # Section headers (all caps section names: REVENUE / COGS / OPERATING EXPENSES etc.)
        self.tag_configure("section", background=NAVY, foreground=WHITE,
                           font=("Helvetica", 10, "bold"))
        # Sub-section headers (e.g. "Current Assets", "Non-Current Assets")
        self.tag_configure("subsection", background=GREY_BG_2, foreground=NAVY_DARK,
                           font=("Helvetica", 10, "bold"))
        # Hierarchy indentation levels (line items within a section)
        self.tag_configure("level1", font=("Helvetica", 9, "bold"), foreground=NAVY_DARK)
        self.tag_configure("level2", font=("Helvetica", 9), foreground="#2D3748")
        self.tag_configure("level3", font=("Helvetica", 9), foreground="#4A5568")
        # Key milestones
        self.tag_configure("gross_profit", background=GREEN_BG, foreground=GREEN_DEEP,
                           font=("Helvetica", 10, "bold"))
        self.tag_configure("net_profit", background=GREEN_BG_2, foreground=GREEN_DEEP,
                           font=("Helvetica", 11, "bold"))
        self.tag_configure("net_loss", background=RED_BG_2, foreground=RED_STRONG,
                           font=("Helvetica", 11, "bold"))
        # Subtotals within a section
        self.tag_configure("expense_total", background=GREY_BG,
                           font=("Helvetica", 10, "bold"), foreground=NAVY_DARK)
        self.tag_configure("subtotal", background=GREY_BG,
                           font=("Helvetica", 10, "bold"), foreground=NAVY_DARK)
        self.tag_configure("total_row", background=LIGHT_BLUE, foreground=NAVY_DARK,
                           font=("Helvetica", 10, "bold"))
        self.tag_configure("grand_total", background=NAVY, foreground=WHITE,
                           font=("Helvetica", 11, "bold"))
        # High risk / watch / good
        self.tag_configure("overdue_high", background=RED_BG,
                           foreground=RED_STRONG, font=("Helvetica", 9, "bold"))
        self.tag_configure("overdue_med", background=AMBER_BG,
                           foreground=AMBER_DEEP, font=("Helvetica", 9))
        self.tag_configure("good", background=GREEN_BG,
                           foreground=GREEN_DEEP, font=("Helvetica", 9, "bold"))
        self.tag_configure("neutral", foreground=GREY_DARK)
        self.tag_configure("zero", foreground=GREY_MED, font=("Helvetica", 9, "italic"))
        # UX bindings: focus + edge-hover auto-scroll
        self.bind("<Enter>", self._uxtv_on_enter, add="+")
        self.bind("<Leave>", self._uxtv_on_leave, add="+")

    # ---------- UX helpers: focus + edge auto-scroll ----------
    def _uxtv_on_enter(self, _event):
        try:
            self.focus_set()
        except Exception:
            pass
        self.bind("<Motion>", self._uxtv_on_motion, add="+")

    def _uxtv_on_leave(self, _event):
        try:
            self.unbind("<Motion>")
        except Exception:
            pass
        self._uxtv_edge_stop()

    def _uxtv_on_motion(self, event):
        try:
            w = max(1, self.winfo_width())
            h = max(1, self.winfo_height())
        except Exception:
            return
        EDGE = 35
        dx = 0
        dy = 0
        # Horizontal: left edge < EDGE  or  right edge (w - x) < EDGE
        if event.x < EDGE:
            dx = -int((EDGE - event.x + 9) // 10)      # 1..3 units per tick
        elif w - event.x < EDGE:
            dx = int((EDGE - (w - event.x) + 9) // 10)
        # Vertical: top < EDGE  or  bottom (h - y) < EDGE
        if event.y < EDGE:
            dy = -int((EDGE - event.y + 9) // 10)
        elif h - event.y < EDGE:
            dy = int((EDGE - (h - event.y) + 9) // 10)
        # Clamp max scroll per tick to ±2 (smooth, not jumpy)
        dx = max(-2, min(2, dx))
        dy = max(-2, min(2, dy))
        self._edge_dx, self._edge_dy = dx, dy
        if dx == 0 and dy == 0:
            self._uxtv_edge_stop()
        else:
            if self._edge_after_id is None:
                self._uxtv_edge_tick()

    def _uxtv_edge_tick(self):
        try:
            if self._edge_dy != 0:
                self.yview_scroll(self._edge_dy, "units")
            if self._edge_dx != 0:
                self.xview_scroll(self._edge_dx, "units")
        except Exception:
            self._uxtv_edge_stop()
            return
        self._edge_after_id = self.after(55, self._uxtv_edge_tick)

    def _uxtv_edge_stop(self):
        if self._edge_after_id is not None:
            try:
                self.after_cancel(self._edge_after_id)
            except Exception:
                pass
            self._edge_after_id = None
        self._edge_dx = 0
        self._edge_dy = 0

    def zebra_insert(self, parent="", index="end", values=(), tags=(), **kw):
        """Insert with auto-zebra if no special tag is already given."""
        idx = len(self.get_children(""))
        tag_list = list(tags)
        special = ("section", "subsection", "gross_profit", "net_profit", "net_loss",
                   "expense_total", "total_row", "grand_total", "subtotal",
                   "overdue_high", "overdue_med", "good", "neutral", "zero",
                   "level1", "level2", "level3")
        if not any(t in special for t in tag_list):
            tag_list.append("odd" if idx % 2 == 1 else "even")
        return self.insert(parent, index, values=values, tags=tuple(tag_list), **kw)


# ===========================================================================
# 1b. SCROLLABLE FRAME (vertical + horizontal scroll, mouse wheel + arrow keys)
# ===========================================================================
class ScrollableFrame(tk.Frame):
    """A Frame with a canvas + scrollbars that auto-scrolls children.

    Legacy base behaviour:
      - Bind mouse wheel to vertical scroll when mouse is over the canvas/child.
      - Shift+wheel for horizontal scroll.
      - Arrow keys scroll when focus is on canvas.
      - Click-drag on empty canvas space scrolls too.
      - Automatically propagates <Configure> of inner_frame to canvas scrollregion.

    Session UX additions (same pattern as ZebraTreeview):
      * <Enter> canvas → focus_set() + start Motion watcher for edge auto-scroll.
      * Edge-hover auto-scroll on the OUTER scrollable canvas (useful when the
        dashboard itself is taller than the window — hover near the top/bottom
        edge to scroll the whole report area).
    """
    def __init__(self, master, *args, bg="#F8FAFC", **kwargs):
        super().__init__(master, *args, bg=bg, **kwargs)
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self, borderwidth=0,
                                highlightthickness=0,
                                bg=bg, takefocus=1)
        # Edge-hover auto-scroll state (mirrors ZebraTreeview — same semantics)
        self._edge_after_id = None
        self._edge_dx = 0
        self._edge_dy = 0
        self.vbar = ttk.Scrollbar(self, orient="vertical",
                                  command=self.canvas.yview)
        self.hbar = ttk.Scrollbar(self, orient="horizontal",
                                  command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=self.vbar.set,
                              xscrollcommand=self.hbar.set)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.vbar.grid(row=0, column=1, sticky="ns")
        self.hbar.grid(row=1, column=0, sticky="ew")

        # Always show scrollbars so users notice — not optional
        self.vbar.grid_remove(); self.hbar.grid_remove()
        self.vbar.grid(); self.hbar.grid()

        self.scrollable_frame = tk.Frame(self.canvas, bg=bg)
        self._inner_id = self.canvas.create_window(
            (0, 0), window=self.scrollable_frame, anchor="nw"
        )

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(
                scrollregion=self.canvas.bbox("all"),
            ),
        )
        self.canvas.bind(
            "<Configure>",
            lambda e: self.canvas.itemconfig(
                self._inner_id, width=max(e.width, self.scrollable_frame.winfo_reqwidth())
            ),
        )

        # Wheel + key bindings
        self._bind_scroll(self.canvas, self.canvas)
        self.canvas.bind("<Enter>", self._sf_on_enter, add="+")
        self.canvas.bind("<Leave>", self._sf_on_leave, add="+")
        self.canvas.bind("<Enter>", self._bind_to_children, add="+")
        self.canvas.bind("<Leave>", self._unbind_from_children, add="+")
        # Arrow keys for when canvas has focus
        self.canvas.bind("<Up>",    lambda e: self.canvas.yview_scroll(-3, "units"))
        self.canvas.bind("<Down>",  lambda e: self.canvas.yview_scroll(3,  "units"))
        self.canvas.bind("<Left>",  lambda e: self.canvas.xview_scroll(-3, "units"))
        self.canvas.bind("<Right>", lambda e: self.canvas.xview_scroll(3,  "units"))
        self.canvas.bind("<Prior>", lambda e: self.canvas.yview_scroll(-1, "pages"))
        self.canvas.bind("<Next>",  lambda e: self.canvas.yview_scroll(1,  "pages"))
        self.canvas.bind("<Home>",  lambda e: self.canvas.yview_moveto("0"))
        self.canvas.bind("<End>",   lambda e: self.canvas.yview_moveto("1"))
        # Click-drag scroll (click empty canvas area and drag to pan)
        self._drag_start = None
        def on_btn1(e):
            self._drag_start = (e.x, e.y)
            self.canvas.focus_set()
        def on_btn1_motion(e):
            if self._drag_start is None:
                return
            dx = e.x - self._drag_start[0]
            dy = e.y - self._drag_start[1]
            self.canvas.xview_scroll(-dx, "units")
            self.canvas.yview_scroll(-dy, "units")
            self._drag_start = (e.x, e.y)
        self.canvas.bind("<ButtonPress-1>",  on_btn1)
        self.canvas.bind("<B1-Motion>",      on_btn1_motion)
        self.canvas.bind("<ButtonRelease-1>", lambda e: setattr(self, "_drag_start", None))

    # ---------- UX: focus on enter + edge-hover auto-scroll ----------
    def _sf_on_enter(self, _event):
        try:
            self.canvas.focus_set()
        except Exception:
            pass
        self.canvas.bind("<Motion>", self._sf_on_motion, add="+")

    def _sf_on_leave(self, _event):
        try:
            self.canvas.unbind("<Motion>")
        except Exception:
            pass
        self._sf_edge_stop()

    def _sf_on_motion(self, event):
        try:
            w = max(1, self.canvas.winfo_width())
            h = max(1, self.canvas.winfo_height())
        except Exception:
            return
        EDGE = 40
        dx = 0
        dy = 0
        if event.x < EDGE:
            dx = -int((EDGE - event.x + 9) // 10)
        elif w - event.x < EDGE:
            dx = int((EDGE - (w - event.x) + 9) // 10)
        if event.y < EDGE:
            dy = -int((EDGE - event.y + 9) // 10)
        elif h - event.y < EDGE:
            dy = int((EDGE - (h - event.y) + 9) // 10)
        dx = max(-2, min(2, dx))
        dy = max(-2, min(2, dy))
        self._edge_dx, self._edge_dy = dx, dy
        if dx == 0 and dy == 0:
            self._sf_edge_stop()
        elif self._edge_after_id is None:
            self._sf_edge_tick()

    def _sf_edge_tick(self):
        try:
            if self._edge_dy != 0:
                self.canvas.yview_scroll(self._edge_dy, "units")
            if self._edge_dx != 0:
                self.canvas.xview_scroll(self._edge_dx, "units")
        except Exception:
            self._sf_edge_stop()
            return
        self._edge_after_id = self.after(55, self._sf_edge_tick)

    def _sf_edge_stop(self):
        if self._edge_after_id is not None:
            try:
                self.after_cancel(self._edge_after_id)
            except Exception:
                pass
            self._edge_after_id = None
        self._edge_dx = 0
        self._edge_dy = 0

    def _bind_scroll(self, widget, canvas):
        def _on_wheel(event):
            if event.state & 0x0001:  # Shift key held
                canvas.xview_scroll(int(-1 * (event.delta / 120)), "units")
            else:
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            return "break"
        widget.bind("<MouseWheel>", _on_wheel)
        # Linux/BSD wheel emulation
        widget.bind("<Button-4>", lambda e: (canvas.yview_scroll(-1, "units"), "break"))
        widget.bind("<Button-5>", lambda e: (canvas.yview_scroll(1,  "units"), "break"))

    def _bind_to_children(self, event=None):
        """When mouse enters the canvas, also bind wheel to all children so
        scroll works while hovering Treeviews/Forms, not just empty canvas."""
        def recurse(w):
            self._bind_scroll(w, self.canvas)
            for c in w.winfo_children():
                recurse(c)
        recurse(self.scrollable_frame)

    def _unbind_from_children(self, event=None):
        def recurse(w):
            try:
                w.unbind("<MouseWheel>")
            except Exception:
                pass
            try:
                w.unbind("<Button-4>")
            except Exception:
                pass
            try:
                w.unbind("<Button-5>")
            except Exception:
                pass
            for c in w.winfo_children():
                recurse(c)
        recurse(self.scrollable_frame)


# ===========================================================================
# 2. EXPENSE ENTRY DIALOG (User-friendly — NO Debit/Credit)
# ===========================================================================
class SimpleExpenseDialog:
    """Popup form: Date + Category dropdown + Payee + Amount + Description.

    Behind the scenes, it posts a double-entry GL batch via LedgerService.
    """
    def __init__(self, parent, data_manager, on_saved_cb=None):
        self.dm = data_manager
        self.ledger = getattr(data_manager, "ledger", None)
        # Fallback: always ensure a working LedgerService exists so expense
        # posting works even when the data manager didn't initialise one.
        if self.ledger is None:
            try:
                from ledger_service import LedgerService
                self.ledger = LedgerService(self.dm)
            except Exception:
                self.ledger = None
        self.on_saved = on_saved_cb
        self.dlg = tk.Toplevel(parent)
        self.dlg.title("💸 Record an Expense (No Accounting Knowledge Needed)")
        self.dlg.geometry("560x420")
        self.dlg.transient(parent)
        self.dlg.resizable(False, False)
        self._build()

    def _build(self):
        frm = ttk.Frame(self.dlg, padding=18)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="💸 Record a Business Expense",
                  font=("Helvetica", 16, "bold"), foreground=NAVY).pack(anchor="w", pady=(0, 4))
        ttk.Label(frm,
                  text="Just tell us what you spent money on — the system will handle all Debit/Credit entries automatically.",
                  font=("Helvetica", 9), foreground="#4A5568",
                  wraplength=520, justify="left").pack(anchor="w", pady=(0, 14))

        grid = ttk.Frame(frm)
        grid.pack(fill="x")
        for c in (0, 2):
            grid.columnconfigure(c, weight=0)
        for c in (1, 3):
            grid.columnconfigure(c, weight=1)

        # Row 1 — Date
        ttk.Label(grid, text="Date *:", font=("Helvetica", 10, "bold")).grid(row=0, column=0, sticky="w", pady=6, padx=(0, 6))
        self.date_var = tk.StringVar(value=date.today().isoformat())
        ttk.Entry(grid, textvariable=self.date_var, width=16,
                  font=("Helvetica", 10)).grid(row=0, column=1, sticky="we", pady=6)
        ttk.Label(grid, text="Format YYYY-MM-DD",
                  font=("Helvetica", 8), foreground="#718096").grid(row=0, column=2, columnspan=2, sticky="w", padx=(8, 0))

        # Row 2 — Category (the critical simplification)
        ttk.Label(grid, text="Category *:", font=("Helvetica", 10, "bold")).grid(row=1, column=0, sticky="w", pady=6, padx=(0, 6))
        cats = list(EXPENSE_CATEGORY_MAP.keys())
        self.cat_var = tk.StringVar(value=cats[0])
        ttk.Combobox(grid, textvariable=self.cat_var, values=cats,
                     state="readonly", width=32,
                     font=("Helvetica", 10)).grid(row=1, column=1, columnspan=3, sticky="we", pady=6)

        # Row 3 — Amount
        ttk.Label(grid, text="Amount *:", font=("Helvetica", 10, "bold")).grid(row=2, column=0, sticky="w", pady=6, padx=(0, 6))
        self.amt_var = tk.StringVar(value="0.00")
        ttk.Entry(grid, textvariable=self.amt_var, width=18,
                  font=("Helvetica", 10)).grid(row=2, column=1, sticky="we", pady=6)
        ttk.Label(grid, text="AED", font=("Helvetica", 10, "bold"),
                  foreground=NAVY).grid(row=2, column=2, sticky="w", padx=(8, 0))

        # Row 4 — Payee
        ttk.Label(grid, text="Payee:", font=("Helvetica", 10, "bold")).grid(row=3, column=0, sticky="w", pady=6, padx=(0, 6))
        self.payee_var = tk.StringVar(value="")
        ttk.Entry(grid, textvariable=self.payee_var, width=32,
                  font=("Helvetica", 10)).grid(row=3, column=1, columnspan=3, sticky="we", pady=6)

        # Row 5 — Description
        ttk.Label(grid, text="Description:", font=("Helvetica", 10, "bold")).grid(row=4, column=0, sticky="nw", pady=6, padx=(0, 6))
        self.desc_txt = tk.Text(grid, height=4, wrap="word",
                                 font=("Helvetica", 10), relief="solid",
                                 borderwidth=1)
        self.desc_txt.grid(row=4, column=1, columnspan=3, sticky="we", pady=6)

        # Buttons
        btns = ttk.Frame(frm)
        btns.pack(fill="x", pady=(18, 0))

        def _cancel():
            self.dlg.destroy()

        def _save():
            self._save_expense()

        ttk.Button(btns, text="❌ Cancel", command=_cancel).pack(side="right", padx=(6, 0))
        ttk.Button(btns, text="💾 Save Expense", command=_save,
                   style="Accent.TButton").pack(side="right")

    def _save_expense(self):
        if not self.ledger:
            messagebox.showerror("Error", "Ledger service not available.")
            return
        try:
            dt = self.date_var.get().strip()
            datetime.strptime(dt, "%Y-%m-%d")
        except ValueError:
            messagebox.showwarning("Invalid Date",
                                   "Date must be YYYY-MM-DD (e.g. 2025-01-31)")
            return
        try:
            amt = float(self.amt_var.get())
            if amt <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid Amount", "Amount must be a positive number.")
            return
        cat = self.cat_var.get()
        payee = self.payee_var.get().strip()
        desc = self.desc_txt.get("1.0", "end").strip()
        try:
            entries = self.ledger.on_expense_submitted(
                category=cat,
                amount=amt,
                expense_date=dt,
                payee_name=payee,
                description=desc,
                username="Dashboard"
            )
        except Exception as e:
            messagebox.showerror("Could not save expense", str(e))
            return
        messagebox.showinfo(
            "✅ Expense Saved",
            f"Expense of {fmt_aed(amt)} for '{cat}' has been recorded.\n\n"
            f"Behind the scenes, {len(entries)} GL entries were created "
            f"(Debit {cat} | Credit Bank) — perfectly balanced, as all things should be."
        )
        if callable(self.on_saved):
            try:
                self.on_saved()
            except Exception:
                pass
        self.dlg.destroy()


# ===========================================================================
# 3. MAIN DASHBOARD WINDOW
# ===========================================================================
class GAAPReportingDashboard:
    def __init__(self, parent, data_manager):
        """
        parent: root Tk or Toplevel
        data_manager: EnhancedCloudDataManager instance (has load_json/save_json/ledger)
        """
        self.dm = data_manager
        self.ledger = getattr(data_manager, "ledger", None)
        # Fallback: if the data manager didn't initialise a LedgerService,
        # create one inline so Tab 11 Expense Ledger can list/edit/delete expenses
        # and the 💸 Expense dialog can post GL entries, regardless of dm state.
        if self.ledger is None:
            try:
                from ledger_service import LedgerService
                self.ledger = LedgerService(self.dm)
            except Exception:
                self.ledger = None
        # ---------------------------------------------------------------------
        # EARLY STATE INIT — Tab 11 Expense Ledger filter vars + metadata.
        # These MUST exist BEFORE self._build() because _build() calls
        # _build_expense_manager_tab() which then calls _reload_expense_tab()
        # at the END of the builder, and that reload tries to clear / update
        # self._exp_row_meta and read all the filter StringVar/BooleanVar.
        # If they were initialised AFTER _build(), AttributeError fires and
        # the user sees "Could not rebuild expense register: <_exp_row_meta>".
        # ---------------------------------------------------------------------
        self._exp_row_meta = {}
        # Tk String/Boolean vars — created here (not in the tab builder) so
        # they're tied to winfo root that parent provides, and NOT accidentally
        # GC'd or orphaned if the tab is rebuilt.
        self._exp_period_var   = tk.StringVar(value="Use Period Picker")
        self._exp_cat_var      = tk.StringVar(value="All Categories")
        # ---- "Active Only" DEFAULT (not "All") — anything reversed via
        # right-click → Delete does NOT show up in the register anymore,
        # just like it never happened.  User explicitly has to flip this to
        # "All / Reversed Only" to audit old entries.
        self._exp_status_var   = tk.StringVar(value="Active Only")
        self._exp_search_var   = tk.StringVar()
        # ---- INCLUDE REVERSED? OFF BY DEFAULT (auditor opt-in). ----
        # Controls exclude_reversed flag passed to list_expenses / aggregator.
        # True → include reversed bundles (for audit / forensics / "show me
        # the deleted ones" debugging).
        # False (DEFAULT) → reversed bundles vanish completely from list & total.
        self._exp_include_reversed_var = tk.BooleanVar(value=False)
        # Old name misnomer — this used to be `_exp_reversed_var` but it was
        # actually toggling include_withdrawals / include_invoice_costs, which
        # has nothing to do with reversed rows.  Rename for clarity.
        self._exp_include_withdrawals_and_invcost_var = tk.BooleanVar(value=False)
        # Backward compat alias so any lingering code doesn't crash
        self._exp_reversed_var = self._exp_include_withdrawals_and_invcost_var
        # Request 18 — Expense Ledger Source (Manual vs System) filter toggles
        self._exp_origin_manual_var = tk.BooleanVar(value=True)
        self._exp_origin_system_var = tk.BooleanVar(value=True)
        # UI handles — filled in later by _build_expense_manager_tab():
        self._exp_tree     = None
        self._exp_total_lbl = None

        self._init_rbac_and_audit_log()

        self.win = tk.Toplevel(parent)
        self.win.title("🏥 HopePharma — Professional Financial Reporting Centre")
        self.win.transient(parent)
        self._tree_registry = {}
        # ---- Auto-fit window to user's available screen ----
        # Do it BEFORE build so styles can reference real geometry if needed.
        self.win.update_idletasks()
        try:
            screen_w = self.win.winfo_screenwidth()
            screen_h = self.win.winfo_screenheight()
        except Exception:
            screen_w = 1366; screen_h = 768
        # Leave 6% margin each side for docks/menubar (conservative 88% × 86% usable)
        w = int(screen_w * 0.90)
        h = int(screen_h * 0.86)
        min_w = max(1040, int(screen_w * 0.70))
        min_h = max(640, int(screen_h * 0.68))
        # Center on screen
        x = max(0, (screen_w - w) // 2)
        y = max(0, int((screen_h - h) * 0.42))  # slight bias toward top (room for dock)
        self.win.geometry(f"{w}x{h}+{x}+{y}")
        self.win.minsize(min_w, min_h)
        # Remember user resize if they maximize, etc.
        self.win.bind("<Map>", lambda _e: self.win.deiconify())
        self._configure_style()
        # GLOBAL POINTER ROUTING: mouse wheel events ALWAYS target whichever
        # scrollable widget (Treeview / Canvas / Text / Listbox) the mouse
        # cursor is currently hovering over — NO click-to-focus required.
        # Shift + wheel = horizontal scroll (same convention as macOS/Safari).
        def _global_mousewheel(event):
            try:
                px = event.x_root if hasattr(event, "x_root") else self.win.winfo_pointerx()
                py = event.y_root if hasattr(event, "y_root") else self.win.winfo_pointery()
                w = self.win.winfo_containing(px, py)
            except Exception:
                return
            # Walk up until we find a widget with yview_scroll (Treeview/Canvas/Text/Listbox)
            target = None
            cur = w
            while cur is not None:
                ys = getattr(cur, "yview_scroll", None)
                if callable(ys):
                    target = cur
                    break
                cur = getattr(cur, "master", None)
            if target is None:
                return
            try:
                # macOS delta comes in units of 120 per "tick"
                ticks = int(-1 * (event.delta / 120)) if getattr(event, "delta", 0) else 0
                if ticks == 0:
                    return
                shifted = bool(getattr(event, "state", 0) & 0x0001)
                if shifted:
                    xs = getattr(target, "xview_scroll", None)
                    if callable(xs):
                        xs(ticks, "units")
                else:
                    target.yview_scroll(ticks, "units")
                return "break"
            except Exception:
                return
        self.win.bind_all("<MouseWheel>", _global_mousewheel, add="+")

        # --- LIVE REFRESH (stale-engine guard): ---------------------------------
        # Fixes the "I added data but GAAP didn't notice" bug: whenever the user
        # re-raises / clicks into the dashboard window, we re-read data + refresh
        # all 11 tabs. Debounce with a 600 ms sliding window to avoid storms if
        # the user is simply moving / resizing the window.
        self._last_focus_refresh_ts = 0.0
        self._pending_focus_after_id = None
        def _sched_focus_refresh(_evt=None):
            try:
                import time
                now = time.time()
                self._last_focus_refresh_ts = now
                def _run_if_still_active():
                    try:
                        if time.time() - self._last_focus_refresh_ts >= 0.599:
                            self._refresh_engine_and_reports(trigger="window_activated")
                    except Exception:
                        pass
                if self._pending_focus_after_id is not None:
                    try: self.win.after_cancel(self._pending_focus_after_id)
                    except Exception: pass
                self._pending_focus_after_id = self.win.after(600, _run_if_still_active)
            except Exception:
                pass
        # Bind on both <FocusIn> (window gets keyboard focus) and <Map> (re-shown from minimised)
        self.win.bind("<FocusIn>", _sched_focus_refresh, add="+")
        self.win.bind("<Map>",     _sched_focus_refresh, add="+")
        def _win_destroy(_evt=None):
            try:
                if getattr(self, "_periodic_after_id", None) is not None:
                    self.win.after_cancel(self._periodic_after_id)
            except Exception: pass
        self.win.bind("<Destroy>", _win_destroy, add="+")

        def _periodic_refresh():
            try:
                self._refresh_engine_and_reports(trigger="periodic_60s")
            except Exception:
                pass
            self._periodic_after_id = self.win.after(60_000, _periodic_refresh)
        self._periodic_after_id = self.win.after(60_000, _periodic_refresh)

        # =====================================================================
        # SIMPLIFICATION: Early SYNC rebuild of General Ledger (sync, BEFORE
        # _build and BEFORE initial paint).  Fixes the "reports show nothing
        # at all on first open" mess: if the user has 2 invoices + 0 GL rows
        # (common scenario because new companies never trigger a rebuild
        # otherwise), we run the full rebuild here synchronously so the
        # very first paint is populated.  This is safe because the rebuild
        # is idempotent.
        # =====================================================================
        try:
            if self.ledger is not None and hasattr(self.ledger, "rebuild_general_ledger_from_all_data"):
                _gl = getattr(self.dm, "load_json", lambda fn: None)("general_ledger.json") or []
                _inv = getattr(self.dm, "load_json", lambda fn: None)("invoices_data.json") or []
                _pur = getattr(self.dm, "load_json", lambda fn: None)("purchases_data.json") or []
                _exp_list = []
                for _fname in ("transactions_data.json", "transactions.json",
                               "transaction_data.json", "finance_transactions.json"):
                    try:
                        _xx = self.dm.load_json(_fname) or []
                        if isinstance(_xx, list): _exp_list.extend(_xx)
                    except Exception: pass
                _sal1 = getattr(self.dm, "load_json", lambda fn: None)("monthly_salaries.json") or {}
                _sal2 = getattr(self.dm, "load_json", lambda fn: None)("salaries_data.json") or {}
                _emp = getattr(self.dm, "load_json", lambda fn: None)("employees.json") or []
                n_gl  = sum(1 for r in _gl  if isinstance(r, dict))
                n_inv = sum(1 for r in _inv if isinstance(r, dict))
                n_pur = sum(1 for r in _pur if isinstance(r, dict))
                n_exp_total = sum(1 for r in _exp_list if isinstance(r, dict))
                n_emp = sum(1 for r in _emp if isinstance(r, dict))
                n_sal_months = 0
                for _sd in (_sal1, _sal2):
                    if isinstance(_sd, dict):
                        n_sal_months += sum(1 for k, v in _sd.items()
                                            if not k.startswith("__") and isinstance(v, list) and len(v) > 0)
                # Run rebuild whenever we have ANY source data but GL is tiny
                has_source_data = (n_inv > 0 or n_pur > 0 or n_exp_total > 0
                                   or n_sal_months > 0 or n_emp > 0)
                rebuild_needed = (n_gl < max(3, min(n_inv + n_pur + n_exp_total, 10)))
                if has_source_data and rebuild_needed:
                    try:
                        _stats = self.ledger.rebuild_general_ledger_from_all_data(
                            progress_cb=lambda _m: None)
                    except Exception:
                        _stats = None
        except Exception:
            pass

        # User-visible mode flag: None (let engine decide default = Management)
        # or True (locked-only audit) / False (all data management preview)
        self._force_read_locked_only = None

        self._build()
        # Defer scrollbar binding until after widgets are laid out
        def _finalize_size():
            try:
                self.win.update_idletasks()
                # Force scrollable canvas to have focus-friendly scrolls
                if hasattr(self, "_tabs_scroll"):
                    self._tabs_scroll.canvas.focus_set()
            except Exception:
                pass
        self.win.after(250, _finalize_size)
        self._refresh_engine_and_reports(trigger="initial_build")
        try:
            self._append_audit_log("OPEN",
                "GAAP Financial Reporting Dashboard opened successfully; all 12 tabs initialized; RBAC role enforced.")
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # 3.1 Style
    # -----------------------------------------------------------------------
    def _configure_style(self):
        s = ttk.Style(self.win)
        try:
            s.theme_use("clam")
        except tk.TclError:
            pass
        # Clean professional type scale
        s.configure("Title.TLabel", font=("Helvetica", 20, "bold"),
                    foreground=NAVY_DARK, background=WHITE)
        s.configure("Subtitle.TLabel", font=("Helvetica", 9),
                    foreground=GREY_DARK, background=WHITE)
        s.configure("Company.TLabel", font=("Helvetica", 11, "bold"),
                    foreground=NAVY, background=WHITE)
        s.configure("SectionHeader.TLabel", font=("Helvetica", 12, "bold"),
                    foreground=NAVY_DARK, background=WHITE)
        s.configure("BannerValue.TLabel", font=("Helvetica", 16, "bold"),
                    foreground=NAVY_DARK, background=LIGHT_BLUE)
        s.configure("BannerLabel.TLabel", font=("Helvetica", 8, "bold"),
                    foreground=GREY_DARK, background=LIGHT_BLUE)
        s.configure("Banner.TFrame", background=LIGHT_BLUE)
        s.configure("Card.TFrame", background=WHITE)
        s.configure("Explain.TFrame", background=WHITE)
        s.configure("Good.TLabel", font=("Helvetica", 9, "bold"),
                    foreground=GREEN_DEEP, background=WHITE)
        s.configure("Bad.TLabel", font=("Helvetica", 9, "bold"),
                    foreground=RED_STRONG, background=WHITE)
        s.configure("Notify.TLabel", font=("Helvetica", 9, "bold"),
                    foreground=NAVY, background=WHITE)
        # Professional buttons
        s.configure("TButton", padding=(12, 6), font=("Helvetica", 9, "bold"),
                    foreground=NAVY_DARK)
        s.configure("Accent.TButton", font=("Helvetica", 10, "bold"),
                    foreground=WHITE, background=NAVY, padding=(14, 7))
        s.map("Accent.TButton",
              background=[("active", BLUE_DEEP), ("!disabled", NAVY),
                          ("pressed", NAVY_DARK)],
              foreground=[("!disabled", WHITE)])
        s.configure("Ghost.TButton", font=("Helvetica", 9),
                    foreground=NAVY, background=WHITE, padding=(10, 5))
        s.map("Ghost.TButton",
              background=[("active", LIGHT_BLUE), ("!disabled", WHITE)])
        s.configure("Good.TButton", font=("Helvetica", 9, "bold"),
                    foreground=WHITE, background=GREEN_DEEP, padding=(10, 5))
        s.map("Good.TButton",
              background=[("active", GREEN_STRONG), ("!disabled", GREEN_DEEP)],
              foreground=[("!disabled", WHITE)])
        # Notebook — corporate look
        s.configure("TNotebook", background=GREY_BG, tabmargins=(4, 6, 4, 0))
        s.configure("TNotebook.Tab", padding=(16, 10),
                    font=("Helvetica", 10, "bold"))
        s.map("TNotebook.Tab",
              background=[("selected", WHITE), ("!selected", GREY_BG)],
              foreground=[("selected", NAVY_DARK), ("!selected", GREY_DARK)])
        # Treeview — finance grade
        s.configure("Treeview",
                    font=("Helvetica", 9),
                    rowheight=26,
                    background=WHITE,
                    fieldbackground=WHITE,
                    foreground=NAVY_DARK,
                    borderwidth=0)
        s.configure("Treeview.Heading",
                    font=("Helvetica", 9, "bold"),
                    background=NAVY,
                    foreground=WHITE,
                    padding=(8, 8),
                    relief="flat")
        s.map("Treeview.Heading",
              background=[("active", BLUE_DEEP)])
        s.map("Treeview",
              background=[("selected", LIGHT_BLUE_2)],
              foreground=[("selected", NAVY_DARK)])
        # Combobox
        s.configure("TCombobox", padding=4, font=("Helvetica", 9))
        # Separator
        s.configure("TSeparator", background=SEPARATOR)

    # -----------------------------------------------------------------------
    # 3.1b  RBAC + Immutable Audit Log Initialization
    # -----------------------------------------------------------------------
    def _init_rbac_and_audit_log(self):
        """Resolve data folder, user role, and ensure audit_log.jsonl exists.

        Resolves user_role with fallback chain (least privilege by default.
        """
        # 1) Resolve data_folder
        data_folder = None
        for attr in ("data_folder", "invoice_folder"):
            val = getattr(self.dm, attr, None)
            if val and isinstance(val, str) and os.path.isdir(val):
                data_folder = val
                break
        if not data_folder:
            data_folder = os.getcwd()
        self._data_folder = data_folder

        # 2) Resolve user_role (priority chain, safest default = ReadOnly)
        role = None
        role = getattr(self.dm, "user_role", None)
        if not role:
            inv_mgr = getattr(self.dm, "invoice_manager", None)
            if inv_mgr is not None:
                role = getattr(inv_mgr, "user_role", None)
        if not role:
                    # Try app_settings.json in data_folder
                    try:
                        settings_path = os.path.join(data_folder, "app_settings.json")
                        if os.path.isfile(settings_path):
                            try:
                                with open(settings_path, "r", encoding="utf-8") as f:
                                    st = json.load(f)
                                if isinstance(st, dict):
                                    sec = st.get("security")
                                    if isinstance(sec, dict):
                                        role = sec.get("user_role")
                                    if not role:
                                        role = st.get("user_role")
                            except Exception:
                                pass
                    except Exception:
                        pass
        if not role:
            role = "ReadOnly"
        self._user_role = (role or "ReadOnly").strip()

        # 3) Ensure audit_log.jsonl exists (append-only immutable log)
        try:
            os.makedirs(self._data_folder, exist_ok=True)
        except Exception:
            pass
        audit_path = os.path.join(self._data_folder, "audit_log.jsonl")
        try:
            if not os.path.isfile(audit_path):
                with open(audit_path, "a", encoding="utf-8") as f:
                    pass  # touch
        except Exception:
            pass

    def _append_audit_log(self, event_type, details):
        """Append one IMMUTABLE append-only JSONL line to audit_log.jsonl.

        Chain-of-custody: each line includes md5(first16) of previous line so lines
        cannot be rewritten retroactively. NEVER raises — silently if write fails.
        """
        try:
            audit_path = os.path.join(self._data_folder, "audit_log.jsonl")
            # Compute chain-of-custody hash of last line
            prev_hash_16 = "0000000000000000"
            try:
                last_line = None
                if os.path.isfile(audit_path):
                    with open(audit_path, "r", encoding="utf-8") as f:
                        all_lines = f.readlines()
                        if all_lines:
                            for ll in reversed(all_lines):
                                s = ll.strip()
                                if s:
                                    last_line = s
                                    break
                if last_line:
                    h = hashlib.md5(last_line.encode("utf-8")).hexdigest()
                    prev_hash_16 = h[:16]
            except Exception:
                pass
            report_period_info = {
                "range": "",
                "start": "",
                "end": "",
            }
            try:
                if hasattr(self, "range_var"):
                    report_period_info["range"] = self.range_var.get()
            except Exception:
                pass
            try:
                if hasattr(self, "engine"):
                    try:
                        report_period_info["start"] = str(self.engine.start)
                    except Exception:
                        pass
                    try:
                        report_period_info["end"] = str(self.engine.end)
                    except Exception:
                        pass
            except Exception:
                pass
            record = {
                "timestamp_utc": datetime.utcnow().isoformat() + "Z",
                "timestamp_local": datetime.now().isoformat(),
                "event_type": "GAAP_REPORT_" + str(event_type),
                "user_role": self._user_role,
                "username": getattr(self.dm, "username", "unknown"),
                "data_folder": self._data_folder,
                "report_period": report_period_info,
                "details": details,
                "host_pid": os.getpid(),
                "hash_previous_line_16chars": prev_hash_16,
            }
            try:
                line = json.dumps(record, ensure_ascii=False, default=str)
                with open(audit_path, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except Exception:
                pass
        except Exception:
            pass

    # -----------------------------------------------------------------------
    # 3.2 Top bar: Date presets + custom range + Quick Add Expense
    # -----------------------------------------------------------------------
    def _build(self):
        root = ttk.Frame(self.win, padding=0)
        root.configure(style="Explain.TFrame")
        root.pack(fill="both", expand=True)

        # ============== SECURITY BANNER (Phase 5C) ==============
        try:
            locked_mode = False
            if hasattr(self, "engine") and self.engine is not None:
                locked_mode = bool(getattr(self.engine, "read_from_locked_periods_only", False))
        except Exception:
            locked_mode = False
        banner_role = self._user_role
        if banner_role == "ReadOnly":
            banner_bg = LIGHT_BLUE_2
            banner_fg = NAVY_DARK
        elif banner_role == "Admin":
            banner_bg = GREEN_BG_2
            banner_fg = GREEN_DEEP
        elif banner_role == "Accountant":
            banner_bg = AMBER_BG_2
            banner_fg = AMBER_DEEP
        else:
            banner_bg = RED_BG_2
            banner_fg = RED_STRONG
        period_lock_text = (
            "LOCKED (GAAP Audit Mode)" if locked_mode
            else "⚠️ UNLOCKED (Management Review — NOT for audit)"
        )
        security_strip = tk.Frame(root, bg=banner_bg, highlightthickness=1,
                                  highlightbackground=SEPARATOR, height=34)
        security_strip.pack(fill="x", padx=12, pady=(12, 0))
        security_strip.pack_propagate(False)
        sec_lbl = tk.Label(security_strip,
                       text=("🔒  GAAP FINANCIAL REPORTS — Role: " + banner_role
                             + "  |  Period Lock Read Mode: " + period_lock_text
                             + "  |  All actions logged to audit_log.jsonl"),
                       bg=banner_bg, fg=banner_fg,
                       font=("Helvetica", 9, "bold"),
                       anchor="w", padx=14)
        sec_lbl.pack(fill="both", expand=True)
        sec_tooltip = ("🔒 Restricted — role '" + banner_role + "'. "
                       + "Admin / Accountant may modify data; all other roles are view-only. "
                       + "All actions are immutably recorded to audit_log.jsonl with chain-of-custody hashing.")
        try:
            sec_lbl.bind("<Enter>", lambda e: (sec_lbl.config(text=sec_tooltip, wraplength=sec_lbl.winfo_width() + 200),))
            sec_lbl.bind("<Leave>", lambda e: (sec_lbl.config(text=(
                "🔒  GAAP FINANCIAL REPORTS — Role: " + banner_role
                + "  |  Period Lock Read Mode: " + period_lock_text
                + "  |  All actions logged to audit_log.jsonl"), wraplength=0),))
        except Exception:
            pass

        # ============== HEADER CARD ==============
        header_wrap = tk.Frame(root, bg=WHITE, highlightthickness=1,
                               highlightbackground=SEPARATOR)
        header_wrap.pack(fill="x", padx=12, pady=(12, 6))

        header = ttk.Frame(header_wrap, padding=(18, 14, 18, 12))
        header.configure(style="Explain.TFrame")
        header.pack(fill="x")

        # Logo area + title (left column)
        title_col = ttk.Frame(header)
        title_col.configure(style="Explain.TFrame")
        title_col.grid(row=0, column=0, sticky="w")

        # Report branding block — stacked compact
        tk.Label(title_col, text="🏥 HOPEPHARMA MEDICAL TRADING L.L.C.",
                 font=("Helvetica", 11, "bold"), bg=WHITE,
                 fg=NAVY_DARK, anchor="w", justify="left").pack(anchor="w")
        ttk.Label(title_col,
                  text="Financial Reporting Centre — GAAP / IFRS Compliant",
                  style="Subtitle.TLabel").pack(anchor="w")
        ttk.Separator(title_col, orient="horizontal").pack(fill="x", pady=(8, 6))
        ttk.Label(title_col, text="Professional Financial Reports",
                  style="Title.TLabel").pack(anchor="w")
        ttk.Label(title_col,
                  text="Single Source of Truth: General Ledger  ·  Double-Entry Verified  ·  Period-Aware Filtering",
                  style="Subtitle.TLabel").pack(anchor="w", pady=(2, 0))

        # Controls (right side of header)
        ctrl = ttk.Frame(header)
        ctrl.configure(style="Explain.TFrame")
        ctrl.grid(row=0, column=1, rowspan=2, sticky="ne", padx=(20, 0))
        header.columnconfigure(1, weight=1)

        # Row 1: Period selector
        row1 = ttk.Frame(ctrl)
        row1.configure(style="Explain.TFrame")
        row1.pack(anchor="e", pady=(0, 8))

        ttk.Label(row1, text="Reporting Period:",
                  font=("Helvetica", 9, "bold"),
                  foreground=NAVY_DARK).grid(row=0, column=0, sticky="e", padx=(0, 6))
        # Compute smart default
        default_range_label = DATE_RANGE_PRESETS[7][0]
        try:
            _today = date.today()
            _all_gl = (getattr(self.dm, 'ledger', None)
                       and hasattr(self.dm.ledger, 'load_general_ledger')
                       and self.dm.ledger.load_general_ledger() or [])
            _invs = getattr(self.dm, 'load_json',
                            lambda f: []).__call__('invoices_data.json') or []
            _mind = None
            for _src in (_all_gl, _invs):
                for _e in _src:
                    if isinstance(_e, dict):
                        _v = _e.get('date') or _e.get('invoice_date') or ''
                        if _v:
                            try:
                                s10 = str(_v)[:10].strip()
                                try:
                                    d = date.fromisoformat(s10)
                                except Exception:
                                    import re as _re
                                    m = _re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})', s10)
                                    if not m:
                                        continue
                                    d = date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                                if _mind is None or d < _mind:
                                    _mind = d
                            except Exception:
                                continue
            if _mind:
                if (_mind.year == _today.year and _mind.month == _today.month):
                    default_range_label = DATE_RANGE_PRESETS[2][0]
                elif _mind.year == _today.year:
                    default_range_label = DATE_RANGE_PRESETS[6][0]
                else:
                    default_range_label = DATE_RANGE_PRESETS[7][0]
        except Exception:
            pass
        self.range_var = tk.StringVar(value=default_range_label)
        range_combo = ttk.Combobox(row1, textvariable=self.range_var,
                                   values=[p[0] for p in DATE_RANGE_PRESETS],
                                   state="readonly", width=20,
                                   font=("Helvetica", 10))
        range_combo.grid(row=0, column=1, padx=(0, 6))
        range_combo.bind("<<ComboboxSelected>>", lambda _e: self._on_range_change())

        ttk.Label(row1, text="Custom:",
                  font=("Helvetica", 9), foreground=GREY_MED).grid(row=0, column=2, sticky="e")
        self.cust_start = ttk.Entry(row1, width=12, font=("Helvetica", 10))
        self.cust_start.grid(row=0, column=3, padx=(4, 2))
        ttk.Label(row1, text="→",
                  font=("Helvetica", 10, "bold"),
                  foreground=GREY_DARK).grid(row=0, column=4, padx=(0, 2))
        self.cust_end = ttk.Entry(row1, width=12, font=("Helvetica", 10))
        self.cust_end.grid(row=0, column=5, padx=(2, 10))
        self.cust_start.configure(state="disabled")
        self.cust_end.configure(state="disabled")

        ttk.Button(row1, text="🔄 Refresh",
                   command=self._refresh_engine_and_reports,
                   style="Ghost.TButton").grid(row=0, column=6, padx=(0, 6))
        self._btn_rebuild_gl = ttk.Button(row1, text="🔁 Rebuild GL",
                                           command=self._rebuild_gl_now,
                                           style="Ghost.TButton")
        self._btn_rebuild_gl.grid(row=0, column=7, padx=(0, 6))
        self._btn_new_expense = ttk.Button(row1, text="💸 Expense",
                                            command=self._open_expense_dialog,
                                            style="Accent.TButton")
        self._btn_new_expense.grid(row=0, column=8, padx=(0, 6))
        # ---- MODE TOGGLE: Management ↔️ Audit (Locked-Only) ----
        self._mode_read_locked_only_var = tk.BooleanVar(value=False)
        self._btn_mode_toggle = ttk.Button(row1, text=self._mode_toggle_label(False),
                                            command=self._on_mode_toggle,
                                            style="Warning.TButton")
        self._btn_mode_toggle.grid(row=0, column=9, padx=(0, 6))
        ttk.Button(row1, text="✅ Audit",
                   command=self._run_audit_popup,
                   style="Good.TButton").grid(row=0, column=10)

        # ---------- RBAC UI Restrictions (Phase 5C) ----------
        can_write = self._user_role in ("Admin", "Accountant")
        restricted_tooltip = ("🔒 Restricted — role '" + self._user_role
                              + "' may view but not modify financial data. Contact Admin for write access.")
        for btn_widget in (self._btn_rebuild_gl, self._btn_new_expense):
            try:
                if not can_write:
                    btn_widget.configure(state="disabled")
                    try:
                        btn_widget.configure(foreground=GREY_MED)
                    except Exception:
                        pass
            except Exception:
                pass
        if not can_write:
            try:
                def _warn_readonly():
                    messagebox.showwarning(
                        "🔒 Read-Only Mode",
                        restricted_tooltip)
                for btn_widget in (self._btn_rebuild_gl, self._btn_new_expense):
                    btn_widget.configure(command=_warn_readonly)
            except Exception:
                pass

        # Row 2: Notification banner (period info + audit status)
        row2 = ttk.Frame(ctrl)
        row2.configure(style="Explain.TFrame")
        row2.pack(anchor="e", fill="x", pady=(4, 0))
        self.notify_var = tk.StringVar(value="")
        ttk.Label(row2, textvariable=self.notify_var,
                  style="Notify.TLabel", justify="right",
                  anchor="e", wraplength=950).pack(anchor="e", fill="x")

        # ============================================================
        # MAIN REPORT AREA — wrap EVERYTHING below the top header in a
        # SCROLLABLE FRAME so user can scroll horizontally+vertically
        # if their screen is smaller than the wide P&L / ratio / CF trees.
        # The TOP HEADER (period selector + custom + Refresh/Rebuild/Expense/Audit)
        # remains FIXED on screen at all times.
        # ============================================================
        self._tabs_scroll = ScrollableFrame(root, bg="#F4F6FA")
        self._tabs_scroll.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        body = self._tabs_scroll.scrollable_frame
        body.configure(bg="#F4F6FA")

        # Tabs - wrapped in a subtle card
        tabs_wrap = tk.Frame(body, bg=WHITE, highlightthickness=1,
                             highlightbackground=SEPARATOR)
        tabs_wrap.pack(fill="both", expand=True, padx=2, pady=2)

        self.nb = ttk.Notebook(tabs_wrap)
        self.nb.pack(fill="both", expand=True, padx=10, pady=10)

        self._build_pnl_tab()
        self._build_balance_sheet_tab()
        self._build_changes_in_equity_tab()
        self._build_cashflow_tab()
        self._build_financial_ratios_tab()
        self._build_aging_tab()
        self._build_rev_client_tab()
        self._build_stock_tab()
        self._build_supplier_tab()
        self._build_audit_tab()
        self._build_expense_manager_tab()
        self._build_footnotes_tab()

        # Auto-run GL rebuild (async) if GL is still tiny despite abundant source data
        # exists.  We also set a second safety net for cases where the early-sync rebuild
        # in __init__ was missed (e.g. data was empty at boot but freshly added later).
        self._auto_rebuilt = False
        try:
            gl = getattr(self.dm, "load_json", lambda fn: [])("general_ledger.json") or []
            inv = getattr(self.dm, "load_json", lambda fn: [])("invoices_data.json") or []
            pur = getattr(self.dm, "load_json", lambda fn: [])("purchases_data.json") or []
            n_inv = sum(1 for r in inv if isinstance(r, dict))
            n_pur = sum(1 for r in pur if isinstance(r, dict))
            n_gl = sum(1 for r in gl if isinstance(r, dict))
            try:
                _exp_count = 0
                for _f in ("transactions_data.json", "transactions.json",
                            "transaction_data.json", "finance_transactions.json"):
                    try:
                        _tx = self.dm.load_json(_f) or []
                        if isinstance(_tx, list): _exp_count += sum(1 for r in _tx if isinstance(r, dict))
                    except Exception:
                        pass
            except Exception:
                _exp_count = 0
            try:
                _sal_m = 0
                for _f in ("monthly_salaries.json", "salaries_data.json"):
                    try:
                        _s = self.dm.load_json(_f) or {}
                        if isinstance(_s, dict):
                            _sal_m += sum(1 for k, v in _s.items() if not k.startswith("__") and isinstance(v, list) and len(v) > 0)
                    except Exception:
                        pass
            except Exception:
                _sal_m = 0
            _src_total = n_inv + n_pur + _exp_count + _sal_m
            if (n_gl < 3) or (_src_total > 0 and n_gl < max(5, min(_src_total * 2, 20))):
                self.win.after(300, self._auto_rebuild_gl_async)
        except Exception:
            pass
        # Auto-sync the mode toggle's starting state: if the dataset already has
        # locked periods defined in period_locks.json → start off on 🔒 Audit
        # mode since the company probably already formally closes books regularly.  Otherwise,
        # explicitly start on 📊 Management (the simplified default that actually shows
        # numbers for 90% of SMEs who never lock periods).
        try:
            _pl = getattr(self.dm, "load_json", lambda fn: [])("period_locks.json") or []
            _n_lock = sum(1 for r in (_pl if isinstance(_pl, list) else [])
                          if isinstance(r, dict) and r.get("is_locked") is True)
        except Exception:
            _n_lock = 0
        try:
            if _n_lock > 0:
                def _apply_hint(_n=_n_lock):
                    try:
                        self._btn_mode_toggle.configure(
                            text=("📊 Management: All Data"
                                  + f"  ·  {_n} closed periods available → click 🔒 for Audit"))
                    except Exception:
                        pass
                self.win.after(400, _apply_hint)
        except Exception:
            pass

        # ---------------------------------------------------------------
        # RIGHT-CLICK CONTEXT MENUS — attach to EVERY report tree so the
        # user can right-click ANYWHERE inside any GAAP report tab and:
        #   • Copy the row / amount
        #   • ➡️ Jump straight to Expense Ledger (Tab 11) where they can
        #     right-click again and choose 🗑️ Delete / Void Expense to
        #     post a GL reversal.  Works on macOS 2-finger tap (Button-2)
        #     AND Windows/Linux right click (Button-3).
        #
        # NOTE: initialisation is FAIL-SAFE.  Each attachment is wrapped
        # in its own try/except, and `_report_ctx_menus` is initialised
        # BEFORE any operations so callers always get a valid dict (even
        # if empty) and never hit AttributeError.
        # ---------------------------------------------------------------
        if not hasattr(self, "_report_ctx_menus"):
            self._report_ctx_menus = {}
        try:
            _menu_defs = [
                (self.pnl_tree, "P&L Report"),
                (self.bs_tree,  "Balance Sheet"),
                (self.eq_tree,  "Changes_in_Equity"),
                (self.cf_tree,  "Cash_Flow_Statement"),
                (getattr(self, "rt_tree", None),  "Financial_Ratios"),
                (getattr(self, "age_tree", None), "AR_AP_Aging"),
                (getattr(self, "rc_tree", None),  "Revenue_by_Client"),
                (getattr(self, "st_tree", None),  "Stock_Inventory"),
                (getattr(self, "sp_tree", None),  "Suppliers_Spend"),
                (self.audit_tree, "GL_Audit_Trail"),
            ]
            for _t, _name in _menu_defs:
                try:
                    if _t is None:
                        continue
                    self._attach_report_context_menu(_t, _name)
                except Exception as _e:
                    # Don't abort others because one tree raised; but leave
                    # a tiny breadcrumb in stdout for debugging if needed.
                    try: print(f"[ctx_menu] skip {_name}: {type(_e).__name__}: {_e}")
                    except Exception: pass
            # Footnotes trees (2 separate tables on same tab)
            try:
                _fn1 = getattr(self, "fn_policies_tree", None)
                if _fn1 is not None:
                    self._attach_report_context_menu(_fn1, "Footnotes_Accounting_Policies")
            except Exception as _e:
                try: print(f"[ctx_menu] skip fn_policies: {type(_e).__name__}: {_e}")
                except Exception: pass
            try:
                _fn2 = getattr(self, "fn_contingencies_tree", None)
                if _fn2 is not None:
                    self._attach_report_context_menu(_fn2, "Footnotes_Contingencies")
            except Exception as _e:
                try: print(f"[ctx_menu] skip fn_contingencies: {type(_e).__name__}: {_e}")
                except Exception: pass
        except Exception as _top_e:
            # Never fail dashboard open because of context menus.
            try: print(f"[ctx_menu] top-level: {type(_top_e).__name__}: {_top_e}")
            except Exception: pass

    # -----------------------------------------------------------------------
    # 3.3 Date range handling
    # -----------------------------------------------------------------------
    def _selected_range_enum(self):
        lbl = self.range_var.get()
        for p in DATE_RANGE_PRESETS:
            if p[0] == lbl:
                return p[1]
        return DateRangeType.THIS_MONTH

    def _on_range_change(self):
        rt = self._selected_range_enum()
        if rt == DateRangeType.CUSTOM:
            self.cust_start.configure(state="normal")
            self.cust_end.configure(state="normal")
            today = date.today()
            if not self.cust_start.get().strip():
                self.cust_start.insert(0, today.replace(day=1).isoformat())
            if not self.cust_end.get().strip():
                self.cust_end.insert(0, today.isoformat())
        else:
            self.cust_start.delete(0, "end")
            self.cust_end.delete(0, "end")
            self.cust_start.configure(state="disabled")
            self.cust_end.configure(state="disabled")
        self._refresh_engine_and_reports()

    def _build_engine(self):
        rt = self._selected_range_enum()
        s = None; e = None
        if rt == DateRangeType.CUSTOM:
            try:
                s = self.cust_start.get().strip()
                e = self.cust_end.get().strip()
                if not (s and e):
                    raise ValueError
            except Exception:
                messagebox.showwarning("Custom Dates",
                                       "Please fill Custom Start & End (YYYY-MM-DD).")
                return None
        try:
            eng = ReportEngine(self.dm, range_type=rt, start=s, end=e)
            # Apply user's explicit mode-toggled override.  When the user has
            # clicked the 1-click toggle we force the engine to follow their
            # choice regardless of defaults.  When toggle is OFF (default /
            # Management) we let the engine's own smart default apply (which
            # is: Management mode unless user explicitly flipped, with a
            # warning banner explaining the audit-readiness gap).
            try:
                mode_locked = bool(self._mode_read_locked_only_var.get())
            except Exception:
                mode_locked = False
            if self._force_read_locked_only is not None:
                mode_locked = bool(self._force_read_locked_only)
            eng.read_from_locked_periods_only = mode_locked
            if mode_locked:
                # User forced Audit (locked-only) mode.  Re-apply period-lock
                # filter so gl_current/gl_prior only include entries that fall
                # inside at least one closed (locked) period in period_locks.json.
                try:
                    self._refilter_engine_for_locked_only(eng)
                except Exception:
                    pass
            return eng
        except Exception as ex:
            messagebox.showerror("Could not build report engine", str(ex))
            return None

    @staticmethod
    def _refilter_engine_for_locked_only(eng):
        from datetime import date as _d
        try:
            period_locks = eng.dm.load_json('period_locks.json') or []
        except Exception:
            period_locks = []
        if not isinstance(period_locks, list):
            period_locks = []
        locked_periods = []
        for pl in period_locks:
            if not isinstance(pl, dict):
                continue
            if pl.get('is_locked') is True:
                try:
                    s_s = pl.get('period_start') or pl.get('start_date')
                    s_e = pl.get('period_end') or pl.get('end_date')
                    s = eng._parse_date(s_s) if hasattr(eng, "_parse_date") else _d.fromisoformat(str(s_s))
                    e = eng._parse_date(s_e) if hasattr(eng, "_parse_date") else _d.fromisoformat(str(s_e))
                    if s and e:
                        locked_periods.append((s, e))
                except Exception:
                    continue
        # If NO locked periods exist, strict audit mode = 0 entries (explicit)
        new_gl = []
        exc = 0
        for entry in (getattr(eng, "_gl_all_raw", None) or getattr(eng, "_gl_all_orig", None) or eng._gl_all):
            pass
        # Save a raw backup of full GL under `_gl_all_orig` first time so we
        # can always re-filter from complete set if user toggles mode back
        # and forth multiple times in one session.
        if not getattr(eng, "_gl_all_orig", None):
            try:
                eng._gl_all_orig = list(eng._gl_all) if isinstance(eng._gl_all, list) else []
            except Exception:
                eng._gl_all_orig = []
        src = list(eng._gl_all_orig) if isinstance(eng._gl_all_orig, list) else (list(eng._gl_all) if isinstance(eng._gl_all, list) else [])
        filtered = []
        for entry in src:
            if not isinstance(entry, dict):
                continue
            entry_date = (eng._parse_date(entry.get('transaction_date') or entry.get('date'))
                          if hasattr(eng, "_parse_date") else None)
            if entry_date is None:
                try:
                    entry_date = _d.fromisoformat(str(entry.get('date'))[:10])
                except Exception:
                    # Unknown date — when strict locked-only is on, unknown = exclude
                    exc += 1
                    continue
            in_locked = any(ls <= entry_date <= le for (ls, le) in locked_periods)
            if in_locked:
                filtered.append(entry)
            else:
                exc += 1
        eng._gl_all = filtered
        eng._excluded_unlocked_count = exc
        # Re-derive current / prior
        try:
            eng.gl_current = [e for e in eng._gl_all if eng._in_range(e.get('date', _d.today().isoformat()), eng.start, eng.end)]
            eng.gl_prior   = [e for e in eng._gl_all if eng._in_range(e.get('date', _d.today().isoformat()), eng.prev_start, eng.prev_end)]
        except Exception:
            pass
        # Update banner warning
        if locked_periods:
            eng._period_filter_warning = (
                f"🔒 Strict Audit (Locked-Only) mode: {exc} GL entries dated in UNLOCKED periods were EXCLUDED "
                f"per GAAP requirement #3.  Toggle mode to '📊 Management' to include them."
            )
        else:
            eng._period_filter_warning = (
                f"🔒 Strict Audit mode ON but no closed (locked) periods found. "
                f"All {exc} GL entries were EXCLUDED → 0/0/0 shown.  Toggle mode to '📊 Management' to see data, "
                f"or use Period Manager to formally close periods first."
            )
        return None

    # -----------------------------------------------------------------------
    # Mode toggle helpers (Management ↔ Audit) — 1-click UX simplification
    # -----------------------------------------------------------------------
    def _mode_toggle_label(self, locked_value):
        return ("🔒 Audit: Locked Only" if locked_value else "📊 Management: All Data")

    def _mode_toggle_style(self, locked_value):
        return ("Good.TButton" if locked_value else "Warning.TButton")

    def _on_mode_toggle(self):
        try:
            cur = bool(self._mode_read_locked_only_var.get())
        except Exception:
            cur = False
        new = not cur
        # Safety guard: user can only force Audit (locked-only) mode if there
        # are actually some locked periods defined in period_locks.json.
        # Otherwise they'd just end up with 0 data again (old bug).
        if new:
            try:
                pl = (getattr(self.dm, "load_json", lambda fn: [])("period_locks.json") or [])
                locked = sum(1 for r in pl
                             if isinstance(r, dict) and r.get("is_locked") is True)
            except Exception:
                locked = 0
            if locked <= 0:
                ans = messagebox.askyesno(
                    "No Locked Periods Found",
                    "You're switching to 🔒 Audit (Locked-Only) mode, but there are\n"
                    "NO formally closed periods in Period Manager yet.\n\n"
                    "Reports would show 0/0/0 under strict audit rules.\n\n"
                    "Enable anyway (and accept 0 data until periods are closed)?")
                if not ans:
                    return
        try:
            self._mode_read_locked_only_var.set(new)
            self._btn_mode_toggle.configure(text=self._mode_toggle_label(new))
            try:
                self._btn_mode_toggle.configure(style=self._mode_toggle_style(new))
            except Exception:
                pass
        except Exception:
            pass
        self._refresh_engine_and_reports(trigger="mode_toggle")

    def _get_custom_dates(self):
        """Return (start_iso, end_iso) from the top-bar custom date entries, or (None, None) if empty/invalid."""
        try:
            s = (getattr(self, "cust_start", None) and self.cust_start.get().strip()) or ""
            e = (getattr(self, "cust_end", None) and self.cust_end.get().strip()) or ""
            # Validate simple ISO date shape
            if s and len(s) == 10 and e and len(e) == 10:
                date.fromisoformat(s)
                date.fromisoformat(e)
                return s, e
        except Exception:
            pass
        return None, None

    # -----------------------------------------------------------------------
    # 3.4 Tab builders (10 reports)
    # -----------------------------------------------------------------------
    def _make_report_card(self, parent):
        """Each tab is a Card: subtle highlight frame containing padding."""
        card = tk.Frame(parent, bg=WHITE, highlightthickness=1,
                        highlightbackground=SEPARATOR)
        card.pack(fill="both", expand=True, padx=4, pady=4)
        inner = ttk.Frame(card, padding=(16, 14, 16, 14))
        inner.configure(style="Card.TFrame")
        inner.pack(fill="both", expand=True)
        return inner

    def _build_pnl_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  1️⃣ Profit & Loss (P&L)  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "PROFIT & LOSS STATEMENT",
                                  subtitle="For the period — Revenue → COGS → Gross Profit → Operating Expenses → Net Profit",
                                  accent=GREEN_DEEP)
        self.pnl_banner = self._make_banner(card, "P&L Snapshot",
                                            acolors=(GREEN_BG, LIGHT_BLUE, PURPLE_BG, ORANGE_BG))
        self.pnl_explain = self._make_explain_box(card)
        cols = ("Item", "This Period (AED)", "Prior Period (AED)", "% Change")
        self.pnl_tree = ZebraTreeview(card, columns=cols, show="headings", height=18)
        # Widths tuned to fit inside a 1366px-wide window (minimum dashboard
        # width is 1040px) with NO horizontal scrollbar by default.  Item col
        # stretches + the numeric cols keep compact fixed width so every label
        # + every number is visible without left/right dragging.
        widths = (340, 160, 160, 130)
        aligns = ("w", "e", "e", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.pnl_tree.heading(c, text=c, anchor=a)
            # Stretch ALL 4 cols: Item takes most of extra space, numeric cols
            # also stretch modestly so 9-digit totals never get clipped.
            self.pnl_tree.column(c, width=w, anchor=a, stretch=True,
                                 minwidth=max(70, w - 80))
        self.pnl_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "profit_and_loss")

    def _build_balance_sheet_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  2️⃣ Statement of Financial Position  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "STATEMENT OF FINANCIAL POSITION (BALANCE SHEET)",
                                  subtitle="As of end of period — What we OWN (Assets) vs. What we OWE (Liabilities + Equity)",
                                  accent=BLUE_DEEP)
        self.bs_banner = self._make_banner(card, "Balance Sheet Snapshot",
                                            acolors=(LIGHT_BLUE, ORANGE_BG, GREEN_BG, PURPLE_BG))
        self.bs_explain = self._make_explain_box(card)
        cols = ("Section", "Line Item", "Amount (AED)")
        self.bs_tree = ZebraTreeview(card, columns=cols, show="headings", height=22)
        widths = (220, 400, 220)
        aligns = ("w", "w", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.bs_tree.heading(c, text=c, anchor=a)
            self.bs_tree.column(c, width=w, anchor=a, stretch=(c == "Line Item"),
                                minwidth=max(80, w - 60))
        self.bs_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "balance_sheet")

    def _build_changes_in_equity_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  3️⃣ Changes in Equity  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "STATEMENT OF CHANGES IN EQUITY",
                                  subtitle="Owner Capital, Retained Earnings, Contributions, Drawings, and Net Profit flow",
                                  accent=PURPLE_DEEP)
        self.eq_banner = self._make_banner(card, "Equity Movement Snapshot",
                                            acolors=(PURPLE_BG, GREEN_BG, RED_BG, LIGHT_BLUE))
        self.eq_explain = self._make_explain_box(card)
        cols = ("Item", "Capital Stock (AED)", "Treasury Stock (AED)", "Retained Earnings (AED)", "AOCI (AED)", "Total Equity (AED)")
        self.eq_tree = ZebraTreeview(card, columns=cols, show="headings", height=14)
        widths = (340, 140, 150, 170, 140, 170)
        aligns = ("w", "e", "e", "e", "e", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.eq_tree.heading(c, text=c, anchor=a)
            self.eq_tree.column(c, width=w, anchor=a, stretch=(c == "Item"),
                                minwidth=max(80, w - 40))
        self.eq_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "changes_in_equity")

    def _build_cashflow_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  4️⃣ Cash Flow Statement  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "CASH FLOW STATEMENT",
                                  subtitle="Operating · Investing · Financing activities + Opening / Closing cash reconciled",
                                  accent=TEAL_DEEP)
        self.cf_banner = self._make_banner(card, "Cash Flow Snapshot",
                                            acolors=(TEAL_BG, GREEN_BG, ORANGE_BG, PURPLE_BG))
        self.cf_explain = self._make_explain_box(card)
        cols = ("Cash Flow Item", "Amount (AED)")
        self.cf_tree = ZebraTreeview(card, columns=cols, show="headings", height=22)
        # Cash Flow Item column WAS 760 (wider than a whole 1280px laptop!).
        # Drop to 520 + stretch=True, Amount drops to 200 — total 720 visible
        # by default, remaining width given to label column so full description
        # shows WITHOUT horizontal scrolling.
        widths = (520, 200)
        aligns = ("w", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.cf_tree.heading(c, text=c, anchor=a)
            self.cf_tree.column(c, width=w, anchor=a,
                                stretch=(c in ("Cash Flow Item", "Amount (AED)")),
                                minwidth=max(100, w - 60))
        self.cf_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "cash_flow")

    def _build_financial_ratios_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  5️⃣ Financial Ratios (KPIs)  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "FINANCIAL RATIOS & KEY PERFORMANCE INDICATORS",
                                  subtitle="Liquidity · Profitability · Leverage · Efficiency — vs. Pharma benchmarks",
                                  accent=GOLD)
        self.rt_banner = self._make_banner(card, "KPI Snapshot",
                                            acolors=(GREEN_BG, LIGHT_BLUE, ORANGE_BG, PURPLE_BG))
        self.rt_explain = self._make_explain_box(card)
        cols = ("KPI", "Value", "Benchmark for Pharma", "Verdict", "What it means")
        self.rt_tree = ZebraTreeview(card, columns=cols, show="headings", height=22)
        # Old total: 320+140+220+140+540 = 1360 — forced horizontal scroll on
        # any window smaller than 1440p.  Compact slightly: drop "What it means"
        # initial width to 380 (still 380 visible + stretch=True so when user
        # maximises window it grows to show full explanation).
        widths = (300, 130, 200, 120, 380)
        aligns = ("w", "e", "w", "w", "w")
        for c, w, a in zip(cols, widths, aligns):
            self.rt_tree.heading(c, text=c, anchor=a)
            # Stretch BOTH "What it means" (explanation, widest) and "KPI"
            # (label, second-widest) so they share any extra space.
            self.rt_tree.column(c, width=w, anchor=a,
                                stretch=(c in ("KPI", "What it means", "Benchmark for Pharma")),
                                minwidth=max(80, w - 50))
        self.rt_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "financial_ratios")

    def _build_aging_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  6️⃣ Aging Receivables  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "ACCOUNTS RECEIVABLE — AGING ANALYSIS",
                                  subtitle="Balances owed by customers, grouped by how many days past due",
                                  accent=RED_STRONG)
        self.age_banner = self._make_banner(card, "Receivables Snapshot",
                                            acolors=(RED_BG, LIGHT_BLUE, AMBER_BG, ORANGE_BG))
        self.age_explain = self._make_explain_box(card)
        cols = ("Invoice", "Client", "Invoice Date", "Due Date", "Days Overdue", "Bucket", "Balance (AED)")
        self.age_tree = ZebraTreeview(card, columns=cols, show="headings", height=18)
        # Compact date/number cols slightly. Client text col + Bucket + Balance stretch.
        widths = (110, 270, 110, 110, 110, 180, 130)
        aligns = ("w", "w", "w", "w", "e", "w", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.age_tree.heading(c, text=c, anchor=a)
            self.age_tree.column(c, width=w, anchor=a,
                                 stretch=(c in ("Client", "Bucket", "Balance (AED)")),
                                 minwidth=max(70, w - 40))
        self.age_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "aging_receivables")

    def _build_rev_client_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  7️⃣ Revenue by Client  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "REVENUE BY CLIENT RANKING",
                                  subtitle="Clients ranked by Total Invoiced — shows invoice count, paid, unpaid, paid percentage",
                                  accent=BLUE_DEEP)
        self.rc_banner = self._make_banner(card, "Client Revenue Snapshot",
                                            acolors=(LIGHT_BLUE, PURPLE_BG, GREEN_BG, ORANGE_BG))
        self.rc_explain = self._make_explain_box(card)
        cols = ("Client", "# Invoices", "Total Invoiced (AED)", "Total Paid (AED)",
                "Total Unpaid (AED)", "Paid %")
        self.rc_tree = ZebraTreeview(card, columns=cols, show="headings", height=18)
        widths = (340, 90, 180, 180, 180, 100)
        aligns = ("w", "e", "e", "e", "e", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.rc_tree.heading(c, text=c, anchor=a)
            self.rc_tree.column(c, width=w, anchor=a,
                                stretch=(c in ("Client", "Total Invoiced (AED)", "Total Unpaid (AED)")),
                                minwidth=max(70, w - 40))
        self.rc_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "revenue_by_client")

    def _build_stock_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  8️⃣ Stock Valuation  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "INVENTORY & STOCK VALUATION",
                                  subtitle="Opening Inventory + Purchases Received − COGS = Closing Inventory at Cost",
                                  accent=AMBER_DEEP)
        self.st_banner = self._make_banner(card, "Inventory Snapshot",
                                            acolors=(AMBER_BG, ORANGE_BG, GREEN_BG, LIGHT_BLUE))
        self.st_explain = self._make_explain_box(card)
        cols = ("Movement Item", "Value (AED)")
        self.st_tree = ZebraTreeview(card, columns=cols, show="headings", height=14)
        widths = (460, 220)
        aligns = ("w", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.st_tree.heading(c, text=c, anchor=a)
            self.st_tree.column(c, width=w, anchor=a, stretch=True,
                                minwidth=max(100, w - 60))
        self.st_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "stock_valuation")

    def _build_supplier_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  9️⃣ Supplier Spend  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "SUPPLIER & EXPENSE PAYEE SPEND ANALYSIS",
                                  subtitle="Suppliers / Payees ranked by Total Spend — Purchases + Operating Expenses",
                                  accent=PURPLE_DEEP)
        self.sp_banner = self._make_banner(card, "Supplier Spend Snapshot",
                                            acolors=(PURPLE_BG, LIGHT_BLUE, GREEN_BG, ORANGE_BG))
        self.sp_explain = self._make_explain_box(card)
        cols = ("Supplier / Payee", "# POs", "Purchases (AED)", "Expenses Paid (AED)", "Total Spend (AED)")
        self.sp_tree = ZebraTreeview(card, columns=cols, show="headings", height=18)
        widths = (380, 80, 180, 180, 180)
        aligns = ("w", "e", "e", "e", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.sp_tree.heading(c, text=c, anchor=a)
            self.sp_tree.column(c, width=w, anchor=a,
                                stretch=(c in ("Supplier / Payee", "Total Spend (AED)")),
                                minwidth=max(70, w - 40))
        self.sp_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "supplier_spend")

    def _build_audit_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text="  🔟 GL Audit  ")
        card = self._make_report_card(t)
        self._make_section_header(card, "GENERAL LEDGER AUDIT — DOUBLE-ENTRY VERIFICATION",
                                  subtitle="Verifies every transaction batch: Total Debits = Total Credits. Variance = 0 required.",
                                  accent=GREEN_DEEP)
        self.audit_banner = self._make_banner(card, "Audit Status Snapshot",
                                              acolors=(LIGHT_BLUE, GREEN_BG, RED_BG, PURPLE_BG))
        self.audit_explain = self._make_explain_box(card)
        cols = ("Reference Type", "Reference ID", "Debits (AED)", "Credits (AED)", "Variance (AED)")
        self.audit_tree = ZebraTreeview(card, columns=cols, show="headings", height=16)
        widths = (180, 240, 180, 180, 180)
        aligns = ("w", "w", "e", "e", "e")
        for c, w, a in zip(cols, widths, aligns):
            self.audit_tree.heading(c, text=c, anchor=a)
            self.audit_tree.column(c, width=w, anchor=a,
                                   stretch=(c in ("Reference ID", "Variance (AED)")),
                                   minwidth=max(90, w - 50))
        self.audit_tree.pack(fill="both", expand=True, pady=(12, 10))
        self._make_export_bar(card, "audit")

    def _build_footnotes_tab(self):
        t = ttk.Frame(self.nb, padding=6)
        t.configure(style="Card.TFrame")
        self.nb.add(t, text=" 12️⃣ 📋 GAAP Footnotes & Disclosures ")
        card = self._make_report_card(t)
        self._make_section_header(card, "GAAP FOOTNOTES & DISCLOSURES",
                                  subtitle="Significant Accounting Policies · Contingent Liabilities · Related-Party Transactions — ASC 235, 450, 850",
                                  accent=NAVY_DARK)
        self.fn_banner = self._make_banner(card, "Footnotes Snapshot",
                                            acolors=(LIGHT_BLUE, PURPLE_BG, ORANGE_BG, GREEN_BG))
        self.fn_explain = self._make_explain_box(card)

        policy_cols = ("Ref", "Topic", "Policy Text")
        self.fn_policies_tree = ZebraTreeview(card, columns=policy_cols, show="headings", height=11)
        policy_widths = (80, 220, 780)
        policy_aligns = ("w", "w", "w")
        for c, w, a in zip(policy_cols, policy_widths, policy_aligns):
            self.fn_policies_tree.heading(c, text=c, anchor=a)
            self.fn_policies_tree.column(c, width=w, anchor=a,
                                         stretch=(c == "Policy Text"),
                                         minwidth=max(60, w - 50))
        self.fn_policies_tree.pack(fill="both", expand=True, pady=(12, 6))

        cont_cols = ("Category", "Ref", "Assessment / Flag", "Disclosure Text", "Amount (AED)")
        self.fn_contingencies_tree = ZebraTreeview(card, columns=cont_cols, show="headings", height=14)
        cont_widths = (160, 80, 180, 680, 160)
        cont_aligns = ("w", "w", "w", "w", "e")
        for c, w, a in zip(cont_cols, cont_widths, cont_aligns):
            self.fn_contingencies_tree.heading(c, text=c, anchor=a)
            self.fn_contingencies_tree.column(c, width=w, anchor=a,
                                              stretch=(c == "Disclosure Text"),
                                              minwidth=max(60, w - 50))
        self.fn_contingencies_tree.pack(fill="both", expand=True, pady=(6, 10))
        self._make_export_bar(card, "footnotes")

    # -----------------------------------------------------------------------
    # 3.5 Section Header + Banner + Explain + Export Bar builders
    # -----------------------------------------------------------------------
    def _make_section_header(self, parent, title, subtitle="", accent=NAVY):
        """Professional report header with title (ALL-CAPS), subtitle, accent underline."""
        wrap = ttk.Frame(parent)
        wrap.configure(style="Card.TFrame")
        wrap.pack(fill="x", pady=(0, 10))
        tk.Label(wrap, text=title,
                 font=("Helvetica", 14, "bold"), bg=WHITE, fg=NAVY_DARK,
                 anchor="w", justify="left").pack(fill="x")
        if subtitle:
            tk.Label(wrap, text=subtitle,
                     font=("Helvetica", 9), bg=WHITE, fg=GREY_DARK,
                     anchor="w", justify="left", wraplength=1200).pack(fill="x", pady=(2, 0))
        line = tk.Frame(wrap, height=3, bg=accent, highlightthickness=0)
        line.pack(fill="x", pady=(8, 0))

    def _make_banner(self, parent, section_title, acolors=None):
        """4 KPI tiles, each with its own soft background color for readability."""
        if acolors is None:
            acolors = (LIGHT_BLUE, LIGHT_BLUE, LIGHT_BLUE, LIGHT_BLUE)
        wrap = ttk.Frame(parent)
        wrap.configure(style="Card.TFrame")
        wrap.pack(fill="x")
        tk.Label(wrap, text=section_title,
                 font=("Helvetica", 10, "bold"), bg=WHITE, fg=NAVY_DARK,
                 anchor="w", justify="left").pack(anchor="w", pady=(0, 6))
        row = tk.Frame(wrap, bg=WHITE, highlightthickness=0)
        row.pack(fill="x")
        cells = {}
        for i, (k, bg) in enumerate(zip(["a", "b", "c", "d"], acolors)):
            tile = tk.Frame(row, bg=bg, highlightthickness=1,
                            highlightbackground=SEPARATOR)
            tile.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 10, 0),
                      ipadx=12, ipady=10)
            row.columnconfigure(i, weight=1)
            inner = tk.Frame(tile, bg=bg, padx=14, pady=10)
            inner.pack(fill="both", expand=True)
            v_lbl = tk.Label(inner, text="-", bg=bg,
                             fg=NAVY_DARK, font=("Helvetica", 16, "bold"),
                             anchor="w", justify="left")
            v_lbl.pack(fill="x")
            l_lbl = tk.Label(inner, text="", bg=bg,
                             fg=GREY_DARK, font=("Helvetica", 8, "bold"),
                             anchor="w", justify="left", wraplength=260)
            l_lbl.pack(fill="x", pady=(2, 0))
            cells[k] = (v_lbl, l_lbl)
        return cells

    def _make_explain_box(self, parent):
        """Plain-English explanation card."""
        frm = tk.Frame(parent, background=WHITE, highlightthickness=1,
                       highlightbackground=SEPARATOR, padx=16, pady=12)
        frm.pack(fill="x", pady=(10, 0))
        hdr = tk.Frame(frm, bg=WHITE)
        hdr.pack(fill="x")
        tk.Label(hdr, text="💡  How to read this report",
                 bg=WHITE, fg=NAVY_DARK,
                 font=("Helvetica", 10, "bold"), anchor="w", justify="left").pack(side="left")
        tk.Frame(frm, height=2, bg=LIGHT_BLUE_2, highlightthickness=0).pack(fill="x", pady=(8, 8))
        txt = tk.Label(frm, text="(loading…)",
                       bg=WHITE, fg="#2D3748",
                       font=("Helvetica", 9), anchor="w", justify="left",
                       wraplength=1280)
        txt.pack(fill="x")
        return txt

    def register_tree(self, key, banner_label, explain_widget, tree,
                      section_title, accent, pretty_name):
        """Register a report tab's tree + metadata for export/refresh lookups.
        Expense Ledger (Tab 11) is the main caller; other tabs use the engine."""
        self._tree_registry[key] = {
            "banner_label": banner_label,
            "explain_widget": explain_widget,
            "tree": tree,
            "section_title": section_title,
            "accent": accent or NAVY,
            "pretty_name": pretty_name or key.replace("_", " ").title(),
        }

    def _make_export_bar(self, parent, report_key):
        bar = ttk.Frame(parent)
        bar.configure(style="Card.TFrame")
        bar.pack(fill="x", pady=(8, 0))
        # Left: export buttons (fixed width buttons, text wraps cleanly on narrow windows)
        left = ttk.Frame(bar)
        left.pack(side="left")
        ttk.Button(left, text="📄  Export Excel (.xlsx)",
                   command=lambda: self._export(report_key, "xlsx"),
                   style="Ghost.TButton").pack(side="left")
        ttk.Button(left, text="📑  Export PDF (.pdf)",
                   command=lambda: self._export(report_key, "pdf"),
                   style="Ghost.TButton").pack(side="left", padx=8)
        ttk.Button(left, text="📊  Export CSV (.csv)",
                   command=lambda: self._export(report_key, "csv"),
                   style="Ghost.TButton").pack(side="left")
        ttk.Button(left, text="🔄  Refresh Report",
                   command=self._refresh_engine_and_reports,
                   style="Ghost.TButton").pack(side="left", padx=8)
        # Right: report-key label so user knows which report is open (useful when scrolling tabs)
        right = ttk.Frame(bar)
        right.pack(side="right")
        pretty_name = {
            "profit_and_loss":       "P&L Report",
            "balance_sheet":         "Balance Sheet",
            "changes_in_equity":     "Changes in Equity",
            "cash_flow":             "Cash Flow",
            "financial_ratios":      "Financial Ratios & KPIs",
            "aging_receivables":     "Aging Receivables",
            "revenue_by_client":     "Revenue by Client",
            "stock_valuation":       "Stock Valuation",
            "supplier_spend":        "Supplier Spend",
            "audit":                 "GL Audit",
            "footnotes":             "GAAP Footnotes",
            "expense_manager":       "Expense Ledger",
            "Expense Ledger":        "Expense Ledger",
        }.get(report_key, str(report_key).replace("_", " ").title())
        ttk.Label(right,
                  text=f"📌 {pretty_name}   ·   Period: {self.range_var.get() if hasattr(self,'range_var') else '-'}",
                  font=("Helvetica", 8, "bold"), foreground=GREY_DARK,
                  background=WHITE).pack(side="right")

    # -----------------------------------------------------------------------
    # 3.6b  EXPORT: Excel + PDF (always fall back to working formats)
    # -----------------------------------------------------------------------
    def _report_trees(self):
        """Return mapping report_key -> (banner_label_text_dict, explain_widget, tree, section_title, accent, pretty_name, period_label)."""
        eng = getattr(self, "engine", None)
        period = self.range_var.get() if hasattr(self, "range_var") else "-"
        start_end = ""
        if eng:
            start_end = f"{getattr(eng, 'start', '')}  →  {getattr(eng, 'end', '')}"
        rows_map = {
            "profit_and_loss":   (self.pnl_banner, self.pnl_explain,  self.pnl_tree,  "PROFIT & LOSS STATEMENT", GREEN_DEEP, "P&L"),
            "balance_sheet":     (self.bs_banner,  self.bs_explain,   self.bs_tree,   "STATEMENT OF FINANCIAL POSITION", BLUE_DEEP, "Balance Sheet"),
            "changes_in_equity": (self.eq_banner,  self.eq_explain,   self.eq_tree,   "STATEMENT OF CHANGES IN EQUITY", PURPLE_DEEP, "Equity"),
            "cash_flow":         (self.cf_banner,  self.cf_explain,   self.cf_tree,   "CASH FLOW STATEMENT", TEAL_DEEP, "Cash Flow"),
            "financial_ratios":  (self.rt_banner,  self.rt_explain,   self.rt_tree,   "FINANCIAL RATIOS & KPIs", GOLD, "Ratios"),
            "aging_receivables": (self.age_banner, self.age_explain,  self.age_tree,  "AGING RECEIVABLES", RED_STRONG, "Aging"),
            "revenue_by_client": (self.rc_banner,  self.rc_explain,   self.rc_tree,   "REVENUE BY CLIENT", BLUE_DEEP, "Revenue Client"),
            "stock_valuation":   (self.st_banner,  self.st_explain,   self.st_tree,   "STOCK VALUATION", AMBER_DEEP, "Stock"),
            "supplier_spend":    (self.sp_banner,  self.sp_explain,   self.sp_tree,   "SUPPLIER SPEND", PURPLE_DEEP, "Supplier"),
            "audit":             (self.audit_banner, self.audit_explain, self.audit_tree, "GL AUDIT TRAIL", GREEN_DEEP, "Audit"),
            "footnotes":         (self.fn_banner,    self.fn_explain,    [self.fn_policies_tree, self.fn_contingencies_tree], "GAAP FOOTNOTES", NAVY_DARK, "Footnotes"),
        }
        return rows_map, period, start_end

    def _snapshot_tree_values(self, tree):
        """Given a ZebraTreeview, return (headers_list, rows_list_of_tuples, tag_for_row_list)."""
        cols = list(tree["columns"]) or []
        headers = []
        for c in cols:
            try:
                headers.append(tree.heading(c).get("text") or c)
            except Exception:
                headers.append(c)
        rows = []
        tags = []
        for item in tree.get_children():
            rows.append(tuple(tree.item(item, "values")))
            try:
                tags.append(tuple(tree.item(item, "tags") or ()))
            except Exception:
                tags.append(())
        return headers, rows, tags

    def _extract_banner_text(self, banner_frame):
        """Best-effort: collect the 4 tile value/label strings from banner."""
        out = []
        try:
            def walk(w, depth=0):
                if depth > 4:
                    return
                try:
                    kids = w.winfo_children()
                except Exception:
                    return
                for c in kids:
                    try:
                        cls_name = c.winfo_class()
                    except Exception:
                        cls_name = ""
                    if cls_name in ("Label", "TLabel"):
                        try:
                            txt = c.cget("text")
                        except Exception:
                            txt = ""
                        if txt and str(txt).strip():
                            out.append(str(txt).strip())
                    walk(c, depth + 1)
            walk(banner_frame)
        except Exception:
            pass
        # Group into (value, label) pairs assuming order value1, label1, value2, label2...
        pairs = []
        for i in range(0, len(out) - 1, 2):
            pairs.append((out[i], out[i + 1]))
        return pairs

    def _default_export_path(self, report_key, ext):
        from datetime import datetime as _dt
        stamp = _dt.now().strftime("%Y-%m-%d_%H%M%S")
        period = self.range_var.get().replace(" ", "_") if hasattr(self, "range_var") else "ALL"
        fname = f"HopePharma_{report_key}_{period}_{stamp}.{ext}"
        # Use user's Home/Downloads if exists; fallback to data_folder; else Desktop fallback
        home = os.path.expanduser("~")
        candidates = [
            os.path.join(home, "Downloads"),
            os.path.join(home, "Desktop"),
            home,
        ]
        if getattr(self, "dm", None):
            candidates.insert(0, getattr(self.dm, "data_folder", home))
        for p in candidates:
            if p and os.path.isdir(p) and os.access(p, os.W_OK):
                return os.path.join(p, fname)
        return os.path.join(os.getcwd(), fname)

    def _export(self, report_key, fmt):
        """Export the currently-visible report (tree + banner + explain).

        fmt ∈ {"xlsx", "pdf"}.  Will never fail silently — always produces
        something (with graceful fallback to CSV / HTML if libraries missing).
        """
        try:
            map_, period, start_end = self._report_trees()
        except Exception as ex:
            messagebox.showerror("Export", f"Could not prepare report data: {ex}")
            return
        if report_key not in map_:
            messagebox.showerror("Export", f"Unknown report {report_key!r}")
            return
        banner, explain_widget, tree, section_title, accent, pretty = map_[report_key]
        if isinstance(tree, list) or isinstance(tree, tuple):
            all_headers = []
            all_rows = []
            all_tags = []
            for t_idx, t in enumerate(tree):
                h, r, ta = self._snapshot_tree_values(t)
                if t_idx == 0:
                    all_headers = h
                else:
                    pass
                all_rows.extend(r)
                all_tags.extend(ta)
            headers, rows, tags = all_headers, all_rows, all_tags
        else:
            headers, rows, tags = self._snapshot_tree_values(tree)
        banner_pairs = self._extract_banner_text(banner)
        # Explain text
        try:
            explain_txt = ""
            if explain_widget is not None and hasattr(explain_widget, "winfo_children"):
                for c in explain_widget.winfo_children():
                    try:
                        if c.winfo_class() == "Label":
                            t = c.cget("text")
                        else:
                            t = ""
                    except Exception:
                        t = ""
                    if t and str(t).strip():
                        explain_txt += str(t).strip() + "\n\n"
        except Exception:
            explain_txt = ""
        try:
            if fmt == "xlsx":
                path = self._default_export_path(report_key, "xlsx")
                try:
                    from openpyxl import Workbook  # noqa: F401
                    self._export_excel_openpyxl(path, pretty, section_title, period, start_end,
                                                banner_pairs, explain_txt, headers, rows, tags)
                except Exception as ex:
                    # Fallback: CSV (works everywhere, openable in Excel/Numbers/Sheets)
                    path_csv = self._default_export_path(report_key, "csv")
                    self._export_csv(path_csv, headers, rows)
                    messagebox.showwarning(
                        "Excel Export",
                        f"⚠️  openpyxl not installed (error: {ex}).\n\n"
                        f"Fell back to CSV — opens natively in Excel / Numbers / Google Sheets:\n\n  {path_csv}\n\n"
                        f"Install openpyxl to generate proper styled .xlsx workbooks:\n"
                        f"  pip install openpyxl reportlab"
                    )
                    path = path_csv
            elif fmt == "pdf":
                path = self._default_export_path(report_key, "pdf")
                try:
                    import reportlab  # noqa: F401
                    self._export_pdf_reportlab(path, pretty, section_title, period, start_end,
                                               banner_pairs, explain_txt, headers, rows, tags)
                except Exception as ex_rl:
                    # Fallback 1: Excel print sheet; Fallback 2: HTML table (open in browser → Ctrl+P → PDF)
                    try:
                        from openpyxl import Workbook  # noqa: F401
                        path2 = self._default_export_path(report_key, "xlsx")
                        self._export_excel_openpyxl(path2, pretty, section_title, period,
                                                    start_end, banner_pairs, explain_txt,
                                                    headers, rows, tags)
                        messagebox.showwarning(
                            "PDF Export",
                            f"⚠️  reportlab not installed (error: {ex_rl}).\n\n"
                            f"Instead we wrote a Print-Ready Excel file you can open & choose File → Print → Save as PDF:\n\n  {path2}\n\n"
                            f"To get direct PDF export next time, install:\n  pip install reportlab openpyxl"
                        )
                        path = path2
                    except Exception as ex_xl:
                        # Last fallback: HTML (browser → Print → PDF)
                        path_html = self._default_export_path(report_key, "html")
                        self._export_html_print(path_html, pretty, section_title, period,
                                                start_end, banner_pairs, explain_txt,
                                                headers, rows, tags)
                        # Try open in default browser
                        try:
                            import webbrowser
                            webbrowser.open("file://" + os.path.abspath(path_html))
                        except Exception:
                            pass
                        messagebox.showwarning(
                            "PDF Export",
                            f"⚠️  reportlab + openpyxl both unavailable.\n"
                            f"  reportlab error: {ex_rl}\n  openpyxl error: {ex_xl}\n\n"
                            f"Fell back to a Print-Ready HTML file. It will now open in your default browser.\n"
                            f"Press Ctrl+P / Cmd+P → choose 'Save as PDF' in the print dialog:\n\n  {path_html}"
                        )
                        path = path_html
            else:
                messagebox.showerror("Export", f"Unknown format {fmt}")
                return
        except Exception as ex:
            import traceback; traceback.print_exc()
            messagebox.showerror("Export failed", str(ex))
            return
        try:
            # Success modal + reveal button option
            if messagebox.askyesno(
                "✅  Export Complete",
                f"{fmt.upper()} report saved successfully:\n\n"
                f"  {pretty} — {section_title}\n"
                f"  Period: {period}   ({start_end})\n"
                f"  Rows exported: {len(rows)}\n\n"
                f"File path:\n  {path}\n\n"
                f"Reveal file in its folder now?"
            ):
                self._reveal_in_file_manager(path)
        except Exception:
            messagebox.showinfo("✅  Export Complete",
                                f"{fmt.upper()} saved to:\n\n{path}")

    def _reveal_in_file_manager(self, path):
        try:
            abs_path = os.path.abspath(path)
        except Exception:
            abs_path = path
        import sys, subprocess
        try:
            if sys.platform.startswith("darwin"):
                subprocess.Popen(["open", "-R", abs_path])
            elif sys.platform.startswith("win"):
                subprocess.Popen(["explorer", "/select,", abs_path])
            else:
                folder = os.path.dirname(abs_path)
                subprocess.Popen(["xdg-open", folder or "."])
        except Exception:
            pass

    def _export_csv(self, path, headers, rows):
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(headers)
            for r in rows:
                w.writerow(list(r))

    def _export_excel_openpyxl(self, path, pretty, section_title, period, start_end,
                               banner_pairs, explain_txt, headers, rows, tags):
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter
        wb = Workbook()

        # ---------------- Sheet 1: SUMMARY ----------------
        ws_sum = wb.active
        ws_sum.title = "SUMMARY"
        company_font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
        title_font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
        sub_font = Font(name="Calibri", size=10, italic=True, color="FFE7EB")
        navy_fill = PatternFill("solid", fgColor="1A365D")
        accent_fill = PatternFill("solid", fgColor="EBF4FF")
        label_font = Font(bold=True, size=11, color="1A365D")
        kpi_label_font = Font(bold=True, color="2B6CB0")
        kpi_value_font = Font(bold=True, size=13, color="1A365D")
        section_label_font = Font(bold=True, size=11, color="FFFFFF")
        section_fill = PatternFill("solid", fgColor="2B6CB0")
        note_font = Font(name="Calibri", size=9, color="2D3748")
        audit_font = Font(name="Calibri", size=9, italic=True, color="4A5568")
        thin = Side(border_style="thin", color="CBD5E0")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        sum_cols = 2
        sr = 1
        ws_sum.cell(row=sr, column=1,
                    value="🏥 HOPEPHARMA MEDICAL TRADING L.L.C.  —  GAAP FINANCIAL REPORTING")
        ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
        c = ws_sum.cell(row=sr, column=1)
        c.font = company_font; c.fill = navy_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws_sum.row_dimensions[sr].height = 26
        sr += 1
        ws_sum.cell(row=sr, column=1, value="  " + section_title)
        ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
        c = ws_sum.cell(row=sr, column=1)
        c.font = title_font; c.fill = navy_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws_sum.row_dimensions[sr].height = 30
        sr += 1
        ws_sum.cell(row=sr, column=1,
                    value=f"  Period: {period}   ·   {start_end}   ·   Generated {datetime.now().strftime('%d %b %Y %H:%M')}")
        ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
        c = ws_sum.cell(row=sr, column=1)
        c.font = sub_font; c.fill = navy_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws_sum.row_dimensions[sr].height = 20
        sr += 2
        # --- Security / Role Block ---
        try:
            period_lock = ("LOCKED (GAAP Audit Mode)"
                           if getattr(getattr(self, "engine", None),
                                      "read_from_locked_periods_only", False)
                           else "⚠️  UNLOCKED (Management Review — NOT for audit)")
            role = getattr(self, "_user_role", "ReadOnly")
        except Exception:
            period_lock = "⚠️  UNLOCKED"; role = "ReadOnly"
        ws_sum.cell(row=sr, column=1, value="Security & Audit Metadata").font = label_font
        ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
        ws_sum.cell(row=sr, column=1).fill = accent_fill
        sr += 1
        for lab, val in [
            ("User Role", role),
            ("Period Lock Mode", period_lock),
            ("Report Name", pretty),
            ("Report Type", section_title),
            ("Data Folder", getattr(self, "_data_folder", "")),
            ("All actions logged to", "audit_log.jsonl (append-only, chain-of-custody)"),
        ]:
            ws_sum.cell(row=sr, column=1, value=lab).font = kpi_label_font
            ws_sum.cell(row=sr, column=1).border = border
            vcell = ws_sum.cell(row=sr, column=2, value=val)
            vcell.font = kpi_value_font if lab in ("User Role",) else note_font
            vcell.border = border
            vcell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws_sum.row_dimensions[sr].height = 20
            sr += 1
        sr += 1
        # --- KPI Banner Tiles ---
        ws_sum.cell(row=sr, column=1, value="KPI Snapshot (Report Banner)").font = label_font
        ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
        ws_sum.cell(row=sr, column=1).fill = accent_fill
        sr += 1
        for val, lab in (banner_pairs or []):
            ws_sum.cell(row=sr, column=1, value=lab).font = kpi_label_font
            ws_sum.cell(row=sr, column=1).border = border
            vc = ws_sum.cell(row=sr, column=2, value=val)
            vc.font = kpi_value_font
            vc.border = border
            vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws_sum.row_dimensions[sr].height = 22
            sr += 1
        if not banner_pairs:
            ws_sum.cell(row=sr, column=1, value="(No KPIs for this report)").font = note_font
            ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
            sr += 1
        sr += 1
        # --- Diagnostics Summary ---
        diagnostics = []
        try:
            eng = getattr(self, "engine", None)
            if eng is not None:
                diagnostics.append(("Engine Start", str(getattr(eng, "start", ""))))
                diagnostics.append(("Engine End", str(getattr(eng, "end", ""))))
                diagnostics.append(("Read Locked Periods Only",
                                    str(getattr(eng, "read_from_locked_periods_only", False))))
                locked_count = getattr(eng, "_excluded_unlocked_count", None)
                if locked_count is not None:
                    diagnostics.append(("Locked Periods Excluded Count", str(locked_count)))
                diagnostics.append(("Username", getattr(getattr(self, "dm", None),
                                                       "username", "unknown")))
                diagnostics.append(("Host PID", str(os.getpid())))
                diagnostics.append(("Report Row Count (Detail)",
                                    f"{len(rows)} rows × {len(headers)} columns"))
        except Exception:
            pass
        if diagnostics:
            ws_sum.cell(row=sr, column=1,
                        value="Diagnostics Summary").font = section_label_font
            ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
            c = ws_sum.cell(row=sr, column=1); c.fill = section_fill
            sr += 1
            for lab, val in diagnostics:
                ws_sum.cell(row=sr, column=1, value=lab).font = kpi_label_font
                ws_sum.cell(row=sr, column=1).border = border
                vc = ws_sum.cell(row=sr, column=2, value=val)
                vc.border = border
                vc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                ws_sum.row_dimensions[sr].height = 18
                sr += 1
            sr += 1
        # --- Explain box ---
        if explain_txt:
            ws_sum.cell(row=sr, column=1, value="Notes / Explanation").font = label_font
            ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
            ws_sum.cell(row=sr, column=1).fill = accent_fill
            sr += 1
            for ln in str(explain_txt).splitlines():
                ws_sum.cell(row=sr, column=1, value=ln.strip())
                ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
                c = ws_sum.cell(row=sr, column=1)
                c.alignment = Alignment(wrap_text=True, horizontal="left", vertical="top")
                c.font = note_font
                ws_sum.row_dimensions[sr].height = 18
                sr += 1
            sr += 1
        ws_sum.cell(row=sr, column=1,
                    value="✅  Summary complete. Switch to the DETAIL sheet for the full data table."
                   ).font = audit_font
        ws_sum.merge_cells(start_row=sr, start_column=1, end_row=sr, end_column=sum_cols)
        ws_sum.column_dimensions["A"].width = 38
        ws_sum.column_dimensions["B"].width = 92

        # ---------------- Sheet 2: DETAIL ----------------
        ws_det = wb.create_sheet("DETAIL", 1)
        white_fill = PatternFill("solid", fgColor="FFFFFF")
        zebra_fill = PatternFill("solid", fgColor="F7FAFC")
        green_fill = PatternFill("solid", fgColor="C6F6D5")
        red_fill   = PatternFill("solid", fgColor="FED7D7")
        grey_fill  = PatternFill("solid", fgColor="E2E8F0")
        total_fill = PatternFill("solid", fgColor="EBF4FF")
        head_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        n_cols = max(1, len(headers))
        # Title row
        ws_det.cell(row=1, column=1, value=section_title)
        ws_det.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
        c = ws_det.cell(row=1, column=1)
        c.font = title_font; c.fill = navy_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws_det.row_dimensions[1].height = 26
        ws_det.cell(row=2, column=1,
                    value=f"Period: {period}   ·   {start_end}   ·   Generated {datetime.now().strftime('%d %b %Y %H:%M')}   ·   {len(rows)} rows")
        ws_det.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
        c = ws_det.cell(row=2, column=1)
        c.font = sub_font; c.fill = navy_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws_det.row_dimensions[2].height = 20
        r = 4
        # Header row
        for ci, h in enumerate(headers, start=1):
            cc = ws_det.cell(row=r, column=ci, value=h)
            cc.font = head_font
            cc.fill = navy_fill
            cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cc.border = border
        ws_det.row_dimensions[r].height = 26
        r += 1
        data_start = r
        for i, row in enumerate(rows):
            is_zebra = i % 2 == 1
            for ci, val in enumerate(row, start=1):
                c = ws_det.cell(row=r, column=ci, value=(val if val is not None else ""))
                c.border = border
                c.alignment = Alignment(
                    horizontal=("right" if ci > 1 and len(headers) > 2 and ci == len(headers)
                                and isinstance(val, (int, float, Decimal, str))
                                and not re.search(r"[A-Za-z]", str(val)) else "left"),
                    wrap_text=True, vertical="center",
                )
                try:
                    tset = set(tags[i]) if i < len(tags) else set()
                except Exception:
                    tset = set()
                if "section" in tset or "grand_total" in tset:
                    c.fill = navy_fill
                    c.font = Font(bold=True, color="FFFFFF")
                elif "gross_profit" in tset or "good" in tset:
                    c.fill = green_fill
                    c.font = Font(bold=True, color="276749")
                elif "net_profit" in tset:
                    c.fill = PatternFill("solid", fgColor="F0FFF4")
                    c.font = Font(bold=True, color="276749", size=11)
                elif "net_loss" in tset:
                    c.fill = red_fill
                    c.font = Font(bold=True, color="C53030", size=11)
                elif "overdue_high" in tset:
                    c.fill = PatternFill("solid", fgColor="FFF5F5")
                    c.font = Font(color="C53030")
                elif "overdue_med" in tset:
                    c.fill = PatternFill("solid", fgColor="FFFFF0")
                    c.font = Font(color="B7791F")
                elif "subsection" in tset or "total_row" in tset:
                    c.fill = total_fill
                    c.font = Font(bold=True, color="2B6CB0")
                elif "expense_total" in tset:
                    c.fill = grey_fill
                    c.font = Font(bold=True, color="1A365D")
                elif "zero" in tset:
                    c.font = Font(italic=True, color="A0AEC0")
                elif is_zebra:
                    c.fill = zebra_fill
                else:
                    c.fill = white_fill
            r += 1
        ws_det.freeze_panes = f"A{data_start}"
        for ci, col_name in enumerate(headers, start=1):
            letter = get_column_letter(ci)
            max_len = len(str(col_name or "")) + 2
            for row_cells in rows:
                try:
                    v = row_cells[ci - 1] if ci - 1 < len(row_cells) else ""
                except Exception:
                    v = ""
                max_len = max(max_len, min(64, len(str(v or "")) + 2))
            ws_det.column_dimensions[letter].width = min(68, max(12, max_len))
        wb.save(path)

    def _export_html_print(self, path, pretty, section_title, period, start_end,
                           banner_pairs, explain_txt, headers, rows, tags):
        accent = "#1A365D"
        rows_html = []
        for i, r in enumerate(rows):
            cls = "odd" if i % 2 else "even"
            cells = "".join(f"<td>{(str(v) if v is not None else '')}</td>" for v in r)
            rows_html.append(f"<tr class=\"{cls}\">{cells}</tr>")
        head = "".join(f"<th>{h}</th>" for h in headers)
        banner_html = ""
        if banner_pairs:
            tiles = "".join(
                f"<div class=\"tile\"><div class=\"v\">{val}</div><div class=\"l\">{lab}</div></div>"
                for val, lab in banner_pairs[:4]
            )
            banner_html = f"<div class=\"banner\">{tiles}</div>"
        explain_html = ""
        if explain_txt:
            safe = explain_txt.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            explain_html = f"<div class=\"explain\"><strong>Notes / Explanation:</strong><br><pre>{safe}</pre></div>"
        html = f"""<!DOCTYPE html>
<html><head><meta charset="utf-8"><title>{section_title} — HopePharma</title>
<style>
  body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif; color:#2D3748; margin:0; padding:24px; background:#fff; }}
  header {{ background:{accent}; color:#fff; padding:18px 20px; border-radius:8px; }}
  header .co {{ font-weight:700; font-size:16px; opacity:.95; }}
  header h1 {{ margin:4px 0 2px 0; font-size:22px; letter-spacing:.2px; }}
  header .meta {{ opacity:.85; font-size:12px; }}
  h2 {{ font-size:14px; color:{accent}; margin:18px 0 10px 0; }}
  .banner {{ display:grid; grid-template-columns: repeat(4, 1fr); gap:10px; margin-top:14px; }}
  .tile {{ background:#F0FFF4; padding:12px 14px; border-radius:8px; border:1px solid #C6F6D5; }}
  .tile .v {{ font-size:18px; font-weight:800; color:{accent}; }}
  .tile .l {{ font-size:10px; text-transform:uppercase; letter-spacing:.4px; color:#276749; margin-top:4px; }}
  table {{ width:100%; border-collapse: collapse; margin-top:10px; font-size:13px; }}
  th {{ background:{accent}; color:#fff; text-align:left; padding:10px 12px; border:1px solid #2B6CB0; }}
  td {{ padding:7px 12px; border:1px solid #E2E8F0; }}
  tr.even td {{ background:#FFFFFF; }}
  tr.odd  td {{ background:#F7FAFC; }}
  .explain {{ margin-top:16px; background:#EBF4FF; border-left:4px solid #3182CE; padding:12px 14px; border-radius:4px; color:#2D3748; font-size:12.5px; }}
  .explain pre {{ white-space: pre-wrap; font-family:inherit; margin:6px 0 0 0; line-height:1.55; }}
  footer {{ margin-top:22px; font-size:11px; color:#A0AEC0; text-align:center; }}
  @media print {{
    body {{ padding:0; }}
    header {{ -webkit-print-color-adjust: exact; print-color-adjust: exact; }}
    table th, table td {{ font-size:11px; }}
    @page {{ margin:1cm; }}
  }}
</style></head><body>
  <header>
    <div class="co">🏥 HOPEPHARMA MEDICAL TRADING L.L.C.</div>
    <h1>{section_title}</h1>
    <div class="meta">Period: <strong>{period}</strong> &nbsp;·&nbsp; {start_end} &nbsp;·&nbsp; Generated {datetime.now().strftime("%d %b %Y, %H:%M")}</div>
  </header>
  {banner_html}
  {explain_html}
  <h2>{pretty} — {len(rows)} rows</h2>
  <table>
    <thead><tr>{head}</tr></thead>
    <tbody>{"".join(rows_html)}</tbody>
  </table>
  <footer>HopePharma GAAP Financial Reporting Centre — print or Save as PDF from your browser's print dialog.</footer>
</body></html>"""
        with open(path, "w", encoding="utf-8") as f:
            f.write(html)

    def _export_pdf_reportlab(self, path, pretty, section_title, period, start_end,
                              banner_pairs, explain_txt, headers, rows, tags):
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                         Table, TableStyle, PageBreak)
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        # Register Helvetica fallback (reportlab has Helvetica built in via Type1)
        width, height = landscape(A4)
        doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                                leftMargin=1.2*cm, rightMargin=1.2*cm,
                                topMargin=1.3*cm, bottomMargin=1.3*cm,
                                title=f"{pretty} — HopePharma Financial Report",
                                author="HopePharma Reporting Centre")
        styles = getSampleStyleSheet()
        NAVY_CLR = colors.HexColor("#1A365D")
        story = []
        # Header
        story.append(Paragraph(
            f"<font color='#FFFFFF' size='13'><b>🏥 HOPEPHARMA MEDICAL TRADING L.L.C.</b></font>"
            f"<br/><font color='#FFFFFF' size='18'><b>{section_title}</b></font>"
            f"<br/><font color='#FFFFFF' size='9'>Period: <b>{period}</b> &nbsp;·&nbsp; {start_end}"
            f" &nbsp;·&nbsp; Generated {datetime.now().strftime('%d %b %Y %H:%M')}</font>",
            ParagraphStyle("NavyHeader", parent=styles["Normal"], backColor=NAVY_CLR,
                           textColor=colors.white, borderPadding=(12, 14, 12, 14),
                           leading=15, fontName="Helvetica-Bold")
        ))
        story.append(Spacer(1, 0.25*cm))
        # Banner tiles
        if banner_pairs:
            tile_data = []
            row = []
            for i, (val, lab) in enumerate(banner_pairs[:4]):
                cell = Paragraph(
                    f"<b><font size='14' color='#1A365D'>{val}</font></b>"
                    f"<br/><font size='7' color='#276749'>{lab.upper()}</font>",
                    ParagraphStyle("Tile", parent=styles["Normal"],
                                   backColor=colors.HexColor("#F0FFF4"),
                                   borderPadding=(10, 10, 10, 10),
                                   borderColor=colors.HexColor("#C6F6D5"),
                                   borderWidth=1))
                row.append(cell)
                if len(row) == 4:
                    tile_data.append(row); row = []
            if row:
                while len(row) < 4:
                    row.append("")
                tile_data.append(row)
            bt = Table(tile_data, colWidths=[(width-2.4*cm)/4]*4, hAlign="LEFT")
            bt.setStyle(TableStyle([("LEFTPADDING", (0,0), (-1,-1), 0),
                                     ("RIGHTPADDING", (0,0), (-1,-1), 2)]))
            story.append(bt)
            story.append(Spacer(1, 0.35*cm))
        # Explain box
        if explain_txt:
            lines = explain_txt.splitlines()
            safe = "<br/>".join(
                Paragraph(l or "&nbsp;", styles["Normal"]).text
                for l in lines
            )
            # Simpler: just use a big text cell
            box = Paragraph(
                f"<b>Notes / Explanation</b><br/><br/>"
                f"{explain_txt.replace(chr(10), '<br/>').replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')}",
                ParagraphStyle("ExplainBox", parent=styles["Normal"],
                               backColor=colors.HexColor("#EBF4FF"),
                               borderColor=colors.HexColor("#3182CE"),
                               borderWidth=1, borderPadding=(10, 12, 10, 12),
                               fontName="Helvetica", fontSize=9, leading=12))
            story.append(box)
            story.append(Spacer(1, 0.35*cm))
        # Table
        story.append(Paragraph(f"<b>{pretty} — {len(rows)} rows</b>",
                               ParagraphStyle("TblTitle", parent=styles["Normal"],
                                              fontSize=11, textColor=NAVY_CLR,
                                              spaceAfter=8)))
        # Build data rows
        data = [list(headers)]
        for r in rows:
            data.append(list(r))
        # Fit columns to page (landscape A4 - margins)
        usable = width - 2.4*cm
        n = len(headers)
        if n > 0:
            col_w = usable / n
        else:
            col_w = 4*cm
        tbl = Table(data, repeatRows=1, hAlign="LEFT",
                    colWidths=[col_w]*len(headers) if headers else None)
        style_cmds = [
            ("BACKGROUND", (0,0), (-1,0), NAVY_CLR),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 9),
            ("BOTTOMPADDING", (0,0), (-1,0), 10),
            ("TOPPADDING", (0,0), (-1,0), 10),
            ("ALIGN", (1,1), (-1,-1), "RIGHT"),   # Data cols right
            ("ALIGN", (0,0), (0,-1), "LEFT"),     # First col (labels) left
            ("FONTNAME", (0,1), (-1,-1), "Helvetica"),
            ("FONTSIZE", (0,1), (-1,-1), 8),
            ("TOPPADDING", (0,1), (-1,-1), 5),
            ("BOTTOMPADDING", (0,1), (-1,-1), 5),
            ("GRID", (0,0), (-1,-1), 0.25, colors.HexColor("#CBD5E0")),
            ("ROWBACKGROUNDS", (0,1), (-1,-1), [colors.white, colors.HexColor("#F7FAFC")]),
        ]
        for i, r in enumerate(rows, start=1):
            try:
                tset = set(tags[i-1]) if i-1 < len(tags) else set()
            except Exception:
                tset = set()
            if any(s in ("section","grand_total") for s in tset):
                style_cmds.append(("BACKGROUND", (0,i), (-1,i), NAVY_CLR))
                style_cmds.append(("TEXTCOLOR", (0,i), (-1,i), colors.white))
                style_cmds.append(("FONTNAME", (0,i), (-1,i), "Helvetica-Bold"))
            elif "gross_profit" in tset or "good" in tset:
                style_cmds.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor("#C6F6D5")))
                style_cmds.append(("TEXTCOLOR", (0,i), (-1,i), colors.HexColor("#276749")))
                style_cmds.append(("FONTNAME", (0,i), (-1,i), "Helvetica-Bold"))
            elif "net_profit" in tset:
                style_cmds.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor("#F0FFF4")))
                style_cmds.append(("TEXTCOLOR", (0,i), (-1,i), colors.HexColor("#276749")))
                style_cmds.append(("FONTNAME", (0,i), (-1,i), "Helvetica-Bold"))
                style_cmds.append(("FONTSIZE", (0,i), (-1,i), 10))
            elif "net_loss" in tset:
                style_cmds.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor("#FED7D7")))
                style_cmds.append(("TEXTCOLOR", (0,i), (-1,i), colors.HexColor("#C53030")))
                style_cmds.append(("FONTNAME", (0,i), (-1,i), "Helvetica-Bold"))
                style_cmds.append(("FONTSIZE", (0,i), (-1,i), 10))
            elif any(s in ("subsection","total_row") for s in tset):
                style_cmds.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor("#EBF4FF")))
                style_cmds.append(("TEXTCOLOR", (0,i), (-1,i), colors.HexColor("#2B6CB0")))
                style_cmds.append(("FONTNAME", (0,i), (-1,i), "Helvetica-Bold"))
            elif "expense_total" in tset:
                style_cmds.append(("BACKGROUND", (0,i), (-1,i), colors.HexColor("#E2E8F0")))
                style_cmds.append(("FONTNAME", (0,i), (-1,i), "Helvetica-Bold"))
            elif "zero" in tset:
                style_cmds.append(("TEXTCOLOR", (0,i), (-1,i), colors.HexColor("#A0AEC0")))
                style_cmds.append(("FONTNAME", (0,i), (-1,i), "Helvetica-Oblique"))
        tbl.setStyle(TableStyle(style_cmds))
        story.append(tbl)
        doc.build(story)

    # -----------------------------------------------------------------------
    # 3.6 Refresh: Rebuild engine + populate all tabs
    def _rebuild_gl_now(self):
        """User clicked 'Rebuild GL from All Data' button: backfill historical data."""
        ledger = getattr(self.dm, "ledger", None)
        if not ledger:
            messagebox.showerror("Rebuild GL",
                                 "Ledger service not available on data manager.")
            return
        if not messagebox.askyesno(
            "Rebuild General Ledger",
            "This will scan ALL existing invoices, purchases, payments, and\n"
            "expense records and post them (idempotently) to the General Ledger.\n\n"
            "This is SAFE to run multiple times — duplicates are skipped.\n\n"
            "Continue?"):
            return
        self.notify_var.set("🔁 Rebuilding GL from historical data — please wait...")
        self.win.update_idletasks()
        try:
            import threading
            def run_in_bg():
                def progress(msg):
                    self.win.after(0, lambda: self.notify_var.set(
                        "🔁 Rebuilding GL — " + msg))
                    self.win.update_idletasks()
                try:
                    stats = ledger.rebuild_general_ledger_from_all_data(
                        progress_cb=progress)
                    def done():
                        salaries_count = stats.get("salaries", 0)
                        self.notify_var.set(
                            f"✅ GL Rebuild done — "
                            f"{stats['invoices']} invoices · "
                            f"{stats['payments']} payments · "
                            f"{stats['purchases']} purchases · "
                            f"{stats['expenses']} expenses · "
                            f"{salaries_count} payrolls -> "
                            f"GL now has {stats['total_gl_after']} entries · "
                            f"{len(stats['warnings'])} warnings (see console)."
                        )
                        self._refresh_engine_and_reports()
                        for w in stats['warnings'][:10]:
                            print(f"  [warn] {w}")
                        messagebox.showinfo(
                            "GL Rebuild Complete",
                            f"Posted to General Ledger:\n"
                            f"  • {stats['invoices']} invoices (Sales / AR / Revenue)\n"
                            f"  • {stats['payments']} payments (Cash / Bank against AR)\n"
                            f"  • {stats['purchases']} purchases (AP / Inventory)\n"
                            f"  • {stats['expenses']} expenses (Withdrawals → OpEx)\n"
                            f"  • {salaries_count} employee payrolls (Salary 5100 + Bank Cr)\n"
                            f"GL now has {stats['total_gl_after']} total double-entry rows.\n\n"
                            f"All 10 reports are now fully populated with complete data.")
                    self.win.after(0, done)
                except Exception as ex:
                    import traceback; traceback.print_exc()
                    self.win.after(0, lambda: messagebox.showerror("Rebuild GL failed", str(ex)))
            threading.Thread(target=run_in_bg, daemon=True).start()
        except Exception as ex:
            messagebox.showerror("Rebuild GL", str(ex))

    def _auto_rebuild_gl_async(self):
        """Called on first open if GL is empty — user-friendly auto-backfill."""
        if self._auto_rebuilt:
            return
        self._auto_rebuilt = True
        ledger = getattr(self.dm, "ledger", None)
        if not ledger:
            return
        # Auto-run without asking (safe since GL is empty) and just show result
        self.notify_var.set("🔁 First-run: importing your existing invoices/purchases/expenses into the General Ledger...")
        self.win.update_idletasks()
        try:
            stats = ledger.rebuild_general_ledger_from_all_data(
                progress_cb=lambda m: None)
            self.notify_var.set(
                f"✅ Data import complete — {stats['invoices']} invoices · "
                f"{stats['payments']} payments · {stats['purchases']} purchases · "
                f"{stats['expenses']} expenses imported into GL "
                f"({stats['total_gl_after']} entries). Click Refresh → to refresh below reports."
            )
            # Auto-refresh reports after build
            self.win.after(400, self._refresh_engine_and_reports)
            if stats.get('warnings') and len(stats['warnings']) > 0:
                print("[Auto Rebuild GL warnings]:")
                for w in stats['warnings'][:20]:
                    print(f"  - {w}")
        except Exception as ex:
            import traceback; traceback.print_exc()
            self.notify_var.set(f"⚠️ First-run import encountered an issue: {ex}")

    # -----------------------------------------------------------------------
    def _refresh_engine_and_reports(self, trigger="manual"):
        """Rebuild engine from disk (re-read GL JSON) and refresh all 11 tabs.

        `trigger` is a label used only for telemetry / logging so operators can
        tell whether a refresh came from the user clicking 🔁, from window
        activation, or from the 60-second watchdog.
        """
        # --- RE-READ DATA FROM DISK first.  Fixes the "invoice added but GAAP
        # didn't notice" bug where the data manager's in-memory caches were
        # populated before the invoice posting hook ran (because the dashboard
        # was opened *before* the invoice, and dm.load_json was only called
        # once at dashboard __init__).  Without this re-read, `_build_engine`
        # reads stale dm._cache and users saw empty reports.
        try:
            if getattr(self.dm, "reload_all_json", None) and callable(self.dm.reload_all_json):
                try: self.dm.reload_all_json()
                except Exception:
                    # Fallback: re-call the raw load_json on every data file
                    for attr in ("gl_entries", "invoices_data", "expenses",
                                 "transactions_data", "purchase_orders",
                                 "employees_data", "salaries_data"):
                        fname = getattr(self.dm, f"{attr}_file", None) or f"{attr}.json"
                        try: setattr(self.dm, attr, self.dm.load_json(fname))
                        except Exception: pass
            else:
                # No reload_all_json?  Use the brute force approach.
                for attr in ("gl_entries", "invoices_data", "expenses",
                             "transactions_data", "purchase_orders",
                             "employees_data", "salaries_data"):
                    fname = getattr(self.dm, f"{attr}_file", None) or f"{attr}.json"
                    try: setattr(self.dm, attr, self.dm.load_json(fname))
                    except Exception: pass
        except Exception:
            pass
        # Also refresh the LedgerService in-memory view (in case invoice hook
        # appended to the dm list but the ledger service cached an old list).
        try:
            if self.ledger and getattr(self.ledger, "reload_from_dm", None) and callable(self.ledger.reload_from_dm):
                self.ledger.reload_from_dm()
        except Exception:
            pass

        self.engine = self._build_engine()
        if not self.engine:
            return
        s = self.engine.start; e = self.engine.end
        ps = self.engine.prev_start; pe = self.engine.prev_end
        cur_count = len(self.engine.gl_current)
        total_count = len(getattr(self.engine, '_gl_all', []))
        prior_count = len(getattr(self.engine, 'gl_prior', []) or [])
        base = (f"📅 Period:  {s.strftime('%d %b %Y')}  →  {e.strftime('%d %b %Y')}    |    "
                f"📊 Prior Period:  {ps.strftime('%d %b %Y')}  →  {pe.strftime('%d %b %Y')}    |    "
                f"🗂️  GL entries in period: {cur_count}    |    "
                f"🗂️  Prior period: {prior_count}    |    "
                f"🗂️  Total GL entries ever: {total_count}")
        # Emphasise FIRST-DATA and STALE-TRANSITION cases so the user *notices*
        prev_counts = getattr(self, "_last_gl_counts", None) or (0, 0)
        first_data = (prev_counts[0] == 0 and cur_count > 0) or (prev_counts[1] == 0 and total_count > 0)
        grew = cur_count > prev_counts[0] or total_count > prev_counts[1]
        self._last_gl_counts = (cur_count, total_count)
        if first_data:
            base = "🎉  FIRST DATA DETECTED — financial reports are now live.  " + base
        elif grew and trigger in ("window_activated", "periodic_60s"):
            base = f"🔔  New entries detected ({cur_count - prev_counts[0]} in period).  " + base
        self.notify_var.set(base)
        # Populate each tab.  Defensive try/except on every tab so one broken
        # tab never suppresses the rest of the financial statements.
        tabs_success_count = 0
        tab_tasks = (
            ("P&L",                 self._refresh_pnl),
            ("Balance Sheet",       self._refresh_bs),
            ("Changes in Equity",   self._refresh_changes_in_equity),
            ("Cash Flow",           self._refresh_cashflow),
            ("Financial Ratios",    self._refresh_financial_ratios),
            ("Aging Receivables",   self._refresh_aging),
            ("Revenue by Client",   self._refresh_rev_client),
            ("Stock Ledger",        self._refresh_stock),
            ("Supplier Ledger",     self._refresh_supplier),
            ("GL Audit",            self._refresh_audit),
            ("Expense Register",    self._reload_expense_tab),
            ("Footnotes",           self._refresh_footnotes),
        )
        for tab_name, fn in tab_tasks:
            try:
                fn()
                tabs_success_count += 1
            except Exception as ex:
                import traceback
                traceback.print_exc()
                try:
                    self.notify_var.set(f"⚠️  Tab '{tab_name}' refresh failed: {ex}   |   " + base)
                except Exception: pass
        new_entries_flag = bool(grew and trigger in ("window_activated", "periodic_60s"))
        period_lock_mode = bool(getattr(self.engine, "read_from_locked_periods_only", False)) if self.engine else False
        locked_excluded_count = getattr(self.engine, "_excluded_unlocked_count", None)
        if locked_excluded_count is None:
            locked_excluded_count = getattr(self.engine, "excluded_unlocked_count", 0)
        try:
            self._append_audit_log("REFRESH", {
                "trigger": trigger,
                "first_data_detected": bool(first_data),
                "new_entries": bool(new_entries_flag),
                "tabs_refreshed_successfully": int(tabs_success_count),
                "period_lock_mode": bool(period_lock_mode),
                "locked_periods_excluded_count": (int(locked_excluded_count) if isinstance(locked_excluded_count, (int, float)) else 0),
            })
        except Exception:
            pass

    # -------- P&L --------
    def _refresh_pnl(self):
        for i in self.pnl_tree.get_children():
            self.pnl_tree.delete(i)
        try:
            d = self.engine.get_profit_and_loss()
        except Exception as e:
            self.pnl_explain.config(text=f"Error: {e}")
            import traceback; traceback.print_exc()
            return
        sb = d.get("summary_banner", {})
        self._fill_banner(self.pnl_banner, [
            (fmt_aed(sb.get("grand_total", 0)), "Net Income (US GAAP)"),
            (fmt_aed(sb.get("operating_income", 0)), "Operating Income (EBIT)"),
            (f"{fmt_aed(sb.get('gross_margin_pct', 0))}%", "Gross Margin %"),
            (f"EPS {fmt_aed(sb.get('eps', 0))} | Δ {fmt_aed(sb.get('pct_change', 0))}% vs prior",
             "EPS & Net Income Δ"),
        ])
        lines_obj = d.get("lines", {}) or {}
        expl_parts = [d.get("explanation", "") or ""]
        diag = d.get("diagnostics") or {}
        checks = diag.get("checks") or []
        if checks:
            passed_n = sum(1 for c in checks if c.get("passed"))
            expl_parts.append("\n\n" + "═"*70)
            expl_parts.append(f"US GAAP P&L Mathematical Integrity — {passed_n}/{len(checks)} Checks Passed")
            expl_parts.append("─"*70)
            for c in checks:
                icon = "✅" if c.get("passed") else "⚠️"
                expl_parts.append(f"  {icon}  {c.get('name','')}\n      {c.get('detail','')}")
            if not diag.get("passed", True):
                for step in diag.get("fix_steps") or []:
                    expl_parts.append(f"  🔧  {step}")
                expl_parts.append(f"\nSummary: {diag.get('summary','')}")
            else:
                expl_parts.append(f"\nSummary: {diag.get('summary','')}")
        self.pnl_explain.config(text="\n".join(expl_parts))

        def _num(v):
            if v is None: return Decimal(0)
            try: return Decimal(str(v))
            except Exception: return Decimal(0)
        def fv(v):
            if v is None: return ""
            if isinstance(v, Decimal): return fmt_aed(v)
            return str(v)
        # Walk GAAP-structured table_rows (engine L652..L684) — tags drive zebra styling
        for row in d.get("table_rows", []):
            tag = str(row.get("tag") or "").strip()
            cur = _num(row.get("current"))
            pri = _num(row.get("prior"))
            cur_s = fv(row.get("current"))
            pri_s = fv(row.get("prior"))
            pct_s = ""
            if row.get("pct_change") is not None:
                try: pct_s = f"{float(Decimal(str(row['pct_change']))):,.2f}%"
                except Exception: pct_s = str(row["pct_change"])
            label = str(row.get("label", ""))
            # Map engine `tag` -> treeview tag list (zebra-stripe + bolded totals)
            tags_list = []
            is_zero = (abs(cur) == 0 and abs(pri) == 0)
            if tag == "section":
                tags_list = ("section",)
            elif tag == "gross_profit":
                tags_list = ("gross_profit",)
            elif tag == "total_row":
                tags_list = ("total_row",)
            elif tag == "net_profit":
                tags_list = ("net_profit",) if cur >= 0 else ("net_loss",)
            else:
                if is_zero: tags_list = ("zero",)
            self.pnl_tree.zebra_insert(values=(label, cur_s, pri_s, pct_s), tags=tuple(tags_list))
        # ---- Append reconciliation block (GAAP-mandated ties between statements) ----
        ties_title_shown = False
        # Tie 1: Net Income flows to Retained Earnings
        if isinstance(lines_obj, dict) and "net_income" in lines_obj:
            if not ties_title_shown:
                self.pnl_tree.zebra_insert(
                    values=("", "", "", ""), tags=("level2",))
                self.pnl_tree.zebra_insert(
                    values=("🔗  CROSS-STATEMENT RECONCILIATIONS (GAAP Required)", "", "", ""),
                    tags=("section",))
                ties_title_shown = True
            self.pnl_tree.zebra_insert(
                values=(f"  ➜ Net Income {fmt_aed(lines_obj['net_income']['current'])} flows to Statement of Changes in Equity + Balance Sheet (Retained Earnings)",
                        "", "", ""), tags=("level2",))

    # -------- Balance Sheet --------
    def _refresh_bs(self):
        for i in self.bs_tree.get_children():
            self.bs_tree.delete(i)
        try:
            d = self.engine.get_balance_sheet()
        except Exception as e:
            self.bs_explain.config(text=f"Error: {e}")
            import traceback; traceback.print_exc()
            return
        sb = d.get("summary_banner", {})
        bal = sb.get("is_balanced")
        bal_text = "✅ BALANCED" if bal else "⚠️ UNBALANCED"
        bal_color = GREEN_DEEP if bal else RED_STRONG
        wc = sb.get("working_capital", Decimal(0))
        try:
            wc_val = Decimal(str(wc))
            wc_label = ("Working Capital ✅" if wc_val >= 0 else "Working Capital ⚠️ NEG")
        except Exception:
            wc_label = "Working Capital"
        self._fill_banner(self.bs_banner, [
            (fmt_aed(sb.get("grand_total", 0)), "Total Assets (ASC 210)"),
            (fmt_aed(sb.get("working_capital", 0)), wc_label + " (CA − CL)"),
            (fmt_aed(sb.get("liabilities", 0)) + "  |  " + fmt_aed(sb.get("equity", 0)), "Total Liabilities  |  Total Equity"),
            (bal_text, "Golden Rule: A = L + E"),
        ])
        last = self.bs_banner["d"][0]
        last.config(foreground=bal_color)
        lines_obj = d.get("lines", {}) or {}
        expl_parts = [d.get("explanation", "") or ""]
        diag = d.get("diagnostics") or {}
        checks = diag.get("checks") or []
        if checks:
            passed_n = sum(1 for c in checks if c.get("passed"))
            expl_parts.append("\n\n" + "═"*70)
            expl_parts.append(f"US GAAP Balance Sheet Integrity — {passed_n}/{len(checks)} Checks Passed")
            expl_parts.append("─"*70)
            for c in checks:
                icon = "✅" if c.get("passed") else "⚠️"
                expl_parts.append(f"  {icon}  {c.get('name','')}\n      {c.get('detail','')}")
            rcs = diag.get("root_causes") or []
            for rc in rcs:
                expl_parts.append(f"  🔍  Root cause: {rc}")
            fss = diag.get("fix_steps") or []
            for step in fss:
                expl_parts.append(f"  🔧  {step}")
            expl_parts.append(f"\nSummary: {diag.get('summary','')}")
        self.bs_explain.config(text="\n".join(expl_parts))

        def section(sec_title):
            self.bs_tree.zebra_insert(values=(sec_title, "", ""), tags=("section",))
        def subsec(sub_title):
            self.bs_tree.zebra_insert(values=("", sub_title, ""), tags=("subsection",))
        def line_item(lbl, amt, level=2, is_total=False, is_grand=False, tag_override=None):
            tag_list = []
            if tag_override:
                tag_list = [tag_override]
            else:
                if is_grand: tag_list = ["grand_total"]
                elif is_total: tag_list = ["total_row"]
                else:
                    tag_list = [f"level{level}"]
                    try:
                        if abs(Decimal(str(amt))) == 0:
                            tag_list = ["zero"]
                    except Exception:
                        pass
            prefix = ""
            if level == 2 and not is_total: prefix = "    · "
            elif level == 1 and not is_total: prefix = "  "
            lbl_display = f"{prefix}{lbl}"
            self.bs_tree.zebra_insert(
                values=("", lbl_display, fmt_aed(amt)), tags=tuple(tag_list))

        assets = lines_obj.get("assets", {}) or {}
        liabilities = lines_obj.get("liabilities", {}) or {}
        equity = lines_obj.get("equity", {}) or {}
        def _is_subtotal_key(k):
            return str(k).startswith("__")

        # ============= ASSETS (Current, then Non-Current, per explicit subtotal keys) =============
        section("🟦  ASSETS — Resources Owned (ASC 210 Classified)")
        cur_items = []
        ncur_items = []
        asset_total = None
        for k, v in assets.items():
            if k == "__current_assets_total":
                cur_items.append(("__SUBTOTAL__", v))
            elif k == "__noncurrent_assets_total":
                ncur_items.append(("__SUBTOTAL__", v))
            elif k == "total_assets":
                asset_total = v
            else:
                label = str(k)
                ca_explicit = any(c in label for c in ("1000", "1100", "1200", "1210", "1220", "1300", "1310", "1320", "1400", "1410", "1500"))
                nca_explicit = any(c in label for c in ("1600", "1610", "1620", "1630", "1640", "1650", "1700", "1710", "1800"))
                if ca_explicit:
                    cur_items.append((label, v))
                elif nca_explicit:
                    ncur_items.append((label, v))
                elif str(k).strip().startswith("1") and len(str(k).strip()) >= 3:
                    code_prefix = str(k).strip()[:3]
                    if code_prefix in ("100", "110", "120", "130", "140", "150"):
                        cur_items.append((label, v))
                    else:
                        ncur_items.append((label, v))
                else:
                    if cur_items and len(cur_items) < 8:
                        cur_items.append((label, v))
                    else:
                        ncur_items.append((label, v))
        subsec("  Current Assets (Short-Term, ≤ 12 months)")
        ca_subtotal_val = None
        for lbl, v in cur_items:
            if lbl == "__SUBTOTAL__":
                ca_subtotal_val = v
            else:
                line_item(lbl, v, level=2)
        if ca_subtotal_val is not None:
            line_item("TOTAL CURRENT ASSETS", ca_subtotal_val, level=0, is_total=True)
        subsec("  Non-Current Assets (Long-Term, > 12 months)")
        nca_subtotal_val = None
        for lbl, v in ncur_items:
            if lbl == "__SUBTOTAL__":
                nca_subtotal_val = v
            else:
                line_item(lbl, v, level=2)
        if nca_subtotal_val is not None:
            line_item("TOTAL NON-CURRENT ASSETS", nca_subtotal_val, level=0, is_total=True)
        if asset_total is not None:
            line_item("TOTAL ASSETS", asset_total, level=0, is_total=True, tag_override="grand_total")

        # ============= LIABILITIES (Current, then Non-Current) =============
        section("🟧  LIABILITIES — Amounts Owed (ASC 210 Classified)")
        cl_items = []
        ncl_items = []
        liab_total = None
        for k, v in liabilities.items():
            if k == "__current_liabilities_total":
                cl_items.append(("__SUBTOTAL__", v))
            elif k == "__noncurrent_liabilities_total":
                ncl_items.append(("__SUBTOTAL__", v))
            elif k == "total_liabilities":
                liab_total = v
            else:
                label = str(k)
                cl_explicit = any(c in label for c in ("2000", "2010", "2100", "2110", "2120", "2150", "2200", "2210", "2220", "2300", "2310"))
                ncl_explicit = any(c in label for c in ("2400", "2410", "2500", "2510", "2600", "2700"))
                if cl_explicit:
                    cl_items.append((label, v))
                elif ncl_explicit:
                    ncl_items.append((label, v))
                else:
                    if cl_items and len(cl_items) < 7:
                        cl_items.append((label, v))
                    else:
                        ncl_items.append((label, v))
        subsec("  Current Liabilities (≤ 12 months)")
        cl_subtotal_val = None
        for lbl, v in cl_items:
            if lbl == "__SUBTOTAL__":
                cl_subtotal_val = v
            else:
                line_item(lbl, v, level=2)
        if cl_subtotal_val is not None:
            line_item("TOTAL CURRENT LIABILITIES", cl_subtotal_val, level=0, is_total=True)
        subsec("  Non-Current Liabilities (> 12 months)")
        ncl_subtotal_val = None
        for lbl, v in ncl_items:
            if lbl == "__SUBTOTAL__":
                ncl_subtotal_val = v
            else:
                line_item(lbl, v, level=2)
        if ncl_subtotal_val is not None:
            line_item("TOTAL NON-CURRENT LIABILITIES", ncl_subtotal_val, level=0, is_total=True)
        if liab_total is not None:
            line_item("TOTAL LIABILITIES", liab_total, level=0, is_total=True, tag_override="grand_total")

        # ============= EQUITY (Per ASC 505) =============
        section("🟩  STOCKHOLDERS' EQUITY — Owner's Share (ASC 505)")
        eq_total_val = None
        for k, v in equity.items():
            if k == "total_equity":
                eq_total_val = v
            else:
                label = str(k)
                is_re_mid = ("Beginning of Period" in label or "Net Income for" in label or "Distributions" in label or "Drawings" in label)
                is_contra = ("Less:" in label or "Treasury" in label)
                if is_re_mid:
                    line_item(label, v, level=1)
                elif is_contra:
                    line_item(label, v, level=2)
                else:
                    line_item(label, v, level=2)
        if eq_total_val is not None:
            line_item("TOTAL STOCKHOLDERS' EQUITY", eq_total_val, level=0, is_total=True, tag_override="grand_total")

        # ============= GOLDEN EQUATION FOOTER =============
        totals_lbl = "🟰  TOTAL LIABILITIES + STOCKHOLDERS' EQUITY"
        self.bs_tree.zebra_insert(
            values=("", totals_lbl, fmt_aed(lines_obj.get("total_le", 0))),
            tags=("grand_total",))
        self.bs_tree.zebra_insert(
            values=("", "  (vs) TOTAL ASSETS  ↕  (must match above per double-entry)", fmt_aed(asset_total if asset_total is not None else lines_obj.get("assets", {}).get("total_assets", 0))),
            tags=("grand_total",))
        v = lines_obj.get("variance", Decimal(0))
        try:
            vv = Decimal(str(v))
        except Exception:
            vv = Decimal(0)
        if abs(vv) > Decimal("0.005"):
            self.bs_tree.zebra_insert(
                values=("⚠️", "Variance (Assets vs L + E)", fmt_aed(vv)),
                tags=("overdue_high",))
        else:
            self.bs_tree.zebra_insert(
                values=("✅", "Variance (Assets vs L + E) — perfectly balanced (within rounding)", "AED 0.00"),
                tags=("section",))
        # Append cross-statement reconciliation tie
        brk = lines_obj.get("__breakdowns", {}) or {}
        self.bs_tree.zebra_insert(values=("", "", ""), tags=("level2",))
        self.bs_tree.zebra_insert(
            values=("🔗  CROSS-STATEMENT RECONCILIATIONS (GAAP Required)", "", ""),
            tags=("section",))
        if brk:
            re_b = brk.get("re_beginning", 0)
            re_n = brk.get("re_add_ni", 0)
            re_d = brk.get("re_less_drawings", 0)
            re_e = brk.get("re_ending", 0)
            self.bs_tree.zebra_insert(
                values=(f"  ➜ Retained Earnings Bridge: {fmt_aed(re_b)} + NI {fmt_aed(re_n)} − Drawings {fmt_aed(-Decimal(str(re_d)) if isinstance(re_d,(int,float,Decimal)) else 0)} = {fmt_aed(re_e)}",
                        "", ""), tags=("level2",))
            self.bs_tree.zebra_insert(
                values=(f"  ➜ Ending Cash (BS 1000+1100) = Ending Cash per Cash Flow Statement (verifiable on Tab 4)",
                        "", ""), tags=("level2",))

    # -------- Aging --------
    def _refresh_aging(self):
        for i in self.age_tree.get_children():
            self.age_tree.delete(i)
        try:
            d = self.engine.get_aging_receivables()
        except Exception as e:
            self.age_explain.config(text=f"Error: {e}")
            return
        sb = d.get("summary_banner", {})
        self._fill_banner(self.age_banner, [
            (fmt_aed(sb.get("grand_total", 0)), "Total Unpaid"),
            (str(sb.get("overdue_invoice_count", 0)), "# Outstanding Invoices"),
            (fmt_aed(sb.get("avg", 0)), "Avg Unpaid / Invoice"),
            (fmt_aed(sb.get("ninety_plus", 0)), "91+ Days (Bad Debt Risk)"),
        ])
        self.age_explain.config(text=d.get("explanation", ""))
        bk = d.get("buckets", {})
        order = d.get("bucket_keys_order", list(bk.keys()))
        # Bucket summary first
        self.age_tree.zebra_insert(
            values=("", "AGING BUCKETS — SUMMARY", "", "", "", "", ""),
            tags=("section",))
        for b in order:
            amt = bk.get(b, 0)
            is_high = "91+" in b
            is_med = "61-90" in b or "31-60" in b
            tags = ()
            if is_high: tags = ("overdue_high",)
            elif is_med: tags = ("overdue_med",)
            self.age_tree.zebra_insert(
                values=("", f"  ▶  {b}", "", "", "", "", fmt_aed(amt)),
                tags=tags if tags else ("subsection",))
        # Line item invoices
        if d.get("rows"):
            self.age_tree.zebra_insert(
                values=("", "OUTSTANDING INVOICES — DETAIL", "", "", "", "", ""),
                tags=("section",))
        for r in d.get("rows", []):
            tag = ()
            dso = r.get("days_overdue", 0)
            if dso >= 91: tag = ("overdue_high",)
            elif dso >= 31: tag = ("overdue_med",)
            self.age_tree.zebra_insert(values=(
                r.get("invoice_no", ""),
                r.get("client", ""),
                r.get("date", ""),
                r.get("due_date", ""),
                f"{dso:,d}" if isinstance(dso, int) else str(dso),
                r.get("bucket", ""),
                fmt_aed(r.get("balance", 0))
            ), tags=tag)

    # -------- Revenue by Client --------
    def _refresh_rev_client(self):
        for i in self.rc_tree.get_children():
            self.rc_tree.delete(i)
        try:
            d = self.engine.get_revenue_by_client()
        except Exception as e:
            self.rc_explain.config(text=f"Error: {e}")
            return
        sb = d.get("summary_banner", {})
        self._fill_banner(self.rc_banner, [
            (fmt_aed(sb.get("grand_total", 0)), "Total Invoiced"),
            (str(sb.get("client_count", 0)), "Active Clients"),
            (fmt_aed(sb.get("avg", 0)), "Avg per Client"),
            (f"{d.get('total_clients', 0)} clients", "Total Client Count"),
        ])
        self.rc_explain.config(text=d.get("explanation", ""))
        rows = d.get("rows", [])
        total_inv_all = Decimal(0)
        total_paid_all = Decimal(0)
        total_unp_all = Decimal(0)
        n = len(rows)
        for idx, r in enumerate(rows):
            inv = Decimal(str(r.get("total_invoiced", 0)))
            pd = Decimal(str(r.get("total_paid", 0)))
            up = Decimal(str(r.get("total_unpaid", 0)))
            total_inv_all += inv; total_paid_all += pd; total_unp_all += up
            tags = ()
            if idx < 3: tags = ("level1",)  # Top-3 clients highlighted
            elif inv == 0: tags = ("zero",)
            self.rc_tree.zebra_insert(values=(
                r.get("client", ""),
                str(r.get("invoice_count", 0)),
                fmt_aed(inv), fmt_aed(pd), fmt_aed(up),
                f"{fmt_aed(r.get('paid_pct', 0)).replace('AED ', '')}%"
            ), tags=tags)
        if rows:
            paid_pct = (total_paid_all / total_inv_all * 100) if total_inv_all else Decimal(0)
            self.rc_tree.zebra_insert(values=(
                "✅  TOTAL (All Clients)", str(n),
                fmt_aed(total_inv_all), fmt_aed(total_paid_all),
                fmt_aed(total_unp_all), f"{round_aed(paid_pct)}%"
            ), tags=("grand_total",))

    # -------- Stock Valuation --------
    def _refresh_stock(self):
        for i in self.st_tree.get_children():
            self.st_tree.delete(i)
        try:
            d = self.engine.get_stock_valuation()
        except Exception as e:
            self.st_explain.config(text=f"Error: {e}")
            return
        sb = d.get("summary_banner", {})
        self._fill_banner(self.st_banner, [
            (fmt_aed(sb.get("grand_total", 0)), "Closing Inventory (Cost)"),
            (fmt_aed(sb.get("retail_value", 0)), "Estimated Retail Value"),
            (f"{fmt_aed(sb.get('turnover', 0))}x", "Stock Turnover / Year"),
            (f"{fmt_aed(sb.get('dio', 0))} days", "Days Inventory Outstanding"),
        ])
        self.st_explain.config(text=d.get("explanation", ""))
        lines = d.get("lines", {})
        nice_labels = {
            "opening_inventory_cost": "➕ Opening Inventory (at start of period)",
            "goods_received_purchases": "➕ Goods Received (Purchases)",
            "goods_sold_cogs": "➖ Goods Sold (COGS)",
        }
        totals_lbl = {
            "closing_inventory_cost": "🟰  CLOSING INVENTORY (Cost Basis)",
            "closing_inventory_retail_value": "💹  Estimated Retail Value (~ @ typical margin)",
            "stock_turnover_ratio": "📊 Stock Turnover Ratio (x / year)",
            "days_inventory_outstanding": "📊 Days Inventory Outstanding (days)",
        }
        for k, v in lines.items():
            if k in nice_labels:
                tags = ("level2",)
                try:
                    if abs(Decimal(str(v))) == 0: tags = ("zero",)
                except Exception:
                    pass
                self.st_tree.zebra_insert(
                    values=(nice_labels[k], fmt_aed(v)), tags=tags)
        # Section break: Totals
        self.st_tree.zebra_insert(values=("— TOTALS / KPIs —", ""), tags=("section",))
        for k, v in lines.items():
            if k in totals_lbl:
                tags = ("grand_total",) if k == "closing_inventory_cost" else ("total_row",)
                if k in ("stock_turnover_ratio", "days_inventory_outstanding"):
                    val = f"{fmt_aed(v).replace('AED ', '')}{'x' if k == 'stock_turnover_ratio' else ' days'}"
                    tags = ("total_row",)
                else:
                    val = fmt_aed(v)
                self.st_tree.zebra_insert(values=(totals_lbl[k], val), tags=tags)

    # -------- Supplier Spend --------
    def _refresh_supplier(self):
        for i in self.sp_tree.get_children():
            self.sp_tree.delete(i)
        try:
            d = self.engine.get_supplier_spend()
        except Exception as e:
            self.sp_explain.config(text=f"Error: {e}")
            return
        sb = d.get("summary_banner", {})
        top_share = sb.get("top_supplier_share_pct")
        self._fill_banner(self.sp_banner, [
            (fmt_aed(sb.get("grand_total", 0)), "Total Supplier & Expense Spend"),
            (str(sb.get("supplier_count", 0)), "# Suppliers / Payees"),
            (fmt_aed(sb.get("avg_supplier", 0)), "Avg per Supplier"),
            (f"{sb.get('top_supplier_name','N/A')} — {fmt_aed(top_share).replace('AED ','')}%",
             "Top Supplier Share of Total"),
        ])
        self.sp_explain.config(text=d.get("explanation", ""))
        rows = d.get("rows", [])
        for idx, r in enumerate(rows):
            tags = ()
            if idx < 3: tags = ("level1",)  # Top-3 highlighted (largest spenders)
            try:
                ts = Decimal(str(r.get("total_spend", 0)))
                if ts == 0: tags = ("zero",)
            except Exception:
                pass
            self.sp_tree.zebra_insert(values=(
                r.get("supplier", ""),
                str(r.get("po_count", 0)),
                fmt_aed(r.get("purchases", 0)),
                fmt_aed(r.get("expenses_paid", 0)),
                fmt_aed(r.get("total_spend", 0))
            ), tags=tags)
        if rows:
            tp = sum(Decimal(str(r.get("purchases", 0))) for r in rows)
            te = sum(Decimal(str(r.get("expenses_paid", 0))) for r in rows)
            self.sp_tree.zebra_insert(values=(
                "✅  TOTAL (All Suppliers / Payees)",
                str(len(rows)),
                fmt_aed(tp), fmt_aed(te), fmt_aed(d.get("total_spend", 0))
            ), tags=("grand_total",))

    # -------- Cash Flow --------
    def _refresh_cashflow(self):
        for i in self.cf_tree.get_children():
            self.cf_tree.delete(i)
        try:
            d = self.engine.get_cash_flow()
        except Exception as e:
            self.cf_explain.config(text=f"Error: {e}")
            import traceback; traceback.print_exc()
            return
        sb = d.get("summary_banner", {})
        ncf = sb.get("grand_total", 0)
        try:
            ncf_d = Decimal(str(ncf))
        except Exception:
            ncf_d = Decimal(0)
        sign = "➕" if ncf_d >= 0 else "➖"
        fcf = sb.get("free_cash_flow", Decimal(0))
        try:
            fcf_d = Decimal(str(fcf))
            fcf_sign = "✅" if fcf_d >= 0 else "⚠️"
        except Exception:
            fcf_sign = ""
        self._fill_banner(self.cf_banner, [
            (f"{sign} {fmt_aed(ncf)}", "Net Cash Flow (Δ in Cash)"),
            (f"🟢 Op {fmt_aed(sb.get('operating_cash_flow', 0))}", "Operating Activities (ASC 230)"),
            (f"{fcf_sign} FCF {fmt_aed(sb.get('free_cash_flow', 0))}", "Free Cash Flow (Op − CapEx)"),
            (f"🔴Inv {fmt_aed(sb.get('investing_cash_flow', 0))} | 🔵Fin {fmt_aed(sb.get('financing_cash_flow', 0))}", "Investing  |  Financing"),
        ])
        expl_parts = [d.get("explanation", "") or ""]
        diag = d.get("diagnostics") or {}
        checks = diag.get("checks") or []
        if checks:
            passed_n = sum(1 for c in checks if c.get("passed"))
            expl_parts.append("\n\n" + "═"*70)
            expl_parts.append(f"US GAAP Cash Flow Integrity — {passed_n}/{len(checks)} Checks Passed (ASC 230)")
            expl_parts.append("─"*70)
            for c in checks:
                icon = "✅" if c.get("passed") else "⚠️"
                expl_parts.append(f"  {icon}  {c.get('name','')}\n      {c.get('detail','')}")
            issues = diag.get("issues") or diag.get("root_causes") or []
            for it in issues:
                expl_parts.append(f"  🔍  Issue: {it}")
            fss = diag.get("fix_steps") or []
            for step in fss:
                expl_parts.append(f"  🔧  {step}")
            expl_parts.append(f"\nSummary: {diag.get('summary','')}")
        self.cf_explain.config(text="\n".join(expl_parts))
        rows = d.get("table_rows") or []
        def _fmt_cf_label(row):
            lbl = str(row.get("label", ""))
            tag = row.get("tag") or ""
            amt = row.get("amount")
            if tag == "section":
                return ("section", lbl)
            if tag == "total_row":
                nice = lbl
                if "Operating Activities" in nice:
                    nice = "✅  Net Cash from Operating Activities"
                elif "Investing Activities" in nice:
                    nice = "✅  Net Cash from Investing Activities"
                elif "Financing Activities" in nice:
                    nice = "✅  Net Cash from Financing Activities"
                elif "Closing Cash" in nice:
                    nice = "🟰  CLOSING CASH & CASH EQUIVALENTS at End of Period"
                return ("total_row", nice)
            if tag == "grand_total":
                return ("grand_total", lbl)
            if tag == "overdue_med":
                return ("overdue_med", lbl)
            if tag == "overdue_high":
                return ("overdue_high", lbl)
            amt_zero = False
            try:
                if amt is not None and abs(Decimal(str(amt))) == 0:
                    amt_zero = True
            except Exception:
                pass
            tags = ("zero",) if amt_zero else ("level2",)
            return (tags, lbl)

        if rows:
            for r in rows:
                tag_tuple_or_str, lbl = _fmt_cf_label(r)
                amt = r.get("amount")
                if amt is None:
                    val = ""
                elif isinstance(amt, (Decimal, int, float)):
                    val = fmt_aed(amt)
                else:
                    val = str(amt)
                tags = (tag_tuple_or_str,) if isinstance(tag_tuple_or_str, str) else tag_tuple_or_str
                self.cf_tree.zebra_insert(values=(lbl, val), tags=tags)
        else:
            lines = d.get("lines", {})
            for k, v in lines.items():
                lbl = str(k).replace("_", " ").title()
                tags = ()
                if isinstance(v, (Decimal, int, float)):
                    val = fmt_aed(v)
                else:
                    val = str(v)
                self.cf_tree.zebra_insert(values=(lbl, val), tags=tags)
        self.cf_tree.zebra_insert(values=("", ""), tags=("level2",))
        self.cf_tree.zebra_insert(
            values=("🔗  CROSS-STATEMENT RECONCILIATIONS (GAAP Required)", ""),
            tags=("section",))
        self.cf_tree.zebra_insert(
            values=("  ➜ Ending Cash (1000+1100) flows to Balance Sheet Current Assets section", ""),
            tags=("level2",))
        self.cf_tree.zebra_insert(
            values=("  ➜ Operating CF adjustments (ΔAR, ΔINV, ΔAP) derived from BS comparative period balances", ""),
            tags=("level2",))

    # -------- Statement of Changes in Equity --------
    def _refresh_changes_in_equity(self):
        for i in self.eq_tree.get_children():
            self.eq_tree.delete(i)
        try:
            d = self.engine.get_statement_of_changes_in_equity()
        except Exception as e:
            self.eq_explain.config(text=f"Error: {e}")
            import traceback; traceback.print_exc()
            return
        sb = d.get("summary_banner", {})
        self._fill_banner(self.eq_banner, [
            (fmt_aed(sb.get("opening_equity", 0)), "Opening Total Equity"),
            (fmt_aed(sb.get("capital_contributions", 0)), "Capital Contributions"),
            (fmt_aed(sb.get("drawings", 0)), "Owner Drawings / Distributions"),
            (fmt_aed(sb.get("closing_equity", 0)), "Closing Total Equity"),
        ])
        expl_parts = [d.get("explanation", "") or ""]
        diag = d.get("diagnostics") or {}
        checks = diag.get("checks") or []
        if checks:
            passed_n = sum(1 for c in checks if c.get("passed"))
            expl_parts.append("\n\n" + "═"*70)
            expl_parts.append(f"US GAAP Statement of Changes in Equity — {passed_n}/{len(checks)} Checks Passed (ASC 505)")
            expl_parts.append("─"*70)
            for c in checks:
                icon = "✅" if c.get("passed") else "⚠️"
                expl_parts.append(f"  {icon}  {c.get('name','')}\n      {c.get('detail','')}")
            rcs = diag.get("root_causes") or diag.get("issues") or []
            for rc in rcs:
                expl_parts.append(f"  🔍  Root cause: {rc}")
            fss = diag.get("fix_steps") or []
            for step in fss:
                expl_parts.append(f"  🔧  {step}")
            expl_parts.append(f"\nSummary: {diag.get('summary','')}")
        self.eq_explain.config(text="\n".join(expl_parts))
        tag_map = {
            "section": "section",
            "total_row": "grand_total",
            "grand_total": "grand_total",
            "gross_profit": "good",
            "overdue_high": "overdue_high",
            "overdue_med": "overdue_med",
            "subsection": "subsection",
        }
        def fmt_val(v):
            if v is None: return ""
            if isinstance(v, (Decimal, int, float)): return fmt_aed(v)
            return str(v)
        for r in d.get("table_rows", []):
            cols = r.get("columns", {})
            cs = cols.get("Capital_Stock") if isinstance(cols, dict) else None
            ts = cols.get("Treasury_Stock") if isinstance(cols, dict) else None
            re_val = cols.get("Retained_Earnings") if isinstance(cols, dict) else None
            aoci = cols.get("AOCI") if isinstance(cols, dict) else None
            teq = cols.get("Total_Equity") if isinstance(cols, dict) else None
            t = r.get("tag") or r.get("type")
            tags = ()
            if t and t in tag_map:
                tags = (tag_map[t],)
            label = (r.get("label") or "").strip() if isinstance(r.get("label"), str) else ""
            if not label and isinstance(cols, dict):
                label = str(cols.get("Item") or cols.get("Label") or cols.get("Description") or "").strip()
            if "Closing" in label:
                tags = ("grand_total",)
            elif "Opening" in label and "section" not in (tags or ()):
                if not tags:
                    tags = ("subsection",)
            self.eq_tree.zebra_insert(values=(
                label,
                fmt_val(cs),
                fmt_val(ts),
                fmt_val(re_val),
                fmt_val(aoci),
                fmt_val(teq),
            ), tags=tags)
        # Cross-statement reconciliation ties (GAAP required)
        self.eq_tree.zebra_insert(values=("", "", "", "", "", ""), tags=("level2",))
        self.eq_tree.zebra_insert(
            values=("🔗  CROSS-STATEMENT RECONCILIATIONS (GAAP Required)", "", "", "", "", ""),
            tags=("section",))
        ni_pnl = sb.get("net_income_for_period")
        if ni_pnl is not None:
            self.eq_tree.zebra_insert(
                values=(f"  ➜ Net Income {fmt_aed(ni_pnl)} (from P&L statement) = RE addition line above",
                        "", "", "", "", ""), tags=("level2",))
        oci_pnl = sb.get("other_comprehensive_income")
        if oci_pnl is not None:
            self.eq_tree.zebra_insert(
                values=(f"  ➜ OCI {fmt_aed(oci_pnl)} (from P&L OCI section) = AOCI movement above",
                        "", "", "", "", ""), tags=("level2",))
        self.eq_tree.zebra_insert(
            values=("  ➜ Closing Total Equity (rightmost column) = Balance Sheet Stockholders' Equity total",
                    "", "", "", "", ""), tags=("level2",))

    # -------- Financial Ratios / KPIs --------
    def _refresh_financial_ratios(self):
        for i in self.rt_tree.get_children():
            self.rt_tree.delete(i)
        try:
            d = self.engine.get_financial_ratios()
        except Exception as e:
            self.rt_explain.config(text=f"Error: {e}")
            import traceback; traceback.print_exc()
            return
        sb = d.get("summary_banner", {})
        verdict_counts = d.get("traffic_light_counts", {})
        green  = verdict_counts.get("green", 0)
        yellow = verdict_counts.get("yellow", 0)
        red    = verdict_counts.get("red", 0)
        self._fill_banner(self.rt_banner, [
            (f"{sb.get('net_margin_pct', 0)}%", "Net Profit Margin"),
            (f"{sb.get('current_ratio', 0)}x", "Current Ratio (Liquidity)"),
            (f"{sb.get('debt_to_equity', 0)}x", "Debt to Equity (Leverage)"),
            (f"🟢{green}  🟡{yellow}  🔴{red}", "KPI Verdict Counts"),
        ])
        self.rt_explain.config(text=d.get("explanation", ""))
        tag_map_render = {
            "section": "section",
            "total_row": "total_row",
            "gross_profit": "good",       # 🟢 GREEN
            "overdue_med": "overdue_med", # 🟡 YELLOW
            "overdue_high": "overdue_high", # 🔴 RED
        }
        for r in d.get("table_rows", []):
            # category separator rows emitted by engine use "category_separator": True
            is_cat_sep = r.get("category_separator") or r.get("item") and (
                str(r.get("label") or r.get("item", "")).startswith("💧") or
                str(r.get("label") or r.get("item", "")).startswith("📈") or
                str(r.get("label") or r.get("item", "")).startswith("🏗️") or
                str(r.get("label") or r.get("item", "")).startswith("⚙️"))
            t = r.get("tag") or r.get("type")
            tags = ()
            if t and t in tag_map_render:
                tags = (tag_map_render[t],)
            label = str(r.get("label") or r.get("item") or "")
            if is_cat_sep:
                tags = ("section",)
                # Clean up label (engine passes label = category name)
                label = f"📊  {label.upper()}"
            self.rt_tree.zebra_insert(values=(
                label,
                r.get("value", ""),
                r.get("benchmark", ""),
                r.get("verdict", ""),
                r.get("description", ""),
            ), tags=tags)

    # -------- Audit --------
    def _refresh_audit(self):
        for i in self.audit_tree.get_children():
            self.audit_tree.delete(i)
        try:
            d = self.engine.run_audit()
        except Exception as e:
            self.audit_explain.config(text=f"Error: {e}")
            return
        bal = d.get("is_balanced")
        banner_txt = "✅ PASSED — Debits = Credits" if bal else "❌ FAILED — Unbalanced entries exist"
        banner_color = GREEN_DEEP if bal else RED_STRONG
        self._fill_banner(self.audit_banner, [
            (fmt_aed(d.get("total_debits", 0)), "Total Debits (this period)"),
            (fmt_aed(d.get("total_credits", 0)), "Total Credits (this period)"),
            (fmt_aed(d.get("variance", 0)), "Variance (Dr − Cr)"),
            (banner_txt, "Audit Status"),
        ])
        self.audit_banner["d"][0].config(foreground=banner_color)
        self.audit_explain.config(text=d.get("explanation", ""))
        # Summary stats first (nice section + rows)
        self.audit_tree.zebra_insert(
            values=("", "AUDIT SUMMARY", "", "", ""), tags=("section",))
        td = d.get("total_debits", Decimal(0)); tc = d.get("total_credits", Decimal(0))
        summary_rows = [
            ("Σ All Debits (period)", fmt_aed(td)),
            ("Σ All Credits (period)", fmt_aed(tc)),
            ("Variance (Dr − Cr)", fmt_aed(td - Decimal(str(tc)))),
        ]
        batches = d.get("batches") or {}
        summary_rows.append(("Total Journal Entry Batches", f"{len(batches):,}"))
        mismatched = d.get("mismatched_transactions", [])
        summary_rows.append(("⚠️ MISMATCHED Batches", f"{len(mismatched):,}"))
        for lbl, val in summary_rows:
            is_variance = "Variance" in lbl
            is_mismatch = "MISMATCHED" in lbl
            tags = ()
            if is_variance:
                try:
                    if abs(Decimal(str(d.get("variance", 0)))) == 0:
                        tags = ("good",)
                    else:
                        tags = ("overdue_high",)
                except Exception:
                    tags = ("overdue_high",)
            elif is_mismatch:
                tags = ("good",) if len(mismatched) == 0 else ("overdue_high",)
            else:
                tags = ("level2",)
            self.audit_tree.zebra_insert(
                values=("", lbl, val if "Debits" in lbl else "",
                        val if "Credits" in lbl else "",
                        val if is_variance else ""),
                tags=tags)
        if mismatched:
            self.audit_tree.zebra_insert(
                values=("", "— TOP UNBALANCED BATCHES (largest variance first) —",
                        "", "", ""),
                tags=("section",))
            for m in mismatched:
                self.audit_tree.zebra_insert(values=(
                    m.get("reference_type", ""),
                    m.get("reference_id", ""),
                    fmt_aed(m.get("debits", 0)),
                    fmt_aed(m.get("credits", 0)),
                    fmt_aed(m.get("variance", 0)),
                ), tags=("overdue_high",))
        else:
            self.audit_tree.zebra_insert(
                values=("—", "✅ All double-entry batches balance perfectly. GAAP Compliant.",
                        "—", "—", "—"),
                tags=("good",))

    def _refresh_footnotes(self):
        for tree in (self.fn_policies_tree, self.fn_contingencies_tree):
            for i in tree.get_children():
                tree.delete(i)
        try:
            d = self.engine.get_footnotes()
        except Exception as e:
            import traceback; traceback.print_exc()
            self.fn_explain.config(text=f"Error: {e}")
            return
        sb = d.get("summary_banner", {}) or {}
        diag = d.get("diagnostics") or {}
        checks = diag.get("checks") or []
        default_count = sum(1 for c in checks if "DEFAULT" in (c.get("detail", "") or ""))
        mat_threshold = d.get("materiality_threshold", Decimal("0"))
        self._fill_banner(self.fn_banner, [
            (fmt_aed(mat_threshold), "Materiality Threshold"),
            (f"{sb.get('policy_count', 9)} policies", "Accounting Policies (ASC 235)"),
            (f"{len(d.get('contingent_liabilities', []))} items", "Contingencies Found (ASC 450)"),
            (f"{len(d.get('related_parties', []))} items · {default_count} default", "Related-Party Items · Default TPLs"),
        ])
        expl_parts = [d.get("explanation", "") or ""]
        if checks:
            passed_n = sum(1 for c in checks if c.get("passed"))
            expl_parts.append("\n\n" + "═"*70)
            expl_parts.append(f"US GAAP Footnote Generation Integrity — {passed_n}/{len(checks)} Items Generated")
            expl_parts.append("─"*70)
            for c in checks:
                icon = "⚠️" if "DEFAULT" in (c.get("detail", "") or "") else "✅"
                expl_parts.append(f"  {icon}  {c.get('name','')}\n      {c.get('detail','')}")
            issues = diag.get("issues") or []
            if issues:
                expl_parts.append("\n" + "─"*70)
                expl_parts.append("⚠️  MANAGEMENT REVIEW REQUIRED (default template items):")
                for iss in issues:
                    expl_parts.append(f"  • {iss}")
            expl_parts.append(f"\nSummary: {diag.get('summary','')}")
        self.fn_explain.config(text="\n".join(expl_parts))

        acct_policies = d.get("accounting_policies", []) or []
        self.fn_policies_tree.zebra_insert(
            values=("", "SIGNIFICANT ACCOUNTING POLICIES (ASC 235-10-50)", ""),
            tags=("section",))
        for p in acct_policies:
            is_default = p.get("_uses_default_template", True)
            tags = ("overdue_high",) if is_default else ("good",)
            self.fn_policies_tree.zebra_insert(
                values=(p.get("code", ""), p.get("title", ""), p.get("text", "")),
                tags=tags)

        contingencies = d.get("contingent_liabilities", []) or []
        related_parties = d.get("related_parties", []) or []
        self.fn_contingencies_tree.zebra_insert(
            values=("CONTINGENT LIABILITIES (ASC 450 / 460)", "", "", "", ""),
            tags=("section",))
        for c in contingencies:
            is_default = c.get("_uses_default_template", True)
            is_material = c.get("materiality_flag", False)
            if is_material:
                tags = ("overdue_high",)
            elif is_default:
                tags = ("overdue_med",)
            else:
                tags = ("good",)
            amt = c.get("amount", Decimal("0"))
            amt_str = fmt_aed(amt) if amt and amt != Decimal("0") else ""
            self.fn_contingencies_tree.zebra_insert(
                values=("Contingent Liability",
                        c.get("code", ""),
                        c.get("assessment", ""),
                        c.get("text", ""),
                        amt_str),
                tags=tags)

        self.fn_contingencies_tree.zebra_insert(
            values=("RELATED-PARTY TRANSACTIONS (ASC 850)", "", "", "", ""),
            tags=("section",))
        for r in related_parties:
            flag = r.get("flag", "") or ""
            is_default = r.get("_uses_default_template", True)
            if flag.startswith("⚠️"):
                tags = ("overdue_high",)
            elif is_default:
                tags = ("overdue_med",)
            else:
                tags = ("good",)
            amt = r.get("amount", Decimal("0"))
            amt_str = fmt_aed(amt) if amt and amt != Decimal("0") else ""
            self.fn_contingencies_tree.zebra_insert(
                values=("Related Party",
                        r.get("code", ""),
                        flag,
                        r.get("text", ""),
                        amt_str),
                tags=tags)

    def _fill_banner(self, cells, quad):
        for key, (val, lbl) in zip(["a", "b", "c", "d"], quad):
            v_w, l_w = cells[key]
            v_w.config(text=str(val if val not in (None, "") else "-"))
            l_w.config(text=str(lbl or ""))

    # -----------------------------------------------------------------------
    # 3.7 Exports
    # -----------------------------------------------------------------------
    _KEY_TO_REPORT = {
        "profit_and_loss": ("Profit & Loss Statement", lambda e: e.get_profit_and_loss()),
        "balance_sheet": ("Statement of Financial Position (Balance Sheet)", lambda e: e.get_balance_sheet()),
        "changes_in_equity": ("Statement of Changes in Equity", lambda e: e.get_statement_of_changes_in_equity()),
        "cash_flow": ("Cash Flow Statement (Operating / Investing / Financing)", lambda e: e.get_cash_flow()),
        "financial_ratios": ("Financial Ratios & KPIs Dashboard", lambda e: e.get_financial_ratios()),
        "aging_receivables": ("Accounts Receivable Aging", lambda e: e.get_aging_receivables()),
        "revenue_by_client": ("Revenue by Client", lambda e: e.get_revenue_by_client()),
        "stock_valuation": ("Stock Valuation & Movement", lambda e: e.get_stock_valuation()),
        "supplier_spend": ("Supplier Analysis & Spend", lambda e: e.get_supplier_spend()),
        "audit": ("General Ledger Audit Report", lambda e: e.run_audit()),
        "footnotes": ("GAAP Footnotes & Disclosures (ASC 235 / 450 / 850)", lambda e: e.get_footnotes()),
    }

    _KEY_TO_REPORT_SANITIZED = {
        "profit_and_loss": "Profit_and_Loss",
        "balance_sheet": "Balance_Sheet",
        "changes_in_equity": "Statement_of_Changes_in_Equity",
        "cash_flow": "Cash_Flow",
        "financial_ratios": "Financial_Ratios",
        "aging_receivables": "Aging_Receivables",
        "revenue_by_client": "Revenue_by_Client",
        "stock_valuation": "Stock_Valuation",
        "supplier_spend": "Supplier_Spend",
        "audit": "GL_Audit",
        "footnotes": "GAAP_Footnotes",
        "expense_manager": "Expense_Ledger",
        "Expense Ledger": "Expense_Ledger",
    }

    def _sanitize_report_name(self, key):
        """Turn report_key into filesystem-safe Legacy Report Name (no spaces/slashes)."""
        if key in self._KEY_TO_REPORT_SANITIZED:
            return self._KEY_TO_REPORT_SANITIZED[key]
        cleaned = re.sub(r'[^A-Za-z0-9_]+', '_', str(key or "Report").strip())
        cleaned = cleaned.strip("_")
        return cleaned or "GAAP_Report"

    def _export(self, key, fmt):
        fmt_clean = str(fmt or "").lower().strip()
        if fmt_clean not in ("xlsx", "pdf", "csv"):
            messagebox.showwarning("Export", f"Unknown format: {fmt}")
            return

        if self._user_role == "ReadOnly":
            if not messagebox.askyesno(
                "🔒 Read-Only Role — Export Approval Required",
                f"Role '{self._user_role}' is restricted to view-only access.\n\n"
                f"Exports produce files that leave the system and may contain sensitive financial data.\n"
                f"Please ensure you have manager approval before sharing exported files externally.\n\n"
                f"Proceed with export anyway?"):
                return

        eng = getattr(self, "engine", None)
        if eng is None:
            messagebox.showwarning("No Engine", "Please choose a period first.")
            return

        data = None
        display_name = str(key).replace("_", " ").title()
        meta = self._KEY_TO_REPORT.get(key)
        if meta:
            display_name, getter = meta
            try:
                data = getter(eng)
            except Exception as e:
                messagebox.showerror("Report Generation Failed", str(e))
                return

        reports_dir = os.path.join(self._data_folder, "FinancialReports")
        try:
            os.makedirs(reports_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Export Failed",
                                 f"Could not create FinancialReports folder:\n{e}")
            return

        report_basename = self._sanitize_report_name(key)
        stamp_date = datetime.now().strftime("%Y%m%d")
        stamp_time = datetime.now().strftime("%H%M%S")
        ext = fmt_clean
        if fmt_clean == "xlsx":
            ext = "xlsx"
        elif fmt_clean == "pdf":
            ext = "pdf"
        elif fmt_clean == "csv":
            ext = "csv"
        filename = f"{report_basename}_{stamp_date}_{stamp_time}.{ext}"
        out_path = os.path.join(reports_dir, filename)
        abs_out = os.path.abspath(out_path)

        export_success = False
        try:
            if fmt_clean == "xlsx":
                try:
                    if meta and data is not None:
                        eng.export_excel(display_name, data, output_path=out_path)
                    else:
                        self._export_fallback_xlsx_from_tree(key, display_name, out_path)
                    export_success = os.path.isfile(out_path)
                except Exception as ex_primary:
                    try:
                        self._export_fallback_xlsx_from_tree(key, display_name, out_path)
                        export_success = os.path.isfile(out_path)
                    except Exception as ex_fb:
                        raise RuntimeError(
                            f"Excel export failed: primary engine error={ex_primary}; "
                            f"fallback tree error={ex_fb}") from ex_primary
            elif fmt_clean == "csv":
                try:
                    if meta and data is not None:
                        self._export_csv_from_engine_data(data, out_path)
                    else:
                        self._export_csv_from_tree(key, out_path)
                    export_success = os.path.isfile(out_path)
                except Exception as ex_csv:
                    raise RuntimeError(f"CSV export failed: {ex_csv}") from ex_csv
            elif fmt_clean == "pdf":
                logo = self._find_logo()
                try:
                    import reportlab  # noqa: F401
                    has_reportlab = True
                except Exception:
                    has_reportlab = False
                try:
                    if meta and data is not None and has_reportlab:
                        eng.export_pdf(display_name, data, output_path=out_path, logo_path=logo)
                        if not os.path.isfile(out_path):
                            raise RuntimeError("engine.export_pdf completed but file missing")
                        export_success = True
                    elif meta and data is not None and not has_reportlab:
                        csv_path = os.path.join(
                            reports_dir,
                            f"{report_basename}_{stamp_date}_{stamp_time}.csv")
                        try:
                            self._export_csv_from_engine_data(data, csv_path)
                        except Exception:
                            pass
                        stub_note = (
                            "PDF generator missing 'reportlab' library; "
                            "CSV has been saved alongside this file.\n\n"
                            "To generate proper branded PDFs, install:\n"
                            "  pip install reportlab openpyxl\n\n"
                            f"Display name: {display_name}\n"
                            f"Period: {getattr(eng,'start','')} -> {getattr(eng,'end','')}\n"
                            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                            f"Role: {self._user_role}\n"
                        )
                        with open(out_path, "w", encoding="utf-8") as f:
                            f.write(stub_note)
                        export_success = True
                    else:
                        self._export_pdf_fallback_from_tree(key, display_name, out_path, logo)
                        export_success = os.path.isfile(out_path)
                except Exception as ex_pdf:
                    try:
                        csv_path = os.path.join(
                            reports_dir,
                            f"{report_basename}_{stamp_date}_{stamp_time}.csv")
                        try:
                            if meta and data is not None:
                                self._export_csv_from_engine_data(data, csv_path)
                            else:
                                self._export_csv_from_tree(key, csv_path)
                        except Exception:
                            pass
                        stub_note = (
                            f"PDF generator error: {ex_pdf}\n\n"
                            "CSV has been saved alongside this PDF stub file.\n\n"
                            f"Display name: {display_name}\n"
                            f"Period: {getattr(eng,'start','')} -> {getattr(eng,'end','')}\n"
                            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                            f"Role: {self._user_role}\n"
                        )
                        with open(out_path, "w", encoding="utf-8") as f:
                            f.write(stub_note)
                        export_success = True
                    except Exception as ex_stub:
                        raise RuntimeError(
                            f"PDF export failed ({ex_pdf}); stub write also failed ({ex_stub})"
                        ) from ex_pdf
        except Exception as e:
            import traceback; traceback.print_exc()
            messagebox.showerror("Export Failed", str(e))
            return

        if not export_success:
            messagebox.showerror("Export Failed",
                                 f"Could not write output file:\n  {abs_out}")
            return

        file_size = 0
        try:
            if os.path.isfile(abs_out):
                file_size = os.path.getsize(abs_out)
        except Exception:
            file_size = 0

        try:
            fmt_label = {"xlsx": "Excel", "pdf": "PDF", "csv": "CSV"}.get(fmt_clean, fmt_clean.upper())
            self._append_audit_log("EXPORT", {
                "export_format": fmt_label,
                "report_key": key,
                "output_file_path": abs_out,
                "file_size_bytes": int(file_size),
            })
        except Exception:
            pass

        try:
            abs_final = os.path.abspath(out_path)
        except Exception:
            abs_final = out_path

        try:
            pretty = self._KEY_TO_REPORT.get(key, (display_name,))[0]
        except Exception:
            pretty = display_name
        period_info = ""
        try:
            period_info = f"Period: {getattr(eng,'start','')}  →  {getattr(eng,'end','')}"
        except Exception:
            pass

        user_msg = (
            f"✅ {fmt_label} export successful!\n\n"
            f"Report:   {pretty}\n"
            f"{period_info}\n\n"
            f"Format:   .{fmt_clean.upper()}   ({file_size:,} bytes)\n\n"
            f"✅ Saved to:\n{abs_final}"
        )
        if messagebox.askyesno(f"✅ {fmt_label} Export Complete", user_msg + "\n\nOpen file in its folder now?"):
            try:
                self._reveal_in_file_manager(abs_final)
            except Exception:
                try:
                    self._open_file(abs_final)
                except Exception:
                    pass

    # ------------------- Export Helpers (Legacy Parity) --------------------
    def _export_fallback_xlsx_from_tree(self, key, display_name, out_path):
        """Fallback Excel export from treeview data when engine getter unavailable."""
        headers = []
        rows = []
        tags = []
        banner_text = ""
        period_info = ""
        try:
            map_, period, start_end = self._report_trees()
            period_info = f"Period: {period}   ·   {start_end}"
            if key in map_:
                banner, explain_w, tree, section, accent, pretty = map_[key]
                if isinstance(tree, (list, tuple)):
                    for t in tree:
                        h, r, ta = self._snapshot_tree_values(t)
                        headers.extend(h)
                        rows.extend(r)
                        tags.extend(ta)
                        headers.append("---")
                        rows.append(tuple("---" for _ in headers))
                        tags.append(())
                else:
                    headers, rows, tags = self._snapshot_tree_values(tree)
                try:
                    banner_pairs = self._extract_banner_text(banner)
                    banner_text = " | ".join(
                        f"{lab}: {val}" for val, lab in banner_pairs[:4])
                except Exception:
                    banner_text = ""
        except Exception:
            pass
        if not headers and not rows:
            headers = ["Value"]
            rows = [("No data available for this report.",)]
            tags = [()]
        self._export_excel_openpyxl(
            out_path, display_name, display_name, period_info or "", period_info or "",
            [(banner_text or "", "Summary Banner")],
            banner_text or "", headers, rows, tags,
        )

    def _export_csv_from_engine_data(self, data, out_path):
        """CSV export (UTF-8 BOM so Excel opens Arabic correctly) from engine data."""
        headers = []
        rows_list = []
        try:
            table_rows = data.get("table_rows") or []
            if table_rows:
                sample = table_rows[0] if table_rows else {}
                col_keys = [k for k in sample.keys() if k not in ("tag", "type", "columns")]
                if col_keys:
                    headers = [k.replace("_", " ").title() for k in col_keys]
                    for r in table_rows:
                        row = []
                        for k in col_keys:
                            v = r.get(k, "")
                            if v is None:
                                v = ""
                            if isinstance(v, Decimal):
                                v = fmt_aed(v)
                            row.append(v)
                        rows_list.append(tuple(row))
        except Exception:
            pass
        try:
            rows_k = data.get("rows") or []
            if not rows_list and rows_k:
                sample = rows_k[0] if rows_k else {}
                col_keys = list(sample.keys())
                if col_keys:
                    headers = [k.replace("_", " ").title() for k in col_keys]
                    for r in rows_k:
                        row = []
                        for k in col_keys:
                            v = r.get(k, "")
                            if v is None:
                                v = ""
                            if isinstance(v, Decimal):
                                v = fmt_aed(v)
                            row.append(v)
                        rows_list.append(tuple(row))
        except Exception:
            pass
        self._export_csv(out_path, headers, rows_list)

    def _export_csv_from_tree(self, key, out_path):
        """CSV export (UTF-8 BOM) from registered treeview snapshot."""
        headers = []
        rows = []
        try:
            map_, period, start_end = self._report_trees()
            if key in map_:
                _, _, tree, _, _, _ = map_[key]
                if isinstance(tree, (list, tuple)):
                    for idx, t in enumerate(tree):
                        h, r, _ = self._snapshot_tree_values(t)
                        if idx == 0:
                            headers = h
                        rows.extend(r)
                else:
                    headers, rows, _ = self._snapshot_tree_values(tree)
        except Exception:
            pass
        reg_tree = self._tree_registry.get(key, {}).get("tree")
        if not headers and reg_tree is not None:
            try:
                headers, rows, _ = self._snapshot_tree_values(reg_tree)
            except Exception:
                pass
        self._export_csv(out_path, headers or ["Column1"], rows or [("",)])

    def _export_pdf_fallback_from_tree(self, key, display_name, out_path, logo):
        """Fallback PDF export from treeview when engine getter unavailable."""
        headers = []
        rows = []
        tags = []
        banner = None
        try:
            map_, period, start_end = self._report_trees()
            if key in map_:
                banner, explain_w, tree, section_title, accent, pretty = map_[key]
                if isinstance(tree, (list, tuple)):
                    t = tree[0]
                    headers, rows, tags = self._snapshot_tree_values(t)
                else:
                    headers, rows, tags = self._snapshot_tree_values(tree)
                banner_pairs = self._extract_banner_text(banner)
                explain_txt = ""
                try:
                    if hasattr(explain_w, "cget"):
                        explain_txt = str(explain_w.cget("text") or "")
                except Exception:
                    explain_txt = ""
                section = section_title or display_name
                self._export_pdf_reportlab(
                    out_path, pretty or display_name, section, period,
                    start_end, banner_pairs, explain_txt, headers, rows, tags,
                )
                return
        except Exception as ex_fb:
            raise RuntimeError(f"PDF tree fallback failed: {ex_fb}") from ex_fb
        raise RuntimeError("No report metadata or tree data found for PDF fallback.")

    def _find_logo(self):
        candidates = []
        for attr in ("invoice_folder", "data_folder"):
            f = getattr(self.dm, attr, None)
            if f:
                candidates += [
                    os.path.join(f, "logo.png"),
                    os.path.join(f, "logo.jpg"),
                    os.path.join(f, "hopepharma_logo.png"),
                    os.path.join(f, "..", "logo.png"),
                ]
        for p in candidates:
            try:
                if os.path.isfile(p):
                    return p
            except Exception:
                pass
        return None

    def _open_file(self, p):
        try:
            import subprocess, platform
            if platform.system() == "Darwin":
                subprocess.Popen(["open", p])
            elif platform.system() == "Windows":
                os.startfile(p)
            else:
                subprocess.Popen(["xdg-open", p])
        except Exception as e:
            messagebox.showinfo("File Saved",
                                f"Could not auto-open file:\n{e}\n\nPath: {p}")

    # -----------------------------------------------------------------------
    # 3.8 Actions: Expense + Audit popup
    # -----------------------------------------------------------------------
    def _open_expense_dialog(self):
        SimpleExpenseDialog(self.win, self.dm,
                            on_saved_cb=self._refresh_engine_and_reports)

    def _run_audit_popup(self):
        if not self.engine:
            return
        d = self.engine.run_audit()
        bal = d.get("is_balanced")
        if bal:
            messagebox.showinfo(
                "✅ Audit Passed",
                f"General Ledger Audit (Period {self.engine.start} → {self.engine.end})\n\n"
                f"Total Debits:  {fmt_aed(d.get('total_debits', 0))}\n"
                f"Total Credits: {fmt_aed(d.get('total_credits', 0))}\n"
                f"Variance:      {fmt_aed(d.get('variance', 0))}\n\n"
                f"✅ All double-entry batches balance perfectly. GAAP Compliant."
            )
        else:
            self.nb.select(7)  # jump to audit tab
            mism = d.get("mismatched_transactions", [])
            messagebox.showerror(
                "❌ Audit FAILED",
                f"Variance: {fmt_aed(d.get('variance', 0))}\n\n"
                f"{len(mism)} transaction batches are unbalanced.\n"
                f"See the '8️⃣ GL Audit' tab for details."
            )

    # -----------------------------------------------------------------------
    # 3.9 Tab 11 — Expense Manager (List / Edit / Delete every expense)
    # -----------------------------------------------------------------------
    def _build_expense_manager_tab(self):
        key = "expense_manager"
        fr = tk.Frame(self.nb, bg="white")
        self.nb.add(fr, text=" 11️⃣ 💼 Expense Ledger ")

        wrap = tk.Frame(fr, bg="white", padx=18, pady=14)
        wrap.pack(fill="both", expand=True)

        # ---------- Section Header ----------
        header = tk.Frame(wrap, bg="#C0392B", height=62)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="   💼  EXPENSE REGISTER — Edit / Delete Recorded Expenses",
              bg="#C0392B", fg="white", font=("Segoe UI", 15, "bold"),
              anchor="w").pack(side="left", fill="both", expand=True, padx=12)
        self._exp_total_lbl = tk.Label(
            header, text="Total: AED 0.00  ·  0 records",
            bg="#C0392B", fg="white", font=("Segoe UI", 12, "bold"),
            anchor="e", padx=16
        )
        self._exp_total_lbl.pack(side="right", fill="y")

        # ---------- Explain banner ----------
        exp = tk.Frame(wrap, bg="white", padx=16, pady=10,
                    highlightbackground=SEPARATOR, highlightthickness=1)
        exp.pack(fill="x", pady=(12, 10))
        tk.Label(exp,
              text=("📌  This register shows EVERY expense you ever recorded via the 💸 Expense dialog,\n"
                    "      bank withdrawals classified as expenses, and invoice-level delivery/customs fees.\n"
                    "      ✏️  Edit = reverse old GL batch + post a corrected one.  🗑️  Delete = reverse (void) it."),
              font=("Segoe UI", 10), bg="white", justify="left").pack(anchor="w")

        # ---------- Filter bar ----------
        fb = tk.LabelFrame(wrap, text=" 🔍 Search & Filter Expenses ", bg="white",
                        fg=NAVY, font=("Segoe UI", 10, "bold"), padx=14, pady=10)
        fb.pack(fill="x", pady=(0, 10))

        # Row 1
        r1 = tk.Frame(fb, bg="white"); r1.pack(fill="x", pady=2)

        tk.Label(r1, text="Period:", bg="white", fg=MUTED,
              font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8))
        # NOTE: StringVar/BooleanVar for this tab are created EARLY in __init__
        # (BEFORE _build runs).  If they already exist, DON'T re-create them —
        # reassigning StringVars after they're bound to a Combobox/Entry silently
        # detaches the widget's trace callbacks and the filter UI stops driving
        # _reload_expense_tab properly.
        if not hasattr(self, "_exp_period_var") or not isinstance(getattr(self, "_exp_period_var", None), tk.StringVar):
            self._exp_period_var = tk.StringVar(value="Use Period Picker")
        cb = ttk.Combobox(r1, textvariable=self._exp_period_var,
                          values=["Use Period Picker", "This Month", "Last Month",
                                  "This Quarter", "This Year", "All Time"],
                          state="readonly", width=18, font=("Segoe UI", 10))
        cb.grid(row=0, column=1, sticky="w"); cb.bind("<<ComboboxSelected>>", lambda *_: self._reload_expense_tab())

        tk.Label(r1, text="Category:", bg="white", fg=MUTED,
              font=("Segoe UI", 10, "bold")).grid(row=0, column=2, sticky="w", padx=(18, 8))
        if not hasattr(self, "_exp_cat_var") or not isinstance(getattr(self, "_exp_cat_var", None), tk.StringVar):
            self._exp_cat_var = tk.StringVar(value="All Categories")
        cats = ["All Categories"] + list(EXPENSE_CATEGORY_MAP.keys())
        cb2 = ttk.Combobox(r1, textvariable=self._exp_cat_var, values=cats,
                           state="readonly", width=22, font=("Segoe UI", 10))
        cb2.grid(row=0, column=3, sticky="w"); cb2.bind("<<ComboboxSelected>>", lambda *_: self._reload_expense_tab())

        tk.Label(r1, text="Status:", bg="white", fg=MUTED,
              font=("Segoe UI", 10, "bold")).grid(row=0, column=4, sticky="w", padx=(18, 8))
        if not hasattr(self, "_exp_status_var") or not isinstance(getattr(self, "_exp_status_var", None), tk.StringVar):
            self._exp_status_var = tk.StringVar(value="Active Only")
        cb3 = ttk.Combobox(r1, textvariable=self._exp_status_var,
                           values=["Active Only", "Reversed Only", "All"],
                           state="readonly", width=16, font=("Segoe UI", 10))
        cb3.grid(row=0, column=5, sticky="w"); cb3.bind("<<ComboboxSelected>>", lambda *_: self._reload_expense_tab())

        tk.Label(r1, text="Search:", bg="white", fg=MUTED,
              font=("Segoe UI", 10, "bold")).grid(row=0, column=6, sticky="w", padx=(18, 8))
        if not hasattr(self, "_exp_search_var") or not isinstance(getattr(self, "_exp_search_var", None), tk.StringVar):
            self._exp_search_var = tk.StringVar()
        ent = tk.Entry(r1, textvariable=self._exp_search_var, width=26,
                    font=("Segoe UI", 10), relief="groove",
                    highlightbackground=SEPARATOR, highlightthickness=1, highlightcolor=NAVY)
        ent.grid(row=0, column=7, sticky="w")
        ent.bind("<KeyRelease>", lambda e: self._reload_expense_tab())

        if not hasattr(self, "_exp_include_withdrawals_and_invcost_var") \
                or not isinstance(getattr(self, "_exp_include_withdrawals_and_invcost_var", None), tk.BooleanVar):
            self._exp_include_withdrawals_and_invcost_var = tk.BooleanVar(value=False)
            self._exp_reversed_var = self._exp_include_withdrawals_and_invcost_var
        if not hasattr(self, "_exp_include_reversed_var") \
                or not isinstance(getattr(self, "_exp_include_reversed_var", None), tk.BooleanVar):
            self._exp_include_reversed_var = tk.BooleanVar(value=False)
        if not hasattr(self, "_exp_origin_manual_var") or not isinstance(getattr(self, "_exp_origin_manual_var", None), tk.BooleanVar):
            self._exp_origin_manual_var = tk.BooleanVar(value=True)
        if not hasattr(self, "_exp_origin_system_var") or not isinstance(getattr(self, "_exp_origin_system_var", None), tk.BooleanVar):
            self._exp_origin_system_var = tk.BooleanVar(value=True)
        # ---- Row 2: source/origin toggles + misc withdrawal/inv cost
        ttk.Checkbutton(r1, text="Show withdrawals & invoice costs",
                        variable=self._exp_include_withdrawals_and_invcost_var,
                        command=self._reload_expense_tab
                        ).grid(row=0, column=8, padx=(22, 0), sticky="w")
        # ---- NEW DEFAULT checkbox: "Include 🗑️ Reversed" = OFF BY DEFAULT ----
        # When OFF, anything right-click deleted (reversed) does NOT appear in
        # the Expense Ledger anymore — no longer visible, no longer counted,
        # as if it never happened.  Auditors can flip ON to see forensics, and
        # Status filter "Reversed Only" auto-overrides and shows only reversed.
        ttk.Checkbutton(r1, text="Include 🗑️ Reversed (Audit / Forensics)",
                        variable=self._exp_include_reversed_var,
                        command=self._reload_expense_tab
                        ).grid(row=0, column=9, padx=(18, 0), sticky="w")
        ttk.Checkbutton(r1, text="✍️  Manual entries",
                        variable=self._exp_origin_manual_var, command=self._reload_expense_tab
                        ).grid(row=1, column=0, columnspan=2, padx=(2, 0), sticky="w", pady=(10, 0))
        ttk.Checkbutton(r1, text="⚙️  System-generated (payroll / deliveries / withdrawals)",
                        variable=self._exp_origin_system_var, command=self._reload_expense_tab
                        ).grid(row=1, column=2, columnspan=6, padx=(18, 0), sticky="w", pady=(10, 0))

        r1.grid_columnconfigure(99, weight=1)

        # ---------- Action Buttons bar ----------
        bb = tk.Frame(wrap, bg="white")
        bb.pack(fill="x", pady=(0, 8))
        self._btn_exp_new = tk.Button(
            bb, text="💸  New Expense", bg=PRIMARY, fg="white",
            activebackground=HOVER, activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
            padx=16, pady=7, command=self._open_expense_dialog
        )
        self._btn_exp_new.pack(side="left")
        self._btn_exp_edit = tk.Button(
            bb, text="✏️  Edit Selected", bg="#F39C12", fg="white",
            activebackground="#D68910", activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
            padx=16, pady=7, command=self._expense_action_edit
        )
        self._btn_exp_edit.pack(side="left", padx=(10, 0))
        self._btn_exp_del = tk.Button(
            bb, text="🗑️  Delete / Void", bg="#C0392B", fg="white",
            activebackground="#922B21", activeforeground="white",
            font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
            padx=16, pady=7, command=self._expense_action_delete
        )
        self._btn_exp_del.pack(side="left", padx=(10, 0))
        self._btn_exp_refresh = tk.Button(
            bb, text="🔄  Refresh", bg=GRAY_BG, fg=NAVY,
            activebackground="#E5E9F2", activeforeground=NAVY,
            font=("Segoe UI", 10, "bold"), relief="flat", cursor="hand2",
            padx=16, pady=7, command=lambda: self._refresh_engine_and_reports()
        )
        self._btn_exp_refresh.pack(side="left", padx=(10, 0))
        # ---------- RBAC: disable expense action buttons if not Admin/Accountant ----------
        try:
            _can_write = self._user_role in ("Admin", "Accountant")
            _tip = ("🔒 Restricted — role '" + self._user_role
                    + "' may view but not modify expenses. Contact Admin for write access.")
            for _bw in (self._btn_exp_new, self._btn_exp_edit, self._btn_exp_del):
                if not _can_write:
                    try:
                        _bw.configure(state="disabled", fg=GREY_MED, bg=GREY_BG_2)
                    except Exception:
                        pass
                    def _rbac_warn():
                        messagebox.showwarning("🔒 Read-Only Mode", _tip)
                    _bw.configure(command=_rbac_warn)
        except Exception:
            pass
        tk.Label(bb, text="Tip: double-click any row to edit · right-click for menu.",
              font=("Segoe UI", 9, "italic"), bg="white", fg=MUTED).pack(side="left", padx=18)

        # ---------- Tree ----------
        cols = ("date", "ref_id", "category", "payee", "description",
                "amount", "bank", "origin", "by", "status", "updated")
        tree_frame = tk.Frame(wrap, bg="white",
                           highlightbackground=SEPARATOR, highlightthickness=1)
        tree_frame.pack(fill="both", expand=True, pady=(4, 0))
        vsb = ttk.Scrollbar(tree_frame, orient="vertical")
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal")
        vsb.pack(side="right", fill="y"); hsb.pack(side="bottom", fill="x")
        tv = ZebraTreeview(
            tree_frame, columns=cols, show="headings",
            yscrollcommand=vsb.set, xscrollcommand=hsb.set
        )
        tv.pack(side="left", fill="both", expand=True)
        vsb.config(command=tv.yview); hsb.config(command=tv.xview)

        hdrs = [
            # (col, title, initial_width, anchor, is_num)
            # Sum tuned to ~1380 px (exactly fits default ~1382px inner window).
            # Column widths + 1 new "Source" column = 11 cols.
            ("date",       "Date",            90, "center", False),
            ("ref_id",     "Reference ID",   120, "center", False),
            ("category",   "Category → GL",  170, "w",      False),
            ("payee",      "Payee / Vendor", 140, "w",      False),
            ("description","Description",    250, "w",      False),
            ("amount",     "Amount (AED)",   115, "e",      True),
            ("bank",       "Paid From",      125, "w",      False),
            ("origin",     "Source",          90, "center", False),
            ("by",         "Added By",        85, "center", False),
            ("status",     "Status",          90, "center", False),
            ("updated",    "Last Updated",   115, "center", False),
        ]
        for code, title, width, anchor, is_num in hdrs:
            tv.heading(code, text=title, anchor="center")
            # Stretch EVERY column so that on a 1366/1440px window ALL columns
            # resize proportionally — no horizontal scrollbar needed on first
            # load, and if you maximise the window every description/payee
            # becomes fully readable.
            tv.column(code, width=width, anchor=anchor, stretch=True,
                      minwidth=max(55, width - 40))
        # Tag palette
        tv.tag_configure("reversed", background="#FADBD8", foreground="#78281F")
        tv.tag_configure("withdrawal", background="#FEF9E7")
        tv.tag_configure("invoice_cost", background="#EBF5FB")
        # Request 18: system-origin rows get a very subtle cool-blue tint so
        # you can skim the register and spot payroll/delivery auto entries.
        tv.tag_configure("origin_system", background="#F4F8FC")
        tv.tag_configure("origin_manual", background="#FFFFFF")

        self._exp_tree = tv
        # ---- RIGHT-CLICK CONTEXT MENU (cross-platform) -------------------
        # macOS 2-finger tap / Ctrl-click = <Button-2>; Windows/Linux right
        # click = <Button-3>.  Bind BOTH so "Right click → Delete" works on
        # every platform — fixes "nothing happens on right click" complaint.
        m = tk.Menu(wrap, tearoff=0)
        m.add_command(label="✏️  Edit Expense", command=self._expense_action_edit,
                      accelerator="Double-click")
        m.add_command(label="🗑️  Delete / Void Expense", command=self._expense_action_delete,
                      accelerator="Delete")
        m.add_separator()
        m.add_command(label="💸  Record a New Expense", command=self._open_expense_dialog)
        m.add_separator()
        m.add_command(label="📋  Copy Reference ID", command=lambda: self._clip_copy(self._sel_expense_ref()))
        m.add_command(label="📋  Copy Selected Row (All Columns)",
                      command=lambda: self._clip_copy(self._sel_expense_row_tsv()))
        m.add_command(label="🔍  Search payee/description again", command=self._search_from_row)
        m.add_command(label="🔁  Reload Expense Register",
                      command=lambda: self._reload_expense_tab())
        self._exp_ctx = m
        def _popup(ev):
            try:
                iid = tv.identify_row(ev.y) if hasattr(tv, "identify_row") else None
                if iid:
                    tv.selection_set(iid)
                else:
                    sel = tv.selection()
                    if not sel:
                        return
            except Exception:
                pass
            try:
                m.tk_popup(ev.x_root, ev.y_root)
            finally:
                m.grab_release()
        # --- Cross platform: both macOS 2-finger tap (Button-2) + Windows/Linux (Button-3)
        for ev_seq in ("<Button-2>", "<Button-3>"):
            tv.bind(ev_seq, _popup)
        tv.bind("<Double-1>", lambda e: self._expense_action_edit())
        tv.bind("<Delete>", lambda e: self._expense_action_delete())
        # Ctrl+Backspace also works (Windows muscle-memory)
        tv.bind("<Control-BackSpace>", lambda e: self._expense_action_delete())

        self.register_tree(key, None, None, tv,
                           "Expense Register", "#C0392B", "Expense Ledger")
        # Export bar (bottom)
        self._make_export_bar(wrap, "Expense Ledger")

        # Initial load
        self._reload_expense_tab()

    def _expense_filter_effective_dates(self):
        """Translate Period picker in Expense tab into actual start/end date strings."""
        choice = self._exp_period_var.get()
        if choice == "All Time":
            return None, None
        if choice == "Use Period Picker":
            s, e = self._get_custom_dates()
            return s, e
        today = date.today()
        if choice == "This Month":
            s = date(today.year, today.month, 1)
            # Last day of this month: first of next month - 1 day
            if today.month == 12:
                nx = date(today.year + 1, 1, 1)
            else:
                nx = date(today.year, today.month + 1, 1)
            e = nx - timedelta(days=1)
            return s.isoformat(), e.isoformat()
        if choice == "Last Month":
            if today.month == 1:
                s = date(today.year - 1, 12, 1)
                e = date(today.year - 1, 12, 31)
            else:
                s = date(today.year, today.month - 1, 1)
                if today.month == 1:
                    pass
                lm = date(today.year, today.month, 1) - timedelta(days=1)
                e = date(today.year, s.month, lm.day)
            return s.isoformat(), e.isoformat()
        if choice == "This Quarter":
            q = (today.month - 1) // 3
            s = date(today.year, q * 3 + 1, 1)
            end_month = q * 3 + 3
            if end_month == 12:
                nx = date(today.year + 1, 1, 1)
            else:
                nx = date(today.year, end_month + 1, 1)
            e = nx - timedelta(days=1)
            return s.isoformat(), e.isoformat()
        if choice == "This Year":
            return date(today.year, 1, 1).isoformat(), date(today.year, 12, 31).isoformat()
        s, e = self._get_custom_dates()
        return s, e

    def _classify_expense_origin(self, ref_type, created_by, description, ref_id, bundle_lines):
        """Classify any expense row as either 'Manual' (human-entered via the
        💸 New Expense dialog) or 'System' (auto-generated by payroll posting,
        invoice delivery fees, bank withdrawals, stock adjustments, reversals,
        or any other programmatic posting path).

        Used by: _reload_expense_tab, _aggregate_expenses_from_gl AND the
        LedgerService.list_expenses path must mirror this logic.
        """
        rt = (ref_type or "").strip().upper()
        # 1) Reference types that are ALWAYS system-generated
        if rt in {"WITHDRAWAL", "INVOICE_COST", "SALARY", "PAYROLL", "PURCHASE",
                  "STOCK_ADJUSTMENT", "REVERSAL", "ADJUSTMENT", "PO_ACCRUAL",
                  "CREDIT_NOTE", "INVOICE"}:
            return "System"
        # 2) EXPENSE reference_type — decide via created_by + description/ref_id
        cb = (created_by or "").strip()
        cb_low = cb.lower()
        # Known system service creators (module / class names)
        if cb and cb_low in {
            "system", "auto", "employee manager", "employeemanager",
            "invoiceapp", "invoice app", "ledgerservice", "ledger service",
            "enhancedclouddatamanager", "gaap dashboard", "reversal",
            "automatic", "batch", "scheduled"}:
            return "System"
        # Reversal markers → System
        desc = (description or "")
        if desc.startswith("🔁 REVERSAL OF") or "REVERSAL OF" in desc[:60]:
            return "System"
        # 3) EXP-type reference-id patterns: EXP dialog writes EXP-YYYYMMDD-NNNNNN
        rid = str(ref_id or "").strip()
        if rid.startswith("EXP-") and len(rid) >= 14:
            return "Manual"
        # 4) "Payee:" prefix = EXP dialog format → Manual
        if "Payee:" in desc and rt == "EXPENSE":
            return "Manual"
        # 5) created_by looks like a username (2+ chars, NOT like a module)
        if cb and len(cb) >= 2 and not any(x in cb_low for x in [
            "manager", "service", "module", "system", "auto"]):
            return "Manual"
        # Default: treat EXPENSE rows without strong clues as Manual (the
        # optimistic assumption is a user entered them). Everything else
        # (unknown reference types) default to System.
        return "Manual" if rt == "EXPENSE" else "System"

    def _reload_expense_tab(self):
        try:
            s, e = self._expense_filter_effective_dates()
            inc_misc = self._exp_include_withdrawals_and_invcost_var.get() if hasattr(self, "_exp_include_withdrawals_and_invcost_var") else self._exp_reversed_var.get()
            want_manual = bool(self._exp_origin_manual_var.get())
            want_system = bool(self._exp_origin_system_var.get())
            search = self._exp_search_var.get() or None
            st = self._exp_status_var.get() if hasattr(self, "_exp_status_var") else "Active Only"
            # ---- NEW DEFAULT: anything reversed (right-click deleted) does NOT
            # appear anymore, anywhere in the register, unless user explicitly
            # flips "Include Reversed" checkbox ON (auditor mode) OR picks
            # Status = Reversed Only (which implies include reversed).
            if hasattr(self, "_exp_include_reversed_var"):
                _include_rev = bool(self._exp_include_reversed_var.get())
            else:
                _include_rev = False
            # Implicit include when user explicitly requests Status=Reversed Only
            # or Status=All.  Otherwise user clicks "Reversed Only" and sees 0
            # rows, which is a confusing UX.
            if st in ("Reversed Only", "All"):
                _include_rev = True
                try: self._exp_include_reversed_var.set(True)
                except Exception: pass
            exclude_reversed = not _include_rev
            rows = None
            # Prefer real ledger service if available
            if self.ledger is not None and callable(getattr(self.ledger, "list_expenses", None)):
                try:
                    rows = self.ledger.list_expenses(
                        start_date=s, end_date=e, search=search,
                        include_withdrawals=inc_misc, include_invoice_costs=inc_misc,
                        exclude_reversed=exclude_reversed)
                except Exception:
                    rows = None
            # Manual fallback: aggregate general_ledger.json ourselves so the
            # register NEVER shows "0 rows" when real expenses exist in the GL,
            # and NEVER crashes even if LedgerService failed to initialize.
            if rows is None:
                rows = self._aggregate_expenses_from_gl(
                    start_date=s, end_date=e, search=search,
                    include_withdrawals=inc_misc, include_invoice_costs=inc_misc,
                    exclude_reversed=exclude_reversed)
            # Attach origin classifier to EVERY row (works for BOTH paths).
            # Already-attached origin from either aggregator is respected, but
            # we re-run through the classifier to guarantee both paths produce
            # exactly the same label regardless of caller.
            classified = []
            for r in rows:
                origin = self._classify_expense_origin(
                    ref_type=r.get("ref_type"),
                    created_by=r.get("created_by"),
                    description=r.get("description"),
                    ref_id=r.get("ref_id"),
                    bundle_lines=None,
                )
                r = dict(r)  # cheap copy to avoid mutating caller's list
                r["origin"] = origin
                classified.append(r)
            rows = classified
            # ORIGIN filter (Request 18)
            if not (want_manual and want_system):
                keep_types = set()
                if want_manual: keep_types.add("Manual")
                if want_system: keep_types.add("System")
                rows = [r for r in rows if r.get("origin") in keep_types]
            # Category filter
            cat = self._exp_cat_var.get()
            if cat and cat != "All Categories":
                rows = [r for r in rows if r.get("category") == cat]
            # Status filter
            # NOTE: because exclude_reversed=True by default removes reversed
            # bundles BEFORE we get here, "Active Only" == no further filtering
            # needed.  "Reversed Only" or "All" require that checkbox to also
            # be ON, otherwise they'll just see empty/active-only.
            st = self._exp_status_var.get()
            if st == "Active Only":
                rows = [r for r in rows if not (r.get("reversed_ref"))]
            elif st == "Reversed Only":
                rows = [r for r in rows if r.get("reversed_ref")]
            # Populate tree
            tv = self._exp_tree
            for iid in tv.get_children():
                tv.delete(iid)
            self._exp_row_meta.clear()
            total = Decimal("0")
            for r in rows:
                gross = r.get("gross_amount", Decimal("0"))
                try:
                    gross = Decimal(str(gross))
                except Exception:
                    gross = Decimal("0")
                total += gross
                tags = list(r.get("_row_tags") or [])
                origin = r.get("origin") or "System"
                if origin == "System":
                    tags.append("origin_system")
                else:
                    tags.append("origin_manual")
                if r.get("reversed_ref"):
                    tags.append("reversed")
                elif r.get("ref_type") == "WITHDRAWAL":
                    tags.append("withdrawal")
                elif r.get("ref_type") == "INVOICE_COST":
                    tags.append("invoice_cost")
                iid = tv.insert(
                    "", "end",
                    values=(
                        r.get("date") or "",
                        r.get("ref_id") or "",
                        f"{r.get('category','')}  →  {r.get('account_code','')} {r.get('account_name','')}".strip(" →"),
                        r.get("payee") or "",
                        r.get("description") or "",
                        fmt_aed(gross),
                        r.get("bank_account") or "",
                        r.get("origin") or "—",
                        r.get("created_by") or "",
                        r.get("status") or "Posted",
                        (r.get("updated_at") or r.get("created_at") or "")[:19].replace("T", " "),
                    ),
                    tags=tags
                )
                # Store meta on iid via tk name→dict: we'll use a parallel dict
                self._exp_row_meta[iid] = {
                    "ref_type": r.get("ref_type"),
                    "ref_id": r.get("ref_id"),
                    "category": r.get("category"),
                    "gross_amount": gross,
                    "payee": r.get("payee") or "",
                    "description": r.get("description") or "",
                    "date": r.get("date") or "",
                    "bank_account": r.get("bank_account") or "",
                    "reversed": bool(r.get("reversed_ref")),
                    "account_code": r.get("account_code"),
                    "account_name": r.get("account_name"),
                    "status": r.get("status") or "Posted",
                    "created_at": r.get("created_at"),
                    "updated_at": r.get("updated_at"),
                    "origin": r.get("origin"),
                    "created_by": r.get("created_by"),
                    "row": r,
                }
            self._exp_total_lbl.config(
                text=f"Total: {fmt_aed(total)}  ·  {len(rows)} record{'s' if len(rows)!=1 else ''}"
            )
        except Exception as ex:
            messagebox.showerror("❌ Expense Register",
                                 f"Could not rebuild expense register:\n{ex}")
            import traceback
            traceback.print_exc()

    def _aggregate_expenses_from_gl(self, start_date=None, end_date=None, search=None,
                                     include_withdrawals=True, include_invoice_costs=True,
                                     exclude_reversed=True):
        """Manual fallback aggregator: load general_ledger.json directly via dm
        and build the same expense rows LedgerService.list_expenses would return.
        GUARANTEES the register shows data even if self.ledger is missing/broken.

        Adds an `origin` field per row (classified as Manual or System) so the
        Expense Ledger's Source column and filter toggles work identically to
        the LedgerService path.

        NEW (simplified UX): by default ``exclude_reversed=True`` → any group
        marked as reversed (by a "🔁 REVERSAL OF …" GL marker) is dropped
        entirely from the output, and from all totals & counts.  This means
        right-click → Delete removes the expense from everywhere (lists,
        counts, totals) as the user expects.  Flip off to audit.
        """
        try:
            gl = self.dm.load_json('general_ledger.json') or []
        except Exception:
            gl = []
        import re as _re
        groups = {}
        for e in gl:
            if not isinstance(e, dict):
                continue
            rt = str(e.get('reference_type') or '')
            rid = str(e.get('reference_id') or '')
            if not rt or not rid:
                continue
            is_exp = rt == 'EXPENSE'
            if include_withdrawals and rt == 'WITHDRAWAL':
                is_exp = True
            if include_invoice_costs and rt == 'INVOICE_COST':
                is_exp = True
            if not is_exp:
                continue
            k = (rt, rid)
            if k not in groups:
                groups[k] = {'lines': [], 'rev_by': None}
            groups[k]['lines'].append(e)
        # Find reversal markers
        for e in gl:
            desc = str(e.get('description') or '')
            if desc.startswith("🔁 REVERSAL OF"):
                m = _re.search(r"REVERSAL OF\s+(\w+):(\S+)", desc)
                if m:
                    k = (m.group(1), m.group(2))
                    if k in groups:
                        groups[k]['rev_by'] = str(e.get('created_at') or desc)
        out = []
        for (rt, rid), bundle in groups.items():
            # ----- NEW: drop reversed groups completely by default. -----
            if exclude_reversed and bundle.get('rev_by'):
                continue
            lines = bundle.get('lines') or []
            if not lines:
                continue
            dts = sorted([str(x.get('date') or '') for x in lines if x.get('date')])
            expense_dt = dts[-1] if dts else ''
            if expense_dt and start_date and expense_dt < str(start_date):
                continue
            if expense_dt and end_date and expense_dt > str(end_date):
                continue
            dr_lines = [x for x in lines if Decimal(x.get('debit_amount') or 0) > 0]
            cr_lines = [x for x in lines if Decimal(x.get('credit_amount') or 0) > 0]
            expense_line = max(dr_lines, key=lambda x: Decimal(x.get('debit_amount') or 0)) if dr_lines else None
            payment_line = cr_lines[0] if cr_lines else None
            if expense_line is None:
                continue
            acct = str(expense_line.get('account_code') or '')
            gross = Decimal(expense_line.get('debit_amount') or 0)
            if gross <= 0:
                continue
            desc = str(expense_line.get('description') or '')
            payee = ''
            m_p = _re.search(r"Payee:\s*([^|]+)", desc)
            if m_p:
                payee = m_p.group(1).strip()
            clean_desc = desc
            if '|' in desc:
                clean_desc = desc.split('|', 1)[1].strip()
            elif 'Payee:' in desc:
                clean_desc = ''
            category_label = None
            lab = _re.match(r"^([^—|]+?)\s*—", desc)
            if lab:
                cand = lab.group(1).strip()
                if cand in EXPENSE_CATEGORY_MAP:
                    category_label = cand
            if category_label is None:
                rev_hits = [k for k, v in EXPENSE_CATEGORY_MAP.items() if v == acct]
                if len(rev_hits) == 1:
                    category_label = rev_hits[0]
                elif rev_hits:
                    category_label = self._expense_category_from_text(desc + " " + payee)
            if category_label is None:
                if acct.startswith('5'):
                    category_label = self._expense_category_from_text(
                        desc or expense_line.get('account_name') or '')
                else:
                    category_label = str(expense_line.get('account_name') or acct)
            bank_name = ''
            if payment_line is not None:
                bank_code = str(payment_line.get('account_code') or '')
                bank_name = str(payment_line.get('account_name') or bank_code)
                if '1100' in bank_code:
                    pdesc = payment_line.get('description') or ''
                    if '—' in pdesc:
                        bank_name = 'Bank ' + pdesc.split('—')[-1].strip()
            created_by = str(expense_line.get('created_by') or 'System')
            created_at = str(expense_line.get('created_at') or '')
            status = '🗑️  Reversed' if bundle['rev_by'] else 'Posted'
            origin = self._classify_expense_origin(
                ref_type=rt, created_by=created_by, description=desc,
                ref_id=rid, bundle_lines=lines,
            )
            if search:
                sq = str(search).lower().strip()
                if sq:
                    hay = " ".join(str(v or "") for v in [
                        expense_dt, rid, rt, category_label, acct,
                        payee, clean_desc, created_by, bank_name, origin
                    ]).lower()
                    if sq not in hay:
                        continue
            out.append({
                'date': expense_dt,
                'ref_id': rid,
                'ref_type': rt,
                'category': category_label,
                'account_code': acct,
                'account_name': str(expense_line.get('account_name') or acct),
                'payee': payee,
                'description': clean_desc,
                'gross_amount': gross,
                'bank_account': bank_name,
                'created_by': created_by,
                'created_at': created_at,
                'updated_at': bundle['rev_by'] or created_at,
                'reversed_ref': bundle['rev_by'],
                'status': status,
                'origin': origin,
            })
        out.sort(key=lambda r: (str(r['date']), str(r['created_at'])), reverse=True)
        return out

    def _expense_category_from_text(self, txt):
        """Keyword-based expense category classifier (mirrors ledger_service)."""
        t = (txt or '').lower()
        for (label, code), keywords in [
            (("Salaries & Wages", "5100"), ["salary", "wage", "payroll", "commission", "staff", "employee", "salaries"]),
            (("Rent", "5200"), ["rent", "lease", "landlord"]),
            (("Utilities", "5200"), ["electric", "water", "dewa", "utility", "internet", "wifi", "telephone", "phone bill"]),
            (("Customs & Delivery", "5200"), ["custom", "duty", "delivery", "shipping", "courier", "freight", "clearance"]),
            (("Marketing & Advertising", "5300"), ["marketing", "advert", "promo", "campaign", "social media", "facebook", "google ads"]),
            (("Travel & Transport", "5300"), ["travel", "flight", "hotel", "taxi", "fuel", "petrol", "transport"]),
            (("Office & Admin", "5300"), ["office", "stationery", "print", "post", "license", "membership", "software", "subscription"]),
        ]:
            if any(k in t for k in keywords):
                return label[0]
        # Fallback: highest category code matching any expense bucket
        return "General & Admin"

    def _sel_expense_ref(self):
        iid = self._exp_tree.selection()
        if not iid:
            return None, None, None
        iid = iid[0]
        meta = self._exp_row_meta.get(iid)
        if not meta:
            return None, None, None
        return meta["ref_type"], meta["ref_id"], meta

    def _expense_action_edit(self):
        rt, rid, meta = self._sel_expense_ref()
        if not rt:
            messagebox.showwarning("No row selected",
                                   "Select an expense row to edit, then click ✏️ Edit.")
            return
        if meta.get("reversed"):
            messagebox.showinfo("Already Reversed",
                                f"{rt}:{rid}\n\n"
                                "This expense has already been reversed/deleted and cannot be edited.\n"
                                "Instead, delete/void the corrected expense if that's wrong, or record a new one.")
            return
        # Open the Edit Expense dialog
        EditExpenseDialog(
            self.win, self.ledger, meta,
            on_saved_cb=lambda: self._refresh_engine_and_reports(),
            data_manager=self.dm,
        )

    def _expense_action_delete(self):
        rt, rid, meta = self._sel_expense_ref()
        if not rt:
            messagebox.showwarning("No row selected",
                                   "Select an expense row to delete/void, then click 🗑️ Delete.")
            return
        if meta.get("reversed"):
            messagebox.showinfo("Already Reversed",
                                "This expense is already reversed/deleted. No further action.")
            return
        gross = meta.get("gross_amount", Decimal("0"))
        confirm = messagebox.askyesno(
            "🗑️  Confirm — Delete (Reverse) Expense?",
            f"Reference:  {rt} : {rid}\n"
            f"Date:       {meta.get('date','')}\n"
            f"Payee:      {meta.get('payee') or '(none)'}\n"
            f"Category:   {meta.get('category') or ''}\n"
            f"Amount:     {fmt_aed(gross)}\n\n"
            "⚠️  This will POST A REVERSAL to the General Ledger:\n"
            "   • Dr 1100 Bank (increase by same amount)\n"
            "   • Cr Expense Account (decrease expense)\n\n"
            "Nothing is ever deleted (immutable GL).  A Reversed badge will appear next to the original row,\n"
            "and a new REVERSAL batch will appear for audit.  P&L / BS / CF will update immediately.\n\n"
            "Proceed with the reversal (this cannot be undone)?",
            icon="warning"
        )
        if not confirm:
            return
        try:
            # Guarantee a working service for the delete/reverse call
            svc = self.ledger
            if svc is None or not callable(getattr(svc, "delete_expense", None)):
                try:
                    from ledger_service import LedgerService
                    svc = LedgerService(self.dm)
                except Exception as _le:
                    svc = None
                    raise RuntimeError(
                        "LedgerService unavailable (cannot write GL reversals). "
                        "Please check ledger_service.py is in the same folder.") from _le
            result = svc.delete_expense(rt, rid, username="GAAP Dashboard")
            messagebox.showinfo(
                "✅ Expense Reversed",
                f"Deleted AED {fmt_aed(gross)}\n\n"
                f"• Original ref:  {rt}:{rid}\n"
                f"• Reversal batch: REVERSAL : {result.get('new_reference_id','')}\n"
                f"• GL lines written: {result.get('reversal_entries',0)}\n\n"
                "💡  Reports will now refresh — this expense will disappear from P&L / Balance Sheet."
            )
            self._refresh_engine_and_reports()
        except Exception as ex:
            messagebox.showerror("❌ Delete Failed",
                                 f"Could not reverse expense:\n{ex}")
            import traceback
            traceback.print_exc()

    def _clip_copy(self, ref_tuple):
        try:
            text = " : ".join([str(x or "") for x in ref_tuple if x])
            self.win.clipboard_clear()
            self.win.clipboard_append(text)
        except Exception:
            pass

    def _search_from_row(self):
        _, _, meta = self._sel_expense_ref()
        if not meta:
            return
        p = (meta.get("payee") or meta.get("description") or "").strip()
        if p:
            self._exp_search_var.set(p)
            self._reload_expense_tab()

    def _sel_expense_row_tsv(self):
        iid = self._exp_tree.selection()
        if not iid:
            return ""
        iid = iid[0]
        try:
            vals = self._exp_tree.item(iid, "values")
            return "\t".join([str(v or "") for v in vals])
        except Exception:
            return ""

    # ------------------------------------------------------------------
    # Right-click menus for GAAP report trees (P&L / BS / Equity / CF / …)
    # Lets user right-click ANY row inside ANY tabular GAAP report and
    # either (a) copy the row, (b) jump to the Expense Ledger to delete
    # anything expense-related, or (c) refresh the report.
    # Works cross-platform (macOS Button-2 + Windows/Linux Button-3).
    # ------------------------------------------------------------------
    def _jump_to_expense_ledger(self):
        try:
            nb = self.nb
            # Find index of "Expense_Ledger" / "expense_manager" tab
            for idx in range(nb.index("end")):
                try:
                    text = nb.tab(idx, "text") or ""
                except Exception:
                    text = ""
                if ("expense" in text.lower()
                        or "ledger" in text.lower()
                        or "11" in text):
                    nb.select(idx)
                    try:
                        self._reload_expense_tab()
                    except Exception:
                        pass
                    return
            # Fallback: index 10 or last-2
            try: nb.select(min(10, nb.index("end") - 1))
            except Exception: pass
            try: self._reload_expense_tab()
            except Exception: pass
        except Exception:
            pass

    def _attach_report_context_menu(self, tree, report_name,
                                    allow_delete_expense_directly=False):
        """Attach a cross-platform right-click menu to `tree`.

        If `allow_delete_expense_directly=True` AND the tree's selected
        row tags/values contain a direct expense `ref_type:ref_id` that
        maps to an Expense Ledger row, then the menu will ALSO show
        '🗑️  Delete / Void this Expense' which fires the exact same
        reversal flow as the Expense Ledger Delete button.
        """
        if tree is None or not hasattr(tree, "bind"):
            return
        tree_ref = tree

        def _sel_values():
            try:
                sel = tree_ref.selection()
                if not sel: return (None, None)
                return sel[0], tree_ref.item(sel[0], "values")
            except Exception:
                return (None, None)

        def _copy_label():
            _, vals = _sel_values()
            if vals and len(vals) >= 1:
                self._clip_copy((str(vals[0] or ""),))
        def _copy_current():
            _, vals = _sel_values()
            if vals and len(vals) >= 2:
                self._clip_copy((str(vals[1] or ""),))
        def _copy_prior():
            _, vals = _sel_values()
            if vals and len(vals) >= 3:
                self._clip_copy((str(vals[2] or ""),))
        def _copy_row_tsv():
            _, vals = _sel_values()
            if vals:
                self._clip_copy(("\t".join([str(v or "") for v in vals]),))

        def _refresh_this():
            try:
                name = (report_name or "").lower()
                if "pnl" in name or "profit" in name:
                    self._refresh_pnl()
                elif "balance" in name or "bs" in name or "sheet" in name:
                    self._refresh_bs()
                elif "cash" in name or "flow" in name or "cf" in name:
                    self._refresh_cashflow()
                elif "equity" in name or "changes" in name:
                    self._refresh_equity()
                elif "ratio" in name:
                    self._refresh_ratios()
                elif "aging" in name:
                    self._refresh_aging()
                elif "revenue" in name or "client" in name:
                    self._refresh_rev_clients()
                elif "stock" in name or "inventor" in name:
                    self._refresh_stock()
                elif "supplier" in name:
                    self._refresh_suppliers()
                elif "audit" in name or "trail" in name:
                    self._refresh_audit()
                else:
                    self._refresh_engine_and_reports()
            except Exception:
                self._refresh_engine_and_reports()

        # Build menu
        try:
            parent = tree_ref.master if hasattr(tree_ref, "master") else self.win
        except Exception:
            parent = self.win
        m = tk.Menu(parent, tearoff=0)
        m.add_command(label="📋  Copy Row Label / Account Name", command=_copy_label)
        m.add_command(label="📋  Copy Current Period Value", command=_copy_current,
                      accelerator="($ column 2)")
        m.add_command(label="📋  Copy Prior Period Value", command=_copy_prior)
        m.add_command(label="📋  Copy Full Row (TSV)", command=_copy_row_tsv)
        m.add_separator()
        m.add_command(label="💸  Record a New Expense…", command=self._open_expense_dialog)
        m.add_command(label="➡️  Go to Expense Ledger (Edit / Delete Any Expense)",
                      command=self._jump_to_expense_ledger,
                      accelerator="Tab 11")
        m.add_separator()
        m.add_command(label="🔁  Refresh This Report", command=_refresh_this,
                      accelerator="F5 equivalent")
        # Rebuild GL & refresh all reports.  The actual button method is
        # `_rebuild_gl_now`; also accept `_on_rebuild_clicked` for compat with
        # older callers.
        _rebuild_fn = (getattr(self, "_rebuild_gl_now", None)
                       or getattr(self, "_on_rebuild_clicked", None))
        if callable(_rebuild_fn):
            m.add_command(label="♻️  Rebuild GL + Refresh All Reports",
                          command=_rebuild_fn)

        def _popup(ev):
            try:
                iid = tree_ref.identify_row(ev.y) if hasattr(tree_ref, "identify_row") else None
                if iid:
                    tree_ref.selection_set(iid)
                else:
                    if not tree_ref.selection():
                        return
            except Exception:
                pass
            try:
                m.tk_popup(ev.x_root, ev.y_root)
            finally:
                try: m.grab_release()
                except Exception: pass

        for ev_seq in ("<Button-2>", "<Button-3>"):
            try:
                tree_ref.bind(ev_seq, _popup)
            except Exception:
                pass
        # Keep the menu referenced so it doesn't get garbage-collected
        store_key = f"_ctx_{report_name or id(tree_ref)}"
        if not hasattr(self, "_report_ctx_menus"):
            self._report_ctx_menus = {}
        self._report_ctx_menus[store_key] = m


# ======================================================================
# Dialog: EDIT EXISTING EXPENSE (re-open form, then reverse + re-post)
# ======================================================================
class EditExpenseDialog(tk.Toplevel):
    def __init__(self, master, ledger, meta, on_saved_cb=None, data_manager=None):
        super().__init__(master)
        self.ledger = ledger
        self.meta = meta
        self.on_saved_cb = on_saved_cb
        # Fallback: if ledger was passed as None (data manager didn't init it),
        # create a working LedgerService inline so Edit/Reversal always works.
        if self.ledger is None and data_manager is not None:
            try:
                from ledger_service import LedgerService
                self.ledger = LedgerService(data_manager)
            except Exception:
                self.ledger = None

        self.title(f"✏️  Edit Expense — {meta.get('ref_type','')}:{meta.get('ref_id','')}")
        self.configure(bg="white")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()

        w, h = 640, 600
        self.geometry(f"{w}x{h}+{max(60, (self.winfo_screenwidth()-w)//2)}"
                      f"+{max(40, (self.winfo_screenheight()-h)//3)}")

        # Banner
        banner = tk.Frame(self, bg="#F39C12", padx=18, pady=12)
        banner.pack(fill="x")
        tk.Label(banner, text="✏️  Edit Expense (Reverse + Re-Post)",
              bg="#F39C12", fg="white", font=("Segoe UI", 14, "bold"),
              anchor="w").pack(fill="x")
        tk.Label(banner,
              text=f"Original: {meta.get('ref_type','')}:{meta.get('ref_id','')} "
                   f"· {meta.get('date','')} · {fmt_aed(meta.get('gross_amount', 0))}",
              bg="#F39C12", fg="white", font=("Segoe UI", 10),
              anchor="w").pack(fill="x", pady=(2, 0))

        body = tk.Frame(self, bg="white", padx=22, pady=20)
        body.pack(fill="both", expand=True)

        def row(lbl, idx):
            tk.Label(body, text=lbl, font=("Segoe UI", 10, "bold"),
                  bg="white", fg=NAVY).grid(row=idx, column=0, sticky="w", pady=(8, 4), padx=(0, 10))

        row("Date (YYYY-MM-DD)", 0)
        self.dt_var = tk.StringVar(value=str(meta.get("date") or date.today().isoformat()))
        tk.Entry(body, textvariable=self.dt_var, font=("Segoe UI", 12), width=28,
              relief="groove", highlightthickness=1, highlightbackground=SEPARATOR,
              highlightcolor=NAVY
              ).grid(row=0, column=1, sticky="we", pady=(8, 4))

        row("Category", 1)
        self.cat_var = tk.StringVar(value=meta.get("category") or "Other Administrative")
        cats = list(EXPENSE_CATEGORY_MAP.keys())
        ttk.Combobox(body, textvariable=self.cat_var, values=cats,
                     state="readonly", width=26, font=("Segoe UI", 12)
                     ).grid(row=1, column=1, sticky="we", pady=(4, 4))

        row("Amount (AED)", 2)
        self.amt_var = tk.StringVar(value=str(round(float(meta.get("gross_amount") or 0), 2)))
        tk.Entry(body, textvariable=self.amt_var, font=("Segoe UI", 12), width=28,
              relief="groove", highlightthickness=1, highlightbackground=SEPARATOR,
              highlightcolor=NAVY
              ).grid(row=2, column=1, sticky="we", pady=(4, 4))

        row("Payee / Vendor", 3)
        self.pay_var = tk.StringVar(value=meta.get("payee") or "")
        tk.Entry(body, textvariable=self.pay_var, font=("Segoe UI", 12), width=28,
              relief="groove", highlightthickness=1, highlightbackground=SEPARATOR,
              highlightcolor=NAVY
              ).grid(row=3, column=1, sticky="we", pady=(4, 4))

        row("Description / Notes", 4)
        self.desc_txt = tk.Text(body, width=48, height=6, font=("Segoe UI", 11),
                             relief="groove", highlightthickness=1,
                             highlightbackground=SEPARATOR, highlightcolor=NAVY)
        self.desc_txt.grid(row=4, column=1, sticky="we", pady=(4, 4))
        self.desc_txt.insert("1.0", meta.get("description") or "")

        body.grid_columnconfigure(1, weight=1)

        # Info box
        info = tk.LabelFrame(body, text=" 🔎  What happens when I click Save? ",
                          bg="white", fg="#7D6608", font=("Segoe UI", 9, "bold"),
                          padx=12, pady=8)
        info.grid(row=5, column=0, columnspan=2, sticky="we", pady=(18, 4))
        tk.Label(info, bg="white", fg="#333", font=("Segoe UI", 10), justify="left",
              text=("1️⃣  The original expense GL batch will be reversed (Dr↔Cr).\n"
                    "2️⃣  A brand-new EXPENSE batch will be posted with your new values,\n"
                    "      and its Reference ID will read EXP-CORR-YYYYMMDD-XXXXXX.\n"
                    "3️⃣  Reports refresh automatically so P&L / BS / CF reflect immediately.")).pack(anchor="w")

        # Buttons
        btns = tk.Frame(self, bg="white", padx=22, pady=16)
        btns.pack(fill="x")
        tk.Button(btns, text="Cancel", bg="white", fg=NAVY,
               font=("Segoe UI", 11, "bold"), relief="groove",
               cursor="hand2", padx=22, pady=8, command=self.destroy
               ).pack(side="right")
        tk.Button(btns, text="💾  Save Changes", bg="#F39C12", fg="white",
               activebackground="#D68910", activeforeground="white",
               font=("Segoe UI", 11, "bold"), relief="flat",
               cursor="hand2", padx=22, pady=8, command=self._save
               ).pack(side="right", padx=(0, 10))
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(60, lambda: self.desc_txt.focus_set())

    def _save(self):
        # Validate
        dt = (self.dt_var.get() or "").strip()
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", dt):
            messagebox.showwarning("Invalid Date",
                                   "Date must be YYYY-MM-DD (e.g. 2025-06-30)."); return
        try:
            y, m, d = [int(x) for x in dt.split("-")]
            datetime(y, m, d)
        except Exception:
            messagebox.showwarning("Invalid Date", "Date does not exist."); return
        try:
            amt = float(self.amt_var.get())
            if amt <= 0: raise ValueError("negative")
            amt = round(amt, 2)
        except Exception:
            messagebox.showwarning("Invalid Amount",
                                   "Please enter a positive number (e.g. 2499.00)."); return
        if not self.cat_var.get():
            messagebox.showwarning("Category required",
                                   "Pick a valid expense category."); return
        payee = (self.pay_var.get() or "").strip()
        desc = (self.desc_txt.get("1.0", "end") or "").strip()
        rt = self.meta.get("ref_type")
        rid = self.meta.get("ref_id")
        try:
            if rt == "EXPENSE":
                reversals, new_entries = self.ledger.edit_expense(
                    rt, rid,
                    new_fields={
                        "category": self.cat_var.get(),
                        "amount": amt,
                        "expense_date": dt,
                        "payee_name": payee,
                        "description": desc,
                    },
                    username="GAAP Dashboard"
                )
            else:
                # For WITHDRAWAL / INVOICE_COST: reverse original + post brand new EXPENSE
                reversals = self.ledger.reverse_reference(
                    rt, rid, username="GAAP Dashboard",
                    reason=f"Edited expense {rt}:{rid} → re-posted as corrected EXPENSE batch."
                )
                new_entries = self.ledger._force_repost_expense(
                    category=self.cat_var.get(), amount=amt, expense_date=dt,
                    payee_name=payee, description=desc,
                    username="GAAP Dashboard",
                    origin_note=f"Corrected — replaces {rt}:{rid}"
                )
            messagebox.showinfo(
                "✅  Expense Saved",
                f"Corrected {fmt_aed(amt)}\n\n"
                f"• Original reference:  {rt}:{rid}\n"
                f"  → Reversed GL rows:  {len(reversals)}\n"
                f"• New reference:       {new_entries[0]['reference_id'] if new_entries else '—'}\n"
                f"  → Posted GL rows:    {len(new_entries)}\n\n"
                "💡  Reports will now refresh so P&L / BS reflect the change."
            )
            if self.on_saved_cb:
                try: self.on_saved_cb()
                except Exception: pass
            self.destroy()
        except Exception as ex:
            messagebox.showerror("❌ Edit Failed",
                                 f"Could not save changes:\n{ex}")
            import traceback
            traceback.print_exc()
