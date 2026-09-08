"""
GSTR-9 & GSTR-9C Preparation Module
====================================
A plug-in Streamlit module for your existing "GST Data Management &
Reconciliation Tool" (gsttool2.py). It does NOT run on its own inside your
main app - you import it and call `render()` from a new sidebar option.

WHAT IT DOES
------------
For a selected Financial Year, it opens the SAME SQLite database your main
app is already using (Sales, Purchase, GSTR-1, GSTR-3B, GSTR-2B tables) and:

  1. Auto-computes every GSTR-9 / GSTR-9C figure that CAN be derived from
     the data you have already entered (outward supply totals, ITC availed
     as per 3B, ITC as per GSTR-2B, tax paid, rate-wise turnover, etc.)
  2. Leaves an editable field, pre-filled with 0 (or a suggested value),
     for every figure your app has no source data for (exports/SEZ/deemed
     exports flags, audited financial-statement turnover, expense-head
     wise ITC, ITC reversals under Rule 37/42/43 etc., HSN summary).
  3. Saves every manual entry back into the SAME database (a new table,
     `gstr9_manual_entries`) so nothing you type is lost between sessions.
  4. Exports a clean, ready-to-use Excel workbook with both GSTR-9 and
     GSTR-9C figures, which you (or your CA) can use to fill the actual
     GST portal offline utility.

IMPORTANT / HONEST LIMITATIONS
-------------------------------
- Your current database schema does not capture: HSN codes, export/SEZ/
  deemed-export flags on sales, reverse-charge flags on purchases, an
  Inputs/Capital-Goods/Input-Services split on purchases, or Cess. Those
  sections are therefore manual-entry only (clearly marked below).
- GSTR-9C requires figures from your AUDITED financial statements (books
  of account) which live outside this app by definition - those fields
  are always manual.
- This tool follows the standard, well-established structure of GSTR-9 /
  GSTR-9C. Please have your CA cross-check the final numbers against the
  official portal utility before filing - GSTR-9C in particular requires
  certification and sign-off outside of any software.

HOW TO INTEGRATE INTO gsttool2.py
-----------------------------------
1. Save this file as `gstr9_9c.py` in the SAME folder as gsttool2.py.

2. Near the top of gsttool2.py, with the other imports, add:

       import gstr9_9c

3. Find this line in gsttool2.py:

       module = st.sidebar.radio(
           "MODULE",
           [
               "🏠 Dashboard",
               ...
               "⚙️ Settings",
           ],

   and add a new entry to that list, e.g. right after "🔄 Reconciliation":

               "📑 GSTR-9 / 9C",

4. Anywhere alongside the other `elif module == "...":` blocks (e.g. right
   before the "EXPORT" section), add:

       elif module == "📑 GSTR-9 / 9C":
           gstr9_9c.render(get_db_path(), load_company_profile)

   `get_db_path()` and `load_company_profile` already exist in
   gsttool2.py, so this module automatically works on whichever company
   database is currently active / logged in - no extra wiring needed.

That's it. Every client's GSTR-9/9C data stays inside their own private
database, exactly like the rest of your app.
"""

import os
import io
import sqlite3
from datetime import date, datetime

import pandas as pd
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# STYLE CONSTANTS (Excel export)
# ============================================================

THIN = Side(style="thin", color="B8BEC9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="2F5FED")
SUB_FILL = PatternFill("solid", fgColor="EEF1FA")
TOTAL_FILL = PatternFill("solid", fgColor="DCE4FA")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=12)
SECTION_FONT = Font(bold=True, color="FFFFFF", size=11)
BOLD = Font(bold=True)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT_WRAP = Alignment(horizontal="left", vertical="center", wrap_text=True)


# ============================================================
# LOCAL DATABASE HELPERS
# (Self-contained on purpose: this module only ever needs a db_path,
#  so it can be dropped into any app without importing internals.)
# ============================================================

def _read(db_path, query, params=()):
    if not db_path or not os.path.exists(db_path):
        return pd.DataFrame()
    conn = sqlite3.connect(db_path)
    try:
        df = pd.read_sql_query(query, conn, params=params)
    except Exception:
        df = pd.DataFrame()
    conn.close()
    return df


def _run(db_path, query, params=()):
    conn = sqlite3.connect(db_path)
    cur = conn.cursor()
    cur.execute(query, params)
    conn.commit()
    conn.close()


def _ensure_manual_table(db_path):
    _run(db_path, """
        CREATE TABLE IF NOT EXISTS gstr9_manual_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            financial_year TEXT NOT NULL,
            field_key TEXT NOT NULL,
            field_value TEXT,
            UNIQUE(financial_year, field_key)
        )
    """)


def _load_manual(db_path, fy):
    _ensure_manual_table(db_path)
    df = _read(
        db_path,
        "SELECT field_key, field_value FROM gstr9_manual_entries WHERE financial_year = ?",
        (fy,)
    )
    if df.empty:
        return {}
    return dict(zip(df["field_key"], df["field_value"]))


def _save_manual(db_path, fy, key, value):
    _ensure_manual_table(db_path)
    _run(db_path, """
        INSERT INTO gstr9_manual_entries (financial_year, field_key, field_value)
        VALUES (?, ?, ?)
        ON CONFLICT(financial_year, field_key) DO UPDATE SET field_value = excluded.field_value
    """, (fy, key, str(value)))


# ============================================================
# FINANCIAL YEAR HELPERS
# ============================================================

def _fy_options():
    """Last 6 Indian financial years, most recent first, e.g. '2024-25'."""
    this_year = datetime.now().year
    start = this_year if datetime.now().month >= 4 else this_year - 1
    return [f"{y}-{str(y + 1)[-2:]}" for y in range(start, start - 6, -1)]


def _fy_bounds(fy_label):
    """'2024-25' -> (date(2024,4,1), date(2025,3,31))"""
    try:
        start_year = int(fy_label[:4])
    except Exception:
        start_year = datetime.now().year
    return date(start_year, 4, 1), date(start_year + 1, 3, 31)


def _fy_periods(fy_label):
    """List of 'YYYY-MM' strings for the 12 months of the FY (Apr-Mar)."""
    start, _ = _fy_bounds(fy_label)
    periods = []
    y, m = start.year, start.month
    for _ in range(12):
        periods.append(f"{y:04d}-{m:02d}")
        m += 1
        if m > 12:
            m = 1
            y += 1
    return periods


# ============================================================
# PULL DATA FROM THE SHARED DATABASE
# ============================================================

def _pull_all(db_path, fy):
    start, end = _fy_bounds(fy)
    periods = _fy_periods(fy)
    placeholders = ",".join("?" for _ in periods)

    sales = _read(
        db_path,
        "SELECT * FROM gst_invoices WHERE date(invoice_date) BETWEEN date(?) AND date(?)",
        (str(start), str(end))
    )
    purchases = _read(
        db_path,
        "SELECT * FROM purchase_invoices WHERE date(invoice_date) BETWEEN date(?) AND date(?)",
        (str(start), str(end))
    )
    gstr3b = _read(
        db_path,
        f"SELECT * FROM gstr3b_summary WHERE return_period IN ({placeholders})",
        tuple(periods)
    )
    gstr2b = _read(
        db_path,
        f"SELECT * FROM gstr_2b WHERE return_period IN ({placeholders})",
        tuple(periods)
    )

    for df in (sales, purchases, gstr3b, gstr2b):
        for col in ("taxable_value", "cgst", "sgst", "igst", "total_amount",
                    "output_tax", "itc_claimed", "tax_paid_cash", "late_fee"):
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return {"sales": sales, "purchases": purchases, "gstr3b": gstr3b, "gstr2b": gstr2b, "periods": periods}


def _sum(df, col):
    if df.empty or col not in df.columns:
        return 0.0
    return float(df[col].sum())


# ============================================================
# GENERIC ROW-INPUT WIDGET
# One "row" of a GST table = Taxable Value / CGST / SGST / IGST / Cess.
# Some columns are computed (read-only), others are manual (editable and
# persisted). Returns the final values used in totals + the set of keys
# that need saving.
# ============================================================

COLS = ("Taxable Value", "CGST", "SGST", "IGST", "Cess")


def _table_row(manual, keystore, prefix, label, computed=None, note=None):
    """
    computed: dict of {col_name: value} for columns that are AUTO from your
              data (shown read-only). Any column in COLS not in `computed`
              is rendered as an editable manual input.
    Returns dict {col_name: value}.
    """
    computed = computed or {}
    st.markdown(f"**{label}**" + (f"  \n<span style='color:#6b7280;font-size:12.5px'>{note}</span>" if note else ""), unsafe_allow_html=True)
    cols = st.columns(len(COLS))
    result = {}
    for i, col_name in enumerate(COLS):
        if col_name in computed:
            val = float(computed[col_name])
            cols[i].metric(col_name, f"₹{val:,.0f}")
            result[col_name] = val
        else:
            key = f"{prefix}__{col_name}".replace(" ", "_")
            default_val = float(manual.get(key, 0.0) or 0.0)
            val = cols[i].number_input(
                col_name, value=default_val, step=1000.0, format="%.2f", key=key
            )
            result[col_name] = val
            keystore.append(key)
    return result


def _save_button(db_path, fy, keystore, label="💾 Save manual entries for this section"):
    if st.button(label, key=f"save_{fy}_{'_'.join(keystore[:1])}_{len(keystore)}"):
        for k in keystore:
            _save_manual(db_path, fy, k, st.session_state.get(k, 0.0))
        st.success("Saved. These values will be remembered next time you open this Financial Year.")


# ============================================================
# GSTR-9 COMPUTATION
# ============================================================

def compute_gstr9(data, manual):
    sales = data["sales"]
    purchases = data["purchases"]
    gstr3b = data["gstr3b"]
    gstr2b = data["gstr2b"]

    g = {"manual_defaults": {}}

    # ---- classify sales: B2B (valid 15-char GSTIN) vs B2C ----
    if not sales.empty:
        gstin_series = sales["gstin"].fillna("").astype(str).str.strip()
        b2b_mask = gstin_series.str.len() == 15
    else:
        b2b_mask = pd.Series([], dtype=bool)

    b2c = sales[~b2b_mask] if not sales.empty else pd.DataFrame()
    b2b = sales[b2b_mask] if not sales.empty else pd.DataFrame()

    def tot(df):
        return {
            "Taxable Value": _sum(df, "taxable_value"),
            "CGST": _sum(df, "cgst"),
            "SGST": _sum(df, "sgst"),
            "IGST": _sum(df, "igst"),
            "Cess": 0.0,
        }

    g["4A_B2C"] = tot(b2c)     # Supplies to unregistered
    g["4B_B2B"] = tot(b2b)     # Supplies to registered

    return g


# ============================================================
# MAIN RENDER FUNCTION
# ============================================================

def render(db_path, load_company_profile=None, section="gstr9"):

    section = "gstr9c" if section == "gstr9c" else "gstr9"
    page_title = "GSTR-9C Reconciliation" if section == "gstr9c" else "GSTR-9 Annual Return"

    st.markdown(
        '<div style="font-size:26px;font-weight:800;color:#1a1d29;border-bottom:3px solid #2f5fed;'
        f'display:inline-block;padding-bottom:4px;margin-bottom:14px;">📑 {page_title}</div>',
        unsafe_allow_html=True
    )

    if not db_path or not os.path.exists(db_path):
        st.warning("No active company database found. Please select a storage folder / company first.")
        return

    profile = {}
    if load_company_profile:
        try:
            profile = load_company_profile()
        except Exception:
            profile = {}

    top1, top2, top3 = st.columns([2, 2, 3])
    with top1:
        fy = st.selectbox("Financial Year", _fy_options(), key="g9_fy")
    with top2:
        st.text_input("GSTIN", value=profile.get("gstin", ""), disabled=True, key="g9_gstin_display")
    with top3:
        st.text_input("Legal Name", value=profile.get("company_name", ""), disabled=True, key="g9_name_display")

    st.caption(
        "Figures below are auto-pulled from your Sales, Purchase, GSTR-1, GSTR-3B and GSTR-2B "
        "records for this Financial Year wherever possible. Fields your data can't answer are "
        "editable and get saved automatically for next time."
    )

    manual = _load_manual(db_path, fy)
    data = _pull_all(db_path, fy)
    g9 = compute_gstr9(data, manual)

    # Keep the shared calculations and saved-entry keys intact while showing
    # only the category selected from the main sidebar.
    if section == "gstr9":
        st.markdown(
            "<style>"
            "button[data-baseweb='tab']:nth-child(8),"
            "button[data-baseweb='tab']:nth-child(9),"
            "button[data-baseweb='tab']:nth-child(10) { display: none; }"
            "</style>",
            unsafe_allow_html=True
        )
        st.caption("GSTR-9 annual return tables")
    else:
        st.markdown(
            "<style>"
            "button[data-baseweb='tab']:nth-child(2),"
            "button[data-baseweb='tab']:nth-child(3),"
            "button[data-baseweb='tab']:nth-child(4),"
            "button[data-baseweb='tab']:nth-child(5),"
            "button[data-baseweb='tab']:nth-child(6),"
            "button[data-baseweb='tab']:nth-child(7) { display: none; }"
            "</style>",
            unsafe_allow_html=True
        )
        st.caption("GSTR-9C reconciliation and certification working papers")

    tabs = st.tabs([
        "Summary",
        "GSTR-9 · Table 4 (Outward Supplies)",
        "GSTR-9 · Table 5 (Non-Taxable Supplies)",
        "GSTR-9 · Table 6 (ITC Availed)",
        "GSTR-9 · Table 7 (ITC Reversed)",
        "GSTR-9 · Table 8 (ITC vs GSTR-2B)",
        "GSTR-9 · Table 9 (Tax Paid)",
        "GSTR-9C · Turnover Reconciliation",
        "GSTR-9C · Tax Paid Reconciliation",
        "GSTR-9C · ITC Reconciliation",
        "Export to Excel",
    ])

    # ---------------- SUMMARY ----------------
    with tabs[0]:
        sales_taxable = _sum(data["sales"], "taxable_value")
        sales_tax = _sum(data["sales"], "cgst") + _sum(data["sales"], "sgst") + _sum(data["sales"], "igst")
        itc_3b = _sum(data["gstr3b"], "itc_claimed")
        itc_2b = _sum(data["gstr2b"], "cgst") + _sum(data["gstr2b"], "sgst") + _sum(data["gstr2b"], "igst")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Outward Taxable Value (books)", f"₹{sales_taxable:,.0f}")
        c2.metric("Output Tax (books)", f"₹{sales_tax:,.0f}")
        c3.metric("ITC availed per GSTR-3B", f"₹{itc_3b:,.0f}")
        c4.metric("ITC as per GSTR-2B", f"₹{itc_2b:,.0f}")
        st.info(
            "Use the tabs above to review and complete each table. Fields shown as metrics are "
            "auto-computed from your data; number-input boxes are editable and saved per Financial Year."
        )

    # ---------------- TABLE 4 ----------------
    with tabs[1]:
        st.subheader("Table 4 — Outward and inward supplies on which tax is payable")
        keys = []
        row_A = _table_row(manual, keys, f"{fy}_4A",
                            "A. Supplies made to un-registered persons (B2C)",
                            computed=g9["4A_B2C"])
        row_B = _table_row(manual, keys, f"{fy}_4B",
                            "B. Supplies made to registered persons (B2B)",
                            computed=g9["4B_B2B"])
        row_C = _table_row(manual, keys, f"{fy}_4C",
                            "C. Zero rated supply (Export) on payment of tax (except SEZ)",
                            note="Not tracked in your Sales register — enter manually.")
        row_D = _table_row(manual, keys, f"{fy}_4D",
                            "D. Supply to SEZs on payment of tax",
                            note="Not tracked in your Sales register — enter manually.")
        row_E = _table_row(manual, keys, f"{fy}_4E",
                            "E. Deemed Exports",
                            note="Not tracked in your Sales register — enter manually.")
        row_F = _table_row(manual, keys, f"{fy}_4F",
                            "F. Advances on which tax paid but invoice not issued")
        row_G = _table_row(manual, keys, f"{fy}_4G",
                            "G. Inward supplies on which tax payable on reverse charge")
        row_G1 = _table_row(manual, keys, f"{fy}_4G1",
                             "G1. Supplies where e-commerce operator pays tax u/s 9(5)")

        rows_H = [row_A, row_B, row_C, row_D, row_E, row_F, row_G, row_G1]
        H = {c: sum(r[c] for r in rows_H) for c in COLS}
        _table_row(manual, [], f"{fy}_4H_display", "H. Sub-total (A to G1)", computed=H)

        row_I = _table_row(manual, keys, f"{fy}_4I", "I. Credit Notes issued (–)")
        row_J = _table_row(manual, keys, f"{fy}_4J", "J. Debit Notes issued (+)")
        row_K = _table_row(manual, keys, f"{fy}_4K", "K. Supplies / tax declared through Amendments (+)")
        row_L = _table_row(manual, keys, f"{fy}_4L", "L. Supplies / tax reduced through Amendments (–)")

        M = {c: (-row_I[c] + row_J[c] + row_K[c] - row_L[c]) for c in COLS}
        _table_row(manual, [], f"{fy}_4M_display", "M. Net adjustment (–I + J + K – L)", computed=M)

        N = {c: H[c] + M[c] for c in COLS}
        _table_row(manual, [], f"{fy}_4N_display", "N. Supplies and advances on which tax is payable (H + M)", computed=N)

        g9["4G_manual"] = row_G
        g9["4G1_manual"] = row_G1
        g9["4N"] = N

        _save_button(db_path, fy, keys)

    # ---------------- TABLE 5 ----------------
    with tabs[2]:
        st.subheader("Table 5 — Outward supplies on which tax is NOT payable")
        nil_pool = float(data["sales"][data["sales"]["gst_rate"] == 0]["taxable_value"].sum()) if not data["sales"].empty and "gst_rate" in data["sales"].columns else 0.0
        st.caption(
            f"💡 Suggestion: your Sales register has ₹{nil_pool:,.0f} of invoices recorded at a 0% GST "
            "rate. Split this figure across D/E/F below as appropriate (or leave as-is if it belongs "
            "entirely under Exempted)."
        )
        keys = []
        row_A = _table_row(manual, keys, f"{fy}_5A", "A. Zero rated supply (Export) without payment of tax")
        row_B = _table_row(manual, keys, f"{fy}_5B", "B. Supply to SEZs without payment of tax")
        row_C = _table_row(manual, keys, f"{fy}_5C", "C. Supplies on which tax payable by recipient (RCM)")
        row_D = _table_row(manual, keys, f"{fy}_5D", "D. Exempted")
        row_E = _table_row(manual, keys, f"{fy}_5E", "E. Nil Rated")
        row_F = _table_row(manual, keys, f"{fy}_5F", "F. Non-GST supply (includes 'no supply')")

        rows_G = [row_A, row_B, row_C, row_D, row_E, row_F]
        G = {c: sum(r[c] for r in rows_G) for c in COLS}
        _table_row(manual, [], f"{fy}_5G_display", "G. Sub-total (A to F)", computed=G)

        row_H = _table_row(manual, keys, f"{fy}_5H", "H. Credit Notes issued (–)")
        row_I = _table_row(manual, keys, f"{fy}_5I", "I. Debit Notes issued (+)")
        row_J = _table_row(manual, keys, f"{fy}_5J", "J. Supplies / tax declared through Amendments (+)")
        row_K = _table_row(manual, keys, f"{fy}_5K", "K. Supplies / tax reduced through Amendments (–)")

        L = {c: (-row_H[c] + row_I[c] + row_J[c] - row_K[c]) for c in COLS}
        _table_row(manual, [], f"{fy}_5L_display", "L. Net adjustment (–H + I + J – K)", computed=L)

        M = {c: G[c] + L[c] for c in COLS}
        _table_row(manual, [], f"{fy}_5M_display", "M. Turnover on which tax is not payable (G + L)", computed=M)

        N4 = g9.get("4N", {c: 0.0 for c in COLS})
        G4 = g9.get("4G_manual", {c: 0.0 for c in COLS})
        G1_4 = g9.get("4G1_manual", {c: 0.0 for c in COLS})
        N5 = {c: N4[c] + M[c] - G4[c] - G1_4[c] for c in COLS}
        _table_row(manual, [], f"{fy}_5N_display",
                   "N. Total Turnover (4N + 5M – 4G – 4G1)", computed=N5)

        g9["5N_total_turnover"] = N5
        _save_button(db_path, fy, keys)

    # ---------------- TABLE 6 ----------------
    with tabs[3]:
        st.subheader("Table 6 — Details of ITC availed during the financial year")
        itc_3b_total = _sum(data["gstr3b"], "itc_claimed")
        purchase_eligible = data["purchases"][data["purchases"].get("itc_eligible", "Yes") == "Yes"] if not data["purchases"].empty else pd.DataFrame()
        purchase_itc = {
            "Taxable Value": _sum(purchase_eligible, "taxable_value"),
            "CGST": _sum(purchase_eligible, "cgst"),
            "SGST": _sum(purchase_eligible, "sgst"),
            "IGST": _sum(purchase_eligible, "igst"),
            "Cess": 0.0,
        }
        keys = []
        itc_3b_split = {
            "Taxable Value": 0.0,
            "CGST": itc_3b_total / 3,
            "SGST": itc_3b_total / 3,
            "IGST": itc_3b_total / 3,
            "Cess": 0.0,
        }
        A = _table_row(manual, [], f"{fy}_6A_display",
                        "A. Total ITC availed through GSTR-3B (sum of Table 4A of all 3B returns)",
                        computed=itc_3b_split,
                        note="Your GSTR-3B summary only stores one aggregate ITC figure per return, "
                             "not split by tax head — shown here divided equally across CGST/SGST/IGST "
                             "as a placeholder. Edit Table 6 rows below with real splits if you have them.")
        A1 = _table_row(manual, keys, f"{fy}_6A1",
                         "A1. ITC of a preceding FY availed in this FY (subset of A, other than reclaim)")
        A2 = {c: A[c] - A1[c] for c in COLS}
        _table_row(manual, [], f"{fy}_6A2_display", "A2. Net ITC of the financial year (A – A1)", computed=A2)

        B = _table_row(manual, [], f"{fy}_6B_display",
                        "B. Inward supplies (other than imports/RCM, incl. from SEZ) — from Purchase register",
                        computed=purchase_itc,
                        note="Auto = eligible Purchase register total. Official form splits this into Inputs / Capital Goods / Input Services — your app doesn't tag purchases by type, so it's shown as one figure.")
        C = _table_row(manual, keys, f"{fy}_6C", "C. RCM ITC — inward supplies from UNregistered persons")
        D = _table_row(manual, keys, f"{fy}_6D", "D. RCM ITC — inward supplies from registered persons")
        E = _table_row(manual, keys, f"{fy}_6E", "E. Import of goods (incl. from SEZ)")
        F = _table_row(manual, keys, f"{fy}_6F", "F. Import of services (excl. from SEZ)")
        Gr = _table_row(manual, keys, f"{fy}_6G", "G. Input Tax Credit received from ISD")
        H = _table_row(manual, keys, f"{fy}_6H", "H. Other / miscellaneous ITC availed")

        rows_I = [B, C, D, E, F, Gr, H]
        I = {c: sum(r[c] for r in rows_I) for c in COLS}
        _table_row(manual, [], f"{fy}_6I_display", "I. Sub-total (B to H)", computed=I)

        J = {c: I[c] - A2[c] for c in COLS}
        _table_row(manual, [], f"{fy}_6J_display", "J. Difference (I – A2)", computed=J)

        K = _table_row(manual, keys, f"{fy}_6K", "K. Transition Credit through TRAN-1")
        L = _table_row(manual, keys, f"{fy}_6L", "L. Transition Credit through TRAN-2")
        M = _table_row(manual, keys, f"{fy}_6M", "M. ITC availed through ITC-01 / ITC-02 / ITC-02A")

        rows_N = [K, L, M]
        Nn = {c: sum(r[c] for r in rows_N) for c in COLS}
        _table_row(manual, [], f"{fy}_6N_display", "N. Sub-total (K to M)", computed=Nn)

        O = {c: I[c] + Nn[c] for c in COLS}
        _table_row(manual, [], f"{fy}_6O_display", "O. Total ITC availed (I + N)", computed=O)

        g9["6B"] = B
        g9["6O_total_itc_availed"] = O
        _save_button(db_path, fy, keys)

    # ---------------- TABLE 7 ----------------
    with tabs[4]:
        st.subheader("Table 7 — ITC Reversed and Ineligible ITC")
        st.caption("None of these are tracked automatically — please fill in from your books / 3B reversal entries.")
        keys = []
        rA = _table_row(manual, keys, f"{fy}_7A", "As per Rule 37")
        rA1 = _table_row(manual, keys, f"{fy}_7A1", "As per Rule 37A")
        rB = _table_row(manual, keys, f"{fy}_7B", "As per Rule 38")
        rC = _table_row(manual, keys, f"{fy}_7C", "As per Rule 39")
        rD = _table_row(manual, keys, f"{fy}_7D", "As per Rule 42")
        rE = _table_row(manual, keys, f"{fy}_7E", "As per Rule 43")
        rF = _table_row(manual, keys, f"{fy}_7F", "As per Section 17(5)")
        rG = _table_row(manual, keys, f"{fy}_7G", "Reversal of TRAN-I credit")
        rH1 = _table_row(manual, keys, f"{fy}_7H1", "Reversal of TRAN-II credit")
        rI_other = _table_row(manual, keys, f"{fy}_7I_other", "Other reversals (specify in your notes)")

        rows = [rA, rA1, rB, rC, rD, rE, rF, rG, rH1, rI_other]
        total_reversed = {c: sum(r[c] for r in rows) for c in COLS}
        _table_row(manual, [], f"{fy}_7total_display", "Total ITC Reversed", computed=total_reversed)

        O6 = g9.get("6O_total_itc_availed", {c: 0.0 for c in COLS})
        net_available = {c: O6[c] - total_reversed[c] for c in COLS}
        _table_row(manual, [], f"{fy}_7net_display", "Net ITC Available for Utilization (6O – Total Reversed)", computed=net_available)

        g9["7_total_reversed"] = total_reversed
        g9["net_itc_available"] = net_available
        _save_button(db_path, fy, keys)

    # ---------------- TABLE 8 ----------------
    with tabs[5]:
        st.subheader("Table 8 — ITC as per GSTR-9 vs ITC as per GSTR-2B")
        itc_2b = {
            "Taxable Value": _sum(data["gstr2b"], "taxable_value"),
            "CGST": _sum(data["gstr2b"], "cgst"),
            "SGST": _sum(data["gstr2b"], "sgst"),
            "IGST": _sum(data["gstr2b"], "igst"),
            "Cess": 0.0,
        }
        keys = []
        A = _table_row(manual, [], f"{fy}_8A_display", "A. ITC as per GSTR-2B", computed=itc_2b)
        B = g9.get("6B", {c: 0.0 for c in COLS})
        _table_row(manual, [], f"{fy}_8B_display", "B. ITC as per sum total of Table 6(B)", computed=B)
        C = _table_row(manual, keys, f"{fy}_8C",
                       "C. ITC on inward supplies received in this FY but availed in the next FY (up to specified period)")
        D = {c: A[c] - (B[c] + C[c]) for c in COLS}
        _table_row(manual, [], f"{fy}_8D_display", "D. Difference (A – (B + C))", computed=D)
        st.caption(
            "A large, unexplained figure in row D usually means: ITC claimed in books/3B that never "
            "showed up in GSTR-2B (vendor non-compliance), or vice-versa. Investigate before filing."
        )
        _save_button(db_path, fy, keys)

    # ---------------- TABLE 9 ----------------
    with tabs[6]:
        st.subheader("Table 9 — Tax paid as declared in returns filed during the year")
        st.caption(
            "Payable is auto-computed from your Sales register. Paid (cash) / Paid (ITC) come from your "
            "GSTR-3B summary, split proportionally across CGST/SGST/IGST as a starting suggestion — "
            "please correct if your actual 3B filings differ."
        )
        payable = {
            "IGST": _sum(data["sales"], "igst"),
            "CGST": _sum(data["sales"], "cgst"),
            "SGST": _sum(data["sales"], "sgst"),
            "Cess": 0.0,
        }
        total_payable = sum(payable.values()) or 1.0
        cash_total = _sum(data["gstr3b"], "tax_paid_cash")
        itc_total = _sum(data["gstr3b"], "itc_claimed")
        late_fee_total = _sum(data["gstr3b"], "late_fee")

        keys = []
        head_rows = {}
        for head in ["Integrated Tax", "Central Tax", "State/UT Tax", "Cess"]:
            tax_key = {"Integrated Tax": "IGST", "Central Tax": "CGST", "State/UT Tax": "SGST", "Cess": "Cess"}[head]
            share = payable[tax_key] / total_payable if total_payable else 0.0
            st.markdown(f"**{head}**")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Payable", f"₹{payable[tax_key]:,.0f}")
            k_cash = f"{fy}_9_{tax_key}_cash"
            k_itc = f"{fy}_9_{tax_key}_itc"
            default_cash = float(manual.get(k_cash, round(cash_total * share, 2)))
            default_itc = float(manual.get(k_itc, round(itc_total * share, 2)))
            paid_cash = c2.number_input("Paid (cash)", value=default_cash, key=k_cash, step=100.0, format="%.2f")
            paid_itc = c3.number_input("Paid (ITC)", value=default_itc, key=k_itc, step=100.0, format="%.2f")
            keys.extend([k_cash, k_itc])
            diff = payable[tax_key] - (paid_cash + paid_itc)
            c4.metric("Difference", f"₹{diff:,.0f}")
            head_rows[head] = {"payable": payable[tax_key], "cash": paid_cash, "itc": paid_itc, "diff": diff}

        st.markdown("**Interest / Late Fee / Penalty / Other**")
        c1, c2 = st.columns(2)
        c1.metric("Late Fee paid (from GSTR-3B)", f"₹{late_fee_total:,.0f}")
        k_interest = f"{fy}_9_interest"
        interest_val = c2.number_input("Interest paid", value=float(manual.get(k_interest, 0.0)), key=k_interest, step=100.0, format="%.2f")
        keys.append(k_interest)

        g9["9_head_rows"] = head_rows
        _save_button(db_path, fy, keys)

    # ---------------- GSTR-9C: TURNOVER RECONCILIATION ----------------
    with tabs[7]:
        st.subheader("GSTR-9C — Reconciliation of Turnover (Audited FS vs Annual Return)")
        st.caption("These figures come from your audited Financial Statements / books of account and must be entered manually — they live outside this app by definition.")
        keys = []
        fs_turnover = st.number_input("Turnover as per Audited Financial Statements",
                                       value=float(manual.get(f"{fy}_9c_fs_turnover", 0.0)),
                                       key=f"{fy}_9c_fs_turnover", step=10000.0, format="%.2f")
        keys.append(f"{fy}_9c_fs_turnover")

        adj_labels = [
            ("composition_turnover", "Less: Turnover under composition scheme"),
            ("sec15_adj", "Adjustments in turnover under Section 15 & rules (+/-)"),
            ("unbilled_open", "Add: Unbilled revenue at the beginning of the FY"),
            ("unbilled_close", "Less: Unbilled revenue at the end of the FY"),
            ("adv_open", "Less: Unadjusted advances as at beginning of the FY"),
            ("adv_close", "Add: Unadjusted advances as at end of the FY"),
            ("deemed_supply", "Add: Deemed Supply under Schedule I"),
            ("cn_post_fy", "Less: Credit notes issued after FY-end but reflected in Annual Return"),
            ("cn_not_gst", "Add: Credit notes in audited FS not permissible under GST"),
            ("trade_discount", "Add: Trade discounts in audited FS not permissible under GST"),
            ("forex_adj", "Adjustments due to foreign exchange fluctuation (+/-)"),
            ("sez_dta_adj", "Adjustments — supply of goods by SEZ units to DTA units (+/-)"),
            ("other_adj", "Any other adjustment not listed above (+/-)"),
        ]
        adj_values = {}
        for k, label in adj_labels:
            key = f"{fy}_9c_{k}"
            val = st.number_input(label, value=float(manual.get(key, 0.0)), key=key, step=1000.0, format="%.2f")
            adj_values[k] = val
            keys.append(key)

        annual_turnover_after_adj = (
            fs_turnover
            - adj_values["composition_turnover"]
            + adj_values["sec15_adj"]
            + adj_values["unbilled_open"]
            - adj_values["unbilled_close"]
            - adj_values["adv_open"]
            + adj_values["adv_close"]
            + adj_values["deemed_supply"]
            - adj_values["cn_post_fy"]
            + adj_values["cn_not_gst"]
            + adj_values["trade_discount"]
            + adj_values["forex_adj"]
            + adj_values["sez_dta_adj"]
            + adj_values["other_adj"]
        )

        turnover_per_gstr9 = sum(g9.get("5N_total_turnover", {c: 0.0 for c in COLS}).values())

        c1, c2, c3 = st.columns(3)
        c1.metric("Annual Turnover after adjustments", f"₹{annual_turnover_after_adj:,.0f}")
        c2.metric("Turnover as declared in GSTR-9 (auto)", f"₹{turnover_per_gstr9:,.0f}")
        c3.metric("Unreconciled difference", f"₹{(turnover_per_gstr9 - annual_turnover_after_adj):,.0f}")

        if abs(turnover_per_gstr9 - annual_turnover_after_adj) > 1:
            reason_key = f"{fy}_9c_turnover_reason"
            st.text_area("Reason for un-reconciled turnover difference (required if non-zero)",
                          value=manual.get(reason_key, ""), key=reason_key)
            keys.append(reason_key)

        g9["turnover_per_gstr9"] = turnover_per_gstr9
        g9["annual_turnover_after_adj"] = annual_turnover_after_adj
        _save_button(db_path, fy, keys)

    # ---------------- GSTR-9C: TAX PAID RECONCILIATION ----------------
    with tabs[8]:
        st.subheader("GSTR-9C — Rate-wise Liability & Tax Paid Reconciliation")
        if not data["sales"].empty and "gst_rate" in data["sales"].columns:
            by_rate = data["sales"].groupby("gst_rate").agg(
                taxable_value=("taxable_value", "sum"),
                cgst=("cgst", "sum"),
                sgst=("sgst", "sum"),
                igst=("igst", "sum"),
            ).reset_index().rename(columns={"gst_rate": "Rate (%)"})
            by_rate["Total Tax"] = by_rate["cgst"] + by_rate["sgst"] + by_rate["igst"]
            st.caption("Auto-computed from your Sales register, grouped by GST rate:")
            st.dataframe(by_rate, use_container_width=True, hide_index=True)
            total_liability_books = float(by_rate["Total Tax"].sum())
        else:
            st.info("No sales data for this Financial Year yet.")
            total_liability_books = 0.0

        total_payable_gstr9 = sum(g9.get("4N", {c: 0.0 for c in COLS}).values())
        c1, c2, c3 = st.columns(3)
        c1.metric("Total tax liability (rate-wise, books)", f"₹{total_liability_books:,.0f}")
        c2.metric("Total tax payable per GSTR-9 (auto)", f"₹{total_payable_gstr9:,.0f}")
        c3.metric("Unreconciled difference", f"₹{(total_payable_gstr9 - total_liability_books):,.0f}")

        keys = []
        if abs(total_payable_gstr9 - total_liability_books) > 1:
            reason_key = f"{fy}_9c_tax_reason"
            st.text_area("Reason for un-reconciled payment of tax (required if non-zero)",
                          value=manual.get(reason_key, ""), key=reason_key)
            keys.append(reason_key)
        _save_button(db_path, fy, keys)

    # ---------------- GSTR-9C: ITC RECONCILIATION ----------------
    with tabs[9]:
        st.subheader("GSTR-9C — Reconciliation of Input Tax Credit (ITC)")
        keys = []
        books_itc = st.number_input(
            "ITC availed as per audited Financial Statements / books of account",
            value=float(manual.get(f"{fy}_9c_books_itc", 0.0)), key=f"{fy}_9c_books_itc",
            step=1000.0, format="%.2f"
        )
        keys.append(f"{fy}_9c_books_itc")

        booked_prior_fy = st.number_input(
            "ITC booked in earlier Financial Years, claimed in this FY (add)",
            value=float(manual.get(f"{fy}_9c_booked_prior", 0.0)), key=f"{fy}_9c_booked_prior",
            step=1000.0, format="%.2f"
        )
        keys.append(f"{fy}_9c_booked_prior")

        booked_defer_next_fy = st.number_input(
            "ITC booked in this FY, to be claimed in a subsequent FY (subtract)",
            value=float(manual.get(f"{fy}_9c_booked_defer", 0.0)), key=f"{fy}_9c_booked_defer",
            step=1000.0, format="%.2f"
        )
        keys.append(f"{fy}_9c_booked_defer")

        itc_per_books_adjusted = books_itc + booked_prior_fy - booked_defer_next_fy
        itc_per_gstr9 = sum(g9.get("6O_total_itc_availed", {c: 0.0 for c in COLS}).values())

        c1, c2, c3 = st.columns(3)
        c1.metric("ITC per books (adjusted)", f"₹{itc_per_books_adjusted:,.0f}")
        c2.metric("ITC claimed in GSTR-9 (auto)", f"₹{itc_per_gstr9:,.0f}")
        c3.metric("Unreconciled ITC", f"₹{(itc_per_gstr9 - itc_per_books_adjusted):,.0f}")

        if abs(itc_per_gstr9 - itc_per_books_adjusted) > 1:
            reason_key = f"{fy}_9c_itc_reason"
            st.text_area("Reason for un-reconciled ITC difference (required if non-zero)",
                          value=manual.get(reason_key, ""), key=reason_key)
            keys.append(reason_key)

        st.divider()
        st.markdown("**Expense-head-wise ITC (optional detail, for your working papers)**")
        st.caption("The official form asks you to reconcile ITC claimed against ITC embedded in each "
                    "major expense head from your P&L. Your app doesn't tag purchases by expense head, "
                    "so fill in whatever you have on hand — this section is informational only.")
        expense_heads = [
            "Purchases", "Freight / Carriage", "Power and Fuel", "Rent and Insurance",
            "Employee Cost", "Conveyance", "Repairs & Maintenance", "Capital Goods",
            "Other Expenses",
        ]
        expense_total = 0.0
        for head in expense_heads:
            key = f"{fy}_9c_exp_{head.replace(' ', '_').replace('/', '_')}"
            val = st.number_input(f"ITC embedded in: {head}", value=float(manual.get(key, 0.0)), key=key, step=500.0, format="%.2f")
            expense_total += val
            keys.append(key)
        st.metric("Total (expense-head-wise)", f"₹{expense_total:,.0f}")

        g9["itc_per_books_adjusted"] = itc_per_books_adjusted
        g9["itc_per_gstr9"] = itc_per_gstr9
        _save_button(db_path, fy, keys)

    # ---------------- EXPORT ----------------
    with tabs[10]:
        st.subheader("Export GSTR-9 / GSTR-9C Working Papers to Excel")
        st.caption("Generates a two-sheet workbook (GSTR-9 and GSTR-9C) with every figure computed above — "
                    "use this as your working paper to fill the official portal offline utility.")
        if st.button("📊 Generate Excel", type="primary", use_container_width=True):
            path = _export_excel(profile, fy, data, manual, g9)
            with open(path, "rb") as f:
                file_bytes = f.read()
            st.success("Workbook generated.")
            st.download_button(
                "⬇️ Download GSTR-9 & 9C Working Papers",
                data=file_bytes,
                file_name=os.path.basename(path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# ============================================================
# EXCEL EXPORT
# ============================================================

def _style_header(ws, row, ncols, text, fill=HEAD_FILL, font=HEAD_FONT):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    cell = ws.cell(row=row, column=1, value=text)
    cell.fill = fill
    cell.font = font
    cell.alignment = CENTER
    for c in range(1, ncols + 1):
        ws.cell(row=row, column=c).fill = fill
        ws.cell(row=row, column=c).border = BORDER


def _write_row(ws, row, label, values, bold=False):
    ws.cell(row=row, column=1, value=label).alignment = LEFT_WRAP
    for i, v in enumerate(values):
        cell = ws.cell(row=row, column=2 + i, value=round(float(v), 2))
        cell.border = BORDER
        if bold:
            cell.font = BOLD
    ws.cell(row=row, column=1).border = BORDER
    if bold:
        ws.cell(row=row, column=1).font = BOLD
        for c in range(1, 2 + len(values)):
            ws.cell(row=row, column=c).fill = TOTAL_FILL


def _export_excel(profile, fy, data, manual, g9):
    wb = Workbook()

    # ---------------- GSTR-9 sheet ----------------
    ws = wb.active
    ws.title = "GSTR-9"
    ws.column_dimensions["A"].width = 55
    for col in "BCDEF":
        ws.column_dimensions[col].width = 16

    r = 1
    _style_header(ws, r, 6, f"GSTR-9 — Annual Return — FY {fy}"); r += 1
    ws.cell(row=r, column=1, value=f"Legal Name: {profile.get('company_name','')}    GSTIN: {profile.get('gstin','')}")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
    r += 2

    headers = ["Description"] + list(COLS)
    for i, h in enumerate(headers):
        c = ws.cell(row=r, column=1 + i, value=h)
        c.font = BOLD
        c.fill = SUB_FILL
        c.border = BORDER
    r += 1

    def rowdict(d):
        return [d.get(c, 0.0) for c in COLS]

    _style_header(ws, r, 6, "Table 4 — Outward Supplies (Tax Payable)"); r += 1
    _write_row(ws, r, "A. Supplies to unregistered persons (B2C)", rowdict(g9.get("4A_B2C", {}))); r += 1
    _write_row(ws, r, "B. Supplies to registered persons (B2B)", rowdict(g9.get("4B_B2B", {}))); r += 1
    _write_row(ws, r, "N. Supplies & advances on which tax is payable", rowdict(g9.get("4N", {})), bold=True); r += 2

    _style_header(ws, r, 6, "Table 5 — Non-Taxable Outward Supplies"); r += 1
    _write_row(ws, r, "N. Total Turnover (4N + 5M − 4G − 4G1)", rowdict(g9.get("5N_total_turnover", {})), bold=True); r += 2

    _style_header(ws, r, 6, "Table 6 — ITC Availed"); r += 1
    _write_row(ws, r, "B. Inward supplies (other than import/RCM) — Purchase register", rowdict(g9.get("6B", {}))); r += 1
    _write_row(ws, r, "O. Total ITC Availed (I + N)", rowdict(g9.get("6O_total_itc_availed", {})), bold=True); r += 2

    _style_header(ws, r, 6, "Table 7 — ITC Reversed / Net ITC Available"); r += 1
    _write_row(ws, r, "Total ITC Reversed", rowdict(g9.get("7_total_reversed", {}))); r += 1
    _write_row(ws, r, "Net ITC Available for Utilization", rowdict(g9.get("net_itc_available", {})), bold=True); r += 2

    _style_header(ws, r, 6, "Table 9 — Tax Paid (by head)"); r += 1
    ws.cell(row=r, column=1, value="Head").font = BOLD
    ws.cell(row=r, column=2, value="Payable").font = BOLD
    ws.cell(row=r, column=3, value="Paid (Cash)").font = BOLD
    ws.cell(row=r, column=4, value="Paid (ITC)").font = BOLD
    ws.cell(row=r, column=5, value="Difference").font = BOLD
    r += 1
    for head, vals in g9.get("9_head_rows", {}).items():
        ws.cell(row=r, column=1, value=head)
        ws.cell(row=r, column=2, value=round(vals["payable"], 2))
        ws.cell(row=r, column=3, value=round(vals["cash"], 2))
        ws.cell(row=r, column=4, value=round(vals["itc"], 2))
        ws.cell(row=r, column=5, value=round(vals["diff"], 2))
        for c in range(1, 6):
            ws.cell(row=r, column=c).border = BORDER
        r += 1

    # ---------------- GSTR-9C sheet ----------------
    ws2 = wb.create_sheet("GSTR-9C")
    ws2.column_dimensions["A"].width = 60
    ws2.column_dimensions["B"].width = 20

    r = 1
    _style_header(ws2, r, 2, f"GSTR-9C — Reconciliation Statement — FY {fy}"); r += 1
    ws2.cell(row=r, column=1, value=f"Legal Name: {profile.get('company_name','')}    GSTIN: {profile.get('gstin','')}")
    ws2.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
    r += 2

    _style_header(ws2, r, 2, "Turnover Reconciliation"); r += 1
    for label, key in [
        ("Turnover as per Audited FS", "annual_turnover_after_adj"),
        ("Turnover as per GSTR-9 (auto)", "turnover_per_gstr9"),
    ]:
        ws2.cell(row=r, column=1, value=label).border = BORDER
        ws2.cell(row=r, column=2, value=round(float(g9.get(key, 0.0)), 2)).border = BORDER
        r += 1
    diff = float(g9.get("turnover_per_gstr9", 0.0)) - float(g9.get("annual_turnover_after_adj", 0.0))
    _write_row(ws2, r, "Unreconciled Turnover", [diff], bold=True); r += 2

    _style_header(ws2, r, 2, "ITC Reconciliation"); r += 1
    for label, key in [
        ("ITC per books (adjusted)", "itc_per_books_adjusted"),
        ("ITC claimed in GSTR-9 (auto)", "itc_per_gstr9"),
    ]:
        ws2.cell(row=r, column=1, value=label).border = BORDER
        ws2.cell(row=r, column=2, value=round(float(g9.get(key, 0.0)), 2)).border = BORDER
        r += 1
    diff_itc = float(g9.get("itc_per_gstr9", 0.0)) - float(g9.get("itc_per_books_adjusted", 0.0))
    _write_row(ws2, r, "Unreconciled ITC", [diff_itc], bold=True); r += 2

    ws2.cell(row=r, column=1, value=(
        "Note: This working paper follows the standard GSTR-9 / GSTR-9C structure and is auto-populated "
        "from the data in your app. GSTR-9C requires certification/sign-off — please have your CA verify "
        "these figures against the official GST portal offline utility before filing."
    ))
    ws2.cell(row=r, column=1).alignment = LEFT_WRAP
    ws2.merge_cells(start_row=r, start_column=1, end_row=r + 3, end_column=2)

    out_dir = "/mnt/user-data/outputs" if os.path.isdir("/mnt/user-data/outputs") else "."
    path = os.path.join(out_dir, f"GSTR9_9C_Working_Papers_{fy}.xlsx")
    wb.save(path)
    return path


# ============================================================
# STANDALONE TEST MODE
# ============================================================

if __name__ == "__main__":
    st.set_page_config(page_title="GSTR-9 / 9C", page_icon="📑", layout="wide")
    st.sidebar.markdown("### Standalone test mode")
    test_db = st.sidebar.text_input("Path to GST_Data.db", value="GST_Data.db")
    render(test_db if os.path.exists(test_db) else None)
