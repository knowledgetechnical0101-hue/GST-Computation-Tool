"""
GSTR-1 JSON Generator — Streamlit module
-----------------------------------------
Lets a user enter B2B / B2CL / B2CS / CDNR / HSN-summary / Documents-Issued
data on screen and converts it into the JSON structure the GST portal's
"Prepare Offline -> Upload" screen expects for GSTR-1 (Form GSTR-1, Tables
4, 5, 7, 9B, 12, 13).

IMPORTANT — read before you wire this into a "live" workflow:
GSTN does not offer a public API that lets an arbitrary script log in and
push a return. The only two ways data reaches the portal are:
  1. Manual upload of a JSON file via the portal's own
     Returns Dashboard -> Prepare Offline -> Upload screen (this is what
     the official "GST Offline Tool" also does — it generates JSON, it
     does not auto-submit it either).
  2. API access through a government-licensed GSP (GST Suvidha Provider),
     which requires GSTN empanelment, a client_id/client_secret, and
     OTP/EVC or DSC based session auth — not something this file can do
     on its own.
This tool therefore does step 1 for you: correct data in, correct JSON
out, ready to upload by hand. Always re-validate the generated file with
the official Offline Tool / portal before filing — this is a best-effort
reverse-engineered schema (GSTN has not published an official one).

Drop this file into your Streamlit app's `pages/` folder (e.g.
`pages/5_GSTR1_JSON_Generator.py`) for a multipage app, or call
`render()` from inside an existing tab in your `app.py`.
"""

import json
import re
import sqlite3
from datetime import date

import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# Reference data
# --------------------------------------------------------------------------

STATE_CODES = {
    "01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
    "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam",
    "19": "West Bengal", "20": "Jharkhand", "21": "Odisha",
    "22": "Chattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "26": "Dadra and Nagar Haveli and Daman and Diu", "27": "Maharashtra",
    "28": "Andhra Pradesh (Old)", "29": "Karnataka", "30": "Goa",
    "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu",
    "34": "Puducherry", "35": "Andaman and Nicobar Islands",
    "36": "Telangana", "37": "Andhra Pradesh", "38": "Ladakh",
    "97": "Other Territory", "99": "Centre Jurisdiction",
}
STATE_OPTIONS = [f"{c} - {n}" for c, n in STATE_CODES.items()]

INVOICE_TYPES = ["R", "DE", "SEWP", "SEWOP", "CBW"]  # Regular, Deemed Export, SEZ w/pay, SEZ w/o pay, Bonded WH
NOTE_TYPES = {"Credit Note": "C", "Debit Note": "D", "Refund Voucher": "R"}
GSTIN_RE = re.compile(r"^\d{2}[A-Z]{5}\d{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$")

B2B_COLS = ["ctin", "inum", "idt", "val", "pos", "rchrg", "inv_typ", "rt", "txval", "csamt"]
B2CL_COLS = ["inum", "idt", "val", "pos", "rt", "txval", "csamt"]
B2CS_COLS = ["sply_ty", "pos", "rt", "txval", "csamt", "typ"]
CDNR_COLS = ["ctin", "nt_num", "nt_dt", "ntty", "pos", "rchrg", "val", "rt", "txval", "csamt"]
HSN_COLS = ["hsn_sc", "desc", "uqc", "qty", "rt", "txval", "iamt", "camt", "samt", "csamt", "segment"]
DOC_COLS = ["nature", "from_num", "to_num", "total", "cancelled"]

DOC_NATURE_CODE = {
    "Invoices for outward supply": 1, "Invoices for inward supply from unregistered person": 2,
    "Revised Invoice": 3, "Debit Note": 4, "Credit Note": 5, "Receipt voucher": 6,
    "Payment Voucher": 7, "Refund voucher": 8, "Delivery Challan for job work": 9,
    "Delivery Challan for supply on approval": 10,
    "Delivery Challan in case of liquid gas": 11,
    "Delivery Challan in cases other than by way of supply (excluding at S no. 9 to 11)": 12,
}


def empty_df(cols):
    return pd.DataFrame({c: pd.Series(dtype="object") for c in cols})


def init_state():
    defaults = {
        "b2b_df": empty_df(B2B_COLS), "b2cl_df": empty_df(B2CL_COLS),
        "b2cs_df": empty_df(B2CS_COLS), "cdnr_df": empty_df(CDNR_COLS),
        "hsn_df": empty_df(HSN_COLS), "doc_df": empty_df(DOC_COLS),
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def editable_table(session_key, column_config, editor_key):
    """Render an editable table with a per-row Delete action."""
    data = st.session_state[session_key].copy()
    data.insert(0, "__delete__", False)
    edited = st.data_editor(
        data,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "__delete__": st.column_config.CheckboxColumn("Delete", default=False),
            **column_config,
        },
        key=editor_key,
    )
    selected = edited.index[edited["__delete__"]].tolist()
    result = edited.drop(columns=["__delete__"])
    if st.button("🗑️ Delete Selected Rows", key=f"delete_{session_key}"):
        if not selected:
            st.warning("Select at least one row using the Delete checkbox.")
        else:
            st.session_state[session_key] = result.drop(index=selected).reset_index(drop=True)
            st.success(f"Deleted {len(selected)} row(s).")
            st.rerun()
    st.caption("Modify any cell directly in the table. Use the Delete checkbox for row deletion.")
    return result


def num(v, default=0.0):
    try:
        if v is None or v == "":
            return default
        return float(v)
    except (TypeError, ValueError):
        return default


def split_tax(rt, txval, pos_code, home_state_code):
    """Given a rate and taxable value, split into IGST or CGST+SGST
    depending on whether the place of supply is the same state as the
    filer (intra-state) or a different one (inter-state)."""
    rt = num(rt)
    txval = num(txval)
    tax_amt = round(txval * rt / 100, 2)
    if pos_code == home_state_code:
        half = round(tax_amt / 2, 2)
        return {"iamt": 0, "camt": half, "samt": half}
    return {"iamt": tax_amt, "camt": 0, "samt": 0}


# --------------------------------------------------------------------------
# JSON builders — one per GSTR-1 table
# --------------------------------------------------------------------------

def build_b2b(df, home_state_code):
    out = {}
    for i, row in df.iterrows():
        ctin = str(row.get("ctin", "")).strip()
        inum = str(row.get("inum", "")).strip()
        if not ctin or not inum:
            continue
        pos = str(row.get("pos", ""))[:2].zfill(2)
        tax = split_tax(row.get("rt"), row.get("txval"), pos, home_state_code)
        item = {
            "num": i + 1,
            "itm_det": {
                "rt": num(row.get("rt")),
                "txval": num(row.get("txval")),
                "iamt": tax["iamt"], "camt": tax["camt"], "samt": tax["samt"],
                "csamt": num(row.get("csamt")),
            },
        }
        inv = {
            "inum": inum,
            "idt": row.get("idt", ""),
            "val": num(row.get("val")),
            "pos": pos,
            "rchrg": row.get("rchrg", "N") or "N",
            "inv_typ": row.get("inv_typ", "R") or "R",
            "itms": [item],
        }
        bucket = out.setdefault(ctin, {"ctin": ctin, "inv": []})
        # merge line items into the same invoice if inum repeats
        existing = next((x for x in bucket["inv"] if x["inum"] == inum), None)
        if existing:
            existing["itms"].append({**item, "num": len(existing["itms"]) + 1})
        else:
            bucket["inv"].append(inv)
    return list(out.values())


def build_b2cl(df):
    out = {}
    for i, row in df.iterrows():
        inum = str(row.get("inum", "")).strip()
        if not inum:
            continue
        pos = str(row.get("pos", ""))[:2].zfill(2)
        item = {
            "num": i + 1,
            "itm_det": {
                "rt": num(row.get("rt")),
                "txval": num(row.get("txval")),
                "iamt": round(num(row.get("txval")) * num(row.get("rt")) / 100, 2),
                "csamt": num(row.get("csamt")),
            },
        }
        inv = {
            "inum": inum, "idt": row.get("idt", ""), "val": num(row.get("val")),
            "pos": pos, "itms": [item],
        }
        bucket = out.setdefault(pos, {"pos": pos, "inv": []})
        bucket["inv"].append(inv)
    return list(out.values())


def build_b2cs(df):
    entries = []
    for _, row in df.iterrows():
        pos = str(row.get("pos", ""))[:2].zfill(2)
        if not pos or pos == "00":
            continue
        rt = num(row.get("rt"))
        txval = num(row.get("txval"))
        sply_ty = row.get("sply_ty", "INTRA") or "INTRA"
        tax = {"iamt": round(txval * rt / 100, 2), "camt": 0, "samt": 0} if sply_ty == "INTER" \
            else {"iamt": 0, "camt": round(txval * rt / 200, 2), "samt": round(txval * rt / 200, 2)}
        entries.append({
            "pos": pos, "rt": rt, "txval": txval, "sply_ty": sply_ty,
            **tax, "csamt": num(row.get("csamt")),
        })
    return entries


def build_cdnr(df, home_state_code):
    out = {}
    for i, row in df.iterrows():
        ctin = str(row.get("ctin", "")).strip()
        nt_num = str(row.get("nt_num", "")).strip()
        if not ctin or not nt_num:
            continue
        pos = str(row.get("pos", ""))[:2].zfill(2)
        tax = split_tax(row.get("rt"), row.get("txval"), pos, home_state_code)
        item = {
            "num": i + 1,
            "itm_det": {
                "rt": num(row.get("rt")), "txval": num(row.get("txval")),
                "iamt": tax["iamt"], "camt": tax["camt"], "samt": tax["samt"],
                "csamt": num(row.get("csamt")),
            },
        }
        note = {
            "nt_num": nt_num, "nt_dt": row.get("nt_dt", ""),
            "ntty": row.get("ntty", "C") or "C", "pos": pos,
            "rchrg": row.get("rchrg", "N") or "N", "inv_typ": "R",
            "val": num(row.get("val")), "itms": [item],
        }
        bucket = out.setdefault(ctin, {"ctin": ctin, "nt": []})
        bucket["nt"].append(note)
    return list(out.values())


def build_hsn(df):
    hsn_b2b, hsn_b2c = [], []
    for _, row in df.iterrows():
        code = str(row.get("hsn_sc", "")).strip()
        if not code:
            continue
        qty = 0 if code.startswith("99") else num(row.get("qty"))
        entry = {
            "hsn_sc": code, "desc": row.get("desc", ""),
            "uqc": row.get("uqc", "OTH") or "OTH", "qty": qty,
            "rt": num(row.get("rt")), "txval": num(row.get("txval")),
            "iamt": num(row.get("iamt")), "camt": num(row.get("camt")),
            "samt": num(row.get("samt")), "csamt": num(row.get("csamt")),
            "num": int(row.get("num") or 1) if str(row.get("num", "")).strip() else 1,
        }
        (hsn_b2c if str(row.get("segment", "B2B")).upper() == "B2C" else hsn_b2b).append(entry)
    result = {"hsn_b2b": hsn_b2b}
    if hsn_b2c:
        result["hsn_b2c"] = hsn_b2c
    return result


def build_doc_issue(df):
    docs = []
    for _, row in df.iterrows():
        nature = row.get("nature", "")
        if not nature:
            continue
        try:
            frm = int(re.sub(r"\D", "", str(row.get("from_num", "0"))) or 0)
            to = int(re.sub(r"\D", "", str(row.get("to_num", "0"))) or 0)
        except ValueError:
            frm, to = 0, 0
        total = int(num(row.get("total"), to - frm + 1 if to >= frm else 0))
        cancelled = int(num(row.get("cancelled"), 0))
        docs.append({
            "num": DOC_NATURE_CODE.get(nature, 1),
            "docs": [{
                "num": DOC_NATURE_CODE.get(nature, 1),
                "from": str(row.get("from_num", "")), "to": str(row.get("to_num", "")),
                "totnum": total, "cancel": cancelled, "net_issue": total - cancelled,
            }],
        })
    return {"doc_det": docs}


# --------------------------------------------------------------------------
# Streamlit UI
# --------------------------------------------------------------------------

def import_sales_register(db_path):
    if not db_path:
        return 0
    connection = sqlite3.connect(db_path)
    sales = pd.read_sql_query(
        "SELECT gstin, invoice_number, invoice_date, total_amount, supply_category, "
        "transaction_type, gst_rate, taxable_value, cgst, sgst, igst "
        "FROM gst_invoices ORDER BY invoice_date, id", connection
    )
    connection.close()
    if sales.empty:
        return 0

    b2b_rows = []
    b2cl_rows = []
    for _, row in sales.iterrows():
        category = str(row.get("supply_category", "B2B") or "B2B").upper()
        target = b2b_rows if category == "B2B" and len(str(row.get("gstin", "")).strip()) == 15 else b2cl_rows
        item = {
            "inum": str(row.get("invoice_number", "")),
            "idt": str(row.get("invoice_date", "")),
            "val": row.get("total_amount", 0),
            "pos": str(row.get("gstin", ""))[:2] or "29",
            "rt": row.get("gst_rate", 0),
            "txval": row.get("taxable_value", 0),
            "csamt": 0,
        }
        if target is b2b_rows:
            item["ctin"] = str(row.get("gstin", ""))
            item.update({"rchrg": "N", "inv_typ": "R"})
        target.append(item)
    st.session_state.b2b_df = pd.DataFrame(b2b_rows, columns=B2B_COLS) if b2b_rows else empty_df(B2B_COLS)
    st.session_state.b2cl_df = pd.DataFrame(b2cl_rows, columns=B2CL_COLS) if b2cl_rows else empty_df(B2CL_COLS)
    return len(sales)


def render(db_path=None):
    st.header("GSTR-1 — Data Entry to Portal JSON")
    st.caption(
        "Fill in the sections that apply to this return period, then generate a "
        "JSON file in the format the GST portal's Offline upload screen expects. "
        "You still upload the file yourself on the portal — see the note above "
        "for why that last step can't be automated."
    )
    init_state()

    if db_path:
        st.info("Load Sales Register data directly into B2B/B2CL rows before generating JSON.")
        if st.button("📥 Import Sales Register into JSON Generator", type="primary"):
            imported = import_sales_register(db_path)
            if imported:
                st.success(f"Imported {imported} sales invoice(s). Review the B2B/B2CL tabs before generating JSON.")
                st.rerun()
            else:
                st.warning("No sales invoices found in the active database.")

    with st.expander("Filer details (required for every table)", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            gstin = st.text_input("Your GSTIN", max_chars=15, placeholder="e.g. 29ABCDE1234F1Z5")
        with c2:
            period = st.date_input("Return period (any date in the month)", value=date.today())
        with c3:
            home_state = st.selectbox("Your home state (for CGST/SGST split)", STATE_OPTIONS,
                                       index=STATE_OPTIONS.index("29 - Karnataka") if "29 - Karnataka" in STATE_OPTIONS else 0)
        home_state_code = home_state.split(" - ")[0]
        fp = period.strftime("%m%Y")
        if gstin and not GSTIN_RE.match(gstin.strip().upper()):
            st.warning("That doesn't look like a valid 15-character GSTIN format.")

    tabs = st.tabs(["B2B (Table 4)", "B2CL (Table 5)", "B2CS (Table 7)",
                     "CDNR (Table 9B)", "HSN Summary (Table 12)",
                     "Documents Issued (Table 13)", "Preview & Download"])

    with tabs[0]:
        st.write("Invoices to GST-registered recipients.")
        st.session_state.b2b_df = editable_table(
            "b2b_df", {
                "ctin": st.column_config.TextColumn("Recipient GSTIN"),
                "inum": st.column_config.TextColumn("Invoice No."),
                "idt": st.column_config.TextColumn("Invoice Date (DD-MM-YYYY)"),
                "val": st.column_config.NumberColumn("Invoice Value"),
                "pos": st.column_config.SelectboxColumn("Place of Supply", options=list(STATE_CODES.keys())),
                "rchrg": st.column_config.SelectboxColumn("Reverse Charge", options=["N", "Y"]),
                "inv_typ": st.column_config.SelectboxColumn("Invoice Type", options=INVOICE_TYPES),
                "rt": st.column_config.NumberColumn("Rate %"),
                "txval": st.column_config.NumberColumn("Taxable Value"),
                "csamt": st.column_config.NumberColumn("Cess Amount"),
            }, "b2b_editor",
        )

    with tabs[1]:
        st.write("Inter-state invoices to unregistered persons above the B2CL threshold.")
        st.session_state.b2cl_df = editable_table(
            "b2cl_df", {
                "inum": st.column_config.TextColumn("Invoice No."),
                "idt": st.column_config.TextColumn("Invoice Date (DD-MM-YYYY)"),
                "val": st.column_config.NumberColumn("Invoice Value"),
                "pos": st.column_config.SelectboxColumn("Place of Supply", options=list(STATE_CODES.keys())),
                "rt": st.column_config.NumberColumn("Rate %"),
                "txval": st.column_config.NumberColumn("Taxable Value"),
                "csamt": st.column_config.NumberColumn("Cess Amount"),
            }, "b2cl_editor",
        )

    with tabs[2]:
        st.write("Everything else to unregistered persons — summarised by state and rate.")
        st.session_state.b2cs_df = editable_table(
            "b2cs_df", {
                "sply_ty": st.column_config.SelectboxColumn("Supply Type", options=["INTRA", "INTER"]),
                "pos": st.column_config.SelectboxColumn("Place of Supply", options=list(STATE_CODES.keys())),
                "rt": st.column_config.NumberColumn("Rate %"),
                "txval": st.column_config.NumberColumn("Taxable Value"),
                "csamt": st.column_config.NumberColumn("Cess Amount"),
                "typ": st.column_config.SelectboxColumn("Type", options=["OE", "E"]),
            }, "b2cs_editor",
        )

    with tabs[3]:
        st.write("Credit / debit notes issued to registered recipients.")
        st.session_state.cdnr_df = editable_table(
            "cdnr_df", {
                "ctin": st.column_config.TextColumn("Recipient GSTIN"),
                "nt_num": st.column_config.TextColumn("Note No."),
                "nt_dt": st.column_config.TextColumn("Note Date (DD-MM-YYYY)"),
                "ntty": st.column_config.SelectboxColumn("Note Type", options=list(NOTE_TYPES.values())),
                "pos": st.column_config.SelectboxColumn("Place of Supply", options=list(STATE_CODES.keys())),
                "rchrg": st.column_config.SelectboxColumn("Reverse Charge", options=["N", "Y"]),
                "val": st.column_config.NumberColumn("Note Value"),
                "rt": st.column_config.NumberColumn("Rate %"),
                "txval": st.column_config.NumberColumn("Taxable Value"),
                "csamt": st.column_config.NumberColumn("Cess Amount"),
            }, "cdnr_editor",
        )

    with tabs[4]:
        st.write("HSN/SAC-wise summary of all outward supplies (mandatory table).")
        st.session_state.hsn_df = editable_table(
            "hsn_df", {
                "hsn_sc": st.column_config.TextColumn("HSN/SAC Code"),
                "desc": st.column_config.TextColumn("Description"),
                "uqc": st.column_config.TextColumn("UQC"),
                "qty": st.column_config.NumberColumn("Total Quantity"),
                "rt": st.column_config.NumberColumn("Rate %"),
                "txval": st.column_config.NumberColumn("Taxable Value"),
                "iamt": st.column_config.NumberColumn("IGST"),
                "camt": st.column_config.NumberColumn("CGST"),
                "samt": st.column_config.NumberColumn("SGST"),
                "csamt": st.column_config.NumberColumn("Cess"),
                "segment": st.column_config.SelectboxColumn("Segment", options=["B2B", "B2C"]),
            }, "hsn_editor",
        )

    with tabs[5]:
        st.write("Summary of document series issued this period (mandatory table).")
        st.session_state.doc_df = editable_table(
            "doc_df", {
                "nature": st.column_config.SelectboxColumn("Nature of Document", options=list(DOC_NATURE_CODE.keys())),
                "from_num": st.column_config.TextColumn("Sr. No. From"),
                "to_num": st.column_config.TextColumn("Sr. No. To"),
                "total": st.column_config.NumberColumn("Total Number"),
                "cancelled": st.column_config.NumberColumn("Cancelled"),
            }, "doc_editor",
        )

    with tabs[6]:
        st.subheader("Preview")
        payload = {"gstin": gstin.strip().upper() if gstin else "", "fp": fp, "version": "GST3.2.4"}

        b2b = build_b2b(st.session_state.b2b_df, home_state_code)
        b2cl = build_b2cl(st.session_state.b2cl_df)
        b2cs = build_b2cs(st.session_state.b2cs_df)
        cdnr = build_cdnr(st.session_state.cdnr_df, home_state_code)
        hsn = build_hsn(st.session_state.hsn_df)
        doc_issue = build_doc_issue(st.session_state.doc_df)

        if b2b:
            payload["b2b"] = b2b
        if b2cl:
            payload["b2cl"] = b2cl
        if b2cs:
            payload["b2cs"] = b2cs
        if cdnr:
            payload["cdnr"] = cdnr
        if hsn.get("hsn_b2b") or hsn.get("hsn_b2c"):
            payload["hsn"] = hsn
        if doc_issue.get("doc_det"):
            payload["doc_issue"] = doc_issue

        errors = []
        if not gstin or not GSTIN_RE.match(gstin.strip().upper()):
            errors.append("A valid 15-character GSTIN is required.")
        if "hsn" not in payload:
            errors.append("HSN Summary (Table 12) is mandatory — add at least one row.")
        if "doc_issue" not in payload:
            errors.append("Documents Issued (Table 13) is mandatory — add at least one row.")

        for e in errors:
            st.error(e)

        st.json(payload, expanded=False)

        col1, col2 = st.columns(2)
        with col1:
            st.metric("B2B invoices", sum(len(c["inv"]) for c in b2b))
            st.metric("B2CL invoices", sum(len(p["inv"]) for p in b2cl))
        with col2:
            st.metric("CDNR notes", sum(len(c["nt"]) for c in cdnr))
            st.metric("HSN lines", len(hsn.get("hsn_b2b", [])) + len(hsn.get("hsn_b2c", [])))

        json_str = json.dumps(payload, indent=2, ensure_ascii=False)
        st.download_button(
            "Download GSTR-1 JSON",
            data=json_str,
            file_name=f"GSTR1_{payload['gstin'] or 'GSTIN'}_{fp}.json",
            mime="application/json",
            disabled=bool(errors),
            use_container_width=True,
        )
        st.info(
            "Next step: log into the GST portal → **Returns Dashboard → "
            "GSTR-1 → Prepare Offline → Upload** and select this file. "
            "The portal (or the official Offline Tool's 'Open' + validation) "
            "will flag any remaining schema issues before you can submit."
        )


if __name__ == "__main__":
    st.set_page_config(page_title="GSTR-1 JSON Generator", layout="wide")
    render()
