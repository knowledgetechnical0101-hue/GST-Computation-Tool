"""
GSTR-1 JSON Generator
======================
Streamlit app that converts a Sales Register (CSV/Excel) into a GST-portal
style B2B/HSN/Doc-Issue JSON, in the same shape as GSTN's "offline" export.

Run with:
    streamlit run gstr1_json_generator.py

IMPORTANT DISCLAIMER
---------------------
The "chksum" values below are locally-generated SHA-256 hashes used only to
give the JSON an internally consistent shape. GSTN's actual checksum
algorithm for the official Returns Offline Tool is not publicly documented,
and the GST portal recomputes/validates checksums itself on upload. Do NOT
treat the chksum fields in this app's output as authoritative -- always
upload through the official GST portal / Offline Tool and let it validate
the data. This tool is for data-preparation convenience only, not a
replacement for professional GST filing advice.
"""

import io
import json
import hashlib
import datetime as dt

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Reference data
# ---------------------------------------------------------------------------

STATE_CODES = {
    "01": "Jammu & Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
    "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam",
    "19": "West Bengal", "20": "Jharkhand", "21": "Odisha",
    "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "26": "Dadra & Nagar Haveli and Daman & Diu", "27": "Maharashtra",
    "28": "Andhra Pradesh (Old)", "29": "Karnataka", "30": "Goa",
    "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu",
    "34": "Puducherry", "35": "Andaman & Nicobar Islands", "36": "Telangana",
    "37": "Andhra Pradesh", "38": "Ladakh", "97": "Other Territory",
}
STATE_NAME_TO_CODE = {v: k for k, v in STATE_CODES.items()}

INVOICE_TYPES = {
    "Regular B2B": "R",
    "SEZ supplies with payment": "SEWP",
    "SEZ supplies without payment": "SEWOP",
    "Deemed Exports": "DE",
}

DOC_SERIES_LABELS = [
    "1 - Invoices for outward supply",
    "2 - Invoices for inward supply from unregistered person",
    "3 - Revised Invoice",
    "4 - Debit Note",
    "5 - Credit Note",
    "6 - Receipt voucher",
    "7 - Payment Voucher",
    "8 - Refund voucher",
    "9 - Delivery Challan for job work",
    "10 - Delivery Challan for supply on approval",
    "11 - Delivery Challan in case of liquid gas",
    "12 - Delivery Challan in cases other than by way of supply",
]

REQUIRED_COLUMNS = [
    "Invoice Number", "Invoice Date", "Customer GSTIN", "Place of Supply",
    "Taxable Value", "GST Rate", "HSN Code", "HSN Description",
]
OPTIONAL_COLUMNS = [
    "Reverse Charge", "Invoice Type", "UQC", "Quantity",
]

SAMPLE_ROW = {
    "Invoice Number": "LSC-029",
    "Invoice Date": "06-06-2026",
    "Customer GSTIN": "29AATFG3032H2Z6",
    "Place of Supply": "Karnataka",
    "Taxable Value": 160215,
    "GST Rate": 18,
    "HSN Code": "995428",
    "HSN Description": "General construction services of other civil engineering works nowhere else classified",
    "Reverse Charge": "N",
    "Invoice Type": "Regular B2B",
    "UQC": "NA",
    "Quantity": 0,
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def make_checksum(payload) -> str:
    """Deterministic local hash (NOT the official GSTN checksum -- see disclaimer)."""
    raw = json.dumps(payload, sort_keys=True, default=str).encode("utf-8")
    return hashlib.sha256(raw).hexdigest() + hashlib.sha256(raw).hexdigest()[:1]


def parse_date(value) -> str:
    """Return DD-MM-YYYY string from whatever pandas gives us."""
    if isinstance(value, str):
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
            try:
                return dt.datetime.strptime(value.strip(), fmt).strftime("%d-%m-%Y")
            except ValueError:
                continue
        # last resort: let pandas guess
        return pd.to_datetime(value, dayfirst=True).strftime("%d-%m-%Y")
    return pd.to_datetime(value).strftime("%d-%m-%Y")


def state_to_code(pos_value: str) -> str:
    pos_value = str(pos_value).strip()
    if pos_value in STATE_CODES:
        return pos_value
    if pos_value in STATE_NAME_TO_CODE:
        return STATE_NAME_TO_CODE[pos_value]
    # try zero-padded numeric
    if pos_value.isdigit():
        return pos_value.zfill(2)
    raise ValueError(f"Could not resolve Place of Supply '{pos_value}' to a state code")


def build_b2b(df: pd.DataFrame, home_state_code: str) -> list:
    """Group rows -> customers (ctin) -> invoices -> rate-wise items."""
    b2b = []
    for ctin, cust_df in df.groupby("Customer GSTIN"):
        invoices = []
        for inum, inv_df in cust_df.groupby("Invoice Number", sort=False):
            first = inv_df.iloc[0]
            pos_code = state_to_code(first["Place of Supply"])
            rchrg = str(first.get("Reverse Charge", "N")).strip().upper()[:1] or "N"
            inv_typ = INVOICE_TYPES.get(str(first.get("Invoice Type", "Regular B2B")), "R")

            itms = []
            total_val = 0.0
            for rt, rate_df in inv_df.groupby("GST Rate"):
                txval = round(float(rate_df["Taxable Value"].sum()), 2)
                rt = float(rt)
                tax_amt = round(txval * rt / 100.0, 2)
                item_det = {"rt": rt, "txval": txval}
                if pos_code == home_state_code:
                    item_det["camt"] = round(tax_amt / 2, 2)
                    item_det["samt"] = round(tax_amt / 2, 2)
                else:
                    item_det["iamt"] = tax_amt
                itms.append({"num": int(rt * 100), "itm_det": item_det})
                total_val += txval + tax_amt

            inv_record = {
                "val": round(total_val),
                "itms": itms,
                "inv_typ": inv_typ,
                "flag": "U",
                "updby": "S",
                "pos": pos_code,
                "idt": parse_date(first["Invoice Date"]),
                "rchrg": rchrg,
                "cflag": "N",
                "inum": str(inum),
            }
            inv_record["chksum"] = make_checksum(inv_record)
            invoices.append(inv_record)

        # keep invoices in a stable, date-then-number order
        invoices.sort(key=lambda r: (dt.datetime.strptime(r["idt"], "%d-%m-%Y"), r["inum"]))
        b2b.append({"ctin": str(ctin), "cfs": "N", "inv": invoices})

    return b2b


def build_hsn(df: pd.DataFrame) -> dict:
    hsn_rows = []
    grouped = df.groupby(["HSN Code", "GST Rate"], dropna=False)
    for i, ((hsn_code, rt), g) in enumerate(grouped, start=1):
        txval = round(float(g["Taxable Value"].sum()), 2)
        rt = float(rt)
        tax_amt = round(txval * rt / 100.0, 2)
        qty = float(g["Quantity"].sum()) if "Quantity" in g else 0
        uqc = str(g["UQC"].iloc[0]) if "UQC" in g and pd.notna(g["UQC"].iloc[0]) else "NA"
        desc = str(g["HSN Description"].iloc[0])
        row = {
            "rt": rt, "uqc": uqc, "qty": qty, "num": i,
            "txval": txval, "iamt": tax_amt, "hsn_sc": str(hsn_code), "desc": desc,
        }
        hsn_rows.append(row)

    hsn_section = {"flag": "N", "hsn_b2b": hsn_rows, "hsn_b2c": []}
    hsn_section["chksum"] = make_checksum(hsn_section)
    return hsn_section


def build_doc_issue(df: pd.DataFrame, cancelled_count: int = 0) -> dict:
    numbers = sorted(df["Invoice Number"].astype(str).unique().tolist())
    doc_det = []
    for doc_num in range(1, 13):
        if doc_num == 1 and numbers:
            docs = [{
                "cancel": cancelled_count,
                "num": 1,
                "totnum": len(numbers),
                "from": numbers[0],
                "to": numbers[-1],
                "net_issue": len(numbers) - cancelled_count,
            }]
        else:
            docs = []
        doc_det.append({"docs": docs, "doc_num": doc_num})

    doc_issue = {"flag": "N", "doc_det": doc_det}
    doc_issue["chksum"] = make_checksum(doc_issue)
    return doc_issue


def build_gstr1_json(df: pd.DataFrame, gstin: str, fp: str, filing_typ: str,
                      gt: float, cur_gt: float, home_state_code: str,
                      cancelled_count: int, include_fil_dt: bool) -> dict:
    payload = {
        "gstin": gstin,
        "fp": fp,
        "filing_typ": filing_typ,
        "gt": gt,
        "cur_gt": cur_gt,
        "b2b": build_b2b(df, home_state_code),
        "hsn": build_hsn(df),
        "doc_issue": build_doc_issue(df, cancelled_count),
    }
    if include_fil_dt:
        payload["fil_dt"] = dt.date.today().strftime("%d-%m-%Y")
    return payload


def validate_sales_register(df: pd.DataFrame) -> list:
    errors = []
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        errors.append(f"Missing required column(s): {', '.join(missing)}")
        return errors  # can't validate further without required columns

    if df["Customer GSTIN"].isna().any():
        errors.append("Some rows have a blank Customer GSTIN.")
    if df["Invoice Number"].isna().any():
        errors.append("Some rows have a blank Invoice Number.")
    if (df["Taxable Value"].fillna(0) < 0).any():
        errors.append("Some rows have a negative Taxable Value.")
    bad_gstins = df.loc[df["Customer GSTIN"].astype(str).str.len() != 15, "Customer GSTIN"].unique()
    if len(bad_gstins) > 0:
        errors.append(f"GSTIN(s) not 15 characters long: {', '.join(map(str, bad_gstins[:5]))}")
    return errors


# ---------------------------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------------------------

st.title("📄 GSTR-1 JSON Generator")
st.caption(
    "Upload your Sales Register and generate a B2B / HSN / Document-Issued "
    "JSON in the GST portal's offline-export shape."
)

# ---- Sidebar: app / filing configuration ---------------------------------
with st.sidebar:
    st.header("⚙️ Filing configuration")

    gstin = st.text_input("Your GSTIN", value="33IMUPS5792F1Z4", max_chars=15)
    legal_name = st.text_input("Legal Name (for your reference only)", value="")

    st.subheader("Return period")
    col1, col2 = st.columns(2)
    with col1:
        month = st.selectbox(
            "Month",
            list(range(1, 13)),
            index=dt.date.today().month - 1,
            format_func=lambda m: dt.date(2000, m, 1).strftime("%B"),
        )
    with col2:
        year = st.number_input("Year", min_value=2017, max_value=2100, value=dt.date.today().year, step=1)
    fp = f"{month:02d}{year}"

    filing_typ = st.selectbox("Filing type", ["M", "Q"], help="M = Monthly, Q = Quarterly")

    home_state_name = st.selectbox(
        "Your home state (registration state)",
        options=list(STATE_CODES.values()),
        index=list(STATE_CODES.values()).index("Karnataka") if "Karnataka" in STATE_CODES.values() else 0,
        help="Used to decide IGST (inter-state) vs CGST+SGST (intra-state) split.",
    )
    home_state_code = STATE_NAME_TO_CODE[home_state_name]

    st.subheader("Turnover")
    gt = st.number_input("Annual Gross Turnover (last FY) - gt", min_value=0.0, value=0.0, step=1000.0)
    cur_gt = st.number_input("Gross Turnover this FY so far - cur_gt", min_value=0.0, value=0.0, step=1000.0)

    cancelled_count = st.number_input(
        "Cancelled invoice count (for doc-issued summary)", min_value=0, value=0, step=1
    )
    include_fil_dt = st.checkbox("Include filing date (fil_dt) in output", value=False)

    st.divider()
    st.subheader("📥 Sales Register")
    uploaded_file = st.file_uploader(
        "Upload sales register (CSV or Excel)", type=["csv", "xlsx", "xls"]
    )

    with st.expander("Expected columns / template"):
        st.write("**Required:**", ", ".join(REQUIRED_COLUMNS))
        st.write("**Optional:**", ", ".join(OPTIONAL_COLUMNS))
        sample_df = pd.DataFrame([SAMPLE_ROW])
        st.dataframe(sample_df, use_container_width=True)
        csv_bytes = sample_df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "Download blank template (CSV)",
            data=csv_bytes,
            file_name="sales_register_template.csv",
            mime="text/csv",
        )

# ---------------------------------------------------------------------------
# Main panel
# ---------------------------------------------------------------------------

if uploaded_file is None:
    st.info("👈 Upload your Sales Register from the sidebar to get started.")
    st.stop()

# Read file
try:
    if uploaded_file.name.lower().endswith(".csv"):
        sales_df = pd.read_csv(uploaded_file)
    else:
        sales_df = pd.read_excel(uploaded_file)
except Exception as e:
    st.error(f"Could not read the uploaded file: {e}")
    st.stop()

# Normalize column names (strip whitespace)
sales_df.columns = [str(c).strip() for c in sales_df.columns]

# Fill optional columns with defaults if missing
if "Reverse Charge" not in sales_df.columns:
    sales_df["Reverse Charge"] = "N"
if "Invoice Type" not in sales_df.columns:
    sales_df["Invoice Type"] = "Regular B2B"
if "UQC" not in sales_df.columns:
    sales_df["UQC"] = "NA"
if "Quantity" not in sales_df.columns:
    sales_df["Quantity"] = 0

st.subheader("1️⃣ Sales Register preview")
st.dataframe(sales_df, use_container_width=True, height=280)

errors = validate_sales_register(sales_df)
if errors:
    st.error("Please fix the following before generating the JSON:")
    for e in errors:
        st.markdown(f"- {e}")
    st.stop()
else:
    st.success(f"Loaded {len(sales_df)} line item(s) across "
               f"{sales_df['Invoice Number'].nunique()} invoice(s) and "
               f"{sales_df['Customer GSTIN'].nunique()} customer(s).")

st.subheader("2️⃣ Generate GSTR-1 style JSON")

if st.button("🚀 Generate JSON", type="primary"):
    try:
        result = build_gstr1_json(
            sales_df, gstin=gstin, fp=fp, filing_typ=filing_typ,
            gt=gt, cur_gt=cur_gt, home_state_code=home_state_code,
            cancelled_count=cancelled_count, include_fil_dt=include_fil_dt,
        )
    except Exception as e:
        st.error(f"Failed to build JSON: {e}")
        st.stop()

    st.session_state["gstr1_json"] = result

if "gstr1_json" in st.session_state:
    result = st.session_state["gstr1_json"]

    total_invoices = sum(len(c["inv"]) for c in result["b2b"])
    total_taxable = sum(
        itm["itm_det"]["txval"]
        for cust in result["b2b"] for inv in cust["inv"] for itm in inv["itms"]
    )
    m1, m2, m3 = st.columns(3)
    m1.metric("Customers (ctin)", len(result["b2b"]))
    m2.metric("Invoices", total_invoices)
    m3.metric("Total Taxable Value", f"₹{total_taxable:,.2f}")

    st.json(result, expanded=False)

    json_str = json.dumps(result, indent=2)
    filename = f"returns_{dt.date.today().strftime('%d%m%Y')}_R1_{gstin}_offline_others_0.json"
    st.download_button(
        "⬇️ Download JSON",
        data=json_str.encode("utf-8"),
        file_name=filename,
        mime="application/json",
    )

    st.warning(
        "⚠️ The `chksum` values are locally generated for internal consistency only "
        "and are **not** the official GSTN checksum algorithm. Always upload this "
        "data through the official GST Returns Offline Tool / portal, which "
        "validates and recomputes checksums itself before filing."
    )
