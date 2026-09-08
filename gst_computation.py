import pandas as pd
import streamlit as st
import os
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def load_computation_ledger(read_query):
    data = read_query("SELECT * FROM gst_computation_ledger WHERE id = 1")
    if data.empty:
        return {
            "opening_credit_ledger": 0.0,
            "closing_credit_ledger": 0.0,
            "opening_cash_ledger": 0.0,
            "closing_cash_ledger": 0.0,
        }
    return data.iloc[0].to_dict()


def save_computation_ledger(run_query, now_stamp, opening_credit, closing_credit, opening_cash, closing_cash):
    run_query("""
        INSERT INTO gst_computation_ledger
            (id, opening_credit_ledger, closing_credit_ledger, opening_cash_ledger, closing_cash_ledger, updated_on)
        VALUES (1, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            opening_credit_ledger=excluded.opening_credit_ledger,
            closing_credit_ledger=excluded.closing_credit_ledger,
            opening_cash_ledger=excluded.opening_cash_ledger,
            closing_cash_ledger=excluded.closing_cash_ledger,
            updated_on=excluded.updated_on
    """, (opening_credit, closing_credit, opening_cash, closing_cash, now_stamp()))


def load_computation_hsn(read_query):
    return read_query("""
        SELECT id, hsn_code, hsn_description, taxable_value, igst, cgst, sgst
        FROM gst_computation_hsn ORDER BY hsn_code
    """)


def save_computation_hsn(run_query, now_stamp, dataframe):
    run_query("DELETE FROM gst_computation_hsn")
    for _, row in dataframe.iterrows():
        hsn_code = str(row.get("hsn_code", "") or "").strip()
        if not hsn_code:
            continue
        run_query("""
            INSERT INTO gst_computation_hsn
                (hsn_code, hsn_description, taxable_value, igst, cgst, sgst, updated_on)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            hsn_code,
            str(row.get("hsn_description", "") or "").strip(),
            float(row.get("taxable_value", 0) or 0),
            float(row.get("igst", 0) or 0),
            float(row.get("cgst", 0) or 0),
            float(row.get("sgst", 0) or 0),
            now_stamp(),
        ))


def load_computation_working(read_query):
    data = read_query("SELECT * FROM gst_computation_working WHERE id = 1")
    if data.empty:
        return {}
    return data.iloc[0].to_dict()


def save_computation_working(run_query, now_stamp, values):
    run_query("""
        INSERT INTO gst_computation_working
            (id, sales_adjustment, purchase_adjustment, output_tax_adjustment,
             itc_adjustment, itc_reversal_adjustment, other_adjustment, updated_on)
        VALUES (1, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(id) DO UPDATE SET
            sales_adjustment=excluded.sales_adjustment,
            purchase_adjustment=excluded.purchase_adjustment,
            output_tax_adjustment=excluded.output_tax_adjustment,
            itc_adjustment=excluded.itc_adjustment,
            itc_reversal_adjustment=excluded.itc_reversal_adjustment,
            other_adjustment=excluded.other_adjustment,
            updated_on=excluded.updated_on
    """, (*values, now_stamp()))


def load_computation_worksheet(read_query):
    return read_query("SELECT particular, source, adjustment FROM gst_computation_worksheet ORDER BY id")


def save_computation_worksheet(run_query, now_stamp, dataframe):
    run_query("DELETE FROM gst_computation_worksheet")
    for _, row in dataframe.iterrows():
        particular = str(row.get("Particular", "") or "").strip()
        if not particular:
            continue
        run_query("""
            INSERT INTO gst_computation_worksheet (particular, source, adjustment, updated_on)
            VALUES (?, ?, ?, ?)
        """, (particular, str(row.get("Source", "") or "").strip(),
               float(row.get("Adjustment", 0) or 0), now_stamp()))


def create_computation_excel(report_path, period_text, worksheet, payable_rows, ledger, hsn_data, company_profile):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "GST Computation"
    border = Border(bottom=Side(style="thin", color="B7B7B7"))
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(bold=True, color="FFFFFF")
    sheet.merge_cells("A1:J1")
    sheet["A1"] = "GST PRE-FILING COMPUTATION WORKING"
    sheet["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    sheet["A1"].fill = header_fill
    sheet["A1"].alignment = Alignment(horizontal="center")
    sheet.merge_cells("A2:J2")
    company_name = str(company_profile.get("company_name", "") or "").strip()
    gstin = str(company_profile.get("gstin", "") or "").strip()
    address = str(company_profile.get("address", "") or "").strip()
    sheet["A2"] = f"Company: {company_name}    GSTIN: {gstin}"
    sheet["A2"].alignment = Alignment(horizontal="center")
    sheet.merge_cells("A3:J3")
    sheet["A3"] = f"Address: {address}    Computation Period: {period_text}"
    sheet["A3"].alignment = Alignment(horizontal="center", wrap_text=True)
    sheet["A4"] = "GST PAYABLE WORKING"
    sheet["A4"].font = Font(bold=True, color="FFFFFF")
    sheet["A4"].fill = header_fill
    headers = ["Particular", "Type", "Amount", "Calculation / Note"]
    for column, value in enumerate(headers, start=1):
        cell = sheet.cell(row=5, column=column, value=value)
        cell.font = header_font
        cell.fill = header_fill
        cell.border = border
    for row_number, row in enumerate(payable_rows, start=6):
        export_values = [
            row["Particular"],
            row["Source"],
            row["Amount"],
            f"Adjustment: {row.get('Adjustment', 0):,.2f}; Working: {row.get('Working', 0):,.2f}",
        ]
        for column, value in enumerate(export_values, start=1):
            cell = sheet.cell(row=row_number, column=column, value=value)
            cell.border = border
            if column == 3:
                cell.number_format = '#,##0.00'
    start = 6 + len(payable_rows) + 2
    sheet.cell(row=start, column=1, value="LEDGER BALANCES").font = Font(bold=True, color="FFFFFF")
    sheet.cell(row=start, column=1).fill = header_fill
    ledger_rows = [
        ("Opening Credit Ledger", ledger.get("opening_credit_ledger", 0)),
        ("Closing Credit Ledger", ledger.get("closing_credit_ledger", 0)),
        ("Opening Cash Ledger", ledger.get("opening_cash_ledger", 0)),
        ("Closing Cash Ledger", ledger.get("closing_cash_ledger", 0)),
    ]
    for row_number, (label, value) in enumerate(ledger_rows, start=start + 1):
        sheet.cell(row=row_number, column=1, value=label)
        sheet.cell(row=row_number, column=2, value=float(value or 0)).number_format = '#,##0.00'
    hsn_start = start + len(ledger_rows) + 3
    sheet.cell(row=hsn_start, column=1, value="HSN / SAC WORKING").font = Font(bold=True, color="FFFFFF")
    sheet.cell(row=hsn_start, column=1).fill = header_fill
    for column, value in enumerate(["HSN / SAC", "Description", "Taxable Value", "IGST", "CGST", "SGST"], start=1):
        cell = sheet.cell(row=hsn_start + 1, column=column, value=value)
        cell.font = header_font
        cell.fill = header_fill
    for row_number, (_, row) in enumerate(hsn_data.iterrows(), start=hsn_start + 2):
        values = [row.get("hsn_code", ""), row.get("hsn_description", ""), row.get("taxable_value", 0), row.get("igst", 0), row.get("cgst", 0), row.get("sgst", 0)]
        for column, value in enumerate(values, start=1):
            sheet.cell(row=row_number, column=column, value=value)
            if column >= 3:
                sheet.cell(row=row_number, column=column).number_format = '#,##0.00'
    for column, width in {"A": 32, "B": 18, "C": 18, "D": 42, "E": 16, "F": 16, "G": 16, "H": 16, "I": 16, "J": 16}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A6"
    workbook.save(report_path)


def render(db_path, load_sales, load_purchases, load_gstr1, load_gstr3b, read_query, run_query, now_stamp, load_company_profile):
    st.markdown('<div class="main-title">GST Computation Sheet</div>', unsafe_allow_html=True)
    st.caption("Sales, Purchase, GSTR-1 and GSTR-3B totals are read automatically from their registers.")

    period_cols = st.columns(2)
    default_from = date(datetime.now().year, 4, 1)
    from_date = period_cols[0].date_input("From Date", value=default_from, key="gst_computation_from_date")
    to_date = period_cols[1].date_input("To Date", value=date.today(), key="gst_computation_to_date")
    if from_date > to_date:
        st.error("From Date cannot be after To Date.")
        return

    period_start = from_date.strftime("%Y-%m")
    period_end = to_date.strftime("%Y-%m")
    sales = load_sales(from_date, to_date)
    purchases = load_purchases(from_date, to_date)
    gstr1 = load_gstr1()
    gstr3b = load_gstr3b()
    if not gstr1.empty and "return_period" in gstr1.columns:
        gstr1 = gstr1[gstr1["return_period"].astype(str).between(period_start, period_end)].copy()
    if not gstr3b.empty and "return_period" in gstr3b.columns:
        gstr3b = gstr3b[gstr3b["return_period"].astype(str).between(period_start, period_end)].copy()
    st.caption(f"Computation period: {from_date:%Y-%m-%d} to {to_date:%Y-%m-%d}")

    sales_taxable = float(sales["taxable_value"].sum()) if not sales.empty else 0.0
    sales_igst = float(sales["igst"].sum()) if not sales.empty else 0.0
    sales_cgst = float(sales["cgst"].sum()) if not sales.empty else 0.0
    sales_sgst = float(sales["sgst"].sum()) if not sales.empty else 0.0
    purchase_taxable = float(purchases["taxable_value"].sum()) if not purchases.empty else 0.0
    purchase_itc = float(purchases[["igst", "cgst", "sgst"]].sum().sum()) if not purchases.empty else 0.0
    gstr1_taxable = float(gstr1["total_taxable_value"].sum()) if not gstr1.empty else 0.0
    gstr1_tax = float(gstr1["total_tax"].sum()) if not gstr1.empty else 0.0
    gstr3b_itc = float(gstr3b["itc_claimed"].sum()) if not gstr3b.empty else 0.0
    eligible_purchases = purchases
    if not purchases.empty and "itc_eligible" in purchases.columns:
        eligible_purchases = purchases[purchases["itc_eligible"].astype(str).str.strip().str.lower().eq("yes")]
    eligible_purchase_itc = float(eligible_purchases[["igst", "cgst", "sgst"]].sum().sum()) if not eligible_purchases.empty else 0.0
    working = load_computation_working(read_query)

    st.subheader("GST Payable Working")
    with st.form("gst_computation_adjustment_form"):
        adjustment_cols = st.columns(6)
        sales_adjustment = adjustment_cols[0].number_input("Sales GST Adj.", value=float(working.get("sales_adjustment", 0) or 0), step=100.0)
        purchase_adjustment = adjustment_cols[1].number_input("Purchase ITC Adj.", value=float(working.get("purchase_adjustment", 0) or 0), step=100.0)
        output_tax_adjustment = adjustment_cols[2].number_input("Output Tax Adj.", value=float(working.get("output_tax_adjustment", 0) or 0), step=100.0)
        itc_adjustment = adjustment_cols[3].number_input("ITC Adj.", value=float(working.get("itc_adjustment", 0) or 0), step=100.0)
        itc_reversal_adjustment = adjustment_cols[4].number_input("ITC Reversal", value=float(working.get("itc_reversal_adjustment", 0) or 0), step=100.0)
        other_adjustment = adjustment_cols[5].number_input("Other Adj.", value=float(working.get("other_adjustment", 0) or 0), step=100.0)
        save_adjustments = st.form_submit_button("Save Working Adjustments", type="primary")
    if save_adjustments:
        save_computation_working(
            run_query, now_stamp, (sales_adjustment, purchase_adjustment, output_tax_adjustment,
                                   itc_adjustment, itc_reversal_adjustment, other_adjustment)
        )
        st.success("GST working adjustments saved.")

    output_tax = sales_igst + sales_cgst + sales_sgst + sales_adjustment + output_tax_adjustment
    net_itc = eligible_purchase_itc + purchase_adjustment + itc_adjustment - itc_reversal_adjustment
    net_payable = max(output_tax - net_itc + other_adjustment, 0.0)
    opening_credit = float(load_computation_ledger(read_query).get("opening_credit_ledger", 0) or 0)
    cash_required = max(net_payable - opening_credit, 0.0)
    closing_credit_working = max(opening_credit - net_payable, 0.0)
    payable_rows = [
        {"Particular": "Output GST from Sales Register", "Source": "Addition", "Amount": sales_igst + sales_cgst + sales_sgst, "Adjustment": sales_adjustment, "Working": sales_igst + sales_cgst + sales_sgst + sales_adjustment},
        {"Particular": "Output Tax Adjustment", "Source": "Addition", "Amount": output_tax_adjustment, "Adjustment": 0.0, "Working": output_tax_adjustment},
        {"Particular": "Eligible ITC from Purchase Register", "Source": "Less", "Amount": eligible_purchase_itc, "Adjustment": purchase_adjustment + itc_adjustment, "Working": -(eligible_purchase_itc + purchase_adjustment + itc_adjustment)},
        {"Particular": "ITC Reversal", "Source": "Less", "Amount": itc_reversal_adjustment, "Adjustment": 0.0, "Working": itc_reversal_adjustment},
        {"Particular": "Other Adjustment", "Source": "Addition", "Amount": other_adjustment, "Adjustment": 0.0, "Working": other_adjustment},
        {"Particular": "Net GST Payable before Filing", "Source": "Result", "Amount": net_payable, "Adjustment": 0.0, "Working": net_payable},
        {"Particular": "Less: Opening Credit Ledger", "Source": "Less", "Amount": opening_credit, "Adjustment": 0.0, "Working": -min(opening_credit, net_payable)},
        {"Particular": "Cash Tax Required", "Source": "Result", "Amount": cash_required, "Adjustment": 0.0, "Working": cash_required},
        {"Particular": "Closing Credit Ledger (Working)", "Source": "Result", "Amount": closing_credit_working, "Adjustment": 0.0, "Working": closing_credit_working},
    ]
    payable_table = pd.DataFrame(payable_rows)
    st.dataframe(payable_table, use_container_width=True, hide_index=True)

    saved_worksheet = load_computation_worksheet(read_query)
    saved_adjustments = {
        str(row.particular): float(row.adjustment or 0)
        for row in saved_worksheet.itertuples(index=False)
    }
    worksheet = pd.DataFrame([
        {"Particular": "Sales Register", "Source": "Sales Register", "Taxable Value": sales_taxable,
         "IGST": sales_igst, "CGST": sales_cgst, "SGST": sales_sgst,
         "GST Total": sales_igst + sales_cgst + sales_sgst, "ITC Claimed": 0.0, "Adjustment": 0.0},
        {"Particular": "Purchase Register", "Source": "Purchase Register", "Taxable Value": purchase_taxable,
         "IGST": float(purchases["igst"].sum()) if not purchases.empty else 0.0,
         "CGST": float(purchases["cgst"].sum()) if not purchases.empty else 0.0,
         "SGST": float(purchases["sgst"].sum()) if not purchases.empty else 0.0,
         "GST Total": eligible_purchase_itc, "ITC Claimed": eligible_purchase_itc, "Adjustment": 0.0},
        {"Particular": "GSTR-1 Summary", "Source": "GSTR-1", "Taxable Value": gstr1_taxable,
         "IGST": 0.0, "CGST": 0.0, "SGST": 0.0, "GST Total": gstr1_tax,
         "ITC Claimed": 0.0, "Adjustment": 0.0},
        {"Particular": "GSTR-3B ITC Claimed", "Source": "GSTR-3B", "Taxable Value": 0.0,
         "IGST": 0.0, "CGST": 0.0, "SGST": 0.0, "GST Total": 0.0,
         "ITC Claimed": gstr3b_itc, "Adjustment": 0.0},
    ])
    worksheet["Adjustment"] = worksheet["Particular"].map(saved_adjustments).fillna(worksheet["Adjustment"])
    worksheet["Working Total"] = worksheet["GST Total"] + worksheet["ITC Claimed"] + worksheet["Adjustment"]

    st.subheader("Computation Worksheet")
    st.caption("Values are refreshed from the Sales, Purchase, GSTR-1 and GSTR-3B registers. Enter only manual adjustments in the worksheet.")
    edited_worksheet = st.data_editor(
        worksheet,
        use_container_width=True,
        hide_index=True,
        key="gst_computation_worksheet",
        column_config={
            "Particular": st.column_config.TextColumn("Particular"),
            "Source": st.column_config.TextColumn("Source"),
            "Taxable Value": st.column_config.NumberColumn("Taxable Value", format="%.2f"),
            "IGST": st.column_config.NumberColumn("IGST", format="%.2f"),
            "CGST": st.column_config.NumberColumn("CGST", format="%.2f"),
            "SGST": st.column_config.NumberColumn("SGST", format="%.2f"),
            "GST Total": st.column_config.NumberColumn("GST Total", format="%.2f"),
            "ITC Claimed": st.column_config.NumberColumn("ITC Claimed", format="%.2f"),
            "Adjustment": st.column_config.NumberColumn("Manual Adjustment", format="%.2f"),
            "Working Total": st.column_config.NumberColumn("Working Total", format="%.2f"),
        },
        disabled=[column for column in worksheet.columns if column not in {"Adjustment"}],
    )
    edited_worksheet["Working Total"] = (
        pd.to_numeric(edited_worksheet["GST Total"], errors="coerce").fillna(0)
        + pd.to_numeric(edited_worksheet["ITC Claimed"], errors="coerce").fillna(0)
        + pd.to_numeric(edited_worksheet["Adjustment"], errors="coerce").fillna(0)
    )
    working_total = float(edited_worksheet["Working Total"].sum())
    st.info(f"Worksheet working total: INR {working_total:,.2f}")
    if st.button("Save GST Computation Details", type="primary", key="save_gst_computation_details"):
        save_computation_worksheet(run_query, now_stamp, edited_worksheet)
        st.success("GST computation details and adjustments saved.")

    st.subheader("Automatic Register Computation")
    totals = st.columns(6)
    totals[0].metric("Sales Taxable", f"INR {sales_taxable:,.2f}")
    totals[1].metric("Purchase Taxable", f"INR {purchase_taxable:,.2f}")
    totals[2].metric("Sales Output GST", f"INR {sales_igst + sales_cgst + sales_sgst:,.2f}")
    totals[3].metric("Purchase ITC", f"INR {purchase_itc:,.2f}")
    totals[4].metric("GSTR-1 Taxable", f"INR {gstr1_taxable:,.2f}")
    totals[5].metric("GSTR-3B ITC Claimed", f"INR {gstr3b_itc:,.2f}")

    ledger = load_computation_ledger(read_query)
    st.subheader("ITC and Cash Ledger Working")
    with st.form("gst_computation_ledger_form"):
        ledger_cols = st.columns(4)
        opening_credit = ledger_cols[0].number_input("Opening Credit Ledger", min_value=0.0, value=float(ledger.get("opening_credit_ledger", 0) or 0), step=100.0)
        closing_credit = ledger_cols[1].number_input("Closing Credit Ledger", min_value=0.0, value=float(ledger.get("closing_credit_ledger", 0) or 0), step=100.0)
        opening_cash = ledger_cols[2].number_input("Opening Cash Ledger", min_value=0.0, value=float(ledger.get("opening_cash_ledger", 0) or 0), step=100.0)
        closing_cash = ledger_cols[3].number_input("Closing Cash Ledger", min_value=0.0, value=float(ledger.get("closing_cash_ledger", 0) or 0), step=100.0)
        save_ledger = st.form_submit_button("Save Ledger Working", type="primary")
    if save_ledger:
        save_computation_ledger(run_query, now_stamp, opening_credit, closing_credit, opening_cash, closing_cash)
        st.success("Ledger working saved.")

    st.subheader("HSN / SAC Bucket Details")
    hsn_data = load_computation_hsn(read_query)
    if hsn_data.empty:
        hsn_data = pd.DataFrame(columns=["hsn_code", "hsn_description", "taxable_value", "igst", "cgst", "sgst"])
    edited_hsn = st.data_editor(
        hsn_data.drop(columns=["id"], errors="ignore"),
        num_rows="dynamic", use_container_width=True, hide_index=True,
        key="gst_computation_hsn_editor",
        column_config={
            "hsn_code": st.column_config.TextColumn("HSN / SAC Code"),
            "hsn_description": st.column_config.TextColumn("Description"),
            "taxable_value": st.column_config.NumberColumn("Taxable Value", format="%.2f"),
            "igst": st.column_config.NumberColumn("IGST", format="%.2f"),
            "cgst": st.column_config.NumberColumn("CGST", format="%.2f"),
            "sgst": st.column_config.NumberColumn("SGST", format="%.2f"),
        },
    )
    if st.button("Save HSN Details", type="primary", key="save_gst_computation_hsn"):
        save_computation_hsn(run_query, now_stamp, edited_hsn)
        st.success("HSN details saved.")
        st.rerun()

    hsn_taxable = float(pd.to_numeric(edited_hsn.get("taxable_value", 0), errors="coerce").fillna(0).sum())
    hsn_tax = float(pd.to_numeric(edited_hsn.get("igst", 0), errors="coerce").fillna(0).sum())
    hsn_tax += float(pd.to_numeric(edited_hsn.get("cgst", 0), errors="coerce").fillna(0).sum())
    hsn_tax += float(pd.to_numeric(edited_hsn.get("sgst", 0), errors="coerce").fillna(0).sum())
    tolerance = 1.0
    sales_match = abs(hsn_taxable - sales_taxable) <= tolerance and abs(hsn_tax - (sales_igst + sales_cgst + sales_sgst)) <= tolerance
    gstr1_match = abs(hsn_taxable - gstr1_taxable) <= tolerance and abs(hsn_tax - gstr1_tax) <= tolerance
    st.subheader("HSN Reconciliation")
    match_cols = st.columns(2)
    if sales_match:
        match_cols[0].success("HSN matches Sales Register totals.")
    else:
        match_cols[0].warning("HSN does not match Sales Register totals.")
    if gstr1_match:
        match_cols[1].success("HSN matches GSTR-1 totals.")
    else:
        match_cols[1].warning("HSN does not match GSTR-1 totals.")
    st.caption(f"HSN buckets: INR {hsn_taxable:,.2f} taxable and INR {hsn_tax:,.2f} tax. Matching tolerance: INR {tolerance:,.2f}.")

    st.subheader("Professional Excel Working Paper")
    report_folder = os.path.join(os.path.dirname(db_path), "Reports")
    os.makedirs(report_folder, exist_ok=True)
    report_path = os.path.join(report_folder, f"GST_Computation_{from_date:%Y%m%d}_{to_date:%Y%m%d}.xlsx")
    if st.button("Create / Update Excel Computation Report", type="primary", key="create_gst_computation_report"):
        create_computation_excel(
            report_path,
            f"{from_date:%Y-%m-%d} to {to_date:%Y-%m-%d}",
            edited_worksheet,
            payable_rows,
            load_computation_ledger(read_query),
            edited_hsn,
            load_company_profile(),
        )
        st.success("Professional GST computation Excel report created.")
        with open(report_path, "rb") as report_file:
            st.download_button(
                "Download GST Computation Excel",
                data=report_file.read(),
                file_name=os.path.basename(report_path),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_gst_computation_report",
            )
