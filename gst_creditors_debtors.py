"""Creditors and debtors ageing management for the GST tool."""

import sqlite3
from datetime import date

import pandas as pd
import streamlit as st


PAYMENT_TABLE = "gst_payment_status"
MANUAL_TABLE = "gst_manual_parties"


def _connect(db_path):
    connection = sqlite3.connect(db_path)
    connection.execute(f"""
        CREATE TABLE IF NOT EXISTS {PAYMENT_TABLE} (
            register_type TEXT NOT NULL,
            invoice_id INTEGER NOT NULL,
            payment_status TEXT NOT NULL DEFAULT 'Pending',
            cleared_date TEXT,
            expected_payment_date TEXT,
            remarks TEXT DEFAULT '',
            PRIMARY KEY (register_type, invoice_id)
        )
    """)
    connection.execute(f"""
        CREATE TABLE IF NOT EXISTS {MANUAL_TABLE} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            register_type TEXT NOT NULL,
            party_name TEXT NOT NULL,
            gstin TEXT DEFAULT '',
            invoice_number TEXT NOT NULL,
            invoice_date TEXT NOT NULL,
            total_amount REAL DEFAULT 0,
            notes TEXT DEFAULT ''
        )
    """)
    connection.commit()
    return connection


def _load_register(db_path, register_type):
    table_name = "purchase_invoices" if register_type == "Creditor" else "gst_invoices"
    party_column = "supplier_name"
    connection = _connect(db_path)
    data = pd.read_sql_query(f"""
        SELECT id, {party_column} AS party_name, gstin, invoice_number,
               invoice_date, taxable_value, cgst, sgst, igst, total_amount
        FROM {table_name}
        ORDER BY date(invoice_date) ASC, id ASC
    """, connection)
    status = pd.read_sql_query(f"""
        SELECT invoice_id, payment_status, cleared_date,
               expected_payment_date, remarks
        FROM {PAYMENT_TABLE}
        WHERE register_type = ?
    """, connection, params=(register_type,))

    manual = pd.read_sql_query(f"""
        SELECT id, party_name, gstin, invoice_number, invoice_date,
               total_amount, notes
        FROM {MANUAL_TABLE}
        WHERE register_type = ?
        ORDER BY date(invoice_date) ASC, id ASC
    """, connection, params=(register_type,))
    connection.close()
    if not manual.empty:
        manual["id"] = -manual["id"].astype(int)
        manual["source"] = "Manual"
    if not data.empty:
        data["source"] = "Register"
    if data.empty:
        data = manual
    elif not manual.empty:
        data = pd.concat([data, manual], ignore_index=True, sort=False)
    if data.empty:
        return data
    data = data.merge(status, how="left", left_on="id", right_on="invoice_id")
    for column, default in {
        "payment_status": "Pending",
        "cleared_date": "",
        "expected_payment_date": "",
        "remarks": "",
    }.items():
        if column not in data.columns:
            data[column] = default
        else:
            data[column] = data[column].fillna(default)
    data["invoice_date"] = pd.to_datetime(data["invoice_date"], errors="coerce")
    data["total_amount"] = pd.to_numeric(data["total_amount"], errors="coerce").fillna(0)
    data["age_days"] = (pd.Timestamp(date.today()) - data["invoice_date"]).dt.days.clip(lower=0).fillna(0).astype(int)
    data["age_bucket"] = data.apply(_age_bucket, axis=1)
    data["register_type"] = register_type
    return data


def _add_manual_party(db_path, register_type, party_name, gstin, invoice_number, invoice_date, total_amount, notes):
    connection = _connect(db_path)
    connection.execute(f"""
        INSERT INTO {MANUAL_TABLE}
            (register_type, party_name, gstin, invoice_number, invoice_date, total_amount, notes)
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """, (register_type, party_name, gstin, invoice_number, str(invoice_date), total_amount, notes))
    connection.commit()
    connection.close()


def _delete_manual_parties(db_path, register_type, manual_ids):
    if not manual_ids:
        return
    placeholders = ",".join("?" for _ in manual_ids)
    connection = _connect(db_path)
    connection.execute(
        f"DELETE FROM {MANUAL_TABLE} WHERE register_type = ? AND id IN ({placeholders})",
        (register_type, *[abs(int(item)) for item in manual_ids]),
    )
    connection.commit()
    connection.close()


def _age_bucket(row):
    if str(row.get("payment_status", "Pending")) == "Cleared":
        return "Cleared"
    age = int(row.get("age_days", 0) or 0)
    if age > 3 * 365:
        return "Pending > 3 Years"
    if age > 365:
        return "Pending > 1 Year"
    if age > 180:
        return "Pending > 180 Days"
    if age > 90:
        return "Pending > 90 Days"
    return "Pending"


def _save_status(db_path, register_type, row):
    connection = _connect(db_path)
    connection.execute(f"""
        INSERT INTO {PAYMENT_TABLE}
            (register_type, invoice_id, payment_status, cleared_date,
             expected_payment_date, remarks)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(register_type, invoice_id) DO UPDATE SET
            payment_status = excluded.payment_status,
            cleared_date = excluded.cleared_date,
            expected_payment_date = excluded.expected_payment_date,
            remarks = excluded.remarks
    """, (
        register_type, int(row["id"]), str(row.get("payment_status", "Pending")),
        str(row.get("cleared_date", "") or ""),
        str(row.get("expected_payment_date", "") or ""),
        str(row.get("remarks", "") or ""),
    ))
    connection.commit()
    connection.close()


def _summary(data):
    if data.empty:
        return {"count": 0, "total": 0.0, "pending": 0.0, "overdue": 0.0}
    pending = data[data["payment_status"] != "Cleared"]
    overdue = pending[pending["age_days"] > 90]
    return {
        "count": len(data),
        "total": float(data["total_amount"].sum()),
        "pending": float(pending["total_amount"].sum()),
        "overdue": float(overdue["total_amount"].sum()),
    }


def _render_register(db_path, register_type):
    title = "Creditors" if register_type == "Creditor" else "Debtors"
    data = _load_register(db_path, register_type)
    st.subheader(f"{title} Register")
    if data.empty:
        st.info(f"No {title.lower()} found in the relevant register.")
        return

    summary = _summary(data)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Invoices", summary["count"])
    c2.metric("Total Value", f"₹ {summary['total']:,.0f}")
    c3.metric("Pending", f"₹ {summary['pending']:,.0f}")
    c4.metric("Pending > 90 Days", f"₹ {summary['overdue']:,.0f}")

    bucket_order = ["Cleared", "Pending", "Pending > 90 Days", "Pending > 180 Days", "Pending > 1 Year", "Pending > 3 Years"]
    counts = data["age_bucket"].value_counts().reindex(bucket_order, fill_value=0)
    st.bar_chart(counts)

    party_summary = data[data["payment_status"] != "Cleared"].groupby(
        ["party_name", "gstin"], dropna=False
    ).agg(Invoices=("id", "count"), Outstanding=("total_amount", "sum"), Oldest_Days=("age_days", "max")).reset_index()
    st.markdown("**Top outstanding parties**")
    st.dataframe(party_summary.sort_values("Outstanding", ascending=False).head(10), use_container_width=True, hide_index=True)

    filters = st.columns(2)
    with filters[0]:
        bucket_filter = st.multiselect("Filter by ageing category", bucket_order, key=f"bucket_{register_type}")
    with filters[1]:
        search = st.text_input("Search party / GSTIN / invoice", key=f"search_{register_type}").strip().lower()
    display = data if not bucket_filter else data[data["age_bucket"].isin(bucket_filter)]
    if search:
        searchable = display.astype(str).agg(" ".join, axis=1).str.lower()
        display = display[searchable.str.contains(search, na=False)]

    editable = display[[
        "id", "party_name", "gstin", "invoice_number", "invoice_date", "total_amount",
        "age_days", "age_bucket", "payment_status", "cleared_date", "expected_payment_date", "remarks"
    ]].copy()
    edited = st.data_editor(
        editable, use_container_width=True, hide_index=True,
        key=f"aging_editor_{register_type}",
        column_config={
            "payment_status": st.column_config.SelectboxColumn("Status", options=["Pending", "Cleared"]),
            "cleared_date": st.column_config.TextColumn("Cleared Date (YYYY-MM-DD)"),
            "expected_payment_date": st.column_config.TextColumn("Expected Payment Date"),
            "total_amount": st.column_config.NumberColumn("Invoice Value", format="₹ %.2f"),
            "age_days": st.column_config.NumberColumn("Age (days)"),
        },
        disabled=[column for column in editable.columns if column not in ["payment_status", "cleared_date", "expected_payment_date", "remarks"]],
    )
    if st.button(f"💾 Save {title} Status & Dates", type="primary", key=f"save_aging_{register_type}"):
        for _, row in edited.iterrows():
            _save_status(db_path, register_type, row)
        st.success(f"{title} payment status updated.")
        st.rerun()

    manual_rows = display[display["source"] == "Manual"].copy()
    if not manual_rows.empty:
        st.markdown("**Delete manually added records**")
        manual_rows.insert(0, "Delete", False)
        selected = st.data_editor(
            manual_rows[["Delete", "id", "party_name", "gstin", "invoice_number", "invoice_date", "total_amount"]],
            use_container_width=True, hide_index=True, key=f"manual_delete_{register_type}",
            disabled=["id", "party_name", "gstin", "invoice_number", "invoice_date", "total_amount"],
        )
        delete_ids = selected.loc[selected["Delete"], "id"].tolist()
        if st.button(f"🗑️ Delete Selected Manual {title}", key=f"delete_manual_{register_type}"):
            _delete_manual_parties(db_path, register_type, delete_ids)
            st.success("Selected manual records deleted.")
            st.rerun()


def render(db_path):
    st.markdown('<div class="main-title">💼 Creditors & Debtors Management</div>', unsafe_allow_html=True)
    st.caption("Creditors come from Purchase Register; debtors come from Sales Register. Ageing is calculated from invoice date until today.")
    if not db_path:
        st.warning("No active company database found. Please select a storage folder first.")
        return

    creditors = _load_register(db_path, "Creditor")
    debtors = _load_register(db_path, "Debtor")
    creditor_summary = _summary(creditors)
    debtor_summary = _summary(debtors)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Creditors Pending", f"₹ {creditor_summary['pending']:,.0f}")
    c2.metric("Debtors Pending", f"₹ {debtor_summary['pending']:,.0f}")
    c3.metric("Net Working Capital Gap", f"₹ {debtor_summary['pending'] - creditor_summary['pending']:,.0f}")
    c4.metric("Overdue > 90 Days", f"₹ {creditor_summary['overdue'] + debtor_summary['overdue']:,.0f}")

    with st.expander("➕ Add New Debtor / Creditor", expanded=False):
        with st.form("new_party_form"):
            new_type = st.selectbox("Record Type", ["Creditor", "Debtor"])
            new_party = st.text_input("Party Name *")
            new_gstin = st.text_input("GSTIN")
            new_invoice = st.text_input("Invoice / Reference Number *")
            new_date = st.date_input("Invoice / Due Date", value=date.today())
            new_amount = st.number_input("Amount", min_value=0.0, step=100.0)
            new_notes = st.text_area("Notes")
            add_new = st.form_submit_button("💾 Add New Record", type="primary")
        if add_new:
            if not new_party.strip() or not new_invoice.strip():
                st.error("Party Name and Invoice / Reference Number are required.")
            else:
                _add_manual_party(
                    db_path, new_type, new_party.strip(), new_gstin.strip().upper(),
                    new_invoice.strip(), new_date, new_amount, new_notes.strip(),
                )
                st.success(f"New {new_type.lower()} record added.")
                st.rerun()

    creditor_tab, debtor_tab = st.tabs(["📥 Creditors", "📤 Debtors"])
    with creditor_tab:
        _render_register(db_path, "Creditor")
    with debtor_tab:
        _render_register(db_path, "Debtor")