"""GST compliance details and portal access helper for the GST tool."""

import hashlib
import secrets
import sqlite3
import webbrowser
from datetime import date

import pandas as pd
import streamlit as st


GST_PORTAL_LOGIN_URL = "https://services.gst.gov.in/services/login"
EWAY_BILL_LOGIN_URL = "https://ewaybillgst.gov.in/Login.aspx"
EINVOICE_LOGIN_URL = "https://einvoice1.gst.gov.in/"
COMPLIANCE_TABLE = "gst_compliance_details"
CREDENTIALS_TABLE = "gst_portal_credentials"


def _connect(db_path):
    connection = sqlite3.connect(db_path)
    connection.execute(f"""
        CREATE TABLE IF NOT EXISTS {COMPLIANCE_TABLE} (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            return_type TEXT NOT NULL,
            return_period TEXT NOT NULL,
            filing_date TEXT,
            acknowledgement_number TEXT,
            status TEXT NOT NULL DEFAULT 'Filed',
            portal_username TEXT,
            portal_password_hash TEXT,
            portal_password_salt TEXT,
            notes TEXT,
            UNIQUE(return_type, return_period)
        )
    """)
    connection.execute(f"""
        CREATE TABLE IF NOT EXISTS {CREDENTIALS_TABLE} (
            portal_key TEXT PRIMARY KEY,
            portal_name TEXT NOT NULL,
            username TEXT NOT NULL DEFAULT '',
            password_hash TEXT NOT NULL DEFAULT '',
            password_salt TEXT NOT NULL DEFAULT '',
            updated_on TEXT
        )
    """)
    connection.commit()
    return connection


def _hash_portal_password(password, salt_hex=None):
    salt_hex = salt_hex or secrets.token_hex(16)
    password_hash = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), 200_000
    ).hex()
    return password_hash, salt_hex


def _load_portal_credentials(db_path, portal_key):
    connection = _connect(db_path)
    row = connection.execute(
        f"SELECT username FROM {CREDENTIALS_TABLE} WHERE portal_key = ?",
        (portal_key,),
    ).fetchone()
    connection.close()
    return row[0] if row else ""


def _save_portal_credentials(db_path, portal_key, portal_name, username, password):
    password_hash = password_salt = ""
    if password:
        password_hash, password_salt = _hash_portal_password(password)

    connection = _connect(db_path)
    connection.execute(f"""
        INSERT INTO {CREDENTIALS_TABLE}
            (portal_key, portal_name, username, password_hash, password_salt, updated_on)
        VALUES (?, ?, ?, ?, ?, datetime('now'))
        ON CONFLICT(portal_key) DO UPDATE SET
            portal_name = excluded.portal_name,
            username = excluded.username,
            password_hash = CASE
                WHEN excluded.password_hash <> '' THEN excluded.password_hash
                ELSE {CREDENTIALS_TABLE}.password_hash
            END,
            password_salt = CASE
                WHEN excluded.password_salt <> '' THEN excluded.password_salt
                ELSE {CREDENTIALS_TABLE}.password_salt
            END,
            updated_on = excluded.updated_on
    """, (portal_key, portal_name, username, password_hash, password_salt))
    connection.commit()
    connection.close()


def _copy_username(username):
    """Copy a username to the Windows clipboard when the desktop allows it."""
    if not username:
        return False
    try:
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()
        root.clipboard_clear()
        root.clipboard_append(username)
        root.update()
        root.destroy()
        return True
    except Exception:
        return False


def _load_details(db_path):
    connection = _connect(db_path)
    details = pd.read_sql_query(
        f"""SELECT return_type, return_period, filing_date,
                   acknowledgement_number, status, portal_username, notes
            FROM {COMPLIANCE_TABLE}
            ORDER BY return_period DESC, return_type""",
        connection,
    )
    connection.close()
    return details


def _save_details(db_path, return_type, return_period, filing_date,
                  acknowledgement_number, status, portal_username,
                  portal_password, notes):
    password_hash = password_salt = ""
    if portal_password:
        password_hash, password_salt = _hash_portal_password(portal_password)

    connection = _connect(db_path)
    connection.execute(f"""
        INSERT INTO {COMPLIANCE_TABLE} (
            return_type, return_period, filing_date, acknowledgement_number,
            status, portal_username, portal_password_hash, portal_password_salt, notes
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(return_type, return_period) DO UPDATE SET
            filing_date = excluded.filing_date,
            acknowledgement_number = excluded.acknowledgement_number,
            status = excluded.status,
            portal_username = excluded.portal_username,
            portal_password_hash = CASE
                WHEN excluded.portal_password_hash <> '' THEN excluded.portal_password_hash
                ELSE {COMPLIANCE_TABLE}.portal_password_hash
            END,
            portal_password_salt = CASE
                WHEN excluded.portal_password_salt <> '' THEN excluded.portal_password_salt
                ELSE {COMPLIANCE_TABLE}.portal_password_salt
            END,
            notes = excluded.notes
    """, (
        return_type, return_period, str(filing_date), acknowledgement_number,
        status, portal_username, password_hash, password_salt, notes
    ))
    connection.commit()
    connection.close()


def render(db_path):
    """Render compliance records and a link to the official GST portal."""
    st.markdown(
        '<div class="main-title">🛡️ GST Compliance & Portal Access</div>',
        unsafe_allow_html=True,
    )
    st.caption(
        "Record GST filing dates and references for each return period. "
        "Each portal has separate credentials. Clicking Login copies that portal's "
        "username and opens its official login page; password and CAPTCHA/OTP remain manual."
    )

    if not db_path:
        st.warning("No active company database found. Please select a storage folder first.")
        return

    st.subheader("Official Portal Logins")
    portal_columns = st.columns(3)
    portals = [
        ("GST Portal", "Main GST services and returns", GST_PORTAL_LOGIN_URL, "gst_portal"),
        ("E-Way Bill", "E-Way Bill System", EWAY_BILL_LOGIN_URL, "eway_bill"),
        ("e-Invoice", "e-Invoice System", EINVOICE_LOGIN_URL, "einvoice"),
    ]
    for column, (portal_name, description, portal_url, button_key) in zip(portal_columns, portals):
        with column:
            st.markdown(f"**{portal_name}**")
            st.caption(description)
            if st.button(
                "🌐 Login + Copy Username",
                use_container_width=True,
                type="primary",
                key=f"login_{button_key}",
            ):
                username = _load_portal_credentials(db_path, button_key)
                copied = _copy_username(username)
                webbrowser.open(portal_url)
                message = f"{portal_name} login page opened."
                if copied:
                    message += " Username copied to clipboard."
                elif not username:
                    message += " Save the username below first."
                st.success(message)
            st.link_button(
                "🔗 Open Login Page", portal_url,
                use_container_width=True,
                key=f"link_{button_key}",
            )

            with st.form(f"credentials_{button_key}"):
                saved_username = _load_portal_credentials(db_path, button_key)
                credential_username = st.text_input(
                    "Username", value=saved_username, key=f"username_{button_key}"
                )
                credential_password = st.text_input(
                    "Password", type="password", key=f"password_{button_key}",
                    help="Stored as a one-way hash; it cannot be auto-entered into the portal."
                )
                save_credentials = st.form_submit_button("💾 Save Credentials")
            if save_credentials:
                _save_portal_credentials(
                    db_path, button_key, portal_name,
                    credential_username.strip(), credential_password,
                )
                st.success(f"{portal_name} credentials saved separately.")

    details_col = st.container()
    with details_col:
        st.subheader("Add / Update Filing Details")
        with st.form("gst_compliance_form"):
            return_type = st.selectbox(
                "GST Return", ["GSTR-1", "GSTR-3B", "GSTR-9", "GSTR-9C", "Other"]
            )
            return_period = st.text_input("Return Period / Financial Year *", placeholder="2025-26 or 2026-08")
            filing_date = st.date_input("GST Filed Date", value=date.today())
            acknowledgement_number = st.text_input("ARN / Acknowledgement Number")
            status = st.selectbox("Filing Status", ["Filed", "Pending", "Amended", "Not Applicable"])
            portal_username = st.text_input("GST Portal Username")
            portal_password = st.text_input(
                "GST Portal Password (stored as a hash; leave blank to keep existing)",
                type="password",
            )
            notes = st.text_area("Notes")
            save = st.form_submit_button("💾 Save Compliance Details", type="primary")

        if save:
            if not return_period.strip():
                st.error("Return Period / Financial Year is required.")
            else:
                _save_details(
                    db_path, return_type, return_period.strip(), filing_date,
                    acknowledgement_number.strip(), status, portal_username.strip(),
                    portal_password, notes.strip(),
                )
                st.success("Compliance details saved securely in the active company database.")

    st.divider()
    st.subheader("GST Filing History")
    details = _load_details(db_path)
    if details.empty:
        st.info("No GST filing details recorded yet.")
    else:
        st.dataframe(details, use_container_width=True, hide_index=True)
