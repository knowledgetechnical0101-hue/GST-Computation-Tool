"""
GST Data Management & Reconciliation Tool
==========================================
A Streamlit desktop application for GST record-keeping, return tracking
and reconciliation.

Modules covered:
    - Sales invoices (GST outward supplies)
    - Purchase invoices + GSTR-2B (inward supplies / ITC)
    - GSTR-1 and GSTR-3B return summaries
    - Reconciliation: Purchase vs GSTR-2B, and GSTR-3B vs Sales
    - Dashboard with charts
    - Professional Excel export with a selectable date range
    - Backup and multi-company support (load any saved .db file)
"""

import streamlit as st
import pandas as pd
import numpy as np
import sqlite3
import os
import json
import shutil
import io
import re
import hashlib
import secrets
from datetime import datetime, date
import gstr9_9c
import gst_compliance
import gst_creditors_debtors
import gstr1_json_generator

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

try:
    import plotly.express as px
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="GST Data Management Tool",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

CONFIG_FILE = "gst_tool_config.json"
DATABASE_NAME = "GST_Data.db"
EXCEL_EXPORT_NAME = "GST_Data.xlsx"

GST_RATES = [0, 0.25, 3, 5, 12, 18, 28]
TRANSACTION_TYPES = ["Intrastate", "Interstate"]
SUPPLY_CATEGORIES = ["B2B", "B2C", "SEZ", "Export", "Deemed Export", "Other"]
SUPPLY_TYPES = ["Goods", "Services", "Goods and Services"]

# ---- Client registration / login ----
# Every registered client gets their own folder under CLIENTS_ROOT_DIR,
# named after their login username, holding a private GST_Data.db,
# a Backups sub-folder, and nothing else. Credentials (never plain-text
# passwords) live centrally in CLIENT_AUTH_FILE.
CLIENTS_ROOT_DIR = "GST_Clients"
CLIENT_AUTH_FILE = "gst_clients_auth.json"
CLIENT_AUTH_TABLE = "client_auth"
PBKDF2_ITERATIONS = 200_000


# ============================================================
# CUSTOM CSS
# ============================================================

st.markdown("""
<style>

@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
}

:root {
    --gst-primary: #2f5fed;
    --gst-primary-dark: #1e3fb8;
    --gst-bg-soft: #f5f7fb;
    --gst-border: #e3e6ee;
    --gst-success: #1a7f37;
    --gst-warning: #b35900;
    --gst-danger: #c62828;
}

.main-title {
    font-size: 28px;
    font-weight: 800;
    color: #1a1d29;
    padding-bottom: 4px;
    border-bottom: 3px solid var(--gst-primary);
    margin-bottom: 18px;
    display: inline-block;
}

.section-title {
    font-size: 20px;
    font-weight: 700;
    color: #1a1d29;
    margin-top: 6px;
}

.gst-subtle {
    color: #6b7280;
    font-size: 14px;
}

.gst-card {
    background: #ffffff;
    border: 1px solid var(--gst-border);
    border-radius: 12px;
    padding: 18px 20px;
    margin-bottom: 14px;
    box-shadow: 0 1px 3px rgba(16, 24, 40, 0.04);
}

.gst-step-badge {
    display: inline-block;
    background: var(--gst-primary);
    color: white;
    font-weight: 700;
    font-size: 12px;
    padding: 3px 10px;
    border-radius: 999px;
    margin-right: 8px;
    letter-spacing: 0.3px;
}

/* Metric cards */
div[data-testid="stMetric"] {
    border: 1px solid var(--gst-border);
    background: #ffffff;
    padding: 14px 16px;
    border-radius: 12px;
    box-shadow: 0 1px 3px rgba(16, 24, 40, 0.04);
}

/* Buttons */
div.stButton > button, div.stFormSubmitButton > button, div.stDownloadButton > button {
    border-radius: 8px;
    font-weight: 600;
    border: 1px solid var(--gst-border);
    transition: all 0.15s ease-in-out;
}

div.stButton > button[kind="primary"], div.stFormSubmitButton > button[kind="primary"],
div.stDownloadButton > button[kind="primary"] {
    background-color: var(--gst-primary);
    border-color: var(--gst-primary);
    color: white;
}

div.stButton > button[kind="primary"]:hover, div.stFormSubmitButton > button[kind="primary"]:hover,
div.stDownloadButton > button[kind="primary"]:hover {
    background-color: var(--gst-primary-dark);
    border-color: var(--gst-primary-dark);
}

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color: var(--gst-bg-soft);
    border-right: 1px solid var(--gst-border);
}

section[data-testid="stSidebar"] .stRadio label {
    font-size: 15px;
}

/* Tabs */
button[data-baseweb="tab"] {
    font-weight: 600;
}

/* Status text */
.status-matched {
    color: var(--gst-success);
    font-weight: 700;
}

.status-mismatch {
    color: var(--gst-warning);
    font-weight: 700;
}

.status-missing {
    color: var(--gst-danger);
    font-weight: 700;
}

hr {
    margin: 0.6em 0;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# SAVE / LOAD CONFIGURATION
# ============================================================

def save_config(path=None, active_db=None, recent_db=None):
    """Config keeps: the default storage folder, the currently active
    database file (which may be a company loaded from elsewhere), and
    a short list of recently used company databases for quick switching."""

    config = load_config_raw()

    if path is not None:
        config["save_path"] = path

    if active_db is not None:
        config["active_db"] = active_db

    if recent_db is not None:
        recent = config.get("recent_companies", [])
        if recent_db not in recent:
            recent.insert(0, recent_db)
        config["recent_companies"] = recent[:8]

    with open(CONFIG_FILE, "w") as file:
        json.dump(config, file, indent=4)


def load_config_raw():

    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as file:
                return json.load(file)
        except Exception:
            return {}

    return {}


def load_config():
    return load_config_raw().get("save_path", "")


# ============================================================
# CLIENT REGISTRATION & LOGIN
# ============================================================
#
# Each client (business) registers once with their own username and
# password. Registration creates a dedicated, private folder for that
# client under CLIENTS_ROOT_DIR — this is where their GST database and
# backups live, kept completely separate from every other client.
# Passwords are never stored in plain text: only a salted PBKDF2-HMAC
# hash is written to disk.

def hash_password(password, salt_hex=None):
    """Returns (hash_hex, salt_hex) for the given password. Pass an
    existing salt_hex back in to verify a password against a stored hash."""

    salt_hex = salt_hex or secrets.token_hex(16)
    derived = hashlib.pbkdf2_hmac(
        "sha256", password.encode("utf-8"), bytes.fromhex(salt_hex), PBKDF2_ITERATIONS
    )
    return derived.hex(), salt_hex


def load_clients():
    if os.path.exists(CLIENT_AUTH_FILE):
        try:
            with open(CLIENT_AUTH_FILE, "r") as file:
                data = json.load(file)
                return data if isinstance(data, dict) else {}
        except Exception:
            return {}
    return {}


def save_clients(clients):
    """Atomically persist the central account index."""
    temp_file = CLIENT_AUTH_FILE + ".tmp"
    with open(temp_file, "w") as file:
        json.dump(clients, file, indent=4)
    os.replace(temp_file, CLIENT_AUTH_FILE)


def ensure_client_auth_table(db_path, record=None):
    """Create/sync authentication metadata inside the client's SQLite DB.

    This makes a saved GST_Data.db portable: the database itself carries a
    salted password hash and account metadata, while the central JSON file
    remains a fast account index. The plain-text password is never stored.
    """
    if not db_path:
        return False
    try:
        os.makedirs(os.path.dirname(os.path.abspath(db_path)) or ".", exist_ok=True)
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {CLIENT_AUTH_TABLE} (
                id INTEGER PRIMARY KEY CHECK (id = 1),
                username TEXT NOT NULL,
                password_hash TEXT NOT NULL,
                salt TEXT NOT NULL,
                display_name TEXT DEFAULT '',
                gstin TEXT DEFAULT '',
                email TEXT DEFAULT '',
                phone TEXT DEFAULT '',
                state TEXT DEFAULT '',
                recovery_hash TEXT DEFAULT '',
                recovery_salt TEXT DEFAULT '',
                created_on TEXT,
                updated_on TEXT
            )
        """)
        if record:
            cur.execute(f"""
                INSERT INTO {CLIENT_AUTH_TABLE}
                (id, username, password_hash, salt, display_name, gstin, email, phone, state,
                 recovery_hash, recovery_salt, created_on, updated_on)
                VALUES (1, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(id) DO UPDATE SET
                    username=excluded.username, password_hash=excluded.password_hash,
                    salt=excluded.salt, display_name=excluded.display_name, gstin=excluded.gstin,
                    email=excluded.email, phone=excluded.phone, state=excluded.state,
                    recovery_hash=excluded.recovery_hash, recovery_salt=excluded.recovery_salt,
                    created_on=excluded.created_on, updated_on=excluded.updated_on
            """, (
                record.get("username", ""), record.get("password_hash", ""), record.get("salt", ""),
                record.get("display_name", ""), record.get("gstin", ""), record.get("email", ""),
                record.get("phone", ""), record.get("state", ""), record.get("recovery_hash", ""),
                record.get("recovery_salt", ""), record.get("created_on", ""), now_stamp()
            ))
        conn.commit()
        conn.close()
        return True
    except Exception:
        try:
            conn.close()
        except Exception:
            pass
        return False


def read_embedded_client_auth(db_path):
    """Read embedded account metadata from a saved GST database, if present."""
    if not db_path or not os.path.exists(db_path):
        return None
    try:
        conn = sqlite3.connect(db_path)
        cur = conn.cursor()
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (CLIENT_AUTH_TABLE,))
        if not cur.fetchone():
            conn.close()
            return None
        cur.execute(f"SELECT username, password_hash, salt, display_name, gstin, email, phone, state, recovery_hash, recovery_salt, created_on FROM {CLIENT_AUTH_TABLE} WHERE id=1")
        row = cur.fetchone()
        conn.close()
        if not row:
            return None
        keys = ["username", "password_hash", "salt", "display_name", "gstin", "email", "phone", "state", "recovery_hash", "recovery_salt", "created_on"]
        return dict(zip(keys, row))
    except Exception:
        return None


def discover_embedded_client(username_key):
    """Find a portable account record by scanning registered client DBs."""
    root = os.path.abspath(CLIENTS_ROOT_DIR)
    if not os.path.isdir(root):
        return None
    for folder_name in os.listdir(root):
        folder = os.path.join(root, folder_name)
        db_path = os.path.join(folder, DATABASE_NAME)
        if not os.path.isfile(db_path):
            continue
        record = read_embedded_client_auth(db_path)
        if record and str(record.get("username", "")).strip().lower() == username_key:
            record["folder"] = folder
            return record
    return None


def _make_recovery_key():
    return secrets.token_urlsafe(18).replace("-", "").replace("_", "")[:16].upper()


def hash_recovery_key(recovery_key, salt_hex=None):
    return hash_password(recovery_key, salt_hex)


def verify_recovery_key(record, recovery_key):
    if not recovery_key or not record.get("recovery_hash") or not record.get("recovery_salt"):
        return False
    computed, _ = hash_recovery_key(recovery_key.strip(), record.get("recovery_salt"))
    return secrets.compare_digest(computed, record.get("recovery_hash", ""))

def update_client_storage_path(username_key, new_folder, move_existing_data=True):
    """Changes a client's data folder and optionally copies its data there."""

    clients = load_clients()
    record = clients.get(username_key)
    if not record:
        return False, "Account not found."

    old_folder = os.path.abspath(record.get("folder", ""))
    new_folder = os.path.abspath(os.path.expanduser(new_folder.strip()))

    if not new_folder:
        return False, "Please enter a data folder path."
    if old_folder == new_folder:
        return True, "The data folder is already set to this location."

    try:
        os.makedirs(new_folder, exist_ok=True)

        if move_existing_data:
            old_db = os.path.join(old_folder, DATABASE_NAME)
            new_db = os.path.join(new_folder, DATABASE_NAME)
            if os.path.exists(old_db) and not os.path.exists(new_db):
                shutil.copy2(old_db, new_db)

            old_backups = os.path.join(old_folder, "Backups")
            new_backups = os.path.join(new_folder, "Backups")
            if os.path.exists(old_backups):
                os.makedirs(new_backups, exist_ok=True)
                for backup_name in os.listdir(old_backups):
                    source = os.path.join(old_backups, backup_name)
                    target = os.path.join(new_backups, backup_name)
                    if os.path.isfile(source) and not os.path.exists(target):
                        shutil.copy2(source, target)
    except OSError as error:
        return False, f"Could not prepare the new data folder: {error}"

    record["folder"] = new_folder
    record["username"] = username_key
    clients[username_key] = record
    save_clients(clients)
    ensure_client_auth_table(os.path.join(new_folder, DATABASE_NAME), record)
    return True, "Data folder updated successfully."


def client_folder_path(username_key):
    return os.path.join(CLIENTS_ROOT_DIR, username_key)


def register_client(username, password, confirm_password, display_name, gstin, email, phone, state):
    """Create a client account and embed its authentication metadata in the DB."""
    username_key = re.sub(r"[^a-z0-9_.-]", "", username.strip().lower())

    if not username_key:
        return False, "Please choose a valid username (letters, numbers, '.', '_', '-' only)."
    if not display_name.strip():
        return False, "Client / business name is required."
    if len(password) < 6:
        return False, "Password must be at least 6 characters long."
    if password != confirm_password:
        return False, "Password and Confirm Password do not match."

    clients = load_clients()
    if username_key in clients:
        return False, "This username is already registered. Please choose another, or log in instead."

    folder = client_folder_path(username_key)
    os.makedirs(folder, exist_ok=True)
    os.makedirs(os.path.join(folder, "Backups"), exist_ok=True)

    pwd_hash, salt_hex = hash_password(password)
    recovery_key = _make_recovery_key()
    recovery_hash, recovery_salt = hash_recovery_key(recovery_key)
    record = {
        "username": username_key,
        "password_hash": pwd_hash,
        "salt": salt_hex,
        "display_name": display_name.strip(),
        "gstin": gstin.strip().upper(),
        "email": email.strip(),
        "phone": phone.strip(),
        "state": state.strip(),
        "folder": folder,
        "recovery_hash": recovery_hash,
        "recovery_salt": recovery_salt,
        "created_on": now_stamp(),
    }
    save_clients({**clients, username_key: record})
    ensure_client_auth_table(os.path.join(folder, DATABASE_NAME), record)
    return True, f"Registration successful. Save this Recovery Key somewhere safe: {recovery_key}"


def authenticate_client(username, password):
    """Authenticate from the central index or portable embedded DB metadata."""
    username_key = username.strip().lower()
    clients = load_clients()
    record = clients.get(username_key)

    # Repair a missing/stale central account index from the saved database.
    if not record:
        record = discover_embedded_client(username_key)
        if record:
            clients[username_key] = record
            save_clients(clients)

    if not record:
        return False, None

    computed_hash, _ = hash_password(password, record.get("salt"))
    if secrets.compare_digest(computed_hash, record.get("password_hash", "")):
        record = {**record, "username": username_key}
        record.setdefault("folder", client_folder_path(username_key))
        # Keep both authentication stores synchronized.
        clients[username_key] = record
        save_clients(clients)
        ensure_client_auth_table(os.path.join(record["folder"], DATABASE_NAME), record)
        return True, record

    # If JSON is stale but the DB has the current credentials, repair JSON.
    folder = record.get("folder") or client_folder_path(username_key)
    embedded = read_embedded_client_auth(os.path.join(folder, DATABASE_NAME))
    if embedded:
        embedded["folder"] = folder
        embedded_hash, _ = hash_password(password, embedded.get("salt"))
        if secrets.compare_digest(embedded_hash, embedded.get("password_hash", "")):
            clients[username_key] = embedded
            save_clients(clients)
            return True, {**embedded, "username": username_key}

    return False, None


def reset_client_password(username, recovery_key, new_password, confirm_new_password):
    """Reset a password using the account's recovery key."""
    username_key = username.strip().lower()
    if len(new_password) < 6:
        return False, "New password must be at least 6 characters long."
    if new_password != confirm_new_password:
        return False, "New password and confirmation do not match."

    clients = load_clients()
    record = clients.get(username_key) or discover_embedded_client(username_key)
    if not record:
        return False, "Account not found. If you only have a saved GST database, use Restore / Repair Account below."

    if not verify_recovery_key(record, recovery_key):
        return False, "Invalid Recovery Key."

    pwd_hash, salt_hex = hash_password(new_password)
    record["password_hash"] = pwd_hash
    record["salt"] = salt_hex
    record["username"] = username_key
    record["updated_on"] = now_stamp()
    clients[username_key] = record
    save_clients(clients)
    ensure_client_auth_table(os.path.join(record.get("folder", client_folder_path(username_key)), DATABASE_NAME), record)
    return True, "Password reset successfully. You can now log in."


def recover_account_from_saved_database(uploaded_bytes, username, gstin, new_password, confirm_new_password):
    """Restore a saved GST_Data.db and create/repair its login account."""
    username_key = re.sub(r"[^a-z0-9_.-]", "", username.strip().lower())
    gstin_key = re.sub(r"\s+", "", gstin.strip().upper())
    if not username_key:
        return False, "Please enter a valid username."
    if not gstin_key:
        return False, "GSTIN is required to verify the saved database."
    if len(new_password) < 6:
        return False, "New password must be at least 6 characters long."
    if new_password != confirm_new_password:
        return False, "New password and confirmation do not match."
    if not uploaded_bytes:
        return False, "Please select your saved GST_Data.db file."

    temp_path = os.path.join(os.path.abspath(CLIENTS_ROOT_DIR), f"._restore_{secrets.token_hex(8)}.db")
    try:
        os.makedirs(os.path.abspath(CLIENTS_ROOT_DIR), exist_ok=True)
        with open(temp_path, "wb") as file:
            file.write(uploaded_bytes)
        conn = sqlite3.connect(temp_path)
        cur = conn.cursor()
        integrity = cur.execute("PRAGMA integrity_check").fetchone()
        if not integrity or integrity[0] != "ok":
            conn.close()
            return False, "The selected database failed SQLite integrity check."
        db_gstin = ""
        try:
            row = cur.execute("SELECT gstin FROM company_profile WHERE id=1").fetchone()
            db_gstin = re.sub(r"\s+", "", str(row[0]).upper()) if row and row[0] else ""
        except Exception:
            pass
        if db_gstin and db_gstin != gstin_key:
            conn.close()
            return False, f"GSTIN mismatch. The saved database belongs to GSTIN {db_gstin}, not {gstin_key}."
        old_auth = read_embedded_client_auth(temp_path)
        conn.close()

        folder = os.path.abspath(client_folder_path(username_key))
        os.makedirs(os.path.join(folder, "Backups"), exist_ok=True)
        target_db = os.path.join(folder, DATABASE_NAME)
        if os.path.abspath(temp_path) != os.path.abspath(target_db):
            shutil.copy2(temp_path, target_db)

        pwd_hash, salt_hex = hash_password(new_password)
        recovery_key = _make_recovery_key()
        recovery_hash, recovery_salt = hash_recovery_key(recovery_key)
        record = {
            "username": username_key,
            "password_hash": pwd_hash,
            "salt": salt_hex,
            "display_name": (old_auth or {}).get("display_name", "") or "Restored GST Account",
            "gstin": gstin_key,
            "email": (old_auth or {}).get("email", ""),
            "phone": (old_auth or {}).get("phone", ""),
            "state": (old_auth or {}).get("state", ""),
            "folder": folder,
            "recovery_hash": recovery_hash,
            "recovery_salt": recovery_salt,
            "created_on": (old_auth or {}).get("created_on", now_stamp()),
        }
        clients = load_clients()
        clients[username_key] = record
        save_clients(clients)
        ensure_client_auth_table(target_db, record)
        return True, f"Account restored successfully. Save this new Recovery Key: {recovery_key}"
    except Exception as error:
        return False, f"Could not restore the saved database: {error}"
    finally:
        try:
            if os.path.exists(temp_path):
                os.remove(temp_path)
        except OSError:
            pass

def change_client_password(username_key, current_password, new_password, confirm_new_password):
    """Change password and keep central + embedded authentication in sync."""
    clients = load_clients()
    record = clients.get(username_key)
    if not record:
        return False, "Account not found."

    computed_hash, _ = hash_password(current_password, record.get("salt"))
    if not secrets.compare_digest(computed_hash, record.get("password_hash", "")):
        return False, "Current password is incorrect."
    if len(new_password) < 6:
        return False, "New password must be at least 6 characters long."
    if new_password != confirm_new_password:
        return False, "New password and confirmation do not match."

    pwd_hash, salt_hex = hash_password(new_password)
    record["password_hash"] = pwd_hash
    record["salt"] = salt_hex
    record["updated_on"] = now_stamp()
    clients[username_key] = record
    save_clients(clients)
    ensure_client_auth_table(os.path.join(record.get("folder", client_folder_path(username_key)), DATABASE_NAME), record)
    return True, "Password updated successfully."

def render_login_register_screen():
    """Renders the full-page login / registration gate. Nothing else in
    the application runs until a client is logged in."""

    st.markdown(
        '<div class="main-title">📊 GST Data Management Tool</div>'
        '<div class="gst-subtle" style="margin-bottom:18px;">Client Login &amp; Registration Portal</div>',
        unsafe_allow_html=True
    )

    tab_login, tab_register, tab_recovery = st.tabs(["🔐 Client Login", "🆕 New Client Registration", "🛠️ Recover / Restore Account"])

    with tab_login:
        st.markdown('<div class="gst-card">', unsafe_allow_html=True)
        st.subheader("Log in to your GST workspace")

        with st.form("client_login_form"):
            login_username = st.text_input("Username")
            login_password = st.text_input("Password", type="password")
            login_submit = st.form_submit_button("🔐 Log In", use_container_width=True, type="primary")

        if login_submit:
            if not login_username or not login_password:
                st.error("Please enter both your username and password.")
            else:
                ok, record = authenticate_client(login_username, login_password)
                if ok:
                    os.makedirs(record["folder"], exist_ok=True)
                    os.makedirs(os.path.join(record["folder"], "Backups"), exist_ok=True)

                    st.session_state.auth_user = record["username"]
                    st.session_state.client_record = record
                    st.session_state.save_path = record["folder"]
                    st.session_state.active_db = ""

                    st.success(f"✅ Welcome back, {record.get('display_name') or record['username']}!")
                    st.rerun()
                else:
                    st.error("Invalid username or password. If you restored or moved your GST database, use the Recover / Restore Account tab.")

        with st.expander("🔑 Forgot Password? Use Recovery Key"):
            st.caption("Your Recovery Key is created during registration or account restoration. It is not your GSTIN or normal password.")
            with st.form("forgot_password_form"):
                fp_username = st.text_input("Username", key="fp_username")
                fp_key = st.text_input("Recovery Key", type="password", key="fp_key")
                fp_new = st.text_input("New Password", type="password", key="fp_new")
                fp_confirm = st.text_input("Confirm New Password", type="password", key="fp_confirm")
                fp_submit = st.form_submit_button("🔄 Reset Password", use_container_width=True)
            if fp_submit:
                ok, msg = reset_client_password(fp_username, fp_key, fp_new, fp_confirm)
                if ok:
                    st.success(msg)
                else:
                    st.error(msg)

        st.markdown('</div>', unsafe_allow_html=True)

    with tab_register:
        st.markdown('<div class="gst-card">', unsafe_allow_html=True)
        st.subheader("Register a new client")
        st.caption(
            "This creates a private, password-protected workspace for the client. "
            "All of their sales, purchase, GSTR-1/2B/3B and reconciliation data will be "
            "stored in a dedicated folder that only they can log into."
        )

        with st.form("client_register_form"):
            col1, col2 = st.columns(2)

            with col1:
                reg_display_name = st.text_input("Client / Business Name *")
                reg_gstin = st.text_input("GSTIN *").upper()
                reg_state = st.text_input("State")
                reg_email = st.text_input("Email")
                reg_phone = st.text_input("Phone")

            with col2:
                reg_username = st.text_input("Choose a Username *")
                reg_password = st.text_input("Choose a Password *", type="password")
                reg_confirm_password = st.text_input("Confirm Password *", type="password")

            register_submit = st.form_submit_button("🆕 Register Client", use_container_width=True, type="primary")

        if register_submit:
            success, message = register_client(
                reg_username, reg_password, reg_confirm_password,
                reg_display_name, reg_gstin, reg_email, reg_phone, reg_state
            )
            if success:
                st.success(f"✅ {message}")
            else:
                st.error(message)

        st.markdown('</div>', unsafe_allow_html=True)

    with tab_recovery:
        st.markdown('<div class="gst-card">', unsafe_allow_html=True)
        st.subheader("Recover / Restore a Saved GST Database")
        st.caption(
            "If your GST data is saved but login says invalid password, you can restore the saved "
            "GST_Data.db here. The GSTIN inside the database is checked when available, then a new "
            "password and Recovery Key are created. This does not delete your GST transactions."
        )
        restore_file = st.file_uploader(
            "Select your saved GST_Data.db", type=["db", "sqlite", "sqlite3"], key="restore_db_file"
        )
        with st.form("restore_account_form"):
            ra_username = st.text_input("Username to use for this account *")
            ra_gstin = st.text_input("GSTIN in the saved database *").upper()
            ra_new = st.text_input("Create New Password *", type="password")
            ra_confirm = st.text_input("Confirm New Password *", type="password")
            ra_submit = st.form_submit_button("🛠️ Restore Account & Database", use_container_width=True, type="primary")
        if ra_submit:
            uploaded_bytes = restore_file.getvalue() if restore_file else None
            ok, msg = recover_account_from_saved_database(
                uploaded_bytes, ra_username, ra_gstin, ra_new, ra_confirm
            )
            if ok:
                st.success(msg)
                st.info("Please save the Recovery Key now. Then return to Client Login and sign in with your new password.")
            else:
                st.error(msg)
        st.markdown('</div>', unsafe_allow_html=True)


# ============================================================
# SESSION STATE
# ============================================================

if "save_path" not in st.session_state:
    st.session_state.save_path = load_config()

if "active_db" not in st.session_state:
    st.session_state.active_db = load_config_raw().get("active_db", "")

if "auth_user" not in st.session_state:
    st.session_state.auth_user = None

if "client_record" not in st.session_state:
    st.session_state.client_record = None


# ============================================================
# DATABASE PATH
# ============================================================

def get_db_path():
    """Returns the database currently in use. If the user has loaded a
    specific company .db file, that takes priority over the default
    storage-folder database."""

    if st.session_state.active_db:
        return st.session_state.active_db

    if not st.session_state.save_path:
        return None

    return os.path.join(st.session_state.save_path, DATABASE_NAME)


def get_active_company_label():

    db_path = get_db_path()

    if not db_path:
        return "No company loaded"

    return os.path.basename(db_path)


# ============================================================
# CREATE DATABASE (ALL TABLES)
# ============================================================

def create_database():

    db_path = get_db_path()

    if not db_path:
        return

    os.makedirs(os.path.dirname(db_path) or ".", exist_ok=True)

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    # ---- Sales invoices (outward supplies) ----
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS gst_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_name TEXT NOT NULL,
            gstin TEXT NOT NULL,
            invoice_number TEXT NOT NULL,
            invoice_date TEXT NOT NULL,
            taxable_value REAL DEFAULT 0,
            gst_rate REAL DEFAULT 0,
            transaction_type TEXT NOT NULL,
            cgst REAL DEFAULT 0,
            sgst REAL DEFAULT 0,
            igst REAL DEFAULT 0,
            total_amount REAL DEFAULT 0,
            supply_category TEXT DEFAULT 'B2B',
            supply_type TEXT DEFAULT 'Goods',
            created_date TEXT,
            modified_date TEXT
        )
    """)

    # ---- Purchase invoices (inward supplies) ----
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS purchase_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_name TEXT NOT NULL,
            gstin TEXT NOT NULL,
            invoice_number TEXT NOT NULL,
            invoice_date TEXT NOT NULL,
            taxable_value REAL DEFAULT 0,
            gst_rate REAL DEFAULT 0,
            transaction_type TEXT NOT NULL,
            cgst REAL DEFAULT 0,
            sgst REAL DEFAULT 0,
            igst REAL DEFAULT 0,
            total_amount REAL DEFAULT 0,
            itc_eligible TEXT DEFAULT 'Yes',
            supply_category TEXT DEFAULT 'B2B',
            supply_type TEXT DEFAULT 'Goods',
            return_period TEXT,
            created_date TEXT,
            modified_date TEXT
        )
    """)

    # ---- GSTR-2B (as reflected on the portal) ----
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS gstr_2b (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier_name TEXT,
            gstin TEXT NOT NULL,
            invoice_number TEXT NOT NULL,
            invoice_date TEXT,
            taxable_value REAL DEFAULT 0,
            cgst REAL DEFAULT 0,
            sgst REAL DEFAULT 0,
            igst REAL DEFAULT 0,
            return_period TEXT,
            itc_eligible TEXT DEFAULT 'Eligible',
            itc_remark TEXT DEFAULT '',
            created_date TEXT
        )
    """)

    # ---- GSTR-1 (invoice-wise, as filed on the portal) ----
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS gstr1_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_name TEXT,
            gstin TEXT NOT NULL,
            invoice_number TEXT NOT NULL,
            invoice_date TEXT,
            taxable_value REAL DEFAULT 0,
            gst_rate REAL DEFAULT 0,
            transaction_type TEXT,
            cgst REAL DEFAULT 0,
            sgst REAL DEFAULT 0,
            igst REAL DEFAULT 0,
            total_amount REAL DEFAULT 0,
            return_period TEXT,
            created_date TEXT,
            modified_date TEXT
        )
    """)

    # ---- GSTR-1 summary (return-level totals as filed) ----
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS gstr1_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            return_period TEXT NOT NULL UNIQUE,
            total_taxable_value REAL DEFAULT 0,
            total_tax REAL DEFAULT 0,
            invoice_count INTEGER DEFAULT 0,
            filed_date TEXT,
            created_date TEXT,
            modified_date TEXT
        )
    """)

    # ---- GSTR-3B summary (return filed) ----
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS gstr3b_summary (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            return_period TEXT NOT NULL UNIQUE,
            outward_taxable_value REAL DEFAULT 0,
            output_tax REAL DEFAULT 0,
            itc_claimed REAL DEFAULT 0,
            itc_reversed REAL DEFAULT 0,
            tax_paid_cash REAL DEFAULT 0,
            late_fee REAL DEFAULT 0,
            filed_date TEXT,
            created_date TEXT,
            modified_date TEXT
        )
    """)

    # ---- Company profile (single row, for report headers) ----
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS company_profile (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            company_name TEXT DEFAULT '',
            gstin TEXT DEFAULT '',
            state TEXT DEFAULT '',
            financial_year TEXT DEFAULT ''
        )
    """)

    cursor.execute("""
        INSERT OR IGNORE INTO company_profile (id, company_name, gstin, state, financial_year)
        VALUES (1, '', '', '', '')
    """)

    # ---- Backward-compatible column upgrades for existing installs ----
    cursor.execute("PRAGMA table_info(gst_invoices)")
    existing_columns = {row[1] for row in cursor.fetchall()}

    for column_name, column_type in {
        "created_date": "TEXT", "modified_date": "TEXT",
        "supply_category": "TEXT DEFAULT 'B2B'",
        "supply_type": "TEXT DEFAULT 'Goods'",
    }.items():
        if column_name not in existing_columns:
            cursor.execute(f"ALTER TABLE gst_invoices ADD COLUMN {column_name} {column_type}")

    cursor.execute("PRAGMA table_info(gstr3b_summary)")
    existing_gstr3b_columns = {row[1] for row in cursor.fetchall()}
    if "itc_reversed" not in existing_gstr3b_columns:
        cursor.execute("ALTER TABLE gstr3b_summary ADD COLUMN itc_reversed REAL DEFAULT 0")

    cursor.execute("PRAGMA table_info(gstr_2b)")
    existing_2b_columns = {row[1] for row in cursor.fetchall()}
    for column_name, column_type in {
        "itc_eligible": "TEXT DEFAULT 'Eligible'",
        "itc_remark": "TEXT DEFAULT ''",
    }.items():
        if column_name not in existing_2b_columns:
            cursor.execute(f"ALTER TABLE gstr_2b ADD COLUMN {column_name} {column_type}")
    cursor.execute("UPDATE gstr_2b SET itc_eligible = 'Eligible' WHERE itc_eligible IS NULL OR TRIM(itc_eligible) = ''")
    cursor.execute("UPDATE gstr_2b SET itc_remark = '' WHERE itc_remark IS NULL")

    cursor.execute("PRAGMA table_info(purchase_invoices)")
    existing_purchase_columns = {row[1] for row in cursor.fetchall()}
    for column_name, column_type in {
        "return_period": "TEXT",
        "supply_category": "TEXT DEFAULT 'B2B'",
        "supply_type": "TEXT DEFAULT 'Goods'",
    }.items():
        if column_name not in existing_purchase_columns:
            cursor.execute(f"ALTER TABLE purchase_invoices ADD COLUMN {column_name} {column_type}")
    cursor.execute("""
        UPDATE purchase_invoices
        SET return_period = strftime('%Y-%m', invoice_date)
        WHERE (return_period IS NULL OR TRIM(return_period) = '')
          AND invoice_date IS NOT NULL
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS gst_credit_debit_notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            register_type TEXT NOT NULL,
            note_type TEXT NOT NULL,
            party_name TEXT DEFAULT '',
            gstin TEXT DEFAULT '',
            note_number TEXT NOT NULL,
            note_date TEXT NOT NULL,
            taxable_value REAL DEFAULT 0,
            cgst REAL DEFAULT 0,
            sgst REAL DEFAULT 0,
            igst REAL DEFAULT 0,
            total_amount REAL DEFAULT 0,
            reason TEXT DEFAULT '',
            created_date TEXT
        )
    """)
    cursor.execute("UPDATE gst_invoices SET supply_category = 'B2B' WHERE supply_category IS NULL OR TRIM(supply_category) = ''")
    cursor.execute("UPDATE purchase_invoices SET supply_category = 'B2B' WHERE supply_category IS NULL OR TRIM(supply_category) = ''")

    conn.commit()
    conn.close()


# ============================================================
# GENERIC HELPERS
# ============================================================

def run_query(query, params=()):
    """Runs an INSERT / UPDATE / DELETE against the active database."""

    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute(query, params)
    conn.commit()
    conn.close()


def read_query(query, params=()):
    """Runs a SELECT against the active database and returns a DataFrame."""

    db_path = get_db_path()

    if not db_path or not os.path.exists(db_path):
        return pd.DataFrame()

    conn = sqlite3.connect(db_path)
    try:
        data = pd.read_sql_query(query, conn, params=params)
    except Exception:
        data = pd.DataFrame()
    conn.close()

    return data


def delete_records_by_ids(table_name, record_ids):
    """Deletes a selected set of records from an internal table."""

    if not record_ids:
        return

    placeholders = ",".join("?" for _ in record_ids)
    run_query(f"DELETE FROM {table_name} WHERE id IN ({placeholders})", tuple(record_ids))


def render_credit_debit_notes(register_type):
    st.subheader(f"{register_type} Credit / Debit Notes")
    with st.form(f"{register_type.lower()}_notes_form"):
        note_type = st.selectbox("Note Type", ["Credit Note", "Debit Note"])
        party_name = st.text_input("Party Name")
        note_gstin = st.text_input("GSTIN").upper()
        note_number = st.text_input("Note Number *")
        note_date = st.date_input("Note Date", value=date.today())
        note_taxable = st.number_input("Taxable Value", min_value=0.0, step=100.0)
        note_cgst = st.number_input("CGST", min_value=0.0, step=10.0)
        note_sgst = st.number_input("SGST", min_value=0.0, step=10.0)
        note_igst = st.number_input("IGST", min_value=0.0, step=10.0)
        note_reason = st.text_input("Reason")
        save_note = st.form_submit_button("💾 Save Note", type="primary")
    if save_note:
        if not note_number.strip():
            st.error("Note Number is required.")
        else:
            run_query("""
                INSERT INTO gst_credit_debit_notes (
                    register_type, note_type, party_name, gstin, note_number,
                    note_date, taxable_value, cgst, sgst, igst, total_amount,
                    reason, created_date
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (register_type, note_type, party_name.strip(), note_gstin.strip(),
                   note_number.strip(), str(note_date), note_taxable, note_cgst,
                   note_sgst, note_igst, note_taxable + note_cgst + note_sgst + note_igst,
                   note_reason.strip(), now_stamp()))
            st.success(f"{note_type} saved for {register_type}.")
            st.rerun()
    notes = read_query(
        "SELECT id, note_type, party_name, gstin, note_number, note_date, taxable_value, total_amount, reason "
        "FROM gst_credit_debit_notes WHERE register_type = ? ORDER BY id DESC", (register_type,)
    )
    if notes.empty:
        st.info("No credit or debit notes recorded yet.")
    else:
        st.dataframe(notes, use_container_width=True, hide_index=True)
        select_all_notes = st.checkbox("Select all notes", key=f"select_all_notes_{register_type}")
        delete_ids = notes["id"].tolist() if select_all_notes else st.multiselect(
            "Select note IDs to delete", notes["id"].tolist(), key=f"delete_notes_{register_type}"
        )
        if st.button("🗑️ Delete Selected Notes", key=f"delete_notes_button_{register_type}"):
            delete_records_by_ids("gst_credit_debit_notes", delete_ids)
            st.success("Selected notes deleted.")
            st.rerun()


def get_duplicate_import_ids(dataframe, table_name, return_period=None):
    """Returns existing record IDs whose GSTIN + invoice number is in an upload."""

    if dataframe.empty or not {"gstin", "invoice_number"}.issubset(dataframe.columns):
        return []

    incoming_keys = {
        f"{str(row.gstin).strip().upper()}|{str(row.invoice_number).strip().upper()}"
        for row in dataframe[["gstin", "invoice_number"]].itertuples(index=False)
        if str(row.gstin).strip() and str(row.invoice_number).strip()
    }
    if not incoming_keys:
        return []

    query = f"SELECT id, gstin, invoice_number{', return_period' if return_period is not None else ''} FROM {table_name}"
    params = ()
    if return_period is not None:
        query += " WHERE return_period = ?"
        params = (return_period,)

    existing = read_query(query, params)
    if existing.empty:
        return []

    existing_keys = (
        existing["gstin"].fillna("").astype(str).str.strip().str.upper() + "|" +
        existing["invoice_number"].fillna("").astype(str).str.strip().str.upper()
    )
    return existing.loc[existing_keys.isin(incoming_keys), "id"].astype(int).tolist()


def overwrite_import_records(record_ids, table_name):
    """Deletes only the existing records matched by the current upload."""

    if not record_ids:
        return

    placeholders = ",".join("?" for _ in record_ids)
    run_query(f"DELETE FROM {table_name} WHERE id IN ({placeholders})", tuple(record_ids))


def get_period_mismatch_rows(dataframe, return_period):
    """Returns upload rows whose invoice date is outside the selected period."""

    if dataframe.empty or "invoice_date" not in dataframe.columns:
        return pd.DataFrame()

    selected_period = str(return_period).strip()
    parsed_dates = pd.to_datetime(
        dataframe["invoice_date"], errors="coerce", format="%Y-%m-%d"
    )
    invoice_periods = parsed_dates.dt.strftime("%Y-%m")
    mismatch_mask = invoice_periods.isna() | (invoice_periods != selected_period)
    return dataframe.loc[mismatch_mask].copy()


def now_stamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_date_input_value(value, fallback=None):
    """Returns a valid Python date for Streamlit date_input widgets."""

    fallback = fallback or date.today()
    parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)
    return fallback if pd.isna(parsed) else parsed.date()


def to_period(invoice_date):
    """Converts any date-like value to a 'YYYY-MM' return period string."""

    try:
        return pd.to_datetime(invoice_date).strftime("%Y-%m")
    except Exception:
        return ""


def normalize_header_name(column_name):
    """Normalizes messy CSV/Excel headers into the same field names used by the app."""

    if pd.isna(column_name):
        return ""

    text = str(column_name).strip().lower()
    text = text.replace("/", "_")
    text = text.replace("-", "_")
    text = text.replace("(", "_").replace(")", "_")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


HEADER_ALIASES = {
    "supplier_name": {
        "supplier_name", "supplier", "suppliername", "vendor", "vendor_name",
        "vendorname", "party_name", "party", "trade_name", "legal_name",
        "trade_legal_name", "seller_name", "seller", "particulars"
    },
    "customer_name": {
        "customer_name", "customer", "customername", "buyer_name", "buyer",
        "buyername", "recipient_name", "receiver_name", "recipient", "receiver"
    },
    "gstin": {
        "gstin", "gstin_no", "gstin_number", "gstinno", "gstin_no_",
        "gstin_uin", "gstin_uin_of_recipient", "gstin_of_recipient",
        "gstin_of_supplier", "supplier_gstin", "recipient_gstin",
        "customer_gstin", "buyer_gstin", "vendor_gstin", "party_gstin"
    },
    "invoice_number": {
        "invoice_number", "invoice_no", "invoiceno", "bill_no",
        "bill_number", "billno", "voucher_no", "voucher_number",
        "voucher", "tax_invoice_no", "tax_invoice_number", "doc_no",
        "document_no", "document_number", "inv_no", "inv_number",
        "supplier_invoice_no"
    },
    "invoice_date": {
        "invoice_date", "invoice_dt", "invoicedate", "bill_date",
        "billdate", "date", "date_of_invoice", "invoice_date_value", "doc_date",
        "document_date", "inv_date"
    },
    "taxable_value": {
        "taxable_value", "taxable_amount", "taxableamt", "taxable_amount_value",
        "taxable", "net_amount", "total_taxable_value", "taxable_value_", "value"
    },
    "gst_rate": {
        "gst_rate", "rate", "gst_rate_percent", "rate_percent", "tax_rate",
        "rate_of_tax", "gst_percent", "tax_percent"
    },
    "transaction_type": {
        "transaction_type", "transaction", "nature", "tax_type", "gst_type",
        "supply_type", "invoice_type", "place_of_supply_type"
    },
    "cgst": {"cgst", "cgst_amount", "cgst_amt", "cgst_tax", "central_tax", "central_tax_"},
    "sgst": {
        "sgst", "sgst_amount", "sgst_amt", "sgst_tax", "state_tax", "state_tax_",
        "sgst_ut", "ut_tax", "state_ut_tax"
    },
    "igst": {"igst", "igst_amount", "igst_amt", "igst_tax", "integrated_tax", "integrated_tax_"},
    "total_amount": {
        "total_amount", "grand_total", "invoice_total", "total_amt",
        "amount_total", "value_after_tax", "invoice_value", "total_invoice_value",
        "total_value"
    },
    "itc_eligible": {
        "itc_eligible", "itceligible", "eligible_for_itc", "eligible",
        "itc_availability", "eligibility_of_itc"
    },
    "return_period": {"return_period", "return_month", "gst_period", "tax_period", "filing_period"},
}


# Ordered fuzzy fallback rules used when a column doesn't exactly match one of
# the aliases above (e.g. real-world GST portal exports use many variations).
# Each rule is (canonical_field, all_of, any_of, none_of). A column matches a
# rule if every token in all_of is present, AND at least one token in any_of
# is present (skip this check if any_of is empty), AND no token in none_of is
# present. Rules are checked in order and the first match wins, so more
# specific rules (e.g. invoice_date) are listed before generic ones.
FUZZY_FIELD_RULES = [
    ("gstin", [], ["gstin"], []),
    ("invoice_date", [], ["invoice_date", "bill_date", "doc_date", "document_date"], []),
    ("invoice_number", [], ["invoice_no", "invoiceno", "bill_no", "billno", "voucher_no", "doc_no", "document_no"], ["date"]),
    ("taxable_value", ["taxable"], [], []),
    ("cgst", [], ["cgst", "central_tax"], []),
    ("sgst", [], ["sgst", "state_tax", "ut_tax"], []),
    ("igst", [], ["igst", "integrated_tax"], []),
    ("total_amount", [], ["invoice_value", "total_value", "grand_total", "invoice_total", "total_amount"], []),
    ("gst_rate", [], ["rate"], []),
    ("customer_name", [], ["recipient_name", "receiver_name", "customer_name", "buyer_name"], []),
    ("supplier_name", [], ["supplier_name", "supplier", "vendor_name", "vendor", "party_name", "trade_name", "seller_name"], []),
    ("itc_eligible", [], ["itc_eligible", "itc_availability"], []),
    ("return_period", [], ["return_period", "tax_period", "filing_period"], []),
]

# Canonical field list used to drive the manual column-mapping UI.
CANONICAL_FIELDS = [
    "supplier_name", "customer_name", "gstin", "invoice_number", "invoice_date",
    "taxable_value", "gst_rate", "transaction_type", "cgst", "sgst", "igst",
    "total_amount", "itc_eligible", "return_period",
]

FIELD_LABELS = {
    "supplier_name": "Supplier / Party Name",
    "customer_name": "Customer Name",
    "gstin": "GSTIN",
    "invoice_number": "Invoice / Bill Number",
    "invoice_date": "Invoice Date",
    "taxable_value": "Taxable Value",
    "gst_rate": "GST Rate (%)",
    "transaction_type": "Transaction Type",
    "cgst": "CGST",
    "sgst": "SGST",
    "igst": "IGST",
    "total_amount": "Total Invoice Value",
    "itc_eligible": "ITC Eligible",
    "return_period": "Return Period",
}


def fuzzy_guess_field(normalized):
    """Best-effort guess of which canonical field a messy column header
    represents, used only when there is no exact alias match."""

    if not normalized:
        return None

    for canonical, all_of, any_of, none_of in FUZZY_FIELD_RULES:
        if all_of and not all(token in normalized for token in all_of):
            continue
        if any_of and not any(token in normalized for token in any_of):
            continue
        if any(token in normalized for token in none_of):
            continue
        return canonical

    return None


def canonicalize_uploaded_header(column_name):
    normalized = normalize_header_name(column_name)

    for canonical, aliases in HEADER_ALIASES.items():
        if normalized in aliases or normalized == canonical:
            return canonical

    guess = fuzzy_guess_field(normalized)
    return guess or normalized


def is_column_effectively_blank(series):
    """True if a column is missing/blank in a way that isna() alone would
    miss — e.g. a column padded with empty strings by normalize_uploaded_dataframe
    when the source file simply doesn't have that field."""

    if series is None:
        return True

    cleaned = series.astype(str).str.strip().str.lower()
    return cleaned.isin(["", "nan", "none", "nat"]).all()


def get_missing_or_blank_columns(dataframe, required_columns):
    missing = []
    for column in required_columns:
        if column not in dataframe.columns or is_column_effectively_blank(dataframe[column]):
            missing.append(column)
    return missing


def list_excel_sheets(file_obj):
    """Returns the sheet names of an uploaded Excel workbook, or an empty
    list if the file can't be read as Excel (e.g. it's a CSV)."""

    try:
        file_obj.seek(0)
        return pd.ExcelFile(file_obj).sheet_names
    except Exception:
        return []


def detect_header_row(file_obj, required_keywords, max_rows=30, sheet_name=None):
    """Return the row most likely to be the GST header row.

    A valid header row needs multiple specific GST fields, not just a single
    generic term like 'name' or 'date'. This prevents raw data rows from being
    treated as column headers.
    """

    file_obj.seek(0)

    try:
        if str(file_obj.name).lower().endswith(".csv"):
            raw = pd.read_csv(file_obj, header=None, nrows=max_rows)
        else:
            raw = pd.read_excel(file_obj, header=None, nrows=max_rows, sheet_name=sheet_name if sheet_name is not None else 0)
    except Exception:
        return 0

    if raw.empty:
        return 0

    normalized_candidates = {normalize_header_name(item) for item in required_keywords}
    best_score = -1
    best_row = 0

    for idx, row in raw.iterrows():
        row_values = [normalize_header_name(cell) for cell in row if pd.notna(cell)]
        if not row_values:
            continue

        matched = set()
        for value in row_values:
            for candidate in normalized_candidates:
                if not candidate:
                    continue
                if value == candidate or value in candidate or candidate in value:
                    matched.add(candidate)
                    break

        # Prefer exact GST field names; avoid generic matches from data rows.
        score = len(matched) * 10
        if {"gstin", "invoice_number", "invoice_no", "invoice_date", "taxable_value", "cgst", "sgst", "igst"} & matched:
            score += 5

        if score > best_score:
            best_score = score
            best_row = idx

    # Require a strong header match, not just a single generic column name.
    return best_row if best_score >= 20 else 0


UPLOAD_HEADER_KEYWORDS = [
    "gstin", "invoice_number", "invoice_no", "bill_no", "voucher_no",
    "supplier_name", "supplier", "recipient_name", "customer_name",
    "particulars", "supplier_invoice_no", "date",
    "taxable_value", "taxable_amount", "cgst", "sgst", "igst",
    "total_amount", "invoice_value", "value",
]


def load_uploaded_file(file_obj, sheet_name=None, header_row_override=None):
    """Read an uploaded CSV or Excel using an auto-detected (or manually
    overridden) header row, optionally from a specific sheet. Returns
    (dataframe, header_row_used)."""

    file_name = str(file_obj.name).lower()
    is_csv = file_name.endswith(".csv")

    if header_row_override is not None:
        header_row = header_row_override
    else:
        header_row = detect_header_row(
            file_obj, UPLOAD_HEADER_KEYWORDS, max_rows=30,
            sheet_name=None if is_csv else sheet_name
        )

    file_obj.seek(0)

    if is_csv:
        df = pd.read_csv(file_obj, header=header_row)
    else:
        df = pd.read_excel(file_obj, header=header_row, sheet_name=sheet_name if sheet_name is not None else 0)

    return df, header_row


def normalize_uploaded_dataframe(dataframe, required_columns=None, allow_optional=False):
    """Standardizes imported CSV/Excel column headers to the app's internal schema."""

    df = dataframe.copy()
    if df.empty:
        return df

    df.columns = [canonicalize_uploaded_header(col) for col in df.columns]

    for canonical in list(df.columns):
        duplicate_cols = [col for col in df.columns if col == canonical]
        if len(duplicate_cols) > 1:
            df[canonical] = df[duplicate_cols].bfill(axis=1).iloc[:, 0]
            for duplicate in duplicate_cols[1:]:
                df = df.drop(columns=[duplicate])

    required_columns = required_columns or []
    for required in required_columns:
        if required not in df.columns:
            df[required] = ""

    if "invoice_date" in df.columns:
        # dayfirst=True because Indian GST documents and portal exports use
        # DD-MM-YYYY; without this, dates like 01-04-2026 (1 Apr) get
        # silently misread as 4 Jan by pandas' default US-style parsing.
        df["invoice_date"] = pd.to_datetime(df["invoice_date"], errors="coerce", dayfirst=True).dt.strftime("%Y-%m-%d")

    for numeric_col in ["taxable_value", "gst_rate", "cgst", "sgst", "igst", "total_amount"]:
        if numeric_col in df.columns:
            df[numeric_col] = pd.to_numeric(df[numeric_col], errors="coerce").fillna(0)

    if "supplier_name" in df.columns:
        df["supplier_name"] = df["supplier_name"].fillna("").astype(str).str.strip()

    if "customer_name" in df.columns:
        df["customer_name"] = df["customer_name"].fillna("").astype(str).str.strip()

    if "gstin" in df.columns:
        df["gstin"] = df["gstin"].fillna("").astype(str).str.upper().str.strip()

    if "invoice_number" in df.columns:
        df["invoice_number"] = df["invoice_number"].fillna("").astype(str).str.strip()

    if "transaction_type" not in df.columns:
        if "igst" in df.columns:
            df["transaction_type"] = df["igst"].apply(lambda value: "Interstate" if float(value or 0) > 0 else "Intrastate")
        else:
            df["transaction_type"] = "Intrastate"

    if "total_amount" not in df.columns and {"taxable_value", "cgst", "sgst", "igst"}.issubset(set(df.columns)):
        df["total_amount"] = df["taxable_value"] + df["cgst"] + df["sgst"] + df["igst"]

    if "gst_rate" not in df.columns and {"taxable_value", "cgst", "sgst", "igst"}.issubset(set(df.columns)):
        total_tax = df["cgst"] + df["sgst"] + df["igst"]
        df["gst_rate"] = ((total_tax / df["taxable_value"]) * 100).where(df["taxable_value"] != 0, 0)

    return df


# ============================================================
# EXCEL / CSV IMPORT WIZARD (shared by Sales / Purchase / 2B / GSTR-1 uploads)
# ============================================================

def excel_csv_upload_and_map(uploader_label, file_key, required_fields, help_text="",
                              party_field=None, party_label=None):
    """Full upload -> sheet select -> header row -> column mapping -> preview
    pipeline shared across every register import. Returns a cleaned, fully
    canonicalized dataframe ready for a bulk_insert_* function, or None if the
    file hasn't been uploaded yet or isn't ready to import.

    This exists specifically so that an uploaded file's columns are always
    explicitly confirmed by the user (with a sensible auto-guess pre-filled)
    instead of silently relying on automatic header matching, which is what
    previously let files with unrecognized column names import zero rows
    without any explanation.

    party_field / party_label fix the real-world import bug where a Sales
    Register's "Customer / Buyer" column and a Purchase Register's
    "Supplier / Vendor" column are, underneath, the SAME required field for
    that register (the app always stores it in one physical column). Left
    to the generic alias matcher, a "Customer Name" header on a Sales
    Register upload got auto-guessed as the separate "customer_name" field
    instead of the "supplier_name" field the Sales table actually requires
    — so the required-field check failed and the import was silently
    rejected even though a party-name column clearly existed. When
    party_field is given, both "Customer ..." and "Supplier ..." style
    headers are coerced onto that one field, and only a single, correctly
    labelled option is shown in the mapping dropdown to avoid re-confusing
    the person doing the mapping.
    """

    if help_text:
        st.caption(help_text)

    uploaded = st.file_uploader(
        uploader_label, type=["csv", "xlsx", "xls"], key=file_key
    )

    if uploaded is None:
        return None

    is_excel = not uploaded.name.lower().endswith(".csv")
    sheet_name = None

    if is_excel:
        sheets = list_excel_sheets(uploaded)
        if len(sheets) > 1:
            sheet_name = st.selectbox(
                "This workbook has multiple sheets — pick the one with your data",
                sheets, key=f"{file_key}_sheet"
            )
        elif sheets:
            sheet_name = sheets[0]

    auto_header_row = detect_header_row(
        uploaded, UPLOAD_HEADER_KEYWORDS, max_rows=30,
        sheet_name=None if not is_excel else sheet_name
    )

    with st.expander(f"⚙️ Header row: auto-detected as row {auto_header_row + 1}. Click to override.", expanded=False):
        manual_override = st.checkbox(
            "The header row above is wrong — let me set it manually",
            key=f"{file_key}_override"
        )
        header_row = auto_header_row
        if manual_override:
            header_row = st.number_input(
                "Header row number (1 = first row of the file/sheet)",
                min_value=1, max_value=30, value=auto_header_row + 1, step=1,
                key=f"{file_key}_header_num"
            ) - 1

    try:
        raw_df, _ = load_uploaded_file(uploaded, sheet_name=sheet_name, header_row_override=header_row)
    except Exception as error:
        st.error(f"Could not read this file: {error}")
        return None

    raw_df = raw_df.dropna(how="all")
    if raw_df.empty:
        st.warning("No data rows were found using this header row. Try overriding it above.")
        return None

    st.markdown('<span class="gst-step-badge">STEP 1</span> Raw preview (first 5 rows as read from the file)', unsafe_allow_html=True)
    st.dataframe(raw_df.head(5), use_container_width=True, hide_index=True)

    st.markdown('<span class="gst-step-badge">STEP 2</span> Confirm column mapping', unsafe_allow_html=True)
    st.caption("Each uploaded column has been auto-matched to a field below. Fix any that look wrong, or set unused columns to 'Ignore'.")

    # If this register has one "party" field (customer OR supplier, but the
    # table only has one physical column for it), collapse the two generic
    # canonical fields into that single, correctly labelled option so the
    # auto-guess and the dropdown both point at the field that's actually
    # required — see the docstring above for why this matters.
    other_party_field = None
    local_fields = list(CANONICAL_FIELDS)
    local_labels = dict(FIELD_LABELS)
    if party_field in ("supplier_name", "customer_name"):
        other_party_field = "customer_name" if party_field == "supplier_name" else "supplier_name"
        local_fields = [field for field in local_fields if field != other_party_field]
        if party_label:
            local_labels[party_field] = party_label

    ignore_option = "— Ignore this column —"
    options = [ignore_option] + [f"{field} ({local_labels.get(field, field)})" for field in local_fields]
    option_by_field = {f"{field} ({local_labels.get(field, field)})": field for field in local_fields}

    mapping_cols = st.columns(2)
    mapping = {}
    for idx, col in enumerate(raw_df.columns):
        guess = canonicalize_uploaded_header(col)
        if other_party_field and guess == other_party_field:
            guess = party_field
        default_label = next((label for label, field in option_by_field.items() if field == guess), ignore_option) if guess in local_fields else ignore_option

        with mapping_cols[idx % 2]:
            chosen_label = st.selectbox(
                f"'{col}'  →",
                options,
                index=options.index(default_label),
                key=f"{file_key}_map_{idx}"
            )
        mapping[col] = option_by_field.get(chosen_label)

    mapped_df = pd.DataFrame(index=raw_df.index)
    for col, target_field in mapping.items():
        if target_field is None:
            continue
        if target_field not in mapped_df.columns:
            mapped_df[target_field] = raw_df[col]
        else:
            existing = mapped_df[target_field]
            blank_mask = existing.isna() | (existing.astype(str).str.strip() == "")
            mapped_df[target_field] = existing.mask(blank_mask, raw_df[col])

    clean_df = normalize_uploaded_dataframe(mapped_df, required_columns=required_fields, allow_optional=True)

    missing = get_missing_or_blank_columns(clean_df, required_fields)
    if missing:
        missing_labels = ", ".join(FIELD_LABELS.get(field, field) for field in missing)
        st.error(f"These required fields still aren't mapped (or are empty for every row): {missing_labels}")
        return None

    st.markdown('<span class="gst-step-badge">STEP 3</span> Final preview — ready to import', unsafe_allow_html=True)
    st.dataframe(clean_df.head(10), use_container_width=True, hide_index=True)
    st.caption(f"{len(clean_df)} row(s) will be checked and imported.")

    return clean_df


def show_import_results(count, skipped, imported_noun="record(s)"):
    """Displays a consistent success/skip summary after a bulk_insert_* call."""

    if count:
        st.success(f"✅ Imported {count} {imported_noun}.")
    else:
        st.warning("No rows were imported.")

    if skipped:
        st.warning(f"⚠️ Skipped {len(skipped)} row(s) that were missing required data.")
        with st.expander("View skipped rows and reasons", expanded=True):
            st.dataframe(pd.DataFrame(skipped), use_container_width=True, hide_index=True)


def bulk_insert_sales(dataframe):
    """Import a sales register file whose headers match the app's invoice register."""

    df = normalize_uploaded_dataframe(
        dataframe,
        required_columns=["supplier_name", "invoice_number", "invoice_date", "taxable_value"],
        allow_optional=True
    )

    if df.empty:
        return 0, []

    df = df.dropna(how="all").copy()

    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    now = now_stamp()

    inserted = 0
    skipped = []
    for position, (_, row) in enumerate(df.iterrows(), start=2):
        if row.isna().all():
            continue

        supplier = str(row.get("supplier_name", "") if pd.notna(row.get("supplier_name", "")) else "").strip()
        gstin = str(row.get("gstin", "") if pd.notna(row.get("gstin", "")) else "").strip().upper()
        invoice_number = str(row.get("invoice_number", "") if pd.notna(row.get("invoice_number", "")) else "").strip()
        invoice_date = str(row.get("invoice_date", "") if pd.notna(row.get("invoice_date", "")) else "").strip()
        taxable_value = float(pd.to_numeric(row.get("taxable_value", 0), errors="coerce") or 0)
        gst_rate = float(pd.to_numeric(row.get("gst_rate", 0), errors="coerce") or 0)
        transaction_type = str(row.get("transaction_type", "Intrastate") if pd.notna(row.get("transaction_type", "Intrastate")) else "Intrastate").strip() or "Intrastate"
        cgst = float(pd.to_numeric(row.get("cgst", 0), errors="coerce") or 0)
        sgst = float(pd.to_numeric(row.get("sgst", 0), errors="coerce") or 0)
        igst = float(pd.to_numeric(row.get("igst", 0), errors="coerce") or 0)
        total_amount = float(pd.to_numeric(row.get("total_amount", taxable_value + cgst + sgst + igst), errors="coerce") or (taxable_value + cgst + sgst + igst))

        reasons = []
        if not supplier:
            reasons.append("missing supplier/party name")
        if not invoice_number:
            reasons.append("missing invoice number")

        supply_category = str(row.get("supply_category", "B2B") or "B2B").strip()
        if not gstin:
            supply_category = "B2C"

        if reasons:
            skipped.append({"Row": position, "Reason": ", ".join(reasons)})
            continue

        cursor.execute("""
            INSERT INTO gst_invoices (
                supplier_name, gstin, invoice_number, invoice_date, taxable_value,
                gst_rate, transaction_type, cgst, sgst, igst, total_amount,
                supply_category, supply_type, created_date, modified_date
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            supplier, gstin, invoice_number, invoice_date, taxable_value,
            gst_rate, transaction_type, cgst, sgst, igst, total_amount,
            supply_category, str(row.get("supply_type", "Goods") or "Goods"), now, now
        ))
        inserted += 1

    conn.commit()
    conn.close()
    return inserted, skipped


def bulk_insert_purchase(dataframe, return_period=None):
    """Import a purchase register file whose headers match the app's purchase register."""

    df = normalize_uploaded_dataframe(
        dataframe,
        required_columns=["supplier_name", "invoice_number", "invoice_date", "taxable_value"],
        allow_optional=True
    )

    if df.empty:
        return 0, []

    df = df.dropna(how="all").copy()

    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    now = now_stamp()

    inserted = 0
    skipped = []
    for position, (_, row) in enumerate(df.iterrows(), start=2):
        if row.isna().all():
            continue

        supplier = str(row.get("supplier_name", "") if pd.notna(row.get("supplier_name", "")) else "").strip()
        gstin = str(row.get("gstin", "") if pd.notna(row.get("gstin", "")) else "").strip().upper()
        invoice_number = str(row.get("invoice_number", "") if pd.notna(row.get("invoice_number", "")) else "").strip()
        invoice_date = str(row.get("invoice_date", "") if pd.notna(row.get("invoice_date", "")) else "").strip()
        taxable_value = float(pd.to_numeric(row.get("taxable_value", 0), errors="coerce") or 0)
        gst_rate = float(pd.to_numeric(row.get("gst_rate", 0), errors="coerce") or 0)
        transaction_type = str(row.get("transaction_type", "Intrastate") if pd.notna(row.get("transaction_type", "Intrastate")) else "Intrastate").strip() or "Intrastate"
        cgst = float(pd.to_numeric(row.get("cgst", 0), errors="coerce") or 0)
        sgst = float(pd.to_numeric(row.get("sgst", 0), errors="coerce") or 0)
        igst = float(pd.to_numeric(row.get("igst", 0), errors="coerce") or 0)
        total_amount = float(pd.to_numeric(row.get("total_amount", taxable_value + cgst + sgst + igst), errors="coerce") or (taxable_value + cgst + sgst + igst))
        itc_eligible = str(row.get("itc_eligible", "Yes") if pd.notna(row.get("itc_eligible", "Yes")) else "Yes").strip() or "Yes"

        reasons = []
        if not supplier:
            reasons.append("missing supplier/party name")
        if not invoice_number:
            reasons.append("missing invoice number")

        supply_category = str(row.get("supply_category", "B2B") or "B2B").strip()
        if not gstin:
            supply_category = "B2C"

        if reasons:
            skipped.append({"Row": position, "Reason": ", ".join(reasons)})
            continue

        cursor.execute("""
            INSERT INTO purchase_invoices (
                supplier_name, gstin, invoice_number, invoice_date, taxable_value,
                gst_rate, transaction_type, cgst, sgst, igst, total_amount,
                itc_eligible, supply_category, supply_type, return_period, created_date, modified_date
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            supplier, gstin, invoice_number, invoice_date, taxable_value,
            gst_rate, transaction_type, cgst, sgst, igst, total_amount,
            itc_eligible, supply_category, str(row.get("supply_type", "Goods") or "Goods"),
            return_period or to_period(invoice_date), now, now
        ))
        inserted += 1

    conn.commit()
    conn.close()
    return inserted, skipped


def bulk_insert_2b(dataframe, return_period):
    """Used by the CSV / Excel uploader for GSTR-2B downloads from the portal."""

    df = normalize_uploaded_dataframe(
        dataframe,
        required_columns=["supplier_name", "gstin", "invoice_number", "invoice_date", "taxable_value"],
        allow_optional=True
    )

    now = now_stamp()
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    inserted = 0
    skipped = []
    for position, (_, row) in enumerate(df.iterrows(), start=2):
        supplier_name = str(row.get("supplier_name", "")).strip()
        gstin = str(row.get("gstin", "")).strip().upper()
        invoice_number = str(row.get("invoice_number", "")).strip()
        invoice_date = str(row.get("invoice_date", "")).strip()
        taxable_value = float(row.get("taxable_value", 0) or 0)
        cgst = float(row.get("cgst", 0) or 0)
        sgst = float(row.get("sgst", 0) or 0)
        igst = float(row.get("igst", 0) or 0)

        reasons = []
        if not gstin:
            reasons.append("missing GSTIN")
        if not invoice_number:
            reasons.append("missing invoice number")

        if reasons:
            skipped.append({"Row": position, "Reason": ", ".join(reasons)})
            continue

        cursor.execute("""
            INSERT INTO gstr_2b (
                supplier_name, gstin, invoice_number, invoice_date, taxable_value,
                cgst, sgst, igst, return_period, itc_eligible, itc_remark, created_date
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Eligible', '', ?)
        """, (
            supplier_name, gstin, invoice_number, invoice_date, taxable_value,
            cgst, sgst, igst, return_period, now
        ))
        inserted += 1

    conn.commit()
    conn.close()
    return inserted, skipped


# ============================================================
# GST CALCULATION
# ============================================================

def calculate_gst(taxable_value, gst_rate, transaction_type):

    gst_amount = taxable_value * gst_rate / 100

    if transaction_type == "Intrastate":
        cgst = gst_amount / 2
        sgst = gst_amount / 2
        igst = 0
    else:
        cgst = 0
        sgst = 0
        igst = gst_amount

    total_amount = taxable_value + cgst + sgst + igst

    return cgst, sgst, igst, total_amount


# ============================================================
# SALES INVOICES (GST_INVOICES)
# ============================================================

def load_data(from_date=None, to_date=None):

    query = "SELECT * FROM gst_invoices"
    params = ()

    if from_date and to_date:
        query += " WHERE date(invoice_date) BETWEEN date(?) AND date(?)"
        params = (str(from_date), str(to_date))

    query += " ORDER BY id DESC"

    data = read_query(query, params)

    if not data.empty:
        for column in ["created_date", "modified_date"]:
            if column not in data.columns:
                data[column] = ""

    return data


def insert_invoice(supplier_name, gstin, invoice_number, invoice_date, taxable_value,
                    gst_rate, transaction_type, cgst, sgst, igst, total_amount,
                    supply_category="B2B", supply_type="Goods"):

    now = now_stamp()

    run_query("""
        INSERT INTO gst_invoices (
            supplier_name, gstin, invoice_number, invoice_date, taxable_value,
            gst_rate, transaction_type, cgst, sgst, igst, total_amount,
            supply_category, supply_type, created_date, modified_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        supplier_name, gstin, invoice_number, str(invoice_date), taxable_value,
        gst_rate, transaction_type, cgst, sgst, igst, total_amount,
        supply_category, supply_type, now, now
    ))


def update_invoice(invoice_id, supplier_name, gstin, invoice_number, invoice_date,
                    taxable_value, gst_rate, transaction_type, cgst, sgst, igst, total_amount,
                    supply_category="B2B", supply_type="Goods"):

    now = now_stamp()

    run_query("""
        UPDATE gst_invoices SET
            supplier_name = ?, gstin = ?, invoice_number = ?, invoice_date = ?,
            taxable_value = ?, gst_rate = ?, transaction_type = ?, cgst = ?,
            sgst = ?, igst = ?, total_amount = ?, supply_category = ?, supply_type = ?, modified_date = ?
        WHERE id = ?
    """, (
        supplier_name, gstin, invoice_number, str(invoice_date), taxable_value,
        gst_rate, transaction_type, cgst, sgst, igst, total_amount,
        supply_category, supply_type, now, invoice_id
    ))


def delete_invoice(invoice_id):
    run_query("DELETE FROM gst_invoices WHERE id = ?", (invoice_id,))


# ============================================================
# PURCHASE INVOICES
# ============================================================

def load_purchase_data(from_date=None, to_date=None):

    query = "SELECT * FROM purchase_invoices"
    params = ()

    if from_date and to_date:
        query += " WHERE date(invoice_date) BETWEEN date(?) AND date(?)"
        params = (str(from_date), str(to_date))

    query += " ORDER BY id DESC"

    return read_query(query, params)


def insert_purchase(supplier_name, gstin, invoice_number, invoice_date, taxable_value,
                     gst_rate, transaction_type, cgst, sgst, igst, total_amount, itc_eligible,
                     supply_category="B2B", supply_type="Goods"):

    now = now_stamp()

    run_query("""
        INSERT INTO purchase_invoices (
            supplier_name, gstin, invoice_number, invoice_date, taxable_value,
            gst_rate, transaction_type, cgst, sgst, igst, total_amount,
            itc_eligible, supply_category, supply_type, return_period, created_date, modified_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        supplier_name, gstin, invoice_number, str(invoice_date), taxable_value,
        gst_rate, transaction_type, cgst, sgst, igst, total_amount, itc_eligible,
        supply_category, supply_type, to_period(invoice_date), now, now
    ))


def update_purchase(purchase_id, supplier_name, gstin, invoice_number, invoice_date,
                     taxable_value, gst_rate, transaction_type, cgst, sgst, igst,
                     total_amount, itc_eligible, supply_category="B2B", supply_type="Goods"):

    now = now_stamp()

    run_query("""
        UPDATE purchase_invoices SET
            supplier_name = ?, gstin = ?, invoice_number = ?, invoice_date = ?,
            taxable_value = ?, gst_rate = ?, transaction_type = ?, cgst = ?,
            sgst = ?, igst = ?, total_amount = ?, itc_eligible = ?,
            supply_category = ?, supply_type = ?, return_period = ?, modified_date = ?
        WHERE id = ?
    """, (
        supplier_name, gstin, invoice_number, str(invoice_date), taxable_value,
        gst_rate, transaction_type, cgst, sgst, igst, total_amount, itc_eligible,
        supply_category, supply_type, to_period(invoice_date), now, purchase_id
    ))


def delete_purchase(purchase_id):
    run_query("DELETE FROM purchase_invoices WHERE id = ?", (purchase_id,))


# ============================================================
# GSTR-2B
# ============================================================

def load_2b_data(return_period=None):

    if return_period:
        return read_query(
            "SELECT * FROM gstr_2b WHERE return_period = ? ORDER BY id DESC",
            (return_period,)
        )

    return read_query("SELECT * FROM gstr_2b ORDER BY id DESC")


def insert_2b_record(supplier_name, gstin, invoice_number, invoice_date,
                      taxable_value, cgst, sgst, igst, return_period):

    run_query("""
        INSERT INTO gstr_2b (
            supplier_name, gstin, invoice_number, invoice_date, taxable_value,
            cgst, sgst, igst, return_period, itc_eligible, itc_remark, created_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'Eligible', '', ?)
    """, (
        supplier_name, gstin, invoice_number, str(invoice_date), taxable_value,
        cgst, sgst, igst, return_period, now_stamp()
    ))


def delete_2b_record(record_id):
    run_query("DELETE FROM gstr_2b WHERE id = ?", (record_id,))


def update_2b_record(record_id, supplier_name, gstin, invoice_number, invoice_date,
                     taxable_value, cgst, sgst, igst, return_period,
                     itc_eligible="Eligible", itc_remark=""):
    run_query("""
        UPDATE gstr_2b SET
            supplier_name = ?, gstin = ?, invoice_number = ?, invoice_date = ?,
            taxable_value = ?, cgst = ?, sgst = ?, igst = ?, return_period = ?,
            itc_eligible = ?, itc_remark = ?
        WHERE id = ?
    """, (
        supplier_name, gstin, invoice_number, str(invoice_date), taxable_value,
        cgst, sgst, igst, return_period, itc_eligible, itc_remark, record_id
    ))


# ============================================================
# GSTR-1 (INVOICE-WISE, LIKE THE SALES REGISTER)
# ============================================================

def load_gstr1_data(return_period=None, from_date=None, to_date=None):

    query = "SELECT * FROM gstr1_invoices"
    conditions = []
    params = []

    if return_period:
        conditions.append("return_period = ?")
        params.append(return_period)

    if from_date and to_date:
        conditions.append("date(invoice_date) BETWEEN date(?) AND date(?)")
        params.extend([str(from_date), str(to_date)])

    if conditions:
        query += " WHERE " + " AND ".join(conditions)

    query += " ORDER BY id DESC"

    return read_query(query, tuple(params))


def insert_gstr1_invoice(customer_name, gstin, invoice_number, invoice_date, taxable_value,
                          gst_rate, transaction_type, cgst, sgst, igst, total_amount, return_period):

    now = now_stamp()

    run_query("""
        INSERT INTO gstr1_invoices (
            customer_name, gstin, invoice_number, invoice_date, taxable_value,
            gst_rate, transaction_type, cgst, sgst, igst, total_amount, return_period,
            created_date, modified_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        customer_name, gstin, invoice_number, str(invoice_date), taxable_value,
        gst_rate, transaction_type, cgst, sgst, igst, total_amount, return_period, now, now
    ))


def update_gstr1_invoice(record_id, customer_name, gstin, invoice_number, invoice_date,
                          taxable_value, gst_rate, transaction_type, cgst, sgst, igst,
                          total_amount, return_period):

    now = now_stamp()

    run_query("""
        UPDATE gstr1_invoices SET
            customer_name = ?, gstin = ?, invoice_number = ?, invoice_date = ?,
            taxable_value = ?, gst_rate = ?, transaction_type = ?, cgst = ?,
            sgst = ?, igst = ?, total_amount = ?, return_period = ?, modified_date = ?
        WHERE id = ?
    """, (
        customer_name, gstin, invoice_number, str(invoice_date), taxable_value,
        gst_rate, transaction_type, cgst, sgst, igst, total_amount, return_period,
        now, record_id
    ))


def delete_gstr1_invoice(record_id):
    run_query("DELETE FROM gstr1_invoices WHERE id = ?", (record_id,))


def bulk_insert_gstr1(dataframe, return_period):
    """Used by the CSV / Excel uploader for GSTR-1 invoice-wise downloads
    from the portal. Accepts either separate cgst/sgst/igst columns, or a
    single tax amount + transaction type — whichever the export provides."""

    df = normalize_uploaded_dataframe(
        dataframe,
        required_columns=["customer_name", "gstin", "invoice_number", "invoice_date", "taxable_value"],
        allow_optional=True
    )

    now = now_stamp()
    db_path = get_db_path()
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    inserted = 0
    skipped = []

    for position, (_, row) in enumerate(df.iterrows(), start=2):
        gstin = str(row.get("gstin", "")).strip().upper()
        invoice_number = str(row.get("invoice_number", "")).strip()

        reasons = []
        if not gstin:
            reasons.append("missing GSTIN")
        if not invoice_number:
            reasons.append("missing invoice number")

        if reasons:
            skipped.append({"Row": position, "Reason": ", ".join(reasons)})
            continue

        taxable_value = float(row.get("taxable_value", 0) or 0)
        cgst = float(row.get("cgst", 0) or 0)
        sgst = float(row.get("sgst", 0) or 0)
        igst = float(row.get("igst", 0) or 0)
        total_tax = cgst + sgst + igst

        gst_rate = round((total_tax / taxable_value) * 100, 2) if taxable_value else 0
        transaction_type = "Interstate" if igst > 0 else "Intrastate"
        total_amount = taxable_value + total_tax

        cursor.execute("""
            INSERT INTO gstr1_invoices (
                customer_name, gstin, invoice_number, invoice_date, taxable_value,
                gst_rate, transaction_type, cgst, sgst, igst, total_amount, return_period,
                created_date, modified_date
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            str(row.get("customer_name", "")),
            gstin,
            invoice_number,
            str(row.get("invoice_date", "")),
            taxable_value, gst_rate, transaction_type, cgst, sgst, igst, total_amount,
            return_period, now, now
        ))
        inserted += 1

    conn.commit()
    conn.close()

    return inserted, skipped


# ============================================================
# GSTR-1 SUMMARY
# ============================================================

def load_gstr1_summary():
    return read_query("SELECT * FROM gstr1_summary ORDER BY return_period DESC")


def save_gstr1(return_period, total_taxable_value, total_tax, invoice_count, filed_date):

    now = now_stamp()

    run_query("""
        INSERT INTO gstr1_summary (
            return_period, total_taxable_value, total_tax, invoice_count,
            filed_date, created_date, modified_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(return_period) DO UPDATE SET
            total_taxable_value = excluded.total_taxable_value,
            total_tax = excluded.total_tax,
            invoice_count = excluded.invoice_count,
            filed_date = excluded.filed_date,
            modified_date = excluded.modified_date
    """, (return_period, total_taxable_value, total_tax, invoice_count, str(filed_date), now, now))


def delete_gstr1(record_id):
    run_query("DELETE FROM gstr1_summary WHERE id = ?", (record_id,))


# ============================================================
# GSTR-3B SUMMARY
# ============================================================

def load_gstr3b_data():
    return read_query("SELECT * FROM gstr3b_summary ORDER BY return_period DESC")


def save_gstr3b(return_period, outward_taxable_value, output_tax, itc_claimed,
                tax_paid_cash, late_fee, filed_date, itc_reversed=0):

    now = now_stamp()

    run_query("""
        INSERT INTO gstr3b_summary (
            return_period, outward_taxable_value, output_tax, itc_claimed,
            itc_reversed, tax_paid_cash, late_fee, filed_date, created_date, modified_date
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(return_period) DO UPDATE SET
            outward_taxable_value = excluded.outward_taxable_value,
            output_tax = excluded.output_tax,
            itc_claimed = excluded.itc_claimed,
            itc_reversed = excluded.itc_reversed,
            tax_paid_cash = excluded.tax_paid_cash,
            late_fee = excluded.late_fee,
            filed_date = excluded.filed_date,
            modified_date = excluded.modified_date
        """, (return_period, outward_taxable_value, output_tax, itc_claimed,
            itc_reversed, tax_paid_cash, late_fee, str(filed_date), now, now))


def update_gstr3b(record_id, return_period, outward_taxable_value, output_tax,
                  itc_claimed, itc_reversed, tax_paid_cash, late_fee, filed_date):
    """Updates one GSTR-3B row, allowing its return period to be corrected."""

    run_query("""
        UPDATE gstr3b_summary
        SET return_period = ?, outward_taxable_value = ?, output_tax = ?,
            itc_claimed = ?, itc_reversed = ?, tax_paid_cash = ?, late_fee = ?,
            filed_date = ?, modified_date = ?
        WHERE id = ?
    """, (return_period, outward_taxable_value, output_tax, itc_claimed,
           itc_reversed, tax_paid_cash, late_fee, str(filed_date), now_stamp(),
           record_id))


def delete_gstr3b(record_id):
    run_query("DELETE FROM gstr3b_summary WHERE id = ?", (record_id,))


def bulk_insert_gstr3b(dataframe):
    """Used by the CSV / Excel uploader for GSTR-3B. Expects one row per
    return period; each row is inserted or, if that period already exists,
    updated (same as the manual save form)."""

    updated = 0

    for _, row in dataframe.iterrows():
        return_period = str(row.get("return_period", "")).strip()

        if not return_period:
            continue

        save_gstr3b(
            return_period,
            float(row.get("outward_taxable_value", 0) or 0),
            float(row.get("output_tax", 0) or 0),
            float(row.get("itc_claimed", 0) or 0),
            float(row.get("tax_paid_cash", 0) or 0),
            float(row.get("late_fee", 0) or 0),
            row.get("filed_date", ""),
            float(row.get("itc_reversed", 0) or 0)
        )
        updated += 1

    return updated


# ============================================================
# COMPANY PROFILE
# ============================================================

def load_company_profile():

    data = read_query("SELECT * FROM company_profile WHERE id = 1")

    if data.empty:
        return {"company_name": "", "gstin": "", "state": "", "financial_year": ""}

    return data.iloc[0].to_dict()


def save_company_profile(company_name, gstin, state, financial_year):

    run_query("""
        UPDATE company_profile
        SET company_name = ?, gstin = ?, state = ?, financial_year = ?
        WHERE id = 1
    """, (company_name, gstin, state, financial_year))


# ============================================================
# BACKUP DATABASE
# ============================================================

def create_backup():

    db_path = get_db_path()

    if not db_path or not os.path.exists(db_path):
        return None

    backup_root = st.session_state.save_path or os.path.dirname(db_path)
    backup_folder = os.path.join(backup_root, "Backups")
    os.makedirs(backup_folder, exist_ok=True)

    company_tag = os.path.splitext(os.path.basename(db_path))[0]
    today = datetime.now().strftime("%d-%m-%Y")
    backup_prefix = f"{company_tag}_Backup_{today}_"

    # Keep one backup per database per calendar day, even when several
    # records are saved during that day.
    todays_backups = sorted(
        filename for filename in os.listdir(backup_folder)
        if filename.startswith(backup_prefix) and filename.lower().endswith(".db")
    )
    if todays_backups:
        return os.path.join(backup_folder, todays_backups[-1])

    timestamp = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    backup_path = os.path.join(backup_folder, f"{company_tag}_Backup_{timestamp}.db")

    shutil.copy2(db_path, backup_path)

    return backup_path


# ============================================================
# RECONCILIATION: PURCHASE REGISTER vs GSTR-2B
# ============================================================

def _invoice_reconciliation_base(df, party_col, source_label):
    """Create a safe invoice-level reconciliation view.

    GST portal exports can contain multiple tax-rate rows for one invoice.
    Reconciliation must therefore aggregate by GSTIN + invoice before merging;
    otherwise a one-to-many merge creates duplicate rows and false mismatches.
    """
    if df.empty:
        return pd.DataFrame(columns=["match_key", "gstin_norm", "invoice_norm", "supplier_name", "invoice_date", "taxable_value", "total_tax"])

    work = df.copy()
    work["gstin_norm"] = work.get("gstin", "").fillna("").astype(str).str.strip().str.upper()
    work["invoice_norm"] = work.get("invoice_number", "").fillna("").astype(str).str.strip().str.upper()
    work["match_key"] = work["gstin_norm"] + "|" + work["invoice_norm"]
    work["taxable_value"] = pd.to_numeric(work.get("taxable_value", 0), errors="coerce").fillna(0)
    for col in ["cgst", "sgst", "igst"]:
        work[col] = pd.to_numeric(work.get(col, 0), errors="coerce").fillna(0)
    work["total_tax"] = work["cgst"] + work["sgst"] + work["igst"]
    work["invoice_date"] = work.get("invoice_date", "").astype(str).str.strip()
    work["supplier_name"] = work.get(party_col, "").fillna("").astype(str).str.strip()

    # Invalid key rows are retained for visibility but are never silently matched.
    valid = work["gstin_norm"].ne("") & work["invoice_norm"].ne("")
    work = work.loc[valid].copy()
    if work.empty:
        return pd.DataFrame(columns=["match_key", "gstin_norm", "invoice_norm", "supplier_name", "invoice_date", "taxable_value", "total_tax"])

    grouped = work.groupby("match_key", as_index=False).agg(
        gstin_norm=("gstin_norm", "first"),
        invoice_norm=("invoice_norm", "first"),
        supplier_name=("supplier_name", "first"),
        invoice_date=("invoice_date", "first"),
        taxable_value=("taxable_value", "sum"),
        total_tax=("total_tax", "sum"),
        source_rows=("match_key", "size"),
    )
    grouped["source_label"] = source_label
    return grouped


def _reconcile_invoice_sets(left_df, right_df, left_party, right_party,
                             left_prefix, right_prefix, tolerance=1.0):
    left = _invoice_reconciliation_base(left_df, left_party, left_prefix)
    right = _invoice_reconciliation_base(right_df, right_party, right_prefix)

    if left.empty and right.empty:
        return pd.DataFrame()

    # Primary match: GSTIN + invoice number. This is the authoritative match.
    merged = pd.merge(left, right, on="match_key", how="outer", suffixes=(f"_{left_prefix}", f"_{right_prefix}"))
    merged["match_basis"] = merged.apply(
        lambda r: "GSTIN + Invoice" if pd.notna(r.get(f"gstin_norm_{left_prefix}")) and pd.notna(r.get(f"gstin_norm_{right_prefix}")) else "",
        axis=1,
    )

    # Secondary diagnostic: when the invoice number occurs exactly once on each
    # side but GSTIN differs, pair it for review instead of showing two unrelated
    # missing records. Never auto-match ambiguous repeated invoice numbers.
    left_unique = left.groupby("invoice_norm").size()
    right_unique = right.groupby("invoice_norm").size()
    left_by_inv = left.set_index("invoice_norm")
    right_by_inv = right.set_index("invoice_norm")
    unmatched_left = merged[f"gstin_norm_{left_prefix}"].isna() | merged[f"gstin_norm_{right_prefix}"].isna()
    fallback_rows = []
    used_left_keys = set()
    used_right_keys = set()

    for inv in set(left_unique[left_unique == 1].index) & set(right_unique[right_unique == 1].index):
        lrow = left_by_inv.loc[inv]
        rrow = right_by_inv.loc[inv]
        lkey = lrow["match_key"]
        rkey = rrow["match_key"]
        if lkey == rkey:
            continue
        fallback_rows.append((lkey, rkey, inv, lrow, rrow))
        used_left_keys.add(lkey)
        used_right_keys.add(rkey)

    if fallback_rows:
        merged = merged[~merged["match_key"].isin(used_left_keys | used_right_keys)].copy()
        fallback = []
        for lkey, rkey, inv, lrow, rrow in fallback_rows:
            row = {"match_key": f"{lkey} ⇄ {rkey}", "match_basis": "Invoice Number Only (GSTIN differs)"}
            for c in ["gstin_norm", "supplier_name", "invoice_date", "taxable_value", "total_tax", "source_rows"]:
                row[f"{c}_{left_prefix}"] = lrow[c]
                row[f"{c}_{right_prefix}"] = rrow[c]
            row[f"invoice_norm_{left_prefix}"] = inv
            row[f"invoice_norm_{right_prefix}"] = inv
            fallback.append(row)
        merged = pd.concat([merged, pd.DataFrame(fallback)], ignore_index=True, sort=False)

    def num(row, col):
        value = row.get(col, 0)
        return float(value) if pd.notna(value) else 0.0

    statuses = []
    for _, row in merged.iterrows():
        in_left = bool(pd.notna(row.get(f"gstin_norm_{left_prefix}")))
        in_right = bool(pd.notna(row.get(f"gstin_norm_{right_prefix}")))
        basis = row.get("match_basis", "")
        taxable_diff = abs(num(row, f"taxable_value_{left_prefix}") - num(row, f"taxable_value_{right_prefix}"))
        tax_diff = abs(num(row, f"total_tax_{left_prefix}") - num(row, f"total_tax_{right_prefix}"))
        if basis.startswith("Invoice Number Only"):
            statuses.append("GSTIN Mismatch" if taxable_diff <= tolerance and tax_diff <= tolerance else "GSTIN Mismatch + Amount Mismatch")
        elif in_left and in_right:
            statuses.append("Matched" if taxable_diff <= tolerance and tax_diff <= tolerance else "Amount Mismatch")
        elif in_left:
            statuses.append(f"Missing in {right_prefix}")
        else:
            statuses.append(f"Missing in {left_prefix}")

    merged["status"] = statuses
    merged["taxable_diff"] = merged.apply(lambda r: num(r, f"taxable_value_{left_prefix}") - num(r, f"taxable_value_{right_prefix}"), axis=1)
    merged["tax_diff"] = merged.apply(lambda r: num(r, f"total_tax_{left_prefix}") - num(r, f"total_tax_{right_prefix}"), axis=1)

    return merged


def reconcile_purchase_2b(purchase_df, gstr2b_df, tolerance=1.0):
    """Invoice-level Purchase Register vs GSTR-2B reconciliation.

    Handles multiple 2B tax-rate rows per invoice and separately flags an
    invoice-number match with a different GSTIN so data-entry GSTIN errors are
    visible instead of being reported as two unrelated missing invoices.
    """
    merged = _reconcile_invoice_sets(
        purchase_df, gstr2b_df, "supplier_name", "supplier_name", "purchase", "2b", tolerance
    )
    if merged.empty:
        return merged
    merged["status"] = merged["status"].replace({
        "Missing in 2b": "Missing in GSTR-2B",
        "Missing in purchase": "Missing in Purchase Register",
    })

    display_columns = [
        "match_key", "match_basis", "status",
        "supplier_name_purchase", "supplier_name_2b",
        "invoice_norm_purchase", "invoice_norm_2b",
        "invoice_date_purchase", "invoice_date_2b",
        "gstin_norm_purchase", "gstin_norm_2b",
        "taxable_value_purchase", "taxable_value_2b", "taxable_diff",
        "total_tax_purchase", "total_tax_2b", "tax_diff",
        "source_rows_purchase", "source_rows_2b",
    ]
    for column in display_columns:
        if column not in merged.columns:
            merged[column] = None
    return merged[display_columns].rename(columns={
        "invoice_norm_purchase": "invoice_number_purchase",
        "invoice_norm_2b": "invoice_number_2b",
        "gstin_norm_purchase": "gstin_purchase",
        "gstin_norm_2b": "gstin_2b",
        "total_tax_purchase": "purchase_total_tax",
        "total_tax_2b": "gstr2b_total_tax",
    })


# ============================================================
# RECONCILIATION: SALES REGISTER vs GSTR-1 (INVOICE-WISE)
# ============================================================

def reconcile_sales_gstr1(sales_df, gstr1_df, tolerance=1.0):
    """Invoice-level Sales Register vs GSTR-1 reconciliation with safe
    aggregation and GSTIN-mismatch diagnostics."""
    merged = _reconcile_invoice_sets(
        sales_df, gstr1_df, "supplier_name", "customer_name", "sales", "gstr1", tolerance
    )
    if merged.empty:
        return merged
    merged["status"] = merged["status"].replace({
        "Missing in gstr1": "Missing in GSTR-1",
        "Missing in sales": "Missing in Sales Register",
    })

    display_columns = [
        "match_key", "match_basis", "status",
        "supplier_name_sales", "supplier_name_gstr1",
        "invoice_norm_sales", "invoice_norm_gstr1",
        "invoice_date_sales", "invoice_date_gstr1",
        "gstin_norm_sales", "gstin_norm_gstr1",
        "taxable_value_sales", "taxable_value_gstr1", "taxable_diff",
        "total_tax_sales", "total_tax_gstr1", "tax_diff",
        "source_rows_sales", "source_rows_gstr1",
    ]
    for column in display_columns:
        if column not in merged.columns:
            merged[column] = None
    return merged[display_columns].rename(columns={
        "invoice_norm_sales": "invoice_number_sales",
        "invoice_norm_gstr1": "invoice_number_gstr1",
        "gstin_norm_sales": "gstin_sales",
        "gstin_norm_gstr1": "gstin_gstr1",
        "total_tax_sales": "sales_total_tax",
        "total_tax_gstr1": "gstr1_total_tax",
    })


# ============================================================
# RECONCILIATION: GSTR-3B vs SALES INVOICES
# ============================================================

def reconcile_3b_sales(gstr3b_df, sales_df, tolerance=1.0):
    """Compares month-wise sales-register totals against the outward
    taxable value and output tax declared in GSTR-3B for that period."""

    if sales_df.empty:
        sales_summary = pd.DataFrame(
            columns=["return_period", "sales_taxable_value", "sales_tax", "invoice_count"]
        )
    else:
        working = sales_df.copy()
        working["return_period"] = working["invoice_date"].apply(to_period)
        working["invoice_tax"] = working["cgst"] + working["sgst"] + working["igst"]

        sales_summary = (
            working.groupby("return_period")
            .agg(
                sales_taxable_value=("taxable_value", "sum"),
                sales_tax=("invoice_tax", "sum"),
                invoice_count=("id", "count")
            )
            .reset_index()
        )

    returns = gstr3b_df.copy()

    if returns.empty:
        returns = pd.DataFrame(
            columns=["return_period", "outward_taxable_value", "output_tax", "itc_claimed",
                     "tax_paid_cash", "late_fee", "filed_date"]
        )

    returns = returns.drop(columns=["id", "created_date", "modified_date"], errors="ignore")

    merged = pd.merge(sales_summary, returns, on="return_period", how="outer")

    def classify(row):
        filed = pd.notna(row.get("output_tax"))
        has_sales = pd.notna(row.get("sales_tax"))

        if has_sales and not filed:
            return "GSTR-3B Not Filed"

        if filed and not has_sales:
            return "Filed, No Sales Recorded"

        taxable_diff = abs((row.get("sales_taxable_value") or 0) - (row.get("outward_taxable_value") or 0))
        tax_diff = abs((row.get("sales_tax") or 0) - (row.get("output_tax") or 0))

        if taxable_diff <= tolerance and tax_diff <= tolerance:
            return "Matched"

        return "Amount Mismatch"

    merged["status"] = merged.apply(classify, axis=1)
    merged["taxable_diff"] = merged.get("sales_taxable_value", 0).fillna(0) - merged.get("outward_taxable_value", 0).fillna(0)
    merged["tax_diff"] = merged.get("sales_tax", 0).fillna(0) - merged.get("output_tax", 0).fillna(0)

    merged = merged.sort_values("return_period", ascending=False)

    return merged


# ============================================================
# EXCEL STYLES (SHARED)
# ============================================================

TITLE_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FILL = PatternFill("solid", fgColor="4472C4")
SUBHEADER_FILL = PatternFill("solid", fgColor="D9EAF7")
MATCH_FILL = PatternFill("solid", fgColor="C6EFCE")
MISMATCH_FILL = PatternFill("solid", fgColor="FFEB9C")
MISSING_FILL = PatternFill("solid", fgColor="FFC7CE")

TITLE_FONT = Font(bold=True, size=16, color="FFFFFF")
HEADER_FONT = Font(bold=True, color="FFFFFF")
BOLD_FONT = Font(bold=True)

THIN_SIDE = Side(style="thin", color="B7B7B7")
BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
CENTER = Alignment(horizontal="center", vertical="center")

STATUS_FILLS = {
    "Matched": MATCH_FILL,
    "Amount Mismatch": MISMATCH_FILL,
    "Missing in GSTR-2B": MISSING_FILL,
    "Missing in Purchase Register": MISSING_FILL,
    "GSTR-3B Not Filed": MISSING_FILL,
    "Filed, No Sales Recorded": MISMATCH_FILL,
}


def write_title(sheet, text, span):
    sheet.merge_cells(f"A1:{span}1")
    sheet["A1"] = text
    sheet["A1"].font = TITLE_FONT
    sheet["A1"].fill = TITLE_FILL
    sheet["A1"].alignment = CENTER

    sheet.merge_cells(f"A2:{span}2")
    sheet["A2"] = "Generated On: " + datetime.now().strftime("%d-%m-%Y %H:%M:%S")
    sheet["A2"].alignment = CENTER


def write_table_sheet(workbook, sheet_name, title, dataframe, currency_columns=None,
                       status_column=None, table_name=None):
    """Writes a DataFrame to a new sheet as a formatted Excel table, with
    optional currency number formatting and status-based row colouring."""

    sheet = workbook.create_sheet(sheet_name)
    currency_columns = currency_columns or []

    columns = list(dataframe.columns)
    last_col_letter = get_column_letter(max(len(columns), 1))

    write_title(sheet, title, last_col_letter)

    header_row = 4

    for col, header in enumerate(columns, start=1):
        cell = sheet.cell(row=header_row, column=col)
        cell.value = header.replace("_", " ").title()
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = CENTER
        cell.border = BORDER

    if dataframe.empty:
        sheet.cell(row=header_row + 1, column=1).value = "No data available."
        return sheet

    for row_offset, (_, row) in enumerate(dataframe.iterrows(), start=1):
        excel_row = header_row + row_offset

        status_value = row.get(status_column) if status_column else None
        row_fill = STATUS_FILLS.get(status_value) if status_column else None

        for col, column_name in enumerate(columns, start=1):
            cell = sheet.cell(row=excel_row, column=col)
            value = row[column_name]

            if pd.isna(value):
                value = ""

            cell.value = value
            cell.border = BORDER

            if column_name in currency_columns:
                cell.number_format = '₹ #,##0.00'

            if row_fill is not None:
                cell.fill = row_fill

    last_row = header_row + len(dataframe)

    if table_name and last_row > header_row:
        table_ref = f"A{header_row}:{last_col_letter}{last_row}"
        table = Table(displayName=table_name, ref=table_ref)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        sheet.add_table(table)

    sheet.freeze_panes = f"A{header_row + 1}"

    return sheet


def autosize_columns(workbook):

    for sheet in workbook.worksheets:
        for column in sheet.columns:
            maximum = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    value_length = len(str(cell.value))
                    if value_length > maximum:
                        maximum = value_length
                except Exception:
                    pass

            sheet.column_dimensions[column_letter].width = min(maximum + 3, 40)


# ============================================================
# PROFESSIONAL EXCEL REPORT (WITH DATE RANGE + MODULE TOGGLES)
# ============================================================

def create_excel_report(from_date=None, to_date=None, include_purchase=True,
                         include_2b=True, include_returns=True, include_reconciliation=True):

    sales_data = load_data(from_date, to_date)
    purchase_data = load_purchase_data(from_date, to_date) if include_purchase else pd.DataFrame()
    gstr2b_data = load_2b_data() if include_2b else pd.DataFrame()
    gstr1_data = load_gstr1_data() if include_returns else pd.DataFrame()
    gstr3b_data = load_gstr3b_data() if include_returns else pd.DataFrame()
    company = load_company_profile()

    if sales_data.empty and purchase_data.empty:
        return None

    report_folder = os.path.join(st.session_state.save_path or ".", "Reports")
    os.makedirs(report_folder, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    report_path = os.path.join(report_folder, f"GST_Report_{timestamp}.xlsx")

    workbook = Workbook()

    # ------------------------------------------------------
    # DASHBOARD / SUMMARY SHEET
    # ------------------------------------------------------

    summary = workbook.active
    summary.title = "Dashboard"

    write_title(summary, "GST DATA MANAGEMENT REPORT", "F")

    if company.get("company_name"):
        summary.merge_cells("A3:F3")
        summary["A3"] = f"{company['company_name']}  |  GSTIN: {company.get('gstin', '')}"
        summary["A3"].font = BOLD_FONT
        summary["A3"].alignment = CENTER

    if from_date and to_date:
        summary.merge_cells("A4:F4")
        summary["A4"] = f"GST Period: {from_date:%Y-%m-%d} to {to_date:%Y-%m-%d}"
        summary["A4"].alignment = CENTER

    summary["A6"] = "SALES SUMMARY"
    summary["A6"].font = HEADER_FONT
    summary["A6"].fill = HEADER_FILL

    sales_gst = sales_data["cgst"].sum() + sales_data["sgst"].sum() + sales_data["igst"].sum() if not sales_data.empty else 0
    purchase_gst = purchase_data["cgst"].sum() + purchase_data["sgst"].sum() + purchase_data["igst"].sum() if not purchase_data.empty else 0

    sales_items = [
        ("Total Sales Invoices", len(sales_data)),
        ("Total Taxable Value (Sales)", sales_data["taxable_value"].sum() if not sales_data.empty else 0),
        ("Total Output GST", sales_gst),
        ("Total Sales Invoice Value", sales_data["total_amount"].sum() if not sales_data.empty else 0),
    ]

    row = 7
    for label, value in sales_items:
        summary.cell(row=row, column=1).value = label
        summary.cell(row=row, column=2).value = value
        summary.cell(row=row, column=1).border = BORDER
        summary.cell(row=row, column=2).border = BORDER
        if "Invoices" not in label:
            summary.cell(row=row, column=2).number_format = '₹ #,##0.00'
        row += 1

    row += 1
    summary.cell(row=row, column=1).value = "PURCHASE SUMMARY"
    summary.cell(row=row, column=1).font = HEADER_FONT
    summary.cell(row=row, column=1).fill = HEADER_FILL
    row += 1

    purchase_items = [
        ("Total Purchase Invoices", len(purchase_data)),
        ("Total Taxable Value (Purchases)", purchase_data["taxable_value"].sum() if not purchase_data.empty else 0),
        ("Total ITC (Input GST)", purchase_gst),
        ("Net GST Payable (Output - ITC)", sales_gst - purchase_gst),
    ]

    for label, value in purchase_items:
        summary.cell(row=row, column=1).value = label
        summary.cell(row=row, column=2).value = value
        summary.cell(row=row, column=1).border = BORDER
        summary.cell(row=row, column=2).border = BORDER
        if "Invoices" not in label:
            summary.cell(row=row, column=2).number_format = '₹ #,##0.00'
        row += 1

    # ------------------------------------------------------
    # SALES / PURCHASE / 2B / RETURNS SHEETS
    # ------------------------------------------------------

    currency_cols_invoice = ["taxable_value", "cgst", "sgst", "igst", "total_amount"]

    if not sales_data.empty:
        write_table_sheet(
            workbook, "Sales Register", "SALES (OUTWARD SUPPLIES) REGISTER",
            sales_data.drop(columns=["id"], errors="ignore"),
            currency_columns=currency_cols_invoice, table_name="SalesTable"
        )

    if include_purchase and not purchase_data.empty:
        write_table_sheet(
            workbook, "Purchase Register", "PURCHASE (INWARD SUPPLIES) REGISTER",
            purchase_data.drop(columns=["id"], errors="ignore"),
            currency_columns=currency_cols_invoice, table_name="PurchaseTable"
        )

    if include_2b and not gstr2b_data.empty:
        write_table_sheet(
            workbook, "GSTR-2B", "GSTR-2B (AS PER PORTAL)",
            gstr2b_data.drop(columns=["id"], errors="ignore"),
            currency_columns=["taxable_value", "cgst", "sgst", "igst"], table_name="GSTR2BTable"
        )

    if include_returns and not gstr1_data.empty:
        write_table_sheet(
            workbook, "GSTR-1", "GSTR-1 (INVOICE-WISE, AS FILED)",
            gstr1_data.drop(columns=["id"], errors="ignore"),
            currency_columns=["taxable_value", "cgst", "sgst", "igst", "total_amount"],
            table_name="GSTR1Table"
        )

    if include_returns and not gstr3b_data.empty:
        write_table_sheet(
            workbook, "GSTR-3B", "GSTR-3B RETURN SUMMARY",
            gstr3b_data.drop(columns=["id"], errors="ignore"),
            currency_columns=["outward_taxable_value", "output_tax", "itc_claimed", "tax_paid_cash", "late_fee"],
            table_name="GSTR3BTable"
        )

    # ------------------------------------------------------
    # RECONCILIATION SHEETS
    # ------------------------------------------------------

    if include_reconciliation:

        if include_purchase and include_2b and not (purchase_data.empty and gstr2b_data.empty):
            recon_p2b = reconcile_purchase_2b(purchase_data, gstr2b_data)
            write_table_sheet(
                workbook, "Recon - Purchase vs 2B", "PURCHASE REGISTER vs GSTR-2B RECONCILIATION",
                recon_p2b,
                currency_columns=["taxable_value_purchase", "taxable_value_2b", "taxable_diff",
                                   "purchase_total_tax", "gstr2b_total_tax", "tax_diff"],
                status_column="status", table_name="ReconPurchase2BTable"
            )

        if include_returns and not (sales_data.empty and gstr1_data.empty):
            recon_1sales = reconcile_sales_gstr1(sales_data, gstr1_data)
            write_table_sheet(
                workbook, "Recon - Sales vs GSTR-1", "SALES REGISTER vs GSTR-1 RECONCILIATION",
                recon_1sales,
                currency_columns=["taxable_value_sales", "taxable_value_gstr1", "taxable_diff",
                                   "sales_total_tax", "gstr1_total_tax", "tax_diff"],
                status_column="status", table_name="ReconSalesGSTR1Table"
            )

        if include_returns and not (gstr3b_data.empty and sales_data.empty):
            recon_3b = reconcile_3b_sales(gstr3b_data, sales_data)
            write_table_sheet(
                workbook, "Recon - 3B vs Sales", "GSTR-3B vs SALES RECONCILIATION",
                recon_3b,
                currency_columns=["sales_taxable_value", "outward_taxable_value", "taxable_diff",
                                   "sales_tax", "output_tax", "tax_diff", "itc_claimed",
                                   "tax_paid_cash", "late_fee"],
                status_column="status", table_name="Recon3BSalesTable"
            )

    autosize_columns(workbook)
    workbook.save(report_path)

    return report_path


# ============================================================
# AUTHENTICATION GATE
# ============================================================
# Nothing below this point executes until a client has successfully
# logged in. This keeps every client's GST data isolated to their own
# folder and inaccessible without their credentials.

if not st.session_state.auth_user:
    render_login_register_screen()
    st.stop()


# ============================================================
# INITIALIZE DATABASE
# ============================================================

if get_db_path():
    create_database()


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.markdown(
    '<div style="font-size:22px; font-weight:800; color:#1a1d29; padding-bottom:2px;">📊 GST Data Tool</div>'
    '<div class="gst-subtle" style="margin-bottom:10px;">Management &amp; Reconciliation</div>',
    unsafe_allow_html=True
)

_client_record = st.session_state.client_record or {}
st.sidebar.success(f"👤 {_client_record.get('display_name') or st.session_state.auth_user}")
if st.sidebar.button("🚪 Log Out", use_container_width=True):
    st.session_state.auth_user = None
    st.session_state.client_record = None
    st.session_state.save_path = ""
    st.session_state.active_db = ""
    st.rerun()

st.sidebar.divider()

module = st.sidebar.radio(
    "MODULE",
    [
        "🏠 Dashboard",
        "🧾 Sales",
        "🛒 Purchase & 2B",
        "📄 Returns (GSTR-1 / 3B)",
        "🔄 Reconciliation",
        "📊 Reports",
        "📥 Export",
        "💾 Backup",
        "📑 GSTR-9",
        "📋 GSTR-9C",
        "🛡️ GST Compliance",
        "💼 Creditors & Debtors",
        "🧾 GSTR-1 JSON Generator",
        "⚙️ Settings",
    ],
    label_visibility="collapsed"
)

st.sidebar.divider()
st.sidebar.markdown("**🏢 Company Profile**")

if get_db_path():
    try:
        profile = load_company_profile()
    except Exception:
        profile = {}
    company_label = profile.get("company_name") or get_active_company_label()
    st.sidebar.success(company_label)
    if profile.get("gstin"):
        st.sidebar.caption(f"GSTIN: {profile['gstin']}")
    else:
        st.sidebar.caption("Set up the company GSTIN under Settings.")
else:
    st.sidebar.warning("Setting up your workspace…")


# ============================================================
# DASHBOARD
# ============================================================

if module == "🏠 Dashboard":

    st.markdown('<div class="main-title">📊 GST Dashboard</div>', unsafe_allow_html=True)

    if not get_db_path():
        st.warning("Please set up a storage folder or load a company file from Settings.")

    else:
        sales_data = load_data()
        purchase_data = load_purchase_data()

        if sales_data.empty and purchase_data.empty:
            st.info("No data available yet. Start by adding invoices under Sales or Purchase & 2B.")

        else:
            # ---- KPI Row ----
            sales_taxable = sales_data["taxable_value"].sum() if not sales_data.empty else 0
            sales_gst = (sales_data["cgst"].sum() + sales_data["sgst"].sum() + sales_data["igst"].sum()) if not sales_data.empty else 0
            purchase_taxable = purchase_data["taxable_value"].sum() if not purchase_data.empty else 0
            purchase_gst = (purchase_data["cgst"].sum() + purchase_data["sgst"].sum() + purchase_data["igst"].sum()) if not purchase_data.empty else 0
            gstr2b_dashboard = load_2b_data()
            itc_reversed = float(
                purchase_data.loc[
                    purchase_data.get("itc_eligible", pd.Series(dtype=str)).astype(str).str.strip().str.lower().eq("no"),
                    ["cgst", "sgst", "igst"]
                ].sum().sum()
            ) if not purchase_data.empty and "itc_eligible" in purchase_data.columns else 0
            gstr3b_dashboard = load_gstr3b_data()
            gstr3b_itc_reversed = float(gstr3b_dashboard["itc_reversed"].sum()) if not gstr3b_dashboard.empty and "itc_reversed" in gstr3b_dashboard.columns else 0
            itc_reversed += gstr3b_itc_reversed
            ineligible_itc_claimed = float(
                gstr2b_dashboard.loc[
                    gstr2b_dashboard.get("itc_eligible", pd.Series(dtype=str)).astype(str).str.strip().str.lower().eq("ineligible but claimed"),
                    ["cgst", "sgst", "igst"]
                ].sum().sum()
            ) if not gstr2b_dashboard.empty and "itc_eligible" in gstr2b_dashboard.columns else 0
            net_payable = sales_gst - purchase_gst

            c1, c2, c3, c4, c5, c6, c7 = st.columns(7)
            c1.metric("Sales Invoices", len(sales_data))
            c2.metric("Purchase Invoices", len(purchase_data))
            c3.metric("Output GST (Sales)", f"₹ {sales_gst:,.0f}")
            c4.metric("ITC Available (Purchases)", f"₹ {purchase_gst:,.0f}")
            c5.metric("ITC Reversed", f"₹ {itc_reversed:,.0f}")
            c6.metric("Ineligible ITC Claimed", f"₹ {ineligible_itc_claimed:,.0f}")
            c7.metric("Net GST Payable", f"₹ {net_payable:,.0f}")

            st.divider()

            chart_col1, chart_col2 = st.columns(2)

            # ---- Monthly Sales vs Purchase trend ----
            with chart_col1:
                st.subheader("📈 Monthly Sales vs Purchase")

                trend_frames = []

                if not sales_data.empty:
                    s = sales_data.copy()
                    s["Month"] = s["invoice_date"].apply(to_period)
                    s = s.groupby("Month")["taxable_value"].sum().reset_index()
                    s["Type"] = "Sales"
                    trend_frames.append(s)

                if not purchase_data.empty:
                    p = purchase_data.copy()
                    p["Month"] = p["invoice_date"].apply(to_period)
                    p = p.groupby("Month")["taxable_value"].sum().reset_index()
                    p["Type"] = "Purchase"
                    trend_frames.append(p)

                if trend_frames:
                    trend = pd.concat(trend_frames)

                    if PLOTLY_AVAILABLE:
                        fig = px.bar(
                            trend, x="Month", y="taxable_value", color="Type",
                            barmode="group", labels={"taxable_value": "Taxable Value (₹)"}
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        pivot = trend.pivot(index="Month", columns="Type", values="taxable_value").fillna(0)
                        st.bar_chart(pivot)

            # ---- GST rate-wise distribution (Sales) ----
            with chart_col2:
                st.subheader("🥧 Sales by GST Rate")

                if not sales_data.empty:
                    rate_dist = sales_data.groupby("gst_rate")["taxable_value"].sum().reset_index()

                    if PLOTLY_AVAILABLE:
                        fig = px.pie(
                            rate_dist, names="gst_rate", values="taxable_value",
                            hole=0.4, labels={"gst_rate": "GST Rate (%)"}
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.bar_chart(rate_dist.set_index("gst_rate"))
                else:
                    st.info("No sales data yet.")

            st.divider()

            chart_col3, chart_col4 = st.columns(2)

            # ---- Reconciliation snapshot: Purchase vs 2B ----
            with chart_col3:
                st.subheader("🔄 Purchase vs GSTR-2B Status")

                gstr2b_data = load_2b_data()

                if purchase_data.empty and gstr2b_data.empty:
                    st.info("Add purchase invoices and GSTR-2B data to see this.")
                else:
                    recon = reconcile_purchase_2b(purchase_data, gstr2b_data)
                    status_counts = recon["status"].value_counts().reset_index()
                    status_counts.columns = ["Status", "Count"]

                    if PLOTLY_AVAILABLE:
                        fig = px.pie(status_counts, names="Status", values="Count", hole=0.4)
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.bar_chart(status_counts.set_index("Status"))

            # ---- Reconciliation snapshot: 3B vs Sales ----
            with chart_col4:
                st.subheader("🔄 GSTR-3B vs Sales Variance")

                gstr3b_data = load_gstr3b_data()

                if gstr3b_data.empty and sales_data.empty:
                    st.info("Add sales invoices and GSTR-3B summaries to see this.")
                else:
                    recon3b = reconcile_3b_sales(gstr3b_data, sales_data)
                    recon3b_display = recon3b[["return_period", "tax_diff", "status"]].dropna(subset=["return_period"])

                    if PLOTLY_AVAILABLE and not recon3b_display.empty:
                        fig = px.bar(
                            recon3b_display, x="return_period", y="tax_diff", color="status",
                            labels={"return_period": "Return Period", "tax_diff": "Tax Variance (₹)"}
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    elif not recon3b_display.empty:
                        st.bar_chart(recon3b_display.set_index("return_period")["tax_diff"])
                    else:
                        st.info("No overlapping periods yet.")

            st.divider()
            st.subheader("🧾 Recent Sales Invoices")
            st.dataframe(sales_data.head(10), use_container_width=True, hide_index=True)


# ============================================================
# SALES MODULE
# ============================================================

elif module == "🧾 Sales":

    st.markdown('<div class="main-title">🧾 Sales (Outward Supplies)</div>', unsafe_allow_html=True)

    if not get_db_path():
        st.error("Please select a storage folder or company from Settings first.")

    else:
        tab_add, tab_upload, tab_edit, tab_delete, tab_register, tab_notes = st.tabs(
            ["➕ Add Invoice", "📤 Upload Sales Register", "✏️ Edit Invoice", "🗑️ Delete Invoice", "📋 Register", "📝 Credit / Debit Notes"]
        )

        # ---- ADD ----
        with tab_add:

            with st.form("add_invoice"):
                col1, col2 = st.columns(2)

                with col1:
                    supplier_name = st.text_input("Customer Name *")
                    gstin = st.text_input("GSTIN *").upper()
                    invoice_number = st.text_input("Invoice Number *")
                    invoice_date = st.date_input("Invoice Date")

                with col2:
                    taxable_value = st.number_input("Taxable Value", min_value=0.0, step=100.0)
                    gst_rate = st.selectbox("GST Rate (%)", GST_RATES, index=GST_RATES.index(18) if 18 in GST_RATES else 0)
                    transaction_type = st.selectbox("Transaction Type", TRANSACTION_TYPES)
                    supply_category = st.selectbox("Supply Category", SUPPLY_CATEGORIES)
                    supply_type = st.selectbox("Supply Type", SUPPLY_TYPES)

                save_invoice = st.form_submit_button("💾 Save Invoice", use_container_width=True, type="primary")

            if save_invoice:
                if not supplier_name:
                    st.error("Customer Name is required.")
                elif not gstin:
                    st.error("GSTIN is required.")
                elif not invoice_number:
                    st.error("Invoice Number is required.")
                else:
                    cgst, sgst, igst, total = calculate_gst(taxable_value, gst_rate, transaction_type)
                    create_backup()
                    insert_invoice(supplier_name, gstin, invoice_number, invoice_date,
                                    taxable_value, gst_rate, transaction_type, cgst, sgst, igst, total,
                                    supply_category, supply_type)

                    st.success("✅ Invoice saved successfully!")
                    st.write(f"CGST: ₹ {cgst:,.2f}  |  SGST: ₹ {sgst:,.2f}  |  IGST: ₹ {igst:,.2f}  |  Total: ₹ {total:,.2f}")

        # ---- BULK UPLOAD SALES REGISTER ----
        with tab_upload:
            st.markdown('<div class="gst-card">', unsafe_allow_html=True)

            sales_upload_period = st.text_input(
                "Period for invoice-date check (YYYY-MM) *",
                value=datetime.now().strftime("%Y-%m"),
                key="sales_upload_period"
            )

            clean_df = excel_csv_upload_and_map(
                "Choose Sales Register CSV or Excel",
                "sales_upload",
                required_fields=["supplier_name", "invoice_number", "invoice_date", "taxable_value"],
                help_text="Works with any column layout from Tally, Excel or the GST portal — you'll confirm the mapping before anything is imported.",
                party_field="supplier_name",
                party_label="Customer / Buyer Name"
            )

            if clean_df is not None:
                if not sales_upload_period.strip():
                    st.error("Please enter the period for the invoice-date check before importing.")
                else:
                    duplicate_ids = get_duplicate_import_ids(clean_df, "gst_invoices")
                    period_mismatch_rows = get_period_mismatch_rows(clean_df, sales_upload_period)
                    period_mismatch_confirmed = period_mismatch_rows.empty

                    if not period_mismatch_rows.empty:
                        st.warning(
                            f"{len(period_mismatch_rows)} uploaded invoice(s) do not belong to "
                            f"the selected period {sales_upload_period}."
                        )
                        st.dataframe(
                            period_mismatch_rows[
                                [column for column in ["gstin", "invoice_number", "invoice_date"]
                                 if column in period_mismatch_rows.columns]
                            ],
                            use_container_width=True,
                            hide_index=True
                        )
                        period_mismatch_confirmed = st.checkbox(
                            "I still want to import these records.",
                            key="sales_period_mismatch_confirm"
                        )

                    overwrite_existing = False
                    if duplicate_ids:
                        import_choice = st.radio(
                            f"{len(duplicate_ids)} uploaded invoice(s) already exist. What should I do?",
                            ["Keep previous data and include this upload", "Overwrite existing matching data"],
                            key="sales_import_choice"
                        )
                        overwrite_existing = import_choice == "Overwrite existing matching data"

                    if st.button("📥 Import Sales Register", use_container_width=True, type="primary", key="sales_import_btn"):
                        if not period_mismatch_confirmed:
                            st.error("Please confirm the period mismatch before importing.")
                        else:
                            if overwrite_existing:
                                create_backup()
                                overwrite_import_records(duplicate_ids, "gst_invoices")
                            count, skipped = bulk_insert_sales(clean_df)
                            show_import_results(count, skipped, imported_noun="sales invoice(s)")
                            if count and not skipped:
                                st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        # ---- EDIT ----
        with tab_edit:

            data = load_data()

            if data.empty:
                st.info("No invoices available.")
            else:
                invoice_list = {
                    f"{row.invoice_number} | {row.supplier_name} | ID {row.id}": row.id
                    for row in data.itertuples()
                }

                selected = st.selectbox("Select Invoice", list(invoice_list.keys()), key="edit_select")
                invoice_id = invoice_list[selected]
                row = data[data["id"] == invoice_id].iloc[0]

                with st.form("edit_invoice"):
                    edit_supplier = st.text_input("Customer Name", value=row["supplier_name"])
                    edit_gstin = st.text_input("GSTIN", value=row["gstin"])
                    edit_invoice_number = st.text_input("Invoice Number", value=row["invoice_number"])
                    edit_invoice_date = st.date_input(
                        "Invoice Date", value=safe_date_input_value(row["invoice_date"])
                    )
                    edit_taxable = st.number_input("Taxable Value", value=float(row["taxable_value"]))

                    current_rate = float(row["gst_rate"])
                    rate_index = GST_RATES.index(current_rate) if current_rate in GST_RATES else 0
                    edit_rate = st.selectbox("GST Rate", GST_RATES, index=rate_index)

                    type_index = TRANSACTION_TYPES.index(row["transaction_type"]) if row["transaction_type"] in TRANSACTION_TYPES else 0
                    edit_type = st.selectbox("Transaction Type", TRANSACTION_TYPES, index=type_index)
                    edit_category = st.selectbox("Supply Category", SUPPLY_CATEGORIES, index=SUPPLY_CATEGORIES.index(row.get("supply_category", "B2B")) if row.get("supply_category", "B2B") in SUPPLY_CATEGORIES else 0)
                    edit_supply_type = st.selectbox("Supply Type", SUPPLY_TYPES, index=SUPPLY_TYPES.index(row.get("supply_type", "Goods")) if row.get("supply_type", "Goods") in SUPPLY_TYPES else 0)

                    update = st.form_submit_button("💾 Update Invoice", use_container_width=True)

                if update:
                    cgst, sgst, igst, total = calculate_gst(edit_taxable, edit_rate, edit_type)
                    create_backup()
                    update_invoice(invoice_id, edit_supplier, edit_gstin.upper(), edit_invoice_number,
                                    edit_invoice_date, edit_taxable, edit_rate, edit_type, cgst, sgst, igst, total,
                                    edit_category, edit_supply_type)

                    st.success("✅ Invoice updated successfully!")
                    st.rerun()

        # ---- DELETE ----
        with tab_delete:

            data = load_data()

            if data.empty:
                st.info("No invoices available.")
            else:
                delete_table = data.copy()
                select_all_sales_delete = st.checkbox("Select all sales invoices", key="select_all_sales_delete")
                delete_table.insert(0, "Delete", select_all_sales_delete)
                edited_delete_table = st.data_editor(
                    delete_table,
                    use_container_width=True,
                    hide_index=True,
                    key="sales_delete_table",
                    column_config={
                        "Delete": st.column_config.CheckboxColumn("Delete", default=False),
                        "id": st.column_config.NumberColumn("ID", disabled=True),
                    },
                    disabled=[column for column in delete_table.columns if column != "Delete"],
                )
                invoice_ids = edited_delete_table.loc[
                    edited_delete_table["Delete"], "id"
                ].astype(int).tolist()

                st.warning("⚠️ Deleting invoices is permanent.")
                confirm = st.checkbox("I confirm that I want to delete the selected invoice(s).", key="delete_confirm")

                if st.button("🗑️ Delete Selected Invoice(s)", use_container_width=True):
                    if not invoice_ids:
                        st.error("Please select at least one invoice.")
                    elif not confirm:
                        st.error("Please confirm deletion.")
                    else:
                        create_backup()
                        delete_records_by_ids("gst_invoices", invoice_ids)
                        st.success(f"✅ Deleted {len(invoice_ids)} invoice(s).")
                        st.rerun()

        # ---- REGISTER ----
        with tab_register:

            data = load_data()

            if data.empty:
                st.info("No invoices available.")
            else:
                search = st.text_input("🔍 Search Customer / GSTIN / Invoice Number", key="sales_search")

                filtered = data.copy()

                if search:
                    search = search.lower()
                    filtered = filtered[
                        filtered["supplier_name"].astype(str).str.lower().str.contains(search, na=False)
                        | filtered["gstin"].astype(str).str.lower().str.contains(search, na=False)
                        | filtered["invoice_number"].astype(str).str.lower().str.contains(search, na=False)
                    ]

                register_table = filtered.copy()
                select_all_sales = st.checkbox("Select all visible sales invoices", key="select_all_sales")
                register_table.insert(0, "Select", select_all_sales)
                edited_register_table = st.data_editor(
                    register_table,
                    use_container_width=True,
                    hide_index=True,
                    key="sales_register_action_table",
                    column_config={
                        "Select": st.column_config.CheckboxColumn("Select", default=False),
                        "id": st.column_config.NumberColumn("ID", disabled=True),
                    },
                    disabled=[column for column in register_table.columns if column != "Select"],
                )
                selected_register_ids = edited_register_table.loc[
                    edited_register_table["Select"], "id"
                ].astype(int).tolist()
                register_action = st.selectbox(
                    "Action for selected sales invoice",
                    ["Choose action", "Edit", "Delete"],
                    key="sales_register_action"
                )

                if register_action == "Edit":
                    if len(selected_register_ids) != 1:
                        st.warning("Select exactly one sales invoice to edit.")
                    else:
                        action_id = selected_register_ids[0]
                        action_row = data[data["id"] == action_id].iloc[0]
                        with st.form("sales_register_edit_form"):
                            action_supplier = st.text_input("Customer Name", value=action_row["supplier_name"])
                            action_gstin = st.text_input("GSTIN", value=action_row["gstin"])
                            action_invoice_number = st.text_input("Invoice Number", value=action_row["invoice_number"])
                            action_invoice_date = st.date_input("Invoice Date", value=safe_date_input_value(action_row["invoice_date"]))
                            action_taxable = st.number_input("Taxable Value", value=float(action_row["taxable_value"]))
                            action_rate = st.selectbox("GST Rate", GST_RATES, index=GST_RATES.index(float(action_row["gst_rate"])) if float(action_row["gst_rate"]) in GST_RATES else 0)
                            action_type = st.selectbox("Transaction Type", TRANSACTION_TYPES, index=TRANSACTION_TYPES.index(action_row["transaction_type"]) if action_row["transaction_type"] in TRANSACTION_TYPES else 0)
                            action_category = st.selectbox("Supply Category", SUPPLY_CATEGORIES, index=SUPPLY_CATEGORIES.index(action_row.get("supply_category", "B2B")) if action_row.get("supply_category", "B2B") in SUPPLY_CATEGORIES else 0)
                            action_supply_type = st.selectbox("Supply Type", SUPPLY_TYPES, index=SUPPLY_TYPES.index(action_row.get("supply_type", "Goods")) if action_row.get("supply_type", "Goods") in SUPPLY_TYPES else 0)
                            action_save = st.form_submit_button("💾 Save Invoice Changes", type="primary")

                        if action_save:
                            cgst, sgst, igst, total = calculate_gst(action_taxable, action_rate, action_type)
                            create_backup()
                            update_invoice(action_id, action_supplier, action_gstin.upper(), action_invoice_number, action_invoice_date, action_taxable, action_rate, action_type, cgst, sgst, igst, total, action_category, action_supply_type)
                            st.success("✅ Sales invoice updated.")
                            st.rerun()

                elif register_action == "Delete":
                    if not selected_register_ids:
                        st.warning("Select at least one sales invoice to delete.")
                    else:
                        st.warning(f"Deleting {len(selected_register_ids)} selected sales invoice(s) is permanent.")
                        action_confirm = st.checkbox("I confirm deletion of the selected sales invoice(s).", key="sales_register_delete_confirm")
                        if st.button("🗑️ Delete Selected Sales Invoice(s)", key="sales_register_delete_button"):
                            if not action_confirm:
                                st.error("Please confirm deletion.")
                            else:
                                create_backup()
                                delete_records_by_ids("gst_invoices", selected_register_ids)
                                st.success(f"✅ Deleted {len(selected_register_ids)} sales invoice(s).")
                                st.rerun()
                st.caption(f"Showing {len(filtered)} of {len(data)} invoices.")

        with tab_notes:
            render_credit_debit_notes("Sales")


# ============================================================
# PURCHASE & 2B MODULE
# ============================================================

elif module == "🛒 Purchase & 2B":

    st.markdown('<div class="main-title">🛒 Purchase Register & GSTR-2B</div>', unsafe_allow_html=True)

    if not get_db_path():
        st.error("Please select a storage folder or company from Settings first.")

    else:
        tab_add, tab_register, tab_purchase_upload, tab_2b_add, tab_2b_upload, tab_2b_register, tab_notes = st.tabs(
            ["➕ Add Purchase", "📋 Purchase Register", "📤 Upload Purchase Register",
             "➕ Add 2B Entry", "📤 Upload 2B (CSV/Excel)", "📄 2B Register", "📝 Credit / Debit Notes"]
        )

        # ---- ADD PURCHASE ----
        with tab_add:

            with st.form("add_purchase"):
                col1, col2 = st.columns(2)

                with col1:
                    p_supplier = st.text_input("Supplier Name *")
                    p_gstin = st.text_input("Supplier GSTIN *").upper()
                    p_invoice_number = st.text_input("Bill / Invoice Number *")
                    p_invoice_date = st.date_input("Bill Date")

                with col2:
                    p_taxable_value = st.number_input("Taxable Value", min_value=0.0, step=100.0, key="p_taxable")
                    p_gst_rate = st.selectbox("GST Rate (%)", GST_RATES, index=GST_RATES.index(18) if 18 in GST_RATES else 0, key="p_rate")
                    p_type = st.selectbox("Transaction Type", TRANSACTION_TYPES, key="p_type")
                    p_itc = st.selectbox("ITC Eligible?", ["Yes", "No"], help="Select 'No' for blocked credit under Section 17(5).")
                    p_supply_category = st.selectbox("Supply Category", SUPPLY_CATEGORIES, key="p_supply_category")
                    p_supply_type = st.selectbox("Supply Type", SUPPLY_TYPES, key="p_supply_type")

                save_purchase = st.form_submit_button("💾 Save Purchase", use_container_width=True, type="primary")

            if save_purchase:
                if not p_supplier or not p_gstin or not p_invoice_number:
                    st.error("Supplier Name, GSTIN and Bill Number are required.")
                else:
                    cgst, sgst, igst, total = calculate_gst(p_taxable_value, p_gst_rate, p_type)
                    create_backup()
                    insert_purchase(p_supplier, p_gstin, p_invoice_number, p_invoice_date,
                                     p_taxable_value, p_gst_rate, p_type, cgst, sgst, igst, total, p_itc,
                                     p_supply_category, p_supply_type)

                    st.success("✅ Purchase saved successfully!")
                    st.write(f"CGST: ₹ {cgst:,.2f}  |  SGST: ₹ {sgst:,.2f}  |  IGST: ₹ {igst:,.2f}  |  Total: ₹ {total:,.2f}")

        # ---- BULK UPLOAD PURCHASE REGISTER ----
        with tab_purchase_upload:
            st.markdown('<div class="gst-card">', unsafe_allow_html=True)

            purchase_upload_period = st.text_input(
                "Period for invoice-date check (YYYY-MM) *",
                value=datetime.now().strftime("%Y-%m"),
                key="purchase_upload_period"
            )

            clean_df = excel_csv_upload_and_map(
                "Choose Purchase Register CSV or Excel",
                "purchase_upload",
                required_fields=["supplier_name", "invoice_number", "invoice_date", "taxable_value"],
                help_text="Works with any column layout — you'll confirm the mapping before anything is imported.",
                party_field="supplier_name",
                party_label="Supplier / Vendor Name"
            )

            if clean_df is not None:
                if not purchase_upload_period.strip():
                    st.error("Please enter the period for the invoice-date check before importing.")
                else:
                    duplicate_ids = get_duplicate_import_ids(clean_df, "purchase_invoices")
                    period_mismatch_rows = get_period_mismatch_rows(clean_df, purchase_upload_period)
                    period_mismatch_confirmed = period_mismatch_rows.empty

                    if not period_mismatch_rows.empty:
                        st.warning(
                            f"{len(period_mismatch_rows)} uploaded invoice(s) do not belong to "
                            f"the selected period {purchase_upload_period}."
                        )
                        st.dataframe(
                            period_mismatch_rows[
                                [column for column in ["gstin", "invoice_number", "invoice_date"]
                                 if column in period_mismatch_rows.columns]
                            ],
                            use_container_width=True,
                            hide_index=True
                        )
                        period_mismatch_confirmed = st.checkbox(
                            "I still want to import these records.",
                            key="purchase_period_mismatch_confirm"
                        )

                    overwrite_existing = False
                    if duplicate_ids:
                        import_choice = st.radio(
                            f"{len(duplicate_ids)} uploaded invoice(s) already exist. What should I do?",
                            ["Keep previous data and include this upload", "Overwrite existing matching data"],
                            key="purchase_import_choice"
                        )
                        overwrite_existing = import_choice == "Overwrite existing matching data"

                    if st.button("📥 Import Purchase Register", use_container_width=True, type="primary", key="purchase_import_btn"):
                        if not period_mismatch_confirmed:
                            st.error("Please confirm the period mismatch before importing.")
                        else:
                            if overwrite_existing:
                                create_backup()
                                overwrite_import_records(duplicate_ids, "purchase_invoices")
                            count, skipped = bulk_insert_purchase(clean_df, purchase_upload_period)
                            show_import_results(count, skipped, imported_noun="purchase invoice(s)")
                            if count and not skipped:
                                st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        # ---- PURCHASE REGISTER (with edit/delete) ----
        with tab_register:

            purchase_data = load_purchase_data()

            if purchase_data.empty:
                st.info("No purchase invoices recorded yet.")
            else:
                search = st.text_input("🔍 Search Supplier / GSTIN / Bill Number", key="purchase_search")

                filtered = purchase_data.copy()
                if search:
                    search = search.lower()
                    filtered = filtered[
                        filtered["supplier_name"].astype(str).str.lower().str.contains(search, na=False)
                        | filtered["gstin"].astype(str).str.lower().str.contains(search, na=False)
                        | filtered["invoice_number"].astype(str).str.lower().str.contains(search, na=False)
                    ]

                register_table = filtered.copy()
                select_all_purchase = st.checkbox("Select all visible purchase invoices", key="select_all_purchase")
                register_table.insert(0, "Select", select_all_purchase)
                edited_register_table = st.data_editor(
                    register_table,
                    use_container_width=True,
                    hide_index=True,
                    key="purchase_register_action_table",
                    column_config={
                        "Select": st.column_config.CheckboxColumn("Select", default=False),
                        "id": st.column_config.NumberColumn("ID", disabled=True),
                    },
                    disabled=[column for column in register_table.columns if column != "Select"],
                )
                selected_register_ids = edited_register_table.loc[
                    edited_register_table["Select"], "id"
                ].astype(int).tolist()
                register_action = st.selectbox(
                    "Action for selected purchase invoice",
                    ["Choose action", "Edit", "Delete"],
                    key="purchase_register_action"
                )

                if register_action == "Edit":
                    if len(selected_register_ids) != 1:
                        st.warning("Select exactly one purchase invoice to edit.")
                    else:
                        action_id = selected_register_ids[0]
                        action_row = purchase_data[purchase_data["id"] == action_id].iloc[0]
                        with st.form("purchase_register_edit_form"):
                            action_supplier = st.text_input("Supplier Name", value=action_row["supplier_name"])
                            action_gstin = st.text_input("GSTIN", value=action_row["gstin"])
                            action_invoice_number = st.text_input("Bill Number", value=action_row["invoice_number"])
                            action_invoice_date = st.date_input("Bill Date", value=safe_date_input_value(action_row["invoice_date"]))
                            action_taxable = st.number_input("Taxable Value", value=float(action_row["taxable_value"]))
                            action_rate = st.selectbox("GST Rate", GST_RATES, index=GST_RATES.index(float(action_row["gst_rate"])) if float(action_row["gst_rate"]) in GST_RATES else 0)
                            action_type = st.selectbox("Transaction Type", TRANSACTION_TYPES, index=TRANSACTION_TYPES.index(action_row["transaction_type"]) if action_row["transaction_type"] in TRANSACTION_TYPES else 0)
                            action_itc = st.selectbox("ITC Eligible?", ["Yes", "No"], index=0 if action_row.get("itc_eligible", "Yes") == "Yes" else 1)
                            action_category = st.selectbox("Supply Category", SUPPLY_CATEGORIES, index=SUPPLY_CATEGORIES.index(action_row.get("supply_category", "B2B")) if action_row.get("supply_category", "B2B") in SUPPLY_CATEGORIES else 0)
                            action_supply_type = st.selectbox("Supply Type", SUPPLY_TYPES, index=SUPPLY_TYPES.index(action_row.get("supply_type", "Goods")) if action_row.get("supply_type", "Goods") in SUPPLY_TYPES else 0)
                            action_save = st.form_submit_button("💾 Save Purchase Changes", type="primary")

                        if action_save:
                            cgst, sgst, igst, total = calculate_gst(action_taxable, action_rate, action_type)
                            create_backup()
                            update_purchase(action_id, action_supplier, action_gstin.upper(), action_invoice_number, action_invoice_date, action_taxable, action_rate, action_type, cgst, sgst, igst, total, action_itc, action_category, action_supply_type)
                            st.success("✅ Purchase invoice updated.")
                            st.rerun()

                elif register_action == "Delete":
                    if not selected_register_ids:
                        st.warning("Select at least one purchase invoice to delete.")
                    else:
                        st.warning(f"Deleting {len(selected_register_ids)} selected purchase invoice(s) is permanent.")
                        action_confirm = st.checkbox("I confirm deletion of the selected purchase invoice(s).", key="purchase_register_delete_confirm")
                        if st.button("🗑️ Delete Selected Purchase Invoice(s)", key="purchase_register_delete_button"):
                            if not action_confirm:
                                st.error("Please confirm deletion.")
                            else:
                                create_backup()
                                delete_records_by_ids("purchase_invoices", selected_register_ids)
                                st.success(f"✅ Deleted {len(selected_register_ids)} purchase invoice(s).")
                                st.rerun()

                st.caption(f"Showing {len(filtered)} of {len(purchase_data)} purchase invoices.")

        with tab_notes:
            render_credit_debit_notes("Purchase")

        # ---- ADD SINGLE 2B ENTRY ----
        with tab_2b_add:

            with st.form("add_2b"):
                col1, col2 = st.columns(2)

                with col1:
                    b_supplier = st.text_input("Supplier Name")
                    b_gstin = st.text_input("Supplier GSTIN *").upper()
                    b_invoice_number = st.text_input("Invoice Number *")
                    b_invoice_date = st.date_input("Invoice Date", key="b_date")

                with col2:
                    b_taxable = st.number_input("Taxable Value", min_value=0.0, step=100.0, key="b_taxable")
                    b_cgst = st.number_input("CGST", min_value=0.0, step=10.0, key="b_cgst")
                    b_sgst = st.number_input("SGST", min_value=0.0, step=10.0, key="b_sgst")
                    b_igst = st.number_input("IGST", min_value=0.0, step=10.0, key="b_igst")

                b_period = st.text_input("Return Period (YYYY-MM) *", value=datetime.now().strftime("%Y-%m"))

                save_2b = st.form_submit_button("💾 Save 2B Entry", use_container_width=True, type="primary")

            if save_2b:
                if not b_gstin or not b_invoice_number or not b_period:
                    st.error("GSTIN, Invoice Number and Return Period are required.")
                else:
                    insert_2b_record(b_supplier, b_gstin, b_invoice_number, b_invoice_date,
                                      b_taxable, b_cgst, b_sgst, b_igst, b_period)
                    st.success("✅ GSTR-2B entry saved.")

        # ---- BULK UPLOAD 2B ----
        with tab_2b_upload:
            st.markdown('<div class="gst-card">', unsafe_allow_html=True)

            st.write(
                "Upload the GSTR-2B data exported from the GST portal as a CSV or Excel file. "
                "Confirm the column mapping below, then import into a specific return period."
            )

            upload_period = st.text_input("Return Period for this upload (YYYY-MM) *", value=datetime.now().strftime("%Y-%m"), key="upload_period")

            clean_df = excel_csv_upload_and_map(
                "Choose a CSV or Excel file",
                "gstr2b_upload",
                required_fields=["gstin", "invoice_number", "invoice_date", "taxable_value"],
                help_text="",
                party_field="supplier_name",
                party_label="Supplier / Trade-Legal Name"
            )

            if clean_df is not None:
                if not upload_period.strip():
                    st.error("Please enter the Return Period (YYYY-MM) before importing.")
                else:
                    duplicate_ids = get_duplicate_import_ids(clean_df, "gstr_2b", upload_period)
                    period_mismatch_rows = get_period_mismatch_rows(clean_df, upload_period)
                    period_mismatch_confirmed = period_mismatch_rows.empty

                    if not period_mismatch_rows.empty:
                        st.warning(
                            f"{len(period_mismatch_rows)} uploaded invoice(s) do not belong to the selected period "
                            f"{upload_period}. Check the invoice dates before importing."
                        )
                        st.dataframe(
                            period_mismatch_rows[
                                [column for column in ["gstin", "invoice_number", "invoice_date"]
                                 if column in period_mismatch_rows.columns]
                            ],
                            use_container_width=True,
                            hide_index=True
                        )
                        period_mismatch_confirmed = st.checkbox(
                            "I still want to import these records into the selected period.",
                            key="gstr2b_period_mismatch_confirm"
                        )

                    overwrite_existing = False
                    if duplicate_ids:
                        import_choice = st.radio(
                            f"{len(duplicate_ids)} GSTR-2B record(s) for {upload_period} already exist. What should I do?",
                            ["Keep previous data and include this upload", "Overwrite existing matching data"],
                            key="gstr2b_import_choice"
                        )
                        overwrite_existing = import_choice == "Overwrite existing matching data"

                    if st.button("📥 Import into GSTR-2B", use_container_width=True, type="primary", key="gstr2b_import_btn"):
                        if not period_mismatch_confirmed:
                            st.error("Please confirm the period mismatch before importing.")
                        elif overwrite_existing:
                            create_backup()
                            overwrite_import_records(duplicate_ids, "gstr_2b")
                        if period_mismatch_confirmed:
                            count, skipped = bulk_insert_2b(clean_df, upload_period)
                            show_import_results(count, skipped, imported_noun=f"GSTR-2B record(s) for {upload_period}")
                            if count and not skipped:
                                st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        # ---- 2B REGISTER ----
        with tab_2b_register:

            gstr2b_data = load_2b_data()

            if gstr2b_data.empty:
                st.info("No GSTR-2B data loaded yet.")
            else:
                period_options = ["All"] + sorted(gstr2b_data["return_period"].dropna().unique().tolist(), reverse=True)
                period_filter = st.selectbox("Filter by Return Period", period_options)

                display_2b = gstr2b_data if period_filter == "All" else gstr2b_data[gstr2b_data["return_period"] == period_filter]
                register_table = display_2b.copy()
                select_all_2b = st.checkbox("Select all visible GSTR-2B records", key="select_all_2b")
                register_table.insert(0, "Select", select_all_2b)
                edited_register_table = st.data_editor(
                    register_table,
                    use_container_width=True,
                    hide_index=True,
                    key="gstr2b_register_action_table",
                    column_config={
                        "Select": st.column_config.CheckboxColumn("Select", default=False),
                        "id": st.column_config.NumberColumn("ID", disabled=True),
                        "itc_eligible": st.column_config.SelectboxColumn(
                            "ITC Status", options=["Eligible", "Ineligible", "Ineligible but Claimed"], required=True
                        ),
                        "itc_remark": st.column_config.TextColumn("ITC Remark"),
                    },
                    disabled=[
                        column for column in register_table.columns
                        if column not in ["Select", "itc_eligible", "itc_remark"]
                    ],
                )
                selected_register_ids = edited_register_table.loc[
                    edited_register_table["Select"], "id"
                ].astype(int).tolist()
                if st.button("💾 Save ITC Status & Remarks", key="gstr2b_save_itc_button"):
                    create_backup()
                    for _, itc_row in edited_register_table.iterrows():
                        original_row = gstr2b_data[gstr2b_data["id"] == int(itc_row["id"])].iloc[0]
                        update_2b_record(
                            int(itc_row["id"]), original_row["supplier_name"],
                            original_row["gstin"], original_row["invoice_number"],
                            original_row["invoice_date"], original_row["taxable_value"],
                            original_row["cgst"], original_row["sgst"], original_row["igst"],
                            original_row["return_period"],
                            str(itc_row.get("itc_eligible", "Eligible") or "Eligible"),
                            str(itc_row.get("itc_remark", "") or "")
                        )
                    st.success("✅ ITC status and remarks saved.")
                    st.rerun()
                register_action = st.selectbox(
                    "Action for selected GSTR-2B record",
                    ["Choose action", "Edit", "Delete"],
                    key="gstr2b_register_action"
                )

                if register_action == "Edit":
                    if len(selected_register_ids) != 1:
                        st.warning("Select exactly one GSTR-2B record to edit.")
                    else:
                        action_id = selected_register_ids[0]
                        action_row = gstr2b_data[gstr2b_data["id"] == action_id].iloc[0]
                        with st.form("gstr2b_register_edit_form"):
                            action_supplier = st.text_input("Supplier Name", value=action_row["supplier_name"] or "")
                            action_gstin = st.text_input("Supplier GSTIN", value=action_row["gstin"])
                            action_invoice_number = st.text_input("Invoice Number", value=action_row["invoice_number"])
                            action_invoice_date = st.date_input("Invoice Date", value=safe_date_input_value(action_row["invoice_date"]))
                            action_taxable = st.number_input("Taxable Value", value=float(action_row["taxable_value"]))
                            action_cgst = st.number_input("CGST", value=float(action_row["cgst"]))
                            action_sgst = st.number_input("SGST", value=float(action_row["sgst"]))
                            action_igst = st.number_input("IGST", value=float(action_row["igst"]))
                            action_period = st.text_input("Return Period (YYYY-MM)", value=action_row["return_period"])
                            action_itc = st.selectbox(
                                "ITC Status", ["Eligible", "Ineligible", "Ineligible but Claimed"],
                                index=["Eligible", "Ineligible", "Ineligible but Claimed"].index(action_row.get("itc_eligible", "Eligible"))
                                if action_row.get("itc_eligible", "Eligible") in ["Eligible", "Ineligible", "Ineligible but Claimed"] else 0
                            )
                            action_remark = st.text_input(
                                "ITC Remark", value=action_row.get("itc_remark", "") or ""
                            )
                            action_save = st.form_submit_button("💾 Save GSTR-2B Changes", type="primary")

                        if action_save:
                            update_2b_record(action_id, action_supplier, action_gstin.upper(), action_invoice_number, action_invoice_date, action_taxable, action_cgst, action_sgst, action_igst, action_period, action_itc, action_remark)
                            st.success("✅ GSTR-2B record updated.")
                            st.rerun()

                elif register_action == "Delete":
                    if not selected_register_ids:
                        st.warning("Select at least one GSTR-2B record to delete.")
                    else:
                        st.warning(f"Deleting {len(selected_register_ids)} selected GSTR-2B record(s) is permanent.")
                        action_confirm = st.checkbox("I confirm deletion of the selected GSTR-2B record(s).", key="gstr2b_register_delete_confirm")
                        if st.button("🗑️ Delete Selected GSTR-2B Record(s)", key="gstr2b_register_delete_button"):
                            if not action_confirm:
                                st.error("Please confirm deletion.")
                            else:
                                create_backup()
                                delete_records_by_ids("gstr_2b", selected_register_ids)
                                st.success(f"✅ Deleted {len(selected_register_ids)} GSTR-2B record(s).")
                                st.rerun()


# ============================================================
# RETURNS MODULE (GSTR-1 / GSTR-3B)
# ============================================================

elif module == "📄 Returns (GSTR-1 / 3B)":

    st.markdown('<div class="main-title">📄 GSTR-1 & GSTR-3B</div>', unsafe_allow_html=True)

    if not get_db_path():
        st.error("Please select a storage folder or company from Settings first.")

    else:
        tab_1, tab_1_upload, tab_3b = st.tabs(["📄 GSTR-1 Summary", "📤 Upload GSTR-1 (Invoice-wise)", "📄 GSTR-3B"])

        # ---- GSTR-1 ----
        with tab_1:

            st.subheader("Add / Update GSTR-1 Summary")

            with st.form("gstr1_form"):
                col1, col2 = st.columns(2)

                with col1:
                    g1_period = st.text_input("Return Period (YYYY-MM) *", value=datetime.now().strftime("%Y-%m"), key="g1_period")
                    g1_taxable = st.number_input("Total Taxable Value", min_value=0.0, step=100.0, key="g1_taxable")

                with col2:
                    g1_tax = st.number_input("Total Tax (CGST+SGST+IGST)", min_value=0.0, step=100.0, key="g1_tax")
                    g1_count = st.number_input("Invoice Count", min_value=0, step=1, key="g1_count")

                g1_filed = st.date_input("Filed Date", key="g1_filed")
                g1_save = st.form_submit_button("💾 Save GSTR-1", use_container_width=True, type="primary")

            if g1_save:
                if not g1_period:
                    st.error("Return Period is required.")
                else:
                    save_gstr1(g1_period, g1_taxable, g1_tax, int(g1_count), g1_filed)
                    st.success(f"✅ GSTR-1 for {g1_period} saved.")
                    st.rerun()

            st.divider()
            st.subheader("GSTR-1 Summary Register")

            gstr1_summary_data = load_gstr1_summary()

            if gstr1_summary_data.empty:
                st.info("No GSTR-1 summaries recorded yet.")
            else:
                st.dataframe(gstr1_summary_data, use_container_width=True, hide_index=True)

                delete_g1_id = st.number_input("Enter ID to delete a GSTR-1 summary record (optional)", min_value=0, step=1, key="del_g1")
                if st.button("🗑️ Delete GSTR-1 Record", use_container_width=True) and delete_g1_id:
                    delete_gstr1(int(delete_g1_id))
                    st.success("✅ Record deleted.")
                    st.rerun()

        # ---- GSTR-1 BULK UPLOAD (invoice-wise) ----
        with tab_1_upload:
            st.markdown('<div class="gst-card">', unsafe_allow_html=True)

            st.write(
                "Upload the GSTR-1 invoice-wise (B2B) data exported from the GST portal as a CSV or Excel file. "
                "This populates the invoice-wise GSTR-1 records used for Sales vs GSTR-1 reconciliation."
            )

            g1_upload_period = st.text_input("Return Period for this upload (YYYY-MM) *", value=datetime.now().strftime("%Y-%m"), key="g1_upload_period")

            g1_clean_df = excel_csv_upload_and_map(
                "Choose a CSV or Excel file",
                "gstr1_invoices_upload",
                required_fields=["gstin", "invoice_number", "invoice_date", "taxable_value"],
                help_text="",
                party_field="customer_name",
                party_label="Customer / Recipient Name"
            )

            if g1_clean_df is not None:
                if not g1_upload_period.strip():
                    st.error("Please enter the Return Period (YYYY-MM) before importing.")
                else:
                    duplicate_ids = get_duplicate_import_ids(g1_clean_df, "gstr1_invoices", g1_upload_period)
                    period_mismatch_rows = get_period_mismatch_rows(g1_clean_df, g1_upload_period)
                    period_mismatch_confirmed = period_mismatch_rows.empty

                    if not period_mismatch_rows.empty:
                        st.warning(
                            f"{len(period_mismatch_rows)} uploaded invoice(s) do not belong to the selected period "
                            f"{g1_upload_period}. Check the invoice dates before importing."
                        )
                        st.dataframe(
                            period_mismatch_rows[
                                [column for column in ["gstin", "invoice_number", "invoice_date"]
                                 if column in period_mismatch_rows.columns]
                            ],
                            use_container_width=True,
                            hide_index=True
                        )
                        period_mismatch_confirmed = st.checkbox(
                            "I still want to import these records into the selected period.",
                            key="gstr1_period_mismatch_confirm"
                        )

                    overwrite_existing = False
                    if duplicate_ids:
                        import_choice = st.radio(
                            f"{len(duplicate_ids)} GSTR-1 invoice(s) for {g1_upload_period} already exist. What should I do?",
                            ["Keep previous data and include this upload", "Overwrite existing matching data"],
                            key="gstr1_import_choice"
                        )
                        overwrite_existing = import_choice == "Overwrite existing matching data"

                    if st.button("📥 Import GSTR-1 Invoices", use_container_width=True, type="primary", key="gstr1_import_btn"):
                        if not period_mismatch_confirmed:
                            st.error("Please confirm the period mismatch before importing.")
                        elif overwrite_existing:
                            create_backup()
                            overwrite_import_records(duplicate_ids, "gstr1_invoices")
                        if period_mismatch_confirmed:
                            count, skipped = bulk_insert_gstr1(g1_clean_df, g1_upload_period)
                            show_import_results(count, skipped, imported_noun=f"GSTR-1 invoice(s) for {g1_upload_period}")
                            if count and not skipped:
                                st.rerun()

            st.markdown('</div>', unsafe_allow_html=True)

        # ---- GSTR-3B ----
        with tab_3b:

            st.subheader("Add / Update GSTR-3B Summary")

            with st.form("gstr3b_form"):
                col1, col2 = st.columns(2)

                with col1:
                    g3_period = st.text_input("Return Period (YYYY-MM) *", value=datetime.now().strftime("%Y-%m"), key="g3_period")
                    g3_outward = st.number_input("Outward Taxable Value", min_value=0.0, step=100.0, key="g3_outward")
                    g3_output_tax = st.number_input("Output Tax Declared", min_value=0.0, step=100.0, key="g3_output")

                with col2:
                    g3_itc = st.number_input("ITC Claimed", min_value=0.0, step=100.0, key="g3_itc")
                    g3_itc_reversed = st.number_input("ITC Reversed", min_value=0.0, step=100.0, key="g3_itc_reversed")
                    g3_paid = st.number_input("Tax Paid in Cash", min_value=0.0, step=100.0, key="g3_paid")
                    g3_late_fee = st.number_input("Late Fee / Interest", min_value=0.0, step=10.0, key="g3_late")

                g3_filed = st.date_input("Filed Date", key="g3_filed")
                g3_save = st.form_submit_button("💾 Save GSTR-3B", use_container_width=True, type="primary")

            if g3_save:
                if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", g3_period.strip()):
                    st.error("Return Period must use the format YYYY-MM, for example 2026-08.")
                else:
                    create_backup()
                    save_gstr3b(g3_period.strip(), g3_outward, g3_output_tax, g3_itc,
                                g3_paid, g3_late_fee, g3_filed, g3_itc_reversed)
                    st.success(f"✅ GSTR-3B for {g3_period.strip()} saved successfully.")

            st.divider()
            st.subheader("GSTR-3B Register")

            gstr3b_data = load_gstr3b_data()

            if gstr3b_data.empty:
                st.info("No GSTR-3B summaries recorded yet.")
            else:
                register_table = gstr3b_data.copy()
                register_table.insert(0, "Select", False)
                edited_register_table = st.data_editor(
                    register_table,
                    use_container_width=True,
                    hide_index=True,
                    key="gstr3b_register_action_table",
                    column_config={
                        "Select": st.column_config.CheckboxColumn("Select", default=False),
                        "id": st.column_config.NumberColumn("ID", disabled=True),
                    },
                    disabled=[column for column in register_table.columns if column != "Select"],
                )
                selected_g3_ids = edited_register_table.loc[
                    edited_register_table["Select"], "id"
                ].astype(int).tolist()
                register_action = st.selectbox(
                    "Action for selected GSTR-3B record",
                    ["Choose action", "Edit", "Delete"],
                    key="gstr3b_register_action"
                )

                if register_action == "Edit":
                    if len(selected_g3_ids) != 1:
                        st.warning("Select exactly one GSTR-3B record to edit.")
                    else:
                        action_id = selected_g3_ids[0]
                        action_row = gstr3b_data[gstr3b_data["id"] == action_id].iloc[0]
                        with st.form("gstr3b_register_edit_form"):
                            edit_period = st.text_input("Return Period (YYYY-MM) *", value=str(action_row["return_period"]))
                            edit_outward = st.number_input("Outward Taxable Value", min_value=0.0, value=float(action_row["outward_taxable_value"] or 0), step=100.0)
                            edit_output_tax = st.number_input("Output Tax Declared", min_value=0.0, value=float(action_row["output_tax"] or 0), step=100.0)
                            edit_itc = st.number_input("ITC Claimed", min_value=0.0, value=float(action_row["itc_claimed"] or 0), step=100.0)
                            edit_itc_reversed = st.number_input("ITC Reversed", min_value=0.0, value=float(action_row.get("itc_reversed", 0) or 0), step=100.0)
                            edit_paid = st.number_input("Tax Paid in Cash", min_value=0.0, value=float(action_row["tax_paid_cash"] or 0), step=100.0)
                            edit_late_fee = st.number_input("Late Fee / Interest", min_value=0.0, value=float(action_row["late_fee"] or 0), step=10.0)
                            edit_filed = st.date_input("Filed Date", value=safe_date_input_value(action_row["filed_date"]))
                            edit_save = st.form_submit_button("💾 Save GSTR-3B Changes", type="primary")

                        if edit_save:
                            if not re.fullmatch(r"\d{4}-(0[1-9]|1[0-2])", edit_period.strip()):
                                st.error("Return Period must use the format YYYY-MM, for example 2026-08.")
                            else:
                                create_backup()
                                update_gstr3b(
                                    action_id, edit_period.strip(), edit_outward, edit_output_tax,
                                    edit_itc, edit_itc_reversed, edit_paid, edit_late_fee, edit_filed
                                )
                                st.success("✅ GSTR-3B record updated successfully.")

                elif register_action == "Delete":
                    if not selected_g3_ids:
                        st.warning("Select at least one GSTR-3B record to delete.")
                    else:
                        st.warning(f"Deleting {len(selected_g3_ids)} selected GSTR-3B record(s) is permanent.")
                        delete_confirmed = st.checkbox(
                            "I confirm deletion of the selected GSTR-3B record(s).",
                            key="gstr3b_register_delete_confirm"
                        )
                        if st.button("🗑️ Delete Selected GSTR-3B Record(s)", key="gstr3b_register_delete_button"):
                            if not delete_confirmed:
                                st.error("Please confirm deletion.")
                            else:
                                create_backup()
                                for record_id in selected_g3_ids:
                                    delete_gstr3b(record_id)
                                st.success(f"✅ Deleted {len(selected_g3_ids)} GSTR-3B record(s).")


# ============================================================
# RECONCILIATION MODULE
# ============================================================

elif module == "🔄 Reconciliation":

    st.markdown('<div class="main-title">🔄 Reconciliation</div>', unsafe_allow_html=True)

    if not get_db_path():
        st.error("Please select a storage folder or company from Settings first.")
    else:
        reconciliation_view = st.radio(
            "Choose reconciliation",
            ["Purchase vs GSTR-2B", "Sales vs GSTR-1", "Sales vs GSTR-3B"],
            horizontal=True,
            key="reconciliation_view"
        )
        rc1, rc2 = st.columns(2)
        with rc1:
            recon_from = st.date_input("From Date", value=date(datetime.now().year, 4, 1), key="recon_from_date")
        with rc2:
            recon_to = st.date_input("To Date", value=date.today(), key="recon_to_date")

        if recon_from > recon_to:
            st.error("'From Date' cannot be after 'To Date'.")
            st.stop()

        def _filter_date(df, date_col="invoice_date"):
            if df.empty or date_col not in df.columns:
                return df
            dates = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
            return df.loc[dates.between(pd.Timestamp(recon_from), pd.Timestamp(recon_to), inclusive="both")].copy()

        if reconciliation_view == "Purchase vs GSTR-2B":
            st.subheader("🛒 Purchase vs GSTR-2B")
            purchase_data = load_purchase_data(recon_from, recon_to)
            gstr2b_data = _filter_date(load_2b_data())
            tolerance = st.number_input(
                "Amount Tolerance (₹)", min_value=0.0, value=1.0, step=0.50,
                help="Differences within this amount are treated as matched.", key="tolerance_p2b"
            )
            recon = reconcile_purchase_2b(purchase_data, gstr2b_data, tolerance)
            if recon.empty:
                st.info("No Purchase Register or GSTR-2B records exist in the selected period.")
            else:
                counts = recon["status"].value_counts()
                a,b,c,d,e = st.columns(5)
                a.metric("Matched", int(counts.get("Matched", 0)))
                b.metric("Amount Mismatch", int(counts.get("Amount Mismatch", 0)))
                c.metric("GSTIN Mismatch", int(counts.get("GSTIN Mismatch", 0) + counts.get("GSTIN Mismatch + Amount Mismatch", 0)))
                d.metric("Missing in 2B", int(counts.get("Missing in 2b", 0)))
                e.metric("Missing in Books", int(counts.get("Missing in purchase", 0)))
                status_filter = st.multiselect("Filter by Status", sorted(recon["status"].dropna().unique()), key="p2b_status_filter")
                display_recon = recon if not status_filter else recon[recon["status"].isin(status_filter)]
                cols = {
                    "match_key":"GSTIN / Invoice Key", "match_basis":"Match Basis", "status":"Status",
                    "supplier_name_purchase":"Purchase Supplier", "supplier_name_2b":"GSTR-2B Supplier",
                    "invoice_number_purchase":"Purchase Invoice No.", "invoice_number_2b":"GSTR-2B Invoice No.",
                    "gstin_purchase":"Purchase GSTIN", "gstin_2b":"GSTR-2B GSTIN",
                    "invoice_date_purchase":"Purchase Date", "invoice_date_2b":"2B Date",
                    "taxable_value_purchase":"Purchase Taxable", "taxable_value_2b":"2B Taxable", "taxable_diff":"Taxable Difference",
                    "purchase_total_tax":"Purchase ITC", "gstr2b_total_tax":"2B ITC", "tax_diff":"Tax Difference",
                    "source_rows_purchase":"Purchase Rows", "source_rows_2b":"2B Rows"
                }
                table = display_recon[[c for c in cols if c in display_recon.columns]].rename(columns=cols)
                st.dataframe(table, use_container_width=True, hide_index=True)
                st.caption("Multiple portal rows for one invoice are aggregated before matching. GSTIN mismatch is a review flag, not an automatic match.")

        elif reconciliation_view == "Sales vs GSTR-1":
            st.subheader("🧾 Sales vs GSTR-1")
            sales_data = load_data(recon_from, recon_to)
            gstr1_data = _filter_date(load_gstr1_data())
            tolerance = st.number_input("Amount Tolerance (₹)", min_value=0.0, value=1.0, step=0.50, key="tolerance_s1")
            recon = reconcile_sales_gstr1(sales_data, gstr1_data, tolerance)
            if recon.empty:
                st.info("No Sales Register or GSTR-1 records exist in the selected period.")
            else:
                counts = recon["status"].value_counts()
                a,b,c,d = st.columns(4)
                a.metric("Matched", int(counts.get("Matched", 0)))
                b.metric("Amount Mismatch", int(counts.get("Amount Mismatch", 0)))
                c.metric("GSTIN Mismatch", int(counts.get("GSTIN Mismatch", 0) + counts.get("GSTIN Mismatch + Amount Mismatch", 0)))
                d.metric("Missing / Filing Gap", int(sum(counts.get(x, 0) for x in ["Missing in gstr1", "Missing in sales"])))
                cols = {
                    "match_key":"GSTIN / Invoice Key", "match_basis":"Match Basis", "status":"Status",
                    "supplier_name_sales":"Sales Party", "supplier_name_gstr1":"GSTR-1 Party",
                    "invoice_number_sales":"Sales Invoice No.", "invoice_number_gstr1":"GSTR-1 Invoice No.",
                    "gstin_sales":"Sales GSTIN", "gstin_gstr1":"GSTR-1 GSTIN",
                    "invoice_date_sales":"Sales Date", "invoice_date_gstr1":"GSTR-1 Date",
                    "taxable_value_sales":"Sales Taxable", "taxable_value_gstr1":"GSTR-1 Taxable", "taxable_diff":"Taxable Difference",
                    "sales_total_tax":"Sales GST", "gstr1_total_tax":"GSTR-1 GST", "tax_diff":"Tax Difference"
                }
                table = recon[[c for c in cols if c in recon.columns]].rename(columns=cols)
                st.dataframe(table, use_container_width=True, hide_index=True)

        else:
            st.subheader("🧾 Sales vs GSTR-3B")
            sales_data = load_data(recon_from, recon_to)
            gstr3b_data = load_gstr3b_data()
            # 3B is period based, so use the selected date range by return period.
            period_start, period_end = recon_from.strftime("%Y-%m"), recon_to.strftime("%Y-%m")
            if not gstr3b_data.empty and "return_period" in gstr3b_data.columns:
                gstr3b_data = gstr3b_data[gstr3b_data["return_period"].between(period_start, period_end)].copy()
            tolerance = st.number_input("Amount Tolerance (₹)", min_value=0.0, value=1.0, step=0.50, key="tolerance_3b")
            recon3b = reconcile_3b_sales(gstr3b_data, sales_data, tolerance)
            if recon3b.empty:
                st.info("No Sales Register or GSTR-3B data exists in the selected period.")
            else:
                counts = recon3b["status"].value_counts()
                a,b,c,d = st.columns(4)
                a.metric("Matched", int(counts.get("Matched", 0)))
                b.metric("Amount Mismatch", int(counts.get("Amount Mismatch", 0)))
                c.metric("3B Not Filed", int(counts.get("GSTR-3B Not Filed", 0)))
                d.metric("Filed, No Sales", int(counts.get("Filed, No Sales Recorded", 0)))
                cols = {"return_period":"GST Period", "sales_taxable_value":"Sales Taxable", "sales_tax":"Sales GST", "invoice_count":"Invoice Count", "outward_taxable_value":"3B Taxable", "output_tax":"3B Output GST", "taxable_diff":"Taxable Difference", "tax_diff":"Tax Difference", "status":"Status"}
                st.dataframe(recon3b[[c for c in cols if c in recon3b.columns]].rename(columns=cols), use_container_width=True, hide_index=True)


# ============================================================
# REPORTS
# ============================================================

elif module == "📊 Reports":

    st.markdown('<div class="main-title">📊 GST Reports & Analysis</div>', unsafe_allow_html=True)

    if not get_db_path():
        st.error("Please select a storage folder or company from Settings first.")
    else:
        c1, c2 = st.columns(2)
        with c1:
            report_from = st.date_input("From Date", value=date(datetime.now().year, 4, 1), key="report_from_date")
        with c2:
            report_to = st.date_input("To Date", value=date.today(), key="report_to_date")

        if report_from > report_to:
            st.error("'From Date' cannot be after 'To Date'.")
            st.stop()

        sales_data = load_data(report_from, report_to)
        purchase_data = load_purchase_data(report_from, report_to)
        gstr2b_data = load_2b_data()
        gstr1_data = load_gstr1_data()
        gstr3b_data = load_gstr3b_data()

        def _period_filter(df):
            if df.empty or "invoice_date" not in df.columns:
                return df
            dates = pd.to_datetime(df["invoice_date"], errors="coerce", dayfirst=True)
            return df.loc[dates.between(pd.Timestamp(report_from), pd.Timestamp(report_to), inclusive="both")].copy()

        gstr2b_period = _period_filter(gstr2b_data)
        gstr1_period = _period_filter(gstr1_data)
        if not gstr3b_data.empty and "return_period" in gstr3b_data.columns:
            p1, p2 = report_from.strftime("%Y-%m"), report_to.strftime("%Y-%m")
            gstr3b_period = gstr3b_data[gstr3b_data["return_period"].between(p1, p2)].copy()
        else:
            gstr3b_period = gstr3b_data

        sales_tax = float((sales_data[["cgst","sgst","igst"]].sum().sum()) if not sales_data.empty else 0)
        purchase_itc = float((purchase_data[["cgst","sgst","igst"]].sum().sum()) if not purchase_data.empty else 0)
        net_gst = sales_tax - purchase_itc
        sales_taxable = float(sales_data["taxable_value"].sum()) if not sales_data.empty else 0
        purchase_taxable = float(purchase_data["taxable_value"].sum()) if not purchase_data.empty else 0
        sales_value = float(sales_data["total_amount"].sum()) if not sales_data.empty else 0
        purchase_value = float(purchase_data["total_amount"].sum()) if not purchase_data.empty else 0

        k1,k2,k3,k4,k5,k6 = st.columns(6)
        k1.metric("Sales Taxable", f"₹ {sales_taxable:,.0f}")
        k2.metric("Purchase Taxable", f"₹ {purchase_taxable:,.0f}")
        k3.metric("Output GST", f"₹ {sales_tax:,.0f}")
        k4.metric("Book ITC", f"₹ {purchase_itc:,.0f}")
        k5.metric("Net GST", f"₹ {net_gst:,.0f}")
        k6.metric("Gross Margin Proxy", f"₹ {sales_taxable-purchase_taxable:,.0f}")

        tab1, tab2, tab3, tab4 = st.tabs(["📋 GST Summary", "💰 Financial Analysis", "📈 Management Analysis", "🧾 ITC & Reconciliation"])

        with tab1:
            st.subheader("Period Summary")
            summary = pd.DataFrame({"Metric":["Sales invoices","Purchase invoices","Sales taxable value","Purchase taxable value","Output GST","Book ITC","Net GST payable / (credit)","Sales invoice value","Purchase invoice value"],"Value":[len(sales_data),len(purchase_data),sales_taxable,purchase_taxable,sales_tax,purchase_itc,net_gst,sales_value,purchase_value]})
            st.dataframe(summary, use_container_width=True, hide_index=True)

            if not sales_data.empty:
                rate = sales_data.groupby("gst_rate").agg(Invoices=("id","count"), Taxable_Value=("taxable_value","sum"), GST=("cgst","sum")).reset_index()
                rate["GST"] = sales_data.groupby("gst_rate")[["cgst","sgst","igst"]].sum().sum(axis=1).values
                st.subheader("Sales GST Rate-wise")
                st.dataframe(rate, use_container_width=True, hide_index=True)

            st.subheader("Monthly Trend")
            frames=[]
            for df, label in [(sales_data,"Sales"),(purchase_data,"Purchases")]:
                if not df.empty:
                    x=df.copy(); x["Month"]=x["invoice_date"].apply(to_period); x=x.groupby("Month")["taxable_value"].sum().rename(label); frames.append(x)
            if frames:
                trend=pd.concat(frames,axis=1).fillna(0).reset_index(); st.dataframe(trend, use_container_width=True, hide_index=True)
                if PLOTLY_AVAILABLE:
                    long=trend.melt(id_vars="Month",var_name="Type",value_name="Taxable Value")
                    st.plotly_chart(px.line(long,x="Month",y="Taxable Value",color="Type",markers=True),use_container_width=True)

        with tab2:
            st.subheader("Financial & Commercial Indicators")
            avg_sales = sales_taxable / len(sales_data) if len(sales_data) else 0
            avg_purchase = purchase_taxable / len(purchase_data) if len(purchase_data) else 0
            purchase_to_sales = purchase_taxable / sales_taxable * 100 if sales_taxable else 0
            gst_burden = net_gst / sales_taxable * 100 if sales_taxable else 0
            itc_coverage = purchase_itc / sales_tax * 100 if sales_tax else 0
            indicators = pd.DataFrame({"Indicator":["Purchase / Sales %","Gross Margin Proxy %","Net GST / Sales Taxable %","Book ITC / Output GST %","Average Sales Invoice","Average Purchase Invoice"],"Value":[purchase_to_sales,(sales_taxable-purchase_taxable)/sales_taxable*100 if sales_taxable else 0,gst_burden,itc_coverage,avg_sales,avg_purchase]})
            st.dataframe(indicators, use_container_width=True, hide_index=True)
            st.caption("Gross Margin Proxy is sales taxable value less purchase taxable value; it is not statutory gross profit and ignores opening/closing stock, other costs and accounting adjustments.")

            if not purchase_data.empty:
                vendor = purchase_data.groupby("supplier_name").agg(Invoices=("id","count"), Taxable_Value=("taxable_value","sum"), ITC=("cgst","sum")).reset_index()
                vendor["ITC"] = purchase_data.groupby("supplier_name")[["cgst","sgst","igst"]].sum().sum(axis=1).values
                vendor["Share %"] = vendor["Taxable_Value"] / purchase_taxable * 100 if purchase_taxable else 0
                st.subheader("Top Suppliers by Purchase Value")
                st.dataframe(vendor.sort_values("Taxable_Value",ascending=False).head(10),use_container_width=True,hide_index=True)

            if not sales_data.empty:
                customer_col = "supplier_name" if "supplier_name" in sales_data.columns else "customer_name"
                customer = sales_data.groupby(customer_col).agg(Invoices=("id","count"), Taxable_Value=("taxable_value","sum"), GST=("cgst","sum")).reset_index()
                customer["GST"] = sales_data.groupby(customer_col)[["cgst","sgst","igst"]].sum().sum(axis=1).values
                customer["Share %"] = customer["Taxable_Value"] / sales_taxable * 100 if sales_taxable else 0
                st.subheader("Top Customers by Sales Value")
                st.dataframe(customer.sort_values("Taxable_Value",ascending=False).head(10),use_container_width=True,hide_index=True)

        with tab3:
            st.subheader("Management Review")
            st.caption("These are decision-support indicators, not proof of fraud, dummy transactions, or accounting irregularity.")
            review=[]
            if purchase_taxable and not sales_taxable: review.append(("Sales conversion","Purchases exist but no sales are recorded in the period.","High"))
            elif sales_taxable and not purchase_taxable: review.append(("Supply coverage","Sales exist but no purchases are recorded; review opening stock / other sourcing.","High"))
            elif purchase_taxable and sales_taxable:
                ratio=purchase_taxable/sales_taxable
                review.append(("Purchase-to-sales","Purchases are {:.1f}% of sales taxable value.".format(ratio*100),"Review" if ratio>1 else "Normal"))
            if purchase_itc > sales_tax and sales_tax>0: review.append(("ITC vs output","Book ITC exceeds output GST for this period; cash liability may be low/credit may arise.","Review"))
            if not purchase_data.empty:
                dup=purchase_data.duplicated(["gstin","invoice_number"],keep=False).sum()
                if dup: review.append(("Duplicate purchase keys",f"{dup} purchase rows share GSTIN + invoice number; verify duplicate import / split invoice.","High"))
            if not sales_data.empty:
                dup=sales_data.duplicated(["gstin","invoice_number"],keep=False).sum()
                if dup: review.append(("Duplicate sales keys",f"{dup} sales rows share GSTIN + invoice number; verify duplicate import / split invoice.","High"))
            review_df=pd.DataFrame(review,columns=["Area","Observation","Priority"])
            st.dataframe(review_df if not review_df.empty else pd.DataFrame({"Area":["No exception flags"],"Observation":["No management exception was triggered by the available data."],"Priority":["Normal"]}),use_container_width=True,hide_index=True)

            monthly=[]
            if not sales_data.empty:
                x=sales_data.copy(); x["Month"]=x["invoice_date"].apply(to_period); monthly.append(x.groupby("Month")["taxable_value"].sum().rename("Sales"))
            if not purchase_data.empty:
                x=purchase_data.copy(); x["Month"]=x["invoice_date"].apply(to_period); monthly.append(x.groupby("Month")["taxable_value"].sum().rename("Purchases"))
            if monthly:
                m=pd.concat(monthly,axis=1).fillna(0); m["Sales/Purchase"] = np.where(m["Purchases"]!=0,m["Sales"]/m["Purchases"],np.nan) if "Sales" in m and "Purchases" in m else np.nan
                st.subheader("Monthly Management Trend"); st.dataframe(m.reset_index(),use_container_width=True,hide_index=True)

        with tab4:
            st.subheader("Purchase vs GSTR-2B")
            p2b = reconcile_purchase_2b(purchase_data, gstr2b_period, 1.0)
            if p2b.empty:
                st.info("No purchase / 2B data for this period.")
            else:
                counts=p2b["status"].value_counts(); st.write({"Matched":int(counts.get("Matched",0)),"Amount mismatch":int(counts.get("Amount Mismatch",0)),"GSTIN mismatch":int(counts.get("GSTIN Mismatch",0)+counts.get("GSTIN Mismatch + Amount Mismatch",0)),"Missing in 2B":int(counts.get("Missing in 2b",0)),"Missing in books":int(counts.get("Missing in purchase",0))})
                st.dataframe(p2b,use_container_width=True,hide_index=True)

            st.subheader("Sales vs GSTR-1")
            s1 = reconcile_sales_gstr1(sales_data, gstr1_period, 1.0)
            if not s1.empty: st.dataframe(s1,use_container_width=True,hide_index=True)

            st.subheader("Sales vs GSTR-3B")
            r3b = reconcile_3b_sales(gstr3b_period, sales_data, 1.0)
            if not r3b.empty: st.dataframe(r3b,use_container_width=True,hide_index=True)


# ============================================================
# EXPORT
# ============================================================

elif module == "📥 Export":

    st.markdown('<div class="main-title">📥 Export GST Data</div>', unsafe_allow_html=True)

    if not get_db_path():
        st.error("Please select a storage folder or company from Settings first.")

    else:
        st.subheader("1. GST Period")

        col1, col2 = st.columns(2)
        with col1:
            export_from = st.date_input("From Date", value=date(datetime.now().year, 4, 1))
        with col2:
            export_to = st.date_input("To Date", value=date.today())

        st.info(
            f"Report period: **{export_from:%Y-%m-%d} to {export_to:%Y-%m-%d}**"
        )
        st.caption("Sales and Purchase registers are filtered by this date range. Returns and 2B data are exported in full.")

        st.subheader("2. Choose What to Include")

        col3, col4, col5, col6 = st.columns(4)
        include_purchase = col3.checkbox("Purchase Register", value=True)
        include_2b = col4.checkbox("GSTR-2B", value=True)
        include_returns = col5.checkbox("GSTR-1 / 3B", value=True)
        include_reconciliation = col6.checkbox("Reconciliation", value=True)

        preview_data = load_data(export_from, export_to)
        st.write(f"Sales invoices in this period: **{len(preview_data)}**")

        if st.button("📊 Generate Professional Excel Report", use_container_width=True, type="primary"):

            if export_from > export_to:
                st.error("'From Date' cannot be after 'To Date'.")
            else:
                report_path = create_excel_report(
                    export_from, export_to,
                    include_purchase, include_2b, include_returns, include_reconciliation
                )

                if report_path:
                    st.success("✅ Professional Excel report created.")
                    st.info(f"Saved at: {report_path}")

                    with open(report_path, "rb") as file:
                        file_bytes = file.read()

                    st.download_button(
                        label="⬇️ Download Excel Report",
                        data=file_bytes,
                        file_name=os.path.basename(report_path),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                else:
                    st.warning("No sales or purchase data found for the selected period.")


# ============================================================
# BACKUP
# ============================================================

elif module == "💾 Backup":

    st.markdown('<div class="main-title">💾 GST Database Backup</div>', unsafe_allow_html=True)

    if get_db_path():

        if st.button("💾 Create Manual Backup", use_container_width=True):
            backup = create_backup()

            if backup:
                st.success("✅ Backup created successfully.")
                st.info(backup)
            else:
                st.warning("No database available.")

        st.divider()

        backup_root = st.session_state.save_path or os.path.dirname(get_db_path() or "")
        backup_folder = os.path.join(backup_root, "Backups") if backup_root else None

        if backup_folder and os.path.exists(backup_folder):
            backups = sorted(os.listdir(backup_folder), reverse=True)

            st.subheader("Available Backups")
            for backup in backups:
                st.write(f"💾 {backup}")

    else:
        st.warning("Select a storage folder or company first.")


# ============================================================
# SETTINGS
# ============================================================

elif module == "⚙️ Settings":

    st.markdown('<div class="main-title">⚙️ Application Settings</div>', unsafe_allow_html=True)

    st.subheader("🏢 Company Profile")

    if get_db_path():
        profile = load_company_profile()

        with st.form("company_profile_form"):
            cp_name = st.text_input("Company Name", value=profile.get("company_name", ""))
            cp_gstin = st.text_input("Company GSTIN", value=profile.get("gstin", "")).upper()
            cp_state = st.text_input("State", value=profile.get("state", ""))
            cp_fy = st.text_input("Financial Year (e.g. 2026-27)", value=profile.get("financial_year", ""))

            cp_save = st.form_submit_button("💾 Save Company Profile", use_container_width=True, type="primary")

        if cp_save:
            save_company_profile(cp_name, cp_gstin, cp_state, cp_fy)
            st.success("✅ Company profile saved.")
    else:
        st.info("Select a storage folder or load a company below to set up its profile.")

    st.divider()

    st.subheader("📁 Data & Backup Storage")

    record = st.session_state.client_record or {}
    current_data_folder = record.get("folder") or st.session_state.save_path or client_folder_path(st.session_state.auth_user)
    st.caption("The SQLite database and all database backups are stored in this folder.")

    with st.form("storage_path_form"):
        storage_path = st.text_input("Data folder path", value=current_data_folder)
        move_existing_data = st.checkbox(
            "Copy the existing database and backups to the new folder",
            value=True
        )
        storage_save = st.form_submit_button(
            "💾 Save Data Location", use_container_width=True, type="primary"
        )

    if storage_save:
        ok, message = update_client_storage_path(
            st.session_state.auth_user, storage_path, move_existing_data
        )
        if ok:
            updated_clients = load_clients()
            updated_record = updated_clients.get(st.session_state.auth_user, record)
            st.session_state.client_record = {**updated_record, "username": st.session_state.auth_user}
            st.session_state.save_path = os.path.abspath(os.path.expanduser(storage_path.strip()))
            st.session_state.active_db = ""
            os.makedirs(os.path.join(st.session_state.save_path, "Backups"), exist_ok=True)
            st.success(message)
            st.rerun()
        else:
            st.error(message)

    st.divider()

    st.subheader("👤 My Account")

    acc_col1, acc_col2 = st.columns(2)
    with acc_col1:
        st.write(f"**Username:** {st.session_state.auth_user}")
        st.write(f"**Client / Business Name:** {record.get('display_name', '—')}")
        st.write(f"**GSTIN on file:** {record.get('gstin') or '—'}")
    with acc_col2:
        st.write(f"**Email:** {record.get('email') or '—'}")
        st.write(f"**Phone:** {record.get('phone') or '—'}")
        st.write(f"**Registered on:** {record.get('created_on', '—')}")

    st.caption(f"📁 Your data folder: `{record.get('folder', st.session_state.save_path)}`")
    st.caption("Every client's data lives in its own private folder — nobody else can see or open it without your username and password.")

    with st.expander("🔑 Change Password"):
        with st.form("change_password_form"):
            cp_current = st.text_input("Current Password", type="password")
            cp_new = st.text_input("New Password", type="password")
            cp_confirm = st.text_input("Confirm New Password", type="password")
            cp_submit = st.form_submit_button("💾 Update Password", use_container_width=True, type="primary")

        if cp_submit:
            ok, msg = change_client_password(st.session_state.auth_user, cp_current, cp_new, cp_confirm)
            if ok:
                st.success(f"✅ {msg}")
            else:
                st.error(msg)

    st.divider()

    if st.button("🚪 Log Out", use_container_width=True):
        st.session_state.auth_user = None
        st.session_state.client_record = None
        st.session_state.save_path = ""
        st.session_state.active_db = ""
        st.rerun()

    st.divider()

    st.subheader("📄 Active Database")
    st.write(get_db_path() or "None selected")

    st.divider()

    st.subheader("ℹ️ Application Information")
    st.write("GST Data Management & Reconciliation Tool")
    st.write("Permanent SQLite storage per client + professional Excel reporting")
    if not PLOTLY_AVAILABLE:
        st.caption("Tip: install 'plotly' (pip install plotly) for richer dashboard charts.")


# ============================================================
# GSTR-1 JSON GENERATOR MODULE
# ============================================================

elif module == "🧾 GSTR-1 JSON Generator":

    gstr1_json_generator.render(get_db_path())


# ============================================================
# CREDITORS & DEBTORS MODULE
# ============================================================

elif module == "💼 Creditors & Debtors":

    gst_creditors_debtors.render(get_db_path())


# ============================================================
# GST COMPLIANCE MODULE
# ============================================================

elif module == "🛡️ GST Compliance":

    gst_compliance.render(get_db_path())


# ============================================================
# GSTR-9 / 9C MODULE
# ============================================================

elif module == "📑 GSTR-9":

    gstr9_9c.render(get_db_path(), load_company_profile, section="gstr9")


# ============================================================
# GSTR-9C MODULE
# ============================================================

elif module == "📋 GSTR-9C":

    gstr9_9c.render(get_db_path(), load_company_profile, section="gstr9c")