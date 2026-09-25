"""
Sweet Mayhem — Order Fulfillment Web App
"""

import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import gspread
import requests
from google.oauth2.service_account import Credentials
from difflib import SequenceMatcher
from datetime import datetime, date, timedelta
from copy import deepcopy
from urllib.parse import quote as url_quote
import io
import math
import re
import html as html_lib
import functools
import secrets
import uuid
import hashlib
import bcrypt
import openpyxl
from openpyxl.styles import Font
from googleapiclient.discovery import build

# ─── Page Config ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Sweet Mayhem — Fulfillment",
    page_icon=":material/storefront:",
    layout="wide",
)

# ─── Constants ────────────────────────────────────────────────────────────────

SHEET_ID = "1t_L1qR3ikD-jjiA2P1tQETQ60hvR1nocX5AimPdSJ7o"
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# ─── Google Sheets ────────────────────────────────────────────────────────────

def _service_account_creds():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    return Credentials.from_service_account_info(creds_dict, scopes=SCOPES)

@st.cache_resource
def _gc():
    return gspread.authorize(_service_account_creds())

@st.cache_resource
def _drive():
    return build("drive", "v3", credentials=_service_account_creds())

@st.cache_resource
def _spreadsheet():
    return _gc().open_by_key(SHEET_ID)

def _resilient_google_call(fn):
    """Retries once on a transient network error, reconnecting fresh clients
    first. Streamlit Cloud's process can sit idle between requests, and a
    pooled connection Google's end has since closed then surfaces on the next
    call as a low-level OSError (e.g. 'Broken pipe') rather than a clean
    Google API error."""
    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except (BrokenPipeError, ConnectionError, OSError):
            _gc.clear()
            _drive.clear()
            _spreadsheet.clear()
            _shipment_tracker_spreadsheet.clear()
            _refunds_spreadsheet.clear()
            _cancelled_orders_spreadsheet.clear()
            return fn(*args, **kwargs)
    return wrapper

@_resilient_google_call
def get_ws():
    return _spreadsheet().get_worksheet(0)

# ─── Users & Access Control ────────────────────────────────────────────────────

USERS_SHEET_NAME = "Users"
ALL_PAGES = [
    "📦 Fulfillment", "🔄 Restock", "➕ Add Product", "📋 View Inventory",
    "📊 Demand & Reorder", "📥 Purchase Orders", "🚢 Shipment Tracker", "🧾 Shipment Details",
    "🚫 Cancelled Orders", "💸 Refunds",
]
ADMIN_PAGE = "👤 Manage Users"

# "Save Desk" was this page's old name — some users' saved Permissions cells
# may still have the old identity string. Translated on load (below) so
# nobody's page access silently breaks; never written back automatically.
PAGE_KEY_MIGRATIONS = {
    "🛟 Save Desk": "🚫 Cancelled Orders",
}

# Page identity strings above (with emoji) are the stored keys used in existing users'
# saved Permissions cells — keep them unchanged. This maps each to a real icon for
# display only.
PAGE_ICONS = {
    "📦 Fulfillment": ":material/local_shipping:",
    "🔄 Restock": ":material/inventory_2:",
    "➕ Add Product": ":material/add_box:",
    "📋 View Inventory": ":material/list_alt:",
    "📊 Demand & Reorder": ":material/insights:",
    "📥 Purchase Orders": ":material/move_to_inbox:",
    "🚫 Cancelled Orders": ":material/cancel:",
    "💸 Refunds": ":material/currency_exchange:",
    "🚢 Shipment Tracker": ":material/directions_boat:",
    "🧾 Shipment Details": ":material/receipt_long:",
    ADMIN_PAGE: ":material/group:",
}

def page_label(p):
    return p.split(" ", 1)[1] if " " in p else p

# Lets the current page survive a real browser refresh (which starts a brand
# new Streamlit session, wiping session_state) by round-tripping it through a
# URL query param instead.
PAGE_SLUGS = {p: re.sub(r"[^a-z0-9]+", "-", page_label(p).lower()).strip("-") for p in ALL_PAGES + [ADMIN_PAGE]}
SLUG_TO_PAGE = {v: k for k, v in PAGE_SLUGS.items()}

@_resilient_google_call
def get_users_ws():
    sh = _spreadsheet()
    try:
        return sh.worksheet(USERS_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=USERS_SHEET_NAME, rows=100, cols=5)
        ws.append_row(["Username", "PasswordHash", "Role", "Permissions", "CreatedAt"])
        return ws

@st.cache_data(ttl=60)
@_resilient_google_call
def load_users():
    ws = get_users_ws()
    data = ws.get_all_values()
    users = {}
    for i, row in enumerate(data[1:], start=2):
        if len(row) < 4 or not row[0].strip():
            continue
        username = row[0].strip()
        users[username.lower()] = {
            "username": username,
            "password_hash": row[1],
            "role": row[2].strip().lower(),
            "permissions": [
                PAGE_KEY_MIGRATIONS.get(p.strip(), p.strip())
                for p in row[3].split(",") if p.strip()
            ],
            "remember_hash": row[5] if len(row) > 5 else "",
            "row": i,
        }
    return users

def hash_password(password):
    return bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode()

def verify_password(password, password_hash):
    try:
        return bcrypt.checkpw(password.encode(), password_hash.encode())
    except (ValueError, TypeError):
        return False

# Remember-me tokens are high-entropy random secrets, not human-chosen passwords —
# a fast hash (unlike bcrypt) is fine here and lets lookup-by-token stay cheap.
def _hash_token(token):
    return hashlib.sha256(token.encode()).hexdigest()

def issue_remember_token(row):
    token = secrets.token_urlsafe(32)
    get_users_ws().update_cell(row, 6, _hash_token(token))
    load_users.clear()
    return token

def revoke_remember_token(row):
    get_users_ws().update_cell(row, 6, "")
    load_users.clear()

def create_user(username, password, role, permissions):
    ws = get_users_ws()
    ws.append_row([
        username, hash_password(password), role, ",".join(permissions),
        datetime.now().strftime("%Y-%m-%d %H:%M"), "",
    ])
    load_users.clear()

def update_user(row, role=None, permissions=None, password=None):
    ws = get_users_ws()
    if role is not None:
        ws.update_cell(row, 3, role)
    if permissions is not None:
        ws.update_cell(row, 4, ",".join(permissions))
    if password is not None:
        ws.update_cell(row, 2, hash_password(password))
        # A password change invalidates any existing remember-me link for this account.
        ws.update_cell(row, 6, "")
    load_users.clear()

def delete_user(row):
    get_users_ws().delete_rows(row)
    load_users.clear()

def effective_pages(user):
    if user["role"] == "admin":
        return ALL_PAGES + [ADMIN_PAGE]
    return [p for p in user["permissions"] if p in ALL_PAGES]

# ─── Refunds ──────────────────────────────────────────────────────────────────
# Own separate spreadsheet (like Shipment Tracker's), not a tab inside Users' —
# one row per refund request, shared across every device, unlike the old
# localStorage-based tool, so Khawla, Sacha, and admins all see the same list
# no matter what they're signed in on.

REFUNDS_SHEET_ID = "1e24MtmjE7BRSeZ_WKGbR58BEt9DO1XmR-ykE03F-pjI"
REFUNDS_SHEET_NAME = "Refunds"
REFUND_STATUSES = ["Pending", "Refunded", "Rejected"]

@st.cache_resource
def _refunds_spreadsheet():
    return _gc().open_by_key(REFUNDS_SHEET_ID)

@_resilient_google_call
def get_refunds_ws():
    sh = _refunds_spreadsheet()
    try:
        return sh.worksheet(REFUNDS_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=REFUNDS_SHEET_NAME, rows=200, cols=7)
        ws.append_row(["Order #", "Customer Name", "Amount", "Status", "Whish Number", "Logged By", "Date Logged"])
        return ws

def _parse_amount(v):
    try:
        return float(str(v).strip())
    except (ValueError, TypeError):
        return None

@st.cache_data(ttl=30)
@_resilient_google_call
def load_refunds():
    ws = get_refunds_ws()
    data = ws.get_all_values()
    refunds = []
    for i, row in enumerate(data[1:], start=2):
        if len(row) < 1 or not row[0].strip():
            continue
        status = row[3].strip() if len(row) > 3 else ""
        refunds.append({
            "row": i,
            "order": row[0].strip(),
            "customer": row[1].strip() if len(row) > 1 else "",
            "amount": _parse_amount(row[2]) if len(row) > 2 else None,
            "status": status if status in REFUND_STATUSES else "Pending",
            "whish": row[4].strip() if len(row) > 4 else "",
            "logged_by": row[5].strip() if len(row) > 5 else "",
            "date_logged": row[6].strip() if len(row) > 6 else "",
        })
    return refunds

def add_refund(order, customer, amount, whish, logged_by):
    ws = get_refunds_ws()
    ws.append_row([order, customer, amount, "Pending", whish, logged_by, datetime.now().strftime("%Y-%m-%d %H:%M")])
    load_refunds.clear()

def update_refund_status(row, status):
    get_refunds_ws().update_cell(row, 4, status)
    load_refunds.clear()

def delete_refund(row):
    get_refunds_ws().delete_rows(row)
    load_refunds.clear()

# ─── Cancelled Orders ───────────────────────────────────────────────────────────
# Own separate spreadsheet, same reasoning as Refunds: the old tool kept
# everything in browser localStorage, so an admin's upload and an employee's
# WhatsApp/call/coupon progress were only ever visible on whichever single
# device did the work. One shared Sheet fixes that the same way.

CANCELLED_ORDERS_SHEET_ID = "14avUX9udcZ4CYDGOUTI8omKpAOlBeotfV_ONCToO-Pk"
CANCELLED_ORDERS_TAB = "Cancelled Orders"
CANCELLED_ORDERS_SETTINGS_TAB = "Settings"
CO_STAGE_ORDER = ["wa1", "call", "coupon"]
CO_STAGE_LABELS = {"wa1": "1st WhatsApp", "call": "WhatsApp call", "coupon": "Coupon WhatsApp"}
CO_DEFAULT_TEMPLATE_1 = "Hi {firstname}, this is {employee} 👋 I noticed that you cancelled your order for {items} (total {total}) — can you tell me what's the problem?"
CO_DEFAULT_TEMPLATE_COUPON = "Hi {firstname}, this is {employee} 👋 Use code \"{code}\" for {percent}% off, the code is valid for 72 hours!"
CO_COLUMNS = [
    "Order ID", "Reference ID", "Customer Name", "Phone", "Address", "Note",
    "USD Price", "Delivery USD", "Total USD", "Payment Status", "Creation Date",
    "Assigned To", "Stage", "Outcome", "Notes", "Date Added",
]
CO_SETTINGS_COLUMNS = [
    "Employee1", "Employee1Weight", "Employee2", "Employee2Weight",
    "RRState1", "RRState2", "CouponCode", "CouponPercent",
    "WaTemplate1", "WaTemplateCoupon",
]

@st.cache_resource
def _cancelled_orders_spreadsheet():
    return _gc().open_by_key(CANCELLED_ORDERS_SHEET_ID)

@_resilient_google_call
def get_co_ws():
    sh = _cancelled_orders_spreadsheet()
    try:
        return sh.worksheet(CANCELLED_ORDERS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=CANCELLED_ORDERS_TAB, rows=1000, cols=len(CO_COLUMNS))
        ws.append_row(CO_COLUMNS)
        return ws

@_resilient_google_call
def get_co_settings_ws():
    sh = _cancelled_orders_spreadsheet()
    try:
        return sh.worksheet(CANCELLED_ORDERS_SETTINGS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=CANCELLED_ORDERS_SETTINGS_TAB, rows=2, cols=len(CO_SETTINGS_COLUMNS))
        ws.append_row(CO_SETTINGS_COLUMNS)
        ws.append_row(["Khawla", "1", "Sacha", "2", "0", "0", "Comeback15", "15", CO_DEFAULT_TEMPLATE_1, CO_DEFAULT_TEMPLATE_COUPON])
        return ws

@st.cache_data(ttl=30)
@_resilient_google_call
def load_co_settings():
    ws = get_co_settings_ws()
    data = ws.get_all_values()
    if len(data) < 2:
        return {
            "employees": ["Khawla", "Sacha"], "weights": [1.0, 2.0], "rr_state": [0.0, 0.0],
            "coupon_code": "Comeback15", "coupon_percent": 15.0,
            "wa_template_1": CO_DEFAULT_TEMPLATE_1, "wa_template_coupon": CO_DEFAULT_TEMPLATE_COUPON,
        }
    row = data[1]
    def _f(idx, default=0.0):
        try:
            return float(row[idx])
        except (IndexError, ValueError):
            return default
    return {
        "employees": [row[0].strip() or "Khawla", row[2].strip() or "Sacha"] if len(row) > 2 else ["Khawla", "Sacha"],
        "weights": [_f(1, 1.0), _f(3, 2.0)],
        "rr_state": [_f(4, 0.0), _f(5, 0.0)],
        "coupon_code": (row[6].strip() if len(row) > 6 and row[6].strip() else "Comeback15"),
        "coupon_percent": _f(7, 15.0),
        "wa_template_1": (row[8] if len(row) > 8 and row[8].strip() else CO_DEFAULT_TEMPLATE_1),
        "wa_template_coupon": (row[9] if len(row) > 9 and row[9].strip() else CO_DEFAULT_TEMPLATE_COUPON),
    }

def save_co_settings(settings):
    ws = get_co_settings_ws()
    ws.update([[
        settings["employees"][0], settings["weights"][0],
        settings["employees"][1], settings["weights"][1],
        settings["rr_state"][0], settings["rr_state"][1],
        settings["coupon_code"], settings["coupon_percent"],
        settings["wa_template_1"], settings["wa_template_coupon"],
    ]], range_name="A2:J2")
    load_co_settings.clear()

# Weighted round-robin (nginx/LVS "smooth" algorithm): each pick bumps every
# employee's running weight by their share, hands the order to whoever's
# highest, then knocks the total back off them — evenly interleaves a 1:2
# split (S,K,S,S,K,S…) instead of clumping one person's share in batches.
def co_next_assignee(settings):
    emps, weights, rr = settings["employees"], settings["weights"], settings["rr_state"]
    total = sum(weights) or len(emps)
    best = 0
    for i in range(len(emps)):
        rr[i] += weights[i] or 0
        if rr[i] > rr[best]:
            best = i
    rr[best] -= total
    return emps[best]

def co_order_number(order_id, reference_id):
    ref = str(reference_id or "").strip()
    m = re.search(r"(\d+)\s*$", ref)
    return "1" + (m.group(1) if m else (ref or order_id))

CO_AVATAR_COLORS = ["#D9531E", "#2A5C8A", "#227A55", "#8A5CA8"]

def co_avatar_color(name, employees):
    idx = employees.index(name) if name in employees else 0
    return CO_AVATAR_COLORS[idx % len(CO_AVATAR_COLORS)]

# Roadrunner packs items into one note like "(1) Product, Color / S sku: ABC-1.
# (1) Other Product / M. (1) Third Product sku: XYZ-3." — the sku suffix is
# inconsistent (some items have it, some don't), so split on each "(qty)"
# marker rather than anchoring the match on "sku:", or items without a sku
# silently vanish instead of falling back.
def co_parse_items(note):
    text = str(note or "").strip()
    if not text:
        return []
    starts = [m.start() for m in re.finditer(r"\(\d+\)", text)]
    if not starts:
        return [text]
    items = []
    for i, start in enumerate(starts):
        end = starts[i + 1] if i + 1 < len(starts) else len(text)
        seg = text[start:end].strip()
        qm = re.match(r"^\((\d+)\)\s*", seg)
        qty = qm.group(1) if qm else "1"
        desc = seg[qm.end():] if qm else seg
        desc = re.sub(r"\s*sku:\s*\S+\s*$", "", desc, flags=re.IGNORECASE)
        desc = re.sub(r"[.\s]+$", "", desc).strip()
        if desc:
            items.append(f"{qty}× {desc}")
    return items if items else [text]

def co_first_name(name):
    parts = str(name or "").strip().split()
    return parts[0] if parts else "there"

def co_fill_template(tpl, order, settings):
    items = ", ".join(co_parse_items(order["note"]))
    code = (settings["coupon_code"] or "").strip() or "the code"
    pct = settings["coupon_percent"]
    pct_str = str(int(pct)) if float(pct).is_integer() else str(pct)
    return (
        str(tpl or "")
        .replace("{name}", order["customer"] or "there")
        .replace("{firstname}", co_first_name(order["customer"]))
        .replace("{employee}", order["assigned_to"] or "")
        .replace("{items}", items)
        .replace("{total}", f"${order['total_usd']:,.2f}")
        .replace("{code}", code)
        .replace("{percent}", pct_str)
    )

# api.whatsapp.com is WhatsApp's own click-to-chat endpoint. The wa.me
# shortlink 302-redirects through it and lands with a "#no_universal_links"
# fragment tacked on by WhatsApp's redirector — linking straight to
# api.whatsapp.com skips that hop entirely.
def co_wa_link(phone, text):
    digits = re.sub(r"\D", "", str(phone or ""))
    url = f"https://api.whatsapp.com/send?phone={digits}"
    if text:
        url += "&text=" + url_quote(text, safe="")
    return url

@st.cache_data(ttl=30)
@_resilient_google_call
def load_cancelled_orders():
    ws = get_co_ws()
    data = ws.get_all_values()
    orders = []
    for i, row in enumerate(data[1:], start=2):
        if len(row) < 1 or not row[0].strip():
            continue
        def _g(idx, default=""):
            return row[idx].strip() if len(row) > idx else default
        stage = _g(12, "wa1")
        orders.append({
            "row": i,
            "order_id": _g(0),
            "reference_id": _g(1),
            "customer": _g(2),
            "phone": _g(3),
            "address": _g(4),
            "note": _g(5),
            "usd_price": _parse_amount(_g(6)) or 0.0,
            "delivery_usd": _parse_amount(_g(7)) or 0.0,
            "total_usd": _parse_amount(_g(8)) or 0.0,
            "payment_status": _g(9),
            "creation_date": _g(10),
            "assigned_to": _g(11),
            "stage": stage if stage in CO_STAGE_ORDER else "wa1",
            "outcome": _g(13) or None,
            "notes": _g(14),
            "date_added": _g(15),
        })
    return orders

def co_status_is_cancelled(raw):
    return bool(re.search("cancel", str(raw or ""), re.IGNORECASE))

def co_payment_status(raw):
    return "Paid" if re.match(r"^paid", str(raw or "").strip(), re.IGNORECASE) else "Unpaid (COD)"

# Roadrunner's daily export — parses .xlsx (via pandas/openpyxl) or .csv, finds
# the header row (it isn't always row 0), maps columns by name, and returns
# only rows whose Status contains "cancel". Cancelled-but-already-tracked
# orders are filtered out by the caller against load_cancelled_orders().
def parse_roadrunner_export(file):
    if file.name.lower().endswith(".csv"):
        raw = pd.read_csv(file, header=None, dtype=str, keep_default_na=False)
    else:
        raw = pd.read_excel(file, header=None, dtype=str, keep_default_na=False)

    header_idx = None
    for i in range(min(10, len(raw))):
        row_vals = [str(v).strip().lower() for v in raw.iloc[i].tolist()]
        if "order id" in row_vals:
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('Couldn\'t find an "Order ID" column — is this the right export?')

    header = [str(v).strip() for v in raw.iloc[header_idx].tolist()]
    body = raw.iloc[header_idx + 1:].reset_index(drop=True)
    body.columns = header

    def col(name):
        return name if name in body.columns else None

    c_order = col("Order ID")
    if not c_order:
        raise ValueError('Couldn\'t find an "Order ID" column — is this the right export?')
    c_ref, c_usd, c_delivery, c_total, c_status, c_created, c_customer, c_phone, c_address, c_note = (
        col("Reference ID"), col("USD Price"), col("Delivery USD"), col("Total USD"),
        col("Status"), col("Creation Date"), col("Customer Name"), col("Phone"),
        col("Address"), col("Note"),
    )

    records, skipped = [], 0
    for _, r in body.iterrows():
        order_id = str(r.get(c_order, "")).strip()
        if not order_id:
            continue
        status_raw = r.get(c_status, "") if c_status else ""
        if not co_status_is_cancelled(status_raw):
            skipped += 1
            continue
        records.append({
            "order_id": order_id,
            "reference_id": str(r.get(c_ref, "")).strip() if c_ref else "",
            "customer": str(r.get(c_customer, "")).strip() if c_customer else "",
            "phone": str(r.get(c_phone, "")).strip() if c_phone else "",
            "address": str(r.get(c_address, "")).strip() if c_address else "",
            "note": str(r.get(c_note, "")).strip() if c_note else "",
            "usd_price": _parse_amount(r.get(c_usd, "")) or 0.0 if c_usd else 0.0,
            "delivery_usd": _parse_amount(r.get(c_delivery, "")) or 0.0 if c_delivery else 0.0,
            "total_usd": _parse_amount(r.get(c_total, "")) or 0.0 if c_total else 0.0,
            "payment_status": co_payment_status(status_raw),
            "creation_date": str(r.get(c_created, "")).strip() if c_created else "",
        })
    return records, len(body), skipped

def add_cancelled_orders_batch(records, settings):
    ws = get_co_ws()
    rows, per_emp = [], {}
    today = datetime.now().strftime("%Y-%m-%d")
    for rec in sorted(records, key=lambda r: r["creation_date"]):
        emp = co_next_assignee(settings)
        per_emp[emp] = per_emp.get(emp, 0) + 1
        rows.append([
            rec["order_id"], rec["reference_id"], rec["customer"], rec["phone"],
            rec["address"], rec["note"], rec["usd_price"], rec["delivery_usd"],
            rec["total_usd"], rec["payment_status"], rec["creation_date"],
            emp, "wa1", "", "", today,
        ])
    if rows:
        ws.append_rows(rows)
    save_co_settings(settings)
    load_cancelled_orders.clear()
    return per_emp

def update_co_order(row, **fields):
    ws = get_co_ws()
    col_map = {"assigned_to": 12, "stage": 13, "outcome": 14, "notes": 15}
    updates = []
    for key, value in fields.items():
        updates.append({"range": gspread.utils.rowcol_to_a1(row, col_map[key]), "values": [[value]]})
    if updates:
        ws.batch_update(updates)
    # Deliberately NOT clearing load_cancelled_orders' cache here — that forced
    # a full Sheet re-read (on top of the write above) on every single Yes/No/
    # notes click, which is what made those buttons feel laggy. The page
    # applies this exact change to its already-loaded list itself instead (see
    # co_local_patches in session_state); the cache still expires normally
    # (ttl=30) so other devices/sessions pick up the change shortly after, and
    # the "Refresh" button forces an immediate full resync for anyone who
    # wants one right away.
    patches = st.session_state.setdefault("co_local_patches", {})
    patches.setdefault(row, {}).update(fields)

# ─── Auth ─────────────────────────────────────────────────────────────────────

def _autofill_fix_script():
    """iOS Safari's saved-password autofill sets an input's value without firing the
    events React listens for, so Streamlit's widget state stays stale — the field
    looks filled but the app still sees the old (often empty) value, which reads as
    a wrong password. The animationstart/:-webkit-autofill trick alone isn't reliable
    enough (Keychain's QuickType-bar fill doesn't always trigger it the same way
    Safari's own form-fill styling does), so this resyncs on three independent
    triggers instead of just one: the autofill animation, losing focus, and — as a
    last-resort net — the moment ANY tap starts anywhere on the page, caught in the
    capture phase so it runs before Streamlit's own click handling does."""
    nonce = f"<!-- {datetime.now().isoformat()} -->"  # forces a fresh iframe reload every render
    script = nonce + """
<script>
try {
    const doc = window.parent.document;
    if (!doc.getElementById('__sm_autofill_fix_style__')) {
        const s = doc.createElement('style');
        s.id = '__sm_autofill_fix_style__';
        s.textContent = '@keyframes smAutoFill { from {} to {} } input:-webkit-autofill { animation-name: smAutoFill; }';
        doc.head.appendChild(s);
    }
    const setter = Object.getOwnPropertyDescriptor(window.parent.HTMLInputElement.prototype, 'value').set;
    function syncAll() {
        doc.querySelectorAll('[data-testid="stTextInputRootElement"] input').forEach(function (inp) {
            setter.call(inp, inp.value);
            inp.dispatchEvent(new Event('input', { bubbles: true }));
        });
    }
    doc.querySelectorAll('[data-testid="stTextInputRootElement"] input').forEach(function (inp) {
        if (inp.dataset.smAutofillHooked) return;
        inp.dataset.smAutofillHooked = '1';
        inp.addEventListener('animationstart', function (e) {
            if (e.animationName === 'smAutoFill') syncAll();
        });
        inp.addEventListener('blur', syncAll);
    });
    if (!doc.__smAutofillCaptureHooked) {
        doc.__smAutofillCaptureHooked = true;
        doc.addEventListener('pointerdown', syncAll, true);
        doc.addEventListener('touchstart', syncAll, true);
    }
} catch (e) {}
</script>
"""
    components.html(script, height=0)

def login_screen():
    if st.session_state.get("authenticated"):
        return True

    users = load_users()

    # Silent re-auth: session_state alone doesn't survive a real page reload (mobile
    # browsers in particular tear the whole session down on refresh/background-resume,
    # unlike a typical desktop soft-reload). A token in the URL does survive a reload,
    # so a matching one here logs the user back in without asking again.
    token = st.query_params.get("t")
    if token:
        token_hash = _hash_token(token)
        matched = next((u for u in users.values() if u["remember_hash"] and u["remember_hash"] == token_hash), None)
        if matched:
            st.session_state.authenticated = True
            st.session_state.username = matched["username"]
            st.session_state.role = matched["role"]
            st.session_state.pages = effective_pages(matched)
            return True
        del st.query_params["t"]  # stale/revoked token — drop it so it stops being tried

    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"] { background: #f4f5f7; }
    [data-testid="stForm"] { border: none !important; padding: 0 !important; background: transparent !important; }
    </style>
    """, unsafe_allow_html=True)

    col = st.columns([1, 1.2, 1])[1]
    with col:
        st.markdown("<br><br>", unsafe_allow_html=True)
        st.markdown(
            '<div style="width:56px;height:56px;border-radius:50%;background:#e6394f;'
            'display:flex;align-items:center;justify-content:center;margin:0 auto 0.9rem;'
            'font-family:\'Inter\',sans-serif;font-weight:700;font-size:1.3rem;color:#fff;">SM</div>',
            unsafe_allow_html=True,
        )
        st.markdown('<p style="font-family:\'Inter\',sans-serif;font-size:1.5rem;font-weight:700;color:#1f232c;text-align:center;letter-spacing:0">Sweet Mayhem</p>', unsafe_allow_html=True)
        st.markdown('<p style="font-size:0.7rem;color:#8b8f9b;text-align:center;letter-spacing:0.15em;text-transform:uppercase;margin-top:-1rem">Fulfillment Studio</p>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        if not users:
            st.info("No accounts exist yet. Create the first admin account to get started.")
            new_user = st.text_input("Admin username")
            new_pass = st.text_input("Admin password", type="password")
            confirm = st.text_input("Confirm password", type="password")
            if st.button("Create Admin Account", type="primary", use_container_width=True):
                if not new_user.strip() or not new_pass:
                    st.error("Username and password are required.")
                elif new_pass != confirm:
                    st.error("Passwords don't match.")
                else:
                    create_user(new_user.strip(), new_pass, "admin", ALL_PAGES)
                    st.success("Admin account created — please sign in.")
                    st.rerun()
        else:
            # Submits both fields as one atomic unit rather than two independent widgets
            # each syncing to the server on their own.
            with st.form("login_form", clear_on_submit=False):
                username = st.text_input("Username", autocomplete="username")
                pwd = st.text_input(
                    "Password", type="password", placeholder="Enter password",
                    autocomplete="current-password",
                )
                submitted = st.form_submit_button("Sign In", type="primary", use_container_width=True)

            if submitted:
                u = users.get(username.strip().lower())
                if u and verify_password(pwd, u["password_hash"]):
                    st.session_state.authenticated = True
                    st.session_state.username = u["username"]
                    st.session_state.role = u["role"]
                    st.session_state.pages = effective_pages(u)
                    st.query_params["t"] = issue_remember_token(u["row"])
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")

    _autofill_fix_script()
    return False

if not login_screen():
    st.stop()

@st.cache_data(ttl=60)
@_resilient_google_call
def load_inventory():
    ws = get_ws()
    data = ws.get_all_values()
    inv = {}
    for i, row in enumerate(data[1:], start=2):
        if len(row) < 4:
            continue
        p, c, s, q = row[0].strip(), row[1].strip(), row[2].strip(), row[3]
        if not p or not c or not s:
            continue
        try:
            qty = int(q)
        except (ValueError, TypeError):
            qty = 0
        key = (p.lower(), c.lower(), s.lower())
        if key not in inv:
            inv[key] = {"product": p, "color": c, "size": s, "qty": qty, "row": i}
    return inv

def batch_update_qty(ws, updates):
    ws.batch_update([
        {"range": f"D{row}", "values": [[qty]]}
        for row, qty in updates
    ])
    load_inventory.clear()

PRODUCT_PRICES_TAB = "Product Prices"

@_resilient_google_call
def get_product_prices_ws():
    sh = _spreadsheet()
    try:
        return sh.worksheet(PRODUCT_PRICES_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=PRODUCT_PRICES_TAB, rows=200, cols=2)
        ws.append_row(["Product", "Unit Price"])
        return ws

@st.cache_data(ttl=120)
@_resilient_google_call
def load_product_prices():
    """Product -> fixed unit price, used to pre-fill Shipment Details' Add New
    Shipment form so prices stay consistent across batches instead of being
    retyped (and potentially mistyped) every time."""
    ws = get_product_prices_ws()
    data = ws.get_all_values()
    prices = {}
    for row in data[1:]:
        if len(row) < 2 or not row[0].strip():
            continue
        try:
            prices[row[0].strip()] = float(row[1])
        except (ValueError, TypeError):
            continue
    return prices

def set_product_prices(updates):
    """updates: dict of product -> unit price. Updates existing rows in place,
    appends new ones for products not yet tracked."""
    ws = get_product_prices_ws()
    data = ws.get_all_values()
    remaining = dict(updates)
    for i, row in enumerate(data[1:], start=2):
        if row and row[0].strip() in remaining:
            ws.update_cell(i, 2, remaining.pop(row[0].strip()))
    if remaining:
        ws.append_rows([[product, price] for product, price in remaining.items()])
    load_product_prices.clear()

# ─── Shopify ──────────────────────────────────────────────────────────────────

_SHOPIFY_STORE   = st.secrets["shopify"]["store"]
_SHOPIFY_TOKEN   = st.secrets["shopify"]["access_token"]
_SHOPIFY_VERSION = st.secrets["shopify"]["api_version"]
_SHOPIFY_HEADERS = {"X-Shopify-Access-Token": _SHOPIFY_TOKEN, "Content-Type": "application/json"}
_SHOPIFY_BASE    = f"https://{_SHOPIFY_STORE}/admin/api/{_SHOPIFY_VERSION}"

def _order_phone(o):
    shipping = o.get("shipping_address") or {}
    customer = o.get("customer") or {}
    return (shipping.get("phone") or o.get("phone") or customer.get("phone") or "").strip()

def _parse_order(o):
    items = [
        {"name": li["name"], "quantity": li.get("fulfillable_quantity", li["quantity"])}
        for li in o["line_items"]
        if li.get("fulfillable_quantity", li["quantity"]) > 0
    ]
    if not items:
        return None
    return {
        "name": o["name"],
        "id": o["id"],
        "email": o.get("email", ""),
        "phone": _order_phone(o),
        "created_at": o["created_at"],
        "line_items": items,
    }

def fetch_shopify_orders():
    resp = requests.get(
        f"{_SHOPIFY_BASE}/orders.json",
        headers=_SHOPIFY_HEADERS,
        params={"fulfillment_status": "unfulfilled", "status": "open", "limit": 250},
    )
    resp.raise_for_status()
    orders = [_parse_order(o) for o in resp.json()["orders"]]
    return sorted([o for o in orders if o], key=lambda x: x["created_at"])

def fetch_shopify_order_by_name(order_name):
    name = order_name.strip().lstrip("#")
    resp = requests.get(
        f"{_SHOPIFY_BASE}/orders.json",
        headers=_SHOPIFY_HEADERS,
        params={"name": f"#{name}", "status": "any", "limit": 5},
    )
    resp.raise_for_status()
    results = resp.json().get("orders", [])
    if not results:
        raise Exception(f"Order #{name} not found in Shopify.")
    o = results[0]
    parsed = _parse_order(o)
    if not parsed:
        raise Exception(f"Order #{name} has no unfulfilled items.")
    return [parsed]

def shopify_fulfill_order(order_id):
    fo_resp = requests.get(f"{_SHOPIFY_BASE}/orders/{order_id}/fulfillment_orders.json", headers=_SHOPIFY_HEADERS)
    if not fo_resp.ok:
        raise Exception(f"Failed to get fulfillment orders: {fo_resp.status_code} {fo_resp.text}")
    fulfillment_orders = fo_resp.json().get("fulfillment_orders", [])
    SKIP = {"closed", "cancelled", "fulfilled", "incomplete"}
    fo_ids = [
        {"fulfillment_order_id": fo["id"]}
        for fo in fulfillment_orders
        if fo["status"] not in SKIP
    ]
    if not fo_ids:
        statuses = [fo["status"] for fo in fulfillment_orders]
        raise Exception(f"No fulfillable fulfillment orders (statuses: {statuses})")
    payload = {
        "fulfillment": {
            "line_items_by_fulfillment_order": fo_ids,
            "notify_customer": False,
        }
    }
    f_resp = requests.post(f"{_SHOPIFY_BASE}/fulfillments.json", headers=_SHOPIFY_HEADERS, json=payload)
    if not f_resp.ok:
        raise Exception(f"Fulfillment failed: {f_resp.status_code} {f_resp.text}")

@st.cache_resource
def get_primary_location_id():
    resp = requests.get(f"{_SHOPIFY_BASE}/locations.json", headers=_SHOPIFY_HEADERS)
    resp.raise_for_status()
    locations = resp.json().get("locations", [])
    if not locations:
        raise Exception("No Shopify locations found.")
    return locations[0]["id"]

@st.cache_data(ttl=600)
def fetch_shopify_variant_map():
    """(product, color, size) [lowercased] -> inventory_item_id, for restock syncing."""
    result = {}
    url = f"{_SHOPIFY_BASE}/products.json"
    params = {"limit": 250, "fields": "id,title,variants"}
    while url:
        resp = requests.get(url, headers=_SHOPIFY_HEADERS, params=params)
        resp.raise_for_status()
        for p in resp.json().get("products", []):
            for v in p.get("variants", []):
                parts = (v.get("title") or "").split(" / ", 1)
                if len(parts) != 2:
                    continue
                color, size = parts[0].strip(), parts[1].strip()
                key = (p["title"].lower(), color.lower(), size.lower())
                if key not in result:
                    result[key] = v["inventory_item_id"]
        next_url = None
        for part in resp.headers.get("Link", "").split(","):
            if 'rel="next"' in part and "<" in part:
                next_url = part[part.find("<") + 1:part.find(">")]
        url, params = next_url, None
    return result

@st.cache_data(ttl=600)
def fetch_shopify_available_by_key():
    """(product, color, size) [lowercased] -> {"available": qty, "continue_oos": bool}.
    Available (not on_hand) is what actually blocks a sale, so it's what
    determines whether a variant counts as out-of-stock for a given day's
    snapshot. continue_oos reflects Shopify's "Continue selling when out of
    stock" setting (inventoryPolicy == CONTINUE) — when on, the variant can
    always be sold regardless of Available, so no day should ever count as
    out-of-stock for it. Matches the same (product, color, size) key format
    load_inventory() uses."""
    location_id = get_primary_location_id()
    result = {}
    cursor = None
    query = """
    query($cursor: String, $locationId: ID!) {
      products(first: 100, after: $cursor) {
        edges {
          node {
            title
            variants(first: 100) {
              edges {
                node {
                  title
                  inventoryPolicy
                  inventoryItem {
                    inventoryLevel(locationId: $locationId) {
                      quantities(names: ["available"]) { quantity }
                    }
                  }
                }
              }
            }
          }
        }
        pageInfo { hasNextPage endCursor }
      }
    }
    """
    while True:
        resp = requests.post(
            f"{_SHOPIFY_BASE}/graphql.json", headers=_SHOPIFY_HEADERS,
            json={"query": query, "variables": {"cursor": cursor, "locationId": _gid("Location", location_id)}},
        )
        resp.raise_for_status()
        data = resp.json()
        if data.get("errors"):
            raise Exception("; ".join(e["message"] for e in data["errors"]))
        products_page = data["data"]["products"]
        for edge in products_page["edges"]:
            p = edge["node"]
            for ve in p["variants"]["edges"]:
                v = ve["node"]
                parts = (v.get("title") or "").split(" / ", 1)
                if len(parts) != 2:
                    continue
                color, size = parts[0].strip(), parts[1].strip()
                level = v["inventoryItem"].get("inventoryLevel")
                qty = level["quantities"][0]["quantity"] if level else 0
                key = (p["title"].lower(), color.lower(), size.lower())
                result[key] = {"available": qty, "continue_oos": v.get("inventoryPolicy") == "CONTINUE"}
        if products_page["pageInfo"]["hasNextPage"]:
            cursor = products_page["pageInfo"]["endCursor"]
        else:
            break
    return result

def find_shopify_inventory_item(variant_map, product, color, size):
    """Same exact-color/size + fuzzy-product-name matching as find_key(), applied to
    Shopify's variant list instead of the Google Sheet inventory."""
    exact = (product.lower(), color.lower(), size.lower())
    if exact in variant_map:
        return variant_map[exact]
    best, ratio = None, 0.75
    for key, item_id in variant_map.items():
        p, c, s = key
        if c != color.lower() or s != size.lower():
            continue
        r = SequenceMatcher(None, product.lower(), p).ratio()
        if r > ratio:
            ratio, best = r, item_id
    return best

def _gid(resource, numeric_id):
    return f"gid://shopify/{resource}/{numeric_id}"

def _current_shopify_onhand(inventory_item_id, location_id):
    query = """
    query currentOnHand($itemId: ID!, $locationId: ID!) {
      inventoryItem(id: $itemId) {
        inventoryLevel(locationId: $locationId) {
          quantities(names: ["on_hand"]) { quantity }
        }
      }
    }
    """
    variables = {
        "itemId": _gid("InventoryItem", inventory_item_id),
        "locationId": _gid("Location", location_id),
    }
    resp = requests.post(
        f"{_SHOPIFY_BASE}/graphql.json",
        headers=_SHOPIFY_HEADERS,
        json={"query": query, "variables": variables},
    )
    if not resp.ok:
        raise Exception(f"{resp.status_code} {resp.text}")
    data = resp.json()
    if data.get("errors"):
        raise Exception("; ".join(e["message"] for e in data["errors"]))
    level = (data.get("data") or {}).get("inventoryItem", {}).get("inventoryLevel")
    if not level:
        raise Exception("No inventory level found for this item at this location.")
    return level["quantities"][0]["quantity"]

def add_shopify_onhand_quantity(inventory_item_id, location_id, delta):
    """Adds delta to whatever Shopify's 'On hand' quantity (not 'Available') currently
    is — does not overwrite it to match the sheet, so manual additions made directly
    in Shopify (e.g. for incoming batches) aren't clobbered."""
    current = _current_shopify_onhand(inventory_item_id, location_id)
    quantity = current + delta
    query = """
    mutation setOnHand($input: InventorySetQuantitiesInput!, $key: String!) {
      inventorySetQuantities(input: $input) @idempotent(key: $key) {
        userErrors { field message }
      }
    }
    """
    variables = {
        "input": {
            "name": "on_hand",
            "reason": "correction",
            "quantities": [{
                "inventoryItemId": _gid("InventoryItem", inventory_item_id),
                "locationId": _gid("Location", location_id),
                "quantity": quantity,
                "changeFromQuantity": current,
            }],
        },
        "key": str(uuid.uuid4()),
    }
    resp = requests.post(
        f"{_SHOPIFY_BASE}/graphql.json",
        headers=_SHOPIFY_HEADERS,
        json={"query": query, "variables": variables},
    )
    if not resp.ok:
        raise Exception(f"{resp.status_code} {resp.text}")
    data = resp.json()
    if data.get("errors"):
        raise Exception("; ".join(e["message"] for e in data["errors"]))
    errors = data.get("data", {}).get("inventorySetQuantities", {}).get("userErrors", [])
    if errors:
        raise Exception("; ".join(e["message"] for e in errors))

# ─── Logic ────────────────────────────────────────────────────────────────────

def parse_lineitem_name(name):
    if not name or str(name).strip().lower() in ("nan", ""):
        return None, None, None
    name = str(name).strip()
    di = name.rfind(" - ")
    if di == -1:
        return None, None, None
    product = name[:di].strip()
    cs = name[di + 3:].strip()
    si = cs.rfind(" / ")
    if si == -1:
        return product, cs, None
    return product, cs[:si].strip(), cs[si + 3:].strip()

def find_key(inv, product, color, size):
    exact = (product.lower(), color.lower(), size.lower())
    if exact in inv:
        return exact
    best, ratio = None, 0.75
    for key in inv:
        ip, ic, is_ = key
        if ic != color.lower() or is_ != size.lower():
            continue
        r = SequenceMatcher(None, product.lower(), ip).ratio()
        if r > ratio:
            ratio, best = r, key
    return best

def load_orders(file):
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file, dtype=str, low_memory=False)
    else:
        df = pd.read_excel(file, dtype=str)
    df.columns = [c.strip() for c in df.columns]
    required = ["Name", "Fulfillment Status", "Created at", "Lineitem quantity", "Lineitem name"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing columns: {missing}")
    unf = set(df.loc[df["Fulfillment Status"].str.strip().str.lower() == "unfulfilled", "Name"])
    if not unf:
        return []
    df_u = df[df["Name"].isin(unf)].copy()
    orders = {}
    for _, row in df_u.iterrows():
        name = str(row["Name"]).strip()
        if name not in orders:
            phone = str(row.get("Shipping Phone") or row.get("Phone") or "").strip()
            orders[name] = {
                "name": name,
                "email": str(row.get("Email", "")).strip(),
                "phone": "" if phone.lower() == "nan" else phone,
                "created_at": str(row.get("Created at", "")).strip(),
                "line_items": [],
            }
        li = str(row.get("Lineitem name", "")).strip()
        try:
            qty = int(float(str(row.get("Lineitem quantity", 1))))
        except Exception:
            qty = 1
        if li and li.lower() != "nan":
            orders[name]["line_items"].append({"name": li, "quantity": qty})
    return sorted(orders.values(), key=lambda x: x["created_at"])

def determine_fulfillable(orders, inv):
    working = deepcopy(inv)
    fulfillable, skipped = [], []
    for order in orders:
        ok, reason, reqs = True, None, {}
        for item in order["line_items"]:
            p, c, s = parse_lineitem_name(item["name"])
            if not p or not s:
                ok, reason = False, f"Can't parse: '{item['name']}'"
                break
            if not c:
                ok, reason = False, f"No color in: '{item['name']}'"
                break
            k = find_key(working, p, c, s)
            if k is None:
                ok, reason = False, f"Not in inventory: {p} — {c} / {s}"
                break
            reqs[k] = reqs.get(k, 0) + item["quantity"]
        if ok:
            for k, need in reqs.items():
                have = working[k]["qty"]
                if have < need:
                    v = working[k]
                    ok, reason = False, f"Low stock: {v['product']} — {v['color']} / {v['size']} (need {need}, have {have})"
                    break
        if ok:
            for k, q in reqs.items():
                working[k]["qty"] -= q
            fulfillable.append(order)
        else:
            order["skip_reason"] = reason
            skipped.append(order)
    return fulfillable, skipped, working

def recalc_inv(orig, fulfillable):
    new = deepcopy(orig)
    for order in fulfillable:
        for item in order["line_items"]:
            p, c, s = parse_lineitem_name(item["name"])
            if p and c and s:
                k = find_key(new, p, c, s)
                if k:
                    new[k]["qty"] -= item["quantity"]
    return new

def make_report(fulfillable):
    rows = []
    for o in fulfillable:
        for item in o["line_items"]:
            rows.append({
                "Order #": o["name"],
                "Created At": o["created_at"],
                "Email": o["email"],
                "Phone": o.get("phone") or "MISSING",
                "Item": item["name"],
                "Qty": item["quantity"],
                "Action": "Fulfill in Shopify",
            })
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ─── Shared status-color styling ───────────────────────────────────────────────
# st.dataframe's grid can't render icon fonts (or even reliably show emoji), so
# status values are plain text and colored via a pandas Styler instead of emoji.

STATUS_COLORS = {
    "Out of Stock": "#fdecea", "Cancelled": "#fdecea", "Rejected": "#fdecea",
    "Reorder Now": "#fdecea",
    "Reorder Soon": "#fff6e0", "Pending": "#fff6e0",
    "In Transit": "#eef2f7",
    "OK": "#eaf7ed", "Received": "#eaf7ed", "Refunded": "#eaf7ed",
}

def style_status(df, column="Status"):
    def _color(val):
        base = str(val).split(" (")[0].strip()
        bg = STATUS_COLORS.get(base, "")
        return f"background-color: {bg}" if bg else ""
    return df.style.map(_color, subset=[column])

REORDER_STATUS_PILL = {
    "Reorder Now": "rr-pill-red", "Out of Stock": "rr-pill-red",
    "Reorder Soon": "rr-pill-amber", "OK": "rr-pill-green",
}

def render_reorder_table(df):
    """Dashboard-style table for Demand & Reorder — same rationale as
    render_shipment_table: pill badges and right-aligned tabular numbers instead
    of a plain st.dataframe grid."""
    def esc(v):
        return html_lib.escape(str(v)) if v not in (None, "") else "—"

    def fmt_int(v):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return "—"
        return f"{int(v):,}"

    def fmt_days(v):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return "—"
        if isinstance(v, float) and math.isinf(v):
            return "∞"
        return f"{v:,.1f}"

    rows_html = []
    for _, r in df.iterrows():
        pill_cls = REORDER_STATUS_PILL.get(r["Status"], "rr-pill-blue")
        rows_html.append(f"""
        <tr>
          <td class="rr-t-strong">{esc(r['Product'])}</td>
          <td>{esc(r['Color'])}</td>
          <td>{esc(r['Size'])}</td>
          <td class="rr-t-num">{fmt_int(r['Current Qty'])}</td>
          <td class="rr-t-num">{fmt_int(r['Units Sold'])}</td>
          <td class="rr-t-num">{fmt_int(r['Days OOS (window)'])}</td>
          <td class="rr-t-num">{r['Daily Demand']:.2f}</td>
          <td class="rr-t-num">{fmt_days(r['Days Left'])}</td>
          <td class="rr-t-num">{fmt_int(r['Incoming Qty'])}</td>
          <td class="rr-t-num rr-t-strong">{fmt_int(r['Reorder Qty'])}</td>
          <td><span class="rr-pill {pill_cls}">{esc(r['Status'])}</span></td>
          <td class="rr-t-trunc">{esc(r['Confidence'])}</td>
        </tr>""")

    headers = [
        "Product", "Color", "Size", "Current Qty", "Units Sold", "Days OOS",
        "Daily Demand", "Days Left", "Incoming Qty", "Reorder Qty", "Status", "Confidence",
    ]
    num_cols = {"Current Qty", "Units Sold", "Days OOS", "Daily Demand", "Days Left", "Incoming Qty", "Reorder Qty"}
    header_html = "".join(
        f'<th class="{"rr-t-num" if h in num_cols else ""}">{h}</th>' for h in headers
    )

    st.markdown(f"""
    <div class="rr-table-wrap">
      <table class="rr-table">
        <thead><tr>{header_html}</tr></thead>
        <tbody>{"".join(rows_html)}</tbody>
      </table>
    </div>
    """, unsafe_allow_html=True)

def render_incoming_table(df):
    """Dashboard-style table for Purchase Orders' variant-level breakdown."""
    def esc(v):
        return html_lib.escape(str(v)) if v not in (None, "") else "—"

    rows_html = []
    for _, r in df.iterrows():
        rows_html.append(f"""
        <tr>
          <td class="rr-t-strong">{esc(r['Product'])}</td>
          <td>{esc(r['Color'])}</td>
          <td>{esc(r['Size'])}</td>
          <td class="rr-t-num rr-t-strong">{r['Incoming Qty']:,}</td>
          <td class="rr-t-trunc">{esc(r['Batches'])}</td>
        </tr>""")

    inc_headers = ["Product", "Color", "Size", "Incoming Qty", "Batches"]
    inc_header_html = "".join(
        f'<th class="{"rr-t-num" if h == "Incoming Qty" else ""}">{h}</th>' for h in inc_headers
    )

    st.markdown(f"""
    <div class="rr-table-wrap">
      <table class="rr-table">
        <thead><tr>{inc_header_html}</tr></thead>
        <tbody>{"".join(rows_html)}</tbody>
      </table>
    </div>
    """, unsafe_allow_html=True)

SHIPMENT_STATUS_PILL = {"Received": "rr-pill-green", "Cancelled": "rr-pill-red", "In Transit": "rr-pill-blue"}

def render_shipment_table(df):
    """Custom dashboard-style table for Shipment Tracker — st.dataframe can't do
    pill badges, truncation-with-tooltip, or per-column typography, so this
    renders the rows as HTML instead, reusing the app's existing design tokens."""
    def esc(v):
        return html_lib.escape(str(v)) if v not in (None, "") else "—"

    def fmt_date(v):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return "—"
        try:
            return pd.Timestamp(v).strftime("%b %d, %Y")
        except Exception:
            return esc(v)

    def fmt_num(v, prefix=""):
        if v is None or (isinstance(v, float) and math.isnan(v)):
            return "—"
        return f"{prefix}{v:,.0f}"

    def truncated(v, maxlen=42):
        text = str(v) if v not in (None, "") else ""
        if not text:
            return "—"
        short = text if len(text) <= maxlen else text[:maxlen - 1] + "…"
        return f'<span title="{html_lib.escape(text)}">{html_lib.escape(short)}</span>'

    rows_html = []
    for _, r in df.iterrows():
        pill_cls = SHIPMENT_STATUS_PILL.get(r["Status"], "rr-pill-amber")
        rows_html.append(f"""
        <tr>
          <td class="rr-t-strong">{esc(r['Batch #'])}</td>
          <td>{esc(r['Brand'])}</td>
          <td>{fmt_date(r['Date Paid'])}</td>
          <td>{fmt_date(r['Date Shipped'])}</td>
          <td>{esc(r['Shipment Type'])}</td>
          <td>{esc(r['Shipping Company'])}</td>
          <td class="rr-t-mono">{esc(r['Tracking #'])}</td>
          <td>{fmt_date(r['Date Received'])}</td>
          <td><span class="rr-pill {pill_cls}">{esc(r['Status'])}</span></td>
          <td class="rr-t-num">{fmt_num(r['Price'], '$')}</td>
          <td class="rr-t-num">{fmt_num(r['# of Cartons'])}</td>
          <td class="rr-t-trunc">{truncated(r['Items Ordered'])}</td>
        </tr>""")

    headers = [
        "Batch #", "Brand", "Date Paid", "Date Shipped", "Shipment Type",
        "Shipping Company", "Tracking #", "Date Received", "Status",
        "Price", "# of Cartons", "Items Ordered",
    ]
    num_cols = {"Price", "# of Cartons"}
    header_html = "".join(
        f'<th class="{"rr-t-num" if h in num_cols else ""}">{h}</th>' for h in headers
    )

    st.markdown(f"""
    <div class="rr-table-wrap">
      <table class="rr-table">
        <thead><tr>{header_html}</tr></thead>
        <tbody>{"".join(rows_html)}</tbody>
      </table>
    </div>
    """, unsafe_allow_html=True)

# ─── Demand & Reorder ─────────────────────────────────────────────────────────

SNAPSHOT_SHEET_NAME = "InventorySnapshots"
MIN_TRACKED_DAYS = 5  # minimum days of stock-history before trusting the adjusted rate

@_resilient_google_call
def get_snapshot_ws():
    sh = _spreadsheet()
    try:
        return sh.worksheet(SNAPSHOT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SNAPSHOT_SHEET_NAME, rows=1000, cols=5)
        ws.append_row(["Date", "Product", "Color", "Size", "Qty"])
        return ws

def _match_shopify_available(available_by_key, product, color, size):
    """Same exact + fuzzy matching as find_shopify_inventory_item(), against
    fetch_shopify_available_by_key()'s (product, color, size) -> {"available",
    "continue_oos"} dict instead of -> item id. Returns that dict, or None."""
    exact = (product.lower(), color.lower(), size.lower())
    if exact in available_by_key:
        return available_by_key[exact]
    best, ratio, best_val = None, 0.75, None
    for key, info in available_by_key.items():
        p, c, s = key
        if c != color.lower() or s != size.lower():
            continue
        r = SequenceMatcher(None, product.lower(), p).ratio()
        if r > ratio:
            ratio, best_val = r, info
    return best_val

def record_snapshot_if_needed(inv):
    """Log today's Shopify Available quantity per variant, once per day —
    Available (not the Studio Inventory count) is what actually blocks a
    sale, so it's the correct signal for whether a variant was genuinely
    out of stock that day, feeding Demand & Reorder's OOS-adjusted demand."""
    today = datetime.now().strftime("%Y-%m-%d")
    ws = get_snapshot_ws()
    existing_dates = set(ws.col_values(1))
    if today in existing_dates:
        return False
    available_by_key = fetch_shopify_available_by_key()
    rows = []
    for v in inv.values():
        info = _match_shopify_available(available_by_key, v["product"], v["color"], v["size"])
        qty = info["available"] if info else v["qty"]  # not found in Shopify — fall back to the Studio count
        rows.append([today, v["product"], v["color"], v["size"], qty])
    if rows:
        ws.append_rows(rows)
    return True

@st.cache_data(ttl=600)
@_resilient_google_call
def load_snapshots():
    ws = get_snapshot_ws()
    data = ws.get_all_values()
    records = []
    for row in data[1:]:
        if len(row) < 5:
            continue
        date_s, p, c, s, q = row[0].strip(), row[1].strip(), row[2].strip(), row[3].strip(), row[4]
        try:
            qty = int(q)
        except (ValueError, TypeError):
            continue
        if not date_s or not p:
            continue
        records.append((date_s, p.lower(), c.lower(), s.lower(), qty))
    return records

def compute_stock_days(records, start_date, end_date):
    """Per variant key: which dates (within [start_date, end_date]) were tracked, and which had stock > 0."""
    start_s, end_s = start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")
    per_key = {}
    for date_s, p, c, s, qty in records:
        if date_s < start_s or date_s > end_s:
            continue
        key = (p, c, s)
        d = per_key.setdefault(key, {"tracked": set(), "in_stock": set()})
        d["tracked"].add(date_s)
        if qty > 0:
            d["in_stock"].add(date_s)
    return per_key

@st.cache_data(ttl=1800)
def fetch_shopify_sales(start_date, end_date):
    """All non-cancelled orders created between start_date and end_date (inclusive)."""
    since = start_date.strftime("%Y-%m-%dT00:00:00Z")
    until = end_date.strftime("%Y-%m-%dT23:59:59Z")
    url = f"{_SHOPIFY_BASE}/orders.json"
    params = {
        "status": "any",
        "created_at_min": since,
        "created_at_max": until,
        "limit": 250,
        "fields": "id,created_at,cancelled_at,line_items",
    }
    orders = []
    while url:
        resp = requests.get(url, headers=_SHOPIFY_HEADERS, params=params)
        resp.raise_for_status()
        orders.extend(resp.json().get("orders", []))
        next_url = None
        for part in resp.headers.get("Link", "").split(","):
            if 'rel="next"' in part and "<" in part:
                next_url = part[part.find("<") + 1:part.find(">")]
        url, params = next_url, None
    return [o for o in orders if not o.get("cancelled_at")]

def aggregate_sales(orders, inv):
    """Units sold per inventory key, matched the same way fulfillment matches line items."""
    sold, unmatched = {}, 0
    for o in orders:
        for li in o.get("line_items", []):
            qty = li.get("quantity", 0)
            p, c, s = parse_lineitem_name(li.get("name", ""))
            k = find_key(inv, p, c, s) if p and c and s else None
            if k is None:
                unmatched += qty
                continue
            sold[k] = sold.get(k, 0) + qty
    return sold, unmatched

def build_reorder_table(inv, sold, stock_days, start_date, end_date, lead_time, coverage_days, incoming=None, continue_oos=None):
    incoming = incoming or {}
    continue_oos = continue_oos or {}
    window_days = max(1, (end_date - start_date).days + 1)
    rows = []
    for key, v in inv.items():
        total_sold = sold.get(key, 0)
        sd = stock_days.get(key, {"tracked": set(), "in_stock": set()})
        days_tracked, days_in_stock = len(sd["tracked"]), len(sd["in_stock"])
        sells_past_zero = bool(continue_oos.get(key))
        if sells_past_zero:
            # "Continue selling when out of stock" is on for this variant, so
            # it's never actually unsellable — no day should be excluded from
            # its demand calculation just because Available hit zero.
            days_in_stock = days_tracked
        days_oos = days_tracked - days_in_stock
        cur_qty = max(0, v["qty"])
        incoming_qty = incoming.get(key, 0)

        adjusted = days_tracked >= MIN_TRACKED_DAYS and days_in_stock > 0
        if adjusted:
            daily_demand = total_sold / days_in_stock
        else:
            daily_demand = total_sold / window_days

        days_left = (cur_qty / daily_demand) if daily_demand > 0 else float("inf")
        reorder_qty = max(0, math.ceil(daily_demand * coverage_days) - cur_qty - incoming_qty)

        if v["qty"] <= 0:
            status = "Out of Stock"
        elif days_left <= lead_time:
            status = "Reorder Now"
        elif days_left <= lead_time + 7:
            status = "Reorder Soon"
        else:
            status = "OK"

        rows.append({
            "Product": v["product"], "Color": v["color"], "Size": v["size"],
            "Current Qty": v["qty"],
            "Units Sold": total_sold,
            "Days OOS (window)": days_oos,
            "Daily Demand": round(daily_demand, 2),
            "Days Left": float("inf") if daily_demand == 0 else round(days_left, 1),
            "Incoming Qty": incoming_qty,
            "Reorder Qty": reorder_qty,
            "Status": status,
            "Confidence": ("Adjusted" if adjusted else "Raw (building history)") + (" · sells past zero" if sells_past_zero else ""),
        })
    return pd.DataFrame(rows)

# ─── Shipment Tracker ───────────────────────────────────────────────────────────

SHIPMENT_TRACKER_SHEET_ID = "1xwLzbuUU_xbE7aetCI5CxSpkpEdRN2AwzsuSkpqeS7g"
SHIPMENT_TRACKER_TAB = "Shipments Tracker"

@st.cache_resource
def _shipment_tracker_spreadsheet():
    return _gc().open_by_key(SHIPMENT_TRACKER_SHEET_ID)

@_resilient_google_call
def get_shipment_tracker_ws():
    return _shipment_tracker_spreadsheet().worksheet(SHIPMENT_TRACKER_TAB)

def _parse_sheet_date(v):
    """Handles the sheet's mixed date storage: real date cells (serial numbers) and
    manually-typed text dates like '27/3/2026'. Returns None if it genuinely can't
    make sense of the value — the caller is responsible for surfacing that, rather
    than letting it silently show up as a blank date."""
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        try:
            return date(1899, 12, 30) + timedelta(days=int(v))
        except (ValueError, OverflowError):
            return None
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    # Tolerates one specific, seen-in-the-wild typo: a missing slash right before a
    # 4-digit year, e.g. "7/72026" meant as "7/7/2026".
    m = re.match(r"^(\d{1,2})/(\d{1,2})(\d{4})$", s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
        except ValueError:
            return None
    return None

def _as_number(v):
    return v if isinstance(v, (int, float)) else None

def _status_display(brand, status):
    if str(brand).strip().upper() == "CANCELLED":
        return "Cancelled"
    s = (status or "").strip()
    low = s.lower()
    if not s:
        return "In Transit"
    if "cancel" in low:
        return "Cancelled"
    if low in ("recieved", "received"):
        return "Received"
    # Anything else that merely mentions "received" (e.g. "Received on stockie",
    # an intermediate holding point) hasn't actually arrived yet — keep it in transit.
    return f"In Transit ({s})"

SHIPMENT_STATUS_OPTIONS = ["In Transit", "Received", "Cancelled"]

def _status_option_default(display_status):
    return display_status if display_status in SHIPMENT_STATUS_OPTIONS else "In Transit"

def _status_option_to_raw(option):
    return {"In Transit": "", "Received": "Received", "Cancelled": "Cancelled"}[option]

SHIPMENT_TYPE_OPTIONS = ["Air", "Sea"]
SHIPPING_COMPANY_OPTIONS = ["WTB"]

# The supplier hands each shipping method a fixed drop-off address — once you
# pick the type + company there's nothing left to type, so it's looked up
# here instead of being a free-text field.
WAREHOUSE_ADDRESS_BY_COMBO = {
    ("Air", "WTB"): (
        "AIR CARGO WAREHOUSE\n\n"
        "地址：广州市天河区沐陂西路八号大院D2-1-2空运仓库\n"
        "入仓号A 1604           （入仓号不需写在外箱，写在运单或者装箱单附带即可）\n"
        "如果货车送货 可以搜索高德地图《中航沈飞地板》\n"
        "周先生  18922246469\n"
        "陈先生  18928745807                                \n"
        "请注意  \n"
        "1  货物外包装四面写唛头   最好纸箱加编织袋包装\n"
        "2  送货请提供客户全名或者ID\n"
        "3  请提供详细装箱单(我司格式) 箱单随货一齐发！！！\n"
        "4 仓库正常收货上班时间周一至周六上午10点到下午18：00，超过正常收货时间将拒收或有偿加班收货，由送货商自行支付\n"
        "5 所有送货卸货自己安排自卸，如需要仓库安排卸货须有偿卸货\n"
        "6 所有送货需要按照要求放置到仓库指定位置,并取得仓库收货回单，方算入仓完成\n"
        "谢谢！"
    ),
}

def _options_with_current(options, current):
    """Keeps a legacy value (e.g. an old 'Rayan' shipment) selectable instead
    of silently swapping it for the first option when the field is opened."""
    current = (current or "").strip()
    return options if (not current or current in options) else options + [current]

@st.cache_data(ttl=120)
@_resilient_google_call
def load_shipments():
    ws = get_shipment_tracker_ws()
    values = ws.get("A2:S500", value_render_option="UNFORMATTED_VALUE")
    if not values:
        return pd.DataFrame(), []
    rows, date_issues = [], []
    date_cols = [("Date Paid", 2), ("Date Shipped", 3), ("Date Received", 9)]
    for sheet_row, row in enumerate(values[1:], start=3):
        row = row + [None] * (19 - len(row))
        batch = row[0]
        if not batch:
            continue
        brand, status = row[1] or "", row[10] or ""

        parsed = {}
        for label, idx in date_cols:
            raw = row[idx]
            parsed[label] = _parse_sheet_date(raw)
            if raw not in (None, "") and parsed[label] is None:
                date_issues.append(f"{batch} — {label}: {raw!r} isn't a recognizable date")

        rows.append({
            "Batch #": batch,
            "Brand": brand,
            "Date Paid": parsed["Date Paid"],
            "Date Shipped": parsed["Date Shipped"],
            "Shipment Type": row[4] or "",
            "Shipping Company": row[5] or "",
            "Warehouse Address": row[6] or "",
            "Shipping Mark": row[7] or "",
            "Tracking #": str(row[8]) if row[8] not in (None, "") else "",
            "Date Received": parsed["Date Received"],
            "Status": _status_display(brand, status),
            "Total Items": _as_number(row[11]),
            "Price": _as_number(row[12]),
            "# of Cartons": _as_number(row[13]),
            "Notes": row[14] or "",
            "Items Ordered": row[15] or "",
            "Img ref.": row[16] or "",
            "Shopify Inventory Status": row[17] or "",
            "Marked as Ordered": row[18] or "",
            "row": sheet_row,
            "raw_status": status,
        })
    return pd.DataFrame(rows), date_issues

def _shipment_row_values(fields):
    """fields: dict of the editable columns (Batch # through Marked as
    Ordered, matching load_shipments' keys) -> value, in sheet column order."""
    def d(v):
        return v.strftime("%Y-%m-%d") if v else ""
    return [
        fields["Batch #"], fields["Brand"], d(fields["Date Paid"]), d(fields["Date Shipped"]),
        fields["Shipment Type"], fields["Shipping Company"], fields["Warehouse Address"],
        fields["Shipping Mark"], fields["Tracking #"], d(fields["Date Received"]),
        fields["raw_status"], fields["Total Items"], fields["Price"], fields["# of Cartons"],
        fields["Notes"], fields["Items Ordered"], fields["Img ref."], fields["Shopify Inventory Status"],
        fields["Marked as Ordered"],
    ]

def update_shipment(row_num, fields):
    """Writes the whole row back in one call so a stale read never overwrites
    an unrelated cell."""
    ws = get_shipment_tracker_ws()
    ws.update(f"A{row_num}:S{row_num}", [_shipment_row_values(fields)])
    load_shipments.clear()

def add_shipment(fields):
    ws = get_shipment_tracker_ws()
    ws.append_row(_shipment_row_values(fields))
    load_shipments.clear()

def _row_to_fields(row):
    """A load_shipments() row -> the fields dict update_shipment() expects,
    for callers that only want to change one or two fields and leave the rest
    of the row exactly as it was."""
    return {
        "Batch #": row["Batch #"], "Brand": row["Brand"],
        "Date Paid": row["Date Paid"], "Date Shipped": row["Date Shipped"],
        "Shipment Type": row["Shipment Type"], "Shipping Company": row["Shipping Company"],
        "Warehouse Address": row["Warehouse Address"], "Shipping Mark": row["Shipping Mark"],
        "Tracking #": row["Tracking #"], "Date Received": row["Date Received"],
        "raw_status": row["raw_status"],
        "Total Items": int(row["Total Items"]) if pd.notna(row["Total Items"]) else 0,
        "Price": float(row["Price"]) if pd.notna(row["Price"]) else 0.0,
        "# of Cartons": int(row["# of Cartons"]) if pd.notna(row["# of Cartons"]) else 0,
        "Notes": row["Notes"], "Items Ordered": row["Items Ordered"],
        "Img ref.": row["Img ref."], "Shopify Inventory Status": row["Shopify Inventory Status"],
        "Marked as Ordered": row["Marked as Ordered"],
    }

SHOPIFY_ADDED_MARKER = "Added to Shopify"
MARKED_ORDERED_MARKER = "Marked as Ordered"

def mark_batch_as_ordered(row):
    """Flags a batch as ordered — independent of Shopify — so Demand & Reorder
    and Purchase Orders count its quantities as incoming. Doesn't touch
    Shopify or the Studio inventory at all."""
    update_shipment(int(row["row"]), {**_row_to_fields(row), "Marked as Ordered": MARKED_ORDERED_MARKER})

def receive_batch_to_shopify(line_items):
    """Adds a batch's line-item quantities to Shopify's on-hand inventory —
    additive, same mechanism as Restock, so it's safe to run before the batch
    has physically arrived (e.g. to start pre-selling incoming stock)."""
    variant_map = fetch_shopify_variant_map()
    location_id = get_primary_location_id()
    synced, unmatched, failed = 0, [], []
    for it in line_items:
        label = f"{it['product']} — {it['color']} / {it['size']}"
        inv_item_id = find_shopify_inventory_item(variant_map, it["product"], it["color"], it["size"])
        if inv_item_id is None:
            unmatched.append(label)
            continue
        try:
            add_shopify_onhand_quantity(inv_item_id, location_id, it["qty"])
            synced += 1
        except Exception as e:
            failed.append(f"{label}: {e}")
    return synced, unmatched, failed

def add_batch_to_shopify(row, line_items):
    """Adds a batch's quantities to Shopify's on-hand inventory and marks it
    so Demand & Reorder and Purchase Orders count it as incoming — nothing
    touches Shopify or those numbers before this runs."""
    synced, unmatched, failed = receive_batch_to_shopify(line_items)
    if synced:
        update_shipment(int(row["row"]), {**_row_to_fields(row), "Shopify Inventory Status": SHOPIFY_ADDED_MARKER})
    return synced, unmatched, failed

def receive_batch_to_studio_inventory(line_items):
    """Adds a batch's line-item quantities into the Studio's own Inventory
    sheet — the one Restock/Fulfillment actually read from — for when the
    stock has genuinely arrived and is physically on hand."""
    inv = load_inventory()
    ws = get_ws()
    row_totals = {}
    unmatched = []
    for it in line_items:
        key = find_key(inv, it["product"], it["color"], it["size"])
        if key is None:
            unmatched.append(f"{it['product']} — {it['color']} / {it['size']}")
            continue
        row_num = inv[key]["row"]
        if row_num not in row_totals:
            row_totals[row_num] = inv[key]["qty"]
        row_totals[row_num] += it["qty"]
    if row_totals:
        batch_update_qty(ws, list(row_totals.items()))
    return len(row_totals), unmatched

def delete_shipment_batch(row_num, batch_name):
    """Removes the batch's row from Shipments Tracker (if it has one) and any
    of its Shipment Line Items rows."""
    if row_num:
        get_shipment_tracker_ws().delete_rows(int(row_num))
        load_shipments.clear()
    save_line_items(batch_name, [])

def delete_shipment_detail_file(file_id):
    _drive().files().delete(fileId=file_id).execute()
    list_shipment_detail_files.clear()

SHIPMENT_LINE_ITEMS_TAB = "Shipment Line Items"

@_resilient_google_call
def get_shipment_line_items_ws():
    sh = _shipment_tracker_spreadsheet()
    try:
        return sh.worksheet(SHIPMENT_LINE_ITEMS_TAB)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SHIPMENT_LINE_ITEMS_TAB, rows=1000, cols=6)
        ws.append_row(["Batch #", "Product", "Color", "Size", "Qty", "Unit Price"])
        return ws

@st.cache_data(ttl=120)
@_resilient_google_call
def load_line_items():
    """Per-(batch, product, color, size) quantities entered in-app — the structured
    alternative to uploading an Excel file for Shipment Details."""
    ws = get_shipment_line_items_ws()
    data = ws.get_all_values()
    items = []
    for row in data[1:]:
        if len(row) < 6 or not row[0].strip():
            continue
        try:
            qty = int(float(row[4]))
        except (ValueError, TypeError):
            qty = 0
        try:
            unit_price = float(row[5])
        except (ValueError, TypeError):
            unit_price = None
        items.append({
            "batch": row[0].strip(), "product": row[1].strip(),
            "color": row[2].strip(), "size": row[3].strip(),
            "qty": qty, "unit_price": unit_price,
        })
    return items

def save_line_items(batch, entries):
    """Replaces this batch's line items with `entries` (list of dicts with product,
    color, size, qty, unit_price) — entries with qty <= 0 are dropped."""
    ws = get_shipment_line_items_ws()
    all_values = ws.get_all_values()
    header = all_values[0] if all_values else ["Batch #", "Product", "Color", "Size", "Qty", "Unit Price"]
    kept = [r for r in all_values[1:] if r and r[0].strip() != batch]
    new_rows = [[batch, e["product"], e["color"], e["size"], e["qty"], e["unit_price"]] for e in entries if e["qty"] > 0]
    ws.clear()
    ws.append_rows([header] + kept + new_rows)
    load_line_items.clear()

SIZE_ORDER = ["XXS", "XS", "S", "M", "L", "XL", "XXL", "2XL", "3XL", "4XL", "5XL"]

def _size_sort_key(size):
    """Sorts sizes as S/M/L/XL.../2XL/3XL instead of alphabetically. Combo sizes
    like 'XS/S' or '2XL/3XL' sort by their first part; anything unrecognized
    sorts after all known sizes, alphabetically."""
    s = str(size).strip().upper()
    first = s.split("/")[0].strip()
    if first in SIZE_ORDER:
        return (0, SIZE_ORDER.index(first))
    return (1, s)

def line_items_to_products(batch, all_items):
    """Groups one batch's line items into the same shape parse_shipment_detail()
    produces, so Shipment Details can render in-app-entered batches with the same
    product_grid() rendering used for uploaded Excel files."""
    batch_items = [it for it in all_items if it["batch"] == batch]
    by_product = {}
    for it in batch_items:
        by_product.setdefault(it["product"], []).append(it)

    products = []
    for name, items in by_product.items():
        sizes = sorted({it["size"] for it in items}, key=_size_sort_key)
        colors = sorted({it["color"] for it in items})
        color_rows = []
        for c in colors:
            qtys = {s: 0 for s in sizes}
            for it in items:
                if it["color"] == c:
                    qtys[it["size"]] = it["qty"]
            color_rows.append((c, qtys))
        total_qty = sum(it["qty"] for it in items)
        unit_price = next((it["unit_price"] for it in items if it["unit_price"]), None)
        subtotal = (unit_price * total_qty) if unit_price is not None else None
        products.append({
            "name": name, "size_cols": [(s, s) for s in sizes], "total_col": None,
            "colors": color_rows, "unit_price": unit_price,
            "subtotal": subtotal, "total_qty": total_qty,
        })
    subtotals = [p["subtotal"] for p in products if p["subtotal"] is not None]
    grand_total = sum(subtotals) if subtotals else None
    return products, grand_total

@st.cache_data(ttl=120)
@_resilient_google_call
def load_packaging_tables():
    ws = get_shipment_tracker_ws()
    values = ws.get("V2:AB100", value_render_option="UNFORMATTED_VALUE")
    used, orders = [], []
    for row in values[1:] if values else []:
        row = row + [None] * (7 - len(row))
        if row[0]:
            used.append({"Batch": row[0], "Small 15×15": _as_number(row[1]), "Big 35×25": _as_number(row[2])})
        if row[4]:
            orders.append({"Order": row[4], "Small 15×15": _as_number(row[5]), "Big 35×25": _as_number(row[6])})
    return pd.DataFrame(used), pd.DataFrame(orders)

def compute_incoming():
    """Sums quantities for batches that have been explicitly Marked as
    Ordered (independent of whether they've also been added to Shopify) and
    aren't yet Received or Cancelled, per (product, color, size) — for
    Purchase Orders and for Demand & Reorder to subtract before suggesting
    how much more to order. A batch created but not yet marked counts toward
    neither — nothing about Demand & Reorder's numbers changes until that
    button is pressed.

    Returns (qty_by_key, detail_rows, open_batches):
      qty_by_key   — keyed like load_inventory() (lowercased tuples) -> qty
      detail_rows  — display-ready per-variant breakdown with contributing batches
      open_batches — set of batch names counted as still incoming
    """
    tracker_df, _ = load_shipments()
    open_batches = (
        set(tracker_df[
            ~tracker_df["Status"].isin(["Received", "Cancelled"])
            & (tracker_df["Marked as Ordered"].fillna("").str.strip() == MARKED_ORDERED_MARKER)
        ]["Batch #"])
        if not tracker_df.empty else set()
    )

    qty_by_key, batches_by_key, display_by_key = {}, {}, {}
    for it in load_line_items():
        if it["batch"] not in open_batches:
            continue
        key = (it["product"].lower(), it["color"].lower(), it["size"].lower())
        qty_by_key[key] = qty_by_key.get(key, 0) + it["qty"]
        batches_by_key.setdefault(key, set()).add(it["batch"])
        display_by_key[key] = (it["product"], it["color"], it["size"])

    detail_rows = [
        {
            "Product": display_by_key[key][0], "Color": display_by_key[key][1], "Size": display_by_key[key][2],
            "Incoming Qty": qty, "Batches": ", ".join(sorted(batches_by_key[key], key=_batch_num)),
        }
        for key, qty in qty_by_key.items()
    ]
    return qty_by_key, detail_rows, open_batches

# ─── Shipment Details ──────────────────────────────────────────────────────────

SHIPMENT_DETAILS_FOLDER_ID = "1fzGaxnG9fu0dgsUgWIGQYXtKgw0u0_Ex"

def _batch_num(name):
    m = re.search(r"batch[_\s]*(\d+)", name, re.IGNORECASE)
    return int(m.group(1)) if m else -1

@st.cache_data(ttl=300)
@_resilient_google_call
def list_shipment_detail_files():
    """Flat list of real shipment files in the Drive folder (recurses one level into
    subfolders — some batches, e.g. split shipments, are grouped in their own subfolder).
    Excel's temporary lock files (~$...) are skipped."""
    drive = _drive()

    def _list_children(folder_id):
        res = drive.files().list(
            q=f"'{folder_id}' in parents and trashed = false",
            fields="files(id, name, mimeType)",
            pageSize=200,
        ).execute()
        return res.get("files", [])

    files = []
    for f in _list_children(SHIPMENT_DETAILS_FOLDER_ID):
        if f["name"].startswith("~$"):
            continue
        if f["mimeType"] == "application/vnd.google-apps.folder":
            files.extend(sub for sub in _list_children(f["id"]) if not sub["name"].startswith("~$"))
        else:
            files.append(f)

    files.sort(key=lambda f: (_batch_num(f["name"]), f["name"]), reverse=True)
    return files

@st.cache_data(ttl=300)
@_resilient_google_call
def fetch_shipment_detail_bytes(file_id):
    return bytes(_drive().files().get_media(fileId=file_id).execute())

def parse_shipment_detail(file_bytes):
    """Parses the repeating product-block layout: a product name row, a size-header row,
    one row per color, a per-product TOTAL row, and a unit-price/subtotal row — repeated
    per product, ending in a GRAND TOTAL row. Column widths vary per product and per file,
    so this walks row-by-row using content markers rather than fixed column positions."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb[wb.sheetnames[0]]

    rows = []
    for r in range(2, ws.max_row + 1):  # row 1 is always the "Batch N — Date" title
        row = {c: ws.cell(row=r, column=c).value
               for c in range(2, ws.max_column + 1)
               if ws.cell(row=r, column=c).value not in (None, "")}
        if row:
            rows.append(row)

    products, grand_total = [], None
    state, current = "EXPECT_PRODUCT", None

    for row in rows:
        b = row.get(2)
        if isinstance(b, str) and b.strip().upper() == "GRAND TOTAL":
            nums = [v for c, v in row.items() if c != 2 and isinstance(v, (int, float))]
            grand_total = nums[-1] if nums else None
            continue

        if state == "EXPECT_PRODUCT":
            if isinstance(b, str) and len(row) == 1:
                current = {"name": b.strip(), "size_cols": [], "total_col": None,
                           "colors": [], "unit_price": None, "subtotal": None, "total_qty": None}
                state = "EXPECT_SIZE_HEADER"
            continue

        if state == "EXPECT_SIZE_HEADER":
            if b is None:
                entries = [(c, str(v).strip()) for c, v in row.items() if c >= 3]
                total_col = next((c for c, v in entries if v.upper() == "TOTAL"), None)
                current["size_cols"] = [(c, v) for c, v in entries if c != total_col]
                current["total_col"] = total_col
                state = "IN_COLORS"
            continue

        if state == "IN_COLORS":
            if isinstance(b, str) and b.strip().upper() == "TOTAL":
                current["total_qty"] = row.get(current["total_col"])
                state = "EXPECT_PRICE"
            elif isinstance(b, str) and current["total_col"]:
                qtys = {c: row.get(c) for c, _ in current["size_cols"]}
                current["colors"].append((b.strip(), qtys))
            continue

        if state == "EXPECT_PRICE":
            if current["total_col"]:
                price_text = row.get(current["total_col"] - 1)
                if isinstance(price_text, str):
                    m = re.search(r"[\d.]+", price_text)
                    if m:
                        current["unit_price"] = float(m.group())
                current["subtotal"] = row.get(current["total_col"])
            products.append(current)
            current, state = None, "EXPECT_PRODUCT"
            continue

    return products, grand_total

def product_grid(product):
    """Wide Color × Size grid, mirroring the original sheet layout, for one product block."""
    data = {label: [qtys.get(col) or 0 for _, qtys in product["colors"]]
            for col, label in product["size_cols"]}
    df = pd.DataFrame(data, index=[c for c, _ in product["colors"]])
    if not df.empty:
        df["Total"] = df.sum(axis=1)
    return df

def format_order_summary(batch_name, products, grand_total):
    """Plain-text order summary, formatted to paste straight into a message to a supplier."""
    lines = [batch_name, ""]
    for p in products:
        lines.append(p["name"])
        for color, qtys in p["colors"]:
            parts = ", ".join(f"{label}={qtys.get(col) or 0}" for col, label in p["size_cols"])
            total = sum(qtys.get(col) or 0 for col, label in p["size_cols"])
            lines.append(f"  {color}: {parts} (Total: {total})")
        price_str = f"${p['unit_price']:,.2f}" if p["unit_price"] is not None else "—"
        subtotal_str = f"${p['subtotal']:,.2f}" if p["subtotal"] is not None else "—"
        lines.append(f"  Unit price: {price_str} | Subtotal: {subtotal_str}")
        lines.append("")
    if grand_total is not None:
        lines.append(f"Grand Total: ${grand_total:,.2f}")
    return "\n".join(lines)

def build_order_excel(batch_name, products, grand_total):
    """One-sheet order form — a Color × Size grid per product, unit price and
    subtotal, and a grand total — ready to download and send to a supplier."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Order"
    bold = Font(bold=True)

    row = 1
    ws.cell(row=row, column=1, value=batch_name).font = Font(bold=True, size=14)
    row += 2

    max_col = 1
    for p in products:
        ws.cell(row=row, column=1, value=p["name"]).font = bold
        row += 1

        ws.cell(row=row, column=1, value="Color").font = bold
        for i, (_, label) in enumerate(p["size_cols"], start=2):
            ws.cell(row=row, column=i, value=label).font = bold
        total_col = len(p["size_cols"]) + 2
        ws.cell(row=row, column=total_col, value="Total").font = bold
        max_col = max(max_col, total_col)
        row += 1

        for color, qtys in p["colors"]:
            ws.cell(row=row, column=1, value=color)
            total = 0
            for i, (col, _) in enumerate(p["size_cols"], start=2):
                qty = qtys.get(col) or 0
                ws.cell(row=row, column=i, value=qty)
                total += qty
            ws.cell(row=row, column=total_col, value=total).font = bold
            row += 1

        ws.cell(row=row, column=1, value="Unit Price").font = bold
        ws.cell(row=row, column=2, value=p["unit_price"])
        ws.cell(row=row, column=3, value="Subtotal").font = bold
        ws.cell(row=row, column=4, value=p["subtotal"])
        row += 2

    ws.cell(row=row, column=1, value="Grand Total").font = Font(bold=True, size=12)
    ws.cell(row=row, column=2, value=grand_total).font = Font(bold=True, size=12)

    for col_idx in range(1, max_col + 1):
        letter = openpyxl.utils.get_column_letter(col_idx)
        ws.column_dimensions[letter].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ─── Session State ────────────────────────────────────────────────────────────

_defaults = dict(
    fulfillable=None, skipped=None, orig_inv=None, new_inv=None,
    ws=None, preview_done=False, removed=set(),
    fulfilled=False, report_buf=None, report_name=None,
    shopify_errors=[],
)
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─── CSS ──────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

:root {
    --rr-red: #e6394f;
    --rr-red-dark: #c72e42;
    --rr-red-soft: #fbe3e6;
    --rr-on-red: #fff;
    --rr-red-text: #e6394f;
    --rr-sidebar-bg: #1f232c;
    --rr-sidebar-text: #aeb2bd;
    --rr-bg: #f4f5f7;
    --rr-border: #e5e7eb;
}

/* ── Base ── */
html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif;
    background: var(--rr-bg);
    color: #1f232c;
}
[data-testid="stAppViewContainer"] > .main {
    background: var(--rr-bg);
}

/* ── Sidebar ── */
[data-testid="stSidebar"] {
    background: #eef0f4 !important;
    border-right: 1px solid var(--rr-border);
}
[data-testid="stSidebarContent"] {
    padding: 1.6rem 1rem;
}
.sidebar-brand {
    font-family: 'Inter', sans-serif;
    font-size: 1.1rem;
    font-weight: 700;
    color: #1f232c;
    letter-spacing: 0;
    text-transform: none;
    margin-bottom: 0.2rem;
}
.sidebar-tagline {
    font-size: 0.65rem;
    color: #6b6f7b;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    margin-bottom: 1.6rem;
}
.sidebar-section {
    font-size: 0.62rem;
    font-weight: 600;
    color: #6b6f7b;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin: 0.9rem 0.8rem 0.3rem;
}
[data-testid="stSidebar"] hr {
    border-color: var(--rr-border);
    margin: 1.1rem 0;
}
[data-testid="stSidebar"] p, [data-testid="stSidebar"] .stCaption {
    color: #6b6f7b !important;
}
[data-testid="stSidebar"] .stButton > button {
    border: 1px solid var(--rr-border) !important;
    color: #1f232c !important;
    background: #fff !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
}
[data-testid="stSidebar"] .stButton > button:hover:not(.st-key-sidebar_nav *) {
    background: var(--rr-red) !important;
    border-color: var(--rr-red) !important;
    color: #fff !important;
}

/* ── Sidebar nav (icon buttons) ──
   Every pseudo-state pairs its own background + text/icon color in one rule so
   there's never a gap where one property is set but not the other (that gap is
   what caused white-on-white at some interaction states before). Icons render
   as inline SVG (fill), labels as text (color) — both are set together. */
.st-key-sidebar_nav .stButton { margin-bottom: 0.5rem; }
.st-key-sidebar_nav .stButton > button {
    border: 1px solid var(--rr-border) !important;
    font-size: 0.9rem !important;
    font-weight: 500 !important;
    padding: 0.7rem 0.8rem !important;
    border-radius: 10px !important;
    width: 100% !important;
    box-shadow: 0 1px 2px rgba(0,0,0,0.05) !important;
    outline: none !important;
}

.st-key-sidebar_nav .stButton > button[kind="secondary"],
.st-key-sidebar_nav .stButton > button[kind="secondary"]:focus,
.st-key-sidebar_nav .stButton > button[kind="secondary"]:focus-visible {
    background: #fff !important;
    border-color: var(--rr-border) !important;
}
.st-key-sidebar_nav .stButton > button[kind="secondary"]:hover {
    background: #f7f8fa !important;
    border-color: #d0d3d9 !important;
}
.st-key-sidebar_nav .stButton > button[kind="secondary"]:active {
    background: var(--rr-red) !important;
    border-color: var(--rr-red) !important;
}
.st-key-sidebar_nav .stButton > button[kind="secondary"] *,
.st-key-sidebar_nav .stButton > button[kind="secondary"]:focus *,
.st-key-sidebar_nav .stButton > button[kind="secondary"]:hover * {
    color: #1f232c !important; fill: #1f232c !important;
}
.st-key-sidebar_nav .stButton > button[kind="secondary"]:active * {
    color: #fff !important; fill: #fff !important;
}

.st-key-sidebar_nav .stButton > button[kind="primary"],
.st-key-sidebar_nav .stButton > button[kind="primary"]:focus,
.st-key-sidebar_nav .stButton > button[kind="primary"]:focus-visible {
    background: var(--rr-red) !important;
    border-color: var(--rr-red) !important;
}
.st-key-sidebar_nav .stButton > button[kind="primary"]:hover,
.st-key-sidebar_nav .stButton > button[kind="primary"]:active {
    background: var(--rr-red-dark) !important;
    border-color: var(--rr-red-dark) !important;
}
.st-key-sidebar_nav .stButton > button[kind="primary"] *,
.st-key-sidebar_nav .stButton > button[kind="primary"]:focus *,
.st-key-sidebar_nav .stButton > button[kind="primary"]:hover *,
.st-key-sidebar_nav .stButton > button[kind="primary"]:active * {
    color: #fff !important; fill: #fff !important;
}

/* ── Buttons ── */
.stButton > button {
    font-family: 'Inter', sans-serif !important;
    font-size: 0.82rem !important;
    font-weight: 600 !important;
    letter-spacing: 0 !important;
    text-transform: none !important;
    border-radius: 6px !important;
    padding: 0.55rem 1.4rem !important;
    transition: all 0.15s ease !important;
    border: 1.5px solid var(--rr-red) !important;
    background: transparent !important;
    box-shadow: none !important;
    outline: none !important;
}
.stButton > button,
.stButton > button:focus,
.stButton > button:focus-visible {
    color: var(--rr-red) !important;
}
.stButton > button *,
.stButton > button:focus *,
.stButton > button:focus-visible * {
    color: var(--rr-red) !important; fill: var(--rr-red) !important;
}
/* Hover is a light tint, not a full fill — a secondary button that's merely
   sitting under the cursor (e.g. right after a click triggers a rerun and
   the mouse hasn't moved) shouldn't look identical to something actually
   selected/pressed. Only :active (an actual mouse-down) and primary buttons
   get the bold solid fill. */
.stButton > button:hover {
    background: var(--rr-red-soft) !important;
    border-color: var(--rr-red) !important;
}
.stButton > button:hover * {
    color: var(--rr-red) !important; fill: var(--rr-red) !important;
}
.stButton > button:active,
.stButton > button[kind="primary"] {
    background: var(--rr-red) !important;
    border-color: var(--rr-red) !important;
}
.stButton > button:active *,
.stButton > button[kind="primary"] * {
    color: white !important; fill: white !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[kind="primary"]:active {
    background: var(--rr-red-dark) !important;
    border-color: var(--rr-red-dark) !important;
}
.stButton > button[kind="primary"]:hover *,
.stButton > button[kind="primary"]:active * {
    color: white !important; fill: white !important;
}

/* ── Inputs ── */
.stTextInput input,
.stTextArea textarea,
.stNumberInput input {
    font-family: 'Inter', sans-serif !important;
    font-size: 0.88rem !important;
    border-radius: 6px !important;
    border: 1px solid var(--rr-border) !important;
    background: #fff !important;
    color: #1f232c !important;
}
.stTextInput input:focus,
.stTextArea textarea:focus {
    border-color: var(--rr-red) !important;
    box-shadow: 0 0 0 2px rgba(230,57,79,0.12) !important;
}
.stSelectbox > div > div {
    border-radius: 6px !important;
    border: 1px solid var(--rr-border) !important;
    font-size: 0.88rem !important;
}
label[data-testid="stWidgetLabel"] p {
    font-size: 0.72rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.06em !important;
    text-transform: uppercase !important;
    color: #6b6f7b !important;
}

/* ── Tabs ── */
.stTabs [data-testid="stTab"] {
    font-size: 0.8rem !important;
    font-weight: 600 !important;
    letter-spacing: 0 !important;
    text-transform: none !important;
    color: #6b6f7b !important;
}
.stTabs [data-testid="stTab"][aria-selected="true"] {
    color: var(--rr-red-text) !important;
    border-bottom-color: var(--rr-red-text) !important;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border: 1.5px dashed var(--rr-border) !important;
    border-radius: 8px !important;
    background: #fff !important;
    padding: 1rem !important;
}

/* ── Dataframe ── */
[data-testid="stDataFrame"] {
    border: 1px solid var(--rr-border) !important;
    border-radius: 8px !important;
    overflow: hidden;
}

/* ── Alerts ── */
[data-testid="stAlert"] {
    border-radius: 6px !important;
    font-size: 0.88rem !important;
}

/* ── Checkbox ── */
.stCheckbox label p {
    font-size: 0.85rem !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    font-weight: 400 !important;
    color: #1f232c !important;
}

/* ── Divider ── */
hr { border-color: var(--rr-border) !important; }

/* ── Brand header (top bar) ── */
.brand-header {
    background: var(--rr-red);
    border-radius: 8px;
    padding: 1.1rem 1.6rem;
    margin-bottom: 1.6rem;
}
.brand-header-eyebrow {
    font-family: 'Inter', sans-serif;
    font-size: 0.65rem;
    font-weight: 600;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: rgba(255,255,255,0.7);
    margin: 0 0 0.2rem;
}
.brand-header h1 {
    font-family: 'Inter', sans-serif;
    font-size: 1.5rem;
    font-weight: 700;
    color: var(--rr-on-red);
    margin: 0;
    letter-spacing: 0;
    line-height: 1.2;
}
.brand-header p {
    font-size: 0.72rem;
    font-weight: 400;
    letter-spacing: 0.04em;
    text-transform: none;
    color: rgba(255,255,255,0.75);
    margin: 0.25rem 0 0;
}

/* ── Stat cards ── */
.stat {
    background: #fff;
    border: 1px solid var(--rr-border);
    border-radius: 8px;
    padding: 1.2rem 1.4rem;
    text-align: left;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}
.stat .num {
    font-family: 'Inter', sans-serif;
    font-size: 2rem;
    font-weight: 700;
    line-height: 1;
    margin: 0;
}
.stat .lbl {
    font-size: 0.68rem;
    font-weight: 600;
    color: #8b8f9b;
    margin: 0.4rem 0 0;
    text-transform: uppercase;
    letter-spacing: 0.08em;
}

/* ── Pills (status/stage badges — Cancelled Orders, Refunds) ── */
.rr-pill {
    display: inline-flex; align-items: center; gap: 5px;
    font-size: 0.66rem; font-weight: 700; padding: 4px 10px; border-radius: 999px;
    text-transform: uppercase; letter-spacing: 0.04em; white-space: nowrap;
}
.rr-pill-amber { background: #fff6e0; color: #966600; }
.rr-pill-green { background: #eaf7ed; color: #227A55; }
.rr-pill-red   { background: #fdecea; color: #B03A3A; }
.rr-pill-blue  { background: #eef2f7; color: #2A5C8A; }

/* ── Dashboard table (Shipment Tracker) ── */
.rr-table-wrap {
    background: #fff;
    border: 1px solid var(--rr-border);
    border-radius: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04);
    overflow-x: auto;
    margin-bottom: 0.5rem;
}
.rr-table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.82rem;
    white-space: nowrap;
}
.rr-table thead th {
    position: sticky;
    top: 0;
    background: #fafbfc;
    color: #6b6f7b;
    font-size: 0.64rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    text-align: left;
    padding: 0.7rem 0.9rem;
    border-bottom: 1px solid var(--rr-border);
}
.rr-table tbody td {
    padding: 0.65rem 0.9rem;
    border-bottom: 1px solid #f0f1f3;
    color: #1f232c;
    vertical-align: middle;
}
.rr-table tbody tr:last-child td { border-bottom: none; }
.rr-table tbody tr:hover td { background: #fafbfc; }
.rr-t-strong { font-weight: 700; }
.rr-t-mono { font-family: 'SFMono-Regular', Consolas, monospace; font-size: 0.78rem; color: #4a4e58; }
.rr-t-num, th.rr-t-num { text-align: right; font-variant-numeric: tabular-nums; }
.rr-t-trunc {
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
    max-width: 220px; display: table-cell;
}

/* ── Avatar circle (assigned employee initial) ── */
.rr-avatar {
    width: 30px; height: 30px; border-radius: 999px; flex: none;
    display: inline-flex; align-items: center; justify-content: center;
    font-size: 12px; font-weight: 700; color: #fff; font-family: 'Inter', sans-serif;
}

/* ── Row cards (Cancelled Orders queue, Refunds' update panel) ── */
[class*="st-key-co_card_"], .st-key-rf_update_panel {
    border-radius: 14px !important;
    transition: box-shadow .15s ease, border-color .15s ease;
}
[class*="st-key-co_card_"]:hover {
    border-color: #d7dae0 !important;
    box-shadow: 0 4px 16px rgba(31,35,44,0.07) !important;
}

/* ── WhatsApp-branded link buttons (st.link_button has no default theming) ── */
.stLinkButton a {
    font-family: 'Inter', sans-serif !important;
    font-size: 0.82rem !important;
    font-weight: 700 !important;
    border-radius: 6px !important;
    background: #1D8A4E !important;
    border: 1.5px solid #1D8A4E !important;
    color: #fff !important;
    box-shadow: none !important;
    transition: filter .15s ease !important;
}
.stLinkButton a:hover { filter: brightness(1.08); color: #fff !important; }
.stLinkButton a p { color: #fff !important; }

/* ── Subheaders ── */
h2 {
    font-family: 'Inter', sans-serif !important;
    font-weight: 700 !important;
    font-size: 1.3rem !important;
    color: #1f232c !important;
    letter-spacing: 0 !important;
}
h3 {
    font-family: 'Inter', sans-serif !important;
    font-size: 0.75rem !important;
    font-weight: 600 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    color: #6b6f7b !important;
}

/* ── Number inputs: hide the +/- stepper buttons ── */
button[data-testid="stNumberInputStepUp"],
button[data-testid="stNumberInputStepDown"] {
    display: none !important;
}
[data-testid="stNumberInputContainer"] {
    border-radius: 0.5rem !important;
}
</style>
""", unsafe_allow_html=True)

# ─── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("""
    <div style="display:flex;align-items:center;gap:0.6rem;margin-bottom:0.2rem;">
      <div style="width:34px;height:34px;min-width:34px;border-radius:50%;background:#e6394f;
                  display:flex;align-items:center;justify-content:center;
                  font-family:'Inter',sans-serif;font-weight:700;font-size:0.8rem;color:#fff;">SM</div>
      <p class="sidebar-brand" style="margin-bottom:0;">Sweet Mayhem</p>
    </div>
    <p class="sidebar-tagline">Fulfillment Studio</p>
    """, unsafe_allow_html=True)

    my_pages = st.session_state.get("pages", [])
    if not my_pages:
        st.warning("Your account has no page access yet. Ask an admin to assign some.")
        st.stop()

    # Always show pages in the same canonical order (grouped by section below),
    # regardless of the order they happen to be stored in for this user.
    _page_order = {p: i for i, p in enumerate(ALL_PAGES + [ADMIN_PAGE])}
    my_pages = sorted(my_pages, key=lambda p: _page_order.get(p, len(_page_order)))

    if "page" not in st.session_state:
        qp_page = SLUG_TO_PAGE.get(st.query_params.get("page"))
        st.session_state.page = qp_page if qp_page in my_pages else my_pages[0]
    elif st.session_state.page not in my_pages:
        st.session_state.page = my_pages[0]
    st.query_params["page"] = PAGE_SLUGS[st.session_state.page]

    NAV_SECTIONS = {
        "📦 Fulfillment": "INVENTORY", "🔄 Restock": "INVENTORY",
        "➕ Add Product": "INVENTORY", "📋 View Inventory": "INVENTORY",
        "📊 Demand & Reorder": "SHIPMENTS", "📥 Purchase Orders": "SHIPMENTS",
        "🚢 Shipment Tracker": "SHIPMENTS", "🧾 Shipment Details": "SHIPMENTS",
        "🚫 Cancelled Orders": "SUPPORT", "💸 Refunds": "SUPPORT",
    }
    with st.container(key="sidebar_nav"):
        last_section = None
        for p in my_pages:
            section = NAV_SECTIONS.get(p)
            if section and section != last_section:
                st.markdown(f'<p class="sidebar-section">{section}</p>', unsafe_allow_html=True)
                last_section = section
            if st.button(
                page_label(p), icon=PAGE_ICONS.get(p, ":material/circle:"),
                key=f"nav_{p}", use_container_width=True,
                type="primary" if st.session_state.page == p else "secondary",
            ):
                st.session_state.page = p
                st.session_state._close_sidebar_on_nav = True
                st.rerun()
    page = st.session_state.page

    if st.session_state.pop("_close_sidebar_on_nav", False):
        # The <!-- nonce --> makes this HTML string unique every render. Without it,
        # an identical srcdoc on a later rerun doesn't reload the iframe, so the
        # <script> only ever fires the very first time this appears — which is why
        # closing worked once after login and then silently stopped.
        components.html(
            f"""
            <!-- {datetime.now().isoformat()} -->
            <script>
            try {{
                const el = window.parent.document.querySelector('[data-testid="stSidebarCollapseButton"] button')
                        || window.parent.document.querySelector('[data-testid="stSidebarCollapseButton"]');
                if (el) el.click();
            }} catch (e) {{}}
            </script>
            """,
            height=0,
        )

    st.divider()
    st.caption(f"Signed in as **{st.session_state.username}**  ·  {st.session_state.role}")
    if st.button("Sign Out", use_container_width=True):
        me = load_users().get(st.session_state.username.lower())
        if me:
            revoke_remember_token(me["row"])
        st.query_params.pop("t", None)
        for k in ("authenticated", "username", "role", "pages"):
            st.session_state.pop(k, None)
        st.rerun()

# ─── Header ───────────────────────────────────────────────────────────────────

st.markdown("""
<div class="brand-header">
  <p class="brand-header-eyebrow">Operations Dashboard</p>
  <h1>Sweet Mayhem</h1>
  <p>Order Fulfillment &amp; Inventory</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: FULFILLMENT
# ─────────────────────────────────────────────────────────────────────────────

if page == "📦 Fulfillment":

    fetch_btn = st.button("Fetch All Unfulfilled", icon=":material/refresh:", use_container_width=True)

    sc1, sc2 = st.columns([3, 1])
    order_input = sc1.text_input("", placeholder="Order number e.g. 17234", label_visibility="collapsed")
    single_btn  = sc2.button("Fetch Order", icon=":material/search:", use_container_width=True)

    st.divider()

    # ── Fetch orders from Shopify ──────────────────────────────────────────────

    if fetch_btn:
        with st.spinner("Fetching unfulfilled orders from Shopify…"):
            try:
                inv, ws  = load_inventory(), get_ws()
                orders   = fetch_shopify_orders()
                if not orders:
                    st.warning("No unfulfilled orders found in Shopify.")
                    st.stop()
                fulfillable, skipped, new_inv = determine_fulfillable(orders, inv)
                st.session_state.update(
                    fulfillable=fulfillable, skipped=skipped,
                    orig_inv=inv, new_inv=new_inv, ws=ws,
                    preview_done=True, removed=set(),
                    fulfilled=False, report_buf=None,
                )
                st.success(f"Loaded {len(orders)} unfulfilled order(s) from Shopify.")
            except Exception as e:
                st.error(str(e))
                st.stop()

    if single_btn:
        if not order_input.strip():
            st.warning("Enter an order number first.")
            st.stop()
        with st.spinner(f"Fetching order #{order_input.strip().lstrip('#')} from Shopify…"):
            try:
                inv, ws  = load_inventory(), get_ws()
                orders   = fetch_shopify_order_by_name(order_input)
                fulfillable, skipped, new_inv = determine_fulfillable(orders, inv)
                st.session_state.update(
                    fulfillable=fulfillable, skipped=skipped,
                    orig_inv=inv, new_inv=new_inv, ws=ws,
                    preview_done=True, removed=set(),
                    fulfilled=False, report_buf=None,
                )
                st.success(f"Loaded order {orders[0]['name']} from Shopify.")
            except Exception as e:
                st.error(str(e))
                st.stop()

    # ── Results ───────────────────────────────────────────────────────────────

    if st.session_state.preview_done:
        fulfillable = [o for o in st.session_state.fulfillable
                       if o["name"] not in st.session_state.removed]
        skipped  = st.session_state.skipped or []
        orig_inv = st.session_state.orig_inv or {}
        cur_inv  = recalc_inv(orig_inv, fulfillable)
        changes  = [k for k in orig_inv if cur_inv.get(k, {}).get("qty") != orig_inv[k]["qty"]]

        no_phone_count = sum(1 for o in fulfillable if not o.get("phone"))

        # Stat cards
        sc1, sc2, sc3, sc4 = st.columns(4)
        sc1.markdown(f'<div class="stat"><p class="num">{len(fulfillable)}</p><p class="lbl">Ready to fulfill</p></div>', unsafe_allow_html=True)
        sc2.markdown(f'<div class="stat"><p class="num">{len(skipped)}</p><p class="lbl">Skipped</p></div>', unsafe_allow_html=True)
        sc3.markdown(f'<div class="stat"><p class="num">{len(changes)}</p><p class="lbl">Inventory changes</p></div>', unsafe_allow_html=True)
        sc4.markdown(f'<div class="stat"><p class="num" style="color:var(--rr-red)">{no_phone_count}</p><p class="lbl">No Phone Number</p></div>', unsafe_allow_html=True)
        st.markdown("")

        tab1, tab2, tab3 = st.tabs([
            ":material/check_circle: To Fulfill", ":material/warning: Skipped",
            ":material/sync_alt: Inventory Changes",
        ])

        # Tab 1 — To Fulfill
        with tab1:
            if not fulfillable:
                st.info("No fulfillable orders (or all removed).")
            else:
                no_phone = [o for o in fulfillable if not o.get("phone")]
                if no_phone:
                    names = ", ".join(f"`{o['name']}`" for o in no_phone)
                    st.warning(
                        f"{len(no_phone)} order(s) have no phone number on file — carriers can "
                        f"fail delivery without one: {names}",
                        icon=":material/phone_disabled:",
                    )

                st.caption("Click :material/close: to remove an order from this run before fulfilling.")
                hc = st.columns([2, 2, 1.3, 4, 1])
                hc[0].markdown("**Order #**")
                hc[1].markdown("**Date**")
                hc[2].markdown("**Phone**")
                hc[3].markdown("**Items**")
                hc[4].markdown("**Remove**")

                for order in fulfillable:
                    date = order["created_at"][:10] if len(order["created_at"]) >= 10 else order["created_at"]
                    items_str = "  ·  ".join(
                        f"{i['name']} ×{i['quantity']}" for i in order["line_items"]
                    )
                    rc = st.columns([2, 2, 1.3, 4, 1])
                    rc[0].markdown(f"`{order['name']}`")
                    rc[1].markdown(date)
                    if order.get("phone"):
                        rc[2].markdown(":material/check_circle:")
                    else:
                        rc[2].markdown(":red[:material/warning: Missing]")
                    rc[3].markdown(items_str)
                    if rc[4].button("", icon=":material/close:", key=f"rm_{order['name']}"):
                        st.session_state.removed.add(order["name"])
                        st.rerun()

                st.divider()
                confirm = st.checkbox(
                    f"I confirm I want to fulfill {len(fulfillable)} order(s) and update inventory"
                )
                if confirm:
                    if st.button("Fulfill These Orders", icon=":material/check_circle:", type="primary", use_container_width=True):
                        errors = []
                        with st.spinner("Updating inventory in Google Sheets…"):
                            try:
                                updates = [(orig_inv[k]["row"], cur_inv[k]["qty"]) for k in changes]
                                batch_update_qty(st.session_state.ws, updates)
                            except Exception as e:
                                errors.append(f"Inventory update failed: {e}")

                        shopify_errs = []
                        with st.spinner("Marking orders as fulfilled in Shopify…"):
                            for order in fulfillable:
                                try:
                                    shopify_fulfill_order(order["id"])
                                except Exception as e:
                                    shopify_errs.append(f"{order['name']}: {e}")

                        buf      = make_report(fulfillable)
                        date_str = datetime.now().strftime("%Y-%m-%d")
                        st.session_state.update(
                            fulfilled=True, preview_done=False,
                            report_buf=buf,
                            report_name=f"fulfilled_{date_str}.xlsx",
                            shopify_errors=shopify_errs,
                        )
                        st.rerun()

        # Tab 2 — Skipped
        with tab2:
            if not skipped:
                st.info("No skipped orders — everything can be fulfilled!")
            else:
                rows = [{"Order #": o["name"],
                         "Date": o["created_at"][:10],
                         "Reason": o.get("skip_reason", "")}
                        for o in skipped]
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

        # Tab 3 — Inventory Changes
        with tab3:
            if not changes:
                st.info("No inventory changes.")
            else:
                rows = []
                for k in changes:
                    o = orig_inv[k]
                    nq = cur_inv[k]["qty"]
                    rows.append({
                        "Product": o["product"], "Color": o["color"],
                        "Size": o["size"], "Before": o["qty"],
                        "After": nq, "Change": f"{nq - o['qty']:+d}",
                    })
                st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    # Post-fulfill download
    if st.session_state.fulfilled and st.session_state.report_buf:
        shopify_errors = st.session_state.get("shopify_errors", [])
        if shopify_errors:
            st.error("Shopify fulfillment failed for some orders. See details below:", icon=":material/error:")
            for err in shopify_errors:
                st.code(err)
        else:
            st.success("Orders marked as fulfilled in Shopify and inventory updated in Google Sheets.", icon=":material/check_circle:")
        st.download_button(
            "Download Fulfillment Report",
            icon=":material/download:",
            data=st.session_state.report_buf,
            file_name=st.session_state.report_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        if st.button("Start New Run", icon=":material/restart_alt:"):
            for k, v in _defaults.items():
                st.session_state[k] = v
            st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: RESTOCK
# ─────────────────────────────────────────────────────────────────────────────

elif page == "🔄 Restock":
    st.subheader("Add Restock Quantities")
    st.caption("Tap a quantity box to add stock, then click Apply Restock.")

    for m in st.session_state.pop("restock_messages", []):
        getattr(st, m["kind"])(m["text"], **({"icon": m["icon"]} if m.get("icon") else {}))
    if st.session_state.pop("restock_balloons", False):
        st.balloons()

    try:
        with st.spinner("Loading inventory…"):
            inv, ws = load_inventory(), get_ws()

        items = list(inv.items())

        restock_products = ["All"] + sorted({v["product"] for _, v in items})
        restock_prod_filter = st.selectbox("Filter by Product", restock_products, key="restock_prod_filter")

        visible_indices = [
            idx for idx, (_, v) in enumerate(items)
            if restock_prod_filter == "All" or v["product"] == restock_prod_filter
        ]

        hc1, hc2, hc3, hc4 = st.columns([3, 2, 1, 2])
        hc1.caption("PRODUCT")
        hc2.caption("COLOR")
        hc3.caption("SIZE")
        hc4.caption("ADD QTY")

        for idx in visible_indices:
            _, item = items[idx]
            c1, c2, c3, c4 = st.columns([3, 2, 1, 2])
            c1.write(item["product"])
            c2.write(item["color"])
            c3.write(item["size"])
            c4.number_input(
                f"Add qty — {item['product']} {item['color']} {item['size']}",
                min_value=0, step=1, value=0,
                key=f"restock_qty_{item['row']}",
                label_visibility="collapsed",
                help=f"Current: {item['qty']}",
            )

        add_qty_by_idx = {}
        for idx, (_, item) in enumerate(items):
            qty = st.session_state.get(f"restock_qty_{item['row']}", 0)
            if qty and qty > 0:
                add_qty_by_idx[idx] = int(qty)

        if add_qty_by_idx:
            st.info(f"{len(add_qty_by_idx)} item(s) with quantities to add.")

        if st.button("Apply Restock", type="primary", use_container_width=True):
            if not add_qty_by_idx:
                st.warning("No quantities entered. Tap a quantity box first.")
            else:
                messages = []
                with st.spinner("Updating Google Sheets…"):
                    try:
                        updates = [
                            (items[idx][1]["row"], items[idx][1]["qty"] + delta)
                            for idx, delta in add_qty_by_idx.items()
                        ]
                        batch_update_qty(ws, updates)
                        messages.append({"kind": "success", "text": f"{len(updates)} item(s) restocked!", "icon": ":material/check_circle:"})
                    except Exception as e:
                        st.error(str(e))
                        st.stop()

                with st.spinner("Adding restocked quantities to Shopify…"):
                    try:
                        variant_map = fetch_shopify_variant_map()
                        location_id = get_primary_location_id()
                        synced, unmatched, failed = 0, [], []
                        for idx, delta in add_qty_by_idx.items():
                            _, item = items[idx]
                            label = f"{item['product']} — {item['color']} / {item['size']}"
                            inv_item_id = find_shopify_inventory_item(
                                variant_map, item["product"], item["color"], item["size"]
                            )
                            if inv_item_id is None:
                                unmatched.append(label)
                                continue
                            try:
                                add_shopify_onhand_quantity(inv_item_id, location_id, delta)
                                synced += 1
                            except Exception as e:
                                failed.append(f"{label}: {e}")
                        if synced:
                            messages.append({"kind": "success", "text": f"{synced} item(s) added to Shopify's on-hand quantity.", "icon": ":material/sync:"})
                        if unmatched:
                            messages.append({"kind": "warning", "text":
                                "Couldn't match to a Shopify variant (Sheet quantity was still "
                                "updated) — check these manually in Shopify:\n\n"
                                + "\n".join(f"- {m}" for m in unmatched)
                            })
                        if failed:
                            messages.append({"kind": "error", "text":
                                "Matched in Shopify but the inventory update failed:\n\n"
                                + "\n".join(f"- {f}" for f in failed)
                            })
                    except Exception as e:
                        messages.append({"kind": "error", "text": str(e)})

                for idx in add_qty_by_idx:
                    st.session_state.pop(f"restock_qty_{items[idx][1]['row']}", None)
                st.session_state["restock_messages"] = messages
                st.session_state["restock_balloons"] = True
                st.rerun()
    except Exception as e:
        st.error(f"Could not load inventory: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: ADD PRODUCT
# ─────────────────────────────────────────────────────────────────────────────

elif page == "➕ Add Product":
    st.subheader("Add New Product to Inventory")

    name = st.text_input("Product Name", placeholder="e.g. High Waist Leggings")
    col1, col2 = st.columns(2)
    with col1:
        colors_raw = st.text_area("Colors (one per line)", placeholder="Black\nNude\nPink", height=160)
    with col2:
        sizes_raw = st.text_area("Sizes (one per line)", value="XS\nS\nM\nL\nXL\nXXL", height=160)
    start_qty = st.number_input("Starting quantity per variant", min_value=0, value=0)

    colors   = [c.strip() for c in colors_raw.splitlines() if c.strip()]
    sizes    = [s.strip() for s in sizes_raw.splitlines() if s.strip()]
    variants = [(name, c, s, start_qty) for c in colors for s in sizes] if name else []

    if variants:
        st.caption(f"Preview — {len(variants)} variant(s):")
        st.dataframe(
            pd.DataFrame(variants, columns=["Product", "Color", "Size", "Qty"]),
            use_container_width=True, hide_index=True,
        )

    if st.button("Add to Inventory", type="primary", disabled=not variants):
        with st.spinner("Adding to Google Sheets…"):
            try:
                ws = get_ws()
                existing_keys = set(load_inventory().keys())
                to_add = [v for v in variants
                          if (v[0].lower(), v[1].lower(), v[2].lower()) not in existing_keys]
                skipped_count = len(variants) - len(to_add)
                if to_add:
                    ws.append_rows([[p, c, s, q] for p, c, s, q in to_add])
                    load_inventory.clear()
                msg = f"{len(to_add)} variant(s) added to inventory."
                if skipped_count:
                    msg += f" ({skipped_count} skipped — already existed.)"
                st.success(msg, icon=":material/check_circle:")
            except Exception as e:
                st.error(str(e))

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: VIEW INVENTORY
# ─────────────────────────────────────────────────────────────────────────────

elif page == "📋 View Inventory":
    st.subheader("Current Inventory")

    try:
        with st.spinner("Loading…"):
            inv = load_inventory()

        df = pd.DataFrame([
            {"Product": v["product"], "Color": v["color"],
             "Size": v["size"], "Qty": v["qty"]}
            for v in inv.values()
        ])

        fc1, fc2 = st.columns(2)
        products = ["All"] + sorted(df["Product"].unique().tolist())
        sel_prod = fc1.selectbox("Filter by Product", products)
        if sel_prod != "All":
            df = df[df["Product"] == sel_prod]
            colors = ["All"] + sorted(df["Color"].unique().tolist())
            sel_col = fc2.selectbox("Filter by Color", colors)
            if sel_col != "All":
                df = df[df["Color"] == sel_col]

        st.dataframe(df, use_container_width=True, hide_index=True)
        st.caption(f"{len(df)} variant(s)  |  {df['Qty'].sum()} total units")

    except Exception as e:
        st.error(f"Could not load inventory: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: DEMAND & REORDER
# ─────────────────────────────────────────────────────────────────────────────

elif page == "📊 Demand & Reorder":
    st.subheader("Demand & Reorder Suggestions")
    st.caption(
        "Sales velocity is adjusted for the days each item was actually out of stock "
        "on Shopify (Available quantity at or below zero), so a stock-out doesn't make "
        "demand look lower than it really is. Logged automatically once a day — "
        "accuracy improves the more often this page is checked."
    )

    today = datetime.now().date()

    try:
        with st.spinner("Loading inventory & recording today's stock snapshot…"):
            inv = load_inventory()
            recorded_today = record_snapshot_if_needed(inv)
            if recorded_today:
                load_snapshots.clear()
            records = load_snapshots()

        tracking_start_date = min(
            (datetime.strptime(r[0], "%Y-%m-%d").date() for r in records),
            default=today,
        )

        d1, d2, d3, d4 = st.columns([1, 1, 1, 1.2])
        start_date = d1.date_input(
            "Sales data from", value=max(tracking_start_date, today - timedelta(days=90)),
            min_value=tracking_start_date, max_value=today,
            help="Set this to a product's launch date to exclude the period before it existed. "
                 f"Can't go earlier than when stock tracking began ({tracking_start_date.strftime('%b %d, %Y')}).",
        )
        end_date = d2.date_input("Sales data to", value=today, min_value=tracking_start_date, max_value=today)
        lead_time = d3.slider("Lead time (days)", min_value=5, max_value=21, value=9,
                               help="Sweet Mayhem's supplier lead time is ~7–10 days.")
        coverage_days = d4.number_input("Target stock coverage (days)", min_value=5, max_value=90, value=25, step=1,
                                         help="Reorder quantity tops stock up to cover this many days of demand.")

        if start_date > end_date:
            st.error("'Sales data from' must be on or before 'Sales data to'.")
            st.stop()

        with st.spinner("Fetching sales history from Shopify…"):
            orders = fetch_shopify_sales(start_date=start_date, end_date=end_date)
            sold, _unmatched = aggregate_sales(orders, inv)

        stock_days = compute_stock_days(records, start_date, end_date)

        tracked_counts = [len(v["tracked"]) for v in stock_days.values()]
        max_tracked = max(tracked_counts) if tracked_counts else 0
        if max_tracked < MIN_TRACKED_DAYS:
            st.info(
                f"{max_tracked} day(s) of stock-history recorded so far. Demand is shown as a "
                f"raw average (unadjusted) until {MIN_TRACKED_DAYS} days are tracked — check back "
                f"as the history builds up.",
                icon=":material/calendar_month:",
            )
        else:
            st.info(
                f"{max_tracked} day(s) of stock-history recorded — adjusted figures below where available.",
                icon=":material/calendar_month:",
            )

        incoming_by_key, _incoming_detail, _open_batches = compute_incoming()

        available_data = fetch_shopify_available_by_key()
        continue_oos_by_key = {}
        for key, v in inv.items():
            info = _match_shopify_available(available_data, v["product"], v["color"], v["size"])
            continue_oos_by_key[key] = bool(info and info.get("continue_oos"))

        df = build_reorder_table(
            inv, sold, stock_days, start_date, end_date, lead_time, coverage_days,
            incoming_by_key, continue_oos_by_key,
        )

        s1, s2, s3, s4 = st.columns(4)
        s1.markdown(f'<div class="stat"><p class="num">{(df["Status"]=="Reorder Now").sum()}</p><p class="lbl">Reorder Now</p></div>', unsafe_allow_html=True)
        s2.markdown(f'<div class="stat"><p class="num">{(df["Status"]=="Reorder Soon").sum()}</p><p class="lbl">Reorder Soon</p></div>', unsafe_allow_html=True)
        s3.markdown(f'<div class="stat"><p class="num">{(df["Status"]=="Out of Stock").sum()}</p><p class="lbl">Out of Stock</p></div>', unsafe_allow_html=True)
        s4.markdown(f'<div class="stat"><p class="num">{int(df["Reorder Qty"].sum())}</p><p class="lbl">Units to Reorder</p></div>', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        fc1, fc2 = st.columns(2)
        status_options = df["Status"].unique().tolist()
        status_filter = fc1.multiselect("Filter by Status", status_options, default=status_options)
        products = ["All"] + sorted(df["Product"].unique().tolist())
        prod_filter = fc2.selectbox("Filter by Product", products)

        fdf = df[df["Status"].isin(status_filter)]
        if prod_filter != "All":
            fdf = fdf[fdf["Product"] == prod_filter]
        fdf = fdf.sort_values("Days Left")

        render_reorder_table(fdf)
        st.caption(
            f"{len(fdf)} variant(s) shown  |  Sales data: {start_date.strftime('%b %d, %Y')} – "
            f"{end_date.strftime('%b %d, %Y')}  |  Lead time: {lead_time} days  |  "
            f"Target coverage: {coverage_days} days"
        )

        reorder_items = fdf[fdf["Reorder Qty"] > 0]
        st.markdown("")
        oc1, oc2 = st.columns(2)
        if oc1.button(
            f"Create Order — {len(reorder_items)} item(s), {int(reorder_items['Reorder Qty'].sum())} units",
            icon=":material/local_shipping:", type="primary", use_container_width=True,
            disabled=reorder_items.empty,
        ):
            prefill = {}
            for _, r in reorder_items.iterrows():
                prefill.setdefault(r["Product"], {})[(r["Color"], r["Size"])] = int(r["Reorder Qty"])
            st.session_state["prefill_reorder"] = prefill
            st.session_state.page = "📥 Purchase Orders"
            st.rerun()

        buf = io.BytesIO()
        df.sort_values("Days Left").to_excel(buf, index=False)
        buf.seek(0)
        oc2.download_button(
            "Download Full Reorder Report",
            icon=":material/download:",
            data=buf,
            file_name=f"reorder_report_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"Could not compute demand & reorder data: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: PURCHASE ORDERS
# ─────────────────────────────────────────────────────────────────────────────

elif page == "📥 Purchase Orders":
    st.subheader("Purchase Orders")
    st.caption(
        "Create a purchase order here — it shows up in Shipment Tracker and "
        "Shipment Details too. New orders stay invisible to Shopify and Demand "
        "& Reorder until you act on them below. Mark as Ordered so Demand & "
        "Reorder counts the quantities as incoming. Add to Shopify's on-hand "
        "inventory whenever you're confident an order is coming — even before "
        "it physically arrives, e.g. to start pre-selling. Add to the Studio's "
        "own inventory separately, once you actually have it on hand. All "
        "three are independent."
    )

    if "shipment_edit_message" in st.session_state:
        st.success(st.session_state.pop("shipment_edit_message"), icon=":material/check_circle:")

    if st.button("Refresh", icon=":material/refresh:"):
        load_shipments.clear()
        load_line_items.clear()
        st.rerun()

    try:
        with st.spinner("Loading purchase orders…"):
            tracker_df, _ = load_shipments()
            qty_by_key, detail_rows, marked_batches = compute_incoming()

        pending_df = tracker_df[~tracker_df["Status"].isin(["Received", "Cancelled"])] if not tracker_df.empty else tracker_df
        total_incoming_units = sum(r["Incoming Qty"] for r in detail_rows)

        s1, s2, s3, s4 = st.columns(4)
        s1.markdown(f'<div class="stat"><p class="num">{len(pending_df)}</p><p class="lbl">Pending Batches</p></div>', unsafe_allow_html=True)
        s2.markdown(f'<div class="stat"><p class="num">{len(marked_batches)}</p><p class="lbl">Marked as Ordered</p></div>', unsafe_allow_html=True)
        s3.markdown(f'<div class="stat"><p class="num">{total_incoming_units:,}</p><p class="lbl">Total Incoming Units</p></div>', unsafe_allow_html=True)
        s4.markdown(f'<div class="stat"><p class="num">{len(detail_rows)}</p><p class="lbl">Variants Incoming</p></div>', unsafe_allow_html=True)
        st.markdown("")

        existing_batch_nums = [
            int(m.group(1)) for m in (
                re.match(r"Batch (\d+)$", str(b).strip(), re.IGNORECASE) for b in tracker_df["Batch #"]
            ) if m
        ] if not tracker_df.empty else []
        next_batch_default = f"Batch {max(existing_batch_nums) + 1}" if existing_batch_nums else "Batch 1"
        po_product_names = sorted({v["product"] for v in load_inventory().values()})

        # Kept in session_state (not popped) so every rerun — including the one
        # triggered by clicking "Create Purchase Order" itself — rebuilds the
        # grids with the same base quantities. A data_editor's returned value
        # falls back to whatever its `data` argument says for any cell the
        # user hasn't explicitly touched, so if this disappeared after the
        # first render, an untouched prefilled cell would silently go back to
        # 0 the moment the create-order rerun re-evaluated the page.
        reorder_prefill = st.session_state.get("prefill_reorder")
        if reorder_prefill and not st.session_state.get("prefill_reorder_applied"):
            st.session_state["new_ship_items"] = list(reorder_prefill.keys())
            st.session_state["prefill_reorder_applied"] = True

        with st.expander("Create Purchase Order", icon=":material/add_circle:", expanded=bool(reorder_prefill)):
            if reorder_prefill:
                st.info(
                    "Pre-filled from your Demand & Reorder list — review the quantities "
                    "below before creating the order.",
                    icon=":material/auto_awesome:",
                )

            c1, c2 = st.columns(2)
            nb_batch = c1.text_input("Batch #", value=next_batch_default, key="new_ship_batch")
            nb_brand = c2.text_input("Brand", value="Sweet Mayhem", key="new_ship_brand")

            nb_items = st.multiselect("Products in this order", options=po_product_names, key="new_ship_items")

            nb_line_items = []
            nb_unit_prices = {}
            inv_for_new = load_inventory()
            fixed_prices = load_product_prices()
            for prod in nb_items:
                prod_variants = [v for v in inv_for_new.values() if v["product"] == prod]
                colors = sorted({v["color"] for v in prod_variants})
                sizes = sorted({v["size"] for v in prod_variants}, key=_size_sort_key)
                st.markdown(f"**{prod}**")
                prod_prefill = (reorder_prefill or {}).get(prod, {})
                if colors and sizes:
                    hdr_cols = st.columns([2] + [1] * len(sizes))
                    hdr_cols[0].caption("")
                    for hc, size in zip(hdr_cols[1:], sizes):
                        hc.caption(size)
                    for color in colors:
                        row_cols = st.columns([2] + [1] * len(sizes))
                        row_cols[0].write(color)
                        for cc, size in zip(row_cols[1:], sizes):
                            cc.number_input(
                                f"{prod} — {color} / {size}",
                                min_value=0, step=1,
                                value=int(prod_prefill.get((color, size), 0)),
                                key=f"new_ship_qty_{prod}__{color}__{size}",
                                label_visibility="collapsed",
                            )
                else:
                    st.caption("No color/size variants found in inventory for this product.")
                unit_price = st.number_input(
                    f"Unit price — {prod} ($)", min_value=0.0, step=0.01, format="%.2f",
                    value=fixed_prices.get(prod, 0.0), key=f"new_ship_price_{prod}",
                    help="Remembered from last time — changing it here updates the fixed price for future shipments too.",
                )
                nb_unit_prices[prod] = unit_price
                for color in colors:
                    for size in sizes:
                        qty = int(st.session_state.get(f"new_ship_qty_{prod}__{color}__{size}", 0) or 0)
                        if qty > 0:
                            nb_line_items.append({
                                "product": prod, "color": color, "size": size,
                                "qty": qty, "unit_price": unit_price,
                            })

            if nb_line_items:
                nb_total_items = sum(e["qty"] for e in nb_line_items)
                nb_price = sum(e["qty"] * e["unit_price"] for e in nb_line_items)
                st.caption(f"Total: **{nb_total_items} items · ${nb_price:,.2f}**")
            else:
                nb_total_items, nb_price = 0, 0.0

            if st.button("Create Purchase Order", type="primary", use_container_width=True, key="new_ship_create"):
                if not nb_batch.strip():
                    st.error("Batch # is required.")
                elif not tracker_df.empty and nb_batch.strip().lower() in tracker_df["Batch #"].str.lower().tolist():
                    st.error(f"'{nb_batch.strip()}' already exists — pick a different Batch #.")
                else:
                    batch_name = nb_batch.strip()
                    add_shipment({
                        "Batch #": batch_name,
                        "Brand": nb_brand.strip(),
                        "Date Paid": datetime.now().date(),
                        "Date Shipped": None,
                        "Shipment Type": "",
                        "Shipping Company": "",
                        "Warehouse Address": "",
                        "Shipping Mark": "",
                        "Tracking #": "",
                        "Date Received": None,
                        "raw_status": "",
                        "Total Items": nb_total_items,
                        "Price": nb_price,
                        "# of Cartons": 0,
                        "Notes": "",
                        "Items Ordered": ", ".join(nb_items),
                        "Img ref.": "",
                        "Shopify Inventory Status": "",
                        "Marked as Ordered": "",
                    })
                    if nb_line_items:
                        save_line_items(batch_name, nb_line_items)
                    priced = {p: v for p, v in nb_unit_prices.items() if v > 0}
                    if priced:
                        set_product_prices(priced)
                    for k in list(st.session_state.keys()):
                        if k.startswith("new_ship_"):
                            del st.session_state[k]
                    st.session_state.pop("prefill_reorder", None)
                    st.session_state.pop("prefill_reorder_applied", None)
                    st.session_state["shipment_edit_message"] = f"{batch_name} created — fill in its shipping details below."
                    st.session_state["shipment_detail_select"] = batch_name
                    st.session_state[f"ship_editing_{batch_name}"] = True
                    st.session_state.page = "🧾 Shipment Details"
                    st.rerun()

        st.markdown("### Receiving")
        if pending_df.empty:
            st.caption("No pending batches to receive.")
        else:
            po_batch = st.selectbox("Batch #", pending_df["Batch #"].tolist(), key="po_receive_batch_select")
            po_row = pending_df[pending_df["Batch #"] == po_batch].iloc[0]
            po_line_items = [it for it in load_line_items() if it["batch"] == po_batch]

            if not po_line_items:
                st.info(f"{po_batch} has no product breakdown yet — this only applies to batches created via \"Create Purchase Order\" above.")
            else:
                po_shopify_done = (po_row["Shopify Inventory Status"] or "").strip() == SHOPIFY_ADDED_MARKER
                po_received_done = po_row["Status"] == "Received"
                po_marked_done = (po_row["Marked as Ordered"] or "").strip() == MARKED_ORDERED_MARKER

                if po_marked_done:
                    st.success(f"{po_batch} is marked as ordered — its quantities count as incoming on Demand & Reorder.", icon=":material/check_circle:")
                else:
                    st.caption(f"{po_batch}: mark it as ordered so Demand & Reorder counts these quantities as incoming, without touching Shopify.")
                if st.button(
                    "Mark as Ordered" if not po_marked_done else "Already marked as ordered",
                    icon=":material/playlist_add_check:", use_container_width=True,
                    type="primary" if not po_marked_done else "secondary",
                    disabled=po_marked_done,
                    key=f"po_mark_ordered_{po_row['row']}",
                ):
                    mark_batch_as_ordered(po_row)
                    st.session_state["shipment_edit_message"] = f"{po_batch} marked as ordered — now counted as incoming on Demand & Reorder."
                    st.rerun()

                st.markdown("")
                rc1, rc2 = st.columns(2)
                with rc1:
                    if po_shopify_done:
                        st.success("Added to Shopify", icon=":material/check_circle:")
                        st.caption("Adding again will add these quantities a second time — only do this if you're sure it's needed.")
                    if st.button(
                        "Add to Shopify Inventory" if not po_shopify_done else "Add to Shopify again",
                        icon=":material/sync:", use_container_width=True,
                        type="secondary" if po_shopify_done else "primary",
                        key=f"po_add_shopify_{po_row['row']}",
                    ):
                        with st.spinner("Adding to Shopify's on-hand inventory…"):
                            synced, unmatched, failed = add_batch_to_shopify(po_row, po_line_items)
                        msg = []
                        if synced:
                            msg.append(f"{synced} item(s) added to Shopify.")
                        if unmatched:
                            msg.append(f"{len(unmatched)} unmatched: " + ", ".join(unmatched))
                        if failed:
                            msg.append(f"{len(failed)} failed: " + ", ".join(failed))
                        st.session_state["shipment_edit_message"] = " ".join(msg) if msg else "Nothing to add."
                        st.rerun()

                with rc2:
                    if po_received_done:
                        st.success("Added to Studio Inventory", icon=":material/check_circle:")
                        st.caption("Adding again will add these quantities a second time — only do this if you're sure it's needed.")
                    if st.button(
                        "Add to Studio Inventory" if not po_received_done else "Add to Studio Inventory again",
                        icon=":material/inventory_2:", use_container_width=True,
                        type="secondary" if po_received_done else "primary",
                        key=f"po_add_studio_{po_row['row']}",
                    ):
                        with st.spinner("Adding to the Studio inventory…"):
                            added, unmatched = receive_batch_to_studio_inventory(po_line_items)
                        update_shipment(int(po_row["row"]), {
                            **_row_to_fields(po_row),
                            "raw_status": "Received",
                            "Date Received": datetime.now().date(),
                        })
                        msg = [f"{added} item(s) added to Studio inventory."]
                        if unmatched:
                            msg.append(f"{len(unmatched)} unmatched: " + ", ".join(unmatched))
                        st.session_state["shipment_edit_message"] = " ".join(msg)
                        st.rerun()

        st.markdown("")
        st.markdown("### Pending Batches")
        if pending_df.empty:
            st.info("No pending purchase orders — every batch is either Received or Cancelled.")
        else:
            render_shipment_table(pending_df)

        st.markdown("")
        st.markdown("### Incoming by Product")
        st.caption("Only includes batches that have been marked as ordered.")
        if not detail_rows:
            st.info("Nothing counted as incoming yet — mark a batch as ordered above.")
        else:
            incoming_df = pd.DataFrame(detail_rows).sort_values(["Product", "Color", "Size"])
            render_incoming_table(incoming_df)

    except Exception as e:
        st.error(f"Could not load purchase orders: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: CANCELLED ORDERS
# ─────────────────────────────────────────────────────────────────────────────

elif page == "🚫 Cancelled Orders":
    st.subheader("Cancelled Orders")
    st.caption("Roadrunner cancellations → Khawla & Sacha's follow-up queue. Shared across every device.")

    is_admin = st.session_state.role == "admin"

    if st.button("Refresh", icon=":material/refresh:"):
        load_cancelled_orders.clear()
        load_co_settings.clear()
        st.session_state.pop("co_local_patches", None)
        st.rerun()

    if "co_upload_msg" in st.session_state:
        st.success(st.session_state.pop("co_upload_msg"), icon=":material/check_circle:")
    if "co_upload_err" in st.session_state:
        st.error(st.session_state.pop("co_upload_err"))

    try:
        with st.spinner("Loading…"):
            co_settings = load_co_settings()
            co_orders = load_cancelled_orders()
    except Exception as e:
        st.error(f"Could not load cancelled orders: {e}")
        st.stop()

    # Edits made THIS session (stage/outcome/notes) are applied on top of the
    # possibly-cached list rather than forcing a fresh Sheet read on every
    # click — see the comment on update_co_order() for why.
    co_patches = st.session_state.get("co_local_patches", {})
    if co_patches:
        for o in co_orders:
            if o["row"] in co_patches:
                o.update(co_patches[o["row"]])

    co_locked_employee = None
    if not is_admin and st.session_state.username:
        for emp in co_settings["employees"]:
            if emp.lower() == st.session_state.username.strip().lower():
                co_locked_employee = emp
                break

    co_visible = [o for o in co_orders if (not co_locked_employee or o["assigned_to"] == co_locked_employee)]

    co_total_usd = sum(o["total_usd"] for o in co_visible)
    co_recovered = [o for o in co_visible if o["outcome"] == "recovered"]
    co_lost = [o for o in co_visible if o["outcome"] == "lost"]
    co_open = [o for o in co_visible if not o["outcome"]]
    co_today = datetime.now().strftime("%Y-%m-%d")
    co_added_today = [o for o in co_visible if o["date_added"] == co_today]
    co_rate = round(len(co_recovered) / len(co_visible) * 100) if co_visible else 0

    ct1, ct2, ct3, ct4, ct5 = st.columns(5)
    ct1.markdown(f'<div class="stat"><p class="num">{len(co_added_today)}</p><p class="lbl">Added today · ${sum(o["total_usd"] for o in co_added_today):,.2f} at risk</p></div>', unsafe_allow_html=True)
    ct2.markdown(f'<div class="stat"><p class="num">{len(co_visible)}</p><p class="lbl">Tracked total · ${co_total_usd:,.2f} cumulative</p></div>', unsafe_allow_html=True)
    ct3.markdown(f'<div class="stat"><p class="num" style="color:var(--rr-red)">{len(co_open)}</p><p class="lbl">Open / in progress</p></div>', unsafe_allow_html=True)
    ct4.markdown(f'<div class="stat"><p class="num" style="color:#227A55">{len(co_recovered)}</p><p class="lbl">Recovered · {co_rate}% save rate</p></div>', unsafe_allow_html=True)
    ct5.markdown(f'<div class="stat"><p class="num" style="color:#B03A3A">{len(co_lost)}</p><p class="lbl">Lost · ${sum(o["total_usd"] for o in co_lost):,.2f} gone</p></div>', unsafe_allow_html=True)

    st.markdown("")

    if is_admin:
        with st.expander("Upload today's Roadrunner export", icon=":material/upload:"):
            st.caption('.xlsx or .csv — the "Orders" export from the dashboard. Already-tracked orders are skipped automatically.')
            co_file = st.file_uploader("Roadrunner export", type=["xlsx", "xls", "csv"], label_visibility="collapsed")
            # The uploaded file stays in the widget's state across the rerun
            # triggered below, so without this guard the same file would be
            # re-parsed and re-added in an infinite loop.
            co_file_id = f"{co_file.name}:{co_file.size}" if co_file is not None else None
            if co_file is not None and co_file_id != st.session_state.get("co_last_uploaded_file_id"):
                try:
                    with st.spinner("Parsing…"):
                        co_records, co_total_rows, co_skipped = parse_roadrunner_export(co_file)
                    co_existing_ids = {o["order_id"] for o in co_orders}
                    co_new_records = [r for r in co_records if r["order_id"] not in co_existing_ids]
                    co_existing_count = len(co_records) - len(co_new_records)
                    if co_new_records:
                        with st.spinner("Assigning and saving…"):
                            co_per_emp = add_cancelled_orders_batch(co_new_records, co_settings)
                        co_bits = ", ".join(f"{n} to {e}" for e, n in co_per_emp.items())
                        co_msg = f"Parsed {co_total_rows} row(s) → {len(co_new_records)} new order(s) added ({co_bits}), {co_existing_count} already tracked"
                    else:
                        co_msg = f"Parsed {co_total_rows} row(s) → 0 new order(s) added, {co_existing_count} already tracked"
                    if co_skipped:
                        co_msg += f", {co_skipped} skipped (not cancelled)"
                    st.session_state["co_last_uploaded_file_id"] = co_file_id
                    st.session_state["co_upload_msg"] = co_msg
                    st.rerun()
                except Exception as e:
                    st.session_state["co_last_uploaded_file_id"] = co_file_id
                    st.session_state["co_upload_err"] = f"Couldn't read that file — {e}"
                    st.rerun()

        with st.expander("Settings", icon=":material/settings:"):
            with st.form("co_settings_form"):
                cs1, cs2 = st.columns(2)
                co_emp1 = cs1.text_input("Employee 1", value=co_settings["employees"][0])
                co_emp1_w = cs1.number_input("Employee 1 workload share", min_value=0.0, step=0.5, value=float(co_settings["weights"][0]))
                co_emp2 = cs2.text_input("Employee 2", value=co_settings["employees"][1])
                co_emp2_w = cs2.number_input("Employee 2 workload share", min_value=0.0, step=0.5, value=float(co_settings["weights"][1]))
                cs3, cs4 = st.columns(2)
                co_coupon_code = cs3.text_input("Reusable coupon code", value=co_settings["coupon_code"])
                co_coupon_pct = cs4.number_input("Discount %", min_value=0.0, max_value=100.0, value=float(co_settings["coupon_percent"]))
                st.caption("New cancellations are split by these shares — changing them only affects orders added after that; use Rebalance below for untouched orders.")
                co_tpl1 = st.text_area("1st WhatsApp message template", value=co_settings["wa_template_1"])
                co_tpl_coupon = st.text_area("Coupon WhatsApp message template", value=co_settings["wa_template_coupon"])
                st.caption("Tokens: {firstname} {name} {employee} {items} {total} {code} {percent}")
                if st.form_submit_button("Save settings", type="primary", use_container_width=True):
                    co_old_names = co_settings["employees"]
                    co_new_names = [co_emp1.strip() or "Employee 1", co_emp2.strip() or "Employee 2"]
                    if co_old_names != co_new_names:
                        co_rename_map = {old: new for old, new in zip(co_old_names, co_new_names) if old != new}
                        if co_rename_map:
                            co_ws = get_co_ws()
                            co_rename_updates = [
                                {"range": gspread.utils.rowcol_to_a1(o["row"], 12), "values": [[co_rename_map[o["assigned_to"]]]]}
                                for o in co_orders if o["assigned_to"] in co_rename_map
                            ]
                            if co_rename_updates:
                                co_ws.batch_update(co_rename_updates)
                    co_new_weights = [co_emp1_w, co_emp2_w]
                    co_rr_state = [0.0, 0.0] if co_new_weights != co_settings["weights"] else co_settings["rr_state"]
                    save_co_settings({
                        "employees": co_new_names, "weights": co_new_weights, "rr_state": co_rr_state,
                        "coupon_code": co_coupon_code, "coupon_percent": co_coupon_pct,
                        "wa_template_1": co_tpl1, "wa_template_coupon": co_tpl_coupon,
                    })
                    load_cancelled_orders.clear()
                    st.success("Settings saved.")
                    st.rerun()

            if st.button("↻ Rebalance untouched orders", use_container_width=True):
                co_untouched = [o for o in co_orders if not o["outcome"] and o["stage"] == "wa1"]
                if not co_untouched:
                    st.info("Nothing untouched to rebalance.")
                else:
                    co_untouched.sort(key=lambda o: o["creation_date"])
                    co_settings["rr_state"] = [0.0, 0.0]
                    co_ws = get_co_ws()
                    co_rb_updates, co_per_emp = [], {}
                    for o in co_untouched:
                        emp = co_next_assignee(co_settings)
                        co_per_emp[emp] = co_per_emp.get(emp, 0) + 1
                        co_rb_updates.append({"range": gspread.utils.rowcol_to_a1(o["row"], 12), "values": [[emp]]})
                    co_ws.batch_update(co_rb_updates)
                    save_co_settings(co_settings)
                    co_bits = ", ".join(f"{n} to {e}" for e, n in co_per_emp.items())
                    st.success(f"Rebalanced {len(co_untouched)} untouched order(s): {co_bits}")
                    st.rerun()

    st.markdown("")

    if not co_locked_employee:
        co_tab_defs = ["All"] + co_settings["employees"]
        co_active_tab = st.radio("Assigned to", co_tab_defs, horizontal=True, label_visibility="collapsed")
    else:
        co_active_tab = "All"

    co_filtered = co_visible
    if not co_locked_employee and co_active_tab != "All":
        co_filtered = [o for o in co_filtered if o["assigned_to"] == co_active_tab]
    co_filtered.sort(key=lambda o: o["creation_date"], reverse=True)

    if co_filtered:
        co_tsv = ["Order ID\tCustomer\tPhone\tAmount USD\tPayment\tCreated\tAddress\tItems\tFollow-up\tCoupon\tNotes"]
        for o in co_filtered:
            co_followup = "Recovered" if o["outcome"] == "recovered" else "Lost" if o["outcome"] == "lost" else CO_STAGE_LABELS[o["stage"]]
            co_tsv.append("\t".join(str(v).replace("\t", " ").replace("\n", " ") for v in [
                o["order_id"], o["customer"], o["phone"], f"{o['total_usd']:.2f}", o["payment_status"],
                (o["creation_date"] or "").split(" ")[0], o["address"], "; ".join(co_parse_items(o["note"])),
                co_followup, co_settings["coupon_code"], o["notes"],
            ]))
        st.download_button(
            "Download this list", "\n".join(co_tsv),
            file_name=f"cancelled_orders_{datetime.now().strftime('%Y-%m-%d')}.tsv",
            mime="text/tab-separated-values", icon=":material/download:",
        )

    if not co_filtered:
        st.info(
            "No cancellations tracked yet — upload today's Roadrunner export above to get started."
            if is_admin else
            "No cancellations tracked yet — ask an admin to upload today's export."
        )
    else:
        for o in co_filtered:
            with st.container(border=True, key=f"co_card_{o['row']}"):
                cc1, cc2, cc3 = st.columns([1.6, 2.6, 1.4])
                with cc1:
                    co_avatar_bg = co_avatar_color(o["assigned_to"], co_settings["employees"])
                    co_initial = (o["assigned_to"] or "?").strip()[:1].upper()
                    st.markdown(
                        f'<div style="display:flex;align-items:center;gap:9px;margin-bottom:2px">'
                        f'<div class="rr-avatar" style="background:{co_avatar_bg}">{co_initial}</div>'
                        f'<div><div style="font-weight:700">#{co_order_number(o["order_id"], o["reference_id"])}</div>'
                        f'<div style="font-size:0.72rem;color:#8b8f9b">{o["assigned_to"] or "—"}</div></div></div>',
                        unsafe_allow_html=True,
                    )
                with cc2:
                    st.markdown(f"**{o['customer'] or '—'}**")
                    if o["phone"]:
                        co_phone_digits = re.sub(r"[^0-9+]", "", o["phone"])
                        st.markdown(f"[:material/call: {o['phone']}](tel:{co_phone_digits})")
                    co_items = co_parse_items(o["note"])
                    if co_items:
                        st.caption(" · ".join(co_items))
                with cc3:
                    st.markdown(f"**${o['total_usd']:,.2f}**")
                    if o["outcome"]:
                        co_pill_cls = "rr-pill-green" if o["outcome"] == "recovered" else "rr-pill-red"
                        co_label = "✓ Recovered" if o["outcome"] == "recovered" else "✕ Lost"
                    else:
                        co_pill_cls = "rr-pill-amber"
                        co_label = CO_STAGE_LABELS[o["stage"]]
                    st.markdown(f'<span class="rr-pill {co_pill_cls}">{co_label}</span>', unsafe_allow_html=True)

                st.markdown("")
                if o["outcome"]:
                    if st.button("Reopen", key=f"co_reopen_{o['row']}"):
                        update_co_order(o["row"], outcome="")
                        st.rerun()
                else:
                    co_stage = o["stage"]
                    ac1, ac2, ac3 = st.columns([2.2, 1.3, 1])
                    with ac1:
                        if o["phone"]:
                            if co_stage == "wa1":
                                co_text = co_fill_template(co_settings["wa_template_1"], o, co_settings)
                                co_wa_label = "💬 Send WhatsApp"
                            elif co_stage == "call":
                                co_text = co_fill_template("Hi {firstname}, this is {employee} 👋", o, co_settings)
                                co_wa_label = "📞 Call on WhatsApp"
                            else:
                                co_text = co_fill_template(co_settings["wa_template_coupon"], o, co_settings)
                                co_wa_label = "💬 Send coupon"
                            st.link_button(co_wa_label, co_wa_link(o["phone"], co_text), use_container_width=True)
                        else:
                            st.caption("No phone")
                    with ac2:
                        if st.button("Recovered", key=f"co_yes_{o['row']}", use_container_width=True):
                            update_co_order(o["row"], outcome="recovered")
                            st.rerun()
                    with ac3:
                        if st.button("No", key=f"co_no_{o['row']}", use_container_width=True):
                            co_idx = CO_STAGE_ORDER.index(co_stage)
                            if co_idx < len(CO_STAGE_ORDER) - 1:
                                update_co_order(o["row"], stage=CO_STAGE_ORDER[co_idx + 1])
                            else:
                                update_co_order(o["row"], outcome="lost")
                            st.rerun()
                co_notes = st.text_input(
                    "Notes", value=o["notes"], key=f"co_notes_{o['row']}",
                    label_visibility="collapsed", placeholder="Call notes…",
                )
                if co_notes != o["notes"]:
                    update_co_order(o["row"], notes=co_notes)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: REFUNDS
# ─────────────────────────────────────────────────────────────────────────────

elif page == "💸 Refunds":
    st.subheader("Refunds")
    st.caption("Customer refund requests → Whish payout tracker. Shared across every device — everyone sees the same list.")

    if st.button("Refresh", icon=":material/refresh:"):
        load_refunds.clear()
        st.rerun()

    try:
        with st.spinner("Loading refunds…"):
            refunds = load_refunds()
    except Exception as e:
        st.error(f"Could not load refunds: {e}")
        st.stop()

    pending = [r for r in refunds if r["status"] == "Pending"]
    refunded = [r for r in refunds if r["status"] == "Refunded"]
    rejected = [r for r in refunds if r["status"] == "Rejected"]
    pending_sum = sum(r["amount"] or 0 for r in pending)
    refunded_sum = sum(r["amount"] or 0 for r in refunded)

    s1, s2, s3 = st.columns(3)
    s1.markdown(f'<div class="stat"><p class="num" style="color:#966600">{len(pending)}</p><p class="lbl">Pending · ${pending_sum:,.2f} owed</p></div>', unsafe_allow_html=True)
    s2.markdown(f'<div class="stat"><p class="num" style="color:#227A55">{len(refunded)}</p><p class="lbl">Refunded · ${refunded_sum:,.2f} paid out</p></div>', unsafe_allow_html=True)
    s3.markdown(f'<div class="stat"><p class="num" style="color:var(--rr-red)">{len(rejected)}</p><p class="lbl">Rejected</p></div>', unsafe_allow_html=True)

    st.markdown("")
    with st.expander("Log a new refund request", icon=":material/add_circle:"):
        with st.form("add_refund_form", clear_on_submit=True):
            c1, c2 = st.columns(2)
            order = c1.text_input("Order # *")
            customer = c2.text_input("Customer name *")
            c3, c4 = st.columns(2)
            amount = c3.number_input("Amount ($) *", min_value=0.0, step=0.01, format="%.2f")
            whish = c4.text_input("Whish number *")
            submitted = st.form_submit_button("Save refund", type="primary", use_container_width=True)
            if submitted:
                if not order.strip() or not customer.strip() or not whish.strip() or amount <= 0:
                    st.error("Order #, Customer name, Amount, and Whish number are all required.")
                else:
                    add_refund(order.strip(), customer.strip(), amount, whish.strip(), st.session_state.username)
                    st.success(f"Refund for {customer.strip()} logged.")
                    st.rerun()

    st.markdown("")
    if not refunds:
        st.info("No refund requests logged yet — use the form above to log one.")
    else:
        fc1, fc2 = st.columns([2, 1])
        search = fc1.text_input("Search", placeholder="Search order # or customer…", label_visibility="collapsed")
        status_filter = fc2.selectbox("Status filter", ["All"] + REFUND_STATUSES, label_visibility="collapsed")

        filtered = refunds
        if status_filter != "All":
            filtered = [r for r in filtered if r["status"] == status_filter]
        if search.strip():
            q = search.strip().lower()
            filtered = [r for r in filtered if q in r["order"].lower() or q in r["customer"].lower()]
        filtered = sorted(filtered, key=lambda r: r["date_logged"], reverse=True)

        if not filtered:
            st.caption("No refunds match this filter.")
        else:
            df = pd.DataFrame([{
                "Order #": r["order"], "Customer": r["customer"],
                "Amount": r["amount"], "Status": r["status"],
                "Whish Number": r["whish"], "Logged By": r["logged_by"],
                "Date": r["date_logged"],
            } for r in filtered])
            st.dataframe(
                style_status(df, column="Status"),
                use_container_width=True, hide_index=True,
                column_config={"Amount": st.column_config.NumberColumn(format="$%.2f")},
            )
            st.caption(f"{len(filtered)} of {len(refunds)} refund(s) shown")

        st.markdown("")
        st.markdown("### Update a Refund")
        with st.container(border=True, key="rf_update_panel"):
            sel_choice = st.selectbox(
                "Refund", refunds, label_visibility="collapsed",
                format_func=lambda r: f"{r['order']} — {r['customer']}",
                key="rf_selected_refund",
            )
            # Streamlit's selectbox can hand back a stale copy of the chosen
            # dict across a rerun (it doesn't reliably re-match a mutated
            # object by value) — re-resolve by row against the just-reloaded
            # list so the panel below always reflects the latest status.
            sel_refund = next((r for r in refunds if r["row"] == sel_choice["row"]), sel_choice)

            rf_pill_cls = {"Pending": "rr-pill-amber", "Refunded": "rr-pill-green", "Rejected": "rr-pill-red"}[sel_refund["status"]]
            st.markdown(f'<span class="rr-pill {rf_pill_cls}">{sel_refund["status"]}</span>', unsafe_allow_html=True)
            st.markdown("")

            d1, d2 = st.columns(2)
            with d1:
                amt_str = f"${sel_refund['amount']:,.2f}" if sel_refund["amount"] is not None else "—"
                st.markdown(f"**Amount:** {amt_str}")
                st.markdown(f"**Whish number:** {sel_refund['whish'] or '—'}")
            with d2:
                st.markdown(f"**Logged by:** {sel_refund['logged_by'] or '—'}")
                st.markdown(f"**Date logged:** {sel_refund['date_logged'] or '—'}")

            new_status = st.selectbox(
                "Status", REFUND_STATUSES,
                index=REFUND_STATUSES.index(sel_refund["status"]),
                key=f"status_select_{sel_refund['row']}",
            )
            if st.button(
                "Update Status", type="primary", use_container_width=True,
                disabled=(new_status == sel_refund["status"]),
            ):
                update_refund_status(sel_refund["row"], new_status)
                st.success(f"Marked {sel_refund['order']} as {new_status}.")
                st.rerun()

            if st.session_state.role == "admin":
                confirm_key = f"confirm_del_refund_{sel_refund['row']}"
                if not st.session_state.get(confirm_key):
                    if st.button("Delete this refund", icon=":material/delete:", use_container_width=True):
                        st.session_state[confirm_key] = True
                        st.rerun()
                else:
                    st.warning(
                        f"Permanently delete the refund for {sel_refund['customer']} "
                        f"({sel_refund['order']})? This can't be undone."
                    )
                    cc1, cc2 = st.columns(2)
                    if cc1.button("Yes, delete", type="primary", use_container_width=True):
                        delete_refund(sel_refund["row"])
                        st.session_state.pop(confirm_key, None)
                        st.success("Refund deleted.")
                        st.rerun()
                    if cc2.button("Cancel", use_container_width=True):
                        st.session_state.pop(confirm_key, None)
                        st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: MANAGE USERS (admin only)
# ─────────────────────────────────────────────────────────────────────────────

elif page == ADMIN_PAGE:
    if st.session_state.role != "admin":
        st.error("Admins only.")
        st.stop()

    st.subheader("Manage Users")
    st.caption("Create accounts and control which pages each person can see. Only admins can reach this page.")

    users = load_users()
    admin_count = sum(1 for u in users.values() if u["role"] == "admin")

    st.markdown("### Existing Users")
    rows = [
        {
            "Username": u["username"],
            "Role": u["role"],
            "Page Access": "All" if u["role"] == "admin" else (", ".join(u["permissions"]) or "None"),
        }
        for u in users.values()
    ]
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.divider()
    st.markdown("### Add New User")
    with st.form("add_user_form", clear_on_submit=True):
        new_username = st.text_input("Username")
        new_password = st.text_input("Password", type="password")
        new_role = st.selectbox("Role", ["user", "admin"])
        new_perms = st.multiselect(
            "Page Access", ALL_PAGES,
            help="Ignored for admins — admins always have access to every page.",
        )
        submitted = st.form_submit_button("Create User", type="primary")
        if submitted:
            if not new_username.strip() or not new_password:
                st.error("Username and password are required.")
            elif new_username.strip().lower() in users:
                st.error("That username already exists.")
            else:
                perms = ALL_PAGES if new_role == "admin" else new_perms
                create_user(new_username.strip(), new_password, new_role, perms)
                st.success(f"User '{new_username.strip()}' created.")
                st.rerun()

    st.divider()
    st.markdown("### Edit or Remove a User")
    if not users:
        st.info("No users yet.")
    else:
        usernames = sorted(u["username"] for u in users.values())
        sel = st.selectbox("Select a user", usernames)
        u = users[sel.lower()]
        is_self = sel.lower() == st.session_state.username.lower()
        is_last_admin = u["role"] == "admin" and admin_count <= 1

        with st.form("edit_user_form"):
            role_options = ["user", "admin"]
            edit_role = st.selectbox(
                "Role", role_options, index=role_options.index(u["role"]) if u["role"] in role_options else 0,
                disabled=is_last_admin,
                help="The last remaining admin can't be demoted." if is_last_admin else None,
            )
            edit_perms = st.multiselect(
                "Page Access", ALL_PAGES, default=[p for p in u["permissions"] if p in ALL_PAGES],
                help="Ignored for admins — admins always have access to every page.",
            )
            new_pw = st.text_input("Reset Password (leave blank to keep current)", type="password")
            save = st.form_submit_button("Save Changes", type="primary")
            if save:
                perms = ALL_PAGES if edit_role == "admin" else edit_perms
                update_user(u["row"], role=edit_role, permissions=perms, password=new_pw or None)
                if is_self:
                    st.session_state.role = edit_role
                    st.session_state.pages = effective_pages({"role": edit_role, "permissions": perms})
                st.success(f"Updated '{u['username']}'.")
                st.rerun()

        if is_last_admin:
            st.caption("Can't delete the last remaining admin.")
        elif is_self:
            st.caption("You can't delete your own account while signed in as it.")
        else:
            if st.button(f"Delete '{u['username']}'", icon=":material/delete:", use_container_width=True):
                delete_user(u["row"])
                st.success(f"Deleted '{u['username']}'.")
                st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: SHIPMENT TRACKER
# ─────────────────────────────────────────────────────────────────────────────

elif page == "🚢 Shipment Tracker":
    st.subheader("Shipment Tracker")
    st.caption(
        "Reads live from the shared Shipments Tracker Google Sheet. To edit a "
        "batch, go to Shipment Details."
    )

    if "shipment_edit_message" in st.session_state:
        st.success(st.session_state.pop("shipment_edit_message"), icon=":material/check_circle:")

    if st.button("Refresh", icon=":material/refresh:"):
        load_shipments.clear()
        load_packaging_tables.clear()
        st.rerun()

    try:
        with st.spinner("Loading shipments…"):
            df, date_issues = load_shipments()

        if date_issues:
            st.warning(
                "Some dates in the Sheet couldn't be read and are showing blank below "
                "— fix them at the source:\n\n" + "\n".join(f"- {i}" for i in date_issues),
                icon=":material/event_busy:",
            )

        if df.empty:
            st.info("No shipments found in the Sheet yet.")
        else:
            total_shipments = len(df)
            received = df["Status"].eq("Received").sum()
            cancelled = df["Status"].eq("Cancelled").sum()
            total_spent = df["Price"].fillna(0).sum()

            s1, s2, s3, s4 = st.columns(4)
            s1.markdown(f'<div class="stat"><p class="num">{total_shipments}</p><p class="lbl">Total Shipments</p></div>', unsafe_allow_html=True)
            s2.markdown(f'<div class="stat"><p class="num">{received}</p><p class="lbl">Received</p></div>', unsafe_allow_html=True)
            s3.markdown(f'<div class="stat"><p class="num">{total_shipments - received - cancelled}</p><p class="lbl">In Transit</p></div>', unsafe_allow_html=True)
            s4.markdown(f'<div class="stat"><p class="num">${total_spent:,.0f}</p><p class="lbl">Total Spent</p></div>', unsafe_allow_html=True)
            st.markdown("")

            fc1, fc2, fc3 = st.columns(3)
            brands = ["All"] + sorted(df["Brand"].replace("", pd.NA).dropna().unique().tolist())
            sel_brand = fc1.selectbox("Brand", brands)
            statuses = ["All"] + sorted(df["Status"].unique().tolist())
            sel_status = fc2.selectbox("Status", statuses)
            types = ["All"] + sorted(df["Shipment Type"].replace("", pd.NA).dropna().unique().tolist())
            sel_type = fc3.selectbox("Shipment Type", types)

            fdf = df.copy()
            if sel_brand != "All":
                fdf = fdf[fdf["Brand"] == sel_brand]
            if sel_status != "All":
                fdf = fdf[fdf["Status"] == sel_status]
            if sel_type != "All":
                fdf = fdf[fdf["Shipment Type"] == sel_type]

            render_shipment_table(fdf)
            st.caption(f"{len(fdf)} of {total_shipments} shipment(s) shown")

            st.divider()
            used_df, orders_df = load_packaging_tables()
            with st.expander("Packaging Usage & Stock Orders", icon=":material/inventory_2:"):
                p1, p2 = st.columns(2)
                with p1:
                    st.markdown("**Packaging Used (per batch)**")
                    st.dataframe(used_df, use_container_width=True, hide_index=True)
                with p2:
                    st.markdown("**Packaging Stock (orders)**")
                    st.dataframe(orders_df, use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Could not load shipment data: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: SHIPMENT DETAILS
# ─────────────────────────────────────────────────────────────────────────────

elif page == "🧾 Shipment Details":
    st.subheader("Shipment Details")
    st.caption(
        "View and edit a batch's product breakdown and order summary — to "
        "send straight to your supplier. To create a new one, go to Purchase "
        "Orders. Also shows shipments uploaded as an Excel file to the Drive "
        "folder."
    )

    if "shipment_edit_message" in st.session_state:
        st.success(st.session_state.pop("shipment_edit_message"), icon=":material/check_circle:")

    if st.button("Refresh", icon=":material/refresh:"):
        list_shipment_detail_files.clear()
        load_line_items.clear()
        load_shipments.clear()
        st.rerun()

    try:
        with st.spinner("Loading shipment list…"):
            files = list_shipment_detail_files()
            all_line_items = load_line_items()
            tracker_df, _ = load_shipments()

        sheet_batches = sorted({it["batch"] for it in all_line_items}, key=_batch_num, reverse=True)
        drive_names = [f["name"] for f in files]
        # Sheet-entered batches take priority over a Drive file for the same batch
        # number — matched by number, not exact name, since a Drive file's name
        # (e.g. "Batch_19_Sep17_2026_SM.xlsx") never matches the sheet's "Batch 19"
        # as a string even when they represent the same shipment (e.g. after
        # importing an old Excel file's breakdown into the sheet).
        sheet_batch_nums = {_batch_num(b) for b in sheet_batches}
        all_names = sheet_batches + [n for n in drive_names if _batch_num(n) not in sheet_batch_nums]

        if not all_names:
            st.info("No shipment details yet — create one from Purchase Orders, or upload an Excel file to the Drive folder.")
        else:
            sel_name = st.selectbox("Shipment", all_names, key="shipment_detail_select")

            tracker_match = tracker_df[tracker_df["Batch #"] == sel_name] if not tracker_df.empty else tracker_df
            has_tracker_row = not tracker_match.empty
            edit_row = tracker_match.iloc[0] if has_tracker_row else None
            editing_key = f"ship_editing_{sel_name}"

            if sel_name in sheet_batches:
                st.caption("Source: entered in-app")
                products, grand_total = line_items_to_products(sel_name, all_line_items)
            else:
                st.caption("Source: uploaded Excel file")
                sel_file = next(f for f in files if f["name"] == sel_name)
                with st.spinner(f"Loading {sel_name}…"):
                    file_bytes = fetch_shipment_detail_bytes(sel_file["id"])
                    products, grand_total = parse_shipment_detail(file_bytes)

            if has_tracker_row:
                hc1, hc2 = st.columns([5, 1])
                with hc1:
                    st.markdown(f"### {html_lib.escape(sel_name)}")
                with hc2:
                    if st.button("Edit", icon=":material/edit:", use_container_width=True):
                        st.session_state[editing_key] = not st.session_state.get(editing_key, False)
                        st.rerun()

                if not st.session_state.get(editing_key):
                    with st.container(border=True):
                        d1, d2 = st.columns(2)
                        d1.text_input("Brand", value=edit_row["Brand"], disabled=True)
                        d2.selectbox(
                            "Status", SHIPMENT_STATUS_OPTIONS,
                            index=SHIPMENT_STATUS_OPTIONS.index(_status_option_default(edit_row["Status"])),
                            disabled=True,
                        )

                        d3, d4, d5 = st.columns(3)
                        d3.date_input("Date Paid", value=edit_row["Date Paid"], disabled=True)
                        d4.date_input("Date Shipped", value=edit_row["Date Shipped"], disabled=True)
                        d5.date_input("Date Received", value=edit_row["Date Received"], disabled=True)

                        d6, d7, d8 = st.columns(3)
                        d6.text_input("Shipment Type", value=edit_row["Shipment Type"], disabled=True)
                        d7.text_input("Shipping Company", value=edit_row["Shipping Company"], disabled=True)
                        d8.text_input("Tracking #", value=edit_row["Tracking #"], disabled=True)

                        d9, d10 = st.columns(2)
                        d9.number_input(
                            "Price ($)", step=0.01, format="%.2f",
                            value=float(edit_row["Price"]) if pd.notna(edit_row["Price"]) else 0.0,
                            disabled=True,
                        )
                        d10.number_input(
                            "# of Cartons", step=1,
                            value=int(edit_row["# of Cartons"]) if pd.notna(edit_row["# of Cartons"]) else 0,
                            disabled=True,
                        )

                        d11, d12 = st.columns(2)
                        d11.text_input("Shipping Mark", value=edit_row["Shipping Mark"], disabled=True)
                        d12.text_input("Img ref.", value=edit_row["Img ref."], disabled=True)

                        st.text_area("Warehouse Address", value=edit_row["Warehouse Address"], height=90, disabled=True)
                        st.text_input("Items Ordered", value=edit_row["Items Ordered"], disabled=True)
                        st.text_area("Notes", value=edit_row["Notes"], height=90, disabled=True)
                else:
                    with st.container(border=True, key="shipment_edit_panel"):
                        st.caption("Items Ordered isn't editable here — it's set from the batch's product breakdown.")

                        # Outside the form so picking a type/company updates the fixed
                        # address preview immediately, instead of only after Save.
                        type_options = _options_with_current(SHIPMENT_TYPE_OPTIONS, edit_row["Shipment Type"])
                        co_options = _options_with_current(SHIPPING_COMPANY_OPTIONS, edit_row["Shipping Company"])

                        ac1, ac2 = st.columns(2)
                        sel_ship_type = ac1.selectbox(
                            "Shipment Type", type_options,
                            index=type_options.index((edit_row["Shipment Type"] or "").strip() or SHIPMENT_TYPE_OPTIONS[0]),
                            key=f"ship_type_select_{edit_row['row']}",
                        )
                        sel_ship_co = ac2.selectbox(
                            "Shipping Company", co_options,
                            index=co_options.index((edit_row["Shipping Company"] or "").strip() or SHIPPING_COMPANY_OPTIONS[0]),
                            key=f"ship_co_select_{edit_row['row']}",
                        )
                        sel_warehouse = WAREHOUSE_ADDRESS_BY_COMBO.get((sel_ship_type, sel_ship_co), "")
                        st.text_area(
                            "Warehouse Address",
                            value=sel_warehouse or "No fixed address on file for this Shipment Type + Shipping Company yet.",
                            height=90, disabled=True,
                        )

                        with st.form(f"edit_shipment_form_{edit_row['row']}"):
                            c1, c2 = st.columns(2)
                            f_brand = c1.text_input("Brand", value=edit_row["Brand"])
                            f_status = c2.selectbox(
                                "Status", SHIPMENT_STATUS_OPTIONS,
                                index=SHIPMENT_STATUS_OPTIONS.index(_status_option_default(edit_row["Status"])),
                            )

                            c3, c4, c5 = st.columns(3)
                            f_date_paid = c3.date_input("Date Paid", value=edit_row["Date Paid"])
                            f_date_shipped = c4.date_input("Date Shipped", value=edit_row["Date Shipped"])
                            f_date_received = c5.date_input("Date Received", value=edit_row["Date Received"])

                            f_tracking = st.text_input("Tracking #", value=edit_row["Tracking #"])

                            c9, c10 = st.columns(2)
                            f_price = c9.number_input(
                                "Price ($)", min_value=0.0, step=0.01, format="%.2f",
                                value=float(edit_row["Price"]) if pd.notna(edit_row["Price"]) else 0.0,
                            )
                            f_cartons = c10.number_input(
                                "# of Cartons", min_value=0, step=1,
                                value=int(edit_row["# of Cartons"]) if pd.notna(edit_row["# of Cartons"]) else 0,
                            )

                            c11, c12 = st.columns(2)
                            f_ship_mark = c11.text_input("Shipping Mark", value=edit_row["Shipping Mark"])
                            f_img_ref = c12.text_input("Img ref.", value=edit_row["Img ref."])

                            f_notes = st.text_area("Notes", value=edit_row["Notes"], height=90)

                            fc1, fc2 = st.columns(2)
                            if fc1.form_submit_button("Save Changes", type="primary", use_container_width=True):
                                update_shipment(int(edit_row["row"]), {
                                    "Batch #": edit_row["Batch #"],
                                    "Brand": f_brand.strip(),
                                    "Date Paid": f_date_paid,
                                    "Date Shipped": f_date_shipped,
                                    "Shipment Type": sel_ship_type,
                                    "Shipping Company": sel_ship_co,
                                    "Warehouse Address": sel_warehouse,
                                    "Shipping Mark": f_ship_mark.strip(),
                                    "Tracking #": f_tracking.strip(),
                                    "Date Received": f_date_received,
                                    "raw_status": _status_option_to_raw(f_status),
                                    "Total Items": int(edit_row["Total Items"]) if pd.notna(edit_row["Total Items"]) else 0,
                                    "Price": f_price,
                                    "# of Cartons": f_cartons,
                                    "Notes": f_notes.strip(),
                                    "Items Ordered": edit_row["Items Ordered"],
                                    "Img ref.": f_img_ref.strip(),
                                    "Shopify Inventory Status": edit_row["Shopify Inventory Status"],
                                    "Marked as Ordered": edit_row["Marked as Ordered"],
                                })
                                st.session_state.pop(editing_key, None)
                                st.session_state["shipment_edit_message"] = f"Batch {edit_row['Batch #']} updated."
                                st.rerun()
                            if fc2.form_submit_button("Cancel", use_container_width=True):
                                st.session_state.pop(editing_key, None)
                                st.rerun()

                st.divider()

            if not products:
                st.warning("Couldn't find any recognizable product blocks in this shipment.")
            else:
                for p in products:
                    st.markdown(f"#### {p['name']}")
                    st.dataframe(product_grid(p), use_container_width=True)
                    price_str = f"${p['unit_price']:,.2f}" if p["unit_price"] is not None else "—"
                    subtotal_str = f"${p['subtotal']:,.2f}" if p["subtotal"] is not None else "—"
                    st.caption(
                        f"Total ordered: {p['total_qty']}  |  Unit price: {price_str}  |  "
                        f"Subtotal: {subtotal_str}"
                    )
                    st.markdown("")

                if grand_total is not None:
                    st.markdown(f"### Grand Total: ${grand_total:,.2f}")

                st.divider()
                st.markdown("##### Order Summary")
                st.caption("Copy this, or download it as an Excel file, to send straight to your supplier.")
                st.code(format_order_summary(sel_name, products, grand_total), language=None)
                st.download_button(
                    "Download as Excel",
                    data=build_order_excel(sel_name, products, grand_total),
                    file_name=f"{sel_name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    icon=":material/download:",
                )

                if st.session_state.role == "admin":
                    st.divider()
                    confirm_key = f"confirm_del_shipment_{sel_name}"
                    if not st.session_state.get(confirm_key):
                        if st.button("Delete this batch", icon=":material/delete:"):
                            st.session_state[confirm_key] = True
                            st.rerun()
                    else:
                        st.warning(
                            f"Permanently delete {sel_name}? This removes it from Shipment "
                            "Tracker and its product breakdown here. This can't be undone."
                        )
                        cc1, cc2 = st.columns(2)
                        if cc1.button("Yes, delete", type="primary", use_container_width=True):
                            if sel_name in sheet_batches:
                                tracker_match = tracker_df[tracker_df["Batch #"] == sel_name]
                                del_row = int(tracker_match.iloc[0]["row"]) if not tracker_match.empty else None
                                delete_shipment_batch(del_row, sel_name)
                            else:
                                delete_shipment_detail_file(sel_file["id"])
                            st.session_state.pop(confirm_key, None)
                            st.session_state.pop("shipment_detail_select", None)
                            st.session_state["shipment_edit_message"] = f"{sel_name} deleted."
                            st.rerun()
                        if cc2.button("Cancel", use_container_width=True):
                            st.session_state.pop(confirm_key, None)
                            st.rerun()

    except Exception as e:
        st.error(f"Could not load shipment details: {e}")
