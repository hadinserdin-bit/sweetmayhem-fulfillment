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
from pathlib import Path
import io
import json
import math
import re
import bcrypt
import openpyxl
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

def get_ws():
    return _gc().open_by_key(SHEET_ID).get_worksheet(0)

# ─── Users & Access Control ────────────────────────────────────────────────────

USERS_SHEET_NAME = "Users"
ALL_PAGES = [
    "📦 Fulfillment", "🔄 Restock", "➕ Add Product", "📋 View Inventory", "📊 Demand & Reorder",
    "🚫 Cancelled Orders", "🚢 Shipment Tracker", "🧾 Shipment Details",
]
ADMIN_PAGE = "👤 Manage Users"
CANCELLED_ORDERS_URL = "https://claude.ai/artifact/1TJ5d5iJuijaHKNTTBSiyZ"

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
    "🚫 Cancelled Orders": ":material/cancel:",
    "🚢 Shipment Tracker": ":material/directions_boat:",
    "🧾 Shipment Details": ":material/receipt_long:",
    ADMIN_PAGE: ":material/group:",
}

def page_label(p):
    return p.split(" ", 1)[1] if " " in p else p

def get_users_ws():
    sh = _gc().open_by_key(SHEET_ID)
    try:
        return sh.worksheet(USERS_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=USERS_SHEET_NAME, rows=100, cols=5)
        ws.append_row(["Username", "PasswordHash", "Role", "Permissions", "CreatedAt"])
        return ws

@st.cache_data(ttl=60)
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

def create_user(username, password, role, permissions):
    ws = get_users_ws()
    ws.append_row([
        username, hash_password(password), role, ",".join(permissions),
        datetime.now().strftime("%Y-%m-%d %H:%M"),
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
    load_users.clear()

def delete_user(row):
    get_users_ws().delete_rows(row)
    load_users.clear()

def effective_pages(user):
    if user["role"] == "admin":
        return ALL_PAGES + [ADMIN_PAGE]
    return [p for p in user["permissions"] if p in ALL_PAGES]

# ─── Auth ─────────────────────────────────────────────────────────────────────

def login_screen():
    if st.session_state.get("authenticated"):
        return True

    st.markdown("""
    <style>
    [data-testid="stAppViewContainer"] { background: #f4f5f7; }
    </style>
    """, unsafe_allow_html=True)

    users = load_users()

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
            username = st.text_input("Username")
            pwd = st.text_input("Password", type="password", placeholder="Enter password")
            if st.button("Sign In", type="primary", use_container_width=True):
                u = users.get(username.strip().lower())
                if u and verify_password(pwd, u["password_hash"]):
                    st.session_state.authenticated = True
                    st.session_state.username = u["username"]
                    st.session_state.role = u["role"]
                    st.session_state.pages = effective_pages(u)
                    st.rerun()
                else:
                    st.error("Incorrect username or password.")
    return False

if not login_screen():
    st.stop()

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
    return inv, ws

def batch_update_qty(ws, updates):
    ws.batch_update([
        {"range": f"D{row}", "values": [[qty]]}
        for row, qty in updates
    ])

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
    "Out of Stock": "#fdecea", "Cancelled": "#fdecea",
    "Reorder Now": "#fdecea",
    "Reorder Soon": "#fff6e0",
    "In Transit": "#eef2f7",
    "OK": "#eaf7ed", "Received": "#eaf7ed",
}

def style_status(df, column="Status"):
    def _color(val):
        base = str(val).split(" (")[0].strip()
        bg = STATUS_COLORS.get(base, "")
        return f"background-color: {bg}" if bg else ""
    return df.style.map(_color, subset=[column])

# ─── Demand & Reorder ─────────────────────────────────────────────────────────

SNAPSHOT_SHEET_NAME = "InventorySnapshots"
MIN_TRACKED_DAYS = 5  # minimum days of stock-history before trusting the adjusted rate

def get_snapshot_ws():
    sh = _gc().open_by_key(SHEET_ID)
    try:
        return sh.worksheet(SNAPSHOT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SNAPSHOT_SHEET_NAME, rows=1000, cols=5)
        ws.append_row(["Date", "Product", "Color", "Size", "Qty"])
        return ws

def record_snapshot_if_needed(inv):
    """Log today's stock level per variant, once per day."""
    today = datetime.now().strftime("%Y-%m-%d")
    ws = get_snapshot_ws()
    existing_dates = set(ws.col_values(1))
    if today in existing_dates:
        return False
    rows = [[today, v["product"], v["color"], v["size"], v["qty"]] for v in inv.values()]
    if rows:
        ws.append_rows(rows)
    return True

@st.cache_data(ttl=600)
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

def build_reorder_table(inv, sold, stock_days, start_date, end_date, lead_time, coverage_days):
    window_days = max(1, (end_date - start_date).days + 1)
    rows = []
    for key, v in inv.items():
        total_sold = sold.get(key, 0)
        sd = stock_days.get(key, {"tracked": set(), "in_stock": set()})
        days_tracked, days_in_stock = len(sd["tracked"]), len(sd["in_stock"])
        days_oos = days_tracked - days_in_stock
        cur_qty = max(0, v["qty"])

        adjusted = days_tracked >= MIN_TRACKED_DAYS and days_in_stock > 0
        if adjusted:
            daily_demand = total_sold / days_in_stock
        else:
            daily_demand = total_sold / window_days

        days_left = (cur_qty / daily_demand) if daily_demand > 0 else float("inf")
        reorder_qty = max(0, math.ceil(daily_demand * coverage_days) - cur_qty)

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
            "Reorder Qty": reorder_qty,
            "Status": status,
            "Confidence": "Adjusted" if adjusted else "Raw (building history)",
        })
    return pd.DataFrame(rows)

# ─── Shipment Tracker ───────────────────────────────────────────────────────────

SHIPMENT_TRACKER_SHEET_ID = "1xwLzbuUU_xbE7aetCI5CxSpkpEdRN2AwzsuSkpqeS7g"
SHIPMENT_TRACKER_TAB = "Shipments Tracker"

def get_shipment_tracker_ws():
    return _gc().open_by_key(SHIPMENT_TRACKER_SHEET_ID).worksheet(SHIPMENT_TRACKER_TAB)

def _parse_sheet_date(v):
    """Handles the sheet's mixed date storage: real date cells (serial numbers) and
    manually-typed text dates like '27/3/2026'."""
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        try:
            return date(1899, 12, 30) + timedelta(days=int(v))
        except (ValueError, OverflowError):
            return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(v).strip(), fmt).date()
        except ValueError:
            continue
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

@st.cache_data(ttl=120)
def load_shipments():
    ws = get_shipment_tracker_ws()
    values = ws.get("A2:R500", value_render_option="UNFORMATTED_VALUE")
    if not values:
        return pd.DataFrame()
    rows = []
    for row in values[1:]:
        row = row + [None] * (18 - len(row))
        batch = row[0]
        if not batch:
            continue
        brand, status = row[1] or "", row[10] or ""
        rows.append({
            "Batch #": batch,
            "Brand": brand,
            "Date Paid": _parse_sheet_date(row[2]),
            "Date Shipped": _parse_sheet_date(row[3]),
            "Shipment Type": row[4] or "",
            "Shipping Company": row[5] or "",
            "Warehouse Address": row[6] or "",
            "Shipping Mark": row[7] or "",
            "Tracking #": str(row[8]) if row[8] not in (None, "") else "",
            "Date Received": _parse_sheet_date(row[9]),
            "Status": _status_display(brand, status),
            "Total Items": _as_number(row[11]),
            "Price": _as_number(row[12]),
            "# of Cartons": _as_number(row[13]),
            "Notes": row[14] or "",
            "Items Ordered": row[15] or "",
            "Img ref.": row[16] or "",
            "Shopify Inventory Status": row[17] or "",
        })
    return pd.DataFrame(rows)

@st.cache_data(ttl=120)
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

# ─── Shipment Details ──────────────────────────────────────────────────────────

SHIPMENT_DETAILS_FOLDER_ID = "1fzGaxnG9fu0dgsUgWIGQYXtKgw0u0_Ex"

def _batch_num(name):
    m = re.search(r"batch[_\s]*(\d+)", name, re.IGNORECASE)
    return int(m.group(1)) if m else -1

@st.cache_data(ttl=300)
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
.stButton > button:hover,
.stButton > button:active,
.stButton > button[kind="primary"] {
    background: var(--rr-red) !important;
    border-color: var(--rr-red) !important;
}
.stButton > button:hover *,
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

    if st.session_state.get("page") not in my_pages:
        st.session_state.page = my_pages[0]

    NAV_SECTIONS = {"🚢 Shipment Tracker": "SHIPMENTS", "🧾 Shipment Details": "SHIPMENTS"}
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
        components.html(
            """
            <script>
            try {
                const el = window.parent.document.querySelector('[data-testid="stSidebarCollapseButton"] button')
                        || window.parent.document.querySelector('[data-testid="stSidebarCollapseButton"]');
                if (el) el.click();
            } catch (e) {}
            </script>
            """,
            height=0,
        )

    st.divider()
    st.caption(f"Signed in as **{st.session_state.username}**  ·  {st.session_state.role}")
    if st.button("Sign Out", use_container_width=True):
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

    c1, c2 = st.columns(2)
    fetch_btn   = c1.button("Fetch All Unfulfilled", icon=":material/refresh:", use_container_width=True)
    preview_btn = c2.button("Preview  (no changes)", icon=":material/visibility:", use_container_width=True, disabled=not st.session_state.preview_done and st.session_state.fulfillable is None)

    sc1, sc2 = st.columns([3, 1])
    order_input = sc1.text_input("", placeholder="Order number e.g. 17234", label_visibility="collapsed")
    single_btn  = sc2.button("Fetch Order", icon=":material/search:", use_container_width=True)

    st.divider()

    # ── Fetch orders from Shopify ──────────────────────────────────────────────

    if fetch_btn:
        with st.spinner("Fetching unfulfilled orders from Shopify…"):
            try:
                inv, ws  = load_inventory()
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
                inv, ws  = load_inventory()
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
    st.caption("Edit the 'Add Qty' column for any item, then click Apply Restock.")

    try:
        with st.spinner("Loading inventory…"):
            inv, ws = load_inventory()

        items = list(inv.items())
        df = pd.DataFrame([
            {
                "Product": v["product"],
                "Color":   v["color"],
                "Size":    v["size"],
                "Current Qty": v["qty"],
                "Add Qty": 0,
            }
            for _, v in items
        ])

        edited = st.data_editor(
            df,
            column_config={
                "Product":     st.column_config.TextColumn(disabled=True),
                "Color":       st.column_config.TextColumn(disabled=True),
                "Size":        st.column_config.TextColumn(disabled=True),
                "Current Qty": st.column_config.NumberColumn(disabled=True),
                "Add Qty":     st.column_config.NumberColumn(min_value=0, step=1),
            },
            hide_index=True,
            use_container_width=True,
        )

        changed = edited[edited["Add Qty"] > 0]
        if not changed.empty:
            st.info(f"{len(changed)} item(s) with quantities to add.")

        if st.button("Apply Restock", type="primary", use_container_width=True):
            if changed.empty:
                st.warning("No quantities entered. Edit the 'Add Qty' column first.")
            else:
                with st.spinner("Updating Google Sheets…"):
                    try:
                        updates = []
                        for idx in changed.index:
                            key, item = items[idx]
                            new_qty = item["qty"] + int(changed.loc[idx, "Add Qty"])
                            updates.append((item["row"], new_qty))
                        batch_update_qty(ws, updates)
                        st.success(f"{len(updates)} item(s) restocked!", icon=":material/check_circle:")
                        st.balloons()
                    except Exception as e:
                        st.error(str(e))
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
                existing = ws.get_all_values()
                existing_keys = set(
                    (r[0].strip().lower(), r[1].strip().lower(), r[2].strip().lower())
                    for r in existing[1:] if len(r) >= 3 and r[0]
                )
                to_add = [v for v in variants
                          if (v[0].lower(), v[1].lower(), v[2].lower()) not in existing_keys]
                skipped_count = len(variants) - len(to_add)
                if to_add:
                    ws.append_rows([[p, c, s, q] for p, c, s, q in to_add])
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
            inv, _ = load_inventory()

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
        "Sales velocity is adjusted for the days each item was actually out of stock, "
        "so a stock-out doesn't make demand look lower than it really is. Stock levels "
        "are logged automatically each time you open this page — accuracy improves the "
        "more often it's checked."
    )

    today = datetime.now().date()
    d1, d2, d3, d4 = st.columns([1, 1, 1, 1.2])
    start_date = d1.date_input(
        "Sales data from", value=today - timedelta(days=90),
        max_value=today,
        help="Set this to a product's launch date to exclude the period before it existed.",
    )
    end_date = d2.date_input("Sales data to", value=today, max_value=today)
    lead_time = d3.slider("Lead time (days)", min_value=5, max_value=21, value=9,
                           help="Sweet Mayhem's supplier lead time is ~7–10 days.")
    coverage_days = d4.number_input("Target stock coverage (days)", min_value=5, max_value=90, value=25, step=1,
                                     help="Reorder quantity tops stock up to cover this many days of demand.")

    if start_date > end_date:
        st.error("'Sales data from' must be on or before 'Sales data to'.")
        st.stop()

    try:
        with st.spinner("Loading inventory & recording today's stock snapshot…"):
            inv, _ = load_inventory()
            recorded_today = record_snapshot_if_needed(inv)
            if recorded_today:
                load_snapshots.clear()

        with st.spinner("Fetching sales history from Shopify…"):
            orders = fetch_shopify_sales(start_date=start_date, end_date=end_date)
            sold, _unmatched = aggregate_sales(orders, inv)

        records = load_snapshots()
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

        df = build_reorder_table(inv, sold, stock_days, start_date, end_date, lead_time, coverage_days)

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

        st.dataframe(
            style_status(fdf),
            use_container_width=True, hide_index=True,
            column_config={"Days Left": st.column_config.NumberColumn(format="%.1f")},
        )
        st.caption(
            f"{len(fdf)} variant(s) shown  |  Sales data: {start_date.strftime('%b %d, %Y')} – "
            f"{end_date.strftime('%b %d, %Y')}  |  Lead time: {lead_time} days  |  "
            f"Target coverage: {coverage_days} days"
        )

        buf = io.BytesIO()
        df.sort_values("Days Left").to_excel(buf, index=False)
        buf.seek(0)
        st.download_button(
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
# PAGE: CANCELLED ORDERS
# ─────────────────────────────────────────────────────────────────────────────

elif page == "🚫 Cancelled Orders":
    cancelled_orders_html = Path(__file__).parent / "cancelled_orders.html"
    if cancelled_orders_html.exists():
        html = cancelled_orders_html.read_text(encoding="utf-8")
        # Tells the embedded tool who's looking at it, so it can hide Upload/Settings
        # for non-admins — see the APP_ROLE check near the top of its own script.
        role_script = f"<script>window.APP_ROLE = {json.dumps(st.session_state.role)};</script>"
        html = html.replace("<body>", "<body>" + role_script, 1)
        components.html(html, height=1400, scrolling=True)
    else:
        st.error("cancelled_orders.html wasn't found next to app.py — the embed can't load.")
        st.link_button("Open Cancelled Orders ↗", CANCELLED_ORDERS_URL, use_container_width=True)

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
        "Reads live from the shared Shipments Tracker Google Sheet. Edit shipment data "
        "there — this page always shows what's currently in the Sheet."
    )

    if st.button("Refresh", icon=":material/refresh:"):
        load_shipments.clear()
        load_packaging_tables.clear()
        st.rerun()

    try:
        with st.spinner("Loading shipments…"):
            df = load_shipments()

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

            display_cols = [
                "Batch #", "Brand", "Date Paid", "Date Shipped", "Shipment Type",
                "Shipping Company", "Tracking #", "Date Received", "Status",
                "Total Items", "Price", "# of Cartons", "Items Ordered", "Shopify Inventory Status",
            ]
            st.dataframe(
                style_status(fdf[display_cols]),
                use_container_width=True, hide_index=True,
                column_config={
                    "Date Paid": st.column_config.DateColumn(format="MMM D, YYYY"),
                    "Date Shipped": st.column_config.DateColumn(format="MMM D, YYYY"),
                    "Date Received": st.column_config.DateColumn(format="MMM D, YYYY"),
                    "Price": st.column_config.NumberColumn(format="$%.0f"),
                },
            )
            st.caption(f"{len(fdf)} of {total_shipments} shipment(s) shown")

            st.divider()
            st.markdown("### Shipment Detail Lookup")
            sel_batch = st.selectbox("Batch #", df["Batch #"].tolist())
            row = df[df["Batch #"] == sel_batch].iloc[0]
            d1, d2 = st.columns(2)
            with d1:
                st.markdown(f"**Shipping Mark:** {row['Shipping Mark'] or '—'}")
                st.markdown(f"**Tracking #:** {row['Tracking #'] or '—'}")
                st.markdown(f"**Notes:** {row['Notes'] or '—'}")
                st.markdown(f"**Img ref.:** {row['Img ref.'] or '—'}")
            with d2:
                st.markdown("**Warehouse Address:**")
                st.text_area(
                    "Warehouse Address", value=row["Warehouse Address"] or "—",
                    height=140, disabled=True, label_visibility="collapsed",
                )

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
        "Reads directly from the Shipment Details folder in Google Drive — pick a "
        "shipment to see exactly what was ordered, by product, color, and size."
    )

    if st.button("Refresh file list", icon=":material/refresh:"):
        list_shipment_detail_files.clear()
        st.rerun()

    try:
        with st.spinner("Loading file list from Drive…"):
            files = list_shipment_detail_files()

        if not files:
            st.info("No shipment detail files found in the folder.")
        else:
            names = [f["name"] for f in files]
            sel_name = st.selectbox("Shipment file", names)
            sel_file = next(f for f in files if f["name"] == sel_name)

            with st.spinner(f"Loading {sel_name}…"):
                file_bytes = fetch_shipment_detail_bytes(sel_file["id"])
                products, grand_total = parse_shipment_detail(file_bytes)

            if not products:
                st.warning("Couldn't find any recognizable product blocks in this file.")
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

    except Exception as e:
        st.error(f"Could not load shipment details: {e}")
