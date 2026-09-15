"""
Sweet Mayhem — Order Fulfillment Web App
"""

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from difflib import SequenceMatcher
from datetime import datetime
from copy import deepcopy
import io

# ─── Page Config ─────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Sweet Mayhem — Fulfillment",
    page_icon="🌸",
    layout="wide",
)

# ─── Constants ────────────────────────────────────────────────────────────────

SHEET_ID = "1t_L1qR3ikD-jjiA2P1tQETQ60hvR1nocX5AimPdSJ7o"
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]

# ─── Google Sheets ────────────────────────────────────────────────────────────

@st.cache_resource
def _gc():
    creds_dict = dict(st.secrets["gcp_service_account"])
    creds_dict["private_key"] = creds_dict["private_key"].replace("\\n", "\n")
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_ws():
    return _gc().open_by_key(SHEET_ID).get_worksheet(0)

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
            orders[name] = {
                "name": name,
                "email": str(row.get("Email", "")).strip(),
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
                "Item": item["name"],
                "Qty": item["quantity"],
                "Action": "Fulfill in Shopify",
            })
    buf = io.BytesIO()
    pd.DataFrame(rows).to_excel(buf, index=False)
    buf.seek(0)
    return buf

# ─── Session State ────────────────────────────────────────────────────────────

_defaults = dict(
    fulfillable=None, skipped=None, orig_inv=None, new_inv=None,
    ws=None, preview_done=False, removed=set(),
    fulfilled=False, report_buf=None, report_name=None,
)
for k, v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ─── CSS ──────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
[data-testid="stSidebar"] { background: #fdf0f5; }
.brand-header {
    background: linear-gradient(135deg, #d63384 0%, #a8265e 100%);
    border-radius: 12px; padding: 1.2rem 1.8rem; margin-bottom: 1.5rem;
}
.brand-header h1 { color: white; margin: 0; font-size: 1.7rem; }
.brand-header p  { color: #f5c6db; margin: 0.2rem 0 0; font-size: 0.9rem; }
.stat {
    background: white; border: 1px solid #dee2e6;
    border-radius: 10px; padding: 1rem 1.4rem; text-align: center;
}
.stat .num { font-size: 2rem; font-weight: 700; margin: 0; }
.stat .lbl { font-size: 0.72rem; color: #6c757d; margin: 0;
             text-transform: uppercase; letter-spacing: .05em; }
</style>
""", unsafe_allow_html=True)

# ─── Sidebar ──────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("### 🌸 Sweet Mayhem")
    page = st.radio(
        "Navigate",
        ["📦 Fulfillment", "🔄 Restock", "➕ Add Product", "📋 View Inventory"],
        label_visibility="collapsed",
    )

# ─── Header ───────────────────────────────────────────────────────────────────

st.markdown("""
<div class="brand-header">
  <h1>🌸 Sweet Mayhem</h1>
  <p>Order Fulfillment Dashboard</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE: FULFILLMENT
# ─────────────────────────────────────────────────────────────────────────────

if page == "📦 Fulfillment":

    uploaded = st.file_uploader(
        "Upload your Shopify orders export (.xlsx or .csv)",
        type=["xlsx", "csv"],
    )

    c1, c2 = st.columns(2)
    preview_btn = c1.button("🔍 Preview  (no changes)", use_container_width=True)
    fulfill_btn = c2.button("✅ Fulfill Orders", use_container_width=True, type="primary")

    st.divider()

    # ── Run logic ─────────────────────────────────────────────────────────────

    if preview_btn or fulfill_btn:
        if not uploaded:
            st.error("Please upload an orders file first.")
            st.stop()
        with st.spinner("Loading inventory from Google Sheets…"):
            try:
                inv, ws = load_inventory()
                orders = load_orders(uploaded)
                if not orders:
                    st.warning("No unfulfilled orders found in this file.")
                    st.stop()
                fulfillable, skipped, new_inv = determine_fulfillable(orders, inv)
                st.session_state.update(
                    fulfillable=fulfillable, skipped=skipped,
                    orig_inv=inv, new_inv=new_inv, ws=ws,
                    preview_done=True, removed=set(),
                    fulfilled=False, report_buf=None,
                )
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

        # Stat cards
        sc1, sc2, sc3 = st.columns(3)
        sc1.markdown(f'<div class="stat"><p class="num" style="color:#198754">{len(fulfillable)}</p><p class="lbl">Ready to fulfill</p></div>', unsafe_allow_html=True)
        sc2.markdown(f'<div class="stat"><p class="num" style="color:#e07b00">{len(skipped)}</p><p class="lbl">Skipped</p></div>', unsafe_allow_html=True)
        sc3.markdown(f'<div class="stat"><p class="num" style="color:#0d6efd">{len(changes)}</p><p class="lbl">Inventory changes</p></div>', unsafe_allow_html=True)
        st.markdown("")

        tab1, tab2, tab3 = st.tabs(["✅ To Fulfill", "⚠️ Skipped", "📊 Inventory Changes"])

        # Tab 1 — To Fulfill
        with tab1:
            if not fulfillable:
                st.info("No fulfillable orders (or all removed).")
            else:
                st.caption("Click ✕ to remove an order from this run before fulfilling.")
                hc = st.columns([2, 2, 5, 1])
                hc[0].markdown("**Order #**")
                hc[1].markdown("**Date**")
                hc[2].markdown("**Items**")
                hc[3].markdown("**Remove**")

                for order in fulfillable:
                    date = order["created_at"][:10] if len(order["created_at"]) >= 10 else order["created_at"]
                    items_str = "  ·  ".join(
                        f"{i['name']} ×{i['quantity']}" for i in order["line_items"]
                    )
                    rc = st.columns([2, 2, 5, 1])
                    rc[0].markdown(f"`{order['name']}`")
                    rc[1].markdown(date)
                    rc[2].markdown(items_str)
                    if rc[3].button("✕", key=f"rm_{order['name']}"):
                        st.session_state.removed.add(order["name"])
                        st.rerun()

                st.divider()
                confirm = st.checkbox(
                    f"I confirm I want to fulfill {len(fulfillable)} order(s) and update inventory"
                )
                if confirm:
                    if st.button("✅ Fulfill These Orders", type="primary", use_container_width=True):
                        with st.spinner("Updating Google Sheets inventory…"):
                            try:
                                updates = [(orig_inv[k]["row"], cur_inv[k]["qty"]) for k in changes]
                                batch_update_qty(st.session_state.ws, updates)
                                buf = make_report(fulfillable)
                                date_str = datetime.now().strftime("%Y-%m-%d")
                                st.session_state.update(
                                    fulfilled=True, preview_done=False,
                                    report_buf=buf,
                                    report_name=f"to_fulfill_{date_str}.xlsx",
                                )
                                st.rerun()
                            except Exception as e:
                                st.error(str(e))

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
        st.success("✅ Orders fulfilled! Inventory updated in Google Sheets.")
        st.download_button(
            "📥 Download Fulfillment Report",
            data=st.session_state.report_buf,
            file_name=st.session_state.report_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        if st.button("🔄 Start New Run"):
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
                        st.success(f"✅ {len(updates)} item(s) restocked!")
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
                msg = f"✅ {len(to_add)} variant(s) added to inventory."
                if skipped_count:
                    msg += f" ({skipped_count} skipped — already existed.)"
                st.success(msg)
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
