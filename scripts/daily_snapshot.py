#!/usr/bin/env python3
"""Daily inventory snapshot logger — run via GitHub Actions cron.

Reads current inventory from the same Google Sheet the fulfillment app uses,
and appends today's stock levels to the InventorySnapshots tab so the
Demand & Reorder page can compute out-of-stock-adjusted demand.
"""
import json
import os
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

SHEET_ID = "1t_L1qR3ikD-jjiA2P1tQETQ60hvR1nocX5AimPdSJ7o"
SNAPSHOT_SHEET_NAME = "InventorySnapshots"
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]


def get_client():
    creds_dict = json.loads(os.environ["GCP_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)


def load_inventory(gc):
    ws = gc.open_by_key(SHEET_ID).get_worksheet(0)
    data = ws.get_all_values()
    inv = {}
    for row in data[1:]:
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
            inv[key] = {"product": p, "color": c, "size": s, "qty": qty}
    return inv


def get_snapshot_ws(gc):
    sh = gc.open_by_key(SHEET_ID)
    try:
        return sh.worksheet(SNAPSHOT_SHEET_NAME)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=SNAPSHOT_SHEET_NAME, rows=1000, cols=5)
        ws.append_row(["Date", "Product", "Color", "Size", "Qty"])
        return ws


def main():
    gc = get_client()
    inv = load_inventory(gc)
    ws = get_snapshot_ws(gc)

    today = datetime.now().strftime("%Y-%m-%d")
    if today in set(ws.col_values(1)):
        print(f"Snapshot for {today} already recorded — skipping.")
        return

    rows = [[today, v["product"], v["color"], v["size"], v["qty"]] for v in inv.values()]
    if not rows:
        print("No inventory rows found — nothing recorded.")
        return

    ws.append_rows(rows)
    print(f"Recorded {len(rows)} variant snapshot(s) for {today}.")


if __name__ == "__main__":
    main()
