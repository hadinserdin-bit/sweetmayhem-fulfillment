#!/usr/bin/env python3
"""Periodic Shopify on-hand sync — run via GitHub Actions cron.

The Google Sheet inventory is the source of truth (restocked via the app,
or edited directly in Sheets). This reconciles Shopify's "On hand" quantity
at the primary location to match, for every variant that's out of sync.
"""
import json
import os
import uuid
from difflib import SequenceMatcher

import gspread
import requests
from google.oauth2.service_account import Credentials

SHEET_ID = "1t_L1qR3ikD-jjiA2P1tQETQ60hvR1nocX5AimPdSJ7o"
SHOPIFY_STORE = "6u6sqq-k5.myshopify.com"
SHOPIFY_API_VERSION = "2026-07"
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]


def get_sheet_inventory():
    creds_dict = json.loads(os.environ["GCP_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    gc = gspread.authorize(creds)
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
            continue
        key = (p.lower(), c.lower(), s.lower())
        if key not in inv:
            inv[key] = {"product": p, "color": c, "size": s, "qty": qty}
    return inv


class Shopify:
    def __init__(self, store, token, api_version):
        self.base = f"https://{store}/admin/api/{api_version}"
        self.headers = {"X-Shopify-Access-Token": token, "Content-Type": "application/json"}

    def graphql(self, query, variables=None):
        resp = requests.post(f"{self.base}/graphql.json", headers=self.headers,
                              json={"query": query, "variables": variables or {}})
        resp.raise_for_status()
        data = resp.json()
        if data.get("errors"):
            raise Exception("; ".join(e["message"] for e in data["errors"]))
        return data["data"]

    def primary_location(self):
        resp = requests.get(f"{self.base}/locations.json", headers=self.headers)
        resp.raise_for_status()
        locations = resp.json().get("locations", [])
        if not locations:
            raise Exception("No Shopify locations found.")
        return locations[0]["id"]

    def variant_map(self):
        """(product, color, size) [lowercased] -> inventory_item_id."""
        result = {}
        url = f"{self.base}/products.json"
        params = {"limit": 250, "fields": "id,title,variants"}
        while url:
            resp = requests.get(url, headers=self.headers, params=params)
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

    def onhand_by_item(self, location_id):
        """inventory_item_id -> current on_hand quantity, at one location."""
        result = {}
        cursor = None
        query = """
        query levels($locationId: ID!, $cursor: String) {
          location(id: $locationId) {
            inventoryLevels(first: 250, after: $cursor) {
              edges {
                node {
                  item { id }
                  quantities(names: ["on_hand"]) { quantity }
                }
              }
              pageInfo { hasNextPage endCursor }
            }
          }
        }
        """
        while True:
            data = self.graphql(query, {
                "locationId": f"gid://shopify/Location/{location_id}",
                "cursor": cursor,
            })
            levels = data["location"]["inventoryLevels"]
            for edge in levels["edges"]:
                node = edge["node"]
                item_id = int(node["item"]["id"].rsplit("/", 1)[-1])
                result[item_id] = node["quantities"][0]["quantity"]
            if levels["pageInfo"]["hasNextPage"]:
                cursor = levels["pageInfo"]["endCursor"]
            else:
                break
        return result

    def set_onhand(self, inventory_item_id, location_id, quantity, current):
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
                    "inventoryItemId": f"gid://shopify/InventoryItem/{inventory_item_id}",
                    "locationId": f"gid://shopify/Location/{location_id}",
                    "quantity": quantity,
                    "changeFromQuantity": current,
                }],
            },
            "key": str(uuid.uuid4()),
        }
        data = self.graphql(query, variables)
        errors = data["inventorySetQuantities"]["userErrors"]
        if errors:
            raise Exception("; ".join(e["message"] for e in errors))


def find_inventory_item(variant_map, product, color, size):
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


def main():
    token = os.environ["SHOPIFY_ACCESS_TOKEN"]
    shopify = Shopify(SHOPIFY_STORE, token, SHOPIFY_API_VERSION)

    sheet_inv = get_sheet_inventory()
    print(f"Loaded {len(sheet_inv)} variant(s) from the Google Sheet.")

    variant_map = shopify.variant_map()
    location_id = shopify.primary_location()
    onhand = shopify.onhand_by_item(location_id)
    print(f"Fetched {len(variant_map)} Shopify variant(s), {len(onhand)} inventory level(s).")

    synced, unmatched, failed, unchanged = 0, [], [], 0
    for item in sheet_inv.values():
        label = f"{item['product']} — {item['color']} / {item['size']}"
        inv_item_id = find_inventory_item(variant_map, item["product"], item["color"], item["size"])
        if inv_item_id is None:
            unmatched.append(label)
            continue
        current = onhand.get(inv_item_id)
        if current is None:
            unmatched.append(label)
            continue
        if current == item["qty"]:
            unchanged += 1
            continue
        try:
            shopify.set_onhand(inv_item_id, location_id, item["qty"], current)
            print(f"Synced: {label}: {current} -> {item['qty']}")
            synced += 1
        except Exception as e:
            failed.append(f"{label}: {e}")

    print(f"\n{synced} synced, {unchanged} already in sync, {len(unmatched)} unmatched, {len(failed)} failed.")
    if unmatched:
        print("Unmatched:\n" + "\n".join(f"  - {m}" for m in unmatched))
    if failed:
        print("Failed:\n" + "\n".join(f"  - {f}" for f in failed))


if __name__ == "__main__":
    main()
