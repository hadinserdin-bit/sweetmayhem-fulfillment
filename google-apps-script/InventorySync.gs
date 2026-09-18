/**
 * Instant Shopify on-hand sync for manual edits made directly in this Google Sheet.
 *
 * Setup (one-time):
 *   1. Open this spreadsheet -> Extensions -> Apps Script.
 *   2. Paste this whole file in as Code.gs (replace the default content).
 *   3. Project Settings (gear icon) -> Script Properties -> Add property:
 *        SHOPIFY_ACCESS_TOKEN = <the shpat_... token from secrets.toml>
 *   4. In the function dropdown at the top, select "setupTrigger" and click Run.
 *      Grant the authorization it asks for (it needs permission to call
 *      external services and read this spreadsheet).
 *   5. Edit a quantity cell in this sheet and check Shopify a few seconds later.
 *
 * Note: this only fires for edits made through the Sheets UI. Restocks made
 * through the Streamlit app already sync themselves directly; a separate
 * GitHub Actions job also reconciles everything every 15 minutes as a backstop.
 */

const SHOPIFY_STORE = '6u6sqq-k5.myshopify.com';
const SHOPIFY_API_VERSION = '2026-07';
const INVENTORY_SHEET_NAME = 'Sheet1';
const QTY_COLUMN = 4; // A=Product, B=Color, C=Size, D=Qty

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(t);
    }
  });
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  Logger.log('Trigger installed.');
}

function onEditInstallable(e) {
  try {
    const sheet = e.range.getSheet();
    if (sheet.getName() !== INVENTORY_SHEET_NAME) return;

    const range = e.range;
    const touchesQtyColumn =
      range.getColumn() <= QTY_COLUMN &&
      range.getColumn() + range.getNumColumns() - 1 >= QTY_COLUMN;
    if (!touchesQtyColumn) return;

    const firstRow = Math.max(range.getRow(), 2); // skip header
    const lastRow = range.getRow() + range.getNumRows() - 1;
    if (firstRow > lastRow) return;

    const token = PropertiesService.getScriptProperties().getProperty('SHOPIFY_ACCESS_TOKEN');
    if (!token) throw new Error('SHOPIFY_ACCESS_TOKEN script property is not set.');

    for (let row = firstRow; row <= lastRow; row++) {
      const values = sheet.getRange(row, 1, 1, 4).getValues()[0];
      const product = values[0], color = values[1], size = values[2], qtyRaw = values[3];
      if (!product || !color || !size) continue;
      const qty = Number(qtyRaw);
      if (!Number.isFinite(qty)) continue;
      syncRowToShopify_(token, String(product), String(color), String(size), qty);
    }
  } catch (err) {
    Logger.log('onEditInstallable failed: ' + err);
  }
}

function syncRowToShopify_(token, product, color, size, qty) {
  const inventoryItemId = findShopifyVariant_(token, product, color, size);
  if (!inventoryItemId) {
    Logger.log('No Shopify match for ' + product + ' / ' + color + ' / ' + size);
    return;
  }
  const locationId = getPrimaryLocationId_(token);
  const current = getCurrentOnHand_(token, inventoryItemId, locationId);
  if (current === qty) return;
  setOnHandQuantity_(token, inventoryItemId, locationId, qty, current);
  Logger.log('Synced ' + product + ' / ' + color + ' / ' + size + ': ' + current + ' -> ' + qty);
}

function shopifyGraphQL_(token, query, variables) {
  const url = 'https://' + SHOPIFY_STORE + '/admin/api/' + SHOPIFY_API_VERSION + '/graphql.json';
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'X-Shopify-Access-Token': token },
    payload: JSON.stringify({ query: query, variables: variables || {} }),
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  if (data.errors) throw new Error(JSON.stringify(data.errors));
  return data.data;
}

function getPrimaryLocationId_(token) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('shopify_location_id');
  if (cached) return cached;

  const url = 'https://' + SHOPIFY_STORE + '/admin/api/' + SHOPIFY_API_VERSION + '/locations.json';
  const resp = UrlFetchApp.fetch(url, {
    headers: { 'X-Shopify-Access-Token': token },
    muteHttpExceptions: true,
  });
  const data = JSON.parse(resp.getContentText());
  const locations = data.locations || [];
  if (!locations.length) throw new Error('No Shopify locations found.');
  const id = String(locations[0].id);
  cache.put('shopify_location_id', id, 21600);
  return id;
}

function getVariantList_(token) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('shopify_variant_list');
  if (cached) return JSON.parse(cached);

  const list = [];
  const query =
    'query($cursor: String) {' +
    '  products(first: 100, after: $cursor) {' +
    '    edges {' +
    '      node {' +
    '        title' +
    '        variants(first: 100) {' +
    '          edges { node { title inventoryItem { id } } }' +
    '        }' +
    '      }' +
    '    }' +
    '    pageInfo { hasNextPage endCursor }' +
    '  }' +
    '}';

  let cursor = null;
  while (true) {
    const data = shopifyGraphQL_(token, query, { cursor: cursor });
    data.products.edges.forEach(function (edge) {
      const p = edge.node;
      p.variants.edges.forEach(function (ve) {
        const v = ve.node;
        const title = v.title || '';
        const sepIdx = title.indexOf(' / ');
        if (sepIdx === -1) return;
        list.push({
          product: p.title,
          color: title.slice(0, sepIdx).trim(),
          size: title.slice(sepIdx + 3).trim(),
          itemId: v.inventoryItem.id.split('/').pop(),
        });
      });
    });
    if (data.products.pageInfo.hasNextPage) {
      cursor = data.products.pageInfo.endCursor;
    } else {
      break;
    }
  }

  cache.put('shopify_variant_list', JSON.stringify(list), 21600);
  return list;
}

function findShopifyVariant_(token, product, color, size) {
  const list = getVariantList_(token);
  const pLower = product.toLowerCase(), cLower = color.toLowerCase(), sLower = size.toLowerCase();

  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (v.product.toLowerCase() === pLower && v.color.toLowerCase() === cLower && v.size.toLowerCase() === sLower) {
      return v.itemId;
    }
  }

  let best = null, bestRatio = 0.75;
  for (let i = 0; i < list.length; i++) {
    const v = list[i];
    if (v.color.toLowerCase() !== cLower || v.size.toLowerCase() !== sLower) continue;
    const ratio = similarityRatio_(pLower, v.product.toLowerCase());
    if (ratio > bestRatio) {
      bestRatio = ratio;
      best = v.itemId;
    }
  }
  return best;
}

function getCurrentOnHand_(token, inventoryItemId, locationId) {
  const query =
    'query($itemId: ID!, $locationId: ID!) {' +
    '  inventoryItem(id: $itemId) {' +
    '    inventoryLevel(locationId: $locationId) {' +
    '      quantities(names: ["on_hand"]) { quantity }' +
    '    }' +
    '  }' +
    '}';
  const data = shopifyGraphQL_(token, query, {
    itemId: 'gid://shopify/InventoryItem/' + inventoryItemId,
    locationId: 'gid://shopify/Location/' + locationId,
  });
  const level = data.inventoryItem && data.inventoryItem.inventoryLevel;
  if (!level) throw new Error('No inventory level found for this item at this location.');
  return level.quantities[0].quantity;
}

function setOnHandQuantity_(token, inventoryItemId, locationId, quantity, current) {
  const mutation =
    'mutation setOnHand($input: InventorySetQuantitiesInput!, $key: String!) {' +
    '  inventorySetQuantities(input: $input) @idempotent(key: $key) {' +
    '    userErrors { field message }' +
    '  }' +
    '}';
  const variables = {
    input: {
      name: 'on_hand',
      reason: 'correction',
      quantities: [{
        inventoryItemId: 'gid://shopify/InventoryItem/' + inventoryItemId,
        locationId: 'gid://shopify/Location/' + locationId,
        quantity: quantity,
        changeFromQuantity: current,
      }],
    },
    key: Utilities.getUuid(),
  };
  const data = shopifyGraphQL_(token, mutation, variables);
  const errors = data.inventorySetQuantities.userErrors;
  if (errors && errors.length) {
    throw new Error(errors.map(function (er) { return er.message; }).join('; '));
  }
}

function levenshtein_(a, b) {
  const dp = [];
  for (let i = 0; i <= a.length; i++) dp[i] = [i];
  for (let j = 0; j <= b.length; j++) dp[0][j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      dp[i][j] = a[i - 1] === b[j - 1]
        ? dp[i - 1][j - 1]
        : 1 + Math.min(dp[i - 1][j - 1], dp[i - 1][j], dp[i][j - 1]);
    }
  }
  return dp[a.length][b.length];
}

function similarityRatio_(a, b) {
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein_(a, b) / maxLen;
}
