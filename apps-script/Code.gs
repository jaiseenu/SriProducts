/**
 * Sri Products — Oil Business Management App
 * Backend: Google Apps Script Web App, storing data in this Sheet's tabs.
 *
 * SETUP: see README.md.
 * If you already ran setupSheets from an earlier version, re-run it —
 * it now also patches missing columns onto existing sheets (e.g. the
 * new Sales.Status column) without touching your existing rows.
 */

// ---------- CONFIG ----------

const SHEETS = {
  USERS: 'Users',
  SESSIONS: 'Sessions',
  COUNTERS: 'Counters',
  CUSTOMERS: 'Customers',
  ITEMS: 'Items',
  PRICE_HISTORY: 'PriceHistory',
  BOTTLE_ADJUSTMENTS: 'BottleAdjustments',
  PRODUCTION: 'Production',
  PRODUCTION_INPUTS: 'ProductionInputs',
  PRODUCTION_OUTPUTS: 'ProductionOutputs',
  SALES: 'Sales',
  SALE_ITEMS: 'SaleItems',
  PAYMENTS: 'Payments',
  QUOTATIONS: 'Quotations',
  QUOTATION_ITEMS: 'QuotationItems',
  STOCK_MOVEMENTS: 'StockMovements',
  STOCK_BALANCES: 'StockBalances',
  SETTINGS: 'Settings'
};

const SCHEMA = {
  Users: ['UserId', 'Name', 'Username', 'PasswordHash', 'Role', 'Active', 'CreatedDate'],
  Sessions: ['Token', 'UserId', 'ExpiresAt'],
  Counters: ['Prefix', 'NextNumber'],
  Customers: ['CustomerId', 'Name', 'Active', 'CreatedDate', 'CreatedBy', 'LastModifiedDate', 'LastModifiedBy'],
  Items: ['ItemId', 'Name', 'Type', 'Unit', 'TaxRate', 'Active', 'CreatedDate'],
  PriceHistory: ['PriceId', 'ItemId', 'EffectiveFrom', 'Price', 'CreatedBy', 'CreatedDate'],
  BottleAdjustments: ['AdjustmentId', 'ItemId', 'EffectiveFrom', 'Amount', 'CreatedBy', 'CreatedDate'],
  Production: ['ProductionId', 'Date', 'Notes', 'CreatedBy', 'CreatedDate', 'LastModifiedBy', 'LastModifiedDate'],
  ProductionInputs: ['ProductionInputId', 'ProductionId', 'ItemId', 'QuantityConsumed'],
  ProductionOutputs: ['ProductionOutputId', 'ProductionId', 'ItemId', 'QuantityProduced'],
  Sales: ['SaleId', 'SaleDate', 'CustomerId', 'Subtotal', 'TaxAmount', 'GrandTotal', 'AmountReceived', 'Outstanding', 'GstEnabled', 'QuotationRef', 'Status', 'VoidReason', 'CreatedBy', 'CreatedDate', 'LastModifiedBy', 'LastModifiedDate'],
  SaleItems: ['SaleItemId', 'SaleId', 'ItemId', 'Quantity', 'Unit', 'BaseRate', 'BottleAdjustment', 'FinalRate', 'Amount', 'TaxRate', 'TaxAmount'],
  Payments: ['PaymentId', 'CustomerId', 'SaleId', 'Amount', 'PaymentDate', 'Method', 'Notes', 'CreatedBy', 'CreatedDate'],
  Quotations: ['QuotationId', 'QuotationDate', 'CustomerId', 'Subtotal', 'TaxAmount', 'GrandTotal', 'ValidUntil', 'Status', 'CreatedBy', 'CreatedDate', 'LastModifiedBy', 'LastModifiedDate'],
  QuotationItems: ['QuotationItemId', 'QuotationId', 'ItemId', 'Quantity', 'Unit', 'BaseRate', 'BottleAdjustment', 'FinalRate', 'Amount', 'TaxRate', 'TaxAmount'],
  StockMovements: ['MovementId', 'Date', 'ItemId', 'Quantity', 'Unit', 'MovementType', 'ReferenceType', 'ReferenceId', 'Notes', 'CreatedBy', 'CreatedDate'],
  StockBalances: ['ItemId', 'Balance', 'LastUpdated'],
  Settings: ['Key', 'Value']
};

const DEFAULT_SETTINGS = {
  companyName: 'Sri Products',
  motto: 'Quality never compromised',
  gstEnabled: 'false',
  saleNumberPrefix: 'SALE-',
  quotationNumberPrefix: 'QT-',
  paymentNumberPrefix: 'PAY-',
  customerNumberPrefix: 'CUS-',
  logoUrl: ''
};

const SESSION_HOURS = 12;

// ---------- SETUP / MIGRATION ----------

function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(SCHEMA).forEach(name => {
    let sheet = ss.getSheetByName(name);
    const headers = SCHEMA[name];
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      return;
    }
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      return;
    }
    // Migration: append any headers this version added that the sheet
    // doesn't have yet, without touching existing columns or rows.
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const missing = headers.filter(h => existing.indexOf(h) === -1);
    if (missing.length) {
      sheet.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
    }
  });
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0) ss.deleteSheet(def);

  const settingsSheet = ss.getSheetByName(SHEETS.SETTINGS);
  if (settingsSheet.getLastRow() <= 1) {
    Object.keys(DEFAULT_SETTINGS).forEach(k => settingsSheet.appendRow([k, DEFAULT_SETTINGS[k]]));
  }

  // StockBalances is new — if it has no data rows yet (fresh sheet, or
  // just migrated in), backfill it from the full StockMovements history
  // so current stock isn't zero after upgrading. Safe to run repeatedly:
  // no-ops once balances exist.
  const balancesSheet = ss.getSheetByName(SHEETS.STOCK_BALANCES);
  if (balancesSheet && balancesSheet.getLastRow() <= 1) {
    rebuildStockBalances();
  }
  Logger.log('Sheets initialized / migrated.');
}

/** Edit these, then run this function once from the editor. */
function createFirstAdmin() {
  const username = 'admin';
  const password = 'ChangeMe123!'; // change immediately after first login
  const sheet = getSheet(SHEETS.USERS);
  const userId = nextId('USR', 5);
  sheet.appendRow([userId, 'Administrator', username, hashPassword(password), 'Admin', true, new Date()]);
  Logger.log('Admin created: ' + username + ' / ' + password);
}

// ---------- SHEET / ID HELPERS ----------

function getSheet(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Sheet not found: ' + name + ' — run setupSheets first.');
  return sheet;
}

function sheetToObjects(sheetName) {
  const sheet = getSheet(sheetName);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1);
  return rows
    .map((row, i) => {
      const obj = {};
      headers.forEach((h, idx) => obj[h] = row[idx]);
      obj._row = i + 2;
      return obj;
    })
    .filter(obj => headers.some(h => obj[h] !== '' && obj[h] !== null && obj[h] !== undefined));
}

function appendObject(sheetName, obj) {
  const sheet = getSheet(sheetName);
  const headers = SCHEMA[sheetName];
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sheet.appendRow(row);
  return obj;
}

function updateObjectRow(sheetName, rowIndex, obj) {
  const sheet = getSheet(sheetName);
  const headers = SCHEMA[sheetName];
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
}

function deleteRows(sheetName, rowIndexes) {
  const sheet = getSheet(sheetName);
  rowIndexes.slice().sort((a, b) => b - a).forEach(idx => sheet.deleteRow(idx));
}

function nextId(prefix, padLength) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet(SHEETS.COUNTERS);
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1, current = 0;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === prefix) { rowIndex = i + 1; current = Number(data[i][1]) || 0; break; }
    }
    const next = current + 1;
    if (rowIndex === -1) sheet.appendRow([prefix, next]);
    else sheet.getRange(rowIndex, 2).setValue(next);
    return prefix + String(next).padStart(padLength, '0');
  } finally {
    lock.releaseLock();
  }
}

// ---------- AUTH ----------

function hashPassword(password) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password);
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
}

function login(username, password) {
  const users = sheetToObjects(SHEETS.USERS);
  const user = users.find(u => u.Username === username && u.Active === true);
  if (!user || user.PasswordHash !== hashPassword(password)) throw new Error('Invalid username or password.');
  const token = Utilities.getUuid();
  const expiresAt = new Date(Date.now() + SESSION_HOURS * 3600 * 1000);
  appendObject(SHEETS.SESSIONS, { Token: token, UserId: user.UserId, ExpiresAt: expiresAt });
  return { token: token, user: { userId: user.UserId, name: user.Name, role: user.Role } };
}

function requireSession(token) {
  if (!token) throw new Error('Not authenticated.');
  const session = sheetToObjects(SHEETS.SESSIONS).find(s => s.Token === token);
  if (!session) throw new Error('Session not found. Please log in again.');
  if (new Date(session.ExpiresAt) < new Date()) throw new Error('Session expired. Please log in again.');
  const user = sheetToObjects(SHEETS.USERS).find(u => u.UserId === session.UserId);
  if (!user || !user.Active) throw new Error('User is not active.');
  return { userId: user.UserId, name: user.Name, role: user.Role };
}

function requireAdmin(user) {
  if (user.role !== 'Admin') throw new Error('This action requires Admin access.');
}

/** Deletes the Session row for this token server-side, so "Log out"
 * actually invalidates the token instead of just clearing it from the
 * device's localStorage — matters most on a shared/borrowed phone,
 * where the old token would otherwise stay valid for up to
 * SESSION_HOURS after the user thinks they've logged out. */
function logout(token) {
  if (!token) return { loggedOut: true };
  const rows = sheetToObjects(SHEETS.SESSIONS);
  const row = rows.find(s => s.Token === token);
  if (row) deleteRows(SHEETS.SESSIONS, [row._row]);
  return { loggedOut: true };
}

// ---------- SETTINGS ----------

function getSettings() {
  const out = {};
  sheetToObjects(SHEETS.SETTINGS).forEach(r => out[r.Key] = r.Value);
  return out;
}

function updateSettings(user, updates) {
  requireAdmin(user);
  const sheet = getSheet(SHEETS.SETTINGS);
  const rows = sheetToObjects(SHEETS.SETTINGS);
  Object.keys(updates).forEach(key => {
    const existing = rows.find(r => r.Key === key);
    if (existing) sheet.getRange(existing._row, 2).setValue(updates[key]);
    else sheet.appendRow([key, updates[key]]);
  });
  return getSettings();
}

// ---------- CUSTOMERS ----------

function listCustomers(search) {
  let customers = sheetToObjects(SHEETS.CUSTOMERS);
  if (search) {
    const s = search.toLowerCase();
    customers = customers.filter(c => String(c.Name).toLowerCase().indexOf(s) !== -1);
  }
  return customers;
}

function createCustomer(user, name) {
  if (!name || !name.trim()) throw new Error('Customer name is required.');
  const customerId = nextId(getSettings().customerNumberPrefix || 'CUS-', 5);
  const now = new Date();
  const obj = { CustomerId: customerId, Name: name.trim(), Active: true, CreatedDate: now, CreatedBy: user.name, LastModifiedDate: now, LastModifiedBy: user.name };
  appendObject(SHEETS.CUSTOMERS, obj);
  return obj;
}

/**
 * Resolves a customer for a Sale/Quotation from either an existing
 * customerId or a typed name. Match on name is case-insensitive/trimmed
 * exact match only — no fuzzy matching, to avoid silently merging two
 * different customers. No match => a new customer is created.
 */
function resolveCustomer(user, customerId, customerName) {
  if (customerId) {
    const c = sheetToObjects(SHEETS.CUSTOMERS).find(c => c.CustomerId === customerId);
    if (!c) throw new Error('Customer not found.');
    return c.CustomerId;
  }
  if (!customerName || !customerName.trim()) throw new Error('Customer name is required.');
  const name = customerName.trim();
  const match = sheetToObjects(SHEETS.CUSTOMERS).find(c => String(c.Name).trim().toLowerCase() === name.toLowerCase());
  if (match) return match.CustomerId;
  return createCustomer(user, name).CustomerId;
}

function getCustomerLedger(customerId) {
  const sales = sheetToObjects(SHEETS.SALES).filter(s => s.CustomerId === customerId);
  const activeSales = sales.filter(s => s.Status !== 'Voided');
  const payments = sheetToObjects(SHEETS.PAYMENTS).filter(p => p.CustomerId === customerId);
  const totalSales = activeSales.reduce((sum, s) => sum + Number(s.GrandTotal || 0), 0);
  const totalPayments = payments.reduce((sum, p) => sum + Number(p.Amount || 0), 0);
  return {
    totalSales: totalSales,
    totalPayments: totalPayments,
    outstanding: round2(totalSales - totalPayments),
    sales: sales.sort((a, b) => new Date(b.SaleDate) - new Date(a.SaleDate)),
    payments: payments.sort((a, b) => new Date(b.PaymentDate) - new Date(a.PaymentDate))
  };
}

// ---------- ITEMS ----------

function listItems(includeInactive) {
  const rows = sheetToObjects(SHEETS.ITEMS);
  return includeInactive ? rows : rows.filter(i => i.Active !== false);
}

function createItem(user, item) {
  requireAdmin(user);
  if (!item.name || !item.type || !item.unit) throw new Error('Item name, type, and unit are required.');
  const obj = { ItemId: nextId('ITM', 4), Name: item.name, Type: item.type, Unit: item.unit, TaxRate: item.taxRate || 0, Active: true, CreatedDate: new Date() };
  appendObject(SHEETS.ITEMS, obj);
  return obj;
}

function updateItem(user, itemId, updates) {
  requireAdmin(user);
  const row = sheetToObjects(SHEETS.ITEMS).find(i => i.ItemId === itemId);
  if (!row) throw new Error('Item not found.');
  if (updates.name !== undefined) row.Name = updates.name;
  if (updates.type !== undefined) row.Type = updates.type;
  if (updates.unit !== undefined) row.Unit = updates.unit;
  if (updates.taxRate !== undefined) row.TaxRate = Number(updates.taxRate);
  updateObjectRow(SHEETS.ITEMS, row._row, row);
  return row;
}

function setItemActive(user, itemId, active) {
  requireAdmin(user);
  const row = sheetToObjects(SHEETS.ITEMS).find(i => i.ItemId === itemId);
  if (!row) throw new Error('Item not found.');
  row.Active = !!active;
  updateObjectRow(SHEETS.ITEMS, row._row, row);
  return row;
}

// ---------- PRICE BOOK ----------

function listPrices(itemId) {
  let rows = sheetToObjects(SHEETS.PRICE_HISTORY);
  if (itemId) rows = rows.filter(r => r.ItemId === itemId);
  return rows.sort((a, b) => new Date(b.EffectiveFrom) - new Date(a.EffectiveFrom));
}

/** One row per active item: its latest price on/before today. For list views. */
function getCurrentPrices() {
  const items = listItems();
  const prices = sheetToObjects(SHEETS.PRICE_HISTORY);
  const today = new Date();
  return items.map(i => {
    const applicable = prices.filter(p => p.ItemId === i.ItemId && new Date(p.EffectiveFrom) <= today)
      .sort((a, b) => new Date(b.EffectiveFrom) - new Date(a.EffectiveFrom));
    return { itemId: i.ItemId, name: i.Name, unit: i.Unit, currentPrice: applicable.length ? Number(applicable[0].Price) : null, effectiveFrom: applicable.length ? applicable[0].EffectiveFrom : null };
  });
}

function getCurrentBottleAdjustments() {
  const items = listItems();
  const rows = sheetToObjects(SHEETS.BOTTLE_ADJUSTMENTS);
  const today = new Date();
  return items.map(i => {
    const applicable = rows.filter(r => r.ItemId === i.ItemId && new Date(r.EffectiveFrom) <= today)
      .sort((a, b) => new Date(b.EffectiveFrom) - new Date(a.EffectiveFrom));
    return { itemId: i.ItemId, name: i.Name, unit: i.Unit, currentAmount: applicable.length ? Number(applicable[0].Amount) : null, effectiveFrom: applicable.length ? applicable[0].EffectiveFrom : null };
  });
}

function addPrice(user, itemId, effectiveFrom, price) {
  requireAdmin(user);
  const dupe = sheetToObjects(SHEETS.PRICE_HISTORY).find(r => r.ItemId === itemId && sameDate(r.EffectiveFrom, effectiveFrom));
  if (dupe) throw new Error('A price for this item is already set for that Effective From date. Edit that entry instead.');
  const obj = { PriceId: nextId('PRC', 5), ItemId: itemId, EffectiveFrom: new Date(effectiveFrom), Price: Number(price), CreatedBy: user.name, CreatedDate: new Date() };
  appendObject(SHEETS.PRICE_HISTORY, obj);
  return obj;
}

function updatePrice(user, priceId, effectiveFrom, price) {
  requireAdmin(user);
  const rows = sheetToObjects(SHEETS.PRICE_HISTORY);
  const row = rows.find(r => r.PriceId === priceId);
  if (!row) throw new Error('Price entry not found.');
  const dupe = rows.find(r => r.PriceId !== priceId && r.ItemId === row.ItemId && sameDate(r.EffectiveFrom, effectiveFrom));
  if (dupe) throw new Error('Another price entry already exists for that date.');
  row.EffectiveFrom = new Date(effectiveFrom);
  row.Price = Number(price);
  updateObjectRow(SHEETS.PRICE_HISTORY, row._row, row);
  return row;
}

function deletePrice(user, priceId) {
  requireAdmin(user);
  const row = sheetToObjects(SHEETS.PRICE_HISTORY).find(r => r.PriceId === priceId);
  if (!row) throw new Error('Price entry not found.');
  deleteRows(SHEETS.PRICE_HISTORY, [row._row]);
  return { deleted: true };
}

function getEffectivePrice(itemId, onDate) {
  const target = new Date(onDate);
  const applicable = sheetToObjects(SHEETS.PRICE_HISTORY).filter(r => r.ItemId === itemId && new Date(r.EffectiveFrom) <= target)
    .sort((a, b) => new Date(b.EffectiveFrom) - new Date(a.EffectiveFrom));
  if (applicable.length === 0) throw new Error('No price set for this item on or before ' + onDate + '.');
  return Number(applicable[0].Price);
}

// ---------- BOTTLE ADJUSTMENTS ----------

function listBottleAdjustments(itemId) {
  let rows = sheetToObjects(SHEETS.BOTTLE_ADJUSTMENTS);
  if (itemId) rows = rows.filter(r => r.ItemId === itemId);
  return rows.sort((a, b) => new Date(b.EffectiveFrom) - new Date(a.EffectiveFrom));
}

function addBottleAdjustment(user, itemId, effectiveFrom, amount) {
  requireAdmin(user);
  const dupe = sheetToObjects(SHEETS.BOTTLE_ADJUSTMENTS).find(r => r.ItemId === itemId && sameDate(r.EffectiveFrom, effectiveFrom));
  if (dupe) throw new Error('A bottle adjustment for this item is already set for that Effective From date.');
  const obj = { AdjustmentId: nextId('BTL', 5), ItemId: itemId, EffectiveFrom: new Date(effectiveFrom), Amount: Number(amount), CreatedBy: user.name, CreatedDate: new Date() };
  appendObject(SHEETS.BOTTLE_ADJUSTMENTS, obj);
  return obj;
}

function updateBottleAdjustment(user, adjustmentId, effectiveFrom, amount) {
  requireAdmin(user);
  const rows = sheetToObjects(SHEETS.BOTTLE_ADJUSTMENTS);
  const row = rows.find(r => r.AdjustmentId === adjustmentId);
  if (!row) throw new Error('Adjustment entry not found.');
  const dupe = rows.find(r => r.AdjustmentId !== adjustmentId && r.ItemId === row.ItemId && sameDate(r.EffectiveFrom, effectiveFrom));
  if (dupe) throw new Error('Another adjustment entry already exists for that date.');
  row.EffectiveFrom = new Date(effectiveFrom);
  row.Amount = Number(amount);
  updateObjectRow(SHEETS.BOTTLE_ADJUSTMENTS, row._row, row);
  return row;
}

function deleteBottleAdjustment(user, adjustmentId) {
  requireAdmin(user);
  const row = sheetToObjects(SHEETS.BOTTLE_ADJUSTMENTS).find(r => r.AdjustmentId === adjustmentId);
  if (!row) throw new Error('Adjustment entry not found.');
  deleteRows(SHEETS.BOTTLE_ADJUSTMENTS, [row._row]);
  return { deleted: true };
}

function getEffectiveBottleAdjustment(itemId, onDate) {
  const target = new Date(onDate);
  const applicable = sheetToObjects(SHEETS.BOTTLE_ADJUSTMENTS).filter(r => r.ItemId === itemId && new Date(r.EffectiveFrom) <= target)
    .sort((a, b) => new Date(b.EffectiveFrom) - new Date(a.EffectiveFrom));
  return applicable.length === 0 ? 0 : Number(applicable[0].Amount);
}

// ---------- STOCK ----------
//
// StockMovements is the full append-only ledger — every change ever
// made, kept forever, for audit history. StockBalances is a *running
// total*, one row per item, updated incrementally by every movement.
// getCurrentStock() reads only StockBalances (one small sheet, O(item
// count)) instead of summing the entire lifetime movement history on
// every dashboard/inventory load. The ledger is still the source of
// truth for "what happened and when" — getStockMovementHistory() still
// reads it directly, since that screen is opened rarely, not on every
// page load, and needs the real rows.

function recordStockMovement(itemId, quantity, unit, movementType, referenceType, referenceId, notes, user) {
  const obj = { MovementId: nextId('MOV', 6), Date: new Date(), ItemId: itemId, Quantity: quantity, Unit: unit, MovementType: movementType, ReferenceType: referenceType, ReferenceId: referenceId, Notes: notes || '', CreatedBy: user.name, CreatedDate: new Date() };
  appendObject(SHEETS.STOCK_MOVEMENTS, obj);
  incrementStockBalance(itemId, Number(quantity));
  return obj;
}

/** Adds delta (positive or negative) to an item's running balance.
 * Locked the same way nextId() is, so two near-simultaneous sales of
 * the same item can't both read the same starting balance and clobber
 * each other's update. */
function incrementStockBalance(itemId, delta) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet(SHEETS.STOCK_BALANCES);
    const values = sheet.getDataRange().getValues();
    let rowIndex = -1, current = 0;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === itemId) { rowIndex = i + 1; current = Number(values[i][1]) || 0; break; }
    }
    const next = round2(current + delta);
    const now = new Date();
    if (rowIndex === -1) sheet.appendRow([itemId, next, now]);
    else sheet.getRange(rowIndex, 2, 1, 2).setValues([[next, now]]);
  } finally {
    lock.releaseLock();
  }
}

/** Recomputes every item's balance from the full StockMovements ledger
 * and overwrites StockBalances with the result. Run this from the
 * Apps Script editor (not exposed to the frontend) in two situations:
 * once automatically on first setupSheets() after upgrading to this
 * version, and any time you suspect StockBalances has drifted from the
 * ledger (e.g. a sheet row was edited by hand outside the app). This
 * is the one place that still does a full-ledger scan — that's fine,
 * because it's a manual recovery action, not something that runs on
 * every page load. */
function rebuildStockBalances() {
  const movements = sheetToObjects(SHEETS.STOCK_MOVEMENTS);
  const totals = {};
  movements.forEach(m => { totals[m.ItemId] = round2((totals[m.ItemId] || 0) + Number(m.Quantity)); });
  const sheet = getSheet(SHEETS.STOCK_BALANCES);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, SCHEMA.StockBalances.length).clearContent();
  const now = new Date();
  const rows = Object.keys(totals).map(itemId => [itemId, totals[itemId], now]);
  if (rows.length) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  Logger.log('Stock balances rebuilt for ' + rows.length + ' items.');
}

function getCurrentStock() {
  const balances = sheetToObjects(SHEETS.STOCK_BALANCES);
  const items = sheetToObjects(SHEETS.ITEMS);
  const balanceByItem = {};
  balances.forEach(b => { balanceByItem[b.ItemId] = Number(b.Balance) || 0; });
  return items.map(i => ({ itemId: i.ItemId, name: i.Name, unit: i.Unit, type: i.Type, active: i.Active !== false, currentStock: round2(balanceByItem[i.ItemId] || 0) }));
}

function getStockMovementHistory(itemId) {
  let rows = sheetToObjects(SHEETS.STOCK_MOVEMENTS);
  if (itemId) rows = rows.filter(r => r.ItemId === itemId);
  return rows.sort((a, b) => new Date(b.Date) - new Date(a.Date));
}

function createStockAdjustment(user, itemId, quantity, reason) {
  requireAdmin(user);
  if (!reason || !reason.trim()) throw new Error('A reason is required for stock adjustments.');
  const item = sheetToObjects(SHEETS.ITEMS).find(i => i.ItemId === itemId);
  if (!item) throw new Error('Item not found.');
  return recordStockMovement(itemId, Number(quantity), item.Unit, 'StockAdjustment', 'Manual', '', reason, user);
}

/** Reverses a manual stock adjustment with an offsetting entry — the
 * ledger itself is never mutated or removed, only added to. "Edit" in
 * the UI is implemented as reverse-then-recreate on top of this. */
function reverseStockAdjustment(user, movementId, reason) {
  requireAdmin(user);
  const row = sheetToObjects(SHEETS.STOCK_MOVEMENTS).find(m => m.MovementId === movementId);
  if (!row) throw new Error('Movement not found.');
  if (row.MovementType !== 'StockAdjustment') throw new Error('Only manual stock adjustments can be reversed here — sales and production are reversed via their own Void/Delete.');
  return recordStockMovement(row.ItemId, -Number(row.Quantity), row.Unit, 'StockAdjustment', 'Manual', '', 'Reversal of ' + movementId + (reason ? ': ' + reason : ''), user);
}

// ---------- SALES ----------

function createSale(user, payload) {
  const saleDate = payload.saleDate ? new Date(payload.saleDate) : new Date();
  const customerId = resolveCustomer(user, payload.customerId, payload.customerName);
  if (!payload.items || payload.items.length === 0) throw new Error('At least one line item is required.');

  const settings = getSettings();
  const gstEnabled = settings.gstEnabled === 'true' || settings.gstEnabled === true;
  const items = sheetToObjects(SHEETS.ITEMS);
  const saleId = nextId(settings.saleNumberPrefix || 'SALE-', 6);
  let subtotal = 0, taxTotal = 0;
  const lineItems = [];

  payload.items.forEach(line => {
    const item = items.find(i => i.ItemId === line.itemId);
    if (!item) throw new Error('Item not found: ' + line.itemId);
    const quantity = Number(line.quantity);
    if (!(quantity > 0)) throw new Error('Quantity must be greater than zero for ' + item.Name);
    const baseRate = getEffectivePrice(line.itemId, saleDate);
    const bottleAdj = line.ownBottle ? getEffectiveBottleAdjustment(line.itemId, saleDate) : 0;
    const finalRate = round2(baseRate - bottleAdj);
    const amount = round2(quantity * finalRate);
    const taxRate = gstEnabled ? Number(item.TaxRate || 0) : 0;
    const taxAmount = round2(amount * taxRate / 100);
    subtotal += amount; taxTotal += taxAmount;
    lineItems.push({ SaleItemId: nextId('SI', 6), SaleId: saleId, ItemId: line.itemId, Quantity: quantity, Unit: item.Unit, BaseRate: baseRate, BottleAdjustment: bottleAdj, FinalRate: finalRate, Amount: amount, TaxRate: taxRate, TaxAmount: taxAmount });
  });

  const grandTotal = round2(subtotal + taxTotal);
  const amountReceived = round2(Number(payload.amountReceived || 0));
  const now = new Date();

  const saleObj = { SaleId: saleId, SaleDate: saleDate, CustomerId: customerId, Subtotal: round2(subtotal), TaxAmount: round2(taxTotal), GrandTotal: grandTotal, AmountReceived: amountReceived, Outstanding: round2(grandTotal - amountReceived), GstEnabled: gstEnabled, QuotationRef: payload.quotationRef || '', Status: 'Active', VoidReason: '', CreatedBy: user.name, CreatedDate: now, LastModifiedBy: user.name, LastModifiedDate: now };
  appendObject(SHEETS.SALES, saleObj);
  lineItems.forEach(li => appendObject(SHEETS.SALE_ITEMS, li));
  lineItems.forEach(li => recordStockMovement(li.ItemId, -Math.abs(li.Quantity), li.Unit, 'Sale', 'Sale', saleId, '', user));

  if (amountReceived > 0) {
    appendObject(SHEETS.PAYMENTS, { PaymentId: nextId(settings.paymentNumberPrefix || 'PAY-', 6), CustomerId: customerId, SaleId: saleId, Amount: amountReceived, PaymentDate: saleDate, Method: payload.paymentMethod || 'Cash', Notes: 'Received at time of sale', CreatedBy: user.name, CreatedDate: now });
  }
  return { sale: saleObj, items: lineItems };
}

/** Edits items/customer/date on an existing, non-voided sale. Reverses
 * the old stock impact and applies the new one; AmountReceived is left
 * untouched (payments are managed separately via recordPayment). */
function updateSale(user, saleId, payload) {
  const saleRow = sheetToObjects(SHEETS.SALES).find(s => s.SaleId === saleId);
  if (!saleRow) throw new Error('Sale not found.');
  if (saleRow.Status === 'Voided') throw new Error('This sale has been voided and cannot be edited.');
  if (!payload.items || payload.items.length === 0) throw new Error('At least one line item is required.');

  const oldItems = sheetToObjects(SHEETS.SALE_ITEMS).filter(si => si.SaleId === saleId);
  oldItems.forEach(li => recordStockMovement(li.ItemId, Math.abs(Number(li.Quantity)), li.Unit, 'SaleEdit', 'Sale', saleId, 'Reversed for edit', user));
  deleteRows(SHEETS.SALE_ITEMS, oldItems.map(r => r._row));

  const customerId = resolveCustomer(user, payload.customerId, payload.customerName);
  const saleDate = payload.saleDate ? new Date(payload.saleDate) : new Date(saleRow.SaleDate);
  const settings = getSettings();
  const gstEnabled = settings.gstEnabled === 'true' || settings.gstEnabled === true;
  const items = sheetToObjects(SHEETS.ITEMS);
  let subtotal = 0, taxTotal = 0;
  const lineItems = [];
  payload.items.forEach(line => {
    const item = items.find(i => i.ItemId === line.itemId);
    if (!item) throw new Error('Item not found: ' + line.itemId);
    const quantity = Number(line.quantity);
    if (!(quantity > 0)) throw new Error('Quantity must be greater than zero for ' + item.Name);
    const baseRate = getEffectivePrice(line.itemId, saleDate);
    const bottleAdj = line.ownBottle ? getEffectiveBottleAdjustment(line.itemId, saleDate) : 0;
    const finalRate = round2(baseRate - bottleAdj);
    const amount = round2(quantity * finalRate);
    const taxRate = gstEnabled ? Number(item.TaxRate || 0) : 0;
    const taxAmount = round2(amount * taxRate / 100);
    subtotal += amount; taxTotal += taxAmount;
    lineItems.push({ SaleItemId: nextId('SI', 6), SaleId: saleId, ItemId: line.itemId, Quantity: quantity, Unit: item.Unit, BaseRate: baseRate, BottleAdjustment: bottleAdj, FinalRate: finalRate, Amount: amount, TaxRate: taxRate, TaxAmount: taxAmount });
  });
  const grandTotal = round2(subtotal + taxTotal);

  saleRow.CustomerId = customerId;
  saleRow.SaleDate = saleDate;
  saleRow.Subtotal = round2(subtotal);
  saleRow.TaxAmount = round2(taxTotal);
  saleRow.GrandTotal = grandTotal;
  saleRow.Outstanding = round2(grandTotal - Number(saleRow.AmountReceived || 0));
  saleRow.LastModifiedBy = user.name;
  saleRow.LastModifiedDate = new Date();
  updateObjectRow(SHEETS.SALES, saleRow._row, saleRow);
  lineItems.forEach(li => appendObject(SHEETS.SALE_ITEMS, li));
  lineItems.forEach(li => recordStockMovement(li.ItemId, -Math.abs(li.Quantity), li.Unit, 'SaleEdit', 'Sale', saleId, 'Applied after edit', user));
  return { sale: saleRow, items: lineItems };
}

/** Void, not delete: reverses stock via an offsetting ledger entry and
 * marks the sale Voided (excluded from revenue totals) but keeps the
 * row — deleting it outright would orphan its stock movements and any
 * payment recorded against it. */
function voidSale(user, saleId, reason) {
  requireAdmin(user);
  if (!reason || !reason.trim()) throw new Error('A reason is required to void a sale.');
  const saleRow = sheetToObjects(SHEETS.SALES).find(s => s.SaleId === saleId);
  if (!saleRow) throw new Error('Sale not found.');
  if (saleRow.Status === 'Voided') throw new Error('This sale is already voided.');
  const items = sheetToObjects(SHEETS.SALE_ITEMS).filter(si => si.SaleId === saleId);
  items.forEach(li => recordStockMovement(li.ItemId, Math.abs(Number(li.Quantity)), li.Unit, 'SaleVoid', 'Sale', saleId, 'Sale voided: ' + reason, user));
  saleRow.Status = 'Voided';
  saleRow.VoidReason = reason;
  saleRow.Outstanding = 0;
  saleRow.LastModifiedBy = user.name;
  saleRow.LastModifiedDate = new Date();
  updateObjectRow(SHEETS.SALES, saleRow._row, saleRow);
  return saleRow;
}

/** Adds customerName and a per-line item summary to raw Sales rows —
 * used everywhere a sale list is shown so the UI never needs a separate
 * round trip just to label a row. */
function enrichSalesWithDetails(sales) {
  const allSaleItems = sheetToObjects(SHEETS.SALE_ITEMS);
  const items = sheetToObjects(SHEETS.ITEMS);
  const customers = sheetToObjects(SHEETS.CUSTOMERS);
  const itemName = id => { const it = items.find(i => i.ItemId === id); return it ? it.Name : id; };
  const custName = id => { const c = customers.find(c => c.CustomerId === id); return c ? c.Name : id; };
  return sales.map(s => Object.assign({}, s, {
    customerName: custName(s.CustomerId),
    items: allSaleItems.filter(si => si.SaleId === s.SaleId).map(si => ({ itemName: itemName(si.ItemId), quantity: si.Quantity, unit: si.Unit }))
  }));
}

function listSales(filters) {
  let sales = sheetToObjects(SHEETS.SALES);
  filters = filters || {};
  if (filters.customerId) sales = sales.filter(s => s.CustomerId === filters.customerId);
  if (filters.fromDate) sales = sales.filter(s => new Date(s.SaleDate) >= new Date(filters.fromDate));
  if (filters.toDate) sales = sales.filter(s => new Date(s.SaleDate) <= new Date(filters.toDate));
  sales = sales.sort((a, b) => new Date(b.SaleDate) - new Date(a.SaleDate));
  return enrichSalesWithDetails(sales);
}

function getSaleDetail(saleId) {
  const sale = sheetToObjects(SHEETS.SALES).find(s => s.SaleId === saleId);
  if (!sale) throw new Error('Sale not found.');
  const items = sheetToObjects(SHEETS.SALE_ITEMS).filter(i => i.SaleId === saleId);
  const customer = sheetToObjects(SHEETS.CUSTOMERS).find(c => c.CustomerId === sale.CustomerId);
  const itemDefs = sheetToObjects(SHEETS.ITEMS);
  const enrichedItems = items.map(i => Object.assign({}, i, { itemName: (itemDefs.find(d => d.ItemId === i.ItemId) || {}).Name || i.ItemId }));
  return { sale: sale, items: enrichedItems, customer: customer };
}

// ---------- PAYMENTS ----------

/**
 * payload: { customerId, amount (total received), paymentDate, method,
 * notes, allocations: [{saleId, amount}] }.
 * Each allocation creates its own Payment row AND updates that Sale's
 * AmountReceived/Outstanding — this is what keeps a sale's own paid
 * status accurate over time instead of only the customer-level ledger.
 * If allocations sum to less than the total amount, the remainder is
 * recorded as one unallocated Payment row (an advance/credit not tied
 * to any specific sale). Allocations summing to more than the total
 * amount is rejected — fix the split or increase the amount.
 */
function recordPayment(user, payload) {
  const customer = sheetToObjects(SHEETS.CUSTOMERS).find(c => c.CustomerId === payload.customerId);
  if (!customer) throw new Error('Customer not found.');
  const totalAmount = round2(Number(payload.amount || 0));
  if (!(totalAmount > 0)) throw new Error('Payment amount must be greater than zero.');

  const settings = getSettings();
  const paymentDate = payload.paymentDate ? new Date(payload.paymentDate) : new Date();
  const method = payload.method || 'Cash';
  const notes = payload.notes || '';
  const now = new Date();

  const allocations = (payload.allocations || [])
    .filter(a => a.saleId && Number(a.amount) > 0)
    .map(a => ({ saleId: a.saleId, amount: round2(Number(a.amount)) }));
  const allocatedTotal = round2(allocations.reduce((s, a) => s + a.amount, 0));
  if (allocatedTotal - totalAmount > 0.01) throw new Error('The amounts applied to sales add up to more than the payment amount.');

  const allSales = sheetToObjects(SHEETS.SALES);
  const createdPayments = [];

  allocations.forEach(a => {
    const saleRow = allSales.find(s => s.SaleId === a.saleId && s.CustomerId === payload.customerId);
    if (!saleRow) throw new Error('Sale not found for allocation: ' + a.saleId);
    if (saleRow.Status === 'Voided') throw new Error('Cannot apply a payment to a voided sale: ' + a.saleId);
    const paymentObj = { PaymentId: nextId(settings.paymentNumberPrefix || 'PAY-', 6), CustomerId: payload.customerId, SaleId: a.saleId, Amount: a.amount, PaymentDate: paymentDate, Method: method, Notes: notes, CreatedBy: user.name, CreatedDate: now };
    appendObject(SHEETS.PAYMENTS, paymentObj);
    createdPayments.push(paymentObj);
    saleRow.AmountReceived = round2(Number(saleRow.AmountReceived || 0) + a.amount);
    saleRow.Outstanding = round2(Number(saleRow.GrandTotal) - saleRow.AmountReceived);
    saleRow.LastModifiedBy = user.name;
    saleRow.LastModifiedDate = now;
    updateObjectRow(SHEETS.SALES, saleRow._row, saleRow);
  });

  const remainder = round2(totalAmount - allocatedTotal);
  if (remainder > 0.004) {
    const paymentObj = { PaymentId: nextId(settings.paymentNumberPrefix || 'PAY-', 6), CustomerId: payload.customerId, SaleId: '', Amount: remainder, PaymentDate: paymentDate, Method: method, Notes: notes || 'Unallocated / advance', CreatedBy: user.name, CreatedDate: now };
    appendObject(SHEETS.PAYMENTS, paymentObj);
    createdPayments.push(paymentObj);
  }
  return { payments: createdPayments };
}

/** One-tap version of recordPayment for a single sale's full outstanding
 * balance — same effect as allocating the whole amount to that one sale. */
function markSaleAsPaid(user, saleId, method) {
  const allSales = sheetToObjects(SHEETS.SALES);
  const saleRow = allSales.find(s => s.SaleId === saleId);
  if (!saleRow) throw new Error('Sale not found.');
  if (saleRow.Status === 'Voided') throw new Error('This sale is voided.');
  const outstanding = round2(Number(saleRow.GrandTotal) - Number(saleRow.AmountReceived || 0));
  if (outstanding <= 0) throw new Error('This sale is already fully paid.');
  const settings = getSettings();
  const now = new Date();
  appendObject(SHEETS.PAYMENTS, { PaymentId: nextId(settings.paymentNumberPrefix || 'PAY-', 6), CustomerId: saleRow.CustomerId, SaleId: saleId, Amount: outstanding, PaymentDate: now, Method: method || 'Cash', Notes: 'Marked as paid', CreatedBy: user.name, CreatedDate: now });
  saleRow.AmountReceived = saleRow.GrandTotal;
  saleRow.Outstanding = 0;
  saleRow.LastModifiedBy = user.name;
  saleRow.LastModifiedDate = now;
  updateObjectRow(SHEETS.SALES, saleRow._row, saleRow);
  return saleRow;
}


function listPayments(filters) {
  let payments = sheetToObjects(SHEETS.PAYMENTS);
  filters = filters || {};
  if (filters.customerId) payments = payments.filter(p => p.CustomerId === filters.customerId);
  return payments.sort((a, b) => new Date(b.PaymentDate) - new Date(a.PaymentDate));
}

// ---------- QUOTATIONS ----------

function createQuotation(user, payload) {
  const quoteDate = payload.quotationDate ? new Date(payload.quotationDate) : new Date();
  const customerId = resolveCustomer(user, payload.customerId, payload.customerName);
  if (!payload.items || payload.items.length === 0) throw new Error('At least one line item is required.');

  const settings = getSettings();
  const gstEnabled = settings.gstEnabled === 'true' || settings.gstEnabled === true;
  const items = sheetToObjects(SHEETS.ITEMS);
  const quotationId = nextId(settings.quotationNumberPrefix || 'QT-', 6);
  let subtotal = 0, taxTotal = 0;
  const lineItems = [];
  payload.items.forEach(line => {
    const item = items.find(i => i.ItemId === line.itemId);
    if (!item) throw new Error('Item not found: ' + line.itemId);
    const quantity = Number(line.quantity);
    if (!(quantity > 0)) throw new Error('Quantity must be greater than zero for ' + item.Name);
    const baseRate = getEffectivePrice(line.itemId, quoteDate);
    const bottleAdj = line.ownBottle ? getEffectiveBottleAdjustment(line.itemId, quoteDate) : 0;
    const finalRate = round2(baseRate - bottleAdj);
    const amount = round2(quantity * finalRate);
    const taxRate = gstEnabled ? Number(item.TaxRate || 0) : 0;
    const taxAmount = round2(amount * taxRate / 100);
    subtotal += amount; taxTotal += taxAmount;
    lineItems.push({ QuotationItemId: nextId('QI', 6), QuotationId: quotationId, ItemId: line.itemId, Quantity: quantity, Unit: item.Unit, BaseRate: baseRate, BottleAdjustment: bottleAdj, FinalRate: finalRate, Amount: amount, TaxRate: taxRate, TaxAmount: taxAmount });
  });

  const obj = { QuotationId: quotationId, QuotationDate: quoteDate, CustomerId: customerId, Subtotal: round2(subtotal), TaxAmount: round2(taxTotal), GrandTotal: round2(subtotal + taxTotal), ValidUntil: payload.validUntil ? new Date(payload.validUntil) : '', Status: 'Draft', CreatedBy: user.name, CreatedDate: new Date(), LastModifiedBy: user.name, LastModifiedDate: new Date() };
  appendObject(SHEETS.QUOTATIONS, obj);
  lineItems.forEach(li => appendObject(SHEETS.QUOTATION_ITEMS, li));
  return { quotation: obj, items: lineItems };
}

function updateQuotation(user, quotationId, payload) {
  const qRow = sheetToObjects(SHEETS.QUOTATIONS).find(q => q.QuotationId === quotationId);
  if (!qRow) throw new Error('Quotation not found.');
  if (qRow.Status === 'Converted') throw new Error('This quotation has been converted to a sale and cannot be edited.');
  if (!payload.items || payload.items.length === 0) throw new Error('At least one line item is required.');

  const oldItems = sheetToObjects(SHEETS.QUOTATION_ITEMS).filter(qi => qi.QuotationId === quotationId);
  deleteRows(SHEETS.QUOTATION_ITEMS, oldItems.map(r => r._row));

  const customerId = resolveCustomer(user, payload.customerId, payload.customerName);
  const quoteDate = payload.quotationDate ? new Date(payload.quotationDate) : new Date(qRow.QuotationDate);
  const settings = getSettings();
  const gstEnabled = settings.gstEnabled === 'true' || settings.gstEnabled === true;
  const items = sheetToObjects(SHEETS.ITEMS);
  let subtotal = 0, taxTotal = 0;
  const lineItems = [];
  payload.items.forEach(line => {
    const item = items.find(i => i.ItemId === line.itemId);
    if (!item) throw new Error('Item not found: ' + line.itemId);
    const quantity = Number(line.quantity);
    if (!(quantity > 0)) throw new Error('Quantity must be greater than zero for ' + item.Name);
    const baseRate = getEffectivePrice(line.itemId, quoteDate);
    const bottleAdj = line.ownBottle ? getEffectiveBottleAdjustment(line.itemId, quoteDate) : 0;
    const finalRate = round2(baseRate - bottleAdj);
    const amount = round2(quantity * finalRate);
    const taxRate = gstEnabled ? Number(item.TaxRate || 0) : 0;
    const taxAmount = round2(amount * taxRate / 100);
    subtotal += amount; taxTotal += taxAmount;
    lineItems.push({ QuotationItemId: nextId('QI', 6), QuotationId: quotationId, ItemId: line.itemId, Quantity: quantity, Unit: item.Unit, BaseRate: baseRate, BottleAdjustment: bottleAdj, FinalRate: finalRate, Amount: amount, TaxRate: taxRate, TaxAmount: taxAmount });
  });

  qRow.CustomerId = customerId;
  qRow.QuotationDate = quoteDate;
  qRow.Subtotal = round2(subtotal);
  qRow.TaxAmount = round2(taxTotal);
  qRow.GrandTotal = round2(subtotal + taxTotal);
  if (payload.validUntil) qRow.ValidUntil = new Date(payload.validUntil);
  qRow.LastModifiedBy = user.name;
  qRow.LastModifiedDate = new Date();
  updateObjectRow(SHEETS.QUOTATIONS, qRow._row, qRow);
  lineItems.forEach(li => appendObject(SHEETS.QUOTATION_ITEMS, li));
  return { quotation: qRow, items: lineItems };
}

function deleteQuotation(user, quotationId) {
  const qRow = sheetToObjects(SHEETS.QUOTATIONS).find(q => q.QuotationId === quotationId);
  if (!qRow) throw new Error('Quotation not found.');
  if (qRow.Status === 'Converted') throw new Error('This quotation has been converted to a sale and cannot be deleted.');
  const items = sheetToObjects(SHEETS.QUOTATION_ITEMS).filter(qi => qi.QuotationId === quotationId);
  deleteRows(SHEETS.QUOTATION_ITEMS, items.map(r => r._row));
  deleteRows(SHEETS.QUOTATIONS, [qRow._row]);
  return { deleted: true };
}

function listQuotations(filters) {
  let rows = sheetToObjects(SHEETS.QUOTATIONS);
  filters = filters || {};
  if (filters.customerId) rows = rows.filter(r => r.CustomerId === filters.customerId);
  if (filters.status) rows = rows.filter(r => r.Status === filters.status);
  return rows.map(effectiveQuotationStatus).sort((a, b) => new Date(b.QuotationDate) - new Date(a.QuotationDate));
}

function effectiveQuotationStatus(q) {
  if ((q.Status === 'Draft' || q.Status === 'Sent') && q.ValidUntil && new Date(q.ValidUntil) < new Date()) {
    return Object.assign({}, q, { Status: 'Expired' });
  }
  return q;
}

function getQuotationDetail(quotationId) {
  const quotation = sheetToObjects(SHEETS.QUOTATIONS).find(q => q.QuotationId === quotationId);
  if (!quotation) throw new Error('Quotation not found.');
  const items = sheetToObjects(SHEETS.QUOTATION_ITEMS).filter(i => i.QuotationId === quotationId);
  const customer = sheetToObjects(SHEETS.CUSTOMERS).find(c => c.CustomerId === quotation.CustomerId);
  const itemDefs = sheetToObjects(SHEETS.ITEMS);
  const enrichedItems = items.map(i => Object.assign({}, i, { itemName: (itemDefs.find(d => d.ItemId === i.ItemId) || {}).Name || i.ItemId }));
  return { quotation: effectiveQuotationStatus(quotation), items: enrichedItems, customer: customer };
}

function updateQuotationStatus(user, quotationId, status) {
  const row = sheetToObjects(SHEETS.QUOTATIONS).find(r => r.QuotationId === quotationId);
  if (!row) throw new Error('Quotation not found.');
  row.Status = status;
  row.LastModifiedBy = user.name;
  row.LastModifiedDate = new Date();
  updateObjectRow(SHEETS.QUOTATIONS, row._row, row);
  return row;
}

function convertQuotationToSale(user, quotationId, paymentAmount, paymentMethod) {
  const detail = getQuotationDetail(quotationId);
  const quotation = detail.quotation;
  if (quotation.Status === 'Converted') throw new Error('This quotation has already been converted to a sale.');
  if (quotation.Status === 'Expired') throw new Error('This quotation has expired and cannot be converted. Create a new quotation instead.');

  const stockLevels = getCurrentStock();
  const stockWarnings = [];
  detail.items.forEach(li => {
    const stockRow = stockLevels.find(s => s.itemId === li.ItemId);
    const available = stockRow ? stockRow.currentStock : 0;
    if (available < Number(li.Quantity)) stockWarnings.push(li.itemName + ': requested ' + li.Quantity + ' ' + li.Unit + ', only ' + available + ' ' + li.Unit + ' in stock.');
  });

  const settings = getSettings();
  const saleId = nextId(settings.saleNumberPrefix || 'SALE-', 6);
  const amountReceived = round2(Number(paymentAmount || 0));
  const now = new Date();

  const saleObj = { SaleId: saleId, SaleDate: now, CustomerId: quotation.CustomerId, Subtotal: quotation.Subtotal, TaxAmount: quotation.TaxAmount, GrandTotal: quotation.GrandTotal, AmountReceived: amountReceived, Outstanding: round2(quotation.GrandTotal - amountReceived), GstEnabled: quotation.TaxAmount > 0, QuotationRef: quotationId, Status: 'Active', VoidReason: '', CreatedBy: user.name, CreatedDate: now, LastModifiedBy: user.name, LastModifiedDate: now };
  appendObject(SHEETS.SALES, saleObj);

  detail.items.forEach(li => {
    appendObject(SHEETS.SALE_ITEMS, { SaleItemId: nextId('SI', 6), SaleId: saleId, ItemId: li.ItemId, Quantity: li.Quantity, Unit: li.Unit, BaseRate: li.BaseRate, BottleAdjustment: li.BottleAdjustment, FinalRate: li.FinalRate, Amount: li.Amount, TaxRate: li.TaxRate, TaxAmount: li.TaxAmount });
    recordStockMovement(li.ItemId, -Math.abs(Number(li.Quantity)), li.Unit, 'Sale', 'Sale', saleId, 'Converted from ' + quotationId, user);
  });

  if (amountReceived > 0) {
    appendObject(SHEETS.PAYMENTS, { PaymentId: nextId(settings.paymentNumberPrefix || 'PAY-', 6), CustomerId: quotation.CustomerId, SaleId: saleId, Amount: amountReceived, PaymentDate: now, Method: paymentMethod || 'Cash', Notes: 'Received on quotation conversion', CreatedBy: user.name, CreatedDate: now });
  }
  updateQuotationStatus(user, quotationId, 'Converted');
  return { sale: saleObj, stockWarnings: stockWarnings };
}

// ---------- PRODUCTION ----------

function createProduction(user, payload) {
  if (!payload.inputs || payload.inputs.length === 0) throw new Error('At least one raw material input is required.');
  if (!payload.outputs || payload.outputs.length === 0) throw new Error('At least one output is required.');
  const productionId = nextId('PRD', 6);
  const date = payload.date ? new Date(payload.date) : new Date();
  const items = sheetToObjects(SHEETS.ITEMS);
  const now = new Date();

  appendObject(SHEETS.PRODUCTION, { ProductionId: productionId, Date: date, Notes: payload.notes || '', CreatedBy: user.name, CreatedDate: now, LastModifiedBy: user.name, LastModifiedDate: now });

  payload.inputs.forEach(inp => {
    const item = items.find(i => i.ItemId === inp.itemId);
    if (!item) throw new Error('Item not found: ' + inp.itemId);
    const qty = Number(inp.quantity);
    if (!(qty > 0)) throw new Error('Input quantity must be greater than zero.');
    appendObject(SHEETS.PRODUCTION_INPUTS, { ProductionInputId: nextId('PI', 6), ProductionId: productionId, ItemId: inp.itemId, QuantityConsumed: qty });
    recordStockMovement(inp.itemId, -qty, item.Unit, 'Production', 'Production', productionId, 'Consumed', user);
  });
  payload.outputs.forEach(out => {
    const item = items.find(i => i.ItemId === out.itemId);
    if (!item) throw new Error('Item not found: ' + out.itemId);
    const qty = Number(out.quantity);
    if (!(qty > 0)) throw new Error('Output quantity must be greater than zero.');
    appendObject(SHEETS.PRODUCTION_OUTPUTS, { ProductionOutputId: nextId('PO', 6), ProductionId: productionId, ItemId: out.itemId, QuantityProduced: qty });
    recordStockMovement(out.itemId, qty, item.Unit, 'Production', 'Production', productionId, 'Produced', user);
  });
  return { productionId: productionId };
}

function updateProduction(user, productionId, payload) {
  const pRow = sheetToObjects(SHEETS.PRODUCTION).find(p => p.ProductionId === productionId);
  if (!pRow) throw new Error('Production entry not found.');
  if (!payload.inputs || !payload.inputs.length) throw new Error('At least one raw material input is required.');
  if (!payload.outputs || !payload.outputs.length) throw new Error('At least one output is required.');

  const items = sheetToObjects(SHEETS.ITEMS);
  const oldInputs = sheetToObjects(SHEETS.PRODUCTION_INPUTS).filter(i => i.ProductionId === productionId);
  const oldOutputs = sheetToObjects(SHEETS.PRODUCTION_OUTPUTS).filter(o => o.ProductionId === productionId);
  oldInputs.forEach(inp => { const item = items.find(i => i.ItemId === inp.ItemId); recordStockMovement(inp.ItemId, Number(inp.QuantityConsumed), item ? item.Unit : '', 'ProductionEdit', 'Production', productionId, 'Reversed for edit', user); });
  oldOutputs.forEach(out => { const item = items.find(i => i.ItemId === out.ItemId); recordStockMovement(out.ItemId, -Number(out.QuantityProduced), item ? item.Unit : '', 'ProductionEdit', 'Production', productionId, 'Reversed for edit', user); });
  deleteRows(SHEETS.PRODUCTION_INPUTS, oldInputs.map(r => r._row));
  deleteRows(SHEETS.PRODUCTION_OUTPUTS, oldOutputs.map(r => r._row));

  payload.inputs.forEach(inp => {
    const item = items.find(i => i.ItemId === inp.itemId);
    if (!item) throw new Error('Item not found: ' + inp.itemId);
    const qty = Number(inp.quantity);
    if (!(qty > 0)) throw new Error('Input quantity must be greater than zero.');
    appendObject(SHEETS.PRODUCTION_INPUTS, { ProductionInputId: nextId('PI', 6), ProductionId: productionId, ItemId: inp.itemId, QuantityConsumed: qty });
    recordStockMovement(inp.itemId, -qty, item.Unit, 'ProductionEdit', 'Production', productionId, 'Applied after edit', user);
  });
  payload.outputs.forEach(out => {
    const item = items.find(i => i.ItemId === out.itemId);
    if (!item) throw new Error('Item not found: ' + out.itemId);
    const qty = Number(out.quantity);
    if (!(qty > 0)) throw new Error('Output quantity must be greater than zero.');
    appendObject(SHEETS.PRODUCTION_OUTPUTS, { ProductionOutputId: nextId('PO', 6), ProductionId: productionId, ItemId: out.itemId, QuantityProduced: qty });
    recordStockMovement(out.itemId, qty, item.Unit, 'ProductionEdit', 'Production', productionId, 'Applied after edit', user);
  });

  pRow.Date = payload.date ? new Date(payload.date) : pRow.Date;
  pRow.Notes = payload.notes || '';
  pRow.LastModifiedBy = user.name;
  pRow.LastModifiedDate = new Date();
  updateObjectRow(SHEETS.PRODUCTION, pRow._row, pRow);
  return { productionId: productionId };
}

/** Reverses stock via offsetting entries, then removes the production
 * record itself — Production has no downstream references the way a
 * Sale's QuotationRef does, so a real delete is safe here. */
function deleteProduction(user, productionId) {
  requireAdmin(user);
  const pRow = sheetToObjects(SHEETS.PRODUCTION).find(p => p.ProductionId === productionId);
  if (!pRow) throw new Error('Production entry not found.');
  const items = sheetToObjects(SHEETS.ITEMS);
  const inputs = sheetToObjects(SHEETS.PRODUCTION_INPUTS).filter(i => i.ProductionId === productionId);
  const outputs = sheetToObjects(SHEETS.PRODUCTION_OUTPUTS).filter(o => o.ProductionId === productionId);
  inputs.forEach(inp => { const item = items.find(i => i.ItemId === inp.ItemId); recordStockMovement(inp.ItemId, Number(inp.QuantityConsumed), item ? item.Unit : '', 'ProductionDelete', 'Production', productionId, 'Production entry deleted', user); });
  outputs.forEach(out => { const item = items.find(i => i.ItemId === out.ItemId); recordStockMovement(out.ItemId, -Number(out.QuantityProduced), item ? item.Unit : '', 'ProductionDelete', 'Production', productionId, 'Production entry deleted', user); });
  deleteRows(SHEETS.PRODUCTION_INPUTS, inputs.map(r => r._row));
  deleteRows(SHEETS.PRODUCTION_OUTPUTS, outputs.map(r => r._row));
  deleteRows(SHEETS.PRODUCTION, [pRow._row]);
  return { deleted: true };
}

function listProduction() {
  const productions = sheetToObjects(SHEETS.PRODUCTION).sort((a, b) => new Date(b.Date) - new Date(a.Date));
  const inputs = sheetToObjects(SHEETS.PRODUCTION_INPUTS);
  const outputs = sheetToObjects(SHEETS.PRODUCTION_OUTPUTS);
  const items = sheetToObjects(SHEETS.ITEMS);
  const nameOf = id => { const it = items.find(i => i.ItemId === id); return it ? it.Name : id; };
  return productions.map(p => ({
    productionId: p.ProductionId, date: p.Date, notes: p.Notes,
    inputs: inputs.filter(i => i.ProductionId === p.ProductionId).map(i => ({ itemId: i.ItemId, itemName: nameOf(i.ItemId), quantity: i.QuantityConsumed })),
    outputs: outputs.filter(o => o.ProductionId === p.ProductionId).map(o => ({ itemId: o.ItemId, itemName: nameOf(o.ItemId), quantity: o.QuantityProduced }))
  }));
}

function getProductionDetail(productionId) {
  const all = listProduction();
  const found = all.find(p => p.productionId === productionId);
  if (!found) throw new Error('Production entry not found.');
  return found;
}

// ---------- DASHBOARD ----------

function getDashboard() {
  const today = new Date(); today.setHours(0, 0, 0, 0);
  const monthStart = new Date(today.getFullYear(), today.getMonth(), 1);
  const sales = sheetToObjects(SHEETS.SALES).filter(s => s.Status !== 'Voided');
  const payments = sheetToObjects(SHEETS.PAYMENTS);
  const todaySales = sales.filter(s => new Date(s.SaleDate) >= today);
  const monthSales = sales.filter(s => new Date(s.SaleDate) >= monthStart);
  const todayPayments = payments.filter(p => new Date(p.PaymentDate) >= today);

  // Split net outstanding into what's owed TO you vs. credit you owe
  // BACK (from an overpayment) — netting them into one number hides
  // the second case entirely.
  const pendingReceivable = round2(sales.reduce((sum, s) => sum + Math.max(0, Number(s.Outstanding || 0)), 0));
  const creditsOwed = round2(sales.reduce((sum, s) => sum + Math.max(0, -Number(s.Outstanding || 0)), 0));

  const recentSales = enrichSalesWithDetails(sheetToObjects(SHEETS.SALES).sort((a, b) => new Date(b.CreatedDate) - new Date(a.CreatedDate)).slice(0, 5));
  const recentPayments = payments.sort((a, b) => new Date(b.CreatedDate) - new Date(a.CreatedDate)).slice(0, 5);

  // Finished-oil / cake-by-product quantities sold this month, per
  // product — from SaleItems on active (non-voided) sales dated this
  // month.
  const monthSaleIds = monthSales.map(s => s.SaleId);
  const monthSaleItems = sheetToObjects(SHEETS.SALE_ITEMS).filter(si => monthSaleIds.indexOf(si.SaleId) !== -1);
  const items = sheetToObjects(SHEETS.ITEMS);
  const qtyByItem = {};
  monthSaleItems.forEach(si => { qtyByItem[si.ItemId] = round2((qtyByItem[si.ItemId] || 0) + Number(si.Quantity)); });
  const productSummary = items
    .filter(i => i.Type === 'FinishedOil' || i.Type === 'CakeByProduct')
    .map(i => ({ itemId: i.ItemId, name: i.Name, type: i.Type, unit: i.Unit, quantitySold: qtyByItem[i.ItemId] || 0 }))
    .filter(p => p.quantitySold > 0);

  return {
    todaySalesTotal: round2(todaySales.reduce((s, x) => s + Number(x.GrandTotal), 0)),
    todaySalesCount: todaySales.length,
    todayPaymentsTotal: round2(todayPayments.reduce((s, x) => s + Number(x.Amount), 0)),
    monthSalesTotal: round2(monthSales.reduce((s, x) => s + Number(x.GrandTotal), 0)),
    pendingReceivable: pendingReceivable,
    creditsOwed: creditsOwed,
    currentStock: getCurrentStock(),
    recentSales: recentSales,
    recentPayments: recentPayments,
    productSummary: productSummary
  };
}

// ---------- UTIL ----------

function round2(n) { return Math.round((Number(n) + Number.EPSILON) * 100) / 100; }
function sameDate(a, b) { const da = new Date(a), db = new Date(b); return da.getFullYear() === db.getFullYear() && da.getMonth() === db.getMonth() && da.getDate() === db.getDate(); }

// ---------- WEB APP ENTRY POINTS ----------

function doGet(e) { return jsonOutput({ ok: true, message: 'Sri Products API is running. Use POST requests.' }); }

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return jsonOutput({ ok: false, error: 'Invalid request body.' }); }
  const action = body.action;
  try {
    let result, user = null;
    if (action !== 'login') user = requireSession(body.token);

    switch (action) {
      case 'login': result = login(body.username, body.password); break;
      case 'logout': result = logout(body.token); break;
      case 'getSettings': result = getSettings(); break;
      case 'updateSettings': result = updateSettings(user, body.updates); break;
      case 'listCustomers': result = listCustomers(body.search); break;
      case 'createCustomer': result = createCustomer(user, body.name); break;
      case 'getCustomerLedger': result = getCustomerLedger(body.customerId); break;
      case 'listItems': result = listItems(body.includeInactive); break;
      case 'createItem': result = createItem(user, body.item); break;
      case 'updateItem': result = updateItem(user, body.itemId, body.updates); break;
      case 'setItemActive': result = setItemActive(user, body.itemId, body.active); break;
      case 'listPrices': result = listPrices(body.itemId); break;
      case 'getCurrentPrices': result = getCurrentPrices(); break;
      case 'addPrice': result = addPrice(user, body.itemId, body.effectiveFrom, body.price); break;
      case 'updatePrice': result = updatePrice(user, body.priceId, body.effectiveFrom, body.price); break;
      case 'deletePrice': result = deletePrice(user, body.priceId); break;
      case 'listBottleAdjustments': result = listBottleAdjustments(body.itemId); break;
      case 'getCurrentBottleAdjustments': result = getCurrentBottleAdjustments(); break;
      case 'addBottleAdjustment': result = addBottleAdjustment(user, body.itemId, body.effectiveFrom, body.amount); break;
      case 'updateBottleAdjustment': result = updateBottleAdjustment(user, body.adjustmentId, body.effectiveFrom, body.amount); break;
      case 'deleteBottleAdjustment': result = deleteBottleAdjustment(user, body.adjustmentId); break;
      case 'getCurrentStock': result = getCurrentStock(); break;
      case 'getStockMovementHistory': result = getStockMovementHistory(body.itemId); break;
      case 'createStockAdjustment': result = createStockAdjustment(user, body.itemId, body.quantity, body.reason); break;
      case 'reverseStockAdjustment': result = reverseStockAdjustment(user, body.movementId, body.reason); break;
      case 'createSale': result = createSale(user, body.payload); break;
      case 'updateSale': result = updateSale(user, body.saleId, body.payload); break;
      case 'voidSale': result = voidSale(user, body.saleId, body.reason); break;
      case 'markSaleAsPaid': result = markSaleAsPaid(user, body.saleId, body.method); break;
      case 'listSales': result = listSales(body.filters); break;
      case 'getSaleDetail': result = getSaleDetail(body.saleId); break;
      case 'recordPayment': result = recordPayment(user, body.payload); break;
      case 'listPayments': result = listPayments(body.filters); break;
      case 'createQuotation': result = createQuotation(user, body.payload); break;
      case 'updateQuotation': result = updateQuotation(user, body.quotationId, body.payload); break;
      case 'deleteQuotation': result = deleteQuotation(user, body.quotationId); break;
      case 'listQuotations': result = listQuotations(body.filters); break;
      case 'getQuotationDetail': result = getQuotationDetail(body.quotationId); break;
      case 'updateQuotationStatus': result = updateQuotationStatus(user, body.quotationId, body.status); break;
      case 'convertQuotationToSale': result = convertQuotationToSale(user, body.quotationId, body.paymentAmount, body.paymentMethod); break;
      case 'createProduction': result = createProduction(user, body.payload); break;
      case 'updateProduction': result = updateProduction(user, body.productionId, body.payload); break;
      case 'deleteProduction': result = deleteProduction(user, body.productionId); break;
      case 'listProduction': result = listProduction(); break;
      case 'getProductionDetail': result = getProductionDetail(body.productionId); break;
      case 'getDashboard': result = getDashboard(); break;
      default: throw new Error('Unknown action: ' + action);
    }
    return jsonOutput({ ok: true, data: result });
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
