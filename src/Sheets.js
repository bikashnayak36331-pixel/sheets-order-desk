/**
 * Order Desk: small helpers for reading and writing the sheets.
 * Names ending in an underscore are private: they do not show up in the Apps Script "Run" menu.
 */

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" is missing. Run Order Desk > Set up / repair sheets.');
  return sh;
}

/** 1-based column number of an Orders column, by header name. */
function col_(name) {
  const i = ORDER_COLUMNS.indexOf(name);
  if (i === -1) throw new Error('Unknown Orders column: ' + name);
  return i + 1;
}

function tz_() {
  return ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone();
}

function formatDate_(date, pattern) {
  return Utilities.formatDate(date, tz_(), pattern || 'yyyy-MM-dd HH:mm');
}

/** Settings sheet (key in column A, value in column B) as a typed object. */
function getSettings_() {
  const values = sheet_(SHEETS.SETTINGS).getDataRange().getValues();
  const raw = {};
  values.slice(1).forEach(function (r) { if (r[0] !== '') raw[String(r[0]).trim()] = r[1]; });
  const get = function (key) {
    if (key in raw) return raw[key];
    const def = DEFAULT_SETTINGS.filter(function (d) { return d[0] === key; })[0];
    return def ? def[1] : '';
  };
  const yes = function (key) { return String(get(key)).trim().toLowerCase() === 'yes'; };
  return {
    businessName: String(get('Business name') || 'My Shop'),
    ownerEmail: String(get('Owner email') || ''),
    currency: String(get('Currency symbol')),
    taxRate: Number(get('Tax rate %')) || 0,
    idPrefix: String(get('Order ID prefix') || 'ORD-'),
    sendEmails: yes('Send customer emails'),
    invoiceFolder: String(get('Invoice folder') || 'Order Desk invoices'),
    invoiceWhenPaid: yes('Create invoice when paid'),
    summaryHour: Math.min(23, Math.max(0, parseInt(get('Daily summary hour (0-23)'), 10) || 7)),
    paymentInstructions: String(get('Payment instructions') || ''),
  };
}

/** Products sheet as a SKU index (see productIndex in Core.js). */
function getProducts_() {
  const sh = sheet_(SHEETS.PRODUCTS);
  const last = sh.getLastRow();
  if (last < 2) return {};
  return productIndex(sh.getRange(2, 1, last - 1, PRODUCT_COLUMNS.length).getValues());
}

/** Write new stock numbers back to the Products sheet. */
function saveStock_(products) {
  const sh = sheet_(SHEETS.PRODUCTS);
  Object.keys(products).forEach(function (sku) {
    const p = products[sku];
    sh.getRange(p.row + 2, 4).setValue(p.stock);
  });
}

/** One Orders row as an object; `row` is the sheet row number. */
function rowToOrder_(values, row, settings) {
  const o = {};
  ORDER_COLUMNS.forEach(function (name, i) { o[name] = values[i]; });
  const created = values[col_('Created') - 1];
  return {
    row: row,
    id: String(o['Order ID']),
    created: created,
    createdText: created instanceof Date ? formatDate_(created, 'yyyy-MM-dd') : String(created),
    customer: String(o['Customer'] || ''),
    email: String(o['Email'] || '').trim(),
    phone: String(o['Phone'] || ''),
    items: String(o['Items'] || ''),
    discountPercent: o['Discount %'] === '' ? 0 : Number(o['Discount %']) || 0,
    total: Number(o['Total']) || 0,
    totalText: formatMoney(toCents(Number(o['Total']) || 0), settings ? settings.currency : '$'),
    status: String(o['Status'] || 'New'),
    stock: String(o['Stock'] || ''),
    notes: String(o['Notes'] || ''),
  };
}

function readOrder_(row, settings) {
  const values = sheet_(SHEETS.ORDERS).getRange(row, 1, 1, ORDER_COLUMNS.length).getValues()[0];
  return rowToOrder_(values, row, settings);
}

function readAllOrders_(settings) {
  const sh = sheet_(SHEETS.ORDERS);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ORDER_COLUMNS.length).getValues()
    .filter(function (r) { return r[0] !== ''; })
    .map(function (r, i) { return rowToOrder_(r, i + 2, settings); });
}

/** Set several Orders cells in one row by column name. */
function setOrderCells_(row, changes) {
  const sh = sheet_(SHEETS.ORDERS);
  Object.keys(changes).forEach(function (name) {
    sh.getRange(row, col_(name)).setValue(changes[name]);
  });
  sh.getRange(row, col_('Last update')).setValue(new Date());
}

/** Append a line to the Log sheet (newest at the bottom, trimmed to the last 2,000 entries). */
function log_(event, orderId, details) {
  const sh = ss_().getSheetByName(SHEETS.LOG);
  if (!sh) return;
  sh.appendRow([new Date(), event, orderId || '', details || '']);
  const extra = sh.getLastRow() - 1 - 2000;
  if (extra > 0) sh.deleteRows(2, extra);
}
