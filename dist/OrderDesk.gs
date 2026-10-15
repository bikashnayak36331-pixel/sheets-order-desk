/**
 * Order Desk v1.0.0: order tracking, stock, PDF invoices and emails for Google Sheets.
 * https://github.com/bikashnayak36331-pixel/sheets-order-desk  (MIT License)
 *
 * Install: in your spreadsheet open Extensions > Apps Script, replace the contents of Code.gs
 * with this whole file, save, reload the spreadsheet, then use the "Order Desk" menu.
 * Built from src/ by scripts/build.mjs. Edit the files in src/, not this one.
 */

// ----- Config.js -----
/**
 * @OnlyCurrentDoc  (limits spreadsheet access to the sheet this script is attached to)
 *
 * Order Desk: configuration shared by every file.
 * Sheet names and column layouts live here so the rest of the code never hard-codes positions.
 */

const SHEETS = {
  ORDERS: 'Orders',
  PRODUCTS: 'Products',
  SETTINGS: 'Settings',
  DASHBOARD: 'Dashboard',
  LOG: 'Log',
};

const ORDER_COLUMNS = [
  'Order ID', 'Created', 'Customer', 'Email', 'Phone', 'Items', 'Discount %', 'Subtotal', 'Discount',
  'Tax', 'Total', 'Status', 'Stock', 'Invoice', 'Notes', 'Last update',
];

/** Values of the Stock column: whether this order's items have been taken out of stock. */
const STOCK_STATE = { RESERVED: 'Reserved', NOT_RESERVED: 'Not reserved', RETURNED: 'Returned' };

const PRODUCT_COLUMNS = ['SKU', 'Name', 'Price', 'Stock', 'Reorder at', 'Active'];

const LOG_COLUMNS = ['Time', 'Event', 'Order', 'Details'];

/** Allowed status changes. Anything else typed into the Status column is undone. */
const STATUS_FLOW = {
  'New': ['Confirmed', 'Paid', 'On hold', 'Cancelled'],
  'On hold': ['Confirmed', 'Paid', 'Cancelled'],
  'Confirmed': ['Paid', 'Cancelled'],
  'Paid': ['Shipped', 'Cancelled'],
  'Shipped': ['Delivered'],
  'Delivered': [],
  'Cancelled': [],
};

const STATUSES = Object.keys(STATUS_FLOW);

/** Statuses that should send the customer an email (when emails are switched on). */
const EMAIL_ON_STATUS = ['New', 'On hold', 'Confirmed', 'Paid', 'Shipped', 'Cancelled'];

/** Settings rows created by Setup, with their default values. */
const DEFAULT_SETTINGS = [
  ['Business name', 'My Shop'],
  ['Owner email', ''],
  ['Currency symbol', '$'],
  ['Tax rate %', 8],
  ['Order ID prefix', 'ORD-'],
  ['Send customer emails', 'No'],
  ['Invoice folder', 'Order Desk invoices'],
  ['Create invoice when paid', 'Yes'],
  ['Daily summary hour (0-23)', 7],
  ['Payment instructions', 'Pay by bank transfer or card. Reply to this email with any questions.'],
];

const SAMPLE_PRODUCTS = [
  ['CAKE-CHOC', 'Chocolate fudge cake (8 in)', 32, 12, 3, 'Yes'],
  ['CAKE-LEMON', 'Lemon drizzle cake (8 in)', 28, 10, 3, 'Yes'],
  ['CUPC-12', 'Box of 12 cupcakes', 30, 20, 5, 'Yes'],
  ['COOKIE-24', 'Cookie tin (24 pieces)', 22, 15, 5, 'Yes'],
  ['BREAD-SOUR', 'Sourdough loaf', 8.5, 30, 8, 'Yes'],
];

// ----- Core.js -----
/**
 * Order Desk: pure business logic.
 *
 * Nothing in this file touches Google services, so every function can be unit-tested in Node.
 * Money is handled in whole cents to avoid floating-point surprises.
 */

/** Convert a price such as 8.5 or "8.50" to whole cents (850). */
function toCents(value) {
  if (typeof value === 'number') {
    if (!isFinite(value)) throw new Error('Not a valid amount: ' + value);
    return Math.round(value * 100);
  }
  const text = String(value === null || value === undefined ? '' : value).trim();
  if (text === '') return 0;
  const cleaned = text.replace(/[^0-9.\-]/g, '');
  const n = Number(cleaned);
  if (!/\d/.test(cleaned) || !isFinite(n)) throw new Error('Not a valid amount: ' + value);
  return Math.round(n * 100);
}

/** Format cents for people: 123450 -> "$1,234.50". */
function formatMoney(cents, symbol) {
  const sign = cents < 0 ? '-' : '';
  const abs = Math.abs(Math.round(cents));
  const whole = Math.floor(abs / 100).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const frac = String(abs % 100).padStart(2, '0');
  return sign + (symbol === undefined ? '$' : symbol) + whole + '.' + frac;
}

/** A percentage such as 8.25 as basis points (825), so tax maths stays in integers. */
function percentToBp(value) {
  const n = Number(String(value === '' || value === null || value === undefined ? 0 : value).replace('%', ''));
  if (!isFinite(n) || n < 0 || n > 100) throw new Error('A percentage must be between 0 and 100, got ' + value);
  return Math.round(n * 100);
}

/**
 * Read an items cell such as "CAKE-CHOC x2, CUPC-12" into [{sku, qty}].
 * Separators: comma, semicolon or new line. Quantity defaults to 1.
 */
function parseItems(text) {
  const parts = String(text || '').split(/[,;\n]+/).map(function (p) { return p.trim(); }).filter(String);
  if (!parts.length) throw new Error('The order has no items. Use a format like "CAKE-CHOC x2, CUPC-12 x1".');
  const merged = {};
  const order = [];
  parts.forEach(function (part) {
    const m = part.match(/^([A-Za-z0-9][A-Za-z0-9_\-]*)\s*(?:[x×*]\s*(\d+))?$/);
    if (!m) throw new Error('Cannot read item "' + part + '". Use SKU x quantity, e.g. CAKE-CHOC x2.');
    const sku = m[1].toUpperCase();
    const qty = m[2] === undefined ? 1 : Number(m[2]);
    if (!(qty >= 1 && qty <= 999)) throw new Error('Quantity for ' + sku + ' must be between 1 and 999.');
    if (!(sku in merged)) { merged[sku] = 0; order.push(sku); }
    merged[sku] += qty;
  });
  return order.map(function (sku) { return { sku: sku, qty: merged[sku] }; });
}

function formatItems(items) {
  return items.map(function (i) { return i.sku + ' x' + i.qty; }).join(', ');
}

/** Build a SKU -> product lookup from Products sheet rows (SKU, Name, Price, Stock, Reorder at, Active). */
function productIndex(rows) {
  const index = {};
  rows.forEach(function (r, i) {
    const sku = String(r[0] || '').trim().toUpperCase();
    if (!sku) return;
    index[sku] = {
      sku: sku,
      name: String(r[1] || sku),
      priceCents: toCents(r[2] || 0),
      stock: Number(r[3]) || 0,
      reorderAt: Number(r[4]) || 0,
      active: String(r[5] || 'Yes').toLowerCase() !== 'no',
      row: i,
    };
  });
  return index;
}

/**
 * Price an order. Returns lines plus subtotal, discount, tax and total in cents.
 * Unknown or inactive SKUs throw an error that names the SKU.
 */
function priceOrder(items, products, options) {
  const taxBp = percentToBp(options && options.taxRate);
  const discountBp = percentToBp(options && options.discountPercent);
  const lines = items.map(function (item) {
    const p = products[item.sku];
    if (!p) throw new Error('Unknown product SKU: ' + item.sku);
    if (!p.active) throw new Error('Product ' + item.sku + ' is not active.');
    return { sku: p.sku, name: p.name, qty: item.qty, unitCents: p.priceCents, lineCents: p.priceCents * item.qty };
  });
  const subtotal = lines.reduce(function (s, l) { return s + l.lineCents; }, 0);
  const discount = Math.round(subtotal * discountBp / 10000);
  const tax = Math.round((subtotal - discount) * taxBp / 10000);
  return { lines: lines, subtotal: subtotal, discount: discount, tax: tax, total: subtotal - discount + tax };
}

/** Items that cannot be fulfilled from current stock. */
function findShortages(items, products) {
  return items
    .filter(function (i) { return products[i.sku] && products[i.sku].stock < i.qty; })
    .map(function (i) { return { sku: i.sku, wanted: i.qty, inStock: products[i.sku].stock }; });
}

/** ORD-0001, ORD-0002, ... continuing from the highest existing number with the same prefix. */
function nextOrderId(existingIds, prefix) {
  prefix = prefix || 'ORD-';
  let max = 0;
  existingIds.forEach(function (id) {
    const s = String(id || '');
    if (s.indexOf(prefix) === 0) {
      const n = parseInt(s.slice(prefix.length), 10);
      if (n > max) max = n;
    }
  });
  return prefix + String(max + 1).padStart(4, '0');
}

function canTransition(from, to) {
  if (from === to) return true;
  const allowed = STATUS_FLOW[from];
  return Array.isArray(allowed) && allowed.indexOf(to) !== -1;
}

function lowStockList(products) {
  return Object.keys(products)
    .map(function (k) { return products[k]; })
    .filter(function (p) { return p.active && p.stock <= p.reorderAt; })
    .map(function (p) { return { sku: p.sku, name: p.name, stock: p.stock, reorderAt: p.reorderAt }; });
}

/** Counts and revenue for orders created in [start, end). Cancelled orders are not revenue. */
function summarizeOrders(orders, start, end) {
  const byStatus = {};
  STATUSES.forEach(function (s) { byStatus[s] = 0; });
  let count = 0;
  let revenue = 0;
  let open = 0;
  orders.forEach(function (o) {
    const created = o.created instanceof Date ? o.created : new Date(o.created);
    if (['New', 'On hold', 'Confirmed', 'Paid'].indexOf(o.status) !== -1) open += 1;
    if (created >= start && created < end) {
      count += 1;
      byStatus[o.status] = (byStatus[o.status] || 0) + 1;
      if (o.status !== 'Cancelled') revenue += toCents(o.total);
    }
  });
  return { count: count, revenueCents: revenue, byStatus: byStatus, openOrders: open };
}

function escapeHtml(text) {
  return String(text === undefined || text === null ? '' : text)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** The invoice as a self-contained HTML page (Apps Script turns it into a PDF). */
function buildInvoiceHtml(order, priced, settings) {
  const cur = settings.currency;
  const paid = order.status === 'Paid' || order.status === 'Shipped' || order.status === 'Delivered';
  const rows = priced.lines.map(function (l) {
    return '<tr><td>' + escapeHtml(l.name) + '<div class="sku">' + escapeHtml(l.sku) + '</div></td>' +
      '<td class="r">' + l.qty + '</td><td class="r">' + formatMoney(l.unitCents, cur) + '</td>' +
      '<td class="r">' + formatMoney(l.lineCents, cur) + '</td></tr>';
  }).join('');
  const discountRow = priced.discount
    ? '<tr><td>Discount</td><td class="r">-' + formatMoney(priced.discount, cur) + '</td></tr>' : '';
  return '<!doctype html><html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#1f2933;margin:40px;font-size:13px}' +
    'table.top{width:100%;border-collapse:collapse}table.top td{vertical-align:top;padding:0}' +
    'h1{color:#0b6e4f;margin:0;font-size:28px;letter-spacing:1px}.biz{font-size:18px;font-weight:bold}' +
    '.muted{color:#616e7c}.badge{display:inline-block;margin-top:6px;padding:3px 10px;border-radius:12px;' +
    'background:#0b6e4f;color:#fff;font-weight:bold;font-size:11px}' +
    'table.items{width:100%;border-collapse:collapse;margin-top:28px}' +
    'table.items th{background:#1f2933;color:#fff;text-align:left;padding:8px;font-size:11px}' +
    'table.items td{padding:8px;border-bottom:1px solid #d9e2ec}.r,table.items th.r{text-align:right}' +
    '.sku{color:#616e7c;font-size:10px}table.totals{margin-left:auto;margin-top:12px}' +
    'table.totals td{padding:4px 8px}.grand td{font-weight:bold;font-size:15px;border-top:2px solid #1f2933}' +
    '.box{margin-top:28px}</style></head><body>' +
    '<table class="top"><tr><td><div class="biz">' + escapeHtml(settings.businessName) + '</div>' +
    '<div class="muted">' + escapeHtml(settings.ownerEmail) + '</div></td>' +
    '<td style="text-align:right"><h1>INVOICE</h1><div>' + escapeHtml(order.id) + '</div>' +
    '<div class="muted">' + escapeHtml(order.createdText) + '</div>' +
    (paid ? '<div class="badge">PAID</div>' : '') + '</td></tr></table>' +
    '<div class="box"><div class="muted">BILL TO</div><div><b>' + escapeHtml(order.customer) + '</b></div>' +
    '<div class="muted">' + escapeHtml(order.email) + (order.phone ? ' · ' + escapeHtml(order.phone) : '') +
    '</div></div>' +
    '<table class="items"><tr><th>ITEM</th><th class="r">QTY</th><th class="r">PRICE</th>' +
    '<th class="r">AMOUNT</th></tr>' + rows + '</table>' +
    '<table class="totals"><tr><td>Subtotal</td><td class="r">' + formatMoney(priced.subtotal, cur) + '</td></tr>' +
    discountRow +
    '<tr><td>Tax (' + settings.taxRate + '%)</td><td class="r">' + formatMoney(priced.tax, cur) + '</td></tr>' +
    '<tr class="grand"><td>Total</td><td class="r">' + formatMoney(priced.total, cur) + '</td></tr></table>' +
    (order.notes ? '<div class="box"><div class="muted">NOTES</div>' + escapeHtml(order.notes) + '</div>' : '') +
    (!paid && settings.paymentInstructions
      ? '<div class="box"><div class="muted">HOW TO PAY</div>' + escapeHtml(settings.paymentInstructions) + '</div>'
      : '') +
    '</body></html>';
}

/** Subject and body of the email a customer gets when their order changes status. */
function buildStatusEmail(order, status, settings) {
  const messages = {
    'New': 'Thank you for your order! We have received it and will confirm it shortly.',
    'On hold': 'Thank you for your order! One or more items are low on stock, so we will be in touch shortly.',
    'Confirmed': 'Thank you! Your order is confirmed and we are getting it ready.',
    'Paid': 'We have received your payment. Thank you! Your invoice is attached when available.',
    'Shipped': 'Good news: your order is on its way.',
    'Cancelled': 'Your order has been cancelled. If this is unexpected, please reply to this email.',
  };
  const first = String(order.customer || 'there').split(' ')[0];
  const subject = settings.businessName + ': order ' + order.id + ' ' + status.toLowerCase();
  const body = 'Hi ' + first + ',\n\n' + (messages[status] || 'Your order status is now: ' + status + '.') +
    '\n\nOrder: ' + order.id + '\nItems: ' + order.items + '\nTotal: ' + order.totalText +
    '\n\nThank you,\n' + settings.businessName;
  return { subject: subject, body: body };
}

/** The owner's daily summary email. */
function buildSummaryEmail(summary, lowStock, settings, dayLabel) {
  const cur = settings.currency;
  const statusText = STATUSES.filter(function (s) { return summary.byStatus[s]; })
    .map(function (s) { return s + ' ' + summary.byStatus[s]; }).join(', ');
  const lines = [
    'Order Desk summary for ' + dayLabel, '',
    'New orders: ' + summary.count,
    'Revenue (excluding cancelled): ' + formatMoney(summary.revenueCents, cur),
    'Open orders right now: ' + summary.openOrders, '',
    'By status: ' + (statusText || 'none'),
  ];
  if (lowStock.length) {
    lines.push('', 'Low stock:');
    lowStock.forEach(function (p) { lines.push('- ' + p.name + ' (' + p.sku + '): ' + p.stock + ' left'); });
  }
  return { subject: settings.businessName + ' daily summary: ' + summary.count + ' new orders', body: lines.join('\n') };
}

// ----- Sheets.js -----
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

// ----- Setup.js -----
/**
 * Order Desk: menu and one-click setup.
 * Setup is safe to run again: it only adds what is missing and never deletes your data.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Order Desk')
    .addItem('Set up / repair sheets', 'setupOrderDesk')
    .addItem('Install automations', 'installTriggers')
    .addSeparator()
    .addItem('Add a test order', 'addTestOrder')
    .addItem('Create invoice for selected order', 'createInvoiceForSelection')
    .addItem('Refresh dashboard', 'refreshDashboard')
    .addItem('Send daily summary now', 'sendDailySummary')
    .addToUi();
}

function setupOrderDesk() {
  const ss = ss_();
  const symbol = '$';

  // Settings first, so the currency symbol can be read for formats below.
  const settings = ensureSheet_(ss, SHEETS.SETTINGS, ['Setting', 'Value']);
  const existing = settings.getDataRange().getValues().map(function (r) { return String(r[0]); });
  DEFAULT_SETTINGS.forEach(function (pair) {
    if (existing.indexOf(pair[0]) === -1) settings.appendRow(pair);
  });
  settings.setColumnWidth(1, 220);
  settings.setColumnWidth(2, 360);
  const yesNo = SpreadsheetApp.newDataValidation().requireValueInList(['Yes', 'No'], true).build();
  settings.getDataRange().getValues().forEach(function (r, i) {
    if (r[0] === 'Send customer emails' || r[0] === 'Create invoice when paid') {
      settings.getRange(i + 1, 2).setDataValidation(yesNo);
    }
  });
  const cur = getSettings_().currency;
  const moneyFormat = '"' + (cur || symbol) + '"#,##0.00';

  const orders = ensureSheet_(ss, SHEETS.ORDERS, ORDER_COLUMNS);
  const widths = { 'Order ID': 90, 'Created': 130, 'Customer': 160, 'Email': 200, 'Items': 240, 'Status': 100,
    'Invoice': 90, 'Notes': 220, 'Last update': 130 };
  ORDER_COLUMNS.forEach(function (name, i) { orders.setColumnWidth(i + 1, widths[name] || 85); });
  const rows = orders.getMaxRows() - 1;
  orders.getRange(2, col_('Created'), rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  orders.getRange(2, col_('Last update'), rows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  ['Subtotal', 'Discount', 'Tax', 'Total'].forEach(function (name) {
    orders.getRange(2, col_(name), rows, 1).setNumberFormat(moneyFormat);
  });
  orders.getRange(2, col_('Status'), rows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(STATUSES, true).setAllowInvalid(false)
      .setHelpText('Pick a status. Order Desk only allows the next steps in the order flow.').build());
  orders.setConditionalFormatRules(statusColours_(orders.getRange(2, col_('Status'), rows, 1)));

  const products = ensureSheet_(ss, SHEETS.PRODUCTS, PRODUCT_COLUMNS);
  if (products.getLastRow() < 2) {
    products.getRange(2, 1, SAMPLE_PRODUCTS.length, PRODUCT_COLUMNS.length).setValues(SAMPLE_PRODUCTS);
  }
  products.setColumnWidth(2, 240);
  products.getRange(2, 3, products.getMaxRows() - 1, 1).setNumberFormat(moneyFormat);
  products.getRange(2, 6, products.getMaxRows() - 1, 1).setDataValidation(yesNo);

  ensureSheet_(ss, SHEETS.LOG, LOG_COLUMNS).setColumnWidth(4, 420);
  ensureSheet_(ss, SHEETS.DASHBOARD, ['Order Desk dashboard', '']);
  refreshDashboard();
  log_('Setup', '', 'Sheets checked and repaired');
  ss.toast('Order Desk is ready. Next: Order Desk > Install automations.', 'Order Desk', 8);
}

/** Get a sheet, creating it with a bold, frozen header row if it does not exist. */
function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  if (current.join('') === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1f2933').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}

function statusColours_(range) {
  const colours = {
    'New': '#e3f2fd', 'On hold': '#fff3cd', 'Confirmed': '#e8eaf6', 'Paid': '#d1f2e3',
    'Shipped': '#e0f7fa', 'Delivered': '#c8e6c9', 'Cancelled': '#f8d7da',
  };
  return Object.keys(colours).map(function (status) {
    return SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(status)
      .setBackground(colours[status]).setRanges([range]).build();
  });
}

// ----- Orders.js -----
/**
 * Order Desk: creating orders (from a Google Form or the menu) and reacting to status changes.
 */

/**
 * Installable "On form submit" trigger. Works with a Google Form linked to this spreadsheet.
 * Recognised questions (case does not matter): Name, Email, Phone, Product, Quantity, Items, Notes.
 * Use either Product + Quantity (one product) or Items ("CAKE-CHOC x2, CUPC-12").
 */
function onFormSubmit(e) {
  const answers = {};
  const named = (e && e.namedValues) || {};
  Object.keys(named).forEach(function (q) {
    answers[q.trim().toLowerCase()] = String((named[q] || [''])[0]).trim();
  });
  const pick = function () {
    for (let i = 0; i < arguments.length; i++) if (answers[arguments[i]]) return answers[arguments[i]];
    return '';
  };
  let items = pick('items', 'order', 'order items');
  if (!items) {
    const product = pick('product', 'item');
    const qty = pick('quantity', 'qty') || '1';
    items = product ? productToSku_(product) + ' x' + qty : '';
  }
  return createOrder_({
    customer: pick('name', 'your name', 'full name', 'customer'),
    email: pick('email', 'email address', 'your email'),
    phone: pick('phone', 'phone number', 'mobile'),
    items: items,
    notes: pick('notes', 'comments', 'message', 'delivery notes'),
    source: 'Form',
  });
}

/** Accept a SKU, a product name, or "Name (SKU)" from a form dropdown. */
function productToSku_(text) {
  const products = getProducts_();
  const t = String(text).trim();
  const inBrackets = t.match(/\(([A-Za-z0-9_\-]+)\)\s*$/);
  if (inBrackets && products[inBrackets[1].toUpperCase()]) return inBrackets[1].toUpperCase();
  if (products[t.toUpperCase()]) return t.toUpperCase();
  const byName = Object.keys(products).filter(function (sku) {
    return products[sku].name.toLowerCase() === t.toLowerCase();
  });
  if (byName.length) return byName[0];
  return t.replace(/\s+/g, '-').toUpperCase(); // unknown: priceOrder will report it clearly
}

/**
 * Create one order row. If an item is unknown, the order is still saved (status On hold) with the
 * problem in Notes, so no customer submission is ever lost.
 */
function createOrder_(input) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // two form submissions at once must not get the same order ID
  try {
    const settings = getSettings_();
    const products = getProducts_();
    const sh = sheet_(SHEETS.ORDERS);
    const last = sh.getLastRow();
    const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return r[0]; }) : [];
    const id = nextOrderId(ids, settings.idPrefix);
    const now = new Date();

    let status = 'New';
    let stock = STOCK_STATE.NOT_RESERVED;
    let notes = input.notes || '';
    let priced = { subtotal: 0, discount: 0, tax: 0, total: 0 };
    let itemsText = String(input.items || '');
    const problem = function (msg) {
      status = 'On hold';
      notes = (notes ? notes + ' | ' : '') + msg;
    };
    try {
      const items = parseItems(itemsText);
      itemsText = formatItems(items);
      priced = priceOrder(items, products, { taxRate: settings.taxRate, discountPercent: input.discountPercent });
      const short = findShortages(items, products);
      if (short.length) {
        problem('Not enough stock: ' + short.map(function (s) {
          return s.sku + ' wanted ' + s.wanted + ', in stock ' + s.inStock;
        }).join('; '));
      } else {
        reserveStock_(items, products);
        stock = STOCK_STATE.RESERVED;
      }
    } catch (err) {
      problem(err.message);
    }
    if (!input.email) problem('No email address');

    const row = ORDER_COLUMNS.map(function (name) {
      switch (name) {
        case 'Order ID': return id;
        case 'Created': return now;
        case 'Customer': return input.customer || '';
        case 'Email': return input.email || '';
        case 'Phone': return input.phone || '';
        case 'Items': return itemsText;
        case 'Discount %': return input.discountPercent || '';
        case 'Subtotal': return priced.subtotal / 100;
        case 'Discount': return priced.discount / 100;
        case 'Tax': return priced.tax / 100;
        case 'Total': return priced.total / 100;
        case 'Status': return status;
        case 'Stock': return stock;
        case 'Invoice': return '';
        case 'Notes': return notes;
        case 'Last update': return now;
        default: return '';
      }
    });
    sh.appendRow(row);
    const rowNumber = sh.getLastRow();
    log_('Order created (' + (input.source || 'Manual') + ')', id, status + ': ' + itemsText);
    const order = readOrder_(rowNumber, settings);
    notifyCustomer_(order, status, settings, null);
    refreshDashboard();
    return order;
  } finally {
    lock.releaseLock();
  }
}

/** Menu: add a realistic test order so you can see everything working. */
function addTestOrder() {
  const products = getProducts_();
  const skus = Object.keys(products).filter(function (s) { return products[s].active; });
  if (!skus.length) throw new Error('Add at least one active product first.');
  const a = skus[Math.floor(Math.random() * skus.length)];
  const b = skus[Math.floor(Math.random() * skus.length)];
  const order = createOrder_({
    customer: 'Test Customer',
    email: getSettings_().ownerEmail || Session.getEffectiveUser().getEmail(),
    items: a + ' x1, ' + b + ' x1',
    notes: 'Test order from the Order Desk menu',
    source: 'Test',
  });
  ss_().toast('Added ' + order.id + ' (' + order.totalText + ').', 'Order Desk', 5);
  return order;
}

/**
 * Installable "On edit" trigger (a simple onEdit cannot send email or create files).
 * Handles changes to Status and Items in the Orders sheet.
 */
function handleEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sh = range.getSheet();
  if (sh.getName() !== SHEETS.ORDERS || range.getRow() < 2) return;
  if (range.getNumRows() !== 1 || range.getNumColumns() !== 1) return;
  const column = range.getColumn();
  if (column === col_('Status')) return changeStatus_(range.getRow(), e.oldValue, e.value, range);
  if (column === col_('Items')) return repriceOrder_(range.getRow(), e.oldValue);
  if (column === col_('Discount %')) return repriceOrder_(range.getRow());
}

/** Apply a status change typed into the sheet, or undo it if it skips a step. */
function changeStatus_(row, oldValue, newValue, range) {
  const from = oldValue === undefined || oldValue === '' ? 'New' : String(oldValue);
  const to = String(newValue === undefined ? '' : newValue);
  const id = String(sheet_(SHEETS.ORDERS).getRange(row, col_('Order ID')).getValue());
  if (!to || !canTransition(from, to)) {
    range.setValue(from);
    const next = (STATUS_FLOW[from] || []).join(', ') || 'none (final status)';
    ss_().toast('An order that is ' + from + ' can move to: ' + next + '.', 'Status not changed', 8);
    log_('Status change refused', id, from + ' -> ' + (to || '(blank)'));
    return false;
  }
  if (from === to) return true;

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const settings = getSettings_();
    let order = readOrder_(row, settings);
    const products = getProducts_();

    // Stock: take it when an on-hold order goes ahead; give it back when an order is cancelled.
    if (to !== 'Cancelled' && order.stock !== STOCK_STATE.RESERVED) {
      const items = parseItems(order.items);
      const short = findShortages(items, products);
      if (short.length) {
        range.setValue(from);
        ss_().toast('Not enough stock for ' + short.map(function (s) { return s.sku; }).join(', ') +
          '. Update the Products sheet first.', 'Status not changed', 8);
        log_('Status change refused', id, 'Not enough stock for ' + to);
        return false;
      }
      reserveStock_(items, products);
      setOrderCells_(row, { 'Stock': STOCK_STATE.RESERVED });
    }
    if (to === 'Cancelled' && order.stock === STOCK_STATE.RESERVED) {
      returnStock_(parseItems(order.items), products);
      setOrderCells_(row, { 'Stock': STOCK_STATE.RETURNED });
    }
    setOrderCells_(row, {});
    log_('Status ' + from + ' -> ' + to, id, '');

    let invoice = null;
    if (to === 'Paid' && settings.invoiceWhenPaid) invoice = createInvoice_(row);
    order = readOrder_(row, settings);
    notifyCustomer_(order, to, settings, invoice);
    refreshDashboard();
    return true;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Recalculate totals after Items or Discount % is edited (only while the order is New or On hold).
 * If stock was already reserved, the old items go back to stock and the new ones are reserved.
 */
function repriceOrder_(row, oldItemsText) {
  const settings = getSettings_();
  const order = readOrder_(row, settings);
  const sh = sheet_(SHEETS.ORDERS);
  const undo = function (message) {
    if (oldItemsText !== undefined) sh.getRange(row, col_('Items')).setValue(oldItemsText);
    ss_().toast(message, 'Order not changed', 8);
    log_('Edit refused', order.id, message);
    return false;
  };
  if (['New', 'On hold'].indexOf(order.status) === -1) {
    return undo('Items of a ' + order.status + ' order are locked.');
  }
  let items;
  let priced;
  const products = getProducts_();
  try {
    items = parseItems(order.items);
    priced = priceOrder(items, products, { taxRate: settings.taxRate, discountPercent: order.discountPercent });
  } catch (err) {
    return undo(err.message);
  }
  if (order.stock === STOCK_STATE.RESERVED && oldItemsText !== undefined) {
    const oldItems = parseItems(oldItemsText);
    returnStock_(oldItems, products);
    const short = findShortages(items, products);
    if (short.length) {
      reserveStock_(oldItems, products); // put the original reservation back
      return undo('Not enough stock for ' + short.map(function (s) { return s.sku; }).join(', ') + '.');
    }
    reserveStock_(items, products);
  }
  setOrderCells_(row, {
    'Items': formatItems(items), 'Subtotal': priced.subtotal / 100, 'Discount': priced.discount / 100,
    'Tax': priced.tax / 100, 'Total': priced.total / 100,
  });
  log_('Order repriced', order.id, formatItems(items) + ' = ' + formatMoney(priced.total, settings.currency));
  refreshDashboard();
  return true;
}

function reserveStock_(items, products) {
  items.forEach(function (i) { products[i.sku].stock -= i.qty; });
  saveStock_(products);
}

function returnStock_(items, products) {
  items.forEach(function (i) { if (products[i.sku]) products[i.sku].stock += i.qty; });
  saveStock_(products);
}

// ----- Documents.js -----
/**
 * Order Desk: PDF invoices in Google Drive and customer emails.
 */

/** Build the invoice PDF for an Orders row, save it to Drive and link it in the Invoice column. */
function createInvoice_(row) {
  const settings = getSettings_();
  const order = readOrder_(row, settings);
  const priced = priceOrder(parseItems(order.items), getProducts_(),
    { taxRate: settings.taxRate, discountPercent: order.discountPercent });
  const html = buildInvoiceHtml(order, priced, settings);
  const pdf = HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf')
    .setName('Invoice ' + order.id + '.pdf');
  const file = invoiceFolder_(settings.invoiceFolder).createFile(pdf);
  sheet_(SHEETS.ORDERS).getRange(row, col_('Invoice'))
    .setFormula('=HYPERLINK("' + file.getUrl() + '","Invoice PDF")');
  log_('Invoice created', order.id, file.getUrl());
  return file;
}

/** Menu: create (or re-create) the invoice for the selected Orders row. */
function createInvoiceForSelection() {
  const range = ss_().getActiveRange();
  if (!range || range.getSheet().getName() !== SHEETS.ORDERS || range.getRow() < 2) {
    throw new Error('Select a cell in the order\'s row on the Orders sheet first.');
  }
  const file = createInvoice_(range.getRow());
  ss_().toast('Saved ' + file.getName() + ' to Drive.', 'Order Desk', 5);
  return file;
}

function invoiceFolder_(name) {
  const found = DriveApp.getFoldersByName(name);
  return found.hasNext() ? found.next() : DriveApp.createFolder(name);
}

/**
 * Email the customer about a status change, if emails are switched on in Settings.
 * With emails off (the default), the email is written to the Log instead, so you can check it first.
 */
function notifyCustomer_(order, status, settings, invoiceFile) {
  if (EMAIL_ON_STATUS.indexOf(status) === -1) return false;
  const email = buildStatusEmail(order, status, settings);
  if (!order.email) {
    log_('Email skipped', order.id, 'No email address');
    return false;
  }
  if (!settings.sendEmails) {
    log_('Email preview (sending is off)', order.id, 'To ' + order.email + ' | ' + email.subject);
    return false;
  }
  if (MailApp.getRemainingDailyQuota() < 1) {
    log_('Email not sent', order.id, 'Daily Gmail quota used up');
    return false;
  }
  const message = { to: order.email, subject: email.subject, body: email.body, name: settings.businessName };
  if (settings.ownerEmail) message.replyTo = settings.ownerEmail;
  if (invoiceFile) message.attachments = [invoiceFile.getAs('application/pdf')];
  MailApp.sendEmail(message);
  log_('Email sent', order.id, 'To ' + order.email + ' | ' + email.subject);
  return true;
}

// ----- Reports.js -----
/**
 * Order Desk: dashboard, daily summary email and triggers.
 */

/** Rewrite the Dashboard sheet with today's numbers. Runs after every order change and hourly. */
function refreshDashboard() {
  const settings = getSettings_();
  const orders = readAllOrders_(settings);
  const products = getProducts_();
  const now = new Date();
  const startOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
  const tomorrow = new Date(startOfDay.getTime() + 24 * 3600 * 1000);
  const today = summarizeOrders(orders, startOfDay, tomorrow);
  const month = summarizeOrders(orders, startOfMonth, tomorrow);
  const low = lowStockList(products);
  const cur = settings.currency;

  const rows = [
    ['Order Desk dashboard', 'Updated ' + formatDate_(now)],
    ['', ''],
    ['Orders today', today.count],
    ['Revenue today', formatMoney(today.revenueCents, cur)],
    ['Orders this month', month.count],
    ['Revenue this month', formatMoney(month.revenueCents, cur)],
    ['Open orders (New, On hold, Confirmed, Paid)', today.openOrders],
    ['Products at or below reorder level', low.length],
    ['', ''],
    ['Status (this month)', 'Orders'],
  ];
  STATUSES.forEach(function (s) { rows.push([s, month.byStatus[s] || 0]); });
  rows.push(['', ''], ['Low stock', 'In stock']);
  if (low.length) low.forEach(function (p) { rows.push([p.name + ' (' + p.sku + ')', p.stock]); });
  else rows.push(['Nothing to reorder', '']);

  const sh = sheet_(SHEETS.DASHBOARD);
  sh.clearContents();
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange(1, 1, 1, 2).setFontWeight('bold').setFontSize(14);
  sh.getRange(10, 1, 1, 2).setFontWeight('bold');
  sh.getRange(10 + STATUSES.length + 2, 1, 1, 2).setFontWeight('bold');
  sh.setColumnWidth(1, 320);
  sh.setColumnWidth(2, 160);
  if (sh.getCharts().length === 0) {
    sh.insertChart(sh.newChart().setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(11, 1, STATUSES.length, 2))
      .setPosition(2, 4, 0, 0).setOption('title', 'Orders this month by status').setOption('legend', 'none')
      .build());
  }
  return { today: today, month: month, lowStock: low };
}

/** Email the owner a summary of the last 24 hours. Scheduled daily by installTriggers. */
function sendDailySummary() {
  const settings = getSettings_();
  const now = new Date();
  const end = new Date(now.getTime() + 1000); // include orders created this very second
  const start = new Date(now.getTime() - 24 * 3600 * 1000);
  const summary = summarizeOrders(readAllOrders_(settings), start, end);
  const low = lowStockList(getProducts_());
  const email = buildSummaryEmail(summary, low, settings, 'the 24 hours to ' + formatDate_(now));
  if (!settings.ownerEmail) {
    log_('Daily summary not sent', '', 'Add your address to "Owner email" in Settings');
    return email;
  }
  MailApp.sendEmail({ to: settings.ownerEmail, subject: email.subject, body: email.body, name: 'Order Desk' });
  log_('Daily summary sent', '', 'To ' + settings.ownerEmail + ' | ' + email.subject);
  return email;
}

/**
 * Menu: install the automations. Safe to run again: old Order Desk triggers are removed first.
 * - On form submit -> onFormSubmit (orders from a linked Google Form)
 * - On edit        -> handleEdit (status changes, item edits)
 * - Daily          -> sendDailySummary at the hour in Settings
 * - Hourly         -> refreshDashboard
 */
function installTriggers() {
  const ss = ss_();
  const ours = ['onFormSubmit', 'handleEdit', 'sendDailySummary', 'refreshDashboard'];
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (ours.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('sendDailySummary').timeBased().everyDays(1).atHour(getSettings_().summaryHour).create();
  ScriptApp.newTrigger('refreshDashboard').timeBased().everyHours(1).create();
  log_('Automations installed', '', ours.join(', '));
  ss.toast('Automations are on: form orders, status emails, daily summary.', 'Order Desk', 8);
}
