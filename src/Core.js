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
