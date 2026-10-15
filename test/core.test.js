'use strict';
process.env.TZ = 'UTC';
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadOrderDesk } = require('./fakes');

const { context: c } = loadOrderDesk();
const plain = (x) => JSON.parse(JSON.stringify(x)); // objects from the vm context have another realm's prototypes

test('money helpers keep exact cents', () => {
  assert.equal(c.toCents(8.5), 850);
  assert.equal(c.toCents('$1,234.56'), 123456);
  assert.equal(c.toCents(0.1 + 0.2), 30);
  assert.equal(c.formatMoney(123456, '$'), '$1,234.56');
  assert.equal(c.formatMoney(-5, '€'), '-€0.05');
  assert.equal(c.formatMoney(100000000, 'Rs '), 'Rs 1,000,000.00');
  assert.equal(c.percentToBp('8.25%'), 825);
  assert.throws(() => c.percentToBp(120), /between 0 and 100/);
  assert.throws(() => c.toCents('abc'), /Not a valid amount/);
});

test('parseItems reads friendly item lists and merges duplicates', () => {
  assert.deepEqual(plain(c.parseItems('cake-choc x2, CUPC-12')), [{ sku: 'CAKE-CHOC', qty: 2 }, { sku: 'CUPC-12', qty: 1 }]);
  assert.deepEqual(plain(c.parseItems('A x1; A x2\nB×3')), [{ sku: 'A', qty: 3 }, { sku: 'B', qty: 3 }]);
  assert.equal(c.formatItems(c.parseItems('a x2,b')), 'A x2, B x1');
  assert.throws(() => c.parseItems(''), /no items/);
  assert.throws(() => c.parseItems('two cakes please'), /Cannot read item/);
  assert.throws(() => c.parseItems('A x0'), /between 1 and 999/);
});

test('priceOrder applies discount before tax, in cents', () => {
  const idx = c.productIndex([['CAKE-CHOC', 'Cake', 32, 5, 1, 'Yes'], ['BREAD', 'Bread', 8.5, 9, 2, 'Yes'],
    ['OLD', 'Old', 3, 1, 0, 'No']]);
  const p = c.priceOrder(c.parseItems('CAKE-CHOC x2, BREAD x3'), idx, { taxRate: 8.25, discountPercent: 10 });
  assert.equal(p.subtotal, 8950); // 64.00 + 25.50
  assert.equal(p.discount, 895);
  assert.equal(p.tax, 665); // 8.25% of 80.55 = 6.645375 -> 6.65
  assert.equal(p.total, 8720);
  assert.throws(() => c.priceOrder(c.parseItems('NOPE'), idx, {}), /Unknown product SKU: NOPE/);
  assert.throws(() => c.priceOrder(c.parseItems('OLD'), idx, {}), /not active/);
  assert.deepEqual(plain(c.findShortages(c.parseItems('CAKE-CHOC x6'), idx)), [{ sku: 'CAKE-CHOC', wanted: 6, inStock: 5 }]);
});

test('order IDs continue from the highest existing number', () => {
  assert.equal(c.nextOrderId([], 'ORD-'), 'ORD-0001');
  assert.equal(c.nextOrderId(['ORD-0009', 'ORD-0002', 'X-0100', ''], 'ORD-'), 'ORD-0010');
});

test('status flow allows only the next steps', () => {
  assert.equal(c.canTransition('New', 'Paid'), true);
  assert.equal(c.canTransition('New', 'Shipped'), false);
  assert.equal(c.canTransition('Shipped', 'Delivered'), true);
  assert.equal(c.canTransition('Delivered', 'New'), false);
  assert.equal(c.canTransition('Cancelled', 'Paid'), false);
});

test('invoice HTML escapes customer text', () => {
  const idx = c.productIndex([['A', 'Apple <pie>', 2, 5, 1, 'Yes']]);
  const priced = c.priceOrder(c.parseItems('A x2'), idx, { taxRate: 0 });
  const html = c.buildInvoiceHtml({ id: 'ORD-0001', customer: '<script>x</script>', email: 'a@b.example',
    createdText: '2026-10-15', status: 'Paid', notes: 'Ring "twice"' }, priced,
  { businessName: 'Shop & Co', currency: '$', taxRate: 0, ownerEmail: '', paymentInstructions: 'Pay' });
  assert.match(html, /&lt;script&gt;x&lt;\/script&gt;/);
  assert.match(html, /Apple &lt;pie&gt;/);
  assert.match(html, /Shop &amp; Co/);
  assert.match(html, /PAID/);
  assert.doesNotMatch(html, /HOW TO PAY/); // already paid
});

test('emails read naturally', () => {
  const e = c.buildStatusEmail({ id: 'ORD-0003', customer: 'Dana Fox', items: 'A x1', totalText: '$9.00' }, 'Shipped',
    { businessName: 'Crumbs Bakery' });
  assert.equal(e.subject, 'Crumbs Bakery: order ORD-0003 shipped');
  assert.match(e.body, /^Hi Dana,/);
  assert.match(e.body, /on its way/);
  const s = c.buildSummaryEmail({ count: 0, revenueCents: 0, openOrders: 0, byStatus: {} }, [], { businessName: 'X', currency: '$' }, 'today');
  assert.match(s.body, /By status: none/);
});
