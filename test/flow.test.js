'use strict';
process.env.TZ = 'UTC';
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadOrderDesk } = require('./fakes');

const NOW = '2026-10-15T10:30:00Z';

function ready(options = {}) {
  const env = loadOrderDesk({ now: NOW, ...options });
  env.context.setupOrderDesk();
  return env;
}
const sheet = (env, name) => env.ss.getSheetByName(name);
const orders = (env) => sheet(env, 'Orders').records();
const stockOf = (env, sku) => sheet(env, 'Products').records().find((p) => p.SKU === sku).Stock;
const logs = (env) => sheet(env, 'Log').records().map((r) => `${r.Event} | ${r.Order} | ${r.Details}`);
const setSetting = (env, key, value) => {
  const sh = sheet(env, 'Settings');
  const i = sh.records().findIndex((r) => r.Setting === key);
  assert.notEqual(i, -1, key);
  sh.getRange(i + 2, 2).setValue(value);
};
const statusCell = (env, row) => `L${row}`; // Status is column L
const form = (env, answers) => env.context.onFormSubmit({
  namedValues: Object.fromEntries(Object.entries(answers).map(([k, v]) => [k, [v]])),
});

test('setup builds every sheet once and is safe to run again', () => {
  const env = ready();
  assert.deepEqual(env.ss.getSheets().map((s) => s.name).sort(), ['Dashboard', 'Log', 'Orders', 'Products', 'Settings']);
  assert.equal(sheet(env, 'Orders').getRange('A1:P1').getValues()[0].join('|'),
    'Order ID|Created|Customer|Email|Phone|Items|Discount %|Subtotal|Discount|Tax|Total|Status|Stock|Invoice|Notes|Last update');
  assert.equal(sheet(env, 'Products').records().length, 5);
  assert.equal(sheet(env, 'Settings').records().length, 10);
  assert.ok(sheet(env, 'Orders').validations.some((v) => v.rule.requireValueInList[0].includes('Delivered')));
  env.context.setupOrderDesk();
  assert.equal(sheet(env, 'Products').records().length, 5, 'sample products are not added twice');
  assert.equal(sheet(env, 'Settings').records().length, 10, 'settings are not duplicated');
  env.context.onOpen();
  assert.equal(env.ss.menu.title, 'Order Desk');
  assert.ok(env.ss.menu.items.some(([, fn]) => fn === 'installTriggers'));
  assert.match(env.ss.toasts.at(-1).msg, /Order Desk is ready/);
});

test('a Google Form order gets an ID, prices, tax and reserves stock', () => {
  const env = ready();
  form(env, { 'Name': 'Dana Fox', 'Email': 'dana@fox.example', 'Product': 'Chocolate fudge cake (8 in)', 'Quantity': '2',
    'Notes': 'Birthday, please write "Happy 30th"' });
  const [o] = orders(env);
  assert.equal(o['Order ID'], 'ORD-0001');
  assert.equal(o.Items, 'CAKE-CHOC x2');
  assert.equal(o.Subtotal, 64);
  assert.equal(o.Tax, 5.12); // 8% default
  assert.equal(o.Total, 69.12);
  assert.equal(o.Status, 'New');
  assert.equal(o.Stock, 'Reserved');
  assert.equal(stockOf(env, 'CAKE-CHOC'), 10);
  assert.equal(env.mail.length, 0, 'emails are off by default');
  assert.ok(logs(env).some((l) => l.startsWith('Email preview (sending is off) | ORD-0001 | To dana@fox.example')));
  const dash = sheet(env, 'Dashboard').getRange('A3:B8').getValues();
  assert.deepEqual(dash.map((r) => r[1]).slice(0, 2).map(String), ['1', '$69.12']);
});

test('multi-item form answers and "Name (SKU)" dropdowns are understood', () => {
  const env = ready();
  form(env, { 'Your name': 'Sam Lee', 'Email address': 'sam@lee.example', 'Items': 'cupc-12 x1; BREAD-SOUR x3, CUPC-12' });
  form(env, { 'Name': 'Ana', 'Email': 'ana@x.example', 'Product': 'Lemon drizzle cake (8 in) (CAKE-LEMON)', 'Quantity': '1' });
  const [a, b] = orders(env);
  assert.equal(a.Items, 'CUPC-12 x2, BREAD-SOUR x3');
  assert.equal(a.Subtotal, 85.5);
  assert.equal(b.Items, 'CAKE-LEMON x1');
  assert.equal(b['Order ID'], 'ORD-0002');
});

test('problems never lose an order: it is saved On hold with a clear note', () => {
  const env = ready();
  form(env, { 'Name': 'Big Party', 'Email': 'p@x.example', 'Product': 'CAKE-CHOC', 'Quantity': '20' });
  form(env, { 'Name': 'Typo', 'Email': 'typo@x.example', 'Items': 'CAKE-BANANA x1' });
  form(env, { 'Name': 'No Email', 'Items': 'BREAD-SOUR' });
  const [short, unknown, noEmail] = orders(env);
  assert.equal(short.Status, 'On hold');
  assert.equal(short.Stock, 'Not reserved');
  assert.match(short.Notes, /Not enough stock: CAKE-CHOC wanted 20, in stock 12/);
  assert.equal(stockOf(env, 'CAKE-CHOC'), 12, 'stock untouched');
  assert.equal(unknown.Status, 'On hold');
  assert.match(unknown.Notes, /Unknown product SKU: CAKE-BANANA/);
  assert.equal(noEmail.Status, 'On hold');
  assert.match(noEmail.Notes, /No email address/);
  assert.ok(logs(env).some((l) => l.startsWith('Email skipped | ORD-0003')));
});

test('status changes: invoice PDF and emails when paid, illegal jumps are undone', () => {
  const env = ready();
  setSetting(env, 'Send customer emails', 'Yes');
  setSetting(env, 'Business name', 'Crumbs Bakery');
  setSetting(env, 'Owner email', 'owner@crumbs.example');
  form(env, { 'Name': 'Dana Fox', 'Email': 'dana@fox.example', 'Items': 'CAKE-CHOC x1, COOKIE-24 x2' });
  assert.equal(env.mail.length, 1);
  assert.equal(env.mail[0].subject, 'Crumbs Bakery: order ORD-0001 new');

  env.edit('Orders', statusCell(env, 2), 'Delivered'); // New -> Delivered skips steps
  assert.equal(orders(env)[0].Status, 'New');
  assert.match(env.ss.toasts.at(-1).msg, /can move to: Confirmed, Paid, On hold, Cancelled/);

  env.edit('Orders', statusCell(env, 2), 'Paid');
  const [o] = orders(env);
  assert.equal(o.Status, 'Paid');
  const folder = env.drive.folders.find((f) => f.name === 'Order Desk invoices');
  assert.equal(folder.files.length, 1);
  const pdf = folder.files[0];
  assert.equal(pdf.getName(), 'Invoice ORD-0001.pdf');
  assert.equal(pdf.blob.getContentType(), 'application/pdf');
  assert.match(pdf.blob.getDataAsString(), /Crumbs Bakery[\s\S]*ORD-0001[\s\S]*PAID[\s\S]*\$82\.08/);
  assert.equal(sheet(env, 'Orders').getRange('N2').getFormula(), `=HYPERLINK("${pdf.getUrl()}","Invoice PDF")`);
  const paidMail = env.mail.at(-1);
  assert.equal(paidMail.subject, 'Crumbs Bakery: order ORD-0001 paid');
  assert.equal(paidMail.replyTo, 'owner@crumbs.example');
  assert.equal(paidMail.attachments.length, 1);

  env.edit('Orders', statusCell(env, 2), 'Shipped');
  env.edit('Orders', statusCell(env, 2), 'Delivered');
  assert.equal(orders(env)[0].Status, 'Delivered');
  assert.equal(env.mail.length, 3, 'New, Paid and Shipped emails (Delivered sends none)');
  env.edit('Orders', statusCell(env, 2), 'Cancelled');
  assert.equal(orders(env)[0].Status, 'Delivered', 'final status cannot change');
});

test('cancelling returns reserved stock exactly once', () => {
  const env = ready();
  form(env, { 'Name': 'A', 'Email': 'a@x.example', 'Items': 'CUPC-12 x4' });
  assert.equal(stockOf(env, 'CUPC-12'), 16);
  env.edit('Orders', statusCell(env, 2), 'Cancelled');
  assert.equal(stockOf(env, 'CUPC-12'), 20);
  assert.equal(orders(env)[0].Stock, 'Returned');
});

test('an on-hold order reserves stock when it goes ahead, or is refused while stock is short', () => {
  const env = ready();
  form(env, { 'Name': 'A', 'Email': 'a@x.example', 'Items': 'CAKE-LEMON x15' });
  env.edit('Orders', statusCell(env, 2), 'Confirmed');
  assert.equal(orders(env)[0].Status, 'On hold');
  assert.match(env.ss.toasts.at(-1).msg, /Not enough stock for CAKE-LEMON/);
  sheet(env, 'Products').getRange('D3').setValue(40); // restock lemon cakes
  env.edit('Orders', statusCell(env, 2), 'Confirmed');
  assert.equal(orders(env)[0].Status, 'Confirmed');
  assert.equal(orders(env)[0].Stock, 'Reserved');
  assert.equal(stockOf(env, 'CAKE-LEMON'), 25);
});

test('editing items or discount reprices the order and moves stock', () => {
  const env = ready();
  form(env, { 'Name': 'A', 'Email': 'a@x.example', 'Items': 'BREAD-SOUR x2' });
  env.edit('Orders', 'F2', 'BREAD-SOUR x4, COOKIE-24');
  let [o] = orders(env);
  assert.equal(o.Items, 'BREAD-SOUR x4, COOKIE-24 x1');
  assert.equal(o.Subtotal, 56);
  assert.equal(stockOf(env, 'BREAD-SOUR'), 26);
  assert.equal(stockOf(env, 'COOKIE-24'), 14);
  env.edit('Orders', 'G2', 10); // 10% discount
  [o] = orders(env);
  assert.equal(o.Discount, 5.6);
  assert.equal(o.Total, 54.43); // (56 - 5.60) * 1.08 = 54.432
  env.edit('Orders', 'F2', 'BREAD-SOUR x999');
  assert.equal(orders(env)[0].Items, 'BREAD-SOUR x4, COOKIE-24 x1', 'refused edit is undone');
  env.edit('Orders', statusCell(env, 2), 'Paid');
  env.edit('Orders', 'F2', 'BREAD-SOUR x1');
  assert.equal(orders(env)[0].Items, 'BREAD-SOUR x4, COOKIE-24 x1', 'paid orders are locked');
});

test('edits elsewhere are ignored', () => {
  const env = ready();
  form(env, { 'Name': 'A', 'Email': 'a@x.example', 'Items': 'BREAD-SOUR' });
  const logCount = logs(env).length;
  env.edit('Products', 'B2', 'Renamed cake');
  env.edit('Orders', 'O2', 'A note');
  env.context.handleEdit({ range: sheet(env, 'Orders').getRange('L2:L3'), value: 'Paid' }); // multi-cell paste
  env.context.handleEdit(undefined);
  const [o] = orders(env);
  assert.equal(o.Status, 'New');
  assert.equal(o.Notes, 'A note');
  assert.equal(logs(env).length, logCount, 'nothing was logged');
});

test('order IDs stay unique across many orders', () => {
  const env = ready();
  for (let i = 0; i < 25; i++) form(env, { 'Name': 'C' + i, 'Email': `c${i}@x.example`, 'Items': 'BREAD-SOUR' });
  const ids = orders(env).map((o) => o['Order ID']);
  assert.equal(new Set(ids).size, 25);
  assert.equal(ids.at(-1), 'ORD-0025');
});

test('installing automations twice leaves exactly four triggers', () => {
  const env = ready();
  setSetting(env, 'Daily summary hour (0-23)', 18);
  env.context.installTriggers();
  env.context.installTriggers();
  const t = env.triggers.map((x) => `${x.fn}:${x.event}:${x.atHour || x.everyHours || ''}`).sort();
  assert.deepEqual(t, ['handleEdit:edit:', 'onFormSubmit:formSubmit:', 'refreshDashboard:time:1', 'sendDailySummary:time:18']);
});

test('daily summary goes to the owner and lists low stock', () => {
  const env = ready();
  env.context.sendDailySummary();
  assert.equal(env.mail.length, 0);
  assert.ok(logs(env).some((l) => l.includes('Add your address to "Owner email"')));
  setSetting(env, 'Owner email', 'owner@crumbs.example');
  form(env, { 'Name': 'A', 'Email': 'a@x.example', 'Items': 'CAKE-CHOC x10' }); // leaves 2, reorder at 3
  form(env, { 'Name': 'B', 'Email': 'b@x.example', 'Items': 'BREAD-SOUR' });
  env.edit('Orders', statusCell(env, 3), 'Cancelled');
  const email = env.context.sendDailySummary();
  const sent = env.mail.at(-1);
  assert.equal(sent.to, 'owner@crumbs.example');
  assert.equal(sent.subject, 'My Shop daily summary: 2 new orders');
  assert.match(email.body, /Revenue \(excluding cancelled\): \$345\.60/);
  assert.match(email.body, /By status: New 1, Cancelled 1/);
  assert.match(email.body, /Chocolate fudge cake \(8 in\) \(CAKE-CHOC\): 2 left/);
});

test('dashboard shows today, the month, status counts and one chart', () => {
  const env = ready();
  form(env, { 'Name': 'A', 'Email': 'a@x.example', 'Items': 'CUPC-12 x2' });
  form(env, { 'Name': 'B', 'Email': 'b@x.example', 'Items': 'CUPC-12' });
  env.edit('Orders', statusCell(env, 3), 'Paid');
  const d = sheet(env, 'Dashboard');
  const v = Object.fromEntries(d.getRange('A3:B17').getValues().map(([k, val]) => [k, String(val)]));
  assert.equal(v['Orders today'], '2');
  assert.equal(v['Revenue today'], '$97.20');
  assert.equal(v['New'], '1');
  assert.equal(v['Paid'], '1');
  assert.equal(d.getCharts().length, 1);
  env.context.refreshDashboard();
  assert.equal(d.getCharts().length, 1, 'chart is not duplicated');
});

test('menu actions: test order and invoice for the selected row', () => {
  const env = ready();
  const order = env.context.addTestOrder();
  assert.equal(order.id, 'ORD-0001');
  assert.equal(order.email, 'owner@shop.example');
  env.ss.activeRange = sheet(env, 'Orders').getRange('C2');
  const file = env.context.createInvoiceForSelection();
  assert.equal(file.getName(), 'Invoice ORD-0001.pdf');
  env.ss.activeRange = sheet(env, 'Products').getRange('A2');
  assert.throws(() => env.context.createInvoiceForSelection(), /Select a cell/);
});
