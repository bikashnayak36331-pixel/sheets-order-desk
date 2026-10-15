'use strict';
process.env.TZ = 'UTC';
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const { loadOrderDesk, SRC_ORDER } = require('./fakes');

test('dist/OrderDesk.gs contains every source file and works on its own', () => {
  const bundle = fs.readFileSync(path.join(__dirname, '..', 'dist', 'OrderDesk.gs'), 'utf8');
  SRC_ORDER.forEach((f) => assert.ok(bundle.includes(`// ----- ${f} -----`), f));
  const env = loadOrderDesk({ bundle: true, now: '2026-10-15T09:00:00Z' });
  env.context.setupOrderDesk();
  env.context.onFormSubmit({ namedValues: { Name: ['Dana'], Email: ['d@x.example'], Items: ['CAKE-CHOC x1'] } });
  const [order] = env.ss.getSheetByName('Orders').records();
  assert.equal(order['Order ID'], 'ORD-0001');
  assert.equal(order.Total, 34.56);
});
