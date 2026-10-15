// Generate README pictures from the real script code running in the test harness:
//   docs/sample-invoice.html  - the exact HTML Order Desk turns into a PDF invoice
//   docs/sheet-preview.html   - the Orders, Products and Dashboard sheets after a few demo orders
// Then: python3 scripts/screenshot.py  (renders both to docs/images/*.png)
import { writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

process.env.TZ = 'UTC';
const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const { loadOrderDesk } = createRequire(import.meta.url)(join(root, 'test', 'fakes.js'));

const env = loadOrderDesk({ now: '2026-10-15T10:30:00Z' });
const c = env.context;
c.setupOrderDesk();
const settings = env.ss.getSheetByName('Settings');
settings.getRange('B2').setValue('Crumbs Bakery');
settings.getRange('B3').setValue('orders@crumbs.example');
const submit = (a) => c.onFormSubmit({ namedValues: Object.fromEntries(Object.entries(a).map(([k, v]) => [k, [v]])) });
submit({ Name: 'Dana Fox', Email: 'dana@fox.example', Items: 'CAKE-CHOC x1, COOKIE-24 x2', Notes: 'Write "Happy 30th, Sam!"' });
submit({ Name: 'Leo Park', Email: 'leo@park.example', Product: 'Box of 12 cupcakes', Quantity: '3' });
submit({ Name: 'Grace Kim', Email: 'grace@kim.example', Items: 'BREAD-SOUR x4' });
submit({ Name: 'Office Party', Email: 'hr@acme.example', Items: 'CAKE-LEMON x15' });
submit({ Name: 'Omar Ali', Email: 'omar@ali.example', Items: 'CAKE-CHOC x9' });
env.edit('Orders', 'L2', 'Paid');
env.edit('Orders', 'L3', 'Confirmed');
env.edit('Orders', 'L4', 'Paid');
env.edit('Orders', 'L4', 'Shipped');

const pdf = env.drive.folders[0].files[0];
writeFileSync(join(root, 'docs', 'sample-invoice.html'), pdf.blob.getDataAsString());

const fmt = (v, money) => {
  if (v && typeof v === 'object' && typeof v.getTime === 'function') return new Date(v.getTime()).toISOString().slice(0, 16).replace('T', ' ');
  if (money && typeof v === 'number') return '$' + v.toFixed(2);
  return String(v ?? '');
};
const table = (name, moneyCols = [], skip = []) => {
  const sh = env.ss.getSheetByName(name);
  const [head, ...rows] = sh.data;
  const keep = head.map((h, i) => i).filter((i) => !skip.includes(head[i]));
  const th = keep.map((i) => `<th>${head[i]}</th>`).join('');
  const tr = rows.map((r) => '<tr>' + keep.map((i) => {
    const h = head[i]; let v = fmt(r[i], moneyCols.includes(h));
    if (h === 'Invoice' && sh.formulas[`${rows.indexOf(r) + 2},${i + 1}`]) v = '<u>Invoice PDF</u>';
    const cls = h === 'Status' ? ` class="s-${String(r[i]).replace(' ', '')}"` : '';
    return `<td${cls}>${v.replace(/</g, '&lt;').replace(/&lt;u>/g, '<u>').replace(/&lt;\/u>/g, '</u>')}</td>`;
  }).join('') + '</tr>').join('');
  return `<h2>${name}</h2><table><tr>${th}</tr>${tr}</table>`;
};
const dash = env.ss.getSheetByName('Dashboard').data.slice(0, 18).map((r) => `<tr><td>${r[0] ?? ''}</td><td>${r[1] ?? ''}</td></tr>`).join('');
const html = `<!doctype html><meta charset="utf-8"><style>
body{font-family:Arial,Helvetica,sans-serif;margin:24px;color:#202124;font-size:12px}
h2{font-size:14px;margin:18px 0 6px}table{border-collapse:collapse}
th{background:#1f2933;color:#fff;font-weight:bold;padding:5px 8px;border:1px solid #c0c0c0;text-align:left}
td{padding:4px 8px;border:1px solid #e0e0e0;white-space:nowrap}
.s-New{background:#e3f2fd}.s-Onhold{background:#fff3cd}.s-Confirmed{background:#e8eaf6}.s-Paid{background:#d1f2e3}
.s-Shipped{background:#e0f7fa}.s-Delivered{background:#c8e6c9}.s-Cancelled{background:#f8d7da}
.note{color:#5f6368;font-size:11px;margin-bottom:8px}.row{display:flex;gap:28px;align-items:flex-start}
</style>
<div class="note">Drawn from the sheet contents after the test harness ran five demo orders through the real script code.
Formatting is approximate; this is not a screenshot of Google Sheets.</div>
${table('Orders', ['Subtotal', 'Discount', 'Tax', 'Total'], ['Phone', 'Last update', 'Discount %'])}
<div class="row"><div>${table('Products', ['Price'])}</div><div><h2>Dashboard</h2><table>${dash}</table></div></div>`;
writeFileSync(join(root, 'docs', 'sheet-preview.html'), html);
console.log('wrote docs/sample-invoice.html and docs/sheet-preview.html');
