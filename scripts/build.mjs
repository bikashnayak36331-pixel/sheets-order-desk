// Bundle src/*.js into one file, dist/OrderDesk.gs, for easy copy-paste into the Apps Script editor.
import { readFileSync, writeFileSync, mkdirSync } from 'node:fs';
import { createRequire } from 'node:module';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const { SRC_ORDER } = createRequire(import.meta.url)(join(root, 'test', 'fakes.js'));
const { version } = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));

const banner = `/**
 * Order Desk v${version}: order tracking, stock, PDF invoices and emails for Google Sheets.
 * https://github.com/bikashnayak36331-pixel/sheets-order-desk  (MIT License)
 *
 * Install: in your spreadsheet open Extensions > Apps Script, replace the contents of Code.gs
 * with this whole file, save, reload the spreadsheet, then use the "Order Desk" menu.
 * Built from src/ by scripts/build.mjs. Edit the files in src/, not this one.
 */
`;
const parts = SRC_ORDER.map((f) => `// ----- ${f} -----\n` + readFileSync(join(root, 'src', f), 'utf8').trim() + '\n');
mkdirSync(join(root, 'dist'), { recursive: true });
writeFileSync(join(root, 'dist', 'OrderDesk.gs'), banner + '\n' + parts.join('\n'));
console.log(`dist/OrderDesk.gs written (${SRC_ORDER.length} files)`);
