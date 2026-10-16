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
    .addItem('Import missed form orders', 'importMissedFormOrdersNow')
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
  orders.getRange(2, col_('Created'), rows, 1).setNumberFormat(DATE_FORMAT);
  orders.getRange(2, col_('Last update'), rows, 1).setNumberFormat(DATE_FORMAT);
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
