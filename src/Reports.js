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
