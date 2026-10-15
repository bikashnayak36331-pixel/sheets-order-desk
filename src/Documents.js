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
