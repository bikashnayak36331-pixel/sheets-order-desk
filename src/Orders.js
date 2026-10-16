/**
 * Order Desk: creating orders (from a Google Form or the menu) and reacting to status changes.
 */

/**
 * Installable "On form submit" trigger. Works with a Google Form linked to this spreadsheet.
 * Recognised questions (case does not matter): Name, Email, Phone, Product, Quantity, Items, Notes.
 * Use either Product + Quantity (one product) or Items ("CAKE-CHOC x2, CUPC-12").
 *
 * The new order's ID is written next to the response (column "Order Desk ID"), so a response is never turned
 * into two orders, even if Google fires the trigger twice, and importMissedFormOrders() can find responses
 * whose trigger never fired.
 */
function onFormSubmit(e) {
  const input = formAnswersToInput_((e && e.namedValues) || {});
  input.source = 'Form';
  if (e && e.range) input.response = { sheet: e.range.getSheet(), row: e.range.getRow() };
  return createOrder_(input);
}

/** Turn form answers ({question: [answer]}) into the input for createOrder_. */
function formAnswersToInput_(named) {
  const answers = {};
  Object.keys(named).forEach(function (q) {
    const v = named[q];
    answers[String(q).trim().toLowerCase()] = String((Array.isArray(v) ? v[0] : v) || '').trim();
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
  return {
    customer: pick('name', 'your name', 'full name', 'customer'),
    email: pick('email', 'email address', 'your email'),
    phone: pick('phone', 'phone number', 'mobile'),
    items: items,
    notes: pick('notes', 'comments', 'message', 'delivery notes'),
  };
}

/**
 * Turn form responses that have no order yet into orders. Google occasionally does not fire the form submit
 * trigger (seen in a live test on 16 Oct 2026), so this runs every hour (hourlyCheck) and from the menu.
 * Responses that were already there when the automations were installed are marked "Not imported" and skipped.
 */
function importMissedFormOrders() {
  let imported = 0;
  formResponseSheets_().forEach(function (sh) {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const idCol = formIdColumn_(sh);
    const values = sh.getRange(1, 1, lastRow, Math.max(idCol, sh.getLastColumn())).getValues();
    const header = values[0];
    for (let r = 1; r < values.length; r++) {
      if (values[r][0] === '' || values[r][idCol - 1] !== '') continue; // empty row, or already handled
      const named = {};
      header.forEach(function (h, i) { if (i > 0 && i !== idCol - 1 && h !== '') named[h] = [values[r][i]]; });
      const input = formAnswersToInput_(named);
      input.source = 'Form, caught up';
      input.response = { sheet: sh, row: r + 1 };
      if (createOrder_(input)) imported += 1;
    }
  });
  if (imported) log_('Missed form orders imported', '', imported + ' order(s) whose form trigger did not run');
  return imported;
}

/** Menu: run the catch-up now and say what happened. */
function importMissedFormOrdersNow() {
  const n = importMissedFormOrders();
  ss_().toast(n ? 'Imported ' + n + ' missed form order(s).' : 'Every form response already has an order.',
    'Order Desk', 6);
  return n;
}

/** Installing the automations: mark responses that already exist so the catch-up never imports them. */
function markExistingFormResponses_() {
  formResponseSheets_().forEach(function (sh) {
    const lastRow = sh.getLastRow();
    if (lastRow < 2) return;
    const idCol = formIdColumn_(sh);
    const values = sh.getRange(2, 1, lastRow - 1, idCol).getValues();
    values.forEach(function (v, i) {
      if (v[0] !== '' && v[idCol - 1] === '') sh.getRange(i + 2, idCol).setValue(FORM_NOT_IMPORTED);
    });
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
    // A form response becomes at most one order: skip it if it already has an Order Desk ID.
    let responseCell = null;
    if (input.response) {
      responseCell = input.response.sheet.getRange(input.response.row, formIdColumn_(input.response.sheet));
      const existing = responseCell.getValue();
      if (existing !== '') {
        log_('Form response skipped', String(existing), 'Row ' + input.response.row + ' already has an order');
        return null;
      }
    }
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
    setDateCell_(sh, rowNumber, col_('Created'), now);
    setDateCell_(sh, rowNumber, col_('Last update'), now);
    if (responseCell) responseCell.setValue(id);
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
