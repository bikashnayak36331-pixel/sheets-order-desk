/**
 * @OnlyCurrentDoc  (limits spreadsheet access to the sheet this script is attached to)
 *
 * Order Desk: configuration shared by every file.
 * Sheet names and column layouts live here so the rest of the code never hard-codes positions.
 */

const SHEETS = {
  ORDERS: 'Orders',
  PRODUCTS: 'Products',
  SETTINGS: 'Settings',
  DASHBOARD: 'Dashboard',
  LOG: 'Log',
};

const ORDER_COLUMNS = [
  'Order ID', 'Created', 'Customer', 'Email', 'Phone', 'Items', 'Discount %', 'Subtotal', 'Discount',
  'Tax', 'Total', 'Status', 'Stock', 'Invoice', 'Notes', 'Last update',
];

/** Values of the Stock column: whether this order's items have been taken out of stock. */
const STOCK_STATE = { RESERVED: 'Reserved', NOT_RESERVED: 'Not reserved', RETURNED: 'Returned' };

const PRODUCT_COLUMNS = ['SKU', 'Name', 'Price', 'Stock', 'Reorder at', 'Active'];

const LOG_COLUMNS = ['Time', 'Event', 'Order', 'Details'];

/** Allowed status changes. Anything else typed into the Status column is undone. */
const STATUS_FLOW = {
  'New': ['Confirmed', 'Paid', 'On hold', 'Cancelled'],
  'On hold': ['Confirmed', 'Paid', 'Cancelled'],
  'Confirmed': ['Paid', 'Cancelled'],
  'Paid': ['Shipped', 'Cancelled'],
  'Shipped': ['Delivered'],
  'Delivered': [],
  'Cancelled': [],
};

const STATUSES = Object.keys(STATUS_FLOW);

/** Statuses that should send the customer an email (when emails are switched on). */
const EMAIL_ON_STATUS = ['New', 'On hold', 'Confirmed', 'Paid', 'Shipped', 'Cancelled'];

/** Settings rows created by Setup, with their default values. */
const DEFAULT_SETTINGS = [
  ['Business name', 'My Shop'],
  ['Owner email', ''],
  ['Currency symbol', '$'],
  ['Tax rate %', 8],
  ['Order ID prefix', 'ORD-'],
  ['Send customer emails', 'No'],
  ['Invoice folder', 'Order Desk invoices'],
  ['Create invoice when paid', 'Yes'],
  ['Daily summary hour (0-23)', 7],
  ['Payment instructions', 'Pay by bank transfer or card. Reply to this email with any questions.'],
];

/** Number format for the Created and Last update columns. Re-applied on every write (see setDateCell_). */
const DATE_FORMAT = 'yyyy-mm-dd hh:mm';

/** Column Order Desk adds to a linked form's response sheet, to remember which responses became orders. */
const FORM_ID_HEADER = 'Order Desk ID';
const FORM_NOT_IMPORTED = 'Not imported (sent before Order Desk was set up)';

/** Statuses whose order total counts as money received. */
const PAID_STATUSES = ['Paid', 'Shipped', 'Delivered'];

const SAMPLE_PRODUCTS = [
  ['CAKE-CHOC', 'Chocolate fudge cake (8 in)', 32, 12, 3, 'Yes'],
  ['CAKE-LEMON', 'Lemon drizzle cake (8 in)', 28, 10, 3, 'Yes'],
  ['CUPC-12', 'Box of 12 cupcakes', 30, 20, 5, 'Yes'],
  ['COOKIE-24', 'Cookie tin (24 pieces)', 22, 15, 5, 'Yes'],
  ['BREAD-SOUR', 'Sourdough loaf', 8.5, 30, 8, 'Yes'],
];
