# Changelog

## 1.1.0 (2026-10-16)

Fixes from the first full test in a real Google account (Sheets, Forms, Drive and Gmail).

- **Form orders can no longer be lost.** In the live test Google did not fire the form submit trigger for one of
  three responses. Each response now gets its order ID written next to it ("Order Desk ID" column), and an hourly
  check (plus the new menu item *Import missed form orders*) turns any response without an ID into an order.
  The same ID check stops a duplicate trigger from creating a second order.
- Responses that already exist when you install the automations are marked "Not imported" and never imported.
- Invoice PDFs: the PAID badge and table header now use borders and text colour, because Google's PDF converter
  drops background colours.
- Customer email subjects read naturally ("order ORD-0004 received", "payment received").
- Dashboard and daily summary: "Paid" now counts only Paid, Shipped and Delivered orders; the summary also shows
  the value still awaiting payment. Before, unpaid and on-hold orders were counted as revenue.
- "Today" and "this month" follow the spreadsheet's time zone.
- Created and Last update keep their `yyyy-mm-dd hh:mm` format (Sheets reset it when a date was written).
- Dashboard: the "Updated" time no longer gets cut off.
- 25 tests (3 new).

**Upgrading from 1.0.0:** replace the code in Code.gs with the new `dist/OrderDesk.gs`, save, then run
*Order Desk → Install automations* once. That swaps the old hourly trigger for the new hourly check and marks
the form responses you already have, so they are not imported twice.

## 1.0.0 (2026-10-15)

First public release.

- One-click setup: Orders, Products, Settings, Dashboard and Log sheets with formats, dropdowns and status colours.
- Orders from a linked Google Form (one product + quantity, or a multi-item list), from the menu, or typed in.
- Order IDs, prices, discount, tax and totals in exact cents; stock reserved, checked and returned.
- Status flow with guard-rails: skipped steps are undone with a message.
- PDF invoices saved to Google Drive and linked in the sheet; customer emails with a safe preview mode.
- Dashboard with a status chart; daily summary email; one-click trigger install.
- Single-file build (`dist/OrderDesk.gs`) for copy-paste installs.
- 22 tests that run the real script code against in-memory fakes of Sheets, Drive, Gmail and triggers.
