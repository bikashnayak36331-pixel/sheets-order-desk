# Changelog

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
