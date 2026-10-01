# Business POS System

Google Apps Script + Google Sheets based Business POS, Inventory, Sales, Purchase, Customer, Supplier and Finance foundation.

## Stack
- Google Apps Script Web App
- Google Sheets database
- Google Drive-ready architecture
- HTML/CSS/Vanilla JavaScript
- Mobile-first responsive UI

## Repository
This repository contains the canonical source for the Apps Script backend and web interface.

## Setup
1. Create a Google Spreadsheet (or let the script create one automatically).
2. Open **Extensions → Apps Script**.
3. Add `Code.gs`, `Index.html`, and `appsscript.json`.
4. In Apps Script, run:
   ```javascript
   setupDatabase("admin","CHANGE-THIS-STRONG-PASSWORD","your@email.com")
   ```
5. Authorize the script.
6. Deploy → New deployment → Web app.
7. Execute as: **Me**.
8. Choose the appropriate access setting for your organization.
9. Open the Web App URL and log in.

## Database
`setupDatabase()` creates the structured Sheets tables defined in `CONFIG.SHEETS`, including Products, Sales, Sale_Items, Purchases, Purchase_Items, Customers, Suppliers, Stock, Payments, Expenses, accounting tables, audit logs and system logs.

## Security
- Passwords are stored as SHA-256 hashes rather than plaintext.
- Sessions use Apps Script CacheService with expiry.
- Role/permission checks are centralized.
- Sensitive runtime Script Properties are not exposed to the browser.
- Change the setup password immediately and do not commit credentials.

## Current foundation
Working backend/API foundation includes authentication, database setup, products, stock, sales, purchases, customers, suppliers, expenses, dashboard, settings, audit logging and health/sync controls.

The Master Prompt's remaining advanced modules should be implemented incrementally without breaking these core flows.
