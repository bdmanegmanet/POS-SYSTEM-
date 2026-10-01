# Business POS + Inventory + Accounting System

Production-oriented Google Apps Script + Google Sheets business management system based on the supplied Master Prompt.

## Architecture
- Backend: Google Apps Script (Code.gs)
- Database: Google Sheets
- Storage: Google Drive
- Frontend: HTML5/CSS3/Vanilla JavaScript
- Render: Node.js Web Service wrapper/proxy for the same frontend

## Implemented modules
Authentication and sessions, role/permission checks, Products, Categories, Brands, Variants, Customers, Suppliers, POS Sales, Sales Returns, Purchases, Purchase Returns, Inventory, Stock Movements, Payments, Accounts, Expenses, Dashboard, Reports, Profit/Loss, Accounting summary, Settings, Sync/Push state, System Health, Audit Logs, System Logs, Backup/Restore, Developer Code Copy/Draft.

## Apps Script setup
1. Create/open a Google Sheet.
2. Open Extensions -> Apps Script.
3. Add Code.gs, Index.html, appsscript.json.
4. Run setupDatabase("admin","CHANGE-THIS-STRONG-PASSWORD","your@email.com").
5. Authorize the script.
6. Deploy as a Web App, executing as the owner, and copy the /exec URL.

## Render Web Service
Render serves Index.html through server.js. Set the Render environment variable GAS_WEB_APP_URL to your Apps Script /exec URL.
The Render server proxies /api requests to Apps Script doPost(), so the same frontend works both inside Apps Script and on Render.

## Data safety
Database initialization only creates missing sheets/headers and default settings/roles; it does not intentionally overwrite existing rows. Product stock changes are recorded in Stock_Movements.

## Developer
The authorized Developer panel can load Code.gs from the configured GitHub source, copy/download it, and save a non-deploying draft in Script Properties.