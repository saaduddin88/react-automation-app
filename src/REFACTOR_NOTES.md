# MuntQ App — modular split

`App.jsx` (was ~3,455 lines) is now split into focused files. Drop all of these into your `src/` folder, replacing the single `App.jsx`. The entry point is unchanged: `App.jsx` still `export default App`, so `index.js` keeps importing `./App`.

## Files
- **App.jsx** — root component: state, localStorage, routes only.
- **Layout.jsx** — Header, NavItem, NavDropdown, Footer.
- **LandingPage.jsx** — landing page + its cards.
- **constants.js** — static data (columns, departments, categories, budget matrix, bank lists, approval steps).
- **utils.js** — shared helpers (dates, amount/money formatting, status helpers, approval-trail builder, etc.).
- **sadadValidation.js** — SADAD column definitions, validators, and the invoice category classifier.
- **SadadPayments.jsx** — SADAD Payments page, screen selector dropdown, workspace shell.
- **IqamaExpenseScreen.jsx** — Residence & Work Permits → Expense tab.
- **SadadPaymentScreen.jsx** — the invoice-identify Payment screen (used by Iqama Payment and STC).
- **AssetRequisition.jsx** — Asset Requisition + request details + approval timeline.
- **RequisitionTracking.jsx** — Requisition Tracking page.
- **Budget.jsx** — Budget upload + budget request details.
- **Reports.jsx** — Reports page.
- **BankTransfers.jsx** — Bank Transfers page.
- **sadadData.js**, **styles.css** — unchanged.

## To add a new SADAD category later
Edit `sadadValidation.js` (classifier rule) and wire the screen in `SadadPayments.jsx` — no need to touch unrelated pages.
