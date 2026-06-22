# DemoFinanceTracker

A personal finance tracker built on Google Sheets and Google Apps Script. No external backend — everything runs within Google Workspace.

## Live Demo

- **Dashboard:** https://sites.google.com/view/demo-finance-tracker/dashboard
- **Google Sheet:** https://docs.google.com/spreadsheets/d/1D-jiNcjuRQHsDUREOKBQdJ7MY-e9M5fr_x6vyf8i6js

## Features

- Real-time dashboard with account balances, credit card utilization, installment tracking, and monthly expense breakdown
- Credit card billing cycle navigator with full transaction history toggle
- Installment tracker with automatic monthly payment insertion via a daily time-based trigger
- Multi-currency support with FX-aware balance calculations
- Privacy toggle to hide all amounts on screen
- iOS Shortcut integration for hands-free transaction logging via a public REST endpoint
- Control center for managing accounts, categories, and installments without touching the spreadsheet
- Two-deployment security model: private dashboard (authenticated) and public API (anonymous)

## Sheet Structure

| Sheet | Purpose |
|---|---|
| 2026 Transactions | Transaction ledger (Date, Account, Amount, Currency, Target, Type, Category, Subcategory, Note) |
| Accounts | Account registry with balance formulas, CC metadata, and FX conversion |
| Account Categories | Lookup list for account types |
| Txn_Type | Transaction type reference (Expense, Income, Transfer, Adjustment) |
| Txn_Categories | Category list with type mapping |
| Txn_Subcategories | Subcategory list with category mapping |
| Currency | Currency list with rates to PHP |
| Installments | Active installment plans with term tracking |

## Apps Script Files

| File | Purpose |
|---|---|
| `Code.js` | `doGet` / `doPost` routing, chart data, transaction and installment queries |
| `Controls.js` | Add account, category, installment, and balance adjustment functions |
| `Trigger.js` | Daily trigger for automatic installment payment insertion |

## Deployment

Uses [clasp](https://github.com/google/clasp) for local development.

```bash
clasp push --force
clasp deploy --deploymentId <id> --description "description"
```

Two named deployments:
- **Private** — serves the dashboard, requires authentication
- **Public API** — handles iOS Shortcut POST and `shortcuts_data` GET, anonymous access

## Configuration

IDs are stored in Apps Script Script Properties (not hardcoded). Run `setProperties()` once in the editor after cloning to configure:

- `SPREADSHEET_ID`
- `PRIVATE_DEPLOYMENT_ID`
- `PUBLIC_DEPLOYMENT_ID`
- `CC_IMAGES_FOLDER_ID`
