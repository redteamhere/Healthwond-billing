# Healthwond Billing System

Professional Windows desktop pharmaceutical billing and inventory software built with VB.NET WinForms, .NET Framework 4.8, and SQLite.

## Current module

Modules 1 to 10 are complete:

- Project scaffold with layered folders
- SQLite schema bootstrap on first run
- Password hashing and login flow for admin and staff
- Session handling and dashboard shell
- Seeded sample users, customers, suppliers, products, and stock ledger data
- Product master with searchable CRUD, stock adjustment logging, GST and pricing maintenance
- Customer master with searchable CRUD, GSTIN, license, address, contact, and outstanding balance maintenance
- Billing workflow with customer selection, product lines, GST totals, round-off, invoice save, stock deduction, and customer balance updates
- Purchase workflow with supplier master maintenance, purchase numbering, batch-wise stock-in, stock ledger posting, and supplier outstanding updates
- Invoice document workflow with generated Excel GST invoices, PDF export, print preview, instant print, and automatic template bootstrap
- Reports workflow with live sales, purchase, GST, stock, outstanding, and profit summary views plus Excel export
- Admin settings workflow with company profile, prefixes, low-stock threshold, currency, and configurable invoice template path management
- Inventory workflow with current stock, batch stock, expiry watchlists, low stock alerts, stock-ledger history, and Excel export
- Invoice history workflow with saved invoice search, reopen-for-edit, re-export, preview, print, and generated file access
- Stock operations workflow with purchase returns, payable correction, manual stock adjustments, and ledger posting
- Clean build through `dotnet build`

## Solution layout

- `HealthwondBilling.sln`
- `HealthwondBilling/`

## Main technologies

- VB.NET WinForms
- .NET Framework 4.8
- SQLite
- ClosedXML

## Run locally

```powershell
dotnet restore
dotnet build .\HealthwondBilling.sln
dotnet run --project .\HealthwondBilling\HealthwondBilling.vbproj
```

## Default demo credentials

- `admin` / `Admin@123`
- `staff` / `Staff@123`

## Notes

- Runtime data is stored under `%LocalAppData%\HealthwondBilling`
- The detailed module README is in `HealthwondBilling/README.md`
- Next modules will cover settlement workflows and advanced analytics
