# Healthwond Billing System

Modules 1 to 6 deliver the foundation, SQLite bootstrap, authentication, dashboard shell, seeded sample data, master-data screens, billing, purchase entry, invoice export/print workflows, and the first reporting module.

## Current scope

- VB.NET WinForms application targeting `.NET Framework 4.8`
- SQLite database auto-created on first run
- Layered folders for forms, services, repositories, database helpers, utilities, assets, and templates
- Password hashing with PBKDF2
- Role-based authentication for `Admin` and `Staff`
- Dashboard metrics for today's sales, total stock, expiry alerts, low stock alerts, and pending payments
- Seeded sample users, customers, suppliers, products, and stock ledger rows
- Product master with searchable CRUD, stock-safe adjustments, barcode, GST, and pricing fields
- Customer master with searchable CRUD, GSTIN, drug license, address, contact data, and dues
- Billing screen with customer selection, product line entry, live GST calculations, round-off, dues handling, and invoice save with stock deduction
- Purchase screen with supplier management, purchase line entry, auto purchase numbering, live GST totals, inventory batch stock-in, and stock ledger posting
- Template-based GST invoice generation with ClosedXML, PDF export, print preview, instant print, and automatic `Templates/GSTInvoiceTemplate.xlsx` creation
- Reports screen with date-filtered sales, purchases, GST, stock, customer outstanding, and estimated profit views plus Excel export to `%LocalAppData%\HealthwondBilling\Reports`

## Default demo credentials

- `admin` / `Admin@123`
- `staff` / `Staff@123`

## NuGet packages

- `System.Data.SQLite.Core`
- `ClosedXML`
- `PdfSharp-MigraDoc-GDI`

## Data paths

- SQLite database: `%LocalAppData%\HealthwondBilling\Database\healthwond.db`
- Logs: `%LocalAppData%\HealthwondBilling\Logs`
- Generated invoices: `%LocalAppData%\HealthwondBilling\Invoices`
- Reports: `%LocalAppData%\HealthwondBilling\Reports`
- Excel templates copied to output under `Templates`

## Build

```powershell
dotnet restore
dotnet build
dotnet run --project .\HealthwondBilling\HealthwondBilling.vbproj
```

## Next modules

1. Settings screen and configuration management
2. Inventory views such as batch stock, expiry stock, and ledger reporting
3. Invoice edit/reprint history and document management
