# Healthwond Billing System

Modules 1 to 3 deliver the foundation, SQLite bootstrap, authentication, dashboard shell, seeded sample data, master-data screens, and the first live billing workflow.

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

## Default demo credentials

- `admin` / `Admin@123`
- `staff` / `Staff@123`

## NuGet packages

- `System.Data.SQLite.Core`
- `ClosedXML`

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

1. Purchase module and stock ledger updates
2. ClosedXML invoice generation and print/PDF flows
3. Reports and settings screens
