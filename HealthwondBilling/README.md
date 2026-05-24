# Healthwond Billing System

Module 1 delivers the project foundation, SQLite bootstrap, authentication, dashboard shell, and sample data.

## Current scope

- VB.NET WinForms application targeting `.NET Framework 4.8`
- SQLite database auto-created on first run
- Layered folders for forms, services, repositories, database helpers, utilities, assets, and templates
- Password hashing with PBKDF2
- Role-based authentication for `Admin` and `Staff`
- Dashboard metrics for today's sales, total stock, expiry alerts, low stock alerts, and pending payments
- Seeded sample users, customers, suppliers, products, and stock ledger rows

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

1. Product and customer master forms
2. Billing workflow with grid entry and GST calculations
3. Purchase module and stock ledger updates
4. ClosedXML invoice generation and print/PDF flows
5. Reports and settings screens
