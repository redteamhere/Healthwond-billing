# Healthwond Billing System

Professional Windows desktop pharmaceutical billing and inventory software built with VB.NET WinForms, .NET Framework 4.8, and SQLite.

## Current module

Modules 1 and 2 are complete:

- Project scaffold with layered folders
- SQLite schema bootstrap on first run
- Password hashing and login flow for admin and staff
- Session handling and dashboard shell
- Seeded sample users, customers, suppliers, products, and stock ledger data
- Product master with searchable CRUD, stock adjustment logging, GST and pricing maintenance
- Customer master with searchable CRUD, GSTIN, license, address, contact, and outstanding balance maintenance
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
- Next module will cover the billing workflow
