Imports HealthwondBilling.Models

Namespace Services

    Public Class AppRuntimeContext

        Public Property Company As CompanyWorkspaceRecord
        Public Property AuthService As AuthService
        Public Property DashboardService As DashboardService
        Public Property ProductService As ProductService
        Public Property CustomerService As CustomerService
        Public Property SupplierService As SupplierService
        Public Property BillingService As BillingService
        Public Property PurchaseService As PurchaseService
        Public Property PurchasePrintService As PurchasePrintService
        Public Property InvoiceExportService As InvoiceExportService
        Public Property ReportService As ReportService
        Public Property InventoryService As InventoryService
        Public Property StockOperationService As StockOperationService
        Public Property SettlementService As SettlementService
        Public Property SettingsService As SettingsService
        Public Property MaintenanceService As MaintenanceService
        Public Property AccountingService As AccountingService
        Public Property UserAdministrationService As UserAdministrationService

    End Class

End Namespace
