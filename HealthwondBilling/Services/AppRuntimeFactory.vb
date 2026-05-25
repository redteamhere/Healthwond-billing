Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class AppRuntimeFactory

        Public Function CreateRuntime(company As CompanyWorkspaceRecord) As AppRuntimeContext
            If company Is Nothing Then
                Throw New ArgumentNullException(NameOf(company))
            End If

            AppPaths.ConfigureCompanyWorkspace(company.WorkspaceId)
            AppPaths.EnsureDirectories()

            Dim connectionFactory As IDbConnectionFactory = New SqliteConnectionFactory(company.DatabaseFilePath)
            Dim bootstrapper As New DatabaseBootstrapper(connectionFactory, company.DatabaseFilePath)
            bootstrapper.Initialize()

            Dim userRepository As New UserRepository(connectionFactory)
            Dim productRepository As New ProductRepository(connectionFactory)
            Dim customerRepository As New CustomerRepository(connectionFactory)
            Dim supplierRepository As New SupplierRepository(connectionFactory)
            Dim invoiceRepository As New InvoiceRepository(connectionFactory)
            Dim purchaseRepository As New PurchaseRepository(connectionFactory)
            Dim reportRepository As New ReportRepository(connectionFactory)
            Dim inventoryRepository As New InventoryRepository(connectionFactory)
            Dim stockOperationRepository As New StockOperationRepository(connectionFactory)
            Dim paymentRepository As New PaymentRepository(connectionFactory)
            Dim settingsRepository As New SettingsRepository(connectionFactory)
            Dim accountingRepository As New AccountingRepository(connectionFactory)

            Dim authService As New AuthService(userRepository)
            Dim dashboardService As New DashboardService(connectionFactory)
            Dim productService As New ProductService(productRepository)
            Dim customerService As New CustomerService(customerRepository)
            Dim supplierService As New SupplierService(supplierRepository)
            Dim billingService As New BillingService(invoiceRepository, customerRepository, productRepository)
            Dim purchaseService As New PurchaseService(purchaseRepository, supplierRepository, productRepository)
            Dim reportService As New ReportService(reportRepository)
            Dim inventoryService As New InventoryService(inventoryRepository)
            Dim stockOperationService As New StockOperationService(stockOperationRepository, productRepository)
            Dim settlementService As New SettlementService(paymentRepository)
            Dim settingsService As New SettingsService(settingsRepository)
            Dim maintenanceService As New MaintenanceService(connectionFactory)
            Dim accountingService As New AccountingService(accountingRepository)
            Dim userAdministrationService As New UserAdministrationService(userRepository)

            Dim settingsProfile As AppSettingsProfile = settingsRepository.GetProfile()
            InvoiceTemplateGenerator.EnsureTemplateExists(settingsService.GetResolvedTemplatePath(settingsProfile))

            Dim invoiceExportService As New InvoiceExportService(invoiceRepository, settingsRepository)
            Dim purchasePrintService As New PurchasePrintService(purchaseRepository)
            accountingRepository.SynchronizeOperationalVouchers()

            Return New AppRuntimeContext With {
                .Company = company,
                .AuthService = authService,
                .DashboardService = dashboardService,
                .ProductService = productService,
                .CustomerService = customerService,
                .SupplierService = supplierService,
                .BillingService = billingService,
                .PurchaseService = purchaseService,
                .PurchasePrintService = purchasePrintService,
                .InvoiceExportService = invoiceExportService,
                .ReportService = reportService,
                .InventoryService = inventoryService,
                .StockOperationService = stockOperationService,
                .SettlementService = settlementService,
                .SettingsService = settingsService,
                .MaintenanceService = maintenanceService,
                .AccountingService = accountingService,
                .UserAdministrationService = userAdministrationService
            }
        End Function

    End Class

End Namespace
