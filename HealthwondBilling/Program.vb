Imports HealthwondBilling.Database
Imports HealthwondBilling.Forms
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Services
Imports HealthwondBilling.Utilities
Imports System.Threading

Friend Module Program

    <STAThread()>
    Friend Sub Main()
        AddHandler Application.ThreadException, AddressOf OnThreadException
        AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf OnUnhandledException

        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)

        Try
            AppPaths.EnsureDirectories()
            InvoiceTemplateGenerator.EnsureTemplateExists()
            AppLogger.Info("Healthwond Billing System is starting.")

            Dim connectionFactory As IDbConnectionFactory = New SqliteConnectionFactory(AppPaths.DatabaseFilePath)
            Dim bootstrapper As New DatabaseBootstrapper(connectionFactory)
            bootstrapper.Initialize()

            Dim userRepository As New UserRepository(connectionFactory)
            Dim productRepository As New ProductRepository(connectionFactory)
            Dim customerRepository As New CustomerRepository(connectionFactory)
            Dim supplierRepository As New SupplierRepository(connectionFactory)
            Dim invoiceRepository As New InvoiceRepository(connectionFactory)
            Dim purchaseRepository As New PurchaseRepository(connectionFactory)
            Dim reportRepository As New ReportRepository(connectionFactory)
            Dim authService As New AuthService(userRepository)
            Dim dashboardService As New DashboardService(connectionFactory)
            Dim productService As New ProductService(productRepository)
            Dim customerService As New CustomerService(customerRepository)
            Dim supplierService As New SupplierService(supplierRepository)
            Dim billingService As New BillingService(invoiceRepository, customerRepository, productRepository)
            Dim purchaseService As New PurchaseService(purchaseRepository, supplierRepository, productRepository)
            Dim invoiceExportService As New InvoiceExportService(invoiceRepository)
            Dim reportService As New ReportService(reportRepository)

            Application.Run(New FrmLogin(authService, dashboardService, productService, customerService, supplierService, billingService, purchaseService, invoiceExportService, reportService))
        Catch ex As Exception
            AppLogger.Error("Application bootstrap failed.", ex)
            MessageBox.Show(
                "The application could not start. Review the latest log file in the HealthwondBilling Logs folder.",
                "Startup Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub OnThreadException(sender As Object, e As ThreadExceptionEventArgs)
        AppLogger.Error("Unhandled UI thread exception.", e.Exception)
        MessageBox.Show(
            "An unexpected error occurred. The event was logged for review.",
            "Application Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Error)
    End Sub

    Private Sub OnUnhandledException(sender As Object, e As UnhandledExceptionEventArgs)
        Dim exceptionValue As Exception = TryCast(e.ExceptionObject, Exception)
        AppLogger.Error("Unhandled non-UI exception.", exceptionValue)
    End Sub

End Module
