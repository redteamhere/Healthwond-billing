Imports HealthwondBilling.Database
Imports HealthwondBilling.Forms
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
            AppLogger.Info("Healthwond Billing System is starting.")

            Dim companyWorkspaceService As New CompanyWorkspaceService()
            Dim runtimeFactory As New AppRuntimeFactory()
            Dim savedCredentialService As New SavedCredentialService()

            Application.Run(New FrmLogin(runtimeFactory, companyWorkspaceService, savedCredentialService))
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
