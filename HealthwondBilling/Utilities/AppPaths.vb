Imports System.IO

Namespace Utilities

    Public NotInheritable Class AppPaths

        Private Sub New()
        End Sub

        Public Shared ReadOnly Property BaseDirectory As String
            Get
                Return AppDomain.CurrentDomain.BaseDirectory
            End Get
        End Property

        Public Shared ReadOnly Property DataRootDirectory As String
            Get
                Return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "HealthwondBilling")
            End Get
        End Property

        Public Shared ReadOnly Property DatabaseDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Database")
            End Get
        End Property

        Public Shared ReadOnly Property DatabaseFilePath As String
            Get
                Return Path.Combine(DatabaseDirectory, "healthwond.db")
            End Get
        End Property

        Public Shared ReadOnly Property LogsDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Logs")
            End Get
        End Property

        Public Shared ReadOnly Property GeneratedInvoicesDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Invoices")
            End Get
        End Property

        Public Shared ReadOnly Property ReportsDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Reports")
            End Get
        End Property

        Public Shared ReadOnly Property TemplatesDirectory As String
            Get
                Return Path.Combine(BaseDirectory, "Templates")
            End Get
        End Property

        Public Shared ReadOnly Property AssetsDirectory As String
            Get
                Return Path.Combine(BaseDirectory, "Assets")
            End Get
        End Property

        Public Shared Sub EnsureDirectories()
            Directory.CreateDirectory(DataRootDirectory)
            Directory.CreateDirectory(DatabaseDirectory)
            Directory.CreateDirectory(LogsDirectory)
            Directory.CreateDirectory(GeneratedInvoicesDirectory)
            Directory.CreateDirectory(ReportsDirectory)
            Directory.CreateDirectory(TemplatesDirectory)
            Directory.CreateDirectory(AssetsDirectory)
        End Sub

    End Class

End Namespace
