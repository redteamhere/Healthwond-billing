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

        Public Shared ReadOnly Property GstInvoiceTemplateFilePath As String
            Get
                Return Path.Combine(TemplatesDirectory, "GSTInvoiceTemplate.xlsx")
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

        Public Shared Function ResolveConfiguredPath(configuredPath As String, fallbackAbsolutePath As String) As String
            Dim normalizedPath As String = If(configuredPath, String.Empty).Trim()
            If normalizedPath = String.Empty Then
                Return fallbackAbsolutePath
            End If

            If Path.IsPathRooted(normalizedPath) Then
                Return normalizedPath
            End If

            Return Path.GetFullPath(Path.Combine(BaseDirectory, normalizedPath))
        End Function

        Public Shared Function ToBaseRelativePath(pathValue As String) As String
            Dim normalizedPath As String = If(pathValue, String.Empty).Trim()
            If normalizedPath = String.Empty Then
                Return String.Empty
            End If

            If Not Path.IsPathRooted(normalizedPath) Then
                Return normalizedPath
            End If

            Dim baseDirectoryPath As String = BaseDirectory
            If Not baseDirectoryPath.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal) Then
                baseDirectoryPath &= Path.DirectorySeparatorChar
            End If

            Dim baseUri As New Uri(baseDirectoryPath, UriKind.Absolute)
            Dim fileUri As New Uri(normalizedPath, UriKind.Absolute)

            If baseUri.IsBaseOf(fileUri) Then
                Return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fileUri).ToString()).Replace("/"c, Path.DirectorySeparatorChar)
            End If

            Return normalizedPath
        End Function

    End Class

End Namespace
