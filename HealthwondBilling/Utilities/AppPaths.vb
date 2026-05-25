Imports System.IO

Namespace Utilities

    Public NotInheritable Class AppPaths

        Private Shared _activeCompanyId As String = "default"

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

        Public Shared ReadOnly Property StartupDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Startup")
            End Get
        End Property

        Public Shared ReadOnly Property CompaniesDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Companies")
            End Get
        End Property

        Public Shared ReadOnly Property LegacyDatabaseDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Database")
            End Get
        End Property

        Public Shared ReadOnly Property LegacyDatabaseFilePath As String
            Get
                Return Path.Combine(LegacyDatabaseDirectory, "healthwond.db")
            End Get
        End Property

        Public Shared ReadOnly Property CompanyRegistryFilePath As String
            Get
                Return Path.Combine(StartupDirectory, "companies.xml")
            End Get
        End Property

        Public Shared ReadOnly Property SavedCredentialsFilePath As String
            Get
                Return Path.Combine(StartupDirectory, "saved-credentials.xml")
            End Get
        End Property

        Public Shared ReadOnly Property ActiveCompanyId As String
            Get
                Return _activeCompanyId
            End Get
        End Property

        Public Shared ReadOnly Property DatabaseDirectory As String
            Get
                Return GetCompanyDatabaseDirectory(_activeCompanyId)
            End Get
        End Property

        Public Shared ReadOnly Property DatabaseFilePath As String
            Get
                Return GetCompanyDatabaseFilePath(_activeCompanyId)
            End Get
        End Property

        Public Shared ReadOnly Property LogsDirectory As String
            Get
                Return Path.Combine(DataRootDirectory, "Logs")
            End Get
        End Property

        Public Shared ReadOnly Property GeneratedInvoicesDirectory As String
            Get
                Return Path.Combine(GetCompanyWorkspaceDirectory(_activeCompanyId), "Invoices")
            End Get
        End Property

        Public Shared ReadOnly Property ReportsDirectory As String
            Get
                Return Path.Combine(GetCompanyWorkspaceDirectory(_activeCompanyId), "Reports")
            End Get
        End Property

        Public Shared ReadOnly Property BackupsDirectory As String
            Get
                Return Path.Combine(GetCompanyWorkspaceDirectory(_activeCompanyId), "Backups")
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

        Public Shared Sub ConfigureCompanyWorkspace(companyId As String)
            _activeCompanyId = SanitizeWorkspaceKey(companyId)
            EnsureCompanyDirectories(_activeCompanyId)
        End Sub

        Public Shared Function GetCompanyWorkspaceDirectory(companyId As String) As String
            Return Path.Combine(CompaniesDirectory, SanitizeWorkspaceKey(companyId))
        End Function

        Public Shared Function GetCompanyDatabaseDirectory(companyId As String) As String
            Return Path.Combine(GetCompanyWorkspaceDirectory(companyId), "Database")
        End Function

        Public Shared Function GetCompanyDatabaseFilePath(companyId As String) As String
            Return Path.Combine(GetCompanyDatabaseDirectory(companyId), "healthwond.db")
        End Function

        Public Shared Function SanitizeWorkspaceKey(value As String) As String
            Dim normalizedValue As String = If(value, String.Empty).Trim().ToLowerInvariant()
            If normalizedValue = String.Empty Then
                normalizedValue = "default"
            End If

            Dim characters As New Text.StringBuilder(normalizedValue.Length)
            For Each characterValue As Char In normalizedValue
                If Char.IsLetterOrDigit(characterValue) Then
                    characters.Append(characterValue)
                ElseIf characterValue = " "c OrElse characterValue = "-"c OrElse characterValue = "_"c Then
                    characters.Append("-"c)
                End If
            Next

            Dim sanitized As String = characters.ToString().Trim("-"c)
            If sanitized = String.Empty Then
                sanitized = "default"
            End If

            Do While sanitized.Contains("--")
                sanitized = sanitized.Replace("--", "-")
            Loop

            Return sanitized
        End Function

        Public Shared Sub EnsureDirectories()
            Directory.CreateDirectory(DataRootDirectory)
            Directory.CreateDirectory(StartupDirectory)
            Directory.CreateDirectory(CompaniesDirectory)
            Directory.CreateDirectory(LegacyDatabaseDirectory)
            Directory.CreateDirectory(LogsDirectory)
            Directory.CreateDirectory(TemplatesDirectory)
            Directory.CreateDirectory(AssetsDirectory)
            EnsureCompanyDirectories(_activeCompanyId)
        End Sub

        Public Shared Sub EnsureCompanyDirectories(companyId As String)
            Dim normalizedCompanyId As String = SanitizeWorkspaceKey(companyId)
            Directory.CreateDirectory(GetCompanyWorkspaceDirectory(normalizedCompanyId))
            Directory.CreateDirectory(GetCompanyDatabaseDirectory(normalizedCompanyId))
            Directory.CreateDirectory(Path.Combine(GetCompanyWorkspaceDirectory(normalizedCompanyId), "Invoices"))
            Directory.CreateDirectory(Path.Combine(GetCompanyWorkspaceDirectory(normalizedCompanyId), "Reports"))
            Directory.CreateDirectory(Path.Combine(GetCompanyWorkspaceDirectory(normalizedCompanyId), "Backups"))
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
