Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities
Imports System.Data.SQLite
Imports System.Globalization
Imports System.IO
Imports System.Linq
Imports System.Xml.Linq

Namespace Services

    Public Class CompanyWorkspaceService

        Private Const DefaultDemoWorkspaceId As String = "default-demonstration"
        Private Const DefaultDemoName As String = "Healthwond Default Demonstration"
        Private Const DefaultEditionLabel As String = "Desktop Edition"

        Public Function LoadCompanies() As IReadOnlyList(Of CompanyWorkspaceRecord)
            Dim state As RegistryState = LoadRegistryState()
            EnsureSeedWorkspaces(state)
            SaveRegistryState(state)
            Return state.Companies.OrderBy(Function(company) If(company.IsDefaultDemo, 1, 0)).ThenBy(Function(company) company.DisplayName).ToList()
        End Function

        Public Function LoadStartupProfile() As StartupShellProfile
            Dim state As RegistryState = LoadRegistryState()
            EnsureSeedWorkspaces(state)
            SaveRegistryState(state)
            Return New StartupShellProfile With {
                .SelectedCompanyId = state.Profile.SelectedCompanyId,
                .EditionLabel = state.Profile.EditionLabel
            }
        End Function

        Public Sub SaveSelectedCompany(companyId As String)
            Dim state As RegistryState = LoadRegistryState()
            state.Profile.SelectedCompanyId = AppPaths.SanitizeWorkspaceKey(companyId)
            SaveRegistryState(state)
        End Sub

        Public Function UpdateEditionLabel(editionLabel As String) As EntityOperationResult
            Dim normalizedEdition As String = If(editionLabel, String.Empty).Trim()
            If normalizedEdition = String.Empty Then
                Return EntityOperationResult.Failure("ERP version label is required.")
            End If

            Dim state As RegistryState = LoadRegistryState()
            state.Profile.EditionLabel = normalizedEdition
            SaveRegistryState(state)
            Return EntityOperationResult.Success("ERP version label updated successfully.")
        End Function

        Public Function CreateCompany(displayName As String) As CompanyWorkspaceOperationResult
            Dim normalizedDisplayName As String = If(displayName, String.Empty).Trim()
            If normalizedDisplayName = String.Empty Then
                Return CompanyWorkspaceOperationResult.Failure("Company name is required.")
            End If

            Try
                Dim state As RegistryState = LoadRegistryState()
                EnsureSeedWorkspaces(state)

                Dim workspaceId As String = GenerateWorkspaceId(normalizedDisplayName, state.Companies)
                Dim company As New CompanyWorkspaceRecord With {
                    .WorkspaceId = workspaceId,
                    .DisplayName = normalizedDisplayName,
                    .DatabaseFilePath = AppPaths.GetCompanyDatabaseFilePath(workspaceId),
                    .CreatedAt = DateTime.Now,
                    .IsDefaultDemo = False
                }

                CreateWorkspaceDatabase(company.DatabaseFilePath, normalizedDisplayName)
                state.Companies.Add(company)
                state.Profile.SelectedCompanyId = company.WorkspaceId
                SaveRegistryState(state)

                AppLogger.Info($"Created workspace '{company.DisplayName}' at '{company.DatabaseFilePath}'.")
                Return CompanyWorkspaceOperationResult.Success("New company workspace created successfully.", company)
            Catch ex As Exception
                AppLogger.Error("Company creation failed.", ex)
                Return CompanyWorkspaceOperationResult.Failure("The company workspace could not be created.")
            End Try
        End Function

        Public Function RestoreDefaultDemonstration() As CompanyWorkspaceOperationResult
            Try
                Dim state As RegistryState = LoadRegistryState()
                EnsureSeedWorkspaces(state)

                Dim company As CompanyWorkspaceRecord =
                    state.Companies.FirstOrDefault(Function(record) record.IsDefaultDemo OrElse String.Equals(record.WorkspaceId, DefaultDemoWorkspaceId, StringComparison.OrdinalIgnoreCase))

                If company Is Nothing Then
                    company = New CompanyWorkspaceRecord With {
                        .WorkspaceId = DefaultDemoWorkspaceId,
                        .DisplayName = DefaultDemoName,
                        .DatabaseFilePath = AppPaths.GetCompanyDatabaseFilePath(DefaultDemoWorkspaceId),
                        .CreatedAt = DateTime.Now,
                        .IsDefaultDemo = True
                    }
                    state.Companies.Add(company)
                End If

                ResetWorkspaceDatabase(company.DatabaseFilePath, DefaultDemoName)
                state.Profile.SelectedCompanyId = company.WorkspaceId
                SaveRegistryState(state)

                AppLogger.Info("Default demonstration workspace was restored.")
                Return CompanyWorkspaceOperationResult.Success("Default demonstration restored successfully.", company)
            Catch ex As Exception
                AppLogger.Error("Default demonstration restore failed.", ex)
                Return CompanyWorkspaceOperationResult.Failure("The default demonstration could not be restored.")
            End Try
        End Function

        Public Function RestoreFromBackup(sourceFilePath As String, displayName As String) As CompanyWorkspaceOperationResult
            Dim normalizedSourcePath As String = If(sourceFilePath, String.Empty).Trim()
            If normalizedSourcePath = String.Empty OrElse Not File.Exists(normalizedSourcePath) Then
                Return CompanyWorkspaceOperationResult.Failure("Choose a valid demonstration database file.")
            End If

            Dim normalizedDisplayName As String = If(displayName, String.Empty).Trim()
            If normalizedDisplayName = String.Empty Then
                normalizedDisplayName = Path.GetFileNameWithoutExtension(normalizedSourcePath)
            End If

            Try
                ValidateDatabaseFile(normalizedSourcePath)

                Dim state As RegistryState = LoadRegistryState()
                EnsureSeedWorkspaces(state)

                Dim workspaceId As String = GenerateWorkspaceId(normalizedDisplayName, state.Companies)
                Dim company As New CompanyWorkspaceRecord With {
                    .WorkspaceId = workspaceId,
                    .DisplayName = normalizedDisplayName,
                    .DatabaseFilePath = AppPaths.GetCompanyDatabaseFilePath(workspaceId),
                    .CreatedAt = DateTime.Now,
                    .IsDefaultDemo = False
                }

                AppPaths.EnsureCompanyDirectories(company.WorkspaceId)
                CopyDatabaseFile(normalizedSourcePath, company.DatabaseFilePath)
                InitializeExistingWorkspace(company.DatabaseFilePath)
                ApplyCompanyName(company.DatabaseFilePath, normalizedDisplayName)

                state.Companies.Add(company)
                state.Profile.SelectedCompanyId = company.WorkspaceId
                SaveRegistryState(state)

                AppLogger.Info($"Imported demonstration workspace '{company.DisplayName}' from '{normalizedSourcePath}'.")
                Return CompanyWorkspaceOperationResult.Success("Demonstration workspace imported successfully.", company)
            Catch ex As Exception
                AppLogger.Error("Demonstration import failed.", ex)
                Return CompanyWorkspaceOperationResult.Failure("The selected demonstration file could not be imported.")
            End Try
        End Function

        Public Function GetCompanyById(companyId As String) As CompanyWorkspaceRecord
            Dim companies As IReadOnlyList(Of CompanyWorkspaceRecord) = LoadCompanies()
            Dim normalizedCompanyId As String = AppPaths.SanitizeWorkspaceKey(companyId)
            Return companies.FirstOrDefault(Function(company) String.Equals(company.WorkspaceId, normalizedCompanyId, StringComparison.OrdinalIgnoreCase))
        End Function

        Private Function LoadRegistryState() As RegistryState
            AppPaths.EnsureDirectories()

            If Not File.Exists(AppPaths.CompanyRegistryFilePath) Then
                Return New RegistryState With {
                    .Profile = New StartupShellProfile With {
                        .EditionLabel = DefaultEditionLabel
                    }
                }
            End If

            Dim document As XDocument = XDocument.Load(AppPaths.CompanyRegistryFilePath)
            Dim root As XElement = document.Root

            Dim state As New RegistryState With {
                .Profile = New StartupShellProfile With {
                    .SelectedCompanyId = Convert.ToString(root.@selectedCompanyId),
                    .EditionLabel = If(Convert.ToString(root.@editionLabel), DefaultEditionLabel)
                }
            }

            For Each companyElement As XElement In root.<Companies>.<Company>
                Dim createdAtValue As DateTime = DateTime.Now
                Dim rawCreatedAt As String = Convert.ToString(companyElement.@createdAt)
                If rawCreatedAt <> String.Empty Then
                    DateTime.TryParseExact(rawCreatedAt, "O", CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, createdAtValue)
                End If

                Dim workspaceId As String = AppPaths.SanitizeWorkspaceKey(Convert.ToString(companyElement.@id))
                state.Companies.Add(New CompanyWorkspaceRecord With {
                    .WorkspaceId = workspaceId,
                    .DisplayName = Convert.ToString(companyElement.@displayName),
                    .DatabaseFilePath = AppPaths.GetCompanyDatabaseFilePath(workspaceId),
                    .IsDefaultDemo = ParseBoolean(Convert.ToString(companyElement.@isDefaultDemo)),
                    .CreatedAt = createdAtValue
                })
            Next

            Return state
        End Function

        Private Sub SaveRegistryState(state As RegistryState)
            AppPaths.EnsureDirectories()

            Dim document As New XDocument(
                New XElement(
                    "CompanyRegistry",
                    New XAttribute("selectedCompanyId", If(state.Profile.SelectedCompanyId, String.Empty)),
                    New XAttribute("editionLabel", If(state.Profile.EditionLabel, DefaultEditionLabel)),
                    New XElement(
                        "Companies",
                        state.Companies.OrderBy(Function(company) company.CreatedAt).
                            Select(Function(company) New XElement(
                                "Company",
                                New XAttribute("id", company.WorkspaceId),
                                New XAttribute("displayName", company.DisplayName),
                                New XAttribute("isDefaultDemo", company.IsDefaultDemo),
                                New XAttribute("createdAt", company.CreatedAt.ToString("O", CultureInfo.InvariantCulture)))))))
            document.Save(AppPaths.CompanyRegistryFilePath)
        End Sub

        Private Sub EnsureSeedWorkspaces(state As RegistryState)
            If state.Profile Is Nothing Then
                state.Profile = New StartupShellProfile With {
                    .EditionLabel = DefaultEditionLabel
                }
            End If

            If String.IsNullOrWhiteSpace(state.Profile.EditionLabel) Then
                state.Profile.EditionLabel = DefaultEditionLabel
            End If

            If state.Companies.Count = 0 Then
                If File.Exists(AppPaths.LegacyDatabaseFilePath) Then
                    Dim legacyCompanyName As String = ReadCompanyName(AppPaths.LegacyDatabaseFilePath)
                    If legacyCompanyName = String.Empty Then
                        legacyCompanyName = "Healthwond Pharmacy"
                    End If

                    Dim workspaceId As String = GenerateWorkspaceId(legacyCompanyName, state.Companies)
                    Dim migratedCompany As New CompanyWorkspaceRecord With {
                        .WorkspaceId = workspaceId,
                        .DisplayName = legacyCompanyName,
                        .DatabaseFilePath = AppPaths.GetCompanyDatabaseFilePath(workspaceId),
                        .CreatedAt = DateTime.Now,
                        .IsDefaultDemo = False
                    }

                    AppPaths.EnsureCompanyDirectories(migratedCompany.WorkspaceId)
                    CopyDatabaseFile(AppPaths.LegacyDatabaseFilePath, migratedCompany.DatabaseFilePath)
                    InitializeExistingWorkspace(migratedCompany.DatabaseFilePath)
                    state.Companies.Add(migratedCompany)
                    AppLogger.Info($"Migrated legacy database into workspace '{migratedCompany.DisplayName}'.")
                End If
            End If

            Dim demoCompany As CompanyWorkspaceRecord =
                state.Companies.FirstOrDefault(Function(company) company.IsDefaultDemo OrElse String.Equals(company.WorkspaceId, DefaultDemoWorkspaceId, StringComparison.OrdinalIgnoreCase))

            If demoCompany Is Nothing Then
                demoCompany = New CompanyWorkspaceRecord With {
                    .WorkspaceId = DefaultDemoWorkspaceId,
                    .DisplayName = DefaultDemoName,
                    .DatabaseFilePath = AppPaths.GetCompanyDatabaseFilePath(DefaultDemoWorkspaceId),
                    .CreatedAt = DateTime.Now,
                    .IsDefaultDemo = True
                }
                state.Companies.Add(demoCompany)
            End If

            If Not File.Exists(demoCompany.DatabaseFilePath) Then
                ResetWorkspaceDatabase(demoCompany.DatabaseFilePath, DefaultDemoName)
            Else
                InitializeExistingWorkspace(demoCompany.DatabaseFilePath)
            End If

            For Each company As CompanyWorkspaceRecord In state.Companies.Where(Function(record) Not record.IsDefaultDemo)
                If Not File.Exists(company.DatabaseFilePath) Then
                    CreateWorkspaceDatabase(company.DatabaseFilePath, company.DisplayName)
                Else
                    InitializeExistingWorkspace(company.DatabaseFilePath)
                End If
            Next

            Dim selectedCompany As CompanyWorkspaceRecord = state.Companies.FirstOrDefault(
                Function(company) String.Equals(company.WorkspaceId, state.Profile.SelectedCompanyId, StringComparison.OrdinalIgnoreCase))

            If selectedCompany Is Nothing Then
                selectedCompany = state.Companies.FirstOrDefault(Function(company) Not company.IsDefaultDemo)
                If selectedCompany Is Nothing Then
                    selectedCompany = state.Companies.First()
                End If
                state.Profile.SelectedCompanyId = selectedCompany.WorkspaceId
            End If
        End Sub

        Private Sub CreateWorkspaceDatabase(databaseFilePath As String, companyName As String)
            ResetWorkspaceDatabase(databaseFilePath, companyName)
        End Sub

        Private Sub ResetWorkspaceDatabase(databaseFilePath As String, companyName As String)
            DeleteWorkspaceFiles(databaseFilePath)
            InitializeExistingWorkspace(databaseFilePath)
            ApplyCompanyName(databaseFilePath, companyName)
        End Sub

        Private Sub InitializeExistingWorkspace(databaseFilePath As String)
            Dim connectionFactory As IDbConnectionFactory = New SqliteConnectionFactory(databaseFilePath)
            Dim bootstrapper As New DatabaseBootstrapper(connectionFactory, databaseFilePath)
            bootstrapper.Initialize()
        End Sub

        Private Sub ApplyCompanyName(databaseFilePath As String, companyName As String)
            Dim connectionFactory As IDbConnectionFactory = New SqliteConnectionFactory(databaseFilePath)
            Dim settingsRepository As New SettingsRepository(connectionFactory)
            Dim profile As AppSettingsProfile = settingsRepository.GetProfile()
            profile.CompanyName = companyName
            settingsRepository.SaveProfile(profile)
        End Sub

        Private Function ReadCompanyName(databaseFilePath As String) As String
            Try
                Using connection As New SQLiteConnection($"Data Source={databaseFilePath};Version=3;Read Only=True;")
                    connection.Open()
                    Using command As SQLiteCommand = connection.CreateCommand()
                        command.CommandText = "SELECT SettingValue FROM Settings WHERE SettingKey = 'CompanyName' LIMIT 1;"
                        Dim value As Object = command.ExecuteScalar()
                        Return Convert.ToString(value).Trim()
                    End Using
                End Using
            Catch
                Return String.Empty
            End Try
        End Function

        Private Sub ValidateDatabaseFile(filePath As String)
            Using connection As New SQLiteConnection($"Data Source={filePath};Version=3;Read Only=True;")
                connection.Open()
                Using command As SQLiteCommand = connection.CreateCommand()
                    command.CommandText = "PRAGMA integrity_check;"
                    Dim value As Object = command.ExecuteScalar()
                    If Not String.Equals(Convert.ToString(value), "ok", StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidDataException("The SQLite database failed integrity validation.")
                    End If
                End Using
            End Using
        End Sub

        Private Sub CopyDatabaseFile(sourceFilePath As String, destinationFilePath As String)
            DeleteWorkspaceFiles(destinationFilePath)
            Directory.CreateDirectory(Path.GetDirectoryName(destinationFilePath))
            File.Copy(sourceFilePath, destinationFilePath, True)
        End Sub

        Private Sub DeleteWorkspaceFiles(databaseFilePath As String)
            Directory.CreateDirectory(Path.GetDirectoryName(databaseFilePath))

            DeleteIfExists(databaseFilePath)
            DeleteIfExists(databaseFilePath & "-wal")
            DeleteIfExists(databaseFilePath & "-shm")
        End Sub

        Private Sub DeleteIfExists(filePath As String)
            If File.Exists(filePath) Then
                File.Delete(filePath)
            End If
        End Sub

        Private Function GenerateWorkspaceId(displayName As String, companies As IEnumerable(Of CompanyWorkspaceRecord)) As String
            Dim baseId As String = AppPaths.SanitizeWorkspaceKey(displayName)
            Dim candidate As String = baseId
            Dim index As Integer = 2

            While companies.Any(Function(company) String.Equals(company.WorkspaceId, candidate, StringComparison.OrdinalIgnoreCase))
                candidate = $"{baseId}-{index}"
                index += 1
            End While

            Return candidate
        End Function

        Private Function ParseBoolean(value As String) As Boolean
            Dim parsedValue As Boolean
            If Boolean.TryParse(value, parsedValue) Then
                Return parsedValue
            End If

            Return False
        End Function

        Private Class RegistryState

            Public Property Profile As StartupShellProfile
            Public ReadOnly Property Companies As List(Of CompanyWorkspaceRecord)

            Public Sub New()
                Companies = New List(Of CompanyWorkspaceRecord)()
            End Sub

        End Class

    End Class

End Namespace
