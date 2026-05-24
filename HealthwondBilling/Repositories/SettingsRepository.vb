Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class SettingsRepository
        Implements ISettingsRepository

        Private Shared ReadOnly SettingDescriptions As IReadOnlyDictionary(Of String, String) =
            New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) From {
                {"CompanyName", "Company display name used on invoices."},
                {"CompanyAddress", "Primary company address displayed on invoices."},
                {"CompanyPhone", "Primary company phone displayed on invoices."},
                {"CompanyGstin", "Company GSTIN displayed on invoices."},
                {"CompanyDrugLicense", "Company drug license displayed on invoices."},
                {"InvoicePrefix", "Prefix for auto-generated invoice numbers."},
                {"PurchasePrefix", "Prefix for auto-generated purchase numbers."},
                {"LowStockThreshold", "Default low stock alert threshold."},
                {"CurrencySymbol", "Default invoice currency symbol."},
                {"InvoiceTemplatePath", "Configured GST invoice template path."}
            }

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetProfile() As AppSettingsProfile Implements ISettingsRepository.GetProfile
            Dim values As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText = "SELECT SettingKey, SettingValue FROM Settings;"

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            values(Convert.ToString(reader("SettingKey"), CultureInfo.InvariantCulture)) =
                                Convert.ToString(reader("SettingValue"), CultureInfo.InvariantCulture)
                        End While
                    End Using
                End Using
            End Using

            Dim profile As New AppSettingsProfile With {
                .CompanyName = GetValue(values, "CompanyName", "Healthwond Pharmacy"),
                .CompanyAddress = GetValue(values, "CompanyAddress", "88 Medical Avenue, Bengaluru"),
                .CompanyPhone = GetValue(values, "CompanyPhone", "080-4000-1122"),
                .CompanyGstin = GetValue(values, "CompanyGstin", "29AAACH2024H1ZX"),
                .CompanyDrugLicense = GetValue(values, "CompanyDrugLicense", "DL-HWB-2024-01"),
                .InvoicePrefix = GetValue(values, "InvoicePrefix", "HWB"),
                .PurchasePrefix = GetValue(values, "PurchasePrefix", "PUR"),
                .LowStockThreshold = ParseInteger(GetValue(values, "LowStockThreshold", "10"), 10),
                .CurrencySymbol = GetValue(values, "CurrencySymbol", "Rs."),
                .InvoiceTemplatePath = GetValue(values, "InvoiceTemplatePath", IO.Path.Combine("Templates", "GSTInvoiceTemplate.xlsx"))
            }

            Return profile
        End Function

        Public Sub SaveProfile(profile As AppSettingsProfile) Implements ISettingsRepository.SaveProfile
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    UpsertSetting(connection, transaction, "CompanyName", profile.CompanyName)
                    UpsertSetting(connection, transaction, "CompanyAddress", profile.CompanyAddress)
                    UpsertSetting(connection, transaction, "CompanyPhone", profile.CompanyPhone)
                    UpsertSetting(connection, transaction, "CompanyGstin", profile.CompanyGstin)
                    UpsertSetting(connection, transaction, "CompanyDrugLicense", profile.CompanyDrugLicense)
                    UpsertSetting(connection, transaction, "InvoicePrefix", profile.InvoicePrefix)
                    UpsertSetting(connection, transaction, "PurchasePrefix", profile.PurchasePrefix)
                    UpsertSetting(connection, transaction, "LowStockThreshold", profile.LowStockThreshold.ToString(CultureInfo.InvariantCulture))
                    UpsertSetting(connection, transaction, "CurrencySymbol", profile.CurrencySymbol)
                    UpsertSetting(connection, transaction, "InvoiceTemplatePath", profile.InvoiceTemplatePath)
                    transaction.Commit()
                End Using
            End Using
        End Sub

        Private Sub UpsertSetting(connection As DbConnection, transaction As DbTransaction, key As String, value As String)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Settings (SettingKey, SettingValue, Description, UpdatedAt) " &
                    "VALUES (@SettingKey, @SettingValue, @Description, @UpdatedAt) " &
                    "ON CONFLICT(SettingKey) DO UPDATE SET " &
                    "SettingValue = excluded.SettingValue, " &
                    "Description = excluded.Description, " &
                    "UpdatedAt = excluded.UpdatedAt;"
                command.AddParameter("@SettingKey", key)
                command.AddParameter("@SettingValue", value)
                command.AddParameter("@Description", SettingDescriptions(key))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Function GetValue(values As IDictionary(Of String, String), key As String, defaultValue As String) As String
            If values.ContainsKey(key) Then
                Return values(key)
            End If

            Return defaultValue
        End Function

        Private Function ParseInteger(value As String, defaultValue As Integer) As Integer
            Dim parsedValue As Integer
            If Integer.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, parsedValue) Then
                Return parsedValue
            End If

            Return defaultValue
        End Function

    End Class

End Namespace
