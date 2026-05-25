Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities
Imports System.Globalization

Namespace Services

    Public Class SettingsService

        Private ReadOnly _settingsRepository As ISettingsRepository

        Public Sub New(settingsRepository As ISettingsRepository)
            _settingsRepository = settingsRepository
        End Sub

        Public Async Function LoadAsync() As Task(Of AppSettingsProfile)
            Return Await Task.Run(Function() _settingsRepository.GetProfile())
        End Function

        Public Async Function SaveAsync(profile As AppSettingsProfile) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeProfile(profile)

                    Dim validationMessage As String = ValidateProfile(profile)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        _settingsRepository.SaveProfile(profile)
                        Dim templateAbsolutePath As String = GetResolvedTemplatePath(profile)
                        InvoiceTemplateGenerator.EnsureTemplateExists(templateAbsolutePath)
                        AppLogger.Info("Application settings were updated.")
                        Return EntityOperationResult.Success("Settings saved successfully.")
                    Catch ex As Exception
                        AppLogger.Error("Settings save failed.", ex)
                        Return EntityOperationResult.Failure("The settings could not be saved.")
                    End Try
                End Function)
        End Function

        Public Function GetResolvedTemplatePath(profile As AppSettingsProfile) As String
            Return AppPaths.ResolveConfiguredPath(profile.InvoiceTemplatePath, AppPaths.GstInvoiceTemplateFilePath)
        End Function

        Public Function GetDefaultTemplateSettingValue() As String
            Return AppPaths.ToBaseRelativePath(AppPaths.GstInvoiceTemplateFilePath)
        End Function

        Private Sub NormalizeProfile(profile As AppSettingsProfile)
            profile.CompanyName = If(profile.CompanyName, String.Empty).Trim()
            profile.CompanyAddress = If(profile.CompanyAddress, String.Empty).Trim()
            profile.CompanyPhone = If(profile.CompanyPhone, String.Empty).Trim()
            profile.CompanyGstin = If(profile.CompanyGstin, String.Empty).Trim().ToUpperInvariant()
            profile.CompanyDrugLicense = If(profile.CompanyDrugLicense, String.Empty).Trim().ToUpperInvariant()
            profile.InvoicePrefix = If(profile.InvoicePrefix, String.Empty).Trim().ToUpperInvariant()
            profile.PurchasePrefix = If(profile.PurchasePrefix, String.Empty).Trim().ToUpperInvariant()
            profile.ReceiptPrefix = If(profile.ReceiptPrefix, String.Empty).Trim().ToUpperInvariant()
            profile.SupplierPaymentPrefix = If(profile.SupplierPaymentPrefix, String.Empty).Trim().ToUpperInvariant()
            profile.CurrencySymbol = If(profile.CurrencySymbol, String.Empty).Trim()
            profile.InvoiceTemplatePath = If(profile.InvoiceTemplatePath, String.Empty).Trim()

            If profile.InvoiceTemplatePath = String.Empty Then
                profile.InvoiceTemplatePath = GetDefaultTemplateSettingValue()
            End If
        End Sub

        Private Function ValidateProfile(profile As AppSettingsProfile) As String
            If Not InputValidator.IsRequiredTextProvided(profile.CompanyName) Then
                Return "Company name is required."
            End If

            If Not InputValidator.IsRequiredTextProvided(profile.InvoicePrefix) Then
                Return "Invoice prefix is required."
            End If

            If Not InputValidator.IsRequiredTextProvided(profile.PurchasePrefix) Then
                Return "Purchase prefix is required."
            End If

            If Not InputValidator.IsRequiredTextProvided(profile.ReceiptPrefix) Then
                Return "Receipt prefix is required."
            End If

            If Not InputValidator.IsRequiredTextProvided(profile.SupplierPaymentPrefix) Then
                Return "Supplier payment prefix is required."
            End If

            If profile.LowStockThreshold < 0 Then
                Return "Low stock threshold cannot be negative."
            End If

            If Not InputValidator.IsRequiredTextProvided(profile.CurrencySymbol) Then
                Return "Currency symbol is required."
            End If

            If Not InputValidator.IsRequiredTextProvided(profile.InvoiceTemplatePath) Then
                Return "Invoice template path is required."
            End If

            Dim extension As String = IO.Path.GetExtension(profile.InvoiceTemplatePath)
            If Not String.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) Then
                Return "Invoice template path must point to an .xlsx file."
            End If

            Return String.Empty
        End Function

    End Class

End Namespace
