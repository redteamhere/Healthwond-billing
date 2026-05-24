Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class SupplierService

        Private ReadOnly _supplierRepository As ISupplierRepository

        Public Sub New(supplierRepository As ISupplierRepository)
            _supplierRepository = supplierRepository
        End Sub

        Public Async Function SearchAsync(searchTerm As String) As Task(Of List(Of SupplierRecord))
            Return Await Task.Run(Function() _supplierRepository.Search(searchTerm))
        End Function

        Public Async Function SaveAsync(supplier As SupplierRecord) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeSupplier(supplier)

                    Dim validationMessage As String = ValidateSupplier(supplier)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        Dim entityId As Integer = _supplierRepository.Save(supplier)
                        Dim successMessage As String = If(supplier.Id > 0, "Supplier updated successfully.", "Supplier created successfully.")
                        AppLogger.Info($"Supplier '{supplier.SupplierName}' saved with Id {entityId}.")
                        Return EntityOperationResult.Success(successMessage, entityId)
                    Catch ex As Exception
                        AppLogger.Error($"Supplier save failed for '{supplier.SupplierName}'.", ex)
                        Return EntityOperationResult.Failure("The supplier could not be saved.")
                    End Try
                End Function)
        End Function

        Public Async Function DeleteAsync(supplier As SupplierRecord) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    If supplier Is Nothing OrElse supplier.Id <= 0 Then
                        Return EntityOperationResult.Failure("Select a supplier to delete.")
                    End If

                    If supplier.OutstandingBalance > 0D Then
                        Return EntityOperationResult.Failure("Clear the supplier outstanding balance before deleting the record.")
                    End If

                    Try
                        If _supplierRepository.Delete(supplier.Id) Then
                            AppLogger.Info($"Supplier '{supplier.SupplierName}' deleted.")
                            Return EntityOperationResult.Success("Supplier deleted successfully.", supplier.Id)
                        End If

                        Return EntityOperationResult.Failure("The supplier could not be deleted.")
                    Catch ex As Exception
                        AppLogger.Error($"Supplier delete failed for '{supplier.SupplierName}'.", ex)
                        Return EntityOperationResult.Failure("The supplier could not be deleted. It may already be used in purchases.")
                    End Try
                End Function)
        End Function

        Private Sub NormalizeSupplier(supplier As SupplierRecord)
            supplier.SupplierName = If(supplier.SupplierName, String.Empty).Trim()
            supplier.Gstin = If(supplier.Gstin, String.Empty).Trim().ToUpperInvariant()
            supplier.DrugLicenseNumber = If(supplier.DrugLicenseNumber, String.Empty).Trim().ToUpperInvariant()
            supplier.Address = If(supplier.Address, String.Empty).Trim()
            supplier.Phone = If(supplier.Phone, String.Empty).Trim()
            supplier.Email = If(supplier.Email, String.Empty).Trim()
        End Sub

        Private Function ValidateSupplier(supplier As SupplierRecord) As String
            If Not InputValidator.IsRequiredTextProvided(supplier.SupplierName) Then
                Return "Supplier name is required."
            End If

            If supplier.Email <> String.Empty AndAlso Not InputValidator.IsValidEmail(supplier.Email) Then
                Return "Enter a valid email address."
            End If

            If supplier.OutstandingBalance < 0D Then
                Return "Outstanding balance cannot be negative."
            End If

            Return String.Empty
        End Function

    End Class

End Namespace
