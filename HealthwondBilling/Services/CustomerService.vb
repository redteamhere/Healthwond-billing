Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class CustomerService

        Private ReadOnly _customerRepository As ICustomerRepository

        Public Sub New(customerRepository As ICustomerRepository)
            _customerRepository = customerRepository
        End Sub

        Public Async Function SearchAsync(searchTerm As String) As Task(Of List(Of CustomerRecord))
            Return Await Task.Run(Function() _customerRepository.Search(searchTerm))
        End Function

        Public Async Function SaveAsync(customer As CustomerRecord) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeCustomer(customer)

                    Dim validationMessage As String = ValidateCustomer(customer)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        Dim entityId As Integer = _customerRepository.Save(customer)
                        Dim successMessage As String = If(customer.Id > 0, "Customer updated successfully.", "Customer created successfully.")
                        AppLogger.Info($"Customer '{customer.CustomerName}' saved with Id {entityId}.")
                        Return EntityOperationResult.Success(successMessage, entityId)
                    Catch ex As Exception
                        AppLogger.Error($"Customer save failed for '{customer.CustomerName}'.", ex)
                        Return EntityOperationResult.Failure("The customer could not be saved.")
                    End Try
                End Function)
        End Function

        Public Async Function DeleteAsync(customer As CustomerRecord) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    If customer Is Nothing OrElse customer.Id <= 0 Then
                        Return EntityOperationResult.Failure("Select a customer to delete.")
                    End If

                    If customer.OutstandingBalance > 0D Then
                        Return EntityOperationResult.Failure("Clear the customer's outstanding balance before deleting the record.")
                    End If

                    Try
                        If _customerRepository.Delete(customer.Id) Then
                            AppLogger.Info($"Customer '{customer.CustomerName}' deleted.")
                            Return EntityOperationResult.Success("Customer deleted successfully.", customer.Id)
                        End If

                        Return EntityOperationResult.Failure("The customer could not be deleted.")
                    Catch ex As Exception
                        AppLogger.Error($"Customer delete failed for '{customer.CustomerName}'.", ex)
                        Return EntityOperationResult.Failure("The customer could not be deleted. It may already be used in invoices.")
                    End Try
                End Function)
        End Function

        Private Sub NormalizeCustomer(customer As CustomerRecord)
            customer.CustomerName = If(customer.CustomerName, String.Empty).Trim()
            customer.Gstin = If(customer.Gstin, String.Empty).Trim().ToUpperInvariant()
            customer.DrugLicenseNumber = If(customer.DrugLicenseNumber, String.Empty).Trim().ToUpperInvariant()
            customer.Address = If(customer.Address, String.Empty).Trim()
            customer.Phone = If(customer.Phone, String.Empty).Trim()
            customer.Email = If(customer.Email, String.Empty).Trim()
        End Sub

        Private Function ValidateCustomer(customer As CustomerRecord) As String
            If Not InputValidator.IsRequiredTextProvided(customer.CustomerName) Then
                Return "Customer name is required."
            End If

            If customer.Email <> String.Empty AndAlso Not InputValidator.IsValidEmail(customer.Email) Then
                Return "Enter a valid email address."
            End If

            If customer.OutstandingBalance < 0D Then
                Return "Outstanding balance cannot be negative."
            End If

            Return String.Empty
        End Function

    End Class

End Namespace
