Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class SettlementService

        Private ReadOnly _paymentRepository As IPaymentRepository

        Public Sub New(paymentRepository As IPaymentRepository)
            _paymentRepository = paymentRepository
        End Sub

        Public Async Function GenerateNextCustomerReceiptNumberAsync(paymentDate As DateTime) As Task(Of String)
            Return Await Task.Run(Function() _paymentRepository.GenerateNextCustomerReceiptNumber(paymentDate.Date))
        End Function

        Public Async Function GenerateNextSupplierPaymentNumberAsync(paymentDate As DateTime) As Task(Of String)
            Return Await Task.Run(Function() _paymentRepository.GenerateNextSupplierPaymentNumber(paymentDate.Date))
        End Function

        Public Async Function SearchCustomerPaymentsAsync(customerId As Integer, fromDate As DateTime, toDate As DateTime, searchTerm As String) As Task(Of List(Of CustomerPaymentHistoryRow))
            Return Await Task.Run(Function() _paymentRepository.SearchCustomerPayments(customerId, fromDate.Date, toDate.Date, NormalizeSearchTerm(searchTerm)))
        End Function

        Public Async Function SearchSupplierPaymentsAsync(supplierId As Integer, fromDate As DateTime, toDate As DateTime, searchTerm As String) As Task(Of List(Of SupplierPaymentHistoryRow))
            Return Await Task.Run(Function() _paymentRepository.SearchSupplierPayments(supplierId, fromDate.Date, toDate.Date, NormalizeSearchTerm(searchTerm)))
        End Function

        Public Async Function SaveCustomerPaymentAsync(draft As CustomerPaymentDraft, createdByUserId As Integer) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeCustomerPaymentDraft(draft)

                    Dim validationMessage As String = ValidateCustomerPaymentDraft(draft)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        Dim paymentId As Integer = _paymentRepository.SaveCustomerPayment(draft, createdByUserId)
                        AppLogger.Info($"Customer payment '{draft.ReceiptNumber}' saved with Id {paymentId}.")
                        Return EntityOperationResult.Success($"Customer collection {draft.ReceiptNumber} saved successfully.", paymentId)
                    Catch ex As Exception
                        AppLogger.Error($"Customer payment save failed for '{draft.ReceiptNumber}'.", ex)
                        Return EntityOperationResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Public Async Function SaveSupplierPaymentAsync(draft As SupplierPaymentDraft, createdByUserId As Integer) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeSupplierPaymentDraft(draft)

                    Dim validationMessage As String = ValidateSupplierPaymentDraft(draft)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        Dim paymentId As Integer = _paymentRepository.SaveSupplierPayment(draft, createdByUserId)
                        AppLogger.Info($"Supplier payment '{draft.PaymentNumber}' saved with Id {paymentId}.")
                        Return EntityOperationResult.Success($"Supplier payment {draft.PaymentNumber} saved successfully.", paymentId)
                    Catch ex As Exception
                        AppLogger.Error($"Supplier payment save failed for '{draft.PaymentNumber}'.", ex)
                        Return EntityOperationResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Private Sub NormalizeCustomerPaymentDraft(draft As CustomerPaymentDraft)
            draft.ReceiptNumber = If(draft.ReceiptNumber, String.Empty).Trim().ToUpperInvariant()
            draft.CustomerName = If(draft.CustomerName, String.Empty).Trim()
            draft.PaymentMode = If(draft.PaymentMode, String.Empty).Trim()
            draft.ReferenceNumber = If(draft.ReferenceNumber, String.Empty).Trim().ToUpperInvariant()
            draft.Notes = If(draft.Notes, String.Empty).Trim()
            draft.Amount = Decimal.Round(Math.Max(0D, draft.Amount), 2, MidpointRounding.AwayFromZero)
            draft.OutstandingBeforePayment = Decimal.Round(Math.Max(0D, draft.OutstandingBeforePayment), 2, MidpointRounding.AwayFromZero)
        End Sub

        Private Function ValidateCustomerPaymentDraft(draft As CustomerPaymentDraft) As String
            If Not InputValidator.IsRequiredTextProvided(draft.ReceiptNumber) Then
                Return "Receipt number is required."
            End If

            If draft.CustomerId <= 0 Then
                Return "Select a customer for the collection entry."
            End If

            If draft.Amount <= 0D Then
                Return "Collection amount must be greater than zero."
            End If

            If draft.OutstandingBeforePayment > 0D AndAlso draft.Amount > draft.OutstandingBeforePayment Then
                Return "Collection amount cannot exceed the customer's current outstanding."
            End If

            Return String.Empty
        End Function

        Private Sub NormalizeSupplierPaymentDraft(draft As SupplierPaymentDraft)
            draft.PaymentNumber = If(draft.PaymentNumber, String.Empty).Trim().ToUpperInvariant()
            draft.SupplierName = If(draft.SupplierName, String.Empty).Trim()
            draft.PaymentMode = If(draft.PaymentMode, String.Empty).Trim()
            draft.ReferenceNumber = If(draft.ReferenceNumber, String.Empty).Trim().ToUpperInvariant()
            draft.Notes = If(draft.Notes, String.Empty).Trim()
            draft.Amount = Decimal.Round(Math.Max(0D, draft.Amount), 2, MidpointRounding.AwayFromZero)
            draft.OutstandingBeforePayment = Decimal.Round(Math.Max(0D, draft.OutstandingBeforePayment), 2, MidpointRounding.AwayFromZero)
        End Sub

        Private Function ValidateSupplierPaymentDraft(draft As SupplierPaymentDraft) As String
            If Not InputValidator.IsRequiredTextProvided(draft.PaymentNumber) Then
                Return "Payment number is required."
            End If

            If draft.SupplierId <= 0 Then
                Return "Select a supplier for the payment entry."
            End If

            If draft.Amount <= 0D Then
                Return "Payment amount must be greater than zero."
            End If

            If draft.OutstandingBeforePayment > 0D AndAlso draft.Amount > draft.OutstandingBeforePayment Then
                Return "Payment amount cannot exceed the supplier outstanding balance."
            End If

            Return String.Empty
        End Function

        Private Function NormalizeSearchTerm(searchTerm As String) As String
            Return If(searchTerm, String.Empty).Trim()
        End Function

    End Class

End Namespace
