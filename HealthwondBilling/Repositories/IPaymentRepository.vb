Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IPaymentRepository
        Function GenerateNextCustomerReceiptNumber(paymentDate As DateTime) As String
        Function GenerateNextSupplierPaymentNumber(paymentDate As DateTime) As String
        Function SearchCustomerPayments(customerId As Integer, fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of CustomerPaymentHistoryRow)
        Function SearchSupplierPayments(supplierId As Integer, fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of SupplierPaymentHistoryRow)
        Function SaveCustomerPayment(draft As CustomerPaymentDraft, createdByUserId As Integer) As Integer
        Function SaveSupplierPayment(draft As SupplierPaymentDraft, createdByUserId As Integer) As Integer
    End Interface

End Namespace
