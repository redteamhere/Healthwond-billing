Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IInvoiceRepository
        Function GenerateNextInvoiceNumber(invoiceDate As DateTime) As String
        Function SaveInvoice(draft As BillingInvoiceDraft, createdByUserId As Integer) As Integer
        Function GetInvoiceDocument(invoiceId As Integer) As InvoiceDocument
    End Interface

End Namespace
