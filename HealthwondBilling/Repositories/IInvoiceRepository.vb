Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IInvoiceRepository
        Function GenerateNextInvoiceNumber(invoiceDate As DateTime) As String
        Function SaveInvoice(draft As BillingInvoiceDraft, createdByUserId As Integer) As Integer
        Function UpdateInvoice(draft As BillingInvoiceDraft, updatedByUserId As Integer) As Integer
        Function SearchInvoices(fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of InvoiceHistoryRow)
        Function LoadInvoiceDraft(invoiceId As Integer) As BillingInvoiceDraft
        Function GetInvoiceDocument(invoiceId As Integer) As InvoiceDocument
    End Interface

End Namespace
