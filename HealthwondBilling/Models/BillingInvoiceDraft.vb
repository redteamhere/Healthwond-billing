Namespace Models

    Public Class BillingInvoiceDraft
        Public Property InvoiceId As Integer
        Public Property InvoiceNumber As String = String.Empty
        Public Property CustomerId As Integer
        Public Property CustomerName As String = String.Empty
        Public Property InvoiceDate As DateTime
        Public Property PaymentMode As String = String.Empty
        Public Property AmountPaid As Decimal
        Public Property Notes As String = String.Empty
        Public Property Items As New List(Of BillingLineItem)
        Public Property Summary As New BillingTotalsSummary()
    End Class

End Namespace
