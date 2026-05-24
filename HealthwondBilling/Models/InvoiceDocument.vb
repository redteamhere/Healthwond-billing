Namespace Models

    Public Class InvoiceDocument
        Public Property InvoiceId As Integer
        Public Property InvoiceNumber As String = String.Empty
        Public Property InvoiceDate As DateTime
        Public Property PaymentMode As String = String.Empty
        Public Property CustomerName As String = String.Empty
        Public Property CustomerGstin As String = String.Empty
        Public Property CustomerDrugLicenseNumber As String = String.Empty
        Public Property CustomerAddress As String = String.Empty
        Public Property CustomerPhone As String = String.Empty
        Public Property CompanyName As String = String.Empty
        Public Property CompanyAddress As String = String.Empty
        Public Property CompanyPhone As String = String.Empty
        Public Property CompanyGstin As String = String.Empty
        Public Property CompanyDrugLicenseNumber As String = String.Empty
        Public Property Notes As String = String.Empty
        Public Property SubTotal As Decimal
        Public Property DiscountAmount As Decimal
        Public Property SchemeAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property RoundOffAmount As Decimal
        Public Property NetAmount As Decimal
        Public Property AmountPaid As Decimal
        Public Property BalanceAmount As Decimal
        Public Property Items As New List(Of InvoiceDocumentItem)()
    End Class

End Namespace
