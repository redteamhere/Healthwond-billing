Namespace Models

    Public Class SalesReportRow
        Public Property InvoiceDate As DateTime
        Public Property InvoiceNumber As String = String.Empty
        Public Property CustomerName As String = String.Empty
        Public Property PaymentMode As String = String.Empty
        Public Property LineCount As Integer
        Public Property TotalUnits As Integer
        Public Property TaxableAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property NetAmount As Decimal
        Public Property AmountPaid As Decimal
        Public Property BalanceAmount As Decimal
    End Class

End Namespace
