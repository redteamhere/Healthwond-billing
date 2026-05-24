Namespace Models

    Public Class InvoiceHistoryRow
        Public Property InvoiceId As Integer
        Public Property InvoiceNumber As String = String.Empty
        Public Property InvoiceDate As DateTime
        Public Property CustomerName As String = String.Empty
        Public Property PaymentMode As String = String.Empty
        Public Property LineCount As Integer
        Public Property TotalUnits As Integer
        Public Property NetAmount As Decimal
        Public Property AmountPaid As Decimal
        Public Property BalanceAmount As Decimal
        Public Property UpdatedAt As DateTime
    End Class

End Namespace
