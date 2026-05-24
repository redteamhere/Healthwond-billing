Namespace Models

    Public Class CustomerPaymentHistoryRow
        Public Property PaymentId As Integer
        Public Property ReceiptNumber As String = String.Empty
        Public Property PaymentDate As DateTime
        Public Property CustomerId As Integer
        Public Property CustomerName As String = String.Empty
        Public Property PaymentMode As String = String.Empty
        Public Property ReferenceNumber As String = String.Empty
        Public Property Amount As Decimal
        Public Property BalanceAfterPayment As Decimal
        Public Property UpdatedAt As DateTime
    End Class

End Namespace
