Namespace Models

    Public Class CustomerPaymentDraft
        Public Property ReceiptNumber As String = String.Empty
        Public Property CustomerId As Integer
        Public Property CustomerName As String = String.Empty
        Public Property PaymentDate As DateTime
        Public Property PaymentMode As String = String.Empty
        Public Property ReferenceNumber As String = String.Empty
        Public Property Amount As Decimal
        Public Property OutstandingBeforePayment As Decimal
        Public Property Notes As String = String.Empty
    End Class

End Namespace
