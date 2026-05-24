Namespace Models

    Public Class SupplierPaymentDraft
        Public Property PaymentNumber As String = String.Empty
        Public Property SupplierId As Integer
        Public Property SupplierName As String = String.Empty
        Public Property PaymentDate As DateTime
        Public Property PaymentMode As String = String.Empty
        Public Property ReferenceNumber As String = String.Empty
        Public Property Amount As Decimal
        Public Property OutstandingBeforePayment As Decimal
        Public Property Notes As String = String.Empty
    End Class

End Namespace
