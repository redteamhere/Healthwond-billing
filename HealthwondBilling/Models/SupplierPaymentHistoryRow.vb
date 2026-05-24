Namespace Models

    Public Class SupplierPaymentHistoryRow
        Public Property PaymentId As Integer
        Public Property PaymentNumber As String = String.Empty
        Public Property PaymentDate As DateTime
        Public Property SupplierId As Integer
        Public Property SupplierName As String = String.Empty
        Public Property PaymentMode As String = String.Empty
        Public Property ReferenceNumber As String = String.Empty
        Public Property Amount As Decimal
        Public Property BalanceAfterPayment As Decimal
        Public Property UpdatedAt As DateTime
    End Class

End Namespace
