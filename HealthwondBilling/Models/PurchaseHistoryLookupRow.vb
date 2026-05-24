Namespace Models

    Public Class PurchaseHistoryLookupRow
        Public Property PurchaseId As Integer
        Public Property PurchaseNumber As String = String.Empty
        Public Property PurchaseDate As DateTime
        Public Property SupplierId As Integer
        Public Property SupplierName As String = String.Empty
        Public Property SupplierInvoiceNumber As String = String.Empty
        Public Property LineCount As Integer
        Public Property TotalUnits As Integer
        Public Property NetAmount As Decimal
        Public Property Notes As String = String.Empty
    End Class

End Namespace
