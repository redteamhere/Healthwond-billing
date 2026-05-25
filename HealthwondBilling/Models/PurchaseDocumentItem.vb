Namespace Models

    Public Class PurchaseDocumentItem
        Public Property LineNumber As Integer
        Public Property ProductName As String = String.Empty
        Public Property Packing As String = String.Empty
        Public Property HsnCode As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property ExpiryDate As DateTime
        Public Property Quantity As Integer
        Public Property FreeQuantity As Integer
        Public Property MRP As Decimal
        Public Property PTR As Decimal
        Public Property GstPercentage As Decimal
        Public Property TaxableAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property LineTotal As Decimal
    End Class

End Namespace
