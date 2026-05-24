Namespace Models

    Public Class InvoiceDocumentItem
        Public Property LineNumber As Integer
        Public Property ProductName As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property ExpiryDate As DateTime
        Public Property Quantity As Integer
        Public Property FreeQuantity As Integer
        Public Property Rate As Decimal
        Public Property MRP As Decimal
        Public Property DiscountPercentage As Decimal
        Public Property DiscountAmount As Decimal
        Public Property SchemeDescription As String = String.Empty
        Public Property GstPercentage As Decimal
        Public Property TaxableAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property LineTotal As Decimal
    End Class

End Namespace
