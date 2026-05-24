Namespace Models

    Public Class PurchaseLineItem
        Public Property LineNumber As Integer
        Public Property ProductId As Integer
        Public Property ProductName As String = String.Empty
        Public Property Packing As String = String.Empty
        Public Property HsnCode As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property ExpiryDate As DateTime
        Public Property CompanyName As String = String.Empty
        Public Property Composition As String = String.Empty
        Public Property Barcode As String = String.Empty
        Public Property ExistingStock As Integer
        Public Property Quantity As Integer
        Public Property FreeQuantity As Integer
        Public Property PTR As Decimal
        Public Property PTS As Decimal
        Public Property MRP As Decimal
        Public Property GstPercentage As Decimal
        Public Property TaxableAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property LineTotal As Decimal
    End Class

End Namespace
