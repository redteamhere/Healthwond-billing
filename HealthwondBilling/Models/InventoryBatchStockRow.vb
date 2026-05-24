Namespace Models

    Public Class InventoryBatchStockRow
        Public Property ProductName As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property ExpiryDate As DateTime
        Public Property CompanyName As String = String.Empty
        Public Property Packing As String = String.Empty
        Public Property CurrentStock As Integer
        Public Property GstPercentage As Decimal
        Public Property MRP As Decimal
        Public Property PTR As Decimal
        Public Property PTS As Decimal
        Public Property Barcode As String = String.Empty
        Public Property StockStatus As String = String.Empty
    End Class

End Namespace
