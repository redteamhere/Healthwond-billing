Namespace Models

    Public Class InventoryCurrentStockRow
        Public Property ProductName As String = String.Empty
        Public Property CompanyName As String = String.Empty
        Public Property Packing As String = String.Empty
        Public Property Composition As String = String.Empty
        Public Property BatchCount As Integer
        Public Property TotalStock As Integer
        Public Property EarliestExpiryDate As DateTime
        Public Property LatestExpiryDate As DateTime
        Public Property StockValueAtPTR As Decimal
        Public Property StockValueAtPTS As Decimal
        Public Property StockStatus As String = String.Empty
    End Class

End Namespace
