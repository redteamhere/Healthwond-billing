Namespace Models

    Public Class InventoryLowStockRow
        Public Property ProductName As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property CompanyName As String = String.Empty
        Public Property CurrentStock As Integer
        Public Property ReorderThreshold As Integer
        Public Property ShortageUnits As Integer
        Public Property ExpiryDate As DateTime
        Public Property StockStatus As String = String.Empty
    End Class

End Namespace
