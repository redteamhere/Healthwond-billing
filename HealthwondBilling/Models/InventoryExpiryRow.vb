Namespace Models

    Public Class InventoryExpiryRow
        Public Property ProductName As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property ExpiryDate As DateTime
        Public Property DaysToExpiry As Integer
        Public Property CompanyName As String = String.Empty
        Public Property CurrentStock As Integer
        Public Property MRP As Decimal
        Public Property PTR As Decimal
        Public Property StockStatus As String = String.Empty
    End Class

End Namespace
