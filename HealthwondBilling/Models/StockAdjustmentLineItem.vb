Namespace Models

    Public Class StockAdjustmentLineItem
        Public Property LineNumber As Integer
        Public Property ProductId As Integer
        Public Property ProductName As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property CurrentStock As Integer
        Public Property AdjustmentMode As StockAdjustmentMode = StockAdjustmentMode.Increase
        Public Property Quantity As Integer
        Public Property ResultingStock As Integer
        Public Property UnitCost As Decimal
        Public Property Remarks As String = String.Empty
    End Class

End Namespace
