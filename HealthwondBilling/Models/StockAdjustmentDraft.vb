Namespace Models

    Public Class StockAdjustmentDraft
        Public Property AdjustmentId As Integer
        Public Property AdjustmentNumber As String = String.Empty
        Public Property AdjustmentDate As DateTime
        Public Property Notes As String = String.Empty
        Public Property Items As New List(Of StockAdjustmentLineItem)()
    End Class

End Namespace
