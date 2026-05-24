Namespace Models

    Public Class StockAdjustmentSaveResult
        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property AdjustmentId As Integer
        Public Property AdjustmentNumber As String = String.Empty

        Public Shared Function Success(adjustmentId As Integer, adjustmentNumber As String, message As String) As StockAdjustmentSaveResult
            Return New StockAdjustmentSaveResult With {
                .IsSuccess = True,
                .AdjustmentId = adjustmentId,
                .AdjustmentNumber = adjustmentNumber,
                .Message = message
            }
        End Function

        Public Shared Function Failure(message As String) As StockAdjustmentSaveResult
            Return New StockAdjustmentSaveResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function
    End Class

End Namespace
