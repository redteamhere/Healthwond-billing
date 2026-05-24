Namespace Models

    Public Class PurchaseReturnSaveResult
        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property ReturnId As Integer
        Public Property ReturnNumber As String = String.Empty

        Public Shared Function Success(returnId As Integer, returnNumber As String, message As String) As PurchaseReturnSaveResult
            Return New PurchaseReturnSaveResult With {
                .IsSuccess = True,
                .ReturnId = returnId,
                .ReturnNumber = returnNumber,
                .Message = message
            }
        End Function

        Public Shared Function Failure(message As String) As PurchaseReturnSaveResult
            Return New PurchaseReturnSaveResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function
    End Class

End Namespace
