Namespace Models

    Public Class PurchaseSaveResult

        Public Property IsSuccess As Boolean
        Public Property PurchaseId As Integer
        Public Property PurchaseNumber As String = String.Empty
        Public Property Message As String = String.Empty

        Public Shared Function Success(purchaseId As Integer, purchaseNumber As String, message As String) As PurchaseSaveResult
            Return New PurchaseSaveResult With {
                .IsSuccess = True,
                .PurchaseId = purchaseId,
                .PurchaseNumber = purchaseNumber,
                .Message = message
            }
        End Function

        Public Shared Function Failure(message As String) As PurchaseSaveResult
            Return New PurchaseSaveResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function

    End Class

End Namespace
