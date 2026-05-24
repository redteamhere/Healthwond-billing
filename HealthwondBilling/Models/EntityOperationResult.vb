Namespace Models

    Public Class EntityOperationResult

        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property EntityId As Integer

        Public Shared Function Success(message As String, Optional entityId As Integer = 0) As EntityOperationResult
            Return New EntityOperationResult With {
                .IsSuccess = True,
                .Message = message,
                .EntityId = entityId
            }
        End Function

        Public Shared Function Failure(message As String) As EntityOperationResult
            Return New EntityOperationResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function

    End Class

End Namespace
