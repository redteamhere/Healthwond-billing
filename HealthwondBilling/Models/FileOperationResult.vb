Namespace Models

    Public Class FileOperationResult

        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property FilePath As String = String.Empty

        Public Shared Function Success(message As String, filePath As String) As FileOperationResult
            Return New FileOperationResult With {
                .IsSuccess = True,
                .Message = message,
                .FilePath = filePath
            }
        End Function

        Public Shared Function Failure(message As String) As FileOperationResult
            Return New FileOperationResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function

    End Class

End Namespace
