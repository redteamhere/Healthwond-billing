Namespace Models

    Public Class LoginResult

        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property User As UserAccount

        Public Shared Function Success(user As UserAccount) As LoginResult
            Return New LoginResult With {
                .IsSuccess = True,
                .Message = "Login successful.",
                .User = user
            }
        End Function

        Public Shared Function Failure(message As String) As LoginResult
            Return New LoginResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function

    End Class

End Namespace
