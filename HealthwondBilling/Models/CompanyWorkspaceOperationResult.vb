Namespace Models

    Public Class CompanyWorkspaceOperationResult

        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property Company As CompanyWorkspaceRecord

        Public Shared Function Success(message As String, company As CompanyWorkspaceRecord) As CompanyWorkspaceOperationResult
            Return New CompanyWorkspaceOperationResult With {
                .IsSuccess = True,
                .Message = message,
                .Company = company
            }
        End Function

        Public Shared Function Failure(message As String) As CompanyWorkspaceOperationResult
            Return New CompanyWorkspaceOperationResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function

    End Class

End Namespace
