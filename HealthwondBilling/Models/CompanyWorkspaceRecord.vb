Namespace Models

    Public Class CompanyWorkspaceRecord

        Public Property WorkspaceId As String = String.Empty
        Public Property DisplayName As String = String.Empty
        Public Property DatabaseFilePath As String = String.Empty
        Public Property IsDefaultDemo As Boolean
        Public Property CreatedAt As DateTime

        Public Overrides Function ToString() As String
            If IsDefaultDemo Then
                Return $"{DisplayName} [Demo]"
            End If

            Return DisplayName
        End Function

    End Class

End Namespace
