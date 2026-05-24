Namespace Utilities

    Public NotInheritable Class SqliteDateHelper

        Private Sub New()
        End Sub

        Public Shared Function ToStorageDate(value As DateTime) As String
            Return value.ToString("yyyy-MM-dd")
        End Function

        Public Shared Function ToStorageDateTime(value As DateTime) As String
            Return value.ToString("yyyy-MM-dd HH:mm:ss")
        End Function

    End Class

End Namespace
