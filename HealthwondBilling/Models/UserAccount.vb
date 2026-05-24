Namespace Models

    Public Class UserAccount
        Public Property Id As Integer
        Public Property Username As String = String.Empty
        Public Property PasswordHash As String = String.Empty
        Public Property PasswordSalt As String = String.Empty
        Public Property FullName As String = String.Empty
        Public Property Role As UserRole
        Public Property IsActive As Boolean
        Public Property LastLoginAt As DateTime?
    End Class

End Namespace
