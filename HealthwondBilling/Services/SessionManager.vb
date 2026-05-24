Imports HealthwondBilling.Models

Namespace Services

    Public NotInheritable Class SessionManager

        Private Shared _currentUser As UserAccount

        Private Sub New()
        End Sub

        Public Shared ReadOnly Property CurrentUser As UserAccount
            Get
                Return _currentUser
            End Get
        End Property

        Public Shared ReadOnly Property IsAuthenticated As Boolean
            Get
                Return _currentUser IsNot Nothing
            End Get
        End Property

        Public Shared ReadOnly Property IsAdmin As Boolean
            Get
                Return _currentUser IsNot Nothing AndAlso _currentUser.Role = UserRole.Admin
            End Get
        End Property

        Public Shared Sub StartSession(user As UserAccount)
            _currentUser = user
        End Sub

        Public Shared Sub Clear()
            _currentUser = Nothing
        End Sub

    End Class

End Namespace
