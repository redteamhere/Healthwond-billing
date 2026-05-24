Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IUserRepository
        Function GetByUsername(username As String) As UserAccount
        Sub UpdateLastLogin(userId As Integer, loggedInAt As DateTime)
    End Interface

End Namespace
