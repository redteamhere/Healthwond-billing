Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IUserRepository
        Function GetAll() As IReadOnlyList(Of UserAccount)
        Function GetById(userId As Integer) As UserAccount
        Function GetByUsername(username As String) As UserAccount
        Function SaveUser(user As UserAccount) As Integer
        Sub UpdateLastLogin(userId As Integer, loggedInAt As DateTime)
        Sub UpdatePassword(userId As Integer, passwordHash As String, passwordSalt As String)
    End Interface

End Namespace
