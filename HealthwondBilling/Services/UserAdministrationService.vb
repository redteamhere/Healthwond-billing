Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class UserAdministrationService

        Private ReadOnly _userRepository As IUserRepository

        Public Sub New(userRepository As IUserRepository)
            _userRepository = userRepository
        End Sub

        Public Async Function LoadUsersAsync() As Task(Of IReadOnlyList(Of UserAccount))
            Return Await Task.Run(Function() _userRepository.GetAll())
        End Function

        Public Async Function SaveUserAsync(user As UserAccount, Optional plainPassword As String = Nothing) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    If user Is Nothing Then
                        Return EntityOperationResult.Failure("Select a valid operator record.")
                    End If

                    If Not InputValidator.IsRequiredTextProvided(user.Username) Then
                        Return EntityOperationResult.Failure("Operator username is required.")
                    End If

                    If Not InputValidator.IsRequiredTextProvided(user.FullName) Then
                        Return EntityOperationResult.Failure("Operator full name is required.")
                    End If

                    Dim existingUser As UserAccount = _userRepository.GetByUsername(user.Username.Trim())
                    If existingUser IsNot Nothing AndAlso existingUser.Id <> user.Id Then
                        Return EntityOperationResult.Failure("That operator username already exists.")
                    End If

                    If user.Id <= 0 Then
                        If Not InputValidator.IsRequiredTextProvided(plainPassword) Then
                            Return EntityOperationResult.Failure("Set an initial password for the new operator.")
                        End If

                        Dim hashResult As PasswordHashResult = PasswordHasher.HashPassword(plainPassword)
                        user.PasswordHash = hashResult.Hash
                        user.PasswordSalt = hashResult.Salt
                    Else
                        Dim currentRecord As UserAccount = _userRepository.GetById(user.Id)
                        If currentRecord Is Nothing Then
                            Return EntityOperationResult.Failure("The selected operator no longer exists.")
                        End If

                        user.PasswordHash = currentRecord.PasswordHash
                        user.PasswordSalt = currentRecord.PasswordSalt
                    End If

                    Dim entityId As Integer = _userRepository.SaveUser(user)
                    Return EntityOperationResult.Success("Operator powers saved successfully.", entityId)
                End Function)
        End Function

        Public Async Function ResetPasswordAsync(userId As Integer, newPassword As String) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    If userId <= 0 Then
                        Return EntityOperationResult.Failure("Select an operator before resetting the password.")
                    End If

                    If Not InputValidator.IsRequiredTextProvided(newPassword) Then
                        Return EntityOperationResult.Failure("Enter a valid replacement password.")
                    End If

                    Dim existingRecord As UserAccount = _userRepository.GetById(userId)
                    If existingRecord Is Nothing Then
                        Return EntityOperationResult.Failure("The selected operator no longer exists.")
                    End If

                    Dim hashResult As PasswordHashResult = PasswordHasher.HashPassword(newPassword)
                    _userRepository.UpdatePassword(userId, hashResult.Hash, hashResult.Salt)
                    Return EntityOperationResult.Success("Operator password reset successfully.")
                End Function)
        End Function

    End Class

End Namespace
