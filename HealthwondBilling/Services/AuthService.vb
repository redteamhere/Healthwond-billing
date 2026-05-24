Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class AuthService

        Private ReadOnly _userRepository As IUserRepository

        Public Sub New(userRepository As IUserRepository)
            _userRepository = userRepository
        End Sub

        Public Async Function AuthenticateAsync(username As String, password As String) As Task(Of LoginResult)
            If Not InputValidator.IsRequiredTextProvided(username) Then
                Return LoginResult.Failure("Enter your username.")
            End If

            If Not InputValidator.IsRequiredTextProvided(password) Then
                Return LoginResult.Failure("Enter your password.")
            End If

            Return Await Task.Run(
                Function()
                    Dim normalizedUsername As String = username.Trim()
                    Dim user As UserAccount = _userRepository.GetByUsername(normalizedUsername)

                    If user Is Nothing Then
                        AppLogger.Warn($"Login failed for unknown username '{normalizedUsername}'.")
                        Return LoginResult.Failure("Invalid username or password.")
                    End If

                    If Not user.IsActive Then
                        AppLogger.Warn($"Login blocked for inactive account '{normalizedUsername}'.")
                        Return LoginResult.Failure("This account is inactive.")
                    End If

                    If Not PasswordHasher.VerifyPassword(password, user.PasswordHash, user.PasswordSalt) Then
                        AppLogger.Warn($"Login failed for username '{normalizedUsername}'.")
                        Return LoginResult.Failure("Invalid username or password.")
                    End If

                    _userRepository.UpdateLastLogin(user.Id, DateTime.Now)
                    user.LastLoginAt = DateTime.Now
                    AppLogger.Info($"User '{normalizedUsername}' signed in as {user.Role.ToFriendlyText()}.")
                    Return LoginResult.Success(user)
                End Function)
        End Function

    End Class

End Namespace
