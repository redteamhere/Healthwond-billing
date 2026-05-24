Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class UserRepository
        Implements IUserRepository

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetByUsername(username As String) As UserAccount Implements IUserRepository.GetByUsername
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT Id, Username, PasswordHash, PasswordSalt, FullName, Role, IsActive, LastLoginAt " &
                        "FROM Users WHERE Username = @Username LIMIT 1;"
                    command.AddParameter("@Username", username.Trim())

                    Using reader = command.ExecuteReader()
                        If Not reader.Read() Then
                            Return Nothing
                        End If

                        Dim user As New UserAccount With {
                            .Id = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .Username = Convert.ToString(reader("Username"), CultureInfo.InvariantCulture),
                            .PasswordHash = Convert.ToString(reader("PasswordHash"), CultureInfo.InvariantCulture),
                            .PasswordSalt = Convert.ToString(reader("PasswordSalt"), CultureInfo.InvariantCulture),
                            .FullName = Convert.ToString(reader("FullName"), CultureInfo.InvariantCulture),
                            .Role = ParseRole(Convert.ToString(reader("Role"), CultureInfo.InvariantCulture)),
                            .IsActive = Convert.ToInt32(reader("IsActive"), CultureInfo.InvariantCulture) = 1
                        }

                        Dim lastLoginRaw As Object = reader("LastLoginAt")
                        If lastLoginRaw IsNot DBNull.Value Then
                            Dim parsedDate As DateTime
                            If DateTime.TryParse(Convert.ToString(lastLoginRaw, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                                user.LastLoginAt = parsedDate
                            End If
                        End If

                        Return user
                    End Using
                End Using
            End Using
        End Function

        Public Sub UpdateLastLogin(userId As Integer, loggedInAt As DateTime) Implements IUserRepository.UpdateLastLogin
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText = "UPDATE Users SET LastLoginAt = @LastLoginAt, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                    command.AddParameter("@LastLoginAt", SqliteDateHelper.ToStorageDateTime(loggedInAt))
                    command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    command.AddParameter("@Id", userId)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Function ParseRole(roleText As String) As UserRole
            If String.Equals(roleText, "Admin", StringComparison.OrdinalIgnoreCase) Then
                Return UserRole.Admin
            End If

            Return UserRole.Staff
        End Function

    End Class

End Namespace
