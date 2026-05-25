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

        Public Function GetAll() As IReadOnlyList(Of UserAccount) Implements IUserRepository.GetAll
            Dim users As New List(Of UserAccount)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT Id, Username, PasswordHash, PasswordSalt, FullName, Role, IsActive, LastLoginAt " &
                        "FROM Users ORDER BY Username;"

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            users.Add(MapUser(reader))
                        End While
                    End Using
                End Using
            End Using

            Return users
        End Function

        Public Function GetById(userId As Integer) As UserAccount Implements IUserRepository.GetById
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT Id, Username, PasswordHash, PasswordSalt, FullName, Role, IsActive, LastLoginAt " &
                        "FROM Users WHERE Id = @Id LIMIT 1;"
                    command.AddParameter("@Id", userId)

                    Using reader = command.ExecuteReader()
                        If reader.Read() Then
                            Return MapUser(reader)
                        End If
                    End Using
                End Using
            End Using

            Return Nothing
        End Function

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

                        Return MapUser(reader)
                    End Using
                End Using
            End Using
        End Function

        Public Function SaveUser(user As UserAccount) As Integer Implements IUserRepository.SaveUser
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    If user.Id <= 0 Then
                        command.CommandText =
                            "INSERT INTO Users (Username, PasswordHash, PasswordSalt, FullName, Role, IsActive, LastLoginAt, CreatedAt, UpdatedAt) " &
                            "VALUES (@Username, @PasswordHash, @PasswordSalt, @FullName, @Role, @IsActive, @LastLoginAt, @CreatedAt, @UpdatedAt);" &
                            "SELECT last_insert_rowid();"
                        command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    Else
                        command.CommandText =
                            "UPDATE Users SET Username = @Username, FullName = @FullName, Role = @Role, IsActive = @IsActive, " &
                            "LastLoginAt = @LastLoginAt, UpdatedAt = @UpdatedAt WHERE Id = @Id;" &
                            "SELECT @Id;"
                        command.AddParameter("@Id", user.Id)
                    End If

                    command.AddParameter("@Username", user.Username.Trim())
                    command.AddParameter("@PasswordHash", user.PasswordHash)
                    command.AddParameter("@PasswordSalt", user.PasswordSalt)
                    command.AddParameter("@FullName", user.FullName.Trim())
                    command.AddParameter("@Role", GetRoleText(user.Role))
                    command.AddParameter("@IsActive", If(user.IsActive, 1, 0))
                    command.AddParameter("@LastLoginAt", If(user.LastLoginAt.HasValue, CType(SqliteDateHelper.ToStorageDateTime(user.LastLoginAt.Value), Object), DBNull.Value))
                    command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
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

        Public Sub UpdatePassword(userId As Integer, passwordHash As String, passwordSalt As String) Implements IUserRepository.UpdatePassword
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "UPDATE Users SET PasswordHash = @PasswordHash, PasswordSalt = @PasswordSalt, UpdatedAt = @UpdatedAt " &
                        "WHERE Id = @Id;"
                    command.AddParameter("@PasswordHash", passwordHash)
                    command.AddParameter("@PasswordSalt", passwordSalt)
                    command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    command.AddParameter("@Id", userId)
                    command.ExecuteNonQuery()
                End Using
            End Using
        End Sub

        Private Function MapUser(reader As DbDataReader) As UserAccount
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
        End Function

        Private Function ParseRole(roleText As String) As UserRole
            If String.Equals(roleText, "Admin", StringComparison.OrdinalIgnoreCase) Then
                Return UserRole.Admin
            End If

            Return UserRole.Staff
        End Function

        Private Function GetRoleText(role As UserRole) As String
            If role = UserRole.Admin Then
                Return "Admin"
            End If

            Return "Staff"
        End Function

    End Class

End Namespace
