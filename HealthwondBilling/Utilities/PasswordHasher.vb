Imports System.Security.Cryptography

Namespace Utilities

    Public Structure PasswordHashResult
        Public Sub New(hashValue As String, saltValue As String)
            Hash = hashValue
            Salt = saltValue
        End Sub

        Public Property Hash As String
        Public Property Salt As String
    End Structure

    Public NotInheritable Class PasswordHasher

        Private Const SaltSize As Integer = 16
        Private Const HashSize As Integer = 32
        Private Const IterationCount As Integer = 100000

        Private Sub New()
        End Sub

        Public Shared Function HashPassword(password As String) As PasswordHashResult
            Dim saltBytes(SaltSize - 1) As Byte
            Using generator As RandomNumberGenerator = RandomNumberGenerator.Create()
                generator.GetBytes(saltBytes)
            End Using

            Dim hashBytes() As Byte
            Using deriveBytes As New Rfc2898DeriveBytes(password, saltBytes, IterationCount)
                hashBytes = deriveBytes.GetBytes(HashSize)
            End Using

            Return New PasswordHashResult(Convert.ToBase64String(hashBytes), Convert.ToBase64String(saltBytes))
        End Function

        Public Shared Function VerifyPassword(password As String, storedHash As String, storedSalt As String) As Boolean
            If String.IsNullOrWhiteSpace(password) OrElse String.IsNullOrWhiteSpace(storedHash) OrElse String.IsNullOrWhiteSpace(storedSalt) Then
                Return False
            End If

            Try
                Dim expectedHash As Byte() = Convert.FromBase64String(storedHash)
                Dim saltBytes As Byte() = Convert.FromBase64String(storedSalt)

                Dim actualHash As Byte()
                Using deriveBytes As New Rfc2898DeriveBytes(password, saltBytes, IterationCount)
                    actualHash = deriveBytes.GetBytes(HashSize)
                End Using

                Return FixedTimeEquals(expectedHash, actualHash)
            Catch ex As FormatException
                AppLogger.Warn("Password verification skipped because stored hash data is invalid.")
                Return False
            End Try
        End Function

        Private Shared Function FixedTimeEquals(left As Byte(), right As Byte()) As Boolean
            If left.Length <> right.Length Then
                Return False
            End If

            Dim difference As Integer = 0
            For index As Integer = 0 To left.Length - 1
                difference = difference Or (left(index) Xor right(index))
            Next

            Return difference = 0
        End Function

    End Class

End Namespace
