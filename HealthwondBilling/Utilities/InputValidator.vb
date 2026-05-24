Imports System.Net.Mail

Namespace Utilities

    Public NotInheritable Class InputValidator

        Private Sub New()
        End Sub

        Public Shared Function IsRequiredTextProvided(value As String) As Boolean
            Return Not String.IsNullOrWhiteSpace(value)
        End Function

        Public Shared Function IsValidEmail(value As String) As Boolean
            If String.IsNullOrWhiteSpace(value) Then
                Return False
            End If

            Try
                Dim address As New MailAddress(value)
                Return String.Equals(address.Address, value.Trim(), StringComparison.OrdinalIgnoreCase)
            Catch
                Return False
            End Try
        End Function

    End Class

End Namespace
