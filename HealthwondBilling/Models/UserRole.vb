Imports System.Runtime.CompilerServices

Namespace Models

    Public Enum UserRole
        Admin = 1
        Staff = 2
    End Enum

    Public Module UserRoleExtensions

        <Extension()>
        Public Function ToFriendlyText(role As UserRole) As String
            Select Case role
                Case UserRole.Admin
                    Return "Administrator"
                Case UserRole.Staff
                    Return "Staff"
                Case Else
                    Return "Unknown"
            End Select
        End Function

    End Module

End Namespace
