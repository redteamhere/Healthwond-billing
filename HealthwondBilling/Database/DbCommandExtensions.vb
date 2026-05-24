Imports System.Data.Common
Imports System.Runtime.CompilerServices

Namespace Database

    Public Module DbCommandExtensions

        <Extension()>
        Public Sub AddParameter(command As DbCommand, parameterName As String, value As Object)
            Dim parameter As DbParameter = command.CreateParameter()
            parameter.ParameterName = parameterName
            parameter.Value = If(value Is Nothing, DirectCast(DBNull.Value, Object), value)
            command.Parameters.Add(parameter)
        End Sub

    End Module

End Namespace
