Imports System.Data.Common

Namespace Database

    Public Interface IDbConnectionFactory
        Function CreateOpenConnection() As DbConnection
    End Interface

End Namespace
