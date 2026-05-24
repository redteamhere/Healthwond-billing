Imports System.Data.Common
Imports System.Data.SQLite

Namespace Database

    Public Class SqliteConnectionFactory
        Implements IDbConnectionFactory

        Private ReadOnly _connectionString As String

        Public Sub New(databaseFilePath As String)
            _connectionString = $"Data Source={databaseFilePath};Version=3;Pooling=True;Journal Mode=WAL;Synchronous=NORMAL;"
        End Sub

        Public Function CreateOpenConnection() As DbConnection Implements IDbConnectionFactory.CreateOpenConnection
            Dim connection As New SQLiteConnection(_connectionString)
            connection.Open()

            Using command As DbCommand = connection.CreateCommand()
                command.CommandText = "PRAGMA foreign_keys = ON; PRAGMA busy_timeout = 5000;"
                command.ExecuteNonQuery()
            End Using

            Return connection
        End Function
    End Class

End Namespace
