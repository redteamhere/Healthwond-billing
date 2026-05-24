Imports HealthwondBilling.Utilities
Imports System.Data.SQLite
Imports System.IO

Namespace Database

    Public Class DatabaseBootstrapper

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Sub Initialize()
            Dim databaseCreated As Boolean = False

            If Not File.Exists(AppPaths.DatabaseFilePath) Then
                SQLiteConnection.CreateFile(AppPaths.DatabaseFilePath)
                databaseCreated = True
            End If

            Dim schemaPath As String = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Database", "schema.sql")
            If Not File.Exists(schemaPath) Then
                Throw New FileNotFoundException("The SQLite schema file was not found.", schemaPath)
            End If

            Dim schemaScript As String = File.ReadAllText(schemaPath)
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText = schemaScript
                    command.ExecuteNonQuery()
                End Using
            End Using

            Dim seedDataService As New SeedDataService(_connectionFactory)
            seedDataService.Seed()

            If databaseCreated Then
                AppLogger.Info($"A new SQLite database was created at '{AppPaths.DatabaseFilePath}'.")
            End If

            AppLogger.Info("Database bootstrap completed.")
        End Sub

    End Class

End Namespace
