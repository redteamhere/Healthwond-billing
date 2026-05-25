Imports HealthwondBilling.Utilities
Imports System.Data.SQLite
Imports System.Data.Common
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

                EnsureSchemaCompatibility(connection)
            End Using

            Dim seedDataService As New SeedDataService(_connectionFactory)
            seedDataService.Seed()

            If databaseCreated Then
                AppLogger.Info($"A new SQLite database was created at '{AppPaths.DatabaseFilePath}'.")
            End If

            AppLogger.Info("Database bootstrap completed.")
        End Sub

        Private Sub EnsureSchemaCompatibility(connection As DbConnection)
            EnsureColumnExists(connection, "Purchases", "SupplierInvoiceDate", "TEXT NULL")
            EnsureColumnExists(connection, "Purchases", "PurchaseOrderNumber", "TEXT NULL")
            EnsureColumnExists(connection, "Purchases", "PurchaseOrderDate", "TEXT NULL")
            EnsureColumnExists(connection, "Purchases", "PlaceOfSupply", "TEXT NULL")
            EnsureColumnExists(connection, "Purchases", "CaseCount", "INTEGER NOT NULL DEFAULT 0")
            EnsureColumnExists(connection, "Purchases", "TransportName", "TEXT NULL")
            EnsureColumnExists(connection, "Purchases", "EwayBillNumber", "TEXT NULL")

            EnsureColumnExists(connection, "PurchaseItems", "ProductName", "TEXT NULL")
            EnsureColumnExists(connection, "PurchaseItems", "Packing", "TEXT NULL")
            EnsureColumnExists(connection, "PurchaseItems", "HsnCode", "TEXT NULL")
        End Sub

        Private Sub EnsureColumnExists(connection As DbConnection, tableName As String, columnName As String, columnDefinition As String)
            Using checkCommand = connection.CreateCommand()
                checkCommand.CommandText = $"PRAGMA table_info({tableName});"

                Using reader = checkCommand.ExecuteReader()
                    While reader.Read()
                        If String.Equals(Convert.ToString(reader("name")), columnName, StringComparison.OrdinalIgnoreCase) Then
                            Return
                        End If
                    End While
                End Using
            End Using

            Using alterCommand = connection.CreateCommand()
                alterCommand.CommandText = $"ALTER TABLE {tableName} ADD COLUMN {columnName} {columnDefinition};"
                alterCommand.ExecuteNonQuery()
            End Using
        End Sub

    End Class

End Namespace
