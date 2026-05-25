Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.SQLite
Imports System.Globalization
Imports System.IO

Namespace Services

    Public Class MaintenanceService

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Async Function CreateDatabaseBackupAsync() As Task(Of FileOperationResult)
            Return Await Task.Run(
                Function()
                    Try
                        AppPaths.EnsureDirectories()

                        Dim backupFilePath As String = Path.Combine(
                            AppPaths.BackupsDirectory,
                            $"healthwond-backup-{DateTime.Now:yyyyMMdd-HHmmss}.db")

                        Using sourceConnection As SQLiteConnection = TryCast(_connectionFactory.CreateOpenConnection(), SQLiteConnection)
                            If sourceConnection Is Nothing Then
                                Throw New InvalidOperationException("The SQLite connection could not be opened for backup.")
                            End If

                            Using destinationConnection As New SQLiteConnection($"Data Source={backupFilePath};Version=3;")
                                destinationConnection.Open()
                                sourceConnection.BackupDatabase(destinationConnection, "main", "main", -1, Nothing, 0)
                            End Using
                        End Using

                        AppLogger.Info($"Database backup created at '{backupFilePath}'.")
                        Return FileOperationResult.Success("Database backup created successfully.", backupFilePath)
                    Catch ex As Exception
                        AppLogger.Error("Database backup failed.", ex)
                        Return FileOperationResult.Failure("The database backup could not be created.")
                    End Try
                End Function)
        End Function

        Public Async Function RestoreDatabaseBackupAsync(backupFilePath As String) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    Dim normalizedBackupPath As String = If(backupFilePath, String.Empty).Trim()
                    If normalizedBackupPath = String.Empty OrElse Not File.Exists(normalizedBackupPath) Then
                        Return EntityOperationResult.Failure("Select a valid backup file to restore.")
                    End If

                    Try
                        ValidateBackupFile(normalizedBackupPath)

                        AppPaths.EnsureDirectories()
                        SQLiteConnection.ClearAllPools()

                        Dim safeguardPath As String = Path.Combine(
                            AppPaths.BackupsDirectory,
                            $"pre-restore-{DateTime.Now:yyyyMMdd-HHmmss}.db")

                        If File.Exists(AppPaths.DatabaseFilePath) Then
                            File.Copy(AppPaths.DatabaseFilePath, safeguardPath, True)
                        End If

                        DeleteSidecarFile(AppPaths.DatabaseFilePath & "-wal")
                        DeleteSidecarFile(AppPaths.DatabaseFilePath & "-shm")
                        File.Copy(normalizedBackupPath, AppPaths.DatabaseFilePath, True)

                        ValidateBackupFile(AppPaths.DatabaseFilePath)
                        AppLogger.Info($"Database restored from backup '{normalizedBackupPath}'.")
                        Return EntityOperationResult.Success("Database backup restored successfully. Restart the application to continue with the restored data.")
                    Catch ex As Exception
                        AppLogger.Error("Database restore failed.", ex)
                        Return EntityOperationResult.Failure("The selected backup could not be restored.")
                    End Try
                End Function)
        End Function

        Public Async Function OptimizeDatabaseAsync() As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    Try
                        Using connection As SQLiteConnection = TryCast(_connectionFactory.CreateOpenConnection(), SQLiteConnection)
                            If connection Is Nothing Then
                                Throw New InvalidOperationException("The SQLite connection could not be opened for optimization.")
                            End If

                            Using command As SQLiteCommand = connection.CreateCommand()
                                command.CommandText = "PRAGMA wal_checkpoint(FULL); VACUUM; ANALYZE;"
                                command.ExecuteNonQuery()
                            End Using
                        End Using

                        AppLogger.Info("Database optimization completed.")
                        Return EntityOperationResult.Success("Database optimization completed successfully.")
                    Catch ex As Exception
                        AppLogger.Error("Database optimization failed.", ex)
                        Return EntityOperationResult.Failure("The database could not be optimized.")
                    End Try
                End Function)
        End Function

        Public Function GetDefaultBackupDirectory() As String
            Return AppPaths.BackupsDirectory
        End Function

        Private Sub ValidateBackupFile(filePath As String)
            Using connection As New SQLiteConnection($"Data Source={filePath};Version=3;Read Only=True;")
                connection.Open()

                Using command As SQLiteCommand = connection.CreateCommand()
                    command.CommandText = "PRAGMA integrity_check;"
                    Dim result As Object = command.ExecuteScalar()
                    Dim integrityMessage As String = Convert.ToString(result, CultureInfo.InvariantCulture)
                    If Not String.Equals(integrityMessage, "ok", StringComparison.OrdinalIgnoreCase) Then
                        Throw New InvalidDataException("The selected SQLite database failed integrity validation.")
                    End If
                End Using
            End Using
        End Sub

        Private Sub DeleteSidecarFile(filePath As String)
            If File.Exists(filePath) Then
                File.Delete(filePath)
            End If
        End Sub

    End Class

End Namespace
