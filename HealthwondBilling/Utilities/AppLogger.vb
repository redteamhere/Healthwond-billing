Imports System.IO
Imports System.Text

Namespace Utilities

    Public NotInheritable Class AppLogger

        Private Shared ReadOnly SyncRoot As New Object()

        Private Sub New()
        End Sub

        Public Shared Sub Info(message As String)
            WriteEntry("INFO", message, Nothing)
        End Sub

        Public Shared Sub Warn(message As String)
            WriteEntry("WARN", message, Nothing)
        End Sub

        Public Shared Sub [Error](message As String, Optional ex As Exception = Nothing)
            WriteEntry("ERROR", message, ex)
        End Sub

        Private Shared Sub WriteEntry(level As String, message As String, ex As Exception)
            Try
                AppPaths.EnsureDirectories()

                Dim logFilePath As String = Path.Combine(AppPaths.LogsDirectory, $"application-{DateTime.Today:yyyyMMdd}.log")
                Dim builder As New StringBuilder()
                builder.AppendLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}")

                If ex IsNot Nothing Then
                    builder.AppendLine(ex.ToString())
                End If

                SyncLock SyncRoot
                    File.AppendAllText(logFilePath, builder.ToString(), Encoding.UTF8)
                End SyncLock
            Catch
                ' Logging must never throw back into the calling workflow.
            End Try
        End Sub

    End Class

End Namespace
