Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports System.Data.Common
Imports System.Globalization

Namespace Services

    Public Class DashboardService

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Async Function GetSummaryAsync() As Task(Of DashboardSummary)
            Return Await Task.Run(Function() GetSummary())
        End Function

        Private Function GetSummary() As DashboardSummary
            Dim summary As New DashboardSummary()

            Using connection = _connectionFactory.CreateOpenConnection()
                Dim lowStockThreshold As Integer = ExecuteIntScalar(connection, "SELECT COALESCE(CAST(SettingValue AS INTEGER), 10) FROM Settings WHERE SettingKey = 'LowStockThreshold' LIMIT 1;", 10)

                summary.TodaySales = ExecuteDecimalScalar(connection, "SELECT COALESCE(SUM(NetAmount), 0) FROM Invoices WHERE date(InvoiceDate) = date('now', 'localtime');")
                summary.TotalStockUnits = ExecuteIntScalar(connection, "SELECT COALESCE(SUM(CurrentStock), 0) FROM Products WHERE IsDeleted = 0;", 0)
                summary.ExpiryAlerts = ExecuteIntScalar(connection, "SELECT COUNT(1) FROM Products WHERE IsDeleted = 0 AND CurrentStock > 0 AND date(ExpiryDate) <= date('now', '+60 day');", 0)
                summary.LowStockAlerts = ExecuteIntScalar(connection, "SELECT COUNT(1) FROM Products WHERE IsDeleted = 0 AND CurrentStock <= @Threshold;", 0, Sub(command) command.AddParameter("@Threshold", lowStockThreshold))
                summary.PendingPayments = ExecuteDecimalScalar(connection, "SELECT COALESCE(SUM(OutstandingBalance), 0) FROM Customers;")
            End Using

            Return summary
        End Function

        Private Function ExecuteIntScalar(connection As DbConnection, sql As String, defaultValue As Integer, Optional parameterize As Action(Of DbCommand) = Nothing) As Integer
            Using command = connection.CreateCommand()
                command.CommandText = sql
                parameterize?.Invoke(command)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return defaultValue
                End If

                Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function ExecuteDecimalScalar(connection As DbConnection, sql As String) As Decimal
            Using command = connection.CreateCommand()
                command.CommandText = sql
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return 0D
                End If

                Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
            End Using
        End Function

    End Class

End Namespace
