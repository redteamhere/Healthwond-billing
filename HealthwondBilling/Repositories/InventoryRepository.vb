Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class InventoryRepository
        Implements IInventoryRepository

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetInventorySummary(expiryWindowInDays As Integer) As InventorySummary Implements IInventoryRepository.GetInventorySummary
            Dim summary As New InventorySummary()

            Using connection = _connectionFactory.CreateOpenConnection()
                Dim lowStockThreshold As Integer = GetLowStockThreshold(connection)

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT " &
                        "COUNT(DISTINCT ProductName || '|' || COALESCE(CompanyName, '') || '|' || COALESCE(Packing, '')) AS DistinctProducts, " &
                        "COUNT(1) AS BatchCount, " &
                        "COALESCE(SUM(CurrentStock), 0) AS TotalUnits, " &
                        "COALESCE(SUM(CASE WHEN CurrentStock > 0 AND date(ExpiryDate) <= date('now', '+' || @ExpiryWindow || ' day') THEN 1 ELSE 0 END), 0) AS ExpiringSoonCount, " &
                        "COALESCE(SUM(CASE WHEN CurrentStock <= @LowStockThreshold THEN 1 ELSE 0 END), 0) AS LowStockCount " &
                        "FROM Products WHERE IsDeleted = 0;"
                    command.AddParameter("@ExpiryWindow", Math.Max(expiryWindowInDays, 0))
                    command.AddParameter("@LowStockThreshold", lowStockThreshold)

                    Using reader = command.ExecuteReader()
                        If reader.Read() Then
                            summary.DistinctProducts = Convert.ToInt32(reader("DistinctProducts"), CultureInfo.InvariantCulture)
                            summary.BatchCount = Convert.ToInt32(reader("BatchCount"), CultureInfo.InvariantCulture)
                            summary.TotalUnits = Convert.ToInt32(reader("TotalUnits"), CultureInfo.InvariantCulture)
                            summary.ExpiringSoonCount = Convert.ToInt32(reader("ExpiringSoonCount"), CultureInfo.InvariantCulture)
                            summary.LowStockCount = Convert.ToInt32(reader("LowStockCount"), CultureInfo.InvariantCulture)
                        End If
                    End Using
                End Using
            End Using

            Return summary
        End Function

        Public Function GetCurrentStock(searchTerm As String) As List(Of InventoryCurrentStockRow) Implements IInventoryRepository.GetCurrentStock
            Dim rows As New List(Of InventoryCurrentStockRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Dim lowStockThreshold As Integer = GetLowStockThreshold(connection)

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT ProductName, COALESCE(CompanyName, '') AS CompanyName, COALESCE(Packing, '') AS Packing, COALESCE(Composition, '') AS Composition, " &
                        "COUNT(1) AS BatchCount, COALESCE(SUM(CurrentStock), 0) AS TotalStock, MIN(ExpiryDate) AS EarliestExpiryDate, MAX(ExpiryDate) AS LatestExpiryDate, " &
                        "COALESCE(SUM(CurrentStock * PTR), 0) AS StockValueAtPTR, COALESCE(SUM(CurrentStock * PTS), 0) AS StockValueAtPTS " &
                        "FROM Products " &
                        "WHERE IsDeleted = 0 AND " & BuildProductSearchCondition() & " " &
                        "GROUP BY ProductName, COALESCE(CompanyName, ''), COALESCE(Packing, ''), COALESCE(Composition, '') " &
                        "ORDER BY ProductName ASC, CompanyName ASC;"
                    AddSearchParameters(command, searchTerm)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            Dim earliestExpiryDate As DateTime = ParseDate(reader("EarliestExpiryDate"))
                            Dim totalStock As Integer = Convert.ToInt32(reader("TotalStock"), CultureInfo.InvariantCulture)

                            rows.Add(New InventoryCurrentStockRow With {
                                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                                .CompanyName = Convert.ToString(reader("CompanyName"), CultureInfo.InvariantCulture),
                                .Packing = Convert.ToString(reader("Packing"), CultureInfo.InvariantCulture),
                                .Composition = Convert.ToString(reader("Composition"), CultureInfo.InvariantCulture),
                                .BatchCount = Convert.ToInt32(reader("BatchCount"), CultureInfo.InvariantCulture),
                                .TotalStock = totalStock,
                                .EarliestExpiryDate = earliestExpiryDate,
                                .LatestExpiryDate = ParseDate(reader("LatestExpiryDate")),
                                .StockValueAtPTR = Convert.ToDecimal(reader("StockValueAtPTR"), CultureInfo.InvariantCulture),
                                .StockValueAtPTS = Convert.ToDecimal(reader("StockValueAtPTS"), CultureInfo.InvariantCulture),
                                .StockStatus = ResolveStockStatus(totalStock, lowStockThreshold, earliestExpiryDate)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetBatchStock(searchTerm As String) As List(Of InventoryBatchStockRow) Implements IInventoryRepository.GetBatchStock
            Dim rows As New List(Of InventoryBatchStockRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Dim lowStockThreshold As Integer = GetLowStockThreshold(connection)

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT ProductName, BatchNumber, ExpiryDate, COALESCE(CompanyName, '') AS CompanyName, COALESCE(Packing, '') AS Packing, CurrentStock, GstPercentage, MRP, PTR, PTS, COALESCE(Barcode, '') AS Barcode " &
                        "FROM Products WHERE IsDeleted = 0 AND " & BuildProductSearchCondition() & " " &
                        "ORDER BY ProductName ASC, date(ExpiryDate) ASC, BatchNumber ASC;"
                    AddSearchParameters(command, searchTerm)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            Dim expiryDate As DateTime = ParseDate(reader("ExpiryDate"))
                            Dim currentStock As Integer = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture)

                            rows.Add(New InventoryBatchStockRow With {
                                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                                .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                                .ExpiryDate = expiryDate,
                                .CompanyName = Convert.ToString(reader("CompanyName"), CultureInfo.InvariantCulture),
                                .Packing = Convert.ToString(reader("Packing"), CultureInfo.InvariantCulture),
                                .CurrentStock = currentStock,
                                .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                                .MRP = Convert.ToDecimal(reader("MRP"), CultureInfo.InvariantCulture),
                                .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                                .PTS = Convert.ToDecimal(reader("PTS"), CultureInfo.InvariantCulture),
                                .Barcode = Convert.ToString(reader("Barcode"), CultureInfo.InvariantCulture),
                                .StockStatus = ResolveStockStatus(currentStock, lowStockThreshold, expiryDate)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetExpiryStock(searchTerm As String, expiryWindowInDays As Integer) As List(Of InventoryExpiryRow) Implements IInventoryRepository.GetExpiryStock
            Dim rows As New List(Of InventoryExpiryRow)()
            Dim safeWindow As Integer = Math.Max(expiryWindowInDays, 0)

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT ProductName, BatchNumber, ExpiryDate, COALESCE(CompanyName, '') AS CompanyName, CurrentStock, MRP, PTR " &
                        "FROM Products " &
                        "WHERE IsDeleted = 0 AND CurrentStock > 0 AND date(ExpiryDate) <= date('now', '+' || @ExpiryWindow || ' day') AND " & BuildProductSearchCondition() & " " &
                        "ORDER BY date(ExpiryDate) ASC, ProductName ASC, BatchNumber ASC;"
                    command.AddParameter("@ExpiryWindow", safeWindow)
                    AddSearchParameters(command, searchTerm)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            Dim expiryDate As DateTime = ParseDate(reader("ExpiryDate"))
                            Dim daysToExpiry As Integer = CInt((expiryDate.Date - DateTime.Today).TotalDays)

                            rows.Add(New InventoryExpiryRow With {
                                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                                .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                                .ExpiryDate = expiryDate,
                                .DaysToExpiry = daysToExpiry,
                                .CompanyName = Convert.ToString(reader("CompanyName"), CultureInfo.InvariantCulture),
                                .CurrentStock = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture),
                                .MRP = Convert.ToDecimal(reader("MRP"), CultureInfo.InvariantCulture),
                                .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                                .StockStatus = ResolveExpiryStatus(daysToExpiry)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetLowStock(searchTerm As String) As List(Of InventoryLowStockRow) Implements IInventoryRepository.GetLowStock
            Dim rows As New List(Of InventoryLowStockRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Dim lowStockThreshold As Integer = GetLowStockThreshold(connection)

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT ProductName, BatchNumber, ExpiryDate, COALESCE(CompanyName, '') AS CompanyName, CurrentStock " &
                        "FROM Products " &
                        "WHERE IsDeleted = 0 AND CurrentStock <= @LowStockThreshold AND " & BuildProductSearchCondition() & " " &
                        "ORDER BY CurrentStock ASC, ProductName ASC, BatchNumber ASC;"
                    command.AddParameter("@LowStockThreshold", lowStockThreshold)
                    AddSearchParameters(command, searchTerm)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            Dim currentStock As Integer = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture)
                            Dim expiryDate As DateTime = ParseDate(reader("ExpiryDate"))

                            rows.Add(New InventoryLowStockRow With {
                                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                                .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                                .CompanyName = Convert.ToString(reader("CompanyName"), CultureInfo.InvariantCulture),
                                .CurrentStock = currentStock,
                                .ReorderThreshold = lowStockThreshold,
                                .ShortageUnits = Math.Max(lowStockThreshold - currentStock, 0),
                                .ExpiryDate = expiryDate,
                                .StockStatus = ResolveStockStatus(currentStock, lowStockThreshold, expiryDate)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetStockLedger(searchTerm As String, fromDate As DateTime, toDate As DateTime) As List(Of InventoryLedgerRow) Implements IInventoryRepository.GetStockLedger
            Dim rows As New List(Of InventoryLedgerRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT sl.TransactionDate, p.ProductName, sl.BatchNumber, sl.TransactionType, sl.ReferenceType, sl.ReferenceId, " &
                        "sl.QuantityIn, sl.QuantityOut, sl.BalanceQuantity, sl.UnitCost, COALESCE(sl.Remarks, '') AS Remarks " &
                        "FROM StockLedger sl " &
                        "INNER JOIN Products p ON p.Id = sl.ProductId " &
                        "WHERE date(sl.TransactionDate) BETWEEN date(@FromDate) AND date(@ToDate) AND " &
                        "(@Search = '' OR p.ProductName LIKE @SearchLike OR sl.BatchNumber LIKE @SearchLike OR sl.TransactionType LIKE @SearchLike OR sl.ReferenceType LIKE @SearchLike OR sl.Remarks LIKE @SearchLike) " &
                        "ORDER BY datetime(sl.TransactionDate) DESC, sl.Id DESC;"
                    command.AddParameter("@FromDate", fromDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    command.AddParameter("@ToDate", toDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    AddSearchParameters(command, searchTerm)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New InventoryLedgerRow With {
                                .TransactionDate = ParseDateTime(reader("TransactionDate")),
                                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                                .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                                .TransactionType = Convert.ToString(reader("TransactionType"), CultureInfo.InvariantCulture),
                                .ReferenceType = Convert.ToString(reader("ReferenceType"), CultureInfo.InvariantCulture),
                                .ReferenceId = Convert.ToInt32(reader("ReferenceId"), CultureInfo.InvariantCulture),
                                .QuantityIn = Convert.ToInt32(reader("QuantityIn"), CultureInfo.InvariantCulture),
                                .QuantityOut = Convert.ToInt32(reader("QuantityOut"), CultureInfo.InvariantCulture),
                                .BalanceQuantity = Convert.ToInt32(reader("BalanceQuantity"), CultureInfo.InvariantCulture),
                                .UnitCost = Convert.ToDecimal(reader("UnitCost"), CultureInfo.InvariantCulture),
                                .Remarks = Convert.ToString(reader("Remarks"), CultureInfo.InvariantCulture)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Private Function BuildProductSearchCondition() As String
            Return "(@Search = '' OR ProductName LIKE @SearchLike OR BatchNumber LIKE @SearchLike OR CompanyName LIKE @SearchLike OR Composition LIKE @SearchLike OR Barcode LIKE @SearchLike OR Packing LIKE @SearchLike)"
        End Function

        Private Sub AddSearchParameters(command As DbCommand, searchTerm As String)
            Dim safeSearchTerm As String = If(searchTerm, String.Empty).Trim()
            command.AddParameter("@Search", safeSearchTerm)
            command.AddParameter("@SearchLike", $"%{safeSearchTerm}%")
        End Sub

        Private Function GetLowStockThreshold(connection As DbConnection) As Integer
            Using command = connection.CreateCommand()
                command.CommandText = "SELECT COALESCE(CAST(SettingValue AS INTEGER), 10) FROM Settings WHERE SettingKey = 'LowStockThreshold' LIMIT 1;"
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return 10
                End If

                Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function ResolveStockStatus(currentStock As Integer, lowStockThreshold As Integer, expiryDate As DateTime) As String
            If currentStock <= 0 Then
                Return "Out of Stock"
            End If

            If expiryDate.Date < DateTime.Today Then
                Return "Expired"
            End If

            If expiryDate.Date <= DateTime.Today.AddDays(60) Then
                Return "Expiring Soon"
            End If

            If currentStock <= lowStockThreshold Then
                Return "Low Stock"
            End If

            Return "Healthy"
        End Function

        Private Function ResolveExpiryStatus(daysToExpiry As Integer) As String
            If daysToExpiry < 0 Then
                Return "Expired"
            End If

            Return "Expiring Soon"
        End Function

        Private Function ParseDate(value As Object) As DateTime
            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return DateTime.Today
        End Function

        Private Function ParseDateTime(value As Object) As DateTime
            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return DateTime.Now
        End Function

    End Class

End Namespace
