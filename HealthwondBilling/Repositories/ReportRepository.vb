Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class ReportRepository
        Implements IReportRepository

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GetSalesReport(fromDate As DateTime, toDate As DateTime) As List(Of SalesReportRow) Implements IReportRepository.GetSalesReport
            Dim rows As New List(Of SalesReportRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT i.InvoiceDate, i.InvoiceNumber, c.CustomerName, COALESCE(i.PaymentMode, '') AS PaymentMode, " &
                        "COUNT(ii.Id) AS LineCount, COALESCE(SUM(ii.Quantity + ii.FreeQuantity), 0) AS TotalUnits, " &
                        "COALESCE(SUM(ii.TaxableAmount), 0) AS TaxableAmount, i.GstAmount, i.NetAmount, i.AmountPaid, i.BalanceAmount " &
                        "FROM Invoices i " &
                        "INNER JOIN Customers c ON c.Id = i.CustomerId " &
                        "LEFT JOIN InvoiceItems ii ON ii.InvoiceId = i.Id " &
                        "WHERE date(i.InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "GROUP BY i.Id, i.InvoiceDate, i.InvoiceNumber, c.CustomerName, i.PaymentMode, i.GstAmount, i.NetAmount, i.AmountPaid, i.BalanceAmount " &
                        "ORDER BY date(i.InvoiceDate) DESC, i.InvoiceNumber DESC;"
                    AddDateParameters(command, fromDate, toDate)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New SalesReportRow With {
                                .InvoiceDate = ParseDate(reader("InvoiceDate")),
                                .InvoiceNumber = Convert.ToString(reader("InvoiceNumber"), CultureInfo.InvariantCulture),
                                .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                                .PaymentMode = Convert.ToString(reader("PaymentMode"), CultureInfo.InvariantCulture),
                                .LineCount = Convert.ToInt32(reader("LineCount"), CultureInfo.InvariantCulture),
                                .TotalUnits = Convert.ToInt32(reader("TotalUnits"), CultureInfo.InvariantCulture),
                                .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                                .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                                .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture),
                                .AmountPaid = Convert.ToDecimal(reader("AmountPaid"), CultureInfo.InvariantCulture),
                                .BalanceAmount = Convert.ToDecimal(reader("BalanceAmount"), CultureInfo.InvariantCulture)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetPurchaseReport(fromDate As DateTime, toDate As DateTime) As List(Of PurchaseReportRow) Implements IReportRepository.GetPurchaseReport
            Dim rows As New List(Of PurchaseReportRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT p.PurchaseDate, p.PurchaseNumber, s.SupplierName, COALESCE(p.SupplierInvoiceNumber, '') AS SupplierInvoiceNumber, " &
                        "COUNT(pi.Id) AS LineCount, COALESCE(SUM(pi.Quantity + pi.FreeQuantity), 0) AS TotalUnits, " &
                        "COALESCE(SUM(pi.TaxableAmount), 0) AS TaxableAmount, p.GstAmount, p.NetAmount " &
                        "FROM Purchases p " &
                        "INNER JOIN Suppliers s ON s.Id = p.SupplierId " &
                        "LEFT JOIN PurchaseItems pi ON pi.PurchaseId = p.Id " &
                        "WHERE date(p.PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "GROUP BY p.Id, p.PurchaseDate, p.PurchaseNumber, s.SupplierName, p.SupplierInvoiceNumber, p.GstAmount, p.NetAmount " &
                        "ORDER BY date(p.PurchaseDate) DESC, p.PurchaseNumber DESC;"
                    AddDateParameters(command, fromDate, toDate)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New PurchaseReportRow With {
                                .PurchaseDate = ParseDate(reader("PurchaseDate")),
                                .PurchaseNumber = Convert.ToString(reader("PurchaseNumber"), CultureInfo.InvariantCulture),
                                .SupplierName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                                .SupplierInvoiceNumber = Convert.ToString(reader("SupplierInvoiceNumber"), CultureInfo.InvariantCulture),
                                .LineCount = Convert.ToInt32(reader("LineCount"), CultureInfo.InvariantCulture),
                                .TotalUnits = Convert.ToInt32(reader("TotalUnits"), CultureInfo.InvariantCulture),
                                .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                                .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                                .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetGstReport(fromDate As DateTime, toDate As DateTime) As List(Of GstReportRow) Implements IReportRepository.GetGstReport
            Dim rows As New List(Of GstReportRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                rows.AddRange(GetSalesGstRows(connection, fromDate, toDate))
                rows.AddRange(GetPurchaseGstRows(connection, fromDate, toDate))
            End Using

            Return rows.OrderBy(Function(row) row.ReportSection).ThenBy(Function(row) row.GstPercentage).ToList()
        End Function

        Public Function GetStockReport() As List(Of StockReportRow) Implements IReportRepository.GetStockReport
            Dim rows As New List(Of StockReportRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Dim lowStockThreshold As Integer = GetLowStockThreshold(connection)

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT ProductName, BatchNumber, ExpiryDate, CompanyName, CurrentStock, GstPercentage, MRP, PTR, PTS " &
                        "FROM Products WHERE IsDeleted = 0 ORDER BY ProductName ASC, BatchNumber ASC;"

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            Dim expiryDate As DateTime = ParseDate(reader("ExpiryDate"))
                            Dim currentStock As Integer = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture)
                            rows.Add(New StockReportRow With {
                                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                                .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                                .ExpiryDate = expiryDate,
                                .CompanyName = ConvertNullableString(reader("CompanyName")),
                                .CurrentStock = currentStock,
                                .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                                .MRP = Convert.ToDecimal(reader("MRP"), CultureInfo.InvariantCulture),
                                .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                                .PTS = Convert.ToDecimal(reader("PTS"), CultureInfo.InvariantCulture),
                                .StockValueAtPTR = currentStock * Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                                .StockValueAtPTS = currentStock * Convert.ToDecimal(reader("PTS"), CultureInfo.InvariantCulture),
                                .StockStatus = ResolveStockStatus(currentStock, lowStockThreshold, expiryDate)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetCustomerOutstandingReport() As List(Of CustomerOutstandingReportRow) Implements IReportRepository.GetCustomerOutstandingReport
            Dim rows As New List(Of CustomerOutstandingReportRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT CustomerName, Gstin, DrugLicenseNumber, Phone, Email, OutstandingBalance " &
                        "FROM Customers WHERE OutstandingBalance > 0 ORDER BY OutstandingBalance DESC, CustomerName ASC;"

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New CustomerOutstandingReportRow With {
                                .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                                .Gstin = ConvertNullableString(reader("Gstin")),
                                .DrugLicenseNumber = ConvertNullableString(reader("DrugLicenseNumber")),
                                .Phone = ConvertNullableString(reader("Phone")),
                                .Email = ConvertNullableString(reader("Email")),
                                .OutstandingBalance = Convert.ToDecimal(reader("OutstandingBalance"), CultureInfo.InvariantCulture)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetProfitLossReport(fromDate As DateTime, toDate As DateTime) As ProfitLossReport Implements IReportRepository.GetProfitLossReport
            Dim report As New ProfitLossReport With {
                .FromDate = fromDate.Date,
                .ToDate = toDate.Date
            }

            Using connection = _connectionFactory.CreateOpenConnection()
                report.SalesTaxableAmount = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(ii.TaxableAmount), 0) FROM InvoiceItems ii INNER JOIN Invoices i ON i.Id = ii.InvoiceId WHERE date(i.InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                report.SalesNetAmount = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(NetAmount), 0) FROM Invoices WHERE date(InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                report.PurchaseTaxableAmount = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(pi.TaxableAmount), 0) FROM PurchaseItems pi INNER JOIN Purchases p ON p.Id = pi.PurchaseId WHERE date(p.PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                report.PurchaseNetAmount = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(NetAmount), 0) FROM Purchases WHERE date(PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                report.EstimatedCostOfGoodsSold = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(QuantityOut * UnitCost), 0) FROM StockLedger WHERE TransactionType = 'SALE' AND date(TransactionDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                report.OutstandingReceivables = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(OutstandingBalance), 0) FROM Customers;",
                    Nothing,
                    Nothing)
                report.OutstandingPayables = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(OutstandingBalance), 0) FROM Suppliers;",
                    Nothing,
                    Nothing)
            End Using

            report.EstimatedGrossProfit = report.SalesTaxableAmount - report.EstimatedCostOfGoodsSold
            report.GrossMarginPercentage = If(report.SalesTaxableAmount <= 0D, 0D, Decimal.Round((report.EstimatedGrossProfit / report.SalesTaxableAmount) * 100D, 2, MidpointRounding.AwayFromZero))

            Return report
        End Function

        Private Function GetSalesGstRows(connection As DbConnection, fromDate As DateTime, toDate As DateTime) As IEnumerable(Of GstReportRow)
            Dim rows As New List(Of GstReportRow)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT ii.GstPercentage, COALESCE(SUM(ii.TaxableAmount), 0) AS TaxableAmount, COALESCE(SUM(ii.GstAmount), 0) AS GstAmount, COUNT(DISTINCT ii.InvoiceId) AS InvoiceCount " &
                    "FROM InvoiceItems ii INNER JOIN Invoices i ON i.Id = ii.InvoiceId " &
                    "WHERE date(i.InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                    "GROUP BY ii.GstPercentage ORDER BY ii.GstPercentage ASC;"
                AddDateParameters(command, fromDate, toDate)

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New GstReportRow With {
                            .ReportSection = "Sales",
                            .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                            .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                            .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                            .InvoiceCount = Convert.ToInt32(reader("InvoiceCount"), CultureInfo.InvariantCulture)
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function GetPurchaseGstRows(connection As DbConnection, fromDate As DateTime, toDate As DateTime) As IEnumerable(Of GstReportRow)
            Dim rows As New List(Of GstReportRow)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT pi.GstPercentage, COALESCE(SUM(pi.TaxableAmount), 0) AS TaxableAmount, COALESCE(SUM(pi.GstAmount), 0) AS GstAmount, COUNT(DISTINCT pi.PurchaseId) AS InvoiceCount " &
                    "FROM PurchaseItems pi INNER JOIN Purchases p ON p.Id = pi.PurchaseId " &
                    "WHERE date(p.PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                    "GROUP BY pi.GstPercentage ORDER BY pi.GstPercentage ASC;"
                AddDateParameters(command, fromDate, toDate)

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New GstReportRow With {
                            .ReportSection = "Purchases",
                            .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                            .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                            .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                            .InvoiceCount = Convert.ToInt32(reader("InvoiceCount"), CultureInfo.InvariantCulture)
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function ExecuteDecimalScalar(connection As DbConnection, sql As String, fromDate As DateTime?, toDate As DateTime?) As Decimal
            Using command = connection.CreateCommand()
                command.CommandText = sql
                If fromDate.HasValue AndAlso toDate.HasValue Then
                    AddDateParameters(command, fromDate.Value, toDate.Value)
                End If

                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return 0D
                End If

                Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub AddDateParameters(command As DbCommand, fromDate As DateTime, toDate As DateTime)
            command.AddParameter("@FromDate", fromDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
            command.AddParameter("@ToDate", toDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
        End Sub

        Private Function ParseDate(value As Object) As DateTime
            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return DateTime.Today
        End Function

        Private Function ConvertNullableString(value As Object) As String
            If value Is Nothing OrElse value Is DBNull.Value Then
                Return String.Empty
            End If

            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        End Function

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

            If expiryDate.Date <= DateTime.Today.AddDays(60) Then
                Return "Expiring Soon"
            End If

            If currentStock <= lowStockThreshold Then
                Return "Low Stock"
            End If

            Return "Normal"
        End Function

    End Class

End Namespace
