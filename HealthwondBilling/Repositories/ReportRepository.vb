Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports System.Data.Common
Imports System.Globalization
Imports System.Linq

Namespace Repositories

    Public Class ReportRepository
        Implements IReportRepository

        Private NotInheritable Class AccountMasterSnapshot
            Public Property AccountId As Integer
            Public Property PartyName As String = String.Empty
            Public Property Gstin As String = String.Empty
            Public Property DrugLicenseNumber As String = String.Empty
            Public Property Phone As String = String.Empty
            Public Property OutstandingBalance As Decimal
        End Class

        Private NotInheritable Class AgingSourceDocument
            Public Property AccountId As Integer
            Public Property DocumentId As Integer
            Public Property DocumentDate As DateTime
            Public Property OpenAmount As Decimal
        End Class

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

        Public Function GetCustomerAgingReport(asOfDate As DateTime) As List(Of AgingReportRow) Implements IReportRepository.GetCustomerAgingReport
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim accounts As List(Of AccountMasterSnapshot) = LoadCustomerOutstandingAccounts(connection)
                Dim documents As List(Of AgingSourceDocument) = LoadCustomerAgingDocuments(connection, asOfDate)
                Dim paymentsByAccount As Dictionary(Of Integer, Decimal) = LoadCustomerPaymentsByAccount(connection, asOfDate)
                Return BuildAgingRows(accounts, documents, paymentsByAccount, asOfDate)
            End Using
        End Function

        Public Function GetSupplierAgingReport(asOfDate As DateTime) As List(Of AgingReportRow) Implements IReportRepository.GetSupplierAgingReport
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim accounts As List(Of AccountMasterSnapshot) = LoadSupplierOutstandingAccounts(connection)
                Dim documents As List(Of AgingSourceDocument) = LoadSupplierAgingDocuments(connection, asOfDate)
                Dim paymentsByAccount As Dictionary(Of Integer, Decimal) = LoadSupplierPaymentsByAccount(connection, asOfDate)
                Return BuildAgingRows(accounts, documents, paymentsByAccount, asOfDate)
            End Using
        End Function

        Public Function GetReportOverview(fromDate As DateTime, toDate As DateTime) As ReportOverview Implements IReportRepository.GetReportOverview
            Dim overview As New ReportOverview With {
                .FromDate = fromDate.Date,
                .ToDate = toDate.Date
            }

            Using connection = _connectionFactory.CreateOpenConnection()
                overview.SalesInvoiceCount = ExecuteIntegerScalar(connection,
                    "SELECT COUNT(1) FROM Invoices WHERE date(InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                overview.PurchaseBillCount = ExecuteIntegerScalar(connection,
                    "SELECT COUNT(1) FROM Purchases WHERE date(PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                overview.SalesUnits = ExecuteIntegerScalar(connection,
                    "SELECT COALESCE(SUM(ii.Quantity + ii.FreeQuantity), 0) FROM InvoiceItems ii INNER JOIN Invoices i ON i.Id = ii.InvoiceId WHERE date(i.InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                overview.PurchaseUnits = ExecuteIntegerScalar(connection,
                    "SELECT COALESCE(SUM(pi.Quantity + pi.FreeQuantity), 0) FROM PurchaseItems pi INNER JOIN Purchases p ON p.Id = pi.PurchaseId WHERE date(p.PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)

                Dim salesNetAmount As Decimal = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(NetAmount), 0) FROM Invoices WHERE date(InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                Dim purchaseNetAmount As Decimal = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(NetAmount), 0) FROM Purchases WHERE date(PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)

                overview.AverageSaleBillValue = If(overview.SalesInvoiceCount <= 0, 0D, Decimal.Round(salesNetAmount / overview.SalesInvoiceCount, 2, MidpointRounding.AwayFromZero))
                overview.AveragePurchaseBillValue = If(overview.PurchaseBillCount <= 0, 0D, Decimal.Round(purchaseNetAmount / overview.PurchaseBillCount, 2, MidpointRounding.AwayFromZero))
                overview.CustomerCollectionsAmount = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(Amount), 0) FROM CustomerPayments WHERE date(PaymentDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                overview.SupplierPaymentsAmount = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(Amount), 0) FROM SupplierPayments WHERE date(PaymentDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                overview.CollectionEfficiencyPercentage = If(salesNetAmount <= 0D, 0D, Decimal.Round((overview.CustomerCollectionsAmount / salesNetAmount) * 100D, 2, MidpointRounding.AwayFromZero))
                overview.SupplierPaymentCoveragePercentage = If(purchaseNetAmount <= 0D, 0D, Decimal.Round((overview.SupplierPaymentsAmount / purchaseNetAmount) * 100D, 2, MidpointRounding.AwayFromZero))
                overview.InventorySkuCount = ExecuteIntegerScalar(connection,
                    "SELECT COUNT(1) FROM Products WHERE IsDeleted = 0 AND CurrentStock > 0;",
                    Nothing,
                    Nothing)
                overview.InventoryStockValueAtPTR = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(CurrentStock * PTR), 0) FROM Products WHERE IsDeleted = 0;",
                    Nothing,
                    Nothing)
                overview.OutstandingReceivables = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(OutstandingBalance), 0) FROM Customers;",
                    Nothing,
                    Nothing)
                overview.OutstandingPayables = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(OutstandingBalance), 0) FROM Suppliers;",
                    Nothing,
                    Nothing)
                overview.NetCashMovement = Decimal.Round(overview.CustomerCollectionsAmount - overview.SupplierPaymentsAmount, 2, MidpointRounding.AwayFromZero)
            End Using

            Return overview
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

        Private Function LoadCustomerOutstandingAccounts(connection As DbConnection) As List(Of AccountMasterSnapshot)
            Dim rows As New List(Of AccountMasterSnapshot)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT Id, CustomerName, COALESCE(Gstin, '') AS Gstin, COALESCE(DrugLicenseNumber, '') AS DrugLicenseNumber, COALESCE(Phone, '') AS Phone, OutstandingBalance " &
                    "FROM Customers WHERE OutstandingBalance > 0 ORDER BY CustomerName ASC;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New AccountMasterSnapshot With {
                            .AccountId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .PartyName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                            .Gstin = ConvertNullableString(reader("Gstin")),
                            .DrugLicenseNumber = ConvertNullableString(reader("DrugLicenseNumber")),
                            .Phone = ConvertNullableString(reader("Phone")),
                            .OutstandingBalance = Convert.ToDecimal(reader("OutstandingBalance"), CultureInfo.InvariantCulture)
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadSupplierOutstandingAccounts(connection As DbConnection) As List(Of AccountMasterSnapshot)
            Dim rows As New List(Of AccountMasterSnapshot)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT Id, SupplierName, COALESCE(Gstin, '') AS Gstin, COALESCE(DrugLicenseNumber, '') AS DrugLicenseNumber, COALESCE(Phone, '') AS Phone, OutstandingBalance " &
                    "FROM Suppliers WHERE OutstandingBalance > 0 ORDER BY SupplierName ASC;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New AccountMasterSnapshot With {
                            .AccountId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .PartyName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                            .Gstin = ConvertNullableString(reader("Gstin")),
                            .DrugLicenseNumber = ConvertNullableString(reader("DrugLicenseNumber")),
                            .Phone = ConvertNullableString(reader("Phone")),
                            .OutstandingBalance = Convert.ToDecimal(reader("OutstandingBalance"), CultureInfo.InvariantCulture)
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadCustomerAgingDocuments(connection As DbConnection, asOfDate As DateTime) As List(Of AgingSourceDocument)
            Dim rows As New List(Of AgingSourceDocument)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT Id, CustomerId, InvoiceDate, BalanceAmount " &
                    "FROM Invoices " &
                    "WHERE BalanceAmount > 0 AND date(InvoiceDate) <= date(@AsOfDate) " &
                    "ORDER BY CustomerId ASC, date(InvoiceDate) ASC, Id ASC;"
                command.AddParameter("@AsOfDate", asOfDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New AgingSourceDocument With {
                            .AccountId = Convert.ToInt32(reader("CustomerId"), CultureInfo.InvariantCulture),
                            .DocumentId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .DocumentDate = ParseDate(reader("InvoiceDate")),
                            .OpenAmount = Convert.ToDecimal(reader("BalanceAmount"), CultureInfo.InvariantCulture)
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadSupplierAgingDocuments(connection As DbConnection, asOfDate As DateTime) As List(Of AgingSourceDocument)
            Dim rows As New List(Of AgingSourceDocument)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT p.Id, p.SupplierId, p.PurchaseDate, " &
                    "MAX(0, p.NetAmount - COALESCE(pr.ReturnAmount, 0)) AS OpenAmount " &
                    "FROM Purchases p " &
                    "LEFT JOIN (" &
                    "    SELECT PurchaseId, COALESCE(SUM(NetAmount), 0) AS ReturnAmount " &
                    "    FROM PurchaseReturns WHERE date(ReturnDate) <= date(@AsOfDate) GROUP BY PurchaseId" &
                    ") pr ON pr.PurchaseId = p.Id " &
                    "WHERE date(p.PurchaseDate) <= date(@AsOfDate) " &
                    "ORDER BY p.SupplierId ASC, date(p.PurchaseDate) ASC, p.Id ASC;"
                command.AddParameter("@AsOfDate", asOfDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        Dim openAmount As Decimal = Convert.ToDecimal(reader("OpenAmount"), CultureInfo.InvariantCulture)
                        If openAmount <= 0D Then
                            Continue While
                        End If

                        rows.Add(New AgingSourceDocument With {
                            .AccountId = Convert.ToInt32(reader("SupplierId"), CultureInfo.InvariantCulture),
                            .DocumentId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .DocumentDate = ParseDate(reader("PurchaseDate")),
                            .OpenAmount = openAmount
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadCustomerPaymentsByAccount(connection As DbConnection, asOfDate As DateTime) As Dictionary(Of Integer, Decimal)
            Return LoadPaymentTotalsByAccount(connection,
                "SELECT CustomerId AS AccountId, COALESCE(SUM(Amount), 0) AS PaymentAmount FROM CustomerPayments WHERE date(PaymentDate) <= date(@AsOfDate) GROUP BY CustomerId;",
                asOfDate)
        End Function

        Private Function LoadSupplierPaymentsByAccount(connection As DbConnection, asOfDate As DateTime) As Dictionary(Of Integer, Decimal)
            Return LoadPaymentTotalsByAccount(connection,
                "SELECT SupplierId AS AccountId, COALESCE(SUM(Amount), 0) AS PaymentAmount FROM SupplierPayments WHERE date(PaymentDate) <= date(@AsOfDate) GROUP BY SupplierId;",
                asOfDate)
        End Function

        Private Function LoadPaymentTotalsByAccount(connection As DbConnection, sql As String, asOfDate As DateTime) As Dictionary(Of Integer, Decimal)
            Dim values As New Dictionary(Of Integer, Decimal)()

            Using command = connection.CreateCommand()
                command.CommandText = sql
                command.AddParameter("@AsOfDate", asOfDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        values(Convert.ToInt32(reader("AccountId"), CultureInfo.InvariantCulture)) =
                            Convert.ToDecimal(reader("PaymentAmount"), CultureInfo.InvariantCulture)
                    End While
                End Using
            End Using

            Return values
        End Function

        Private Function BuildAgingRows(accounts As IEnumerable(Of AccountMasterSnapshot), documents As IEnumerable(Of AgingSourceDocument), paymentsByAccount As IDictionary(Of Integer, Decimal), asOfDate As DateTime) As List(Of AgingReportRow)
            Dim rows As New List(Of AgingReportRow)()
            Dim docsByAccount = documents.GroupBy(Function(item) item.AccountId).ToDictionary(Function(group) group.Key, Function(group) group.OrderBy(Function(item) item.DocumentDate).ThenBy(Function(item) item.DocumentId).ToList())

            For Each account As AccountMasterSnapshot In accounts
                Dim row As New AgingReportRow With {
                    .PartyName = account.PartyName,
                    .Gstin = account.Gstin,
                    .DrugLicenseNumber = account.DrugLicenseNumber,
                    .Phone = account.Phone
                }

                Dim remainingPayment As Decimal = If(paymentsByAccount.ContainsKey(account.AccountId), paymentsByAccount(account.AccountId), 0D)
                Dim accountDocuments As List(Of AgingSourceDocument) = If(docsByAccount.ContainsKey(account.AccountId), docsByAccount(account.AccountId), New List(Of AgingSourceDocument)())
                Dim accountedOutstanding As Decimal = 0D
                Dim oldestOpenDate As DateTime? = Nothing

                For Each document As AgingSourceDocument In accountDocuments
                    Dim openAmount As Decimal = Decimal.Round(Math.Max(0D, document.OpenAmount), 2, MidpointRounding.AwayFromZero)
                    If openAmount <= 0D Then
                        Continue For
                    End If

                    If remainingPayment > 0D Then
                        Dim appliedAmount As Decimal = Math.Min(remainingPayment, openAmount)
                        openAmount = Decimal.Round(openAmount - appliedAmount, 2, MidpointRounding.AwayFromZero)
                        remainingPayment = Decimal.Round(remainingPayment - appliedAmount, 2, MidpointRounding.AwayFromZero)
                    End If

                    If openAmount <= 0D Then
                        Continue For
                    End If

                    row.OpenDocumentCount += 1
                    accountedOutstanding += openAmount

                    If Not oldestOpenDate.HasValue OrElse document.DocumentDate < oldestOpenDate.Value Then
                        oldestOpenDate = document.DocumentDate
                    End If

                    Dim ageInDays As Integer = Math.Max(0, CInt((asOfDate.Date - document.DocumentDate.Date).TotalDays))
                    If ageInDays <= 30 Then
                        row.Days0To30Amount += openAmount
                    ElseIf ageInDays <= 60 Then
                        row.Days31To60Amount += openAmount
                    ElseIf ageInDays <= 90 Then
                        row.Days61To90Amount += openAmount
                    Else
                        row.DaysAbove90Amount += openAmount
                    End If
                Next

                row.OldestOpenDate = If(oldestOpenDate.HasValue, oldestOpenDate.Value, asOfDate.Date)
                row.AgeInDays = If(oldestOpenDate.HasValue, Math.Max(0, CInt((asOfDate.Date - oldestOpenDate.Value.Date).TotalDays)), 0)

                Dim unallocatedAmount As Decimal = Decimal.Round(Math.Max(0D, account.OutstandingBalance - accountedOutstanding), 2, MidpointRounding.AwayFromZero)
                If unallocatedAmount > 0D Then
                    row.UnallocatedAmount = unallocatedAmount
                End If

                row.OutstandingBalance =
                    Decimal.Round(
                        row.Days0To30Amount +
                        row.Days31To60Amount +
                        row.Days61To90Amount +
                        row.DaysAbove90Amount +
                        row.UnallocatedAmount,
                        2,
                        MidpointRounding.AwayFromZero)

                rows.Add(row)
            Next

            Return rows.
                OrderByDescending(Function(item) item.DaysAbove90Amount).
                ThenByDescending(Function(item) item.OutstandingBalance).
                ThenBy(Function(item) item.PartyName).
                ToList()
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

        Private Function ExecuteIntegerScalar(connection As DbConnection, sql As String, fromDate As DateTime?, toDate As DateTime?) As Integer
            Using command = connection.CreateCommand()
                command.CommandText = sql
                If fromDate.HasValue AndAlso toDate.HasValue Then
                    AddDateParameters(command, fromDate.Value, toDate.Value)
                End If

                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return 0
                End If

                Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
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
