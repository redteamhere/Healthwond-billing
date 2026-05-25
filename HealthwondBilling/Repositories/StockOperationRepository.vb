Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class StockOperationRepository
        Implements IStockOperationRepository

        Private NotInheritable Class PurchaseReturnContext
            Public Property PurchaseId As Integer
            Public Property SupplierId As Integer
            Public Property PurchaseNumber As String = String.Empty
        End Class

        Private NotInheritable Class ProductStockSnapshot
            Public Property ProductId As Integer
            Public Property ProductName As String = String.Empty
            Public Property BatchNumber As String = String.Empty
            Public Property CurrentStock As Integer
            Public Property PTR As Decimal
        End Class

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GenerateNextPurchaseReturnNumber(returnDate As DateTime) As String Implements IStockOperationRepository.GenerateNextPurchaseReturnNumber
            Return GenerateNumber("PurchaseReturns", "ReturnNumber", GetSettingValue("PurchaseReturnPrefix", "PRN"), returnDate)
        End Function

        Public Function SearchPurchasesForReturn(fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of PurchaseHistoryLookupRow) Implements IStockOperationRepository.SearchPurchasesForReturn
            Dim rows As New List(Of PurchaseHistoryLookupRow)()
            Dim normalizedSearch As String = If(searchTerm, String.Empty).Trim()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT p.Id, p.PurchaseNumber, p.PurchaseDate, p.SupplierId, s.SupplierName, COALESCE(p.SupplierInvoiceNumber, '') AS SupplierInvoiceNumber, " &
                        "COUNT(pi.Id) AS LineCount, COALESCE(SUM(pi.Quantity + pi.FreeQuantity), 0) AS TotalUnits, p.NetAmount, COALESCE(p.Notes, '') AS Notes " &
                        "FROM Purchases p " &
                        "INNER JOIN Suppliers s ON s.Id = p.SupplierId " &
                        "LEFT JOIN PurchaseItems pi ON pi.PurchaseId = p.Id " &
                        "WHERE date(p.PurchaseDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "AND (@Search = '' OR p.PurchaseNumber LIKE @SearchLike OR s.SupplierName LIKE @SearchLike OR COALESCE(p.SupplierInvoiceNumber, '') LIKE @SearchLike OR COALESCE(p.Notes, '') LIKE @SearchLike) " &
                        "GROUP BY p.Id, p.PurchaseNumber, p.PurchaseDate, p.SupplierId, s.SupplierName, p.SupplierInvoiceNumber, p.NetAmount, p.Notes " &
                        "ORDER BY date(p.PurchaseDate) DESC, p.Id DESC;"
                    command.AddParameter("@FromDate", fromDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    command.AddParameter("@ToDate", toDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    command.AddParameter("@Search", normalizedSearch)
                    command.AddParameter("@SearchLike", $"%{normalizedSearch}%")

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New PurchaseHistoryLookupRow With {
                                .PurchaseId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                                .PurchaseNumber = Convert.ToString(reader("PurchaseNumber"), CultureInfo.InvariantCulture),
                                .PurchaseDate = ParseDate(reader("PurchaseDate")),
                                .SupplierId = Convert.ToInt32(reader("SupplierId"), CultureInfo.InvariantCulture),
                                .SupplierName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                                .SupplierInvoiceNumber = Convert.ToString(reader("SupplierInvoiceNumber"), CultureInfo.InvariantCulture),
                                .LineCount = Convert.ToInt32(reader("LineCount"), CultureInfo.InvariantCulture),
                                .TotalUnits = Convert.ToInt32(reader("TotalUnits"), CultureInfo.InvariantCulture),
                                .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture),
                                .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetPurchaseReturnLines(purchaseId As Integer) As List(Of PurchaseReturnLineItem) Implements IStockOperationRepository.GetPurchaseReturnLines
            Dim rows As New List(Of PurchaseReturnLineItem)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT pi.Id AS PurchaseItemId, pi.PurchaseId, pi.ProductId, COALESCE(prod.ProductName, '') AS ProductName, pi.BatchNumber, pi.ExpiryDate, " &
                        "pi.Quantity, pi.FreeQuantity, COALESCE(pri.ReturnedQuantity, 0) AS ReturnedQuantity, COALESCE(pri.ReturnedFreeQuantity, 0) AS ReturnedFreeQuantity, " &
                        "pi.Quantity - COALESCE(pri.ReturnedQuantity, 0) AS RemainingQuantity, pi.FreeQuantity - COALESCE(pri.ReturnedFreeQuantity, 0) AS RemainingFreeQuantity, " &
                        "COALESCE(prod.CurrentStock, 0) AS CurrentStock, pi.PTR, pi.GstPercentage " &
                        "FROM PurchaseItems pi " &
                        "LEFT JOIN Products prod ON prod.Id = pi.ProductId " &
                        "LEFT JOIN (" &
                        "    SELECT PurchaseItemId, COALESCE(SUM(ReturnQuantity), 0) AS ReturnedQuantity, COALESCE(SUM(ReturnFreeQuantity), 0) AS ReturnedFreeQuantity " &
                        "    FROM PurchaseReturnItems GROUP BY PurchaseItemId" &
                        ") pri ON pri.PurchaseItemId = pi.Id " &
                        "WHERE pi.PurchaseId = @PurchaseId " &
                        "ORDER BY pi.Id ASC;"
                    command.AddParameter("@PurchaseId", purchaseId)

                    Using reader = command.ExecuteReader()
                        Dim lineNumber As Integer = 1
                        While reader.Read()
                            Dim remainingQuantity As Integer = Convert.ToInt32(reader("RemainingQuantity"), CultureInfo.InvariantCulture)
                            Dim remainingFreeQuantity As Integer = Convert.ToInt32(reader("RemainingFreeQuantity"), CultureInfo.InvariantCulture)
                            If remainingQuantity <= 0 AndAlso remainingFreeQuantity <= 0 Then
                                Continue While
                            End If

                            rows.Add(New PurchaseReturnLineItem With {
                                .LineNumber = lineNumber,
                                .PurchaseItemId = Convert.ToInt32(reader("PurchaseItemId"), CultureInfo.InvariantCulture),
                                .PurchaseId = Convert.ToInt32(reader("PurchaseId"), CultureInfo.InvariantCulture),
                                .ProductId = Convert.ToInt32(reader("ProductId"), CultureInfo.InvariantCulture),
                                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                                .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                                .ExpiryDate = ParseDate(reader("ExpiryDate")),
                                .PurchasedQuantity = Convert.ToInt32(reader("Quantity"), CultureInfo.InvariantCulture),
                                .PurchasedFreeQuantity = Convert.ToInt32(reader("FreeQuantity"), CultureInfo.InvariantCulture),
                                .AlreadyReturnedQuantity = Convert.ToInt32(reader("ReturnedQuantity"), CultureInfo.InvariantCulture),
                                .AlreadyReturnedFreeQuantity = Convert.ToInt32(reader("ReturnedFreeQuantity"), CultureInfo.InvariantCulture),
                                .RemainingQuantity = remainingQuantity,
                                .RemainingFreeQuantity = remainingFreeQuantity,
                                .CurrentStock = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture),
                                .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                                .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture)
                            })
                            lineNumber += 1
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function SavePurchaseReturn(draft As PurchaseReturnDraft, createdByUserId As Integer) As Integer Implements IStockOperationRepository.SavePurchaseReturn
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim context As PurchaseReturnContext = LoadPurchaseReturnContext(connection, transaction, draft.PurchaseId)
                    If context.SupplierId <> draft.SupplierId Then
                        Throw New InvalidOperationException("The selected supplier does not match the original purchase.")
                    End If

                    Dim returnId As Integer = InsertPurchaseReturn(connection, transaction, draft, createdByUserId, context.SupplierId)

                    For Each item As PurchaseReturnLineItem In draft.Items.Where(Function(line) line.ReturnQuantity + line.ReturnFreeQuantity > 0)
                        Dim latestLine As PurchaseReturnLineItem = GetPurchaseReturnLine(connection, transaction, item.PurchaseItemId)
                        ValidatePurchaseReturnLine(item, latestLine)

                        Dim productSnapshot As ProductStockSnapshot = GetProductStockSnapshot(connection, transaction, item.ProductId)
                        Dim totalReturnedUnits As Integer = item.ReturnQuantity + item.ReturnFreeQuantity
                        If productSnapshot.CurrentStock < totalReturnedUnits Then
                            Throw New InvalidOperationException($"Insufficient stock available to return '{item.ProductName}'.")
                        End If

                        Dim newStock As Integer = productSnapshot.CurrentStock - totalReturnedUnits
                        UpdateProductStock(connection, transaction, item.ProductId, newStock)
                        InsertPurchaseReturnItem(connection, transaction, returnId, item)
                        InsertLedgerEntry(
                            connection,
                            transaction,
                            item.ProductId,
                            item.BatchNumber,
                            "PURCHASE_RETURN",
                            "PurchaseReturn",
                            returnId,
                            0,
                            totalReturnedUnits,
                            newStock,
                            item.PTR,
                            $"Returned against purchase {context.PurchaseNumber}",
                            draft.ReturnDate)
                    Next

                    UpdateSupplierOutstanding(connection, transaction, context.SupplierId, -draft.Summary.NetAmount)
                    AccountingPostingHelper.PostPurchaseReturnVoucher(connection, transaction, returnId, draft, createdByUserId)
                    transaction.Commit()
                    Return returnId
                End Using
            End Using
        End Function

        Public Function GenerateNextStockAdjustmentNumber(adjustmentDate As DateTime) As String Implements IStockOperationRepository.GenerateNextStockAdjustmentNumber
            Return GenerateNumber("StockAdjustments", "AdjustmentNumber", GetSettingValue("StockAdjustmentPrefix", "ADJ"), adjustmentDate)
        End Function

        Public Function SaveStockAdjustment(draft As StockAdjustmentDraft, createdByUserId As Integer) As Integer Implements IStockOperationRepository.SaveStockAdjustment
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim adjustmentId As Integer = InsertStockAdjustment(connection, transaction, draft, createdByUserId)

                    For Each item As StockAdjustmentLineItem In draft.Items
                        Dim productSnapshot As ProductStockSnapshot = GetProductStockSnapshot(connection, transaction, item.ProductId)
                        Dim signedQuantity As Integer = If(item.AdjustmentMode = StockAdjustmentMode.Increase, item.Quantity, -item.Quantity)
                        Dim resultingStock As Integer = productSnapshot.CurrentStock + signedQuantity
                        If resultingStock < 0 Then
                            Throw New InvalidOperationException($"Insufficient stock for '{item.ProductName}'.")
                        End If

                        UpdateProductStock(connection, transaction, item.ProductId, resultingStock)
                        InsertStockAdjustmentItem(connection, transaction, adjustmentId, item, resultingStock)
                        InsertLedgerEntry(
                            connection,
                            transaction,
                            item.ProductId,
                            item.BatchNumber,
                            If(item.AdjustmentMode = StockAdjustmentMode.Increase, "ADJUSTMENT_IN", "ADJUSTMENT_OUT"),
                            "StockAdjustment",
                            adjustmentId,
                            Math.Max(signedQuantity, 0),
                            Math.Abs(Math.Min(signedQuantity, 0)),
                            resultingStock,
                            item.UnitCost,
                            item.Remarks,
                            draft.AdjustmentDate)
                    Next

                    transaction.Commit()
                    Return adjustmentId
                End Using
            End Using
        End Function

        Private Function GenerateNumber(tableName As String, numberColumnName As String, prefix As String, valueDate As DateTime) As String
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim monthPart As String = valueDate.ToString("yyyyMM", CultureInfo.InvariantCulture)
                Dim pattern As String = $"{prefix}-{monthPart}-%"
                Dim lastNumber As Integer = 0

                Using command = connection.CreateCommand()
                    command.CommandText = $"SELECT {numberColumnName} FROM {tableName} WHERE {numberColumnName} LIKE @Pattern ORDER BY Id DESC LIMIT 1;"
                    command.AddParameter("@Pattern", pattern)
                    Dim result As Object = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                        Dim previousNumber As String = Convert.ToString(result, CultureInfo.InvariantCulture)
                        Dim parts As String() = previousNumber.Split("-"c)
                        Dim parsedNumber As Integer
                        If parts.Length >= 3 AndAlso Integer.TryParse(parts(parts.Length - 1), parsedNumber) Then
                            lastNumber = parsedNumber
                        End If
                    End If
                End Using

                Return $"{prefix}-{monthPart}-{(lastNumber + 1).ToString("0000", CultureInfo.InvariantCulture)}"
            End Using
        End Function

        Private Function GetSettingValue(settingKey As String, defaultValue As String) As String
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText = "SELECT SettingValue FROM Settings WHERE SettingKey = @SettingKey LIMIT 1;"
                    command.AddParameter("@SettingKey", settingKey)
                    Dim result As Object = command.ExecuteScalar()
                    If result Is Nothing OrElse result Is DBNull.Value Then
                        Return defaultValue
                    End If

                    Return Convert.ToString(result, CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Private Function LoadPurchaseReturnContext(connection As DbConnection, transaction As DbTransaction, purchaseId As Integer) As PurchaseReturnContext
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT Id, SupplierId, PurchaseNumber FROM Purchases WHERE Id = @Id LIMIT 1;"
                command.AddParameter("@Id", purchaseId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException($"Purchase Id {purchaseId} was not found.")
                    End If

                    Return New PurchaseReturnContext With {
                        .PurchaseId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                        .SupplierId = Convert.ToInt32(reader("SupplierId"), CultureInfo.InvariantCulture),
                        .PurchaseNumber = Convert.ToString(reader("PurchaseNumber"), CultureInfo.InvariantCulture)
                    }
                End Using
            End Using
        End Function

        Private Function InsertPurchaseReturn(connection As DbConnection, transaction As DbTransaction, draft As PurchaseReturnDraft, createdByUserId As Integer, supplierId As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO PurchaseReturns (ReturnNumber, PurchaseId, SupplierId, ReturnDate, SubTotal, GstAmount, RoundOffAmount, NetAmount, Notes, CreatedBy, CreatedAt, UpdatedAt) " &
                    "VALUES (@ReturnNumber, @PurchaseId, @SupplierId, @ReturnDate, @SubTotal, @GstAmount, @RoundOffAmount, @NetAmount, @Notes, @CreatedBy, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@ReturnNumber", draft.ReturnNumber)
                command.AddParameter("@PurchaseId", draft.PurchaseId)
                command.AddParameter("@SupplierId", supplierId)
                command.AddParameter("@ReturnDate", SqliteDateHelper.ToStorageDate(draft.ReturnDate))
                command.AddParameter("@SubTotal", draft.Summary.SubTotal)
                command.AddParameter("@GstAmount", draft.Summary.GstAmount)
                command.AddParameter("@RoundOffAmount", draft.Summary.RoundOffAmount)
                command.AddParameter("@NetAmount", draft.Summary.NetAmount)
                command.AddParameter("@Notes", draft.Notes)
                command.AddParameter("@CreatedBy", createdByUserId)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function GetPurchaseReturnLine(connection As DbConnection, transaction As DbTransaction, purchaseItemId As Integer) As PurchaseReturnLineItem
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT pi.Id AS PurchaseItemId, pi.PurchaseId, pi.ProductId, COALESCE(prod.ProductName, '') AS ProductName, pi.BatchNumber, pi.ExpiryDate, " &
                    "pi.Quantity, pi.FreeQuantity, COALESCE(pri.ReturnedQuantity, 0) AS ReturnedQuantity, COALESCE(pri.ReturnedFreeQuantity, 0) AS ReturnedFreeQuantity, " &
                    "pi.Quantity - COALESCE(pri.ReturnedQuantity, 0) AS RemainingQuantity, pi.FreeQuantity - COALESCE(pri.ReturnedFreeQuantity, 0) AS RemainingFreeQuantity, " &
                    "COALESCE(prod.CurrentStock, 0) AS CurrentStock, pi.PTR, pi.GstPercentage " &
                    "FROM PurchaseItems pi " &
                    "LEFT JOIN Products prod ON prod.Id = pi.ProductId " &
                    "LEFT JOIN (" &
                    "    SELECT PurchaseItemId, COALESCE(SUM(ReturnQuantity), 0) AS ReturnedQuantity, COALESCE(SUM(ReturnFreeQuantity), 0) AS ReturnedFreeQuantity " &
                    "    FROM PurchaseReturnItems GROUP BY PurchaseItemId" &
                    ") pri ON pri.PurchaseItemId = pi.Id " &
                    "WHERE pi.Id = @PurchaseItemId LIMIT 1;"
                command.AddParameter("@PurchaseItemId", purchaseItemId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException("A selected purchase line could not be found.")
                    End If

                    Return New PurchaseReturnLineItem With {
                        .PurchaseItemId = Convert.ToInt32(reader("PurchaseItemId"), CultureInfo.InvariantCulture),
                        .PurchaseId = Convert.ToInt32(reader("PurchaseId"), CultureInfo.InvariantCulture),
                        .ProductId = Convert.ToInt32(reader("ProductId"), CultureInfo.InvariantCulture),
                        .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                        .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                        .ExpiryDate = ParseDate(reader("ExpiryDate")),
                        .PurchasedQuantity = Convert.ToInt32(reader("Quantity"), CultureInfo.InvariantCulture),
                        .PurchasedFreeQuantity = Convert.ToInt32(reader("FreeQuantity"), CultureInfo.InvariantCulture),
                        .AlreadyReturnedQuantity = Convert.ToInt32(reader("ReturnedQuantity"), CultureInfo.InvariantCulture),
                        .AlreadyReturnedFreeQuantity = Convert.ToInt32(reader("ReturnedFreeQuantity"), CultureInfo.InvariantCulture),
                        .RemainingQuantity = Convert.ToInt32(reader("RemainingQuantity"), CultureInfo.InvariantCulture),
                        .RemainingFreeQuantity = Convert.ToInt32(reader("RemainingFreeQuantity"), CultureInfo.InvariantCulture),
                        .CurrentStock = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture),
                        .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                        .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture)
                    }
                End Using
            End Using
        End Function

        Private Sub ValidatePurchaseReturnLine(requestedLine As PurchaseReturnLineItem, persistedLine As PurchaseReturnLineItem)
            If requestedLine.ReturnQuantity < 0 OrElse requestedLine.ReturnFreeQuantity < 0 Then
                Throw New InvalidOperationException($"Return quantities cannot be negative for '{requestedLine.ProductName}'.")
            End If

            If requestedLine.ReturnQuantity > persistedLine.RemainingQuantity Then
                Throw New InvalidOperationException($"Return quantity exceeds remaining purchase quantity for '{requestedLine.ProductName}'.")
            End If

            If requestedLine.ReturnFreeQuantity > persistedLine.RemainingFreeQuantity Then
                Throw New InvalidOperationException($"Return free quantity exceeds remaining free quantity for '{requestedLine.ProductName}'.")
            End If
        End Sub

        Private Function GetProductStockSnapshot(connection As DbConnection, transaction As DbTransaction, productId As Integer) As ProductStockSnapshot
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT Id, ProductName, BatchNumber, CurrentStock, PTR FROM Products WHERE Id = @Id LIMIT 1;"
                command.AddParameter("@Id", productId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException("The selected product was not found.")
                    End If

                    Return New ProductStockSnapshot With {
                        .ProductId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                        .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                        .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                        .CurrentStock = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture),
                        .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture)
                    }
                End Using
            End Using
        End Function

        Private Sub UpdateProductStock(connection As DbConnection, transaction As DbTransaction, productId As Integer, currentStock As Integer)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "UPDATE Products SET CurrentStock = @CurrentStock, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@CurrentStock", currentStock)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", productId)

                If command.ExecuteNonQuery() = 0 Then
                    Throw New InvalidOperationException("The selected product could not be updated.")
                End If
            End Using
        End Sub

        Private Sub InsertPurchaseReturnItem(connection As DbConnection, transaction As DbTransaction, returnId As Integer, item As PurchaseReturnLineItem)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO PurchaseReturnItems (PurchaseReturnId, PurchaseItemId, ProductId, BatchNumber, ReturnQuantity, ReturnFreeQuantity, PTR, GstPercentage, TaxableAmount, GstAmount, LineTotal) " &
                    "VALUES (@PurchaseReturnId, @PurchaseItemId, @ProductId, @BatchNumber, @ReturnQuantity, @ReturnFreeQuantity, @PTR, @GstPercentage, @TaxableAmount, @GstAmount, @LineTotal);"
                command.AddParameter("@PurchaseReturnId", returnId)
                command.AddParameter("@PurchaseItemId", item.PurchaseItemId)
                command.AddParameter("@ProductId", item.ProductId)
                command.AddParameter("@BatchNumber", item.BatchNumber)
                command.AddParameter("@ReturnQuantity", item.ReturnQuantity)
                command.AddParameter("@ReturnFreeQuantity", item.ReturnFreeQuantity)
                command.AddParameter("@PTR", item.PTR)
                command.AddParameter("@GstPercentage", item.GstPercentage)
                command.AddParameter("@TaxableAmount", item.TaxableAmount)
                command.AddParameter("@GstAmount", item.GstAmount)
                command.AddParameter("@LineTotal", item.LineTotal)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Function InsertStockAdjustment(connection As DbConnection, transaction As DbTransaction, draft As StockAdjustmentDraft, createdByUserId As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO StockAdjustments (AdjustmentNumber, AdjustmentDate, Notes, CreatedBy, CreatedAt, UpdatedAt) " &
                    "VALUES (@AdjustmentNumber, @AdjustmentDate, @Notes, @CreatedBy, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@AdjustmentNumber", draft.AdjustmentNumber)
                command.AddParameter("@AdjustmentDate", SqliteDateHelper.ToStorageDate(draft.AdjustmentDate))
                command.AddParameter("@Notes", draft.Notes)
                command.AddParameter("@CreatedBy", createdByUserId)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub InsertStockAdjustmentItem(connection As DbConnection, transaction As DbTransaction, adjustmentId As Integer, item As StockAdjustmentLineItem, resultingStock As Integer)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO StockAdjustmentItems (StockAdjustmentId, ProductId, BatchNumber, AdjustmentMode, Quantity, BalanceQuantity, UnitCost, Remarks) " &
                    "VALUES (@StockAdjustmentId, @ProductId, @BatchNumber, @AdjustmentMode, @Quantity, @BalanceQuantity, @UnitCost, @Remarks);"
                command.AddParameter("@StockAdjustmentId", adjustmentId)
                command.AddParameter("@ProductId", item.ProductId)
                command.AddParameter("@BatchNumber", item.BatchNumber)
                command.AddParameter("@AdjustmentMode", item.AdjustmentMode.ToString())
                command.AddParameter("@Quantity", item.Quantity)
                command.AddParameter("@BalanceQuantity", resultingStock)
                command.AddParameter("@UnitCost", item.UnitCost)
                command.AddParameter("@Remarks", item.Remarks)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub InsertLedgerEntry(connection As DbConnection, transaction As DbTransaction, productId As Integer, batchNumber As String, transactionType As String, referenceType As String, referenceId As Integer, quantityIn As Integer, quantityOut As Integer, balanceQuantity As Integer, unitCost As Decimal, remarks As String, transactionDate As DateTime)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO StockLedger (ProductId, BatchNumber, TransactionType, ReferenceType, ReferenceId, QuantityIn, QuantityOut, BalanceQuantity, UnitCost, Remarks, TransactionDate, CreatedAt) " &
                    "VALUES (@ProductId, @BatchNumber, @TransactionType, @ReferenceType, @ReferenceId, @QuantityIn, @QuantityOut, @BalanceQuantity, @UnitCost, @Remarks, @TransactionDate, @CreatedAt);"
                command.AddParameter("@ProductId", productId)
                command.AddParameter("@BatchNumber", batchNumber)
                command.AddParameter("@TransactionType", transactionType)
                command.AddParameter("@ReferenceType", referenceType)
                command.AddParameter("@ReferenceId", referenceId)
                command.AddParameter("@QuantityIn", quantityIn)
                command.AddParameter("@QuantityOut", quantityOut)
                command.AddParameter("@BalanceQuantity", balanceQuantity)
                command.AddParameter("@UnitCost", unitCost)
                command.AddParameter("@Remarks", remarks)
                command.AddParameter("@TransactionDate", SqliteDateHelper.ToStorageDate(transactionDate))
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub UpdateSupplierOutstanding(connection As DbConnection, transaction As DbTransaction, supplierId As Integer, amountDelta As Decimal)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "UPDATE Suppliers SET OutstandingBalance = OutstandingBalance + @AmountDelta, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@AmountDelta", amountDelta)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", supplierId)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Function ParseDate(value As Object) As DateTime
            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return DateTime.Today
        End Function

    End Class

End Namespace
