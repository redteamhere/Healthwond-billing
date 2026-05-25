Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class InvoiceRepository
        Implements IInvoiceRepository

        Private NotInheritable Class ExistingInvoiceSnapshot
            Public Property CustomerId As Integer
            Public Property BalanceAmount As Decimal
        End Class

        Private NotInheritable Class RestorableInvoiceItem
            Public Property ProductId As Integer
            Public Property ProductName As String = String.Empty
            Public Property BatchNumber As String = String.Empty
            Public Property QuantityToRestore As Integer
            Public Property CurrentStock As Integer
            Public Property PTR As Decimal
        End Class

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GenerateNextInvoiceNumber(invoiceDate As DateTime) As String Implements IInvoiceRepository.GenerateNextInvoiceNumber
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim prefix As String = GetSetting(connection, "InvoicePrefix", "HWB")
                Dim monthPart As String = invoiceDate.ToString("yyyyMM", CultureInfo.InvariantCulture)
                Dim pattern As String = $"{prefix}-{monthPart}-%"
                Dim lastNumber As Integer = 0

                Using command = connection.CreateCommand()
                    command.CommandText = "SELECT InvoiceNumber FROM Invoices WHERE InvoiceNumber LIKE @Pattern ORDER BY Id DESC LIMIT 1;"
                    command.AddParameter("@Pattern", pattern)
                    Dim result As Object = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                        Dim invoiceNumber As String = Convert.ToString(result, CultureInfo.InvariantCulture)
                        Dim parts As String() = invoiceNumber.Split("-"c)
                        Dim parsedNumber As Integer
                        If parts.Length >= 3 AndAlso Integer.TryParse(parts(parts.Length - 1), parsedNumber) Then
                            lastNumber = parsedNumber
                        End If
                    End If
                End Using

                Return $"{prefix}-{monthPart}-{(lastNumber + 1).ToString("0000", CultureInfo.InvariantCulture)}"
            End Using
        End Function

        Public Function SaveInvoice(draft As BillingInvoiceDraft, createdByUserId As Integer) As Integer Implements IInvoiceRepository.SaveInvoice
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim invoiceId As Integer = InsertInvoice(connection, transaction, draft, createdByUserId)

                    For Each item As BillingLineItem In draft.Items
                        InsertInvoiceItem(connection, transaction, invoiceId, item)
                        ApplyStockReduction(connection, transaction, invoiceId, item)
                    Next

                    AdjustCustomerOutstanding(connection, transaction, draft.CustomerId, draft.Summary.BalanceAmount)
                    AccountingPostingHelper.PostInvoiceVoucher(connection, transaction, invoiceId, draft, createdByUserId)
                    transaction.Commit()
                    Return invoiceId
                End Using
            End Using
        End Function

        Public Function UpdateInvoice(draft As BillingInvoiceDraft, updatedByUserId As Integer) As Integer Implements IInvoiceRepository.UpdateInvoice
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim existingInvoice As ExistingInvoiceSnapshot = LoadExistingInvoiceSnapshot(connection, transaction, draft.InvoiceId)

                    RestorePriorInvoiceStock(connection, transaction, draft.InvoiceId)
                    AdjustCustomerOutstanding(connection, transaction, existingInvoice.CustomerId, -existingInvoice.BalanceAmount)
                    DeleteInvoiceItems(connection, transaction, draft.InvoiceId)
                    UpdateInvoiceHeader(connection, transaction, draft, updatedByUserId)

                    For Each item As BillingLineItem In draft.Items
                        InsertInvoiceItem(connection, transaction, draft.InvoiceId, item)
                        ApplyStockReduction(connection, transaction, draft.InvoiceId, item)
                    Next

                    AdjustCustomerOutstanding(connection, transaction, draft.CustomerId, draft.Summary.BalanceAmount)
                    AccountingPostingHelper.PostInvoiceVoucher(connection, transaction, draft.InvoiceId, draft, updatedByUserId)
                    transaction.Commit()
                    Return draft.InvoiceId
                End Using
            End Using
        End Function

        Public Function SearchInvoices(fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of InvoiceHistoryRow) Implements IInvoiceRepository.SearchInvoices
            Dim rows As New List(Of InvoiceHistoryRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT i.Id, i.InvoiceNumber, i.InvoiceDate, c.CustomerName, COALESCE(i.PaymentMode, '') AS PaymentMode, " &
                        "COUNT(ii.Id) AS LineCount, COALESCE(SUM(ii.Quantity + ii.FreeQuantity), 0) AS TotalUnits, i.NetAmount, i.AmountPaid, i.BalanceAmount, i.UpdatedAt " &
                        "FROM Invoices i " &
                        "INNER JOIN Customers c ON c.Id = i.CustomerId " &
                        "LEFT JOIN InvoiceItems ii ON ii.InvoiceId = i.Id " &
                        "WHERE date(i.InvoiceDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "AND (@Search = '' OR i.InvoiceNumber LIKE @SearchLike OR c.CustomerName LIKE @SearchLike OR COALESCE(i.PaymentMode, '') LIKE @SearchLike OR COALESCE(i.Notes, '') LIKE @SearchLike) " &
                        "GROUP BY i.Id, i.InvoiceNumber, i.InvoiceDate, c.CustomerName, i.PaymentMode, i.NetAmount, i.AmountPaid, i.BalanceAmount, i.UpdatedAt " &
                        "ORDER BY date(i.InvoiceDate) DESC, i.Id DESC;"
                    command.AddParameter("@FromDate", fromDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    command.AddParameter("@ToDate", toDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture))
                    command.AddParameter("@Search", If(searchTerm, String.Empty).Trim())
                    command.AddParameter("@SearchLike", $"%{If(searchTerm, String.Empty).Trim()}%")

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New InvoiceHistoryRow With {
                                .InvoiceId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                                .InvoiceNumber = Convert.ToString(reader("InvoiceNumber"), CultureInfo.InvariantCulture),
                                .InvoiceDate = ParseDate(reader("InvoiceDate")),
                                .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                                .PaymentMode = Convert.ToString(reader("PaymentMode"), CultureInfo.InvariantCulture),
                                .LineCount = Convert.ToInt32(reader("LineCount"), CultureInfo.InvariantCulture),
                                .TotalUnits = Convert.ToInt32(reader("TotalUnits"), CultureInfo.InvariantCulture),
                                .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture),
                                .AmountPaid = Convert.ToDecimal(reader("AmountPaid"), CultureInfo.InvariantCulture),
                                .BalanceAmount = Convert.ToDecimal(reader("BalanceAmount"), CultureInfo.InvariantCulture),
                                .UpdatedAt = ParseDateTime(reader("UpdatedAt"))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function LoadInvoiceDraft(invoiceId As Integer) As BillingInvoiceDraft Implements IInvoiceRepository.LoadInvoiceDraft
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim draft As BillingInvoiceDraft = LoadInvoiceDraftHeader(connection, invoiceId)
                draft.Items = LoadInvoiceDraftItems(connection, invoiceId)
                Return draft
            End Using
        End Function

        Public Function GetInvoiceDocument(invoiceId As Integer) As InvoiceDocument Implements IInvoiceRepository.GetInvoiceDocument
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim document As InvoiceDocument = LoadInvoiceHeader(connection, invoiceId)
                document.CompanyName = GetSetting(connection, "CompanyName", "Healthwond Pharmacy")
                document.CompanyAddress = GetSetting(connection, "CompanyAddress", "Address not configured")
                document.CompanyPhone = GetSetting(connection, "CompanyPhone", String.Empty)
                document.CompanyGstin = GetSetting(connection, "CompanyGstin", String.Empty)
                document.CompanyDrugLicenseNumber = GetSetting(connection, "CompanyDrugLicense", String.Empty)
                document.Items = LoadInvoiceItems(connection, invoiceId)
                Return document
            End Using
        End Function

        Private Function InsertInvoice(connection As DbConnection, transaction As DbTransaction, draft As BillingInvoiceDraft, createdByUserId As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Invoices (InvoiceNumber, CustomerId, InvoiceDate, PaymentMode, SubTotal, DiscountAmount, SchemeAmount, GstAmount, RoundOffAmount, NetAmount, AmountPaid, BalanceAmount, Notes, CreatedBy, CreatedAt, UpdatedAt) " &
                    "VALUES (@InvoiceNumber, @CustomerId, @InvoiceDate, @PaymentMode, @SubTotal, @DiscountAmount, @SchemeAmount, @GstAmount, @RoundOffAmount, @NetAmount, @AmountPaid, @BalanceAmount, @Notes, @CreatedBy, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@InvoiceNumber", draft.InvoiceNumber)
                command.AddParameter("@CustomerId", draft.CustomerId)
                command.AddParameter("@InvoiceDate", SqliteDateHelper.ToStorageDate(draft.InvoiceDate))
                command.AddParameter("@PaymentMode", draft.PaymentMode)
                command.AddParameter("@SubTotal", draft.Summary.SubTotal)
                command.AddParameter("@DiscountAmount", draft.Summary.DiscountAmount)
                command.AddParameter("@SchemeAmount", draft.Summary.SchemeAmount)
                command.AddParameter("@GstAmount", draft.Summary.GstAmount)
                command.AddParameter("@RoundOffAmount", draft.Summary.RoundOffAmount)
                command.AddParameter("@NetAmount", draft.Summary.NetAmount)
                command.AddParameter("@AmountPaid", draft.Summary.AmountPaid)
                command.AddParameter("@BalanceAmount", draft.Summary.BalanceAmount)
                command.AddParameter("@Notes", draft.Notes)
                command.AddParameter("@CreatedBy", createdByUserId)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub UpdateInvoiceHeader(connection As DbConnection, transaction As DbTransaction, draft As BillingInvoiceDraft, updatedByUserId As Integer)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "UPDATE Invoices SET CustomerId = @CustomerId, InvoiceDate = @InvoiceDate, PaymentMode = @PaymentMode, SubTotal = @SubTotal, DiscountAmount = @DiscountAmount, " &
                    "SchemeAmount = @SchemeAmount, GstAmount = @GstAmount, RoundOffAmount = @RoundOffAmount, NetAmount = @NetAmount, AmountPaid = @AmountPaid, BalanceAmount = @BalanceAmount, " &
                    "Notes = @Notes, CreatedBy = @CreatedBy, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@CustomerId", draft.CustomerId)
                command.AddParameter("@InvoiceDate", SqliteDateHelper.ToStorageDate(draft.InvoiceDate))
                command.AddParameter("@PaymentMode", draft.PaymentMode)
                command.AddParameter("@SubTotal", draft.Summary.SubTotal)
                command.AddParameter("@DiscountAmount", draft.Summary.DiscountAmount)
                command.AddParameter("@SchemeAmount", draft.Summary.SchemeAmount)
                command.AddParameter("@GstAmount", draft.Summary.GstAmount)
                command.AddParameter("@RoundOffAmount", draft.Summary.RoundOffAmount)
                command.AddParameter("@NetAmount", draft.Summary.NetAmount)
                command.AddParameter("@AmountPaid", draft.Summary.AmountPaid)
                command.AddParameter("@BalanceAmount", draft.Summary.BalanceAmount)
                command.AddParameter("@Notes", draft.Notes)
                command.AddParameter("@CreatedBy", updatedByUserId)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", draft.InvoiceId)

                If command.ExecuteNonQuery() <> 1 Then
                    Throw New InvalidOperationException($"Invoice Id {draft.InvoiceId} could not be updated.")
                End If
            End Using
        End Sub

        Private Sub InsertInvoiceItem(connection As DbConnection, transaction As DbTransaction, invoiceId As Integer, item As BillingLineItem)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO InvoiceItems (InvoiceId, ProductId, BatchNumber, ExpiryDate, Quantity, FreeQuantity, Rate, MRP, DiscountPercentage, DiscountAmount, SchemeDescription, GstPercentage, TaxableAmount, GstAmount, LineTotal) " &
                    "VALUES (@InvoiceId, @ProductId, @BatchNumber, @ExpiryDate, @Quantity, @FreeQuantity, @Rate, @MRP, @DiscountPercentage, @DiscountAmount, @SchemeDescription, @GstPercentage, @TaxableAmount, @GstAmount, @LineTotal);"
                command.AddParameter("@InvoiceId", invoiceId)
                command.AddParameter("@ProductId", item.ProductId)
                command.AddParameter("@BatchNumber", item.BatchNumber)
                command.AddParameter("@ExpiryDate", SqliteDateHelper.ToStorageDate(item.ExpiryDate))
                command.AddParameter("@Quantity", item.Quantity)
                command.AddParameter("@FreeQuantity", item.FreeQuantity)
                command.AddParameter("@Rate", item.Rate)
                command.AddParameter("@MRP", item.MRP)
                command.AddParameter("@DiscountPercentage", item.DiscountPercentage)
                command.AddParameter("@DiscountAmount", item.DiscountAmount)
                command.AddParameter("@SchemeDescription", item.SchemeDescription)
                command.AddParameter("@GstPercentage", item.GstPercentage)
                command.AddParameter("@TaxableAmount", item.TaxableAmount)
                command.AddParameter("@GstAmount", item.GstAmount)
                command.AddParameter("@LineTotal", item.LineTotal)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub ApplyStockReduction(connection As DbConnection, transaction As DbTransaction, invoiceId As Integer, item As BillingLineItem)
            Dim quantityOut As Integer = item.Quantity + item.FreeQuantity
            Dim currentStock As Integer

            Using selectCommand = connection.CreateCommand()
                selectCommand.Transaction = transaction
                selectCommand.CommandText = "SELECT CurrentStock FROM Products WHERE Id = @Id LIMIT 1;"
                selectCommand.AddParameter("@Id", item.ProductId)
                Dim result As Object = selectCommand.ExecuteScalar()

                If result Is Nothing OrElse result Is DBNull.Value Then
                    Throw New InvalidOperationException($"Product '{item.ProductName}' is no longer available.")
                End If

                currentStock = Convert.ToInt32(result, CultureInfo.InvariantCulture)
            End Using

            If currentStock < quantityOut Then
                Throw New InvalidOperationException($"Insufficient stock for '{item.ProductName}'.")
            End If

            Dim newStock As Integer = currentStock - quantityOut

            Using updateCommand = connection.CreateCommand()
                updateCommand.Transaction = transaction
                updateCommand.CommandText = "UPDATE Products SET CurrentStock = @CurrentStock, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                updateCommand.AddParameter("@CurrentStock", newStock)
                updateCommand.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                updateCommand.AddParameter("@Id", item.ProductId)
                updateCommand.ExecuteNonQuery()
            End Using

            Using ledgerCommand = connection.CreateCommand()
                ledgerCommand.Transaction = transaction
                ledgerCommand.CommandText =
                    "INSERT INTO StockLedger (ProductId, BatchNumber, TransactionType, ReferenceType, ReferenceId, QuantityIn, QuantityOut, BalanceQuantity, UnitCost, Remarks, TransactionDate, CreatedAt) " &
                    "VALUES (@ProductId, @BatchNumber, @TransactionType, @ReferenceType, @ReferenceId, 0, @QuantityOut, @BalanceQuantity, @UnitCost, @Remarks, @TransactionDate, @CreatedAt);"
                ledgerCommand.AddParameter("@ProductId", item.ProductId)
                ledgerCommand.AddParameter("@BatchNumber", item.BatchNumber)
                ledgerCommand.AddParameter("@TransactionType", "SALE")
                ledgerCommand.AddParameter("@ReferenceType", "Invoice")
                ledgerCommand.AddParameter("@ReferenceId", invoiceId)
                ledgerCommand.AddParameter("@QuantityOut", quantityOut)
                ledgerCommand.AddParameter("@BalanceQuantity", newStock)
                ledgerCommand.AddParameter("@UnitCost", item.PTR)
                ledgerCommand.AddParameter("@Remarks", $"Sold through invoice {invoiceId}")
                ledgerCommand.AddParameter("@TransactionDate", SqliteDateHelper.ToStorageDate(DateTime.Today))
                ledgerCommand.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                ledgerCommand.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub RestorePriorInvoiceStock(connection As DbConnection, transaction As DbTransaction, invoiceId As Integer)
            Dim itemsToRestore As List(Of RestorableInvoiceItem) = LoadRestorableInvoiceItems(connection, transaction, invoiceId)

            For Each item As RestorableInvoiceItem In itemsToRestore
                Dim restoredStock As Integer = item.CurrentStock + item.QuantityToRestore

                Using updateCommand = connection.CreateCommand()
                    updateCommand.Transaction = transaction
                    updateCommand.CommandText = "UPDATE Products SET CurrentStock = @CurrentStock, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                    updateCommand.AddParameter("@CurrentStock", restoredStock)
                    updateCommand.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    updateCommand.AddParameter("@Id", item.ProductId)
                    updateCommand.ExecuteNonQuery()
                End Using

                Using ledgerCommand = connection.CreateCommand()
                    ledgerCommand.Transaction = transaction
                    ledgerCommand.CommandText =
                        "INSERT INTO StockLedger (ProductId, BatchNumber, TransactionType, ReferenceType, ReferenceId, QuantityIn, QuantityOut, BalanceQuantity, UnitCost, Remarks, TransactionDate, CreatedAt) " &
                        "VALUES (@ProductId, @BatchNumber, @TransactionType, @ReferenceType, @ReferenceId, @QuantityIn, 0, @BalanceQuantity, @UnitCost, @Remarks, @TransactionDate, @CreatedAt);"
                    ledgerCommand.AddParameter("@ProductId", item.ProductId)
                    ledgerCommand.AddParameter("@BatchNumber", item.BatchNumber)
                    ledgerCommand.AddParameter("@TransactionType", "SALE_REVERSAL")
                    ledgerCommand.AddParameter("@ReferenceType", "InvoiceEdit")
                    ledgerCommand.AddParameter("@ReferenceId", invoiceId)
                    ledgerCommand.AddParameter("@QuantityIn", item.QuantityToRestore)
                    ledgerCommand.AddParameter("@BalanceQuantity", restoredStock)
                    ledgerCommand.AddParameter("@UnitCost", item.PTR)
                    ledgerCommand.AddParameter("@Remarks", $"Reversed invoice {invoiceId} before update")
                    ledgerCommand.AddParameter("@TransactionDate", SqliteDateHelper.ToStorageDate(DateTime.Today))
                    ledgerCommand.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    ledgerCommand.ExecuteNonQuery()
                End Using
            Next
        End Sub

        Private Function LoadRestorableInvoiceItems(connection As DbConnection, transaction As DbTransaction, invoiceId As Integer) As List(Of RestorableInvoiceItem)
            Dim items As New List(Of RestorableInvoiceItem)()

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT ii.ProductId, MAX(COALESCE(p.ProductName, '')) AS ProductName, MAX(ii.BatchNumber) AS BatchNumber, " &
                    "COALESCE(SUM(ii.Quantity + ii.FreeQuantity), 0) AS QuantityToRestore, MAX(COALESCE(p.CurrentStock, 0)) AS CurrentStock, MAX(COALESCE(p.PTR, 0)) AS PTR " &
                    "FROM InvoiceItems ii " &
                    "INNER JOIN Products p ON p.Id = ii.ProductId " &
                    "WHERE ii.InvoiceId = @InvoiceId " &
                    "GROUP BY ii.ProductId;"
                command.AddParameter("@InvoiceId", invoiceId)

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        items.Add(New RestorableInvoiceItem With {
                            .ProductId = Convert.ToInt32(reader("ProductId"), CultureInfo.InvariantCulture),
                            .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                            .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                            .QuantityToRestore = Convert.ToInt32(reader("QuantityToRestore"), CultureInfo.InvariantCulture),
                            .CurrentStock = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture),
                            .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture)
                        })
                    End While
                End Using
            End Using

            Return items
        End Function

        Private Sub DeleteInvoiceItems(connection As DbConnection, transaction As DbTransaction, invoiceId As Integer)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "DELETE FROM InvoiceItems WHERE InvoiceId = @InvoiceId;"
                command.AddParameter("@InvoiceId", invoiceId)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub AdjustCustomerOutstanding(connection As DbConnection, transaction As DbTransaction, customerId As Integer, amountDelta As Decimal)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "UPDATE Customers SET OutstandingBalance = OutstandingBalance + @AmountDelta, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@AmountDelta", amountDelta)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", customerId)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Function LoadExistingInvoiceSnapshot(connection As DbConnection, transaction As DbTransaction, invoiceId As Integer) As ExistingInvoiceSnapshot
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT CustomerId, BalanceAmount FROM Invoices WHERE Id = @Id LIMIT 1;"
                command.AddParameter("@Id", invoiceId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException($"Invoice Id {invoiceId} was not found.")
                    End If

                    Return New ExistingInvoiceSnapshot With {
                        .CustomerId = Convert.ToInt32(reader("CustomerId"), CultureInfo.InvariantCulture),
                        .BalanceAmount = Convert.ToDecimal(reader("BalanceAmount"), CultureInfo.InvariantCulture)
                    }
                End Using
            End Using
        End Function

        Private Function LoadInvoiceDraftHeader(connection As DbConnection, invoiceId As Integer) As BillingInvoiceDraft
            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT i.Id, i.InvoiceNumber, i.CustomerId, c.CustomerName, i.InvoiceDate, COALESCE(i.PaymentMode, '') AS PaymentMode, i.AmountPaid, COALESCE(i.Notes, '') AS Notes " &
                    "FROM Invoices i " &
                    "INNER JOIN Customers c ON c.Id = i.CustomerId " &
                    "WHERE i.Id = @Id LIMIT 1;"
                command.AddParameter("@Id", invoiceId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException($"Invoice Id {invoiceId} was not found.")
                    End If

                    Return New BillingInvoiceDraft With {
                        .InvoiceId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                        .InvoiceNumber = Convert.ToString(reader("InvoiceNumber"), CultureInfo.InvariantCulture),
                        .CustomerId = Convert.ToInt32(reader("CustomerId"), CultureInfo.InvariantCulture),
                        .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                        .InvoiceDate = ParseDate(reader("InvoiceDate")),
                        .PaymentMode = Convert.ToString(reader("PaymentMode"), CultureInfo.InvariantCulture),
                        .AmountPaid = Convert.ToDecimal(reader("AmountPaid"), CultureInfo.InvariantCulture),
                        .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture)
                    }
                End Using
            End Using
        End Function

        Private Function LoadInvoiceDraftItems(connection As DbConnection, invoiceId As Integer) As List(Of BillingLineItem)
            Dim items As New List(Of BillingLineItem)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT ii.ProductId, COALESCE(p.ProductName, '') AS ProductName, ii.BatchNumber, ii.ExpiryDate, COALESCE(p.Packing, '') AS Packing, COALESCE(p.Barcode, '') AS Barcode, " &
                    "COALESCE(p.CurrentStock, 0) + (ii.Quantity + ii.FreeQuantity) AS AvailableStock, ii.Quantity, ii.FreeQuantity, ii.Rate, COALESCE(p.PTR, 0) AS PTR, ii.MRP, " &
                    "ii.DiscountPercentage, ii.DiscountAmount, COALESCE(ii.SchemeDescription, '') AS SchemeDescription, ii.GstPercentage, ii.TaxableAmount, ii.GstAmount, ii.LineTotal " &
                    "FROM InvoiceItems ii " &
                    "LEFT JOIN Products p ON p.Id = ii.ProductId " &
                    "WHERE ii.InvoiceId = @InvoiceId ORDER BY ii.Id ASC;"
                command.AddParameter("@InvoiceId", invoiceId)

                Using reader = command.ExecuteReader()
                    Dim lineNumber As Integer = 1
                    While reader.Read()
                        items.Add(New BillingLineItem With {
                            .LineNumber = lineNumber,
                            .ProductId = Convert.ToInt32(reader("ProductId"), CultureInfo.InvariantCulture),
                            .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                            .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                            .ExpiryDate = ParseDate(reader("ExpiryDate")),
                            .Packing = Convert.ToString(reader("Packing"), CultureInfo.InvariantCulture),
                            .Barcode = Convert.ToString(reader("Barcode"), CultureInfo.InvariantCulture),
                            .AvailableStock = Convert.ToInt32(reader("AvailableStock"), CultureInfo.InvariantCulture),
                            .Quantity = Convert.ToInt32(reader("Quantity"), CultureInfo.InvariantCulture),
                            .FreeQuantity = Convert.ToInt32(reader("FreeQuantity"), CultureInfo.InvariantCulture),
                            .Rate = Convert.ToDecimal(reader("Rate"), CultureInfo.InvariantCulture),
                            .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                            .MRP = Convert.ToDecimal(reader("MRP"), CultureInfo.InvariantCulture),
                            .DiscountPercentage = Convert.ToDecimal(reader("DiscountPercentage"), CultureInfo.InvariantCulture),
                            .DiscountAmount = Convert.ToDecimal(reader("DiscountAmount"), CultureInfo.InvariantCulture),
                            .SchemeDescription = Convert.ToString(reader("SchemeDescription"), CultureInfo.InvariantCulture),
                            .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                            .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                            .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                            .LineTotal = Convert.ToDecimal(reader("LineTotal"), CultureInfo.InvariantCulture)
                        })
                        lineNumber += 1
                    End While
                End Using
            End Using

            Return items
        End Function

        Private Function LoadInvoiceHeader(connection As DbConnection, invoiceId As Integer) As InvoiceDocument
            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT i.Id, i.InvoiceNumber, i.InvoiceDate, i.PaymentMode, i.Notes, i.SubTotal, i.DiscountAmount, i.SchemeAmount, i.GstAmount, i.RoundOffAmount, i.NetAmount, i.AmountPaid, i.BalanceAmount, " &
                    "c.CustomerName, c.Gstin, c.DrugLicenseNumber, c.Address, c.Phone " &
                    "FROM Invoices i " &
                    "INNER JOIN Customers c ON c.Id = i.CustomerId " &
                    "WHERE i.Id = @Id LIMIT 1;"
                command.AddParameter("@Id", invoiceId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException($"Invoice Id {invoiceId} was not found.")
                    End If

                    Return New InvoiceDocument With {
                        .InvoiceId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                        .InvoiceNumber = Convert.ToString(reader("InvoiceNumber"), CultureInfo.InvariantCulture),
                        .InvoiceDate = ParseDate(reader("InvoiceDate")),
                        .PaymentMode = ConvertNullableString(reader("PaymentMode")),
                        .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                        .CustomerGstin = ConvertNullableString(reader("Gstin")),
                        .CustomerDrugLicenseNumber = ConvertNullableString(reader("DrugLicenseNumber")),
                        .CustomerAddress = ConvertNullableString(reader("Address")),
                        .CustomerPhone = ConvertNullableString(reader("Phone")),
                        .Notes = ConvertNullableString(reader("Notes")),
                        .SubTotal = Convert.ToDecimal(reader("SubTotal"), CultureInfo.InvariantCulture),
                        .DiscountAmount = Convert.ToDecimal(reader("DiscountAmount"), CultureInfo.InvariantCulture),
                        .SchemeAmount = Convert.ToDecimal(reader("SchemeAmount"), CultureInfo.InvariantCulture),
                        .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                        .RoundOffAmount = Convert.ToDecimal(reader("RoundOffAmount"), CultureInfo.InvariantCulture),
                        .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture),
                        .AmountPaid = Convert.ToDecimal(reader("AmountPaid"), CultureInfo.InvariantCulture),
                        .BalanceAmount = Convert.ToDecimal(reader("BalanceAmount"), CultureInfo.InvariantCulture)
                    }
                End Using
            End Using
        End Function

        Private Function LoadInvoiceItems(connection As DbConnection, invoiceId As Integer) As List(Of InvoiceDocumentItem)
            Dim items As New List(Of InvoiceDocumentItem)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT Id, BatchNumber, ExpiryDate, Quantity, FreeQuantity, Rate, MRP, DiscountPercentage, DiscountAmount, SchemeDescription, GstPercentage, TaxableAmount, GstAmount, LineTotal, " &
                    "(SELECT ProductName FROM Products WHERE Products.Id = InvoiceItems.ProductId) AS ProductName " &
                    "FROM InvoiceItems WHERE InvoiceId = @InvoiceId ORDER BY Id ASC;"
                command.AddParameter("@InvoiceId", invoiceId)

                Using reader = command.ExecuteReader()
                    Dim lineNumber As Integer = 1
                    While reader.Read()
                        items.Add(New InvoiceDocumentItem With {
                            .LineNumber = lineNumber,
                            .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                            .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                            .ExpiryDate = ParseDate(reader("ExpiryDate")),
                            .Quantity = Convert.ToInt32(reader("Quantity"), CultureInfo.InvariantCulture),
                            .FreeQuantity = Convert.ToInt32(reader("FreeQuantity"), CultureInfo.InvariantCulture),
                            .Rate = Convert.ToDecimal(reader("Rate"), CultureInfo.InvariantCulture),
                            .MRP = Convert.ToDecimal(reader("MRP"), CultureInfo.InvariantCulture),
                            .DiscountPercentage = Convert.ToDecimal(reader("DiscountPercentage"), CultureInfo.InvariantCulture),
                            .DiscountAmount = Convert.ToDecimal(reader("DiscountAmount"), CultureInfo.InvariantCulture),
                            .SchemeDescription = ConvertNullableString(reader("SchemeDescription")),
                            .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                            .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                            .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                            .LineTotal = Convert.ToDecimal(reader("LineTotal"), CultureInfo.InvariantCulture)
                        })
                        lineNumber += 1
                    End While
                End Using
            End Using

            Return items
        End Function

        Private Function GetSetting(connection As DbConnection, settingKey As String, defaultValue As String) As String
            Using command = connection.CreateCommand()
                command.CommandText = "SELECT SettingValue FROM Settings WHERE SettingKey = @SettingKey LIMIT 1;"
                command.AddParameter("@SettingKey", settingKey)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return defaultValue
                End If

                Return Convert.ToString(result, CultureInfo.InvariantCulture)
            End Using
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

        Private Function ConvertNullableString(value As Object) As String
            If value Is Nothing OrElse value Is DBNull.Value Then
                Return String.Empty
            End If

            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace
