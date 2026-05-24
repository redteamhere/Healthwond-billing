Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class InvoiceRepository
        Implements IInvoiceRepository

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

                    UpdateCustomerOutstanding(connection, transaction, draft.CustomerId, draft.Summary.BalanceAmount)
                    transaction.Commit()
                    Return invoiceId
                End Using
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
                selectCommand.CommandText = "SELECT CurrentStock FROM Products WHERE Id = @Id AND IsDeleted = 0 LIMIT 1;"
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

        Private Sub UpdateCustomerOutstanding(connection As DbConnection, transaction As DbTransaction, customerId As Integer, balanceToAdd As Decimal)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "UPDATE Customers SET OutstandingBalance = OutstandingBalance + @BalanceToAdd, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@BalanceToAdd", balanceToAdd)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", customerId)
                command.ExecuteNonQuery()
            End Using
        End Sub

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

    End Class

End Namespace
