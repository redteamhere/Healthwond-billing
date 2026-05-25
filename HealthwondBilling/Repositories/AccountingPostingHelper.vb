Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization
Imports System.Linq

Namespace Repositories

    Public NotInheritable Class AccountingPostingHelper

        Private NotInheritable Class PostingLine
            Public Property LedgerId As Integer
            Public Property EntryType As String = String.Empty
            Public Property Amount As Decimal
            Public Property Remarks As String = String.Empty
        End Class

        Private Sub New()
        End Sub

        Public Shared Sub PostInvoiceVoucher(connection As DbConnection, transaction As DbTransaction, invoiceId As Integer, draft As BillingInvoiceDraft, createdByUserId As Integer)
            Dim customerLedgerId As Integer = EnsurePartyLedger(connection, transaction, "Customer", draft.CustomerId, draft.CustomerName)
            Dim salesLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Sales Account")
            Dim outputGstLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Output GST")
            Dim roundOffLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Round Off")
            Dim cashLedgerId As Integer = ResolveSettlementLedgerId(connection, transaction, draft.PaymentMode)
            Dim taxableAmount As Decimal = Decimal.Round(Math.Max(draft.Summary.SubTotal - draft.Summary.DiscountAmount, 0D), 2, MidpointRounding.AwayFromZero)

            Dim lines As New List(Of PostingLine)()
            AddLine(lines, customerLedgerId, "Dr", draft.Summary.NetAmount, $"Invoice {draft.InvoiceNumber}")
            AddLine(lines, salesLedgerId, "Cr", taxableAmount, $"Sales invoice {draft.InvoiceNumber}")
            AddLine(lines, outputGstLedgerId, "Cr", draft.Summary.GstAmount, $"GST on invoice {draft.InvoiceNumber}")
            AddRoundOffLine(lines, roundOffLedgerId, draft.Summary.RoundOffAmount, "Cr", "Dr", $"Round off for invoice {draft.InvoiceNumber}")

            If draft.Summary.AmountPaid > 0D Then
                AddLine(lines, cashLedgerId, "Dr", draft.Summary.AmountPaid, $"Received on invoice {draft.InvoiceNumber}")
                AddLine(lines, customerLedgerId, "Cr", draft.Summary.AmountPaid, $"Settlement against invoice {draft.InvoiceNumber}")
            End If

            SaveSourceVoucher(
                connection,
                transaction,
                draft.InvoiceNumber,
                "Sales",
                draft.InvoiceDate,
                draft.InvoiceNumber,
                $"Sales invoice raised for {draft.CustomerName}",
                "Invoice",
                invoiceId,
                createdByUserId,
                lines)
        End Sub

        Public Shared Sub PostPurchaseVoucher(connection As DbConnection, transaction As DbTransaction, purchaseId As Integer, draft As PurchaseDraft, createdByUserId As Integer)
            Dim supplierLedgerId As Integer = EnsurePartyLedger(connection, transaction, "Supplier", draft.SupplierId, draft.SupplierName)
            Dim purchaseLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Purchase Account")
            Dim inputGstLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Input GST")
            Dim roundOffLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Round Off")

            Dim lines As New List(Of PostingLine)()
            AddLine(lines, purchaseLedgerId, "Dr", draft.Summary.SubTotal, $"Purchase {draft.PurchaseNumber}")
            AddLine(lines, inputGstLedgerId, "Dr", draft.Summary.GstAmount, $"GST on purchase {draft.PurchaseNumber}")
            AddRoundOffLine(lines, roundOffLedgerId, draft.Summary.RoundOffAmount, "Dr", "Cr", $"Round off for purchase {draft.PurchaseNumber}")
            AddLine(lines, supplierLedgerId, "Cr", draft.Summary.NetAmount, $"Payable for purchase {draft.PurchaseNumber}")

            SaveSourceVoucher(
                connection,
                transaction,
                draft.PurchaseNumber,
                "Purchase",
                draft.PurchaseDate,
                draft.SupplierInvoiceNumber,
                $"Purchase booked from {draft.SupplierName}",
                "Purchase",
                purchaseId,
                createdByUserId,
                lines)
        End Sub

        Public Shared Sub PostPurchaseReturnVoucher(connection As DbConnection, transaction As DbTransaction, returnId As Integer, draft As PurchaseReturnDraft, createdByUserId As Integer)
            Dim supplierLedgerId As Integer = EnsurePartyLedger(connection, transaction, "Supplier", draft.SupplierId, draft.SupplierName)
            Dim purchaseReturnLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Purchase Return Account")
            Dim inputGstLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Input GST")
            Dim roundOffLedgerId As Integer = EnsureSystemLedger(connection, transaction, "Round Off")

            Dim lines As New List(Of PostingLine)()
            AddLine(lines, supplierLedgerId, "Dr", draft.Summary.NetAmount, $"Purchase return {draft.ReturnNumber}")
            AddLine(lines, purchaseReturnLedgerId, "Cr", draft.Summary.SubTotal, $"Purchase return {draft.ReturnNumber}")
            AddLine(lines, inputGstLedgerId, "Cr", draft.Summary.GstAmount, $"Input GST reversal for {draft.ReturnNumber}")
            AddRoundOffLine(lines, roundOffLedgerId, draft.Summary.RoundOffAmount, "Cr", "Dr", $"Round off for purchase return {draft.ReturnNumber}")

            SaveSourceVoucher(
                connection,
                transaction,
                draft.ReturnNumber,
                "Purchase Return",
                draft.ReturnDate,
                draft.PurchaseNumber,
                $"Purchase return booked for {draft.SupplierName}",
                "PurchaseReturn",
                returnId,
                createdByUserId,
                lines)
        End Sub

        Public Shared Sub PostCustomerPaymentVoucher(connection As DbConnection, transaction As DbTransaction, paymentId As Integer, draft As CustomerPaymentDraft, createdByUserId As Integer)
            Dim customerLedgerId As Integer = EnsurePartyLedger(connection, transaction, "Customer", draft.CustomerId, draft.CustomerName)
            Dim settlementLedgerId As Integer = ResolveSettlementLedgerId(connection, transaction, draft.PaymentMode)

            Dim lines As New List(Of PostingLine)()
            AddLine(lines, settlementLedgerId, "Dr", draft.Amount, $"Collection {draft.ReceiptNumber}")
            AddLine(lines, customerLedgerId, "Cr", draft.Amount, $"Collection from {draft.CustomerName}")

            SaveSourceVoucher(
                connection,
                transaction,
                draft.ReceiptNumber,
                "Receipt",
                draft.PaymentDate,
                draft.ReferenceNumber,
                $"Collection received from {draft.CustomerName}",
                "CustomerPayment",
                paymentId,
                createdByUserId,
                lines)
        End Sub

        Public Shared Sub PostSupplierPaymentVoucher(connection As DbConnection, transaction As DbTransaction, paymentId As Integer, draft As SupplierPaymentDraft, createdByUserId As Integer)
            Dim supplierLedgerId As Integer = EnsurePartyLedger(connection, transaction, "Supplier", draft.SupplierId, draft.SupplierName)
            Dim settlementLedgerId As Integer = ResolveSettlementLedgerId(connection, transaction, draft.PaymentMode)

            Dim lines As New List(Of PostingLine)()
            AddLine(lines, supplierLedgerId, "Dr", draft.Amount, $"Payment {draft.PaymentNumber}")
            AddLine(lines, settlementLedgerId, "Cr", draft.Amount, $"Payment to {draft.SupplierName}")

            SaveSourceVoucher(
                connection,
                transaction,
                draft.PaymentNumber,
                "Payment",
                draft.PaymentDate,
                draft.ReferenceNumber,
                $"Payment made to {draft.SupplierName}",
                "SupplierPayment",
                paymentId,
                createdByUserId,
                lines)
        End Sub

        Public Shared Function SaveManualVoucher(connection As DbConnection, transaction As DbTransaction, draft As AccountingVoucherDraft, createdByUserId As Integer) As Integer
            Return SaveSourceVoucher(
                connection,
                transaction,
                draft.VoucherNumber,
                draft.VoucherType,
                draft.VoucherDate,
                draft.ReferenceNumber,
                draft.Narration,
                Nothing,
                0,
                createdByUserId,
                draft.Lines.Select(
                    Function(line) New PostingLine With {
                        .LedgerId = line.LedgerId,
                        .EntryType = line.EntryType,
                        .Amount = line.Amount,
                        .Remarks = line.Remarks
                    }).ToList())
        End Function

        Public Shared Sub SyncPartyOpeningBalances(connection As DbConnection, transaction As DbTransaction)
            SyncCustomerOpeningBalances(connection, transaction)
            SyncSupplierOpeningBalances(connection, transaction)
        End Sub

        Private Shared Function SaveSourceVoucher(connection As DbConnection, transaction As DbTransaction, voucherNumber As String, voucherType As String, voucherDate As DateTime, referenceNumber As String, narration As String, sourceType As String, sourceId As Integer, createdByUserId As Integer, lines As List(Of PostingLine)) As Integer
            Dim effectiveLines As List(Of PostingLine) = lines.
                Where(Function(line) line IsNot Nothing AndAlso line.LedgerId > 0 AndAlso line.Amount > 0D AndAlso (line.EntryType = "Dr" OrElse line.EntryType = "Cr")).
                Select(Function(line) New PostingLine With {
                    .LedgerId = line.LedgerId,
                    .EntryType = line.EntryType,
                    .Amount = Decimal.Round(line.Amount, 2, MidpointRounding.AwayFromZero),
                    .Remarks = If(line.Remarks, String.Empty).Trim()
                }).
                ToList()

            If effectiveLines.Count = 0 Then
                Throw New InvalidOperationException("No accounting lines were generated for the voucher.")
            End If

            DeleteExistingSourceVoucher(connection, transaction, sourceType, sourceId)

            Dim createdAt As String = SqliteDateHelper.ToStorageDateTime(DateTime.Now)
            Dim voucherId As Integer

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO AccountingVouchers (VoucherNumber, VoucherType, VoucherDate, ReferenceNumber, Narration, SourceType, SourceId, CreatedBy, CreatedAt, UpdatedAt) " &
                    "VALUES (@VoucherNumber, @VoucherType, @VoucherDate, @ReferenceNumber, @Narration, @SourceType, @SourceId, @CreatedBy, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@VoucherNumber", voucherNumber)
                command.AddParameter("@VoucherType", voucherType)
                command.AddParameter("@VoucherDate", SqliteDateHelper.ToStorageDate(voucherDate))
                command.AddParameter("@ReferenceNumber", If(referenceNumber, String.Empty).Trim())
                command.AddParameter("@Narration", If(narration, String.Empty).Trim())
                command.AddParameter("@SourceType", ToNullableValue(sourceType))
                command.AddParameter("@SourceId", If(sourceType Is Nothing OrElse sourceId <= 0, CType(DBNull.Value, Object), sourceId))
                command.AddParameter("@CreatedBy", createdByUserId)
                command.AddParameter("@CreatedAt", createdAt)
                command.AddParameter("@UpdatedAt", createdAt)
                voucherId = Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using

            For Each line As PostingLine In effectiveLines
                Using command = connection.CreateCommand()
                    command.Transaction = transaction
                    command.CommandText =
                        "INSERT INTO AccountingVoucherEntries (VoucherId, LedgerId, EntryType, Amount, Remarks) " &
                        "VALUES (@VoucherId, @LedgerId, @EntryType, @Amount, @Remarks);"
                    command.AddParameter("@VoucherId", voucherId)
                    command.AddParameter("@LedgerId", line.LedgerId)
                    command.AddParameter("@EntryType", line.EntryType)
                    command.AddParameter("@Amount", line.Amount)
                    command.AddParameter("@Remarks", line.Remarks)
                    command.ExecuteNonQuery()
                End Using
            Next

            Return voucherId
        End Function

        Private Shared Sub DeleteExistingSourceVoucher(connection As DbConnection, transaction As DbTransaction, sourceType As String, sourceId As Integer)
            If String.IsNullOrWhiteSpace(sourceType) OrElse sourceId <= 0 Then
                Return
            End If

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "DELETE FROM AccountingVouchers WHERE SourceType = @SourceType AND SourceId = @SourceId;"
                command.AddParameter("@SourceType", sourceType)
                command.AddParameter("@SourceId", sourceId)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Shared Function EnsurePartyLedger(connection As DbConnection, transaction As DbTransaction, entityType As String, entityId As Integer, partyName As String) As Integer
            Dim normalizedName As String = If(partyName, String.Empty).Trim()
            If normalizedName = String.Empty Then
                normalizedName = $"{entityType} {entityId.ToString(CultureInfo.InvariantCulture)}"
            End If

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT Id FROM Ledgers WHERE LinkedEntityType = @LinkedEntityType AND LinkedEntityId = @LinkedEntityId LIMIT 1;"
                command.AddParameter("@LinkedEntityType", entityType)
                command.AddParameter("@LinkedEntityId", entityId)
                Dim result As Object = command.ExecuteScalar()
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Dim ledgerId As Integer = Convert.ToInt32(result, CultureInfo.InvariantCulture)
                    Using updateCommand = connection.CreateCommand()
                        updateCommand.Transaction = transaction
                        updateCommand.CommandText = "UPDATE Ledgers SET LedgerName = @LedgerName, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                        updateCommand.AddParameter("@LedgerName", normalizedName)
                        updateCommand.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                        updateCommand.AddParameter("@Id", ledgerId)
                        updateCommand.ExecuteNonQuery()
                    End Using
                    Return ledgerId
                End If
            End Using

            Dim groupName As String = If(String.Equals(entityType, "Customer", StringComparison.OrdinalIgnoreCase), "Sundry Debtors", "Sundry Creditors")
            Dim openingType As String = If(String.Equals(entityType, "Customer", StringComparison.OrdinalIgnoreCase), "Dr", "Cr")

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Ledgers (LedgerName, AccountGroupId, OpeningBalance, OpeningBalanceType, IsSystem, IsPartyLedger, LinkedEntityType, LinkedEntityId, Notes, CreatedAt, UpdatedAt) " &
                    "VALUES (@LedgerName, (SELECT Id FROM AccountGroups WHERE GroupName = @GroupName LIMIT 1), 0, @OpeningBalanceType, 0, 1, @LinkedEntityType, @LinkedEntityId, '', @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@LedgerName", normalizedName)
                command.AddParameter("@GroupName", groupName)
                command.AddParameter("@OpeningBalanceType", openingType)
                command.AddParameter("@LinkedEntityType", entityType)
                command.AddParameter("@LinkedEntityId", entityId)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Shared Function EnsureSystemLedger(connection As DbConnection, transaction As DbTransaction, ledgerName As String) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT Id FROM Ledgers WHERE LedgerName = @LedgerName LIMIT 1;"
                command.AddParameter("@LedgerName", ledgerName)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Throw New InvalidOperationException($"System ledger '{ledgerName}' was not found.")
                End If

                Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Shared Function ResolveSettlementLedgerId(connection As DbConnection, transaction As DbTransaction, paymentMode As String) As Integer
            Dim normalizedMode As String = If(paymentMode, String.Empty).Trim().ToUpperInvariant()
            Dim useBank As Boolean =
                normalizedMode.Contains("BANK") OrElse
                normalizedMode.Contains("CARD") OrElse
                normalizedMode.Contains("UPI") OrElse
                normalizedMode.Contains("ONLINE") OrElse
                normalizedMode.Contains("CHEQUE") OrElse
                normalizedMode.Contains("RTGS") OrElse
                normalizedMode.Contains("NEFT")

            Return EnsureSystemLedger(connection, transaction, If(useBank, "Bank Account", "Cash in Hand"))
        End Function

        Private Shared Sub AddLine(lines As IList(Of PostingLine), ledgerId As Integer, entryType As String, amount As Decimal, remarks As String)
            Dim normalizedAmount As Decimal = Decimal.Round(Math.Max(amount, 0D), 2, MidpointRounding.AwayFromZero)
            If normalizedAmount <= 0D Then
                Return
            End If

            lines.Add(New PostingLine With {
                .LedgerId = ledgerId,
                .EntryType = entryType,
                .Amount = normalizedAmount,
                .Remarks = remarks
            })
        End Sub

        Private Shared Sub AddRoundOffLine(lines As IList(Of PostingLine), roundOffLedgerId As Integer, roundOffAmount As Decimal, positiveEntryType As String, negativeEntryType As String, remarks As String)
            If roundOffAmount = 0D Then
                Return
            End If

            Dim entryType As String = If(roundOffAmount > 0D, positiveEntryType, negativeEntryType)
            AddLine(lines, roundOffLedgerId, entryType, Math.Abs(roundOffAmount), remarks)
        End Sub

        Private Shared Function ToNullableValue(value As String) As Object
            Dim normalizedValue As String = If(value, String.Empty).Trim()
            If normalizedValue = String.Empty Then
                Return DBNull.Value
            End If

            Return normalizedValue
        End Function

        Private Shared Sub SyncCustomerOpeningBalances(connection As DbConnection, transaction As DbTransaction)
            Dim rows As New List(Of Tuple(Of Integer, String, Decimal, Decimal))()

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT c.Id, c.CustomerName, c.OutstandingBalance, " &
                    "COALESCE((SELECT SUM(BalanceAmount) FROM Invoices WHERE CustomerId = c.Id), 0) AS DocumentOutstanding " &
                    "FROM Customers c;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(Tuple.Create(
                            Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                            Convert.ToDecimal(reader("OutstandingBalance"), CultureInfo.InvariantCulture),
                            Convert.ToDecimal(reader("DocumentOutstanding"), CultureInfo.InvariantCulture)))
                    End While
                End Using
            End Using

            For Each row In rows
                Dim ledgerId As Integer = EnsurePartyLedger(connection, transaction, "Customer", row.Item1, row.Item2)
                If Not CanApplyOpeningBalance(connection, transaction, ledgerId) Then
                    Continue For
                End If

                UpdateLedgerOpeningBalance(connection, transaction, ledgerId, row.Item3 - row.Item4, "Dr")
            Next
        End Sub

        Private Shared Sub SyncSupplierOpeningBalances(connection As DbConnection, transaction As DbTransaction)
            Dim rows As New List(Of Tuple(Of Integer, String, Decimal, Decimal))()

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT s.Id, s.SupplierName, s.OutstandingBalance, " &
                    "COALESCE((SELECT SUM(NetAmount) FROM Purchases WHERE SupplierId = s.Id), 0) " &
                    "- COALESCE((SELECT SUM(NetAmount) FROM PurchaseReturns WHERE SupplierId = s.Id), 0) " &
                    "- COALESCE((SELECT SUM(Amount) FROM SupplierPayments WHERE SupplierId = s.Id), 0) AS DocumentOutstanding " &
                    "FROM Suppliers s;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(Tuple.Create(
                            Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                            Convert.ToDecimal(reader("OutstandingBalance"), CultureInfo.InvariantCulture),
                            Convert.ToDecimal(reader("DocumentOutstanding"), CultureInfo.InvariantCulture)))
                    End While
                End Using
            End Using

            For Each row In rows
                Dim ledgerId As Integer = EnsurePartyLedger(connection, transaction, "Supplier", row.Item1, row.Item2)
                If Not CanApplyOpeningBalance(connection, transaction, ledgerId) Then
                    Continue For
                End If

                UpdateLedgerOpeningBalance(connection, transaction, ledgerId, row.Item3 - row.Item4, "Cr")
            Next
        End Sub

        Private Shared Function CanApplyOpeningBalance(connection As DbConnection, transaction As DbTransaction, ledgerId As Integer) As Boolean
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT COUNT(1) FROM AccountingVoucherEntries WHERE LedgerId = @LedgerId;"
                command.AddParameter("@LedgerId", ledgerId)
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture) = 0
            End Using
        End Function

        Private Shared Sub UpdateLedgerOpeningBalance(connection As DbConnection, transaction As DbTransaction, ledgerId As Integer, signedDifference As Decimal, positiveBalanceType As String)
            Dim openingBalance As Decimal = Decimal.Round(Math.Abs(signedDifference), 2, MidpointRounding.AwayFromZero)
            Dim openingType As String = positiveBalanceType

            If signedDifference < 0D Then
                openingType = If(String.Equals(positiveBalanceType, "Dr", StringComparison.OrdinalIgnoreCase), "Cr", "Dr")
            End If

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "UPDATE Ledgers SET OpeningBalance = @OpeningBalance, OpeningBalanceType = @OpeningBalanceType, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@OpeningBalance", openingBalance)
                command.AddParameter("@OpeningBalanceType", openingType)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", ledgerId)
                command.ExecuteNonQuery()
            End Using
        End Sub

    End Class

End Namespace
