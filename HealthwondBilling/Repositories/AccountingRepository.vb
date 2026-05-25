Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization
Imports System.Linq

Namespace Repositories

    Public Class AccountingRepository
        Implements IAccountingRepository

        Private NotInheritable Class SignedBalance
            Public Property Amount As Decimal
            Public Property BalanceType As String = "Dr"
        End Class

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function LoadAccountGroups() As List(Of AccountGroupRecord) Implements IAccountingRepository.LoadAccountGroups
            Dim rows As New List(Of AccountGroupRecord)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT Id, GroupName, Nature, DisplayOrder, IsSystem " &
                        "FROM AccountGroups ORDER BY DisplayOrder ASC, GroupName ASC;"

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New AccountGroupRecord With {
                                .Id = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                                .GroupName = Convert.ToString(reader("GroupName"), CultureInfo.InvariantCulture),
                                .Nature = Convert.ToString(reader("Nature"), CultureInfo.InvariantCulture),
                                .DisplayOrder = Convert.ToInt32(reader("DisplayOrder"), CultureInfo.InvariantCulture),
                                .IsSystem = Convert.ToInt32(reader("IsSystem"), CultureInfo.InvariantCulture) = 1
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function LoadLedgers(searchTerm As String) As List(Of LedgerRecord) Implements IAccountingRepository.LoadLedgers
            Dim rows As New List(Of LedgerRecord)()
            Dim normalizedSearch As String = If(searchTerm, String.Empty).Trim()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT l.Id, l.LedgerName, l.AccountGroupId, ag.GroupName, l.OpeningBalance, l.OpeningBalanceType, l.IsSystem, l.IsPartyLedger, " &
                        "COALESCE(l.LinkedEntityType, '') AS LinkedEntityType, COALESCE(l.LinkedEntityId, 0) AS LinkedEntityId, COALESCE(l.Notes, '') AS Notes, " &
                        "COALESCE(SUM(CASE WHEN ave.EntryType = 'Dr' THEN ave.Amount ELSE -ave.Amount END), 0) AS NetMovement " &
                        "FROM Ledgers l " &
                        "INNER JOIN AccountGroups ag ON ag.Id = l.AccountGroupId " &
                        "LEFT JOIN AccountingVoucherEntries ave ON ave.LedgerId = l.Id " &
                        "WHERE (@Search = '' OR l.LedgerName LIKE @SearchLike OR ag.GroupName LIKE @SearchLike) " &
                        "GROUP BY l.Id, l.LedgerName, l.AccountGroupId, ag.GroupName, l.OpeningBalance, l.OpeningBalanceType, l.IsSystem, l.IsPartyLedger, l.LinkedEntityType, l.LinkedEntityId, l.Notes " &
                        "ORDER BY ag.DisplayOrder ASC, l.LedgerName ASC;"
                    command.AddParameter("@Search", normalizedSearch)
                    command.AddParameter("@SearchLike", $"%{normalizedSearch}%")

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            Dim signedBalance As Decimal = GetOpeningBalanceSigned(
                                Convert.ToDecimal(reader("OpeningBalance"), CultureInfo.InvariantCulture),
                                Convert.ToString(reader("OpeningBalanceType"), CultureInfo.InvariantCulture)) +
                                Convert.ToDecimal(reader("NetMovement"), CultureInfo.InvariantCulture)

                            Dim balance As SignedBalance = ToSignedBalance(signedBalance)

                            rows.Add(New LedgerRecord With {
                                .Id = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                                .LedgerName = Convert.ToString(reader("LedgerName"), CultureInfo.InvariantCulture),
                                .AccountGroupId = Convert.ToInt32(reader("AccountGroupId"), CultureInfo.InvariantCulture),
                                .AccountGroupName = Convert.ToString(reader("GroupName"), CultureInfo.InvariantCulture),
                                .OpeningBalance = Convert.ToDecimal(reader("OpeningBalance"), CultureInfo.InvariantCulture),
                                .OpeningBalanceType = Convert.ToString(reader("OpeningBalanceType"), CultureInfo.InvariantCulture),
                                .IsSystem = Convert.ToInt32(reader("IsSystem"), CultureInfo.InvariantCulture) = 1,
                                .IsPartyLedger = Convert.ToInt32(reader("IsPartyLedger"), CultureInfo.InvariantCulture) = 1,
                                .LinkedEntityType = Convert.ToString(reader("LinkedEntityType"), CultureInfo.InvariantCulture),
                                .LinkedEntityId = Convert.ToInt32(reader("LinkedEntityId"), CultureInfo.InvariantCulture),
                                .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture),
                                .CurrentBalance = balance.Amount,
                                .CurrentBalanceType = balance.BalanceType
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function SaveLedger(record As LedgerRecord) As Integer Implements IAccountingRepository.SaveLedger
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    EnsureLedgerNameUnique(connection, transaction, record)

                    If record.Id > 0 Then
                        EnsureLedgerEditable(connection, transaction, record.Id)
                        UpdateLedger(connection, transaction, record)
                        transaction.Commit()
                        Return record.Id
                    End If

                    Dim ledgerId As Integer = InsertLedger(connection, transaction, record)
                    transaction.Commit()
                    Return ledgerId
                End Using
            End Using
        End Function

        Public Function GenerateNextVoucherNumber(voucherType As String, voucherDate As DateTime) As String Implements IAccountingRepository.GenerateNextVoucherNumber
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim prefix As String = ResolveVoucherPrefix(voucherType)
                Dim monthPart As String = voucherDate.ToString("yyyyMM", CultureInfo.InvariantCulture)
                Dim pattern As String = $"{prefix}-{monthPart}-%"
                Dim lastNumber As Integer = 0

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT VoucherNumber FROM AccountingVouchers " &
                        "WHERE VoucherNumber LIKE @Pattern " &
                        "ORDER BY Id DESC LIMIT 1;"
                    command.AddParameter("@Pattern", pattern)
                    Dim result As Object = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                        Dim parts As String() = Convert.ToString(result, CultureInfo.InvariantCulture).Split("-"c)
                        Dim parsedNumber As Integer
                        If parts.Length >= 3 AndAlso Integer.TryParse(parts(parts.Length - 1), parsedNumber) Then
                            lastNumber = parsedNumber
                        End If
                    End If
                End Using

                Return $"{prefix}-{monthPart}-{(lastNumber + 1).ToString("0000", CultureInfo.InvariantCulture)}"
            End Using
        End Function

        Public Function SaveManualVoucher(draft As AccountingVoucherDraft, createdByUserId As Integer) As Integer Implements IAccountingRepository.SaveManualVoucher
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim voucherId As Integer = AccountingPostingHelper.SaveManualVoucher(connection, transaction, draft, createdByUserId)
                    transaction.Commit()
                    Return voucherId
                End Using
            End Using
        End Function

        Public Function LoadVouchers(fromDate As DateTime, toDate As DateTime, voucherType As String, searchTerm As String) As List(Of VoucherHistoryRow) Implements IAccountingRepository.LoadVouchers
            Dim rows As New List(Of VoucherHistoryRow)()
            Dim normalizedType As String = If(voucherType, String.Empty).Trim()
            Dim normalizedSearch As String = If(searchTerm, String.Empty).Trim()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT av.Id, av.VoucherNumber, av.VoucherType, av.VoucherDate, COALESCE(av.ReferenceNumber, '') AS ReferenceNumber, COALESCE(av.Narration, '') AS Narration, " &
                        "COALESCE(av.SourceType, '') AS SourceType, " &
                        "COALESCE(SUM(CASE WHEN ave.EntryType = 'Dr' THEN ave.Amount ELSE 0 END), 0) AS TotalDebit, " &
                        "COALESCE(SUM(CASE WHEN ave.EntryType = 'Cr' THEN ave.Amount ELSE 0 END), 0) AS TotalCredit, " &
                        "COUNT(ave.Id) AS EntryCount " &
                        "FROM AccountingVouchers av " &
                        "INNER JOIN AccountingVoucherEntries ave ON ave.VoucherId = av.Id " &
                        "WHERE date(av.VoucherDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "AND (@VoucherType = '' OR av.VoucherType = @VoucherType) " &
                        "AND (@Search = '' OR av.VoucherNumber LIKE @SearchLike OR COALESCE(av.ReferenceNumber, '') LIKE @SearchLike OR COALESCE(av.Narration, '') LIKE @SearchLike) " &
                        "GROUP BY av.Id, av.VoucherNumber, av.VoucherType, av.VoucherDate, av.ReferenceNumber, av.Narration, av.SourceType " &
                        "ORDER BY date(av.VoucherDate) DESC, av.Id DESC;"
                    command.AddParameter("@FromDate", SqliteDateHelper.ToStorageDate(fromDate.Date))
                    command.AddParameter("@ToDate", SqliteDateHelper.ToStorageDate(toDate.Date))
                    command.AddParameter("@VoucherType", normalizedType)
                    command.AddParameter("@Search", normalizedSearch)
                    command.AddParameter("@SearchLike", $"%{normalizedSearch}%")

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New VoucherHistoryRow With {
                                .VoucherId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                                .VoucherNumber = Convert.ToString(reader("VoucherNumber"), CultureInfo.InvariantCulture),
                                .VoucherType = Convert.ToString(reader("VoucherType"), CultureInfo.InvariantCulture),
                                .VoucherDate = ParseDate(reader("VoucherDate")),
                                .ReferenceNumber = Convert.ToString(reader("ReferenceNumber"), CultureInfo.InvariantCulture),
                                .Narration = Convert.ToString(reader("Narration"), CultureInfo.InvariantCulture),
                                .SourceType = Convert.ToString(reader("SourceType"), CultureInfo.InvariantCulture),
                                .TotalDebit = Convert.ToDecimal(reader("TotalDebit"), CultureInfo.InvariantCulture),
                                .TotalCredit = Convert.ToDecimal(reader("TotalCredit"), CultureInfo.InvariantCulture),
                                .EntryCount = Convert.ToInt32(reader("EntryCount"), CultureInfo.InvariantCulture)
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function LoadLedgerStatement(ledgerId As Integer, fromDate As DateTime, toDate As DateTime) As List(Of LedgerStatementRow) Implements IAccountingRepository.LoadLedgerStatement
            Dim rows As New List(Of LedgerStatementRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Dim ledger As LedgerRecord = LoadLedgers(String.Empty).FirstOrDefault(Function(item) item.Id = ledgerId)
                If ledger Is Nothing Then
                    Return rows
                End If

                Dim runningSignedBalance As Decimal = GetOpeningBalanceSigned(ledger.OpeningBalance, ledger.OpeningBalanceType)
                rows.Add(New LedgerStatementRow With {
                    .VoucherDate = fromDate.Date,
                    .VoucherNumber = "OPENING",
                    .VoucherType = "Opening",
                    .ReferenceNumber = String.Empty,
                    .Particulars = "Opening Balance",
                    .DebitAmount = If(ledger.OpeningBalanceType = "Dr", ledger.OpeningBalance, 0D),
                    .CreditAmount = If(ledger.OpeningBalanceType = "Cr", ledger.OpeningBalance, 0D),
                    .RunningBalance = Math.Abs(runningSignedBalance),
                    .RunningBalanceType = If(runningSignedBalance < 0D, "Cr", "Dr")
                })

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT av.VoucherDate, av.VoucherNumber, av.VoucherType, COALESCE(av.ReferenceNumber, '') AS ReferenceNumber, " &
                        "COALESCE(ave.Remarks, av.Narration, '') AS Particulars, ave.EntryType, ave.Amount " &
                        "FROM AccountingVoucherEntries ave " &
                        "INNER JOIN AccountingVouchers av ON av.Id = ave.VoucherId " &
                        "WHERE ave.LedgerId = @LedgerId " &
                        "AND date(av.VoucherDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "ORDER BY date(av.VoucherDate) ASC, av.Id ASC, ave.Id ASC;"
                    command.AddParameter("@LedgerId", ledgerId)
                    command.AddParameter("@FromDate", SqliteDateHelper.ToStorageDate(fromDate.Date))
                    command.AddParameter("@ToDate", SqliteDateHelper.ToStorageDate(toDate.Date))

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            Dim entryType As String = Convert.ToString(reader("EntryType"), CultureInfo.InvariantCulture)
                            Dim amount As Decimal = Convert.ToDecimal(reader("Amount"), CultureInfo.InvariantCulture)

                            If entryType = "Dr" Then
                                runningSignedBalance += amount
                            Else
                                runningSignedBalance -= amount
                            End If

                            Dim balance As SignedBalance = ToSignedBalance(runningSignedBalance)
                            rows.Add(New LedgerStatementRow With {
                                .VoucherDate = ParseDate(reader("VoucherDate")),
                                .VoucherNumber = Convert.ToString(reader("VoucherNumber"), CultureInfo.InvariantCulture),
                                .VoucherType = Convert.ToString(reader("VoucherType"), CultureInfo.InvariantCulture),
                                .ReferenceNumber = Convert.ToString(reader("ReferenceNumber"), CultureInfo.InvariantCulture),
                                .Particulars = Convert.ToString(reader("Particulars"), CultureInfo.InvariantCulture),
                                .DebitAmount = If(entryType = "Dr", amount, 0D),
                                .CreditAmount = If(entryType = "Cr", amount, 0D),
                                .RunningBalance = balance.Amount,
                                .RunningBalanceType = balance.BalanceType
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function GetAccountingOverview(fromDate As DateTime, toDate As DateTime) As AccountingOverview Implements IAccountingRepository.GetAccountingOverview
            Dim overview As New AccountingOverview With {
                .FromDate = fromDate.Date,
                .ToDate = toDate.Date
            }

            Using connection = _connectionFactory.CreateOpenConnection()
                overview.VoucherCount = ExecuteIntegerScalar(connection,
                    "SELECT COUNT(1) FROM AccountingVouchers WHERE date(VoucherDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                overview.ManualVoucherCount = ExecuteIntegerScalar(connection,
                    "SELECT COUNT(1) FROM AccountingVouchers WHERE date(VoucherDate) BETWEEN date(@FromDate) AND date(@ToDate) AND (SourceType IS NULL OR SourceType = '');",
                    fromDate,
                    toDate)
                overview.AutoVoucherCount = ExecuteIntegerScalar(connection,
                    "SELECT COUNT(1) FROM AccountingVouchers WHERE date(VoucherDate) BETWEEN date(@FromDate) AND date(@ToDate) AND SourceType IS NOT NULL AND SourceType <> '';",
                    fromDate,
                    toDate)
                overview.TotalDebit = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(ave.Amount), 0) FROM AccountingVoucherEntries ave INNER JOIN AccountingVouchers av ON av.Id = ave.VoucherId WHERE ave.EntryType = 'Dr' AND date(av.VoucherDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)
                overview.TotalCredit = ExecuteDecimalScalar(connection,
                    "SELECT COALESCE(SUM(ave.Amount), 0) FROM AccountingVoucherEntries ave INNER JOIN AccountingVouchers av ON av.Id = ave.VoucherId WHERE ave.EntryType = 'Cr' AND date(av.VoucherDate) BETWEEN date(@FromDate) AND date(@ToDate);",
                    fromDate,
                    toDate)

                Dim cashBalance As SignedBalance = LoadLedgerBalanceByName(connection, "Cash in Hand")
                overview.CashBalance = cashBalance.Amount
                overview.CashBalanceType = cashBalance.BalanceType

                Dim bankBalance As SignedBalance = LoadLedgerBalanceByName(connection, "Bank Account")
                overview.BankBalance = bankBalance.Amount
                overview.BankBalanceType = bankBalance.BalanceType

                Dim receivableBalance As SignedBalance = LoadGroupBalance(connection, "Sundry Debtors")
                overview.ReceivableBalance = receivableBalance.Amount
                overview.ReceivableBalanceType = receivableBalance.BalanceType

                Dim payableBalance As SignedBalance = LoadGroupBalance(connection, "Sundry Creditors")
                overview.PayableBalance = payableBalance.Amount
                overview.PayableBalanceType = payableBalance.BalanceType
            End Using

            Return overview
        End Function

        Public Sub SynchronizeOperationalVouchers() Implements IAccountingRepository.SynchronizeOperationalVouchers
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    AccountingPostingHelper.SyncPartyOpeningBalances(connection, transaction)

                    For Each draft In LoadInvoiceDraftsForSync(connection, transaction)
                        AccountingPostingHelper.PostInvoiceVoucher(connection, transaction, draft.InvoiceId, draft, draft.CreatedBy)
                    Next

                    For Each draft In LoadPurchaseDraftsForSync(connection, transaction)
                        AccountingPostingHelper.PostPurchaseVoucher(connection, transaction, draft.PurchaseId, draft, draft.CreatedBy)
                    Next

                    For Each draft In LoadPurchaseReturnDraftsForSync(connection, transaction)
                        AccountingPostingHelper.PostPurchaseReturnVoucher(connection, transaction, draft.ReturnId, draft, draft.CreatedBy)
                    Next

                    For Each draft In LoadCustomerPaymentsForSync(connection, transaction)
                        AccountingPostingHelper.PostCustomerPaymentVoucher(connection, transaction, draft.PaymentId, draft.Entry, draft.CreatedBy)
                    Next

                    For Each draft In LoadSupplierPaymentsForSync(connection, transaction)
                        AccountingPostingHelper.PostSupplierPaymentVoucher(connection, transaction, draft.PaymentId, draft.Entry, draft.CreatedBy)
                    Next

                    transaction.Commit()
                End Using
            End Using
        End Sub

        Private Function InsertLedger(connection As DbConnection, transaction As DbTransaction, record As LedgerRecord) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Ledgers (LedgerName, AccountGroupId, OpeningBalance, OpeningBalanceType, IsSystem, IsPartyLedger, LinkedEntityType, LinkedEntityId, Notes, CreatedAt, UpdatedAt) " &
                    "VALUES (@LedgerName, @AccountGroupId, @OpeningBalance, @OpeningBalanceType, 0, 0, NULL, NULL, @Notes, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@LedgerName", record.LedgerName)
                command.AddParameter("@AccountGroupId", record.AccountGroupId)
                command.AddParameter("@OpeningBalance", Decimal.Round(Math.Max(record.OpeningBalance, 0D), 2, MidpointRounding.AwayFromZero))
                command.AddParameter("@OpeningBalanceType", record.OpeningBalanceType)
                command.AddParameter("@Notes", record.Notes)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub UpdateLedger(connection As DbConnection, transaction As DbTransaction, record As LedgerRecord)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "UPDATE Ledgers SET LedgerName = @LedgerName, AccountGroupId = @AccountGroupId, OpeningBalance = @OpeningBalance, OpeningBalanceType = @OpeningBalanceType, Notes = @Notes, UpdatedAt = @UpdatedAt " &
                    "WHERE Id = @Id;"
                command.AddParameter("@LedgerName", record.LedgerName)
                command.AddParameter("@AccountGroupId", record.AccountGroupId)
                command.AddParameter("@OpeningBalance", Decimal.Round(Math.Max(record.OpeningBalance, 0D), 2, MidpointRounding.AwayFromZero))
                command.AddParameter("@OpeningBalanceType", record.OpeningBalanceType)
                command.AddParameter("@Notes", record.Notes)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", record.Id)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub EnsureLedgerNameUnique(connection As DbConnection, transaction As DbTransaction, record As LedgerRecord)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT COUNT(1) FROM Ledgers WHERE LOWER(LedgerName) = LOWER(@LedgerName) AND Id <> @Id;"
                command.AddParameter("@LedgerName", record.LedgerName)
                command.AddParameter("@Id", record.Id)
                If Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture) > 0 Then
                    Throw New InvalidOperationException("Another ledger already uses the same name.")
                End If
            End Using
        End Sub

        Private Sub EnsureLedgerEditable(connection As DbConnection, transaction As DbTransaction, ledgerId As Integer)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT IsSystem, IsPartyLedger FROM Ledgers WHERE Id = @Id LIMIT 1;"
                command.AddParameter("@Id", ledgerId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException("The selected ledger was not found.")
                    End If

                    Dim isSystem As Boolean = Convert.ToInt32(reader("IsSystem"), CultureInfo.InvariantCulture) = 1
                    Dim isPartyLedger As Boolean = Convert.ToInt32(reader("IsPartyLedger"), CultureInfo.InvariantCulture) = 1
                    If isSystem OrElse isPartyLedger Then
                        Throw New InvalidOperationException("System and party ledgers cannot be edited manually.")
                    End If
                End Using
            End Using
        End Sub

        Private Function LoadLedgerBalanceByName(connection As DbConnection, ledgerName As String) As SignedBalance
            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT COALESCE(CASE WHEN l.OpeningBalanceType = 'Dr' THEN l.OpeningBalance ELSE -l.OpeningBalance END, 0) + " &
                    "COALESCE(SUM(CASE WHEN ave.EntryType = 'Dr' THEN ave.Amount ELSE -ave.Amount END), 0) AS SignedBalance " &
                    "FROM Ledgers l " &
                    "LEFT JOIN AccountingVoucherEntries ave ON ave.LedgerId = l.Id " &
                    "WHERE l.LedgerName = @LedgerName " &
                    "GROUP BY l.Id, l.OpeningBalance, l.OpeningBalanceType;"
                command.AddParameter("@LedgerName", ledgerName)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return New SignedBalance()
                End If

                Return ToSignedBalance(Convert.ToDecimal(result, CultureInfo.InvariantCulture))
            End Using
        End Function

        Private Function LoadGroupBalance(connection As DbConnection, groupName As String) As SignedBalance
            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT COALESCE(SUM(CASE WHEN l.OpeningBalanceType = 'Dr' THEN l.OpeningBalance ELSE -l.OpeningBalance END), 0) + " &
                    "COALESCE(SUM(CASE WHEN ave.EntryType = 'Dr' THEN ave.Amount ELSE -ave.Amount END), 0) AS SignedBalance " &
                    "FROM Ledgers l " &
                    "INNER JOIN AccountGroups ag ON ag.Id = l.AccountGroupId " &
                    "LEFT JOIN AccountingVoucherEntries ave ON ave.LedgerId = l.Id " &
                    "WHERE ag.GroupName = @GroupName;"
                command.AddParameter("@GroupName", groupName)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return New SignedBalance()
                End If

                Return ToSignedBalance(Convert.ToDecimal(result, CultureInfo.InvariantCulture))
            End Using
        End Function

        Private Function ExecuteIntegerScalar(connection As DbConnection, sql As String, fromDate As DateTime, toDate As DateTime) As Integer
            Using command = connection.CreateCommand()
                command.CommandText = sql
                command.AddParameter("@FromDate", SqliteDateHelper.ToStorageDate(fromDate.Date))
                command.AddParameter("@ToDate", SqliteDateHelper.ToStorageDate(toDate.Date))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function ExecuteDecimalScalar(connection As DbConnection, sql As String, fromDate As DateTime, toDate As DateTime) As Decimal
            Using command = connection.CreateCommand()
                command.CommandText = sql
                command.AddParameter("@FromDate", SqliteDateHelper.ToStorageDate(fromDate.Date))
                command.AddParameter("@ToDate", SqliteDateHelper.ToStorageDate(toDate.Date))
                Return Convert.ToDecimal(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Shared Function ResolveVoucherPrefix(voucherType As String) As String
            Select Case If(voucherType, String.Empty).Trim().ToUpperInvariant()
                Case "RECEIPT"
                    Return "RV"
                Case "PAYMENT"
                    Return "PV"
                Case "CONTRA"
                    Return "CON"
                Case "DEBIT NOTE"
                    Return "DBN"
                Case "CREDIT NOTE"
                    Return "CRN"
                Case "EXPENSE"
                    Return "EXP"
                Case Else
                    Return "JRN"
            End Select
        End Function

        Private Shared Function GetOpeningBalanceSigned(openingBalance As Decimal, openingBalanceType As String) As Decimal
            Return If(String.Equals(openingBalanceType, "Cr", StringComparison.OrdinalIgnoreCase), -openingBalance, openingBalance)
        End Function

        Private Shared Function ToSignedBalance(signedValue As Decimal) As SignedBalance
            Return New SignedBalance With {
                .Amount = Decimal.Round(Math.Abs(signedValue), 2, MidpointRounding.AwayFromZero),
                .BalanceType = If(signedValue < 0D, "Cr", "Dr")
            }
        End Function

        Private Shared Function ParseDate(value As Object) As DateTime
            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return DateTime.Today
        End Function

        Private NotInheritable Class SyncInvoiceDraft
            Inherits BillingInvoiceDraft
            Public Property CreatedBy As Integer
        End Class

        Private NotInheritable Class SyncPurchaseDraft
            Inherits PurchaseDraft
            Public Property CreatedBy As Integer
        End Class

        Private NotInheritable Class SyncPurchaseReturnDraft
            Inherits PurchaseReturnDraft
            Public Property CreatedBy As Integer
        End Class

        Private NotInheritable Class SyncCustomerPaymentDraft
            Public Property PaymentId As Integer
            Public Property CreatedBy As Integer
            Public Property Entry As CustomerPaymentDraft
        End Class

        Private NotInheritable Class SyncSupplierPaymentDraft
            Public Property PaymentId As Integer
            Public Property CreatedBy As Integer
            Public Property Entry As SupplierPaymentDraft
        End Class

        Private Function LoadInvoiceDraftsForSync(connection As DbConnection, transaction As DbTransaction) As List(Of SyncInvoiceDraft)
            Dim rows As New List(Of SyncInvoiceDraft)()
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT i.Id, i.InvoiceNumber, i.CustomerId, c.CustomerName, i.InvoiceDate, COALESCE(i.PaymentMode, '') AS PaymentMode, i.AmountPaid, COALESCE(i.Notes, '') AS Notes, " &
                    "i.SubTotal, i.DiscountAmount, i.SchemeAmount, i.GstAmount, i.RoundOffAmount, i.NetAmount, i.BalanceAmount, i.CreatedBy " &
                    "FROM Invoices i INNER JOIN Customers c ON c.Id = i.CustomerId ORDER BY i.Id ASC;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New SyncInvoiceDraft With {
                            .InvoiceId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .InvoiceNumber = Convert.ToString(reader("InvoiceNumber"), CultureInfo.InvariantCulture),
                            .CustomerId = Convert.ToInt32(reader("CustomerId"), CultureInfo.InvariantCulture),
                            .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                            .InvoiceDate = ParseDate(reader("InvoiceDate")),
                            .PaymentMode = Convert.ToString(reader("PaymentMode"), CultureInfo.InvariantCulture),
                            .AmountPaid = Convert.ToDecimal(reader("AmountPaid"), CultureInfo.InvariantCulture),
                            .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture),
                            .CreatedBy = Convert.ToInt32(reader("CreatedBy"), CultureInfo.InvariantCulture),
                            .Summary = New BillingTotalsSummary With {
                                .SubTotal = Convert.ToDecimal(reader("SubTotal"), CultureInfo.InvariantCulture),
                                .DiscountAmount = Convert.ToDecimal(reader("DiscountAmount"), CultureInfo.InvariantCulture),
                                .SchemeAmount = Convert.ToDecimal(reader("SchemeAmount"), CultureInfo.InvariantCulture),
                                .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                                .RoundOffAmount = Convert.ToDecimal(reader("RoundOffAmount"), CultureInfo.InvariantCulture),
                                .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture),
                                .AmountPaid = Convert.ToDecimal(reader("AmountPaid"), CultureInfo.InvariantCulture),
                                .BalanceAmount = Convert.ToDecimal(reader("BalanceAmount"), CultureInfo.InvariantCulture)
                            }
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadPurchaseDraftsForSync(connection As DbConnection, transaction As DbTransaction) As List(Of SyncPurchaseDraft)
            Dim rows As New List(Of SyncPurchaseDraft)()
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT p.Id, p.PurchaseNumber, p.SupplierId, s.SupplierName, p.PurchaseDate, COALESCE(p.SupplierInvoiceNumber, '') AS SupplierInvoiceNumber, " &
                    "COALESCE(p.Notes, '') AS Notes, p.SubTotal, p.DiscountAmount, p.GstAmount, p.RoundOffAmount, p.NetAmount, p.CreatedBy " &
                    "FROM Purchases p INNER JOIN Suppliers s ON s.Id = p.SupplierId ORDER BY p.Id ASC;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New SyncPurchaseDraft With {
                            .PurchaseId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .PurchaseNumber = Convert.ToString(reader("PurchaseNumber"), CultureInfo.InvariantCulture),
                            .SupplierId = Convert.ToInt32(reader("SupplierId"), CultureInfo.InvariantCulture),
                            .SupplierName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                            .PurchaseDate = ParseDate(reader("PurchaseDate")),
                            .SupplierInvoiceNumber = Convert.ToString(reader("SupplierInvoiceNumber"), CultureInfo.InvariantCulture),
                            .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture),
                            .CreatedBy = Convert.ToInt32(reader("CreatedBy"), CultureInfo.InvariantCulture),
                            .Summary = New PurchaseTotalsSummary With {
                                .SubTotal = Convert.ToDecimal(reader("SubTotal"), CultureInfo.InvariantCulture),
                                .DiscountAmount = Convert.ToDecimal(reader("DiscountAmount"), CultureInfo.InvariantCulture),
                                .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                                .RoundOffAmount = Convert.ToDecimal(reader("RoundOffAmount"), CultureInfo.InvariantCulture),
                                .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture)
                            }
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadPurchaseReturnDraftsForSync(connection As DbConnection, transaction As DbTransaction) As List(Of SyncPurchaseReturnDraft)
            Dim rows As New List(Of SyncPurchaseReturnDraft)()
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT pr.Id, pr.ReturnNumber, pr.PurchaseId, p.PurchaseNumber, pr.SupplierId, s.SupplierName, pr.ReturnDate, COALESCE(pr.Notes, '') AS Notes, " &
                    "pr.SubTotal, pr.GstAmount, pr.RoundOffAmount, pr.NetAmount, pr.CreatedBy " &
                    "FROM PurchaseReturns pr " &
                    "INNER JOIN Purchases p ON p.Id = pr.PurchaseId " &
                    "INNER JOIN Suppliers s ON s.Id = pr.SupplierId " &
                    "ORDER BY pr.Id ASC;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New SyncPurchaseReturnDraft With {
                            .ReturnId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .ReturnNumber = Convert.ToString(reader("ReturnNumber"), CultureInfo.InvariantCulture),
                            .PurchaseId = Convert.ToInt32(reader("PurchaseId"), CultureInfo.InvariantCulture),
                            .PurchaseNumber = Convert.ToString(reader("PurchaseNumber"), CultureInfo.InvariantCulture),
                            .SupplierId = Convert.ToInt32(reader("SupplierId"), CultureInfo.InvariantCulture),
                            .SupplierName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                            .ReturnDate = ParseDate(reader("ReturnDate")),
                            .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture),
                            .CreatedBy = Convert.ToInt32(reader("CreatedBy"), CultureInfo.InvariantCulture),
                            .Summary = New PurchaseReturnSummary With {
                                .SubTotal = Convert.ToDecimal(reader("SubTotal"), CultureInfo.InvariantCulture),
                                .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                                .RoundOffAmount = Convert.ToDecimal(reader("RoundOffAmount"), CultureInfo.InvariantCulture),
                                .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture)
                            }
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadCustomerPaymentsForSync(connection As DbConnection, transaction As DbTransaction) As List(Of SyncCustomerPaymentDraft)
            Dim rows As New List(Of SyncCustomerPaymentDraft)()
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT cp.Id, cp.ReceiptNumber, cp.CustomerId, c.CustomerName, cp.PaymentDate, COALESCE(cp.PaymentMode, '') AS PaymentMode, " &
                    "COALESCE(cp.ReferenceNumber, '') AS ReferenceNumber, cp.Amount, COALESCE(cp.Notes, '') AS Notes, cp.CreatedBy " &
                    "FROM CustomerPayments cp INNER JOIN Customers c ON c.Id = cp.CustomerId ORDER BY cp.Id ASC;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New SyncCustomerPaymentDraft With {
                            .PaymentId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .CreatedBy = Convert.ToInt32(reader("CreatedBy"), CultureInfo.InvariantCulture),
                            .Entry = New CustomerPaymentDraft With {
                                .ReceiptNumber = Convert.ToString(reader("ReceiptNumber"), CultureInfo.InvariantCulture),
                                .CustomerId = Convert.ToInt32(reader("CustomerId"), CultureInfo.InvariantCulture),
                                .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                                .PaymentDate = ParseDate(reader("PaymentDate")),
                                .PaymentMode = Convert.ToString(reader("PaymentMode"), CultureInfo.InvariantCulture),
                                .ReferenceNumber = Convert.ToString(reader("ReferenceNumber"), CultureInfo.InvariantCulture),
                                .Amount = Convert.ToDecimal(reader("Amount"), CultureInfo.InvariantCulture),
                                .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture)
                            }
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

        Private Function LoadSupplierPaymentsForSync(connection As DbConnection, transaction As DbTransaction) As List(Of SyncSupplierPaymentDraft)
            Dim rows As New List(Of SyncSupplierPaymentDraft)()
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT sp.Id, sp.PaymentNumber, sp.SupplierId, s.SupplierName, sp.PaymentDate, COALESCE(sp.PaymentMode, '') AS PaymentMode, " &
                    "COALESCE(sp.ReferenceNumber, '') AS ReferenceNumber, sp.Amount, COALESCE(sp.Notes, '') AS Notes, sp.CreatedBy " &
                    "FROM SupplierPayments sp INNER JOIN Suppliers s ON s.Id = sp.SupplierId ORDER BY sp.Id ASC;"

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        rows.Add(New SyncSupplierPaymentDraft With {
                            .PaymentId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .CreatedBy = Convert.ToInt32(reader("CreatedBy"), CultureInfo.InvariantCulture),
                            .Entry = New SupplierPaymentDraft With {
                                .PaymentNumber = Convert.ToString(reader("PaymentNumber"), CultureInfo.InvariantCulture),
                                .SupplierId = Convert.ToInt32(reader("SupplierId"), CultureInfo.InvariantCulture),
                                .SupplierName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                                .PaymentDate = ParseDate(reader("PaymentDate")),
                                .PaymentMode = Convert.ToString(reader("PaymentMode"), CultureInfo.InvariantCulture),
                                .ReferenceNumber = Convert.ToString(reader("ReferenceNumber"), CultureInfo.InvariantCulture),
                                .Amount = Convert.ToDecimal(reader("Amount"), CultureInfo.InvariantCulture),
                                .Notes = Convert.ToString(reader("Notes"), CultureInfo.InvariantCulture)
                            }
                        })
                    End While
                End Using
            End Using

            Return rows
        End Function

    End Class

End Namespace
