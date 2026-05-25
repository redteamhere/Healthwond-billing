Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class AccountingService

        Private Shared ReadOnly VoucherTypes As IReadOnlyList(Of String) =
            New List(Of String) From {
                "Journal",
                "Receipt",
                "Payment",
                "Contra",
                "Debit Note",
                "Credit Note",
                "Expense"
            }

        Private ReadOnly _accountingRepository As IAccountingRepository

        Public Sub New(accountingRepository As IAccountingRepository)
            _accountingRepository = accountingRepository
        End Sub

        Public Function GetVoucherTypes() As IReadOnlyList(Of String)
            Return VoucherTypes
        End Function

        Public Async Function LoadAccountGroupsAsync() As Task(Of List(Of AccountGroupRecord))
            Return Await Task.Run(Function() _accountingRepository.LoadAccountGroups())
        End Function

        Public Async Function LoadLedgersAsync(searchTerm As String) As Task(Of List(Of LedgerRecord))
            Return Await Task.Run(Function() _accountingRepository.LoadLedgers(If(searchTerm, String.Empty).Trim()))
        End Function

        Public Async Function SaveLedgerAsync(record As LedgerRecord) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeLedger(record)
                    Dim validationMessage As String = ValidateLedger(record)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        Dim ledgerId As Integer = _accountingRepository.SaveLedger(record)
                        AppLogger.Info($"Ledger '{record.LedgerName}' saved with Id {ledgerId}.")
                        Return EntityOperationResult.Success($"Ledger {record.LedgerName} saved successfully.", ledgerId)
                    Catch ex As Exception
                        AppLogger.Error($"Ledger save failed for '{record.LedgerName}'.", ex)
                        Return EntityOperationResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Public Async Function GenerateNextVoucherNumberAsync(voucherType As String, voucherDate As DateTime) As Task(Of String)
            Return Await Task.Run(Function() _accountingRepository.GenerateNextVoucherNumber(voucherType, voucherDate.Date))
        End Function

        Public Async Function SaveManualVoucherAsync(draft As AccountingVoucherDraft, createdByUserId As Integer) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeVoucherDraft(draft)
                    Dim validationMessage As String = ValidateVoucherDraft(draft)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        Dim voucherId As Integer = _accountingRepository.SaveManualVoucher(draft, createdByUserId)
                        AppLogger.Info($"Manual voucher '{draft.VoucherNumber}' saved with Id {voucherId}.")
                        Return EntityOperationResult.Success($"Voucher {draft.VoucherNumber} saved successfully.", voucherId)
                    Catch ex As Exception
                        AppLogger.Error($"Voucher save failed for '{draft.VoucherNumber}'.", ex)
                        Return EntityOperationResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Public Async Function LoadVouchersAsync(fromDate As DateTime, toDate As DateTime, voucherType As String, searchTerm As String) As Task(Of List(Of VoucherHistoryRow))
            Return Await Task.Run(Function() _accountingRepository.LoadVouchers(fromDate.Date, toDate.Date, If(voucherType, String.Empty).Trim(), If(searchTerm, String.Empty).Trim()))
        End Function

        Public Async Function LoadLedgerStatementAsync(ledgerId As Integer, fromDate As DateTime, toDate As DateTime) As Task(Of List(Of LedgerStatementRow))
            Return Await Task.Run(Function() _accountingRepository.LoadLedgerStatement(ledgerId, fromDate.Date, toDate.Date))
        End Function

        Public Async Function LoadOverviewAsync(fromDate As DateTime, toDate As DateTime) As Task(Of AccountingOverview)
            Return Await Task.Run(Function() _accountingRepository.GetAccountingOverview(fromDate.Date, toDate.Date))
        End Function

        Public Sub RecalculateVoucherTotals(lines As IEnumerable(Of VoucherLineItem), ByRef totalDebit As Decimal, ByRef totalCredit As Decimal)
            totalDebit = 0D
            totalCredit = 0D

            For Each line As VoucherLineItem In lines
                If line Is Nothing OrElse line.LedgerId <= 0 Then
                    Continue For
                End If

                Dim normalizedAmount As Decimal = Decimal.Round(Math.Max(line.Amount, 0D), 2, MidpointRounding.AwayFromZero)
                If String.Equals(line.EntryType, "Cr", StringComparison.OrdinalIgnoreCase) Then
                    totalCredit += normalizedAmount
                Else
                    totalDebit += normalizedAmount
                End If
            Next

            totalDebit = Decimal.Round(totalDebit, 2, MidpointRounding.AwayFromZero)
            totalCredit = Decimal.Round(totalCredit, 2, MidpointRounding.AwayFromZero)
        End Sub

        Private Sub NormalizeLedger(record As LedgerRecord)
            record.LedgerName = If(record.LedgerName, String.Empty).Trim()
            record.OpeningBalance = Decimal.Round(Math.Max(record.OpeningBalance, 0D), 2, MidpointRounding.AwayFromZero)
            record.OpeningBalanceType = NormalizeEntryType(record.OpeningBalanceType)
            record.Notes = If(record.Notes, String.Empty).Trim()
        End Sub

        Private Function ValidateLedger(record As LedgerRecord) As String
            If Not InputValidator.IsRequiredTextProvided(record.LedgerName) Then
                Return "Ledger name is required."
            End If

            If record.AccountGroupId <= 0 Then
                Return "Select an account group for the ledger."
            End If

            If record.OpeningBalance < 0D Then
                Return "Opening balance cannot be negative."
            End If

            Return String.Empty
        End Function

        Private Sub NormalizeVoucherDraft(draft As AccountingVoucherDraft)
            draft.VoucherNumber = If(draft.VoucherNumber, String.Empty).Trim().ToUpperInvariant()
            draft.VoucherType = NormalizeVoucherType(draft.VoucherType)
            draft.ReferenceNumber = If(draft.ReferenceNumber, String.Empty).Trim().ToUpperInvariant()
            draft.Narration = If(draft.Narration, String.Empty).Trim()

            Dim lineNumber As Integer = 1
            For Each line As VoucherLineItem In draft.Lines
                line.LineNumber = lineNumber
                line.LedgerName = If(line.LedgerName, String.Empty).Trim()
                line.EntryType = NormalizeEntryType(line.EntryType)
                line.Amount = Decimal.Round(Math.Max(line.Amount, 0D), 2, MidpointRounding.AwayFromZero)
                line.Remarks = If(line.Remarks, String.Empty).Trim()
                lineNumber += 1
            Next
        End Sub

        Private Function ValidateVoucherDraft(draft As AccountingVoucherDraft) As String
            If Not InputValidator.IsRequiredTextProvided(draft.VoucherNumber) Then
                Return "Voucher number is required."
            End If

            If Not VoucherTypes.Contains(draft.VoucherType) Then
                Return "Select a valid voucher type."
            End If

            If draft.Lines Is Nothing OrElse draft.Lines.Count < 2 Then
                Return "Add at least two voucher lines."
            End If

            Dim totalDebit As Decimal
            Dim totalCredit As Decimal

            For Each line As VoucherLineItem In draft.Lines
                If line.LedgerId <= 0 Then
                    Return "Every voucher line must have a ledger."
                End If

                If line.Amount <= 0D Then
                    Return $"Voucher line {line.LineNumber.ToString(Globalization.CultureInfo.InvariantCulture)} must have an amount greater than zero."
                End If
            Next

            RecalculateVoucherTotals(draft.Lines, totalDebit, totalCredit)
            If totalDebit <= 0D OrElse totalCredit <= 0D Then
                Return "Voucher must contain both debit and credit amounts."
            End If

            If totalDebit <> totalCredit Then
                Return "Debit and credit totals must match."
            End If

            Return String.Empty
        End Function

        Private Function NormalizeVoucherType(voucherType As String) As String
            Dim normalizedValue As String = If(voucherType, String.Empty).Trim()

            For Each value As String In VoucherTypes
                If String.Equals(value, normalizedValue, StringComparison.OrdinalIgnoreCase) Then
                    Return value
                End If
            Next

            Return normalizedValue
        End Function

        Private Function NormalizeEntryType(entryType As String) As String
            Return If(String.Equals(If(entryType, String.Empty).Trim(), "Cr", StringComparison.OrdinalIgnoreCase), "Cr", "Dr")
        End Function

    End Class

End Namespace
