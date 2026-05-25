Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IAccountingRepository
        Function LoadAccountGroups() As List(Of AccountGroupRecord)
        Function LoadLedgers(searchTerm As String) As List(Of LedgerRecord)
        Function SaveLedger(record As LedgerRecord) As Integer
        Function GenerateNextVoucherNumber(voucherType As String, voucherDate As DateTime) As String
        Function SaveManualVoucher(draft As AccountingVoucherDraft, createdByUserId As Integer) As Integer
        Function LoadVouchers(fromDate As DateTime, toDate As DateTime, voucherType As String, searchTerm As String) As List(Of VoucherHistoryRow)
        Function LoadLedgerStatement(ledgerId As Integer, fromDate As DateTime, toDate As DateTime) As List(Of LedgerStatementRow)
        Function GetAccountingOverview(fromDate As DateTime, toDate As DateTime) As AccountingOverview
        Sub SynchronizeOperationalVouchers()
    End Interface

End Namespace
