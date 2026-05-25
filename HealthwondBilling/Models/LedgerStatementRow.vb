Namespace Models

    Public Class LedgerStatementRow
        Public Property VoucherDate As DateTime
        Public Property VoucherNumber As String = String.Empty
        Public Property VoucherType As String = String.Empty
        Public Property ReferenceNumber As String = String.Empty
        Public Property Particulars As String = String.Empty
        Public Property DebitAmount As Decimal
        Public Property CreditAmount As Decimal
        Public Property RunningBalance As Decimal
        Public Property RunningBalanceType As String = "Dr"
    End Class

End Namespace
