Namespace Models

    Public Class AccountingOverview
        Public Property FromDate As DateTime
        Public Property ToDate As DateTime
        Public Property VoucherCount As Integer
        Public Property ManualVoucherCount As Integer
        Public Property AutoVoucherCount As Integer
        Public Property TotalDebit As Decimal
        Public Property TotalCredit As Decimal
        Public Property CashBalance As Decimal
        Public Property CashBalanceType As String = "Dr"
        Public Property BankBalance As Decimal
        Public Property BankBalanceType As String = "Dr"
        Public Property ReceivableBalance As Decimal
        Public Property ReceivableBalanceType As String = "Dr"
        Public Property PayableBalance As Decimal
        Public Property PayableBalanceType As String = "Cr"
    End Class

End Namespace
