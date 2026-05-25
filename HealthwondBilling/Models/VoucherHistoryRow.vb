Namespace Models

    Public Class VoucherHistoryRow
        Public Property VoucherId As Integer
        Public Property VoucherNumber As String = String.Empty
        Public Property VoucherType As String = String.Empty
        Public Property VoucherDate As DateTime
        Public Property ReferenceNumber As String = String.Empty
        Public Property Narration As String = String.Empty
        Public Property SourceType As String = String.Empty
        Public Property TotalDebit As Decimal
        Public Property TotalCredit As Decimal
        Public Property EntryCount As Integer
    End Class

End Namespace
