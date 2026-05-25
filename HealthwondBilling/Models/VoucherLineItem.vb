Namespace Models

    Public Class VoucherLineItem
        Public Property LineNumber As Integer
        Public Property LedgerId As Integer
        Public Property LedgerName As String = String.Empty
        Public Property EntryType As String = "Dr"
        Public Property Amount As Decimal
        Public Property Remarks As String = String.Empty
    End Class

End Namespace
