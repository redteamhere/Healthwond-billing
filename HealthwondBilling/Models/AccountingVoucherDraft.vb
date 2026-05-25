Namespace Models

    Public Class AccountingVoucherDraft
        Public Property VoucherId As Integer
        Public Property VoucherNumber As String = String.Empty
        Public Property VoucherType As String = String.Empty
        Public Property VoucherDate As DateTime
        Public Property ReferenceNumber As String = String.Empty
        Public Property Narration As String = String.Empty
        Public Property SourceType As String = String.Empty
        Public Property SourceId As Integer
        Public Property Lines As New List(Of VoucherLineItem)()
    End Class

End Namespace
