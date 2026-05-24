Namespace Models

    Public Class PurchaseReturnDraft
        Public Property ReturnId As Integer
        Public Property ReturnNumber As String = String.Empty
        Public Property PurchaseId As Integer
        Public Property PurchaseNumber As String = String.Empty
        Public Property SupplierId As Integer
        Public Property SupplierName As String = String.Empty
        Public Property ReturnDate As DateTime
        Public Property Notes As String = String.Empty
        Public Property Items As New List(Of PurchaseReturnLineItem)()
        Public Property Summary As New PurchaseReturnSummary()
    End Class

End Namespace
