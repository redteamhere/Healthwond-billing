Namespace Models

    Public Class PurchaseDraft
        Public Property PurchaseId As Integer
        Public Property PurchaseNumber As String = String.Empty
        Public Property SupplierId As Integer
        Public Property SupplierName As String = String.Empty
        Public Property PurchaseDate As DateTime
        Public Property SupplierInvoiceNumber As String = String.Empty
        Public Property Notes As String = String.Empty
        Public Property Items As New List(Of PurchaseLineItem)()
        Public Property Summary As New PurchaseTotalsSummary()
    End Class

End Namespace
