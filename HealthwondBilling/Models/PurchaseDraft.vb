Namespace Models

    Public Class PurchaseDraft
        Public Property PurchaseId As Integer
        Public Property PurchaseNumber As String = String.Empty
        Public Property SupplierId As Integer
        Public Property SupplierName As String = String.Empty
        Public Property PurchaseDate As DateTime
        Public Property SupplierInvoiceNumber As String = String.Empty
        Public Property SupplierInvoiceDate As DateTime?
        Public Property PurchaseOrderNumber As String = String.Empty
        Public Property PurchaseOrderDate As DateTime?
        Public Property PlaceOfSupply As String = String.Empty
        Public Property CaseCount As Integer
        Public Property TransportName As String = String.Empty
        Public Property EwayBillNumber As String = String.Empty
        Public Property Notes As String = String.Empty
        Public Property Items As New List(Of PurchaseLineItem)()
        Public Property Summary As New PurchaseTotalsSummary()
    End Class

End Namespace
