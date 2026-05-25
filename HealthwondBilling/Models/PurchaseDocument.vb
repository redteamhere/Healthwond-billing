Namespace Models

    Public Class PurchaseDocument
        Public Property PurchaseId As Integer
        Public Property PurchaseNumber As String = String.Empty
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

        Public Property SupplierName As String = String.Empty
        Public Property SupplierAddress As String = String.Empty
        Public Property SupplierPhone As String = String.Empty
        Public Property SupplierEmail As String = String.Empty
        Public Property SupplierGstin As String = String.Empty
        Public Property SupplierDrugLicenseNumber As String = String.Empty

        Public Property CompanyName As String = String.Empty
        Public Property CompanyAddress As String = String.Empty
        Public Property CompanyPhone As String = String.Empty
        Public Property CompanyGstin As String = String.Empty
        Public Property CompanyDrugLicenseNumber As String = String.Empty

        Public Property SubTotal As Decimal
        Public Property DiscountAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property RoundOffAmount As Decimal
        Public Property NetAmount As Decimal
        Public Property TotalLines As Integer
        Public Property TotalUnits As Integer
        Public Property Items As New List(Of PurchaseDocumentItem)()
        Public Property TaxLines As New List(Of PurchaseTaxSummaryLine)()
    End Class

End Namespace
