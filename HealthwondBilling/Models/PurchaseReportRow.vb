Namespace Models

    Public Class PurchaseReportRow
        Public Property PurchaseDate As DateTime
        Public Property PurchaseNumber As String = String.Empty
        Public Property SupplierName As String = String.Empty
        Public Property SupplierInvoiceNumber As String = String.Empty
        Public Property LineCount As Integer
        Public Property TotalUnits As Integer
        Public Property TaxableAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property NetAmount As Decimal
    End Class

End Namespace
