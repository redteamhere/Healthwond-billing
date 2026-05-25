Namespace Models

    Public Class ReportOverview
        Public Property FromDate As DateTime
        Public Property ToDate As DateTime
        Public Property SalesInvoiceCount As Integer
        Public Property PurchaseBillCount As Integer
        Public Property SalesUnits As Integer
        Public Property PurchaseUnits As Integer
        Public Property AverageSaleBillValue As Decimal
        Public Property AveragePurchaseBillValue As Decimal
        Public Property CustomerCollectionsAmount As Decimal
        Public Property SupplierPaymentsAmount As Decimal
        Public Property CollectionEfficiencyPercentage As Decimal
        Public Property SupplierPaymentCoveragePercentage As Decimal
        Public Property InventorySkuCount As Integer
        Public Property InventoryStockValueAtPTR As Decimal
        Public Property OutstandingReceivables As Decimal
        Public Property OutstandingPayables As Decimal
        Public Property NetCashMovement As Decimal
    End Class

End Namespace
