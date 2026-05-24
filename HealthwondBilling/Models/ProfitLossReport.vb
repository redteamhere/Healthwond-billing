Namespace Models

    Public Class ProfitLossReport
        Public Property FromDate As DateTime
        Public Property ToDate As DateTime
        Public Property SalesTaxableAmount As Decimal
        Public Property SalesNetAmount As Decimal
        Public Property PurchaseTaxableAmount As Decimal
        Public Property PurchaseNetAmount As Decimal
        Public Property EstimatedCostOfGoodsSold As Decimal
        Public Property EstimatedGrossProfit As Decimal
        Public Property GrossMarginPercentage As Decimal
        Public Property OutstandingReceivables As Decimal
        Public Property OutstandingPayables As Decimal
    End Class

End Namespace
