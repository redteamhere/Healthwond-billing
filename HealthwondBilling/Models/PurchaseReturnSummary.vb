Namespace Models

    Public Class PurchaseReturnSummary
        Public Property SubTotal As Decimal
        Public Property GstAmount As Decimal
        Public Property RoundOffAmount As Decimal
        Public Property NetAmount As Decimal
        Public Property TotalUnits As Integer
        Public Property TotalLines As Integer
        Public Property TaxSummaryText As String = String.Empty
    End Class

End Namespace
