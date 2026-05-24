Namespace Models

    Public Class PurchaseReturnLineItem
        Public Property LineNumber As Integer
        Public Property PurchaseItemId As Integer
        Public Property PurchaseId As Integer
        Public Property ProductId As Integer
        Public Property ProductName As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property ExpiryDate As DateTime
        Public Property PurchasedQuantity As Integer
        Public Property PurchasedFreeQuantity As Integer
        Public Property AlreadyReturnedQuantity As Integer
        Public Property AlreadyReturnedFreeQuantity As Integer
        Public Property RemainingQuantity As Integer
        Public Property RemainingFreeQuantity As Integer
        Public Property CurrentStock As Integer
        Public Property PTR As Decimal
        Public Property GstPercentage As Decimal
        Public Property ReturnQuantity As Integer
        Public Property ReturnFreeQuantity As Integer
        Public Property TaxableAmount As Decimal
        Public Property GstAmount As Decimal
        Public Property LineTotal As Decimal
    End Class

End Namespace
