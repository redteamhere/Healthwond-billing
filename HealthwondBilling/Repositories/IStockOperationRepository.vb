Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IStockOperationRepository
        Function GenerateNextPurchaseReturnNumber(returnDate As DateTime) As String
        Function SearchPurchasesForReturn(fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of PurchaseHistoryLookupRow)
        Function GetPurchaseReturnLines(purchaseId As Integer) As List(Of PurchaseReturnLineItem)
        Function SavePurchaseReturn(draft As PurchaseReturnDraft, createdByUserId As Integer) As Integer
        Function GenerateNextStockAdjustmentNumber(adjustmentDate As DateTime) As String
        Function SaveStockAdjustment(draft As StockAdjustmentDraft, createdByUserId As Integer) As Integer
    End Interface

End Namespace
