Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IPurchaseRepository
        Function GenerateNextPurchaseNumber(purchaseDate As DateTime) As String
        Function SavePurchase(draft As PurchaseDraft, createdByUserId As Integer) As Integer
        Function GetPurchaseDocument(purchaseId As Integer) As PurchaseDocument
    End Interface

End Namespace
