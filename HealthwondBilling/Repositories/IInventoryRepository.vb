Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IInventoryRepository
        Function GetInventorySummary(expiryWindowInDays As Integer) As InventorySummary
        Function GetCurrentStock(searchTerm As String) As List(Of InventoryCurrentStockRow)
        Function GetBatchStock(searchTerm As String) As List(Of InventoryBatchStockRow)
        Function GetExpiryStock(searchTerm As String, expiryWindowInDays As Integer) As List(Of InventoryExpiryRow)
        Function GetLowStock(searchTerm As String) As List(Of InventoryLowStockRow)
        Function GetStockLedger(searchTerm As String, fromDate As DateTime, toDate As DateTime) As List(Of InventoryLedgerRow)
    End Interface

End Namespace
