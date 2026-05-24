Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports System.Threading.Tasks

Namespace Services

    Public Class InventoryService

        Private ReadOnly _inventoryRepository As IInventoryRepository

        Public Sub New(inventoryRepository As IInventoryRepository)
            _inventoryRepository = inventoryRepository
        End Sub

        Public Async Function GetInventorySummaryAsync(expiryWindowInDays As Integer) As Task(Of InventorySummary)
            Dim safeWindow As Integer = Math.Max(expiryWindowInDays, 0)
            Return Await Task.Run(Function() _inventoryRepository.GetInventorySummary(safeWindow))
        End Function

        Public Async Function GetCurrentStockAsync(searchTerm As String) As Task(Of List(Of InventoryCurrentStockRow))
            Return Await Task.Run(Function() _inventoryRepository.GetCurrentStock(NormalizeSearchTerm(searchTerm)))
        End Function

        Public Async Function GetBatchStockAsync(searchTerm As String) As Task(Of List(Of InventoryBatchStockRow))
            Return Await Task.Run(Function() _inventoryRepository.GetBatchStock(NormalizeSearchTerm(searchTerm)))
        End Function

        Public Async Function GetExpiryStockAsync(searchTerm As String, expiryWindowInDays As Integer) As Task(Of List(Of InventoryExpiryRow))
            Dim safeWindow As Integer = Math.Max(expiryWindowInDays, 0)
            Return Await Task.Run(Function() _inventoryRepository.GetExpiryStock(NormalizeSearchTerm(searchTerm), safeWindow))
        End Function

        Public Async Function GetLowStockAsync(searchTerm As String) As Task(Of List(Of InventoryLowStockRow))
            Return Await Task.Run(Function() _inventoryRepository.GetLowStock(NormalizeSearchTerm(searchTerm)))
        End Function

        Public Async Function GetStockLedgerAsync(searchTerm As String, fromDate As DateTime, toDate As DateTime) As Task(Of List(Of InventoryLedgerRow))
            Dim safeFromDate As DateTime = fromDate.Date
            Dim safeToDate As DateTime = toDate.Date

            If safeFromDate > safeToDate Then
                Throw New ArgumentException("The stock ledger date range is invalid.")
            End If

            Return Await Task.Run(Function() _inventoryRepository.GetStockLedger(NormalizeSearchTerm(searchTerm), safeFromDate, safeToDate))
        End Function

        Private Function NormalizeSearchTerm(searchTerm As String) As String
            Return If(searchTerm, String.Empty).Trim()
        End Function

    End Class

End Namespace
