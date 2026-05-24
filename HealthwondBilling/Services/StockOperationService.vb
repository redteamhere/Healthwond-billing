Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities
Imports System.Globalization
Imports System.Linq

Namespace Services

    Public Class StockOperationService

        Private ReadOnly _stockOperationRepository As IStockOperationRepository
        Private ReadOnly _productRepository As IProductRepository

        Public Sub New(stockOperationRepository As IStockOperationRepository, productRepository As IProductRepository)
            _stockOperationRepository = stockOperationRepository
            _productRepository = productRepository
        End Sub

        Public Async Function GenerateNextPurchaseReturnNumberAsync(returnDate As DateTime) As Task(Of String)
            Return Await Task.Run(Function() _stockOperationRepository.GenerateNextPurchaseReturnNumber(returnDate.Date))
        End Function

        Public Async Function SearchPurchasesForReturnAsync(fromDate As DateTime, toDate As DateTime, searchTerm As String) As Task(Of List(Of PurchaseHistoryLookupRow))
            Return Await Task.Run(Function() _stockOperationRepository.SearchPurchasesForReturn(fromDate.Date, toDate.Date, NormalizeSearchTerm(searchTerm)))
        End Function

        Public Async Function GetPurchaseReturnLinesAsync(purchaseId As Integer) As Task(Of List(Of PurchaseReturnLineItem))
            Return Await Task.Run(Function() _stockOperationRepository.GetPurchaseReturnLines(purchaseId))
        End Function

        Public Sub RecalculatePurchaseReturnLine(line As PurchaseReturnLineItem)
            line.ReturnQuantity = Math.Max(line.ReturnQuantity, 0)
            line.ReturnFreeQuantity = Math.Max(line.ReturnFreeQuantity, 0)
            line.PTR = Math.Max(line.PTR, 0D)
            line.GstPercentage = Math.Max(0D, Math.Min(100D, line.GstPercentage))

            line.TaxableAmount = Decimal.Round(line.ReturnQuantity * line.PTR, 2, MidpointRounding.AwayFromZero)
            line.GstAmount = Decimal.Round(line.TaxableAmount * (line.GstPercentage / 100D), 2, MidpointRounding.AwayFromZero)
            line.LineTotal = Decimal.Round(line.TaxableAmount + line.GstAmount, 2, MidpointRounding.AwayFromZero)
        End Sub

        Public Function CalculatePurchaseReturnTotals(items As IEnumerable(Of PurchaseReturnLineItem)) As PurchaseReturnSummary
            Dim summary As New PurchaseReturnSummary()
            Dim selectedLines As List(Of PurchaseReturnLineItem) = items.
                Where(Function(item) item IsNot Nothing AndAlso (item.ReturnQuantity > 0 OrElse item.ReturnFreeQuantity > 0)).
                ToList()

            For Each line As PurchaseReturnLineItem In selectedLines
                RecalculatePurchaseReturnLine(line)
            Next

            summary.TotalLines = selectedLines.Count
            summary.TotalUnits = selectedLines.Sum(Function(line) line.ReturnQuantity + line.ReturnFreeQuantity)
            summary.SubTotal = Decimal.Round(selectedLines.Sum(Function(line) line.TaxableAmount), 2, MidpointRounding.AwayFromZero)
            summary.GstAmount = Decimal.Round(selectedLines.Sum(Function(line) line.GstAmount), 2, MidpointRounding.AwayFromZero)

            Dim exactNet As Decimal = selectedLines.Sum(Function(line) line.LineTotal)
            Dim roundedNet As Decimal = Decimal.Round(exactNet, 0, MidpointRounding.AwayFromZero)
            summary.RoundOffAmount = Decimal.Round(roundedNet - exactNet, 2, MidpointRounding.AwayFromZero)
            summary.NetAmount = Decimal.Round(roundedNet, 2, MidpointRounding.AwayFromZero)

            Dim groupedTax = selectedLines.
                GroupBy(Function(line) line.GstPercentage).
                OrderBy(Function(group) group.Key).
                Select(Function(group) $"{group.Key.ToString("N2", CultureInfo.InvariantCulture)}% : Taxable {group.Sum(Function(line) line.TaxableAmount):N2}, GST {group.Sum(Function(line) line.GstAmount):N2}")
            summary.TaxSummaryText = If(groupedTax.Any(), String.Join(Environment.NewLine, groupedTax), "No return lines selected.")

            Return summary
        End Function

        Public Async Function SavePurchaseReturnAsync(draft As PurchaseReturnDraft, createdByUserId As Integer) As Task(Of PurchaseReturnSaveResult)
            Return Await Task.Run(
                Function()
                    NormalizePurchaseReturnDraft(draft)

                    Dim validationMessage As String = ValidatePurchaseReturnDraft(draft)
                    If validationMessage <> String.Empty Then
                        Return PurchaseReturnSaveResult.Failure(validationMessage)
                    End If

                    draft.Summary = CalculatePurchaseReturnTotals(draft.Items)

                    Try
                        Dim returnId As Integer = _stockOperationRepository.SavePurchaseReturn(draft, createdByUserId)
                        AppLogger.Info($"Purchase return '{draft.ReturnNumber}' saved with Id {returnId}.")
                        Return PurchaseReturnSaveResult.Success(returnId, draft.ReturnNumber, $"Purchase return {draft.ReturnNumber} saved successfully.")
                    Catch ex As Exception
                        AppLogger.Error($"Purchase return save failed for '{draft.ReturnNumber}'.", ex)
                        Return PurchaseReturnSaveResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Public Async Function LoadProductsAsync() As Task(Of List(Of ProductRecord))
            Return Await Task.Run(Function() _productRepository.Search(String.Empty))
        End Function

        Public Async Function GenerateNextStockAdjustmentNumberAsync(adjustmentDate As DateTime) As Task(Of String)
            Return Await Task.Run(Function() _stockOperationRepository.GenerateNextStockAdjustmentNumber(adjustmentDate.Date))
        End Function

        Public Function CreateAdjustmentLineFromProduct(product As ProductRecord, adjustmentMode As StockAdjustmentMode) As StockAdjustmentLineItem
            Dim line As New StockAdjustmentLineItem With {
                .ProductId = product.Id,
                .ProductName = product.ProductName,
                .BatchNumber = product.BatchNumber,
                .CurrentStock = product.CurrentStock,
                .AdjustmentMode = adjustmentMode,
                .Quantity = 1,
                .UnitCost = product.PTR,
                .Remarks = String.Empty
            }
            RecalculateStockAdjustmentLine(line)
            Return line
        End Function

        Public Sub RecalculateStockAdjustmentLine(line As StockAdjustmentLineItem)
            line.Quantity = Math.Max(line.Quantity, 0)
            line.UnitCost = Math.Max(line.UnitCost, 0D)
            Dim delta As Integer = If(line.AdjustmentMode = StockAdjustmentMode.Increase, line.Quantity, -line.Quantity)
            line.ResultingStock = line.CurrentStock + delta
        End Sub

        Public Async Function SaveStockAdjustmentAsync(draft As StockAdjustmentDraft, createdByUserId As Integer) As Task(Of StockAdjustmentSaveResult)
            Return Await Task.Run(
                Function()
                    NormalizeStockAdjustmentDraft(draft)

                    Dim validationMessage As String = ValidateStockAdjustmentDraft(draft)
                    If validationMessage <> String.Empty Then
                        Return StockAdjustmentSaveResult.Failure(validationMessage)
                    End If

                    Try
                        Dim adjustmentId As Integer = _stockOperationRepository.SaveStockAdjustment(draft, createdByUserId)
                        AppLogger.Info($"Stock adjustment '{draft.AdjustmentNumber}' saved with Id {adjustmentId}.")
                        Return StockAdjustmentSaveResult.Success(adjustmentId, draft.AdjustmentNumber, $"Stock adjustment {draft.AdjustmentNumber} saved successfully.")
                    Catch ex As Exception
                        AppLogger.Error($"Stock adjustment save failed for '{draft.AdjustmentNumber}'.", ex)
                        Return StockAdjustmentSaveResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Private Sub NormalizePurchaseReturnDraft(draft As PurchaseReturnDraft)
            draft.ReturnNumber = If(draft.ReturnNumber, String.Empty).Trim().ToUpperInvariant()
            draft.PurchaseNumber = If(draft.PurchaseNumber, String.Empty).Trim().ToUpperInvariant()
            draft.SupplierName = If(draft.SupplierName, String.Empty).Trim()
            draft.Notes = If(draft.Notes, String.Empty).Trim()

            For Each line As PurchaseReturnLineItem In draft.Items
                line.ProductName = If(line.ProductName, String.Empty).Trim()
                line.BatchNumber = If(line.BatchNumber, String.Empty).Trim().ToUpperInvariant()
                RecalculatePurchaseReturnLine(line)
            Next
        End Sub

        Private Function ValidatePurchaseReturnDraft(draft As PurchaseReturnDraft) As String
            If Not InputValidator.IsRequiredTextProvided(draft.ReturnNumber) Then
                Return "Return number is required."
            End If

            If draft.PurchaseId <= 0 Then
                Return "Select a purchase to return against."
            End If

            If draft.SupplierId <= 0 Then
                Return "The selected purchase does not have a valid supplier."
            End If

            If draft.Items Is Nothing OrElse Not draft.Items.Any(Function(line) line.ReturnQuantity > 0 OrElse line.ReturnFreeQuantity > 0) Then
                Return "Enter at least one return quantity."
            End If

            For Each line As PurchaseReturnLineItem In draft.Items.Where(Function(item) item.ReturnQuantity > 0 OrElse item.ReturnFreeQuantity > 0)
                If line.ReturnQuantity > line.RemainingQuantity Then
                    Return $"Return quantity exceeds remaining quantity for '{line.ProductName}'."
                End If

                If line.ReturnFreeQuantity > line.RemainingFreeQuantity Then
                    Return $"Return free quantity exceeds remaining free quantity for '{line.ProductName}'."
                End If

                If line.ReturnQuantity + line.ReturnFreeQuantity > line.CurrentStock Then
                    Return $"Current stock is insufficient to return '{line.ProductName}'."
                End If
            Next

            Return String.Empty
        End Function

        Private Sub NormalizeStockAdjustmentDraft(draft As StockAdjustmentDraft)
            draft.AdjustmentNumber = If(draft.AdjustmentNumber, String.Empty).Trim().ToUpperInvariant()
            draft.Notes = If(draft.Notes, String.Empty).Trim()

            For Each line As StockAdjustmentLineItem In draft.Items
                line.ProductName = If(line.ProductName, String.Empty).Trim()
                line.BatchNumber = If(line.BatchNumber, String.Empty).Trim().ToUpperInvariant()
                line.Remarks = If(line.Remarks, String.Empty).Trim()
                RecalculateStockAdjustmentLine(line)
            Next
        End Sub

        Private Function ValidateStockAdjustmentDraft(draft As StockAdjustmentDraft) As String
            If Not InputValidator.IsRequiredTextProvided(draft.AdjustmentNumber) Then
                Return "Adjustment number is required."
            End If

            If draft.Items Is Nothing OrElse draft.Items.Count = 0 Then
                Return "Add at least one adjustment line."
            End If

            For Each line As StockAdjustmentLineItem In draft.Items
                If line.ProductId <= 0 Then
                    Return "One adjustment line is missing a product."
                End If

                If line.Quantity <= 0 Then
                    Return $"Adjustment quantity must be greater than zero for '{line.ProductName}'."
                End If

                If line.AdjustmentMode = StockAdjustmentMode.Decrease AndAlso line.ResultingStock < 0 Then
                    Return $"Insufficient stock for '{line.ProductName}'."
                End If
            Next

            Return String.Empty
        End Function

        Private Function NormalizeSearchTerm(searchTerm As String) As String
            Return If(searchTerm, String.Empty).Trim()
        End Function

    End Class

End Namespace
