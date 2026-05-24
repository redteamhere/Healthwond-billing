Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities
Imports System.Globalization
Imports System.Linq

Namespace Services

    Public Class PurchaseService

        Private ReadOnly _purchaseRepository As IPurchaseRepository
        Private ReadOnly _supplierRepository As ISupplierRepository
        Private ReadOnly _productRepository As IProductRepository

        Public Sub New(purchaseRepository As IPurchaseRepository, supplierRepository As ISupplierRepository, productRepository As IProductRepository)
            _purchaseRepository = purchaseRepository
            _supplierRepository = supplierRepository
            _productRepository = productRepository
        End Sub

        Public Async Function LoadSuppliersAsync() As Task(Of List(Of SupplierRecord))
            Return Await Task.Run(Function() _supplierRepository.Search(String.Empty))
        End Function

        Public Async Function LoadProductsAsync() As Task(Of List(Of ProductRecord))
            Return Await Task.Run(Function() _productRepository.Search(String.Empty))
        End Function

        Public Async Function GenerateNextPurchaseNumberAsync(purchaseDate As DateTime) As Task(Of String)
            Return Await Task.Run(Function() _purchaseRepository.GenerateNextPurchaseNumber(purchaseDate))
        End Function

        Public Function CreateLineFromProduct(product As ProductRecord) As PurchaseLineItem
            Dim line As New PurchaseLineItem With {
                .ProductId = product.Id,
                .ProductName = product.ProductName,
                .Packing = product.Packing,
                .HsnCode = product.HsnCode,
                .BatchNumber = product.BatchNumber,
                .ExpiryDate = product.ExpiryDate,
                .CompanyName = product.CompanyName,
                .Composition = product.Composition,
                .Barcode = product.Barcode,
                .ExistingStock = product.CurrentStock,
                .Quantity = 1,
                .FreeQuantity = 0,
                .PTR = product.PTR,
                .PTS = product.PTS,
                .MRP = product.MRP,
                .GstPercentage = product.GstPercentage
            }
            RecalculateLine(line)
            Return line
        End Function

        Public Sub RecalculateLine(line As PurchaseLineItem)
            line.Quantity = Math.Max(line.Quantity, 0)
            line.FreeQuantity = Math.Max(line.FreeQuantity, 0)
            line.PTR = Math.Max(line.PTR, 0D)
            line.PTS = Math.Max(line.PTS, 0D)
            line.MRP = Math.Max(line.MRP, 0D)
            line.GstPercentage = Math.Max(0D, Math.Min(100D, line.GstPercentage))

            line.TaxableAmount = Decimal.Round(line.Quantity * line.PTR, 2, MidpointRounding.AwayFromZero)
            line.GstAmount = Decimal.Round(line.TaxableAmount * (line.GstPercentage / 100D), 2, MidpointRounding.AwayFromZero)
            line.LineTotal = Decimal.Round(line.TaxableAmount + line.GstAmount, 2, MidpointRounding.AwayFromZero)
        End Sub

        Public Function CalculateTotals(items As IEnumerable(Of PurchaseLineItem)) As PurchaseTotalsSummary
            Dim summary As New PurchaseTotalsSummary()
            Dim lines As List(Of PurchaseLineItem) = items.Where(Function(item) item IsNot Nothing AndAlso item.ProductId > 0).ToList()

            For Each item As PurchaseLineItem In lines
                RecalculateLine(item)
            Next

            summary.SubTotal = Decimal.Round(lines.Sum(Function(item) item.Quantity * item.PTR), 2, MidpointRounding.AwayFromZero)
            summary.DiscountAmount = 0D
            summary.GstAmount = Decimal.Round(lines.Sum(Function(item) item.GstAmount), 2, MidpointRounding.AwayFromZero)

            Dim exactNet As Decimal = lines.Sum(Function(item) item.LineTotal)
            Dim roundedNet As Decimal = Decimal.Round(exactNet, 0, MidpointRounding.AwayFromZero)
            summary.RoundOffAmount = Decimal.Round(roundedNet - exactNet, 2, MidpointRounding.AwayFromZero)
            summary.NetAmount = Decimal.Round(roundedNet, 2, MidpointRounding.AwayFromZero)

            Dim groupedTax = lines.
                GroupBy(Function(item) item.GstPercentage).
                OrderBy(Function(group) group.Key).
                Select(Function(group) $"{group.Key.ToString("N2", CultureInfo.InvariantCulture)}% : Taxable {group.Sum(Function(item) item.TaxableAmount):N2}, GST {group.Sum(Function(item) item.GstAmount):N2}")

            summary.TaxSummaryText = If(groupedTax.Any(), String.Join(Environment.NewLine, groupedTax), "No tax lines yet.")
            Return summary
        End Function

        Public Async Function SavePurchaseAsync(draft As PurchaseDraft, createdByUserId As Integer) As Task(Of PurchaseSaveResult)
            Return Await Task.Run(
                Function()
                    NormalizeDraft(draft)

                    Dim validationMessage As String = ValidateDraft(draft)
                    If validationMessage <> String.Empty Then
                        Return PurchaseSaveResult.Failure(validationMessage)
                    End If

                    draft.Summary = CalculateTotals(draft.Items)

                    Try
                        Dim purchaseId As Integer = _purchaseRepository.SavePurchase(draft, createdByUserId)
                        AppLogger.Info($"Purchase '{draft.PurchaseNumber}' saved with Id {purchaseId}.")
                        Return PurchaseSaveResult.Success(purchaseId, draft.PurchaseNumber, $"Purchase {draft.PurchaseNumber} saved successfully.")
                    Catch ex As Exception
                        AppLogger.Error($"Purchase save failed for '{draft.PurchaseNumber}'.", ex)
                        Return PurchaseSaveResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Private Sub NormalizeDraft(draft As PurchaseDraft)
            draft.PurchaseNumber = If(draft.PurchaseNumber, String.Empty).Trim().ToUpperInvariant()
            draft.SupplierInvoiceNumber = If(draft.SupplierInvoiceNumber, String.Empty).Trim().ToUpperInvariant()
            draft.Notes = If(draft.Notes, String.Empty).Trim()

            For Each item As PurchaseLineItem In draft.Items
                item.ProductName = If(item.ProductName, String.Empty).Trim()
                item.Packing = If(item.Packing, String.Empty).Trim()
                item.HsnCode = If(item.HsnCode, String.Empty).Trim().ToUpperInvariant()
                item.BatchNumber = If(item.BatchNumber, String.Empty).Trim().ToUpperInvariant()
                item.CompanyName = If(item.CompanyName, String.Empty).Trim()
                item.Composition = If(item.Composition, String.Empty).Trim()
                item.Barcode = If(item.Barcode, String.Empty).Trim()
                If item.ExpiryDate = DateTime.MinValue Then
                    item.ExpiryDate = DateTime.Today
                End If
                RecalculateLine(item)
            Next
        End Sub

        Private Function ValidateDraft(draft As PurchaseDraft) As String
            If Not InputValidator.IsRequiredTextProvided(draft.PurchaseNumber) Then
                Return "Purchase number is required."
            End If

            If draft.SupplierId <= 0 Then
                Return "Select a supplier."
            End If

            If draft.Items Is Nothing OrElse draft.Items.Count = 0 Then
                Return "Add at least one product line."
            End If

            For Each item As PurchaseLineItem In draft.Items
                If item.ProductId <= 0 Then
                    Return "One purchase line is missing a source product."
                End If

                If Not InputValidator.IsRequiredTextProvided(item.ProductName) Then
                    Return "Each purchase line must have a product name."
                End If

                If Not InputValidator.IsRequiredTextProvided(item.BatchNumber) Then
                    Return $"Batch number is required for '{item.ProductName}'."
                End If

                If item.Quantity <= 0 Then
                    Return $"Quantity must be greater than zero for '{item.ProductName}'."
                End If
            Next

            Return String.Empty
        End Function

    End Class

End Namespace
