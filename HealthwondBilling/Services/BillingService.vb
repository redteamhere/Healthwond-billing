Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities
Imports System.Globalization
Imports System.Linq

Namespace Services

    Public Class BillingService

        Private ReadOnly _invoiceRepository As IInvoiceRepository
        Private ReadOnly _customerRepository As ICustomerRepository
        Private ReadOnly _productRepository As IProductRepository

        Public Sub New(invoiceRepository As IInvoiceRepository, customerRepository As ICustomerRepository, productRepository As IProductRepository)
            _invoiceRepository = invoiceRepository
            _customerRepository = customerRepository
            _productRepository = productRepository
        End Sub

        Public Async Function LoadCustomersAsync() As Task(Of List(Of CustomerRecord))
            Return Await Task.Run(Function() _customerRepository.Search(String.Empty))
        End Function

        Public Async Function LoadProductsAsync() As Task(Of List(Of ProductRecord))
            Return Await Task.Run(Function() _productRepository.Search(String.Empty))
        End Function

        Public Async Function GenerateNextInvoiceNumberAsync(invoiceDate As DateTime) As Task(Of String)
            Return Await Task.Run(Function() _invoiceRepository.GenerateNextInvoiceNumber(invoiceDate))
        End Function

        Public Async Function LoadInvoiceHistoryAsync(fromDate As DateTime, toDate As DateTime, searchTerm As String) As Task(Of List(Of InvoiceHistoryRow))
            Dim safeSearchTerm As String = If(searchTerm, String.Empty).Trim()
            Return Await Task.Run(Function() _invoiceRepository.SearchInvoices(fromDate.Date, toDate.Date, safeSearchTerm))
        End Function

        Public Async Function GetInvoiceDraftAsync(invoiceId As Integer) As Task(Of BillingInvoiceDraft)
            Return Await Task.Run(Function() _invoiceRepository.LoadInvoiceDraft(invoiceId))
        End Function

        Public Async Function GetInvoiceDocumentAsync(invoiceId As Integer) As Task(Of InvoiceDocument)
            Return Await Task.Run(Function() _invoiceRepository.GetInvoiceDocument(invoiceId))
        End Function

        Public Function CreateLineFromProduct(product As ProductRecord) As BillingLineItem
            Dim line As New BillingLineItem With {
                .ProductId = product.Id,
                .ProductName = product.ProductName,
                .BatchNumber = product.BatchNumber,
                .ExpiryDate = product.ExpiryDate,
                .Packing = product.Packing,
                .Barcode = product.Barcode,
                .AvailableStock = product.CurrentStock,
                .Quantity = 1,
                .FreeQuantity = 0,
                .Rate = product.PTS,
                .PTR = product.PTR,
                .MRP = product.MRP,
                .DiscountPercentage = 0D,
                .DiscountAmount = 0D,
                .SchemeDescription = String.Empty,
                .GstPercentage = product.GstPercentage
            }
            RecalculateLine(line)
            Return line
        End Function

        Public Sub RecalculateLine(line As BillingLineItem)
            line.Quantity = Math.Max(line.Quantity, 0)
            line.FreeQuantity = Math.Max(line.FreeQuantity, 0)
            line.Rate = Math.Max(line.Rate, 0D)
            line.PTR = Math.Max(line.PTR, 0D)
            line.MRP = Math.Max(line.MRP, 0D)
            line.DiscountPercentage = Math.Max(0D, Math.Min(100D, line.DiscountPercentage))
            line.GstPercentage = Math.Max(0D, Math.Min(100D, line.GstPercentage))

            Dim gross As Decimal = line.Quantity * line.Rate
            line.DiscountAmount = Decimal.Round(gross * (line.DiscountPercentage / 100D), 2, MidpointRounding.AwayFromZero)
            line.TaxableAmount = Decimal.Round(gross - line.DiscountAmount, 2, MidpointRounding.AwayFromZero)
            line.GstAmount = Decimal.Round(line.TaxableAmount * (line.GstPercentage / 100D), 2, MidpointRounding.AwayFromZero)
            line.LineTotal = Decimal.Round(line.TaxableAmount + line.GstAmount, 2, MidpointRounding.AwayFromZero)
        End Sub

        Public Function CalculateTotals(items As IEnumerable(Of BillingLineItem), amountPaid As Decimal) As BillingTotalsSummary
            Dim summary As New BillingTotalsSummary()
            Dim lines As List(Of BillingLineItem) = items.Where(Function(item) item IsNot Nothing AndAlso item.ProductId > 0).ToList()

            For Each item As BillingLineItem In lines
                RecalculateLine(item)
            Next

            summary.SubTotal = Decimal.Round(lines.Sum(Function(item) item.Quantity * item.Rate), 2, MidpointRounding.AwayFromZero)
            summary.DiscountAmount = Decimal.Round(lines.Sum(Function(item) item.DiscountAmount), 2, MidpointRounding.AwayFromZero)
            summary.SchemeAmount = Decimal.Round(lines.Sum(Function(item) item.FreeQuantity * item.Rate), 2, MidpointRounding.AwayFromZero)
            summary.GstAmount = Decimal.Round(lines.Sum(Function(item) item.GstAmount), 2, MidpointRounding.AwayFromZero)

            Dim exactNet As Decimal = lines.Sum(Function(item) item.LineTotal)
            Dim roundedNet As Decimal = Decimal.Round(exactNet, 0, MidpointRounding.AwayFromZero)
            summary.RoundOffAmount = Decimal.Round(roundedNet - exactNet, 2, MidpointRounding.AwayFromZero)
            summary.NetAmount = Decimal.Round(roundedNet, 2, MidpointRounding.AwayFromZero)
            summary.AmountPaid = Decimal.Round(Math.Max(amountPaid, 0D), 2, MidpointRounding.AwayFromZero)
            summary.BalanceAmount = Decimal.Round(Math.Max(summary.NetAmount - summary.AmountPaid, 0D), 2, MidpointRounding.AwayFromZero)

            Dim groupedTax = lines.
                GroupBy(Function(item) item.GstPercentage).
                OrderBy(Function(group) group.Key).
                Select(Function(group) $"{group.Key.ToString("N2", CultureInfo.InvariantCulture)}% : Taxable {group.Sum(Function(item) item.TaxableAmount):N2}, GST {group.Sum(Function(item) item.GstAmount):N2}")

            summary.TaxSummaryText = If(groupedTax.Any(), String.Join(Environment.NewLine, groupedTax), "No tax lines yet.")
            Return summary
        End Function

        Public Async Function SaveInvoiceAsync(draft As BillingInvoiceDraft, createdByUserId As Integer) As Task(Of InvoiceSaveResult)
            Return Await Task.Run(
                Function()
                    NormalizeDraft(draft)

                    Dim validationMessage As String = ValidateDraft(draft)
                    If validationMessage <> String.Empty Then
                        Return InvoiceSaveResult.Failure(validationMessage)
                    End If

                    draft.Summary = CalculateTotals(draft.Items, draft.AmountPaid)
                    If draft.Summary.AmountPaid > draft.Summary.NetAmount Then
                        Return InvoiceSaveResult.Failure("Amount paid cannot exceed the invoice net amount.")
                    End If

                    Try
                        Dim invoiceId As Integer
                        Dim successMessage As String

                        If draft.InvoiceId > 0 Then
                            invoiceId = _invoiceRepository.UpdateInvoice(draft, createdByUserId)
                            successMessage = $"Invoice {draft.InvoiceNumber} updated successfully."
                            AppLogger.Info($"Invoice '{draft.InvoiceNumber}' updated with Id {invoiceId}.")
                        Else
                            invoiceId = _invoiceRepository.SaveInvoice(draft, createdByUserId)
                            successMessage = $"Invoice {draft.InvoiceNumber} saved successfully."
                            AppLogger.Info($"Invoice '{draft.InvoiceNumber}' saved with Id {invoiceId}.")
                        End If

                        Return InvoiceSaveResult.Success(invoiceId, draft.InvoiceNumber, successMessage)
                    Catch ex As Exception
                        AppLogger.Error($"Invoice save failed for '{draft.InvoiceNumber}'.", ex)
                        Return InvoiceSaveResult.Failure(ex.Message)
                    End Try
                End Function)
        End Function

        Private Sub NormalizeDraft(draft As BillingInvoiceDraft)
            draft.InvoiceNumber = If(draft.InvoiceNumber, String.Empty).Trim().ToUpperInvariant()
            draft.PaymentMode = If(draft.PaymentMode, String.Empty).Trim()
            draft.Notes = If(draft.Notes, String.Empty).Trim()
            draft.AmountPaid = Math.Max(draft.AmountPaid, 0D)

            For Each item As BillingLineItem In draft.Items
                item.SchemeDescription = If(item.SchemeDescription, String.Empty).Trim()
                RecalculateLine(item)
            Next
        End Sub

        Private Function ValidateDraft(draft As BillingInvoiceDraft) As String
            If Not InputValidator.IsRequiredTextProvided(draft.InvoiceNumber) Then
                Return "Invoice number is required."
            End If

            If draft.CustomerId <= 0 Then
                Return "Select a customer."
            End If

            If draft.Items Is Nothing OrElse draft.Items.Count = 0 Then
                Return "Add at least one product."
            End If

            For Each item As BillingLineItem In draft.Items
                If item.ProductId <= 0 Then
                    Return "One invoice line is missing a product."
                End If

                If item.Quantity <= 0 Then
                    Return $"Quantity must be greater than zero for '{item.ProductName}'."
                End If

                If item.Quantity + item.FreeQuantity > item.AvailableStock Then
                    Return $"Insufficient stock for '{item.ProductName}'."
                End If
            Next

            Return String.Empty
        End Function

    End Class

End Namespace
