Imports ClosedXML.Excel
Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities
Imports PdfSharp.Drawing
Imports PdfSharp.Pdf
Imports System.Diagnostics
Imports System.IO
Imports System.Linq
Imports System.Windows.Forms

Namespace Services

    Public Class InvoiceExportService

        Private ReadOnly _invoiceRepository As IInvoiceRepository
        Private ReadOnly _settingsRepository As ISettingsRepository

        Public Sub New(invoiceRepository As IInvoiceRepository, settingsRepository As ISettingsRepository)
            _invoiceRepository = invoiceRepository
            _settingsRepository = settingsRepository
        End Sub

        Public Async Function GenerateInvoiceFilesAsync(invoiceId As Integer) As Task(Of InvoiceExportResult)
            Return Await Task.Run(
                Function()
                    Try
                        Dim settingsProfile As AppSettingsProfile = _settingsRepository.GetProfile()
                        Dim templatePath As String = AppPaths.ResolveConfiguredPath(settingsProfile.InvoiceTemplatePath, AppPaths.GstInvoiceTemplateFilePath)
                        InvoiceTemplateGenerator.EnsureTemplateExists(templatePath)
                        Dim document As InvoiceDocument = _invoiceRepository.GetInvoiceDocument(invoiceId)
                        Dim excelFilePath As String = GenerateExcelInvoice(document, templatePath)
                        Dim pdfFilePath As String = GeneratePdfInvoice(document)
                        Return InvoiceExportResult.Success($"Invoice {document.InvoiceNumber} exported successfully.", excelFilePath, pdfFilePath)
                    Catch ex As Exception
                        AppLogger.Error($"Invoice export failed for invoice Id {invoiceId}.", ex)
                        Return InvoiceExportResult.Failure("Invoice files could not be generated.")
                    End Try
                End Function)
        End Function

        Public Sub ShowPrintPreview(invoiceId As Integer)
            Dim document As InvoiceDocument = _invoiceRepository.GetInvoiceDocument(invoiceId)
            Using printDocument As New InvoicePrintDocument(document)
                Using previewDialog As New PrintPreviewDialog()
                    previewDialog.Document = printDocument
                    previewDialog.Width = 1200
                    previewDialog.Height = 800
                    previewDialog.ShowDialog()
                End Using
            End Using
        End Sub

        Public Sub PrintInvoice(invoiceId As Integer)
            Dim document As InvoiceDocument = _invoiceRepository.GetInvoiceDocument(invoiceId)
            Using printDocument As New InvoicePrintDocument(document)
                printDocument.Print()
            End Using
        End Sub

        Public Sub OpenInvoiceFolder()
            Process.Start(New ProcessStartInfo With {
                .FileName = AppPaths.GeneratedInvoicesDirectory,
                .UseShellExecute = True
            })
        End Sub

        Public Function GetExcelFilePath(invoiceNumber As String) As String
            Return Path.Combine(AppPaths.GeneratedInvoicesDirectory, $"{invoiceNumber}.xlsx")
        End Function

        Public Function GetPdfFilePath(invoiceNumber As String) As String
            Return Path.Combine(AppPaths.GeneratedInvoicesDirectory, $"{invoiceNumber}.pdf")
        End Function

        Public Sub OpenGeneratedInvoiceFile(invoiceNumber As String, openPdf As Boolean)
            Dim filePath As String = If(openPdf, GetPdfFilePath(invoiceNumber), GetExcelFilePath(invoiceNumber))

            If Not File.Exists(filePath) Then
                Throw New FileNotFoundException("The requested invoice file was not found.", filePath)
            End If

            Process.Start(New ProcessStartInfo With {
                .FileName = filePath,
                .UseShellExecute = True
            })
        End Sub

        Private Function GenerateExcelInvoice(document As InvoiceDocument, templatePath As String) As String
            Dim filePath As String = GetExcelFilePath(document.InvoiceNumber)

            Using workbook As New XLWorkbook(templatePath)
                Dim sheet As IXLWorksheet = workbook.Worksheet("Invoice")
                sheet.Cell("A1").Value = document.CompanyName
                sheet.Cell("A4").Value = "Seller"
                sheet.Cell("A5").Value = document.CompanyAddress
                sheet.Cell("A6").Value = $"Phone: {document.CompanyPhone}"
                sheet.Cell("A7").Value = $"GSTIN: {document.CompanyGstin}"
                sheet.Cell("A8").Value = $"Drug Lic.: {document.CompanyDrugLicenseNumber}"

                sheet.Cell("G4").Value = "Invoice Details"
                sheet.Cell("G5").Value = $"Invoice No: {document.InvoiceNumber}"
                sheet.Cell("G6").Value = $"Date: {document.InvoiceDate:dd-MMM-yyyy}"
                sheet.Cell("G7").Value = $"Payment: {document.PaymentMode}"
                sheet.Cell("G8").Value = $"Balance: {document.BalanceAmount:N2}"

                Dim customerLabelRow As Integer = 9
                sheet.Range(customerLabelRow, 1, customerLabelRow + 2, 12).Style.Border.OutsideBorder = XLBorderStyleValues.Thin
                sheet.Cell(customerLabelRow, 1).Value = "Bill To"
                sheet.Cell(customerLabelRow, 1).Style.Font.Bold = True
                sheet.Cell(customerLabelRow + 1, 1).Value = document.CustomerName
                sheet.Cell(customerLabelRow + 1, 2).Value = document.CustomerAddress
                sheet.Cell(customerLabelRow + 2, 1).Value = $"Phone: {document.CustomerPhone}"
                sheet.Cell(customerLabelRow + 2, 4).Value = $"GSTIN: {document.CustomerGstin}"
                sheet.Cell(customerLabelRow + 2, 8).Value = $"Drug Lic.: {document.CustomerDrugLicenseNumber}"

                Dim itemStartRow As Integer = 11
                Dim currentRow As Integer = itemStartRow

                For Each item As InvoiceDocumentItem In document.Items
                    sheet.Cell(currentRow, 1).Value = item.LineNumber
                    sheet.Cell(currentRow, 2).Value = item.ProductName
                    sheet.Cell(currentRow, 3).Value = item.BatchNumber
                    sheet.Cell(currentRow, 4).Value = item.ExpiryDate.ToString("MM/yy")
                    sheet.Cell(currentRow, 5).Value = item.Quantity
                    sheet.Cell(currentRow, 6).Value = item.FreeQuantity
                    sheet.Cell(currentRow, 7).Value = item.Rate
                    sheet.Cell(currentRow, 8).Value = item.DiscountPercentage
                    sheet.Cell(currentRow, 9).Value = item.GstPercentage
                    sheet.Cell(currentRow, 10).Value = item.TaxableAmount
                    sheet.Cell(currentRow, 11).Value = item.GstAmount
                    sheet.Cell(currentRow, 12).Value = item.LineTotal
                    sheet.Range(currentRow, 1, currentRow, 12).Style.Border.OutsideBorder = XLBorderStyleValues.Thin
                    sheet.Range(currentRow, 1, currentRow, 12).Style.Border.InsideBorder = XLBorderStyleValues.Thin
                    currentRow += 1
                Next

                currentRow += 1
                sheet.Cell(currentRow, 10).Value = "Subtotal"
                sheet.Cell(currentRow, 12).Value = document.SubTotal
                currentRow += 1
                sheet.Cell(currentRow, 10).Value = "Discount"
                sheet.Cell(currentRow, 12).Value = document.DiscountAmount
                currentRow += 1
                sheet.Cell(currentRow, 10).Value = "GST"
                sheet.Cell(currentRow, 12).Value = document.GstAmount
                currentRow += 1
                sheet.Cell(currentRow, 10).Value = "Round Off"
                sheet.Cell(currentRow, 12).Value = document.RoundOffAmount
                currentRow += 1
                sheet.Cell(currentRow, 10).Value = "Net Amount"
                sheet.Cell(currentRow, 12).Value = document.NetAmount
                sheet.Range(currentRow - 4, 10, currentRow, 12).Style.Border.OutsideBorder = XLBorderStyleValues.Thin

                currentRow += 2
                sheet.Cell(currentRow, 1).Value = $"Amount Paid: {document.AmountPaid:N2}"
                sheet.Cell(currentRow, 4).Value = $"Balance: {document.BalanceAmount:N2}"

                If Not String.IsNullOrWhiteSpace(document.Notes) Then
                    currentRow += 2
                    sheet.Range(currentRow, 1, currentRow, 12).Merge()
                    sheet.Cell(currentRow, 1).Value = $"Notes: {document.Notes}"
                End If

                sheet.Range(itemStartRow, 7, currentRow, 12).Style.NumberFormat.Format = "#,##0.00"
                workbook.SaveAs(filePath)
            End Using

            Return filePath
        End Function

        Private Function GeneratePdfInvoice(document As InvoiceDocument) As String
            Dim filePath As String = GetPdfFilePath(document.InvoiceNumber)
            Dim pdfDocument As New PdfDocument()
            pdfDocument.Info.Title = $"Invoice {document.InvoiceNumber}"

            Dim page As PdfPage = pdfDocument.AddPage()
            page.Orientation = PdfSharp.PageOrientation.Landscape
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)

            Dim titleFont As New XFont("Segoe UI", 18, XFontStyle.Bold)
            Dim headerFont As New XFont("Segoe UI", 10, XFontStyle.Bold)
            Dim bodyFont As New XFont("Segoe UI", 9, XFontStyle.Regular)

            Dim y As Double = 30
            gfx.DrawString(document.CompanyName, titleFont, XBrushes.Black, New XRect(20, y, page.Width.Point - 40, 24), XStringFormats.TopCenter)
            y += 24
            gfx.DrawString("GST TAX INVOICE", headerFont, XBrushes.Black, New XRect(20, y, page.Width.Point - 40, 18), XStringFormats.TopCenter)
            y += 26

            gfx.DrawRectangle(XPens.Black, 20, y, 430, 72)
            gfx.DrawRectangle(XPens.Black, 470, y, 300, 72)
            gfx.DrawString("Seller", headerFont, XBrushes.Black, New XRect(26, y + 6, 100, 14), XStringFormats.TopLeft)
            gfx.DrawString(document.CompanyAddress, bodyFont, XBrushes.Black, New XRect(26, y + 22, 410, 14), XStringFormats.TopLeft)
            gfx.DrawString($"Phone: {document.CompanyPhone}", bodyFont, XBrushes.Black, New XRect(26, y + 36, 410, 14), XStringFormats.TopLeft)
            gfx.DrawString($"GSTIN: {document.CompanyGstin}", bodyFont, XBrushes.Black, New XRect(26, y + 50, 410, 14), XStringFormats.TopLeft)
            gfx.DrawString($"Drug Lic.: {document.CompanyDrugLicenseNumber}", bodyFont, XBrushes.Black, New XRect(230, y + 50, 206, 14), XStringFormats.TopLeft)

            gfx.DrawString("Invoice Details", headerFont, XBrushes.Black, New XRect(476, y + 6, 120, 14), XStringFormats.TopLeft)
            gfx.DrawString($"Invoice No: {document.InvoiceNumber}", bodyFont, XBrushes.Black, New XRect(476, y + 22, 280, 14), XStringFormats.TopLeft)
            gfx.DrawString($"Date: {document.InvoiceDate:dd-MMM-yyyy}", bodyFont, XBrushes.Black, New XRect(476, y + 36, 280, 14), XStringFormats.TopLeft)
            gfx.DrawString($"Payment: {document.PaymentMode}", bodyFont, XBrushes.Black, New XRect(476, y + 50, 280, 14), XStringFormats.TopLeft)

            y += 84
            gfx.DrawRectangle(XPens.Black, 20, y, 750, 56)
            gfx.DrawString("Bill To", headerFont, XBrushes.Black, New XRect(26, y + 6, 100, 14), XStringFormats.TopLeft)
            gfx.DrawString(document.CustomerName, bodyFont, XBrushes.Black, New XRect(26, y + 22, 240, 14), XStringFormats.TopLeft)
            gfx.DrawString(document.CustomerAddress, bodyFont, XBrushes.Black, New XRect(180, y + 22, 360, 14), XStringFormats.TopLeft)
            gfx.DrawString($"Phone: {document.CustomerPhone}", bodyFont, XBrushes.Black, New XRect(26, y + 38, 140, 14), XStringFormats.TopLeft)
            gfx.DrawString($"GSTIN: {document.CustomerGstin}", bodyFont, XBrushes.Black, New XRect(180, y + 38, 200, 14), XStringFormats.TopLeft)
            gfx.DrawString($"Drug Lic.: {document.CustomerDrugLicenseNumber}", bodyFont, XBrushes.Black, New XRect(430, y + 38, 220, 14), XStringFormats.TopLeft)

            y += 72
            Dim headers As String() = {"#", "Product", "Batch", "Expiry", "Qty", "Free", "Rate", "Disc%", "GST%", "Taxable", "GST", "Total"}
            Dim widths As Double() = {24, 170, 56, 56, 34, 34, 54, 46, 46, 64, 58, 64}
            Dim x As Double = 20

            For index As Integer = 0 To headers.Length - 1
                gfx.DrawRectangle(XBrushes.DarkBlue, x, y, widths(index), 20)
                gfx.DrawString(headers(index), bodyFont, XBrushes.White, New XRect(x, y + 3, widths(index), 14), XStringFormats.TopCenter)
                x += widths(index)
            Next

            y += 20
            For Each item As InvoiceDocumentItem In document.Items
                x = 20
                Dim values As String() = {
                    item.LineNumber.ToString(),
                    item.ProductName,
                    item.BatchNumber,
                    item.ExpiryDate.ToString("MM/yy"),
                    item.Quantity.ToString(),
                    item.FreeQuantity.ToString(),
                    item.Rate.ToString("N2"),
                    item.DiscountPercentage.ToString("N2"),
                    item.GstPercentage.ToString("N2"),
                    item.TaxableAmount.ToString("N2"),
                    item.GstAmount.ToString("N2"),
                    item.LineTotal.ToString("N2")
                }

                For index As Integer = 0 To values.Length - 1
                    gfx.DrawRectangle(XPens.Black, x, y, widths(index), 18)
                    Dim format As XStringFormat = If(index = 1, XStringFormats.TopLeft, XStringFormats.TopCenter)
                    gfx.DrawString(values(index), bodyFont, XBrushes.Black, New XRect(x + 2, y + 2, widths(index) - 4, 14), format)
                    x += widths(index)
                Next
                y += 18
            Next

            y += 12
            gfx.DrawString($"Subtotal: {document.SubTotal:N2}", headerFont, XBrushes.Black, New XRect(560, y, 180, 14), XStringFormats.TopLeft)
            y += 14
            gfx.DrawString($"Discount: {document.DiscountAmount:N2}", bodyFont, XBrushes.Black, New XRect(560, y, 180, 14), XStringFormats.TopLeft)
            y += 14
            gfx.DrawString($"GST: {document.GstAmount:N2}", bodyFont, XBrushes.Black, New XRect(560, y, 180, 14), XStringFormats.TopLeft)
            y += 14
            gfx.DrawString($"Round Off: {document.RoundOffAmount:N2}", bodyFont, XBrushes.Black, New XRect(560, y, 180, 14), XStringFormats.TopLeft)
            y += 14
            gfx.DrawString($"Net Amount: {document.NetAmount:N2}", headerFont, XBrushes.Black, New XRect(560, y, 180, 14), XStringFormats.TopLeft)
            y += 14
            gfx.DrawString($"Amount Paid: {document.AmountPaid:N2}", bodyFont, XBrushes.Black, New XRect(560, y, 180, 14), XStringFormats.TopLeft)
            y += 14
            gfx.DrawString($"Balance: {document.BalanceAmount:N2}", headerFont, XBrushes.Black, New XRect(560, y, 180, 14), XStringFormats.TopLeft)

            If Not String.IsNullOrWhiteSpace(document.Notes) Then
                y += 24
                gfx.DrawString($"Notes: {document.Notes}", bodyFont, XBrushes.Black, New XRect(20, y, 500, 24), XStringFormats.TopLeft)
            End If

            pdfDocument.Save(filePath)
            Return filePath
        End Function

    End Class

End Namespace
