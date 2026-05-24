Imports HealthwondBilling.Models
Imports System.Drawing.Printing
Imports System.Linq

Namespace Utilities

    Public Class InvoicePrintDocument
        Inherits PrintDocument

        Private ReadOnly _document As InvoiceDocument
        Private _currentItemIndex As Integer

        Public Sub New(document As InvoiceDocument)
            _document = document
            DefaultPageSettings.Landscape = True
            DocumentName = $"Invoice {_document.InvoiceNumber}"
        End Sub

        Protected Overrides Sub OnBeginPrint(e As PrintEventArgs)
            MyBase.OnBeginPrint(e)
            _currentItemIndex = 0
        End Sub

        Protected Overrides Sub OnPrintPage(e As PrintPageEventArgs)
            MyBase.OnPrintPage(e)

            Dim g As Graphics = e.Graphics
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

            Dim left As Single = e.MarginBounds.Left
            Dim top As Single = e.MarginBounds.Top
            Dim width As Single = e.MarginBounds.Width
            Dim y As Single = top

            Using titleFont As New Font("Segoe UI Semibold", 18.0F, FontStyle.Bold),
                  headerFont As New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold),
                  bodyFont As New Font("Segoe UI", 9.0F, FontStyle.Regular),
                  smallFont As New Font("Segoe UI", 8.0F, FontStyle.Regular),
                  linePen As New Pen(Color.Black, 1.0F)

                g.DrawString(_document.CompanyName, titleFont, Brushes.Black, New RectangleF(left, y, width, 28))
                y += 28
                g.DrawString("GST TAX INVOICE", headerFont, Brushes.Black, New RectangleF(left, y, width, 18), New StringFormat With {.Alignment = StringAlignment.Center})
                y += 26

                Dim sellerRect As New RectangleF(left, y, width * 0.55F, 82)
                Dim invoiceRect As New RectangleF(left + width * 0.57F, y, width * 0.43F, 82)
                g.DrawRectangle(linePen, sellerRect.X, sellerRect.Y, sellerRect.Width, sellerRect.Height)
                g.DrawRectangle(linePen, invoiceRect.X, invoiceRect.Y, invoiceRect.Width, invoiceRect.Height)

                DrawMultiLine(g, "Seller", headerFont, bodyFont, sellerRect,
                    _document.CompanyAddress,
                    $"Phone: {_document.CompanyPhone}",
                    $"GSTIN: {_document.CompanyGstin}",
                    $"Drug Lic.: {_document.CompanyDrugLicenseNumber}")

                DrawKeyValueLines(g, invoiceRect, headerFont, bodyFont, New String() {
                    $"Invoice No: {_document.InvoiceNumber}",
                    $"Date: {_document.InvoiceDate:dd-MMM-yyyy}",
                    $"Payment: {_document.PaymentMode}",
                    $"Balance: {_document.BalanceAmount:N2}"
                })

                y += 92

                Dim customerRect As New RectangleF(left, y, width, 68)
                g.DrawRectangle(linePen, customerRect.X, customerRect.Y, customerRect.Width, customerRect.Height)
                DrawMultiLine(g, "Bill To", headerFont, bodyFont, customerRect,
                    _document.CustomerName,
                    _document.CustomerAddress,
                    $"Phone: {_document.CustomerPhone}",
                    $"GSTIN: {_document.CustomerGstin}    Drug Lic.: {_document.CustomerDrugLicenseNumber}")
                y += 78

                Dim columns As Single() = {32, 190, 78, 72, 42, 42, 62, 54, 54, 78, 68, 78}
                Dim x As Single = left
                Dim headers As String() = {"#", "Product", "Batch", "Expiry", "Qty", "Free", "Rate", "Disc%", "GST%", "Taxable", "GST", "Total"}

                For index As Integer = 0 To headers.Length - 1
                    Dim rect As New RectangleF(x, y, columns(index), 24)
                    g.FillRectangle(New SolidBrush(Color.FromArgb(23, 37, 84)), rect)
                    g.DrawRectangle(linePen, rect.X, rect.Y, rect.Width, rect.Height)
                    g.DrawString(headers(index), smallFont, Brushes.White, rect, New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                    x += columns(index)
                Next

                y += 24
                Dim maxBottom As Single = e.MarginBounds.Bottom - 120

                While _currentItemIndex < _document.Items.Count
                    Dim item As InvoiceDocumentItem = _document.Items(_currentItemIndex)
                    x = left
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

                    Dim rowHeight As Single = 22
                    If y + rowHeight > maxBottom Then
                        e.HasMorePages = True
                        Return
                    End If

                    For index As Integer = 0 To values.Length - 1
                        Dim rect As New RectangleF(x, y, columns(index), rowHeight)
                        g.DrawRectangle(linePen, rect.X, rect.Y, rect.Width, rect.Height)
                        Dim format As New StringFormat With {.LineAlignment = StringAlignment.Center}
                        If index = 1 Then
                            format.Alignment = StringAlignment.Near
                        Else
                            format.Alignment = StringAlignment.Center
                        End If
                        g.DrawString(values(index), smallFont, Brushes.Black, rect, format)
                        x += columns(index)
                    Next

                    y += rowHeight
                    _currentItemIndex += 1
                End While

                y += 16
                g.DrawString($"SubTotal: {_document.SubTotal:N2}", headerFont, Brushes.Black, left + width - 210, y)
                y += 18
                g.DrawString($"Discount: {_document.DiscountAmount:N2}", bodyFont, Brushes.Black, left + width - 210, y)
                y += 18
                g.DrawString($"GST: {_document.GstAmount:N2}", bodyFont, Brushes.Black, left + width - 210, y)
                y += 18
                g.DrawString($"Round Off: {_document.RoundOffAmount:N2}", bodyFont, Brushes.Black, left + width - 210, y)
                y += 18
                g.DrawString($"Net Amount: {_document.NetAmount:N2}", headerFont, Brushes.Black, left + width - 210, y)
                y += 18
                g.DrawString($"Amount Paid: {_document.AmountPaid:N2}", bodyFont, Brushes.Black, left + width - 210, y)
                y += 18
                g.DrawString($"Balance: {_document.BalanceAmount:N2}", headerFont, Brushes.Black, left + width - 210, y)

                If Not String.IsNullOrWhiteSpace(_document.Notes) Then
                    y += 28
                    g.DrawString($"Notes: {_document.Notes}", bodyFont, Brushes.Black, New RectangleF(left, y, width - 220, 40))
                End If
            End Using

            e.HasMorePages = False
        End Sub

        Private Sub DrawMultiLine(g As Graphics, headingText As String, headingFont As Font, bodyFont As Font, area As RectangleF, ParamArray lines() As String)
            Dim y As Single = area.Y + 6
            g.DrawString(headingText, headingFont, Brushes.Black, area.X + 6, y)
            y += 18

            For Each line As String In lines.Where(Function(value) Not String.IsNullOrWhiteSpace(value))
                g.DrawString(line, bodyFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 16))
                y += 14
            Next
        End Sub

        Private Sub DrawKeyValueLines(g As Graphics, area As RectangleF, headingFont As Font, bodyFont As Font, lines As IEnumerable(Of String))
            Dim y As Single = area.Y + 6
            g.DrawString("Invoice Details", headingFont, Brushes.Black, area.X + 6, y)
            y += 18

            For Each line As String In lines
                g.DrawString(line, bodyFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 16))
                y += 14
            Next
        End Sub

    End Class

End Namespace
