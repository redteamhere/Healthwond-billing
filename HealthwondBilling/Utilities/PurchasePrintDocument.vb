Imports HealthwondBilling.Models
Imports System.Drawing.Printing
Imports System.Globalization
Imports System.Linq

Namespace Utilities

    Public Class PurchasePrintDocument
        Inherits PrintDocument

        Private ReadOnly _document As PurchaseDocument
        Private _currentItemIndex As Integer

        Public Sub New(document As PurchaseDocument)
            _document = document
            DefaultPageSettings.Landscape = True
            DocumentName = $"Purchase {_document.PurchaseNumber}"
        End Sub

        Protected Overrides Sub OnBeginPrint(e As PrintEventArgs)
            MyBase.OnBeginPrint(e)
            _currentItemIndex = 0
        End Sub

        Protected Overrides Sub OnPrintPage(e As PrintPageEventArgs)
            MyBase.OnPrintPage(e)

            Dim g As Graphics = e.Graphics
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.ClearTypeGridFit

            Dim margin As Rectangle = e.MarginBounds
            Dim left As Single = margin.Left
            Dim top As Single = margin.Top
            Dim width As Single = margin.Width
            Dim y As Single = top

            Using titleFont As New Font("Segoe UI Semibold", 24.0F, FontStyle.Bold),
                  bannerFont As New Font("Segoe UI Semibold", 11.0F, FontStyle.Bold),
                  sectionFont As New Font("Segoe UI Semibold", 10.0F, FontStyle.Bold),
                  bodyFont As New Font("Segoe UI", 8.9F, FontStyle.Regular),
                  smallFont As New Font("Segoe UI", 8.0F, FontStyle.Regular),
                  totalFont As New Font("Segoe UI Semibold", 11.0F, FontStyle.Bold),
                  linePen As New Pen(Color.Black, 1.0F),
                  headerBrush As New SolidBrush(Color.FromArgb(236, 236, 236))

                g.DrawRectangle(linePen, left, top, width, margin.Height)

                g.DrawString("GST PURCHASE INVOICE", bannerFont, Brushes.Black, New RectangleF(left, y, width, 18), New StringFormat With {.Alignment = StringAlignment.Center})
                g.DrawString("Original for Buyer", smallFont, Brushes.Black, New RectangleF(left, y, width - 6, 18), New StringFormat With {.Alignment = StringAlignment.Far})
                y += 18

                g.DrawString(_document.SupplierName.ToUpperInvariant(), titleFont, Brushes.DarkBlue, New RectangleF(left, y, width, 36), New StringFormat With {.Alignment = StringAlignment.Center})
                y += 34

                Dim supplierContact As String = JoinNonEmpty("  ", _document.SupplierAddress, FormatPhoneEmail(_document.SupplierPhone, _document.SupplierEmail))
                g.DrawString(supplierContact, bodyFont, Brushes.Black, New RectangleF(left + 16, y, width - 32, 16), New StringFormat With {.Alignment = StringAlignment.Center})
                y += 14
                g.DrawString($"D.L. No.: {_document.SupplierDrugLicenseNumber}    GSTIN: {_document.SupplierGstin}", bodyFont, Brushes.Black, New RectangleF(left + 16, y, width - 32, 16), New StringFormat With {.Alignment = StringAlignment.Center})
                y += 22

                Dim leftBoxWidth As Single = width * 0.36F
                Dim middleBoxWidth As Single = width * 0.37F
                Dim rightBoxWidth As Single = width - leftBoxWidth - middleBoxWidth
                Dim headerBoxHeight As Single = 134

                Dim vendorRect As New RectangleF(left, y, leftBoxWidth, headerBoxHeight)
                Dim receiverRect As New RectangleF(left + leftBoxWidth, y, middleBoxWidth, headerBoxHeight)
                Dim detailRect As New RectangleF(left + leftBoxWidth + middleBoxWidth, y, rightBoxWidth, headerBoxHeight)

                g.DrawRectangle(linePen, vendorRect.X, vendorRect.Y, vendorRect.Width, vendorRect.Height)
                g.DrawRectangle(linePen, receiverRect.X, receiverRect.Y, receiverRect.Width, receiverRect.Height)
                g.DrawRectangle(linePen, detailRect.X, detailRect.Y, detailRect.Width, detailRect.Height)

                DrawPartySection(g, vendorRect, sectionFont, bodyFont, "Details of Supplier / Vendor :", _document.SupplierName, _document.SupplierAddress, _document.SupplierDrugLicenseNumber, _document.SupplierPhone, _document.SupplierGstin)
                DrawPartySection(g, receiverRect, sectionFont, bodyFont, "Details of Receiver / Billed to :", _document.CompanyName, _document.CompanyAddress, _document.CompanyDrugLicenseNumber, _document.CompanyPhone, _document.CompanyGstin)
                DrawDocumentMeta(g, detailRect, sectionFont, bodyFont)

                y += headerBoxHeight

                Dim headers As String() = {"Sr", "Pack", "Qty", "Free", "Description of Goods", "HSN", "Batch", "Exp", "M.R.P", "Rate", "GST%", "GST Amt.", "Taxable Amt.", "Net Amount"}
                Dim baseWidths As Single() = {30, 62, 42, 44, 256, 74, 88, 52, 58, 58, 52, 74, 90, 94}
                Dim scale As Single = width / baseWidths.Sum()
                Dim widths As Single() = baseWidths.Select(Function(value) value * scale).ToArray()
                Dim x As Single = left
                Dim headerHeight As Single = 22

                For index As Integer = 0 To headers.Length - 1
                    Dim rect As New RectangleF(x, y, widths(index), headerHeight)
                    g.FillRectangle(headerBrush, rect)
                    g.DrawRectangle(linePen, rect.X, rect.Y, rect.Width, rect.Height)
                    g.DrawString(headers(index), smallFont, Brushes.Black, rect, New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                    x += widths(index)
                Next

                y += headerHeight
                Dim maxBottom As Single = margin.Bottom - 210

                While _currentItemIndex < _document.Items.Count
                    Dim item As PurchaseDocumentItem = _document.Items(_currentItemIndex)
                    Dim rowHeight As Single = 20
                    If y + rowHeight > maxBottom Then
                        e.HasMorePages = True
                        Return
                    End If

                    x = left
                    Dim values As String() = {
                        item.LineNumber.ToString(CultureInfo.InvariantCulture),
                        item.Packing,
                        item.Quantity.ToString(CultureInfo.InvariantCulture),
                        item.FreeQuantity.ToString(CultureInfo.InvariantCulture),
                        item.ProductName,
                        item.HsnCode,
                        item.BatchNumber,
                        item.ExpiryDate.ToString("MM/yy", CultureInfo.InvariantCulture),
                        item.MRP.ToString("N2", CultureInfo.InvariantCulture),
                        item.PTR.ToString("N2", CultureInfo.InvariantCulture),
                        item.GstPercentage.ToString("N2", CultureInfo.InvariantCulture),
                        item.GstAmount.ToString("N2", CultureInfo.InvariantCulture),
                        item.TaxableAmount.ToString("N2", CultureInfo.InvariantCulture),
                        item.LineTotal.ToString("N2", CultureInfo.InvariantCulture)
                    }

                    For index As Integer = 0 To values.Length - 1
                        Dim rect As New RectangleF(x, y, widths(index), rowHeight)
                        g.DrawRectangle(linePen, rect.X, rect.Y, rect.Width, rect.Height)
                        Dim format As New StringFormat With {.Alignment = If(index = 4, StringAlignment.Near, StringAlignment.Center), .LineAlignment = StringAlignment.Center}
                        Dim drawRect As RectangleF = If(index = 4, New RectangleF(rect.X + 3, rect.Y, rect.Width - 6, rect.Height), rect)
                        g.DrawString(values(index), smallFont, Brushes.Black, drawRect, format)
                        x += widths(index)
                    Next

                    y += rowHeight
                    _currentItemIndex += 1
                End While

                Dim taxSummaryWidth As Single = width * 0.38F
                Dim countWidth As Single = width * 0.22F
                Dim totalWidth As Single = width - taxSummaryWidth - countWidth
                Dim summaryTop As Single = y + 12
                Dim summaryHeight As Single = 92

                DrawTaxSummary(g, New RectangleF(left, summaryTop, taxSummaryWidth, summaryHeight), sectionFont, smallFont, linePen, headerBrush)
                DrawCountSummary(g, New RectangleF(left + taxSummaryWidth, summaryTop, countWidth, summaryHeight), sectionFont, bodyFont, linePen)
                DrawTotalsSummary(g, New RectangleF(left + taxSummaryWidth + countWidth, summaryTop, totalWidth, summaryHeight), sectionFont, bodyFont, totalFont, linePen)

                Dim footerTop As Single = summaryTop + summaryHeight + 8
                Dim notesRect As New RectangleF(left, footerTop, width * 0.62F, 74)
                Dim signRect As New RectangleF(left + width * 0.62F, footerTop, width * 0.38F, 74)
                g.DrawRectangle(linePen, notesRect.X, notesRect.Y, notesRect.Width, notesRect.Height)
                g.DrawRectangle(linePen, signRect.X, signRect.Y, signRect.Width, signRect.Height)

                g.DrawString("Terms & Notes", sectionFont, Brushes.Black, New RectangleF(notesRect.X + 6, notesRect.Y + 6, notesRect.Width - 12, 16))
                Dim noteText As String = If(String.IsNullOrWhiteSpace(_document.Notes), "Goods once received should be checked against the supplier invoice. Record kept for purchase audit and GST reconciliation.", _document.Notes)
                g.DrawString(noteText, bodyFont, Brushes.Black, New RectangleF(notesRect.X + 6, notesRect.Y + 24, notesRect.Width - 12, 42))

                g.DrawString("For Purchase Department", sectionFont, Brushes.Black, New RectangleF(signRect.X + 6, signRect.Y + 8, signRect.Width - 12, 16), New StringFormat With {.Alignment = StringAlignment.Center})
                g.DrawString("Authorized Signatory", bodyFont, Brushes.Black, New RectangleF(signRect.X + 6, signRect.Y + 44, signRect.Width - 12, 16), New StringFormat With {.Alignment = StringAlignment.Center})
            End Using

            e.HasMorePages = False
        End Sub

        Private Sub DrawPartySection(graphics As Graphics, area As RectangleF, sectionFont As Font, bodyFont As Font, headingText As String, name As String, address As String, drugLicenseNumber As String, phone As String, gstin As String)
            Dim y As Single = area.Y + 6
            graphics.DrawString(headingText, sectionFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 16))
            y += 18
            graphics.DrawString(name, sectionFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 16))
            y += 18
            graphics.DrawString(address, bodyFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 30))
            y += 34
            graphics.DrawString($"D.L. No.: {drugLicenseNumber}", bodyFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 14))
            y += 14
            graphics.DrawString($"PH. No.: {phone}", bodyFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 14))
            y += 14
            graphics.DrawString($"GSTIN: {gstin}", bodyFont, Brushes.Black, New RectangleF(area.X + 6, y, area.Width - 12, 14))
        End Sub

        Private Sub DrawDocumentMeta(graphics As Graphics, area As RectangleF, sectionFont As Font, bodyFont As Font)
            Dim lines As New List(Of Tuple(Of String, String)) From {
                Tuple.Create("Purchase No.", _document.PurchaseNumber),
                Tuple.Create("Date", _document.PurchaseDate.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture)),
                Tuple.Create("Supplier Invoice", _document.SupplierInvoiceNumber),
                Tuple.Create("Invoice Date", FormatOptionalDate(_document.SupplierInvoiceDate)),
                Tuple.Create("P.O. No.", _document.PurchaseOrderNumber),
                Tuple.Create("P.O. Date", FormatOptionalDate(_document.PurchaseOrderDate)),
                Tuple.Create("Place of Supply", _document.PlaceOfSupply),
                Tuple.Create("Cases", If(_document.CaseCount <= 0, String.Empty, _document.CaseCount.ToString(CultureInfo.InvariantCulture))),
                Tuple.Create("Transport", _document.TransportName),
                Tuple.Create("Eway Bill No.", _document.EwayBillNumber)
            }

            Dim y As Single = area.Y + 8
            For Each entry In lines
                graphics.DrawString(entry.Item1, bodyFont, Brushes.Black, New RectangleF(area.X + 8, y, area.Width * 0.48F, 13))
                graphics.DrawString(":", bodyFont, Brushes.Black, New RectangleF(area.X + area.Width * 0.52F - 8, y, 10, 13))
                graphics.DrawString(entry.Item2, bodyFont, Brushes.Black, New RectangleF(area.X + area.Width * 0.56F, y, area.Width * 0.4F, 13))
                y += 12
            Next
        End Sub

        Private Sub DrawTaxSummary(graphics As Graphics, area As RectangleF, sectionFont As Font, smallFont As Font, linePen As Pen, headerBrush As Brush)
            graphics.DrawRectangle(linePen, area.X, area.Y, area.Width, area.Height)

            Dim headerHeight As Single = 20
            Dim headers As String() = {"GST Slab", "Taxable", "GST"}
            Dim widths As Single() = {area.Width * 0.28F, area.Width * 0.36F, area.Width * 0.36F}
            Dim x As Single = area.X
            For index As Integer = 0 To headers.Length - 1
                Dim rect As New RectangleF(x, area.Y, widths(index), headerHeight)
                graphics.FillRectangle(headerBrush, rect)
                graphics.DrawRectangle(linePen, rect.X, rect.Y, rect.Width, rect.Height)
                graphics.DrawString(headers(index), smallFont, Brushes.Black, rect, New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                x += widths(index)
            Next

            Dim y As Single = area.Y + headerHeight
            Dim visibleLines As List(Of PurchaseTaxSummaryLine) = If(_document.TaxLines.Count = 0, New List(Of PurchaseTaxSummaryLine) From {New PurchaseTaxSummaryLine()}, _document.TaxLines)
            For Each taxLine As PurchaseTaxSummaryLine In visibleLines
                x = area.X
                Dim values As String() = {
                    If(taxLine.GstPercentage <= 0D, "0.00%", taxLine.GstPercentage.ToString("N2", CultureInfo.InvariantCulture) & "%"),
                    taxLine.TaxableAmount.ToString("N2", CultureInfo.InvariantCulture),
                    taxLine.GstAmount.ToString("N2", CultureInfo.InvariantCulture)
                }

                For index As Integer = 0 To values.Length - 1
                    Dim rect As New RectangleF(x, y, widths(index), 18)
                    graphics.DrawRectangle(linePen, rect.X, rect.Y, rect.Width, rect.Height)
                    graphics.DrawString(values(index), smallFont, Brushes.Black, rect, New StringFormat With {.Alignment = StringAlignment.Center, .LineAlignment = StringAlignment.Center})
                    x += widths(index)
                Next

                y += 18
                If y + 18 > area.Bottom Then
                    Exit For
                End If
            Next

            graphics.DrawString("Tax Summary", sectionFont, Brushes.Black, New RectangleF(area.X + 6, area.Bottom - 18, area.Width - 12, 12))
        End Sub

        Private Sub DrawCountSummary(graphics As Graphics, area As RectangleF, sectionFont As Font, bodyFont As Font, linePen As Pen)
            graphics.DrawRectangle(linePen, area.X, area.Y, area.Width, area.Height)
            graphics.DrawString("Purchase Summary", sectionFont, Brushes.Black, New RectangleF(area.X + 8, area.Y + 8, area.Width - 16, 16))
            graphics.DrawString($"Total Items : {_document.TotalLines}", bodyFont, Brushes.Black, New RectangleF(area.X + 8, area.Y + 32, area.Width - 16, 16))
            graphics.DrawString($"Total Units : {_document.TotalUnits}", bodyFont, Brushes.Black, New RectangleF(area.X + 8, area.Y + 50, area.Width - 16, 16))
            graphics.DrawString($"Supplier Invoice : {_document.SupplierInvoiceNumber}", bodyFont, Brushes.Black, New RectangleF(area.X + 8, area.Y + 68, area.Width - 16, 16))
        End Sub

        Private Sub DrawTotalsSummary(graphics As Graphics, area As RectangleF, sectionFont As Font, bodyFont As Font, totalFont As Font, linePen As Pen)
            graphics.DrawRectangle(linePen, area.X, area.Y, area.Width, area.Height)

            Dim y As Single = area.Y + 8
            DrawSummaryLine(graphics, sectionFont, bodyFont, area, y, "Taxable Total", _document.SubTotal.ToString("N2", CultureInfo.InvariantCulture))
            y += 18
            DrawSummaryLine(graphics, sectionFont, bodyFont, area, y, "Discount", _document.DiscountAmount.ToString("N2", CultureInfo.InvariantCulture))
            y += 18
            DrawSummaryLine(graphics, sectionFont, bodyFont, area, y, "GST", _document.GstAmount.ToString("N2", CultureInfo.InvariantCulture))
            y += 18
            DrawSummaryLine(graphics, sectionFont, bodyFont, area, y, "Round Off", _document.RoundOffAmount.ToString("N2", CultureInfo.InvariantCulture))
            y += 22
            graphics.DrawString("Grand Total", totalFont, Brushes.Black, New RectangleF(area.X + 8, y, area.Width * 0.48F, 18))
            graphics.DrawString(_document.NetAmount.ToString("N2", CultureInfo.InvariantCulture), totalFont, Brushes.Black, New RectangleF(area.X + area.Width * 0.5F, y, area.Width * 0.42F, 18), New StringFormat With {.Alignment = StringAlignment.Far})
        End Sub

        Private Sub DrawSummaryLine(graphics As Graphics, sectionFont As Font, bodyFont As Font, area As RectangleF, y As Single, labelText As String, valueText As String)
            graphics.DrawString(labelText, bodyFont, Brushes.Black, New RectangleF(area.X + 8, y, area.Width * 0.52F, 16))
            graphics.DrawString(valueText, bodyFont, Brushes.Black, New RectangleF(area.X + area.Width * 0.54F, y, area.Width * 0.36F, 16), New StringFormat With {.Alignment = StringAlignment.Far})
        End Sub

        Private Function JoinNonEmpty(separator As String, ParamArray values() As String) As String
            Return String.Join(separator, values.Where(Function(value) Not String.IsNullOrWhiteSpace(value)))
        End Function

        Private Function FormatPhoneEmail(phone As String, emailAddress As String) As String
            If String.IsNullOrWhiteSpace(phone) Then
                Return If(emailAddress, String.Empty)
            End If

            If String.IsNullOrWhiteSpace(emailAddress) Then
                Return $"Phone: {phone}"
            End If

            Return $"Phone: {phone}  Email: {emailAddress}"
        End Function

        Private Function FormatOptionalDate(dateValue As DateTime?) As String
            If Not dateValue.HasValue Then
                Return String.Empty
            End If

            Return dateValue.Value.ToString("dd-MM-yyyy", CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace
