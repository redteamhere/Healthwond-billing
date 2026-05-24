Imports ClosedXML.Excel
Imports System.IO

Namespace Utilities

    Public NotInheritable Class InvoiceTemplateGenerator

        Private Sub New()
        End Sub

        Public Shared Sub EnsureTemplateExists(Optional templateFilePath As String = Nothing)
            Dim targetPath As String = If(String.IsNullOrWhiteSpace(templateFilePath), AppPaths.GstInvoiceTemplateFilePath, templateFilePath)

            Directory.CreateDirectory(Path.GetDirectoryName(targetPath))

            If File.Exists(targetPath) Then
                Return
            End If

            Using workbook As New XLWorkbook()
                Dim sheet As IXLWorksheet = workbook.Worksheets.Add("Invoice")

                sheet.Column(1).Width = 6
                sheet.Column(2).Width = 28
                sheet.Column(3).Width = 12
                sheet.Column(4).Width = 12
                sheet.Column(5).Width = 8
                sheet.Column(6).Width = 8
                sheet.Column(7).Width = 10
                sheet.Column(8).Width = 10
                sheet.Column(9).Width = 10
                sheet.Column(10).Width = 12
                sheet.Column(11).Width = 12
                sheet.Column(12).Width = 12

                sheet.Range("A1:L1").Merge().Style.Font.SetBold().Font.SetFontSize(18)
                sheet.Range("A1:L1").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                sheet.Cell("A1").Value = "Healthwond Billing System"

                sheet.Range("A2:L2").Merge().Style.Font.SetBold().Font.SetFontSize(11)
                sheet.Range("A2:L2").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                sheet.Cell("A2").Value = "GST TAX INVOICE"

                sheet.Range("A4:F8").Style.Border.OutsideBorder = XLBorderStyleValues.Thin
                sheet.Range("G4:L8").Style.Border.OutsideBorder = XLBorderStyleValues.Thin

                sheet.Cell("A10").Value = "Sr"
                sheet.Cell("B10").Value = "Product"
                sheet.Cell("C10").Value = "Batch"
                sheet.Cell("D10").Value = "Expiry"
                sheet.Cell("E10").Value = "Qty"
                sheet.Cell("F10").Value = "Free"
                sheet.Cell("G10").Value = "Rate"
                sheet.Cell("H10").Value = "Disc %"
                sheet.Cell("I10").Value = "GST %"
                sheet.Cell("J10").Value = "Taxable"
                sheet.Cell("K10").Value = "GST"
                sheet.Cell("L10").Value = "Total"

                sheet.Range("A10:L10").Style.Font.SetBold()
                sheet.Range("A10:L10").Style.Fill.BackgroundColor = XLColor.FromHtml("#172554")
                sheet.Range("A10:L10").Style.Font.FontColor = XLColor.White
                sheet.Range("A10:L10").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)

                sheet.PageSetup.PageOrientation = XLPageOrientation.Landscape
                sheet.PageSetup.Margins.Top = 0.35
                sheet.PageSetup.Margins.Bottom = 0.35
                sheet.PageSetup.Margins.Left = 0.25
                sheet.PageSetup.Margins.Right = 0.25
                sheet.PageSetup.FitToPages(1, 0)

                workbook.SaveAs(targetPath)
            End Using
        End Sub

    End Class

End Namespace
