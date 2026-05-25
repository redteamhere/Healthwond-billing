Imports ClosedXML.Excel
Imports HealthwondBilling.Models
Imports System.IO
Imports System.Globalization
Imports System.Windows.Forms

Namespace Utilities

    Public NotInheritable Class ReportExportHelper

        Private Sub New()
        End Sub

        Public Shared Function ExportGrid(reportName As String, grid As DataGridView) As String
            Dim filePath As String = Path.Combine(AppPaths.ReportsDirectory, $"{reportName}-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx")

            Using workbook As New XLWorkbook()
                Dim worksheet As IXLWorksheet = workbook.Worksheets.Add("Report")
                worksheet.Cell(1, 1).Value = reportName
                worksheet.Range(1, 1, 1, Math.Max(grid.Columns.Count, 1)).Merge()
                worksheet.Cell(1, 1).Style.Font.Bold = True
                worksheet.Cell(1, 1).Style.Font.FontSize = 16

                For columnIndex As Integer = 0 To grid.Columns.Count - 1
                    worksheet.Cell(3, columnIndex + 1).Value = grid.Columns(columnIndex).HeaderText
                    worksheet.Cell(3, columnIndex + 1).Style.Font.Bold = True
                    worksheet.Cell(3, columnIndex + 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#172554")
                    worksheet.Cell(3, columnIndex + 1).Style.Font.FontColor = XLColor.White
                Next

                Dim rowIndex As Integer = 4
                For Each row As DataGridViewRow In grid.Rows
                    If row.IsNewRow Then
                        Continue For
                    End If

                    For columnIndex As Integer = 0 To grid.Columns.Count - 1
                        worksheet.Cell(rowIndex, columnIndex + 1).Value = Convert.ToString(row.Cells(columnIndex).Value)
                    Next
                    rowIndex += 1
                Next

                worksheet.Columns().AdjustToContents()
                workbook.SaveAs(filePath)
            End Using

            Return filePath
        End Function

        Public Shared Function ExportProfitLoss(report As ProfitLossReport) As String
            Dim filePath As String = Path.Combine(AppPaths.ReportsDirectory, $"ProfitLoss-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx")

            Using workbook As New XLWorkbook()
                Dim worksheet As IXLWorksheet = workbook.Worksheets.Add("ProfitLoss")
                worksheet.Cell("A1").Value = "Profit / Loss Report"
                worksheet.Cell("A1").Style.Font.Bold = True
                worksheet.Cell("A1").Style.Font.FontSize = 16
                worksheet.Cell("A3").Value = "From Date"
                worksheet.Cell("B3").Value = report.FromDate.ToString("dd-MMM-yyyy")
                worksheet.Cell("A4").Value = "To Date"
                worksheet.Cell("B4").Value = report.ToDate.ToString("dd-MMM-yyyy")

                Dim metrics = New Dictionary(Of String, Decimal) From {
                    {"Sales Taxable Amount", report.SalesTaxableAmount},
                    {"Sales Net Amount", report.SalesNetAmount},
                    {"Purchase Taxable Amount", report.PurchaseTaxableAmount},
                    {"Purchase Net Amount", report.PurchaseNetAmount},
                    {"Estimated Cost of Goods Sold", report.EstimatedCostOfGoodsSold},
                    {"Estimated Gross Profit", report.EstimatedGrossProfit},
                    {"Gross Margin %", report.GrossMarginPercentage},
                    {"Outstanding Receivables", report.OutstandingReceivables},
                    {"Outstanding Payables", report.OutstandingPayables}
                }

                Dim rowIndex As Integer = 6
                For Each metric In metrics
                    worksheet.Cell(rowIndex, 1).Value = metric.Key
                    worksheet.Cell(rowIndex, 2).Value = metric.Value
                    rowIndex += 1
                Next

                worksheet.Column(2).Style.NumberFormat.Format = "#,##0.00"
                worksheet.Columns().AdjustToContents()
                workbook.SaveAs(filePath)
            End Using

            Return filePath
        End Function

        Public Shared Function ExportOverview(overview As ReportOverview) As String
            Dim filePath As String = Path.Combine(AppPaths.ReportsDirectory, $"Overview-{DateTime.Now:yyyyMMdd-HHmmss}.xlsx")

            Using workbook As New XLWorkbook()
                Dim worksheet As IXLWorksheet = workbook.Worksheets.Add("Overview")
                worksheet.Cell("A1").Value = "Business Overview"
                worksheet.Cell("A1").Style.Font.Bold = True
                worksheet.Cell("A1").Style.Font.FontSize = 16
                worksheet.Cell("A3").Value = "From Date"
                worksheet.Cell("B3").Value = overview.FromDate.ToString("dd-MMM-yyyy")
                worksheet.Cell("A4").Value = "To Date"
                worksheet.Cell("B4").Value = overview.ToDate.ToString("dd-MMM-yyyy")

                Dim metrics = New Dictionary(Of String, Object) From {
                    {"Sales Invoice Count", overview.SalesInvoiceCount},
                    {"Purchase Bill Count", overview.PurchaseBillCount},
                    {"Sales Units", overview.SalesUnits},
                    {"Purchase Units", overview.PurchaseUnits},
                    {"Average Sale Bill", overview.AverageSaleBillValue},
                    {"Average Purchase Bill", overview.AveragePurchaseBillValue},
                    {"Collections Received", overview.CustomerCollectionsAmount},
                    {"Supplier Payments", overview.SupplierPaymentsAmount},
                    {"Collection Efficiency %", overview.CollectionEfficiencyPercentage},
                    {"Supplier Payment Coverage %", overview.SupplierPaymentCoveragePercentage},
                    {"Inventory SKU Count", overview.InventorySkuCount},
                    {"Inventory Stock Value At PTR", overview.InventoryStockValueAtPTR},
                    {"Outstanding Receivables", overview.OutstandingReceivables},
                    {"Outstanding Payables", overview.OutstandingPayables},
                    {"Net Cash Movement", overview.NetCashMovement}
                }

                Dim rowIndex As Integer = 6
                For Each metric In metrics
                    worksheet.Cell(rowIndex, 1).Value = metric.Key
                    If TypeOf metric.Value Is Integer Then
                        worksheet.Cell(rowIndex, 2).Value = Convert.ToInt32(metric.Value, CultureInfo.InvariantCulture)
                    Else
                        worksheet.Cell(rowIndex, 2).Value = Convert.ToDecimal(metric.Value, CultureInfo.InvariantCulture)
                    End If
                    rowIndex += 1
                Next

                worksheet.Column(2).Style.NumberFormat.Format = "#,##0.00"
                worksheet.Columns().AdjustToContents()
                workbook.SaveAs(filePath)
            End Using

            Return filePath
        End Function

    End Class

End Namespace
