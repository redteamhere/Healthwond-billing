Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports System.Threading.Tasks

Namespace Services

    Public Class ReportService

        Private ReadOnly _reportRepository As IReportRepository

        Public Sub New(reportRepository As IReportRepository)
            _reportRepository = reportRepository
        End Sub

        Public Async Function GetSalesReportAsync(fromDate As DateTime, toDate As DateTime) As Task(Of List(Of SalesReportRow))
            Return Await Task.Run(Function() _reportRepository.GetSalesReport(fromDate, toDate))
        End Function

        Public Async Function GetPurchaseReportAsync(fromDate As DateTime, toDate As DateTime) As Task(Of List(Of PurchaseReportRow))
            Return Await Task.Run(Function() _reportRepository.GetPurchaseReport(fromDate, toDate))
        End Function

        Public Async Function GetGstReportAsync(fromDate As DateTime, toDate As DateTime) As Task(Of List(Of GstReportRow))
            Return Await Task.Run(Function() _reportRepository.GetGstReport(fromDate, toDate))
        End Function

        Public Async Function GetStockReportAsync() As Task(Of List(Of StockReportRow))
            Return Await Task.Run(Function() _reportRepository.GetStockReport())
        End Function

        Public Async Function GetCustomerOutstandingReportAsync() As Task(Of List(Of CustomerOutstandingReportRow))
            Return Await Task.Run(Function() _reportRepository.GetCustomerOutstandingReport())
        End Function

        Public Async Function GetProfitLossReportAsync(fromDate As DateTime, toDate As DateTime) As Task(Of ProfitLossReport)
            Return Await Task.Run(Function() _reportRepository.GetProfitLossReport(fromDate, toDate))
        End Function

    End Class

End Namespace
