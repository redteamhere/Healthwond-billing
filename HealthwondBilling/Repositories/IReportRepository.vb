Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IReportRepository
        Function GetSalesReport(fromDate As DateTime, toDate As DateTime) As List(Of SalesReportRow)
        Function GetPurchaseReport(fromDate As DateTime, toDate As DateTime) As List(Of PurchaseReportRow)
        Function GetGstReport(fromDate As DateTime, toDate As DateTime) As List(Of GstReportRow)
        Function GetStockReport() As List(Of StockReportRow)
        Function GetCustomerOutstandingReport() As List(Of CustomerOutstandingReportRow)
        Function GetCustomerAgingReport(asOfDate As DateTime) As List(Of AgingReportRow)
        Function GetSupplierAgingReport(asOfDate As DateTime) As List(Of AgingReportRow)
        Function GetReportOverview(fromDate As DateTime, toDate As DateTime) As ReportOverview
        Function GetProfitLossReport(fromDate As DateTime, toDate As DateTime) As ProfitLossReport
    End Interface

End Namespace
