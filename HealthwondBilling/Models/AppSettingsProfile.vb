Namespace Models

    Public Class AppSettingsProfile
        Public Property CompanyName As String = String.Empty
        Public Property CompanyAddress As String = String.Empty
        Public Property CompanyPhone As String = String.Empty
        Public Property CompanyGstin As String = String.Empty
        Public Property CompanyDrugLicense As String = String.Empty
        Public Property InvoicePrefix As String = String.Empty
        Public Property PurchasePrefix As String = String.Empty
        Public Property ReceiptPrefix As String = String.Empty
        Public Property SupplierPaymentPrefix As String = String.Empty
        Public Property LowStockThreshold As Integer
        Public Property CurrencySymbol As String = String.Empty
        Public Property InvoiceTemplatePath As String = String.Empty
    End Class

End Namespace
