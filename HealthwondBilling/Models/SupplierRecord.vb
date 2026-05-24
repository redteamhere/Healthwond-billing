Namespace Models

    Public Class SupplierRecord
        Public Property Id As Integer
        Public Property SupplierName As String = String.Empty
        Public Property Gstin As String = String.Empty
        Public Property DrugLicenseNumber As String = String.Empty
        Public Property Address As String = String.Empty
        Public Property Phone As String = String.Empty
        Public Property Email As String = String.Empty
        Public Property OutstandingBalance As Decimal
        Public Property CreatedAt As DateTime?
        Public Property UpdatedAt As DateTime?

        Public Overrides Function ToString() As String
            Return $"{SupplierName} | {Phone}"
        End Function
    End Class

End Namespace
