Namespace Models

    Public Class CustomerRecord
        Public Property Id As Integer
        Public Property CustomerName As String = String.Empty
        Public Property Gstin As String = String.Empty
        Public Property DrugLicenseNumber As String = String.Empty
        Public Property Address As String = String.Empty
        Public Property Phone As String = String.Empty
        Public Property Email As String = String.Empty
        Public Property OutstandingBalance As Decimal
        Public Property CreatedAt As DateTime?
        Public Property UpdatedAt As DateTime?
    End Class

End Namespace
