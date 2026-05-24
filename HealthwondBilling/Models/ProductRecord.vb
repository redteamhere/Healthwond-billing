Namespace Models

    Public Class ProductRecord
        Public Property Id As Integer
        Public Property ProductName As String = String.Empty
        Public Property Packing As String = String.Empty
        Public Property HsnCode As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property ExpiryDate As DateTime
        Public Property GstPercentage As Decimal
        Public Property MRP As Decimal
        Public Property PTR As Decimal
        Public Property PTS As Decimal
        Public Property CompanyName As String = String.Empty
        Public Property Composition As String = String.Empty
        Public Property CurrentStock As Integer
        Public Property Barcode As String = String.Empty
        Public Property IsDeleted As Boolean
        Public Property CreatedAt As DateTime?
        Public Property UpdatedAt As DateTime?
    End Class

End Namespace
