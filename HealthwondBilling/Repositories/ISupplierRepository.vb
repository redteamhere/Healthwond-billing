Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface ISupplierRepository
        Function Search(searchTerm As String) As List(Of SupplierRecord)
        Function Save(supplier As SupplierRecord) As Integer
        Function Delete(supplierId As Integer) As Boolean
    End Interface

End Namespace
