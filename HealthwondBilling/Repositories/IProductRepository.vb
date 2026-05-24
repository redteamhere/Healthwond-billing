Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface IProductRepository
        Function Search(searchTerm As String) As List(Of ProductRecord)
        Function Save(product As ProductRecord) As Integer
        Function SoftDelete(productId As Integer) As Boolean
    End Interface

End Namespace
