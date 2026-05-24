Imports HealthwondBilling.Models

Namespace Repositories

    Public Interface ICustomerRepository
        Function Search(searchTerm As String) As List(Of CustomerRecord)
        Function Save(customer As CustomerRecord) As Integer
        Function Delete(customerId As Integer) As Boolean
    End Interface

End Namespace
