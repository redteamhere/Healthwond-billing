Namespace Models

    Public Class InventoryLedgerRow
        Public Property TransactionDate As DateTime
        Public Property ProductName As String = String.Empty
        Public Property BatchNumber As String = String.Empty
        Public Property TransactionType As String = String.Empty
        Public Property ReferenceType As String = String.Empty
        Public Property ReferenceId As Integer
        Public Property QuantityIn As Integer
        Public Property QuantityOut As Integer
        Public Property BalanceQuantity As Integer
        Public Property UnitCost As Decimal
        Public Property Remarks As String = String.Empty
    End Class

End Namespace
