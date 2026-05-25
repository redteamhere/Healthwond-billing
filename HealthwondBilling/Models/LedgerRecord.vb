Namespace Models

    Public Class LedgerRecord
        Public Property Id As Integer
        Public Property LedgerName As String = String.Empty
        Public Property AccountGroupId As Integer
        Public Property AccountGroupName As String = String.Empty
        Public Property OpeningBalance As Decimal
        Public Property OpeningBalanceType As String = "Dr"
        Public Property IsSystem As Boolean
        Public Property IsPartyLedger As Boolean
        Public Property LinkedEntityType As String = String.Empty
        Public Property LinkedEntityId As Integer
        Public Property Notes As String = String.Empty
        Public Property CurrentBalance As Decimal
        Public Property CurrentBalanceType As String = "Dr"
    End Class

End Namespace
