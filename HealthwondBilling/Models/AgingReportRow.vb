Namespace Models

    Public Class AgingReportRow
        Public Property PartyName As String = String.Empty
        Public Property Gstin As String = String.Empty
        Public Property DrugLicenseNumber As String = String.Empty
        Public Property Phone As String = String.Empty
        Public Property OpenDocumentCount As Integer
        Public Property OldestOpenDate As DateTime
        Public Property AgeInDays As Integer
        Public Property Days0To30Amount As Decimal
        Public Property Days31To60Amount As Decimal
        Public Property Days61To90Amount As Decimal
        Public Property DaysAbove90Amount As Decimal
        Public Property UnallocatedAmount As Decimal
        Public Property OutstandingBalance As Decimal
    End Class

End Namespace
