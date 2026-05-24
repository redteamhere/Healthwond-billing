Namespace Models

    Public Class InvoiceSaveResult
        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property InvoiceId As Integer
        Public Property InvoiceNumber As String = String.Empty

        Public Shared Function Success(invoiceId As Integer, invoiceNumber As String, message As String) As InvoiceSaveResult
            Return New InvoiceSaveResult With {
                .IsSuccess = True,
                .InvoiceId = invoiceId,
                .InvoiceNumber = invoiceNumber,
                .Message = message
            }
        End Function

        Public Shared Function Failure(message As String) As InvoiceSaveResult
            Return New InvoiceSaveResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function
    End Class

End Namespace
