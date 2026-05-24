Namespace Models

    Public Class InvoiceExportResult

        Public Property IsSuccess As Boolean
        Public Property Message As String = String.Empty
        Public Property ExcelFilePath As String = String.Empty
        Public Property PdfFilePath As String = String.Empty

        Public Shared Function Success(message As String, excelFilePath As String, pdfFilePath As String) As InvoiceExportResult
            Return New InvoiceExportResult With {
                .IsSuccess = True,
                .Message = message,
                .ExcelFilePath = excelFilePath,
                .PdfFilePath = pdfFilePath
            }
        End Function

        Public Shared Function Failure(message As String) As InvoiceExportResult
            Return New InvoiceExportResult With {
                .IsSuccess = False,
                .Message = message
            }
        End Function

    End Class

End Namespace
