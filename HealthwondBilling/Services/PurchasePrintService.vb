Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities
Imports System.Windows.Forms

Namespace Services

    Public Class PurchasePrintService

        Private ReadOnly _purchaseRepository As IPurchaseRepository

        Public Sub New(purchaseRepository As IPurchaseRepository)
            _purchaseRepository = purchaseRepository
        End Sub

        Public Sub ShowPrintPreview(purchaseId As Integer)
            Dim document As PurchaseDocument = _purchaseRepository.GetPurchaseDocument(purchaseId)

            Using printDocument As New PurchasePrintDocument(document)
                Using previewDialog As New PrintPreviewDialog()
                    previewDialog.Document = printDocument
                    previewDialog.Width = 1280
                    previewDialog.Height = 900
                    previewDialog.ShowDialog()
                End Using
            End Using
        End Sub

        Public Sub PrintPurchase(purchaseId As Integer)
            Dim document As PurchaseDocument = _purchaseRepository.GetPurchaseDocument(purchaseId)

            Using printDocument As New PurchasePrintDocument(document)
                printDocument.Print()
            End Using
        End Sub

    End Class

End Namespace
