Imports HealthwondBilling.Models
Imports HealthwondBilling.Repositories
Imports HealthwondBilling.Utilities

Namespace Services

    Public Class ProductService

        Private ReadOnly _productRepository As IProductRepository

        Public Sub New(productRepository As IProductRepository)
            _productRepository = productRepository
        End Sub

        Public Async Function SearchAsync(searchTerm As String) As Task(Of List(Of ProductRecord))
            Return Await Task.Run(Function() _productRepository.Search(searchTerm))
        End Function

        Public Async Function SaveAsync(product As ProductRecord) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    NormalizeProduct(product)

                    Dim validationMessage As String = ValidateProduct(product)
                    If validationMessage <> String.Empty Then
                        Return EntityOperationResult.Failure(validationMessage)
                    End If

                    Try
                        Dim entityId As Integer = _productRepository.Save(product)
                        Dim successMessage As String = If(product.Id > 0, "Product updated successfully.", "Product created successfully.")
                        AppLogger.Info($"Product '{product.ProductName}' saved with Id {entityId}.")
                        Return EntityOperationResult.Success(successMessage, entityId)
                    Catch ex As Exception
                        AppLogger.Error($"Product save failed for '{product.ProductName}'.", ex)
                        Return EntityOperationResult.Failure("The product could not be saved.")
                    End Try
                End Function)
        End Function

        Public Async Function DeleteAsync(product As ProductRecord) As Task(Of EntityOperationResult)
            Return Await Task.Run(
                Function()
                    If product Is Nothing OrElse product.Id <= 0 Then
                        Return EntityOperationResult.Failure("Select a product to delete.")
                    End If

                    If product.CurrentStock > 0 Then
                        Return EntityOperationResult.Failure("Set the current stock to zero before deleting the product.")
                    End If

                    Try
                        If _productRepository.SoftDelete(product.Id) Then
                            AppLogger.Info($"Product '{product.ProductName}' was marked deleted.")
                            Return EntityOperationResult.Success("Product deleted successfully.", product.Id)
                        End If

                        Return EntityOperationResult.Failure("The product could not be deleted.")
                    Catch ex As Exception
                        AppLogger.Error($"Product delete failed for '{product.ProductName}'.", ex)
                        Return EntityOperationResult.Failure("The product could not be deleted.")
                    End Try
                End Function)
        End Function

        Private Sub NormalizeProduct(product As ProductRecord)
            product.ProductName = If(product.ProductName, String.Empty).Trim()
            product.Packing = If(product.Packing, String.Empty).Trim()
            product.HsnCode = If(product.HsnCode, String.Empty).Trim().ToUpperInvariant()
            product.BatchNumber = If(product.BatchNumber, String.Empty).Trim().ToUpperInvariant()
            product.CompanyName = If(product.CompanyName, String.Empty).Trim()
            product.Composition = If(product.Composition, String.Empty).Trim()
            product.Barcode = If(product.Barcode, String.Empty).Trim()

            If product.ExpiryDate = DateTime.MinValue Then
                product.ExpiryDate = DateTime.Today
            End If
        End Sub

        Private Function ValidateProduct(product As ProductRecord) As String
            If Not InputValidator.IsRequiredTextProvided(product.ProductName) Then
                Return "Product name is required."
            End If

            If Not InputValidator.IsRequiredTextProvided(product.BatchNumber) Then
                Return "Batch number is required."
            End If

            If product.GstPercentage < 0D OrElse product.GstPercentage > 100D Then
                Return "GST percentage must be between 0 and 100."
            End If

            If product.MRP < 0D OrElse product.PTR < 0D OrElse product.PTS < 0D Then
                Return "Price fields cannot be negative."
            End If

            If product.CurrentStock < 0 Then
                Return "Current stock cannot be negative."
            End If

            Return String.Empty
        End Function

    End Class

End Namespace
