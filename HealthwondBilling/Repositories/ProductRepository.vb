Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class ProductRepository
        Implements IProductRepository

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function Search(searchTerm As String) As List(Of ProductRecord) Implements IProductRepository.Search
            Dim products As New List(Of ProductRecord)()
            Dim normalizedSearch As String = If(searchTerm, String.Empty).Trim()
            Dim searchLike As String = $"%{normalizedSearch}%"

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT Id, ProductName, Packing, HsnCode, BatchNumber, ExpiryDate, GstPercentage, MRP, PTR, PTS, CompanyName, Composition, CurrentStock, Barcode, IsDeleted, CreatedAt, UpdatedAt " &
                        "FROM Products " &
                        "WHERE IsDeleted = 0 " &
                        "AND (@Search = '' OR ProductName LIKE @SearchLike OR BatchNumber LIKE @SearchLike OR CompanyName LIKE @SearchLike OR Composition LIKE @SearchLike OR Barcode LIKE @SearchLike OR HsnCode LIKE @SearchLike) " &
                        "ORDER BY ProductName ASC, BatchNumber ASC;"
                    command.AddParameter("@Search", normalizedSearch)
                    command.AddParameter("@SearchLike", searchLike)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            products.Add(MapProduct(reader))
                        End While
                    End Using
                End Using
            End Using

            Return products
        End Function

        Public Function Save(product As ProductRecord) As Integer Implements IProductRepository.Save
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim productId As Integer
                    Dim previousStock As Integer = 0

                    If product.Id > 0 Then
                        previousStock = GetCurrentStock(connection, transaction, product.Id)
                        UpdateProduct(connection, transaction, product)
                        productId = product.Id
                    Else
                        productId = InsertProduct(connection, transaction, product)
                    End If

                    Dim stockDelta As Integer = product.CurrentStock - previousStock
                    If product.Id = 0 AndAlso product.CurrentStock > 0 Then
                        InsertLedgerEntry(connection, transaction, productId, product.BatchNumber, "OPENING", "ProductMaster", productId, product.CurrentStock, 0, product.CurrentStock, product.PTR, "Product created from master screen")
                    ElseIf product.Id > 0 AndAlso stockDelta <> 0 Then
                        InsertLedgerEntry(
                            connection,
                            transaction,
                            productId,
                            product.BatchNumber,
                            "ADJUSTMENT",
                            "ProductMaster",
                            productId,
                            Math.Max(stockDelta, 0),
                            Math.Abs(Math.Min(stockDelta, 0)),
                            product.CurrentStock,
                            product.PTR,
                            "Stock adjusted from product master")
                    End If

                    transaction.Commit()
                    Return productId
                End Using
            End Using
        End Function

        Public Function SoftDelete(productId As Integer) As Boolean Implements IProductRepository.SoftDelete
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText = "UPDATE Products SET IsDeleted = 1, UpdatedAt = @UpdatedAt WHERE Id = @Id AND IsDeleted = 0;"
                    command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    command.AddParameter("@Id", productId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Function InsertProduct(connection As DbConnection, transaction As DbTransaction, product As ProductRecord) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Products (ProductName, Packing, HsnCode, BatchNumber, ExpiryDate, GstPercentage, MRP, PTR, PTS, CompanyName, Composition, CurrentStock, Barcode, IsDeleted, CreatedAt, UpdatedAt) " &
                    "VALUES (@ProductName, @Packing, @HsnCode, @BatchNumber, @ExpiryDate, @GstPercentage, @MRP, @PTR, @PTS, @CompanyName, @Composition, @CurrentStock, @Barcode, 0, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                AddProductParameters(command, product)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub UpdateProduct(connection As DbConnection, transaction As DbTransaction, product As ProductRecord)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "UPDATE Products SET ProductName = @ProductName, Packing = @Packing, HsnCode = @HsnCode, BatchNumber = @BatchNumber, ExpiryDate = @ExpiryDate, GstPercentage = @GstPercentage, MRP = @MRP, PTR = @PTR, PTS = @PTS, CompanyName = @CompanyName, Composition = @Composition, CurrentStock = @CurrentStock, Barcode = @Barcode, UpdatedAt = @UpdatedAt " &
                    "WHERE Id = @Id AND IsDeleted = 0;"
                AddProductParameters(command, product)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", product.Id)

                If command.ExecuteNonQuery() = 0 Then
                    Throw New InvalidOperationException("The selected product could not be updated.")
                End If
            End Using
        End Sub

        Private Sub AddProductParameters(command As DbCommand, product As ProductRecord)
            command.AddParameter("@ProductName", product.ProductName)
            command.AddParameter("@Packing", product.Packing)
            command.AddParameter("@HsnCode", product.HsnCode)
            command.AddParameter("@BatchNumber", product.BatchNumber)
            command.AddParameter("@ExpiryDate", SqliteDateHelper.ToStorageDate(product.ExpiryDate))
            command.AddParameter("@GstPercentage", product.GstPercentage)
            command.AddParameter("@MRP", product.MRP)
            command.AddParameter("@PTR", product.PTR)
            command.AddParameter("@PTS", product.PTS)
            command.AddParameter("@CompanyName", product.CompanyName)
            command.AddParameter("@Composition", product.Composition)
            command.AddParameter("@CurrentStock", product.CurrentStock)
            command.AddParameter("@Barcode", product.Barcode)
        End Sub

        Private Function GetCurrentStock(connection As DbConnection, transaction As DbTransaction, productId As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT CurrentStock FROM Products WHERE Id = @Id AND IsDeleted = 0 LIMIT 1;"
                command.AddParameter("@Id", productId)

                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Throw New InvalidOperationException("The selected product was not found.")
                End If

                Return Convert.ToInt32(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub InsertLedgerEntry(connection As DbConnection, transaction As DbTransaction, productId As Integer, batchNumber As String, transactionType As String, referenceType As String, referenceId As Integer, quantityIn As Integer, quantityOut As Integer, balanceQuantity As Integer, unitCost As Decimal, remarks As String)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO StockLedger (ProductId, BatchNumber, TransactionType, ReferenceType, ReferenceId, QuantityIn, QuantityOut, BalanceQuantity, UnitCost, Remarks, TransactionDate, CreatedAt) " &
                    "VALUES (@ProductId, @BatchNumber, @TransactionType, @ReferenceType, @ReferenceId, @QuantityIn, @QuantityOut, @BalanceQuantity, @UnitCost, @Remarks, @TransactionDate, @CreatedAt);"
                command.AddParameter("@ProductId", productId)
                command.AddParameter("@BatchNumber", batchNumber)
                command.AddParameter("@TransactionType", transactionType)
                command.AddParameter("@ReferenceType", referenceType)
                command.AddParameter("@ReferenceId", referenceId)
                command.AddParameter("@QuantityIn", quantityIn)
                command.AddParameter("@QuantityOut", quantityOut)
                command.AddParameter("@BalanceQuantity", balanceQuantity)
                command.AddParameter("@UnitCost", unitCost)
                command.AddParameter("@Remarks", remarks)
                command.AddParameter("@TransactionDate", SqliteDateHelper.ToStorageDate(DateTime.Today))
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Function MapProduct(reader As DbDataReader) As ProductRecord
            Return New ProductRecord With {
                .Id = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                .ProductName = Convert.ToString(reader("ProductName"), CultureInfo.InvariantCulture),
                .Packing = ConvertNullableString(reader("Packing")),
                .HsnCode = ConvertNullableString(reader("HsnCode")),
                .BatchNumber = Convert.ToString(reader("BatchNumber"), CultureInfo.InvariantCulture),
                .ExpiryDate = ParseDate(reader("ExpiryDate")),
                .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                .MRP = Convert.ToDecimal(reader("MRP"), CultureInfo.InvariantCulture),
                .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                .PTS = Convert.ToDecimal(reader("PTS"), CultureInfo.InvariantCulture),
                .CompanyName = ConvertNullableString(reader("CompanyName")),
                .Composition = ConvertNullableString(reader("Composition")),
                .CurrentStock = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture),
                .Barcode = ConvertNullableString(reader("Barcode")),
                .IsDeleted = Convert.ToInt32(reader("IsDeleted"), CultureInfo.InvariantCulture) = 1,
                .CreatedAt = ParseNullableDateTime(reader("CreatedAt")),
                .UpdatedAt = ParseNullableDateTime(reader("UpdatedAt"))
            }
        End Function

        Private Function ParseDate(value As Object) As DateTime
            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return DateTime.Today
        End Function

        Private Function ParseNullableDateTime(value As Object) As DateTime?
            If value Is Nothing OrElse value Is DBNull.Value Then
                Return Nothing
            End If

            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return Nothing
        End Function

        Private Function ConvertNullableString(value As Object) As String
            If value Is Nothing OrElse value Is DBNull.Value Then
                Return String.Empty
            End If

            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace
