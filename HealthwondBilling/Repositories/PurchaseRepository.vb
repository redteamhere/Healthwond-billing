Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization
Imports System.Linq

Namespace Repositories

    Public Class PurchaseRepository
        Implements IPurchaseRepository

        Private NotInheritable Class InventoryProductSnapshot
            Public Property Id As Integer
            Public Property CurrentStock As Integer
        End Class

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GenerateNextPurchaseNumber(purchaseDate As DateTime) As String Implements IPurchaseRepository.GenerateNextPurchaseNumber
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim prefix As String = GetSetting(connection, "PurchasePrefix", "PUR")
                Dim monthPart As String = purchaseDate.ToString("yyyyMM", CultureInfo.InvariantCulture)
                Dim pattern As String = $"{prefix}-{monthPart}-%"
                Dim lastNumber As Integer = 0

                Using command = connection.CreateCommand()
                    command.CommandText = "SELECT PurchaseNumber FROM Purchases WHERE PurchaseNumber LIKE @Pattern ORDER BY Id DESC LIMIT 1;"
                    command.AddParameter("@Pattern", pattern)
                    Dim result As Object = command.ExecuteScalar()
                    If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                        Dim purchaseNumber As String = Convert.ToString(result, CultureInfo.InvariantCulture)
                        Dim parts As String() = purchaseNumber.Split("-"c)
                        Dim parsedNumber As Integer
                        If parts.Length >= 3 AndAlso Integer.TryParse(parts(parts.Length - 1), parsedNumber) Then
                            lastNumber = parsedNumber
                        End If
                    End If
                End Using

                Return $"{prefix}-{monthPart}-{(lastNumber + 1).ToString("0000", CultureInfo.InvariantCulture)}"
            End Using
        End Function

        Public Function SavePurchase(draft As PurchaseDraft, createdByUserId As Integer) As Integer Implements IPurchaseRepository.SavePurchase
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim purchaseId As Integer = InsertPurchase(connection, transaction, draft, createdByUserId)

                    For Each item As PurchaseLineItem In draft.Items
                        Dim productSnapshot As InventoryProductSnapshot = ResolveInventoryProduct(connection, transaction, item)
                        InsertPurchaseItem(connection, transaction, purchaseId, productSnapshot.Id, item)
                        InsertStockLedgerEntry(connection, transaction, purchaseId, draft.PurchaseDate, productSnapshot, item)
                    Next

                    UpdateSupplierOutstanding(connection, transaction, draft.SupplierId, draft.Summary.NetAmount)
                    AccountingPostingHelper.PostPurchaseVoucher(connection, transaction, purchaseId, draft, createdByUserId)
                    transaction.Commit()
                    Return purchaseId
                End Using
            End Using
        End Function

        Public Function GetPurchaseDocument(purchaseId As Integer) As PurchaseDocument Implements IPurchaseRepository.GetPurchaseDocument
            Using connection = _connectionFactory.CreateOpenConnection()
                Dim document As PurchaseDocument = LoadPurchaseHeader(connection, purchaseId)
                document.CompanyName = GetSetting(connection, "CompanyName", "Healthwond Pharmacy")
                document.CompanyAddress = GetSetting(connection, "CompanyAddress", "Address not configured")
                document.CompanyPhone = GetSetting(connection, "CompanyPhone", String.Empty)
                document.CompanyGstin = GetSetting(connection, "CompanyGstin", String.Empty)
                document.CompanyDrugLicenseNumber = GetSetting(connection, "CompanyDrugLicense", String.Empty)
                document.Items = LoadPurchaseItems(connection, purchaseId)
                document.TaxLines = LoadPurchaseTaxLines(connection, purchaseId)
                document.TotalLines = document.Items.Count
                document.TotalUnits = document.Items.Sum(Function(item) item.Quantity + item.FreeQuantity)
                Return document
            End Using
        End Function

        Private Function InsertPurchase(connection As DbConnection, transaction As DbTransaction, draft As PurchaseDraft, createdByUserId As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Purchases (PurchaseNumber, SupplierId, PurchaseDate, SupplierInvoiceNumber, SupplierInvoiceDate, PurchaseOrderNumber, PurchaseOrderDate, PlaceOfSupply, CaseCount, TransportName, EwayBillNumber, SubTotal, DiscountAmount, GstAmount, RoundOffAmount, NetAmount, Notes, CreatedBy, CreatedAt, UpdatedAt) " &
                    "VALUES (@PurchaseNumber, @SupplierId, @PurchaseDate, @SupplierInvoiceNumber, @SupplierInvoiceDate, @PurchaseOrderNumber, @PurchaseOrderDate, @PlaceOfSupply, @CaseCount, @TransportName, @EwayBillNumber, @SubTotal, @DiscountAmount, @GstAmount, @RoundOffAmount, @NetAmount, @Notes, @CreatedBy, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@PurchaseNumber", draft.PurchaseNumber)
                command.AddParameter("@SupplierId", draft.SupplierId)
                command.AddParameter("@PurchaseDate", SqliteDateHelper.ToStorageDate(draft.PurchaseDate))
                command.AddParameter("@SupplierInvoiceNumber", draft.SupplierInvoiceNumber)
                command.AddParameter("@SupplierInvoiceDate", ToNullableStorageDate(draft.SupplierInvoiceDate))
                command.AddParameter("@PurchaseOrderNumber", draft.PurchaseOrderNumber)
                command.AddParameter("@PurchaseOrderDate", ToNullableStorageDate(draft.PurchaseOrderDate))
                command.AddParameter("@PlaceOfSupply", draft.PlaceOfSupply)
                command.AddParameter("@CaseCount", draft.CaseCount)
                command.AddParameter("@TransportName", draft.TransportName)
                command.AddParameter("@EwayBillNumber", draft.EwayBillNumber)
                command.AddParameter("@SubTotal", draft.Summary.SubTotal)
                command.AddParameter("@DiscountAmount", draft.Summary.DiscountAmount)
                command.AddParameter("@GstAmount", draft.Summary.GstAmount)
                command.AddParameter("@RoundOffAmount", draft.Summary.RoundOffAmount)
                command.AddParameter("@NetAmount", draft.Summary.NetAmount)
                command.AddParameter("@Notes", draft.Notes)
                command.AddParameter("@CreatedBy", createdByUserId)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function ResolveInventoryProduct(connection As DbConnection, transaction As DbTransaction, item As PurchaseLineItem) As InventoryProductSnapshot
            Dim existingProduct As InventoryProductSnapshot = FindInventoryProduct(connection, transaction, item.ProductName, item.BatchNumber)
            Dim quantityIn As Integer = item.Quantity + item.FreeQuantity

            If existingProduct Is Nothing Then
                Dim newProductId As Integer = InsertInventoryProduct(connection, transaction, item, quantityIn)
                Return New InventoryProductSnapshot With {
                    .Id = newProductId,
                    .CurrentStock = quantityIn
                }
            End If

            Dim newStock As Integer = existingProduct.CurrentStock + quantityIn
            UpdateInventoryProduct(connection, transaction, existingProduct.Id, item, newStock)
            existingProduct.CurrentStock = newStock
            Return existingProduct
        End Function

        Private Function FindInventoryProduct(connection As DbConnection, transaction As DbTransaction, productName As String, batchNumber As String) As InventoryProductSnapshot
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "SELECT Id, CurrentStock FROM Products " &
                    "WHERE IsDeleted = 0 AND LOWER(ProductName) = LOWER(@ProductName) AND UPPER(BatchNumber) = @BatchNumber " &
                    "LIMIT 1;"
                command.AddParameter("@ProductName", productName)
                command.AddParameter("@BatchNumber", batchNumber.ToUpperInvariant())

                Using reader = command.ExecuteReader()
                    If reader.Read() Then
                        Return New InventoryProductSnapshot With {
                            .Id = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                            .CurrentStock = Convert.ToInt32(reader("CurrentStock"), CultureInfo.InvariantCulture)
                        }
                    End If
                End Using
            End Using

            Return Nothing
        End Function

        Private Function InsertInventoryProduct(connection As DbConnection, transaction As DbTransaction, item As PurchaseLineItem, initialStock As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Products (ProductName, Packing, HsnCode, BatchNumber, ExpiryDate, GstPercentage, MRP, PTR, PTS, CompanyName, Composition, CurrentStock, Barcode, IsDeleted, CreatedAt, UpdatedAt) " &
                    "VALUES (@ProductName, @Packing, @HsnCode, @BatchNumber, @ExpiryDate, @GstPercentage, @MRP, @PTR, @PTS, @CompanyName, @Composition, @CurrentStock, @Barcode, 0, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                AddInventoryProductParameters(command, item, initialStock)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub UpdateInventoryProduct(connection As DbConnection, transaction As DbTransaction, productId As Integer, item As PurchaseLineItem, currentStock As Integer)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "UPDATE Products SET ProductName = @ProductName, Packing = @Packing, HsnCode = @HsnCode, BatchNumber = @BatchNumber, ExpiryDate = @ExpiryDate, GstPercentage = @GstPercentage, MRP = @MRP, PTR = @PTR, PTS = @PTS, CompanyName = @CompanyName, Composition = @Composition, CurrentStock = @CurrentStock, Barcode = @Barcode, UpdatedAt = @UpdatedAt " &
                    "WHERE Id = @Id AND IsDeleted = 0;"
                AddInventoryProductParameters(command, item, currentStock)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", productId)

                If command.ExecuteNonQuery() = 0 Then
                    Throw New InvalidOperationException($"Inventory update failed for '{item.ProductName}'.")
                End If
            End Using
        End Sub

        Private Sub AddInventoryProductParameters(command As DbCommand, item As PurchaseLineItem, currentStock As Integer)
            command.AddParameter("@ProductName", item.ProductName)
            command.AddParameter("@Packing", item.Packing)
            command.AddParameter("@HsnCode", item.HsnCode)
            command.AddParameter("@BatchNumber", item.BatchNumber)
            command.AddParameter("@ExpiryDate", SqliteDateHelper.ToStorageDate(item.ExpiryDate))
            command.AddParameter("@GstPercentage", item.GstPercentage)
            command.AddParameter("@MRP", item.MRP)
            command.AddParameter("@PTR", item.PTR)
            command.AddParameter("@PTS", item.PTS)
            command.AddParameter("@CompanyName", item.CompanyName)
            command.AddParameter("@Composition", item.Composition)
            command.AddParameter("@CurrentStock", currentStock)
            command.AddParameter("@Barcode", item.Barcode)
        End Sub

        Private Sub InsertPurchaseItem(connection As DbConnection, transaction As DbTransaction, purchaseId As Integer, productId As Integer, item As PurchaseLineItem)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO PurchaseItems (PurchaseId, ProductId, ProductName, Packing, HsnCode, BatchNumber, ExpiryDate, Quantity, FreeQuantity, PTR, PTS, MRP, GstPercentage, TaxableAmount, GstAmount, LineTotal) " &
                    "VALUES (@PurchaseId, @ProductId, @ProductName, @Packing, @HsnCode, @BatchNumber, @ExpiryDate, @Quantity, @FreeQuantity, @PTR, @PTS, @MRP, @GstPercentage, @TaxableAmount, @GstAmount, @LineTotal);"
                command.AddParameter("@PurchaseId", purchaseId)
                command.AddParameter("@ProductId", productId)
                command.AddParameter("@ProductName", item.ProductName)
                command.AddParameter("@Packing", item.Packing)
                command.AddParameter("@HsnCode", item.HsnCode)
                command.AddParameter("@BatchNumber", item.BatchNumber)
                command.AddParameter("@ExpiryDate", SqliteDateHelper.ToStorageDate(item.ExpiryDate))
                command.AddParameter("@Quantity", item.Quantity)
                command.AddParameter("@FreeQuantity", item.FreeQuantity)
                command.AddParameter("@PTR", item.PTR)
                command.AddParameter("@PTS", item.PTS)
                command.AddParameter("@MRP", item.MRP)
                command.AddParameter("@GstPercentage", item.GstPercentage)
                command.AddParameter("@TaxableAmount", item.TaxableAmount)
                command.AddParameter("@GstAmount", item.GstAmount)
                command.AddParameter("@LineTotal", item.LineTotal)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Function LoadPurchaseHeader(connection As DbConnection, purchaseId As Integer) As PurchaseDocument
            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT p.Id, p.PurchaseNumber, p.PurchaseDate, COALESCE(p.SupplierInvoiceNumber, '') AS SupplierInvoiceNumber, p.SupplierInvoiceDate, " &
                    "COALESCE(p.PurchaseOrderNumber, '') AS PurchaseOrderNumber, p.PurchaseOrderDate, COALESCE(p.PlaceOfSupply, '') AS PlaceOfSupply, " &
                    "COALESCE(p.CaseCount, 0) AS CaseCount, COALESCE(p.TransportName, '') AS TransportName, COALESCE(p.EwayBillNumber, '') AS EwayBillNumber, " &
                    "COALESCE(p.Notes, '') AS Notes, p.SubTotal, p.DiscountAmount, p.GstAmount, p.RoundOffAmount, p.NetAmount, " &
                    "s.SupplierName, COALESCE(s.Address, '') AS Address, COALESCE(s.Phone, '') AS Phone, COALESCE(s.Email, '') AS Email, " &
                    "COALESCE(s.Gstin, '') AS Gstin, COALESCE(s.DrugLicenseNumber, '') AS DrugLicenseNumber " &
                    "FROM Purchases p " &
                    "INNER JOIN Suppliers s ON s.Id = p.SupplierId " &
                    "WHERE p.Id = @Id LIMIT 1;"
                command.AddParameter("@Id", purchaseId)

                Using reader = command.ExecuteReader()
                    If Not reader.Read() Then
                        Throw New InvalidOperationException($"Purchase Id {purchaseId} was not found.")
                    End If

                    Return New PurchaseDocument With {
                        .PurchaseId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                        .PurchaseNumber = ConvertNullableString(reader("PurchaseNumber")),
                        .PurchaseDate = ParseDate(reader("PurchaseDate")),
                        .SupplierInvoiceNumber = ConvertNullableString(reader("SupplierInvoiceNumber")),
                        .SupplierInvoiceDate = ParseNullableDate(reader("SupplierInvoiceDate")),
                        .PurchaseOrderNumber = ConvertNullableString(reader("PurchaseOrderNumber")),
                        .PurchaseOrderDate = ParseNullableDate(reader("PurchaseOrderDate")),
                        .PlaceOfSupply = ConvertNullableString(reader("PlaceOfSupply")),
                        .CaseCount = Convert.ToInt32(reader("CaseCount"), CultureInfo.InvariantCulture),
                        .TransportName = ConvertNullableString(reader("TransportName")),
                        .EwayBillNumber = ConvertNullableString(reader("EwayBillNumber")),
                        .Notes = ConvertNullableString(reader("Notes")),
                        .SubTotal = Convert.ToDecimal(reader("SubTotal"), CultureInfo.InvariantCulture),
                        .DiscountAmount = Convert.ToDecimal(reader("DiscountAmount"), CultureInfo.InvariantCulture),
                        .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                        .RoundOffAmount = Convert.ToDecimal(reader("RoundOffAmount"), CultureInfo.InvariantCulture),
                        .NetAmount = Convert.ToDecimal(reader("NetAmount"), CultureInfo.InvariantCulture),
                        .SupplierName = ConvertNullableString(reader("SupplierName")),
                        .SupplierAddress = ConvertNullableString(reader("Address")),
                        .SupplierPhone = ConvertNullableString(reader("Phone")),
                        .SupplierEmail = ConvertNullableString(reader("Email")),
                        .SupplierGstin = ConvertNullableString(reader("Gstin")),
                        .SupplierDrugLicenseNumber = ConvertNullableString(reader("DrugLicenseNumber"))
                    }
                End Using
            End Using
        End Function

        Private Function LoadPurchaseItems(connection As DbConnection, purchaseId As Integer) As List(Of PurchaseDocumentItem)
            Dim items As New List(Of PurchaseDocumentItem)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT COALESCE(pi.ProductName, prod.ProductName, '') AS ProductName, COALESCE(pi.Packing, prod.Packing, '') AS Packing, " &
                    "COALESCE(pi.HsnCode, prod.HsnCode, '') AS HsnCode, pi.BatchNumber, pi.ExpiryDate, pi.Quantity, pi.FreeQuantity, pi.MRP, pi.PTR, " &
                    "pi.GstPercentage, pi.TaxableAmount, pi.GstAmount, pi.LineTotal " &
                    "FROM PurchaseItems pi " &
                    "LEFT JOIN Products prod ON prod.Id = pi.ProductId " &
                    "WHERE pi.PurchaseId = @PurchaseId " &
                    "ORDER BY pi.Id ASC;"
                command.AddParameter("@PurchaseId", purchaseId)

                Using reader = command.ExecuteReader()
                    Dim lineNumber As Integer = 1
                    While reader.Read()
                        items.Add(New PurchaseDocumentItem With {
                            .LineNumber = lineNumber,
                            .ProductName = ConvertNullableString(reader("ProductName")),
                            .Packing = ConvertNullableString(reader("Packing")),
                            .HsnCode = ConvertNullableString(reader("HsnCode")),
                            .BatchNumber = ConvertNullableString(reader("BatchNumber")),
                            .ExpiryDate = ParseDate(reader("ExpiryDate")),
                            .Quantity = Convert.ToInt32(reader("Quantity"), CultureInfo.InvariantCulture),
                            .FreeQuantity = Convert.ToInt32(reader("FreeQuantity"), CultureInfo.InvariantCulture),
                            .MRP = Convert.ToDecimal(reader("MRP"), CultureInfo.InvariantCulture),
                            .PTR = Convert.ToDecimal(reader("PTR"), CultureInfo.InvariantCulture),
                            .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                            .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                            .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture),
                            .LineTotal = Convert.ToDecimal(reader("LineTotal"), CultureInfo.InvariantCulture)
                        })
                        lineNumber += 1
                    End While
                End Using
            End Using

            Return items
        End Function

        Private Function LoadPurchaseTaxLines(connection As DbConnection, purchaseId As Integer) As List(Of PurchaseTaxSummaryLine)
            Dim lines As New List(Of PurchaseTaxSummaryLine)()

            Using command = connection.CreateCommand()
                command.CommandText =
                    "SELECT pi.GstPercentage, COALESCE(SUM(pi.TaxableAmount), 0) AS TaxableAmount, COALESCE(SUM(pi.GstAmount), 0) AS GstAmount " &
                    "FROM PurchaseItems pi " &
                    "WHERE pi.PurchaseId = @PurchaseId " &
                    "GROUP BY pi.GstPercentage " &
                    "ORDER BY pi.GstPercentage ASC;"
                command.AddParameter("@PurchaseId", purchaseId)

                Using reader = command.ExecuteReader()
                    While reader.Read()
                        lines.Add(New PurchaseTaxSummaryLine With {
                            .GstPercentage = Convert.ToDecimal(reader("GstPercentage"), CultureInfo.InvariantCulture),
                            .TaxableAmount = Convert.ToDecimal(reader("TaxableAmount"), CultureInfo.InvariantCulture),
                            .GstAmount = Convert.ToDecimal(reader("GstAmount"), CultureInfo.InvariantCulture)
                        })
                    End While
                End Using
            End Using

            Return lines
        End Function

        Private Sub InsertStockLedgerEntry(connection As DbConnection, transaction As DbTransaction, purchaseId As Integer, purchaseDate As DateTime, productSnapshot As InventoryProductSnapshot, item As PurchaseLineItem)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO StockLedger (ProductId, BatchNumber, TransactionType, ReferenceType, ReferenceId, QuantityIn, QuantityOut, BalanceQuantity, UnitCost, Remarks, TransactionDate, CreatedAt) " &
                    "VALUES (@ProductId, @BatchNumber, @TransactionType, @ReferenceType, @ReferenceId, @QuantityIn, 0, @BalanceQuantity, @UnitCost, @Remarks, @TransactionDate, @CreatedAt);"
                command.AddParameter("@ProductId", productSnapshot.Id)
                command.AddParameter("@BatchNumber", item.BatchNumber)
                command.AddParameter("@TransactionType", "PURCHASE")
                command.AddParameter("@ReferenceType", "Purchase")
                command.AddParameter("@ReferenceId", purchaseId)
                command.AddParameter("@QuantityIn", item.Quantity + item.FreeQuantity)
                command.AddParameter("@BalanceQuantity", productSnapshot.CurrentStock)
                command.AddParameter("@UnitCost", item.PTR)
                command.AddParameter("@Remarks", $"Received through purchase {purchaseId}")
                command.AddParameter("@TransactionDate", SqliteDateHelper.ToStorageDate(purchaseDate))
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub UpdateSupplierOutstanding(connection As DbConnection, transaction As DbTransaction, supplierId As Integer, amountToAdd As Decimal)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "UPDATE Suppliers SET OutstandingBalance = OutstandingBalance + @AmountToAdd, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@AmountToAdd", amountToAdd)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", supplierId)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Function GetSetting(connection As DbConnection, settingKey As String, defaultValue As String) As String
            Using command = connection.CreateCommand()
                command.CommandText = "SELECT SettingValue FROM Settings WHERE SettingKey = @SettingKey LIMIT 1;"
                command.AddParameter("@SettingKey", settingKey)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Return defaultValue
                End If

                Return Convert.ToString(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function ToNullableStorageDate(dateValue As DateTime?) As Object
            If Not dateValue.HasValue Then
                Return DBNull.Value
            End If

            Return SqliteDateHelper.ToStorageDate(dateValue.Value.Date)
        End Function

        Private Function ParseDate(value As Object) As DateTime
            Dim parsedDate As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
                Return parsedDate
            End If

            Return DateTime.Today
        End Function

        Private Function ParseNullableDate(value As Object) As DateTime?
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
