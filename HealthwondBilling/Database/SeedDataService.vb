Imports HealthwondBilling.Utilities
Imports System.Data.Common

Namespace Database

    Public Class SeedDataService

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Sub Seed()
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    EnsureDefaultSettings(connection, transaction)
                    EnsureDefaultAccounts(connection, transaction)
                    EnsureDefaultUsers(connection, transaction)
                    EnsureSampleSuppliers(connection, transaction)
                    EnsureSampleCustomers(connection, transaction)
                    EnsureSampleProducts(connection, transaction)
                    transaction.Commit()
                End Using
            End Using
        End Sub

        Private Sub EnsureDefaultSettings(connection As DbConnection, transaction As DbTransaction)
            Dim createdAt As String = SqliteDateHelper.ToStorageDateTime(DateTime.Now)

            InsertSetting(connection, transaction, "CompanyName", "Healthwond Pharmacy", "Company display name used on invoices.", createdAt)
            InsertSetting(connection, transaction, "CompanyAddress", "88 Medical Avenue, Bengaluru", "Primary company address displayed on invoices.", createdAt)
            InsertSetting(connection, transaction, "CompanyPhone", "080-4000-1122", "Primary company phone displayed on invoices.", createdAt)
            InsertSetting(connection, transaction, "CompanyGstin", "29AAACH2024H1ZX", "Company GSTIN displayed on invoices.", createdAt)
            InsertSetting(connection, transaction, "CompanyDrugLicense", "DL-HWB-2024-01", "Company drug license displayed on invoices.", createdAt)
            InsertSetting(connection, transaction, "InvoicePrefix", "HWB", "Prefix for auto-generated invoice numbers.", createdAt)
            InsertSetting(connection, transaction, "PurchasePrefix", "PUR", "Prefix for auto-generated purchase numbers.", createdAt)
            InsertSetting(connection, transaction, "ReceiptPrefix", "RCPT", "Prefix for auto-generated customer collection receipts.", createdAt)
            InsertSetting(connection, transaction, "SupplierPaymentPrefix", "SPAY", "Prefix for auto-generated supplier payment numbers.", createdAt)
            InsertSetting(connection, transaction, "PurchaseReturnPrefix", "PRN", "Prefix for auto-generated purchase return numbers.", createdAt)
            InsertSetting(connection, transaction, "StockAdjustmentPrefix", "ADJ", "Prefix for auto-generated stock adjustment numbers.", createdAt)
            InsertSetting(connection, transaction, "LowStockThreshold", "10", "Default low stock alert threshold.", createdAt)
            InsertSetting(connection, transaction, "CurrencySymbol", "Rs.", "Default invoice currency symbol.", createdAt)
            InsertSetting(connection, transaction, "InvoiceTemplatePath", "Templates/GSTInvoiceTemplate.xlsx", "Default GST invoice template path.", createdAt)
        End Sub

        Private Sub EnsureDefaultAccounts(connection As DbConnection, transaction As DbTransaction)
            Dim createdAt As String = SqliteDateHelper.ToStorageDateTime(DateTime.Now)

            InsertAccountGroup(connection, transaction, "Cash & Bank", "Asset", 10, createdAt)
            InsertAccountGroup(connection, transaction, "Sundry Debtors", "Asset", 20, createdAt)
            InsertAccountGroup(connection, transaction, "Input Taxes", "Asset", 30, createdAt)
            InsertAccountGroup(connection, transaction, "Sundry Creditors", "Liability", 40, createdAt)
            InsertAccountGroup(connection, transaction, "Output Taxes", "Liability", 50, createdAt)
            InsertAccountGroup(connection, transaction, "Sales Accounts", "Income", 60, createdAt)
            InsertAccountGroup(connection, transaction, "Purchase Accounts", "Expense", 70, createdAt)
            InsertAccountGroup(connection, transaction, "Indirect Expenses", "Expense", 80, createdAt)

            InsertSystemLedger(connection, transaction, "Cash in Hand", "Cash & Bank", createdAt)
            InsertSystemLedger(connection, transaction, "Bank Account", "Cash & Bank", createdAt)
            InsertSystemLedger(connection, transaction, "Output GST", "Output Taxes", createdAt)
            InsertSystemLedger(connection, transaction, "Input GST", "Input Taxes", createdAt)
            InsertSystemLedger(connection, transaction, "Sales Account", "Sales Accounts", createdAt)
            InsertSystemLedger(connection, transaction, "Purchase Account", "Purchase Accounts", createdAt)
            InsertSystemLedger(connection, transaction, "Purchase Return Account", "Purchase Accounts", createdAt)
            InsertSystemLedger(connection, transaction, "Round Off", "Indirect Expenses", createdAt)
        End Sub

        Private Sub InsertSetting(connection As DbConnection, transaction As DbTransaction, key As String, value As String, description As String, updatedAt As String)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT OR IGNORE INTO Settings (SettingKey, SettingValue, Description, UpdatedAt) " &
                    "VALUES (@SettingKey, @SettingValue, @Description, @UpdatedAt);"
                command.AddParameter("@SettingKey", key)
                command.AddParameter("@SettingValue", value)
                command.AddParameter("@Description", description)
                command.AddParameter("@UpdatedAt", updatedAt)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub InsertAccountGroup(connection As DbConnection, transaction As DbTransaction, groupName As String, nature As String, displayOrder As Integer, createdAt As String)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT OR IGNORE INTO AccountGroups (GroupName, Nature, DisplayOrder, IsSystem, CreatedAt, UpdatedAt) " &
                    "VALUES (@GroupName, @Nature, @DisplayOrder, 1, @CreatedAt, @UpdatedAt);"
                command.AddParameter("@GroupName", groupName)
                command.AddParameter("@Nature", nature)
                command.AddParameter("@DisplayOrder", displayOrder)
                command.AddParameter("@CreatedAt", createdAt)
                command.AddParameter("@UpdatedAt", createdAt)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub InsertSystemLedger(connection As DbConnection, transaction As DbTransaction, ledgerName As String, groupName As String, createdAt As String)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT OR IGNORE INTO Ledgers (LedgerName, AccountGroupId, OpeningBalance, OpeningBalanceType, IsSystem, IsPartyLedger, LinkedEntityType, LinkedEntityId, Notes, CreatedAt, UpdatedAt) " &
                    "VALUES (@LedgerName, (SELECT Id FROM AccountGroups WHERE GroupName = @GroupName LIMIT 1), 0, 'Dr', 1, 0, NULL, NULL, '', @CreatedAt, @UpdatedAt);"
                command.AddParameter("@LedgerName", ledgerName)
                command.AddParameter("@GroupName", groupName)
                command.AddParameter("@CreatedAt", createdAt)
                command.AddParameter("@UpdatedAt", createdAt)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub EnsureDefaultUsers(connection As DbConnection, transaction As DbTransaction)
            If GetTableCount(connection, transaction, "Users") > 0 Then
                Return
            End If

            InsertUser(connection, transaction, "admin", "Admin@123", "System Administrator", "Admin")
            InsertUser(connection, transaction, "staff", "Staff@123", "Billing Staff", "Staff")
        End Sub

        Private Sub InsertUser(connection As DbConnection, transaction As DbTransaction, username As String, password As String, fullName As String, role As String)
            Dim hashResult As PasswordHashResult = PasswordHasher.HashPassword(password)
            Dim timestamp As String = SqliteDateHelper.ToStorageDateTime(DateTime.Now)

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Users (Username, PasswordHash, PasswordSalt, FullName, Role, IsActive, CreatedAt, UpdatedAt) " &
                    "VALUES (@Username, @PasswordHash, @PasswordSalt, @FullName, @Role, 1, @CreatedAt, @UpdatedAt);"
                command.AddParameter("@Username", username)
                command.AddParameter("@PasswordHash", hashResult.Hash)
                command.AddParameter("@PasswordSalt", hashResult.Salt)
                command.AddParameter("@FullName", fullName)
                command.AddParameter("@Role", role)
                command.AddParameter("@CreatedAt", timestamp)
                command.AddParameter("@UpdatedAt", timestamp)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub EnsureSampleSuppliers(connection As DbConnection, transaction As DbTransaction)
            If GetTableCount(connection, transaction, "Suppliers") > 0 Then
                Return
            End If

            Dim timestamp As String = SqliteDateHelper.ToStorageDateTime(DateTime.Now)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Suppliers (SupplierName, Gstin, DrugLicenseNumber, Address, Phone, Email, OutstandingBalance, CreatedAt, UpdatedAt) " &
                    "VALUES " &
                    "(@Name1, @Gstin1, @License1, @Address1, @Phone1, @Email1, @Balance1, @CreatedAt, @UpdatedAt)," &
                    "(@Name2, @Gstin2, @License2, @Address2, @Phone2, @Email2, @Balance2, @CreatedAt, @UpdatedAt);"
                command.AddParameter("@Name1", "MediTrade Distributors")
                command.AddParameter("@Gstin1", "29AACCM1234F1Z2")
                command.AddParameter("@License1", "DL-2024-4581")
                command.AddParameter("@Address1", "45 Market Road, Bengaluru")
                command.AddParameter("@Phone1", "9988776655")
                command.AddParameter("@Email1", "accounts@meditrade.example")
                command.AddParameter("@Balance1", 15250D)
                command.AddParameter("@Name2", "Sterling Pharma Hub")
                command.AddParameter("@Gstin2", "29AACCS4455Q1Z7")
                command.AddParameter("@License2", "DL-2024-9923")
                command.AddParameter("@Address2", "12 Industrial Estate, Mysuru")
                command.AddParameter("@Phone2", "9876501234")
                command.AddParameter("@Email2", "sales@sterlinghub.example")
                command.AddParameter("@Balance2", 0D)
                command.AddParameter("@CreatedAt", timestamp)
                command.AddParameter("@UpdatedAt", timestamp)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub EnsureSampleCustomers(connection As DbConnection, transaction As DbTransaction)
            If GetTableCount(connection, transaction, "Customers") > 0 Then
                Return
            End If

            Dim timestamp As String = SqliteDateHelper.ToStorageDateTime(DateTime.Now)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Customers (CustomerName, Gstin, DrugLicenseNumber, Address, Phone, Email, OutstandingBalance, CreatedAt, UpdatedAt) " &
                    "VALUES " &
                    "(@Name1, @Gstin1, @License1, @Address1, @Phone1, @Email1, @Balance1, @CreatedAt, @UpdatedAt)," &
                    "(@Name2, @Gstin2, @License2, @Address2, @Phone2, @Email2, @Balance2, @CreatedAt, @UpdatedAt);"
                command.AddParameter("@Name1", "Sunrise Clinic")
                command.AddParameter("@Gstin1", "29AABCS7788L1ZQ")
                command.AddParameter("@License1", "CL-4479")
                command.AddParameter("@Address1", "88 Hospital Street, Bengaluru")
                command.AddParameter("@Phone1", "9090909090")
                command.AddParameter("@Email1", "purchase@sunriseclinic.example")
                command.AddParameter("@Balance1", 3250D)
                command.AddParameter("@Name2", "Wellcare Medicals")
                command.AddParameter("@Gstin2", "29AABCW9101H1ZT")
                command.AddParameter("@License2", "WM-1032")
                command.AddParameter("@Address2", "14 Residency Road, Bengaluru")
                command.AddParameter("@Phone2", "9345612780")
                command.AddParameter("@Email2", "contact@wellcare.example")
                command.AddParameter("@Balance2", 0D)
                command.AddParameter("@CreatedAt", timestamp)
                command.AddParameter("@UpdatedAt", timestamp)
                command.ExecuteNonQuery()
            End Using
        End Sub

        Private Sub EnsureSampleProducts(connection As DbConnection, transaction As DbTransaction)
            If GetTableCount(connection, transaction, "Products") > 0 Then
                Return
            End If

            InsertSeedProduct(connection, transaction, "Paracetamol 500", "10x15 Tab", "30049069", "PCM2401", DateTime.Today.AddMonths(11), 5D, 32.5D, 24D, 26.5D, "Healthwond Labs", "Paracetamol 500mg", 120, "890100000001")
            InsertSeedProduct(connection, transaction, "Amoxicillin 250", "10x10 Cap", "30041010", "AMX2409", DateTime.Today.AddDays(50), 12D, 96D, 74D, 78D, "MediLife Pharma", "Amoxicillin 250mg", 18, "890100000002")
            InsertSeedProduct(connection, transaction, "Vitamin C Syrup", "100 ml", "21069099", "VCS118", DateTime.Today.AddDays(18), 12D, 82D, 61D, 66D, "NutriWell", "Vitamin C with Zinc", 6, "890100000003")
        End Sub

        Private Sub InsertSeedProduct(connection As DbConnection, transaction As DbTransaction, productName As String, packing As String, hsnCode As String, batchNumber As String, expiryDate As DateTime, gstPercentage As Decimal, mrp As Decimal, ptr As Decimal, pts As Decimal, companyName As String, composition As String, currentStock As Integer, barcode As String)
            Dim timestamp As String = SqliteDateHelper.ToStorageDateTime(DateTime.Now)
            Dim productId As Integer

            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO Products (ProductName, Packing, HsnCode, BatchNumber, ExpiryDate, GstPercentage, MRP, PTR, PTS, CompanyName, Composition, CurrentStock, Barcode, IsDeleted, CreatedAt, UpdatedAt) " &
                    "VALUES (@ProductName, @Packing, @HsnCode, @BatchNumber, @ExpiryDate, @GstPercentage, @MRP, @PTR, @PTS, @CompanyName, @Composition, @CurrentStock, @Barcode, 0, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@ProductName", productName)
                command.AddParameter("@Packing", packing)
                command.AddParameter("@HsnCode", hsnCode)
                command.AddParameter("@BatchNumber", batchNumber)
                command.AddParameter("@ExpiryDate", SqliteDateHelper.ToStorageDate(expiryDate))
                command.AddParameter("@GstPercentage", gstPercentage)
                command.AddParameter("@MRP", mrp)
                command.AddParameter("@PTR", ptr)
                command.AddParameter("@PTS", pts)
                command.AddParameter("@CompanyName", companyName)
                command.AddParameter("@Composition", composition)
                command.AddParameter("@CurrentStock", currentStock)
                command.AddParameter("@Barcode", barcode)
                command.AddParameter("@CreatedAt", timestamp)
                command.AddParameter("@UpdatedAt", timestamp)
                productId = Convert.ToInt32(command.ExecuteScalar())
            End Using

            Using ledgerCommand = connection.CreateCommand()
                ledgerCommand.Transaction = transaction
                ledgerCommand.CommandText =
                    "INSERT INTO StockLedger (ProductId, BatchNumber, TransactionType, ReferenceType, ReferenceId, QuantityIn, QuantityOut, BalanceQuantity, UnitCost, Remarks, TransactionDate, CreatedAt) " &
                    "VALUES (@ProductId, @BatchNumber, @TransactionType, @ReferenceType, @ReferenceId, @QuantityIn, @QuantityOut, @BalanceQuantity, @UnitCost, @Remarks, @TransactionDate, @CreatedAt);"
                ledgerCommand.AddParameter("@ProductId", productId)
                ledgerCommand.AddParameter("@BatchNumber", batchNumber)
                ledgerCommand.AddParameter("@TransactionType", "OPENING")
                ledgerCommand.AddParameter("@ReferenceType", "Seed")
                ledgerCommand.AddParameter("@ReferenceId", 0)
                ledgerCommand.AddParameter("@QuantityIn", currentStock)
                ledgerCommand.AddParameter("@QuantityOut", 0)
                ledgerCommand.AddParameter("@BalanceQuantity", currentStock)
                ledgerCommand.AddParameter("@UnitCost", ptr)
                ledgerCommand.AddParameter("@Remarks", "Initial sample data")
                ledgerCommand.AddParameter("@TransactionDate", SqliteDateHelper.ToStorageDate(DateTime.Today))
                ledgerCommand.AddParameter("@CreatedAt", timestamp)
                ledgerCommand.ExecuteNonQuery()
            End Using
        End Sub

        Private Function GetTableCount(connection As DbConnection, transaction As DbTransaction, tableName As String) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = $"SELECT COUNT(1) FROM {tableName};"
                Return Convert.ToInt32(command.ExecuteScalar())
            End Using
        End Function

    End Class

End Namespace
