Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class PaymentRepository
        Implements IPaymentRepository

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function GenerateNextCustomerReceiptNumber(paymentDate As DateTime) As String Implements IPaymentRepository.GenerateNextCustomerReceiptNumber
            Using connection = _connectionFactory.CreateOpenConnection()
                Return GenerateNextDocumentNumber(connection, "ReceiptPrefix", "RCPT", "CustomerPayments", "ReceiptNumber", paymentDate)
            End Using
        End Function

        Public Function GenerateNextSupplierPaymentNumber(paymentDate As DateTime) As String Implements IPaymentRepository.GenerateNextSupplierPaymentNumber
            Using connection = _connectionFactory.CreateOpenConnection()
                Return GenerateNextDocumentNumber(connection, "SupplierPaymentPrefix", "SPAY", "SupplierPayments", "PaymentNumber", paymentDate)
            End Using
        End Function

        Public Function SearchCustomerPayments(customerId As Integer, fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of CustomerPaymentHistoryRow) Implements IPaymentRepository.SearchCustomerPayments
            Dim rows As New List(Of CustomerPaymentHistoryRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT cp.Id, cp.ReceiptNumber, cp.PaymentDate, cp.CustomerId, c.CustomerName, COALESCE(cp.PaymentMode, '') AS PaymentMode, " &
                        "COALESCE(cp.ReferenceNumber, '') AS ReferenceNumber, cp.Amount, cp.BalanceAfterPayment, cp.UpdatedAt " &
                        "FROM CustomerPayments cp " &
                        "INNER JOIN Customers c ON c.Id = cp.CustomerId " &
                        "WHERE date(cp.PaymentDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "AND (@CustomerId = 0 OR cp.CustomerId = @CustomerId) " &
                        "AND (@Search = '' OR cp.ReceiptNumber LIKE @SearchLike OR c.CustomerName LIKE @SearchLike OR COALESCE(cp.PaymentMode, '') LIKE @SearchLike OR COALESCE(cp.ReferenceNumber, '') LIKE @SearchLike OR COALESCE(cp.Notes, '') LIKE @SearchLike) " &
                        "ORDER BY date(cp.PaymentDate) DESC, cp.Id DESC;"
                    command.AddParameter("@FromDate", SqliteDateHelper.ToStorageDate(fromDate.Date))
                    command.AddParameter("@ToDate", SqliteDateHelper.ToStorageDate(toDate.Date))
                    command.AddParameter("@CustomerId", customerId)
                    command.AddParameter("@Search", searchTerm)
                    command.AddParameter("@SearchLike", $"%{searchTerm}%")

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New CustomerPaymentHistoryRow With {
                                .PaymentId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                                .ReceiptNumber = Convert.ToString(reader("ReceiptNumber"), CultureInfo.InvariantCulture),
                                .PaymentDate = ParseDate(reader("PaymentDate")),
                                .CustomerId = Convert.ToInt32(reader("CustomerId"), CultureInfo.InvariantCulture),
                                .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                                .PaymentMode = ConvertNullableString(reader("PaymentMode")),
                                .ReferenceNumber = ConvertNullableString(reader("ReferenceNumber")),
                                .Amount = Convert.ToDecimal(reader("Amount"), CultureInfo.InvariantCulture),
                                .BalanceAfterPayment = Convert.ToDecimal(reader("BalanceAfterPayment"), CultureInfo.InvariantCulture),
                                .UpdatedAt = ParseDateTime(reader("UpdatedAt"))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function SearchSupplierPayments(supplierId As Integer, fromDate As DateTime, toDate As DateTime, searchTerm As String) As List(Of SupplierPaymentHistoryRow) Implements IPaymentRepository.SearchSupplierPayments
            Dim rows As New List(Of SupplierPaymentHistoryRow)()

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT sp.Id, sp.PaymentNumber, sp.PaymentDate, sp.SupplierId, s.SupplierName, COALESCE(sp.PaymentMode, '') AS PaymentMode, " &
                        "COALESCE(sp.ReferenceNumber, '') AS ReferenceNumber, sp.Amount, sp.BalanceAfterPayment, sp.UpdatedAt " &
                        "FROM SupplierPayments sp " &
                        "INNER JOIN Suppliers s ON s.Id = sp.SupplierId " &
                        "WHERE date(sp.PaymentDate) BETWEEN date(@FromDate) AND date(@ToDate) " &
                        "AND (@SupplierId = 0 OR sp.SupplierId = @SupplierId) " &
                        "AND (@Search = '' OR sp.PaymentNumber LIKE @SearchLike OR s.SupplierName LIKE @SearchLike OR COALESCE(sp.PaymentMode, '') LIKE @SearchLike OR COALESCE(sp.ReferenceNumber, '') LIKE @SearchLike OR COALESCE(sp.Notes, '') LIKE @SearchLike) " &
                        "ORDER BY date(sp.PaymentDate) DESC, sp.Id DESC;"
                    command.AddParameter("@FromDate", SqliteDateHelper.ToStorageDate(fromDate.Date))
                    command.AddParameter("@ToDate", SqliteDateHelper.ToStorageDate(toDate.Date))
                    command.AddParameter("@SupplierId", supplierId)
                    command.AddParameter("@Search", searchTerm)
                    command.AddParameter("@SearchLike", $"%{searchTerm}%")

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            rows.Add(New SupplierPaymentHistoryRow With {
                                .PaymentId = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                                .PaymentNumber = Convert.ToString(reader("PaymentNumber"), CultureInfo.InvariantCulture),
                                .PaymentDate = ParseDate(reader("PaymentDate")),
                                .SupplierId = Convert.ToInt32(reader("SupplierId"), CultureInfo.InvariantCulture),
                                .SupplierName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
                                .PaymentMode = ConvertNullableString(reader("PaymentMode")),
                                .ReferenceNumber = ConvertNullableString(reader("ReferenceNumber")),
                                .Amount = Convert.ToDecimal(reader("Amount"), CultureInfo.InvariantCulture),
                                .BalanceAfterPayment = Convert.ToDecimal(reader("BalanceAfterPayment"), CultureInfo.InvariantCulture),
                                .UpdatedAt = ParseDateTime(reader("UpdatedAt"))
                            })
                        End While
                    End Using
                End Using
            End Using

            Return rows
        End Function

        Public Function SaveCustomerPayment(draft As CustomerPaymentDraft, createdByUserId As Integer) As Integer Implements IPaymentRepository.SaveCustomerPayment
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim currentOutstanding As Decimal = LoadCustomerOutstanding(connection, transaction, draft.CustomerId)
                    Dim paymentAmount As Decimal = Decimal.Round(draft.Amount, 2, MidpointRounding.AwayFromZero)

                    If currentOutstanding <= 0D Then
                        Throw New InvalidOperationException("The selected customer does not have an outstanding balance.")
                    End If

                    If paymentAmount > currentOutstanding Then
                        Throw New InvalidOperationException("Received amount cannot exceed the current outstanding balance.")
                    End If

                    Dim balanceAfterPayment As Decimal = Decimal.Round(currentOutstanding - paymentAmount, 2, MidpointRounding.AwayFromZero)
                    Dim paymentId As Integer = InsertCustomerPayment(connection, transaction, draft, paymentAmount, balanceAfterPayment, createdByUserId)
                    UpdateCustomerOutstanding(connection, transaction, draft.CustomerId, balanceAfterPayment)
                    AccountingPostingHelper.PostCustomerPaymentVoucher(connection, transaction, paymentId, draft, createdByUserId)
                    transaction.Commit()
                    Return paymentId
                End Using
            End Using
        End Function

        Public Function SaveSupplierPayment(draft As SupplierPaymentDraft, createdByUserId As Integer) As Integer Implements IPaymentRepository.SaveSupplierPayment
            Using connection = _connectionFactory.CreateOpenConnection()
                Using transaction = connection.BeginTransaction()
                    Dim currentOutstanding As Decimal = LoadSupplierOutstanding(connection, transaction, draft.SupplierId)
                    Dim paymentAmount As Decimal = Decimal.Round(draft.Amount, 2, MidpointRounding.AwayFromZero)

                    If currentOutstanding <= 0D Then
                        Throw New InvalidOperationException("The selected supplier does not have an outstanding balance.")
                    End If

                    If paymentAmount > currentOutstanding Then
                        Throw New InvalidOperationException("Payment amount cannot exceed the supplier outstanding balance.")
                    End If

                    Dim balanceAfterPayment As Decimal = Decimal.Round(currentOutstanding - paymentAmount, 2, MidpointRounding.AwayFromZero)
                    Dim paymentId As Integer = InsertSupplierPayment(connection, transaction, draft, paymentAmount, balanceAfterPayment, createdByUserId)
                    UpdateSupplierOutstanding(connection, transaction, draft.SupplierId, balanceAfterPayment)
                    AccountingPostingHelper.PostSupplierPaymentVoucher(connection, transaction, paymentId, draft, createdByUserId)
                    transaction.Commit()
                    Return paymentId
                End Using
            End Using
        End Function

        Private Function GenerateNextDocumentNumber(connection As DbConnection, settingKey As String, defaultPrefix As String, tableName As String, columnName As String, documentDate As DateTime) As String
            Dim prefix As String = GetSetting(connection, settingKey, defaultPrefix)
            Dim monthPart As String = documentDate.ToString("yyyyMM", CultureInfo.InvariantCulture)
            Dim pattern As String = $"{prefix}-{monthPart}-%"
            Dim lastNumber As Integer = 0

            Using command = connection.CreateCommand()
                command.CommandText = $"SELECT {columnName} FROM {tableName} WHERE {columnName} LIKE @Pattern ORDER BY Id DESC LIMIT 1;"
                command.AddParameter("@Pattern", pattern)
                Dim result As Object = command.ExecuteScalar()
                If result IsNot Nothing AndAlso result IsNot DBNull.Value Then
                    Dim documentNumber As String = Convert.ToString(result, CultureInfo.InvariantCulture)
                    Dim parts As String() = documentNumber.Split("-"c)
                    Dim parsedNumber As Integer
                    If parts.Length >= 3 AndAlso Integer.TryParse(parts(parts.Length - 1), parsedNumber) Then
                        lastNumber = parsedNumber
                    End If
                End If
            End Using

            Return $"{prefix}-{monthPart}-{(lastNumber + 1).ToString("0000", CultureInfo.InvariantCulture)}"
        End Function

        Private Function InsertCustomerPayment(connection As DbConnection, transaction As DbTransaction, draft As CustomerPaymentDraft, amount As Decimal, balanceAfterPayment As Decimal, createdByUserId As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO CustomerPayments (ReceiptNumber, CustomerId, PaymentDate, PaymentMode, ReferenceNumber, Amount, BalanceAfterPayment, Notes, CreatedBy, CreatedAt, UpdatedAt) " &
                    "VALUES (@ReceiptNumber, @CustomerId, @PaymentDate, @PaymentMode, @ReferenceNumber, @Amount, @BalanceAfterPayment, @Notes, @CreatedBy, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@ReceiptNumber", draft.ReceiptNumber)
                command.AddParameter("@CustomerId", draft.CustomerId)
                command.AddParameter("@PaymentDate", SqliteDateHelper.ToStorageDate(draft.PaymentDate))
                command.AddParameter("@PaymentMode", draft.PaymentMode)
                command.AddParameter("@ReferenceNumber", draft.ReferenceNumber)
                command.AddParameter("@Amount", amount)
                command.AddParameter("@BalanceAfterPayment", balanceAfterPayment)
                command.AddParameter("@Notes", draft.Notes)
                command.AddParameter("@CreatedBy", createdByUserId)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function InsertSupplierPayment(connection As DbConnection, transaction As DbTransaction, draft As SupplierPaymentDraft, amount As Decimal, balanceAfterPayment As Decimal, createdByUserId As Integer) As Integer
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText =
                    "INSERT INTO SupplierPayments (PaymentNumber, SupplierId, PaymentDate, PaymentMode, ReferenceNumber, Amount, BalanceAfterPayment, Notes, CreatedBy, CreatedAt, UpdatedAt) " &
                    "VALUES (@PaymentNumber, @SupplierId, @PaymentDate, @PaymentMode, @ReferenceNumber, @Amount, @BalanceAfterPayment, @Notes, @CreatedBy, @CreatedAt, @UpdatedAt);" &
                    "SELECT last_insert_rowid();"
                command.AddParameter("@PaymentNumber", draft.PaymentNumber)
                command.AddParameter("@SupplierId", draft.SupplierId)
                command.AddParameter("@PaymentDate", SqliteDateHelper.ToStorageDate(draft.PaymentDate))
                command.AddParameter("@PaymentMode", draft.PaymentMode)
                command.AddParameter("@ReferenceNumber", draft.ReferenceNumber)
                command.AddParameter("@Amount", amount)
                command.AddParameter("@BalanceAfterPayment", balanceAfterPayment)
                command.AddParameter("@Notes", draft.Notes)
                command.AddParameter("@CreatedBy", createdByUserId)
                command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function LoadCustomerOutstanding(connection As DbConnection, transaction As DbTransaction, customerId As Integer) As Decimal
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT OutstandingBalance FROM Customers WHERE Id = @Id LIMIT 1;"
                command.AddParameter("@Id", customerId)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Throw New InvalidOperationException("The selected customer could not be found.")
                End If

                Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Function LoadSupplierOutstanding(connection As DbConnection, transaction As DbTransaction, supplierId As Integer) As Decimal
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "SELECT OutstandingBalance FROM Suppliers WHERE Id = @Id LIMIT 1;"
                command.AddParameter("@Id", supplierId)
                Dim result As Object = command.ExecuteScalar()
                If result Is Nothing OrElse result Is DBNull.Value Then
                    Throw New InvalidOperationException("The selected supplier could not be found.")
                End If

                Return Convert.ToDecimal(result, CultureInfo.InvariantCulture)
            End Using
        End Function

        Private Sub UpdateCustomerOutstanding(connection As DbConnection, transaction As DbTransaction, customerId As Integer, balanceAfterPayment As Decimal)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "UPDATE Customers SET OutstandingBalance = @OutstandingBalance, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@OutstandingBalance", balanceAfterPayment)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", customerId)

                If command.ExecuteNonQuery() <> 1 Then
                    Throw New InvalidOperationException("Customer outstanding balance could not be updated.")
                End If
            End Using
        End Sub

        Private Sub UpdateSupplierOutstanding(connection As DbConnection, transaction As DbTransaction, supplierId As Integer, balanceAfterPayment As Decimal)
            Using command = connection.CreateCommand()
                command.Transaction = transaction
                command.CommandText = "UPDATE Suppliers SET OutstandingBalance = @OutstandingBalance, UpdatedAt = @UpdatedAt WHERE Id = @Id;"
                command.AddParameter("@OutstandingBalance", balanceAfterPayment)
                command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                command.AddParameter("@Id", supplierId)

                If command.ExecuteNonQuery() <> 1 Then
                    Throw New InvalidOperationException("Supplier outstanding balance could not be updated.")
                End If
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

        Private Function ParseDate(value As Object) As DateTime
            Dim parsedValue As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedValue) Then
                Return parsedValue
            End If

            Return DateTime.Today
        End Function

        Private Function ParseDateTime(value As Object) As DateTime
            Dim parsedValue As DateTime
            If DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedValue) Then
                Return parsedValue
            End If

            Return DateTime.Now
        End Function

        Private Function ConvertNullableString(value As Object) As String
            If value Is Nothing OrElse value Is DBNull.Value Then
                Return String.Empty
            End If

            Return Convert.ToString(value, CultureInfo.InvariantCulture)
        End Function

    End Class

End Namespace
