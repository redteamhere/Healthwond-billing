Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class CustomerRepository
        Implements ICustomerRepository

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function Search(searchTerm As String) As List(Of CustomerRecord) Implements ICustomerRepository.Search
            Dim customers As New List(Of CustomerRecord)()
            Dim normalizedSearch As String = If(searchTerm, String.Empty).Trim()
            Dim searchLike As String = $"%{normalizedSearch}%"

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT Id, CustomerName, Gstin, DrugLicenseNumber, Address, Phone, Email, OutstandingBalance, CreatedAt, UpdatedAt " &
                        "FROM Customers " &
                        "WHERE @Search = '' OR CustomerName LIKE @SearchLike OR Gstin LIKE @SearchLike OR Phone LIKE @SearchLike OR Email LIKE @SearchLike OR DrugLicenseNumber LIKE @SearchLike " &
                        "ORDER BY CustomerName ASC;"
                    command.AddParameter("@Search", normalizedSearch)
                    command.AddParameter("@SearchLike", searchLike)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            customers.Add(MapCustomer(reader))
                        End While
                    End Using
                End Using
            End Using

            Return customers
        End Function

        Public Function Save(customer As CustomerRecord) As Integer Implements ICustomerRepository.Save
            Using connection = _connectionFactory.CreateOpenConnection()
                If customer.Id > 0 Then
                    Using command = connection.CreateCommand()
                        command.CommandText =
                            "UPDATE Customers SET CustomerName = @CustomerName, Gstin = @Gstin, DrugLicenseNumber = @DrugLicenseNumber, Address = @Address, Phone = @Phone, Email = @Email, OutstandingBalance = @OutstandingBalance, UpdatedAt = @UpdatedAt " &
                            "WHERE Id = @Id;"
                        AddCustomerParameters(command, customer)
                        command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                        command.AddParameter("@Id", customer.Id)

                        If command.ExecuteNonQuery() = 0 Then
                            Throw New InvalidOperationException("The selected customer could not be updated.")
                        End If

                        Return customer.Id
                    End Using
                End If

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "INSERT INTO Customers (CustomerName, Gstin, DrugLicenseNumber, Address, Phone, Email, OutstandingBalance, CreatedAt, UpdatedAt) " &
                        "VALUES (@CustomerName, @Gstin, @DrugLicenseNumber, @Address, @Phone, @Email, @OutstandingBalance, @CreatedAt, @UpdatedAt);" &
                        "SELECT last_insert_rowid();"
                    AddCustomerParameters(command, customer)
                    command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function Delete(customerId As Integer) As Boolean Implements ICustomerRepository.Delete
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText = "DELETE FROM Customers WHERE Id = @Id;"
                    command.AddParameter("@Id", customerId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Sub AddCustomerParameters(command As DbCommand, customer As CustomerRecord)
            command.AddParameter("@CustomerName", customer.CustomerName)
            command.AddParameter("@Gstin", customer.Gstin)
            command.AddParameter("@DrugLicenseNumber", customer.DrugLicenseNumber)
            command.AddParameter("@Address", customer.Address)
            command.AddParameter("@Phone", customer.Phone)
            command.AddParameter("@Email", customer.Email)
            command.AddParameter("@OutstandingBalance", customer.OutstandingBalance)
        End Sub

        Private Function MapCustomer(reader As DbDataReader) As CustomerRecord
            Return New CustomerRecord With {
                .Id = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                .CustomerName = Convert.ToString(reader("CustomerName"), CultureInfo.InvariantCulture),
                .Gstin = ConvertNullableString(reader("Gstin")),
                .DrugLicenseNumber = ConvertNullableString(reader("DrugLicenseNumber")),
                .Address = ConvertNullableString(reader("Address")),
                .Phone = ConvertNullableString(reader("Phone")),
                .Email = ConvertNullableString(reader("Email")),
                .OutstandingBalance = Convert.ToDecimal(reader("OutstandingBalance"), CultureInfo.InvariantCulture),
                .CreatedAt = ParseNullableDateTime(reader("CreatedAt")),
                .UpdatedAt = ParseNullableDateTime(reader("UpdatedAt"))
            }
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
