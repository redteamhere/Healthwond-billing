Imports HealthwondBilling.Database
Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Data.Common
Imports System.Globalization

Namespace Repositories

    Public Class SupplierRepository
        Implements ISupplierRepository

        Private ReadOnly _connectionFactory As IDbConnectionFactory

        Public Sub New(connectionFactory As IDbConnectionFactory)
            _connectionFactory = connectionFactory
        End Sub

        Public Function Search(searchTerm As String) As List(Of SupplierRecord) Implements ISupplierRepository.Search
            Dim suppliers As New List(Of SupplierRecord)()
            Dim normalizedSearch As String = If(searchTerm, String.Empty).Trim()
            Dim searchLike As String = $"%{normalizedSearch}%"

            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText =
                        "SELECT Id, SupplierName, Gstin, DrugLicenseNumber, Address, Phone, Email, OutstandingBalance, CreatedAt, UpdatedAt " &
                        "FROM Suppliers " &
                        "WHERE @Search = '' OR SupplierName LIKE @SearchLike OR Gstin LIKE @SearchLike OR Phone LIKE @SearchLike OR Email LIKE @SearchLike OR DrugLicenseNumber LIKE @SearchLike " &
                        "ORDER BY SupplierName ASC;"
                    command.AddParameter("@Search", normalizedSearch)
                    command.AddParameter("@SearchLike", searchLike)

                    Using reader = command.ExecuteReader()
                        While reader.Read()
                            suppliers.Add(MapSupplier(reader))
                        End While
                    End Using
                End Using
            End Using

            Return suppliers
        End Function

        Public Function Save(supplier As SupplierRecord) As Integer Implements ISupplierRepository.Save
            Using connection = _connectionFactory.CreateOpenConnection()
                If supplier.Id > 0 Then
                    Using command = connection.CreateCommand()
                        command.CommandText =
                            "UPDATE Suppliers SET SupplierName = @SupplierName, Gstin = @Gstin, DrugLicenseNumber = @DrugLicenseNumber, Address = @Address, Phone = @Phone, Email = @Email, OutstandingBalance = @OutstandingBalance, UpdatedAt = @UpdatedAt " &
                            "WHERE Id = @Id;"
                        AddSupplierParameters(command, supplier)
                        command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                        command.AddParameter("@Id", supplier.Id)

                        If command.ExecuteNonQuery() = 0 Then
                            Throw New InvalidOperationException("The selected supplier could not be updated.")
                        End If

                        Return supplier.Id
                    End Using
                End If

                Using command = connection.CreateCommand()
                    command.CommandText =
                        "INSERT INTO Suppliers (SupplierName, Gstin, DrugLicenseNumber, Address, Phone, Email, OutstandingBalance, CreatedAt, UpdatedAt) " &
                        "VALUES (@SupplierName, @Gstin, @DrugLicenseNumber, @Address, @Phone, @Email, @OutstandingBalance, @CreatedAt, @UpdatedAt);" &
                        "SELECT last_insert_rowid();"
                    AddSupplierParameters(command, supplier)
                    command.AddParameter("@CreatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    command.AddParameter("@UpdatedAt", SqliteDateHelper.ToStorageDateTime(DateTime.Now))
                    Return Convert.ToInt32(command.ExecuteScalar(), CultureInfo.InvariantCulture)
                End Using
            End Using
        End Function

        Public Function Delete(supplierId As Integer) As Boolean Implements ISupplierRepository.Delete
            Using connection = _connectionFactory.CreateOpenConnection()
                Using command = connection.CreateCommand()
                    command.CommandText = "DELETE FROM Suppliers WHERE Id = @Id;"
                    command.AddParameter("@Id", supplierId)
                    Return command.ExecuteNonQuery() > 0
                End Using
            End Using
        End Function

        Private Sub AddSupplierParameters(command As DbCommand, supplier As SupplierRecord)
            command.AddParameter("@SupplierName", supplier.SupplierName)
            command.AddParameter("@Gstin", supplier.Gstin)
            command.AddParameter("@DrugLicenseNumber", supplier.DrugLicenseNumber)
            command.AddParameter("@Address", supplier.Address)
            command.AddParameter("@Phone", supplier.Phone)
            command.AddParameter("@Email", supplier.Email)
            command.AddParameter("@OutstandingBalance", supplier.OutstandingBalance)
        End Sub

        Private Function MapSupplier(reader As DbDataReader) As SupplierRecord
            Return New SupplierRecord With {
                .Id = Convert.ToInt32(reader("Id"), CultureInfo.InvariantCulture),
                .SupplierName = Convert.ToString(reader("SupplierName"), CultureInfo.InvariantCulture),
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
