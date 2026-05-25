Imports HealthwondBilling.Models
Imports HealthwondBilling.Utilities
Imports System.Security.Cryptography
Imports System.Text
Imports System.Linq
Imports System.Xml.Linq

Namespace Services

    Public Class SavedCredentialService

        Public Function Load(companyId As String) As SavedCredentialRecord
            AppPaths.EnsureDirectories()

            If Not IO.File.Exists(AppPaths.SavedCredentialsFilePath) Then
                Return Nothing
            End If

            Dim document As XDocument = XDocument.Load(AppPaths.SavedCredentialsFilePath)
            Dim credentialElement As XElement =
                document.Root.<Credential>.
                    FirstOrDefault(Function(element) String.Equals(Convert.ToString(element.@companyId), companyId, StringComparison.OrdinalIgnoreCase))

            If credentialElement Is Nothing Then
                Return Nothing
            End If

            Dim protectedValue As String = Convert.ToString(credentialElement.@secret)
            If String.IsNullOrWhiteSpace(protectedValue) Then
                Return Nothing
            End If

            Try
                Return New SavedCredentialRecord With {
                    .CompanyId = companyId,
                    .Username = Convert.ToString(credentialElement.@username),
                    .Password = DecryptPassword(protectedValue)
                }
            Catch ex As Exception
                AppLogger.Warn($"Saved credentials for company '{companyId}' could not be opened. They will be ignored.")
                Return Nothing
            End Try
        End Function

        Public Sub Save(companyId As String, username As String, password As String)
            AppPaths.EnsureDirectories()

            Dim document As XDocument = LoadDocument()
            Dim root As XElement = document.Root
            Dim existing As XElement =
                root.<Credential>.
                    FirstOrDefault(Function(element) String.Equals(Convert.ToString(element.@companyId), companyId, StringComparison.OrdinalIgnoreCase))

            If existing IsNot Nothing Then
                existing.Remove()
            End If

            root.Add(New XElement(
                "Credential",
                New XAttribute("companyId", companyId),
                New XAttribute("username", username.Trim()),
                New XAttribute("secret", EncryptPassword(password))))
            document.Save(AppPaths.SavedCredentialsFilePath)
        End Sub

        Public Function Clear(companyId As String) As Boolean
            AppPaths.EnsureDirectories()

            If Not IO.File.Exists(AppPaths.SavedCredentialsFilePath) Then
                Return False
            End If

            Dim document As XDocument = XDocument.Load(AppPaths.SavedCredentialsFilePath)
            Dim credentialElement As XElement =
                document.Root.<Credential>.
                    FirstOrDefault(Function(element) String.Equals(Convert.ToString(element.@companyId), companyId, StringComparison.OrdinalIgnoreCase))

            If credentialElement Is Nothing Then
                Return False
            End If

            credentialElement.Remove()
            document.Save(AppPaths.SavedCredentialsFilePath)
            Return True
        End Function

        Private Function LoadDocument() As XDocument
            If IO.File.Exists(AppPaths.SavedCredentialsFilePath) Then
                Return XDocument.Load(AppPaths.SavedCredentialsFilePath)
            End If

            Return New XDocument(New XElement("SavedCredentials"))
        End Function

        Private Function EncryptPassword(password As String) As String
            Dim plainBytes As Byte() = Encoding.UTF8.GetBytes(password)

            Using algorithm As Aes = Aes.Create()
                algorithm.Key = CreateEncryptionKey()
                algorithm.GenerateIV()

                Using encryptor As ICryptoTransform = algorithm.CreateEncryptor()
                    Dim cipherBytes As Byte() = encryptor.TransformFinalBlock(plainBytes, 0, plainBytes.Length)
                    Dim payload(cipherBytes.Length + algorithm.IV.Length - 1) As Byte
                    Buffer.BlockCopy(algorithm.IV, 0, payload, 0, algorithm.IV.Length)
                    Buffer.BlockCopy(cipherBytes, 0, payload, algorithm.IV.Length, cipherBytes.Length)
                    Return Convert.ToBase64String(payload)
                End Using
            End Using
        End Function

        Private Function DecryptPassword(payloadText As String) As String
            Dim payload As Byte() = Convert.FromBase64String(payloadText)

            Using algorithm As Aes = Aes.Create()
                algorithm.Key = CreateEncryptionKey()

                Dim iv(algorithm.BlockSize \ 8 - 1) As Byte
                Buffer.BlockCopy(payload, 0, iv, 0, iv.Length)
                algorithm.IV = iv

                Dim cipherBytes(payload.Length - iv.Length - 1) As Byte
                Buffer.BlockCopy(payload, iv.Length, cipherBytes, 0, cipherBytes.Length)

                Using decryptor As ICryptoTransform = algorithm.CreateDecryptor()
                    Dim plainBytes As Byte() = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length)
                    Return Encoding.UTF8.GetString(plainBytes)
                End Using
            End Using
        End Function

        Private Function CreateEncryptionKey() As Byte()
            Dim source As String = $"{Environment.MachineName}|{Environment.UserName}|HealthwondBilling"
            Using sha256 As SHA256 = SHA256.Create()
                Return sha256.ComputeHash(Encoding.UTF8.GetBytes(source))
            End Using
        End Function

    End Class

End Namespace
